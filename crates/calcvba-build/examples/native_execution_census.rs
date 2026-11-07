//! Execute isolated, deterministic, zero-argument conformance procedures on
//! VM3, JIT and AOT. Admission and execution are reported separately. Each child
//! owns its runtime thread and DLLs; the supervisor bounds crashes and loops.
use calcvba_build::{NativeLibraryBuildOptions, build_native_library, build_oximage};
use calcvba_bundle::{VarTypeRef, coreir::CoreProc};
use calcvba_engine::{Engine, HostConfig};
use calcvba_oxir::{OxImage, OxInst, OxNativeCallee};
use calcvba_symbol::manifest::{
    ModuleAttributes, ModuleKind, ModuleUnit, ProjectKind, SymbolProjectManifest,
};
use calcvba_syntax::{SyntaxKind, SyntaxNode};
use serde::{Deserialize, Serialize};
use std::{
    path::{Path, PathBuf},
    process::{Child, Command},
    time::{Duration, Instant},
};

#[derive(Debug, Serialize, Deserialize)]
struct Row {
    file: String,
    procedure: String,
    status: String,
    detail: String,
    vm3: Option<String>,
    jit: Option<String>,
    aot: Option<String>,
}
impl Row {
    fn skipped(path: &Path, procedure: &str, detail: impl ToString) -> Self {
        Self {
            file: path.display().to_string(),
            procedure: procedure.into(),
            status: "not-run".into(),
            detail: detail.to_string(),
            vm3: None,
            jit: None,
            aot: None,
        }
    }
}
const OBSERVED: &str = "CvbCensusObserved";
const PROBE: &str = "CvbCensusProbe";
const VALUE_HELPER: &str = r#"
Private Function CvbCensusValue(ByVal v As Variant) As String
    Dim item As Variant, text As String, dimension As Long, lo As Long, hi As Long
    On Error GoTo unavailable
    If IsObject(v) Then
        CvbCensusValue = "object:" & TypeName(v)
    ElseIf IsArray(v) Then
        text = "array:" & CStr(VarType(v)) & ":"
        For dimension = 1 To 8
            On Error Resume Next
            Err.Clear
            lo = LBound(v, dimension)
            hi = UBound(v, dimension)
            If Err.Number <> 0 Then Exit For
            On Error GoTo unavailable
            text = text & CStr(lo) & "," & CStr(hi) & ";"
        Next
        Err.Clear
        On Error GoTo unavailable
        For Each item In v
            text = text & CvbCensusValue(item) & ";"
        Next
        CvbCensusValue = text
    ElseIf IsNull(v) Then
        CvbCensusValue = "null"
    ElseIf IsEmpty(v) Then
        CvbCensusValue = "empty"
    Else
        text = CStr(v)
        CvbCensusValue = CStr(VarType(v)) & ":" & CStr(Len(text)) & ":" & text
    End If
    Exit Function
unavailable:
    CvbCensusValue = "unavailable:" & CStr(Err.Number)
End Function
"#;

fn unsupported_snapshot(ty: &VarTypeRef) -> bool {
    match ty {
        VarTypeRef::Udt(_) => true,
        VarTypeRef::Array(element) | VarTypeRef::FixedArray { element, .. } => {
            matches!(**element, VarTypeRef::FixedString(_)) || unsupported_snapshot(element)
        }
        _ => false,
    }
}
fn ident(name: &str) -> bool {
    !name.is_empty() && name.chars().all(|c| c.is_ascii_alphanumeric() || c == '_')
}
fn collect_exits(node: SyntaxNode<'_>, offsets: &mut Vec<usize>) {
    if node.kind() == SyntaxKind::ExitStmt
        && (node.text().to_ascii_lowercase().contains("sub")
            || node.text().to_ascii_lowercase().contains("function"))
    {
        offsets.push(node.text_range().0 as usize);
    }
    for child in node.child_nodes() {
        collect_exits(child, offsets);
    }
}
fn instrument(source: &str, procedure: &CoreProc, globals: &[String]) -> Result<String, String> {
    let parsed = calcvba_syntax::parse(source);
    let root = parsed.syntax();
    let node = root
        .child_nodes()
        .into_iter()
        .find(|node| {
            node.proc_name_token()
                .is_some_and(|token| token.text.eq_ignore_ascii_case(&procedure.name))
        })
        .ok_or("procedure source unavailable")?;
    let block = node.body_block().ok_or("procedure body unavailable")?;
    let mut names = Vec::new();
    for local in &procedure.locals {
        if unsupported_snapshot(&local.ty) {
            return Err(
                "record/fixed-string-array snapshot is not a legal Office Variant argument".into(),
            );
        }
        if ident(&local.name) {
            names.push(local.name.clone());
        }
    }
    let snapshots = names
        .iter()
        .map(|name| format!("CvbCensusValue({name})"))
        .collect::<Vec<_>>();
    let capture = format!(
        "CvbCensusError = Err.Number : CvbCensusCaptured = {}",
        if snapshots.is_empty() {
            "\"\"".into()
        } else {
            snapshots.join(" & \"|\" & ")
        }
    );
    let mut edits = vec![(block.text_range().1 as usize, format!("\n{capture}\n"))];
    let mut exits = Vec::new();
    collect_exits(block, &mut exits);
    edits.extend(
        exits
            .into_iter()
            .map(|offset| (offset, format!("{capture} : "))),
    );
    let first = root
        .child_nodes()
        .into_iter()
        .find(|n| n.proc_name_token().is_some())
        .ok_or("no procedures")?
        .text_range()
        .0 as usize;
    edits.push((
        first,
        format!("Private CvbCensusCaptured As String\nPrivate CvbCensusError As Long\nPublic {OBSERVED} As String\n"),
    ));
    edits.sort_by_key(|edit| std::cmp::Reverse(edit.0));
    let mut output = source.to_string();
    for (offset, text) in edits {
        output.insert_str(offset, &text);
    }
    let call = if procedure.return_local.is_some() {
        let set = if procedure
            .return_local
            .and_then(|id| {
                procedure
                    .locals
                    .get(id.0.saturating_sub(procedure.params.len()))
            })
            .is_some_and(|local| matches!(local.ty, VarTypeRef::Object(_)))
        {
            "Set "
        } else {
            ""
        };
        format!("{set}discarded = {}()", procedure.name)
    } else {
        format!("Call {}()", procedure.name)
    };
    let global_snapshot = globals
        .iter()
        .map(|name| format!(" & \"|g:{name}=\" & CvbCensusValue({name})"))
        .collect::<String>();
    output.push_str(&format!(
        r#"
Public Function {PROBE}() As String
    Dim discarded As Variant
    On Error GoTo failed
    {call}
    {OBSERVED} = "ok:error=" & CStr(CvbCensusError) & ":" & CvbCensusCaptured & "|return=" & CvbCensusValue(discarded){global_snapshot}
    {PROBE} = {OBSERVED}
    Exit Function
failed:
    {OBSERVED} = "error:" & CStr(Err.Number)
    {PROBE} = {OBSERVED}
End Function
{VALUE_HELPER}
"#
    ));
    Ok(output)
}

fn manifest(source: String) -> SymbolProjectManifest {
    SymbolProjectManifest {
        project_name: "Census".into(),
        project_kind: ProjectKind::Source,
        modules: vec![ModuleUnit {
            module_name: "Code".into(),
            module_kind: ModuleKind::Procedural,
            attributes: ModuleAttributes::named("Code"),
            source,
        }],
        references: vec![],
        reference_projects: vec![],
        conditional_constants: Default::default(),
        conditional_compilation_target: Default::default(),
    }
}

fn pure_builtin(id: calcvba_bundle::NativeImplId) -> bool {
    use calcvba_bundle::NativeImplId::*;
    matches!(
        id,
        Len | LenB
            | Left
            | LeftB
            | Right
            | RightB
            | Mid
            | MidStmt
            | LSetStmt
            | RSetStmt
            | InStr
            | InStrRev
            | LCase
            | UCase
            | Split
            | Join
            | Replace
            | Trim
            | LTrim
            | RTrim
            | StrComp
            | Like
            | Chr
            | Asc
            | ChrW
            | AscW
            | Space
            | StringRepeat
            | StrReverse
            | StrConv
            | Format
            | FormatNumber
            | FormatCurrency
            | FormatPercent
            | FormatDateTime
            | Filter
            | Abs
            | Int
            | Fix
            | Sgn
            | Round
            | Sqr
            | Sin
            | Cos
            | Log
            | Exp
            | Atn
            | Tan
            | DateSerial
            | TimeSerial
            | DateValue
            | TimeValue
            | DateAdd
            | DateDiff
            | Year
            | Month
            | Day
            | Weekday
            | Hour
            | Minute
            | Second
            | MonthName
            | WeekdayName
            | DatePart
            | Hex
            | Oct
            | CStr
            | Str
            | Val
            | CDate
            | CVErr
            | CBool
            | CByte
            | CInt
            | CLng
            | CLngLng
            | CLngPtr
            | CSng
            | CDbl
            | CCur
            | CDec
            | CVar
            | Fv
            | Pv
            | Pmt
            | Npv
            | Irr
            | Mirr
            | Rate
            | NPer
            | IPmt
            | PPmt
            | Sln
            | Syd
            | Ddb
            | IsArray
            | VarType
            | TypeName
            | IsNumeric
            | IsError
            | IsDate
            | IsObject
            | IsNull
            | IsEmpty
            | IsMissing
            | ErrorText
            | IIf
            | Choose
            | Switch
            | Rgb
            | QbColor
    )
}

#[cfg(windows)]
fn native_value(path: &Path) -> Result<String, String> {
    use windows_sys::Win32::{
        Foundation::{FreeLibrary, SysFreeString, SysStringLen},
        System::LibraryLoader::{GetProcAddress, LoadLibraryW},
    };
    let path: Vec<u16> = path
        .to_string_lossy()
        .encode_utf16()
        .chain(Some(0))
        .collect();
    // SAFETY: this child owns the generated DLL and fixed BSTR/no-argument ABI;
    // all calls, cleanup and unload take place on its one owning main thread.
    unsafe {
        let library = LoadLibraryW(path.as_ptr());
        if library.is_null() {
            return Err("LoadLibrary failed".into());
        }
        let call: unsafe extern "C" fn() -> *mut u16 = std::mem::transmute(
            GetProcAddress(library, c"CensusProbe".as_ptr().cast()).ok_or("missing probe")?,
        );
        let shutdown: unsafe extern "C" fn() -> i32 = std::mem::transmute(
            GetProcAddress(library, c"CalcVbaShutdown".as_ptr().cast())
                .ok_or("missing shutdown")?,
        );
        let error: unsafe extern "C" fn(*mut usize, *mut usize) -> i32 = std::mem::transmute(
            GetProcAddress(library, c"CalcVbaGetLastError".as_ptr().cast())
                .ok_or("missing error accessor")?,
        );
        let pointer = call();
        let number = error(std::ptr::null_mut(), std::ptr::null_mut());
        let value = if pointer.is_null() {
            String::new()
        } else {
            String::from_utf16_lossy(std::slice::from_raw_parts(
                pointer,
                SysStringLen(pointer) as usize,
            ))
        };
        SysFreeString(pointer);
        let stopped = shutdown();
        FreeLibrary(library);
        if number != 0 || stopped != 0 {
            Err(format!(
                "native boundary error {number}, shutdown {stopped}"
            ))
        } else {
            Ok(value)
        }
    }
}
#[cfg(not(windows))]
fn native_value(_: &Path) -> Result<String, String> {
    Err("Windows AOT execution required".into())
}

fn worker(path: &Path, directory: &Path) -> Result<Vec<Row>, Box<dyn std::error::Error>> {
    std::fs::create_dir_all(directory)?;
    if path
        .file_stem()
        .unwrap()
        .to_string_lossy()
        .ends_with(".class")
    {
        return Ok(vec![Row::skipped(
            path,
            "",
            "class module needs its source project",
        )]);
    }
    let source = std::fs::read_to_string(path)?;
    let core = match calcvba_bind::bind_program(
        &manifest(source.clone()),
        &calcvba_symbol::CatalogTypeLibResolver,
    ) {
        Ok(core) => core,
        Err(error) => {
            return Ok(vec![Row::skipped(
                path,
                "",
                format!("source diagnostic: {error}"),
            )]);
        }
    };
    let program = calcvba_oxir::elaborate::elaborate(&core).map_err(|e| e.to_string())?;
    let roots: Vec<_> = (0..program.funcs.len()).map(calcvba_oxir::FuncId).collect();
    let refusal = calcvba_build::check_native_program(&program, &roots)
        .err()
        .map(|e| e.to_string());
    let effect = if !program.external_calls.is_empty() {
        Some("Declare requires a separately controlled native fixture".into())
    } else {
        program
            .funcs
            .iter()
            .flat_map(|f| &f.blocks)
            .flat_map(|b| &b.instrs)
            .find_map(|inst| match inst {
                OxInst::CallExtern { import, .. } => {
                    match calcvba_rt_abi::library::resolve_import(&program, *import) {
                        Ok(resolved) if !pure_builtin(resolved.implementation) => Some(format!(
                            "host-dependent or nondeterministic builtin {:?}",
                            resolved.implementation
                        )),
                        Err(_) => Some("library import requires its controlled fixture".into()),
                        _ => None,
                    }
                }
                OxInst::CallNative {
                    callee: OxNativeCallee::Builtin(id),
                    ..
                } if !pure_builtin(*id) => {
                    Some(format!("host-dependent or nondeterministic builtin {id:?}"))
                }
                _ => None,
            })
    };
    let globals: Vec<_> = core
        .globals
        .iter()
        .filter(|g| ident(&g.name) && !unsupported_snapshot(&g.ty))
        .map(|g| g.name.clone())
        .collect();
    let mut rows = Vec::new();
    for (index, procedure) in core.procs.iter().enumerate() {
        let mut row = Row::skipped(path, &procedure.name, "");
        if let Some(reason) = refusal.as_ref().or(effect.as_ref()) {
            row.detail = reason.clone();
            rows.push(row);
            continue;
        }
        if !procedure.params.is_empty() || !ident(&procedure.name) {
            row.detail = "requires arguments or host/project entry context".into();
            rows.push(row);
            continue;
        }
        let instrumented = match instrument(&source, procedure, &globals) {
            Ok(value) => value,
            Err(reason) => {
                row.detail = reason;
                rows.push(row);
                continue;
            }
        };
        let case = directory.join(index.to_string());
        std::fs::create_dir_all(&case)?;
        std::fs::write(case.join("Code.bas"), instrumented)?;
        let project = case.join("Census.basproj");
        std::fs::write(
            &project,
            format!(
                r#"<Project><PropertyGroup><OutputType>NativeLibrary</OutputType><ProjectName>Census</ProjectName></PropertyGroup><ItemGroup><Module Include="Code.bas"/><NativeExport Include="CensusProbe"><Module>Code</Module><Procedure>{PROBE}</Procedure></NativeExport></ItemGroup></Project>"#
            ),
        )?;
        let result = (|| -> Result<(), String> {
            let image_path = build_oximage(&project, &case).map_err(|e| e.to_string())?;
            let mut image =
                OxImage::from_bytes(&std::fs::read(image_path).map_err(|e| e.to_string())?)
                    .map_err(|e| e.to_string())?;
            let program = &mut image.programs[image.entry];
            program.entry = Some(calcvba_oxir::FuncId(
                program
                    .funcs
                    .iter()
                    .position(|f| f.name.eq_ignore_ascii_case(PROBE))
                    .ok_or("probe entry absent")?,
            ));
            let observed = program
                .globals
                .iter()
                .position(|g| g.name.eq_ignore_ascii_case(OBSERVED))
                .ok_or("observed global absent")?;
            let run = |config| {
                Engine::new(config)
                    .execute_image_with_variant_snapshot(&image)
                    .map_err(|e| e.to_string())
                    .and_then(|values| {
                        values
                            .get(observed)
                            .and_then(|v| v.as_bstr())
                            .map(|s| s.to_string())
                            .ok_or_else(|| "probe did not return a String".into())
                    })
            };
            row.vm3 = Some(run(HostConfig::vm3())?);
            row.jit = Some(run(HostConfig::jit())?);
            let built = build_native_library(&NativeLibraryBuildOptions {
                project_path: project,
                out_dir: case.join("native"),
                template_path: None,
            })
            .map_err(|e| e.to_string())?;
            row.aot = Some(native_value(&built.dll_path)?);
            row.status = if row.vm3 == row.jit && row.vm3 == row.aot {
                "agreed"
            } else {
                "disagreed"
            }
            .into();
            Ok(())
        })();
        if let Err(error) = result {
            row.detail = error;
        }
        rows.push(row);
    }
    Ok(rows)
}
fn collect(root: &Path, paths: &mut Vec<PathBuf>) -> std::io::Result<()> {
    for entry in std::fs::read_dir(root)? {
        let path = entry?.path();
        if path.is_dir() {
            collect(&path, paths)?;
        } else if path
            .extension()
            .is_some_and(|s| s.eq_ignore_ascii_case("bas"))
        {
            paths.push(path);
        }
    }
    Ok(())
}
struct Running {
    child: Child,
    started: Instant,
    path: PathBuf,
    output: PathBuf,
}
fn main() -> Result<(), Box<dyn std::error::Error>> {
    let args: Vec<_> = std::env::args().skip(1).collect();
    if args.first().is_some_and(|s| s == "--worker") {
        let directory = PathBuf::from(&args[2]);
        let result = worker(Path::new(&args[1]), &directory).unwrap_or_else(|e| {
            vec![Row::skipped(
                Path::new(&args[1]),
                "",
                format!("worker: {e}"),
            )]
        });
        std::fs::write(
            directory.join("result.json"),
            serde_json::to_vec_pretty(&result)?,
        )?;
        return Ok(());
    }
    let root = PathBuf::from(args.first().map(String::as_str).unwrap_or("conformance"));
    let output = PathBuf::from(
        args.get(1)
            .map(String::as_str)
            .unwrap_or("temp/native-execution-census"),
    );
    std::fs::create_dir_all(&output)?;
    let mut paths = Vec::new();
    collect(&root, &mut paths)?;
    paths.sort();
    let mut pending = paths.into_iter().enumerate();
    let mut active: Vec<Running> = Vec::new();
    let mut rows = Vec::new();
    loop {
        while active.len() < 2 {
            let Some((index, path)) = pending.next() else {
                break;
            };
            let directory = output.join(index.to_string());
            std::fs::create_dir_all(&directory)?;
            let child = Command::new(std::env::current_exe()?)
                .arg("--worker")
                .arg(&path)
                .arg(&directory)
                .stdout(std::process::Stdio::null())
                .stderr(std::process::Stdio::null())
                .spawn()?;
            active.push(Running {
                child,
                started: Instant::now(),
                path,
                output: directory.join("result.json"),
            });
        }
        if active.is_empty() {
            break;
        }
        for index in (0..active.len()).rev() {
            let timed_out = active[index].started.elapsed() > Duration::from_secs(45);
            if timed_out || active[index].child.try_wait()?.is_some() {
                let mut running = active.remove(index);
                if timed_out {
                    running.child.kill()?;
                    running.child.wait()?;
                }
                match std::fs::read(&running.output)
                    .ok()
                    .and_then(|bytes| serde_json::from_slice::<Vec<Row>>(&bytes).ok())
                {
                    Some(result) => rows.extend(result),
                    None => rows.push(Row::skipped(
                        &running.path,
                        "",
                        if timed_out {
                            "owned worker timed out"
                        } else {
                            "owned worker crashed"
                        },
                    )),
                }
                std::fs::write(
                    output.join("results.json"),
                    serde_json::to_vec_pretty(&rows)?,
                )?;
            }
        }
        std::thread::sleep(Duration::from_millis(50));
    }
    rows.sort_by(|a, b| (&a.file, &a.procedure).cmp(&(&b.file, &b.procedure)));
    std::fs::write(
        output.join("results.json"),
        serde_json::to_vec_pretty(&rows)?,
    )?;
    for status in ["agreed", "disagreed", "not-run"] {
        println!(
            "{status}: {}",
            rows.iter().filter(|r| r.status == status).count()
        );
    }
    println!(
        "{} procedure/module outcomes; bounded workers; deterministic zero-argument entries; snapshots include locals, scalar globals and return/error values",
        rows.len()
    );
    Ok(())
}
