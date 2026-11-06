//! Source transformation tool. Excel's ABI is represented by ordinary VBA.
use calcvba_build::NativeLibraryBuildOptions;
use calcvba_bundle::{
    ProcedureKind, VarTypeRef,
    coreir::{CoreConst, CoreParam, CoreProc},
};
use calcvba_project::ModuleKind;
use std::{error::Error, fmt::Write as _, fs, path::PathBuf};

type Result<T> = std::result::Result<T, Box<dyn Error>>;
const PLAIN: &str = include_str!("../../../examples/native/xll/PlainXll.bas");
const VALUES: &str = include_str!("../../../examples/native/xll/XlValues.bas");
const APPLICATION: &str = include_str!("../vba/Application.bas");

fn main() {
    if let Err(error) = run() {
        eprintln!("calcvba-udf: {error}");
        std::process::exit(1);
    }
}
fn run() -> Result<()> {
    let mut args = std::env::args().skip(1);
    let mut project = None;
    let mut output = None;
    let mut selected = Vec::new();
    let mut prefix = String::new();
    let mut slots = 16usize;
    let mut emit_only = false;
    while let Some(arg) = args.next() {
        match arg.as_str() {
            "--project" => project = args.next().map(PathBuf::from),
            "--out" => output = args.next().map(PathBuf::from),
            "--udf" => selected.push(args.next().ok_or("--udf needs Module.Function")?),
            "--prefix" => prefix = args.next().ok_or("--prefix needs text")?,
            "--paramarray-slots" => {
                slots = args
                    .next()
                    .ok_or("--paramarray-slots needs a number")?
                    .parse()?
            }
            "--emit-only" => emit_only = true,
            "--help" | "-h" => {
                println!(
                    "calcvba-udf --project input.basproj --out NEW_DIRECTORY [--udf Module.Function ...] [--prefix Native_] [--paramarray-slots 16] [--emit-only]\nEmits bound-signature VBA shims, registration.tsv, NativeExport entries and a native .xll. Default: all public standard-module functions. ParamArray capacity is explicit, maximum 32 total parameters. Refuses document modules, forms, events and project references."
                );
                return Ok(());
            }
            _ => return Err(format!("unknown argument {arg}; use --help").into()),
        }
    }
    let project = project.ok_or("--project is required")?.canonicalize()?;
    let output = output.ok_or("--out is required")?;
    if output.exists() {
        return Err("output directory already exists; choose a new directory".into());
    }
    if !(1..=32).contains(&slots) {
        return Err("ParamArray slots must be between 1 and 32".into());
    }
    let parsed = calcvba_project::parse_basproj_xml(&fs::read_to_string(&project)?)?;
    if !parsed.project_references.is_empty() {
        return Err("native project references are not supported".into());
    }
    for module in &parsed.modules {
        if matches!(
            module.kind,
            calcvba_project::BasProjModuleKind::DocumentModule
                | calcvba_project::BasProjModuleKind::FormModule
        ) {
            return Err(format!(
                "{}: document modules and forms cannot be wrapped",
                module.include
            )
            .into());
        }
    }
    let loaded = calcvba_project::load_basproj(&project)?;
    for module in &loaded.manifest.modules {
        if matches!(
            module.module_kind,
            ModuleKind::Document | ModuleKind::Form | ModuleKind::Extension
        ) {
            return Err(format!(
                "{}: document modules, forms and extensions cannot be wrapped",
                module.module_name
            )
            .into());
        }
        if ["UdfRuntime", "XlValues"]
            .iter()
            .any(|name| module.module_name.eq_ignore_ascii_case(name))
        {
            return Err(format!("reserved generated module name {}", module.module_name).into());
        }
        if calcvba_syntax::lexer::tokenize(&module.source)
            .iter()
            .any(|(kind, _)| {
                matches!(
                    kind,
                    calcvba_syntax::SyntaxKind::KwEvent
                        | calcvba_syntax::SyntaxKind::KwWithEvents
                        | calcvba_syntax::SyntaxKind::KwRaiseEvent
                )
            })
        {
            return Err(format!(
                "{}: events and WithEvents cannot be wrapped",
                module.module_name
            )
            .into());
        }
    }
    let mut closure = calcvba_project::load_project_closure(&project)?;
    if closure.len() != 1 {
        return Err("native project references are not supported".into());
    }
    let manifest = closure.last_mut().ok_or("empty project")?;
    if manifest.references.iter().any(|r| {
        matches!(
            r,
            calcvba_symbol::manifest::ProjectReference::Project { .. }
                | calcvba_symbol::manifest::ProjectReference::HostInjected { .. }
        )
    }) {
        return Err("native project references are not supported".into());
    }
    // Bind the real bodies against the very same ordinary VBA Application property
    // that will be compiled in the output project. No activation of Excel occurs.
    manifest.modules.push(calcvba_symbol::manifest::ModuleUnit {
        module_name: "UdfRuntime".into(),
        module_kind: calcvba_symbol::manifest::ModuleKind::Procedural,
        attributes: calcvba_symbol::manifest::ModuleAttributes::named("UdfRuntime"),
        source: "Public Property Get Application() As Object\nEnd Property\n".into(),
    });
    // The property uses only Object, so signature binding needs no host type.
    let (core, procedures) = calcvba_bind::bind_program_with_procedures(
        manifest,
        &calcvba_symbol::CatalogTypeLibResolver,
    )?;
    let mut functions = Vec::new();
    for entry in procedures {
        if entry.module == "UdfRuntime" {
            continue;
        }
        let qualified = format!("{}.{}", entry.module, entry.name);
        let wanted =
            selected.is_empty() || selected.iter().any(|s| s.eq_ignore_ascii_case(&qualified));
        if !wanted {
            continue;
        }
        let proc = &core.procs[entry.index];
        if entry.visibility != calcvba_symbol::model::Visibility::Public
            || entry.option_private_module
            || proc.kind != ProcedureKind::Function
        {
            if !selected.is_empty() {
                return Err(format!(
                    "{qualified}: UDF must be a Public Function in a public standard module"
                )
                .into());
            }
            continue;
        }
        validate_signature(proc, &core)?;
        let worksheet = format!("{prefix}{}", entry.name);
        if functions
            .iter()
            .any(|(_, _, name): &(String, &CoreProc, String)| name.eq_ignore_ascii_case(&worksheet))
        {
            return Err(format!(
                "duplicate worksheet name {worksheet}; select qualified UDFs explicitly"
            )
            .into());
        }
        functions.push((qualified, proc, worksheet));
    }
    for name in &selected {
        if !functions
            .iter()
            .any(|(qualified, _, _)| name.eq_ignore_ascii_case(qualified))
        {
            return Err(format!("UDF {name} was not found").into());
        }
    }
    if functions.is_empty() {
        return Err("no public UDFs selected".into());
    }
    let mut runtime = PLAIN[..PLAIN
        .find("Public Function AddDouble")
        .ok_or("invalid runtime template")?]
        .replace("PlainXll", "UdfRuntime");
    runtime = runtime
        .replace("Private Function ExcelCall", "Public Function ExcelCall")
        .replace("Private Sub ExcelFree", "Public Sub ExcelFree");
    runtime = runtime.replace("    argsPointer = VarPtr(args)", "    If callback = 0^ Then Err.Raise 5&, \"UdfRuntime.ExcelCall\", \"Excel callback is unavailable\"\n    argsPointer = VarPtr(args)");
    runtime = runtime
        .replace("1 To 7", &format!("1 To {}", functions.len()))
        .replace("To 7&", &format!("To {}&", functions.len()));
    let start = runtime
        .find("Private Function RegistrationTable")
        .ok_or("registration table missing")?;
    let end = start
        + runtime[start..]
            .find("End Function")
            .ok_or("registration table end missing")?
        + "End Function".len();
    let mut table = String::from(
        "Private Function RegistrationTable(ByVal index As Long) As RegistrationRow\nSelect Case index\n",
    );
    let mut wrappers = String::new();
    let mut registration = String::from("export\tworksheet\ttypes\targuments\tsource\n");
    let mut exports = String::new();
    for (i, (qualified, proc, worksheet)) in functions.iter().enumerate() {
        let export = format!("Udf{}", i + 1);
        let count = if proc.params.last().is_some_and(|p| p.variadic) {
            proc.params.len() - 1 + slots
        } else {
            proc.params.len()
        };
        if count > 32 {
            return Err(
                format!("{qualified}: {count} arguments exceeds the native limit of 32").into(),
            );
        }
        let types = format!(
            "Q{}",
            (0..count)
                .map(|p| {
                    if proc.params.get(p).is_some_and(|p| {
                        !p.variadic && !matches!(p.ty, VarTypeRef::Variant | VarTypeRef::Object(_))
                    }) {
                        'Q'
                    } else {
                        'U'
                    }
                })
                .collect::<String>()
        );
        let names = (0..count)
            .map(|p| {
                proc.params
                    .get(p)
                    .filter(|p| !p.variadic)
                    .map(|p| p.name.clone())
                    .unwrap_or_else(|| format!("item{}", p + 1))
            })
            .collect::<Vec<_>>()
            .join(",");
        writeln!(
            table,
            "Case {}\nRegistrationTable.exportName={}\nRegistrationTable.typeText={}\nRegistrationTable.worksheetName={}\nRegistrationTable.argumentNames={}",
            i + 1,
            quote(&export),
            quote(&types),
            quote(worksheet),
            quote(&names)
        )?;
        writeln!(
            registration,
            "{export}\t{worksheet}\t{types}\t{names}\t{qualified}"
        )?;
        wrappers.push_str(&wrapper(qualified, proc, &export, count)?);
        export_line(&mut exports, &export, "UdfRuntime", &export)?;
    }
    table.push_str("Case Else\nErr.Raise 9&\nEnd Select\nEnd Function");
    runtime.replace_range(start..end, &table);
    runtime.push_str(
        "\nPublic Sub xlAutoFree12(ByVal operand As LongPtr)\nXlReleaseReturn operand\nEnd Sub\n",
    );
    runtime.push_str(
        &PLAIN[PLAIN
            .find("Public Function xlAddInManagerInfo12")
            .ok_or("add-in information template missing")?..],
    );
    runtime = runtime.replace(
        "    xlAutoClose = 1&",
        "    Set cachedApplication = Nothing\n    xlAutoClose = 1&",
    );
    let (declarations, application) = APPLICATION
        .split_once("'--- procedures ---")
        .ok_or("Application template missing separator")?;
    runtime = runtime.replacen(
        "Option Explicit",
        &declarations.replace("Attribute VB_Name = \"UdfRuntime\"", ""),
        1,
    );
    runtime.push_str(application);
    runtime.push_str(&wrappers);
    let mut xml = format!(
        "<Project><PropertyGroup><OutputType>NativeLibrary</OutputType><ProjectName>{}</ProjectName>",
        xml_escape(&format!("{}Udf", loaded.manifest.project_name))
    );
    if !manifest.conditional_constants.is_empty() {
        let constants = manifest
            .conditional_constants
            .iter()
            .map(|(key, value)| format!("{key}={value}"))
            .collect::<Vec<_>>()
            .join(";");
        write!(
            xml,
            "<DefineConstants>{}</DefineConstants>",
            xml_escape(&constants)
        )?;
    }
    xml.push_str("</PropertyGroup><ItemGroup>\n");
    // All validation above precedes the first output write.
    fs::create_dir_all(&output)?;
    for module in &loaded.manifest.modules {
        let ext = if module.module_kind == ModuleKind::Class {
            "cls"
        } else {
            "bas"
        };
        let name = format!("{}.{ext}", module.module_name);
        fs::write(output.join(&name), &module.source)?;
        let tag = if ext == "cls" {
            "ClassModule"
        } else {
            "Module"
        };
        writeln!(xml, "<{tag} Include=\"{}\"/>", xml_escape(&name))?;
    }
    for reference in &manifest.references {
        use calcvba_symbol::manifest::ProjectReference;
        let ProjectReference::TypeLibrary {
            name,
            guid,
            version_major,
            version_minor,
            lcid,
            import_lib,
        } = reference
        else {
            if let ProjectReference::Native { name, path } = reference {
                write!(xml, "<NativeReference Include=\"{}\">", xml_escape(name))?;
                if let Some(path) = path {
                    let path = project.parent().unwrap().join(path);
                    write!(
                        xml,
                        "<Path>{}</Path>",
                        xml_escape(&path.display().to_string())
                    )?;
                }
                xml.push_str("</NativeReference>\n");
            }
            continue;
        };
        write!(xml, "<COMReference Include=\"{}\">", xml_escape(name))?;
        for (tag, value) in [
            ("Guid", guid.clone()),
            ("VersionMajor", version_major.map(|v| v.to_string())),
            ("VersionMinor", version_minor.map(|v| v.to_string())),
            ("Lcid", lcid.map(|v| v.to_string())),
            (
                "ImportLib",
                import_lib.as_ref().map(|p| {
                    let p = PathBuf::from(p);
                    let relative = project.parent().unwrap().join(&p);
                    if p.is_absolute() || !relative.exists() {
                        p
                    } else {
                        relative
                    }
                    .display()
                    .to_string()
                }),
            ),
        ] {
            if let Some(value) = value {
                write!(xml, "<{tag}>{}</{tag}>", xml_escape(&value))?;
            }
        }
        xml.push_str("</COMReference>\n");
    }
    for name in [
        "xlAutoOpen",
        "xlAutoClose",
        "xlAutoFree12",
        "xlAddInManagerInfo12",
    ] {
        export_line(&mut exports, name, "UdfRuntime", name)?;
    }
    xml.push_str("<Module Include=\"UdfRuntime.bas\"/><Module Include=\"XlValues.bas\"/>\n");
    xml.push_str(&exports);
    xml.push_str("</ItemGroup></Project>\n");
    fs::write(output.join("UdfRuntime.bas"), runtime)?;
    fs::write(output.join("XlValues.bas"), VALUES)?;
    fs::write(output.join("registration.tsv"), registration)?;
    let generated = output.join("Udf.basproj");
    fs::write(&generated, xml)?;
    if !emit_only {
        let built = calcvba_build::build_native_library(&NativeLibraryBuildOptions {
            project_path: generated.clone(),
            out_dir: output.join("native"),
            template_path: None,
        })?;
        let xll = output.join(format!("{}Udf.xll", loaded.manifest.project_name));
        fs::copy(&built.dll_path, &xll)?;
        println!("{}", xll.display());
    } else {
        println!("{}", generated.display());
    }
    Ok(())
}
fn quote(s: &str) -> String {
    format!("\"{}\"", s.replace('"', "\"\""))
}
fn xml_escape(s: &str) -> String {
    s.replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
}
fn export_line(out: &mut String, name: &str, module: &str, procedure: &str) -> std::fmt::Result {
    writeln!(
        out,
        "<NativeExport Include=\"{}\"><Module>{module}</Module><Procedure>{procedure}</Procedure></NativeExport>",
        xml_escape(name)
    )
}
fn type_name(ty: &VarTypeRef) -> Result<String> {
    Ok(match ty {
        VarTypeRef::Builtin(t) => format!("{t:?}"),
        VarTypeRef::Variant => "Variant".into(),
        VarTypeRef::Object(n)
            if n.rsplit('.').next().is_some_and(|s| {
                s.eq_ignore_ascii_case("Range") || s.eq_ignore_ascii_case("_Range")
            }) =>
        {
            n.clone()
        }
        VarTypeRef::Array(e) if **e == VarTypeRef::Variant => "Variant()".into(),
        _ => return Err(format!("unsupported worksheet signature type {ty:?}").into()),
    })
}
fn validate_signature(proc: &CoreProc, core: &calcvba_bundle::coreir::CoreProgram) -> Result<()> {
    let local = proc.return_local.ok_or("UDF has no result")?;
    for ty in proc
        .params
        .iter()
        .map(|p| &p.ty)
        .chain(Some(&proc.locals[local.0 - proc.params.len()].ty))
    {
        type_name(ty)?;
        if let VarTypeRef::Object(name) = ty {
            let project_class = core.classes.iter().any(|class| {
                class.name.eq_ignore_ascii_case(name)
                    || format!("{}.{}", core.unit_name, class.name).eq_ignore_ascii_case(name)
            });
            let excel_range = core.object_interfaces.iter().any(|interface| {
                interface.name.eq_ignore_ascii_case(name)
                    && interface.iid.is_some_and(|iid| {
                        iid.data1 == 0x00020846
                            && iid.data2 == 0
                            && iid.data3 == 0
                            && iid.data4 == [0xc0, 0, 0, 0, 0, 0, 0, 0x46]
                    })
            });
            if project_class || !excel_range {
                return Err(format!(
                    "{name}: only the bound Excel Range interface can receive worksheet references"
                )
                .into());
            }
        }
    }
    Ok(())
}
fn literal(value: &CoreConst) -> String {
    match value {
        CoreConst::Empty => "Empty".into(),
        CoreConst::Null => "Null".into(),
        CoreConst::Nothing => "Nothing".into(),
        CoreConst::Bool(v) => if *v { "True" } else { "False" }.into(),
        CoreConst::I16(v) => format!("{v}%"),
        CoreConst::I32(v) => format!("{v}&"),
        CoreConst::I64(v) => format!("{v}^"),
        CoreConst::F64(v) => format!("{}#", f64::from_bits(*v)),
        CoreConst::F32(v) => format!("{}!", f32::from_bits(*v)),
        CoreConst::Currency(v) => format!("UdfCurrency({v}^)"),
        CoreConst::Date(v) => format!("CDate({}#)", f64::from_bits(*v)),
        CoreConst::Str(s) => {
            let mut parts = Vec::new();
            let mut plain = String::new();
            for c in s.chars() {
                if c.is_control() {
                    if !plain.is_empty() {
                        parts.push(quote(&plain));
                        plain.clear();
                    }
                    parts.push(format!("ChrW$({}&)", u32::from(c)));
                } else {
                    plain.push(c);
                }
            }
            if !plain.is_empty() || parts.is_empty() {
                parts.push(quote(&plain));
            }
            parts.join(" & ")
        }
    }
}
fn wrapper(qualified: &str, proc: &CoreProc, export: &str, count: usize) -> Result<String> {
    let params = (0..count)
        .map(|i| format!("ByVal p{i} As LongPtr"))
        .collect::<Vec<_>>()
        .join(", ");
    let mut out = format!(
        "\nPublic Function {export}({params}) As LongPtr\nOn Error GoTo Failed\nDim result As Variant, oper As XLOPER12\n"
    );
    let fixed = proc.params.iter().filter(|p| !p.variadic).count();
    for i in 0..count {
        let p = proc.params.get(i).filter(|p| !p.variadic);
        let ty = p
            .map(|p| type_name(&p.ty))
            .transpose()?
            .unwrap_or_else(|| "Variant".into());
        if ty == "Variant()" {
            writeln!(out, "Dim a{i}() As Variant")?;
        } else {
            writeln!(out, "Dim a{i} As {ty}")?;
        }
        if let Some(p) = p.filter(|p| p.optional) {
            writeln!(out, "If UdfMissing(p{i}) Then")?;
            if let Some(default) = &p.default {
                let set = if matches!(default, CoreConst::Nothing) {
                    "Set "
                } else {
                    ""
                };
                writeln!(out, "{set}a{i}={}", literal(default))?;
            } else if p.ty == VarTypeRef::Variant {
                writeln!(out, "a{i}=UdfValue(p{i})")?;
            }
            out.push_str("Else\n");
            argument(&mut out, i, Some(p))?;
            out.push_str("End If\n");
        } else {
            argument(&mut out, i, p)?;
        }
    }
    let result_ty = &proc.locals[proc.return_local.unwrap().0 - proc.params.len()].ty;
    let object_result = matches!(result_ty, VarTypeRef::Object(_));
    if object_result {
        out.push_str("Dim returned As Object\n");
    }
    let call = |names: &str| {
        if object_result {
            format!("Set returned={qualified}({names})\nresult=returned.Value2")
        } else {
            format!("result={qualified}({names})")
        }
    };
    if proc.params.last().is_some_and(|p| p.variadic) {
        writeln!(
            out,
            "Dim tail As Long\ntail={}\nDo While tail>0&",
            count - fixed
        )?;
        out.push_str("Select Case tail\n");
        for n in 1..=count - fixed {
            writeln!(
                out,
                "Case {n}\nIf Not UdfMissing(p{}) Then Exit Do",
                fixed + n - 1
            )?;
        }
        out.push_str("End Select\ntail=tail-1&\nLoop\nSelect Case tail\n");
        for n in 0..=count - fixed {
            let names = (0..fixed + n)
                .map(|i| format!("a{i}"))
                .collect::<Vec<_>>()
                .join(",");
            writeln!(out, "Case {n}\n{}", call(&names))?;
        }
        out.push_str("End Select\n");
    } else {
        let names = (0..fixed)
            .map(|i| format!("a{i}"))
            .collect::<Vec<_>>()
            .join(",");
        writeln!(out, "{}", call(&names))?;
    }
    writeln!(
        out,
        "oper=XlFromVariant(result)\n{export}=XlReturn(oper)\nExit Function\nFailed:\nErr.Clear\noper=XlError(15&)\n{export}=XlReturn(oper)\nEnd Function"
    )?;
    Ok(out)
}
fn argument(out: &mut String, i: usize, p: Option<&CoreParam>) -> Result<()> {
    match p.map(|p| &p.ty) {
        Some(VarTypeRef::Object(_)) => writeln!(out, "Set a{i}=UdfRange(p{i})")?,
        Some(VarTypeRef::Variant) | None => writeln!(
            out,
            "If UdfIsReference(p{i}) Then\nSet a{i}=UdfRange(p{i})\nElse\na{i}=UdfValue(p{i})\nEnd If"
        )?,
        _ => writeln!(out, "a{i}=UdfValue(p{i})")?,
    }
    Ok(())
}
