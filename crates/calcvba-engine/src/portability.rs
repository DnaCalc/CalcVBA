//! Source-less portability analysis and enforcement for compiled [`OxImage`] / [`OxProgram`] units.
//!
//! Uses the same portability matrix, Win32 `Declare` shim table, portable COM registry,
//! and mode enforcement rules as the source analyzer in `calcvba_symbol::portability`.

mod effects;

use std::collections::HashMap;

use calcvba_bundle::{DeclareParamType, ExportTarget, NativeImplId};
use calcvba_diagnostics::Diagnostic;
use calcvba_oxir::com::ComInterface;
use calcvba_oxir::inst::{OxAsNew, OxInst};
use calcvba_oxir::value::{OxArg, OxCallArg, OxConst, OxNativeCallee, OxOperand};
use calcvba_oxir::{OxImage, OxProgram, OxTerminator};
use calcvba_symbol::portability::{
    BrowserServices, FeatureClass, Finding, Host, Mode, PortabilityEffect, Report, Support,
    com_activation_support, declare_support, enforce_report, looks_like_windows_path,
    lookup_matrix, verdict_for_support,
};

/// Inspect a compiled `.oxi` image against the portability matrix for `host`.
pub fn check_image(image: &OxImage, host: Host) -> Report {
    check_image_with_browser_services(image, host, BrowserServices::default())
}

/// Decode and inspect serialized `.oxi` bytes against the portability matrix for `host`.
pub fn check_image_bytes(bytes: &[u8], host: Host) -> Result<Report, String> {
    let image = OxImage::from_bytes(bytes).map_err(|error| error.to_string())?;
    Ok(check_image(&image, host))
}

/// Inspect a compiled `.oxi` image with explicit browser host services.
pub fn check_image_with_browser_services(
    image: &OxImage,
    host: Host,
    browser: BrowserServices,
) -> Report {
    let programs = image.program_refs();
    check_programs(&programs, host, browser)
}

/// Combine source locations with effects over the compiled closure. A static
/// check never resolves registered type libraries or activates host objects.
pub fn check_source_projects(
    projects: &[calcvba_symbol::manifest::SymbolProjectManifest],
    host: Host,
) -> Result<Report, String> {
    struct SourceTypes;
    impl calcvba_symbol::TypeLibResolver for SourceTypes {
        fn resolve(
            &self,
            _request: &calcvba_com::TypeLibResolveRequest,
        ) -> Option<calcvba_com::TypeLibMetadataBlob> {
            None
        }

        fn dictionary_source(&self) -> calcvba_symbol::DictionarySource {
            calcvba_symbol::DictionarySource::Engine
        }
    }

    let mut report = calcvba_symbol::portability::check_project(projects, host)?;
    if report.analysis_complete
        && let Ok(core) = calcvba_bind::bind_projects(projects, &SourceTypes)
        && let Ok(programs) = core
            .iter()
            .map(calcvba_oxir::elaborate::elaborate)
            .collect::<Result<Vec<_>, _>>()
    {
        let programs: Vec<_> = programs.iter().collect();
        let bound = check_programs(&programs, host, BrowserServices::default());
        report.procedures = bound.procedures;
        report.effects_complete = bound.effects_complete;
    }
    Ok(report)
}

/// Enforce `mode` against a compiled `.oxi` image before linking or execution.
pub fn enforce_image(
    mode: Mode,
    image: &OxImage,
    host: Host,
    browser: BrowserServices,
) -> Result<Vec<Diagnostic>, Box<Diagnostic>> {
    if mode != Mode::Strict {
        return Ok(Vec::new());
    }
    enforce_report(
        mode,
        check_image_with_browser_services(image, host, browser),
    )
}

/// Inspect a slice of elaborated [`OxProgram`]s against the portability matrix.
pub fn check_programs(programs: &[&OxProgram], host: Host, browser: BrowserServices) -> Report {
    let mut report = Report::new(host);
    let library = calcvba_bundle::vba_library_bundle();
    let offsets: Vec<_> = programs
        .iter()
        .scan(0, |next, unit| {
            let offset = *next;
            *next += unit.funcs.len();
            Some(offset)
        })
        .collect();
    let mut graph = Vec::new();

    for (unit, program) in programs.iter().enumerate() {
        let project = if program.unit_name.is_empty() {
            "OxImage"
        } else {
            program.unit_name.as_str()
        };

        // 1. All Declare Lib external calls (including unreachable ones retained in the image).
        let mut declares = HashMap::new();
        for desc in &program.external_calls {
            let params: Vec<(&str, bool)> = desc
                .param_types
                .iter()
                .zip(
                    desc.param_by_ref
                        .iter()
                        .copied()
                        .chain(std::iter::repeat(true)),
                )
                .map(|(ty, by_ref)| (declare_param_name(*ty), by_ref))
                .collect();
            let result = desc.return_type.map(declare_param_name);
            let entry = if desc.alias.is_empty() {
                desc.declared_name.as_str()
            } else {
                desc.alias.as_str()
            };
            let (class, support) = declare_support(
                &desc.library,
                Some(entry),
                &params,
                result,
                desc.ordinal_alias,
                host,
            );
            let detail = format!("{}!{entry}", desc.library);
            declares.insert(
                desc.descriptor_id,
                PortabilityEffect {
                    feature: "Declare".into(),
                    detail: Some(detail.clone()),
                    verdict: support,
                },
            );
            record(
                &mut report,
                project,
                None,
                "Declare",
                class,
                support,
                Some(detail),
                None,
            );
        }

        // 2. Class `As New` COM fields.
        for class in &program.classes {
            for field in &class.as_new_fields {
                if let OxAsNew::ComClass { prog_id } = &field.binding {
                    let support = com_activation_support(false, Some(prog_id), true, host);
                    record(
                        &mut report,
                        project,
                        Some(&class.name),
                        "CreateObject",
                        FeatureClass::Windows,
                        support,
                        Some(prog_id.clone()),
                        None,
                    );
                }
            }
        }

        // 3. Retained early-bound COM interfaces.
        for interface in &program.com_interfaces {
            let ComInterface::Com(meta) = interface else {
                continue;
            };
            let support = if host == Host::Windows {
                Support::Native
            } else {
                Support::Reject
            };
            record(
                &mut report,
                project,
                None,
                "Early-bound reference",
                FeatureClass::Windows,
                support,
                Some(meta.name.clone()),
                None,
            );
        }

        // 4. Instructions across every compiled procedure.
        for func in &program.funcs {
            let first_finding = report.findings.len();
            let module = func
                .name
                .split_once('.')
                .map(|(module, _)| module)
                .unwrap_or(func.name.as_str());
            let mut procedure = effects::Procedure {
                project: project.into(),
                module: module.into(),
                name: func.name.rsplit('.').next().unwrap_or(&func.name).into(),
                entry: func.entry.0,
                blocks: Vec::with_capacity(func.blocks.len()),
                dynamic_control: false,
            };
            for block in &func.blocks {
                let mut footprint = effects::Block::default();
                match &block.terminator {
                    OxTerminator::Jump(target) => footprint.successors.push(target.0),
                    OxTerminator::Branch {
                        cond,
                        then_blk,
                        else_blk,
                    } => match cond {
                        OxOperand::Const(OxConst::Bool(true)) => {
                            footprint.successors.push(then_blk.0)
                        }
                        OxOperand::Const(OxConst::Bool(false)) => {
                            footprint.successors.push(else_blk.0)
                        }
                        _ => footprint.successors.extend([then_blk.0, else_blk.0]),
                    },
                    OxTerminator::GoSub { .. }
                    | OxTerminator::GoSubReturn
                    | OxTerminator::Resume
                    | OxTerminator::ResumeNext
                    | OxTerminator::ResumeLabel(_) => procedure.dynamic_control = true,
                    _ => {}
                }
                footprint
                    .successors
                    .retain(|target| *target < func.blocks.len());
                for inst in &block.instrs {
                    let start = report.findings.len();
                    check_inst(
                        &mut report,
                        project,
                        module,
                        program,
                        library,
                        inst,
                        host,
                        browser,
                    );
                    for finding in &report.findings[start..] {
                        if finding.detail.as_deref() == Some("dynamic or omitted ProgID") {
                            footprint.unknown = true;
                        } else if finding.class != FeatureClass::Pure
                            && finding.feature != "Windows path"
                        {
                            let effect = PortabilityEffect {
                                feature: finding.feature.clone(),
                                detail: finding.detail.clone(),
                                verdict: finding.verdict,
                            };
                            if !matches!(inst, OxInst::AsNew { .. }) {
                                footprint.must.insert(effect.clone());
                            }
                            footprint.may.insert(effect);
                        }
                    }
                    match inst {
                        OxInst::SetErrorHandler(_) => procedure.dynamic_control = true,
                        OxInst::CallProc { proc, .. } if proc.0 < program.funcs.len() => {
                            footprint.calls.push(offsets[unit] + proc.0)
                        }
                        OxInst::CallExtern { import, .. } => {
                            if let Some(import) = program.imports.get(import.0)
                                && !import.unit.eq_ignore_ascii_case("VBA")
                            {
                                if let Some(target) = resolve_import(programs, &offsets, import) {
                                    footprint.calls.push(target);
                                } else {
                                    footprint.unknown = true;
                                }
                            }
                        }
                        OxInst::CallNative {
                            callee: OxNativeCallee::Declare { descriptor_id },
                            ..
                        } => {
                            if let Some(effect) = declares.get(descriptor_id) {
                                footprint.may.insert(effect.clone());
                                footprint.must.insert(effect.clone());
                            }
                        }
                        OxInst::NewObject { class, .. } => {
                            if let Some(init) = program
                                .classes
                                .get(class.0)
                                .and_then(|class| class.initialize)
                                && init.0 < program.funcs.len()
                            {
                                footprint.calls.push(offsets[unit] + init.0);
                            }
                        }
                        OxInst::CallProcRef { .. }
                        | OxInst::CallByName { .. }
                        | OxInst::ComCallEarly { .. }
                        | OxInst::ComCallLate { .. }
                        | OxInst::NewExtern { .. }
                        | OxInst::PredeclaredExtern { .. } => footprint.unknown = true,
                        _ => {}
                    }
                }
                procedure.blocks.push(footprint);
            }
            for finding in &mut report.findings[first_finding..] {
                finding.procedure = Some(procedure.name.clone());
            }
            graph.push(procedure);
        }
    }

    report.procedures = effects::solve(graph);
    report.effects_complete = true;
    report
}

fn resolve_import(
    programs: &[&OxProgram],
    offsets: &[usize],
    import: &calcvba_bundle::BundleImport,
) -> Option<usize> {
    let (unit, program) = programs
        .iter()
        .enumerate()
        .find(|(_, program)| program.unit_name.eq_ignore_ascii_case(&import.unit))?;
    let export = program
        .exports
        .iter()
        .find(|export| export.token.matches(&import.token))?;
    match export.target {
        ExportTarget::Proc(index) if index < program.funcs.len() => Some(offsets[unit] + index),
        _ => None,
    }
}

fn declare_param_name(ty: DeclareParamType) -> &'static str {
    match ty {
        DeclareParamType::Any => "Any",
        DeclareParamType::Byte => "Byte",
        DeclareParamType::Boolean => "Boolean",
        DeclareParamType::Integer => "Integer",
        DeclareParamType::Long => "Long",
        DeclareParamType::LongLong => "LongLong",
        DeclareParamType::LongPtr => "LongPtr",
        DeclareParamType::Single => "Single",
        DeclareParamType::Double => "Double",
        DeclareParamType::Currency => "Currency",
        DeclareParamType::Date => "Date",
        DeclareParamType::String | DeclareParamType::Variant | DeclareParamType::Array => "",
    }
}

#[allow(clippy::too_many_arguments)]
fn record(
    report: &mut Report,
    project: &str,
    module: Option<&str>,
    feature: &str,
    class: FeatureClass,
    support: Support,
    detail: Option<String>,
    mac_dialect: Option<&'static str>,
) {
    report.verdict = report.verdict.max(verdict_for_support(support));
    report.findings.push(Finding {
        project: project.to_owned(),
        module: module.map(str::to_owned),
        procedure: None,
        line: None,
        feature: feature.to_owned(),
        class,
        verdict: support,
        detail,
        mac_dialect,
    });
}

#[allow(clippy::too_many_arguments)]
fn record_matrix(
    report: &mut Report,
    project: &str,
    module: Option<&str>,
    area: &str,
    name: &str,
    label: Option<&str>,
    detail: Option<String>,
    host: Host,
    browser: BrowserServices,
) {
    if let Some((class, support, mac_dialect)) = lookup_matrix(area, name, host, browser) {
        record(
            report,
            project,
            module,
            label.unwrap_or_else(|| name.trim_end_matches("/$")),
            class,
            support,
            detail,
            Some(mac_dialect),
        );
    }
}

fn check_operand_path(
    report: &mut Report,
    project: &str,
    module: &str,
    operand: &OxOperand,
    host: Host,
) {
    if let OxOperand::Const(OxConst::Str(text)) = operand
        && looks_like_windows_path(text)
    {
        let support = if host == Host::Windows {
            Support::Native
        } else {
            Support::Shim
        };
        record(
            report,
            project,
            Some(module),
            "Windows path",
            FeatureClass::Shim,
            support,
            Some(text.clone()),
            None,
        );
    }
}

enum CallArgView<'a> {
    ConstStr(&'a str),
    Omitted,
    Dynamic,
}

fn view_call_args(args: &[OxCallArg]) -> Vec<CallArgView<'_>> {
    args.iter()
        .map(|arg| match arg {
            OxCallArg::Operand(OxOperand::Const(OxConst::Str(text)))
            | OxCallArg::ByValOverride(OxOperand::Const(OxConst::Str(text))) => {
                CallArgView::ConstStr(text.as_str())
            }
            OxCallArg::Operand(OxOperand::Const(OxConst::NullStr))
            | OxCallArg::ByValOverride(OxOperand::Const(OxConst::NullStr)) => {
                CallArgView::ConstStr("")
            }
            OxCallArg::Omitted => CallArgView::Omitted,
            OxCallArg::Named {
                value: OxOperand::Const(OxConst::Str(text)),
                ..
            } => CallArgView::ConstStr(text.as_str()),
            OxCallArg::Named {
                value: OxOperand::Const(OxConst::NullStr),
                ..
            } => CallArgView::ConstStr(""),
            _ => CallArgView::Dynamic,
        })
        .collect()
}

fn view_proc_args(args: &[OxArg]) -> Vec<CallArgView<'_>> {
    args.iter()
        .map(|arg| match arg {
            OxArg::ByVal(OxOperand::Const(OxConst::Str(text))) => {
                CallArgView::ConstStr(text.as_str())
            }
            OxArg::ByVal(OxOperand::Const(OxConst::NullStr)) => CallArgView::ConstStr(""),
            OxArg::Omitted => CallArgView::Omitted,
            _ => CallArgView::Dynamic,
        })
        .collect()
}

#[allow(clippy::too_many_arguments)]
fn check_builtin(
    report: &mut Report,
    project: &str,
    module: &str,
    id: NativeImplId,
    args: &[CallArgView<'_>],
    host: Host,
    browser: BrowserServices,
) {
    use NativeImplId::*;
    match id {
        CreateObject | GetObject => {
            let is_get = id == GetObject;
            let prog_id = match args.get(usize::from(is_get)) {
                Some(CallArgView::ConstStr(text)) => Some(*text),
                _ => None,
            };
            let can_activate = if is_get {
                matches!(args.first(), Some(CallArgView::ConstStr("")))
            } else {
                match args.get(1) {
                    None | Some(CallArgView::Omitted) | Some(CallArgView::ConstStr("")) => true,
                    Some(_) => false,
                }
            };
            let support = com_activation_support(is_get, prog_id, can_activate, host);
            let label = if is_get { "GetObject" } else { "CreateObject" };
            let detail = Some(
                prog_id
                    .map(str::to_owned)
                    .unwrap_or_else(|| "dynamic or omitted ProgID".into()),
            );
            record(
                report,
                project,
                Some(module),
                label,
                FeatureClass::Windows,
                support,
                detail,
                None,
            );
        }
        Shell => record_matrix(
            report,
            project,
            Some(module),
            "Interaction",
            "Shell",
            None,
            None,
            host,
            browser,
        ),
        SendKeys => record_matrix(
            report,
            project,
            Some(module),
            "Interaction",
            "SendKeys",
            None,
            None,
            host,
            browser,
        ),
        AppActivate => record_matrix(
            report,
            project,
            Some(module),
            "Interaction",
            "AppActivate",
            None,
            None,
            host,
            browser,
        ),
        MacScript => record_matrix(
            report,
            project,
            Some(module),
            "Interaction",
            "MacScript",
            None,
            None,
            host,
            browser,
        ),
        Environ => record_matrix(
            report,
            project,
            Some(module),
            "Interaction",
            "Environ/$",
            None,
            None,
            host,
            browser,
        ),
        GetSetting => record_matrix(
            report,
            project,
            Some(module),
            "Interaction",
            "GetSetting",
            None,
            None,
            host,
            browser,
        ),
        GetAllSettings => record_matrix(
            report,
            project,
            Some(module),
            "Interaction",
            "GetAllSettings",
            None,
            None,
            host,
            browser,
        ),
        SaveSetting => record_matrix(
            report,
            project,
            Some(module),
            "Interaction",
            "SaveSetting",
            None,
            None,
            host,
            browser,
        ),
        DeleteSetting => record_matrix(
            report,
            project,
            Some(module),
            "Interaction",
            "DeleteSetting",
            None,
            None,
            host,
            browser,
        ),
        Dir => record_matrix(
            report,
            project,
            Some(module),
            "FileSystem",
            "Dir",
            None,
            None,
            host,
            browser,
        ),
        FileChDir => record_matrix(
            report,
            project,
            Some(module),
            "FileSystem",
            "ChDir",
            None,
            None,
            host,
            browser,
        ),
        FileChDrive => record_matrix(
            report,
            project,
            Some(module),
            "FileSystem",
            "ChDrive",
            None,
            None,
            host,
            browser,
        ),
        FileCurDir => record_matrix(
            report,
            project,
            Some(module),
            "FileSystem",
            "CurDir/$",
            None,
            None,
            host,
            browser,
        ),
        FileCopy => record_matrix(
            report,
            project,
            Some(module),
            "FileSystem",
            "FileCopy",
            None,
            None,
            host,
            browser,
        ),
        FileDateTime => record_matrix(
            report,
            project,
            Some(module),
            "FileSystem",
            "FileDateTime",
            None,
            None,
            host,
            browser,
        ),
        FileLen => record_matrix(
            report,
            project,
            Some(module),
            "FileSystem",
            "FileLen",
            None,
            None,
            host,
            browser,
        ),
        FileGetAttr => record_matrix(
            report,
            project,
            Some(module),
            "FileSystem",
            "GetAttr",
            None,
            None,
            host,
            browser,
        ),
        FileSetAttr => record_matrix(
            report,
            project,
            Some(module),
            "FileSystem",
            "SetAttr",
            None,
            None,
            host,
            browser,
        ),
        FileKill => record_matrix(
            report,
            project,
            Some(module),
            "FileSystem",
            "Kill",
            None,
            None,
            host,
            browser,
        ),
        FileMkDir => record_matrix(
            report,
            project,
            Some(module),
            "FileSystem",
            "MkDir",
            None,
            None,
            host,
            browser,
        ),
        FileRmDir => record_matrix(
            report,
            project,
            Some(module),
            "FileSystem",
            "RmDir",
            None,
            None,
            host,
            browser,
        ),
        FileOpen => record_matrix(
            report,
            project,
            Some(module),
            "Statements",
            "Open / Close / Get / Put",
            Some("Open"),
            None,
            host,
            browser,
        ),
        FileClose => record_matrix(
            report,
            project,
            Some(module),
            "Statements",
            "Open / Close / Get / Put",
            Some("Close"),
            None,
            host,
            browser,
        ),
        FileGetInto | FileRead => record_matrix(
            report,
            project,
            Some(module),
            "Statements",
            "Open / Close / Get / Put",
            Some("Get"),
            None,
            host,
            browser,
        ),
        FilePut => record_matrix(
            report,
            project,
            Some(module),
            "Statements",
            "Open / Close / Get / Put",
            Some("Put"),
            None,
            host,
            browser,
        ),
        FileLock => record_matrix(
            report,
            project,
            Some(module),
            "Statements",
            "Open / Close / Get / Put",
            Some("Lock"),
            None,
            host,
            browser,
        ),
        FileUnlock => record_matrix(
            report,
            project,
            Some(module),
            "Statements",
            "Open / Close / Get / Put",
            Some("Unlock"),
            None,
            host,
            browser,
        ),
        FileRename => record_matrix(
            report,
            project,
            Some(module),
            "Statements",
            "Name ... As",
            Some("Name"),
            None,
            host,
            browser,
        ),
        FilePrint => record_matrix(
            report,
            project,
            Some(module),
            "Statements",
            "Print # / Write # / Line Input #",
            Some("Print #"),
            None,
            host,
            browser,
        ),
        FileWrite => record_matrix(
            report,
            project,
            Some(module),
            "Statements",
            "Print # / Write # / Line Input #",
            Some("Write #"),
            None,
            host,
            browser,
        ),
        FileLineInput => record_matrix(
            report,
            project,
            Some(module),
            "Statements",
            "Print # / Write # / Line Input #",
            Some("Line Input #"),
            None,
            host,
            browser,
        ),
        FileInput | FileInputChars => record_matrix(
            report,
            project,
            Some(module),
            "_HiddenModule",
            "Input/$",
            Some("Input #"),
            None,
            host,
            browser,
        ),
        FileInputBytes => record_matrix(
            report,
            project,
            Some(module),
            "_HiddenModule",
            "InputB/$",
            None,
            None,
            host,
            browser,
        ),
        FileWidth | FileWidthFunction => record_matrix(
            report,
            project,
            Some(module),
            "_HiddenModule",
            "Width",
            None,
            None,
            host,
            browser,
        ),
        Now => record_matrix(
            report,
            project,
            Some(module),
            "DateTime",
            "Now",
            None,
            None,
            host,
            browser,
        ),
        DateNow | DateNowText => record_matrix(
            report,
            project,
            Some(module),
            "DateTime",
            "Date/$",
            None,
            None,
            host,
            browser,
        ),
        TimeNow | TimeNowText => record_matrix(
            report,
            project,
            Some(module),
            "DateTime",
            "Time/$",
            None,
            None,
            host,
            browser,
        ),
        Timer => record_matrix(
            report,
            project,
            Some(module),
            "DateTime",
            "Timer",
            None,
            None,
            host,
            browser,
        ),
        Command => record_matrix(
            report,
            project,
            Some(module),
            "Interaction",
            "Command/$",
            None,
            None,
            host,
            browser,
        ),
        Beep => record_matrix(
            report,
            project,
            Some(module),
            "Interaction",
            "Beep",
            None,
            None,
            host,
            browser,
        ),
        MsgBox => record_matrix(
            report,
            project,
            Some(module),
            "Interaction",
            "MsgBox",
            None,
            None,
            host,
            browser,
        ),
        InputBox => record_matrix(
            report,
            project,
            Some(module),
            "Interaction",
            "InputBox",
            None,
            None,
            host,
            browser,
        ),
        IMEStatus => record_matrix(
            report,
            project,
            Some(module),
            "Information",
            "IMEStatus",
            None,
            None,
            host,
            browser,
        ),
        Load => record_matrix(
            report,
            project,
            Some(module),
            "VBEGlobal",
            "Load",
            None,
            None,
            host,
            browser,
        ),
        Unload => record_matrix(
            report,
            project,
            Some(module),
            "VBEGlobal",
            "Unload",
            None,
            None,
            host,
            browser,
        ),
        UserForms => record_matrix(
            report,
            project,
            Some(module),
            "VBEGlobal",
            "UserForms",
            None,
            None,
            host,
            browser,
        ),
        _ => {}
    }
}

#[allow(clippy::too_many_arguments)]
fn check_inst(
    report: &mut Report,
    project: &str,
    module: &str,
    program: &OxProgram,
    library: &calcvba_bundle::Bundle,
    inst: &OxInst,
    host: Host,
    browser: BrowserServices,
) {
    match inst {
        OxInst::LoadProcRef { .. } => record_matrix(
            report,
            project,
            Some(module),
            "Language",
            "AddressOf (callbacks)",
            Some("AddressOf"),
            None,
            host,
            browser,
        ),
        OxInst::Pow { .. } => record_matrix(
            report,
            project,
            Some(module),
            "Runtime",
            "^ operator",
            Some("^"),
            None,
            host,
            browser,
        ),
        OxInst::AsNew {
            binding: OxAsNew::ComClass { prog_id },
            ..
        } => {
            let support = com_activation_support(false, Some(prog_id), true, host);
            record(
                report,
                project,
                Some(module),
                "CreateObject",
                FeatureClass::Windows,
                support,
                Some(prog_id.clone()),
                None,
            );
        }
        OxInst::Assign { value, .. } => check_operand_path(report, project, module, value, host),
        OxInst::CallNative { callee, args, .. } => {
            for arg in args {
                match arg {
                    OxCallArg::Operand(op)
                    | OxCallArg::ByValOverride(op)
                    | OxCallArg::Named { value: op, .. } => {
                        check_operand_path(report, project, module, op, host);
                    }
                    _ => {}
                }
            }
            if let OxNativeCallee::Builtin(id) = callee {
                check_builtin(
                    report,
                    project,
                    module,
                    *id,
                    &view_call_args(args),
                    host,
                    browser,
                );
            }
        }
        OxInst::CallExtern { import, args, .. } => {
            for arg in args {
                if let OxArg::ByVal(op) = arg {
                    check_operand_path(report, project, module, op, host);
                }
            }
            let Some(imp) = program.imports.get(import.0) else {
                return;
            };
            if !imp.unit.eq_ignore_ascii_case("VBA") {
                return;
            }
            let id = library
                .exports
                .iter()
                .find(|e| e.token.matches(&imp.token))
                .and_then(|export| match export.target {
                    calcvba_bundle::ExportTarget::Proc(p) => {
                        library.procedures.get(p).and_then(|proc| proc.native)
                    }
                    _ => None,
                });
            if let Some(calcvba_bundle::NativeBody::Library(id)) = id {
                check_builtin(
                    report,
                    project,
                    module,
                    id,
                    &view_proc_args(args),
                    host,
                    browser,
                );
            }
        }
        _ => {}
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::{Engine, ExecBackend, HostConfig, RuntimeProfileId};

    #[test]
    fn prepare_image_session_bytes_enforces_portability_mode_before_linking() {
        let manifest = calcvba_symbol::manifest::SymbolProjectManifest {
            project_name: "ImageProbe".into(),
            project_kind: calcvba_symbol::manifest::ProjectKind::Source,
            modules: vec![calcvba_symbol::manifest::ModuleUnit {
                module_name: "Main".into(),
                module_kind: calcvba_symbol::manifest::ModuleKind::Procedural,
                attributes: calcvba_symbol::manifest::ModuleAttributes::named("Main"),
                source: "Private Declare PtrSafe Sub Absent Lib \"kernel32\" ()\n\
                         Public Sub Main()\n\
                         End Sub\n"
                    .into(),
            }],
            references: Vec::new(),
            reference_projects: Vec::new(),
            conditional_constants: Default::default(),
            conditional_compilation_target: Default::default(),
            codepage: None,
        };
        let core =
            calcvba_bind::bind_program(&manifest, &calcvba_symbol::CatalogTypeLibResolver).unwrap();
        let ox = calcvba_oxir::elaborate::elaborate(&core).unwrap();
        let bytes = OxImage::new(vec![ox]).to_bytes().unwrap();

        let mut engine = Engine::new(HostConfig {
            backend: ExecBackend::Vm3,
        });
        engine.set_runtime_profile(RuntimeProfileId::LinuxStdio);
        engine.set_portability_mode(Mode::Strict);
        let err = match engine.prepare_image_session_bytes(&bytes) {
            Ok(_) => panic!("expected prepare_image_session_bytes rejection"),
            Err(err) => err,
        };
        let rendered = err.diagnostic().render_human();
        assert!(
            rendered.contains("PORT-E-UNSUPPORTED") && rendered.contains("kernel32!Absent"),
            "{rendered}"
        );
    }
}
