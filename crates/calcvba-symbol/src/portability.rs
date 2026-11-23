//! Static VBA portability analysis. Inspecting a project never resolves or loads
//! native libraries and never executes the workbook.

mod matrix;

use crate::cond_comp::{base_cc_constants_for_target, preprocess};
use crate::manifest::{ModuleKind, ModuleUnit, ProjectReference, SymbolProjectManifest};
use calcvba_com::platform::portable::BuiltinProgId;
use calcvba_syntax::lexer::tokenize;
use calcvba_syntax::{SyntaxKind as K, SyntaxNode, SyntaxToken, parse};
use serde::Serialize;
use std::collections::{HashMap, HashSet};

#[derive(Clone, Copy, Debug, Eq, PartialEq, Serialize)]
#[serde(rename_all = "lowercase")]
pub enum Host {
    Windows,
    MacOs,
    Linux,
    Wasi,
    Browser,
}

impl Host {
    pub fn from_name(name: &str) -> Option<Self> {
        Some(match name {
            "windows" => Self::Windows,
            "macos" => Self::MacOs,
            "linux" => Self::Linux,
            "wasi" => Self::Wasi,
            "browser" => Self::Browser,
            _ => return None,
        })
    }

    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Windows => "windows",
            Self::MacOs => "macos",
            Self::Linux => "linux",
            Self::Wasi => "wasi",
            Self::Browser => "browser",
        }
    }

    pub const fn current() -> Self {
        if cfg!(windows) {
            Self::Windows
        } else if cfg!(target_os = "macos") {
            Self::MacOs
        } else if cfg!(target_os = "wasi") {
            Self::Wasi
        } else if cfg!(target_family = "wasm") {
            Self::Browser
        } else {
            Self::Linux
        }
    }
}

/// Select how much host-specific VBA a source project can use.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum Mode {
    WindowsFaithful,
    Strict,
    Native,
}

impl Mode {
    pub fn from_name(name: &str) -> Option<Self> {
        Some(match name {
            "windows-faithful" => Self::WindowsFaithful,
            "strict" => Self::Strict,
            "native" => Self::Native,
            _ => return None,
        })
    }

    pub fn for_project(project: &SymbolProjectManifest) -> Self {
        if project.project_kind == crate::manifest::ProjectKind::Host {
            Self::WindowsFaithful
        } else {
            Self::Native
        }
    }

    pub fn check(
        self,
        projects: &[SymbolProjectManifest],
        host: Host,
    ) -> Result<Vec<calcvba_diagnostics::Diagnostic>, Box<calcvba_diagnostics::Diagnostic>> {
        self.check_with_browser_dialogs(projects, host, false)
    }

    /// A browser worker may admit modal dialogs when it has a synchronous reply channel.
    pub fn check_with_browser_dialogs(
        self,
        projects: &[SymbolProjectManifest],
        host: Host,
        browser_dialogs: bool,
    ) -> Result<Vec<calcvba_diagnostics::Diagnostic>, Box<calcvba_diagnostics::Diagnostic>> {
        self.check_with_browser_services(
            projects,
            host,
            BrowserServices {
                dialogs: browser_dialogs,
                ..BrowserServices::default()
            },
        )
    }

    pub fn check_with_browser_services(
        self,
        projects: &[SymbolProjectManifest],
        host: Host,
        browser: BrowserServices,
    ) -> Result<Vec<calcvba_diagnostics::Diagnostic>, Box<calcvba_diagnostics::Diagnostic>> {
        use calcvba_diagnostics::{
            Diagnostic, DiagnosticPhase, DiagnosticSeverity, DiagnosticSource,
        };
        if self == Self::Native {
            return Ok(Vec::new());
        }

        let report = check_project_inner(projects, host, browser).map_err(|message| {
            Box::new(Diagnostic::error(
                "PORT-E-CHECK",
                DiagnosticPhase::ProjectLoad,
                message,
            ))
        })?;
        let mut warnings = Vec::new();
        for finding in report.findings {
            let location = finding.module.unwrap_or(finding.project);
            let mut source = DiagnosticSource::module(location);
            source.line = finding.line.and_then(|line| u32::try_from(line).ok());
            source.column = finding.line.map(|_| 1);

            let is_shell = finding.feature == "Shell";
            let feature = match finding.detail {
                Some(detail) => format!("{} ({detail})", finding.feature),
                None => finding.feature,
            };
            let reason = if finding.verdict == Support::Reject {
                Some("the host has no supported implementation")
            } else if self == Self::Strict && finding.class == FeatureClass::Windows {
                Some("strict mode refuses Windows-specific constructs")
            } else if self == Self::Strict
                && matches!(finding.verdict, Support::Shim | Support::Replace)
            {
                Some("strict mode requires native or modelled behaviour")
            } else if self == Self::WindowsFaithful && is_shell && host != Host::Windows {
                Some("the host has no Windows command mapping")
            } else {
                None
            };
            if let Some(reason) = reason {
                return Err(Box::new(
                    Diagnostic::error(
                        "PORT-E-UNSUPPORTED",
                        DiagnosticPhase::ProjectLoad,
                        format!("{feature} on {}: {reason}", host.as_str()),
                    )
                    .with_source(source),
                ));
            }
            if self == Self::Strict && finding.verdict == Support::Host {
                let mut warning = Diagnostic::error(
                    "PORT-W-HOST-DEPENDENT",
                    DiagnosticPhase::ProjectLoad,
                    format!("{feature} on {} depends on host values", host.as_str()),
                )
                .with_source(source);
                warning.severity = DiagnosticSeverity::Warning;
                warnings.push(warning);
            }
        }
        Ok(warnings)
    }
}

/// Services supplied by a browser embedding.
#[derive(Clone, Copy, Debug, Default)]
pub struct BrowserServices {
    pub dialogs: bool,
    pub settings: bool,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq, Serialize)]
#[serde(rename_all = "lowercase")]
pub enum FeatureClass {
    Pure,
    Locale,
    Shim,
    Host,
    Windows,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq, Serialize)]
#[serde(rename_all = "lowercase")]
pub enum Support {
    Native,
    Model,
    Shim,
    Host,
    Replace,
    Reject,
}

struct Feature {
    area: &'static str,
    name: &'static str,
    class: FeatureClass,
    support: [Support; 5],
    mac_dialect: &'static str,
}

#[derive(Clone, Copy, Debug, Eq, Ord, PartialEq, PartialOrd, Serialize)]
pub enum Verdict {
    #[serde(rename = "portable")]
    Portable,
    #[serde(rename = "host-dependent values")]
    HostDependentValues,
    #[serde(rename = "portable with shims")]
    PortableWithShims,
    #[serde(rename = "needs Windows")]
    NeedsWindows,
}

#[derive(Debug, Serialize)]
pub struct Finding {
    pub project: String,
    pub module: Option<String>,
    pub line: Option<usize>,
    pub feature: String,
    pub class: FeatureClass,
    pub verdict: Support,
    pub detail: Option<String>,
    pub mac_dialect: Option<&'static str>,
}

#[derive(Debug, Serialize)]
pub struct Report {
    pub host: Host,
    pub verdict: Verdict,
    pub findings: Vec<Finding>,
    pub parse_errors: Vec<ParseIssue>,
}

#[derive(Debug, Serialize)]
pub struct ParseIssue {
    pub project: String,
    pub module: String,
    pub line: usize,
    pub message: String,
}

pub fn check_project(projects: &[SymbolProjectManifest], host: Host) -> Result<Report, String> {
    check_project_inner(projects, host, BrowserServices::default())
}

fn check_project_inner(
    projects: &[SymbolProjectManifest],
    host: Host,
    browser: BrowserServices,
) -> Result<Report, String> {
    let mut report = Report {
        host,
        verdict: Verdict::Portable,
        findings: Vec::new(),
        parse_errors: Vec::new(),
    };
    for project in projects {
        let constants = base_cc_constants_for_target(
            &project.conditional_constants,
            project.conditional_compilation_target,
        );
        let prepared = project
            .modules
            .iter()
            .map(|module| {
                let source = preprocess(&module.source, &constants).map_err(|error| {
                    format!("{}.{}: {error}", project.project_name, module.module_name)
                })?;
                let tree = parse(&source);
                Ok((module, tree))
            })
            .collect::<Result<Vec<_>, String>>()?;

        let mut project_names = HashSet::new();
        let mut project_consts = HashMap::new();
        for (module, tree) in &prepared {
            project_names.insert(crate::model::fold_identifier(&module.module_name));
            if !module.attributes.vb_name.is_empty() {
                project_names.insert(crate::model::fold_identifier(&module.attributes.vb_name));
            }
            collect_public_exports(
                module,
                tree.syntax(),
                false,
                &mut project_names,
                &mut project_consts,
            );
        }
        collect_referenced_project_exports(
            project,
            projects,
            &mut project_names,
            &mut project_consts,
        )?;

        for (module, tree) in &prepared {
            let mut scanner = Scanner::new(&mut report, project, Some(module), browser);
            for error in tree.errors() {
                let line = scanner.line(error.offset);
                scanner.report.parse_errors.push(ParseIssue {
                    project: project.project_name.clone(),
                    module: module.module_name.clone(),
                    line,
                    message: error.message.clone(),
                });
            }
            if module.module_kind == ModuleKind::Form {
                scanner.matrix("VBEGlobal", "UserForms", None, None);
            }
            let mut names = project_names.clone();
            let mut consts = project_consts.clone();
            for node in tree.syntax().child_nodes() {
                collect_declarations(node, &mut names);
                collect_string_consts(node, &mut consts);
            }
            scanner.scan_directives(&module.source);
            scanner.walk(tree.syntax(), None, &names, &consts);
        }

        let mut scanner = Scanner::new(&mut report, project, None, browser);
        for reference in &project.references {
            match reference {
                ProjectReference::TypeLibrary { name, guid, .. } => {
                    let detail = Some(
                        guid.as_ref()
                            .map_or_else(|| name.clone(), |guid| format!("{name} ({guid})")),
                    );
                    let normalized_guid = guid.as_deref().unwrap_or("").trim_matches(['{', '}']);
                    let (class, support) = if normalized_guid
                        .eq_ignore_ascii_case("000204ef-0000-0000-c000-000000000046")
                    {
                        (FeatureClass::Pure, Support::Native)
                    } else if [
                        "00020813-0000-0000-c000-000000000046",
                        "00020430-0000-0000-c000-000000000046",
                        "2df8d04c-5bfa-101b-bde5-00aa0044de52",
                    ]
                    .iter()
                    .any(|id| normalized_guid.eq_ignore_ascii_case(id))
                    {
                        (FeatureClass::Host, Support::Host)
                    } else if normalized_guid
                        .eq_ignore_ascii_case("420b2830-e718-11cf-893d-00a0c9054228")
                    {
                        let support = if scanner.report.host == Host::Windows {
                            Support::Native
                        } else if scripting_reference_uses_only_dictionary(
                            &prepared,
                            &project_names,
                        ) {
                            Support::Replace
                        } else {
                            Support::Reject
                        };
                        (FeatureClass::Windows, support)
                    } else {
                        (FeatureClass::Windows, scanner.windows_native())
                    };
                    scanner.record("Early-bound reference", None, class, support, detail, None);
                }
                ProjectReference::Native { name, path } => {
                    let library = path.as_deref().unwrap_or(name);
                    scanner.declare(None, library, None, None);
                }
                ProjectReference::HostInjected {
                    referenced_project_name,
                } => scanner.record(
                    "Early-bound reference",
                    None,
                    FeatureClass::Host,
                    Support::Host,
                    Some(referenced_project_name.clone()),
                    None,
                ),
                ProjectReference::Project {
                    referenced_project_name,
                } => scanner.record(
                    "Early-bound reference",
                    None,
                    FeatureClass::Pure,
                    Support::Native,
                    Some(referenced_project_name.clone()),
                    None,
                ),
            }
        }
    }
    Ok(report)
}

pub fn check_source(name: &str, source: &str, host: Host) -> Result<Report, String> {
    use crate::manifest as sym;
    check_project(
        &[SymbolProjectManifest {
            project_name: name.into(),
            project_kind: sym::ProjectKind::Source,
            modules: vec![ModuleUnit {
                module_name: name.into(),
                module_kind: ModuleKind::Procedural,
                attributes: sym::ModuleAttributes::named(name),
                source: source.into(),
            }],
            references: Vec::new(),
            reference_projects: Vec::new(),
            conditional_constants: Default::default(),
            conditional_compilation_target: Default::default(),
        }],
        host,
    )
}

struct Scanner<'a> {
    report: &'a mut Report,
    project: &'a str,
    module: Option<&'a str>,
    line_starts: Vec<usize>,
    browser: BrowserServices,
}

impl<'a> Scanner<'a> {
    fn new(
        report: &'a mut Report,
        project: &'a SymbolProjectManifest,
        module: Option<&'a ModuleUnit>,
        browser: BrowserServices,
    ) -> Self {
        let mut line_starts = vec![0];
        if let Some(module) = module {
            let mut offset = 0;
            for (kind, text) in tokenize(&module.source) {
                offset += text.len();
                if matches!(kind, K::Newline | K::LineContinuation) {
                    line_starts.push(offset);
                }
            }
        }
        Self {
            report,
            project: &project.project_name,
            module: module.map(|m| m.module_name.as_str()),
            line_starts,
            browser,
        }
    }

    fn line(&self, offset: u32) -> usize {
        self.line_starts
            .partition_point(|start| *start <= offset as usize)
    }

    fn record(
        &mut self,
        feature: &str,
        offset: Option<u32>,
        class: FeatureClass,
        support: Support,
        detail: Option<String>,
        mac_dialect: Option<&'static str>,
    ) {
        let verdict = match support {
            Support::Native | Support::Model => Verdict::Portable,
            Support::Shim | Support::Replace => Verdict::PortableWithShims,
            Support::Host => Verdict::HostDependentValues,
            Support::Reject => Verdict::NeedsWindows,
        };
        self.report.verdict = self.report.verdict.max(verdict);
        self.report.findings.push(Finding {
            project: self.project.into(),
            module: self.module.map(str::to_owned),
            line: offset.map(|offset| self.line(offset)),
            feature: feature.into(),
            class,
            verdict: support,
            detail,
            mac_dialect,
        });
    }

    fn matrix(&mut self, area: &str, name: &str, offset: Option<u32>, label: Option<&str>) {
        if let Some(row) = matrix::MATRIX
            .iter()
            .find(|row| row.area == area && row.name == name)
        {
            self.record(
                label.unwrap_or(row.name),
                offset,
                row.class,
                self.support(row),
                None,
                Some(row.mac_dialect),
            );
        }
    }

    fn support(&self, row: &Feature) -> Support {
        if self.report.host == Host::Browser
            && (row.area == "FileSystem"
                || matches!(
                    row.name,
                    "Open / Close / Get / Put"
                        | "Print # / Write # / Line Input #"
                        | "Name ... As"
                        | "Input/$"
                        | "InputB/$"
                        | "Width"
                        | "Environ/$"
                )
                || (!self.browser.settings
                    && matches!(
                        row.name,
                        "DeleteSetting" | "GetAllSettings" | "GetSetting" | "SaveSetting"
                    ))
                || (!self.browser.dialogs && matches!(row.name, "MsgBox" | "InputBox")))
        {
            Support::Reject
        } else {
            row.support[self.report.host as usize]
        }
    }

    fn windows_native(&self) -> Support {
        if self.report.host == Host::Windows {
            Support::Native
        } else {
            Support::Reject
        }
    }

    fn declare(
        &mut self,
        offset: Option<u32>,
        library: &str,
        entry: Option<&str>,
        node: Option<SyntaxNode<'_>>,
    ) {
        let lib = library.to_ascii_lowercase();
        let (class, support) = if lib.ends_with(".dylib") {
            (
                FeatureClass::Pure,
                if self.report.host == Host::MacOs {
                    Support::Native
                } else {
                    Support::Reject
                },
            )
        } else if lib.ends_with(".so") || lib.contains(".so.") {
            (
                FeatureClass::Shim,
                if self.report.host == Host::Linux {
                    Support::Native
                } else {
                    Support::Reject
                },
            )
        } else {
            (
                FeatureClass::Windows,
                if matches!(self.report.host, Host::Linux | Host::MacOs)
                    && entry
                        .zip(node)
                        .is_some_and(|(entry, node)| has_win32_shim(library, entry, node))
                {
                    Support::Replace
                } else {
                    self.windows_native()
                },
            )
        };
        let detail = match entry {
            Some(entry) => format!("{library}!{entry}"),
            None => library.to_owned(),
        };
        let label = if node.is_some() {
            "Declare"
        } else {
            "Native reference"
        };
        self.record(label, offset, class, support, Some(detail), None);
    }

    fn scan_directives(&mut self, source: &str) {
        for (line, offset) in self.line_starts.clone().into_iter().enumerate() {
            let end = self
                .line_starts
                .get(line + 1)
                .copied()
                .unwrap_or(source.len());
            let text = source[offset..end].trim_start().to_ascii_lowercase();
            if text.starts_with("#if ") || text.starts_with("#elseif ") {
                self.matrix(
                    "Language",
                    "#If Win64 / Win32 / Mac / VBA7",
                    Some(offset as u32),
                    Some("#If"),
                );
            }
        }
    }

    fn walk(
        &mut self,
        node: SyntaxNode<'_>,
        parent: Option<SyntaxNode<'_>>,
        names: &HashSet<String>,
        consts: &HashMap<String, String>,
    ) {
        let offset = node.first_token().map(|t| t.offset);
        let statement = match node.kind() {
            K::OpenStmt => Some(("Open / Close / Get / Put", "Open")),
            K::CloseStmt => Some(("Open / Close / Get / Put", "Close")),
            K::GetStmt => Some(("Open / Close / Get / Put", "Get")),
            K::PutStmt => Some(("Open / Close / Get / Put", "Put")),
            K::LockStmt => {
                let is_unlock = node
                    .first_significant_token()
                    .is_some_and(|token| token.text.eq_ignore_ascii_case("unlock"));
                Some((
                    "Open / Close / Get / Put",
                    if is_unlock { "Unlock" } else { "Lock" },
                ))
            }
            K::PrintStmt => Some(("Print # / Write # / Line Input #", "Print #")),
            K::WriteStmt => Some(("Print # / Write # / Line Input #", "Write #")),
            K::LineInputStmt => Some(("Print # / Write # / Line Input #", "Line Input #")),
            K::NameStmt => Some(("Name ... As", "Name")),
            K::OptionStmt
                if node
                    .text()
                    .split_whitespace()
                    .any(|word| word.eq_ignore_ascii_case("text")) =>
            {
                Some(("Option Compare Text, Like", "Option Compare Text"))
            }
            _ => None,
        };
        if let Some((matrix, label)) = statement {
            self.matrix("Statements", matrix, offset, Some(label));
        }
        match node.kind() {
            K::DeclareStmt => {
                let library = node.child_node(K::LibClause).and_then(clause_string);
                let alias = node.child_node(K::AliasClause).and_then(clause_string);
                let name = node.proc_name_token();
                self.declare(
                    offset,
                    library.as_deref().unwrap_or("<unknown>"),
                    alias.as_deref().or(name.map(|t| t.text)),
                    Some(node),
                );
            }
            K::AddressOfExpr => self.matrix(
                "Language",
                "AddressOf (callbacks)",
                offset,
                Some("AddressOf"),
            ),
            K::TypeRef if node.text().trim().eq_ignore_ascii_case("LongPtr") => self.matrix(
                "Language",
                "LongPtr / pointer width",
                offset,
                Some("LongPtr"),
            ),
            K::BinaryExpr => match node.binary_op_token().map(|t| t.kind) {
                Some(K::Caret) => self.matrix("Runtime", "^ operator", offset, Some("^")),
                Some(K::KwLike) => self.matrix(
                    "Statements",
                    "Option Compare Text, Like",
                    offset,
                    Some("Like"),
                ),
                _ => {}
            },
            K::LiteralExpr => {
                if let Some(path) = node.literal_token().and_then(string_literal)
                    && looks_like_windows_path(&path)
                {
                    let support = if self.report.host == Host::Windows {
                        Support::Native
                    } else {
                        Support::Shim
                    };
                    self.record(
                        "Windows path",
                        offset,
                        FeatureClass::Shim,
                        support,
                        Some(path),
                        None,
                    );
                }
            }
            K::IdentExpr | K::MemberExpr => self.library_use(node, parent, names, consts),
            K::ErrorStmt => self.matrix("Conversion", "Error/$", offset, Some("Error")),
            K::InputStmt => self.matrix("_HiddenModule", "Input/$", offset, Some("Input #")),
            K::SeekStmt => self.matrix("FileSystem", "Seek", offset, None),
            K::WidthStmt => self.matrix("_HiddenModule", "Width", offset, None),
            _ => {}
        }

        let scoped_names;
        let scoped_consts;
        let (names, consts) =
            if matches!(node.kind(), K::SubDecl | K::FunctionDecl | K::PropertyDecl) {
                let mut local_names = names.clone();
                let mut local_consts = consts.clone();
                for parameter in node.param_list().into_iter().flat_map(|list| list.params()) {
                    if let Some(name) = parameter.declarator_name() {
                        let folded = crate::model::fold_identifier(name.text);
                        local_consts.remove(&folded);
                        local_names.insert(folded);
                    }
                }
                if let Some(body) = node.body_block() {
                    collect_locals(body, &mut local_names, &mut local_consts);
                }
                scoped_names = local_names;
                scoped_consts = local_consts;
                (&scoped_names, &scoped_consts)
            } else {
                (names, consts)
            };
        for child in node.child_nodes() {
            self.walk(child, Some(node), names, consts);
        }
    }

    fn library_use(
        &mut self,
        node: SyntaxNode<'_>,
        parent: Option<SyntaxNode<'_>>,
        names: &HashSet<String>,
        consts: &HashMap<String, String>,
    ) {
        let (token, qualifier) = if node.kind() == K::MemberExpr {
            let Some(receiver) = node
                .member_receiver()
                .and_then(|receiver| library_qualifier(receiver, names))
            else {
                return;
            };
            (node.member_name_token(), Some(receiver))
        } else {
            if parent.is_some_and(|parent| {
                parent.kind() == K::MemberExpr && !node.text().trim().eq_ignore_ascii_case("Err")
            }) {
                return;
            }
            (node.ident_name_token(), None)
        };
        let Some(token) = token else { return };
        let name = token.text.trim_matches(['[', ']']);
        if qualifier.is_none() && names.contains(&crate::model::fold_identifier(name)) {
            return;
        }
        let Some(row) = matrix::MATRIX.iter().find(|row| {
            row.name.trim_end_matches("/$").eq_ignore_ascii_case(name)
                && match qualifier.as_deref() {
                    Some("vba") | None => global_area(row.area),
                    Some(area) => row.area.eq_ignore_ascii_case(area),
                }
        }) else {
            return;
        };

        let mut support = self.support(row);
        let mut detail = None;
        if matches!(row.name, "CreateObject" | "GetObject") {
            let args = parent.and_then(|parent| match parent.kind() {
                K::IndexExpr => parent.index_arg_list(),
                K::CallExpr => parent.child_node(K::ArgList),
                K::CallStmt => parent.call_arg_list(),
                _ => None,
            });
            let args = args.map(|args| args.arg_items()).unwrap_or_default();
            let prog_id =
                string_argument(&args, usize::from(row.name == "GetObject"), "Class", consts);
            let can_activate = if row.name == "CreateObject" {
                match argument_item(&args, 1, "ServerName") {
                    None | Some(calcvba_syntax::red::ArgItem::Omitted) => true,
                    Some(server) => arg_string(server, consts).as_deref() == Some(""),
                }
            } else {
                string_argument(&args, 0, "PathName", consts).as_deref() == Some("")
            };
            let portable = match prog_id.as_deref().and_then(BuiltinProgId::resolve) {
                Some(BuiltinProgId::Dictionary) => true,
                Some(BuiltinProgId::RegExp) => row.name == "CreateObject",
                Some(BuiltinProgId::FileSystemObject | BuiltinProgId::XmlDocument) => {
                    self.report.host != Host::Browser
                }
                None => false,
            };
            if support == Support::Replace && !(can_activate && portable) {
                support = Support::Reject;
            }
            detail = Some(prog_id.unwrap_or_else(|| "dynamic or omitted ProgID".into()));
        }
        self.record(
            row.name.trim_end_matches("/$"),
            Some(token.offset),
            row.class,
            support,
            detail,
            Some(row.mac_dialect),
        );
    }
}

fn has_win32_shim(library: &str, entry: &str, node: SyntaxNode<'_>) -> bool {
    use crate::scanner::parameter_has_modifier;
    use calcvba_runtime::win32_declare::Win32Declare;

    fn abi_type(node: SyntaxNode<'_>) -> &'static str {
        let text = node.text();
        [
            "Any", "Byte", "Boolean", "Integer", "Long", "LongLong", "LongPtr", "Single", "Double",
            "Currency", "Date",
        ]
        .into_iter()
        .find(|name| name.eq_ignore_ascii_case(text.trim()))
        .unwrap_or("")
    }

    let mut params = Vec::new();
    for param in node.param_list().into_iter().flat_map(|list| list.params()) {
        if param.array_bounds().is_some()
            || parameter_has_modifier(param, K::KwOptional)
            || parameter_has_modifier(param, K::KwParamArray)
        {
            return false;
        }
        let ty = param.child_node(K::TypeRef).map_or("", abi_type);
        params.push((ty, !parameter_has_modifier(param, K::KwByVal)));
    }
    let result = node
        .child_tokens()
        .iter()
        .any(|token| token.kind == K::KwFunction)
        .then(|| node.return_type().map_or("", abi_type));
    Win32Declare::resolve(library, entry, &params, result).is_some()
}

fn global_area(area: &str) -> bool {
    matches!(
        area,
        "_HiddenModule"
            | "Conversion"
            | "DateTime"
            | "FileSystem"
            | "Financial"
            | "Information"
            | "Interaction"
            | "Math"
            | "Strings"
            | "VBEGlobal"
    )
}

fn library_qualifier(node: SyntaxNode<'_>, names: &HashSet<String>) -> Option<String> {
    let name = match node.kind() {
        K::IdentExpr => {
            let raw = node.ident_name_token()?.text.trim_matches(['[', ']']);
            let folded = crate::model::fold_identifier(raw);
            if names.contains(&folded) {
                return None;
            }
            folded
        }
        K::MemberExpr
            if node
                .member_receiver()
                .and_then(|receiver| library_qualifier(receiver, names))
                .as_deref()
                == Some("vba") =>
        {
            node.member_name_token()?
                .text
                .trim_matches(['[', ']'])
                .to_ascii_lowercase()
        }
        _ => return None,
    };
    Some(if name == "err" {
        "errobject".into()
    } else {
        name
    })
}

fn procedure_name(node: SyntaxNode<'_>) -> Option<SyntaxToken<'_>> {
    if matches!(
        node.kind(),
        K::SubDecl | K::FunctionDecl | K::PropertyDecl | K::DeclareStmt
    ) {
        node.proc_name_token()
    } else {
        None
    }
}

fn add_name(names: &mut HashSet<String>, name: Option<SyntaxToken<'_>>) {
    if let Some(name) = name {
        names.insert(crate::model::fold_identifier(
            name.text.trim_matches(['[', ']']),
        ));
    }
}

fn is_private_decl(node: SyntaxNode<'_>) -> bool {
    node.child_tokens().iter().any(|t| t.kind == K::KwPrivate)
}

fn is_friend_decl(node: SyntaxNode<'_>) -> bool {
    node.child_tokens().iter().any(|t| t.kind == K::KwFriend)
}

fn is_explicit_public_decl(node: SyntaxNode<'_>) -> bool {
    node.child_tokens()
        .iter()
        .any(|t| t.kind == K::KwPublic || t.text.eq_ignore_ascii_case("global"))
}

fn has_option_private_module(syntax: SyntaxNode<'_>) -> bool {
    syntax.child_nodes().into_iter().any(|node| {
        node.kind() == K::OptionStmt
            && node
                .text()
                .split_whitespace()
                .any(|word| word.eq_ignore_ascii_case("private"))
    })
}

fn source_has_global_namespace(syntax: SyntaxNode<'_>) -> bool {
    syntax.child_nodes().into_iter().any(|node| {
        node.kind() == K::AttributeStmt
            && node
                .text()
                .to_ascii_lowercase()
                .contains("vb_globalnamespace")
            && node.text().to_ascii_lowercase().contains("true")
    })
}

fn collect_public_exports(
    module: &ModuleUnit,
    syntax: SyntaxNode<'_>,
    cross_project: bool,
    names: &mut HashSet<String>,
    consts: &mut HashMap<String, String>,
) {
    let exposes_unqualified = module.module_kind == ModuleKind::Procedural
        || module.attributes.vb_global_namespace
        || source_has_global_namespace(syntax);
    if !exposes_unqualified {
        return;
    }
    if cross_project
        && (module.attributes.option_private_module || has_option_private_module(syntax))
    {
        return;
    }
    for node in syntax.child_nodes() {
        if is_private_decl(node) || (cross_project && is_friend_decl(node)) {
            continue;
        }
        match node.kind() {
            K::SubDecl | K::FunctionDecl | K::PropertyDecl | K::DeclareStmt => {
                add_name(names, procedure_name(node));
            }
            K::DimStmt => {
                if is_explicit_public_decl(node) {
                    for decl in node.declarators() {
                        add_name(names, decl.declarator_name());
                    }
                }
            }
            K::ConstStmt => {
                if is_explicit_public_decl(node) {
                    for decl in node.declarators() {
                        add_name(names, decl.declarator_name());
                    }
                    collect_string_consts(node, consts);
                }
            }
            K::EnumBlock => {
                add_name(names, node.name_token());
                for member in node.enum_members() {
                    add_name(names, member.declarator_name());
                }
            }
            K::TypeBlock => {
                add_name(names, node.name_token());
            }
            _ => {}
        }
    }
}

fn collect_referenced_project_exports(
    project: &SymbolProjectManifest,
    projects: &[SymbolProjectManifest],
    names: &mut HashSet<String>,
    consts: &mut HashMap<String, String>,
) -> Result<(), String> {
    for reference in &project.references {
        let ProjectReference::Project {
            referenced_project_name,
        } = reference
        else {
            continue;
        };
        if let Some(target_project) = projects.iter().find(|other| {
            other
                .project_name
                .eq_ignore_ascii_case(referenced_project_name)
        }) {
            let constants = base_cc_constants_for_target(
                &target_project.conditional_constants,
                target_project.conditional_compilation_target,
            );
            for module in &target_project.modules {
                let source = preprocess(&module.source, &constants).map_err(|error| {
                    format!(
                        "{}.{}: {error}",
                        target_project.project_name, module.module_name
                    )
                })?;
                let tree = parse(&source);
                collect_public_exports(module, tree.syntax(), true, names, consts);
            }
        } else if let Some(ref_project) = project.reference_projects.iter().find(|other| {
            other
                .project_name
                .eq_ignore_ascii_case(referenced_project_name)
        }) {
            let target = ref_project
                .conditional_compilation_target
                .unwrap_or(project.conditional_compilation_target);
            let constants = base_cc_constants_for_target(&Default::default(), target);
            for module in &ref_project.modules {
                let source = preprocess(&module.source, &constants).map_err(|error| {
                    format!(
                        "{}.{}: {error}",
                        ref_project.project_name, module.module_name
                    )
                })?;
                let tree = parse(&source);
                collect_public_exports(module, tree.syntax(), true, names, consts);
            }
        }
    }
    Ok(())
}

fn collect_declarations(node: SyntaxNode<'_>, names: &mut HashSet<String>) {
    add_name(names, procedure_name(node));
    for decl in node.declarators() {
        add_name(names, decl.declarator_name());
    }
    match node.kind() {
        K::EnumBlock => {
            add_name(names, node.name_token());
            for member in node.enum_members() {
                add_name(names, member.declarator_name());
            }
        }
        K::TypeBlock | K::EventDecl => {
            add_name(names, node.name_token());
        }
        _ => {}
    }
}

fn collect_string_consts(node: SyntaxNode<'_>, consts: &mut HashMap<String, String>) {
    if node.kind() != K::ConstStmt {
        return;
    }
    for decl in node.declarators() {
        let Some(name) = decl.declarator_name() else {
            continue;
        };
        let folded = crate::model::fold_identifier(name.text.trim_matches(['[', ']']));
        if let Some(value) = decl
            .expr_children()
            .into_iter()
            .next()
            .and_then(|expr| expr_string(expr, consts))
        {
            consts.insert(folded, value);
        }
    }
}

fn collect_locals(
    node: SyntaxNode<'_>,
    names: &mut HashSet<String>,
    consts: &mut HashMap<String, String>,
) {
    if node.kind() == K::DimStmt {
        for decl in node.declarators() {
            if let Some(name) = decl.declarator_name() {
                consts.remove(&crate::model::fold_identifier(
                    name.text.trim_matches(['[', ']']),
                ));
            }
        }
    }
    collect_declarations(node, names);
    collect_string_consts(node, consts);
    for child in node.child_nodes() {
        collect_locals(child, names, consts);
    }
}

fn scripting_reference_uses_only_dictionary(
    prepared: &[(&ModuleUnit, calcvba_syntax::Parse)],
    project_names: &HashSet<String>,
) -> bool {
    const NON_DICT_SCRIPTING_TYPES: &[&str] = &[
        "filesystemobject",
        "drive",
        "drives",
        "folder",
        "folders",
        "file",
        "files",
        "textstream",
        "encoder",
    ];

    fn check_type_name(raw: &str, project_names: &HashSet<String>) -> bool {
        let trimmed = raw.trim();
        let trimmed = trimmed
            .strip_prefix("New ")
            .or_else(|| trimmed.strip_prefix("new "))
            .or_else(|| trimmed.strip_prefix("NEW "))
            .unwrap_or(trimmed)
            .trim();
        let lower = trimmed.to_ascii_lowercase();
        if let Some(member) = lower.strip_prefix("scripting.") {
            return member == "dictionary" || member == "idictionary";
        }
        if !lower.contains('.')
            && NON_DICT_SCRIPTING_TYPES.contains(&lower.as_str())
            && !project_names.contains(&lower)
        {
            return false;
        }
        true
    }

    fn walk_types(node: SyntaxNode<'_>, project_names: &HashSet<String>) -> bool {
        if matches!(node.kind(), K::TypeRef | K::NewExpr)
            && !check_type_name(&node.text(), project_names)
        {
            return false;
        }
        node.child_nodes()
            .into_iter()
            .all(|child| walk_types(child, project_names))
    }

    prepared
        .iter()
        .all(|(_, tree)| walk_types(tree.syntax(), project_names))
}

fn string_literal(token: SyntaxToken<'_>) -> Option<String> {
    if token.kind != K::StringLiteral {
        return None;
    }
    Some(
        token
            .text
            .strip_prefix('"')?
            .strip_suffix('"')?
            .replace("\"\"", "\""),
    )
}

fn clause_string(node: SyntaxNode<'_>) -> Option<String> {
    node.child_tokens().into_iter().find_map(string_literal)
}

fn argument_item<'a>(
    args: &'a [calcvba_syntax::red::ArgItem<'a>],
    position: usize,
    parameter: &str,
) -> Option<&'a calcvba_syntax::red::ArgItem<'a>> {
    use calcvba_syntax::red::ArgItem;
    args.iter()
        .find(|arg| {
            matches!(arg, ArgItem::Named { name, .. }
            if name.text.eq_ignore_ascii_case(parameter))
        })
        .or_else(|| {
            args.get(position)
                .filter(|arg| matches!(arg, ArgItem::Positional(..) | ArgItem::Omitted))
        })
}

fn string_argument(
    args: &[calcvba_syntax::red::ArgItem<'_>],
    position: usize,
    parameter: &str,
    consts: &HashMap<String, String>,
) -> Option<String> {
    argument_item(args, position, parameter).and_then(|arg| arg_string(arg, consts))
}

fn expr_string(mut node: SyntaxNode<'_>, consts: &HashMap<String, String>) -> Option<String> {
    while node.kind() == K::ParenExpr {
        node = node.expr_children().into_iter().next()?;
    }
    match node.kind() {
        K::LiteralExpr => node.literal_token().and_then(string_literal),
        K::IdentExpr => {
            let name = node.ident_name_token()?.text.trim_matches(['[', ']']);
            consts.get(&crate::model::fold_identifier(name)).cloned()
        }
        _ => None,
    }
}

fn arg_string(
    arg: &calcvba_syntax::red::ArgItem<'_>,
    consts: &HashMap<String, String>,
) -> Option<String> {
    use calcvba_syntax::red::ArgItem;
    let value = match arg {
        ArgItem::Positional(value, _) | ArgItem::Named { value, .. } => value,
        ArgItem::Omitted => return None,
    };
    expr_string(*value, consts)
}

fn looks_like_windows_path(text: &str) -> bool {
    let bytes = text.as_bytes();
    text.starts_with("\\\\")
        || matches!(bytes, [drive, b':', b'\\' | b'/', ..] if drive.is_ascii_alphabetic())
        || (!text.chars().any(char::is_control)
            && text
                .split_once('\\')
                .is_some_and(|(head, tail)| !head.is_empty() && !tail.is_empty()))
}
