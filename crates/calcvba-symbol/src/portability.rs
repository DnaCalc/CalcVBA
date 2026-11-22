//! Static VBA portability analysis. Inspecting a project never resolves or loads
//! native libraries and never executes the workbook.

mod matrix;

use crate::cond_comp::{base_cc_constants_for_target, preprocess};
use crate::manifest::{ModuleKind, ModuleUnit, ProjectReference, SymbolProjectManifest};
use calcvba_com::platform::portable::BuiltinProgId;
use calcvba_syntax::lexer::tokenize;
use calcvba_syntax::{SyntaxKind as K, SyntaxNode, SyntaxToken, parse};
use serde::Serialize;
use std::collections::HashSet;

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
        use calcvba_diagnostics::{
            Diagnostic, DiagnosticPhase, DiagnosticSeverity, DiagnosticSource,
        };
        if self == Self::Native {
            return Ok(Vec::new());
        }

        let report = check_project(projects, host).map_err(|message| {
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

        let mut procedures = HashSet::new();
        for (module, tree) in &prepared {
            if module.module_kind == ModuleKind::Procedural {
                for node in tree.syntax().child_nodes() {
                    if let Some(name) = procedure_name(node)
                        && !node.child_tokens().iter().any(|t| t.kind == K::KwPrivate)
                    {
                        procedures.insert(crate::model::fold_identifier(name.text));
                    }
                }
            }
        }

        for (module, tree) in &prepared {
            let mut scanner = Scanner::new(&mut report, project, Some(module));
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
            let mut names = procedures.clone();
            for node in tree.syntax().child_nodes() {
                collect_declarations(node, &mut names);
            }
            scanner.scan_directives(&module.source);
            scanner.walk(tree.syntax(), None, &names);
        }

        let mut scanner = Scanner::new(&mut report, project, None);
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
                        "00020813-0000-0000-c000-000000000046", // Excel
                        "00020430-0000-0000-c000-000000000046", // OLE Automation
                        "2df8d04c-5bfa-101b-bde5-00aa0044de52", // Office
                    ]
                    .iter()
                    .any(|id| normalized_guid.eq_ignore_ascii_case(id))
                    {
                        (FeatureClass::Host, Support::Host)
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
}

impl<'a> Scanner<'a> {
    fn new(
        report: &'a mut Report,
        project: &'a SymbolProjectManifest,
        module: Option<&'a ModuleUnit>,
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
                row.support[self.report.host as usize],
                None,
                Some(row.mac_dialect),
            );
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
    ) {
        let offset = node.first_token().map(|t| t.offset);
        let statement = match node.kind() {
            K::OpenStmt => Some(("Open / Close / Get / Put", "Open")),
            K::CloseStmt => Some(("Open / Close / Get / Put", "Close")),
            K::GetStmt => Some(("Open / Close / Get / Put", "Get")),
            K::PutStmt => Some(("Open / Close / Get / Put", "Put")),
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
            K::IdentExpr | K::MemberExpr => self.library_use(node, parent, names),
            K::InputStmt => self.matrix("_HiddenModule", "Input/$", offset, Some("Input #")),
            K::SeekStmt => self.matrix("FileSystem", "Seek", offset, None),
            K::WidthStmt => self.matrix("_HiddenModule", "Width", offset, None),
            _ => {}
        }

        let scoped;
        let names = if matches!(node.kind(), K::SubDecl | K::FunctionDecl | K::PropertyDecl) {
            scoped = {
                let mut local = names.clone();
                for parameter in node.param_list().into_iter().flat_map(|list| list.params()) {
                    add_name(&mut local, parameter.declarator_name());
                }
                if let Some(body) = node.body_block() {
                    collect_locals(body, &mut local);
                }
                local
            };
            &scoped
        } else {
            names
        };
        for child in node.child_nodes() {
            self.walk(child, Some(node), names);
        }
    }

    fn library_use(
        &mut self,
        node: SyntaxNode<'_>,
        parent: Option<SyntaxNode<'_>>,
        names: &HashSet<String>,
    ) {
        let (token, qualifier) = if node.kind() == K::MemberExpr {
            let Some(receiver) = node.member_receiver().and_then(library_qualifier) else {
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

        let mut support = row.support[self.report.host as usize];
        let mut detail = None;
        if matches!(row.name, "CreateObject" | "GetObject") {
            let args = parent.and_then(|parent| match parent.kind() {
                K::IndexExpr => parent.index_arg_list(),
                K::CallStmt => parent.call_arg_list(),
                _ => None,
            });
            let args = args.map(|args| args.arg_items()).unwrap_or_default();
            let prog_id = string_argument(&args, usize::from(row.name == "GetObject"), "Class");
            let can_activate = row.name == "CreateObject"
                || string_argument(&args, 0, "PathName").as_deref() == Some("");
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

fn library_qualifier(node: SyntaxNode<'_>) -> Option<String> {
    let name = match node.kind() {
        K::IdentExpr => node.ident_name_token()?.text.to_ascii_lowercase(),
        K::MemberExpr
            if node
                .member_receiver()
                .and_then(library_qualifier)
                .as_deref()
                == Some("vba") =>
        {
            node.member_name_token()?.text.to_ascii_lowercase()
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
        names.insert(crate::model::fold_identifier(name.text));
    }
}

fn collect_declarations(node: SyntaxNode<'_>, names: &mut HashSet<String>) {
    add_name(names, procedure_name(node));
    for decl in node.declarators() {
        add_name(names, decl.declarator_name());
    }
}

fn collect_locals(node: SyntaxNode<'_>, names: &mut HashSet<String>) {
    collect_declarations(node, names);
    for child in node.child_nodes() {
        collect_locals(child, names);
    }
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

fn string_argument(
    args: &[calcvba_syntax::red::ArgItem<'_>],
    position: usize,
    parameter: &str,
) -> Option<String> {
    use calcvba_syntax::red::ArgItem;
    args.iter()
        .find(|arg| {
            matches!(arg, ArgItem::Named { name, .. }
            if name.text.eq_ignore_ascii_case(parameter))
        })
        .or_else(|| {
            args.get(position)
                .filter(|arg| matches!(arg, ArgItem::Positional(..)))
        })
        .and_then(arg_string)
}

fn arg_string(arg: &calcvba_syntax::red::ArgItem<'_>) -> Option<String> {
    use calcvba_syntax::red::ArgItem;
    let value = match arg {
        ArgItem::Positional(value, _) | ArgItem::Named { value, .. } => value,
        ArgItem::Omitted => return None,
    };
    value.literal_token().and_then(string_literal)
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
