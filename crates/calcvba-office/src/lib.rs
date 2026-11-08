//! Office VBA source ingestion and write-back, without Office or platform APIs.
//! The normative format is MS-OVBA; CFB and OPC carry the same VBA project.
#![forbid(unsafe_code)]

pub mod compression;
mod directory;
#[cfg(feature = "files")]
mod folder;
mod storage;
use calcvba_symbol::manifest as sym;
use std::collections::{BTreeMap, BTreeSet};
#[cfg(feature = "files")]
pub use storage::OfficeFile;
pub use storage::ReadLimits;

#[derive(Debug, thiserror::Error)]
pub enum Error {
    #[error("Office VBA I/O: {0}")]
    Io(#[from] std::io::Error),
    #[cfg(feature = "files")]
    #[error("Office package: {0}")]
    Zip(#[from] zip::result::ZipError),
    #[error("invalid Office VBA project: {0}")]
    Invalid(String),
    #[error("Office VBA exceeds {0} limit")]
    Limit(&'static str),
    #[error("unsupported project codepage {0}")]
    Codepage(u16),
    #[error("text cannot be represented exactly in project codepage {0}")]
    Encoding(u16),
    #[error("workbook has no VBA project")]
    NoProject,
    #[error("locked VBA projects cannot be read or changed")]
    Locked,
    #[error("unsupported Office VBA operation: {0}")]
    Unsupported(String),
}
pub type Result<T> = std::result::Result<T, Error>;
pub(crate) fn invalid(message: impl std::fmt::Display) -> Error {
    Error::Invalid(message.to_string())
}

pub(crate) fn decode(bytes: &[u8], cp: u16) -> Result<String> {
    codepage::to_encoding(cp)
        .ok_or(Error::Codepage(cp))?
        .decode_without_bom_handling_and_without_replacement(bytes)
        .map(|s| s.into_owned())
        .ok_or(Error::Encoding(cp))
}
pub(crate) fn encode(text: &str, cp: u16) -> Result<Vec<u8>> {
    let encoding = codepage::to_encoding(cp).ok_or(Error::Codepage(cp))?;
    let (bytes, _, errors) = encoding.encode(text);
    // Some legacy encodings map several characters to the same byte sequence.
    // Round-trip comparison also rejects those lossy mappings.
    if errors || decode(&bytes, cp)? != text {
        return Err(Error::Encoding(cp));
    }
    Ok(bytes.into_owned())
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ProjectProperties {
    pub name: String,
    pub codepage: u16,
    pub sys_kind: u32,
    pub compat_version: Option<u32>,
    pub lcid: u32,
    pub lcid_invoke: u32,
    pub description: String,
    pub help_file: String,
    pub help_context: u32,
    pub lib_flags: u32,
    pub version_major: u32,
    pub version_minor: u16,
    pub constants: String,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ModuleKind {
    Standard,
    Class,
    Document,
    Form,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Module {
    pub name: String,
    pub stream_name: String,
    pub kind: ModuleKind,
    pub description: String,
    pub help_context: u32,
    pub read_only: bool,
    pub private: bool,
    pub source: String,
    /// The host's document class/version suffix in PROJECT (e.g. 0x00000000).
    pub document_version: Option<String>,
    pub(crate) text_offset: u32,
}

impl Module {
    /// Body as shown in the VBE code pane. Export-only attributes remain in `source`.
    pub fn code(&self) -> String {
        self.source
            .lines()
            .filter(|line| {
                !line
                    .trim_start()
                    .to_ascii_lowercase()
                    .starts_with("attribute ")
            })
            .collect::<Vec<_>>()
            .join("\r\n")
    }
    /// Byte offset of the compressed source in the module's stream (MS-OVBA MODULEOFFSET):
    /// the compiled performance cache comes before it.
    pub fn text_offset(&self) -> u32 {
        self.text_offset
    }
    pub fn attribute(&self, name: &str) -> Option<&str> {
        self.source.lines().find_map(|line| {
            let line = line.trim();
            let (lhs, rhs) = line.split_once('=')?;
            let lhs = lhs.trim();
            let key = lhs
                .get(..10)
                .filter(|p| p.eq_ignore_ascii_case("Attribute "))?;
            let _ = key;
            lhs[10..]
                .trim()
                .eq_ignore_ascii_case(name)
                .then(|| rhs.trim())
        })
    }
    fn boolean_attribute(&self, name: &str) -> bool {
        self.attribute(name)
            .is_some_and(|s| s.eq_ignore_ascii_case("True"))
    }
    /// Document host type comes from the persisted VB_Base identity, not its name.
    pub fn host_document_type(&self) -> Result<Option<String>> {
        if self.kind != ModuleKind::Document {
            return Ok(None);
        }
        let base = self.attribute("VB_Base").unwrap_or("").to_ascii_lowercase();
        let ty = if base.contains("00020819-") {
            "Excel.Workbook"
        } else if base.contains("00020820-") {
            "Excel.Worksheet"
        } else if base.contains("00020821-") {
            "Excel.Chart"
        } else {
            return Err(Error::Unsupported(format!(
                "document {} has unknown VB_Base",
                self.name
            )));
        };
        Ok(Some(ty.into()))
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Reference {
    pub name: String,
    pub kind: ReferenceKind,
}
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum ReferenceKind {
    Registered {
        libid: String,
    },
    Project {
        absolute: String,
        relative: String,
        major: u32,
        minor: u16,
    },
    Control {
        original: Option<String>,
        twiddled: String,
        extended_name: Option<String>,
        extended: String,
        guid: [u8; 16],
        cookie: u32,
    },
}
impl Reference {
    pub fn compiler_reference(&self) -> Result<sym::ProjectReference> {
        if matches!(self.kind, ReferenceKind::Project { .. }) {
            if self.name.is_empty() {
                return Err(invalid("referenced project has no name"));
            }
            Ok(sym::ProjectReference::Project {
                referenced_project_name: self.name.clone(),
            })
        } else {
            self.type_library()
        }
    }
    pub fn type_library(&self) -> Result<sym::ProjectReference> {
        let libid = match &self.kind {
            ReferenceKind::Registered { libid } => libid,
            ReferenceKind::Control {
                original, extended, ..
            } => original.as_ref().unwrap_or(extended),
            ReferenceKind::Project { .. } => {
                return Err(Error::Unsupported(format!(
                    "referenced VBA project {} requires a project graph",
                    self.name
                )));
            }
        };
        let fields = libid.split('#').collect::<Vec<_>>();
        if fields.len() < 5 {
            return Err(invalid(format!("reference libid {libid}")));
        }
        let guid = fields[0]
            .get(3..)
            .filter(|s| s.starts_with('{') && s.ends_with('}'))
            .ok_or_else(|| invalid("reference GUID"))?;
        let (major, minor) = fields[1]
            .split_once('.')
            .ok_or_else(|| invalid("reference version"))?;
        Ok(sym::ProjectReference::TypeLibrary {
            name: if self.name.is_empty() {
                fields[4].into()
            } else {
                self.name.clone()
            },
            guid: Some(guid.trim_matches(['{', '}']).into()),
            version_major: Some(u16::from_str_radix(major, 16).map_err(invalid)?),
            version_minor: Some(u16::from_str_radix(minor, 16).map_err(invalid)?),
            lcid: Some(u32::from_str_radix(fields[2], 16).map_err(invalid)?),
            // The original path is retained in libid. Do not force a foreign
            // machine's Office install path on the host's typelib resolver.
            import_lib: None,
        })
    }
}

#[derive(Debug, Clone)]
pub struct Project {
    pub properties: ProjectProperties,
    pub references: Vec<Reference>,
    pub modules: Vec<Module>,
    /// PROJECT text, including host extenders and workspace preferences.
    pub project_stream: String,
    pub(crate) root: String,
    pub(crate) original: Vec<u8>,
}
impl Project {
    /// Reads a `vbaProject.bin` blob: the compound storage, the `dir` stream, the code page,
    /// the references and every module's decompressed source. Nothing is compiled or run.
    pub fn read_bin(bytes: &[u8]) -> Result<Self> {
        storage::read_project(bytes, &ReadLimits::default())
    }
    /// [`Project::read_bin`] under the caller's size limits.
    pub fn read_bin_with_limits(bytes: &[u8], limits: &ReadLimits) -> Result<Self> {
        storage::read_project(bytes, limits)
    }
    /// Sets the code (the body without `Attribute` lines) of the module `name`, adding a
    /// standard or class module of that name when the project has none. An existing module
    /// keeps its kind, stream and `Attribute` lines; a new one gets the attributes the VBE
    /// writes for its kind. [`Project::to_bin`] then writes the changed project.
    pub fn put_module(&mut self, name: &str, kind: ModuleKind, code: &str) -> Result<()> {
        let mut changed = self.clone();
        changed.put_module_unchecked(name, kind, code)?;
        changed.validate()?;
        *self = changed;
        Ok(())
    }
    fn put_module_unchecked(&mut self, name: &str, kind: ModuleKind, code: &str) -> Result<()> {
        if !valid_name(name) {
            return Err(invalid(format!("module name {name:?}")));
        }
        if code.lines().any(|line| {
            line.trim_start()
                .get(..10)
                .is_some_and(|p| p.eq_ignore_ascii_case("Attribute "))
        }) {
            return Err(invalid("module code must not carry Attribute lines"));
        }
        let body = code.replace("\r\n", "\n").replace('\n', "\r\n");
        if let Some(module) = self
            .modules
            .iter_mut()
            .find(|m| m.name.eq_ignore_ascii_case(name))
        {
            if module.kind != kind {
                return Err(Error::Unsupported(format!(
                    "module {} is a {:?} module, not {kind:?}",
                    module.name, module.kind
                )));
            }
            let attributes = module
                .source
                .lines()
                .filter(|line| {
                    line.trim_start()
                        .to_ascii_lowercase()
                        .starts_with("attribute ")
                })
                .map(|line| format!("{line}\r\n"))
                .collect::<String>();
            module.source = attributes + &body;
        } else {
            let attributes = match kind {
                ModuleKind::Standard => format!("Attribute VB_Name = \"{name}\"\r\n"),
                ModuleKind::Class => format!(
                    "Attribute VB_Name = \"{name}\"\r\n\
                     Attribute VB_Base = \"0{{FCFB3D2A-A0FA-1068-A738-08002B3371B5}}\"\r\n\
                     Attribute VB_GlobalNameSpace = False\r\n\
                     Attribute VB_Creatable = False\r\n\
                     Attribute VB_PredeclaredId = False\r\n\
                     Attribute VB_Exposed = False\r\n\
                     Attribute VB_TemplateDerived = False\r\n\
                     Attribute VB_Customizable = False\r\n"
                ),
                ModuleKind::Document | ModuleKind::Form => {
                    return Err(Error::Unsupported(format!(
                        "a new {kind:?} module needs its host document or designer"
                    )));
                }
            };
            self.modules.push(Module {
                name: name.into(),
                stream_name: name.into(),
                kind,
                description: String::new(),
                help_context: 0,
                read_only: false,
                private: false,
                source: attributes + &body,
                document_version: None,
                text_offset: 0,
            });
        }
        Ok(())
    }
    pub fn to_bin(&self) -> Result<Vec<u8>> {
        storage::write_project(self)
    }
    pub fn to_manifest(&self) -> Result<sym::SymbolProjectManifest> {
        self.validate()?;
        let mut modules = Vec::new();
        for module in &self.modules {
            let mut attrs = sym::ModuleAttributes::named(&module.name);
            attrs.vb_global_namespace = module.boolean_attribute("VB_GlobalNameSpace");
            attrs.vb_creatable = module.boolean_attribute("VB_Creatable");
            attrs.vb_predeclared_id = module.boolean_attribute("VB_PredeclaredId");
            attrs.vb_exposed = module.boolean_attribute("VB_Exposed");
            attrs.option_private_module = module
                .code()
                .lines()
                .any(|l| l.trim().eq_ignore_ascii_case("Option Private Module"));
            attrs.host_document_type = module.host_document_type()?;
            attrs.description = module
                .attribute("VB_Description")
                .and_then(|s| s.strip_prefix('"'))
                .and_then(|s| s.strip_suffix('"'))
                .map(|s| s.replace("\"\"", "\""));
            // Office .cls preambles say MultiUse even for private classes.
            // The VB_* attributes determine whether a class is exposed/creatable.
            if matches!(module.kind, ModuleKind::Class | ModuleKind::Document) {
                attrs.instancing = Some(if !attrs.vb_exposed {
                    sym::Instancing::Private
                } else if !attrs.vb_creatable {
                    sym::Instancing::PublicNotCreatable
                } else if attrs.vb_global_namespace {
                    sym::Instancing::GlobalMultiUse
                } else {
                    sym::Instancing::MultiUse
                });
            }
            modules.push(sym::ModuleUnit {
                module_name: module.name.clone(),
                module_kind: match module.kind {
                    ModuleKind::Standard => sym::ModuleKind::Procedural,
                    ModuleKind::Class => sym::ModuleKind::Class,
                    ModuleKind::Document => sym::ModuleKind::Document,
                    ModuleKind::Form => sym::ModuleKind::Form,
                },
                attributes: attrs,
                source: module.source.clone(),
            });
        }
        let mut references = Vec::new();
        // The VBA library is built into the compiler. The Office host library
        // is intrinsic and is absent from dir's explicit reference records.
        // VB_Base identifies the Excel document interfaces; version selection
        // belongs to the host's resolver, not to the saved performance cache.
        if modules.iter().any(|m| {
            m.attributes
                .host_document_type
                .as_deref()
                .is_some_and(|s| s.starts_with("Excel."))
        }) {
            references.push(sym::ProjectReference::TypeLibrary {
                name: "Excel".into(),
                guid: Some("00020813-0000-0000-C000-000000000046".into()),
                version_major: None,
                version_minor: None,
                lcid: None,
                import_lib: None,
            });
        }
        for reference in &self.references {
            references.push(reference.compiler_reference()?);
        }
        Ok(sym::SymbolProjectManifest {
            project_name: self.properties.name.clone(),
            project_kind: sym::ProjectKind::Host,
            modules,
            references,
            reference_projects: Vec::new(),
            conditional_constants: parse_constants(&self.properties.constants)?,
            conditional_compilation_target: Default::default(),
        })
    }
    pub(crate) fn validate(&self) -> Result<()> {
        let mut names = BTreeSet::new();
        let mut streams = BTreeSet::new();
        for module in &self.modules {
            if !valid_name(&module.name)
                || !valid_name(&module.stream_name)
                || !names.insert(module.name.to_lowercase())
                || !streams.insert(module.stream_name.to_lowercase())
            {
                return Err(invalid("invalid or duplicate module identity"));
            }
            let attribute_name = module
                .attribute("VB_Name")
                .and_then(|s| s.strip_prefix('"'))
                .and_then(|s| s.strip_suffix('"'));
            if attribute_name != Some(module.name.as_str()) {
                return Err(invalid(format!(
                    "module {} VB_Name differs from directory",
                    module.name
                )));
            }
        }
        Ok(())
    }
}
pub(crate) fn valid_name(name: &str) -> bool {
    !name.is_empty()
        && name != "."
        && name != ".."
        && !name.chars().any(|c| {
            c.is_control()
                || matches!(
                    c,
                    '/' | '\\' | ':' | '!' | '"' | '<' | '>' | '|' | '?' | '*'
                )
        })
}
fn parse_constants(text: &str) -> Result<BTreeMap<String, i32>> {
    let mut result = BTreeMap::new();
    for item in text.split(':').map(str::trim).filter(|s| !s.is_empty()) {
        let (key, value) = item
            .split_once('=')
            .ok_or_else(|| invalid("conditional constant"))?;
        let value = value.trim();
        let value = if value.eq_ignore_ascii_case("True") {
            -1
        } else if value.eq_ignore_ascii_case("False") {
            0
        } else if let Some(hex) = value
            .strip_prefix("&H")
            .or_else(|| value.strip_prefix("&h"))
        {
            u32::from_str_radix(hex, 16)
                .map(|n| n as i32)
                .map_err(invalid)?
        } else {
            value.parse().map_err(invalid)?
        };
        result.insert(key.trim().into(), value);
    }
    Ok(result)
}
