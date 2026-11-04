//! VBE source files in UTF-8, with MSBuild-shaped packaging and references.
use crate::{
    Module, ModuleKind, Project, ProjectProperties, Reference, ReferenceKind, Result, invalid,
};
use std::{
    collections::BTreeMap,
    io::Cursor,
    path::{Path, PathBuf},
};

#[derive(Default)]
struct Item {
    tag: String,
    include: String,
    values: BTreeMap<String, String>,
}
#[derive(Default)]
struct Xml {
    properties: BTreeMap<String, String>,
    items: Vec<Item>,
}
fn parse(xml: &str) -> Result<Xml> {
    let mut reader = quick_xml::Reader::from_str(xml);
    let mut result = Xml::default();
    let mut stack = Vec::<String>::new();
    let mut item: Option<Item> = None;
    let mut text = String::new();
    loop {
        let event = reader.read_event().map_err(invalid)?;
        let empty = matches!(&event, quick_xml::events::Event::Empty(_));
        match event {
            quick_xml::events::Event::Start(e) | quick_xml::events::Event::Empty(e) => {
                let tag = String::from_utf8(e.local_name().as_ref().to_vec()).map_err(invalid)?;
                if stack.last().is_some_and(|s| s == "ItemGroup") {
                    let mut include = String::new();
                    for attr in e.attributes() {
                        let attr = attr.map_err(invalid)?;
                        if attr.key.as_ref() == b"Include" {
                            include = attr.unescape_value().map_err(invalid)?.into_owned();
                        }
                    }
                    item = Some(Item {
                        tag: tag.clone(),
                        include,
                        ..Default::default()
                    });
                }
                text.clear();
                if !empty {
                    stack.push(tag);
                } else if let Some(value) = item.take() {
                    result.items.push(value);
                }
            }
            quick_xml::events::Event::Text(e) => text.push_str(&e.unescape().map_err(invalid)?),
            quick_xml::events::Event::CData(e) => {
                text.push_str(std::str::from_utf8(&e).map_err(invalid)?)
            }
            quick_xml::events::Event::End(e) => {
                let tag = String::from_utf8(e.local_name().as_ref().to_vec()).map_err(invalid)?;
                if stack.pop().as_deref() != Some(tag.as_str()) {
                    return Err(invalid("project XML nesting"));
                }
                if stack.last().is_some_and(|s| s == "PropertyGroup") {
                    result.properties.insert(tag, text.clone());
                } else if item.as_ref().is_some_and(|i| i.tag == tag) {
                    result.items.push(item.take().unwrap());
                } else if let Some(item) = &mut item {
                    item.values.insert(tag, text.clone());
                }
                text.clear();
            }
            quick_xml::events::Event::Eof => break,
            _ => {}
        }
    }
    if !stack.is_empty() || item.is_some() {
        return Err(invalid("incomplete project XML"));
    }
    Ok(result)
}
fn escaped(text: &str) -> String {
    quick_xml::escape::escape(text).into_owned()
}
fn constant_separators(text: &str, from: char, to: char) -> String {
    let mut result = String::new();
    let mut quoted = false;
    let mut chars = text.chars().peekable();
    while let Some(c) = chars.next() {
        if c == '"' {
            result.push(c);
            if quoted && chars.peek() == Some(&'"') {
                result.push(chars.next().unwrap());
            } else {
                quoted = !quoted;
            }
        } else {
            result.push(if !quoted && c == from { to } else { c });
        }
    }
    result
}
fn prop(xml: &mut String, key: &str, value: impl std::fmt::Display) {
    xml.push_str(&format!(
        "    <{key}>{}</{key}>\n",
        escaped(&value.to_string())
    ));
}
fn number<T: std::str::FromStr>(props: &BTreeMap<String, String>, key: &str) -> Result<T> {
    props
        .get(key)
        .ok_or_else(|| invalid(format!("missing project property {key}")))?
        .trim()
        .parse()
        .map_err(|_| invalid(format!("invalid project property {key}")))
}
fn value(props: &BTreeMap<String, String>, key: &str) -> String {
    props.get(key).cloned().unwrap_or_default()
}
fn insert_attribute(source: &mut String, attribute: &str) -> Result<()> {
    let name = source
        .find("Attribute VB_Name")
        .ok_or_else(|| invalid("VB_Name"))?;
    let end = source[name..]
        .find('\n')
        .map(|n| name + n + 1)
        .unwrap_or(source.len());
    source.insert_str(end, &format!("{attribute}\r\n"));
    Ok(())
}
fn inside(root: &Path, include: &str) -> Result<PathBuf> {
    let path = root.join(include.replace('\\', "/")).canonicalize()?;
    if !path.starts_with(root.canonicalize()?) {
        return Err(invalid("source path escapes project folder"));
    }
    Ok(path)
}

impl Project {
    /// Export to a new directory. Source Attributes are never duplicated in XML.
    pub fn export_folder(&self, destination: impl AsRef<Path>) -> Result<PathBuf> {
        self.validate()?;
        crate::storage::check_protection(&self.project_stream)?;
        if !crate::valid_name(&self.properties.name) {
            return Err(invalid("project name cannot be used as a filename"));
        }
        let destination = destination.as_ref();
        if destination.exists() {
            return Err(invalid("project export directory already exists"));
        }
        let mut files = Vec::<(String, Vec<u8>)>::new();
        let mut xml = "<Project Sdk=\"CalcVBA.Sdk\">\n  <PropertyGroup>\n".to_owned();
        let p = &self.properties;
        prop(&mut xml, "OutputType", "HostModule");
        prop(&mut xml, "ProjectName", &p.name);
        prop(&mut xml, "Codepage", p.codepage);
        prop(&mut xml, "SysKind", p.sys_kind);
        prop(&mut xml, "Lcid", p.lcid);
        prop(&mut xml, "LcidInvoke", p.lcid_invoke);
        prop(&mut xml, "ProjectDescription", &p.description);
        prop(&mut xml, "HelpFile", &p.help_file);
        prop(&mut xml, "HelpContextId", p.help_context);
        prop(&mut xml, "LibFlags", p.lib_flags);
        prop(&mut xml, "VersionMajor", p.version_major);
        prop(&mut xml, "VersionMinor", p.version_minor);
        if let Some(v) = p.compat_version {
            prop(&mut xml, "CompatibilityVersion", v);
        }
        prop(
            &mut xml,
            "DefineConstants",
            constant_separators(&p.constants, ':', ';'),
        );
        xml.push_str("  </PropertyGroup>\n  <ItemGroup>\n");
        for module in &self.modules {
            let (tag, extension) = match module.kind {
                ModuleKind::Standard => ("Module", "bas"),
                ModuleKind::Class => ("ClassModule", "cls"),
                ModuleKind::Document => ("DocumentModule", "cls"),
                ModuleKind::Form => ("FormModule", "frm"),
            };
            let file = format!("{}.{extension}", module.name);
            let mut source = if module.kind == ModuleKind::Form {
                let (source, frx) = crate::storage::export_form(self, module)?;
                files.push((format!("{}.frx", module.name), frx));
                source
            } else if matches!(module.kind, ModuleKind::Class | ModuleKind::Document) {
                format!(
                    "VERSION 1.0 CLASS\r\nBEGIN\r\n  MultiUse = -1  'True\r\nEND\r\n{}",
                    module.source
                )
            } else {
                module.source.clone()
            };
            if !module.description.is_empty() && module.attribute("VB_Description").is_none() {
                insert_attribute(
                    &mut source,
                    &format!(
                        "Attribute VB_Description = \"{}\"",
                        module.description.replace('"', "\"\"")
                    ),
                )?;
            }
            if module.help_context != 0 && module.attribute("VB_HelpID").is_none() {
                insert_attribute(
                    &mut source,
                    &format!("Attribute VB_HelpID = {}", module.help_context),
                )?;
            }
            files.push((file.clone(), source.into_bytes()));
            xml.push_str(&format!("    <{tag} Include=\"{}\">\n", escaped(&file)));
            prop(&mut xml, "StreamName", &module.stream_name);
            if let Some(version) = &module.document_version {
                prop(&mut xml, "DocumentVersion", version);
            }
            prop(&mut xml, "ReadOnly", module.read_only);
            prop(&mut xml, "Private", module.private);
            xml.push_str(&format!("    </{tag}>\n"));
        }
        // Intrinsic host metadata is a reference for the folder compiler. It is
        // not written into dir as an additional explicit Office reference.
        if self.modules.iter().any(|m| {
            m.kind == ModuleKind::Document
                && m.host_document_type()
                    .ok()
                    .flatten()
                    .is_some_and(|s| s.starts_with("Excel."))
        }) {
            xml.push_str("    <COMReference Include=\"Excel\">\n");
            prop(&mut xml, "Guid", "00020813-0000-0000-C000-000000000046");
            prop(&mut xml, "Intrinsic", true);
            xml.push_str("    </COMReference>\n");
        }
        for reference in &self.references {
            let tag = if matches!(reference.kind, ReferenceKind::Project { .. }) {
                "VbaProjectReference"
            } else {
                "COMReference"
            };
            xml.push_str(&format!(
                "    <{tag} Include=\"{}\">\n",
                escaped(&reference.name)
            ));
            if let Ok(calcvba_symbol::manifest::ProjectReference::TypeLibrary {
                guid,
                version_major,
                version_minor,
                lcid,
                ..
            }) = reference.type_library()
            {
                prop(&mut xml, "Guid", guid.unwrap_or_default());
                prop(&mut xml, "VersionMajor", version_major.unwrap_or_default());
                prop(&mut xml, "VersionMinor", version_minor.unwrap_or_default());
                prop(&mut xml, "Lcid", lcid.unwrap_or_default());
            }
            match &reference.kind {
                ReferenceKind::Registered { libid } => prop(&mut xml, "Libid", libid),
                ReferenceKind::Project {
                    absolute,
                    relative,
                    major,
                    minor,
                } => {
                    prop(&mut xml, "AbsoluteLibid", absolute);
                    prop(&mut xml, "RelativeLibid", relative);
                    prop(&mut xml, "VersionMajor", major);
                    prop(&mut xml, "VersionMinor", minor);
                }
                ReferenceKind::Control {
                    original,
                    twiddled,
                    extended_name,
                    extended,
                    guid,
                    cookie,
                } => {
                    prop(&mut xml, "ReferenceKind", "Control");
                    if let Some(v) = original {
                        prop(&mut xml, "OriginalLibid", v);
                    }
                    prop(&mut xml, "TwiddledLibid", twiddled);
                    if let Some(v) = extended_name {
                        prop(&mut xml, "ExtendedName", v);
                    }
                    prop(&mut xml, "ExtendedLibid", extended);
                    prop(&mut xml, "Cookie", cookie);
                    prop(
                        &mut xml,
                        "ControlGuid",
                        guid.iter().map(|b| format!("{b:02x}")).collect::<String>(),
                    );
                }
            }
            xml.push_str(&format!("    </{tag}>\n"));
        }
        let mut section = String::new();
        for line in self.project_stream.lines() {
            if line.starts_with('[') && line.ends_with(']') {
                section = line.into();
                continue;
            }
            let Some((key, val)) = line.split_once('=') else {
                continue;
            };
            if section.is_empty()
                && matches!(
                    key,
                    "Module"
                        | "Class"
                        | "Document"
                        | "BaseClass"
                        | "Name"
                        | "Description"
                        | "HelpContextID"
                        | "HelpFile"
                )
            {
                continue;
            }
            xml.push_str(&format!(
                "    <OfficeProjectProperty Include=\"{}\">\n",
                escaped(&format!("{section}|{key}"))
            ));
            prop(&mut xml, "Value", val);
            xml.push_str("    </OfficeProjectProperty>\n");
        }
        xml.push_str("  </ItemGroup>\n</Project>\n");
        if xml.chars().any(
            |c| matches!(c,'\0'..='\u{8}'|'\u{b}'|'\u{c}'|'\u{e}'..='\u{1f}'|'\u{fffe}'|'\u{ffff}'),
        ) {
            return Err(invalid(
                "project metadata contains a character forbidden in MSBuild XML",
            ));
        }
        let name = format!("{}.basproj", p.name);
        // Check the generated file through the ordinary project parser too.
        calcvba_project::parse_basproj_xml(&xml).map_err(invalid)?;
        files.push((name.clone(), xml.into_bytes()));
        std::fs::create_dir(destination)?;
        for (name, bytes) in files {
            std::fs::write(destination.join(name), bytes)?;
        }
        Ok(destination.join(name))
    }

    /// Read an exported folder project. The UTF-8 source headers own VB_* values.
    pub fn from_folder(path: impl AsRef<Path>) -> Result<Self> {
        let path = path.as_ref();
        let root = path.parent().unwrap_or(Path::new("."));
        let xml = std::fs::read_to_string(path)?;
        let bas = calcvba_project::parse_basproj_xml(&xml).map_err(invalid)?;
        if bas.sdk.split('/').next() != Some("CalcVBA.Sdk") {
            return Err(invalid("Office source folders require CalcVBA.Sdk"));
        }
        let xml = parse(&xml)?;
        let p = &xml.properties;
        let properties = ProjectProperties {
            name: value(p, "ProjectName"),
            codepage: number(p, "Codepage")?,
            sys_kind: number(p, "SysKind")?,
            lcid: number(p, "Lcid")?,
            lcid_invoke: number(p, "LcidInvoke")?,
            description: value(p, "ProjectDescription"),
            help_file: value(p, "HelpFile"),
            help_context: number(p, "HelpContextId")?,
            lib_flags: number(p, "LibFlags")?,
            version_major: number(p, "VersionMajor")?,
            version_minor: number(p, "VersionMinor")?,
            compat_version: p
                .get("CompatibilityVersion")
                .map(|_| number(p, "CompatibilityVersion"))
                .transpose()?,
            constants: constant_separators(&value(p, "DefineConstants"), ';', ':'),
        };
        let mut modules = Vec::new();
        let mut references = Vec::new();
        let mut extra = Vec::new();
        let mut forms = Vec::new();
        for item in xml.items {
            let v = &item.values;
            match item.tag.as_str() {
                "Module" | "ClassModule" | "DocumentModule" | "FormModule" => {
                    let file = inside(root, &item.include)?;
                    let exported = std::fs::read_to_string(&file)?;
                    let source = if exported.trim_start().starts_with("VERSION ") {
                        let at = exported
                            .find("Attribute VB_Name")
                            .ok_or_else(|| invalid("VB_Name"))?;
                        exported[at..].to_owned()
                    } else {
                        exported.clone()
                    };
                    let kind = match item.tag.as_str() {
                        "Module" => ModuleKind::Standard,
                        "ClassModule" => ModuleKind::Class,
                        "DocumentModule" => ModuleKind::Document,
                        _ => ModuleKind::Form,
                    };
                    let mut module = Module {
                        name: String::new(),
                        stream_name: value(v, "StreamName"),
                        kind,
                        description: String::new(),
                        help_context: 0,
                        private: value(v, "Private") == "true",
                        read_only: value(v, "ReadOnly") == "true",
                        source,
                        text_offset: 0,
                        document_version: v.get("DocumentVersion").cloned(),
                    };
                    module.name = module
                        .attribute("VB_Name")
                        .ok_or_else(|| invalid("VB_Name"))?
                        .trim_matches('"')
                        .into();
                    if let Some(description) = module.attribute("VB_Description") {
                        module.description = description
                            .strip_prefix('"')
                            .and_then(|value| value.strip_suffix('"'))
                            .ok_or_else(|| invalid("VB_Description string"))?
                            .replace("\"\"", "\"");
                    }
                    if let Some(help) = module.attribute("VB_HelpID") {
                        module.help_context = help.parse().map_err(invalid)?;
                    }
                    if module.stream_name.is_empty() {
                        module.stream_name = module.name.clone();
                    }
                    if kind == ModuleKind::Form {
                        forms.push((
                            module.name.clone(),
                            exported,
                            std::fs::read(inside(
                                root,
                                &Path::new(&item.include)
                                    .with_extension("frx")
                                    .to_string_lossy(),
                            )?)?,
                        ));
                    }
                    modules.push(module);
                }
                "COMReference" if value(v, "Intrinsic") != "true" => {
                    let kind = if value(v, "ReferenceKind") == "Control" {
                        let hex = value(v, "ControlGuid");
                        if hex.len() != 32 || !hex.bytes().all(|b| b.is_ascii_hexdigit()) {
                            return Err(invalid("control reference GUID"));
                        }
                        let mut guid = [0; 16];
                        for (i, byte) in guid.iter_mut().enumerate() {
                            *byte =
                                u8::from_str_radix(&hex[i * 2..i * 2 + 2], 16).map_err(invalid)?;
                        }
                        ReferenceKind::Control {
                            original: v.get("OriginalLibid").cloned(),
                            twiddled: value(v, "TwiddledLibid"),
                            extended_name: v.get("ExtendedName").cloned(),
                            extended: value(v, "ExtendedLibid"),
                            guid,
                            cookie: number(v, "Cookie")?,
                        }
                    } else {
                        ReferenceKind::Registered {
                            libid: value(v, "Libid"),
                        }
                    };
                    references.push(Reference {
                        name: item.include,
                        kind,
                    });
                }
                "VbaProjectReference" => references.push(Reference {
                    name: item.include,
                    kind: ReferenceKind::Project {
                        absolute: value(v, "AbsoluteLibid"),
                        relative: value(v, "RelativeLibid"),
                        major: number(v, "VersionMajor")?,
                        minor: number(v, "VersionMinor")?,
                    },
                }),
                "OfficeProjectProperty" => extra.push((item.include, value(v, "Value"))),
                _ => {}
            }
        }
        let mut project_stream = String::new();
        let mut section = String::new();
        // Root properties precede named sections. Module declarations are
        // regenerated by the writer from the folder's explicit item list.
        for (key, val) in &extra {
            if let Some((s, k)) = key.split_once('|')
                && s.is_empty()
            {
                project_stream.push_str(&format!("{k}={val}\r\n"));
            }
        }
        project_stream.push_str(&format!(
            "Name=\"{}\"\r\nDescription=\"{}\"\r\nHelpContextID=\"{}\"\r\nHelpFile=\"{}\"\r\n",
            properties.name, properties.description, properties.help_context, properties.help_file
        ));
        for (key, val) in extra {
            if let Some((s, k)) = key.split_once('|')
                && !s.is_empty()
            {
                if section != s {
                    section = s.into();
                    project_stream.push_str(&format!("\r\n{s}\r\n"));
                }
                project_stream.push_str(&format!("{k}={val}\r\n"));
            }
        }
        crate::storage::check_protection(&project_stream)?;
        let mut cfb =
            cfb::CompoundFile::create_with_version(cfb::Version::V3, Cursor::new(Vec::new()))?;
        cfb.create_storage("/VBA")?;
        for (name, frm, frx) in forms {
            crate::storage::import_form(&mut cfb, &name, &frm, &frx, properties.codepage)?;
        }
        let project = Self {
            properties,
            references,
            modules,
            project_stream,
            root: "/".into(),
            original: cfb.into_inner().into_inner(),
        };
        project.validate()?;
        Ok(project)
    }
}
#[cfg(test)]
mod tests {
    use super::constant_separators;
    #[test]
    fn conditional_string_literals_keep_their_punctuation_and_escaped_quotes() {
        let source = "A = 1 : Text = \"a:b;c\"\"d\" : B = 2";
        let xml = constant_separators(source, ':', ';');
        assert_eq!(xml, "A = 1 ; Text = \"a:b;c\"\"d\" ; B = 2");
        assert_eq!(constant_separators(&xml, ';', ':'), source);
    }
}
