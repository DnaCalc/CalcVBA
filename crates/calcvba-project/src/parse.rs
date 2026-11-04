//! XML parser for `.basproj` files using `quick-xml`.

use quick_xml::events::Event;
use quick_xml::reader::Reader;

use crate::error::BasProjError;
use crate::model::*;

/// Parse a `.basproj` XML string into a `BasProj` intermediate representation.
pub fn parse_basproj_xml(xml: &str) -> Result<BasProj, BasProjError> {
    let mut reader = Reader::from_str(xml);

    let mut sdk = String::new();
    let mut properties = BasProjProperties::default();
    let mut modules = Vec::new();
    let mut project_references = Vec::new();
    let mut com_references = Vec::new();
    let mut native_references = Vec::new();
    let mut native_exports = Vec::new();
    let mut imports_to_process: Vec<String> = Vec::new();

    let mut in_project = false;
    let mut in_property_group = false;
    let mut in_item_group = false;
    let mut current_property: Option<String> = None;
    let mut current_item: Option<ItemContext> = None;
    let mut current_item_metadata_name: Option<String> = None;

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => {
                let tag = String::from_utf8_lossy(e.name().as_ref()).to_string();
                match tag.as_str() {
                    "Project" => {
                        in_project = true;
                        for attr in e.attributes().flatten() {
                            let key = String::from_utf8_lossy(attr.key.as_ref()).to_string();
                            if key == "Sdk" {
                                sdk = String::from_utf8_lossy(&attr.value).to_string();
                            }
                        }
                    }
                    "PropertyGroup" if in_project => {
                        in_property_group = true;
                    }
                    "ItemGroup" if in_project => {
                        in_item_group = true;
                    }
                    // Property elements inside PropertyGroup
                    _ if in_property_group => {
                        current_property = Some(tag);
                    }
                    // Metadata elements inside an item (must check BEFORE item tags,
                    // because e.g. <Module> can be both an item tag and a metadata
                    // element inside <NativeExport>)
                    _ if current_item.is_some() => {
                        let item = current_item.as_mut().expect("current item");
                        // Presence matters for refused source attributes and
                        // export options, including paired empty elements.
                        item.set_metadata(&tag, "");
                        current_item_metadata_name = Some(tag);
                    }
                    // Item elements inside ItemGroup
                    "Module" | "ClassModule" | "DocumentModule" | "FormModule"
                    | "ProjectReference" | "COMReference" | "NativeReference" | "NativeExport"
                        if in_item_group =>
                    {
                        if tag == "NativeExport" {
                            validate_native_export_attributes(e)?;
                        }
                        let include = extract_include_attr(e)?;
                        current_item = Some(ItemContext::new(&tag, include));
                    }
                    _ => {}
                }
            }
            Ok(Event::Empty(ref e)) => {
                let tag = String::from_utf8_lossy(e.name().as_ref()).to_string();
                match tag.as_str() {
                    _ if current_item.is_some() => {
                        current_item
                            .as_mut()
                            .expect("current item")
                            .set_metadata(&tag, "");
                    }
                    "Import" if in_project => {
                        // Collect import paths for later processing
                        for attr in e.attributes().flatten() {
                            let key = String::from_utf8_lossy(attr.key.as_ref()).to_string();
                            if key == "Project" {
                                imports_to_process
                                    .push(String::from_utf8_lossy(&attr.value).to_string());
                            }
                        }
                    }
                    // Self-closing item elements (e.g., <Module Include="..." />)
                    "Module" | "ClassModule" | "DocumentModule" | "FormModule"
                    | "ProjectReference" | "COMReference" | "NativeReference" | "NativeExport"
                        if in_item_group =>
                    {
                        if tag == "NativeExport" {
                            validate_native_export_attributes(e)?;
                        }
                        let include = extract_include_attr(e)?;
                        let ctx = ItemContext::new(&tag, include);
                        emit_item(
                            ctx,
                            &mut modules,
                            &mut project_references,
                            &mut com_references,
                            &mut native_references,
                            &mut native_exports,
                        )?;
                    }
                    _ => {}
                }
            }
            Ok(Event::Text(ref e)) => {
                let text = e.unescape().map_err(|err| {
                    BasProjError::XmlParse(format!("XML text unescape error: {err}"))
                })?;
                let text = text.trim().to_string();
                if !text.is_empty() {
                    if let Some(ref prop_name) = current_property {
                        apply_property(&mut properties, prop_name, &text)?;
                    }
                    if let Some(ref meta_name) = current_item_metadata_name
                        && let Some(ref mut ctx) = current_item
                    {
                        ctx.set_metadata(meta_name, &text);
                    }
                }
            }
            Ok(Event::End(ref e)) => {
                let tag = String::from_utf8_lossy(e.name().as_ref()).to_string();

                // If we're closing a metadata element inside an item, just
                // clear the metadata name — don't emit the item yet.
                if current_item_metadata_name.as_deref() == Some(tag.as_str()) {
                    current_item_metadata_name = None;
                } else if current_property.as_deref() == Some(tag.as_str()) {
                    current_property = None;
                } else {
                    match tag.as_str() {
                        "Project" => {
                            in_project = false;
                        }
                        "PropertyGroup" => {
                            in_property_group = false;
                        }
                        "ItemGroup" => {
                            in_item_group = false;
                        }
                        _ if current_item.is_some() => {
                            // Closing tag matches the item element
                            if let Some(ref ctx) = current_item
                                && ctx.tag == tag
                            {
                                let Some(ctx) = current_item.take() else {
                                    return Err(BasProjError::XmlParse(format!(
                                        "unexpected closing item tag </{tag}>"
                                    )));
                                };
                                emit_item(
                                    ctx,
                                    &mut modules,
                                    &mut project_references,
                                    &mut com_references,
                                    &mut native_references,
                                    &mut native_exports,
                                )?;
                            }
                        }
                        _ => {}
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(BasProjError::XmlParse(format!("XML parse error: {e}")));
            }
            _ => {}
        }
    }

    if !in_project && sdk.is_empty() {
        // Allow import files without Sdk attribute (they have bare <Project> root)
    }

    Ok(BasProj {
        sdk,
        properties,
        modules,
        project_references,
        com_references,
        native_references,
        native_exports,
    })
}

/// Merge items from an imported `BasProj` into a parent `BasProj`.
pub fn merge_import(parent: &mut BasProj, imported: BasProj) {
    parent.modules.extend(imported.modules);
    parent
        .project_references
        .extend(imported.project_references);
    parent.com_references.extend(imported.com_references);
    parent.native_references.extend(imported.native_references);
    parent.native_exports.extend(imported.native_exports);
    // Properties from imported files do NOT override parent properties
    // (MSBuild convention: property imports use .props files which are imported
    // before the project's own PropertyGroups, but we keep it simple here).
}

// --- Internal helpers ---

/// Temporary context for accumulating item metadata during parsing.
#[derive(Debug)]
struct ItemContext {
    tag: String,
    include: String,
    metadata: std::collections::HashMap<String, String>,
}

impl ItemContext {
    fn new(tag: &str, include: String) -> Self {
        Self {
            tag: tag.to_string(),
            include,
            metadata: std::collections::HashMap::new(),
        }
    }

    fn set_metadata(&mut self, key: &str, value: &str) {
        self.metadata.insert(key.to_string(), value.to_string());
    }

    fn meta_str(&self, key: &str) -> Option<String> {
        self.metadata.get(key).cloned()
    }

    fn meta_u16(&self, key: &str) -> Option<u16> {
        self.metadata.get(key).and_then(|v| v.parse().ok())
    }

    fn meta_u32(&self, key: &str) -> Option<u32> {
        self.metadata.get(key).and_then(|v| v.parse().ok())
    }
}

fn extract_include_attr(e: &quick_xml::events::BytesStart<'_>) -> Result<String, BasProjError> {
    for attr in e.attributes().flatten() {
        let key = String::from_utf8_lossy(attr.key.as_ref()).to_string();
        if key == "Include" {
            return attr
                .unescape_value()
                .map(|value| value.into_owned())
                .map_err(|error| {
                    BasProjError::XmlParse(format!("XML Include attribute unescape error: {error}"))
                });
        }
    }
    let tag = String::from_utf8_lossy(e.name().as_ref()).to_string();
    Err(BasProjError::MissingAttribute {
        element: tag,
        attribute: "Include".to_string(),
    })
}

fn validate_native_export_attributes(
    e: &quick_xml::events::BytesStart<'_>,
) -> Result<(), BasProjError> {
    for attribute in e.attributes() {
        let attribute = attribute.map_err(|error| BasProjError::XmlParse(error.to_string()))?;
        if attribute.key.as_ref() != b"Include" {
            return Err(BasProjError::XmlParse(format!(
                "unsupported NativeExport attribute {}",
                String::from_utf8_lossy(attribute.key.as_ref())
            )));
        }
    }
    Ok(())
}

fn apply_property(
    props: &mut BasProjProperties,
    name: &str,
    value: &str,
) -> Result<(), BasProjError> {
    match name {
        "OutputType" => {
            props.output_type = match value {
                "HostModule" => Some(OutputType::HostModule),
                "Library" => Some(OutputType::Library),
                "Exe" => Some(OutputType::Exe),
                "NativeLibrary" => Some(OutputType::NativeLibrary),
                "ComServer" => Some(OutputType::ComServer),
                "ComExe" => Some(OutputType::ComExe),
                _ => return Err(BasProjError::InvalidOutputType(value.to_string())),
            };
        }
        "BuildTarget" => {
            props.build_target = match value {
                "Bundle" => Some(BuildTarget::Bundle),
                "WrapperExe" => Some(BuildTarget::WrapperExe),
                "WrapperLibrary" => Some(BuildTarget::WrapperLibrary),
                "WrappedComServer" | "WrapperComServer" => Some(BuildTarget::WrappedComServer),
                _ => return Err(BasProjError::InvalidBuildTarget(value.to_string())),
            };
        }
        "ProjectName" => props.project_name = Some(value.to_string()),
        "EntryPoint" => props.entry_point = Some(value.to_string()),
        "RuntimeFlavor" => {
            props.runtime_flavor = match value {
                "Lite" => Some(RuntimeFlavor::Lite),
                "Jit" => Some(RuntimeFlavor::Jit),
                _ => return Err(BasProjError::InvalidRuntimeFlavor(value.to_string())),
            };
        }
        "DefaultRuntimeProfile" => props.default_runtime_profile = Some(value.to_string()),
        "DefaultRootObject" => props.default_root_object = Some(value.to_string()),
        "DefineConstants" => props.define_constants = Some(value.to_string()),
        _ => {} // Unknown properties are silently ignored
    }
    Ok(())
}

fn emit_item(
    ctx: ItemContext,
    modules: &mut Vec<BasProjModule>,
    project_refs: &mut Vec<BasProjProjectReference>,
    com_refs: &mut Vec<BasProjComReference>,
    native_refs: &mut Vec<BasProjNativeReference>,
    native_exports: &mut Vec<BasProjNativeExport>,
) -> Result<(), BasProjError> {
    if matches!(
        ctx.tag.as_str(),
        "Module" | "ClassModule" | "DocumentModule" | "FormModule"
    ) {
        for key in ctx.metadata.keys() {
            if key.to_ascii_lowercase().starts_with("vb") || key == "Description" {
                return Err(BasProjError::XmlParse(format!(
                    "{key} belongs in the source file's Attribute lines, not .basproj metadata"
                )));
            }
        }
    }
    match ctx.tag.as_str() {
        "Module" => {
            modules.push(BasProjModule {
                kind: BasProjModuleKind::Module,
                include: ctx.include,
                host_document_type: None,
                instancing: None,
                prog_id: None,
            });
        }
        "ClassModule" | "FormModule" => {
            let instancing = ctx
                .meta_str("Instancing")
                .map(|s| match s.as_str() {
                    "Private" | "1" => Ok(Instancing::Private),
                    "PublicNotCreatable" | "2" => Ok(Instancing::PublicNotCreatable),
                    "SingleUse" | "3" => Ok(Instancing::SingleUse),
                    "GlobalSingleUse" | "4" => Ok(Instancing::GlobalSingleUse),
                    "MultiUse" | "5" => Ok(Instancing::MultiUse),
                    "GlobalMultiUse" | "6" => Ok(Instancing::GlobalMultiUse),
                    _ => Err(BasProjError::XmlParse(format!("invalid Instancing {s}"))),
                })
                .transpose()?;
            let prog_id = ctx.meta_str("ProgId");
            modules.push(BasProjModule {
                kind: if ctx.tag == "FormModule" {
                    BasProjModuleKind::FormModule
                } else {
                    BasProjModuleKind::ClassModule
                },
                include: ctx.include,
                host_document_type: None,
                instancing,
                prog_id,
            });
        }
        "DocumentModule" => {
            let doc_type = ctx.meta_str("HostDocumentType");
            modules.push(BasProjModule {
                kind: BasProjModuleKind::DocumentModule,
                include: ctx.include,
                host_document_type: doc_type,
                instancing: None,
                prog_id: None,
            });
        }
        "ProjectReference" => {
            let kind = match ctx.meta_str("Kind").as_deref() {
                Some("HostInjected") => BasProjProjectReferenceKind::HostInjected,
                _ => BasProjProjectReferenceKind::Project,
            };
            project_refs.push(BasProjProjectReference {
                include: ctx.include,
                kind,
            });
        }
        "COMReference" => {
            let guid = ctx.meta_str("Guid");
            let version_major = ctx.meta_u16("VersionMajor");
            let version_minor = ctx.meta_u16("VersionMinor");
            let lcid = ctx.meta_u32("Lcid");
            let import_lib = ctx.meta_str("ImportLib");
            com_refs.push(BasProjComReference {
                include: ctx.include,
                guid,
                version_major,
                version_minor,
                lcid,
                import_lib,
            });
        }
        "NativeReference" => {
            let path = ctx.meta_str("Path");
            native_refs.push(BasProjNativeReference {
                include: ctx.include,
                path,
            });
        }
        "NativeExport" => {
            for key in ctx.metadata.keys() {
                if !matches!(
                    key.as_str(),
                    "Module"
                        | "Procedure"
                        | "CallingConvention"
                        | "Ordinal"
                        | "StringEncoding"
                        | "ErrorPolicy"
                ) {
                    return Err(BasProjError::XmlParse(format!(
                        "unsupported NativeExport metadata {key}"
                    )));
                }
            }
            let cc = match ctx.meta_str("CallingConvention").as_deref() {
                None => None,
                Some("Stdcall") => Some(CallingConvention::Stdcall),
                Some("Cdecl") => Some(CallingConvention::Cdecl),
                Some(other) => {
                    return Err(BasProjError::XmlParse(format!(
                        "invalid native calling convention {other}"
                    )));
                }
            };
            let module = ctx.meta_str("Module");
            let procedure = ctx.meta_str("Procedure");
            let string_encoding = match ctx.meta_str("StringEncoding").as_deref() {
                None | Some("Bstr") => StringEncoding::Bstr,
                Some("Utf16") => StringEncoding::Utf16,
                Some("Ansi") => StringEncoding::Ansi,
                Some(other) => {
                    return Err(BasProjError::XmlParse(format!(
                        "invalid native StringEncoding {other}"
                    )));
                }
            };
            let error_policy = match ctx.meta_str("ErrorPolicy").as_deref() {
                None | Some("Zero") => ErrorPolicy::Zero,
                Some("HResult") => ErrorPolicy::HResult,
                Some(other) => {
                    return Err(BasProjError::XmlParse(format!(
                        "invalid native ErrorPolicy {other}"
                    )));
                }
            };
            let ordinal = ctx
                .meta_str("Ordinal")
                .map(|s| {
                    s.parse::<u16>().map_err(|_| {
                        BasProjError::XmlParse(format!("invalid native export ordinal {s}"))
                    })
                })
                .transpose()?;
            native_exports.push(BasProjNativeExport {
                include: ctx.include,
                module,
                procedure,
                calling_convention: cc,
                ordinal,
                string_encoding,
                error_policy,
            });
        }
        _ => {}
    }
    Ok(())
}
