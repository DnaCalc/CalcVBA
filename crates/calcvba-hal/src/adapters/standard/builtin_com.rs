use calcvba_com::call::ComFault;
use calcvba_com::{DynamicCallArg, DynamicCallKind, DynamicCallRequest, DynamicMemberSelector};
use calcvba_runtime::{
    ObjectRef, ObjectRefExt, RuntimeCallResult, VarType, Variant, coerce::coerce_to,
    variant_to_vba_string,
};
use std::cell::RefCell;
use std::rc::Rc;

use super::StandardHostServices;
use crate::{HalError, HalResult, model::CapabilityId};

pub(super) enum BuiltinObject {
    FileSystem,
    Document(RefCell<Document>),
    Element { name: String, text: String },
}

pub(super) struct Document {
    asynchronous: bool,
    root: Variant,
}

impl BuiltinObject {
    fn name(&self) -> &'static str {
        match self {
            Self::FileSystem => "FileSystemObject",
            Self::Document(_) => "DOMDocument",
            Self::Element { .. } => "IXMLDOMElement",
        }
    }

    pub(super) fn create(host: &StandardHostServices, prog_id: &str) -> HalResult<Option<Variant>> {
        use calcvba_com::platform::portable::BuiltinProgId;
        let object = match BuiltinProgId::resolve(prog_id) {
            Some(BuiltinProgId::FileSystemObject) if host.native_fs_enabled() => Self::FileSystem,
            Some(BuiltinProgId::XmlDocument) => Self::Document(RefCell::new(Document {
                asynchronous: true,
                root: Variant::nothing(),
            })),
            _ => return Ok(None),
        };
        object.retain(host).map(Some)
    }

    fn retain(self, host: &StandardHostServices) -> HalResult<Variant> {
        let identity = host
            .borrow_projection(CapabilityId::ComActivationDispatch, "create_object")?
            .allocate_handle();
        Ok(Variant::from_object_ref(
            ObjectRef::from_compat_identity_with_local_state(identity, Rc::new(self)),
        ))
    }

    pub(super) fn invoke(
        host: &StandardHostServices,
        request: &DynamicCallRequest,
    ) -> HalResult<Option<RuntimeCallResult>> {
        let Some(object) = request.object.local_state::<Self>() else {
            return Ok(None);
        };
        let member = match &request.member {
            DynamicMemberSelector::Name(name) | DynamicMemberSelector::TokenNamed { name, .. } => {
                name.to_ascii_lowercase()
            }
            _ => return Err(fault(host, 438)),
        };
        object
            .call(host, &member, request.call_kind_hint, &request.args)
            .map(RuntimeCallResult::value)
            .map(Some)
    }

    pub(super) fn type_name(object: &ObjectRef) -> Option<&'static str> {
        object.local_state::<Self>().map(|state| state.name())
    }

    fn call(
        &self,
        host: &StandardHostServices,
        member: &str,
        kind: Option<DynamicCallKind>,
        args: &[DynamicCallArg],
    ) -> HalResult<Variant> {
        use DynamicCallKind::{PropertyLet, PropertySet};
        let put = matches!(kind, Some(PropertyLet | PropertySet));
        match self {
            Self::FileSystem if !put && matches!(member, "fileexists" | "folderexists") => {
                let file = member == "fileexists";
                let name = if file { "filespec" } else { "folderspec" };
                let value = argument(host, args, name)?;
                let path = variant_to_vba_string(value).map_err(|_| fault(host, 13))?;
                let found = host
                    .map_path(
                        &path.as_str(),
                        CapabilityId::FileSystemIo,
                        "file_exists",
                        false,
                    )
                    .ok()
                    .and_then(|path| std::fs::metadata(path).ok())
                    .is_some_and(|entry| {
                        if file {
                            entry.is_file()
                        } else {
                            entry.is_dir()
                        }
                    });
                Ok(Variant::from_bool(found))
            }
            Self::Document(doc) if member == "async" && kind == Some(PropertyLet) => {
                let value = argument(host, args, "")?;
                let value = coerce_to(value, VarType::Boolean).map_err(|_| fault(host, 13))?;
                doc.borrow_mut().asynchronous = value.as_bool().unwrap_or(false);
                Ok(Variant::empty())
            }
            Self::Document(doc) if member == "async" && !put && args.is_empty() => {
                Ok(Variant::from_bool(doc.borrow().asynchronous))
            }
            Self::Document(doc) if member == "loadxml" && !put => {
                let xml = variant_to_vba_string(argument(host, args, "bstrXML")?)
                    .map_err(|_| fault(host, 13))?;
                let element = String::from_utf16(xml.payload_units())
                    .ok()
                    .and_then(|text| document_element(&text));
                let loaded = element.is_some();
                doc.borrow_mut().root = match element {
                    Some(root) => root.retain(host)?,
                    None => Variant::nothing(),
                };
                Ok(Variant::from_bool(loaded))
            }
            Self::Document(doc) if member == "documentelement" && !put && args.is_empty() => {
                Ok(doc.borrow().root.clone())
            }
            Self::Element { name, .. } if member == "nodename" && !put && args.is_empty() => {
                Ok(Variant::from_string(name.as_str()))
            }
            Self::Element { text, .. } if member == "text" && !put && args.is_empty() => {
                Ok(Variant::from_string(text.as_str()))
            }
            _ => Err(fault(host, 438)),
        }
    }
}

fn argument<'a>(
    host: &StandardHostServices,
    args: &'a [DynamicCallArg],
    parameter: &str,
) -> HalResult<&'a Variant> {
    let arg = match args {
        [arg] => arg,
        [] => return Err(fault(host, 449)),
        _ => return Err(fault(host, 450)),
    };
    if arg
        .name
        .as_ref()
        .is_some_and(|name| !name.eq_ignore_ascii_case(parameter))
    {
        return Err(fault(host, 448));
    }
    arg.value
        .as_ref()
        .map(|value| value.variant())
        .ok_or_else(|| fault(host, 449))
}

fn fault(host: &StandardHostServices, number: i32) -> HalError {
    host.com_call_fault("dispatch_invoke", ComFault::number(number))
}

fn document_element(xml: &str) -> Option<BuiltinObject> {
    let parsed = roxmltree::Document::parse(xml).ok()?;
    let root = parsed.root_element();
    let name = xml[root.range().start + 1..]
        .split([' ', '\t', '\r', '\n', '/', '>'])
        .next()?
        .to_owned();
    let text: String = root
        .descendants()
        .filter(roxmltree::Node::is_text)
        .filter_map(|node| node.text())
        .collect();
    Some(BuiltinObject::Element {
        name,
        text: text.trim_matches([' ', '\t', '\r', '\n']).to_owned(),
    })
}
