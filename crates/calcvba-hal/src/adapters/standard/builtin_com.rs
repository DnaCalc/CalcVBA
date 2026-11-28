mod ado_stream;
mod text_stream;
mod wsh_shell;
mod xml;

use calcvba_com::call::ComFault;
use calcvba_com::{DynamicCallArg, DynamicCallKind, DynamicCallRequest, DynamicMemberSelector};
use calcvba_runtime::{
    ObjectRef, ObjectRefExt, RuntimeCallResult, VarType, Variant, coerce::coerce_to,
    variant_to_vba_string,
};
use std::cell::RefCell;
use std::io;
use std::path::{Path, PathBuf};
use std::rc::Rc;

use super::StandardHostServices;
use crate::{HalError, HalResult, model::CapabilityId};
use ado_stream::AdoStream;
use text_stream::TextStream;
use xml::XmlElement;

pub(super) enum BuiltinObject {
    FileSystem,
    TextStream(RefCell<TextStream>),
    AdoStream(RefCell<AdoStream>),
    WshShell,
    Document(RefCell<Document>),
    Element(XmlElement),
}

pub(super) struct Document {
    asynchronous: bool,
    root: Variant,
}

impl BuiltinObject {
    fn name(&self) -> &'static str {
        match self {
            Self::FileSystem => "FileSystemObject",
            Self::TextStream(_) => "TextStream",
            Self::AdoStream(_) => "Stream",
            Self::WshShell => "WshShell",
            Self::Document(_) => "DOMDocument",
            Self::Element(_) => "IXMLDOMElement",
        }
    }

    pub(super) fn create(host: &StandardHostServices, prog_id: &str) -> HalResult<Option<Variant>> {
        use calcvba_com::platform::portable::BuiltinProgId;
        let object = match BuiltinProgId::resolve(prog_id) {
            Some(BuiltinProgId::FileSystemObject) if host.native_fs_enabled() => Self::FileSystem,
            Some(BuiltinProgId::AdoStream) if host.native_fs_enabled() => {
                Self::AdoStream(RefCell::default())
            }
            Some(BuiltinProgId::WshShell) if host.policy.windows_faithful => Self::WshShell,
            Some(BuiltinProgId::XmlDocument) => Self::Document(RefCell::new(Document {
                asynchronous: true,
                root: Variant::nothing(),
            })),
            _ => return Ok(None),
        };
        object.retain(host).map(Some)
    }

    fn retain(self, host: &StandardHostServices) -> HalResult<Variant> {
        let identity = match &self {
            Self::Element(node) => node.identity(host)?,
            _ => host
                .borrow_projection(CapabilityId::ComActivationDispatch, "create_object")?
                .allocate_handle(),
        };
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
            Self::FileSystem if member == "createtextfile" && !put => {
                TextStream::create(host, args)
            }
            Self::FileSystem if member == "opentextfile" && !put => TextStream::open(host, args),
            Self::TextStream(stream) if !put => stream.borrow_mut().call(host, member, args),
            Self::AdoStream(stream) => stream.borrow_mut().call(host, member, put, args),
            Self::WshShell => wsh_shell::call(host, member, put, args),
            Self::FileSystem if !put && matches!(member, "getfilename" | "getextensionname") => {
                let extension = member == "getextensionname";
                let parameter = if extension { "path" } else { "pathspec" };
                let path = variant_to_vba_string(argument(host, args, parameter)?)
                    .map_err(|_| fault(host, 13))?;
                let name = windows_file_name(path.payload_units());
                let result = if extension {
                    name.iter()
                        .rposition(|&unit| unit == u16::from(b'.'))
                        .map_or(&[][..], |dot| &name[dot + 1..])
                } else {
                    name
                };
                Ok(Variant::from_utf16_units(result))
            }
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
                    .and_then(|text| XmlElement::parse(&text).map(Self::Element));
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
            Self::Document(doc) if member == "selectsinglenode" && !put => {
                let root = doc
                    .borrow()
                    .root
                    .as_object_ref()
                    .and_then(|object| object.local_state::<Self>());
                match root.as_deref() {
                    Some(Self::Element(node)) => node.select(host, args, true),
                    _ => Ok(Variant::nothing()),
                }
            }
            Self::Element(node) if member == "selectsinglenode" && !put => {
                node.select(host, args, false)
            }
            Self::Element(node) if member == "nodename" && !put && args.is_empty() => {
                Ok(Variant::from_string(node.name.as_str()))
            }
            Self::Element(node) if member == "text" && !put && args.is_empty() => {
                Ok(Variant::from_string(node.text.as_str()))
            }
            _ => Err(fault(host, 438)),
        }
    }
}

fn windows_file_name(path: &[u16]) -> &[u16] {
    let path = match path {
        [drive, colon, rest @ ..]
            if u8::try_from(*drive).is_ok_and(|letter| letter.is_ascii_alphabetic())
                && *colon == u16::from(b':') =>
        {
            rest
        }
        _ => path,
    };
    path.rsplit(|&unit| unit == u16::from(b'\\') || unit == u16::from(b'/'))
        .next()
        .unwrap_or_default()
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

fn ordered_arguments<'a, const N: usize>(
    host: &StandardHostServices,
    args: &'a [DynamicCallArg],
    parameters: [&str; N],
) -> HalResult<[Option<&'a Variant>; N]> {
    let mut values = [None; N];
    let mut supplied = [false; N];
    let mut next = 0;
    for argument in args {
        let index = if let Some(name) = &argument.name {
            parameters
                .iter()
                .position(|parameter| name.eq_ignore_ascii_case(parameter))
                .ok_or_else(|| fault(host, 448))?
        } else {
            while next < N && supplied[next] {
                next += 1;
            }
            if next == N {
                return Err(fault(host, 450));
            }
            next
        };
        if supplied[index] {
            return Err(fault(host, 450));
        }
        supplied[index] = true;
        values[index] = argument.value.as_ref().map(|value| value.variant());
    }
    Ok(values)
}

fn com_file_path(
    host: &StandardHostServices,
    path: Option<&Variant>,
    operation: &'static str,
    write: bool,
) -> HalResult<PathBuf> {
    let capability = CapabilityId::FileSystemIo;
    if !host.native_fs_enabled() || !host.supports(capability) {
        return Err(host.unsupported(capability, operation));
    }
    if write && !host.policy.allow_filesystem_mutation {
        return Err(host.denied(capability, operation));
    }
    let path = path.ok_or_else(|| fault(host, 449))?;
    let path = variant_to_vba_string(path).map_err(|_| fault(host, 13))?;
    host.map_path(&path.as_str(), capability, operation, write)
}

fn com_file_error(
    host: &StandardHostServices,
    operation: &'static str,
    path: &Path,
    error: io::Error,
) -> HalError {
    let code = match error.kind() {
        io::ErrorKind::NotFound
            if path
                .parent()
                .is_some_and(|parent| !parent.as_os_str().is_empty() && !parent.is_dir()) =>
        {
            76
        }
        io::ErrorKind::NotADirectory => 76,
        io::ErrorKind::NotFound => 53,
        io::ErrorKind::AlreadyExists => 58,
        io::ErrorKind::PermissionDenied => 70,
        _ => 75,
    };
    HalError::adapter_fault(
        host.profile,
        CapabilityId::FileSystemIo,
        operation,
        error.to_string(),
    )
    .with_host_error_code(code)
}

fn fault(host: &StandardHostServices, number: i32) -> HalError {
    host.com_call_fault("dispatch_invoke", ComFault::number(number))
}
