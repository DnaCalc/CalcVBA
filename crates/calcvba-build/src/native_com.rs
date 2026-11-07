//! Ordinary compiler-generated VBA shims give COM the existing native ABI.
//! The server template owns only factories, Automation type information and
//! its STA message loop. It cannot load or interpret an OxImage.
use crate::{BuildError, ComInvokeKind, ComParamType, ComServerDescriptor};
use calcvba_project::{
    CallingConvention, ErrorPolicy, LoadedProject, NativeExportDescriptor, StringEncoding,
};
use calcvba_rt_abi::native_export::{NativeComClass, NativeComServer};
use calcvba_symbol::manifest::{ModuleAttributes, ModuleKind, ModuleUnit, SymbolProjectManifest};
use std::{collections::HashSet, path::Path};

const MODULE: &str = "__CalcVbaComEntries";
pub(super) struct Prepared {
    descriptor: ComServerDescriptor,
    server: NativeComServer,
    pub receiver_exports: HashSet<String>,
}
fn error(message: impl ToString) -> BuildError {
    BuildError::NativeExport(message.to_string())
}
fn declaration(ty: ComParamType) -> Result<(&'static str, &'static str), BuildError> {
    use ComParamType::*;
    Ok(match ty {
        Long => ("ByVal", "Long"),
        Integer => ("ByVal", "Integer"),
        Byte => ("ByVal", "Byte"),
        Boolean => ("ByVal", "Boolean"),
        Double => ("ByVal", "Double"),
        Single => ("ByVal", "Single"),
        Currency => ("ByVal", "Currency"),
        Date => ("ByVal", "Date"),
        String => ("ByVal", "String"),
        LongLong => ("ByVal", "LongLong"),
        LongPtr => ("ByVal", "LongPtr"),
        ByRefLong => ("ByRef", "Long"),
        ByRefInteger => ("ByRef", "Integer"),
        ByRefByte => ("ByRef", "Byte"),
        ByRefBoolean => ("ByRef", "Boolean"),
        ByRefDouble => ("ByRef", "Double"),
        ByRefSingle => ("ByRef", "Single"),
        ByRefCurrency => ("ByRef", "Currency"),
        ByRefDate => ("ByRef", "Date"),
        ByRefString => ("ByRef", "String"),
        ByRefLongLong => ("ByRef", "LongLong"),
        ByRefLongPtr => ("ByRef", "LongPtr"),
        Variant => ("ByVal", "Variant"),
        ByRefVariant => ("ByRef", "Variant"),
        Object => ("ByVal", "Object"),
        ByRefObject => ("ByRef", "Object"),
        _ => {
            return Err(error(format!(
                "ComExe Automation boundary {ty:?} is not implemented"
            )));
        }
    })
}
fn export(name: String) -> NativeExportDescriptor {
    NativeExportDescriptor {
        exported_name: name.clone(),
        module_name: MODULE.into(),
        procedure_name: name,
        calling_convention: CallingConvention::Stdcall,
        ordinal: None,
        string_encoding: StringEncoding::Bstr,
        error_policy: ErrorPolicy::HResult,
        kind: None,
        param_types: None,
        return_type: None,
    }
}
pub(super) fn prepare(
    loaded: &mut LoadedProject,
    root: &mut SymbolProjectManifest,
) -> Result<Prepared, BuildError> {
    if !loaded.native_exports.is_empty() {
        return Err(error(
            "ComExe selects exposed classes; NativeExport declarations require NativeLibrary",
        ));
    }
    if root
        .modules
        .iter()
        .any(|module| module.module_name.eq_ignore_ascii_case(MODULE))
    {
        return Err(error(format!("reserved ComExe module {MODULE}")));
    }
    let descriptor = crate::build_com_descriptor(root)?;
    let mut source = String::from(
        "Option Explicit\nPublic Sub CvbComRelease(ByRef instance As Object)\nSet instance = Nothing\nEnd Sub\nPrivate Function CvbComIsMissing(ByVal v As Variant) As Boolean\nIf IsError(v) Then CvbComIsMissing = (CStr(v) = \"Error 448\")\nEnd Function\n",
    );
    loaded.native_exports.push(export("CvbComRelease".into()));
    let mut server = NativeComServer {
        classes: Vec::new(),
    };
    let mut receiver_exports = HashSet::new();
    for (class_index, class) in descriptor.classes.iter().enumerate() {
        if !class.events.is_empty() || !class.implemented_interfaces.is_empty() {
            return Err(error(format!(
                "ComExe class {}: exported events/custom interfaces are not implemented",
                class.class_name
            )));
        }
        if !class.creatable {
            continue;
        }
        let constructor = format!("CvbComNew{class_index}");
        source.push_str(&format!(
            "Public Function {constructor}() As {}\nSet {constructor} = New {}\nEnd Function\n",
            class.class_name, class.class_name
        ));
        loaded.native_exports.push(export(constructor.clone()));
        let mut entries = Vec::new();
        for (member_index, member) in class.members.iter().enumerate() {
            // An omitted Optional Variant arrives as the Automation Missing value (Error 448).
            // The generated entry tests each one and calls the member without it, so the
            // member's own IsMissing and default handling decide, exactly as for a VBA caller.
            let mut optional = Vec::new();
            for (index, flag) in member.parameter_optional.iter().enumerate() {
                if !*flag {
                    continue;
                }
                if !matches!(
                    member.parameter_types.get(index),
                    Some(ComParamType::Variant)
                ) {
                    return Err(error(format!(
                        "ComExe {}.{}: Optional parameter {} is not a Variant; only Optional Variant Automation parameters are implemented",
                        class.class_name, member.name, index
                    )));
                }
                optional.push(index);
            }
            if optional.len() > 4 {
                return Err(error(format!(
                    "ComExe {}.{}: more than four Optional parameters are not implemented",
                    class.class_name, member.name
                )));
            }
            if member.vtable_slot != Some(7 + member_index as u16) {
                return Err(error("ComExe requires contiguous dual-interface slots"));
            }
            let name = format!("CvbComEntry{class_index}_{member_index}");
            let mut params = format!("ByVal instance As {}", class.class_name);
            for (index, ty) in member.parameter_types.iter().enumerate() {
                let (passing, ty) = declaration(*ty)?;
                params.push_str(&format!(", {passing} arg{index} As {ty}"));
            }
            let function = member.return_type;
            let object_result = match function {
                Some(ty) => {
                    let (passing, ty) = declaration(ty)?;
                    if passing != "ByVal" {
                        return Err(error("ComExe retval cannot be ByRef"));
                    }
                    ty == "Object"
                }
                None => false,
            };
            // One statement per combination of omitted optional arguments.
            let statement = |omitted: &[usize]| -> Result<String, BuildError> {
                let mut args: Vec<String> = (0..member.parameter_types.len())
                    .map(|index| {
                        if omitted.contains(&index) {
                            String::new()
                        } else {
                            format!("arg{index}")
                        }
                    })
                    .collect();
                let target = format!("instance.{}", member.name);
                let trim = |args: &mut Vec<String>| {
                    while args.last().is_some_and(String::is_empty) {
                        args.pop();
                    }
                };
                let call = match member.invoke_kind {
                    ComInvokeKind::Method | ComInvokeKind::PropertyGet => {
                        trim(&mut args);
                        if args.is_empty() && member.invoke_kind == ComInvokeKind::PropertyGet {
                            target
                        } else {
                            format!("{target}({})", args.join(", "))
                        }
                    }
                    ComInvokeKind::PropertyPut | ComInvokeKind::PropertyPutRef => {
                        let value = args
                            .pop()
                            .ok_or_else(|| error("property put has no value"))?;
                        trim(&mut args);
                        let target = if args.is_empty() {
                            target
                        } else {
                            format!("{target}({})", args.join(", "))
                        };
                        let set = if member.invoke_kind == ComInvokeKind::PropertyPutRef {
                            "Set "
                        } else {
                            ""
                        };
                        format!("{set}{target} = {value}")
                    }
                };
                Ok(if function.is_some() {
                    // An Object result is assigned with Set; Let into an Object never completes.
                    let set = if object_result { "Set " } else { "" };
                    format!("{set}{name} = {call}")
                } else if member.invoke_kind == ComInvokeKind::Method {
                    format!("Call {call}")
                } else {
                    call
                })
            };
            let mut body = String::new();
            if optional.is_empty() {
                body.push_str(&statement(&[])?);
                body.push('\n');
            } else {
                for (position, index) in optional.iter().enumerate() {
                    body.push_str(&format!(
                        "Dim missing{position} As Boolean\nmissing{position} = CvbComIsMissing(arg{index})\n"
                    ));
                }
                for mask in (0..(1usize << optional.len())).rev() {
                    let omitted: Vec<usize> = optional
                        .iter()
                        .enumerate()
                        .filter(|(position, _)| mask & (1 << position) != 0)
                        .map(|(_, index)| *index)
                        .collect();
                    let condition = if mask == 0 {
                        "True".to_string()
                    } else {
                        optional
                            .iter()
                            .enumerate()
                            .map(|(position, _)| {
                                if mask & (1 << position) != 0 {
                                    format!("missing{position}")
                                } else {
                                    format!("Not missing{position}")
                                }
                            })
                            .collect::<Vec<_>>()
                            .join(" And ")
                    };
                    let keyword = if mask == (1 << optional.len()) - 1 {
                        "If"
                    } else if mask == 0 {
                        "Else"
                    } else {
                        "ElseIf"
                    };
                    if keyword == "Else" {
                        body.push_str("Else\n");
                    } else {
                        body.push_str(&format!("{keyword} {condition} Then\n"));
                    }
                    body.push_str(&statement(&omitted)?);
                    body.push('\n');
                }
                body.push_str("End If\n");
            }
            if let Some(ty) = function {
                let (_, ty) = declaration(ty)?;
                source.push_str(&format!(
                    "Public Function {name}({params}) As {ty}\n{body}End Function\n"
                ));
            } else {
                source.push_str(&format!("Public Sub {name}({params})\n{body}End Sub\n"));
            }
            loaded.native_exports.push(export(name.clone()));
            receiver_exports.insert(name.clone());
            entries.push(name);
        }
        let single_use = loaded
            .class_module_metadata
            .get(&class.class_name)
            .and_then(|metadata| metadata.instancing)
            .is_some_and(|mode| {
                matches!(
                    mode,
                    calcvba_project::Instancing::SingleUse
                        | calcvba_project::Instancing::GlobalSingleUse
                )
            });
        server.classes.push(NativeComClass {
            name: class.class_name.clone(),
            clsid: class.clsid.clone(),
            iid: class.default_interface_iid.clone(),
            constructor,
            release: "CvbComRelease".into(),
            entries,
            single_use,
        });
    }
    if server.classes.is_empty() {
        return Err(error("ComExe requires at least one creatable class"));
    }
    root.modules.push(ModuleUnit {
        module_name: MODULE.into(),
        module_kind: ModuleKind::Procedural,
        attributes: ModuleAttributes::named(MODULE),
        source: source.clone(),
    });
    loaded
        .manifest
        .modules
        .push(calcvba_project::manifest::ModuleUnit {
            module_name: MODULE.into(),
            module_kind: calcvba_project::manifest::ModuleKind::Procedural,
            attributes: calcvba_project::manifest::ModuleAttributes {
                vb_name: MODULE.into(),
                ..Default::default()
            },
            source,
        });
    Ok(Prepared {
        descriptor,
        server,
        receiver_exports,
    })
}
pub(super) fn write(directory: &Path, name: &str, prepared: &Prepared) -> Result<(), BuildError> {
    crate::write_bytes(
        &directory.join(format!("{name}.native-com.json")),
        &serde_json::to_vec_pretty(&prepared.server).map_err(BuildError::DescriptorSerialize)?,
    )?;
    crate::write_bytes(
        &directory.join(format!("{name}.comserver.json")),
        &serde_json::to_vec_pretty(&prepared.descriptor)
            .map_err(BuildError::DescriptorSerialize)?,
    )?;
    #[cfg(windows)]
    crate::oleaut_typelib::emit_native_typelib(
        &prepared.descriptor,
        &directory.join(format!("{name}.tlb")),
    )?;
    #[cfg(not(windows))]
    return Err(error("ComExe type-library output requires Windows"));
    #[cfg(windows)]
    Ok(())
}
