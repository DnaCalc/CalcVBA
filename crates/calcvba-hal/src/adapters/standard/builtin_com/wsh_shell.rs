use calcvba_com::DynamicCallArg;
use calcvba_runtime::{Variant, variant_to_vba_string};

use super::{argument, fault, ordered_arguments, text_stream::long_arg};
use crate::HalResult;
use crate::adapters::standard::{StandardHostServices, process::fold_env_name};
use crate::model::CapabilityId;
use crate::traits::UiInteractionHal;

pub(super) fn call(
    host: &StandardHostServices,
    member: &str,
    put: bool,
    args: &[DynamicCallArg],
) -> HalResult<Variant> {
    match (member, put) {
        ("expandenvironmentstrings", false) => {
            let value = argument(host, args, "strString")?;
            let text = variant_to_vba_string(value).map_err(|_| fault(host, 13))?;
            Ok(Variant::from_utf16_units(&expand(
                host,
                text.payload_units(),
            )))
        }
        ("currentdirectory", false) if args.is_empty() => {
            host.mapped_cur_dir(None).map(Variant::from_string)
        }
        ("currentdirectory", true) => {
            let value =
                variant_to_vba_string(argument(host, args, "")?).map_err(|_| fault(host, 13))?;
            let path = value.as_str();
            host.mapped_ch_dir(&path)?;
            if let [drive, b':', ..] = path.as_bytes()
                && drive.is_ascii_alphabetic()
            {
                host.mapped_ch_drive(char::from(*drive))?;
            }
            Ok(Variant::empty())
        }
        ("specialfolders", false) => {
            let value = argument(host, args, "index")?;
            let folder = variant_to_vba_string(value).map_err(|_| fault(host, 13))?;
            let selected = match folder.as_str().to_ascii_lowercase().as_str() {
                "appdata" => environment(host, "APPDATA"),
                "desktop" => profile_folder(host, "Desktop"),
                "mydocuments" => profile_folder(host, "Documents"),
                _ => None,
            };
            Ok(Variant::from_string(selected.unwrap_or_default()))
        }
        ("popup", false) => {
            let [prompt, seconds, title, style] = ordered_arguments(
                host,
                args,
                ["strText", "nSecondsToWait", "strTitle", "nType"],
            )?;
            let prompt = prompt.ok_or_else(|| fault(host, 449))?;
            if long_arg(host, seconds, 0)? != 0 || title.is_some() {
                return Err(host.unsupported(CapabilityId::UiInteraction, "wscript_popup"));
            }
            host.msg_box_variant(prompt.clone(), Variant::from_i32(long_arg(host, style, 0)?))
        }
        _ => Err(fault(host, 438)),
    }
}

fn expand(host: &StandardHostServices, text: &[u16]) -> Vec<u16> {
    const PERCENT: u16 = b'%' as u16;
    let environment = host.windows_environment();
    let mut expanded = Vec::with_capacity(text.len());
    let mut pending = text;
    while let Some(start) = pending.iter().position(|&unit| unit == PERCENT) {
        expanded.extend_from_slice(&pending[..start]);
        pending = &pending[start..];
        let Some(length) = pending[1..].iter().position(|&unit| unit == PERCENT) else {
            break;
        };
        let end = length + 1;
        let value = String::from_utf16(&pending[1..end]).ok().and_then(|name| {
            let name = fold_env_name(&name);
            environment
                .iter()
                .find(|(key, _)| key.to_str().is_some_and(|key| fold_env_name(key) == name))
                .map(|(_, value)| value.to_string_lossy())
        });
        if let Some(value) = value {
            expanded.extend(value.encode_utf16());
        } else {
            expanded.extend_from_slice(&pending[..=end]);
        }
        pending = &pending[end + 1..];
    }
    expanded.extend_from_slice(pending);
    expanded
}

fn environment(host: &StandardHostServices, key: &str) -> Option<String> {
    host.windows_environment()
        .into_iter()
        .find(|(name, _)| name.to_str().is_some_and(|name| fold_env_name(name) == key))
        .map(|(_, value)| value.to_string_lossy().into_owned())
}

fn profile_folder(host: &StandardHostServices, name: &str) -> Option<String> {
    let profile = environment(host, "USERPROFILE")?;
    let path = format!("{}\\{name}", profile.trim_end_matches('\\'));
    let directory = host
        .map_path(
            &path,
            CapabilityId::FileSystemIo,
            "wscript_special_folders",
            false,
        )
        .ok()?;
    directory.is_dir().then_some(path)
}
