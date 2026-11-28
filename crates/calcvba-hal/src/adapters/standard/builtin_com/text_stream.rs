use std::fs::{self, File, OpenOptions};
use std::io::Write;

use calcvba_com::DynamicCallArg;
use calcvba_runtime::ansi::{ansi_decode_units, ansi_encode_units};
use calcvba_runtime::{VarType, Variant, coerce::coerce_to, variant_to_vba_string};

use super::{BuiltinObject, com_file_error, com_file_path, fault, ordered_arguments};
use crate::HalResult;
use crate::adapters::standard::StandardHostServices;

pub(in crate::adapters::standard) enum TextStream {
    Input { units: Vec<u16>, offset: usize },
    Output { file: File, unicode: bool },
    Closed,
}

impl TextStream {
    pub(super) fn create(
        host: &StandardHostServices,
        args: &[DynamicCallArg],
    ) -> HalResult<Variant> {
        let [filename, overwrite, unicode] =
            ordered_arguments(host, args, ["filename", "overwrite", "unicode"])?;
        let file_path = com_file_path(host, filename, "fso_create_text_file", true)?;
        let overwrite = bool_arg(host, overwrite, true)?;
        let unicode = bool_arg(host, unicode, false)?;
        let mut options = OpenOptions::new();
        options.write(true);
        if overwrite {
            options.create(true).truncate(true);
        } else {
            options.create_new(true);
        }
        let mut file = options
            .open(&file_path)
            .map_err(|error| com_file_error(host, "fso_create_text_file", &file_path, error))?;
        if unicode {
            file.write_all(&[0xFF, 0xFE])
                .map_err(|error| com_file_error(host, "fso_create_text_file", &file_path, error))?;
        }
        BuiltinObject::TextStream(Self::Output { file, unicode }.into()).retain(host)
    }

    pub(super) fn open(host: &StandardHostServices, args: &[DynamicCallArg]) -> HalResult<Variant> {
        let [filename, mode, create, format] =
            ordered_arguments(host, args, ["filename", "iomode", "create", "format"])?;
        let mode = long_arg(host, mode, 1)?;
        let create = bool_arg(host, create, false)?;
        let format = long_arg(host, format, 0)?;
        let path = com_file_path(host, filename, "fso_open_text_file", mode != 1 || create)?;
        let unicode = format == -1;
        let stream = match mode {
            1 => {
                let mut data = match fs::read(&path) {
                    Ok(bytes) => bytes,
                    Err(error) if create && error.kind() == std::io::ErrorKind::NotFound => {
                        File::create_new(&path).map_err(|error| {
                            com_file_error(host, "fso_open_text_file", &path, error)
                        })?;
                        Vec::new()
                    }
                    Err(error) => {
                        return Err(com_file_error(host, "fso_open_text_file", &path, error));
                    }
                };
                let units = if unicode {
                    if data.starts_with(&[0xFF, 0xFE]) {
                        data.drain(..2);
                    }
                    data.as_chunks::<2>()
                        .0
                        .iter()
                        .copied()
                        .map(u16::from_le_bytes)
                        .collect()
                } else {
                    ansi_decode_units(&data)
                };
                Self::Input { units, offset: 0 }
            }
            2 | 8 => {
                let mut options = OpenOptions::new();
                options
                    .create(create)
                    .write(true)
                    .truncate(mode == 2)
                    .append(mode == 8);
                let mut file = options
                    .open(&path)
                    .map_err(|error| com_file_error(host, "fso_open_text_file", &path, error))?;
                if unicode && file.metadata().map(|meta| meta.len()).unwrap_or(0) == 0 {
                    file.write_all(&[0xFF, 0xFE]).map_err(|error| {
                        com_file_error(host, "fso_open_text_file", &path, error)
                    })?;
                }
                Self::Output { file, unicode }
            }
            _ => return Err(fault(host, 5)),
        };
        BuiltinObject::TextStream(stream.into()).retain(host)
    }

    pub(super) fn call(
        &mut self,
        host: &StandardHostServices,
        member: &str,
        args: &[DynamicCallArg],
    ) -> HalResult<Variant> {
        if member == "close" && args.is_empty() {
            *self = Self::Closed;
            return Ok(Variant::empty());
        }
        match self {
            Self::Input { units, offset } => match member {
                "readall" if args.is_empty() => {
                    let text = Variant::from_utf16_units(&units[*offset..]);
                    *offset = units.len();
                    Ok(text)
                }
                "readline" if args.is_empty() => {
                    let remainder = &units[*offset..];
                    let end = remainder
                        .iter()
                        .position(|unit| [10, 13].contains(unit))
                        .unwrap_or(remainder.len());
                    let text = Variant::from_utf16_units(&remainder[..end]);
                    let newline = match remainder.get(end..) {
                        Some([13, 10, ..]) => 2,
                        Some([10 | 13, ..]) => 1,
                        _ => 0,
                    };
                    *offset += end + newline;
                    Ok(text)
                }
                "atendofstream" if args.is_empty() => {
                    Ok(Variant::from_bool(*offset == units.len()))
                }
                _ => Err(fault(host, 438)),
            },
            Self::Output { file, unicode } if matches!(member, "write" | "writeline") => {
                let [arg] = ordered_arguments(host, args, ["text"])?;
                let mut units = match arg {
                    Some(text) => variant_to_vba_string(text)
                        .map_err(|_| fault(host, 13))?
                        .payload_units()
                        .to_vec(),
                    None if member == "writeline" => Vec::new(),
                    None => return Err(fault(host, 449)),
                };
                if member == "writeline" {
                    units.extend([13, 10]);
                }
                let bytes = if *unicode {
                    units.into_iter().flat_map(u16::to_le_bytes).collect()
                } else {
                    ansi_encode_units(&units)
                };
                file.write_all(&bytes).map_err(|error| {
                    com_file_error(
                        host,
                        "fso_text_stream_write",
                        std::path::Path::new(""),
                        error,
                    )
                })?;
                Ok(Variant::empty())
            }
            _ => Err(fault(host, 438)),
        }
    }
}

fn bool_arg(
    host: &StandardHostServices,
    value: Option<&Variant>,
    default: bool,
) -> HalResult<bool> {
    value.map_or(Ok(default), |value| {
        coerce_to(value, VarType::Boolean)
            .ok()
            .and_then(|value| value.as_bool())
            .ok_or_else(|| fault(host, 13))
    })
}

pub(super) fn long_arg(
    host: &StandardHostServices,
    value: Option<&Variant>,
    default: i32,
) -> HalResult<i32> {
    value.map_or(Ok(default), |value| {
        coerce_to(value, VarType::Long)
            .ok()
            .and_then(|value| value.as_i32())
            .ok_or_else(|| fault(host, 13))
    })
}
