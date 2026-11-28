use std::fs::{self, OpenOptions};
use std::io::Write;

use calcvba_com::DynamicCallArg;
use calcvba_runtime::{VarType, Variant, safe_array::SafeArray, variant_to_vba_string};

use super::{com_file_error, com_file_path, fault, ordered_arguments, text_stream::long_arg};
use crate::HalResult;
use crate::adapters::standard::StandardHostServices;

#[derive(Default)]
pub(in crate::adapters::standard) struct AdoStream {
    contents: Vec<u8>,
    position: usize,
    open: bool,
    binary: bool,
    charset: Charset,
}

#[derive(Clone, Copy, Default)]
enum Charset {
    #[default]
    Unicode,
    Utf8,
}

impl Charset {
    fn name(self) -> &'static str {
        match self {
            Self::Unicode => "Unicode",
            Self::Utf8 => "utf-8",
        }
    }

    fn marker(self) -> &'static [u8] {
        match self {
            Self::Unicode => &[0xFF, 0xFE],
            Self::Utf8 => &[0xEF, 0xBB, 0xBF],
        }
    }
}

impl AdoStream {
    pub(super) fn call(
        &mut self,
        host: &StandardHostServices,
        member: &str,
        put: bool,
        args: &[DynamicCallArg],
    ) -> HalResult<Variant> {
        match (member, put) {
            ("state", false) if args.is_empty() => {
                return Ok(Variant::from_i32(i32::from(self.open)));
            }
            ("type", false) if args.is_empty() => {
                return Ok(Variant::from_i32(if self.binary { 1 } else { 2 }));
            }
            ("charset", false) if args.is_empty() => {
                return Ok(Variant::from_string(self.charset.name()));
            }
            ("type", true) if self.position == 0 => {
                let [value] = ordered_arguments(host, args, [""])?;
                self.binary = match long_arg(host, value, 0)? {
                    1 => true,
                    2 => false,
                    _ => return Err(fault(host, 3001)),
                };
                return Ok(Variant::empty());
            }
            ("charset", true) if self.position == 0 => {
                let [value] = ordered_arguments(host, args, [""])?;
                let value = value.ok_or_else(|| fault(host, 449))?;
                let label = variant_to_vba_string(value).map_err(|_| fault(host, 3001))?;
                self.charset = match label.as_str().to_ascii_lowercase().as_str() {
                    "unicode" | "utf-16" | "utf-16le" => Charset::Unicode,
                    "utf-8" => Charset::Utf8,
                    _ => return Err(fault(host, 3001)),
                };
                return Ok(Variant::empty());
            }
            ("open", false) if args.is_empty() => {
                if self.open {
                    return Err(fault(host, 3705));
                }
                self.open = true;
                return Ok(Variant::empty());
            }
            _ => {}
        }
        if !self.open {
            return Err(fault(host, 3704));
        }
        match (member, put) {
            ("close", false) if args.is_empty() => {
                self.open = false;
                self.contents.clear();
                self.position = 0;
            }
            ("size", false) if args.is_empty() => {
                return Ok(Variant::from_i32(self.contents.len() as i32));
            }
            ("position", false) if args.is_empty() => {
                return Ok(Variant::from_i32(self.position as i32));
            }
            ("eos", false) if args.is_empty() => {
                return Ok(Variant::from_bool(self.position == self.contents.len()));
            }
            ("position", true) => {
                let [value] = ordered_arguments(host, args, [""])?;
                let position =
                    usize::try_from(long_arg(host, value, -1)?).map_err(|_| fault(host, 3001))?;
                self.grow(host, position)?;
                self.position = position;
            }
            ("seteos", false) if args.is_empty() => self.contents.truncate(self.position),
            ("write", false) if self.binary || self.contents.is_empty() => {
                let [value] = ordered_arguments(host, args, ["buffer"])?;
                let array = value.ok_or_else(|| fault(host, 449))?;
                let (bounds, length) = array
                    .safearray_bounds_len()
                    .ok_or_else(|| fault(host, 3001))?;
                if bounds.len() != 1 || array.array_element_vartype() != Some(VarType::Byte as u16)
                {
                    return Err(fault(host, 3001));
                }
                let data: Vec<u8> = (0..length)
                    .map(|index| {
                        array
                            .safearray_element(index)
                            .and_then(Result::ok)
                            .and_then(|byte| byte.as_u8())
                            .ok_or_else(|| fault(host, 3001))
                    })
                    .collect::<HalResult<_>>()?;
                self.write(host, &data)?;
                self.binary = true;
            }
            ("read", false) if self.binary => {
                let [count] = ordered_arguments(host, args, ["numbytes"])?;
                let count = long_arg(host, count, -1)?;
                let data = &self.contents[self.position..];
                if data.is_empty() {
                    return Ok(Variant::null());
                }
                let count = match count {
                    -1 => data.len(),
                    0.. => data.len().min(count as usize),
                    _ => return Err(fault(host, 3001)),
                };
                let array = SafeArray::from_typed_variants(
                    VarType::Byte as u16,
                    data[..count]
                        .iter()
                        .copied()
                        .map(Variant::from_u8)
                        .collect(),
                )
                .map_err(|_| fault(host, 3001))?;
                self.position += count;
                return Ok(Variant::from_safearray(array));
            }
            ("writetext", false) if !self.binary => {
                let [value, options] = ordered_arguments(host, args, ["data", "options"])?;
                let value = value.ok_or_else(|| fault(host, 449))?;
                let text = variant_to_vba_string(value).map_err(|_| fault(host, 3001))?;
                let mut units = text.payload_units().to_vec();
                match long_arg(host, options, 0)? {
                    0 => {}
                    1 => units.extend([13, 10]),
                    _ => return Err(fault(host, 3001)),
                }
                let mut data = if self.position == 0 {
                    self.charset.marker().to_vec()
                } else {
                    Vec::new()
                };
                match self.charset {
                    Charset::Unicode => data.extend(units.into_iter().flat_map(u16::to_le_bytes)),
                    Charset::Utf8 => {
                        let text = String::from_utf16(&units).map_err(|_| fault(host, 3001))?;
                        data.extend(text.bytes());
                    }
                }
                self.write(host, &data)?;
            }
            ("readtext", false) if !self.binary => {
                let [count] = ordered_arguments(host, args, ["numchars"])?;
                let count = long_arg(host, count, -1)?;
                if count < -1 {
                    return Err(fault(host, 3001));
                }
                let data = &self.contents[self.position..];
                let (data, skipped) = match data.strip_prefix(self.charset.marker()) {
                    Some(contents) if self.position == 0 => (contents, self.charset.marker().len()),
                    _ => (data, 0),
                };
                let (units, used) = match self.charset {
                    Charset::Unicode => {
                        let (chunks, remainder) = data.as_chunks::<2>();
                        if !remainder.is_empty() {
                            return Err(fault(host, 3001));
                        }
                        let units = chunks.iter().copied().map(u16::from_le_bytes);
                        let units: Vec<u16> = units
                            .take(if count == -1 {
                                usize::MAX
                            } else {
                                count as usize
                            })
                            .collect();
                        let used = units.len() * 2;
                        (units, used)
                    }
                    Charset::Utf8 => {
                        let text = std::str::from_utf8(data).map_err(|_| fault(host, 3001))?;
                        let used = if count == -1 {
                            data.len()
                        } else {
                            text.char_indices()
                                .nth(count as usize)
                                .map_or(data.len(), |(index, _)| index)
                        };
                        (text[..used].encode_utf16().collect(), used)
                    }
                };
                self.position += skipped + used;
                return Ok(Variant::from_utf16_units(&units));
            }
            ("loadfromfile", false) => {
                let [filename] = ordered_arguments(host, args, ["filename"])?;
                let path = com_file_path(host, filename, "ado_stream_load", false)?;
                let contents = fs::read(&path)
                    .map_err(|error| com_file_error(host, "ado_stream_load", &path, error))?;
                if contents.len() > i32::MAX as usize {
                    return Err(fault(host, 3734));
                }
                self.contents = contents;
                self.position = 0;
            }
            ("savetofile", false) => {
                let [filename, options] =
                    ordered_arguments(host, args, ["filename", "saveoptions"])?;
                let path = com_file_path(host, filename, "ado_stream_save", true)?;
                let mut file_options = OpenOptions::new();
                file_options.write(true);
                match long_arg(host, options, 1)? {
                    1 => {
                        file_options.create_new(true);
                    }
                    2 => {
                        file_options.create(true).truncate(true);
                    }
                    _ => return Err(fault(host, 3001)),
                }
                file_options
                    .open(&path)
                    .and_then(|mut file| file.write_all(&self.contents))
                    .map_err(|error| com_file_error(host, "ado_stream_save", &path, error))?;
                self.position = 0;
            }
            ("type" | "charset" | "read" | "readtext" | "write" | "writetext", _) => {
                return Err(fault(host, 3219));
            }
            _ => return Err(fault(host, 438)),
        }
        Ok(Variant::empty())
    }

    fn grow(&mut self, host: &StandardHostServices, length: usize) -> HalResult<()> {
        if length > i32::MAX as usize {
            return Err(fault(host, 3734));
        }
        if length > self.contents.len() {
            self.contents
                .try_reserve(length - self.contents.len())
                .map_err(|_| fault(host, 3734))?;
            self.contents.resize(length, 0);
        }
        Ok(())
    }

    fn write(&mut self, host: &StandardHostServices, data: &[u8]) -> HalResult<()> {
        let end = self
            .position
            .checked_add(data.len())
            .ok_or_else(|| fault(host, 3734))?;
        self.grow(host, end)?;
        self.contents[self.position..end].copy_from_slice(data);
        self.position = end;
        Ok(())
    }
}
