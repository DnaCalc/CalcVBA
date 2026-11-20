//! `OxImage` — the linkable set of `OxProgram`s for one VBA *project closure*, and the vm3
//! runtime artifact (`.oxi`).
//!
//! A whole VBA project is exactly one [`OxProgram`]; cross-project references add more, with the
//! entry (the COM-server / startup project) LAST — the same ordering [`crate`]'s consumer
//! `calcvba_vm3::Vm3::link` uses. This is the vm3 runtime artifact: a build emits it (by
//! elaborating each project's `CoreProgram`) for the in-process COM server to load and link.
//! `OxProgram` is fully serde-serializable, so the image round-trips through JSON.

use crate::program::OxProgram;
use serde::{Deserialize, Serialize};

/// On-disk format tag for an `OxImage` artifact (rejects a stale/foreign blob via its header).
pub const OX_IMAGE_FORMAT: &str = "oxvba.ox-image";
/// On-disk schema version.
pub const OX_IMAGE_VERSION: u32 = 3;

/// Largest `.oxi` blob [`OxImage::from_bytes`] decodes. Decoded size is proportional to the
/// input, so this bounds the allocation made for untrusted package metadata.
pub const MAX_OX_IMAGE_BYTES: usize = 256 * 1024 * 1024;
/// Most programs (projects in one closure) an image may carry.
pub const MAX_OX_IMAGE_PROGRAMS: usize = 4096;

/// A linkable image: the elaborated programs of a project closure plus the entry index.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub struct OxImage {
    pub format: String,
    pub version: u32,
    /// Index into `programs` of the entry project (whose `Main`/exported classes the host
    /// drives). The cross-project entry is LAST, matching `Vm3::link`.
    pub entry: usize,
    pub programs: Vec<OxProgram>,
}

/// Errors building or loading an [`OxImage`].
#[derive(Debug)]
pub enum OxImageError {
    Serialize(serde_json::Error),
    Deserialize(serde_json::Error),
    UnknownFormat {
        format: String,
    },
    UnsupportedVersion {
        version: u32,
    },
    Empty,
    /// The blob is larger than [`MAX_OX_IMAGE_BYTES`]; nothing was decoded.
    TooLarge {
        len: usize,
        max: usize,
    },
    /// More programs than [`MAX_OX_IMAGE_PROGRAMS`].
    TooManyPrograms {
        count: usize,
        max: usize,
    },
    /// A local, temp or global type names a record layout the program does not carry.
    BadRecordLayout {
        program: usize,
        layout: usize,
        layouts: usize,
    },
    EntryOutOfRange {
        entry: usize,
        count: usize,
    },
    /// A contained program failed structural verification (e.g. an out-of-range
    /// entry/func/block reference in a corrupt or foreign `.oxi`).
    VerificationFailed {
        program: usize,
        errors: Vec<crate::verify::VerifyError>,
    },
}

impl std::fmt::Display for OxImageError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            OxImageError::Serialize(e) => write!(f, "OxImage serialize: {e}"),
            OxImageError::Deserialize(e) => write!(f, "OxImage deserialize: {e}"),
            OxImageError::UnknownFormat { format } => {
                write!(f, "not an OxImage artifact: format `{format}`")
            }
            OxImageError::UnsupportedVersion { version } => {
                write!(f, "unsupported OxImage version {version}")
            }
            OxImageError::Empty => write!(f, "OxImage has no programs"),
            OxImageError::TooLarge { len, max } => {
                write!(f, "OxImage of {len} bytes exceeds the {max}-byte limit")
            }
            OxImageError::TooManyPrograms { count, max } => {
                write!(f, "OxImage has {count} programs (limit {max})")
            }
            OxImageError::BadRecordLayout {
                program,
                layout,
                layouts,
            } => write!(
                f,
                "OxImage program {program} names record layout {layout} of {layouts}"
            ),
            OxImageError::EntryOutOfRange { entry, count } => {
                write!(f, "OxImage entry {entry} out of range ({count} programs)")
            }
            OxImageError::VerificationFailed { program, errors } => {
                write!(f, "OxImage program {program} failed verification:")?;
                for e in errors {
                    write!(f, " {e};")?;
                }
                Ok(())
            }
        }
    }
}

impl std::error::Error for OxImageError {}

impl OxImage {
    /// Build an image from the elaborated programs of a closure, with the entry as the LAST
    /// program (the closure ordering the build and `Vm3::link` use). An empty `programs` yields
    /// `entry = 0`, which [`validate`](Self::validate) then rejects.
    pub fn new(programs: Vec<OxProgram>) -> Self {
        let entry = programs.len().saturating_sub(1);
        OxImage {
            format: OX_IMAGE_FORMAT.to_string(),
            version: OX_IMAGE_VERSION,
            entry,
            programs,
        }
    }

    /// Serialize to the `.oxi` byte form (pretty JSON).
    pub fn to_bytes(&self) -> Result<Vec<u8>, OxImageError> {
        serde_json::to_vec_pretty(self).map_err(OxImageError::Serialize)
    }

    /// Deserialize + validate + structurally verify an `.oxi` blob. Verification
    /// is a trust boundary: a corrupt or foreign `.oxi` with a valid header but
    /// out-of-range entry/func/block references would otherwise panic the host
    /// (an OOB index) when a frame is built, so every contained program is run
    /// through [`verify_program`](crate::verify::verify_program) here.
    pub fn from_bytes(bytes: &[u8]) -> Result<Self, OxImageError> {
        Self::from_bytes_within(bytes, MAX_OX_IMAGE_BYTES)
    }

    fn from_bytes_within(bytes: &[u8], max: usize) -> Result<Self, OxImageError> {
        if bytes.len() > max {
            return Err(OxImageError::TooLarge {
                len: bytes.len(),
                max,
            });
        }
        // serde_json refuses nesting deeper than its recursion limit, so a recursive or
        // malicious type/record graph fails here deterministically instead of on the stack.
        let image: Self = serde_json::from_slice(bytes).map_err(OxImageError::Deserialize)?;
        image.validate()?;
        for (index, program) in image.programs.iter().enumerate() {
            check_record_layout_refs(index, program)?;
            crate::verify::verify_program(program).map_err(|errors| {
                OxImageError::VerificationFailed {
                    program: index,
                    errors,
                }
            })?;
        }
        Ok(image)
    }

    /// Reject a stale/foreign blob, an empty image, or an out-of-range entry.
    pub fn validate(&self) -> Result<(), OxImageError> {
        if self.format != OX_IMAGE_FORMAT {
            return Err(OxImageError::UnknownFormat {
                format: self.format.clone(),
            });
        }
        if self.version != OX_IMAGE_VERSION {
            return Err(OxImageError::UnsupportedVersion {
                version: self.version,
            });
        }
        if self.programs.is_empty() {
            return Err(OxImageError::Empty);
        }
        if self.programs.len() > MAX_OX_IMAGE_PROGRAMS {
            return Err(OxImageError::TooManyPrograms {
                count: self.programs.len(),
                max: MAX_OX_IMAGE_PROGRAMS,
            });
        }
        if self.entry >= self.programs.len() {
            return Err(OxImageError::EntryOutOfRange {
                entry: self.entry,
                count: self.programs.len(),
            });
        }
        Ok(())
    }

    /// The program references in link order (entry last), for `Vm3::link`.
    pub fn program_refs(&self) -> Vec<&OxProgram> {
        self.programs.iter().collect()
    }
}

/// Every `Record` type of a program's locals, temps and globals names one of its layouts.
fn check_record_layout_refs(index: usize, program: &OxProgram) -> Result<(), OxImageError> {
    fn layout_of(ty: &crate::ty::OxTy) -> Option<usize> {
        match ty {
            crate::ty::OxTy::Record(id) => Some(id.0),
            crate::ty::OxTy::Array(element, _) => layout_of(element),
            _ => None,
        }
    }
    let layouts = program.record_layouts.len();
    let types =
        program
            .globals
            .iter()
            .map(|global| &global.ty)
            .chain(program.funcs.iter().flat_map(|func| {
                func.locals
                    .iter()
                    .map(|local| &local.ty)
                    .chain(func.temps.iter())
            }));
    for ty in types {
        if let Some(layout) = layout_of(ty).filter(|layout| *layout >= layouts) {
            return Err(OxImageError::BadRecordLayout {
                program: index,
                layout,
                layouts,
            });
        }
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn bounded_decoding_rejects_oversized_truncated_and_nested_blobs() {
        let bytes = OxImage::new(vec![OxProgram::default()])
            .to_bytes()
            .expect("serialize");
        assert!(matches!(
            OxImage::from_bytes_within(&bytes, bytes.len() - 1),
            Err(OxImageError::TooLarge { len, max }) if len == bytes.len() && max == len - 1
        ));
        assert!(OxImage::from_bytes_within(&bytes, bytes.len()).is_ok());
        assert!(matches!(
            OxImage::from_bytes(&bytes[..bytes.len() / 2]),
            Err(OxImageError::Deserialize(_))
        ));
        // A deeply nested blob fails in the decoder, not on the stack.
        let nested = format!("{}{}", "[".repeat(100_000), "]".repeat(100_000));
        assert!(matches!(
            OxImage::from_bytes(nested.as_bytes()),
            Err(OxImageError::Deserialize(_))
        ));
    }

    #[test]
    fn bounded_decoding_rejects_too_many_programs() {
        let image = OxImage::new(vec![OxProgram::default(); MAX_OX_IMAGE_PROGRAMS + 1]);
        let bytes = image.to_bytes().expect("serialize");
        assert!(matches!(
            OxImage::from_bytes(&bytes),
            Err(OxImageError::TooManyPrograms { count, max })
                if count == MAX_OX_IMAGE_PROGRAMS + 1 && max == MAX_OX_IMAGE_PROGRAMS
        ));
    }

    #[test]
    fn bounded_decoding_rejects_a_record_type_without_its_layout() {
        let mut program = OxProgram::default();
        program.globals.push(crate::program::OxGlobal {
            name: "g".into(),
            ty: crate::ty::OxTy::Array(
                Box::new(crate::ty::OxTy::Record(crate::ty::RecordLayoutId(2))),
                crate::ty::ArrayShape::Dynamic,
            ),
            array_element: None,
        });
        program.record_layouts.push(Vec::new());
        let bytes = OxImage::new(vec![program]).to_bytes().expect("serialize");
        assert!(matches!(
            OxImage::from_bytes(&bytes),
            Err(OxImageError::BadRecordLayout {
                program: 0,
                layout: 2,
                layouts: 1
            })
        ));
    }

    #[test]
    fn round_trips_through_bytes() {
        let image = OxImage::new(vec![OxProgram::default(), OxProgram::default()]);
        assert_eq!(image.entry, 1, "entry is the last program");
        let bytes = image.to_bytes().expect("serialize");
        let back = OxImage::from_bytes(&bytes).expect("deserialize");
        assert_eq!(image, back);
        assert_eq!(back.program_refs().len(), 2);
    }

    #[test]
    fn from_bytes_rejects_out_of_range_program_entry() {
        // A .oxi with a valid header but a program whose `entry` FuncId is out of
        // range used to pass load (only header/count/entry were checked) and then
        // panic when a frame was built. It must now be rejected at load.
        let program = OxProgram {
            entry: Some(crate::FuncId(5)), // no funcs -> out of range
            ..OxProgram::default()
        };
        let bytes = OxImage::new(vec![program]).to_bytes().expect("serialize");
        assert!(
            matches!(
                OxImage::from_bytes(&bytes),
                Err(OxImageError::VerificationFailed { program: 0, .. })
            ),
            "a corrupt program.entry must be rejected at load, not reach the runtime"
        );
    }

    #[test]
    fn rejects_empty_and_bad_header() {
        assert!(matches!(
            OxImage::new(Vec::new()).validate(),
            Err(OxImageError::Empty)
        ));
        let mut bad = OxImage::new(vec![OxProgram::default()]);
        bad.format = "nope".into();
        assert!(matches!(
            OxImage::from_bytes(&serde_json::to_vec(&bad).unwrap()),
            Err(OxImageError::UnknownFormat { .. })
        ));
    }
}
