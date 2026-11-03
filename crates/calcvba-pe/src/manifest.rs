use crate::{Result, refuse};

/// Identifies the initial helper signatures, not a PE build timestamp. Change
/// this when any helper signature/semantics changes incompatibly.
pub const HELPER_ABI: u64 = abi_hash(include_bytes!("../../../assets/native/helpers-v1.txt"));
const fn abi_hash(bytes: &[u8]) -> u64 {
    let mut hash = 0xcbf29ce484222325u64;
    let mut at = 0;
    while at < bytes.len() {
        // Ignore CR so a Windows checkout has the same contract hash.
        if bytes[at] != b'\r' {
            hash = (hash ^ bytes[at] as u64).wrapping_mul(0x100000001b3);
        }
        at += 1;
    }
    hash
}
const MAGIC: &[u8; 8] = b"CVBHLP01";
pub struct HelperSpec<'a> {
    pub name: &'a str,
    pub version: u16,
}
pub fn shipped_helper_specs() -> Vec<HelperSpec<'static>> {
    include_str!("../../../assets/native/helpers-v1.txt")
        .lines()
        .skip(1)
        .map(|line| {
            let mut fields = line.split('|');
            HelperSpec {
                name: fields.next().expect("helper contract name"),
                version: fields
                    .next()
                    .expect("helper contract version")
                    .parse()
                    .expect("helper ABI version"),
            }
        })
        .collect()
}
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Helper {
    pub name: String,
    pub version: u16,
    pub rva: u32,
}
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Manifest {
    pub abi: u64,
    pub helpers: Vec<Helper>,
}
impl Manifest {
    pub fn resolve(&self, name: &str, version: u16) -> Result<u32> {
        self.helpers
            .iter()
            .find(|h| h.name == name && h.version == version)
            .map(|h| h.rva)
            .ok_or_else(|| refuse(format!("missing/incompatible helper {name} v{version}")))
    }
    pub(crate) fn encode(&self) -> Result<Vec<u8>> {
        let mut bytes = MAGIC.to_vec();
        bytes.extend_from_slice(&self.abi.to_le_bytes());
        bytes.extend_from_slice(&(self.helpers.len() as u32).to_le_bytes());
        for h in &self.helpers {
            crate::valid_name(&h.name)?;
            bytes.extend_from_slice(
                &u16::try_from(h.name.len())
                    .map_err(|_| refuse("helper name too long"))?
                    .to_le_bytes(),
            );
            bytes.extend_from_slice(&h.version.to_le_bytes());
            bytes.extend_from_slice(&h.rva.to_le_bytes());
            bytes.extend_from_slice(h.name.as_bytes());
        }
        Ok(bytes)
    }
    pub fn decode(bytes: &[u8]) -> Result<Self> {
        use crate::image::{get16, get32, get64, span};
        if span(bytes, 0, 8)? != MAGIC {
            return Err(refuse("invalid helper manifest magic"));
        }
        let abi = get64(bytes, 8)?;
        let count = get32(bytes, 16)? as usize;
        if count > bytes.len() / 8 {
            return Err(refuse("invalid helper count"));
        }
        let mut helpers = Vec::new();
        let mut at = 20;
        for _ in 0..count {
            let len = get16(bytes, at)? as usize;
            let version = get16(bytes, at + 2)?;
            let rva = get32(bytes, at + 4)?;
            let name = std::str::from_utf8(span(bytes, at + 8, len)?)
                .map_err(|_| refuse("invalid helper name"))?
                .to_owned();
            crate::valid_name(&name)?;
            if helpers.last().is_some_and(|h: &Helper| h.name >= name) {
                return Err(refuse("unsorted/duplicate helper manifest"));
            }
            helpers.push(Helper { name, version, rva });
            at += 8 + len;
        }
        if bytes.get(at..).is_some_and(|b| b.iter().any(|v| *v != 0)) {
            return Err(refuse("unexpected manifest trailing bytes"));
        }
        Ok(Self { abi, helpers })
    }
}
