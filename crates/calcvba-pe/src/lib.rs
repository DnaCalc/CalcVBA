//! Checked PE32+ template finishing and program patching. No compiler or linker
//! process is used here. Existing linked code/data RVAs are never moved.

mod image;
mod manifest;

pub use image::{Image, Section};
pub use manifest::{HELPER_ABI, Helper, HelperSpec, Manifest, shipped_helper_specs};
use std::collections::{BTreeMap, BTreeSet};

#[derive(Debug, thiserror::Error)]
#[error("native PE: {0}")]
pub struct Error(pub String);
pub type Result<T> = std::result::Result<T, Error>;
fn refuse(message: impl Into<String>) -> Error {
    Error(message.into())
}

/// Generated code is position independent. Offsets below are relative to text.
#[derive(Debug, Clone)]
pub struct Export {
    pub name: String,
    pub ordinal: Option<u16>,
    pub offset: u32,
}
#[derive(Debug, Clone)]
pub struct Unwind {
    pub begin: u32,
    pub end: u32,
    /// Windows x64 UNWIND_INFO emitted by the code generator.
    pub info: Vec<u8>,
}
#[derive(Debug, Default)]
pub struct Program {
    pub text: Vec<u8>,
    pub read_only: Vec<u8>,
    pub data: Vec<u8>,
    pub image: Vec<u8>,
    pub exports: Vec<Export>,
    pub unwind: Vec<Unwind>,
    /// Offsets into data holding preferred-base absolute pointers (DIR64).
    pub data_relocations: Vec<u32>,
}
#[derive(Debug, Clone, Copy)]
pub struct Layout {
    pub text_rva: u32,
    pub read_only_rva: u32,
    pub data_rva: u32,
}
#[derive(Debug)]
pub struct Patched {
    pub bytes: Vec<u8>,
    pub map: ExportMap,
}
pub type ExportMap = Vec<(String, u16, u32)>;

/// Developer-time only: resolve helper exports into a manifest, remove the
/// public export directory, and verify that the linked DLL is patchable.
pub fn finish_template(bytes: &[u8], specs: &[HelperSpec<'_>]) -> Result<Vec<u8>> {
    let image = Image::parse(bytes)?;
    image.check_template(false)?;
    let exports = image.named_exports()?;
    let abi = exports
        .get("cvb_template_abi")
        .ok_or_else(|| refuse("template has no embedded runtime ABI hash"))?;
    if image::get64(image.at_rva(*abi, 8)?, 0)? != HELPER_ABI {
        return Err(refuse("linked template runtime ABI hash mismatch"));
    }
    let mut helpers = Vec::new();
    let mut seen = BTreeSet::new();
    for spec in specs {
        if !seen.insert(spec.name) {
            return Err(refuse("duplicate helper"));
        }
        let rva = *exports
            .get(spec.name)
            .ok_or_else(|| refuse(format!("missing helper {}", spec.name)))?;
        if !image.executable_rva(rva) {
            return Err(refuse("helper is not executable code"));
        }
        helpers.push(Helper {
            name: spec.name.into(),
            version: spec.version,
            rva,
        });
    }
    helpers.sort_by(|a, b| a.name.cmp(&b.name));
    let manifest = Manifest {
        abi: HELPER_ABI,
        helpers,
    }
    .encode()?;
    let mut writer = Writer::new(image)?;
    writer.add(".cvbhlp", &manifest, 0x40000040)?;
    writer.set_directory(0, 0, 0);
    writer.finish()
}

pub struct Template<'a> {
    image: Image<'a>,
    pub manifest: Manifest,
}
impl<'a> Template<'a> {
    pub fn is_dll(&self) -> bool {
        self.image.is_dll
    }
    pub fn parse(bytes: &'a [u8], expected_abi: u64) -> Result<Self> {
        let image = Image::parse(bytes)?;
        image.check_template(true)?;
        let manifest = Manifest::decode(image.section_data(".cvbhlp")?)?;
        if manifest.abi != expected_abi {
            return Err(refuse("runtime helper ABI mismatch"));
        }
        for helper in &manifest.helpers {
            if !image.executable_rva(helper.rva) {
                return Err(refuse("invalid helper RVA"));
            }
        }
        Ok(Self { image, manifest })
    }
    /// Read-only constants are placed before export/unwind metadata, so their
    /// RVA is known before compilation; later sections depend on final sizes.
    pub fn layout(&self, text_len: usize) -> Result<Layout> {
        self.layout_with_data(text_len, 0)
    }
    pub fn layout_with_data(&self, text_len: usize, data_len: usize) -> Result<Layout> {
        let text_rva = self.image.next_rva_without_reloc()?;
        let data_rva = image::align(
            add(text_rva, text_len.max(1))?,
            self.image.section_alignment,
        )?;
        let read_only_rva = image::align(
            add(data_rva, data_len.max(1))?,
            self.image.section_alignment,
        )?;
        Ok(Layout {
            text_rva,
            read_only_rva,
            data_rva,
        })
    }
    pub fn text_rva(&self) -> Result<u32> {
        self.image.next_rva_without_reloc()
    }
    pub fn patch(&self, dll_name: &str, program: &Program) -> Result<Patched> {
        valid_name(dll_name)?;
        if program.text.is_empty() || program.exports.is_empty() {
            return Err(refuse("program needs code and named exports"));
        }
        let mut writer = Writer::new(Image::parse(self.image.bytes)?)?;
        let text = writer.add(".cvbtext", &program.text, 0x60000020)?;
        // Put mutable slots before the variable-sized export/unwind tables.
        // Both data RVAs are then known before resolving code relocations.
        let data_rva = writer.add(".cvbdata", &program.data, 0xc0000040)?;
        let rdata_rva = writer.next_rva()?;
        let mut rdata = program.read_only.clone();
        pad(&mut rdata, 4);
        let exports_offset = rdata.len();
        let (exports, map) =
            export_directory(dll_name, text, program, add(rdata_rva, exports_offset)?)?;
        rdata.extend_from_slice(&exports);
        writer.set_directory(0, add(rdata_rva, exports_offset)?, exports.len() as u32);
        let mut functions = self.image.runtime_functions()?;
        for unwind in &program.unwind {
            if unwind.begin >= unwind.end || unwind.end as usize > program.text.len() {
                return Err(refuse("invalid generated function range"));
            }
            validate_unwind(&unwind.info)?;
            pad(&mut rdata, 4);
            let info_rva = add(rdata_rva, rdata.len())?;
            rdata.extend_from_slice(&unwind.info);
            functions.push([
                add(text, unwind.begin as usize)?,
                add(text, unwind.end as usize)?,
                info_rva,
            ]);
        }
        functions.sort_by_key(|f| f[0]);
        for pair in functions.windows(2) {
            if pair[0][1] > pair[1][0] {
                return Err(refuse("overlapping runtime functions"));
            }
        }
        pad(&mut rdata, 4);
        let pdata = add(rdata_rva, rdata.len())?;
        for function in &functions {
            for value in function {
                rdata.extend_from_slice(&value.to_le_bytes());
            }
        }
        writer.set_directory(3, pdata, (functions.len() * 12) as u32);
        writer.add(".cvbrdat", &rdata, 0x40000040)?;
        for offset in &program.data_relocations {
            if (*offset as usize)
                .checked_add(8)
                .is_none_or(|end| end > program.data.len())
            {
                return Err(refuse("DIR64 outside generated data"));
            }
            writer.relocations.insert(add(data_rva, *offset as usize)?);
        }
        writer.add(".cvbimg", &program.image, 0x40000040)?;
        Ok(Patched {
            bytes: writer.finish()?,
            map,
        })
    }
}

/// Fill the executable startup's fixed compiled entries. All targets must be
/// executable RVAs in this image; a DLL or an already-filled startup is refused.
pub fn patch_executable_entries(
    bytes: &mut [u8],
    map: &[(String, u16, u32)],
    main: &str,
) -> Result<()> {
    let (offset, entries, checksum) = {
        let image = Image::parse(bytes)?;
        if image.is_dll {
            return Err(refuse("executable startup requires an EXE template"));
        }
        let section = image
            .sections
            .iter()
            .find(|s| s.name == ".rdata")
            .ok_or_else(|| refuse("executable template has no startup data"))?;
        let data = image.section_data(".rdata")?;
        let matches: Vec<_> = data
            .windows(16)
            .enumerate()
            .filter(|(_, word)| *word == b"CVBEXE_ENTRY_V1!")
            .map(|(at, _)| at)
            .collect();
        if matches.len() != 1 {
            return Err(refuse(
                "executable template must contain one startup record",
            ));
        }
        let at = matches[0] + 16;
        let mut entries = Vec::new();
        for (index, name) in [main, "CalcVbaShutdown", "CalcVbaGetLastError"]
            .iter()
            .enumerate()
        {
            if image::get32(data, at + index * 4)? != 0 {
                return Err(refuse("executable startup was already filled"));
            }
            let rva = map
                .iter()
                .find(|(export, _, _)| export == name)
                .map(|(_, _, rva)| *rva)
                .ok_or_else(|| refuse(format!("missing executable entry {name}")))?;
            if !image.executable_rva(rva) {
                return Err(refuse("executable startup target is not code"));
            }
            entries.push(rva);
        }
        (
            section.raw_offset as usize + at,
            entries,
            image.optional + 64,
        )
    };
    for (index, rva) in entries.into_iter().enumerate() {
        image::put32(bytes, offset + index * 4, rva);
    }
    image::put32(bytes, checksum, 0);
    Ok(())
}

fn validate_unwind(info: &[u8]) -> Result<()> {
    // Initial backend emits version 1 without exception handlers/chained info.
    if info.len() < 4 || info[0] != 1 || info.len() != (4 + usize::from(info[2]) * 2 + 3) & !3 {
        return Err(refuse("unsupported/truncated generated UNWIND_INFO"));
    }
    Ok(())
}
fn valid_name(name: &str) -> Result<()> {
    if name.is_empty() || !name.is_ascii() || name.bytes().any(|b| b == 0) {
        return Err(refuse(
            "export/DLL names must be nonempty ASCII without NUL",
        ));
    }
    Ok(())
}
fn export_directory(
    dll: &str,
    text: u32,
    program: &Program,
    rva: u32,
) -> Result<(Vec<u8>, ExportMap)> {
    let mut names = BTreeMap::new();
    let mut used = BTreeSet::new();
    for export in &program.exports {
        valid_name(&export.name)?;
        if export.offset as usize >= program.text.len() {
            return Err(refuse("export outside generated code"));
        }
        if names
            .insert(
                &export.name,
                (export.ordinal, add(text, export.offset as usize)?),
            )
            .is_some()
        {
            return Err(refuse("duplicate export name"));
        }
        if let Some(ordinal) = export.ordinal
            && (ordinal == 0 || !used.insert(ordinal))
        {
            return Err(refuse("duplicate/zero export ordinal"));
        }
    }
    let mut next = 1u16;
    let mut map = Vec::new();
    for (name, (ordinal, address)) in names {
        let ordinal = if let Some(ordinal) = ordinal {
            ordinal
        } else {
            while used.contains(&next) {
                next = next
                    .checked_add(1)
                    .ok_or_else(|| refuse("export ordinal space exhausted"))?;
            }
            used.insert(next);
            next
        };
        map.push((name.clone(), ordinal, address));
    }
    let base = *used.first().expect("nonempty exports");
    let count = usize::from(*used.last().expect("nonempty exports") - base) + 1;
    let n = map.len();
    let mut bytes = vec![0; 40 + count * 4 + n * 6];
    let dll_offset = bytes.len();
    bytes.extend_from_slice(dll.as_bytes());
    bytes.push(0);
    image::put32(&mut bytes, 12, add(rva, dll_offset)?);
    image::put32(&mut bytes, 16, u32::from(base));
    image::put32(&mut bytes, 20, count as u32);
    image::put32(&mut bytes, 24, n as u32);
    image::put32(&mut bytes, 28, add(rva, 40)?);
    image::put32(&mut bytes, 32, add(rva, 40 + count * 4)?);
    image::put32(&mut bytes, 36, add(rva, 40 + count * 4 + n * 4)?);
    for (i, (name, ordinal, address)) in map.iter().enumerate() {
        image::put32(&mut bytes, 40 + usize::from(ordinal - base) * 4, *address);
        let name_rva = add(rva, bytes.len())?;
        image::put32(&mut bytes, 40 + count * 4 + i * 4, name_rva);
        image::put16(&mut bytes, 40 + count * 4 + n * 4 + i * 2, ordinal - base);
        bytes.extend_from_slice(name.as_bytes());
        bytes.push(0);
    }
    Ok((bytes, map))
}
fn add(rva: u32, len: usize) -> Result<u32> {
    rva.checked_add(u32::try_from(len).map_err(|_| refuse("image exceeds RVA range"))?)
        .ok_or_else(|| refuse("image RVA overflow"))
}
fn pad(bytes: &mut Vec<u8>, alignment: usize) {
    bytes.resize((bytes.len() + alignment - 1) & !(alignment - 1), 0);
}

struct Writer {
    bytes: Vec<u8>,
    sections: Vec<Section>,
    optional: usize,
    table: usize,
    section_alignment: u32,
    file_alignment: u32,
    header_limit: u32,
    relocations: BTreeSet<u32>,
}
impl Writer {
    fn new(image: Image<'_>) -> Result<Self> {
        let relocations = image.relocations()?;
        let mut sections = image.sections.clone();
        let reloc = sections
            .pop()
            .ok_or_else(|| refuse("missing relocation section"))?;
        if reloc.name != ".reloc" {
            return Err(refuse("relocation section must be last"));
        }
        let mut bytes = image.bytes.to_vec();
        bytes.truncate(reloc.raw_offset as usize);
        let mut writer = Self {
            bytes,
            sections,
            optional: image.optional,
            table: image.table,
            section_alignment: image.section_alignment,
            file_alignment: image.file_alignment,
            header_limit: image.header_size,
            relocations,
        };
        writer.set_directory(5, 0, 0);
        Ok(writer)
    }
    fn next_rva(&self) -> Result<u32> {
        let last = self
            .sections
            .last()
            .ok_or_else(|| refuse("empty template"))?;
        image::align(
            add(last.rva, last.virtual_size.max(last.raw_size) as usize)?,
            self.section_alignment,
        )
    }
    fn set_directory(&mut self, index: usize, rva: u32, len: u32) {
        image::put32(&mut self.bytes, self.optional + 112 + index * 8, rva);
        image::put32(&mut self.bytes, self.optional + 116 + index * 8, len);
    }
    fn add(&mut self, name: &str, data: &[u8], characteristics: u32) -> Result<u32> {
        if self.sections.iter().any(|s| s.name == name) {
            return Err(refuse("duplicate generated section"));
        }
        if self.table + (self.sections.len() + 1) * 40 > self.header_limit as usize {
            return Err(refuse("template has insufficient section headers"));
        }
        let rva = self.next_rva()?;
        let offset = image::align(
            u32::try_from(self.bytes.len()).map_err(|_| refuse("file too large"))?,
            self.file_alignment,
        )?;
        let size = u32::try_from(data.len().max(1)).map_err(|_| refuse("section too large"))?;
        let raw_size = image::align(size, self.file_alignment)?;
        self.bytes.resize(offset as usize, 0);
        self.bytes.extend_from_slice(data);
        self.bytes
            .resize(add(offset, raw_size as usize)? as usize, 0);
        self.sections.push(Section {
            name: name.into(),
            rva,
            virtual_size: size,
            raw_offset: offset,
            raw_size,
            characteristics,
        });
        Ok(rva)
    }
    fn finish(mut self) -> Result<Vec<u8>> {
        let mut blocks = BTreeMap::<u32, Vec<u16>>::new();
        for rva in &self.relocations {
            blocks
                .entry(rva & !0xfff)
                .or_default()
                .push(0xa000 | (*rva & 0xfff) as u16);
        }
        let mut reloc = Vec::new();
        for (page, mut entries) in blocks {
            if entries.len() % 2 != 0 {
                entries.push(0);
            }
            reloc.extend_from_slice(&page.to_le_bytes());
            reloc.extend_from_slice(&(8 + entries.len() as u32 * 2).to_le_bytes());
            for entry in entries {
                reloc.extend_from_slice(&entry.to_le_bytes());
            }
        }
        let reloc_rva = self.add(".reloc", &reloc, 0x42000040)?;
        self.set_directory(5, reloc_rva, reloc.len() as u32);
        for (i, s) in self.sections.iter().enumerate() {
            let at = self.table + i * 40;
            self.bytes[at..at + 40].fill(0);
            self.bytes[at..at + s.name.len()].copy_from_slice(s.name.as_bytes());
            for (off, value) in [
                (8, s.virtual_size),
                (12, s.rva),
                (16, s.raw_size),
                (20, s.raw_offset),
                (36, s.characteristics),
            ] {
                image::put32(&mut self.bytes, at + off, value);
            }
        }
        image::put16(
            &mut self.bytes,
            self.optional - 18,
            self.sections.len() as u16,
        );
        let image_size = self.next_rva()?;
        image::put32(&mut self.bytes, self.optional + 56, image_size);
        let code = self
            .sections
            .iter()
            .filter(|s| s.characteristics & 0x20 != 0)
            .map(|s| u64::from(s.raw_size))
            .sum::<u64>();
        let data = self
            .sections
            .iter()
            .filter(|s| s.characteristics & 0x40 != 0)
            .map(|s| u64::from(s.raw_size))
            .sum::<u64>();
        image::put32(
            &mut self.bytes,
            self.optional + 4,
            u32::try_from(code).map_err(|_| refuse("code size overflow"))?,
        );
        image::put32(
            &mut self.bytes,
            self.optional + 8,
            u32::try_from(data).map_err(|_| refuse("data size overflow"))?,
        );
        image::put32(&mut self.bytes, self.optional + 64, 0);
        let checksum = image::checksum(&self.bytes);
        image::put32(&mut self.bytes, self.optional + 64, checksum);
        Image::parse(&self.bytes)?;
        Ok(self.bytes)
    }
}
