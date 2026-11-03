use crate::{Result, add, refuse};
use std::collections::{BTreeMap, BTreeSet};

#[derive(Debug, Clone)]
pub struct Section {
    pub name: String,
    pub rva: u32,
    pub virtual_size: u32,
    pub raw_offset: u32,
    pub raw_size: u32,
    pub characteristics: u32,
}
pub struct Image<'a> {
    pub(crate) bytes: &'a [u8],
    pub(crate) optional: usize,
    pub(crate) table: usize,
    pub section_alignment: u32,
    pub file_alignment: u32,
    pub header_size: u32,
    pub preferred_base: u64,
    pub sections: Vec<Section>,
}
impl<'a> Image<'a> {
    pub fn parse(bytes: &'a [u8]) -> Result<Self> {
        if span(bytes, 0, 2)? != b"MZ" {
            return Err(refuse("missing DOS signature"));
        }
        let pe = get32(bytes, 0x3c)? as usize;
        if span(bytes, pe, 4)? != b"PE\0\0" || get16(bytes, pe + 4)? != 0x8664 {
            return Err(refuse("expected AMD64 PE"));
        }
        let count = get16(bytes, pe + 6)? as usize;
        let optional_size = get16(bytes, pe + 20)? as usize;
        let optional = pe + 24;
        if optional_size < 240
            || get16(bytes, optional)? != 0x20b
            || get16(bytes, pe + 22)? & 0x2000 == 0
        {
            return Err(refuse("expected PE32+ DLL with complete optional header"));
        }
        span(bytes, optional, optional_size)?;
        if get32(bytes, optional + 108)? != 16 {
            return Err(refuse("expected sixteen data directories"));
        }
        let section_alignment = get32(bytes, optional + 32)?;
        let file_alignment = get32(bytes, optional + 36)?;
        if !(512..=65536).contains(&file_alignment)
            || !file_alignment.is_power_of_two()
            || !section_alignment.is_power_of_two()
            || section_alignment < 4096
            || section_alignment < file_alignment
        {
            return Err(refuse("unsupported section/file alignment"));
        }
        let header_size = get32(bytes, optional + 60)?;
        if header_size % file_alignment != 0 {
            return Err(refuse("unaligned PE headers"));
        }
        span(bytes, 0, header_size as usize)?;
        let table = optional + optional_size;
        if count == 0 || count > 96 || table + count * 40 > header_size as usize {
            return Err(refuse("invalid section table"));
        }
        let mut sections = Vec::<Section>::new();
        for i in 0..count {
            let at = table + i * 40;
            let name_bytes = span(bytes, at, 8)?;
            let name_len = name_bytes.iter().position(|b| *b == 0).unwrap_or(8);
            let name = std::str::from_utf8(&name_bytes[..name_len])
                .map_err(|_| refuse("invalid section name"))?
                .into();
            let section = Section {
                name,
                virtual_size: get32(bytes, at + 8)?,
                rva: get32(bytes, at + 12)?,
                raw_size: get32(bytes, at + 16)?,
                raw_offset: get32(bytes, at + 20)?,
                characteristics: get32(bytes, at + 36)?,
            };
            if section.rva < align(header_size, section_alignment)?
                || !section.rva.is_multiple_of(section_alignment)
                || !section.raw_size.is_multiple_of(file_alignment)
                || !section.raw_offset.is_multiple_of(file_alignment)
            {
                return Err(refuse("invalid/unaligned section range"));
            }
            if section.raw_size != 0 {
                if section.raw_offset < header_size {
                    return Err(refuse("section overlaps headers"));
                }
                span(
                    bytes,
                    section.raw_offset as usize,
                    section.raw_size as usize,
                )?;
            }
            let end = add(
                section.rva,
                section.virtual_size.max(section.raw_size) as usize,
            )?;
            if end > get32(bytes, optional + 56)? {
                return Err(refuse("section exceeds SizeOfImage"));
            }
            for previous in &sections {
                let previous_end = add(
                    previous.rva,
                    previous.virtual_size.max(previous.raw_size) as usize,
                )?;
                if section.rva < previous_end && previous.rva < end {
                    return Err(refuse("overlapping virtual sections"));
                }
                if section.name == previous.name {
                    return Err(refuse("duplicate section name"));
                }
                if section.raw_size != 0
                    && previous.raw_size != 0
                    && section.raw_offset < add(previous.raw_offset, previous.raw_size as usize)?
                    && previous.raw_offset < add(section.raw_offset, section.raw_size as usize)?
                {
                    return Err(refuse("overlapping raw sections"));
                }
            }
            sections.push(section);
        }
        let image = Self {
            bytes,
            optional,
            table,
            section_alignment,
            file_alignment,
            header_size,
            preferred_base: get64(bytes, optional + 24)?,
            sections,
        };
        for i in 0..16 {
            let (rva, size) = image.directory(i)?;
            if size != 0 {
                if i == 4 {
                    span(bytes, rva as usize, size as usize)?;
                } else {
                    image.at_rva(rva, size as usize)?;
                }
            }
        }
        Ok(image)
    }
    pub(crate) fn check_template(&self, finished: bool) -> Result<()> {
        if self.directory(4)?.1 != 0 {
            return Err(refuse("signed templates are unsupported"));
        }
        if get16(self.bytes, self.optional + 70)? & 0x4000 != 0 {
            return Err(refuse("CFG template is unsupported"));
        }
        // Current MSVC CRT objects retain CF_INSTRUMENTED (0x100) even with
        // /GUARD:NO. That marker alone does not enable CFG. Refuse an enabled
        // image, any actual guard table, or other unhandled guard features.
        let (load, size) = self.directory(10)?;
        if size >= 148 {
            let config = self.at_rva(load, size as usize)?;
            if get32(config, 144)? & !0x100 != 0
                || get64(config, 128)? != 0
                || get64(config, 136)? != 0
            {
                return Err(refuse("CFG load configuration is unsupported"));
            }
        }
        let reloc = self
            .sections
            .last()
            .ok_or_else(|| refuse("empty template"))?;
        if reloc.name != ".reloc" || self.directory(5)?.0 != reloc.rva {
            return Err(refuse("template requires final .reloc section"));
        }
        for previous in &self.sections[..self.sections.len() - 1] {
            if add(
                previous.rva,
                previous.virtual_size.max(previous.raw_size) as usize,
            )? > reloc.rva
                || add(previous.raw_offset, previous.raw_size as usize)? > reloc.raw_offset
            {
                return Err(refuse(".reloc must be last in file and memory"));
            }
        }
        if add(reloc.raw_offset, reloc.raw_size as usize)? as usize != self.bytes.len() {
            return Err(refuse("template overlay is unsupported"));
        }
        let needed = if finished { 4 } else { 5 };
        if self.table + (self.sections.len() + needed) * 40 > self.header_size as usize
            || self.sections.len() + needed > 96
        {
            return Err(refuse("reserve more section headers in developer link"));
        }
        if finished && self.directory(0)?.1 != 0 {
            return Err(refuse("finished template must have no public exports"));
        }
        if self.directory(5)?.1 == 0 {
            return Err(refuse("template must retain base relocations"));
        }
        self.runtime_functions()?;
        self.relocations()?;
        Ok(())
    }
    pub fn directory(&self, index: usize) -> Result<(u32, u32)> {
        if index >= 16 {
            return Err(refuse("invalid directory index"));
        }
        Ok((
            get32(self.bytes, self.optional + 112 + index * 8)?,
            get32(self.bytes, self.optional + 116 + index * 8)?,
        ))
    }
    pub fn at_rva(&self, rva: u32, len: usize) -> Result<&'a [u8]> {
        for s in &self.sections {
            if let Some(offset) = rva.checked_sub(s.rva)
                && add(offset, len)? <= s.raw_size
            {
                return span(
                    self.bytes,
                    add(s.raw_offset, offset as usize)? as usize,
                    len,
                );
            }
        }
        Err(refuse(format!("unmapped RVA {rva:#x} length {len}")))
    }
    pub fn section_data(&self, name: &str) -> Result<&'a [u8]> {
        let s = self
            .sections
            .iter()
            .find(|s| s.name == name)
            .ok_or_else(|| refuse(format!("missing section {name}")))?;
        self.at_rva(s.rva, s.virtual_size as usize)
    }
    pub fn executable_rva(&self, rva: u32) -> bool {
        self.sections.iter().any(|s| {
            s.characteristics & 0x20000000 != 0
                && rva >= s.rva
                && rva - s.rva < s.virtual_size.min(s.raw_size)
        })
    }
    fn string(&self, rva: u32) -> Result<&'a str> {
        let s = self
            .sections
            .iter()
            .find(|s| rva >= s.rva && rva - s.rva < s.raw_size)
            .ok_or_else(|| refuse("unmapped string"))?;
        let bytes = self.at_rva(rva, (s.raw_size - (rva - s.rva)) as usize)?;
        let end = bytes
            .iter()
            .position(|b| *b == 0)
            .ok_or_else(|| refuse("unterminated PE string"))?;
        std::str::from_utf8(&bytes[..end]).map_err(|_| refuse("invalid PE string"))
    }
    pub fn named_exports(&self) -> Result<BTreeMap<String, u32>> {
        let (rva, len) = self.directory(0)?;
        if len == 0 {
            return Ok(BTreeMap::new());
        }
        let dir = self.at_rva(rva, 40)?;
        let function_count = get32(dir, 20)? as usize;
        let n = get32(dir, 24)? as usize;
        let functions = self.at_rva(
            get32(dir, 28)?,
            function_count
                .checked_mul(4)
                .ok_or_else(|| refuse("export overflow"))?,
        )?;
        let names = self.at_rva(
            get32(dir, 32)?,
            n.checked_mul(4).ok_or_else(|| refuse("export overflow"))?,
        )?;
        let ordinals = self.at_rva(
            get32(dir, 36)?,
            n.checked_mul(2).ok_or_else(|| refuse("export overflow"))?,
        )?;
        let mut exports = BTreeMap::new();
        for i in 0..n {
            let name = self.string(get32(names, i * 4)?)?.to_owned();
            let index = get16(ordinals, i * 2)? as usize;
            if index >= function_count {
                return Err(refuse("export ordinal out of bounds"));
            }
            let target = get32(functions, index * 4)?;
            if target >= rva && target < add(rva, len as usize)? {
                return Err(refuse("forwarded exports are unsupported"));
            }
            if exports.insert(name, target).is_some() {
                return Err(refuse("duplicate PE export"));
            }
        }
        Ok(exports)
    }
    pub fn runtime_functions(&self) -> Result<Vec<[u32; 3]>> {
        let (rva, len) = self.directory(3)?;
        if len % 12 != 0 {
            return Err(refuse("truncated exception table"));
        }
        let bytes = if len == 0 {
            &[]
        } else {
            self.at_rva(rva, len as usize)?
        };
        let mut result: Vec<[u32; 3]> = Vec::new();
        for at in (0..len as usize).step_by(12) {
            let f = [
                get32(bytes, at)?,
                get32(bytes, at + 4)?,
                get32(bytes, at + 8)?,
            ];
            if f[0] >= f[1]
                || !self.executable_rva(f[0])
                || !self.executable_rva(f[1] - 1)
                || result.last().is_some_and(|prev| prev[1] > f[0])
            {
                return Err(refuse("invalid/unsorted exception table"));
            }
            self.at_rva(f[2], 4)?;
            result.push(f);
        }
        Ok(result)
    }
    pub fn relocations(&self) -> Result<BTreeSet<u32>> {
        let (rva, len) = self.directory(5)?;
        let bytes = if len == 0 {
            &[]
        } else {
            self.at_rva(rva, len as usize)?
        };
        let mut result = BTreeSet::new();
        let mut at = 0;
        while at < bytes.len() {
            let page = get32(bytes, at)?;
            let size = get32(bytes, at + 4)? as usize;
            if page % 4096 != 0 || size < 8 || !size.is_multiple_of(4) {
                return Err(refuse("invalid relocation block"));
            }
            let block = span(bytes, at, size)?;
            for off in (8..size).step_by(2) {
                let entry = get16(block, off)?;
                match entry >> 12 {
                    0 => {}
                    10 => {
                        let address = add(page, (entry & 0xfff) as usize)?;
                        self.at_rva(address, 8)?;
                        if !result.insert(address) {
                            return Err(refuse("duplicate relocation"));
                        }
                    }
                    _ => return Err(refuse("unsupported base relocation kind")),
                }
            }
            at += size;
        }
        Ok(result)
    }
    pub(crate) fn next_rva_without_reloc(&self) -> Result<u32> {
        let last = self
            .sections
            .iter()
            .rev()
            .find(|s| s.name != ".reloc")
            .ok_or_else(|| refuse("empty template"))?;
        align(
            add(last.rva, last.virtual_size.max(last.raw_size) as usize)?,
            self.section_alignment,
        )
    }
}
pub(crate) fn span(bytes: &[u8], at: usize, len: usize) -> Result<&[u8]> {
    bytes
        .get(
            at..at
                .checked_add(len)
                .ok_or_else(|| refuse("file offset overflow"))?,
        )
        .ok_or_else(|| refuse("truncated PE data"))
}
pub(crate) fn get16(bytes: &[u8], at: usize) -> Result<u16> {
    Ok(u16::from_le_bytes(
        span(bytes, at, 2)?.try_into().expect("checked width"),
    ))
}
pub(crate) fn get32(bytes: &[u8], at: usize) -> Result<u32> {
    Ok(u32::from_le_bytes(
        span(bytes, at, 4)?.try_into().expect("checked width"),
    ))
}
pub(crate) fn get64(bytes: &[u8], at: usize) -> Result<u64> {
    Ok(u64::from_le_bytes(
        span(bytes, at, 8)?.try_into().expect("checked width"),
    ))
}
pub(crate) fn put16(bytes: &mut [u8], at: usize, value: u16) {
    bytes[at..at + 2].copy_from_slice(&value.to_le_bytes());
}
pub(crate) fn put32(bytes: &mut [u8], at: usize, value: u32) {
    bytes[at..at + 4].copy_from_slice(&value.to_le_bytes());
}
pub(crate) fn align(value: u32, alignment: u32) -> Result<u32> {
    Ok(value
        .checked_add(alignment - 1)
        .ok_or_else(|| refuse("alignment overflow"))?
        & !(alignment - 1))
}
pub fn checksum(bytes: &[u8]) -> u32 {
    let mut sum = 0u64;
    for chunk in bytes.chunks(2) {
        sum += u64::from(u16::from_le_bytes([chunk[0], *chunk.get(1).unwrap_or(&0)]));
        sum = (sum & 0xffff) + (sum >> 16);
    }
    sum = (sum & 0xffff) + (sum >> 16);
    (sum as u32).wrapping_add(bytes.len() as u32)
}
