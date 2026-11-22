//! Developer-linked GNU/Linux x86-64 PIE and shared-object templates.
//! The finisher records helpers. The patcher appends mapped sections, keeps
//! all linked addresses fixed, and resolves exports with an ELF SysV hash.
use crate::{
    Error, ExportMap, Helper, Layout, Manifest, Patched, Program, Result,
    manifest::{abi_hash, helper_specs},
};
use object::{
    LittleEndian, Object, ObjectSection, ObjectSymbol, U16, U32, U64,
    elf::*,
    pod::{bytes_of_slice, from_bytes_mut, slice_from_bytes_mut},
    read::elf::ElfFile64,
};
use std::collections::{BTreeMap, BTreeSet};

const LE: LittleEndian = LittleEndian;
const PAGE: u32 = 4096;
const CONTRACT: &str = include_str!("../../../assets/native/helpers-linux-x64-v1.txt");
pub const FREE_BSTR: &str = "CalcVbaFreeBstr";
pub const HELPER_ABI: u64 = abi_hash(CONTRACT.as_bytes());
const ENTRIES: &[u8; 16] = b"CVBEXE_ENTRY_V1!";
type File<'a> = ElfFile64<'a, LittleEndian>;
type Header = FileHeader64<LittleEndian>;
type Segment = ProgramHeader64<LittleEndian>;
type Section = SectionHeader64<LittleEndian>;
type Symbol = Sym64<LittleEndian>;
type Dynamic = Dyn64<LittleEndian>;

fn error(message: impl std::fmt::Display) -> Error {
    Error(format!("ELF: {message}"))
}
fn u16le(n: u16) -> U16<LittleEndian> {
    U16::new(LE, n)
}
fn u32le(n: u32) -> U32<LittleEndian> {
    U32::new(LE, n)
}
fn u64le(n: u64) -> U64<LittleEndian> {
    U64::new(LE, n)
}
fn align(n: u32, alignment: u32) -> Result<u32> {
    n.checked_add(alignment - 1)
        .map(|n| n & !(alignment - 1))
        .ok_or_else(|| error("alignment overflow"))
}
fn at(n: u64) -> Result<usize> {
    usize::try_from(n).map_err(error)
}
fn end(start: u32, len: usize) -> Result<u32> {
    start
        .checked_add(u32::try_from(len).map_err(error)?)
        .ok_or_else(|| error("section address overflow"))
}
fn range(bytes: &[u8], offset: u64, size: u64) -> Result<&[u8]> {
    let offset = at(offset)?;
    let size = at(size)?;
    bytes
        .get(
            offset
                ..offset
                    .checked_add(size)
                    .ok_or_else(|| error("range overflow"))?,
        )
        .ok_or_else(|| error("range outside file"))
}
fn section(file: &File<'_>, name: &str) -> Result<usize> {
    file.section_by_name(name)
        .map(|section| section.index().0)
        .ok_or_else(|| error(format!("missing {name} section")))
}
fn executable(file: &File<'_>, address: u64) -> bool {
    file.elf_program_headers().iter().any(|p| {
        p.p_type.get(LE) == PT_LOAD
            && p.p_flags.get(LE) & PF_X != 0
            && address
                .checked_sub(p.p_vaddr.get(LE))
                .is_some_and(|offset| offset < p.p_filesz.get(LE))
    })
}
fn load_end(file: &File<'_>) -> Result<u32> {
    let mut end = 0;
    for p in file.elf_program_headers() {
        let kind = p.p_type.get(LE);
        let off = p.p_offset.get(LE);
        let addr = p.p_vaddr.get(LE);
        let size = p.p_filesz.get(LE);
        range(file.data(), off, size)?;
        if kind != PT_LOAD {
            continue;
        }
        let flags = p.p_flags.get(LE);
        let memory = p.p_memsz.get(LE);
        if (end == 0 && (addr != 0 || off != 0))
            || addr < u64::from(end)
            || addr % u64::from(PAGE) != off % u64::from(PAGE)
            || size > memory
            || p.p_align.get(LE) != u64::from(PAGE)
            || flags & (PF_W | PF_X) == PF_W | PF_X
        {
            return Err(error("unsupported load segment"));
        }
        end = align(
            u32::try_from(
                addr.checked_add(memory)
                    .ok_or_else(|| error("load overflow"))?,
            )
            .map_err(error)?,
            PAGE,
        )?;
    }
    if end == 0 {
        return Err(error("no load segments"));
    }
    Ok(end)
}
fn parse(bytes: &[u8]) -> Result<File<'_>> {
    let file = File::parse(bytes).map_err(error)?;
    let head = file.elf_header();
    if head.e_type.get(LE) != ET_DYN
        || head.e_machine.get(LE) != EM_X86_64
        || head.e_ehsize.get(LE) as usize != size_of::<Header>()
        || head.e_phentsize.get(LE) as usize != size_of::<Segment>()
        || head.e_shentsize.get(LE) as usize != size_of::<Section>()
        || head.e_shnum.get(LE) as usize != file.elf_section_table().iter().len()
    {
        return Err(error(
            "expected an x86-64 PIE or shared object with ELF64 headers",
        ));
    }
    load_end(&file)?;
    let entry = head.e_entry.get(LE);
    if is_library(&file) && entry != 0 || !is_library(&file) && !executable(&file, entry) {
        return Err(error("invalid shared-object or PIE entry"));
    }
    Ok(file)
}

fn is_library(file: &File<'_>) -> bool {
    !file
        .elf_program_headers()
        .iter()
        .any(|p| p.p_type.get(LE) == PT_INTERP)
}

/// Record the developer-linked helper addresses and hide them from dlsym.
pub fn finish_template(bytes: &[u8]) -> Result<Vec<u8>> {
    let file = parse(bytes)?;
    if file.section_by_name(".cvbhlp").is_some() {
        return Err(error("template already finished"));
    }
    let names: BTreeMap<_, _> = file
        .symbols()
        .filter(|symbol| symbol.is_definition())
        .filter_map(|symbol| {
            Some((
                symbol.name().ok()?,
                (symbol.address(), symbol.section_index()?),
            ))
        })
        .collect();
    let (addr, index) = names
        .get("cvb_template_abi")
        .ok_or_else(|| error("linked template has no runtime ABI hash"))?;
    let abi_section = file.section_by_index(*index).map_err(error)?;
    let abi_bytes = abi_section
        .data_range(*addr, 8)
        .map_err(error)?
        .ok_or_else(|| error("ABI hash outside section"))?;
    if u64::from_le_bytes(abi_bytes.try_into().map_err(error)?) != HELPER_ABI {
        return Err(error("linked template runtime ABI hash mismatch"));
    }
    let mut helpers = Vec::new();
    for spec in helper_specs(CONTRACT) {
        let (addr, _) = names
            .get(spec.name)
            .filter(|(addr, _)| executable(&file, *addr))
            .ok_or_else(|| error(format!("missing executable helper {}", spec.name)))?;
        helpers.push(Helper {
            name: spec.name.into(),
            version: spec.version,
            rva: u32::try_from(*addr).map_err(error)?,
        });
    }
    helpers.sort_by(|left, right| left.name.cmp(&right.name));
    let manifest = Manifest {
        abi: HELPER_ABI,
        helpers,
    };
    let mut writer = Writer::new(&file)?;
    if is_library(&file) {
        hide_helpers(&file, &mut writer, &manifest)?;
    }
    writer.add(".cvbhlp", &manifest.encode()?, 0, 0, 1)?;
    writer.finish()
}

fn cstr(table: &[u8], offset: u32) -> Result<&[u8]> {
    let tail = table
        .get(offset as usize..)
        .ok_or_else(|| error("invalid symbol name offset"))?;
    let len = tail
        .iter()
        .position(|byte| *byte == 0)
        .ok_or_else(|| error("unterminated symbol name"))?;
    Ok(&tail[..len])
}

fn sysv_hash(symbols: &[Symbol], strings: &[u8], buckets: usize) -> Result<Vec<u8>> {
    let count = u32::try_from(symbols.len()).map_err(error)?;
    let mut table = vec![0u32; 2 + buckets + symbols.len()];
    table[0] = u32::try_from(buckets).map_err(error)?;
    table[1] = count;
    for (index, symbol) in symbols.iter().enumerate().skip(1) {
        if symbol.st_name.get(LE) == 0 {
            continue;
        }
        let name = cstr(strings, symbol.st_name.get(LE))?;
        let bucket = classic_hash(name) as usize % buckets + 2;
        table[2 + buckets + index] = table[bucket];
        table[bucket] = index as u32;
    }
    Ok(table.into_iter().flat_map(u32::to_le_bytes).collect())
}

fn classic_hash(name: &[u8]) -> u32 {
    let mut hash = 0u32;
    for &byte in name {
        hash = (hash << 4).wrapping_add(byte.into());
        hash ^= (hash & 0xf000_0000) >> 24;
        hash &= 0x0fff_ffff;
    }
    hash
}

struct DynamicTables<'a> {
    hash: usize,
    symbols: usize,
    strings: usize,
    versions: usize,
    dynamic: usize,
    str_data: &'a [u8],
    sym_data: &'a [Symbol],
    ver_data: &'a [u8],
    buckets: usize,
}
impl<'a> DynamicTables<'a> {
    fn parse(file: &File<'a>, writer: &Writer) -> Result<Self> {
        let hash = section(file, ".hash")?;
        let symbols = section(file, ".dynsym")?;
        let strings = section(file, ".dynstr")?;
        let versions = section(file, ".gnu.version")?;
        let dynamic = section(file, ".dynamic")?;
        let dynamic_section = &writer.sections[dynamic];
        let mut segments = file
            .elf_program_headers()
            .iter()
            .filter(|p| p.p_type.get(LE) == PT_DYNAMIC);
        let segment = segments.next().ok_or_else(|| error("missing PT_DYNAMIC"))?;
        if segments.next().is_some()
            || segment.p_offset != dynamic_section.sh_offset
            || segment.p_vaddr != dynamic_section.sh_addr
            || segment.p_filesz != dynamic_section.sh_size
        {
            return Err(error("PT_DYNAMIC does not match .dynamic"));
        }
        let str_data = writer.original(file, strings)?;
        let sym_data = file.elf_dynamic_symbol_table().symbols();
        let ver_data = writer.original(file, versions)?;
        let hash_data = writer.original(file, hash)?;
        let buckets = crate::image::get32(hash_data, 0)? as usize;
        if buckets == 0
            || crate::image::get32(hash_data, 4)? as usize != sym_data.len()
            || hash_data.len() != (2 + buckets + sym_data.len()) * 4
            || ver_data.len() != sym_data.len() * 2
            || file.section_by_name(".gnu.hash").is_some()
        {
            return Err(error("invalid SysV dynamic symbol tables"));
        }
        Ok(Self {
            hash,
            symbols,
            strings,
            versions,
            dynamic,
            str_data,
            sym_data,
            ver_data,
            buckets,
        })
    }
}

fn hide_helpers(file: &File<'_>, writer: &mut Writer, manifest: &Manifest) -> Result<()> {
    let tables = DynamicTables::parse(file, writer)?;
    let mut definitions = BTreeSet::from([b"cvb_template_abi".as_slice()]);
    definitions.extend(manifest.helpers.iter().map(|helper| helper.name.as_bytes()));
    let mut symbols = tables.sym_data.to_vec();
    let mut hidden = BTreeSet::new();
    for (index, symbol) in symbols.iter_mut().enumerate() {
        if symbol.st_shndx.get(LE) != SHN_UNDEF {
            let name = cstr(tables.str_data, symbol.st_name.get(LE))?;
            // The Rust runtime can retain other no_mangle symbols. They are
            // implementation details too, and must never enter the public ABI.
            definitions.remove(name);
            hidden.insert(index);
            symbol.st_name = u32le(0);
            symbol.st_shndx = u16le(SHN_UNDEF);
            symbol.st_value = u64le(0);
            symbol.st_size = u64le(0);
        }
    }
    if !definitions.is_empty() {
        return Err(error("helper missing from linked dynamic symbols"));
    }
    // The developer link binds private definitions locally. If a dynamic
    // relocation still references a helper, stripping that definition is unsafe.
    for (_, relocation) in file.dynamic_relocations().into_iter().flatten() {
        if let object::RelocationTarget::Symbol(index) = relocation.target()
            && hidden.contains(&index.0)
        {
            return Err(error("dynamic relocation references a private helper"));
        }
    }
    let hash = sysv_hash(&symbols, tables.str_data, tables.buckets)?;
    writer.replace(tables.hash, &hash)?;
    writer.replace(tables.symbols, bytes_of_slice(&symbols))?;
    Ok(())
}

pub struct Template<'a> {
    file: File<'a>,
    pub manifest: Manifest,
}
impl<'a> Template<'a> {
    pub fn parse(bytes: &'a [u8]) -> Result<Self> {
        let file = parse(bytes)?;
        let data = file
            .section_by_name(".cvbhlp")
            .ok_or_else(|| error("expected a finished template"))?
            .data()
            .map_err(error)?;
        let manifest = Manifest::decode(data)?;
        if manifest.abi != HELPER_ABI
            || manifest
                .helpers
                .iter()
                .any(|helper| !executable(&file, helper.rva.into()))
        {
            return Err(error("invalid helper ABI or address"));
        }
        if file.section_by_name(".cvbtext").is_some() {
            return Err(error("template already contains a compiled program"));
        }
        Ok(Self { file, manifest })
    }
    pub fn is_library(&self) -> bool {
        is_library(&self.file)
    }
    pub fn layout_with_data(&self, text_len: usize, data_len: usize) -> Result<Layout> {
        let text_rva = load_end(&self.file)?;
        let data_rva = align(end(text_rva, text_len.max(1))?, PAGE)?;
        let read_only_rva = align(end(data_rva, data_len.max(1))?, PAGE)?;
        Ok(Layout {
            text_rva,
            data_rva,
            read_only_rva,
        })
    }
    pub fn patch(&self, program: &Program, startup: &str) -> Result<Patched> {
        if !program.data_relocations.is_empty() || !program.unwind.is_empty() {
            return Err(error("unexpected PE relocation or unwind records"));
        }
        if program.text.is_empty() || program.exports.is_empty() {
            return Err(error("program needs code and named exports"));
        }
        let layout = self.layout_with_data(program.text.len(), program.data.len())?;
        let mut writer = Writer::new(&self.file)?;
        let mut segments = self.file.elf_program_headers().to_vec();
        segments.retain(|p| p.p_type.get(LE) != PT_PHDR);
        let free_bstr = if self.is_library() {
            Some((FREE_BSTR, self.manifest.resolve("cvb_free_bstr", 1)?))
        } else {
            None
        };
        let map = crate::export_map(program, layout.text_rva, free_bstr)?;
        let offset = writer.pad(PAGE as usize)?;
        writer.add(
            ".cvbtext",
            &program.text,
            layout.text_rva,
            SHF_ALLOC | SHF_EXECINSTR,
            PAGE,
        )?;
        segments.push(segment(
            PT_LOAD,
            PF_R | PF_X,
            offset,
            layout.text_rva,
            program.text.len(),
            PAGE,
        ));
        let offset = writer.pad(PAGE as usize)?;
        // Give the writable segment a mapped page even for raw code-only programs.
        let data = if program.data.is_empty() {
            &[0]
        } else {
            &program.data[..]
        };
        writer.add(
            ".cvbdata",
            data,
            layout.data_rva,
            SHF_ALLOC | SHF_WRITE,
            PAGE,
        )?;
        segments.push(segment(
            PT_LOAD,
            PF_R | PF_W,
            offset,
            layout.data_rva,
            data.len(),
            PAGE,
        ));

        let ro = writer.pad(PAGE as usize)?;
        writer.add(
            ".cvbrdat",
            &program.read_only,
            layout.read_only_rva,
            SHF_ALLOC,
            PAGE,
        )?;
        let image_addr = writer.address(ro, layout.read_only_rva, 8)?;
        writer.add(".cvbimg", &program.image, image_addr, SHF_ALLOC, 8)?;
        if self.is_library() {
            let tables = DynamicTables::parse(&self.file, &writer)?;
            writer.add_exports(&tables, &map, ro, layout.read_only_rva)?;
        } else {
            writer.startup(&self.file, &map, startup)?;
        }
        if !program.elf_unwind.is_empty() {
            writer.add_unwind(&self.file, program, &mut segments, ro, layout)?;
        }
        let phdr = writer.pad(8)?;
        let phdr_addr = end(layout.read_only_rva, phdr - ro)?;
        let count = segments.len() + 2;
        let size = count * size_of::<Segment>();
        let ro_size = phdr - ro + size;
        end(layout.read_only_rva, ro_size)?;
        segments.insert(0, segment(PT_PHDR, PF_R, phdr, phdr_addr, size, 8));
        segments.push(segment(
            PT_LOAD,
            PF_R,
            ro,
            layout.read_only_rva,
            ro_size,
            PAGE,
        ));
        writer.bytes.extend_from_slice(bytes_of_slice(&segments));
        let header = writer.header()?;
        header.e_phoff = u64le(phdr as u64);
        header.e_phnum = u16le(u16::try_from(count).map_err(error)?);
        Ok(Patched {
            bytes: writer.finish()?,
            map,
        })
    }
}

fn delta(target: u32, base: u32) -> Result<i32> {
    i32::try_from(i64::from(target) - i64::from(base)).map_err(error)
}

fn segment(
    kind: u32,
    flags: u32,
    offset: usize,
    addr: u32,
    size: usize,
    alignment: u32,
) -> Segment {
    Segment {
        p_type: u32le(kind),
        p_flags: u32le(flags),
        p_offset: u64le(offset as u64),
        p_vaddr: u64le(addr.into()),
        p_paddr: u64le(addr.into()),
        p_filesz: u64le(size as u64),
        p_memsz: u64le(size as u64),
        p_align: u64le(alignment.into()),
    }
}

struct Writer {
    bytes: Vec<u8>,
    sections: Vec<Section>,
    shstr: Vec<u8>,
    shstr_index: usize,
}
impl Writer {
    fn new(file: &File<'_>) -> Result<Self> {
        let mut writer = Self {
            bytes: file.data().to_vec(),
            sections: file.elf_section_table().iter().copied().collect(),
            shstr: Vec::new(),
            shstr_index: file.elf_header().e_shstrndx.get(LE) as usize,
        };
        writer.shstr = writer.original(file, writer.shstr_index)?.to_vec();
        Ok(writer)
    }
    fn original<'a>(&self, file: &File<'a>, index: usize) -> Result<&'a [u8]> {
        let section = self
            .sections
            .get(index)
            .ok_or_else(|| error("invalid section index"))?;
        range(
            file.data(),
            section.sh_offset.get(LE),
            section.sh_size.get(LE),
        )
    }
    fn header(&mut self) -> Result<&mut Header> {
        from_bytes_mut(&mut self.bytes)
            .map(|(head, _)| head)
            .map_err(|_| error("invalid ELF header"))
    }
    fn pad(&mut self, alignment: usize) -> Result<usize> {
        let offset = self
            .bytes
            .len()
            .checked_next_multiple_of(alignment)
            .ok_or_else(|| error("file offset overflow"))?;
        self.bytes.resize(offset, 0);
        Ok(offset)
    }
    fn address(&mut self, start: usize, addr: u32, alignment: usize) -> Result<u32> {
        end(addr, self.pad(alignment)? - start)
    }
    fn add(
        &mut self,
        name: &str,
        data: &[u8],
        addr: u32,
        flags: u32,
        alignment: u32,
    ) -> Result<usize> {
        let index = self.sections.len();
        let name_offset = u32::try_from(self.shstr.len()).map_err(error)?;
        self.shstr.extend_from_slice(name.as_bytes());
        self.shstr.push(0);
        let at = self.pad(alignment as usize)?;
        self.sections.push(Section {
            sh_name: u32le(name_offset),
            sh_type: u32le(SHT_PROGBITS),
            sh_flags: u64le(flags.into()),
            sh_addr: u64le(addr.into()),
            sh_offset: u64le(at as u64),
            sh_size: u64le(data.len() as u64),
            sh_link: u32le(0),
            sh_info: u32le(0),
            sh_addralign: u64le(alignment.into()),
            sh_entsize: u64le(0),
        });
        self.bytes.extend_from_slice(data);
        Ok(index)
    }
    fn section_bytes(&mut self, index: usize) -> Result<&mut [u8]> {
        let section = &self.sections[index];
        let offset = at(section.sh_offset.get(LE))?;
        let len = at(section.sh_size.get(LE))?;
        let end = offset
            .checked_add(len)
            .ok_or_else(|| error("section bounds overflow"))?;
        self.bytes
            .get_mut(offset..end)
            .ok_or_else(|| error("section bytes outside file"))
    }
    fn replace(&mut self, index: usize, data: &[u8]) -> Result<()> {
        let destination = self.section_bytes(index)?;
        if destination.len() != data.len() {
            return Err(error("unexpected section size"));
        }
        destination.copy_from_slice(data);
        Ok(())
    }
    fn relocate(&mut self, index: usize, data: &[u8], start: usize, base: u32) -> Result<u32> {
        let alignment = at(self.sections[index].sh_addralign.get(LE))?.max(1);
        if !alignment.is_power_of_two() || alignment > PAGE as usize {
            return Err(error("invalid dynamic table alignment"));
        }
        let offset = self.pad(alignment)?;
        let addr = end(base, offset - start)?;
        let section = &mut self.sections[index];
        section.sh_addr = u64le(addr.into());
        section.sh_offset = u64le(offset as u64);
        section.sh_size = u64le(data.len() as u64);
        self.bytes.extend_from_slice(data);
        Ok(addr)
    }
    fn add_exports(
        &mut self,
        tables: &DynamicTables<'_>,
        map: &ExportMap,
        ro: usize,
        base: u32,
    ) -> Result<()> {
        let mut symbols = tables.sym_data.to_vec();
        let mut strings = tables.str_data.to_vec();
        let mut versions = tables.ver_data.to_vec();
        if symbols
            .iter()
            .any(|symbol| symbol.st_shndx.get(LE) != SHN_UNDEF)
        {
            return Err(error(
                "finished shared object still has public runtime definitions",
            ));
        }
        for (name, _, address) in map {
            let text = self
                .sections
                .iter()
                .position(|section| {
                    section.sh_flags.get(LE) & u64::from(SHF_EXECINSTR) != 0
                        && u64::from(*address)
                            .checked_sub(section.sh_addr.get(LE))
                            .is_some_and(|offset| offset < section.sh_size.get(LE))
                })
                .ok_or_else(|| error("dynamic export address outside a code section"))?;
            symbols.push(Symbol {
                st_name: u32le(u32::try_from(strings.len()).map_err(error)?),
                st_info: (STB_GLOBAL << 4) | STT_FUNC,
                st_other: STV_DEFAULT,
                st_shndx: u16le(u16::try_from(text).map_err(error)?),
                st_value: u64le((*address).into()),
                st_size: u64le(0),
            });
            strings.extend_from_slice(name.as_bytes());
            strings.push(0);
            versions.extend_from_slice(&VER_NDX_GLOBAL.to_le_bytes());
        }
        let hash = sysv_hash(&symbols, &strings, (symbols.len() / 2).max(1))?;
        let string_size = u32::try_from(strings.len()).map_err(error)?;
        let updates = [
            (DT_HASH, self.relocate(tables.hash, &hash, ro, base)?),
            (
                DT_SYMTAB,
                self.relocate(tables.symbols, bytes_of_slice(&symbols), ro, base)?,
            ),
            (
                DT_STRTAB,
                self.relocate(tables.strings, &strings, ro, base)?,
            ),
            (DT_STRSZ, string_size),
            (
                DT_VERSYM,
                self.relocate(tables.versions, &versions, ro, base)?,
            ),
        ];
        let data = self.section_bytes(tables.dynamic)?;
        let count = data.len() / size_of::<Dynamic>();
        let (entries, remainder) = slice_from_bytes_mut::<Dynamic>(data, count)
            .map_err(|_| error("unsupported dynamic entries"))?;
        if !remainder.is_empty() {
            return Err(error("invalid dynamic size"));
        }
        for (tag, value) in updates {
            let mut matches = entries
                .iter_mut()
                .take_while(|d| d.d_tag.get(LE) != u64::from(DT_NULL))
                .filter(|d| d.d_tag.get(LE) == u64::from(tag));
            let entry = matches
                .next()
                .ok_or_else(|| error(format!("missing dynamic tag {tag}")))?;
            if matches.next().is_some() {
                return Err(error("duplicate dynamic tag"));
            }
            entry.d_val = u64le(value.into());
        }
        Ok(())
    }
    fn add_unwind(
        &mut self,
        file: &File<'_>,
        program: &Program,
        segments: &mut [Segment],
        ro: usize,
        layout: Layout,
    ) -> Result<()> {
        let header = section(file, ".eh_frame_hdr")?;
        let old_addr = u32::try_from(self.sections[header].sh_addr.get(LE)).map_err(error)?;
        let original = self.original(file, header)?;
        // Version 1, pcrel/sdata4 pointer, udata4 count, datarel/sdata4 table.
        const FORMAT: [u8; 4] = [1, 0x1b, 0x03, 0x3b];
        let count = crate::image::get32(original, 8)? as usize;
        if original.get(..4) != Some(&FORMAT) || original.len() != 12 + count * 8 {
            return Err(error("unsupported .eh_frame_hdr encoding"));
        }
        let read_i32 = |bytes: &[u8], at| crate::image::get32(bytes, at).map(|v| v as i32);
        let frame = old_addr
            .checked_add(4)
            .and_then(|base| base.checked_add_signed(read_i32(original, 4).ok()?))
            .ok_or_else(|| error("invalid .eh_frame pointer"))?;
        let mut entries = Vec::with_capacity(count);
        for pair in original[12..].chunks_exact(8) {
            let address = |offset| -> Result<u32> {
                old_addr
                    .checked_add_signed(read_i32(pair, offset)?)
                    .ok_or_else(|| error("invalid unwind index address"))
            };
            entries.push((address(0)?, address(4)?));
        }
        let frame_addr = self.address(ro, layout.read_only_rva, 8)?;
        let mut encoded = program.elf_unwind.clone();
        let mut pos = 0;
        while pos < encoded.len() {
            let len = crate::image::get32(&encoded, pos)? as usize;
            let record = pos
                .checked_add(4 + len)
                .filter(|end| *end <= encoded.len())
                .ok_or_else(|| error("truncated generated unwind entry"))?;
            if len == 0 && record == encoded.len() {
                break;
            }
            if len < 4 {
                return Err(error("invalid generated unwind entry length"));
            }
            let cie = crate::image::get32(&encoded, pos + 4)? as usize;
            if cie != 0 {
                if len < 12 || cie > pos + 4 {
                    return Err(error("invalid generated FDE"));
                }
                let field = pos + 8;
                let ip = u32::try_from(field)
                    .map_err(error)?
                    .checked_add_signed(read_i32(&encoded, field)?)
                    .ok_or_else(|| error("invalid generated unwind address"))?;
                let length = crate::image::get32(&encoded, field + 4)?;
                if ip < layout.text_rva
                    || length == 0
                    || ip
                        .checked_add(length)
                        .is_none_or(|last| last > layout.text_rva + program.text.len() as u32)
                {
                    return Err(error("generated unwind range outside code"));
                }
                let relative = delta(ip, end(frame_addr, field)?)?;
                encoded[field..field + 4].copy_from_slice(&relative.to_le_bytes());
                entries.push((ip, end(frame_addr, pos)?));
            }
            pos = record;
        }
        self.add(".cvbframe", &encoded, frame_addr, SHF_ALLOC, 8)?;
        let header_addr = self.address(ro, layout.read_only_rva, 4)?;
        entries.sort_unstable_by_key(|entry| entry.0);
        let mut new = Vec::from(FORMAT);
        new.extend_from_slice(&delta(frame, end(header_addr, 4)?)?.to_le_bytes());
        new.extend_from_slice(&u32::try_from(entries.len()).map_err(error)?.to_le_bytes());
        for (ip, fde) in entries {
            new.extend_from_slice(&delta(ip, header_addr)?.to_le_bytes());
            new.extend_from_slice(&delta(fde, header_addr)?.to_le_bytes());
        }
        let offset = self.bytes.len();
        if self.relocate(header, &new, ro, layout.read_only_rva)? != header_addr {
            return Err(error("unexpected unwind header alignment"));
        }
        let mut phdrs = segments
            .iter_mut()
            .filter(|p| p.p_type.get(LE) == PT_GNU_EH_FRAME);
        let phdr = phdrs
            .next()
            .ok_or_else(|| error("missing PT_GNU_EH_FRAME"))?;
        if phdrs.next().is_some() {
            return Err(error("duplicate PT_GNU_EH_FRAME"));
        }
        *phdr = segment(PT_GNU_EH_FRAME, PF_R, offset, header_addr, new.len(), 4);
        Ok(())
    }
    fn startup(&mut self, file: &File<'_>, map: &ExportMap, startup: &str) -> Result<()> {
        let index = section(file, ".cvbentry")?;
        let section = &self.sections[index];
        let addr = section.sh_addr.get(LE);
        let mapped = file.elf_program_headers().iter().any(|p| {
            p.p_type.get(LE) == PT_LOAD
                && p.p_flags.get(LE) == PF_R
                && addr.checked_sub(p.p_vaddr.get(LE)).is_some_and(|offset| {
                    p.p_offset.get(LE).checked_add(offset) == Some(section.sh_offset.get(LE))
                        && offset
                            .checked_add(section.sh_size.get(LE))
                            .is_some_and(|end| end <= p.p_filesz.get(LE))
                })
        });
        if !mapped {
            return Err(error(
                "executable startup record is outside read-only load data",
            ));
        }
        let data = self.section_bytes(index)?;
        if data.len() != 28 || &data[..16] != ENTRIES || data[16..] != [0; 12] {
            return Err(error("invalid executable startup record"));
        }
        for (cell, name) in
            data[16..]
                .chunks_exact_mut(4)
                .zip([startup, "CalcVbaShutdown", "CalcVbaGetLastError"])
        {
            let (_, _, target) = map
                .iter()
                .find(|(export, _, _)| export == name)
                .ok_or_else(|| error(format!("missing startup entry {name}")))?;
            let delta = i64::from(*target) - i64::try_from(addr).map_err(error)?;
            cell.copy_from_slice(&i32::try_from(delta).map_err(error)?.to_le_bytes());
        }
        Ok(())
    }
    fn finish(mut self) -> Result<Vec<u8>> {
        let section = &mut self.sections[self.shstr_index];
        section.sh_offset = u64le(self.bytes.len() as u64);
        section.sh_size = u64le(self.shstr.len() as u64);
        self.bytes.extend_from_slice(&self.shstr);
        let offset = self.pad(8)?;
        self.bytes.extend_from_slice(bytes_of_slice(&self.sections));
        let count = u16::try_from(self.sections.len()).map_err(error)?;
        if count >= SHN_LORESERVE {
            return Err(error("too many sections"));
        }
        let header = self.header()?;
        header.e_shoff = u64le(offset as u64);
        header.e_shnum = u16le(count);
        parse(&self.bytes)?;
        Ok(self.bytes)
    }
}
