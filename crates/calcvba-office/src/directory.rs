//! MS-OVBA 2.3.4.2 version-independent records. Lengths are checked before reads.
use crate::{
    Error, Module, ModuleKind, ProjectProperties, Reference, ReferenceKind, Result, invalid,
};

struct Reader<'a> {
    bytes: &'a [u8],
    at: usize,
    cp: u16,
}
impl<'a> Reader<'a> {
    fn take(&mut self, n: usize) -> Result<&'a [u8]> {
        let end = self
            .at
            .checked_add(n)
            .filter(|end| *end <= self.bytes.len())
            .ok_or_else(|| invalid(format!("truncated dir stream at {}", self.at)))?;
        let value = &self.bytes[self.at..end];
        self.at = end;
        Ok(value)
    }
    fn u16(&mut self) -> Result<u16> {
        Ok(u16::from_le_bytes(self.take(2)?.try_into().unwrap()))
    }
    fn u32(&mut self) -> Result<u32> {
        Ok(u32::from_le_bytes(self.take(4)?.try_into().unwrap()))
    }
    fn peek(&self) -> Result<u16> {
        self.bytes
            .get(self.at..self.at + 2)
            .map(|b| u16::from_le_bytes([b[0], b[1]]))
            .ok_or_else(|| invalid("missing dir record"))
    }
    fn id(&mut self, expected: u16) -> Result<()> {
        let actual = self.u16()?;
        if actual != expected {
            return Err(invalid(format!(
                "dir record {actual:#06x}, expected {expected:#06x}"
            )));
        }
        Ok(())
    }
    fn sized(&mut self) -> Result<&'a [u8]> {
        let n = self.u32()? as usize;
        self.take(n)
    }
    fn record(&mut self, id: u16) -> Result<&'a [u8]> {
        self.id(id)?;
        self.sized()
    }
    fn fixed(&mut self, id: u16, size: u32) -> Result<()> {
        self.id(id)?;
        if self.u32()? != size {
            return Err(invalid(format!("dir record {id:#06x} size")));
        }
        Ok(())
    }
    fn long(&mut self, id: u16) -> Result<u32> {
        self.fixed(id, 4)?;
        self.u32()
    }
    fn short(&mut self, id: u16) -> Result<u16> {
        self.fixed(id, 2)?;
        self.u16()
    }
    fn text(&mut self, id: u16) -> Result<String> {
        let cp = self.cp;
        crate::decode(self.record(id)?, cp)
    }
    fn sized_text(&mut self) -> Result<String> {
        let cp = self.cp;
        crate::decode(self.sized()?, cp)
    }
    fn unicode(&mut self, id: u16) -> Result<String> {
        let bytes = self.record(id)?;
        if bytes.len() % 2 != 0 {
            return Err(invalid("odd UTF-16 dir record"));
        }
        String::from_utf16(
            &bytes
                .as_chunks::<2>()
                .0
                .iter()
                .map(|b| u16::from_le_bytes([b[0], b[1]]))
                .collect::<Vec<_>>(),
        )
        .map_err(|_| invalid("invalid UTF-16 dir record"))
    }
    fn pair(&mut self, mbcs: u16, unicode: u16) -> Result<String> {
        let ansi = self.text(mbcs)?;
        let wide = self.unicode(unicode)?;
        // Unicode names can represent a name unavailable in the legacy code page.
        Ok(if wide.is_empty() { ansi } else { wide })
    }
    fn reference_name(&mut self) -> Result<Option<String>> {
        if self.peek()? == 0x16 {
            Ok(Some(self.pair(0x16, 0x3e)?))
        } else {
            Ok(None)
        }
    }
}

pub(crate) struct ParsedDirectory {
    pub properties: ProjectProperties,
    pub references: Vec<Reference>,
    pub modules: Vec<Module>,
    // Decoded against PROJECT after its codepage is known.
    pub help_file_bytes: Vec<u8>,
}

pub(crate) fn read(bytes: &[u8], module_limit: usize) -> Result<ParsedDirectory> {
    let mut r = Reader {
        bytes,
        at: 0,
        cp: 1252,
    };
    let sys_kind = r.long(1)?;
    if sys_kind > 3 {
        return Err(invalid("PROJECTSYSKIND"));
    }
    let compat_version = if r.peek()? == 0x4a {
        Some(r.long(0x4a)?)
    } else {
        None
    };
    let lcid = r.long(2)?;
    let lcid_invoke = r.long(0x14)?;
    let codepage = r.short(3)?;
    if codepage::to_encoding(codepage).is_none() {
        return Err(Error::Codepage(codepage));
    }
    r.cp = codepage;
    let name = r.text(4)?;
    let description = r.pair(5, 0x40)?;
    let help_file_bytes = r.record(6)?.to_vec();
    if r.record(0x3d)? != help_file_bytes {
        return Err(invalid("project HelpFile copies differ"));
    }
    let help_context = r.long(7)?;
    let lib_flags = r.long(8)?;
    r.fixed(9, 4)?;
    let version_major = r.u32()?;
    let version_minor = r.u16()?;
    let constants = if r.peek()? == 0xc {
        r.pair(0xc, 0x3c)?
    } else {
        String::new()
    };
    let props = ProjectProperties {
        sys_kind,
        compat_version,
        lcid,
        lcid_invoke,
        codepage,
        name,
        description,
        help_file: String::new(),
        help_context,
        lib_flags,
        version_major,
        version_minor,
        constants,
    };
    let mut references = Vec::new();
    while r.peek()? != 0xf {
        let name = r.reference_name()?.unwrap_or_default();
        let kind = match r.u16()? {
            0xd => {
                let _size = r.u32()?;
                let libid = r.sized_text()?;
                r.take(6)?;
                ReferenceKind::Registered { libid }
            }
            0xe => {
                let _size = r.u32()?;
                let absolute = r.sized_text()?;
                let relative = r.sized_text()?;
                let major = r.u32()?;
                let minor = r.u16()?;
                ReferenceKind::Project {
                    absolute,
                    relative,
                    major,
                    minor,
                }
            }
            id @ (0x33 | 0x2f) => {
                let original = if id == 0x33 {
                    let value = r.sized_text()?;
                    r.id(0x2f)?;
                    Some(value)
                } else {
                    None
                };
                let _size = r.u32()?;
                let twiddled = r.sized_text()?;
                r.take(6)?;
                let extended_name = r.reference_name()?;
                r.id(0x30)?;
                let _size = r.u32()?;
                let extended = r.sized_text()?;
                r.take(6)?;
                let guid = r.take(16)?.try_into().unwrap();
                let cookie = r.u32()?;
                ReferenceKind::Control {
                    original,
                    twiddled,
                    extended_name,
                    extended,
                    guid,
                    cookie,
                }
            }
            id => {
                return Err(invalid(format!(
                    "unsupported dir reference record {id:#06x}"
                )));
            }
        };
        references.push(Reference { name, kind });
    }
    let count = usize::from(r.short(0xf)?);
    if count > module_limit {
        return Err(Error::Limit("module count"));
    }
    r.short(0x13)?; // project cookie
    let mut modules = Vec::with_capacity(count);
    for _ in 0..count {
        let mut name = r.text(0x19)?;
        if r.peek()? == 0x47 {
            let unicode = r.unicode(0x47)?;
            if !unicode.is_empty() {
                name = unicode;
            }
        }
        let stream_name = r.pair(0x1a, 0x32)?;
        let description = r.pair(0x1c, 0x48)?;
        let text_offset = r.long(0x31)?;
        let help_context = r.long(0x1e)?;
        r.short(0x2c)?;
        let kind = match r.u16()? {
            0x21 => ModuleKind::Standard,
            0x22 => ModuleKind::Class,
            _ => return Err(invalid("MODULETYPE")),
        };
        r.u32()?;
        let read_only = r.peek()? == 0x25;
        if read_only {
            r.fixed(0x25, 0)?;
        }
        let private = r.peek()? == 0x28;
        if private {
            r.fixed(0x28, 0)?;
        }
        r.fixed(0x2b, 0)?;
        modules.push(Module {
            name,
            stream_name,
            kind,
            description,
            help_context,
            read_only,
            private,
            text_offset,
            source: String::new(),
            document_version: None,
        });
    }
    r.fixed(0x10, 0)?;
    if bytes[r.at..].iter().any(|b| *b != 0) {
        return Err(invalid("trailing dir bytes"));
    }
    Ok(ParsedDirectory {
        properties: props,
        references,
        modules,
        help_file_bytes,
    })
}

struct Writer {
    bytes: Vec<u8>,
    cp: u16,
}
impl Writer {
    fn u16(&mut self, n: u16) {
        self.bytes.extend_from_slice(&n.to_le_bytes());
    }
    fn u32(&mut self, n: u32) {
        self.bytes.extend_from_slice(&n.to_le_bytes());
    }
    fn record(&mut self, id: u16, bytes: &[u8]) -> Result<()> {
        self.u16(id);
        self.sized(bytes)
    }
    fn sized(&mut self, bytes: &[u8]) -> Result<()> {
        self.u32(
            bytes
                .len()
                .try_into()
                .map_err(|_| Error::Limit("dir record"))?,
        );
        self.bytes.extend_from_slice(bytes);
        Ok(())
    }
    fn long(&mut self, id: u16, n: u32) -> Result<()> {
        self.record(id, &n.to_le_bytes())
    }
    fn short(&mut self, id: u16, n: u16) -> Result<()> {
        self.record(id, &n.to_le_bytes())
    }
    fn text(&mut self, id: u16, text: &str) -> Result<()> {
        self.record(id, &crate::encode(text, self.cp)?)
    }
    fn sized_text(&mut self, text: &str) -> Result<()> {
        self.sized(&crate::encode(text, self.cp)?)
    }
    fn pair(&mut self, mbcs: u16, unicode: u16, text: &str) -> Result<()> {
        self.text(mbcs, text)?;
        self.record(
            unicode,
            &text
                .encode_utf16()
                .flat_map(u16::to_le_bytes)
                .collect::<Vec<_>>(),
        )
    }
    fn reference_name(&mut self, name: &str) -> Result<()> {
        if !name.is_empty() {
            self.pair(0x16, 0x3e, name)?;
        }
        Ok(())
    }
}

// MS-OVBA 2.3.4.2.1.12 stores signed numeric project arguments, not the
// arbitrary constants allowed by source-level #Const statements.
fn project_constants(text: &str, cp: u16) -> Result<String> {
    let mut names = std::collections::BTreeSet::new();
    let mut result = Vec::new();
    for item in text.split(':').map(str::trim).filter(|s| !s.is_empty()) {
        let values = crate::parse_constants(item)
            .map_err(|_| invalid("project compilation arguments require numeric values"))?;
        let (name, value) = values
            .into_iter()
            .next()
            .ok_or_else(|| invalid("project compilation constant"))?;
        let mut chars = name.chars();
        if !chars.next().is_some_and(char::is_alphabetic)
            || !chars.all(|c| c.is_alphanumeric() || c == '_')
            || !names.insert(name.to_lowercase())
            || i16::try_from(value).is_err()
        {
            return Err(invalid(
                "project compilation arguments require unique VBA names and signed 16-bit integer values",
            ));
        }
        result.push(format!("{name} = {value}"));
    }
    let result = result.join(" : ");
    if crate::encode(&result, cp)?.len() > 1015 {
        return Err(Error::Limit("project compilation arguments"));
    }
    Ok(result)
}

pub(crate) fn write(
    p: &ProjectProperties,
    refs: &[Reference],
    modules: &[Module],
) -> Result<Vec<u8>> {
    let mut w = Writer {
        bytes: Vec::new(),
        cp: p.codepage,
    };
    w.long(1, p.sys_kind)?;
    if let Some(compat) = p.compat_version {
        w.long(0x4a, compat)?;
    }
    w.long(2, p.lcid)?;
    w.long(0x14, p.lcid_invoke)?;
    w.short(3, p.codepage)?;
    w.text(4, &p.name)?;
    w.pair(5, 0x40, &p.description)?;
    // Excel's source-only loader interprets these fields as UTF-16 even
    // though MS-OVBA describes duplicated MBCS. A cached native save uses
    // MBCS; the reader accepts both and checks against PROJECT's path.
    let help_file = p
        .help_file
        .encode_utf16()
        .flat_map(u16::to_le_bytes)
        .collect::<Vec<_>>();
    w.record(6, &help_file)?;
    w.record(0x3d, &help_file)?;
    w.long(7, p.help_context)?;
    w.long(8, p.lib_flags)?;
    w.u16(9);
    w.u32(4);
    w.u32(p.version_major);
    w.u16(p.version_minor);
    w.pair(0xc, 0x3c, &project_constants(&p.constants, p.codepage)?)?;
    for reference in refs {
        w.reference_name(&reference.name)?;
        match &reference.kind {
            ReferenceKind::Registered { libid } => {
                let bytes = crate::encode(libid, w.cp)?;
                w.u16(0xd);
                w.u32(bytes.len() as u32 + 10);
                w.sized(&bytes)?;
                w.u32(0);
                w.u16(0);
            }
            ReferenceKind::Project {
                absolute,
                relative,
                major,
                minor,
            } => {
                let a = crate::encode(absolute, w.cp)?;
                let b = crate::encode(relative, w.cp)?;
                w.u16(0xe);
                w.u32((a.len() + b.len() + 14) as u32);
                w.sized(&a)?;
                w.sized(&b)?;
                w.u32(*major);
                w.u16(*minor);
            }
            ReferenceKind::Control {
                original,
                twiddled,
                extended_name,
                extended,
                guid,
                cookie,
            } => {
                if let Some(original) = original {
                    w.text(0x33, original)?;
                }
                let t = crate::encode(twiddled, w.cp)?;
                w.u16(0x2f);
                w.u32(t.len() as u32 + 10);
                w.sized(&t)?;
                w.u32(0);
                w.u16(0);
                if let Some(name) = extended_name {
                    w.reference_name(name)?;
                }
                let e = crate::encode(extended, w.cp)?;
                w.u16(0x30);
                w.u32(e.len() as u32 + 30);
                w.sized_text(extended)?;
                w.u32(0);
                w.u16(0);
                w.bytes.extend_from_slice(guid);
                w.u32(*cookie);
            }
        }
    }
    w.short(
        0xf,
        modules
            .len()
            .try_into()
            .map_err(|_| Error::Limit("module count"))?,
    )?;
    w.short(0x13, 0)?;
    for m in modules {
        w.pair(0x19, 0x47, &m.name)?;
        w.pair(0x1a, 0x32, &m.stream_name)?;
        w.pair(0x1c, 0x48, &m.description)?;
        w.long(0x31, 0)?;
        w.long(0x1e, m.help_context)?;
        w.short(0x2c, 0)?;
        w.record(
            if m.kind == ModuleKind::Standard {
                0x21
            } else {
                0x22
            },
            &[],
        )?;
        if m.read_only {
            w.record(0x25, &[])?;
        }
        if m.private {
            w.record(0x28, &[])?;
        }
        w.record(0x2b, &[])?;
    }
    w.record(0x10, &[])?;
    Ok(w.bytes)
}
