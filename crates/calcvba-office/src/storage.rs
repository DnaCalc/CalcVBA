use crate::{Error, ModuleKind, Project, Result, compression, decode, directory, encode, invalid};
use std::{
    io::{Cursor, Read, Write},
    path::Path,
};
type Compound = cfb::CompoundFile<Cursor<Vec<u8>>>;

#[derive(Debug, Clone, Copy)]
pub struct ReadLimits {
    pub workbook_bytes: usize,
    pub project_bytes: usize,
    pub directory_bytes: usize,
    pub modules: usize,
    pub module_bytes: usize,
    pub total_source_bytes: usize,
}
impl Default for ReadLimits {
    fn default() -> Self {
        Self {
            workbook_bytes: 256 * 1024 * 1024,
            project_bytes: 16 * 1024 * 1024,
            directory_bytes: 2 * 1024 * 1024,
            modules: 1024,
            module_bytes: 4 * 1024 * 1024,
            total_source_bytes: 16 * 1024 * 1024,
        }
    }
}

#[derive(Debug, Clone)]
enum Container {
    Zip { original: Vec<u8>, part: String },
    Compound { original: Vec<u8>, root: String },
    Project,
}

/// A workbook together with its standalone VBA project. Loading never executes VBA.
#[derive(Debug, Clone)]
pub struct OfficeFile {
    pub project: Project,
    container: Container,
    documents: std::collections::BTreeMap<String, String>,
}
fn documents(project: &Project) -> std::collections::BTreeMap<String, String> {
    project
        .modules
        .iter()
        .filter(|m| m.kind == ModuleKind::Document)
        .map(|m| {
            (
                m.name.to_lowercase(),
                m.attribute("VB_Base").unwrap_or("").to_lowercase(),
            )
        })
        .collect()
}
impl OfficeFile {
    pub fn open(path: impl AsRef<Path>) -> Result<Self> {
        let path = path.as_ref();
        let limits = ReadLimits::default();
        if std::fs::metadata(path)?.len() > limits.workbook_bytes as u64 {
            return Err(Error::Limit("workbook bytes"));
        }
        Self::read(&std::fs::read(path)?, &limits)
    }
    pub fn read(bytes: &[u8], limits: &ReadLimits) -> Result<Self> {
        if bytes.len() > limits.workbook_bytes {
            return Err(Error::Limit("workbook bytes"));
        }
        if bytes.starts_with(b"PK") {
            let mut zip = zip::ZipArchive::new(Cursor::new(bytes))?;
            let root_rels = zip_bytes(&mut zip, "_rels/.rels", limits.directory_bytes)?;
            let workbook =
                relationship(&root_rels, "officeDocument", "")?.ok_or(Error::NoProject)?;
            let parent = workbook.rsplit_once('/').map_or("", |(p, _)| p);
            let name = workbook.rsplit('/').next().unwrap();
            let rels_name = if parent.is_empty() {
                format!("_rels/{name}.rels")
            } else {
                format!("{parent}/_rels/{name}.rels")
            };
            let rels = zip_bytes(&mut zip, &rels_name, limits.directory_bytes)?;
            let part = relationship(&rels, "vbaProject", parent)?.ok_or(Error::NoProject)?;
            let project = read_project(&zip_bytes(&mut zip, &part, limits.project_bytes)?, limits)?;
            let documents = documents(&project);
            Ok(Self {
                project,
                container: Container::Zip {
                    original: bytes.to_vec(),
                    part,
                },
                documents,
            })
        } else {
            let mut cfb = cfb::CompoundFile::open(Cursor::new(bytes.to_vec()))?;
            if cfb.is_stream("/EncryptedPackage") {
                return Err(Error::Unsupported("encrypted workbook".into()));
            }
            let root = project_root(&cfb)?;
            let bin = if root == "/" {
                bytes.to_vec()
            } else {
                extract_project(&mut cfb, &root, limits.project_bytes)?
            };
            let project = read_project(&bin, limits)?;
            let container = if root == "/" {
                Container::Project
            } else {
                Container::Compound {
                    original: bytes.to_vec(),
                    root,
                }
            };
            let documents = documents(&project);
            Ok(Self {
                project,
                container,
                documents,
            })
        }
    }
    pub fn to_bytes(&self) -> Result<Vec<u8>> {
        if documents(&self.project) != self.documents {
            return Err(Error::Unsupported(
                "document module identity changed; update the host workbook documents explicitly"
                    .into(),
            ));
        }
        let bin = self.project.to_bin()?;
        match &self.container {
            Container::Project => Ok(bin),
            Container::Zip { original, part } => {
                let mut input = zip::ZipArchive::new(Cursor::new(original))?;
                // Changing a signed macro project invalidates its signature.
                // Refuse until the signature relationship can be removed correctly.
                if input.file_names().any(|n| {
                    let name = n.to_ascii_lowercase();
                    name.contains("vbaprojectsignature") || name.starts_with("_xmlsignatures/")
                }) {
                    return Err(Error::Unsupported("signed VBA project write-back".into()));
                }
                let mut out = zip::ZipWriter::new(Cursor::new(Vec::new()));
                out.set_raw_comment(input.comment().to_vec().into_boxed_slice());
                for index in 0..input.len() {
                    let file = input.by_index(index)?;
                    if file.name() == part {
                        let options = zip::write::SimpleFileOptions::default()
                            .compression_method(file.compression());
                        out.start_file(part, options)?;
                        out.write_all(&bin)?;
                    } else {
                        out.raw_copy_file(file)?;
                    }
                }
                Ok(out.finish()?.into_inner())
            }
            Container::Compound { original, root } => {
                let mut out = cfb::CompoundFile::open(Cursor::new(original.clone()))?;
                let mut source = cfb::CompoundFile::open(Cursor::new(bin))?;
                out.remove_storage_all(root)?;
                copy_storage(&mut source, "/", &mut out, root, usize::MAX)?;
                Ok(out.into_inner().into_inner())
            }
        }
    }
    /// Write only after the entire replacement has been validated and encoded.
    /// `create_new` prevents accidental replacement of the input workbook.
    pub fn save_as(&self, path: impl AsRef<Path>) -> Result<()> {
        let bytes = self.to_bytes()?;
        let mut file = std::fs::OpenOptions::new()
            .write(true)
            .create_new(true)
            .open(path)?;
        file.write_all(&bytes)?;
        file.sync_all()?;
        Ok(())
    }
}

fn zip_bytes<R: Read + std::io::Seek>(
    zip: &mut zip::ZipArchive<R>,
    name: &str,
    limit: usize,
) -> Result<Vec<u8>> {
    let file = zip.by_name(name)?;
    if file.size() > limit as u64 {
        return Err(Error::Limit("package part bytes"));
    }
    let mut bytes = Vec::new();
    file.take(limit.saturating_add(1) as u64)
        .read_to_end(&mut bytes)?;
    if bytes.len() > limit {
        return Err(Error::Limit("package part bytes"));
    }
    Ok(bytes)
}
fn relationship(bytes: &[u8], suffix: &str, parent: &str) -> Result<Option<String>> {
    let mut reader = quick_xml::Reader::from_reader(bytes);
    let mut result = None;
    loop {
        match reader.read_event().map_err(invalid)? {
            quick_xml::events::Event::Empty(e) | quick_xml::events::Event::Start(e)
                if e.local_name().as_ref() == b"Relationship" =>
            {
                let mut ty = String::new();
                let mut target = String::new();
                let mut external = false;
                for attr in e.attributes() {
                    let attr = attr.map_err(invalid)?;
                    let value = attr.unescape_value().map_err(invalid)?.into_owned();
                    match attr.key.as_ref() {
                        b"Type" => ty = value,
                        b"Target" => target = value,
                        b"TargetMode" => external = value != "Internal",
                        _ => {}
                    }
                }
                if ty.rsplit('/').next() == Some(suffix) {
                    if result.is_some() || external || target.contains(['\\', '?', '#', '%', ':']) {
                        return Err(invalid("ambiguous or external package relationship"));
                    }
                    let path = if target.starts_with('/') {
                        target
                    } else {
                        format!("{parent}/{target}")
                    };
                    let mut segments = Vec::new();
                    for segment in path.split('/') {
                        match segment {
                            "" | "." => {}
                            ".." => {
                                if segments.pop().is_none() {
                                    return Err(invalid("package target escapes root"));
                                }
                            }
                            s => segments.push(s),
                        }
                    }
                    result = Some(segments.join("/"));
                }
            }
            quick_xml::events::Event::Eof => break,
            _ => {}
        }
    }
    Ok(result)
}

fn project_root(cfb: &Compound) -> Result<String> {
    let dirs = cfb
        .walk()
        .filter(|e| {
            e.is_stream()
                && e.path()
                    .file_name()
                    .is_some_and(|s| s.eq_ignore_ascii_case("dir"))
                && e.path()
                    .parent()
                    .and_then(Path::file_name)
                    .is_some_and(|s| s.eq_ignore_ascii_case("VBA"))
        })
        .map(|e| {
            e.path()
                .parent()
                .unwrap()
                .parent()
                .unwrap()
                .to_string_lossy()
                .replace('\\', "/")
        })
        .collect::<Vec<_>>();
    if dirs.len() != 1 {
        return Err(if dirs.is_empty() {
            Error::NoProject
        } else {
            invalid("multiple VBA project storages")
        });
    }
    Ok(dirs[0].clone())
}
pub(crate) fn stream(cfb: &mut Compound, path: impl AsRef<Path>, limit: usize) -> Result<Vec<u8>> {
    let input = cfb.open_stream(path)?;
    if input.len() > limit as u64 {
        return Err(Error::Limit("project stream bytes"));
    }
    let mut bytes = Vec::new();
    input
        .take(limit.saturating_add(1) as u64)
        .read_to_end(&mut bytes)?;
    if bytes.len() > limit {
        return Err(Error::Limit("project stream bytes"));
    }
    Ok(bytes)
}
fn extract_project(cfb: &mut Compound, root: &str, limit: usize) -> Result<Vec<u8>> {
    let mut out =
        cfb::CompoundFile::create_with_version(cfb::Version::V3, Cursor::new(Vec::new()))?;
    copy_storage(cfb, root, &mut out, "/", limit)?;
    Ok(out.into_inner().into_inner())
}
fn copy_storage(
    source: &mut Compound,
    root: &str,
    out: &mut Compound,
    dest: &str,
    limit: usize,
) -> Result<()> {
    let entries = source.walk_storage(root)?.collect::<Vec<_>>();
    let mut total = 0usize;
    for entry in entries {
        let relative = entry.path().strip_prefix(root).map_err(invalid)?;
        let target = Path::new(dest).join(relative);
        if entry.is_storage() || entry.path() == Path::new(root) {
            out.create_storage_all(&target)?;
            out.set_storage_clsid(&target, *entry.clsid())?;
        } else {
            total = total
                .checked_add(entry.len() as usize)
                .filter(|n| *n <= limit)
                .ok_or(Error::Limit("project bytes"))?;
            let bytes = stream(source, entry.path(), limit)?;
            out.create_stream(&target)?.write_all(&bytes)?;
        }
        out.set_state_bits(&target, entry.state_bits())?;
    }
    Ok(())
}

pub(crate) fn read_project(bytes: &[u8], limits: &ReadLimits) -> Result<Project> {
    if bytes.len() > limits.project_bytes {
        return Err(Error::Limit("project bytes"));
    }
    let mut cfb = cfb::CompoundFile::open(Cursor::new(bytes.to_vec()))?;
    let root = project_root(&cfb)?;
    let vba = Path::new(&root).join("VBA");
    let dir = compression::decompress(
        &stream(&mut cfb, vba.join("dir"), limits.project_bytes)?,
        limits.directory_bytes,
    )?;
    let directory::ParsedDirectory {
        mut properties,
        references,
        mut modules,
        help_file_bytes,
    } = directory::read(&dir, limits.modules)?;
    let project_stream = decode(
        &stream(
            &mut cfb,
            Path::new(&root).join("PROJECT"),
            limits.project_bytes,
        )?,
        properties.codepage,
    )?;
    check_protection(&project_stream)?;
    properties.help_file = if let Some(path) = project_stream
        .lines()
        .find_map(|line| line.strip_prefix("HelpFile=\"")?.strip_suffix('"'))
    {
        let wide = path
            .encode_utf16()
            .flat_map(u16::to_le_bytes)
            .collect::<Vec<_>>();
        if help_file_bytes != encode(path, properties.codepage)? && help_file_bytes != wide {
            return Err(invalid("PROJECT and dir HelpFile paths differ"));
        }
        path.to_owned()
    } else {
        decode(&help_file_bytes, properties.codepage)?
    };
    let mut total = 0usize;
    for module in &mut modules {
        if !crate::valid_name(&module.stream_name) {
            return Err(invalid("module stream name"));
        }
        let bytes = stream(
            &mut cfb,
            vba.join(&module.stream_name),
            limits.project_bytes,
        )?;
        let source = bytes
            .get(module.text_offset as usize..)
            .ok_or_else(|| invalid("module text offset"))?;
        let source = compression::decompress(source, limits.module_bytes)?;
        total = total
            .checked_add(source.len())
            .filter(|n| *n <= limits.total_source_bytes)
            .ok_or(Error::Limit("total source bytes"))?;
        // Raw final chunks can contain zero padding (MS-OVBA 2.4.1.3.10).
        module.source = decode(&source, properties.codepage)?
            .trim_end_matches('\0')
            .to_owned();
        for line in project_stream.lines() {
            if let Some(value) = line.strip_prefix("Document=") {
                let (name, version) = value
                    .split_once('/')
                    .ok_or_else(|| invalid("PROJECT document"))?;
                if name == module.name {
                    module.kind = ModuleKind::Document;
                    module.document_version = Some(version.into());
                }
            } else if let Some(value) = line.strip_prefix("BaseClass=")
                && value == module.name
            {
                module.kind = ModuleKind::Form;
            }
        }
    }
    let project = Project {
        properties,
        references,
        modules,
        project_stream,
        root,
        original: bytes.to_vec(),
    };
    project.validate()?;
    Ok(project)
}

pub(crate) fn write_project(project: &Project) -> Result<Vec<u8>> {
    project.validate()?;
    check_protection(&project.project_stream)?;
    let cp = project.properties.codepage;
    let mut modules = project.modules.clone();
    for module in &mut modules {
        if let Some(description) = module
            .attribute("VB_Description")
            .and_then(|s| s.strip_prefix('"'))
            .and_then(|s| s.strip_suffix('"'))
        {
            module.description = description.replace("\"\"", "\"");
        }
        if let Some(help) = module.attribute("VB_HelpID") {
            module.help_context = help.parse().map_err(invalid)?;
        }
    }
    // Encode every replacement before touching the compound file.
    let sources = project
        .modules
        .iter()
        .map(|m| compression::compress(&encode(&m.source, cp)?))
        .collect::<Result<Vec<_>>>()?;
    let dir = compression::compress(&directory::write(
        &project.properties,
        &project.references,
        &modules,
    )?)?;
    let mut cfb = cfb::CompoundFile::open(Cursor::new(project.original.clone()))?;
    if cfb.walk().any(|entry| {
        entry
            .name()
            .to_ascii_lowercase()
            .contains("digitalsignature")
    }) {
        return Err(Error::Unsupported(
            "signed compound VBA project write-back".into(),
        ));
    }
    let vba = Path::new(&project.root).join("VBA");
    // Keep the owning storage while replacing its source and cache streams.
    if !cfb.is_storage(&vba) {
        cfb.create_storage(&vba)?;
    }
    let obsolete = cfb
        .walk_storage(&vba)?
        .filter(|e| {
            e.is_stream()
                && e.name() != "dir"
                && e.name() != "_VBA_PROJECT"
                && !project
                    .modules
                    .iter()
                    .any(|m| m.stream_name.eq_ignore_ascii_case(e.name()))
        })
        .map(|e| e.path().to_owned())
        .collect::<Vec<_>>();
    for path in obsolete {
        cfb.remove_stream(path)?;
    }
    // MS-OVBA 2.3.4.1: signature 0x61CC, followed by the cache version.
    // Preserve implementation-reserved flags, invalidate the version and omit
    // the performance cache. The source streams contain no compiled prefix.
    let mut header = [0xcc, 0x61, 0xff, 0xff, 0, 0, 0];
    if let Ok(mut cache) = cfb.open_stream(vba.join("_VBA_PROJECT")) {
        cache.read_exact(&mut header)?;
    }
    header[..4].copy_from_slice(&[0xcc, 0x61, 0xff, 0xff]);
    cfb.create_stream(vba.join("_VBA_PROJECT"))?
        .write_all(&header)?;
    cfb.create_stream(vba.join("dir"))?.write_all(&dir)?;
    for (module, source) in project.modules.iter().zip(sources) {
        cfb.create_stream(vba.join(&module.stream_name))?
            .write_all(&source)?;
    }
    cfb.create_stream(Path::new(&project.root).join("PROJECT"))?
        .write_all(&encode(&updated_project_stream(project), cp)?)?;
    let mut names = Vec::new();
    for module in &project.modules {
        names.extend_from_slice(&encode(&module.name, cp)?);
        names.push(0);
        names.extend(module.name.encode_utf16().flat_map(u16::to_le_bytes));
        names.extend_from_slice(&[0, 0]);
    }
    names.extend_from_slice(&[0, 0]);
    cfb.create_stream(Path::new(&project.root).join("PROJECTwm"))?
        .write_all(&names)?;
    Ok(cfb.into_inner().into_inner())
}
fn updated_project_stream(project: &Project) -> String {
    let mut lines = Vec::new();
    let original = project.project_stream.lines().collect::<Vec<_>>();
    let section = original
        .iter()
        .position(|line| line.starts_with('['))
        .unwrap_or(original.len());
    lines.push(
        original[..section]
            .iter()
            .find(|line| line.starts_with("ID="))
            .copied()
            .unwrap_or("ID=\"{00000000-0000-0000-0000-000000000000}\"")
            .into(),
    );
    lines.extend(
        original[..section]
            .iter()
            .filter(|line| line.starts_with("Package="))
            .map(|line| (*line).to_owned()),
    );
    for module in &project.modules {
        lines.push(match module.kind {
            ModuleKind::Standard => format!("Module={}", module.name),
            ModuleKind::Class => format!("Class={}", module.name),
            ModuleKind::Document => format!(
                "Document={}/{}",
                module.name,
                module.document_version.as_deref().unwrap_or("&H00000000")
            ),
            ModuleKind::Form => format!("BaseClass={}", module.name),
        });
    }
    if !project.properties.help_file.is_empty() {
        lines.push(format!(
            "HelpFile=\"{}\"",
            project.properties.help_file.replace('"', "\"\"")
        ));
    }
    lines.extend(
        original[..section]
            .iter()
            .filter(|line| line.starts_with("ExeName32="))
            .map(|line| (*line).to_owned()),
    );
    lines.push(format!("Name=\"{}\"", project.properties.name));
    lines.push(format!(
        "HelpContextID=\"{}\"",
        project.properties.help_context
    ));
    if !project.properties.description.is_empty() {
        lines.push(format!(
            "Description=\"{}\"",
            project.properties.description
        ));
    }
    for line in &original[..section] {
        if [
            "ID=",
            "Module=",
            "Class=",
            "Document=",
            "BaseClass=",
            "Package=",
            "ExeName32=",
            "Name=",
            "Description=",
            "HelpContextID=",
            "HelpFile=",
        ]
        .iter()
        .any(|prefix| line.starts_with(prefix))
        {
            continue;
        }
        lines.push((*line).into());
    }
    lines.extend(original[section..].iter().map(|line| (*line).to_owned()));
    lines.join("\r\n") + "\r\n"
}

pub(crate) fn check_protection(project: &str) -> Result<()> {
    for line in project.lines() {
        if let Some(value) = line.strip_prefix("CMG=") {
            let bytes = decrypt(value.trim_matches('"'))?;
            if bytes.len() != 4 {
                return Err(invalid("project protection state length"));
            }
            if bytes[0] & 7 != 0 {
                return Err(Error::Locked);
            }
        }
    }
    Ok(())
}

pub(crate) fn export_form(project: &Project, module: &crate::Module) -> Result<(String, Vec<u8>)> {
    let mut source = cfb::CompoundFile::open(Cursor::new(project.original.clone()))?;
    let root = Path::new(&project.root).join(&module.name);
    let frame = decode(
        &stream(
            &mut source,
            root.join("\u{3}VBFrame"),
            ReadLimits::default().module_bytes,
        )?,
        project.properties.codepage,
    )?;
    let mut form =
        cfb::CompoundFile::create_with_version(cfb::Version::V3, Cursor::new(Vec::new()))?;
    copy_storage(
        &mut source,
        &root.to_string_lossy(),
        &mut form,
        "/",
        ReadLimits::default().project_bytes,
    )?;
    form.remove_stream("/\u{3}VBFrame")?;
    form.set_storage_clsid(
        "/",
        "c62a69f0-16dc-11ce-9e98-00aa00574a4f"
            .parse()
            .map_err(invalid)?,
    )?;
    let storage = form.into_inner().into_inner();
    let dimension = |key: &str| {
        frame
            .lines()
            .find_map(|line| {
                let (name, value) = line.trim().split_once('=')?;
                (name.trim() == key)
                    .then(|| value.trim().parse::<i32>().ok())
                    .flatten()
            })
            .unwrap_or(0)
    };
    let mut frx = b"LB\x08\0".to_vec();
    frx.extend_from_slice(&(storage.len() as u32).to_le_bytes());
    frx.extend_from_slice(&[0; 8]);
    frx.extend_from_slice(&dimension("ClientWidth").to_le_bytes());
    frx.extend_from_slice(&dimension("ClientHeight").to_le_bytes());
    frx.extend_from_slice(&storage);
    let mut header = Vec::new();
    for line in frame.lines() {
        if line.trim().starts_with("TypeInfoVer") {
            continue;
        }
        if line.trim().eq_ignore_ascii_case("End") {
            header.push(format!(
                "   OleObjectBlob   =   \"{}.frx\":0000",
                module.name
            ));
        }
        header.push(line.into());
    }
    Ok((header.join("\r\n") + "\r\n" + &module.source, frx))
}

pub(crate) fn import_form(
    out: &mut Compound,
    name: &str,
    frm: &str,
    frx: &[u8],
    cp: u16,
) -> Result<()> {
    if frx.len() < 24
        || &frx[..4] != b"LB\x08\0"
        || u32::from_le_bytes(frx[4..8].try_into().unwrap()) as usize != frx.len() - 24
    {
        return Err(invalid("VBE form resource header"));
    }
    let mut source = cfb::CompoundFile::open(Cursor::new(frx[24..].to_vec()))?;
    let root = format!("/{name}");
    copy_storage(
        &mut source,
        "/",
        out,
        &root,
        ReadLimits::default().project_bytes,
    )?;
    out.set_storage_clsid(
        &root,
        "00000000-0000-0000-0000-000000000000"
            .parse()
            .map_err(invalid)?,
    )?;
    let mut frame = Vec::new();
    let mut blobs = 0;
    for line in frm.lines() {
        if line.trim_start().starts_with("Attribute ") {
            break;
        }
        if line.trim_start().starts_with("OleObjectBlob") {
            if !line.trim_end().ends_with(":0000") {
                return Err(Error::Unsupported("nonzero form resource offset".into()));
            }
            blobs += 1;
            continue;
        }
        frame.push(line);
    }
    if blobs != 1 {
        return Err(invalid("form must name one OleObjectBlob at offset 0000"));
    }
    out.create_stream(Path::new(&root).join("\u{3}VBFrame"))?
        .write_all(&encode(&(frame.join("\r\n") + "\r\n"), cp)?)?;
    Ok(())
}
fn decrypt(hex: &str) -> Result<Vec<u8>> {
    if !hex.len().is_multiple_of(2) {
        return Err(invalid("project protection encoding"));
    }
    let bytes = hex
        .as_bytes()
        .chunks_exact(2)
        .map(|c| {
            let s = std::str::from_utf8(c).map_err(invalid)?;
            u8::from_str_radix(s, 16).map_err(invalid)
        })
        .collect::<Result<Vec<_>>>()?;
    if bytes.len() < 8 {
        return Err(invalid("project protection header"));
    }
    let seed = bytes[0];
    let version = bytes[1] ^ seed;
    if version != 2 {
        return Err(invalid("project protection version"));
    }
    let key = bytes[2] ^ seed;
    let mut enc1 = bytes[2];
    let mut enc2 = bytes[1];
    let mut previous = key;
    let ignored = ((seed & 6) >> 1) as usize;
    let mut decoded = Vec::new();
    for byte in bytes.into_iter().skip(3) {
        let plain = byte ^ enc2.wrapping_add(previous);
        enc2 = enc1;
        enc1 = byte;
        previous = plain;
        decoded.push(plain);
    }
    let data = decoded
        .get(ignored..)
        .ok_or_else(|| invalid("project protection ignored bytes"))?;
    let length = data
        .get(..4)
        .ok_or_else(|| invalid("project protection length"))?;
    let length = u32::from_le_bytes(length.try_into().unwrap()) as usize;
    if length.checked_add(4) != Some(data.len()) {
        return Err(invalid("project protection payload length"));
    }
    Ok(data[4..].to_vec())
}

#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn published_unlocked_protection_example() {
        assert_eq!(
            decrypt("0705D8E3D8EDDBF1DBF1DBF1DBF1").unwrap(),
            [0, 0, 0, 0]
        );
        check_protection("CMG=\"0705D8E3D8EDDBF1DBF1DBF1DBF1\"\r\n").unwrap();
    }
}
