//! A Cranelift Module that lays code/data into an ordinary PE image. Compilation
//! and relocation finish at build time; no executable allocator or COFF linker.

use cranelift_codegen::{
    Context,
    binemit::Reloc,
    ir,
    isa::{self, unwind::UnwindInfo},
    settings::{self, Configurable},
};
use cranelift_control::ControlPlane;
use cranelift_module::{
    DataDescription, DataId, FuncId, Init, Linkage, Module, ModuleDeclarations, ModuleError,
    ModuleReloc, ModuleRelocTarget, ModuleResult,
};
use oxvba_pe::{Export, Program, Template, Unwind};
use std::{collections::BTreeMap, sync::Arc};

fn error(message: impl Into<String>) -> ModuleError {
    ModuleError::Backend(std::io::Error::other(message.into()).into())
}

#[derive(Debug, thiserror::Error)]
#[error("native AOT: {0}")]
pub struct Error(pub String);
pub type Result<T> = std::result::Result<T, Error>;
struct Function {
    bytes: Vec<u8>,
    alignment: usize,
    relocs: Vec<ModuleReloc>,
    unwind: Option<Vec<u8>>,
}
struct Data {
    bytes: Vec<u8>,
    alignment: usize,
}
pub struct PeModule {
    isa: Arc<dyn isa::TargetIsa>,
    declarations: ModuleDeclarations,
    functions: BTreeMap<FuncId, Function>,
    data: BTreeMap<DataId, Data>,
    helpers: BTreeMap<FuncId, u16>,
}
impl PeModule {
    pub fn new() -> Result<Self> {
        if !cfg!(all(target_os = "windows", target_arch = "x86_64")) {
            return Err(Error("only Windows x64 is implemented".into()));
        }
        let mut flags = settings::builder();
        for (name, value) in [
            ("opt_level", "speed"),
            // Cranelift's PIC mode routes even colocated data through a GOT.
            // All our symbols are near, in this image: non-PIC mode selects
            // RIP-relative LEA instead. finish rejects absolute relocations,
            // so the emitted image still relocates without code fixups.
            ("is_pic", "false"),
            ("unwind_info", "true"),
            ("use_colocated_libcalls", "true"),
            ("enable_verifier", "true"),
        ] {
            flags.set(name, value).map_err(|e| Error(e.to_string()))?;
        }
        let isa = cranelift_native::builder()
            .map_err(|e| Error(e.into()))?
            .finish(settings::Flags::new(flags))
            .map_err(|e| Error(e.to_string()))?;
        Ok(Self {
            isa,
            declarations: ModuleDeclarations::default(),
            functions: BTreeMap::new(),
            data: BTreeMap::new(),
            helpers: BTreeMap::new(),
        })
    }
    pub fn declare_helper(
        &mut self,
        name: &str,
        version: u16,
        signature: &ir::Signature,
    ) -> Result<FuncId> {
        let id = self
            .declare_function(name, Linkage::Import, signature)
            .map_err(|e| Error(e.to_string()))?;
        if self
            .helpers
            .insert(id, version)
            .is_some_and(|old| old != version)
        {
            return Err(Error("incompatible helper version".into()));
        }
        Ok(id)
    }
    /// Return relocatable code with checked, direct rel32 calls. Export IDs
    /// denote actual generated wrappers, never an entry token or thunk slot.
    pub fn finish(
        self,
        template: &Template<'_>,
        exports: &[(String, Option<u16>, FuncId)],
    ) -> Result<Program> {
        let mut text = Vec::new();
        let mut offsets = BTreeMap::new();
        let mut unwind = Vec::new();
        for (id, function) in &self.functions {
            pad(&mut text, function.alignment);
            let offset =
                u32::try_from(text.len()).map_err(|_| Error("code exceeds RVA range".into()))?;
            offsets.insert(*id, offset);
            text.extend_from_slice(&function.bytes);
            if let Some(info) = &function.unwind {
                unwind.push(Unwind {
                    begin: offset,
                    end: offset
                        .checked_add(function.bytes.len() as u32)
                        .ok_or_else(|| Error("function range overflow".into()))?,
                    info: info.clone(),
                });
            }
        }
        let layout = template
            .layout(text.len())
            .map_err(|e| Error(e.to_string()))?;
        let mut data = Vec::new();
        let mut data_offsets = BTreeMap::new();
        for (id, value) in &self.data {
            pad(&mut data, value.alignment);
            data_offsets.insert(
                *id,
                u32::try_from(data.len()).map_err(|_| Error("data exceeds RVA range".into()))?,
            );
            data.extend_from_slice(&value.bytes);
        }
        let resolve = |name: &ModuleRelocTarget| -> Result<u32> {
            match name {
                ModuleRelocTarget::User {
                    namespace: 0,
                    index,
                } => {
                    let id = FuncId::from_u32(*index);
                    if let Some(offset) = offsets.get(&id) {
                        return layout
                            .text_rva
                            .checked_add(*offset)
                            .ok_or_else(|| Error("code RVA overflow".into()));
                    }
                    let decl = self
                        .declarations
                        .get_functions()
                        .find_map(|(declared, decl)| (declared == id).then_some(decl))
                        .ok_or_else(|| {
                            Error("relocation refers to an undeclared function".into())
                        })?;
                    let name = decl
                        .name
                        .as_deref()
                        .ok_or_else(|| Error("anonymous unresolved procedure".into()))?;
                    let version = self
                        .helpers
                        .get(&id)
                        .ok_or_else(|| Error(format!("unresolved helper {name}")))?;
                    template
                        .manifest
                        .resolve(name, *version)
                        .map_err(|e| Error(e.to_string()))
                }
                ModuleRelocTarget::User {
                    namespace: 1,
                    index,
                } => {
                    let offset = data_offsets
                        .get(&DataId::from_u32(*index))
                        .ok_or_else(|| Error("unresolved data".into()))?;
                    layout
                        .read_only_rva
                        .checked_add(*offset)
                        .ok_or_else(|| Error("data RVA overflow".into()))
                }
                ModuleRelocTarget::FunctionOffset(id, offset) => {
                    if self
                        .functions
                        .get(id)
                        .is_none_or(|f| *offset as usize >= f.bytes.len())
                    {
                        return Err(Error("relocation target outside function".into()));
                    }
                    let begin = offsets
                        .get(id)
                        .ok_or_else(|| Error("unresolved function offset".into()))?;
                    layout
                        .text_rva
                        .checked_add(*begin)
                        .and_then(|r| r.checked_add(*offset))
                        .ok_or_else(|| Error("function RVA overflow".into()))
                }
                other => Err(Error(format!("unsupported relocation target {other}"))),
            }
        };
        for (id, function) in &self.functions {
            let base = offsets[id] as usize;
            for relocation in &function.relocs {
                if !matches!(
                    relocation.kind,
                    Reloc::X86PCRel4 | Reloc::X86CallPCRel4 | Reloc::X86CallPLTRel4
                ) {
                    return Err(Error(format!(
                        "unsupported code relocation {:?}",
                        relocation.kind
                    )));
                }
                let at = base
                    .checked_add(relocation.offset as usize)
                    .ok_or_else(|| Error("relocation offset overflow".into()))?;
                if (relocation.offset as usize)
                    .checked_add(4)
                    .is_none_or(|end| end > function.bytes.len())
                {
                    return Err(Error("relocation outside function".into()));
                }
                let target = i64::from(resolve(&relocation.name)?);
                let source = i64::from(layout.text_rva) + at as i64;
                let relative = target
                    .checked_add(relocation.addend)
                    .and_then(|t| t.checked_sub(source))
                    .ok_or_else(|| Error("relative relocation overflow".into()))?;
                let relative = i32::try_from(relative)
                    .map_err(|_| Error("rel32 target out of range".into()))?;
                text[at..at + 4].copy_from_slice(&relative.to_le_bytes());
            }
        }
        let exports = exports
            .iter()
            .map(|(name, ordinal, id)| {
                let offset = *offsets
                    .get(id)
                    .ok_or_else(|| Error(format!("export {name} has no compiled body")))?;
                Ok(Export {
                    name: name.clone(),
                    ordinal: *ordinal,
                    offset,
                })
            })
            .collect::<Result<Vec<_>>>()?;
        Ok(Program {
            text,
            read_only: data,
            exports,
            unwind,
            ..Program::default()
        })
    }
}
fn pad(bytes: &mut Vec<u8>, alignment: usize) {
    bytes.resize((bytes.len() + alignment - 1) & !(alignment - 1), 0);
}
impl Module for PeModule {
    fn isa(&self) -> &dyn isa::TargetIsa {
        &*self.isa
    }
    fn declarations(&self) -> &ModuleDeclarations {
        &self.declarations
    }
    fn declare_function(
        &mut self,
        name: &str,
        linkage: Linkage,
        signature: &ir::Signature,
    ) -> ModuleResult<FuncId> {
        self.declarations
            .declare_function(name, linkage, signature)
            .map(|(id, _)| id)
    }
    fn declare_anonymous_function(&mut self, signature: &ir::Signature) -> ModuleResult<FuncId> {
        self.declarations.declare_anonymous_function(signature)
    }
    fn declare_data(
        &mut self,
        name: &str,
        linkage: Linkage,
        writable: bool,
        tls: bool,
    ) -> ModuleResult<DataId> {
        if writable || tls {
            return Err(error("writable/TLS data is not yet implemented"));
        }
        self.declarations
            .declare_data(name, linkage, writable, tls)
            .map(|(id, _)| id)
    }
    fn declare_anonymous_data(&mut self, writable: bool, tls: bool) -> ModuleResult<DataId> {
        if writable || tls {
            return Err(error("writable/TLS data is not yet implemented"));
        }
        self.declarations.declare_anonymous_data(writable, tls)
    }
    fn declare_func_in_func(&mut self, id: FuncId, function: &mut ir::Function) -> ir::FuncRef {
        let decl = self.declarations.get_function_decl(id);
        let signature = function.import_signature(decl.signature.clone());
        let name = function.declare_imported_user_function(ir::UserExternalName {
            namespace: 0,
            index: id.as_u32(),
        });
        // Every target is in this single image, including template helpers.
        function.import_function(ir::ExtFuncData {
            name: ir::ExternalName::user(name),
            signature,
            colocated: true,
            patchable: false,
        })
    }
    fn define_function_with_control_plane(
        &mut self,
        id: FuncId,
        context: &mut Context,
        plane: &mut ControlPlane,
    ) -> ModuleResult<()> {
        if !self
            .declarations
            .get_function_decl(id)
            .linkage
            .is_definable()
        {
            return Err(error("cannot define an imported function"));
        }
        if self.functions.contains_key(&id) {
            return Err(error("duplicate function definition"));
        }
        context.compile(&*self.isa, plane)?;
        let compiled = context.compiled_code().expect("successful compilation");
        let unwind = match compiled.create_unwind_info(&*self.isa)? {
            Some(UnwindInfo::WindowsX64(info)) => {
                let mut bytes = vec![0; info.emit_size()];
                info.emit(&mut bytes);
                Some(bytes)
            }
            None => None,
            _ => return Err(error("expected Windows x64 unwind information")),
        };
        let relocs = compiled
            .buffer
            .relocs()
            .iter()
            .map(|r| ModuleReloc::from_mach_reloc(r, &context.func, id))
            .collect();
        self.functions.insert(
            id,
            Function {
                bytes: compiled.code_buffer().to_vec(),
                alignment: (compiled.buffer.alignment as usize).max(16),
                relocs,
                unwind,
            },
        );
        Ok(())
    }
    fn define_function_bytes(
        &mut self,
        _id: FuncId,
        _alignment: u64,
        _bytes: &[u8],
        _relocs: &[ModuleReloc],
    ) -> ModuleResult<()> {
        Err(error(
            "raw functions require explicit unwind information; use Cranelift definitions",
        ))
    }
    fn define_data(&mut self, id: DataId, description: &DataDescription) -> ModuleResult<()> {
        if !self.declarations.get_data_decl(id).linkage.is_definable()
            || self.data.contains_key(&id)
        {
            return Err(error("imported or duplicate data definition"));
        }
        if !description.function_relocs.is_empty()
            || !description.data_relocs.is_empty()
            || description.custom_segment_section.is_some()
        {
            return Err(error(
                "data relocations/custom sections are not yet implemented",
            ));
        }
        let bytes = match &description.init {
            Init::Bytes { contents } => contents.to_vec(),
            Init::Zeros { size } => vec![0; *size],
            Init::Uninitialized => return Err(error("undefined data")),
        };
        let alignment = description.align.unwrap_or(8);
        if !alignment.is_power_of_two() || alignment > 4096 {
            return Err(error("unsupported data alignment"));
        }
        self.data.insert(
            id,
            Data {
                bytes,
                alignment: alignment as usize,
            },
        );
        Ok(())
    }
}
