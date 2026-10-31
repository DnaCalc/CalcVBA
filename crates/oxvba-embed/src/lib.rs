//! Version-one value-call boundary for spreadsheet function hosts.
//! See `docs/spec/UDF_EMBEDDING_V1.md` for scope, ownership and error mapping.

use oxvba_bundle::{ExportTarget, ExportToken, ProcedureKind};
use oxvba_hal::{HostServices, adapters::builder::HostBuilder};
use oxvba_oxir::{FuncId, OxProgram, OxTy};
use oxvba_runtime::{
    VarType, Variant,
    safe_array::{SafeArray, SafeArrayBound},
};
use oxvba_symbol::manifest as sym;
use std::sync::Arc;

/// Spreadsheet values; array cells are row-major and contain scalar values.
#[derive(Debug, Clone, PartialEq)]
pub enum Value {
    Empty,
    Number(f64),
    Text(String),
    Boolean(bool),
    /// VBA CVErr / Excel error code, e.g. 2042 for #N/A.
    Error(i32),
    Array {
        rows: u32,
        columns: u32,
        cells: Vec<Value>,
    },
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Backend {
    Interpreter,
    Jit,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Parameter {
    pub name: String,
    pub type_name: String,
    pub by_ref: bool,
    pub optional: bool,
    pub param_array: bool,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Function {
    pub name: String,
    pub qualified_name: String,
    pub parameters: Vec<Parameter>,
    pub return_type: String,
}

#[derive(Debug, thiserror::Error)]
pub enum Error {
    #[error("compile: {0}")]
    Compile(String),
    #[error("unknown public function: {0}")]
    UnknownFunction(String),
    #[error("expected {expected} arguments, received {received}")]
    ArgumentCount { expected: usize, received: usize },
    #[error("unsupported signature: {0}")]
    UnsupportedSignature(String),
    #[error("value conversion: {0}")]
    ValueConversion(String),
    #[error("execution budget exhausted")]
    BudgetExceeded,
    #[error("JIT backend is not enabled in this build")]
    BackendUnavailable,
    #[error("backend: {0}")]
    Backend(String),
    #[error("backend does not support this module: {0}")]
    UnsupportedBackend(String),
    #[error("VBA error {number}: {description}")]
    Runtime {
        number: i32,
        description: String,
        vba_source: String,
    },
}

/// Compiled semantic module, owned by the caller. It neither executes Main nor
/// registers names in a global registry. Public functions come from compiler exports.
pub struct Module {
    program: OxProgram,
    functions: Vec<(Function, FuncId)>,
    host: Arc<dyn HostServices>,
}

impl Module {
    pub fn load(module_name: &str, source: &str) -> Result<Self, Error> {
        let manifest = sym::SymbolProjectManifest {
            project_name: "UdfProject".into(),
            project_kind: sym::ProjectKind::Library,
            modules: vec![sym::ModuleUnit {
                module_name: module_name.into(),
                module_kind: sym::ModuleKind::Procedural,
                attributes: sym::ModuleAttributes::named(module_name),
                source: source.into(),
            }],
            references: vec![],
            reference_projects: vec![],
            conditional_constants: Default::default(),
            conditional_compilation_target: Default::default(),
        };
        let core = oxvba_bind::bind_program(&manifest, &oxvba_symbol::CatalogTypeLibResolver)
            .map_err(|err| Error::Compile(err.to_string()))?;
        let program = oxvba_oxir::elaborate::elaborate(&core)
            .map_err(|err| Error::Compile(err.to_string()))?;
        oxvba_oxir::verify::verify_program(&program)
            .map_err(|errors| Error::Compile(format!("invalid compiled module: {errors:?}")))?;
        // Descriptor lifetime for classes/UDTs belongs to full project sessions.
        if !program.classes.is_empty() || !program.record_layouts.is_empty() {
            return Err(Error::UnsupportedSignature(
                "classes and UDTs require the project-session API".into(),
            ));
        }
        let mut functions = Vec::new();
        for export in &program.exports {
            let (ExportToken::ModuleFunc { module, member, .. }, ExportTarget::Proc(index)) =
                (&export.token, export.target)
            else {
                continue;
            };
            let proc = &program.funcs[index];
            if proc.kind != ProcedureKind::Function {
                continue;
            }
            let parameters = proc
                .locals
                .iter()
                .take(proc.param_count)
                .map(|local| {
                    let info = local.param.expect("compiler parameter metadata");
                    Parameter {
                        name: local.name.clone(),
                        type_name: type_name(&local.ty),
                        by_ref: info.by_ref,
                        optional: info.optional,
                        param_array: info.variadic,
                    }
                })
                .collect();
            functions.push((
                Function {
                    name: member.clone(),
                    qualified_name: format!("{module}.{member}"),
                    parameters,
                    return_type: proc
                        .return_local
                        .map(|id| type_name(&proc.locals[id.0].ty))
                        .unwrap_or_default(),
                },
                FuncId(index),
            ));
        }
        Ok(Self {
            program,
            functions,
            host: HostBuilder::new().build(),
        })
    }

    /// Replace the default Null HAL with explicit host-owned capabilities.
    pub fn with_host(mut self, host: Arc<dyn HostServices>) -> Self {
        self.host = host;
        self
    }

    pub fn functions(&self) -> impl ExactSizeIterator<Item = &Function> {
        self.functions.iter().map(|(function, _)| function)
    }

    /// JIT compilation happens here once; unsupported images return an error.
    pub fn prepare(&self, backend: Backend) -> Result<Session<'_>, Error> {
        let executable = match backend {
            Backend::Interpreter => Executable::Interpreter,
            #[cfg(feature = "jit")]
            Backend::Jit => Executable::Jit(Box::new(
                oxvba_jit::JitEngine
                    .compile_image(&[&self.program])
                    .map_err(jit_error)?,
            )),
            #[cfg(not(feature = "jit"))]
            Backend::Jit => return Err(Error::BackendUnavailable),
        };
        Ok(Session {
            module: self,
            executable,
        })
    }
}

enum Executable<'a> {
    Interpreter,
    #[cfg(feature = "jit")]
    Jit(Box<oxvba_jit::CompiledImage<'a>>),
    #[cfg(not(feature = "jit"))]
    _Lifetime(std::marker::PhantomData<&'a ()>),
}

/// Borrows its semantic module and retains compiled code until dropped. Each
/// invocation uses fresh globals/Err state and private copies of host arguments.
pub struct Session<'a> {
    module: &'a Module,
    executable: Executable<'a>,
}

impl Session<'_> {
    pub fn call(&self, name: &str, args: &[Value], max_steps: u64) -> Result<Value, Error> {
        let (metadata, id) = self
            .module
            .functions
            .iter()
            .find(|(function, _)| {
                function.name.eq_ignore_ascii_case(name)
                    || function.qualified_name.eq_ignore_ascii_case(name)
            })
            .ok_or_else(|| Error::UnknownFunction(name.into()))?;
        if args.len() != metadata.parameters.len() {
            return Err(Error::ArgumentCount {
                expected: metadata.parameters.len(),
                received: args.len(),
            });
        }
        let function = &self.module.program.funcs[id.0];
        let mut values = Vec::with_capacity(args.len());
        for (value, local) in args.iter().zip(&function.locals) {
            if local.param.is_some_and(|param| param.variadic) {
                return Err(Error::UnsupportedSignature(
                    "ParamArray packing is not part of v1".into(),
                ));
            }
            values.push(coerce_argument(to_variant(value)?, &local.ty)?);
        }
        let value = match &self.executable {
            Executable::Interpreter => {
                let mut vm = oxvba_vm3::Vm3::link_with_budget(
                    &[&self.module.program],
                    &*self.module.host,
                    Some(max_steps),
                )
                .map_err(vm_error)?;
                vm.invoke_function_values(*id, values).map_err(vm_error)?
            }
            #[cfg(feature = "jit")]
            Executable::Jit(compiled) => {
                let mut outcome = compiled
                    .invoke_function_values(&*self.module.host, *id, &values, max_steps)
                    .map_err(jit_error)?;
                if outcome.raised {
                    return Err(Error::Runtime {
                        number: outcome.err.number,
                        description: outcome.err.description,
                        vba_source: outcome.err.source,
                    });
                }
                outcome.values.pop().unwrap_or_else(Variant::empty)
            }
            #[cfg(not(feature = "jit"))]
            Executable::_Lifetime(_) => unreachable!(),
        };
        from_variant(&value, false)
    }
}

fn vm_error(error: oxvba_vm3::Vm3Error) -> Error {
    match error {
        oxvba_vm3::Vm3Error::BudgetExceeded => Error::BudgetExceeded,
        oxvba_vm3::Vm3Error::Fault(fault) => Error::Runtime {
            number: fault.code,
            description: fault.message,
            vba_source: fault.source.unwrap_or_else(|| "UdfProject".into()),
        },
        other => Error::Backend(other.to_string()),
    }
}
#[cfg(feature = "jit")]
fn jit_error(error: oxvba_jit::JitError) -> Error {
    match error {
        oxvba_jit::JitError::BudgetExceeded => Error::BudgetExceeded,
        oxvba_jit::JitError::Unsupported(reason) => Error::UnsupportedBackend(reason),
        other => Error::Backend(other.to_string()),
    }
}

fn coerce_argument(value: Variant, ty: &OxTy) -> Result<Variant, Error> {
    let target = match ty {
        OxTy::Variant => return Ok(value),
        OxTy::Array(element, _)
            if **element == OxTy::Variant && value.vtype() == VarType::ArrayVariant =>
        {
            return Ok(value);
        }
        OxTy::Bool => VarType::Boolean,
        OxTy::Byte => VarType::Byte,
        OxTy::Integer => VarType::Integer,
        OxTy::Long => VarType::Long,
        OxTy::LongLong => VarType::LongLong,
        OxTy::Single => VarType::Single,
        OxTy::Double => VarType::Double,
        OxTy::Currency => VarType::Currency,
        OxTy::Date => VarType::Date,
        OxTy::Str => VarType::String,
        other => {
            return Err(Error::UnsupportedSignature(format!(
                "parameter type {other:?}"
            )));
        }
    };
    oxvba_runtime::coerce::coerce_to(&value, target).map_err(Error::ValueConversion)
}

fn type_name(ty: &OxTy) -> String {
    match ty {
        OxTy::Bool => "Boolean",
        OxTy::Byte => "Byte",
        OxTy::Integer => "Integer",
        OxTy::Long => "Long",
        OxTy::LongLong => "LongLong",
        OxTy::Single => "Single",
        OxTy::Double => "Double",
        OxTy::Currency => "Currency",
        OxTy::Date => "Date",
        OxTy::Str => "String",
        OxTy::Variant => "Variant",
        OxTy::Array(element, _) => return format!("{}()", type_name(element)),
        other => return format!("Unsupported({other:?})"),
    }
    .into()
}

fn to_variant(value: &Value) -> Result<Variant, Error> {
    Ok(match value {
        Value::Empty => Variant::empty(),
        Value::Number(value) => {
            if !value.is_finite() {
                return Err(Error::ValueConversion(
                    "spreadsheet numbers must be finite".into(),
                ));
            }
            Variant::from_f64(*value)
        }
        Value::Text(value) => Variant::from_string(value.as_str()),
        Value::Boolean(value) => Variant::from_bool(*value),
        Value::Error(code) => Variant::from_error_code(*code),
        Value::Array {
            rows,
            columns,
            cells,
        } => {
            let count = (*rows as usize).checked_mul(*columns as usize);
            if *rows == 0
                || *columns == 0
                || count != Some(cells.len())
                || cells.iter().any(|cell| matches!(cell, Value::Array { .. }))
            {
                return Err(Error::ValueConversion(
                    "array must be a nonempty rectangle of scalar cells".into(),
                ));
            }
            let values = cells
                .iter()
                .map(to_variant)
                .collect::<Result<Vec<_>, _>>()?;
            let array = SafeArray::from_shape_and_variants(
                vec![
                    SafeArrayBound {
                        count: *rows,
                        lower: 1,
                    },
                    SafeArrayBound {
                        count: *columns,
                        lower: 1,
                    },
                ],
                values,
            )
            .map_err(Error::ValueConversion)?;
            Variant::from_safearray(array)
        }
    })
}

fn from_variant(value: &Variant, inside_array: bool) -> Result<Value, Error> {
    Ok(match value.vtype() {
        VarType::Empty => Value::Empty,
        VarType::Boolean => Value::Boolean(value.as_bool().expect("Boolean carrier")),
        VarType::Error => Value::Error(value.as_error_code().expect("Error carrier")),
        VarType::String => {
            let text = value.as_bstr().expect("String carrier");
            if !text.byte_len().is_multiple_of(2) {
                return Err(Error::ValueConversion(
                    "result text has an odd BSTR byte length".into(),
                ));
            }
            Value::Text(
                String::from_utf16(&text.to_utf16_units()).map_err(|_| {
                    Error::ValueConversion("result text is not valid UTF-16".into())
                })?,
            )
        }
        VarType::ArrayVariant if !inside_array => {
            let array = value
                .as_safearray()
                .ok_or_else(|| Error::ValueConversion("unallocated result array".into()))?;
            let bounds = array
                .bounds()
                .ok_or_else(|| Error::ValueConversion("array has no bounds".into()))?;
            let (rows, columns) = match bounds.as_slice() {
                [columns] => (1, columns.count),
                [rows, columns] => (rows.count, columns.count),
                _ => {
                    return Err(Error::ValueConversion(
                        "only one or two dimensional result arrays are supported".into(),
                    ));
                }
            };
            if rows == 0 || columns == 0 {
                return Err(Error::ValueConversion(
                    "result array must be a nonempty rectangle".into(),
                ));
            }
            let cells = (0..array.len())
                .map(|index| {
                    array
                        .variant_element(index)
                        .map_err(Error::ValueConversion)
                        .and_then(|value| from_variant(&value, true))
                })
                .collect::<Result<Vec<_>, _>>()?;
            Value::Array {
                rows,
                columns,
                cells,
            }
        }
        VarType::Byte
        | VarType::Integer
        | VarType::Long
        | VarType::LongLong
        | VarType::Single
        | VarType::Double
        | VarType::Date
        | VarType::Currency
        | VarType::Decimal => {
            let number = oxvba_runtime::coerce::coerce_to(value, VarType::Double)
                .map_err(Error::ValueConversion)?;
            Value::Number(number.as_f64().expect("Double coercion"))
        }
        other => {
            return Err(Error::ValueConversion(format!(
                "unsupported result carrier {other:?}"
            )));
        }
    })
}
