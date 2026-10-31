//! Version-one value-call boundary for spreadsheet function hosts.
//! See `docs/spec/UDF_EMBEDDING_V1.md` for scope, ownership and error mapping.

use oxvba_bundle::coreir::CoreConst;
mod cache;
mod spreadsheet;
pub use cache::{CacheStats, CodeCache, OwnedSession};
pub use spreadsheet::{Range, SpreadsheetContext};

use oxvba_bundle::{ExportTarget, ExportToken, ProcedureKind};
use oxvba_hal::{HostServices, adapters::builder::HostBuilder};
use oxvba_oxir::{FuncId, OxProgram, OxTy};
use oxvba_runtime::{
    VarType, Variant,
    safe_array::{SafeArray, SafeArrayBound},
};
use oxvba_symbol::manifest as sym;
use std::{cell::RefCell, rc::Rc, sync::Arc};

/// Spreadsheet values; array cells are row-major and contain scalar values.
#[derive(Debug, Clone, PartialEq)]
pub enum Value {
    Empty,
    /// An omitted positional argument, distinct from a blank cell.
    Missing,
    /// A live read-only reference in the host's sheet namespace.
    Range(Range),
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
    #[error("spreadsheet host refused the call: {0}")]
    HostRefused(String),
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
#[derive(Clone)]
pub struct Module {
    program: OxProgram,
    functions: Vec<(Function, FuncId, Vec<Option<CoreConst>>)>,
    host: Arc<dyn HostServices>,
    spreadsheet: Option<Arc<spreadsheet::SpreadsheetHost>>,
}

impl Module {
    pub fn load(module_name: &str, source: &str) -> Result<Self, Error> {
        Self::load_inner(module_name, source, false)
    }

    /// Compile against the bounded, host-supplied read-only Application/Range surface.
    pub fn load_spreadsheet(module_name: &str, source: &str) -> Result<Self, Error> {
        Self::load_inner(module_name, source, true)
    }

    fn load_inner(
        module_name: &str,
        source: &str,
        spreadsheet_objects: bool,
    ) -> Result<Self, Error> {
        let manifest = sym::SymbolProjectManifest {
            project_name: "UdfProject".into(),
            project_kind: sym::ProjectKind::Library,
            modules: vec![sym::ModuleUnit {
                module_name: module_name.into(),
                module_kind: sym::ModuleKind::Procedural,
                attributes: sym::ModuleAttributes::named(module_name),
                source: source.into(),
            }],
            references: if spreadsheet_objects {
                vec![
                    sym::ProjectReference::HostInjected {
                        referenced_project_name: "DnaCalc.Application".into(),
                    },
                    sym::ProjectReference::TypeLibrary {
                        name: "Excel".into(),
                        guid: None,
                        version_major: None,
                        version_minor: None,
                        lcid: None,
                        import_lib: None,
                    },
                ]
            } else {
                vec![]
            },
            reference_projects: vec![],
            conditional_constants: Default::default(),
            conditional_compilation_target: Default::default(),
        };
        let resolver: &dyn oxvba_symbol::TypeLibResolver = if spreadsheet_objects {
            &spreadsheet::SpreadsheetTypes
        } else {
            &oxvba_symbol::CatalogTypeLibResolver
        };
        let core = oxvba_bind::bind_program(&manifest, resolver)
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
        if program.funcs.iter().any(|function| {
            function
                .blocks
                .iter()
                .any(|block| matches!(block.terminator, oxvba_oxir::OxTerminator::Halt))
        }) {
            return Err(Error::UnsupportedSignature(
                "End requires whole-project reset semantics".into(),
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
                core.procs[index]
                    .params
                    .iter()
                    .map(|p| p.default.clone())
                    .collect(),
            ));
        }
        let spreadsheet =
            spreadsheet_objects.then(|| Arc::new(spreadsheet::SpreadsheetHost::new()));
        let host: Arc<dyn HostServices> = match &spreadsheet {
            Some(host) => host.clone(),
            None => HostBuilder::new().build(),
        };
        Ok(Self {
            program,
            functions,
            host,
            spreadsheet,
        })
    }

    /// Replace the default Null HAL with explicit host-owned capabilities.
    pub fn with_host(mut self, host: Arc<dyn HostServices>) -> Self {
        self.host = host;
        self.spreadsheet = None;
        self
    }

    pub fn functions(&self) -> impl ExactSizeIterator<Item = &Function> {
        self.functions.iter().map(|(function, _, _)| function)
    }

    /// JIT compilation happens once; unsupported images return an error.
    pub fn prepare(&self, backend: Backend) -> Result<Session<'_>, Error> {
        PreparedCode::compile(self, backend, None)?
            .into_session(&*self.host, self.spreadsheet.as_deref())
    }
}

#[derive(Clone)]
enum PreparedCode<'a> {
    Interpreter(&'a Module),
    #[cfg(feature = "jit")]
    Jit(&'a Module, Rc<oxvba_jit::CompiledImage<'a>>),
}
impl<'a> PreparedCode<'a> {
    fn compile(
        module: &'a Module,
        backend: Backend,
        configuration: Option<&JitConfiguration>,
    ) -> Result<Self, Error> {
        match backend {
            Backend::Interpreter => Ok(Self::Interpreter(module)),
            #[cfg(feature = "jit")]
            Backend::Jit => {
                let image = match configuration {
                    Some(configuration) => oxvba_jit::JitEngine
                        .compile_image_with_configuration(&[&module.program], configuration),
                    None => oxvba_jit::JitEngine.compile_image(&[&module.program]),
                }
                .map_err(jit_error)?;
                Ok(Self::Jit(module, Rc::new(image)))
            }
            #[cfg(not(feature = "jit"))]
            Backend::Jit => {
                let _ = configuration;
                Err(Error::BackendUnavailable)
            }
        }
    }
    fn into_session(
        self,
        host: &'a dyn HostServices,
        spreadsheet: Option<&'a spreadsheet::SpreadsheetHost>,
    ) -> Result<Session<'a>, Error> {
        let (module, executable) = match self {
            Self::Interpreter(module) => (module, Executable::Interpreter(RefCell::new(None))),
            #[cfg(feature = "jit")]
            Self::Jit(module, image) => {
                let state = image.start_session(host).map_err(jit_error)?;
                (module, Executable::Jit(image, RefCell::new(state)))
            }
        };
        Ok(Session {
            module,
            host,
            spreadsheet,
            executable,
        })
    }
}

#[cfg(feature = "jit")]
use oxvba_jit::JitConfiguration;
#[cfg(not(feature = "jit"))]
struct JitConfiguration;

enum Executable<'a> {
    Interpreter(RefCell<Option<oxvba_vm3::Vm3<'a>>>),
    #[cfg(feature = "jit")]
    Jit(
        Rc<oxvba_jit::CompiledImage<'a>>,
        RefCell<oxvba_jit::JitSession<'a>>,
    ),
    #[cfg(not(feature = "jit"))]
    _Lifetime(std::marker::PhantomData<&'a ()>),
}

/// Borrows its semantic module and retains compiled code until dropped. Each
/// invocation retains project storage and uses private copies of host arguments.
pub struct Session<'a> {
    module: &'a Module,
    host: &'a dyn HostServices,
    spreadsheet: Option<&'a spreadsheet::SpreadsheetHost>,
    executable: Executable<'a>,
}

impl<'a> Session<'a> {
    pub fn backend(&self) -> Backend {
        match &self.executable {
            Executable::Interpreter(_) => Backend::Interpreter,
            #[cfg(feature = "jit")]
            Executable::Jit(..) => Backend::Jit,
            #[cfg(not(feature = "jit"))]
            Executable::_Lifetime(_) => unreachable!(),
        }
    }

    pub fn call_with_context(
        &self,
        name: &str,
        args: &[Value],
        max_steps: u64,
        context: &mut dyn SpreadsheetContext,
    ) -> Result<Value, Error> {
        spreadsheet::with_context(context, || self.call_inner(name, args, max_steps))
    }

    fn argument(&self, value: &Value) -> Result<Variant, Error> {
        match value {
            Value::Range(range) => self
                .spreadsheet
                .ok_or_else(|| {
                    Error::UnsupportedSignature("Range requires spreadsheet host objects".into())
                })?
                .range_variant(*range),
            value => to_variant(value),
        }
    }

    /// Clear project storage without releasing retained compiled code.
    pub fn reset(&self) -> Result<(), Error> {
        match &self.executable {
            Executable::Interpreter(state) => {
                *state
                    .try_borrow_mut()
                    .map_err(|_| Error::Backend("session is already executing".into()))? = None
            }
            #[cfg(feature = "jit")]
            Executable::Jit(image, state) => {
                *state
                    .try_borrow_mut()
                    .map_err(|_| Error::Backend("session is already executing".into()))? =
                    image.start_session(self.host).map_err(jit_error)?
            }
            #[cfg(not(feature = "jit"))]
            Executable::_Lifetime(_) => unreachable!(),
        }
        Ok(())
    }

    pub fn call(&self, name: &str, args: &[Value], max_steps: u64) -> Result<Value, Error> {
        if self.spreadsheet.is_some() {
            spreadsheet::without_context(|| self.call_inner(name, args, max_steps))
        } else {
            self.call_inner(name, args, max_steps)
        }
    }

    fn call_inner(&self, name: &str, args: &[Value], max_steps: u64) -> Result<Value, Error> {
        let (metadata, id, defaults) = self
            .module
            .functions
            .iter()
            .find(|(function, _, _)| {
                function.name.eq_ignore_ascii_case(name)
                    || function.qualified_name.eq_ignore_ascii_case(name)
            })
            .ok_or_else(|| Error::UnknownFunction(name.into()))?;
        let function = &self.module.program.funcs[id.0];
        let variadic = metadata.parameters.last().is_some_and(|p| p.param_array);
        let fixed = function.param_count - usize::from(variadic);
        if (!variadic && args.len() > fixed)
            || metadata
                .parameters
                .iter()
                .take(fixed)
                .enumerate()
                .any(|(i, p)| {
                    !p.optional && args.get(i).is_none_or(|a| matches!(a, Value::Missing))
                })
        {
            return Err(Error::ArgumentCount {
                expected: fixed,
                received: args.len(),
            });
        }
        let mut values = Vec::with_capacity(function.param_count);
        for (index, local) in function.locals.iter().take(fixed).enumerate() {
            let value = match args.get(index) {
                Some(Value::Range(range))
                    if !matches!(local.ty, OxTy::Variant | OxTy::Object(_)) =>
                {
                    let value = self
                        .spreadsheet
                        .ok_or_else(|| {
                            Error::UnsupportedSignature(
                                "Range requires spreadsheet host objects".into(),
                            )
                        })?
                        .read_range(*range)?;
                    to_variant(&value)?
                }
                Some(value) if !matches!(value, Value::Missing) => self.argument(value)?,
                _ => match &defaults[index] {
                    Some(value) => default_variant(value),
                    None if local.ty == OxTy::Variant => Variant::from_error_code(-2147352572),
                    None if matches!(local.ty, OxTy::Object(_)) => Variant::nothing(),
                    None => Variant::empty(),
                },
            };
            values.push(coerce_argument(value, &local.ty)?);
        }
        if variadic {
            let tail = args
                .iter()
                .skip(fixed)
                .map(|value| self.argument(value))
                .collect::<Result<Vec<_>, _>>()?;
            values.push(Variant::from_safearray(SafeArray::from_variants(tail)));
        }
        let value = match &self.executable {
            Executable::Interpreter(state) => {
                let mut state = state
                    .try_borrow_mut()
                    .map_err(|_| Error::Backend("session is already executing".into()))?;
                if state.is_none() {
                    *state = Some(
                        oxvba_vm3::Vm3::link_with_budget(
                            &[&self.module.program],
                            self.host,
                            Some(max_steps),
                        )
                        .map_err(vm_error)?,
                    );
                } else {
                    state
                        .as_mut()
                        .expect("linked session")
                        .set_execution_budget(max_steps);
                }
                state
                    .as_mut()
                    .expect("linked session")
                    .invoke_function_values(*id, values)
                    .map_err(vm_error)?
            }
            #[cfg(feature = "jit")]
            Executable::Jit(compiled, state) => {
                let mut state = state
                    .try_borrow_mut()
                    .map_err(|_| Error::Backend("session is already executing".into()))?;
                let mut outcome = compiled
                    .invoke_in_session(&mut state, *id, &values, max_steps)
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
        match self.spreadsheet {
            Some(host) => host.result(&value),
            None => from_variant(&value, false),
        }
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
    use oxvba_bundle::NumericCoerceTarget as Target;
    let target = match ty {
        OxTy::Variant => return Ok(value),
        OxTy::Object(_) if value.vtype() == VarType::Object => return Ok(value),
        OxTy::Array(element, _)
            if **element == OxTy::Variant && value.vtype() == VarType::ArrayVariant =>
        {
            return Ok(value);
        }
        OxTy::Str => return oxvba_eval::arith::coerce_string(&value).map_err(argument_error),
        OxTy::Bool => Target::Boolean,
        OxTy::Byte => Target::Byte,
        OxTy::Integer => Target::Integer,
        OxTy::Long => Target::Long,
        OxTy::LongLong => Target::LongLong,
        OxTy::Single => Target::Single,
        OxTy::Double => Target::Double,
        OxTy::Currency => Target::Currency,
        OxTy::Date => Target::Date,
        other => {
            return Err(Error::UnsupportedSignature(format!(
                "parameter type {other:?}"
            )));
        }
    };
    oxvba_eval::arith::coerce_numeric(&value, target).map_err(argument_error)
}

fn argument_error(error: oxvba_eval::arith::ArithError) -> Error {
    Error::Runtime {
        number: error.code,
        description: error.message,
        vba_source: "UdfProject".into(),
    }
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
        OxTy::Object(_) => "Object",
        OxTy::Array(element, _) => return format!("{}()", type_name(element)),
        other => return format!("Unsupported({other:?})"),
    }
    .into()
}

fn to_variant(value: &Value) -> Result<Variant, Error> {
    Ok(match value {
        Value::Range(_) => {
            return Err(Error::ValueConversion(
                "Range requires a live spreadsheet invocation".into(),
            ));
        }
        Value::Empty => Variant::empty(),
        Value::Missing => Variant::from_error_code(-2147352572),
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

fn default_variant(value: &CoreConst) -> Variant {
    match value {
        CoreConst::Empty => Variant::empty(),
        CoreConst::Null => Variant::null(),
        CoreConst::Nothing => Variant::nothing(),
        CoreConst::Bool(v) => Variant::from_bool(*v),
        CoreConst::I16(v) => Variant::from_i16(*v),
        CoreConst::I32(v) => Variant::from_i32(*v),
        CoreConst::I64(v) => Variant::from_i64(*v),
        CoreConst::F64(v) => Variant::from_f64(f64::from_bits(*v)),
        CoreConst::F32(v) => Variant::from_f32(f32::from_bits(*v)),
        CoreConst::Currency(v) => Variant::from_currency_scaled_i64(*v),
        CoreConst::Date(v) => Variant::from_date_f64(f64::from_bits(*v)),
        CoreConst::Str(v) => Variant::from_string(v.as_str()),
    }
}
