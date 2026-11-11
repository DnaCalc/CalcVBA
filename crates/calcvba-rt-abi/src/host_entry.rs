//! Versioned, allocation-free scalar host-entry convention and admission facts.
//! This is a bounded entry family, not an alternative semantic IR. Backends
//! consume the original OxFunc and the same checked admission facts.
use calcvba_bundle::{NumericCoerceTarget as N, NumericMode};
use calcvba_oxir::{
    ArithOp, OxCoerceTarget, OxConst, OxFunc, OxInst, OxOperand, OxPlace, OxTerminator, OxTy,
};

pub use calcvba_host_api::{
    Arguments, MAX_TYPED_ARGS, Scalar, ScalarBits, TYPED_ENTRY_ABI, TypedContext,
};
pub const MAX_TYPED_SLOTS: usize = 128;
pub use calcvba_host_api::ValueType as ScalarType;
pub mod activation;
mod aggregate;
pub use aggregate::array_flat_index;
pub mod collection;
mod iteration;
mod member;
pub mod value;

/// The object-default interpretation of an index operation. Native lowering
/// uses this only for a project object; arrays retain their ordinary path.
pub fn index_member_call(inst: &OxInst) -> Option<OxInst> {
    use calcvba_com::TypeLibMemberInvokeKind as Kind;
    let (dst, recv, indices, value) = match inst {
        OxInst::ArrayGet {
            dst,
            array,
            indices,
        } => (Some(*dst), array.clone(), indices, None),
        OxInst::ArraySet {
            array,
            indices,
            value,
        } => (None, OxOperand::Use(*array), indices, Some(value)),
        _ => return None,
    };
    let mut args: Vec<_> = indices
        .iter()
        .cloned()
        .map(calcvba_oxir::OxCallArg::Operand)
        .collect();
    if let Some(value) = value {
        args.push(calcvba_oxir::OxCallArg::Operand(value.clone()));
    }
    Some(OxInst::ComCallLate {
        dst,
        recv,
        name: String::new(),
        default_member: true,
        invoke_kind: if value.is_some() {
            Kind::PropertyPut
        } else {
            Kind::PropertyGet
        },
        args,
    })
}
pub fn owning_builtin(id: calcvba_bundle::NativeImplId) -> bool {
    use calcvba_bundle::NativeImplId::*;
    matches!(
        id,
        IsMissing
            | IsArray
            | VarType
            | CStr
            | CVar
            | Left
            | Right
            | Mid
            | UCase
            | LCase
            | Trim
            | LTrim
            | RTrim
            | Replace
            | Len
            | InStr
            | InStrRev
            | StrComp
            | Join
            | Split
            | CDate
            | CBool
            | CByte
            | CCur
            | CDbl
            | CInt
            | CLng
            | CLngLng
            | CSng
            | LenB
            | LeftB
            | RightB
            | Chr
            | Asc
            | ChrW
            | AscW
            | Space
            | StringRepeat
            | StrReverse
            | StrConv
            | Format
            | FormatNumber
            | FormatCurrency
            | FormatPercent
            | FormatDateTime
            | Filter
            | Int
            | Fix
            | Sgn
            | Round
            | Sin
            | Cos
            | Log
            | Exp
            | Atn
            | Tan
            | DateSerial
            | TimeSerial
            | DateValue
            | TimeValue
            | DateAdd
            | DateDiff
            | Year
            | Month
            | Day
            | Weekday
            | Hour
            | Minute
            | Second
            | MonthName
            | WeekdayName
            | DatePart
            | Hex
            | Oct
            | Str
            | Val
            | CVErr
            | CDec
            | IsNumeric
            | IsError
            | IsDate
            | IsObject
            | IsNull
            | IsEmpty
            | TypeName
            | Rgb
            | CreateObject
    )
}

/// Shared operations whose mechanics belong to the retained call/error stack.
/// Arithmetic and host callbacks still execute in the selected backend.
pub fn stack_instruction(inst: &OxInst, layout: &ScalarLayout) -> bool {
    if matches!(inst, OxInst::ComCallLate { .. }) {
        return layout.owning;
    }
    matches!(
        inst,
        OxInst::CallProc { .. }
            | OxInst::Predeclared { .. }
            | OxInst::WithEventsGet { .. }
            | OxInst::WithEventsSet { .. }
            | OxInst::WithEventsClearOwner { .. }
            | OxInst::NewObject { .. }
            | OxInst::NewExtern { .. }
            | OxInst::CallNative { .. }
            | OxInst::ComCallLate { .. }
            | OxInst::FieldGet { .. }
            | OxInst::FieldSet { .. }
            | OxInst::FieldArrayGet { .. }
            | OxInst::FieldArraySet { .. }
            | OxInst::DrainTerminations
            | OxInst::CallExtern { .. }
            | OxInst::SetErrorHandler(_)
            | OxInst::ClearErr
            | OxInst::ErrFieldGet { .. }
            | OxInst::ErlGet { .. }
            | OxInst::SetLineNumber { .. }
    )
}
pub fn stack_terminator(term: &OxTerminator) -> bool {
    matches!(
        term,
        OxTerminator::FaultDispatch { .. }
            | OxTerminator::Resume
            | OxTerminator::ResumeNext
            | OxTerminator::ResumeLabel(_)
            | OxTerminator::Raise { .. }
    )
}

fn scalar_type(ty: &OxTy) -> Option<ScalarType> {
    match ty {
        OxTy::Long => Some(ScalarType::Long),
        OxTy::Double => Some(ScalarType::Double),
        OxTy::Bool => Some(ScalarType::Boolean),
        OxTy::Str => Some(ScalarType::String),
        OxTy::Variant => Some(ScalarType::Variant),
        OxTy::Record(_) => Some(ScalarType::Record),
        OxTy::Array(_, _) => Some(ScalarType::Array),
        OxTy::Object(_) => Some(ScalarType::HostObject),
        _ => None,
    }
}

pub fn native_scalar_type(ty: &OxTy) -> Option<ScalarType> {
    scalar_type(ty).or(match ty {
        OxTy::Byte => Some(ScalarType::Byte),
        OxTy::Integer => Some(ScalarType::Integer),
        OxTy::LongLong => Some(ScalarType::LongLong),
        OxTy::Single => Some(ScalarType::Single),
        OxTy::Currency => Some(ScalarType::Currency),
        OxTy::Date => Some(ScalarType::Date),
        OxTy::FixedStr(_) => Some(ScalarType::String),
        _ => None,
    })
}
fn extended_numeric(ty: ScalarType) -> bool {
    matches!(
        ty,
        ScalarType::Byte
            | ScalarType::Integer
            | ScalarType::LongLong
            | ScalarType::Single
            | ScalarType::Currency
            | ScalarType::Date
    )
}
#[derive(Clone, Copy)]
struct LayoutLimits {
    arguments: usize,
    slots: usize,
    native_numbers: bool,
}
const HOST_LIMITS: LayoutLimits = LayoutLimits {
    arguments: MAX_TYPED_ARGS,
    slots: MAX_TYPED_SLOTS,
    native_numbers: false,
};
const NATIVE_LIMITS: LayoutLimits = LayoutLimits {
    arguments: 32,
    slots: 4096,
    native_numbers: true,
};

fn physical_operand_type(
    func: &OxFunc,
    physical: &[Option<ScalarType>],
    global_base: usize,
    op: &OxOperand,
) -> Option<ScalarType> {
    match op {
        OxOperand::Use(OxPlace::Local(id)) if id.0 < func.locals.len() => physical[id.0],
        OxOperand::Use(OxPlace::Temp(id)) => func
            .locals
            .len()
            .checked_add(id.0)
            .and_then(|i| physical.get(i))
            .copied()
            .flatten(),
        OxOperand::Use(OxPlace::Global(id)) => physical.get(global_base + id.0).copied().flatten(),
        OxOperand::Const(OxConst::I16(_) | OxConst::I32(_)) => Some(ScalarType::Long),
        OxOperand::Const(OxConst::F64(_)) => Some(ScalarType::Double),
        OxOperand::Const(OxConst::Bool(_)) => Some(ScalarType::Boolean),
        _ => None,
    }
}

fn logical_type(lhs: ScalarType, rhs: ScalarType) -> Option<ScalarType> {
    match (lhs, rhs) {
        (ScalarType::Boolean, ScalarType::Boolean) => Some(ScalarType::Boolean),
        (ScalarType::Long, ScalarType::Long | ScalarType::Boolean)
        | (ScalarType::Boolean, ScalarType::Long) => Some(ScalarType::Long),
        _ => None,
    }
}

/// Transfer the caller's budget into a stack-owned typed activation.
/// # Safety
/// State follows RawExecState's live unique same-thread contract; context is a
/// nonaliasing, aligned, writable TypedContext for this synchronous call.
pub unsafe extern "C" fn typed_context_start(
    state: *mut crate::RawExecState,
    context: *mut TypedContext,
) -> i32 {
    crate::with_status(|| {
        // SAFETY: the generated adapter retains unique state/context storage.
        let Some(state) = (unsafe { crate::state_from_raw(state) }) else {
            return crate::ST_FAULT;
        };
        if context.is_null() {
            return crate::ST_FAULT;
        }
        // SAFETY: caller supplies aligned writable stack storage.
        unsafe {
            (*context).remaining_blocks = state.remaining_steps.unwrap_or(u64::MAX);
        }
        crate::ST_OK
    })
}

/// Rejoin the dynamic error/budget convention after a typed activation.
/// # Safety
/// State and context follow typed_context_start's contract and remain live;
/// context contains the budget returned by that activation. A nonempty source
/// names one live readable allocation of source_len UTF-8 bytes.
pub unsafe extern "C" fn typed_context_finish(
    state: *mut crate::RawExecState,
    context: *const TypedContext,
    status: i32,
    source: *const u8,
    source_len: usize,
) -> i32 {
    crate::with_status(|| {
        if context.is_null() {
            return crate::ST_FAULT;
        }
        // SAFETY: caller owns both nonaliasing live objects for the call.
        let Some(exec) = (unsafe { crate::state_from_raw(state) }) else {
            return crate::ST_FAULT;
        };
        if exec.remaining_steps.is_some() {
            // SAFETY: the context is live and generated entry has returned.
            exec.remaining_steps = Some(unsafe { (*context).remaining_blocks });
        }
        match status {
            0 => crate::ST_OK,
            -1 => crate::ST_ABORT,
            number => {
                let source = if source_len == 0 {
                    ""
                } else {
                    if source.is_null() {
                        return crate::ST_FAULT;
                    }
                    // SAFETY: generated code retains this UTF-8 program name
                    // for the image lifetime; caller guarantees its full extent.
                    let bytes = unsafe { std::slice::from_raw_parts(source, source_len) };
                    let Ok(source) = std::str::from_utf8(bytes) else {
                        return crate::ST_FAULT;
                    };
                    source
                };
                let mut fault = crate::Fault::new(number, crate::default_error_message(number));
                fault.source = Some(source.into());
                exec.err_engine
                    .raise(fault, exec.default_error_source.clone());
                crate::ST_FAULT
            }
        }
    })
}

/// Inspectable physical layout for a scalar entry. Construction checks every
/// instruction, including unreachable blocks. Unsupported bodies are refused.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ScalarLayout {
    pub library_news: std::collections::HashSet<calcvba_oxir::ImportId>,
    pub slots: Vec<ScalarType>,
    pub local_count: usize,
    pub parameters: Vec<ScalarType>,
    /// Required nominal identity for each object parameter (None for Object).
    pub parameter_iids: Vec<Option<[u8; 16]>>,
    /// Nominal destination views, keyed by physical local/temp/global slot.
    pub slot_iids: std::collections::HashMap<usize, [u8; 16]>,
    pub parameter_classes: Vec<Option<String>>,
    pub assignment_classes: std::collections::HashMap<String, String>,
    pub class_calls: Vec<(OxInst, calcvba_oxir::FuncId, calcvba_oxir::ClassId)>,
    pub class_news: std::collections::HashMap<calcvba_oxir::ClassId, Option<calcvba_oxir::FuncId>>,
    pub by_ref_parameters: Vec<bool>,
    pub optional_parameters: Vec<bool>,
    pub param_array: Option<usize>,
    pub array_slots: std::collections::HashMap<usize, u16>,
    pub array_record_slots: std::collections::HashMap<
        usize,
        (
            calcvba_oxir::RecordLayoutId,
            std::sync::Arc<calcvba_runtime::vba_record::VbaRecordLayout>,
        ),
    >,
    pub record_slots: std::collections::HashMap<
        usize,
        std::sync::Arc<calcvba_runtime::vba_record::VbaRecordLayout>,
    >,
    pub has_result: bool,
    pub calls: std::collections::HashMap<calcvba_oxir::FuncId, CallSignature>,
    pub extern_calls: std::collections::HashMap<calcvba_oxir::ImportId, LinkedCall>,
    pub result: ScalarType,
    pub return_slot: usize,
    pub specialized_variant_temps: Vec<usize>,
    literal_integer_temps: std::collections::HashSet<usize>,
    pub host_calls: std::collections::HashMap<calcvba_oxir::ComMethodRef, HostCallPlan>,
    /// Runtime value coercion selects only compiler-retained default getters.
    pub default_calls: std::collections::HashMap<[u8; 16], HostCallPlan>,
    pub default_faults: std::collections::HashMap<[u8; 16], i32>,
    /// Globals have virtual physical indices after the actual frame slots.
    pub global_base: usize,
    pub assignment_iids: std::collections::HashMap<String, Option<[u8; 16]>>,
    unsupported_globals: std::collections::HashSet<usize>,
    owning: bool,
    native_numbers: bool,
}
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CallSignature {
    pub parameters: Vec<ScalarType>,
    pub by_ref: Vec<bool>,
    pub optional: Vec<bool>,
    pub result: Option<ScalarType>,
}
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct LinkedCall {
    pub target: LinkedTarget,
    pub signature: CallSignature,
}
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum LinkedTarget {
    Procedure {
        program: usize,
        procedure: calcvba_oxir::FuncId,
    },
    Builtin(calcvba_bundle::NativeImplId),
}
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct HostCallPlan {
    pub native: Option<calcvba_com::TypeLibMemberMetadata>,
    pub iid: [u8; 16],
    pub slot: usize,
    pub long_argument: bool,
    pub parameters: Vec<ScalarType>,
    pub result: ScalarType,
    pub result_iid: Option<[u8; 16]>,
}
impl ScalarLayout {
    pub fn new(func: &OxFunc) -> Result<Self, String> {
        Self::with_interfaces(func, &[])
    }
    pub fn with_interfaces(
        func: &OxFunc,
        interfaces: &[calcvba_oxir::ComInterface],
    ) -> Result<Self, String> {
        Self::with_storage(
            func,
            interfaces,
            None,
            false,
            Default::default(),
            HOST_LIMITS,
        )
    }
    pub fn with_program(func: &OxFunc, program: &calcvba_oxir::OxProgram) -> Result<Self, String> {
        Self::with_programs(func, program, &[program])
    }
    pub fn with_programs(
        func: &OxFunc,
        program: &calcvba_oxir::OxProgram,
        programs: &[&calcvba_oxir::OxProgram],
    ) -> Result<Self, String> {
        let mut external = std::collections::HashMap::new();
        for inst in func.blocks.iter().flat_map(|b| &b.instrs) {
            if let OxInst::CallExtern { import, args, .. } = inst {
                let request = program
                    .imports
                    .get(import.0)
                    .ok_or("invalid typed import")?;
                if request.unit.eq_ignore_ascii_case("VBA")
                    && !programs
                        .iter()
                        .any(|p| p.unit_name.eq_ignore_ascii_case("VBA"))
                {
                    let native = calcvba_bundle::NativeImplId::ALL
                        .iter()
                        .copied()
                        .filter(|id| owning_builtin(*id))
                        .find(|native| {
                            native.library_member().is_some_and(|(module, primary)| {
                                std::iter::once(primary)
                                    .chain(native.library_member_aliases().iter().copied())
                                    .any(|member| {
                                        request.token.matches(
                                            &calcvba_bundle::ExportToken::ModuleFunc {
                                                module: module.into(),
                                                member: member.into(),
                                                kind: calcvba_bundle::ProjectMemberKind::Method,
                                            },
                                        )
                                    })
                            })
                        });
                    if let Some(native) = native {
                        external.insert(
                            *import,
                            LinkedCall {
                                target: LinkedTarget::Builtin(native),
                                signature: CallSignature {
                                    parameters: vec![ScalarType::Variant; args.len()],
                                    by_ref: vec![true; args.len()],
                                    optional: vec![true; args.len()],
                                    result: Some(ScalarType::Variant),
                                },
                            },
                        );
                        continue;
                    }
                }
                let (unit, target) = programs
                    .iter()
                    .enumerate()
                    .find(|(_, p)| p.unit_name.eq_ignore_ascii_case(&request.unit))
                    .ok_or_else(|| {
                        format!(
                            "unlinked typed source project {} ({:?})",
                            request.unit, request.token
                        )
                    })?;
                let export = target
                    .exports
                    .iter()
                    .find(|e| e.token.matches(&request.token))
                    .ok_or("unresolved typed source export")?;
                let calcvba_bundle::ExportTarget::Proc(id) = export.target else {
                    return Err("typed class import remains open".into());
                };
                let callee = target
                    .funcs
                    .get(id)
                    .ok_or("invalid typed exported procedure")?;
                external.insert(
                    *import,
                    LinkedCall {
                        target: LinkedTarget::Procedure {
                            program: unit,
                            procedure: calcvba_oxir::FuncId(id),
                        },
                        signature: call_signature(callee)?,
                    },
                );
            }
        }
        Self::with_storage(
            func,
            &program.com_interfaces,
            Some(program),
            true,
            external,
            HOST_LIMITS,
        )
    }
    /// Native exports have a generated ABI signature, rather than the host's
    /// fixed four-slot Arguments carrier. Source-project calls remain local.
    pub fn with_native_program(
        func: &OxFunc,
        program: &calcvba_oxir::OxProgram,
    ) -> Result<Self, String> {
        use calcvba_bundle::NativeImplId;
        let numeric_operand = |operand: &OxOperand| {
            let ty = match operand {
                OxOperand::Use(OxPlace::Local(id)) => native_scalar_type(&func.locals[id.0].ty),
                OxOperand::Use(OxPlace::Temp(id)) => native_scalar_type(&func.temps[id.0]),
                OxOperand::Use(OxPlace::Global(id)) => {
                    native_scalar_type(&program.globals[id.0].ty)
                }
                OxOperand::Const(
                    OxConst::Bool(_)
                    | OxConst::I16(_)
                    | OxConst::I32(_)
                    | OxConst::I64(_)
                    | OxConst::F32(_)
                    | OxConst::F64(_),
                ) => Some(ScalarType::Double),
                _ => None,
            };
            ty.is_some_and(|ty| {
                extended_numeric(ty)
                    || matches!(
                        ty,
                        ScalarType::Long | ScalarType::Double | ScalarType::Boolean
                    )
            })
        };
        let mut external = std::collections::HashMap::new();
        for inst in func.blocks.iter().flat_map(|block| &block.instrs) {
            if let OxInst::CallExtern { dst, import, args } = inst {
                let native = crate::library::resolve_import(program, *import)?.implementation;
                let result = match dst {
                    Some(OxPlace::Local(id)) => native_scalar_type(&func.locals[id.0].ty),
                    Some(OxPlace::Temp(id)) => native_scalar_type(&func.temps[id.0]),
                    Some(OxPlace::Global(id)) => native_scalar_type(&program.globals[id.0].ty),
                    None => None,
                };
                let numeric = match native {
                    NativeImplId::Abs | NativeImplId::Sqr => Some(ScalarType::Double),
                    NativeImplId::CByte => Some(ScalarType::Byte),
                    NativeImplId::CInt => Some(ScalarType::Integer),
                    NativeImplId::CLngLng | NativeImplId::CLngPtr => Some(ScalarType::LongLong),
                    NativeImplId::CSng => Some(ScalarType::Single),
                    NativeImplId::CCur => Some(ScalarType::Currency),
                    NativeImplId::CDate => Some(ScalarType::Date),
                    NativeImplId::CLng => Some(ScalarType::Long),
                    NativeImplId::CDbl => Some(ScalarType::Double),
                    NativeImplId::CBool => Some(ScalarType::Boolean),
                    _ => None,
                };
                let math = matches!(native, NativeImplId::Abs | NativeImplId::Sqr)
                    && matches!(result, None | Some(ScalarType::Double))
                    && matches!(args.as_slice(), [calcvba_oxir::OxArg::ByVal(operand)] if numeric_operand(operand));
                let mut linked = LinkedCall {
                    target: LinkedTarget::Builtin(native),
                    signature: CallSignature {
                        parameters: vec![
                            if math {
                                ScalarType::Double
                            } else {
                                ScalarType::Variant
                            };
                            args.len()
                        ],
                        by_ref: vec![!math; args.len()],
                        optional: vec![true; args.len()],
                        result: if matches!(native, NativeImplId::Abs | NativeImplId::Sqr) {
                            result.or(numeric)
                        } else {
                            numeric.or(result)
                        },
                    },
                };
                if let Some(previous) = external.get(import) {
                    let previous: &LinkedCall = previous;
                    let count = linked
                        .signature
                        .parameters
                        .len()
                        .max(previous.signature.parameters.len());
                    if linked
                        .signature
                        .parameters
                        .iter()
                        .chain(&previous.signature.parameters)
                        .any(|ty| *ty != ScalarType::Double)
                    {
                        linked.signature.parameters = vec![ScalarType::Variant; count];
                        linked.signature.by_ref = vec![true; count];
                    }
                    linked.signature.optional = vec![true; count];
                    linked.signature.result = linked.signature.result.or(previous.signature.result);
                }
                external.insert(*import, linked);
            }
        }
        Self::with_storage(
            func,
            &program.com_interfaces,
            Some(program),
            true,
            external,
            NATIVE_LIMITS,
        )
    }
    fn with_storage(
        func: &OxFunc,
        interfaces: &[calcvba_oxir::ComInterface],
        program: Option<&calcvba_oxir::OxProgram>,
        owning: bool,
        extern_calls: std::collections::HashMap<calcvba_oxir::ImportId, LinkedCall>,
        limits: LayoutLimits,
    ) -> Result<Self, String> {
        let LayoutLimits {
            arguments: argument_limit,
            slots: slot_limit,
            native_numbers,
        } = limits;
        let storage_type = |ty: &OxTy| {
            if native_numbers {
                native_scalar_type(ty)
            } else {
                scalar_type(ty)
            }
        };
        let globals = program.map_or(&[][..], |p| &p.globals);
        let functions = program.map_or(&[][..], |p| &p.funcs);
        let record_layouts = program.map_or(&[][..], |p| &p.record_layouts);
        let classes = program.map_or(&[][..], |p| &p.classes);
        use calcvba_com::{TypeLibParamType as P, TypeLibWireType};
        let mut host_calls = std::collections::HashMap::new();
        for block in &func.blocks {
            for inst in &block.instrs {
                let OxInst::ComCallEarly {
                    method,
                    invoke_kind,
                    args,
                    ..
                } = inst
                else {
                    continue;
                };
                let descriptor = interfaces
                    .get(method.iface.0)
                    .and_then(|i| i.com_members())
                    .and_then(|m| m.get(method.member))
                    .ok_or("typed host call has no retained interface descriptor")?;
                let native = descriptor.callconv_is_stdcall
                    || descriptor.is_dual
                    || descriptor.source_typekind.is_some();
                let native_signature = if native {
                    if !owning {
                        return Err("native typed COM requires an owning activation".into());
                    }
                    Some(
                        calcvba_com::typed_slots::plan::visible_signature(descriptor, args.len())
                            .ok_or("unsupported native typed COM signature")?,
                    )
                } else {
                    None
                };
                let from_param = |ty| match ty {
                    P::Long => Some(ScalarType::Long),
                    P::Double => Some(ScalarType::Double),
                    P::Boolean => Some(ScalarType::Boolean),
                    P::String => Some(ScalarType::String),
                    P::Variant => Some(ScalarType::Variant),
                    P::Object => Some(ScalarType::HostObject),
                    _ => None,
                };
                let parameters = if let Some((parameters, _)) = &native_signature {
                    parameters.clone()
                } else {
                    descriptor
                        .parameter_types
                        .iter()
                        .copied()
                        .enumerate()
                        .map(|(i, ty)| {
                            let ty =
                                from_param(ty).ok_or("unsupported typed host argument carrier")?;
                            Ok(if let Some(Some(iid)) = descriptor.parameter_iids.get(i) {
                                ScalarType::Interface(iid_bytes(*iid))
                            } else {
                                ty
                            })
                        })
                        .collect::<Result<Vec<_>, &str>>()?
                };
                let result = if let Some((_, result)) = &native_signature {
                    *result
                } else {
                    match descriptor.return_type {
                        Some(ty) => {
                            from_param(ty).ok_or("unsupported typed host result carrier")?
                        }
                        None => ScalarType::Void,
                    }
                };
                if descriptor.invoke_kind != *invoke_kind
                    || parameters.len() > MAX_TYPED_ARGS
                    || (!native && descriptor.parameter_optional.iter().any(|p| *p))
                    || (!native
                        && descriptor
                            .parameter_wire_types
                            .iter()
                            .zip(&descriptor.parameter_types)
                            .any(|(wire, ty)| {
                                !matches!(wire, TypeLibWireType::Automation(p) if p == ty)
                                    && !matches!(wire, TypeLibWireType::InterfacePointer { .. })
                            }))
                {
                    return Err("invalid typed host member signature".into());
                }
                let result_iid = match &descriptor.return_wire_type {
                    Some(TypeLibWireType::InterfacePointer { name })
                        if name.eq_ignore_ascii_case("IDispatch")
                            || name.eq_ignore_ascii_case("IUnknown") =>
                    {
                        None
                    }
                    Some(TypeLibWireType::InterfacePointer { name }) => Some(
                        interfaces
                            .iter()
                            .find(|i| i.name().eq_ignore_ascii_case(name))
                            .and_then(|i| match i {
                                calcvba_oxir::ComInterface::Com(meta) => meta.iid,
                                _ => None,
                            })
                            .map(iid_bytes)
                            .ok_or("host result has no retained nominal interface")?,
                    ),
                    Some(TypeLibWireType::Automation(p)) if Some(*p) != descriptor.return_type => {
                        return Err("typed host result wire mismatch".into());
                    }
                    _ => None,
                };
                if !owning
                    && matches!(
                        result,
                        ScalarType::HostObject | ScalarType::String | ScalarType::Variant
                    )
                {
                    return Err("owning member results require an owning frame".into());
                }
                let iid = descriptor
                    .interface_iid
                    .ok_or("typed host slot requires interface identity")?;
                let mut bytes = [0u8; 16];
                bytes[..4].copy_from_slice(&iid.data1.to_le_bytes());
                bytes[4..6].copy_from_slice(&iid.data2.to_le_bytes());
                bytes[6..8].copy_from_slice(&iid.data3.to_le_bytes());
                bytes[8..].copy_from_slice(&iid.data4);
                let slot = calcvba_com::typed_slots::plan::table_slot(
                    descriptor,
                    interfaces[method.iface.0]
                        .com_members()
                        .ok_or("typed host interface is missing")?,
                )
                .ok_or("typed host table index is missing or exceeds declared table")?;
                let plan = HostCallPlan {
                    native: (native
                        && calcvba_com::typed_slots::scalar_signature(descriptor).is_none())
                    .then(|| descriptor.clone()),
                    iid: bytes,
                    slot,
                    long_argument: parameters == [ScalarType::Long] && result == ScalarType::Double,
                    parameters,
                    result,
                    result_iid,
                };
                if let Some(old) = host_calls.get_mut(method) {
                    let old: &mut HostCallPlan = old;
                    if old.parameters.len() < plan.parameters.len() {
                        old.parameters = plan.parameters;
                    }
                } else {
                    host_calls.insert(*method, plan);
                }
            }
        }
        let mut default_calls = std::collections::HashMap::new();
        let mut default_faults = std::collections::HashMap::new();
        if owning {
            for interface in interfaces {
                for member in interface.com_members().into_iter().flatten().filter(|m| {
                    m.is_default_member
                        && m.invoke_kind == calcvba_com::TypeLibMemberInvokeKind::PropertyGet
                }) {
                    let native = member.callconv_is_stdcall
                        || member.is_dual
                        || member.source_typekind.is_some();
                    let native_signature = native
                        .then(|| calcvba_com::typed_slots::plan::visible_signature(member, 0))
                        .flatten();
                    if (!native && !member.parameter_types.is_empty())
                        || (native && native_signature.is_none())
                    {
                        if (!native
                            || (calcvba_com::typed_slots::plan::validate(member)
                                && calcvba_com::typed_slots::plan::visible_indices(member)
                                    .iter()
                                    .any(|&i| {
                                        calcvba_com::typed_slots::plan::default_value(member, i)
                                            .is_none()
                                    })))
                            && let Some(iid) = member.interface_iid.map(iid_bytes)
                        {
                            default_faults.insert(iid, 450);
                        }
                        continue;
                    }
                    let result = if let Some((_, result)) = &native_signature {
                        *result
                    } else {
                        match member.return_type {
                            Some(P::Long) => ScalarType::Long,
                            Some(P::Double) => ScalarType::Double,
                            Some(P::Boolean) => ScalarType::Boolean,
                            Some(P::String) => ScalarType::String,
                            Some(P::Variant) => ScalarType::Variant,
                            Some(P::Object) => ScalarType::HostObject,
                            _ => continue,
                        }
                    };
                    let Some(iid) = member.interface_iid.map(iid_bytes) else {
                        continue;
                    };
                    let Some(slot) = calcvba_com::typed_slots::plan::table_slot(
                        member,
                        interface.com_members().unwrap_or(&[]),
                    ) else {
                        continue;
                    };
                    let plan = HostCallPlan {
                        native: (native
                            && calcvba_com::typed_slots::scalar_signature(member).is_none())
                        .then(|| member.clone()),
                        iid,
                        slot,
                        long_argument: false,
                        parameters: vec![],
                        result,
                        result_iid: None,
                    };
                    if let Some(old) = default_calls.insert(iid, plan.clone())
                        && old != plan
                    {
                        return Err("ambiguous runtime default member".into());
                    }
                }
            }
        }
        let scalar = |ty: &OxTy| {
            storage_type(ty)
                .filter(|ty| {
                    owning
                        || !matches!(
                            ty,
                            ScalarType::String
                                | ScalarType::Variant
                                | ScalarType::Record
                                | ScalarType::Array
                        )
                })
                .ok_or_else(|| format!("typed entry carrier {ty:?} is not supported"))
        };
        if func.param_count > argument_limit
            || func.param_count > func.locals.len()
            || func
                .locals
                .len()
                .saturating_add(func.temps.len())
                .saturating_add(globals.len())
                > slot_limit
        {
            return Err("typed entry exceeds argument/slot bounds".into());
        }
        let result = func.return_local;
        if result.is_none() && !owning {
            return Err("typed scalar entries require a Function/Property Get result".into());
        }
        let mut calls = std::collections::HashMap::new();
        for inst in func.blocks.iter().flat_map(|b| &b.instrs) {
            if let OxInst::CallProc { proc, .. } = inst {
                let callee = functions.get(proc.0).ok_or("invalid typed callee")?;
                if callee.param_count > argument_limit || callee.param_count > callee.locals.len() {
                    return Err("typed callee exceeds argument bounds".into());
                }
                let mut by_ref = Vec::new();
                let mut optional = Vec::new();
                let mut parameters = Vec::new();
                for local in &callee.locals[..callee.param_count] {
                    let p = local.param.ok_or("missing callee parameter metadata")?;
                    by_ref.push(p.by_ref);
                    optional.push(p.optional);
                    parameters.push(scalar(&local.ty)?);
                }
                let result = callee
                    .return_local
                    .map(|id| {
                        callee
                            .locals
                            .get(id.0)
                            .ok_or_else(|| "invalid callee result".to_string())
                            .and_then(|l| scalar(&l.ty))
                    })
                    .transpose()?;
                calls.insert(
                    *proc,
                    CallSignature {
                        parameters,
                        by_ref,
                        optional,
                        result,
                    },
                );
            }
        }
        if owning
            && classes.iter().flat_map(|class| &class.fields).any(|field| {
                if native_numbers {
                    native_scalar_type(&field.ty).is_none()
                } else {
                    !matches!(
                        field.ty,
                        OxTy::Long
                            | OxTy::Double
                            | OxTy::Bool
                            | OxTy::Str
                            | OxTy::Variant
                            | OxTy::Object(_)
                            | OxTy::Array(_, _)
                            | OxTy::Record(_)
                    )
                }
            })
        {
            return Err("typed source class field has no owning carrier".into());
        }
        let mut class_calls = Vec::new();
        if owning {
            for inst in func.blocks.iter().flat_map(|b| &b.instrs) {
                let translated = if native_numbers && index_member_call(inst).is_some() {
                    index_member_call(inst)
                } else if let OxInst::CallByName {
                    dst,
                    object,
                    name: OxOperand::Const(OxConst::Str(name)),
                    calltype,
                    args,
                } = inst
                {
                    let kind = match calltype {
                        OxOperand::Const(OxConst::I16(value)) => Some(i32::from(*value)),
                        OxOperand::Const(OxConst::I32(value)) => Some(*value),
                        _ => None,
                    };
                    kind.and_then(|kind| match kind {
                        1 => Some(calcvba_com::TypeLibMemberInvokeKind::Method),
                        2 => Some(calcvba_com::TypeLibMemberInvokeKind::PropertyGet),
                        4 => Some(calcvba_com::TypeLibMemberInvokeKind::PropertyPut),
                        8 => Some(calcvba_com::TypeLibMemberInvokeKind::PropertyPutRef),
                        _ => None,
                    })
                    .map(|invoke_kind| OxInst::ComCallLate {
                        dst: *dst,
                        recv: object.clone(),
                        name: name.clone(),
                        invoke_kind,
                        default_member: false,
                        args: args.clone(),
                    })
                } else {
                    None
                };
                let lookup = translated.as_ref().unwrap_or(inst);
                if let OxInst::ComCallLate {
                    recv: OxOperand::Use(place),
                    name,
                    invoke_kind,
                    default_member,
                    args,
                    ..
                } = lookup
                {
                    let ty = match place {
                        OxPlace::Local(id) => func.locals.get(id.0).map(|l| &l.ty),
                        OxPlace::Temp(id) => func.temps.get(id.0),
                        OxPlace::Global(id) => globals.get(id.0).map(|g| &g.ty),
                    };
                    if matches!(inst, OxInst::ArrayGet { .. } | OxInst::ArraySet { .. })
                        && !matches!(ty, Some(OxTy::Variant | OxTy::Object(_)))
                    {
                        continue;
                    }
                    let declared = match ty {
                        Some(OxTy::Object(calcvba_oxir::ObjClass::Class(id))) => Some(*id),
                        _ => None,
                    };
                    let interface_name = match ty {
                        Some(OxTy::Object(calcvba_oxir::ObjClass::Iface(id))) => {
                            interfaces.get(id.0).map(|interface| interface.name())
                        }
                        _ => declared
                            .and_then(|id| classes.get(id.0))
                            .filter(|candidate| {
                                classes.iter().any(|class| {
                                    class
                                        .implements
                                        .iter()
                                        .any(|name| name.eq_ignore_ascii_case(&candidate.name))
                                })
                            })
                            .map(|class| class.name.as_str()),
                    };
                    for (index, class) in classes.iter().enumerate() {
                        let id = calcvba_oxir::ClassId(index);
                        if let Some(interface) = interface_name {
                            if !class.name.eq_ignore_ascii_case(interface)
                                && !class
                                    .implements
                                    .iter()
                                    .any(|name| name.eq_ignore_ascii_case(interface))
                            {
                                continue;
                            }
                        } else if declared.is_some_and(|declared| declared != id) {
                            continue;
                        }
                        let methods: Vec<_> = class
                            .methods
                            .iter()
                            .filter(|m| {
                                (if matches!(inst, OxInst::CallByName { .. }) {
                                    m.externally_visible
                                } else {
                                    declared.is_some()
                                        || interface_name.is_some()
                                        || m.externally_visible
                                }) && (if *default_member {
                                    m.is_default_member
                                } else {
                                    m.name.eq_ignore_ascii_case(name)
                                }) && (matches!(inst, OxInst::ArraySet { .. })
                                    && matches!(
                                        m.kind,
                                        calcvba_bundle::ProjectMemberKind::PropertySet
                                    )
                                    || matches!(
                                        (m.kind, invoke_kind),
                                        (
                                            calcvba_bundle::ProjectMemberKind::Method,
                                            calcvba_com::TypeLibMemberInvokeKind::Method
                                        ) | (
                                            calcvba_bundle::ProjectMemberKind::PropertyGet,
                                            calcvba_com::TypeLibMemberInvokeKind::PropertyGet
                                        ) | (
                                            calcvba_bundle::ProjectMemberKind::PropertyLet,
                                            calcvba_com::TypeLibMemberInvokeKind::PropertyPut
                                        ) | (
                                            calcvba_bundle::ProjectMemberKind::PropertySet,
                                            calcvba_com::TypeLibMemberInvokeKind::PropertyPutRef
                                        )
                                    ))
                            })
                            .collect();
                        if methods.is_empty() {
                            if declared.is_some()
                                && interface_name.is_none()
                                && !matches!(inst, OxInst::CallByName { .. })
                            {
                                return Err("unresolved typed source class member".into());
                            }
                            continue;
                        }
                        if args.iter().any(|a| {
                            !matches!(
                                a,
                                calcvba_oxir::OxCallArg::Operand(_)
                                    | calcvba_oxir::OxCallArg::ByRef(_)
                                    | calcvba_oxir::OxCallArg::Omitted
                            )
                        }) {
                            return Err("named source class arguments remain open".into());
                        }
                        for method in methods {
                            calls.insert(
                                method.proc,
                                call_signature(
                                    functions.get(method.proc.0).ok_or("invalid class method")?,
                                )?,
                            );
                            class_calls.push((inst.clone(), method.proc, id));
                        }
                    }
                }
            }
        }
        let mut class_news = std::collections::HashMap::new();
        if owning {
            for class in classes {
                if let Some(proc) = class.terminate {
                    calls.insert(
                        proc,
                        call_signature(functions.get(proc.0).ok_or("invalid class finalizer")?)?,
                    );
                }
            }
            for inst in func.blocks.iter().flat_map(|b| &b.instrs) {
                if let OxInst::NewObject { class, .. } | OxInst::Predeclared { class, .. } = inst {
                    let plan = classes.get(class.0).ok_or("invalid class constructor")?;
                    if matches!(inst, OxInst::Predeclared { .. })
                        && !plan
                            .fields
                            .iter()
                            .any(|field| field.token == calcvba_bundle::HOST_DOCUMENT_FIELD)
                    {
                        return Err("owning predeclared instances require an explicitly bound host document".into());
                    }
                    class_news.insert(*class, plan.initialize);
                    if let Some(proc) = plan.initialize {
                        calls.insert(
                            proc,
                            call_signature(
                                functions.get(proc.0).ok_or("invalid class initializer")?,
                            )?,
                        );
                    }
                }
            }
        }
        let mut physical = func
            .locals
            .iter()
            .map(|l| scalar(&l.ty).map(Some))
            .collect::<Result<Vec<_>, _>>()?;
        let escapes = calcvba_oxir::escape_facts(func);
        let mut integer_writers = std::collections::HashMap::<usize, usize>::new();
        for inst in func.blocks.iter().flat_map(|block| &block.instrs) {
            if let Some(OxPlace::Temp(id)) = instruction_destination(inst) {
                *integer_writers.entry(id.0).or_default() += 1;
            }
        }
        let mut integer_constants = std::collections::HashMap::<usize, i16>::new();
        loop {
            let mut changed = false;
            for inst in func.blocks.iter().flat_map(|b| &b.instrs) {
                let (dst, value) = match inst {
                    OxInst::Assign {
                        dst: OxPlace::Temp(dst),
                        value: OxOperand::Const(OxConst::I16(value)),
                    } => (*dst, Some(*value)),
                    OxInst::Neg {
                        dst: OxPlace::Temp(dst),
                        src,
                        ..
                    } => {
                        let value = match src {
                            OxOperand::Const(OxConst::I16(value)) => Some(*value),
                            OxOperand::Use(OxPlace::Temp(id)) => {
                                integer_constants.get(&id.0).copied()
                            }
                            _ => None,
                        }
                        .and_then(i16::checked_neg);
                        (*dst, value)
                    }
                    _ => continue,
                };
                if owning
                    && func.temps.get(dst.0) == Some(&OxTy::Integer)
                    && integer_writers.get(&dst.0) == Some(&1)
                    && !escapes.temps[dst.0]
                    && let Some(value) = value
                    && integer_constants.insert(dst.0, value).is_none()
                {
                    changed = true;
                }
            }
            if !changed {
                break;
            }
        }
        let literal_integer_temps: std::collections::HashSet<_> = integer_constants
            .keys()
            .map(|id| func.locals.len() + id)
            .collect();
        physical.extend(func.temps.iter().enumerate().map(|(index, ty)| {
            if literal_integer_temps.contains(&(func.locals.len() + index)) {
                return Some(ScalarType::Long);
            }
            if native_numbers && matches!(ty, OxTy::FixedStr(_)) {
                let destination = OxPlace::Temp(calcvba_oxir::TempId(index));
                let writers: Vec<_> = func
                    .blocks
                    .iter()
                    .flat_map(|b| &b.instrs)
                    .filter(|i| instruction_destination(i) == Some(destination))
                    .collect();
                if !writers.is_empty()
                    && writers.iter().all(|i| {
                        matches!(
                            i,
                            OxInst::RecordGet { .. }
                                | OxInst::Coerce {
                                    target: calcvba_oxir::OxCoerceTarget::FixedStr(_),
                                    ..
                                }
                        )
                    })
                {
                    return Some(ScalarType::String);
                }
            }
            if !owning && *ty == OxTy::Variant {
                None
            } else {
                storage_type(ty)
            }
        }));
        let global_base = physical.len();
        let unsupported_globals = globals
            .iter()
            .enumerate()
            .filter(|(_, g)| storage_type(&g.ty).is_none())
            .map(|(i, _)| i)
            .collect();
        // Unused globals need not share the current frame's carrier family.
        // Access to these placeholders is refused by slot() before execution.
        physical.extend(
            globals
                .iter()
                .map(|g| Some(storage_type(&g.ty).unwrap_or(ScalarType::Long))),
        );
        // Widening arithmetic is semantically Variant, but numeric operands
        // with at least one Double prove a Double payload. Specialize only compiler temporaries whose writers
        // prove Double; never a user-declared Variant or Long-only widening.
        loop {
            let mut changed = false;
            for block in &func.blocks {
                for inst in &block.instrs {
                    let source_result = match inst {
                        OxInst::CallNative {
                            dst: Some(OxPlace::Temp(id)),
                            callee:
                                calcvba_oxir::OxNativeCallee::Builtin(
                                    calcvba_bundle::NativeImplId::CreateObject,
                                ),
                            ..
                        } if owning => Some((*id, Some(ScalarType::HostObject))),
                        OxInst::Not {
                            dst: OxPlace::Temp(id),
                            src,
                        } => physical_operand_type(func, &physical, global_base, src)
                            .filter(|ty| matches!(ty, ScalarType::Boolean | ScalarType::Long))
                            .map(|ty| (*id, Some(ty))),
                        OxInst::Logical {
                            dst: OxPlace::Temp(id),
                            lhs,
                            rhs,
                            ..
                        } => physical_operand_type(func, &physical, global_base, lhs)
                            .zip(physical_operand_type(func, &physical, global_base, rhs))
                            .and_then(|(left, right)| logical_type(left, right))
                            .map(|ty| (*id, Some(ty))),
                        OxInst::CallProc {
                            dst: Some(OxPlace::Temp(id)),
                            proc,
                            ..
                        } => Some((*id, calls.get(proc).and_then(|c| c.result))),
                        OxInst::CallExtern {
                            dst: Some(OxPlace::Temp(id)),
                            import,
                            ..
                        } => Some((
                            *id,
                            extern_calls.get(import).and_then(|c| c.signature.result),
                        )),
                        _ => None,
                    };
                    if let Some((id, next)) = source_result {
                        let index = func
                            .locals
                            .len()
                            .checked_add(id.0)
                            .ok_or("invalid temp index")?;
                        let slot = physical.get_mut(index).ok_or("invalid temp index")?;
                        if slot.is_none() || *slot == Some(ScalarType::Variant) {
                            changed |= next.is_some() && next != *slot;
                            *slot = next;
                        }
                    }
                    if let OxInst::ComCallEarly {
                        dst: Some(OxPlace::Temp(id)),
                        method,
                        ..
                    } = inst
                        && host_calls.contains_key(method)
                    {
                        let index = func
                            .locals
                            .len()
                            .checked_add(id.0)
                            .ok_or("invalid temp index")?;
                        let slot = physical.get_mut(index).ok_or("invalid temp index")?;
                        let next = Some(host_calls[method].result);
                        if (slot.is_none() || *slot == Some(ScalarType::Variant)) && next != *slot {
                            *slot = next;
                            changed = true;
                        }
                    }
                    if let OxInst::Arith {
                        dst: OxPlace::Temp(id),
                        lhs,
                        rhs,
                        mode: NumericMode::Widening,
                        ..
                    } = inst
                    {
                        let index = func
                            .locals
                            .len()
                            .checked_add(id.0)
                            .ok_or("invalid temp index")?;
                        let Some(slot) = physical.get(index) else {
                            return Err("invalid temp index".into());
                        };
                        if slot.is_some() && *slot != Some(ScalarType::Variant) {
                            continue;
                        }
                        let ty = |o: &OxOperand| match o {
                            OxOperand::Use(OxPlace::Local(i)) if i.0 < func.locals.len() => {
                                physical[i.0]
                            }
                            OxOperand::Use(OxPlace::Temp(i)) => func
                                .locals
                                .len()
                                .checked_add(i.0)
                                .and_then(|i| physical.get(i))
                                .copied()
                                .flatten(),
                            OxOperand::Use(OxPlace::Global(i)) => {
                                physical.get(global_base + i.0).copied().flatten()
                            }
                            OxOperand::Const(OxConst::F64(_)) => Some(ScalarType::Double),
                            OxOperand::Const(OxConst::I16(_) | OxConst::I32(_)) => {
                                Some(ScalarType::Long)
                            }
                            _ => None,
                        };
                        let (left, right) = (ty(lhs), ty(rhs));
                        let numeric =
                            |ty| matches!(ty, Some(ScalarType::Long | ScalarType::Double));
                        if numeric(left)
                            && numeric(right)
                            && (left == Some(ScalarType::Double)
                                || right == Some(ScalarType::Double))
                        {
                            physical[index] = Some(ScalarType::Double);
                            changed = true;
                        }
                    }
                }
            }
            if !changed {
                break;
            }
        }
        if let Some(index) = physical.iter().position(Option::is_none) {
            let ty = func
                .locals
                .iter()
                .map(|l| &l.ty)
                .chain(func.temps.iter())
                .chain(globals.iter().map(|g| &g.ty))
                .nth(index);
            return Err(format!(
                "typed entry {} has an unsupported carrier at slot {index}: {ty:?}",
                func.name
            ));
        }
        let slots = physical
            .into_iter()
            .collect::<Option<Vec<_>>>()
            .ok_or("typed entry has an unproven carrier")?;
        let parameters = slots[..func.param_count].to_vec();
        let mut record_slots = std::collections::HashMap::new();
        let mut array_slots = std::collections::HashMap::new();
        let mut array_record_slots = std::collections::HashMap::new();
        for (index, ty) in func
            .locals
            .iter()
            .map(|l| &l.ty)
            .chain(func.temps.iter())
            .chain(globals.iter().map(|g| &g.ty))
            .enumerate()
        {
            if let OxTy::Record(id) = ty {
                let fields = record_layouts
                    .get(id.0)
                    .ok_or("missing typed record layout")?;
                record_slots.insert(index, calcvba_bundle::vba_record_layout_for_fields(fields)?);
            }
            if let OxTy::Array(element, _) = ty {
                if let OxTy::Record(id) = element.as_ref() {
                    let fields = record_layouts
                        .get(id.0)
                        .ok_or("missing array record layout")?;
                    array_record_slots.insert(
                        index,
                        (*id, calcvba_bundle::vba_record_layout_for_fields(fields)?),
                    );
                }
                use calcvba_runtime::safe_array::*;
                let vt = match element.as_ref() {
                    OxTy::Variant => VT_VARIANT_VALUE,
                    OxTy::Byte => VT_UI1_VALUE,
                    OxTy::Integer => VT_I2_VALUE,
                    OxTy::Long => VT_I4_VALUE,
                    OxTy::LongLong => VT_I8_VALUE,
                    OxTy::Single => VT_R4_VALUE,
                    OxTy::Double => VT_R8_VALUE,
                    OxTy::Currency => VT_CY_VALUE,
                    OxTy::Date => VT_DATE_VALUE,
                    OxTy::Decimal => VT_DECIMAL_VALUE,
                    OxTy::Bool => VT_BOOL_VALUE,
                    OxTy::Str | OxTy::FixedStr(_) => VT_BSTR_VALUE,
                    OxTy::Object(_) => VT_DISPATCH_VALUE,
                    OxTy::Record(_) => VT_RECORD_VALUE,
                    _ => return Err("unsupported typed array element metadata".into()),
                };
                array_slots.insert(index, vt);
            }
        }
        let slot_iids = func
            .locals
            .iter()
            .map(|l| &l.ty)
            .chain(func.temps.iter())
            .chain(globals.iter().map(|g| &g.ty))
            .enumerate()
            .filter_map(|(index, ty)| {
                let OxTy::Object(calcvba_oxir::ObjClass::ComIface(id)) = ty else {
                    return None;
                };
                Some((|| {
                    let Some(calcvba_oxir::ComInterface::Com(meta)) = interfaces.get(id.0) else {
                        return Err("typed object slot has no interface descriptor");
                    };
                    Ok((
                        index,
                        iid_bytes(
                            meta.iid
                                .ok_or("typed object slot has no interface identity")?,
                        ),
                    ))
                })())
            })
            .collect::<Result<std::collections::HashMap<_, _>, _>>()?;
        if func.locals[..func.param_count].iter().any(|local| {
            matches!(local.ty, OxTy::Object(calcvba_oxir::ObjClass::Iface(_)))
                || (!owning && matches!(local.ty, OxTy::Object(calcvba_oxir::ObjClass::Class(_))))
        }) {
            return Err("typed project-object parameter ABI remains open".into());
        }
        let parameter_iids = (0..func.param_count)
            .map(|i| slot_iids.get(&i).copied())
            .collect();
        if !owning && parameters.contains(&ScalarType::Boolean) {
            return Err("typed Boolean parameter ABI remains open".into());
        }
        for local in func.locals.iter().take(func.param_count) {
            let p = local.param.ok_or("missing parameter metadata")?;
            if (p.by_ref || p.optional || p.variadic) && !owning {
                return Err("typed entry requires required ByVal parameters".into());
            }
        }
        let specialized_variant_temps = func
            .temps
            .iter()
            .enumerate()
            .filter(|(i, ty)| {
                **ty == OxTy::Variant && slots[func.locals.len() + i] != ScalarType::Variant
            })
            .map(|(i, _)| func.locals.len() + i)
            .collect();
        let mut assignment_iids = std::collections::HashMap::new();
        let mut assignment_classes = std::collections::HashMap::new();
        if owning {
            for inst in func.blocks.iter().flat_map(|b| &b.instrs) {
                if let OxInst::ValidateAssignment {
                    target_type_name,
                    target_kind,
                    ..
                } = inst
                {
                    if *target_kind == calcvba_bundle::AssignmentTargetKind::Scalar {
                        continue;
                    }
                    let bare = target_type_name
                        .rsplit('.')
                        .next()
                        .unwrap_or(target_type_name);
                    if let Some(class) = classes.iter().find(|class| {
                        class.name.eq_ignore_ascii_case(target_type_name)
                            || (!target_type_name.contains('.')
                                && class
                                    .name
                                    .rsplit('.')
                                    .next()
                                    .is_some_and(|name| name.eq_ignore_ascii_case(bare)))
                            || program.is_some_and(|p| {
                                format!("{}.{}", p.unit_name, class.name)
                                    .eq_ignore_ascii_case(target_type_name)
                            })
                    }) {
                        assignment_classes.insert(target_type_name.clone(), class.name.clone());
                        assignment_iids.insert(target_type_name.clone(), None);
                        continue;
                    }
                    let iid = if bare.is_empty()
                        || bare.eq_ignore_ascii_case("Object")
                        || bare.eq_ignore_ascii_case("Variant")
                        || target_type_name.eq_ignore_ascii_case("VBA.Collection")
                        || target_type_name.eq_ignore_ascii_case("Collection")
                    {
                        None
                    } else {
                        let exact = interfaces
                            .iter()
                            .any(|i| i.name().eq_ignore_ascii_case(target_type_name));
                        let candidates: Vec<_> = interfaces
                            .iter()
                            .filter(|i| {
                                if exact {
                                    i.name().eq_ignore_ascii_case(target_type_name)
                                } else {
                                    !target_type_name.contains('.')
                                        && i.name()
                                            .rsplit('.')
                                            .next()
                                            .is_some_and(|n| n.eq_ignore_ascii_case(bare))
                                }
                            })
                            .filter_map(|i| match i {
                                calcvba_oxir::ComInterface::Com(meta) => meta.iid,
                                _ => None,
                            })
                            .collect();
                        let first = *candidates
                            .first()
                            .ok_or("owning assignment has no interface identity")?;
                        if candidates.iter().any(|iid| *iid != first) {
                            return Err("ambiguous owning assignment interface".into());
                        }
                        let mut bytes = [0; 16];
                        bytes[..4].copy_from_slice(&first.data1.to_le_bytes());
                        bytes[4..6].copy_from_slice(&first.data2.to_le_bytes());
                        bytes[6..8].copy_from_slice(&first.data3.to_le_bytes());
                        bytes[8..].copy_from_slice(&first.data4);
                        Some(bytes)
                    };
                    assignment_iids.insert(target_type_name.clone(), iid);
                }
            }
        }
        let layout = Self {
            library_news: program.map(|p| p.imports.iter().enumerate().filter_map(|(i, import)| {
                let library = calcvba_bundle::vba_library_bundle();
                (import.unit.eq_ignore_ascii_case("VBA") && library.exports.iter().any(|export| export.token.matches(&import.token)
                    && matches!(export.target, calcvba_bundle::ExportTarget::Class(id) if library.classes.get(id).is_some_and(|c| c.name.eq_ignore_ascii_case("Collection")))))
                    .then_some(calcvba_oxir::ImportId(i))
            }).collect()).unwrap_or_default(),
            parameter_classes: func.locals[..func.param_count]
                .iter()
                .map(|l| {
                    if let OxTy::Object(calcvba_oxir::ObjClass::Class(id)) = l.ty {
                        classes.get(id.0).map(|c| c.name.clone())
                    } else {
                        None
                    }
                })
                .collect(),
            assignment_classes,
            class_calls,
            class_news,
            param_array: func.locals[..func.param_count]
                .iter()
                .position(|l| l.param.is_some_and(|p| p.variadic)),
            record_slots,
            array_record_slots,
            array_slots,
            result: result
                .map(|id| slots.get(id.0).copied().ok_or("invalid return slot"))
                .transpose()?
                .unwrap_or(ScalarType::Void),
            has_result: result.is_some(),
            by_ref_parameters: func.locals[..func.param_count]
                .iter()
                .map(|l| l.param.is_some_and(|p| p.by_ref))
                .collect(),
            optional_parameters: func.locals[..func.param_count]
                .iter()
                .map(|l| l.param.is_some_and(|p| p.optional))
                .collect(),
            calls,
            extern_calls,
            slots,
            local_count: func.locals.len(),
            parameters,
            parameter_iids,
            slot_iids,
            return_slot: result.map_or(usize::MAX, |id| id.0),
            specialized_variant_temps,
            literal_integer_temps,
            host_calls,
            default_calls,
            default_faults,
            global_base,
            assignment_iids,
            unsupported_globals,
            owning,
            native_numbers,
        };
        if !owning
            && matches!(
                layout.result,
                ScalarType::Boolean
                    | ScalarType::HostObject
                    | ScalarType::String
                    | ScalarType::Variant
                    | ScalarType::Void
                    | ScalarType::Record
                    | ScalarType::Array
            )
        {
            return Err("typed Boolean/object result ABI remains open".into());
        }
        let ids = func
            .blocks
            .iter()
            .map(|b| b.id)
            .collect::<std::collections::HashSet<_>>();
        if ids.len() != func.blocks.len() || !ids.contains(&func.entry) {
            return Err("invalid typed block identity".into());
        }
        for block in &func.blocks {
            if block.fault_target.is_some_and(|id| !ids.contains(&id))
                || calcvba_oxir::terminator_successors(&block.terminator)
                    .iter()
                    .any(|id| !ids.contains(id))
            {
                return Err("invalid typed block target".into());
            }
            for inst in &block.instrs {
                layout
                    .check(inst)
                    .map_err(|e| format!("typed entry {}: {e}; instruction {inst:?}", func.name))?;
            }
            match &block.terminator {
                OxTerminator::Jump(_)
                | OxTerminator::Return
                | OxTerminator::FaultDispatch { .. }
                | OxTerminator::Unreachable => {}
                OxTerminator::GoSub { .. } | OxTerminator::GoSubReturn if native_numbers => {}
                OxTerminator::Resume | OxTerminator::ResumeNext | OxTerminator::ResumeLabel(_)
                    if owning => {}
                OxTerminator::Raise {
                    number,
                    source,
                    description,
                    help_file,
                    help_context,
                    ..
                } if owning => {
                    if layout.operand_type(number)? != ScalarType::Long
                        || help_context
                            .as_ref()
                            .is_some_and(|o| layout.operand_type(o) != Ok(ScalarType::Long))
                        || [source, description, help_file].iter().any(|o| {
                            o.as_ref().is_some_and(|o| {
                                !layout.native_numbers
                                    && !matches!(o, OxOperand::Const(OxConst::Str(_)))
                            })
                        })
                    {
                        return Err(
                            "typed Raise requires Long number/help and constant text".into()
                        );
                    }
                }
                OxTerminator::Branch { cond, .. }
                    if layout.operand_type(cond)? == ScalarType::Boolean => {}
                _ => return Err("typed entry terminator requires the dynamic runtime".into()),
            }
        }
        Ok(layout)
    }
    pub fn owns_values(&self) -> bool {
        self.owning
    }
    pub fn slot(&self, place: OxPlace) -> Result<usize, String> {
        let index = match place {
            OxPlace::Local(id) if id.0 < self.local_count => id.0,
            OxPlace::Temp(id) => self
                .local_count
                .checked_add(id.0)
                .ok_or("typed slot overflow")?,
            OxPlace::Global(id) if !self.unsupported_globals.contains(&id.0) => self
                .global_base
                .checked_add(id.0)
                .ok_or("typed global overflow")?,
            _ => return Err("typed entry cannot access project storage".into()),
        };
        self.slots.get(index).ok_or("typed slot out of bounds")?;
        Ok(index)
    }
    /// Integer literal temporaries may use a Long physical carrier, while their
    /// Variant subtype remains Integer (including the implicit For step).
    pub fn is_integer_literal_operand(&self, value: &OxOperand) -> bool {
        match value {
            OxOperand::Const(OxConst::I16(_)) => true,
            OxOperand::Use(place) => self
                .slot(*place)
                .is_ok_and(|slot| self.literal_integer_temps.contains(&slot)),
            _ => false,
        }
    }
    pub fn operand_type(&self, value: &OxOperand) -> Result<ScalarType, String> {
        match value {
            OxOperand::Use(p) => Ok(self.slots[self.slot(*p)?]),
            OxOperand::Const(OxConst::F64(_)) => Ok(ScalarType::Double),
            OxOperand::Const(OxConst::I16(_) | OxConst::I32(_)) => Ok(ScalarType::Long),
            OxOperand::Const(OxConst::Bool(_)) => Ok(ScalarType::Boolean),
            OxOperand::Const(OxConst::I64(_)) if self.native_numbers => Ok(ScalarType::LongLong),
            OxOperand::Const(OxConst::F32(_)) if self.native_numbers => Ok(ScalarType::Single),
            OxOperand::Const(OxConst::Currency(_)) if self.native_numbers => {
                Ok(ScalarType::Currency)
            }
            OxOperand::Const(OxConst::Date(_)) if self.native_numbers => Ok(ScalarType::Date),
            OxOperand::Const(OxConst::Nothing) => Ok(ScalarType::HostObject),
            OxOperand::Const(OxConst::Str(_)) if self.owning => Ok(ScalarType::String),
            OxOperand::Const(OxConst::Empty | OxConst::Null) if self.owning => {
                Ok(ScalarType::Variant)
            }
            _ => Err("typed operand requires the dynamic runtime".into()),
        }
    }
    pub fn comparison_type(&self, lhs: &OxOperand, rhs: &OxOperand) -> Result<ScalarType, String> {
        let left = self.operand_type(lhs)?;
        let right = self.operand_type(rhs)?;
        if self.owning
            && (left == ScalarType::HostObject
                || right == ScalarType::HostObject
                || left == ScalarType::Variant
                || right == ScalarType::Variant)
        {
            return Ok(ScalarType::Variant);
        }
        if left == ScalarType::HostObject || right == ScalarType::HostObject {
            return Err("typed object value comparison requires default-member semantics".into());
        }
        if left == right {
            if matches!(left, ScalarType::Record | ScalarType::Array) {
                return if self.owning {
                    Ok(ScalarType::Variant)
                } else {
                    Err("aggregate comparison requires value semantics".into())
                };
            }
            return Ok(left);
        }
        if matches!(
            (left, right),
            (ScalarType::Long, ScalarType::Double) | (ScalarType::Double, ScalarType::Long)
        ) {
            return Ok(ScalarType::Double);
        }
        Err("typed comparison requires compatible scalar operands".into())
    }
    fn accepts(&self, dst: OxPlace, src: &OxOperand) -> Result<(), String> {
        let slot = self.slot(dst)?;
        let to = self.slots[slot];
        let from = self.operand_type(src)?;
        if self.specialized_variant_temps.contains(&slot) && from != to {
            return Err("specialized temporary needs a proven Double writer".into());
        }
        if to == from
            || (self.native_numbers
                && matches!(
                    (to, from),
                    (ScalarType::HostObject, ScalarType::Variant)
                        | (ScalarType::Variant, ScalarType::HostObject)
                ))
            || (self.owning && !self.native_numbers && to == ScalarType::Variant)
            || (to == ScalarType::Double && from == ScalarType::Long)
            || (self.native_numbers
                && to == ScalarType::Long
                && matches!(from, ScalarType::Integer | ScalarType::Byte))
            || (self.native_numbers
                && to == ScalarType::Variant
                && matches!(
                    from,
                    ScalarType::Long
                        | ScalarType::Double
                        | ScalarType::Boolean
                        | ScalarType::Integer
                        | ScalarType::Byte
                        | ScalarType::LongLong
                        | ScalarType::Single
                        | ScalarType::Currency
                        | ScalarType::Date
                ))
        {
            Ok(())
        } else {
            Err("typed narrowing/coercion requires the dynamic runtime".into())
        }
    }
    fn check(&self, inst: &OxInst) -> Result<(), String> {
        if self.native_numbers {
            match inst {
                OxInst::CallByName {
                    dst,
                    object,
                    name,
                    calltype,
                    args,
                } => {
                    for operand in [object, name, calltype] {
                        self.operand_type(operand)?;
                    }
                    if let Some(dst) = dst {
                        self.slot(*dst)?;
                    }
                    if args.len() > 32 {
                        return Err("native CallByName exceeds 32 arguments".into());
                    }
                    for argument in args {
                        match argument {
                            calcvba_oxir::OxCallArg::Operand(operand)
                            | calcvba_oxir::OxCallArg::Named { value: operand, .. } => {
                                self.operand_type(operand)?;
                            }
                            calcvba_oxir::OxCallArg::ByRef(place) => {
                                self.slot(*place)?;
                            }
                            _ => {}
                        }
                    }
                    return Ok(());
                }
                OxInst::AsNew { place, .. } => {
                    if self.slots[self.slot(*place)?] != ScalarType::HostObject {
                        return Err("As New requires an Object cell".into());
                    }
                    return Ok(());
                }
                OxInst::TypeOfIs { dst, object, .. } => {
                    self.slot(*dst)?;
                    self.operand_type(object)?;
                    return Ok(());
                }
                OxInst::ArrayErase { array, .. } => {
                    self.slot(*array)?;
                    return Ok(());
                }
                OxInst::ErrFieldSet { src, .. } => {
                    self.operand_type(src)?;
                    return Ok(());
                }

                OxInst::ArrayLiteral {
                    dst,
                    values,
                    aliases,
                    ..
                } => {
                    self.slot(*dst)?;
                    for alias in aliases.iter().flatten() {
                        self.slot(*alias)?;
                    }
                    for value in values {
                        self.operand_type(value)?;
                    }
                    return Ok(());
                }
                OxInst::ForEachInit { iter, source } => {
                    self.slot(*iter)?;
                    self.operand_type(source)?;
                    return Ok(());
                }
                OxInst::ForEachNext {
                    iter,
                    item,
                    has_value,
                } => {
                    self.slot(*iter)?;
                    self.slot(*item)?;
                    self.slot(*has_value)?;
                    return Ok(());
                }
                OxInst::Arith { dst, lhs, rhs, .. } if value::needs_helper(inst, self) => {
                    self.slot(*dst)?;
                    self.operand_type(lhs)?;
                    self.operand_type(rhs)?;
                    return Ok(());
                }
                OxInst::CallNative {
                    dst,
                    callee: calcvba_oxir::OxNativeCallee::Builtin(_),
                    args,
                } => {
                    if let Some(dst) = dst {
                        self.slot(*dst)?;
                    }
                    for argument in args {
                        match argument {
                            calcvba_oxir::OxCallArg::Operand(operand) => {
                                self.operand_type(operand)?;
                            }
                            calcvba_oxir::OxCallArg::ByRef(place) => {
                                self.slot(*place)?;
                            }
                            calcvba_oxir::OxCallArg::Const(_)
                            | calcvba_oxir::OxCallArg::Omitted => {}
                            _ => return Err("native library named argument is not admitted".into()),
                        }
                    }
                    return Ok(());
                }
                OxInst::Pow { dst, lhs, rhs } => {
                    self.slot(*dst)?;
                    self.operand_type(lhs)?;
                    self.operand_type(rhs)?;
                    return Ok(());
                }
                _ => {}
            }
        }
        if self.native_numbers
            && let OxInst::Logical { dst, lhs, rhs, .. } = inst
        {
            self.slot(*dst)?;
            for operand in [lhs, rhs] {
                let ty = self.operand_type(operand)?;
                if !matches!(
                    ty,
                    ScalarType::Long
                        | ScalarType::Double
                        | ScalarType::Boolean
                        | ScalarType::Variant
                ) && !extended_numeric(ty)
                {
                    return Err(
                        "native logical operation requires numeric or Variant inputs".into(),
                    );
                }
            }
            return Ok(());
        }
        if self.native_numbers
            && matches!(
                inst,
                OxInst::Ptr { .. }
                    | OxInst::CallNative {
                        callee: calcvba_oxir::OxNativeCallee::Declare { .. },
                        ..
                    }
            )
        {
            return Ok(()); // The direct native backend checks the exact foreign signature.
        }
        if self.native_numbers {
            let extended = |op: &OxOperand| self.operand_type(op).is_ok_and(extended_numeric);
            let dst_extended = instruction_destination(inst).is_some_and(|dst| {
                self.slot(dst)
                    .is_ok_and(|i| extended_numeric(self.slots[i]))
            });
            let uses_extended = match inst {
                OxInst::Assign { value, .. } => extended(value),
                OxInst::Coerce { src, .. }
                | OxInst::Box { src, .. }
                | OxInst::Unbox { src, .. }
                | OxInst::Truthy { src, .. }
                | OxInst::Neg { src, .. }
                | OxInst::Not { src, .. } => extended(src),
                OxInst::Arith { lhs, rhs, .. }
                | OxInst::Div { lhs, rhs, .. }
                | OxInst::Compare { lhs, rhs, .. } => extended(lhs) || extended(rhs),
                _ => false,
            };
            if (dst_extended || uses_extended)
                && matches!(
                    inst,
                    OxInst::Assign { .. }
                        | OxInst::Coerce { .. }
                        | OxInst::Box { .. }
                        | OxInst::Unbox { .. }
                        | OxInst::Truthy { .. }
                        | OxInst::Neg { .. }
                        | OxInst::Not { .. }
                        | OxInst::Arith { .. }
                        | OxInst::Div { .. }
                        | OxInst::Compare { .. }
                )
            {
                return Ok(());
            }
        }
        if let Some(dst) = instruction_destination(inst)
            && self.literal_integer_temps.contains(&self.slot(dst)?)
            && !matches!(
                inst,
                OxInst::Assign {
                    value: OxOperand::Const(OxConst::I16(_)),
                    ..
                } | OxInst::Neg { .. }
            )
        {
            return Err("Integer literal temporary has a conflicting writer".into());
        }
        if matches!(inst, OxInst::Neg { dst, .. } if self.literal_integer_temps.contains(&self.slot(*dst)?))
        {
            return Ok(());
        }
        if self.owning && (iteration::admit(self, inst)? || aggregate::admit(self, inst)?) {
            return Ok(());
        }
        match inst {
            OxInst::FieldArrayGet {
                dst,
                object,
                indices,
                ..
            } if self.owning => {
                self.slot(*dst)?;
                self.operand_type(object)?;
                for index in indices {
                    self.operand_type(index)?;
                }
                Ok(())
            }
            OxInst::FieldArraySet {
                object,
                indices,
                value,
                ..
            } if self.owning => {
                self.operand_type(object)?;
                self.operand_type(value)?;
                for index in indices {
                    self.operand_type(index)?;
                }
                Ok(())
            }
            OxInst::NewExtern { dst, import }
                if self.owning && self.library_news.contains(import) =>
            {
                self.slot(*dst).map(|_| ())
            }
            OxInst::CallNative {
                dst,
                callee: calcvba_oxir::OxNativeCallee::Builtin(id),
                args,
            } if self.owning && owning_builtin(*id) => {
                if let Some(dst) = dst {
                    self.slot(*dst)?;
                }
                for arg in args {
                    match arg {
                        calcvba_oxir::OxCallArg::Operand(op) => {
                            self.operand_type(op)?;
                        }
                        calcvba_oxir::OxCallArg::ByRef(p) => {
                            self.slot(*p)?;
                        }
                        calcvba_oxir::OxCallArg::Omitted | calcvba_oxir::OxCallArg::Const(_) => {}
                        _ => return Err("named native library arguments remain open".into()),
                    }
                }
                Ok(())
            }
            OxInst::Predeclared { dst, class }
                if self.owning && self.class_news.contains_key(class) =>
            {
                self.slot(*dst)?;
                Ok(())
            }
            OxInst::WithEventsGet { dst, owner, .. }
            | OxInst::WithEventsClearOwner { dst, owner }
                if self.owning =>
            {
                self.slot(*dst)?;
                if !matches!(
                    self.operand_type(owner)?,
                    ScalarType::HostObject | ScalarType::Variant
                ) {
                    return Err("invalid WithEvents owner".into());
                }
                Ok(())
            }
            OxInst::WithEventsSet {
                dst, owner, value, ..
            } if self.owning => {
                self.slot(*dst)?;
                if !matches!(
                    self.operand_type(owner)?,
                    ScalarType::HostObject | ScalarType::Variant
                ) || !matches!(
                    self.operand_type(value)?,
                    ScalarType::HostObject | ScalarType::Variant
                ) {
                    return Err("invalid WithEvents assignment".into());
                }
                Ok(())
            }
            OxInst::ComCallLate { dst, .. }
                if self.owning && self.class_calls.iter().any(|(site, _, _)| site == inst) =>
            {
                if let Some(dst) = dst {
                    self.slot(*dst)?;
                }
                Ok(())
            }
            OxInst::ComCallLate {
                dst, recv, args, ..
            } if self.owning => {
                if !matches!(
                    self.operand_type(recv)?,
                    ScalarType::HostObject | ScalarType::Variant
                ) || args.len()
                    > if self.native_numbers {
                        32
                    } else {
                        MAX_TYPED_ARGS
                    }
                {
                    return Err("invalid dynamic member receiver/arity".into());
                }
                if let Some(dst) = dst {
                    self.slot(*dst)?;
                }
                for arg in args {
                    match arg {
                        calcvba_oxir::OxCallArg::Operand(op)
                        | calcvba_oxir::OxCallArg::Named { value: op, .. } => {
                            self.operand_type(op)?;
                        }
                        calcvba_oxir::OxCallArg::ByRef(place) => {
                            self.slot(*place)?;
                        }
                        calcvba_oxir::OxCallArg::Omitted | calcvba_oxir::OxCallArg::Const(_) => {}
                    }
                }
                Ok(())
            }
            OxInst::NewObject { dst, class }
                if self.owning
                    && self.class_news.contains_key(class)
                    && self.slots[self.slot(*dst)?] == ScalarType::HostObject =>
            {
                Ok(())
            }
            OxInst::FieldGet { dst, object, .. } if self.owning => {
                self.slot(*dst)?;
                self.operand_type(object).map(|_| ())
            }
            OxInst::FieldSet { object, value, .. } if self.owning => {
                self.operand_type(object)?;
                self.operand_type(value).map(|_| ())
            }
            OxInst::CallProc { dst, args, .. } | OxInst::CallExtern { dst, args, .. } => {
                let value_builtin = matches!(inst, OxInst::CallExtern { import, .. }
                    if self.extern_calls.get(import).is_some_and(|call| matches!(call.target, LinkedTarget::Builtin(_))));
                let call = match inst {
                    OxInst::CallProc { proc, .. } => self.calls.get(proc),
                    OxInst::CallExtern { import, .. } => {
                        self.extern_calls.get(import).map(|c| &c.signature)
                    }
                    _ => None,
                }
                .ok_or("unplanned typed call")?;
                if args.len() > call.parameters.len()
                    || call.parameters[args.len()..]
                        .iter()
                        .zip(&call.optional[args.len()..])
                        .any(|(ty, opt)| !*opt || *ty != ScalarType::Variant)
                {
                    return Err("typed call arity mismatch".into());
                }
                for (i, ((arg, ty), by_ref)) in args
                    .iter()
                    .zip(&call.parameters)
                    .zip(&call.by_ref)
                    .enumerate()
                {
                    if matches!(arg, calcvba_oxir::OxArg::Omitted)
                        && call.optional[i]
                        && *ty == ScalarType::Variant
                    {
                        continue;
                    }
                    let (from, alias) = match arg {
                        calcvba_oxir::OxArg::ByVal(op) => (self.operand_type(op)?, false),
                        calcvba_oxir::OxArg::ByRef(p) => (self.slots[self.slot(*p)?], true),
                        _ => return Err("omitted typed arguments remain open".into()),
                    };
                    let private_variant =
                        (self.owning || value_builtin) && !alias && *ty == ScalarType::Variant;
                    let private_object = self.owning
                        && !alias
                        && from == ScalarType::Variant
                        && *ty == ScalarType::HostObject;
                    let native_numeric = self.native_numbers
                        && !alias
                        && (extended_numeric(from)
                            || matches!(
                                from,
                                ScalarType::Long | ScalarType::Double | ScalarType::Boolean
                            ))
                        && (extended_numeric(*ty)
                            || matches!(
                                ty,
                                ScalarType::Long | ScalarType::Double | ScalarType::Boolean
                            ));
                    if (from != *ty && !private_variant && !private_object && !native_numeric)
                        || (alias && !by_ref)
                    {
                        return Err("typed call carrier/passing mismatch".into());
                    }
                }
                if let Some(dst) = dst
                    && Some(self.slots[self.slot(*dst)?]) != call.result
                {
                    return Err("typed call result mismatch".into());
                }
                Ok(())
            }
            OxInst::CallNative {
                dst: Some(dst),
                callee:
                    calcvba_oxir::OxNativeCallee::Builtin(
                        native @ (calcvba_bundle::NativeImplId::IsMissing
                        | calcvba_bundle::NativeImplId::IsArray),
                    ),
                args,
            } if self.owning => {
                if args.len() != 1
                    || !matches!(
                        self.slots[self.slot(*dst)?],
                        ScalarType::Boolean | ScalarType::Variant
                    )
                {
                    return Err("invalid typed IsMissing signature".into());
                }
                match &args[0] {
                    calcvba_oxir::OxCallArg::Operand(op)
                        if *native == calcvba_bundle::NativeImplId::IsArray
                            || self.operand_type(op)? == ScalarType::Variant =>
                    {
                        Ok(())
                    }
                    calcvba_oxir::OxCallArg::ByRef(p)
                        if *native == calcvba_bundle::NativeImplId::IsArray
                            || self.slots[self.slot(*p)?] == ScalarType::Variant =>
                    {
                        Ok(())
                    }
                    _ => Err("IsMissing requires a retained Variant".into()),
                }
            }
            OxInst::SetErrorHandler(_) | OxInst::ClearErr | OxInst::SetLineNumber { .. }
                if self.owning =>
            {
                Ok(())
            }
            OxInst::ErrFieldGet {
                dst,
                field:
                    calcvba_oxir::ErrField::Number
                    | calcvba_oxir::ErrField::HelpContext
                    | calcvba_oxir::ErrField::LastDllError,
            }
            | OxInst::ErlGet { dst }
                if self.owning && self.slots[self.slot(*dst)?] == ScalarType::Long =>
            {
                Ok(())
            }
            OxInst::ErrFieldGet {
                dst,
                field:
                    calcvba_oxir::ErrField::Description
                    | calcvba_oxir::ErrField::Source
                    | calcvba_oxir::ErrField::HelpFile,
            } if self.owning && self.slots[self.slot(*dst)?] == ScalarType::String => Ok(()),
            OxInst::ComCallEarly {
                dst,
                method,
                recv,
                args,
                ..
            } => {
                let plan = self.host_calls.get(method).ok_or("unplanned host slot")?;
                let arity = if let Some(native) = &plan.native {
                    calcvba_com::typed_slots::plan::visible_signature(native, args.len()).is_some()
                        && args.len() <= plan.parameters.len()
                } else {
                    args.len() == plan.parameters.len()
                };
                if self.operand_type(recv)? != ScalarType::HostObject || !arity {
                    return Err("typed host call signature mismatch".into());
                }
                match dst {
                    Some(dst) if self.slots[self.slot(*dst)?] == plan.result => {}
                    None => {}
                    _ => return Err("typed host result destination mismatch".into()),
                }
                for (index, (arg, ty)) in args.iter().zip(&plan.parameters).enumerate() {
                    let from = match arg {
                        calcvba_oxir::OxCallArg::Operand(op) => self.operand_type(op)?,
                        calcvba_oxir::OxCallArg::ByRef(p) => self.slots[self.slot(*p)?],
                        calcvba_oxir::OxCallArg::Const(_) => ScalarType::Long,
                        calcvba_oxir::OxCallArg::Omitted
                            if plan.native.as_ref().is_some_and(|member| {
                                let indices =
                                    calcvba_com::typed_slots::plan::visible_indices(member);
                                calcvba_com::typed_slots::plan::default_value(
                                    member,
                                    indices[index],
                                )
                                .is_some()
                            }) =>
                        {
                            ty.carrier()
                        }
                        _ => {
                            return Err(
                                "typed host argument is not a required positional value".into()
                            );
                        }
                    };
                    let boxed_native_input = plan.native.is_some()
                        && ty.carrier() == ScalarType::Variant
                        && matches!(arg, calcvba_oxir::OxCallArg::Operand(_));
                    if from != ty.carrier() && !boxed_native_input {
                        return Err("typed host argument carrier mismatch".into());
                    }
                }
                Ok(())
            }
            OxInst::StmtBoundary {
                clear_temps_from, ..
            } if *clear_temps_from <= self.global_base - self.local_count => Ok(()),
            OxInst::ValidateAssignment {
                src,
                intent:
                    calcvba_bundle::AssignmentIntent::Let | calcvba_bundle::AssignmentIntent::Implicit,
                target_kind: calcvba_bundle::AssignmentTargetKind::Scalar,
                ..
            } if self.operand_type(src)? == ScalarType::Record => Ok(()),
            OxInst::ValidateAssignment {
                src,
                intent: calcvba_bundle::AssignmentIntent::Set,
                target_kind:
                    calcvba_bundle::AssignmentTargetKind::Object
                    | calcvba_bundle::AssignmentTargetKind::Variant,
                target_type_name,
                ..
            } if self.assignment_iids.contains_key(target_type_name)
                && matches!(
                    self.operand_type(src)?,
                    ScalarType::HostObject | ScalarType::Variant
                ) =>
            {
                Ok(())
            }
            OxInst::DrainTerminations => Ok(()),
            OxInst::VariantChanged {
                dst,
                current,
                original,
            } if self.native_numbers => {
                if self.slots[self.slot(*dst)?] != ScalarType::Boolean {
                    return Err("VariantChanged requires Boolean output".into());
                }
                self.operand_type(current)?;
                self.operand_type(original).map(|_| ())
            }
            OxInst::ValidateAssignment { src, .. } if self.native_numbers => {
                self.operand_type(src).map(|_| ())
            }
            OxInst::Assign { dst, value } => self.accepts(*dst, value),
            OxInst::Box { dst, src, .. }
                if self.owning && self.slots[self.slot(*dst)?] == ScalarType::Variant =>
            {
                self.operand_type(src).map(|_| ())
            }
            OxInst::Unbox { dst, src, .. }
                if self.owning
                    && (self.operand_type(src)? == ScalarType::Variant
                        || self.operand_type(src)? == self.slots[self.slot(*dst)?]) =>
            {
                self.slot(*dst).map(|_| ())
            }
            OxInst::Coerce { dst, src, target }
                if self.owning
                    && (matches!(
                        target,
                        OxCoerceTarget::Numeric(N::Long | N::Double | N::Boolean)
                            | OxCoerceTarget::Str
                            | OxCoerceTarget::ImplicitVariant
                    ) || (self.native_numbers
                        && matches!(target, OxCoerceTarget::FixedStr(_)))) =>
            {
                self.slot(*dst)?;
                self.operand_type(src)?;
                Ok(())
            }
            OxInst::Concat { dst, lhs, rhs }
                if self.owning
                    && self.slots[self.slot(*dst)?] == ScalarType::String
                    && self.operand_type(lhs).is_ok()
                    && self.operand_type(rhs).is_ok() =>
            {
                Ok(())
            }
            OxInst::Coerce {
                dst,
                src,
                target: OxCoerceTarget::Numeric(N::Double),
            } if self.slots[self.slot(*dst)?] == ScalarType::Double => self.accepts(*dst, src),
            OxInst::Coerce {
                dst,
                src,
                target: OxCoerceTarget::Numeric(N::Long),
            } if self.slots[self.slot(*dst)?] == ScalarType::Long => self.accepts(*dst, src),
            OxInst::Arith {
                dst,
                lhs,
                rhs,
                mode,
                op,
            } => {
                let ty = self.slots[self.slot(*dst)?];
                if self.owning
                    && matches!(
                        (ty, mode),
                        (ScalarType::Long, NumericMode::Checked(N::Long))
                            | (ScalarType::Double, NumericMode::Checked(N::Double))
                    )
                    && matches!(op, ArithOp::Add | ArithOp::Sub | ArithOp::Mul)
                    && self.operand_type(lhs).is_ok()
                    && self.operand_type(rhs).is_ok()
                    && [self.operand_type(lhs)?, self.operand_type(rhs)?]
                        .iter()
                        .any(|t| {
                            matches!(
                                t,
                                ScalarType::Variant
                                    | ScalarType::String
                                    | ScalarType::HostObject
                                    | ScalarType::Record
                                    | ScalarType::Array
                            )
                        })
                {
                    return Ok(());
                }
                if self.owning
                    && ty == ScalarType::Variant
                    && *mode == NumericMode::Widening
                    && matches!(op, ArithOp::Add | ArithOp::Sub | ArithOp::Mul)
                    && self.operand_type(lhs).is_ok()
                    && self.operand_type(rhs).is_ok()
                {
                    return Ok(());
                }
                if !matches!(
                    (ty, mode),
                    (ScalarType::Long, NumericMode::Checked(N::Long))
                        | (
                            ScalarType::Double,
                            NumericMode::Checked(N::Double) | NumericMode::Widening
                        )
                ) || !(matches!(op, ArithOp::Add | ArithOp::Sub | ArithOp::Mul)
                    || (ty == ScalarType::Long && matches!(op, ArithOp::IntDiv | ArithOp::Mod)))
                {
                    return Err("typed arithmetic mode requires the dynamic runtime".into());
                }
                let left = self.operand_type(lhs)?;
                let right = self.operand_type(rhs)?;
                if matches!(mode, NumericMode::Widening)
                    && (left != ScalarType::Double && right != ScalarType::Double)
                {
                    return Err("typed widening needs a proven Double operand".into());
                }
                // Arithmetic operands are read after numeric promotion. The
                // writer proof concerns the result, not each operand's tag.
                for from in [left, right] {
                    if from != ty && !(ty == ScalarType::Double && from == ScalarType::Long) {
                        return Err("typed arithmetic needs compatible numeric operands".into());
                    }
                }
                Ok(())
            }
            OxInst::Div { dst, lhs, rhs } if self.slots[self.slot(*dst)?] == ScalarType::Double => {
                if self.owning {
                    self.operand_type(lhs)?;
                    self.operand_type(rhs)?;
                    return Ok(());
                }
                self.accepts(*dst, lhs)?;
                self.accepts(*dst, rhs)
            }
            OxInst::CompareObjectIs { dst, lhs, rhs }
                if self.owning && self.slots[self.slot(*dst)?] == ScalarType::Boolean =>
            {
                self.operand_type(lhs)?;
                self.operand_type(rhs)?;
                Ok(())
            }
            OxInst::Compare { dst, lhs, rhs, .. }
                if self.slots[self.slot(*dst)?] == ScalarType::Boolean
                    || (self.owning && self.slots[self.slot(*dst)?] == ScalarType::Variant) =>
            {
                self.comparison_type(lhs, rhs).map(|_| ())
            }
            OxInst::Logical { dst, lhs, rhs, .. }
                if logical_type(self.operand_type(lhs)?, self.operand_type(rhs)?)
                    == Some(self.slots[self.slot(*dst)?]) =>
            {
                Ok(())
            }
            OxInst::Not { dst, src }
                if self.slots[self.slot(*dst)?] == self.operand_type(src)?
                    && matches!(
                        self.operand_type(src)?,
                        ScalarType::Long | ScalarType::Boolean
                    ) =>
            {
                Ok(())
            }
            OxInst::Not { dst, src }
                if self.owning
                    && matches!(
                        self.slots[self.slot(*dst)?],
                        ScalarType::Boolean | ScalarType::Long | ScalarType::Variant
                    ) =>
            {
                self.operand_type(src).map(|_| ())
            }
            OxInst::Neg { dst, src, mode }
                if self.owning
                    && matches!(
                        (self.slots[self.slot(*dst)?], mode),
                        (ScalarType::Long, NumericMode::Checked(N::Long))
                            | (ScalarType::Double, NumericMode::Checked(N::Double))
                            | (ScalarType::Variant, NumericMode::Widening)
                    ) =>
            {
                self.operand_type(src).map(|_| ())
            }
            OxInst::Truthy { dst, src } if self.slots[self.slot(*dst)?] == ScalarType::Boolean => {
                if !self.owning && self.operand_type(src)? == ScalarType::HostObject {
                    Err("typed object condition requires default-member semantics".into())
                } else {
                    Ok(())
                }
            }
            _ => Err(format!(
                "typed instruction {inst:?} requires the dynamic runtime"
            )),
        }
    }
    pub fn read(&self, slots: &[ScalarBits], operand: &OxOperand, to: ScalarType) -> ScalarBits {
        let (bits, from) = match operand {
            OxOperand::Use(place) => {
                let slot = self.slot(*place).expect("admitted slot");
                (slots[slot], self.slots[slot])
            }
            OxOperand::Const(OxConst::F64(bits)) => {
                (ScalarBits::from_bits(*bits), ScalarType::Double)
            }
            OxOperand::Const(OxConst::I16(value)) => {
                (ScalarBits::long(i32::from(*value)), ScalarType::Long)
            }
            OxOperand::Const(OxConst::I32(value)) => (ScalarBits::long(*value), ScalarType::Long),
            OxOperand::Const(OxConst::Bool(value)) => (
                ScalarBits::long(if *value { -1 } else { 0 }),
                ScalarType::Boolean,
            ),
            OxOperand::Const(OxConst::Nothing) => {
                (ScalarBits::object(std::ptr::null()), ScalarType::HostObject)
            }
            _ => unreachable!("admitted operand"),
        };
        if to == ScalarType::Double && from == ScalarType::Long {
            ScalarBits::double(f64::from(bits.as_long()))
        } else {
            bits
        }
    }
}

fn iid_bytes(iid: calcvba_com::ComInterfaceIid) -> [u8; 16] {
    let mut bytes = [0; 16];
    bytes[..4].copy_from_slice(&iid.data1.to_le_bytes());
    bytes[4..6].copy_from_slice(&iid.data2.to_le_bytes());
    bytes[6..8].copy_from_slice(&iid.data3.to_le_bytes());
    bytes[8..].copy_from_slice(&iid.data4);
    bytes
}
/// Nominal source class checks never interpret a native/portable host table as
/// a project allocation, and never route an object through another session.
pub fn class_matches(object: &calcvba_runtime::ObjectRef, required: &str) -> bool {
    if !object.is_project_instance()
        || object.object_owner_id()
            != Some(calcvba_runtime::object_owner::RuntimeObjectOwner::current().id())
    {
        return false;
    }
    let descriptor = object.class_descriptor();
    descriptor.name.eq_ignore_ascii_case(required)
        || descriptor.project_identity.is_some_and(|p| {
            format!("{}.{}", p.unit_name, descriptor.name).eq_ignore_ascii_case(required)
        })
}

fn instruction_destination(inst: &OxInst) -> Option<OxPlace> {
    match inst {
        OxInst::Assign { dst, .. }
        | OxInst::Box { dst, .. }
        | OxInst::Unbox { dst, .. }
        | OxInst::Coerce { dst, .. }
        | OxInst::Truthy { dst, .. }
        | OxInst::Arith { dst, .. }
        | OxInst::Div { dst, .. }
        | OxInst::Neg { dst, .. }
        | OxInst::Not { dst, .. }
        | OxInst::Compare { dst, .. }
        | OxInst::CompareObjectIs { dst, .. }
        | OxInst::Concat { dst, .. }
        | OxInst::VariantChanged { dst, .. }
        | OxInst::ErrFieldGet { dst, .. }
        | OxInst::ErlGet { dst }
        | OxInst::NewObject { dst, .. }
        | OxInst::FieldGet { dst, .. }
        | OxInst::NewRecord { dst, .. }
        | OxInst::RecordGet { dst, .. }
        | OxInst::RecordArrayGet { dst, .. }
        | OxInst::ArrayLiteral { dst, .. }
        | OxInst::ArrayAppend { dst, .. }
        | OxInst::ArrayGet { dst, .. }
        | OxInst::Bound { dst, .. } => Some(*dst),
        OxInst::CallProc { dst, .. }
        | OxInst::CallExtern { dst, .. }
        | OxInst::CallNative { dst, .. }
        | OxInst::ComCallEarly { dst, .. }
        | OxInst::ComCallLate { dst, .. } => *dst,
        OxInst::RecordSet { record, .. }
        | OxInst::RecordLSet { record, .. }
        | OxInst::RecordArraySet { record, .. } => Some(*record),
        OxInst::ArraySet { array, .. } => Some(*array),
        _ => None,
    }
}
fn call_signature(callee: &OxFunc) -> Result<CallSignature, String> {
    if callee.param_count > MAX_TYPED_ARGS || callee.param_count > callee.locals.len() {
        return Err("typed imported argument bound".into());
    }
    let mut parameters = Vec::new();
    let mut by_ref = Vec::new();
    let mut optional = Vec::new();
    for local in &callee.locals[..callee.param_count] {
        let p = local.param.ok_or("missing imported parameter metadata")?;
        parameters.push(scalar_type(&local.ty).ok_or("unsupported imported carrier")?);
        by_ref.push(p.by_ref);
        optional.push(p.optional);
    }
    let result = callee
        .return_local
        .map(|id| {
            callee
                .locals
                .get(id.0)
                .and_then(|l| scalar_type(&l.ty))
                .ok_or("unsupported imported result")
        })
        .transpose()?;
    Ok(CallSignature {
        parameters,
        by_ref,
        optional,
        result,
    })
}

#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn owning_object_roots_keep_object_temporaries_and_declared_variant_locals() {
        use calcvba_oxir::{BlockId, LocalId, OxBlock, OxLocal, OxProgram, TempId};
        let call = |dst| OxInst::CallNative {
            dst: Some(dst),
            callee: calcvba_oxir::OxNativeCallee::Builtin(
                calcvba_bundle::NativeImplId::CreateObject,
            ),
            args: vec![calcvba_oxir::OxCallArg::Operand(OxOperand::Const(
                OxConst::Str("Host.Widget".into()),
            ))],
        };
        let program = OxProgram {
            funcs: vec![OxFunc {
                name: "Main".into(),
                kind: calcvba_bundle::ProcedureKind::Sub,
                locals: vec![OxLocal {
                    name: "declared".into(),
                    ty: OxTy::Variant,
                    array_element: None,
                    param: None,
                    escaped: false,
                }],
                temps: vec![OxTy::Variant],
                param_count: 0,
                return_local: None,
                entry: BlockId(0),
                blocks: vec![OxBlock {
                    id: BlockId(0),
                    fault_target: None,
                    instrs: vec![
                        call(OxPlace::Temp(TempId(0))),
                        call(OxPlace::Local(LocalId(0))),
                    ],
                    terminator: OxTerminator::Return,
                }],
            }],
            ..Default::default()
        };
        let layout = ScalarLayout::with_program(&program.funcs[0], &program).unwrap();
        assert_eq!(
            layout.operand_type(&OxOperand::temp(TempId(0))).unwrap(),
            ScalarType::HostObject
        );
        assert_eq!(
            layout.operand_type(&OxOperand::local(LocalId(0))).unwrap(),
            ScalarType::Variant
        );
    }

    #[test]
    fn dynamic_adapter_keeps_callee_source_and_budget_abort_does_not_replace_err() {
        let host = calcvba_hal::adapters::builder::HostBuilder::new().build();
        let mut state = crate::ExecState::new(&*host);
        state.default_error_source = "Caller".into();
        state.remaining_steps = Some(8);
        let raw = crate::exec_state_as_raw(&mut state);
        let mut context = TypedContext {
            remaining_blocks: 0,
        };
        // SAFETY: unique live state and initialized nonaliasing stack context.
        let status = unsafe { typed_context_start(raw, &mut context) };
        assert_eq!(status, crate::ST_OK);
        context.step().unwrap();
        let source = b"NumericLibrary";
        // SAFETY: the same unique state/context are live, as is the full UTF-8 source slice.
        let status =
            unsafe { typed_context_finish(raw, &context, 6, source.as_ptr(), source.len()) };
        assert_eq!(status, crate::ST_FAULT);
        assert_eq!(state.remaining_steps, Some(7));
        assert_eq!(state.err_engine.err.source, "NumericLibrary");
        assert_eq!(state.err_engine.err.description, "Overflow");
        let before = state.err_engine.err.clone();
        context.remaining_blocks = 0;
        // SAFETY: no state borrow overlaps the synchronous helper call.
        let status =
            unsafe { typed_context_finish(raw, &context, -1, source.as_ptr(), source.len()) };
        assert_eq!(status, crate::ST_ABORT);
        assert_eq!(state.err_engine.err, before);
    }

    #[test]
    fn malformed_scalar_places_and_blocks_are_refused_before_codegen() {
        use calcvba_oxir::{BlockId, LocalId, OxBlock, OxLocal, TempId};
        let mut func = OxFunc {
            name: "Malformed".into(),
            kind: calcvba_bundle::ProcedureKind::Function,
            locals: vec![OxLocal {
                name: "Malformed".into(),
                ty: OxTy::Long,
                array_element: None,
                param: None,
                escaped: false,
            }],
            temps: vec![OxTy::Long],
            param_count: 0,
            return_local: Some(LocalId(0)),
            entry: BlockId(0),
            blocks: vec![OxBlock {
                id: BlockId(0),
                instrs: vec![OxInst::Assign {
                    dst: OxPlace::Local(LocalId(0)),
                    value: OxOperand::temp(TempId(usize::MAX)),
                }],
                fault_target: None,
                terminator: OxTerminator::Return,
            }],
        };
        assert!(ScalarLayout::new(&func).is_err());
        func.blocks[0].instrs.clear();
        func.entry = BlockId(9);
        assert!(ScalarLayout::new(&func).is_err());
        func.entry = BlockId(0);
        func.param_count = 2;
        assert!(ScalarLayout::new(&func).is_err());

        // Double + integer has a proven Double result; Long-only widening
        // does not. A second writer must not invalidate a successful proof.
        func.param_count = 0;
        func.locals[0].ty = OxTy::Double;
        func.temps[0] = OxTy::Variant;
        func.blocks[0].instrs = vec![OxInst::Arith {
            dst: OxPlace::Temp(TempId(0)),
            op: ArithOp::Add,
            lhs: OxOperand::Const(OxConst::F64(1.0f64.to_bits())),
            rhs: OxOperand::Const(OxConst::I32(2)),
            mode: NumericMode::Widening,
        }];
        assert!(ScalarLayout::new(&func).is_ok());
        if let OxInst::Arith { lhs, .. } = &mut func.blocks[0].instrs[0] {
            *lhs = OxOperand::Const(OxConst::I32(1));
        }
        assert!(ScalarLayout::new(&func).is_err());
        if let OxInst::Arith { lhs, .. } = &mut func.blocks[0].instrs[0] {
            *lhs = OxOperand::Const(OxConst::F64(1.0f64.to_bits()));
        }
        func.blocks[0].instrs.push(OxInst::Assign {
            dst: OxPlace::Temp(TempId(0)),
            value: OxOperand::Const(OxConst::I32(3)),
        });
        assert!(ScalarLayout::new(&func).is_err());

        // Object value comparisons/conditions require default-value dispatch,
        // never integer truth/comparison of the physical object pointer.
        func.locals.push(OxLocal {
            name: "ObjectValue".into(),
            ty: OxTy::Object(calcvba_oxir::ObjClass::Untyped),
            array_element: None,
            param: None,
            escaped: false,
        });
        func.temps = vec![OxTy::Bool];
        let object = OxOperand::Use(OxPlace::Local(LocalId(1)));
        func.blocks[0].instrs = vec![OxInst::Truthy {
            dst: OxPlace::Temp(TempId(0)),
            src: object.clone(),
        }];
        assert!(ScalarLayout::new(&func).is_err());
        func.blocks[0].instrs = vec![OxInst::Compare {
            coercion: calcvba_bundle::ComparisonCoercion::Strict,
            dst: OxPlace::Temp(TempId(0)),
            op: calcvba_oxir::CmpOp::Eq,
            lhs: object.clone(),
            rhs: object,
            mode: calcvba_bundle::StringCompareMode::Binary,
        }];
        assert!(ScalarLayout::new(&func).is_err());
    }
}
