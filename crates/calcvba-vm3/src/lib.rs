//! `calcvba-vm3` — the typed-CFG reference interpreter for OxIR.
//!
//! VM3 is the sole product interpreter and the JIT parity reference. It uses a
//! typed register/place model, heap frame stack, ByRef aliases, error/Resume
//! routing, arrays/records/objects/classes, project events and shared
//! runtime/evaluation/library/host substrates. The retired Op/Bundle VM2 path is
//! not an execution fallback.
//!
//! Historical note: `vm2` (the `calcvba-vm2` crate) was the predecessor Op/Bundle
//! interpreter and has been **deleted**. Comments below that say "mirrors vm2's
//! X", "matching vm2", or reference a `vm2` op/method name describe the frame,
//! call, lifecycle and error conventions vm3 inherited from that design — the
//! behavior they document is vm3's own; there is no live `vm2` to consult.
//!
//! VM3 is broad but remains in progress for the complete verified OxIR and
//! Windows interop profiles. Its golden snapshot is regression evidence; public
//! VBA specifications and reproducible Excel/VBA behavior remain semantic
//! authority. See `docs/ARCHITECTURE.md`.

use calcvba_runtime::ObjectRefExt;
use std::borrow::Cow;
use std::collections::HashMap;
use std::ptr::NonNull;
use std::rc::Rc;

mod bytecode;
use crate::bytecode::{
    CompiledFunc, FastArg, FastCallArg, FastConst, InstRef, LocalSlotInit, PackedOp, PackedPlace,
    TAG_CONST, TAG_GLOBAL, TAG_LOCAL, TAG_TEMP, VmOp,
};

use calcvba_bundle::{
    ArrayElementType, AssignmentIntent, AssignmentTargetKind, ComparisonCoercion, NativeImplId,
    NumericCoerceTarget, NumericMode, ProjectMemberKind, array_element_type_for_vartype,
    default_array_element, redim_safearray_from_elements, safearray_vartype_for_element,
    vba_record_layout_for_fields,
};
use calcvba_com::{DynamicMemberSelector, TypeLibMemberInvokeKind};
use calcvba_eval::arith::{self, ArithError};
use calcvba_eval::collection::{CollectionMethod, add_collection_item, dispatch_collection};
use calcvba_hal::HostServices;
use calcvba_hal::traits::DynLinkDescriptorView;
#[cfg(test)]
use calcvba_oxir::OxBlock;
use calcvba_oxir::inst::OxAsNew;
use calcvba_oxir::value::{
    ArithOp, BoundWhich, CmpOp, ErrField, LogicalOp, OxAddrBase, OxAddrStep, OxArg, OxCallArg,
    OxCoerceTarget, OxConst, OxNativeCallee, OxOperand, OxPlace, PtrKind,
};
use calcvba_oxir::{BlockId, FuncId, ImportId, LocalId, OxInst, OxProgram, OxTerminator, OxTy};
use calcvba_rt_abi::{
    ErrorMode, EventBinding, ExecState, Fault, FaultAction, LoadedProgram, MarshalArgRef,
    ObjectMemberRoute, ProcInvoker, ResumePoint, ResumeTarget, SavedErrState,
    VBA_COLLECTION_ROUTE_KEY, array_subscript, collection_member_name, marshal_ox_args,
    marshal_ox_call_args, object_member_route, object_receiver, ox_const_variant,
    project_has_default_get, read_default_value, resolve_project_default_member,
    runtime_class_descriptors_for_program, variant_changed,
};
use calcvba_runtime::collection::CollectionCursor;
use calcvba_runtime::object_ref::{
    ObjectRef, RUNTIME_CLASS_LIFECYCLE_NONE, RUNTIME_IUNKNOWN_INTERFACE_DESCRIPTOR,
    RuntimeClassDescriptor, borrow_project_field_get_into, borrow_project_field_set,
    borrow_project_field_set_trivial, borrow_project_instance_info, borrow_with_native_collection,
};
#[cfg(test)]
use calcvba_runtime::object_ref::{
    RuntimeClassActivationDescriptor, RuntimeInterfaceId, RuntimeInterfaceKind,
    RuntimeMemberBacking, RuntimeMemberInvokeKind, RuntimeProjectClassIdentity, RuntimeValueType,
};
use calcvba_runtime::safe_array::{SafeArray, SafeArrayBound};
use calcvba_runtime::variant::VarType;
use calcvba_runtime::{
    CallbackExecutor, CallbackRegistration, Variant, VbaRecord, register_callback,
};

/// `DISP_E_PARAMNOTFOUND` — the sentinel an omitted optional argument carries into a
/// callee slot, so `IsMissing`/`IsError` observe it exactly as vm2 does.
const MISSING_ARG: i32 = 0x8002_0004u32 as i32;

/// The runtime QI descriptor for a built-in `Collection` (name + IUnknown). Fully `'static`,
/// so `New Collection` needs no per-instance descriptor leak.
static VBA_COLLECTION_DESCRIPTOR: RuntimeClassDescriptor = RuntimeClassDescriptor {
    name: "Collection",
    runtime_type_name: None,
    project_identity: None,
    default_interface_iid: Some(calcvba_runtime::com::IID_VBA_COLLECTION),
    source_interface_iid: None,
    predeclared: false,
    lifecycle: RUNTIME_CLASS_LIFECYCLE_NONE,
    record_members: &[],
    project_records: &[],
    fields: &[],
    as_new_fields: &[],
    implements: &[],
    interfaces: &[RUNTIME_IUNKNOWN_INTERFACE_DESCRIPTOR],
    events: &[],
};

/// A failure to execute an OxIR program on vm3.
#[derive(Debug, Clone)]
pub enum Vm3Error {
    /// A host budget stopped execution; VBA On Error cannot catch it.
    BudgetExceeded,
    /// An uncaught VBA run-time fault propagated to the top level.
    Fault(Fault),
    /// An OxIR construct vm3 does not yet execute (honest and explicit — never a
    /// silent mis-execution). Lands in a later milestone (calls/builtins: M2-b;
    /// error/`Resume`: M2-c; objects/COM/arrays: M3).
    Unimplemented { what: &'static str },
    /// A structurally invalid program (should not occur for verifier-clean OxIR).
    Malformed(String),
    /// An error a WithEvents handler did not handle: VBA reports it as a top-level run-time
    /// error and ends the run; no `On Error` handler catches it (Excel, probe events8).
    EventAbort(Fault),
    /// `End` inside a nested activation: it unwinds every enclosing activation to the
    /// outermost `run_loop`, which ends the run as `End` there does. Never returned from the VM.
    Halted,
}

impl std::fmt::Display for Vm3Error {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Vm3Error::BudgetExceeded => write!(f, "execution budget exhausted"),
            Vm3Error::Fault(fault) => write!(f, "uncaught error {}: {}", fault.code, fault.message),
            Vm3Error::Unimplemented { what } => write!(f, "vm3 does not yet execute: {what}"),
            Vm3Error::Malformed(m) => write!(f, "malformed OxIR: {m}"),
            Vm3Error::EventAbort(fault) => write!(
                f,
                "uncaught error {} in an event handler: {}",
                fault.code, fault.message
            ),
            Vm3Error::Halted => write!(f, "the run ended with End"),
        }
    }
}

impl std::error::Error for Vm3Error {}

/// A resolved runtime storage location on the frame stack — what an [`OxPlace`]
/// denotes once ByRef aliasing is applied. `Local`/`Temp` name a specific frame by
/// index so a callee's ByRef parameter can point at one of its caller's cells (which
/// always outlives it, since callers sit below callees and pop later).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
enum Loc {
    /// `(program, slot)`: the program tags the slot so a ByRef alias or a call's result dst
    /// captured in one program stays bound to that program's globals regardless of the
    /// executing `cur` (mirrors vm2's `Place::Global(bundle, slot)`).
    Global(usize, usize),
    Local(usize, usize),
    Temp(usize, usize),
}

impl Loc {
    fn frame_index(self) -> Option<usize> {
        match self {
            Loc::Global(_, _) => None,
            Loc::Local(frame, _) | Loc::Temp(frame, _) => Some(frame),
        }
    }
}

fn is_cleared_temp(loc: Loc, frame_index: usize, first_temp: usize) -> bool {
    matches!(loc, Loc::Temp(frame, temp) if frame == frame_index && temp >= first_temp)
}

/// `For Each` retains an array storage location so element writes stay visible.
/// The iterator is keyed in [`Vm3::for_each`] by its resolved [`Loc`].
struct ForEachState {
    array_source: Option<Loc>,
    elements: Vec<Variant>,
    position: usize,
    /// Direct built-in `VBA.Collection` cursor (no trait-object boxing or event scope).
    collection_cursor: Option<CollectionCursor>,
    /// A live foreign COM enumerator, asked for one item at a time.
    enumerator: Option<Box<dyn calcvba_com::call::ComEnumerator>>,
}

/// A materialized sequence or live object enumerator for `For Each`.
enum ForEachItems {
    Snapshot(Vec<Variant>),
    CollectionCursor(CollectionCursor),
    Enumerator(Box<dyn calcvba_com::call::ComEnumerator>),
}

/// One procedure activation: its dispatch position, value slots, ByRef aliasing, and
/// the linkage back to its caller. The activation stack holds these so dispatch is an
/// explicit loop (no native recursion → deep VBA recursion is bounded by the frame
/// ceiling, error 28, not a host stack overflow — and matches vm2's iterative model).
struct Frame {
    /// The program (index into `Vm3::programs`) whose `funcs`/`globals` this frame executes
    /// against. `run_loop` sets `cur` from it each iteration, so a cross-program call/return
    /// re-targets globals and instruction fetch automatically. Single-project runs are all `0`.
    prog: usize,
    /// The function this frame is executing.
    func: FuncId,
    /// The current block, bytecode PC inside `CompiledFunc::ops`, and block-entry flag.
    block: BlockId,
    pc: usize,
    at_block_start: bool,
    /// Unified frame slot storage: `[0..num_locals]` are locals (parameters first, then
    /// declared locals) and `[num_locals..]` are temporaries. Sized once from the function's
    /// local + temp tables so a live cell never moves (`VarPtr` / `VT_BYREF` pointer stability).
    slots: Vec<Variant>,
    num_locals: usize,
    /// High-water mark of temp slots written since the last `StmtBoundary`, as a flat index
    /// in `num_locals..=slots.len()`, so statement temp clearing only scans dirty temps.
    temp_hwm: usize,
    /// Strings `StrPtr` converted from non-String operands, kept for the statement (VBA's
    /// conversion temporaries).
    pointer_strings: Vec<Variant>,
    /// The array locks each `MakeRef` temporary holds while its call runs, by temp index;
    /// released by `ReleaseRef`, by clearing the temp at a statement boundary, and on exit.
    ref_locks: Vec<(usize, calcvba_runtime::safe_array::ArrayLock)>,
    /// The array locks `For Each` loops hold, by iterator temp index; released by
    /// `ForEachEnd` (`Exit For` or running out) and on exit, never at a statement boundary:
    /// a `GoTo` out of the loop leaves the array locked until the procedure returns.
    loop_locks: Vec<(usize, calcvba_runtime::safe_array::ArrayLock)>,
    /// ByRef parameters: a parameter's frame-local index → the caller location it
    /// aliases, resolved to its ultimate backing at call time (so aliases never chain).
    /// Stored as a tiny `Vec` (zero allocation when empty, linear scan outperforms hashing
    /// for 1-4 ByRef params).
    aliases: Vec<(usize, Loc)>,
    /// Bitmask of aliased local indices `0..64` (`u64::MAX` if any local `>= 64` is aliased)
    /// so non-aliased locals skip scanning `aliases` even when another parameter is `ByRef`.
    local_alias_mask: u64,
    /// Where this call's return value is written (resolved in the caller at call time);
    /// `None` for a statement call or the entry/initializer frame.
    dst: Option<Loc>,
    /// The local holding this function's result (`None` for a `Sub`).
    return_local: Option<LocalId>,
    /// Whether an unhandled fault unwinding this frame should drain terminations (true for
    /// project methods pushed directly by `CallComEarly`/`CallComLate`, matching
    /// `run_proc_with_me`'s epilogue drain timing).
    drain_on_fault: bool,
    /// The most recent numeric line label executed in this procedure activation.
    /// This is copied to the VM's public `Erl` value only when an error is caught.
    current_line: i32,
    /// The caller's error mode, restored when this frame returns (each callee starts
    /// with no handler).
    saved_error_mode: ErrorMode,
    /// The caller's active-error latch, restored on return (each callee starts with no
    /// active error). Keeping it per-activation is what makes a propagated `Resume`
    /// re-run the *caller's* call-site statement.
    saved_active_error: Option<ResumePoint>,
    /// The `GoSub` Resumption List — a per-activation LIFO stack of return blocks
    /// (MS-VBAL §5.4.2.14). `GoSub` pushes its `ret`; `Return` pops the most recent.
    gosub_stack: Vec<BlockId>,
}

impl Frame {
    #[inline]
    fn local(&self, i: usize) -> Option<&Variant> {
        (i < self.num_locals).then(|| &self.slots[i])
    }

    #[inline]
    fn local_mut(&mut self, i: usize) -> Option<&mut Variant> {
        (i < self.num_locals).then(|| &mut self.slots[i])
    }

    #[inline]
    fn temp(&self, i: usize) -> Option<&Variant> {
        self.slots.get(self.num_locals + i)
    }

    #[inline]
    fn temp_mut(&mut self, i: usize) -> Option<&mut Variant> {
        let idx = self.num_locals + i;
        if idx < self.slots.len() {
            if idx + 1 > self.temp_hwm {
                self.temp_hwm = idx + 1;
            }
            Some(&mut self.slots[idx])
        } else {
            None
        }
    }

    #[inline]
    fn alias_of(&self, local: usize) -> Option<Loc> {
        let mask = self.local_alias_mask;
        if mask == 0 {
            return None;
        }
        if local < 64 && (mask & (1u64 << local)) == 0 {
            return None;
        }
        for &(li, loc) in &self.aliases {
            if li == local {
                return Some(loc);
            }
        }
        None
    }

    #[inline]
    fn set_alias(&mut self, local: usize, loc: Loc) {
        if local < 64 {
            self.local_alias_mask |= 1u64 << local;
        } else {
            self.local_alias_mask = u64::MAX;
        }
        for entry in &mut self.aliases {
            if entry.0 == local {
                entry.1 = loc;
                return;
            }
        }
        self.aliases.push((local, loc));
    }
}

/// Zero-allocation stack buffer for small `Variant` argument lists without pre-constructing
/// unused `Variant::empty()` entries.
struct StackVariantBuf<const N: usize> {
    slots: [std::mem::MaybeUninit<Variant>; N],
    len: usize,
}

impl<const N: usize> StackVariantBuf<N> {
    #[inline]
    fn new() -> Self {
        Self {
            slots: [const { std::mem::MaybeUninit::uninit() }; N],
            len: 0,
        }
    }

    #[inline]
    fn push(&mut self, value: Variant) {
        debug_assert!(self.len < N);
        self.slots[self.len].write(value);
        self.len += 1;
    }

    #[inline]
    fn as_mut_slice(&mut self) -> &mut [Variant] {
        // SAFETY: `0..self.len` were initialized via `push` and have not been dropped.
        unsafe {
            std::slice::from_raw_parts_mut(self.slots.as_mut_ptr().cast::<Variant>(), self.len)
        }
    }
}

impl<const N: usize> Drop for StackVariantBuf<N> {
    #[inline]
    fn drop(&mut self) {
        for i in 0..self.len {
            // SAFETY: `0..self.len` are initialized and dropped once here.
            unsafe { self.slots[i].assume_init_drop() };
        }
    }
}

type ReentrantProjectEventSink<'h> =
    Rc<dyn Fn(Vm3<'h>, ObjectRef, i32, Vec<Variant>) -> (Vm3<'h>, Result<(), String>) + 'h>;

/// Out-of-line per-VM bytecode cache, typed-entry table, and frame pool
/// (kept boxed so `Vm3` stays compact on the stack and in host enums).
struct Vm3Cache<'h> {
    typed_entries: Vec<Vec<Option<typed::TypedInterpreter<'h>>>>,
    /// Lazily compiled compact bytecode per `(program, func)` pair.
    compiled: Vec<Vec<Option<Rc<CompiledFunc>>>>,
    /// Reusable pooled frames whose `slots` stay filled with `Empty` `Variant`s across calls.
    frame_pool: Vec<Frame>,
}

/// The vm3 interpreter over one or more typed OxIR programs.
pub struct Vm3<'h> {
    reentrant_project_event_sink: Option<ReentrantProjectEventSink<'h>>,
    has_typed_entries: bool,
    cache: Box<Vm3Cache<'h>>,
    /// Index into `programs` of the executing program (`0` for single-project activation).
    cur: usize,
    /// Shared runtime state below vm3 and the JIT.
    exec: ExecState<'h>,
    /// The activation stack. `frames[0]` is the entry (`Main`) frame and is never
    /// popped — it backs the result snapshot; deeper frames are `CallProc` callees.
    frames: Vec<Frame>,
    /// `For Each` iterator state, keyed by the loop variable's resolved [`Loc`] (so
    /// reentrant/nested loops that reuse a slot number never alias) — mirrors vm2's
    /// `for_each` map.
    for_each: HashMap<Loc, ForEachState>,
    /// ParamArray packs whose elements alias caller slots. Keyed by the resolved
    /// location that currently stores the ParamArray SAFEARRAY.
    param_array_aliases: HashMap<Loc, Vec<Option<Loc>>>,
    /// `As New` object slots, keyed by the resolved storage location. A read of
    /// one of these slots creates a fresh instance when the slot is Empty/Nothing.
    as_new_slots: HashMap<Loc, OxAsNew>,
    /// Nested `run_loop` activations (host entries, object dispatch, event handlers, native
    /// callbacks): `End` below the outermost unwinds with [`Vm3Error::Halted`].
    run_depth: usize,
}

impl<'h> Vm3<'h> {
    fn idle(host: &'h dyn HostServices) -> Self {
        Self {
            reentrant_project_event_sink: None,
            has_typed_entries: false,
            cache: Box::new(Vm3Cache {
                typed_entries: Vec::new(),
                compiled: Vec::new(),
                frame_pool: Vec::new(),
            }),
            cur: 0,
            exec: ExecState::new(host),
            frames: Vec::new(),
            for_each: HashMap::new(),
            param_array_aliases: HashMap::new(),
            as_new_slots: HashMap::new(),
            run_depth: 0,
        }
    }

    /// The currently-executing program's static IR (`programs[cur]`). A `&'h` copy, so callers
    /// can hold it across `&mut self` mutations exactly as the old `self.program` field allowed.
    #[inline]
    fn cur_program(&self) -> &'h OxProgram {
        self.exec.programs[self.cur].program
    }

    fn project_typed_receiver(&self, operand: &OxOperand) -> bool {
        let Some(frame) = self.frames.last() else {
            return false;
        };
        let function = &self.cur_program().funcs[frame.func.0];
        let ty = match operand {
            OxOperand::Use(OxPlace::Local(id)) => function.locals.get(id.0).map(|local| &local.ty),
            OxOperand::Use(OxPlace::Temp(id)) => function.temps.get(id.0),
            OxOperand::Use(OxPlace::Global(id)) => self
                .cur_program()
                .globals
                .get(id.0)
                .map(|global| &global.ty),
            _ => None,
        };
        matches!(
            ty,
            Some(OxTy::Object(
                calcvba_oxir::ObjClass::Class(_) | calcvba_oxir::ObjClass::Iface(_)
            ))
        )
    }

    fn operand_is_variant(&self, operand: &OxOperand) -> bool {
        let Some(frame) = self.frames.last() else {
            return false;
        };
        let function = &self.cur_program().funcs[frame.func.0];
        let ty = match operand {
            OxOperand::Use(OxPlace::Local(id)) => function.locals.get(id.0).map(|local| &local.ty),
            OxOperand::Use(OxPlace::Temp(id)) => function.temps.get(id.0),
            OxOperand::Use(OxPlace::Global(id)) => self
                .cur_program()
                .globals
                .get(id.0)
                .map(|global| &global.ty),
            _ => None,
        };
        ty == Some(&OxTy::Variant)
    }

    /// Whether `operand` is a place declared `As Object` (an untyped object variable).
    fn operand_is_untyped_object(&self, operand: &OxOperand) -> bool {
        let Some(frame) = self.frames.last() else {
            return false;
        };
        let function = &self.cur_program().funcs[frame.func.0];
        let ty = match operand {
            OxOperand::Use(OxPlace::Local(id)) => function.locals.get(id.0).map(|local| &local.ty),
            OxOperand::Use(OxPlace::Temp(id)) => function.temps.get(id.0),
            OxOperand::Use(OxPlace::Global(id)) => self
                .cur_program()
                .globals
                .get(id.0)
                .map(|global| &global.ty),
            _ => None,
        };
        matches!(ty, Some(OxTy::Object(calcvba_oxir::ObjClass::Untyped)))
    }

    /// Initial slot value for a statically typed VBA variable. Dynamic array variables are
    /// arrays before allocation for `IsArray`, but have no SAFEARRAY descriptor/bounds until
    /// `ReDim`.
    fn initial_value_for_slot(
        program: &OxProgram,
        ty: &OxTy,
        array_element: Option<&ArrayElementType>,
    ) -> Result<Variant, Vm3Error> {
        match ty {
            OxTy::Array(element, _) => {
                if let OxTy::Record(id) = element.as_ref()
                    && let Ok(layout) =
                        calcvba_rt_abi::vba_record_layout_for_program_record(program, *id)
                {
                    return Ok(Variant::unallocated_record_array(layout));
                }
                let element = (match element.as_ref() {
                    OxTy::Record(id) => program
                        .record_layout(*id)
                        .map(|fields| ArrayElementType::Record(fields.to_vec())),
                    _ => None,
                })
                .or_else(|| array_element.cloned())
                .or_else(|| Self::array_element_type_for_ox_ty(element))
                .unwrap_or(ArrayElementType::Variant);
                calcvba_bundle::array_runtime::unallocated_typed_array(&element)
                    .map_err(|e| Vm3Error::Fault(Fault::new(7, e)))
            }
            OxTy::Object(_) => Ok(Variant::nothing()),
            _ => Ok(Variant::empty()),
        }
    }

    fn array_element_type_for_ox_ty(ty: &OxTy) -> Option<ArrayElementType> {
        let element = match ty {
            OxTy::Bool => ArrayElementType::Boolean,
            OxTy::Byte => ArrayElementType::Byte,
            OxTy::Integer => ArrayElementType::Integer,
            OxTy::Long => ArrayElementType::Long,
            OxTy::LongLong => ArrayElementType::LongLong,
            OxTy::Single => ArrayElementType::Single,
            OxTy::Double => ArrayElementType::Double,
            OxTy::Currency => ArrayElementType::Currency,
            OxTy::Date => ArrayElementType::Date,
            OxTy::Str => ArrayElementType::String,
            OxTy::FixedStr(len) => ArrayElementType::FixedString(*len as usize),
            OxTy::Variant => ArrayElementType::Variant,
            _ => ArrayElementType::Variant,
        };
        Some(element)
    }

    /// Build a program's mutable runtime tables: one shared `&'static` runtime descriptor per
    /// class from `calcvba_rt_abi::runtime_class_descriptors_for_program`, the
    /// `(binding,event)→handler` route table, and the zeroed module-global slots.
    fn build_loaded(program: &'h OxProgram) -> Result<LoadedProgram<'h>, Vm3Error> {
        let class_descriptors = runtime_class_descriptors_for_program(program);
        let mut event_routes: HashMap<(i32, i32), usize> = HashMap::new();
        for route in &program.event_routes {
            event_routes.insert((route.binding, route.event), route.handler);
        }
        Ok(LoadedProgram {
            program,
            globals: program
                .globals
                .iter()
                .map(|global| {
                    Self::initial_value_for_slot(program, &global.ty, global.array_element.as_ref())
                })
                .collect::<Result<_, _>>()?,
            class_descriptors,
            predeclared_singletons: HashMap::new(),
            event_routes,
        })
    }

    /// The index of the loaded program declaring unit `name` (its `unit_name`), for resolving a
    /// cross-project import/reference.
    fn program_index_by_unit(&self, unit: &str) -> Option<usize> {
        self.exec
            .programs
            .iter()
            .position(|lp| lp.program.unit_name.eq_ignore_ascii_case(unit))
    }

    /// Every cross-project import (a non-`VBA` `unit`) must name a loaded program, else the link
    /// is unresolved — reported naming the missing unit (never a silent mis-link). `VBA` imports
    /// are the synthetic built-in library, resolved separately in `call_extern`.
    fn validate_links(&self) -> Result<(), Vm3Error> {
        for lp in &self.exec.programs {
            for imp in &lp.program.imports {
                if imp.unit.eq_ignore_ascii_case("VBA") {
                    continue;
                }
                if self.program_index_by_unit(&imp.unit).is_none() {
                    return Err(Vm3Error::Malformed(format!(
                        "unresolved reference to unit '{}'",
                        imp.unit
                    )));
                }
            }
        }
        Ok(())
    }

    /// Run `program` to completion and return the finished VM (read the result snapshot
    /// with [`Vm3::slot`]). Mirrors vm2: the global initializer runs first, then `Main`
    /// in an entry frame that is never popped.
    pub fn run(program: &'h OxProgram, host: &'h dyn HostServices) -> Result<Self, Vm3Error> {
        let mut vm = Self::activate(program, host)?;
        vm.run_entry()?;
        Ok(vm)
    }

    /// Link `programs` into a runnable image and [`activate`](Self::activate) it (without running
    /// the entry — the host drives that). A single program is the common case: a whole VBA
    /// *project* is exactly ONE `OxProgram` with many classes, so it links trivially and this is
    /// what the single-project product paths (CLI run, single-class COM server) need.
    ///
    /// Multiple programs are cross-project references (project A referencing project B's
    /// procs/classes). The entry is the LAST program (mirrors vm2's `Vm::link`); cross-project
    /// imports are resolved by `unit_name` and must all be present, else the link fails naming
    /// the missing unit. Each program's module-global initializer runs in its own program
    /// context; the entry's `Main` is run later by `run_entry`.
    pub fn link(programs: &[&'h OxProgram], host: &'h dyn HostServices) -> Result<Self, Vm3Error> {
        Self::link_with_budget(programs, host, None)
    }

    /// Link with a budget covering initialization and all nested calls.
    pub fn link_with_budget(
        programs: &[&'h OxProgram],
        host: &'h dyn HostServices,
        steps: Option<u64>,
    ) -> Result<Self, Vm3Error> {
        if programs.is_empty() {
            return Err(Vm3Error::Malformed(
                "Vm3::link requires at least one program".into(),
            ));
        }
        let entry = programs.len() - 1;
        let allow_typed = programs
            .iter()
            .all(|p| p.classes.is_empty() && p.record_layouts.is_empty());
        let typed_entries: Vec<Vec<_>> = programs
            .iter()
            .map(|p| {
                p.funcs
                    .iter()
                    .map(|f| {
                        allow_typed
                            .then(|| typed::TypedInterpreter::bind(f).ok())
                            .flatten()
                    })
                    .collect()
            })
            .collect();
        let mut exec = ExecState::new(host);
        exec.remaining_steps = steps;
        exec.programs = programs
            .iter()
            .map(|p| Self::build_loaded(p))
            .collect::<Result<_, _>>()?;
        let compiled = programs.iter().map(|p| vec![None; p.funcs.len()]).collect();
        let has_typed_entries = typed_entries
            .iter()
            .any(|prog| prog.iter().any(Option::is_some));
        let mut vm = Vm3 {
            reentrant_project_event_sink: None,
            has_typed_entries,
            cache: Box::new(Vm3Cache {
                typed_entries,
                compiled,
                frame_pool: Vec::new(),
            }),
            cur: entry,
            exec,
            frames: Vec::new(),
            for_each: HashMap::new(),
            param_array_aliases: HashMap::new(),
            as_new_slots: HashMap::new(),
            run_depth: 0,
        };
        // Every cross-project reference must resolve before any code runs.
        vm.validate_links()?;
        // Run each program's module-global initializer in ITS OWN program context (cur = i), so
        // a global write lands in that program's globals (the entry's initializer runs last).
        for i in 0..vm.exec.programs.len() {
            if let Some(init) = vm.exec.programs[i].program.global_initializer {
                vm.cur = i;
                let frame = vm.new_frame(init)?;
                vm.frames.push(frame);
                let r = vm.run_loop(0);
                // The initializer writes module globals; its own frame is discarded.
                if let Some(frame) = vm.frames.pop() {
                    vm.recycle_frame(frame);
                }
                vm.prune_param_array_aliases_from_depth(vm.frames.len());
                r?;
            }
        }
        vm.cur = entry;
        // Isolate this activation from any prior run on the shared thread-local termination
        // queue, before any entry/invoke runs — matching vm2's per-run reset.
        calcvba_runtime::reset_pending_terminations();
        Ok(vm)
    }

    /// Host session factory (W7): mint a project-class instance by NAME (running its
    /// `Class_Initialize`) and return it as a held object. The COM-server (`calcvba-comhost`)
    /// activation path calls this. A class the program doesn't declare is "can't create object"
    /// (429). The instance's state lives on the object box exactly as `New`-minted ones.
    pub fn create_project_instance(&mut self, class_name: &str) -> Result<Variant, Vm3Error> {
        let idx = self
            .cur_program()
            .classes
            .iter()
            .position(|c| c.name.eq_ignore_ascii_case(class_name))
            .ok_or_else(|| {
                Vm3Error::Fault(Fault::new(429, "ActiveX component can't create object"))
            })?;
        self.new_project_instance(idx)
    }

    /// Host session invoke (W7): call a member by NAME on a held instance with pre-marshaled
    /// by-value args (the COM-server dispatch path). Resolves the member proc by name + accessor
    /// kind (with vm2's get↔method fallback) and runs it via [`Self::run_proc_with_values`]; a
    /// built-in `Collection` receiver routes to the shared keyed dispatch. Missing member → 438.
    pub fn invoke_member_values(
        &mut self,
        object: ObjectRef,
        name: &str,
        kind_hint: Option<ProjectMemberKind>,
        args: Vec<Variant>,
    ) -> Result<Variant, Vm3Error> {
        self.invoke_member_values_captured(object, name, kind_hint, args, None)
    }

    /// Release one host owner using the same WithEvents cleanup and finalizer
    /// boundaries as a VBA slot overwrite. Other held owners keep it alive.
    pub fn release_project_instance(&mut self, object: ObjectRef) {
        let value = Variant::from_object_ref(object);
        self.clear_withevents_owners_before_releasing_values(std::iter::once(&value));
        drop(value);
        self.maybe_drain();
    }
    /// Run queued finalizers after host storage released its owning values.
    pub fn drain_object_terminations(&mut self) {
        self.maybe_drain();
    }
    pub fn default_instance_context(&mut self) -> calcvba_rt_abi::DefaultInstanceContext {
        self.exec.default_instance_context()
    }

    /// Invoke pre-marshaled host arguments and retain their final values,
    /// including changes made before a fault. Equal alias indices share a ByRef cell.
    pub fn invoke_member_values_with_writebacks(
        &mut self,
        object: ObjectRef,
        name: &str,
        kind_hint: Option<ProjectMemberKind>,
        args: &mut [Variant],
        aliases: &[usize],
    ) -> Result<Variant, Vm3Error> {
        self.invoke_member_values_captured(
            object,
            name,
            kind_hint,
            args.to_vec(),
            Some((args, aliases)),
        )
    }

    /// Invoke a member already resolved to its [`calcvba_runtime::RuntimeMemberDescriptor`]
    /// (by DISPID or vtable slot, e.g. from `IDispatch::Invoke`), preserving ByRef aliases and
    /// write-backs on success or fault.
    pub fn invoke_member_descriptor_captured(
        &mut self,
        object: ObjectRef,
        member: &'static calcvba_runtime::RuntimeMemberDescriptor,
        invoke_kind: calcvba_runtime::RuntimeMemberInvokeKind,
        args: &mut [Variant],
        aliases: &[usize],
    ) -> Result<Variant, Vm3Error> {
        let obj_bundle = object.bundle_id() as usize;
        match member.backing {
            calcvba_runtime::RuntimeMemberBacking::Proc(proc) => {
                let values = args.to_vec();
                self.run_proc_with_values_captured(
                    obj_bundle,
                    FuncId(proc),
                    Variant::from_object_ref(object),
                    values,
                    false,
                    Some((args, aliases)),
                )
            }
            calcvba_runtime::RuntimeMemberBacking::Field(field_token) => self
                .dispatch_project_field_member_values(
                    &object,
                    member,
                    field_token,
                    invoke_kind,
                    args.to_vec(),
                ),
            calcvba_runtime::RuntimeMemberBacking::None => {
                Err(Vm3Error::Fault(Fault::new(438, default_error_message(438))))
            }
        }
    }

    fn invoke_member_values_captured(
        &mut self,
        object: ObjectRef,
        name: &str,
        kind_hint: Option<ProjectMemberKind>,
        args: Vec<Variant>,
        capture: Option<(&mut [Variant], &[usize])>,
    ) -> Result<Variant, Vm3Error> {
        if calcvba_lib::objects::is_object(&object) {
            return calcvba_rt_abi::invoke_library_object(
                &mut self.exec,
                &object,
                name,
                kind_hint.unwrap_or(ProjectMemberKind::Method),
                &args,
                &[],
            )
            .map_err(Vm3Error::Fault);
        }
        if object.route_key() == VBA_COLLECTION_ROUTE_KEY {
            return self.dispatch_collection_values(&object, name, args);
        }
        // Resolve via the object's shared class descriptor + dispatch plan cache (dc-yqq0.22),
        // running in the object's OWN program (bundle_id), not the executing cur.
        let obj_bundle = object.bundle_id() as usize;
        let requested_kind = kind_hint.map(calcvba_rt_abi::runtime_invoke_kind);
        let member = object
            .class_descriptor()
            .dispatch_interface()
            .and_then(|dispatch| {
                let plan = self
                    .exec
                    .dispatch_plan_cache
                    .resolve_project_member_for_host(dispatch, name, requested_kind)?;
                dispatch.member_at(plan.member_index)
            })
            .ok_or_else(|| Vm3Error::Fault(Fault::new(438, default_error_message(438))))?;
        match member.backing {
            calcvba_runtime::RuntimeMemberBacking::Proc(proc) => {
                // Host args are already by-value (resolution-order is moot), but run in the
                // object's program (target_prog = obj_bundle).
                self.run_proc_with_values_captured(
                    obj_bundle,
                    FuncId(proc),
                    Variant::from_object_ref(object),
                    args,
                    false,
                    capture,
                )
            }
            calcvba_runtime::RuntimeMemberBacking::Field(field_token) => {
                let effective_kind = requested_kind.unwrap_or(member.invoke_kind);
                self.dispatch_project_field_member_values(
                    &object,
                    member,
                    field_token,
                    effective_kind,
                    args,
                )
            }
            calcvba_runtime::RuntimeMemberBacking::None => {
                Err(Vm3Error::Fault(Fault::new(438, default_error_message(438))))
            }
        }
    }

    /// A built-in `Collection` member call with already-marshaled by-value args (the
    /// [`Self::invoke_member_values`] counterpart of [`Self::dispatch_collection_method`], whose
    /// args arrive as `OxCallArg`s).
    fn dispatch_collection_values(
        &mut self,
        object: &ObjectRef,
        name: &str,
        args: Vec<Variant>,
    ) -> Result<Variant, Vm3Error> {
        let native = Self::vba_collection_native_method(name).ok_or_else(|| {
            Vm3Error::Fault(Fault::new(
                438,
                format!("Collection doesn't support '{name}'"),
            ))
        })?;
        let method = match native {
            calcvba_bundle::NativeMethodId::CollectionAdd => CollectionMethod::Add,
            calcvba_bundle::NativeMethodId::CollectionItem => CollectionMethod::Item,
            calcvba_bundle::NativeMethodId::CollectionCount => CollectionMethod::Count,
            calcvba_bundle::NativeMethodId::CollectionRemove => CollectionMethod::Remove,
            calcvba_bundle::NativeMethodId::CollectionNewEnum => {
                return self.collection_new_enum_object(object);
            }
        };
        object
            .with_native_collection(|d| dispatch_collection(method, d, &args))
            .ok_or_else(|| Vm3Error::Fault(Fault::new(424, "Object required")))?
            .map_err(|error| Vm3Error::Fault(Fault::standard(error.vba_number())))
    }

    /// Register a host event sink (W7): every project `RaiseEvent` is delivered to `sink` as
    /// `(source, event_id, args)` after the internal `WithEvents` fan-out. The COM server uses
    /// this to forward project events to its connection-point clients.
    pub fn set_project_event_sink<F>(&mut self, sink: F)
    where
        F: FnMut(ObjectRef, i32, Vec<Variant>) -> Result<(), String> + 'h,
    {
        self.reentrant_project_event_sink = None;
        self.exec.events.project_event_sink = Some(Box::new(sink));
    }

    /// Remove the host event sink (W7).
    pub fn clear_project_event_sink(&mut self) {
        self.exec.events.project_event_sink = None;
        self.reentrant_project_event_sink = None;
    }

    /// Transfer the suspended VM to a synchronous host event callback and recover
    /// it on return. The callback may invoke members, including nested events,
    /// while exclusively owning the same frames, globals and object state.
    pub fn set_reentrant_project_event_sink<F>(&mut self, sink: F)
    where
        F: Fn(Self, ObjectRef, i32, Vec<Variant>) -> (Self, Result<(), String>) + 'h,
    {
        self.exec.events.project_event_sink = None;
        self.reentrant_project_event_sink = Some(Rc::new(sink));
    }

    /// Build the VM for a SINGLE program and run its module-global initializer, but do NOT run
    /// the entry (`Main`) — the front half of [`Vm3::run`], split out so a long-lived host
    /// session can activate once and issue many member invokes. A whole VBA project is exactly
    /// one `OxProgram`; this is [`Vm3::link`] over `[program]` (which leaks the `&'static` class
    /// descriptors, builds the event routes, runs the initializer, and resets the drain queue
    /// exactly once).
    pub fn activate(program: &'h OxProgram, host: &'h dyn HostServices) -> Result<Self, Vm3Error> {
        Self::link(&[program], host)
    }

    /// Run the program entry (`Sub Main`) in an entry frame that is never popped — it stays as
    /// `frames[0]` for the result snapshot — then drain any parked `Class_Terminate`s. A no-op
    /// when the program has no entry. The back half of [`Vm3::run`], split out so a session can
    /// `activate()` without running `Main`.
    pub fn run_entry(&mut self) -> Result<(), Vm3Error> {
        if let Some(entry) = self.cur_program().entry {
            let frame = self.new_frame(entry)?;
            self.frames.push(frame);
            // The entry frame is never popped — it stays as `frames[0]` for the snapshot.
            let r = self.run_loop(0);
            // Run any `Class_Terminate`s parked while the run unwound — including objects an
            // uncaught fault released as it propagated out of called procs (vm2 drains on the
            // fault path; without this a Terminate would be lost on an error exit). On a clean
            // finish this is a no-op (statement boundaries already drained; the entry frame's
            // own locals stay live, un-terminated, exactly as vm2 leaves them).
            self.maybe_drain();
            r?;
        }
        Ok(())
    }

    /// Start a new host-call budget while retaining the linked project's state.
    pub fn set_execution_budget(&mut self, steps: u64) {
        self.exec.remaining_steps = Some(steps);
    }

    /// Seat an unhandled typed host-entry fault in this session's shared Err.
    pub fn record_typed_entry_error(&mut self, number: i32) {
        self.exec.err_engine.raise(
            Fault::new(number, calcvba_rt_abi::default_error_message(number)),
            self.exec.default_error_source.clone(),
        );
    }
    /// Canonical storage for a retained typed activation. The caller must keep
    /// the session leased and exclude reset/replacement while retaining pointers.
    pub fn typed_globals_raw(&mut self) -> (*mut Variant, usize) {
        self.typed_globals_raw_at(self.exec.programs.len() - 1)
    }
    /// Canonical storage for one retained linked unit under the same lease.
    pub fn typed_globals_raw_at(&mut self, unit: usize) -> (*mut Variant, usize) {
        let globals = &mut self
            .exec
            .programs
            .get_mut(unit)
            .expect("linked session unit")
            .globals;
        // Vec's raw accessor does not materialize an exclusive slice that would
        // invalidate a parked activation's existing raw storage authority.
        (globals.as_mut_ptr(), globals.len())
    }
    /// Stable same-thread Err storage; retaining callers must lease this session
    /// and exclude reset and unrelated execution, as for typed_globals_raw.
    pub fn typed_error_raw(&mut self) -> *mut calcvba_rt_abi::ErrEngine {
        &raw mut *self.exec.err_engine
    }
    /// Stable library state under the typed activation's same-thread session lease.
    pub fn typed_library_raw(&mut self) -> *mut calcvba_lib::LibContext {
        &raw mut *self.exec.lib
    }
    pub fn typed_class_context(&mut self) -> Option<calcvba_rt_abi::TypedClassContext> {
        self.exec.typed_class_context()
    }
    pub fn typed_event_state(
        &self,
    ) -> std::rc::Rc<std::cell::RefCell<calcvba_rt_abi::automatic_events::AutomaticEventState>>
    {
        self.exec.events.automatic.clone()
    }

    /// Invoke a standard-module function with already coerced value arguments.
    /// No hidden Me slot is inserted and host ByRef arguments are private copies.
    pub fn invoke_function_values(
        &mut self,
        func: FuncId,
        args: Vec<Variant>,
    ) -> Result<Variant, Vm3Error> {
        let target = self
            .cur_program()
            .funcs
            .get(func.0)
            .ok_or_else(|| Vm3Error::Malformed("function index out of range".into()))?;
        if target.param_count != args.len() {
            return Err(Vm3Error::Malformed(
                "function invocation signature mismatch".into(),
            ));
        }
        if let Some(result) = self.invoke_scalar_values(self.cur, func, &args) {
            return result;
        }
        let base = self.frames.len();
        let mut frame = self.new_frame(func)?;
        for (i, value) in args.into_iter().enumerate() {
            if let Some(slot) = frame.local_mut(i) {
                *slot = value;
            }
        }
        let saved = self.exec.err_engine.enter_activation();
        frame.saved_error_mode = saved.error_mode;
        frame.saved_active_error = saved.active_error;
        self.frames.push(frame);
        let outcome = self.run_loop(base);
        let value = self
            .frames
            .get(base)
            .and_then(|frame| frame.return_local.and_then(|id| frame.local(id.0).cloned()))
            .unwrap_or_else(Variant::empty);
        self.exec.err_engine.restore(saved);
        self.truncate_frames_with_withevents_cleanup(base);
        self.prune_param_array_aliases_from_depth(base);
        self.maybe_drain();
        outcome.map(|()| value)
    }

    /// The universal value route adapts exact scalar values to the same unboxed
    /// primary interpreter used by typed hosts. No source/name reconstruction.
    fn invoke_scalar_values(
        &mut self,
        program: usize,
        func: FuncId,
        args: &[Variant],
    ) -> Option<Result<Variant, Vm3Error>> {
        use calcvba_rt_abi::host_entry::{MAX_TYPED_ARGS, ScalarBits, ScalarType, TypedContext};
        let entry = self
            .cache
            .typed_entries
            .get(program)?
            .get(func.0)?
            .as_ref()?;
        if args.len() != entry.layout().parameters.len() {
            return None;
        }
        let mut bits = [ScalarBits::ZERO; MAX_TYPED_ARGS];
        for (i, (value, ty)) in args.iter().zip(&entry.layout().parameters).enumerate() {
            bits[i] = match ty {
                ScalarType::Long => ScalarBits::long(value.as_i32()?),
                ScalarType::Double => ScalarBits::double(value.as_f64()?),
                _ => return None,
            };
        }
        let result_ty = entry.layout().result;
        let source = &self.exec.programs[program].program.unit_name;
        let mut context = TypedContext {
            remaining_blocks: self.exec.remaining_steps.unwrap_or(u64::MAX),
        };
        let saved = self.exec.err_engine.enter_activation();
        let result = entry.invoke(&bits[..args.len()], &mut context);
        self.exec.err_engine.restore(saved);
        if self.exec.remaining_steps.is_some() {
            self.exec.remaining_steps = Some(context.remaining_blocks);
        }
        Some(match result {
            Ok(value) => Ok(match result_ty {
                ScalarType::Long => Variant::from_i32(value.as_long()),
                ScalarType::Double => Variant::from_f64(value.as_double()),
                _ => {
                    unreachable!("admitted scalar result")
                }
            }),
            Err(-1) => Err(Vm3Error::BudgetExceeded),
            Err(number) => {
                let mut fault = Fault::new(number, calcvba_rt_abi::default_error_message(number));
                fault.source = Some(source.as_str().into());
                self.exec.err_engine.raise(fault.clone(), source.clone());
                Err(Vm3Error::Fault(fault))
            }
        })
    }

    /// The result snapshot slot `i`: module globals occupy `[0, globals.len())`; higher
    /// indices are the entry (`Main`) frame's locals (the same layout vm2 exposes).
    pub fn slot(&self, i: usize) -> Option<Variant> {
        let global_count = self.exec.programs[self.cur].globals.len();
        if i < global_count {
            self.exec.programs[self.cur].globals.get(i).cloned()
        } else {
            let rel = i - global_count;
            self.frames.first()?.local(rel).cloned()
        }
    }

    /// The final `Err` state (number / description / source) for the error axis.
    pub fn err_number(&self) -> i32 {
        self.exec.err_engine.err.number
    }
    pub fn err_description(&self) -> &calcvba_runtime::Utf16String {
        &self.exec.err_engine.err.description
    }
    pub fn err_source(&self) -> &calcvba_runtime::Utf16String {
        &self.exec.err_engine.err.source
    }
    pub fn erl(&self) -> i32 {
        self.exec.err_engine.erl_line
    }
    /// `Err.LastDllError` — the OS last-error captured after the most recent `Declare Lib`
    /// call (M3-7); `0` until a Declare runs.
    pub fn last_dll_error(&self) -> i32 {
        self.exec.err_engine.last_dll_error
    }

    #[inline]
    fn get_or_compile(&mut self, prog: usize, func: FuncId) -> Rc<CompiledFunc> {
        if let Some(code) = &self.cache.compiled[prog][func.0] {
            return Rc::clone(code);
        }
        self.compile_func_cold(prog, func)
    }

    #[inline]
    fn ensure_compiled(&mut self, prog: usize, func: FuncId) {
        if self.cache.compiled[prog][func.0].is_none() {
            let _ = self.compile_func_cold(prog, func);
        }
    }

    #[cold]
    #[inline(never)]
    fn compile_func_cold(&mut self, prog: usize, func: FuncId) -> Rc<CompiledFunc> {
        let program = self.exec.programs[prog].program;
        let code = Rc::new(CompiledFunc::compile(program, &program.funcs[func.0]));
        self.cache.compiled[prog][func.0] = Some(Rc::clone(&code));
        code
    }

    fn recycle_frame(&mut self, mut frame: Frame) {
        if !self.exec.events.withevents.is_empty() {
            self.clear_withevents_owners_in_frame_before_drop(&frame);
        }
        let reset_end = frame.temp_hwm.max(frame.num_locals).min(frame.slots.len());
        for slot in &mut frame.slots[..reset_end] {
            if slot.raw_vt() != 0 {
                slot.set_empty();
            }
        }
        if self.cache.frame_pool.len() < 64 && frame.slots.capacity() <= 256 {
            if !frame.pointer_strings.is_empty() {
                frame.pointer_strings.clear();
            }
            if !frame.ref_locks.is_empty() {
                frame.ref_locks.clear();
            }
            if !frame.loop_locks.is_empty() {
                frame.loop_locks.clear();
            }
            if !frame.aliases.is_empty() {
                frame.aliases.clear();
            }
            frame.local_alias_mask = 0;
            if !frame.gosub_stack.is_empty() {
                frame.gosub_stack.clear();
            }
            frame.saved_active_error = None;
            self.cache.frame_pool.push(frame);
        }
    }

    fn new_frame(&mut self, func: FuncId) -> Result<Frame, Vm3Error> {
        self.new_frame_in(self.cur, func)
    }

    /// Build a fresh frame for `func` in program `prog` (locals sized from that program's
    /// function); the frame carries `prog`, so `run_loop` executes it against that program's
    /// globals/code regardless of the resolving `cur`. Used to run a cross-project callee (an
    /// object method, or a referenced project's proc) whose ARGUMENTS are resolved in the
    /// caller's program (`cur`) BEFORE this frame runs — matching vm2's resolve-then-switch.
    fn new_frame_in(&mut self, prog: usize, func: FuncId) -> Result<Frame, Vm3Error> {
        self.new_frame_pooled(prog, func, 0, None, None)
    }

    fn new_frame_pooled(
        &mut self,
        prog: usize,
        func: FuncId,
        skip_optional_below: usize,
        dst: Option<Loc>,
        receiver: Option<Variant>,
    ) -> Result<Frame, Vm3Error> {
        let drain_on_fault = receiver.is_some();
        self.ensure_compiled(prog, func);
        let program = self.exec.programs[prog].program;
        let f = &program.funcs[func.0];
        let code = self.cache.compiled[prog][func.0]
            .as_deref()
            .expect("compiled above");
        let num_slots = code.num_slots;
        let num_locals = code.num_locals;
        let pc = code.block_start_pc.get(f.entry.0).copied().unwrap_or(0) as usize;
        let mut frame = if let Some(mut pooled) = self.cache.frame_pool.pop() {
            if pooled.slots.len() < num_slots {
                pooled.slots.resize_with(num_slots, Variant::empty);
            } else if pooled.slots.len() > num_slots {
                pooled.slots.truncate(num_slots);
            }
            pooled.prog = prog;
            pooled.func = func;
            pooled.block = f.entry;
            pooled.pc = pc;
            pooled.at_block_start = true;
            pooled.num_locals = num_locals;
            pooled.temp_hwm = num_locals;
            pooled.dst = dst;
            pooled.return_local = f.return_local;
            pooled.drain_on_fault = drain_on_fault;
            pooled.current_line = 0;
            pooled.saved_error_mode = ErrorMode::None;
            pooled.saved_active_error = None;
            pooled
        } else {
            Frame {
                prog,
                func,
                block: f.entry,
                pc,
                at_block_start: true,
                slots: vec![Variant::empty(); num_slots],
                num_locals,
                temp_hwm: num_locals,
                pointer_strings: Vec::new(),
                ref_locks: Vec::new(),
                loop_locks: Vec::new(),
                aliases: Vec::new(),
                local_alias_mask: 0,
                dst,
                return_local: f.return_local,
                drain_on_fault,
                current_line: 0,
                saved_error_mode: ErrorMode::None,
                saved_active_error: None,
                gosub_stack: Vec::new(),
            }
        };
        if !code.non_empty_locals.is_empty() {
            for init in &code.non_empty_locals {
                match *init {
                    LocalSlotInit::Object(idx) => {
                        if idx != 0 || !drain_on_fault {
                            frame.slots[idx] = Variant::nothing();
                        }
                    }
                    LocalSlotInit::Array(idx) => {
                        let local = &f.locals[idx];
                        frame.slots[idx] = Self::initial_value_for_slot(
                            program,
                            &local.ty,
                            local.array_element.as_ref(),
                        )?;
                    }
                }
            }
        }
        if code.has_optional_params {
            for (index, local) in f
                .locals
                .iter()
                .take(f.param_count)
                .enumerate()
                .skip(skip_optional_below)
            {
                if let Some(value) = omitted_param_value(local) {
                    frame.slots[index] = value?;
                }
            }
        }
        if let Some(receiver) = receiver {
            frame.slots[0] = receiver;
        }
        Ok(frame)
    }

    /// The block-threaded dispatch loop: run the top frame until the frame at depth
    /// `base` returns. A `CallProc` pushes a callee and the loop simply continues with
    /// it; `Return` pops back to the caller. There is no native recursion, so deep VBA
    /// recursion is bounded by the frame ceiling (error 28), never a host stack
    /// overflow — and the model mirrors vm2's iterative dispatch.
    fn run_loop(&mut self, base: usize) -> Result<(), Vm3Error> {
        self.run_depth += 1;
        let result = self.run_loop_body(base);
        self.run_depth -= 1;
        match result {
            // `End` in a nested activation (a handler, a callback, a called method) reached the
            // outermost loop: end the run here as `End` at this level would.
            Err(Vm3Error::Halted) if self.run_depth == 0 => {
                self.exec.events.pending_abort = None;
                self.halt_to(base);
                Ok(())
            }
            other => other,
        }
    }

    /// VBA `End`: stop the *entire* program at any call depth, with no return to the caller and
    /// no finalization. Unwind to the base frame (which stays on the stack to back the snapshot).
    fn halt_to(&mut self, base: usize) {
        self.truncate_frames_with_withevents_cleanup(base + 1);
        self.prune_param_array_aliases_from_depth(self.frames.len());
        // Re-derive `cur` to the surviving (base/entry) frame's program: `End` may fire inside a
        // cross-program callee, and the post-run snapshot reader (`slot`) reads
        // `programs[cur]`'s globals, which must be the entry program's, to stay consistent with
        // the entry frame's locals it also reads.
        if let Some(frame) = self.frames.get(base) {
            self.cur = frame.prog;
        }
    }

    fn run_loop_body(&mut self, base: usize) -> Result<(), Vm3Error> {
        let _policy = self.exec.host.enter_com_policy();
        'dispatch: while self.frames.len() > base {
            let top = self.frames.len() - 1;
            let cur_prog = self.frames[top].prog;
            let cur_func = self.frames[top].func;
            self.cur = cur_prog;
            let code = self.get_or_compile(cur_prog, cur_func);
            let mut cur_block = self.frames[top].block.0 as u32;
            let mut pc = self.frames[top].pc;
            if self.frames[top].at_block_start {
                self.frames[top].at_block_start = false;
                if !self.exec.consume_execution_step() {
                    return Err(Vm3Error::BudgetExceeded);
                }
            }
            if (cur_block as usize) >= code.block_start_pc.len() {
                return Err(Vm3Error::Malformed(format!(
                    "block {cur_block} out of range"
                )));
            }
            let ops = code.ops.as_slice();
            let mut fast_frame = self.as_new_slots.is_empty()
                && self.param_array_aliases.is_empty()
                && self.exec.events.withevents.is_empty()
                && self.frames[top].local_alias_mask == 0;

            macro_rules! try_vm {
                ($expr:expr) => {
                    match $expr {
                        Ok(val) => val,
                        Err(Vm3Error::Fault(fault)) => {
                            let fr = &mut self.frames[top];
                            fr.block = BlockId(cur_block as usize);
                            fr.pc = pc;
                            fr.at_block_start = false;
                            self.route_fault(fault)?;
                            continue 'dispatch;
                        }
                        Err(other) => return Err(other),
                    }
                };
            }

            loop {
                let op = ops[pc];
                pc += 1;
                match op {
                    VmOp::LoadI32 { dst, val } => {
                        if fast_frame && dst.tag() <= TAG_LOCAL {
                            let fr = &mut self.frames[top];
                            let di = dst.index();
                            let slot = &mut fr.slots[di];
                            if slot.raw_vt() < 8 {
                                slot.set_i32_over_scalar(val);
                                if dst.tag() == TAG_TEMP && di + 1 > fr.temp_hwm {
                                    fr.temp_hwm = di + 1;
                                }
                                continue;
                            }
                        }
                        try_vm!(self.write_packed_i32(top, cur_prog, dst, val));
                    }
                    VmOp::LoadF64 {
                        dst,
                        bits_lo,
                        bits_hi,
                    } => {
                        let val = f64::from_bits(u64::from(bits_lo) | (u64::from(bits_hi) << 32));
                        if fast_frame && dst.tag() <= TAG_LOCAL {
                            let fr = &mut self.frames[top];
                            let di = dst.index();
                            let slot = &mut fr.slots[di];
                            if slot.raw_vt() < 8 {
                                slot.set_f64_over_scalar(val);
                                if dst.tag() == TAG_TEMP && di + 1 > fr.temp_hwm {
                                    fr.temp_hwm = di + 1;
                                }
                                continue;
                            }
                        }
                        try_vm!(self.write_packed_f64(top, cur_prog, dst, val));
                    }
                    VmOp::LoadBool { dst, val } => {
                        if fast_frame && dst.tag() <= TAG_LOCAL {
                            let fr = &mut self.frames[top];
                            let di = dst.index();
                            let slot = &mut fr.slots[di];
                            if slot.raw_vt() < 8 || slot.raw_vt() == 11 {
                                slot.set_bool_over_scalar(val);
                                if dst.tag() == TAG_TEMP && di + 1 > fr.temp_hwm {
                                    fr.temp_hwm = di + 1;
                                }
                                continue;
                            }
                        }
                        try_vm!(self.write_packed_bool(top, cur_prog, dst, val));
                    }
                    VmOp::Assign { dst, src } => {
                        if fast_frame && dst.tag() <= TAG_LOCAL && src.tag() <= TAG_LOCAL {
                            let fr = &mut self.frames[top];
                            let sv = &fr.slots[src.index()];
                            if Variant::raw_vt_is_trivial_copy(sv.raw_vt()) {
                                let core = sv.core();
                                let di = dst.index();
                                let slot = &mut fr.slots[di];
                                let old_vt = slot.raw_vt();
                                if old_vt < 8 || old_vt == 11 {
                                    slot.overwrite_trivial_core(core);
                                    if dst.tag() == TAG_TEMP && di + 1 > fr.temp_hwm {
                                        fr.temp_hwm = di + 1;
                                    }
                                    continue;
                                }
                            }
                        }
                        let trivial =
                            self.borrow_packed_op(top, cur_prog, &code, src)
                                .and_then(|v| {
                                    Variant::raw_vt_is_trivial_copy(v.raw_vt()).then(|| v.core())
                                });
                        if let Some(core) = trivial
                            && let Some(slot) = self.fast_dst_slot_mut(top, cur_prog, dst)
                        {
                            slot.overwrite_trivial_core(core);
                        } else {
                            let v = try_vm!(self.eval_packed_op(top, cur_prog, &code, src));
                            try_vm!(self.write_packed_assigned(top, cur_prog, dst, v));
                        }
                    }
                    VmOp::ArithLongImm { dst, lhs, imm, op } => {
                        if fast_frame && lhs.tag() <= TAG_LOCAL && dst.tag() <= TAG_LOCAL {
                            let fr = &mut self.frames[top];
                            let lv = &fr.slots[lhs.index()];
                            if lv.raw_vt() == 3 {
                                let a = lv.as_i32().unwrap();
                                let res = try_vm!(fast_i32_binop(a, imm, op));
                                let di = dst.index();
                                let slot = &mut fr.slots[di];
                                if slot.raw_vt() < 8 {
                                    slot.set_i32_over_scalar(res);
                                    if dst.tag() == TAG_TEMP && di + 1 > fr.temp_hwm {
                                        fr.temp_hwm = di + 1;
                                    }
                                    continue;
                                }
                            }
                        }
                        let fast = self
                            .borrow_packed_op(top, cur_prog, &code, lhs)
                            .and_then(try_fast_i32);
                        if let Some(a) = fast {
                            let res = try_vm!(fast_i32_binop(a, imm, op));
                            try_vm!(self.write_packed_i32(top, cur_prog, dst, res));
                        } else {
                            let l =
                                try_vm!(self.eval_packed_arith_operand(top, cur_prog, &code, lhs));
                            let r = Variant::from_i32(imm);
                            let out = eval_arith_op(
                                &l,
                                &r,
                                op,
                                NumericMode::Checked(NumericCoerceTarget::Long),
                            );
                            try_vm!(self.store_packed_arith(top, cur_prog, dst, out));
                        }
                    }
                    VmOp::ArithLong { dst, lhs, rhs, op } => {
                        if fast_frame
                            && lhs.tag() <= TAG_LOCAL
                            && rhs.tag() <= TAG_LOCAL
                            && dst.tag() <= TAG_LOCAL
                        {
                            let fr = &mut self.frames[top];
                            let lv = &fr.slots[lhs.index()];
                            let rv = &fr.slots[rhs.index()];
                            if lv.raw_vt() == 3 && rv.raw_vt() == 3 {
                                let a = lv.as_i32().unwrap();
                                let b = rv.as_i32().unwrap();
                                let res = try_vm!(fast_i32_binop(a, b, op));
                                let di = dst.index();
                                let slot = &mut fr.slots[di];
                                if slot.raw_vt() < 8 {
                                    slot.set_i32_over_scalar(res);
                                    if dst.tag() == TAG_TEMP && di + 1 > fr.temp_hwm {
                                        fr.temp_hwm = di + 1;
                                    }
                                    continue;
                                }
                            }
                        }
                        let fast = match (
                            self.borrow_packed_op(top, cur_prog, &code, lhs)
                                .and_then(try_fast_i32),
                            self.borrow_packed_op(top, cur_prog, &code, rhs)
                                .and_then(try_fast_i32),
                        ) {
                            (Some(a), Some(b)) => Some((a, b)),
                            _ => None,
                        };
                        if let Some((a, b)) = fast {
                            let res = try_vm!(fast_i32_binop(a, b, op));
                            try_vm!(self.write_packed_i32(top, cur_prog, dst, res));
                        } else {
                            let l =
                                try_vm!(self.eval_packed_arith_operand(top, cur_prog, &code, lhs));
                            let r =
                                try_vm!(self.eval_packed_arith_operand(top, cur_prog, &code, rhs));
                            let out = eval_arith_op(
                                &l,
                                &r,
                                op,
                                NumericMode::Checked(NumericCoerceTarget::Long),
                            );
                            try_vm!(self.store_packed_arith(top, cur_prog, dst, out));
                        }
                    }
                    VmOp::ArithDouble { dst, lhs, rhs, op } => {
                        if fast_frame
                            && lhs.tag() <= TAG_LOCAL
                            && rhs.tag() <= TAG_LOCAL
                            && dst.tag() <= TAG_LOCAL
                        {
                            let fr = &mut self.frames[top];
                            let lv = &fr.slots[lhs.index()];
                            let rv = &fr.slots[rhs.index()];
                            if let (Some(a), Some(b)) = (try_fast_f64(lv), try_fast_f64(rv)) {
                                let res = match op {
                                    ArithOp::Add => Some(a + b),
                                    ArithOp::Sub => Some(a - b),
                                    ArithOp::Mul => Some(a * b),
                                    _ => None,
                                }
                                .filter(|r| r.is_finite());
                                if let Some(val) = res {
                                    let di = dst.index();
                                    let slot = &mut fr.slots[di];
                                    if slot.raw_vt() < 8 {
                                        slot.set_f64_over_scalar(val);
                                        if dst.tag() == TAG_TEMP && di + 1 > fr.temp_hwm {
                                            fr.temp_hwm = di + 1;
                                        }
                                        continue;
                                    }
                                }
                            }
                        }
                        let fast = match (
                            self.borrow_packed_op(top, cur_prog, &code, lhs)
                                .and_then(try_fast_f64),
                            self.borrow_packed_op(top, cur_prog, &code, rhs)
                                .and_then(try_fast_f64),
                        ) {
                            (Some(a), Some(b)) => match op {
                                ArithOp::Add => Some(a + b),
                                ArithOp::Sub => Some(a - b),
                                ArithOp::Mul => Some(a * b),
                                _ => None,
                            }
                            .filter(|r| r.is_finite()),
                            _ => None,
                        };
                        if let Some(res) = fast {
                            try_vm!(self.write_packed_f64(top, cur_prog, dst, res));
                        } else {
                            let l =
                                try_vm!(self.eval_packed_arith_operand(top, cur_prog, &code, lhs));
                            let r =
                                try_vm!(self.eval_packed_arith_operand(top, cur_prog, &code, rhs));
                            let out = eval_arith_op(
                                &l,
                                &r,
                                op,
                                NumericMode::Checked(NumericCoerceTarget::Double),
                            );
                            try_vm!(self.store_packed_arith(top, cur_prog, dst, out));
                        }
                    }
                    VmOp::DivDouble { dst, lhs, rhs } => {
                        let fast = match (
                            self.borrow_packed_op(top, cur_prog, &code, lhs)
                                .and_then(try_fast_f64),
                            self.borrow_packed_op(top, cur_prog, &code, rhs)
                                .and_then(try_fast_f64),
                        ) {
                            (Some(a), Some(b)) if b != 0.0 => {
                                let q = a / b;
                                q.is_finite().then_some(q)
                            }
                            _ => None,
                        };
                        if let Some(res) = fast {
                            try_vm!(self.write_packed_f64(top, cur_prog, dst, res));
                        } else {
                            let l =
                                try_vm!(self.eval_packed_arith_operand(top, cur_prog, &code, lhs));
                            let r =
                                try_vm!(self.eval_packed_arith_operand(top, cur_prog, &code, rhs));
                            try_vm!(self.store_packed_arith(
                                top,
                                cur_prog,
                                dst,
                                arith::div(&l, &r)
                            ));
                        }
                    }
                    VmOp::ArithWidening { dst, lhs, rhs, op } => {
                        enum FastWide {
                            I32(i32),
                            F64(f64),
                        }
                        let fast = match (
                            self.borrow_packed_op(top, cur_prog, &code, lhs),
                            self.borrow_packed_op(top, cur_prog, &code, rhs),
                        ) {
                            (Some(lv), Some(rv)) => {
                                let lvt = lv.raw_vt();
                                let rvt = rv.raw_vt();
                                if lvt == 3 && rvt == 3 {
                                    let a = lv.as_i32().unwrap();
                                    let b = rv.as_i32().unwrap();
                                    match op {
                                        ArithOp::Add => a.checked_add(b).map(FastWide::I32),
                                        ArithOp::Sub => a.checked_sub(b).map(FastWide::I32),
                                        ArithOp::Mul => a.checked_mul(b).map(FastWide::I32),
                                        ArithOp::IntDiv if b != 0 => {
                                            a.checked_div(b).map(FastWide::I32)
                                        }
                                        ArithOp::Mod if b != 0 => {
                                            Some(FastWide::I32(if a == i32::MIN && b == -1 {
                                                0
                                            } else {
                                                a % b
                                            }))
                                        }
                                        _ => None,
                                    }
                                } else if (lvt == 5 || rvt == 5)
                                    && matches!(op, ArithOp::Add | ArithOp::Sub | ArithOp::Mul)
                                    && let (Some(a), Some(b)) = (try_fast_f64(lv), try_fast_f64(rv))
                                {
                                    let res = match op {
                                        ArithOp::Add => a + b,
                                        ArithOp::Sub => a - b,
                                        ArithOp::Mul => a * b,
                                        _ => unreachable!(),
                                    };
                                    res.is_finite().then_some(FastWide::F64(res))
                                } else if ((lvt == 3 && matches!(rvt, 0 | 2))
                                    || (rvt == 3 && matches!(lvt, 0 | 2)))
                                    && let (Some(a), Some(b)) = (try_fast_i32(lv), try_fast_i32(rv))
                                {
                                    match op {
                                        ArithOp::Add => a.checked_add(b).map(FastWide::I32),
                                        ArithOp::Sub => a.checked_sub(b).map(FastWide::I32),
                                        ArithOp::Mul => a.checked_mul(b).map(FastWide::I32),
                                        ArithOp::IntDiv if b != 0 => {
                                            a.checked_div(b).map(FastWide::I32)
                                        }
                                        ArithOp::Mod if b != 0 => {
                                            Some(FastWide::I32(if a == i32::MIN && b == -1 {
                                                0
                                            } else {
                                                a % b
                                            }))
                                        }
                                        _ => None,
                                    }
                                } else if matches!(op, ArithOp::Add | ArithOp::Sub | ArithOp::Mul)
                                    && ((matches!(lvt, 2 | 3 | 5) && rvt == 8)
                                        || (lvt == 8 && matches!(rvt, 2 | 3 | 5)))
                                    && let (Some(a), Some(b)) = (
                                        if lvt == 8 {
                                            lv.borrowed_string_units()
                                                .and_then(try_parse_ascii_int_f64)
                                        } else {
                                            try_fast_f64(lv)
                                        },
                                        if rvt == 8 {
                                            rv.borrowed_string_units()
                                                .and_then(try_parse_ascii_int_f64)
                                        } else {
                                            try_fast_f64(rv)
                                        },
                                    )
                                {
                                    let res = match op {
                                        ArithOp::Add => a + b,
                                        ArithOp::Sub => a - b,
                                        ArithOp::Mul => a * b,
                                        _ => unreachable!(),
                                    };
                                    res.is_finite().then_some(FastWide::F64(res))
                                } else {
                                    None
                                }
                            }
                            _ => None,
                        };
                        match fast {
                            Some(FastWide::I32(v)) => {
                                try_vm!(self.write_packed_i32(top, cur_prog, dst, v));
                            }
                            Some(FastWide::F64(v)) => {
                                try_vm!(self.write_packed_f64(top, cur_prog, dst, v));
                            }
                            None => {
                                let l = try_vm!(
                                    self.eval_packed_arith_operand(top, cur_prog, &code, lhs)
                                );
                                let r = try_vm!(
                                    self.eval_packed_arith_operand(top, cur_prog, &code, rhs)
                                );
                                let out = eval_arith_op(&l, &r, op, NumericMode::Widening);
                                try_vm!(self.store_packed_arith(top, cur_prog, dst, out));
                            }
                        }
                    }
                    VmOp::Neg { dst, src, mode } => {
                        let fast = self
                            .borrow_packed_op(top, cur_prog, &code, src)
                            .and_then(|v| match (mode, v.raw_vt()) {
                                (NumericMode::Checked(NumericCoerceTarget::Long), 3) => {
                                    v.as_i32().and_then(i32::checked_neg).map(Variant::from_i32)
                                }
                                (NumericMode::Checked(NumericCoerceTarget::Double), 5) => v
                                    .as_f64()
                                    .filter(|x| x.is_finite())
                                    .map(|x| Variant::from_f64(-x)),
                                _ => None,
                            });
                        if let Some(v) = fast {
                            try_vm!(self.write_packed_place(top, cur_prog, dst, v));
                        } else {
                            let val =
                                try_vm!(self.eval_packed_arith_value(top, cur_prog, &code, src));
                            try_vm!(self.store_packed_arith(
                                top,
                                cur_prog,
                                dst,
                                arith::neg(&val, mode)
                            ));
                        }
                    }
                    VmOp::Concat {
                        dst,
                        lhs,
                        rhs,
                        coerce_to_string,
                    } => {
                        let fast: Option<Variant> = {
                            let lu = self.borrow_packed_str_units(top, cur_prog, &code, lhs);
                            let ru = self.borrow_packed_str_units(top, cur_prog, &code, rhs);
                            match (lu, ru) {
                                (Some(a), Some(b)) => Some(Variant::from_utf16_concat(a, b)),
                                (Some(a), None) => self
                                    .borrow_packed_op(top, cur_prog, &code, rhs)
                                    .and_then(|rv| {
                                        (rv.raw_vt() == 3).then(|| rv.as_i32()).flatten()
                                    })
                                    .map(|n| {
                                        let mut buf = [0u16; 12];
                                        let digits = format_i32_utf16(n, &mut buf);
                                        Variant::from_utf16_concat(a, digits)
                                    }),
                                _ => None,
                            }
                        };
                        if let Some(res) = fast {
                            try_vm!(self.write_packed_place(top, cur_prog, dst, res));
                        } else {
                            let l =
                                try_vm!(self.eval_packed_arith_value(top, cur_prog, &code, lhs));
                            let r =
                                try_vm!(self.eval_packed_arith_value(top, cur_prog, &code, rhs));
                            let result = arith::concat(&l, &r).and_then(|value| {
                                if coerce_to_string && value.vtype() != VarType::String {
                                    arith::coerce_string(&value)
                                } else {
                                    Ok(value)
                                }
                            });
                            try_vm!(self.store_packed_arith(top, cur_prog, dst, result));
                        }
                    }
                    VmOp::Compare {
                        dst,
                        lhs,
                        rhs,
                        op,
                        mode,
                        coercion,
                    } => {
                        let fast = if coercion != ComparisonCoercion::String {
                            match (
                                self.borrow_packed_op(top, cur_prog, &code, lhs),
                                self.borrow_packed_op(top, cur_prog, &code, rhs),
                            ) {
                                (Some(lv), Some(rv)) => try_fast_cmp(lv, rv, op),
                                _ => None,
                            }
                        } else {
                            None
                        };
                        if let Some(b) = fast {
                            try_vm!(self.write_packed_bool(top, cur_prog, dst, b));
                        } else {
                            let l =
                                try_vm!(self.eval_packed_arith_value(top, cur_prog, &code, lhs));
                            let r =
                                try_vm!(self.eval_packed_arith_value(top, cur_prog, &code, rhs));
                            let out = arith::compare_declared(&l, &r, mode, cmp_op(op), coercion);
                            try_vm!(self.store_packed_arith(top, cur_prog, dst, out));
                        }
                    }
                    VmOp::CmpBranch(idx) => {
                        let data = &code.cmp_branches[idx as usize];
                        let fast = if data.coercion != ComparisonCoercion::String {
                            match (
                                self.borrow_packed_op(top, cur_prog, &code, data.lhs),
                                self.borrow_packed_op(top, cur_prog, &code, data.rhs),
                            ) {
                                (Some(lv), Some(rv)) => try_fast_cmp(lv, rv, data.op),
                                _ => None,
                            }
                        } else {
                            None
                        };
                        let taken = match fast {
                            Some(b) => b,
                            None => {
                                let l = try_vm!(
                                    self.eval_packed_arith_value(top, cur_prog, &code, data.lhs)
                                );
                                let r = try_vm!(
                                    self.eval_packed_arith_value(top, cur_prog, &code, data.rhs)
                                );
                                let cmp_val = try_vm!(
                                    arith::compare_declared(
                                        &l,
                                        &r,
                                        data.mode,
                                        cmp_op(data.op),
                                        data.coercion,
                                    )
                                    .map_err(arith_fault)
                                );
                                try_vm!(arith::is_truthy(&cmp_val).map_err(arith_fault))
                            }
                        };
                        let steps = if taken {
                            cur_block = data.then_block;
                            pc = data.then_pc as usize;
                            u64::from(data.then_steps)
                        } else {
                            cur_block = data.else_block;
                            pc = data.else_pc as usize;
                            u64::from(data.else_steps)
                        };
                        if !self.exec.consume_execution_steps(steps) {
                            return Err(Vm3Error::BudgetExceeded);
                        }
                    }
                    VmOp::ForCheck(idx) => {
                        let data = &code.for_checks[idx as usize];
                        let max_steps = u64::from(data.body_steps.max(data.after_steps));
                        if self.exec.remaining_steps.is_none_or(|r| r >= max_steps)
                            && let Some(nn_v) =
                                self.borrow_packed_op(top, cur_prog, &code, data.nonneg)
                            && nn_v.raw_vt() == 11
                            && let Some(nn) = nn_v.as_bool()
                            && let Some(var_v) = self.borrow_packed_op(
                                top,
                                cur_prog,
                                &code,
                                PackedOp::from_place(data.var),
                            )
                            && let Some(lim_v) =
                                self.borrow_packed_op(top, cur_prog, &code, data.limit)
                        {
                            let take_body = if let (Some(v), Some(l)) =
                                (try_fast_i32(var_v), try_fast_i32(lim_v))
                            {
                                Some(if nn { v <= l } else { v >= l })
                            } else if let (Some(v), Some(l)) =
                                (try_fast_f64(var_v), try_fast_f64(lim_v))
                            {
                                (!v.is_nan() && !l.is_nan()).then_some(if nn {
                                    v <= l
                                } else {
                                    v >= l
                                })
                            } else {
                                None
                            };
                            if let Some(cont) = take_body {
                                let steps = if cont {
                                    cur_block = data.body_block;
                                    pc = data.body_pc as usize;
                                    u64::from(data.body_steps)
                                } else {
                                    cur_block = data.after_block;
                                    pc = data.after_pc as usize;
                                    u64::from(data.after_steps)
                                };
                                let _ = self.exec.consume_execution_steps(steps);
                            }
                        }
                    }
                    VmOp::ForStepCheck(idx) => {
                        let data = &code.for_step_checks[idx as usize];
                        let max_steps = u64::from(data.cont_steps.max(data.exit_steps));
                        // Ultra-fast direct-slot path for local `Long` For counters (step/limit may be I16 or I32).
                        if data.var.tag() == TAG_LOCAL
                            && data.step.tag() == TAG_TEMP
                            && data.limit.tag() == TAG_TEMP
                            && data.nonneg.tag() == TAG_TEMP
                            && fast_frame
                            && self.exec.remaining_steps.is_none_or(|r| r >= max_steps)
                        {
                            let fr = &mut self.frames[top];
                            let vi = data.var.index();
                            let slots = fr.slots.as_mut_slice();
                            let var_s = &slots[vi];
                            let step_s = &slots[data.step.index()];
                            let lim_s = &slots[data.limit.index()];
                            let nn_s = &slots[data.nonneg.index()];
                            if var_s.raw_vt() == 3
                                && nn_s.raw_vt() == 11
                                && let Some(s) = try_fast_i32(step_s)
                                && let Some(lim) = try_fast_i32(lim_s)
                            {
                                let v = var_s.as_i32().unwrap();
                                let nn = nn_s.as_bool().unwrap();
                                if let Some(next_v) = v.checked_add(s) {
                                    slots[vi].set_i32_over_scalar(next_v);
                                    let cont = if nn { next_v <= lim } else { next_v >= lim };
                                    let steps = if cont {
                                        cur_block = data.body_block;
                                        pc = data.body_pc as usize;
                                        u64::from(data.cont_steps)
                                    } else {
                                        cur_block = data.after_block;
                                        pc = data.after_pc as usize;
                                        u64::from(data.exit_steps)
                                    };
                                    let _ = self.exec.consume_execution_steps(steps);
                                    continue;
                                }
                                if data.to_step_steps > 0
                                    && !self
                                        .exec
                                        .consume_execution_steps(u64::from(data.to_step_steps))
                                {
                                    return Err(Vm3Error::BudgetExceeded);
                                }
                                cur_block = data.step_block;
                                pc = data.slow_pc as usize;
                                try_vm!(Err(arith_fault(ArithError::overflow())));
                            }
                        }
                        if self.exec.remaining_steps.is_none_or(|r| r >= max_steps)
                            && let Some(nn_v) =
                                self.borrow_packed_op(top, cur_prog, &code, data.nonneg)
                            && nn_v.raw_vt() == 11
                            && let Some(nn) = nn_v.as_bool()
                            && let Some(s) = self
                                .borrow_packed_op(top, cur_prog, &code, data.step)
                                .and_then(try_fast_i32)
                            && let Some(lim) = self
                                .borrow_packed_op(top, cur_prog, &code, data.limit)
                                .and_then(try_fast_i32)
                            && let Some(v) = self
                                .borrow_packed_op(
                                    top,
                                    cur_prog,
                                    &code,
                                    PackedOp::from_place(data.var),
                                )
                                .and_then(try_fast_i32)
                        {
                            if let Some(next_v) = v.checked_add(s) {
                                if let Some(slot) = self.fast_dst_slot_mut(top, cur_prog, data.var)
                                {
                                    slot.set_i32(next_v);
                                    let cont = if nn { next_v <= lim } else { next_v >= lim };
                                    let steps = if cont {
                                        cur_block = data.body_block;
                                        pc = data.body_pc as usize;
                                        u64::from(data.cont_steps)
                                    } else {
                                        cur_block = data.after_block;
                                        pc = data.after_pc as usize;
                                        u64::from(data.exit_steps)
                                    };
                                    let _ = self.exec.consume_execution_steps(steps);
                                    continue;
                                }
                            } else {
                                if data.to_step_steps > 0
                                    && !self
                                        .exec
                                        .consume_execution_steps(u64::from(data.to_step_steps))
                                {
                                    return Err(Vm3Error::BudgetExceeded);
                                }
                                cur_block = data.step_block;
                                pc = data.slow_pc as usize;
                                try_vm!(Err(arith_fault(ArithError::overflow())));
                            }
                        }
                        if data.to_step_steps > 0 {
                            if !self
                                .exec
                                .consume_execution_steps(u64::from(data.to_step_steps))
                            {
                                return Err(Vm3Error::BudgetExceeded);
                            }
                            cur_block = data.step_block;
                            pc = data.slow_pc as usize;
                        }
                    }
                    VmOp::Truthy { dst, src } => {
                        let fast = self
                            .borrow_packed_op(top, cur_prog, &code, src)
                            .and_then(|v| match v.raw_vt() {
                                11 => v.as_bool(),
                                3 => v.as_i32().map(|n| n != 0),
                                2 => v.as_i16().map(|n| n != 0),
                                0 | 1 => Some(false),
                                _ => None,
                            });
                        if let Some(b) = fast {
                            try_vm!(self.write_packed_bool(top, cur_prog, dst, b));
                        } else {
                            let val =
                                try_vm!(self.eval_packed_arith_value(top, cur_prog, &code, src));
                            let out = arith::is_truthy(&val).map(Variant::from_bool);
                            try_vm!(self.store_packed_arith(top, cur_prog, dst, out));
                        }
                    }
                    VmOp::CoerceNum { dst, src, target } => {
                        enum FastNum {
                            I32(i32),
                            F64(f64),
                            Bool(bool),
                            Other(Variant),
                        }
                        let fast = self
                            .borrow_packed_op(top, cur_prog, &code, src)
                            .and_then(|v| match (target, v.raw_vt()) {
                                (NumericCoerceTarget::Long, 3) => v.as_i32().map(FastNum::I32),
                                (NumericCoerceTarget::Long, 2) => {
                                    v.as_i16().map(|n| FastNum::I32(i32::from(n)))
                                }
                                (NumericCoerceTarget::Long, 0) => Some(FastNum::I32(0)),
                                (NumericCoerceTarget::Double, 5) => {
                                    v.as_f64().filter(|x| x.is_finite()).map(FastNum::F64)
                                }
                                (NumericCoerceTarget::Double, 3) => {
                                    v.as_i32().map(|n| FastNum::F64(f64::from(n)))
                                }
                                (NumericCoerceTarget::Double, 2) => {
                                    v.as_i16().map(|n| FastNum::F64(f64::from(n)))
                                }
                                (NumericCoerceTarget::Double, 0) => Some(FastNum::F64(0.0)),
                                (NumericCoerceTarget::Integer, 2) => {
                                    v.as_i16().map(|n| FastNum::Other(Variant::from_i16(n)))
                                }
                                (NumericCoerceTarget::Integer, 3) => v
                                    .as_i32()
                                    .and_then(|n| i16::try_from(n).ok())
                                    .map(|n| FastNum::Other(Variant::from_i16(n))),
                                (NumericCoerceTarget::Integer, 0) => {
                                    Some(FastNum::Other(Variant::from_i16(0)))
                                }
                                (NumericCoerceTarget::Boolean, 11) => {
                                    v.as_bool().map(FastNum::Bool)
                                }
                                (NumericCoerceTarget::Boolean, 3) => {
                                    v.as_i32().map(|n| FastNum::Bool(n != 0))
                                }
                                (NumericCoerceTarget::Boolean, 0) => Some(FastNum::Bool(false)),
                                _ => None,
                            });
                        match fast {
                            Some(FastNum::I32(v)) => {
                                try_vm!(self.write_packed_i32(top, cur_prog, dst, v));
                            }
                            Some(FastNum::F64(v)) => {
                                try_vm!(self.write_packed_f64(top, cur_prog, dst, v));
                            }
                            Some(FastNum::Bool(v)) => {
                                try_vm!(self.write_packed_bool(top, cur_prog, dst, v));
                            }
                            Some(FastNum::Other(v)) => {
                                try_vm!(self.write_packed_place(top, cur_prog, dst, v));
                            }
                            None => {
                                let val = try_vm!(
                                    self.eval_packed_arith_value(top, cur_prog, &code, src)
                                );
                                let out = arith::coerce_numeric(&val, target);
                                try_vm!(self.store_packed_arith(top, cur_prog, dst, out));
                            }
                        }
                    }
                    VmOp::CoerceStr { dst, src } => {
                        let fast = if src.tag() == TAG_CONST {
                            match &code.consts[src.index()] {
                                FastConst::Str(units) => Some(Variant::from_utf16_units(units)),
                                FastConst::NullStr => Some(Variant::null_string()),
                                _ => None,
                            }
                        } else {
                            self.borrow_packed_op(top, cur_prog, &code, src)
                                .filter(|v| v.raw_vt() == 8)
                                .cloned()
                        };
                        if let Some(v) = fast {
                            try_vm!(self.write_packed_place(top, cur_prog, dst, v));
                        } else {
                            let val =
                                try_vm!(self.eval_packed_arith_value(top, cur_prog, &code, src));
                            let out = arith::coerce_string(&val);
                            try_vm!(self.store_packed_arith(top, cur_prog, dst, out));
                        }
                    }
                    VmOp::Not { dst, src } => {
                        let fast = self
                            .borrow_packed_op(top, cur_prog, &code, src)
                            .and_then(|v| match v.raw_vt() {
                                11 => v.as_bool().map(|b| Variant::from_bool(!b)),
                                3 => v.as_i32().map(|n| Variant::from_i32(!n)),
                                _ => None,
                            });
                        if let Some(v) = fast {
                            try_vm!(self.write_packed_place(top, cur_prog, dst, v));
                        } else {
                            let val =
                                try_vm!(self.eval_packed_arith_value(top, cur_prog, &code, src));
                            try_vm!(self.store_packed_arith(top, cur_prog, dst, arith::not(&val)));
                        }
                    }
                    VmOp::Logical { dst, lhs, rhs, op } => {
                        let fast = match (
                            self.borrow_packed_op(top, cur_prog, &code, lhs),
                            self.borrow_packed_op(top, cur_prog, &code, rhs),
                        ) {
                            (Some(lv), Some(rv)) if lv.raw_vt() == 11 && rv.raw_vt() == 11 => {
                                let a = lv.as_bool().unwrap();
                                let b = rv.as_bool().unwrap();
                                Some(Variant::from_bool(match op {
                                    LogicalOp::And => a && b,
                                    LogicalOp::Or => a || b,
                                    LogicalOp::Xor => a ^ b,
                                    LogicalOp::Eqv => a == b,
                                    LogicalOp::Imp => !a || b,
                                }))
                            }
                            (Some(lv), Some(rv)) if lv.raw_vt() == 3 && rv.raw_vt() == 3 => {
                                let a = lv.as_i32().unwrap();
                                let b = rv.as_i32().unwrap();
                                Some(Variant::from_i32(match op {
                                    LogicalOp::And => a & b,
                                    LogicalOp::Or => a | b,
                                    LogicalOp::Xor => a ^ b,
                                    LogicalOp::Eqv => !(a ^ b),
                                    LogicalOp::Imp => (!a) | b,
                                }))
                            }
                            _ => None,
                        };
                        if let Some(v) = fast {
                            try_vm!(self.write_packed_place(top, cur_prog, dst, v));
                        } else {
                            let l =
                                try_vm!(self.eval_packed_arith_value(top, cur_prog, &code, lhs));
                            let r =
                                try_vm!(self.eval_packed_arith_value(top, cur_prog, &code, rhs));
                            let out = match op {
                                LogicalOp::And => arith::and(&l, &r),
                                LogicalOp::Or => arith::or(&l, &r),
                                LogicalOp::Xor => arith::xor(&l, &r),
                                LogicalOp::Eqv => arith::eqv(&l, &r),
                                LogicalOp::Imp => arith::imp(&l, &r),
                            };
                            try_vm!(self.store_packed_arith(top, cur_prog, dst, out));
                        }
                    }
                    VmOp::ArrayGet1D { dst, array, index } => {
                        let arr_loc = self.resolve_packed(top, cur_prog, array);
                        let is_arr = self
                            .read_loc_ref_fast(arr_loc)
                            .is_some_and(|v| matches!(v.raw_vt() & 0xF000, 0x2000 | 0x6000));
                        if is_arr {
                            let sub =
                                try_vm!(self.eval_fast_subscript(top, cur_prog, &code, index));
                            if let Some((arr_ref, dst_slot)) =
                                self.borrow_loc_and_fast_dst_mut(top, cur_prog, arr_loc, dst)
                                && let Some(res) = arr_ref.safearray_1d_read_into(sub, dst_slot)
                            {
                                match res {
                                    Ok(()) => continue,
                                    Err((num, msg)) => {
                                        try_vm!(Err(Vm3Error::Fault(Fault::new(num, msg))));
                                    }
                                }
                            }
                            let fast = self
                                .read_loc_ref_fast(arr_loc)
                                .and_then(|arr| arr.safearray_1d_read(sub));
                            match fast {
                                Some(Ok(val)) => {
                                    try_vm!(self.write_packed_place(top, cur_prog, dst, val));
                                    continue;
                                }
                                Some(Err((num, msg))) => {
                                    try_vm!(Err(Vm3Error::Fault(Fault::new(num, msg))));
                                }
                                None => {}
                            }
                        }
                        let (iref, count) = code.fallback_inst_for_pc((pc - 1) as u32);
                        let switched = try_vm!(self.exec_slow_inst_range(
                            top, cur_prog, cur_func, cur_block, pc, iref, count
                        ));
                        if switched {
                            continue 'dispatch;
                        }
                    }
                    VmOp::ArraySet1D {
                        array,
                        index,
                        value,
                    } => {
                        if self.param_array_aliases.is_empty() {
                            let arr_loc = self.resolve_packed(top, cur_prog, array);
                            let is_arr = self
                                .read_loc_ref_fast(arr_loc)
                                .is_some_and(|v| matches!(v.raw_vt() & 0xF000, 0x2000 | 0x6000));
                            if is_arr {
                                let scalar_val = self
                                    .borrow_packed_op(top, cur_prog, &code, value)
                                    .and_then(|v| match v.raw_vt() {
                                        3 => v.as_i32().map(FastArrScalar::I32),
                                        5 => v.as_f64().map(FastArrScalar::F64),
                                        _ => None,
                                    });
                                if let Some(sv) = scalar_val {
                                    let sub = try_vm!(
                                        self.eval_fast_subscript(top, cur_prog, &code, index)
                                    );
                                    let fast =
                                        self.read_loc_mut_fast(arr_loc).and_then(|arr| match sv {
                                            FastArrScalar::I32(n) => {
                                                arr.safearray_1d_write_i32(sub, n)
                                            }
                                            FastArrScalar::F64(f) => {
                                                arr.safearray_1d_write_f64(sub, f)
                                            }
                                        });
                                    match fast {
                                        Some(Ok(())) => continue,
                                        Some(Err((num, msg))) => {
                                            try_vm!(Err(Vm3Error::Fault(Fault::new(num, msg))));
                                        }
                                        None => {}
                                    }
                                }
                                let val = try_vm!(self.eval_packed_op(top, cur_prog, &code, value));
                                let sub =
                                    try_vm!(self.eval_fast_subscript(top, cur_prog, &code, index));
                                let fast = self
                                    .read_loc_mut_fast(arr_loc)
                                    .and_then(|arr| arr.safearray_1d_write(sub, &val));
                                match fast {
                                    Some(Ok(())) => continue,
                                    Some(Err((num, msg))) => {
                                        try_vm!(Err(Vm3Error::Fault(Fault::new(num, msg))));
                                    }
                                    None => {}
                                }
                            }
                        }
                        let (iref, count) = code.fallback_inst_for_pc((pc - 1) as u32);
                        let switched = try_vm!(self.exec_slow_inst_range(
                            top, cur_prog, cur_func, cur_block, pc, iref, count
                        ));
                        if switched {
                            continue 'dispatch;
                        }
                    }
                    VmOp::RecordGet { dst, record, field } => {
                        let loc = self.resolve_packed(top, cur_prog, record);
                        if let Some(rec) = self.read_loc_ref_fast(loc)
                            && rec.raw_vt() == VarType::Record as u16
                        {
                            let val = try_vm!(
                                rec.read_record_field_variant(field as usize)
                                    .map_err(|e| Vm3Error::Fault(Fault::new(13, e)))
                            );
                            try_vm!(self.write_packed_place(top, cur_prog, dst, val));
                            continue;
                        }
                        let (iref, count) = code.fallback_inst_for_pc((pc - 1) as u32);
                        let switched = try_vm!(self.exec_slow_inst_range(
                            top, cur_prog, cur_func, cur_block, pc, iref, count
                        ));
                        if switched {
                            continue 'dispatch;
                        }
                    }
                    VmOp::RecordSet {
                        record,
                        field,
                        value,
                    } => {
                        let loc = self.resolve_packed(top, cur_prog, record);
                        if self
                            .read_loc_ref_fast(loc)
                            .is_some_and(|r| r.raw_vt() == VarType::Record as u16)
                        {
                            let val = try_vm!(self.eval_packed_op(top, cur_prog, &code, value));
                            if let Some(rec) = self.read_loc_mut_fast(loc)
                                && rec.raw_vt() == VarType::Record as u16
                            {
                                try_vm!(
                                    rec.write_record_field_variant(field as usize, &val)
                                        .map_err(|e| Vm3Error::Fault(Fault::new(13, e)))
                                );
                                continue;
                            }
                        }
                        let (iref, count) = code.fallback_inst_for_pc((pc - 1) as u32);
                        let switched = try_vm!(self.exec_slow_inst_range(
                            top, cur_prog, cur_func, cur_block, pc, iref, count
                        ));
                        if switched {
                            continue 'dispatch;
                        }
                    }
                    VmOp::ArrayRecordFieldGet1D(data_idx) => {
                        let data = &code.arr_rec_gets[data_idx as usize];
                        let arr_loc = self.resolve_packed(top, cur_prog, data.array);
                        if self
                            .read_loc_ref_fast(arr_loc)
                            .is_some_and(|v| v.raw_vt() == 0x2024)
                        {
                            let sub =
                                try_vm!(self.eval_fast_subscript(top, cur_prog, &code, data.index));
                            if let Some((arr_ref, dst_slot)) =
                                self.borrow_loc_and_fast_dst_mut(top, cur_prog, arr_loc, data.dst)
                                && let Some(res) = arr_ref.safearray_1d_record_field_read_into(
                                    sub,
                                    data.field as usize,
                                    dst_slot,
                                )
                            {
                                match res {
                                    Ok(()) => continue,
                                    Err((num, msg)) => {
                                        try_vm!(Err(Vm3Error::Fault(Fault::new(num, msg))));
                                    }
                                }
                            }
                            let fast = self.read_loc_ref_fast(arr_loc).and_then(|arr| {
                                arr.safearray_1d_record_field_read(sub, data.field as usize)
                            });
                            match fast {
                                Some(Ok(val)) => {
                                    try_vm!(self.write_packed_place(top, cur_prog, data.dst, val));
                                    continue;
                                }
                                Some(Err((num, msg))) => {
                                    try_vm!(Err(Vm3Error::Fault(Fault::new(num, msg))));
                                }
                                None => {}
                            }
                        }
                        let switched = try_vm!(self.exec_slow_inst_range(
                            top,
                            cur_prog,
                            cur_func,
                            cur_block,
                            pc,
                            data.first_inst,
                            data.orig_count as usize
                        ));
                        if switched {
                            continue 'dispatch;
                        }
                    }
                    VmOp::ArrayRecordFieldSet1D(data_idx) => {
                        let data = &code.arr_rec_sets[data_idx as usize];
                        let arr_loc = self.resolve_packed(top, cur_prog, data.array);
                        if self.param_array_aliases.is_empty()
                            && self
                                .read_loc_ref_fast(arr_loc)
                                .is_some_and(|v| v.raw_vt() == 0x2024)
                        {
                            let scalar_val = self
                                .borrow_packed_op(top, cur_prog, &code, data.value)
                                .and_then(|v| match v.raw_vt() {
                                    5 => v.as_f64().map(FastArrScalar::F64),
                                    3 => v.as_i32().map(FastArrScalar::I32),
                                    _ => None,
                                });
                            if let Some(sv) = scalar_val {
                                let sub = try_vm!(
                                    self.eval_fast_subscript(top, cur_prog, &code, data.index)
                                );
                                let fast =
                                    self.read_loc_mut_fast(arr_loc).and_then(|arr| match sv {
                                        FastArrScalar::F64(f) => arr
                                            .safearray_1d_record_field_write_f64(
                                                sub,
                                                data.field as usize,
                                                f,
                                            ),
                                        FastArrScalar::I32(n) => arr
                                            .safearray_1d_record_field_write_i32(
                                                sub,
                                                data.field as usize,
                                                n,
                                            ),
                                    });
                                match fast {
                                    Some(Ok(())) => continue,
                                    Some(Err((num, msg))) => {
                                        try_vm!(Err(Vm3Error::Fault(Fault::new(num, msg))));
                                    }
                                    None => {}
                                }
                            }
                            let val =
                                try_vm!(self.eval_packed_op(top, cur_prog, &code, data.value));
                            let sub =
                                try_vm!(self.eval_fast_subscript(top, cur_prog, &code, data.index));
                            let fast = self.read_loc_mut_fast(arr_loc).and_then(|arr| {
                                arr.safearray_1d_record_field_write(sub, data.field as usize, &val)
                            });
                            match fast {
                                Some(Ok(())) => continue,
                                Some(Err((num, msg))) => {
                                    try_vm!(Err(Vm3Error::Fault(Fault::new(num, msg))));
                                }
                                None => {}
                            }
                        }
                        let switched = try_vm!(self.exec_slow_inst_range(
                            top,
                            cur_prog,
                            cur_func,
                            cur_block,
                            pc,
                            data.first_inst,
                            data.orig_count as usize
                        ));
                        if switched {
                            continue 'dispatch;
                        }
                    }
                    VmOp::FieldGet { dst, object, field } => {
                        let done = if let Some(obj_place) = object.as_place()
                            && self.as_new_slots.is_empty()
                        {
                            let obj_loc = self.resolve_packed(top, cur_prog, obj_place);
                            self.borrow_loc_and_fast_dst_mut(top, cur_prog, obj_loc, dst)
                                .is_some_and(|(obj_ref, dst_slot)| {
                                    obj_ref.byref_target().is_none()
                                        && borrow_project_field_get_into(obj_ref, field, dst_slot)
                                })
                        } else {
                            false
                        };
                        if !done {
                            let recv = try_vm!(self.eval_packed_op(top, cur_prog, &code, object));
                            let instance = try_vm!(variant_to_object(&recv));
                            let val = try_vm!(self.read_project_field_as_new(&instance, field));
                            try_vm!(self.write_packed_place(top, cur_prog, dst, val));
                        }
                    }
                    VmOp::FieldSet {
                        object,
                        field,
                        value,
                    } => {
                        let done = if let (Some(obj_ref), Some(val_ref)) = (
                            self.borrow_packed_op(top, cur_prog, &code, object),
                            self.borrow_packed_op(top, cur_prog, &code, value),
                        ) {
                            borrow_project_field_set_trivial(obj_ref, field, val_ref)
                        } else {
                            false
                        };
                        if !done {
                            let v = try_vm!(self.eval_packed_op(top, cur_prog, &code, value));
                            let can_borrow = self
                                .borrow_packed_op(top, cur_prog, &code, object)
                                .and_then(borrow_project_instance_info)
                                .is_some();
                            if can_borrow {
                                let obj = self
                                    .borrow_packed_op(top, cur_prog, &code, object)
                                    .expect("verified project instance above");
                                borrow_project_field_set(obj, field, v);
                            } else {
                                let recv =
                                    try_vm!(self.eval_packed_op(top, cur_prog, &code, object));
                                let instance = try_vm!(variant_to_object(&recv));
                                instance.project_field_set(field, v);
                            }
                        }
                    }
                    VmOp::CallProcFast(idx) => {
                        let fr = &mut self.frames[top];
                        fr.block = BlockId(cur_block as usize);
                        fr.pc = pc;
                        fr.at_block_start = false;
                        let pushed = try_vm!(self.exec_call_proc_fast(top, cur_prog, &code, idx));
                        if pushed {
                            continue 'dispatch;
                        }
                        fast_frame = self.as_new_slots.is_empty()
                            && self.param_array_aliases.is_empty()
                            && self.exec.events.withevents.is_empty()
                            && self.frames[top].local_alias_mask == 0;
                    }
                    VmOp::CallExternLib(idx) => {
                        let fr = &mut self.frames[top];
                        fr.block = BlockId(cur_block as usize);
                        fr.pc = pc;
                        fr.at_block_start = false;
                        try_vm!(self.exec_call_extern_lib(top, cur_prog, &code, idx));
                        if !code.lib_calls[idx as usize].pure_no_events {
                            fast_frame = self.as_new_slots.is_empty()
                                && self.param_array_aliases.is_empty()
                                && self.exec.events.withevents.is_empty()
                                && self.frames[top].local_alias_mask == 0;
                        }
                    }
                    VmOp::CallComEarly(idx) => {
                        let fr = &mut self.frames[top];
                        fr.block = BlockId(cur_block as usize);
                        fr.pc = pc;
                        fr.at_block_start = false;
                        let pushed =
                            try_vm!(self.exec_call_com_early(top, cur_prog, cur_func, &code, idx));
                        if pushed {
                            continue 'dispatch;
                        }
                        fast_frame = self.as_new_slots.is_empty()
                            && self.param_array_aliases.is_empty()
                            && self.exec.events.withevents.is_empty()
                            && self.frames[top].local_alias_mask == 0;
                    }
                    VmOp::CallComLate(idx) => {
                        let fr = &mut self.frames[top];
                        fr.block = BlockId(cur_block as usize);
                        fr.pc = pc;
                        fr.at_block_start = false;
                        let pushed =
                            try_vm!(self.exec_call_com_late(top, cur_prog, cur_func, &code, idx));
                        if pushed {
                            continue 'dispatch;
                        }
                        fast_frame = self.as_new_slots.is_empty()
                            && self.param_array_aliases.is_empty()
                            && self.exec.events.withevents.is_empty()
                            && self.frames[top].local_alias_mask == 0;
                    }
                    VmOp::StmtBoundary { clear_temps_from } => {
                        let fr = &mut self.frames[top];
                        let first_temp = clear_temps_from as usize;
                        let first_slot = fr.num_locals + first_temp;
                        if first_slot < fr.temp_hwm {
                            let end = fr.temp_hwm.min(fr.slots.len());
                            for slot in &mut fr.slots[first_slot..end] {
                                let vt = slot.raw_vt();
                                if vt != 0 {
                                    if vt < 8 || vt == 11 {
                                        slot.set_empty_over_scalar();
                                    } else {
                                        slot.set_empty();
                                    }
                                }
                            }
                            fr.temp_hwm = first_slot;
                        }
                        if !fr.pointer_strings.is_empty()
                            || !fr.ref_locks.is_empty()
                            || !self.for_each.is_empty()
                            || !self.param_array_aliases.is_empty()
                        {
                            self.clear_statement_temps(first_temp);
                        }
                        if self.exec.object_owner.has_pending() {
                            let fr = &mut self.frames[top];
                            fr.block = BlockId(cur_block as usize);
                            fr.pc = pc;
                            fr.at_block_start = false;
                            self.maybe_drain();
                            fast_frame = self.as_new_slots.is_empty()
                                && self.param_array_aliases.is_empty()
                                && self.exec.events.withevents.is_empty()
                                && self.frames[top].local_alias_mask == 0;
                        }
                    }
                    VmOp::Jump {
                        pc: target_pc,
                        block: target_block,
                        steps,
                    } => {
                        cur_block = target_block;
                        pc = target_pc as usize;
                        if !self.exec.consume_execution_steps(u64::from(steps)) {
                            return Err(Vm3Error::BudgetExceeded);
                        }
                    }
                    VmOp::Branch(idx) => {
                        let data = &code.branches[idx as usize];
                        let taken = match self
                            .borrow_packed_op(top, cur_prog, &code, data.cond)
                            .and_then(Variant::as_bool)
                        {
                            Some(b) => b,
                            None => {
                                let v =
                                    try_vm!(self.eval_packed_op(top, cur_prog, &code, data.cond));
                                v.as_bool().ok_or_else(|| {
                                    Vm3Error::Malformed(
                                        "Branch condition is not a pre-computed Boolean".into(),
                                    )
                                })?
                            }
                        };
                        let steps = if taken {
                            cur_block = data.then_block;
                            pc = data.then_pc as usize;
                            u64::from(data.then_steps)
                        } else {
                            cur_block = data.else_block;
                            pc = data.else_pc as usize;
                            u64::from(data.else_steps)
                        };
                        if !self.exec.consume_execution_steps(steps) {
                            return Err(Vm3Error::BudgetExceeded);
                        }
                    }
                    VmOp::Return => {
                        if self.frames.len() == base + 1 {
                            break 'dispatch;
                        }
                        self.do_return()?;
                        continue 'dispatch;
                    }
                    VmOp::SlowInst(iref) => {
                        let switched = try_vm!(
                            self.exec_slow_inst(top, cur_prog, cur_func, cur_block, pc, iref)
                        );
                        if switched {
                            continue 'dispatch;
                        }
                        fast_frame = self.as_new_slots.is_empty()
                            && self.param_array_aliases.is_empty()
                            && self.exec.events.withevents.is_empty()
                            && self.frames[top].local_alias_mask == 0;
                    }
                    VmOp::SlowTerm(blk_idx) => {
                        let fr = &mut self.frames[top];
                        fr.block = BlockId(cur_block as usize);
                        fr.pc = pc;
                        fr.at_block_start = false;
                        if self.exec_slow_term(base, top, cur_prog, cur_func, blk_idx)? {
                            break 'dispatch;
                        }
                        continue 'dispatch;
                    }
                }
            }
        }
        Ok(())
    }

    #[inline]
    fn resolve_packed(&self, top: usize, cur_prog: usize, p: PackedPlace) -> Loc {
        match p.tag() {
            TAG_TEMP => Loc::Temp(top, p.index() - self.frames[top].num_locals),
            TAG_LOCAL => self.frames[top]
                .alias_of(p.index())
                .unwrap_or(Loc::Local(top, p.index())),
            _ => Loc::Global(cur_prog, p.index()),
        }
    }

    #[inline]
    fn read_loc_ref_fast(&self, loc: Loc) -> Option<&Variant> {
        match loc {
            Loc::Global(p, g) => self.exec.programs[p].globals.get(g),
            Loc::Local(fi, li) => self.frames.get(fi).and_then(|f| f.local(li)),
            Loc::Temp(fi, ti) => self.frames.get(fi).and_then(|f| f.temp(ti)),
        }
    }

    #[inline]
    fn read_loc_mut_fast(&mut self, loc: Loc) -> Option<&mut Variant> {
        match loc {
            Loc::Global(p, g) => self.exec.programs[p].globals.get_mut(g),
            Loc::Local(fi, li) => self.frames.get_mut(fi).and_then(|f| f.local_mut(li)),
            Loc::Temp(fi, ti) => self.frames.get_mut(fi).and_then(|f| f.temp_mut(ti)),
        }
    }

    #[inline]
    fn borrow_packed_op<'a>(
        &'a self,
        top: usize,
        cur_prog: usize,
        code: &'a CompiledFunc,
        op: PackedOp,
    ) -> Option<&'a Variant> {
        let slot = match op.tag() {
            TAG_TEMP => &self.frames[top].slots[op.index()],
            TAG_LOCAL => {
                let fr = &self.frames[top];
                if let Some(loc) = fr.alias_of(op.index()) {
                    self.read_loc_ref_fast(loc)?
                } else {
                    &fr.slots[op.index()]
                }
            }
            TAG_GLOBAL => &self.exec.programs[cur_prog].globals[op.index()],
            _ => {
                return match &code.consts[op.index()] {
                    FastConst::Scalar(v) => Some(v),
                    FastConst::Str(_) | FastConst::NullStr => None,
                };
            }
        };
        if slot.byref_target().is_some() {
            return None;
        }
        if !self.as_new_slots.is_empty() && is_possibly_nothing_for_as_new(slot) {
            return None;
        }
        Some(slot)
    }

    #[inline]
    fn borrow_packed_str_units<'a>(
        &'a self,
        top: usize,
        cur_prog: usize,
        code: &'a CompiledFunc,
        op: PackedOp,
    ) -> Option<&'a [u16]> {
        if op.tag() == TAG_CONST {
            return match &code.consts[op.index()] {
                FastConst::Str(units) => Some(units),
                FastConst::NullStr => Some(&[]),
                FastConst::Scalar(v) if v.raw_vt() == 0 => Some(&[]),
                FastConst::Scalar(_) => None,
            };
        }
        let v = self.borrow_packed_op(top, cur_prog, code, op)?;
        match v.raw_vt() {
            8 if v.string_byte_len().is_some_and(|n| n % 2 == 0) => v.borrowed_string_units(),
            0 => Some(&[]),
            _ => None,
        }
    }

    #[inline]
    fn eval_packed_op(
        &mut self,
        top: usize,
        cur_prog: usize,
        code: &CompiledFunc,
        op: PackedOp,
    ) -> Result<Variant, Vm3Error> {
        if op.tag() == TAG_CONST {
            return Ok(match &code.consts[op.index()] {
                FastConst::Scalar(v) => v.clone(),
                FastConst::Str(units) => Variant::from_utf16_units(units),
                FastConst::NullStr => Variant::null_string(),
            });
        }
        if let Some(v) = self.borrow_packed_op(top, cur_prog, code, op) {
            return Ok(v.clone());
        }
        let place = op.as_place().expect("non-const PackedOp");
        let loc = self.resolve_packed(top, cur_prog, place);
        self.read_loc_as_new(loc)
    }

    #[inline]
    fn eval_packed_place(
        &mut self,
        top: usize,
        cur_prog: usize,
        place: PackedPlace,
    ) -> Result<Variant, Vm3Error> {
        let loc = self.resolve_packed(top, cur_prog, place);
        self.read_loc_as_new(loc)
    }

    #[inline]
    fn eval_packed_arith_operand(
        &mut self,
        top: usize,
        cur_prog: usize,
        code: &CompiledFunc,
        op: PackedOp,
    ) -> Result<Variant, Vm3Error> {
        let v = self.eval_packed_op(top, cur_prog, code, op)?;
        if v.raw_vt() == 9 || v.raw_vt() == 13 {
            self.arithmetic_operand(v)
        } else {
            Ok(v)
        }
    }

    #[inline]
    fn eval_packed_arith_value(
        &mut self,
        top: usize,
        cur_prog: usize,
        code: &CompiledFunc,
        op: PackedOp,
    ) -> Result<Variant, Vm3Error> {
        let v = self.eval_packed_op(top, cur_prog, code, op)?;
        if v.raw_vt() == 9 || v.raw_vt() == 13 {
            self.arithmetic_value(v)
        } else {
            Ok(v)
        }
    }

    #[inline]
    fn eval_fast_subscript(
        &mut self,
        top: usize,
        cur_prog: usize,
        code: &CompiledFunc,
        op: PackedOp,
    ) -> Result<i32, Vm3Error> {
        if let Some(n) = self
            .borrow_packed_op(top, cur_prog, code, op)
            .and_then(try_fast_i32)
        {
            return Ok(n);
        }
        let v = self.eval_packed_op(top, cur_prog, code, op)?;
        array_subscript(&v).map_err(arith_fault)
    }

    #[inline]
    fn fast_dst_slot_mut(
        &mut self,
        top: usize,
        cur_prog: usize,
        dst: PackedPlace,
    ) -> Option<&mut Variant> {
        if !self.param_array_aliases.is_empty() {
            return None;
        }
        let has_withevents = !self.exec.events.withevents.is_empty();
        let slot = match dst.tag() {
            TAG_TEMP => {
                let fr = &mut self.frames[top];
                let idx = dst.index();
                if idx + 1 > fr.temp_hwm {
                    fr.temp_hwm = idx + 1;
                }
                &mut fr.slots[idx]
            }
            TAG_LOCAL => {
                let idx = dst.index();
                let alias = {
                    let fr = &self.frames[top];
                    if fr.local_alias_mask == 0
                        || (idx < 64 && (fr.local_alias_mask & (1u64 << idx)) == 0)
                    {
                        None
                    } else {
                        fr.alias_of(idx)
                    }
                };
                if let Some(loc) = alias {
                    self.read_loc_mut_fast(loc)?
                } else {
                    &mut self.frames[top].slots[idx]
                }
            }
            _ => &mut self.exec.programs[cur_prog].globals[dst.index()],
        };
        let old_vt = slot.raw_vt();
        if old_vt < 8 || old_vt == 11 {
            return Some(slot);
        }
        let hi = old_vt & 0xF000;
        if hi == 0x2000 || hi == 0x4000 || hi == 0x6000 {
            return None;
        }
        if has_withevents && (old_vt == 9 || old_vt == 13) {
            return None;
        }
        Some(slot)
    }

    /// Borrow a source storage cell (`src_loc`) immutably and a distinct fast-writable
    /// destination slot (`dst`) mutably at the same time, enabling in-place reads from
    /// 1D arrays, UDT array fields, and project object fields with zero intermediate `Variant`s.
    #[inline]
    fn borrow_loc_and_fast_dst_mut(
        &mut self,
        top: usize,
        cur_prog: usize,
        src_loc: Loc,
        dst: PackedPlace,
    ) -> Option<(&Variant, &mut Variant)> {
        if !self.param_array_aliases.is_empty() {
            return None;
        }
        let has_withevents = !self.exec.events.withevents.is_empty();
        let (dst_tag, dst_idx) = (dst.tag(), dst.index());
        // Resolve dst into a top-frame flat slot index (handles non-aliased locals and all temps).
        let dst_top_slot = match dst_tag {
            TAG_TEMP => Some(dst_idx),
            TAG_LOCAL => {
                let fr = &self.frames[top];
                if fr.local_alias_mask == 0
                    || (dst_idx < 64 && (fr.local_alias_mask & (1u64 << dst_idx)) == 0)
                {
                    Some(dst_idx)
                } else {
                    None
                }
            }
            _ => None,
        };
        let (src_ref, dst_slot): (&Variant, &mut Variant) = match (src_loc, dst_top_slot) {
            (Loc::Local(fi, li), Some(di)) if fi == top => {
                if li == di {
                    return None;
                }
                let fr = &mut self.frames[top];
                if dst_tag == TAG_TEMP && di + 1 > fr.temp_hwm {
                    fr.temp_hwm = di + 1;
                }
                let slots = fr.slots.as_mut_slice();
                if li < di {
                    let (left, right) = slots.split_at_mut(di);
                    (&left[li], &mut right[0])
                } else {
                    let (left, right) = slots.split_at_mut(li);
                    (&right[0], &mut left[di])
                }
            }
            (Loc::Temp(fi, ti), Some(di)) if fi == top => {
                let si = self.frames[top].num_locals + ti;
                if si == di {
                    return None;
                }
                let fr = &mut self.frames[top];
                if dst_tag == TAG_TEMP && di + 1 > fr.temp_hwm {
                    fr.temp_hwm = di + 1;
                }
                let slots = fr.slots.as_mut_slice();
                if si < di {
                    let (left, right) = slots.split_at_mut(di);
                    (&left[si], &mut right[0])
                } else {
                    let (left, right) = slots.split_at_mut(si);
                    (&right[0], &mut left[di])
                }
            }
            (Loc::Local(fi, li), Some(di)) if fi < top => {
                let (lower, upper) = self.frames.split_at_mut(top);
                let src = lower[fi].local(li)?;
                let fr = &mut upper[0];
                if dst_tag == TAG_TEMP && di + 1 > fr.temp_hwm {
                    fr.temp_hwm = di + 1;
                }
                (src, &mut fr.slots[di])
            }
            (Loc::Temp(fi, ti), Some(di)) if fi < top => {
                let (lower, upper) = self.frames.split_at_mut(top);
                let src = lower[fi].temp(ti)?;
                let fr = &mut upper[0];
                if dst_tag == TAG_TEMP && di + 1 > fr.temp_hwm {
                    fr.temp_hwm = di + 1;
                }
                (src, &mut fr.slots[di])
            }
            (Loc::Global(p, g), Some(di)) => {
                let src = self.exec.programs[p].globals.get(g)?;
                let fr = &mut self.frames[top];
                if dst_tag == TAG_TEMP && di + 1 > fr.temp_hwm {
                    fr.temp_hwm = di + 1;
                }
                (src, &mut fr.slots[di])
            }
            (Loc::Local(fi, li), None) if dst_tag == TAG_GLOBAL => {
                let src = self.frames.get(fi)?.local(li)?;
                let dst_slot = self.exec.programs[cur_prog].globals.get_mut(dst_idx)?;
                (src, dst_slot)
            }
            (Loc::Temp(fi, ti), None) if dst_tag == TAG_GLOBAL => {
                let src = self.frames.get(fi)?.temp(ti)?;
                let dst_slot = self.exec.programs[cur_prog].globals.get_mut(dst_idx)?;
                (src, dst_slot)
            }
            _ => return None,
        };
        let old_vt = dst_slot.raw_vt();
        if old_vt < 8 || old_vt == 11 {
            return Some((src_ref, dst_slot));
        }
        let hi = old_vt & 0xF000;
        if hi == 0x2000 || hi == 0x4000 || hi == 0x6000 {
            return None;
        }
        if has_withevents && (old_vt == 9 || old_vt == 13) {
            return None;
        }
        Some((src_ref, dst_slot))
    }

    #[inline]
    fn write_packed_date_f64(
        &mut self,
        top: usize,
        cur_prog: usize,
        dst: PackedPlace,
        val: f64,
    ) -> Result<(), Vm3Error> {
        if let Some(slot) = self.fast_dst_slot_mut(top, cur_prog, dst) {
            slot.set_date_f64(val);
            return Ok(());
        }
        let loc = self.resolve_packed(top, cur_prog, dst);
        self.write_loc(loc, Variant::from_date_f64(val))
    }

    #[inline]
    fn write_packed_i32(
        &mut self,
        top: usize,
        cur_prog: usize,
        dst: PackedPlace,
        val: i32,
    ) -> Result<(), Vm3Error> {
        if let Some(slot) = self.fast_dst_slot_mut(top, cur_prog, dst) {
            slot.set_i32(val);
            return Ok(());
        }
        let loc = self.resolve_packed(top, cur_prog, dst);
        self.write_loc(loc, Variant::from_i32(val))
    }

    #[inline]
    fn write_packed_f64(
        &mut self,
        top: usize,
        cur_prog: usize,
        dst: PackedPlace,
        val: f64,
    ) -> Result<(), Vm3Error> {
        if let Some(slot) = self.fast_dst_slot_mut(top, cur_prog, dst) {
            slot.set_f64(val);
            return Ok(());
        }
        let loc = self.resolve_packed(top, cur_prog, dst);
        self.write_loc(loc, Variant::from_f64(val))
    }

    #[inline]
    fn write_packed_bool(
        &mut self,
        top: usize,
        cur_prog: usize,
        dst: PackedPlace,
        val: bool,
    ) -> Result<(), Vm3Error> {
        if let Some(slot) = self.fast_dst_slot_mut(top, cur_prog, dst) {
            slot.set_bool(val);
            return Ok(());
        }
        let loc = self.resolve_packed(top, cur_prog, dst);
        self.write_loc(loc, Variant::from_bool(val))
    }

    #[inline]
    fn write_packed_place(
        &mut self,
        top: usize,
        cur_prog: usize,
        dst: PackedPlace,
        val: Variant,
    ) -> Result<(), Vm3Error> {
        if let Some(slot) = self.fast_dst_slot_mut(top, cur_prog, dst) {
            *slot = val;
            return Ok(());
        }
        let loc = self.resolve_packed(top, cur_prog, dst);
        self.write_loc(loc, val)
    }

    #[inline]
    fn write_packed_assigned(
        &mut self,
        top: usize,
        cur_prog: usize,
        dst: PackedPlace,
        mut val: Variant,
    ) -> Result<(), Vm3Error> {
        if (val.raw_vt() & 0xF000) == 0x2000 && dst.tag() != TAG_TEMP {
            let loc = self.resolve_packed(top, cur_prog, dst);
            if !matches!(loc, Loc::Temp(..))
                && let Some(target) = self.read_loc_ref_fast(loc)
            {
                val.adopt_assignment_target_fixedness(target);
            }
            return self.write_loc(loc, val);
        }
        self.write_packed_place(top, cur_prog, dst, val)
    }

    #[inline]
    fn store_packed_arith(
        &mut self,
        top: usize,
        cur_prog: usize,
        dst: PackedPlace,
        out: Result<Variant, ArithError>,
    ) -> Result<(), Vm3Error> {
        let v = out.map_err(|e| Vm3Error::Fault(Fault::from_arith(e)))?;
        let non_finite_double = v.raw_vt() == 5 && v.as_f64().is_some_and(|f| !f.is_finite());
        self.write_packed_place(top, cur_prog, dst, v)?;
        if non_finite_double {
            return Err(Vm3Error::Fault(Fault::from_arith(ArithError::overflow())));
        }
        Ok(())
    }

    #[cold]
    #[inline(never)]
    fn exec_slow_inst(
        &mut self,
        top: usize,
        cur_prog: usize,
        cur_func: FuncId,
        cur_block: u32,
        pc: usize,
        iref: InstRef,
    ) -> Result<bool, Vm3Error> {
        let fr = &mut self.frames[top];
        fr.block = BlockId(cur_block as usize);
        fr.pc = pc;
        fr.at_block_start = false;
        let inst = &self.exec.programs[cur_prog].program.funcs[cur_func.0].blocks
            [iref.block as usize]
            .instrs[iref.inst as usize];
        self.exec(inst)?;
        Ok(self.frames.len() - 1 != top)
    }

    #[cold]
    #[inline(never)]
    #[allow(clippy::too_many_arguments)]
    fn exec_slow_inst_range(
        &mut self,
        top: usize,
        cur_prog: usize,
        cur_func: FuncId,
        cur_block: u32,
        pc: usize,
        first_inst: InstRef,
        count: usize,
    ) -> Result<bool, Vm3Error> {
        let fr = &mut self.frames[top];
        fr.block = BlockId(cur_block as usize);
        fr.pc = pc;
        fr.at_block_start = false;
        let instrs = &self.exec.programs[cur_prog].program.funcs[cur_func.0].blocks
            [first_inst.block as usize]
            .instrs;
        let start = first_inst.inst as usize;
        for inst in &instrs[start..start + count] {
            self.exec(inst)?;
        }
        Ok(self.frames.len() - 1 != top)
    }

    #[cold]
    #[inline(never)]
    fn exec_slow_term(
        &mut self,
        base: usize,
        top: usize,
        cur_prog: usize,
        cur_func: FuncId,
        blk_idx: u32,
    ) -> Result<bool, Vm3Error> {
        let term = &self.exec.programs[cur_prog].program.funcs[cur_func.0].blocks[blk_idx as usize]
            .terminator;
        match term {
            OxTerminator::Jump(b) => self.goto(top, *b),
            OxTerminator::Branch {
                cond,
                then_blk,
                else_blk,
            } => {
                let v = self.operand(cond)?;
                let taken = v.as_bool().ok_or_else(|| {
                    Vm3Error::Malformed("Branch condition is not a pre-computed Boolean".into())
                })?;
                self.goto(top, if taken { *then_blk } else { *else_blk });
            }
            OxTerminator::Return => {
                if self.frames.len() == base + 1 {
                    return Ok(true);
                }
                self.do_return()?;
            }
            OxTerminator::Halt => {
                if self.run_depth > 1 {
                    return Err(Vm3Error::Halted);
                }
                self.halt_to(base);
                return Ok(true);
            }
            OxTerminator::FaultDispatch {
                resume,
                resume_next,
            } => match self
                .exec
                .err_engine
                .dispatch_fault(*resume, *resume_next, self.frames[top].current_line)
                .map_err(|msg| Vm3Error::Malformed(msg.into()))?
            {
                FaultAction::Propagate(fault) => self.propagate_fault(fault, base)?,
                FaultAction::ResumeNext(block) | FaultAction::Handle(block) => {
                    self.goto(top, block);
                }
            },
            OxTerminator::Resume => match self.exec.err_engine.resume(ResumeTarget::Same) {
                Ok(block) => self.goto(top, block),
                Err(fault) => self.route_fault(fault)?,
            },
            OxTerminator::ResumeNext => match self.exec.err_engine.resume(ResumeTarget::Next) {
                Ok(block) => self.goto(top, block),
                Err(fault) => self.route_fault(fault)?,
            },
            OxTerminator::ResumeLabel(b) => {
                match self.exec.err_engine.resume(ResumeTarget::Label(*b)) {
                    Ok(block) => self.goto(top, block),
                    Err(fault) => self.route_fault(fault)?,
                }
            }
            OxTerminator::Raise {
                number,
                source,
                description,
                help_file,
                help_context,
                inherit,
            } => {
                let number = self.operand(number)?;
                let mut value = |operand: &Option<OxOperand>| {
                    operand.as_ref().map(|op| self.operand(op)).transpose()
                };
                let args = [
                    Some(number),
                    value(source)?,
                    value(description)?,
                    value(help_file)?,
                    value(help_context)?,
                ];
                let fault = self
                    .exec
                    .err_engine
                    .fault_from_raise_args(args.each_ref().map(Option::as_ref), *inherit);
                self.route_fault(fault)?;
            }
            OxTerminator::GoSub { target, ret } => {
                self.frames[top].gosub_stack.push(*ret);
                self.goto(top, *target);
            }
            OxTerminator::GoSubReturn => match self.frames[top].gosub_stack.pop() {
                Some(ret) => self.goto(top, ret),
                None => self.raise_runtime_error(3)?,
            },
            OxTerminator::Unreachable => {
                return Err(Vm3Error::Malformed(
                    "reached an Unreachable terminator".into(),
                ));
            }
        }
        Ok(false)
    }

    fn exec_call_proc_fast(
        &mut self,
        top: usize,
        cur_prog: usize,
        code: &CompiledFunc,
        call_idx: u32,
    ) -> Result<bool, Vm3Error> {
        let call = &code.proc_calls[call_idx as usize];
        let program = self.exec.programs[cur_prog].program;
        let callee = program
            .funcs
            .get(call.proc.0)
            .ok_or_else(|| Vm3Error::Malformed(format!("call to unknown proc {}", call.proc.0)))?;
        let dst_loc = call.dst.map(|p| self.resolve_packed(top, cur_prog, p));
        if self.has_typed_entries
            && call.all_byval
            && self.cache.typed_entries[cur_prog][call.proc.0].is_some()
            && call.args.len() == callee.param_count
        {
            self.guard_call_depth()?;
            let mut values =
                std::array::from_fn::<_, { calcvba_rt_abi::host_entry::MAX_TYPED_ARGS }, _>(|_| {
                    Variant::empty()
                });
            for (i, arg) in call.args.iter().enumerate() {
                let FastArg::ByVal(op) = *arg else {
                    unreachable!("checked all_byval")
                };
                values[i] = self.eval_packed_op(top, cur_prog, code, op)?;
            }
            if let Some(result) =
                self.invoke_scalar_values(cur_prog, call.proc, &values[..call.args.len()])
            {
                let value = result?;
                if let Some(dst) = dst_loc {
                    self.write_loc(dst, value)?;
                }
                return Ok(false);
            }
        }
        self.guard_call_depth()?;
        let mut frame =
            self.new_frame_pooled(cur_prog, call.proc, call.args.len(), dst_loc, None)?;
        let frame_index = self.frames.len();
        let mut pending_param_array_aliases: Vec<(Loc, Vec<Option<Loc>>)> = Vec::new();
        for (i, arg) in call.args.iter().enumerate() {
            let step: Result<(), Vm3Error> = match *arg {
                FastArg::ByVal(op) => {
                    if self.param_array_aliases.is_empty()
                        && let Some(src_ref) = self.borrow_packed_op(top, cur_prog, code, op)
                        && Variant::raw_vt_is_trivial_copy(src_ref.raw_vt())
                    {
                        let core = src_ref.core();
                        if let Some(slot) = frame.local_mut(i) {
                            slot.overwrite_trivial_core(core);
                        }
                        Ok(())
                    } else {
                        let pa = if !self.param_array_aliases.is_empty()
                            && let Some(place) = op.as_place()
                        {
                            self.param_array_aliases
                                .get(&self.resolve_packed(top, cur_prog, place))
                                .cloned()
                        } else {
                            None
                        };
                        match self.eval_packed_op(top, cur_prog, code, op) {
                            Ok(v) => {
                                if let Some(slot) = frame.local_mut(i) {
                                    *slot = v;
                                }
                                if let Some(aliases) = pa {
                                    pending_param_array_aliases
                                        .push((Loc::Local(frame_index, i), aliases));
                                }
                                Ok(())
                            }
                            Err(e) => Err(e),
                        }
                    }
                }
                FastArg::ByRef(place) => {
                    let param_ty = callee.locals.get(i).map(|local| &local.ty);
                    let ox_place = place.to_ox(code.num_locals);
                    match self.byref_reference(&ox_place, param_ty) {
                        Ok(Some(reference)) => {
                            frame.slots[i] = reference;
                            Ok(())
                        }
                        Ok(None) => {
                            frame.set_alias(i, self.resolve_packed(top, cur_prog, place));
                            Ok(())
                        }
                        Err(e) => Err(e),
                    }
                }
                FastArg::Omitted => {
                    if let Some(slot) = frame.local_mut(i) {
                        match callee.locals.get(i).and_then(omitted_param_value) {
                            Some(Ok(v)) => {
                                *slot = v;
                                Ok(())
                            }
                            Some(Err(e)) => Err(e),
                            None => {
                                *slot = Variant::from_error_code(MISSING_ARG);
                                Ok(())
                            }
                        }
                    } else {
                        Ok(())
                    }
                }
            };
            if let Err(e) = step {
                self.recycle_frame(frame);
                return Err(e);
            }
        }
        let saved_err = self.exec.err_engine.enter_activation();
        frame.saved_error_mode = saved_err.error_mode;
        frame.saved_active_error = saved_err.active_error;
        self.frames.push(frame);
        for (loc, aliases) in pending_param_array_aliases {
            self.param_array_aliases.insert(loc, aliases);
        }
        Ok(true)
    }

    fn exec_call_extern_lib(
        &mut self,
        top: usize,
        cur_prog: usize,
        code: &CompiledFunc,
        call_idx: u32,
    ) -> Result<(), Vm3Error> {
        let call = &code.lib_calls[call_idx as usize];
        // Ultra-hot pure builtins inlined directly over borrowed operands (zero BSTR cloning
        // for variables and zero BSTR allocation for string literals).
        match call.id {
            NativeImplId::Sqr
            | NativeImplId::Chr
            | NativeImplId::Len
            | NativeImplId::CDbl
            | NativeImplId::CLng
            | NativeImplId::UCase
            | NativeImplId::Trim
            | NativeImplId::LTrim
            | NativeImplId::RTrim
            | NativeImplId::Asc
                if matches!(call.args.as_slice(), [FastArg::ByVal(_)]) =>
            {
                let FastArg::ByVal(arg_op) = call.args[0] else {
                    unreachable!();
                };
                match call.id {
                    NativeImplId::Sqr => {
                        if let Some(x) = self
                            .borrow_packed_op(top, cur_prog, code, arg_op)
                            .and_then(try_fast_f64)
                            && x >= 0.0
                        {
                            if let Some(dst) = call.dst {
                                self.write_packed_f64(top, cur_prog, dst, x.sqrt())?;
                            }
                            return Ok(());
                        }
                    }
                    NativeImplId::Chr => {
                        if let Some(n) = self
                            .borrow_packed_op(top, cur_prog, code, arg_op)
                            .and_then(try_fast_i32)
                            && (0..=127).contains(&n)
                        {
                            if let Some(dst) = call.dst {
                                let u = [n as u16];
                                self.write_packed_place(
                                    top,
                                    cur_prog,
                                    dst,
                                    Variant::from_utf16_units(&u),
                                )?;
                            }
                            return Ok(());
                        }
                    }
                    NativeImplId::Len => {
                        if let Some(units) = self
                            .borrow_packed_op(top, cur_prog, code, arg_op)
                            .and_then(Variant::borrowed_string_units)
                        {
                            if let Some(dst) = call.dst {
                                self.write_packed_i32(top, cur_prog, dst, units.len() as i32)?;
                            }
                            return Ok(());
                        }
                    }
                    NativeImplId::CDbl => {
                        if let Some(x) = self
                            .borrow_packed_op(top, cur_prog, code, arg_op)
                            .and_then(try_fast_f64)
                            && x.is_finite()
                        {
                            if let Some(dst) = call.dst {
                                self.write_packed_f64(top, cur_prog, dst, x)?;
                            }
                            return Ok(());
                        }
                    }
                    NativeImplId::CLng => {
                        if let Some(n) = self
                            .borrow_packed_op(top, cur_prog, code, arg_op)
                            .and_then(try_fast_i32)
                        {
                            if let Some(dst) = call.dst {
                                self.write_packed_i32(top, cur_prog, dst, n)?;
                            }
                            return Ok(());
                        }
                    }
                    NativeImplId::UCase => {
                        if let Some(units) =
                            self.borrow_packed_str_units(top, cur_prog, code, arg_op)
                        {
                            let res = calcvba_lib::ucase_units(units);
                            if let Some(dst) = call.dst {
                                self.write_packed_place(top, cur_prog, dst, res)?;
                            }
                            return Ok(());
                        }
                    }
                    NativeImplId::Trim | NativeImplId::LTrim | NativeImplId::RTrim => {
                        if let Some(units) =
                            self.borrow_packed_str_units(top, cur_prog, code, arg_op)
                        {
                            let (left, right) = match call.id {
                                NativeImplId::Trim => (true, true),
                                NativeImplId::LTrim => (true, false),
                                _ => (false, true),
                            };
                            let res = calcvba_lib::trim_units(units, left, right);
                            if let Some(dst) = call.dst {
                                self.write_packed_place(top, cur_prog, dst, res)?;
                            }
                            return Ok(());
                        }
                    }
                    NativeImplId::Asc => {
                        if let Some(units) =
                            self.borrow_packed_str_units(top, cur_prog, code, arg_op)
                            && let Some(&first) = units.first()
                            && first < 0x80
                        {
                            if let Some(dst) = call.dst {
                                self.write_packed_place(
                                    top,
                                    cur_prog,
                                    dst,
                                    Variant::from_i16(first as i16),
                                )?;
                            }
                            return Ok(());
                        }
                    }
                    _ => {}
                }
            }
            NativeImplId::Mid if matches!(call.args.len(), 2 | 3) => {
                if let FastArg::ByVal(s_op) = call.args[0]
                    && let FastArg::ByVal(start_op) = call.args[1]
                    && let Some(units) = self.borrow_packed_str_units(top, cur_prog, code, s_op)
                    && let Some(start) = self
                        .borrow_packed_op(top, cur_prog, code, start_op)
                        .and_then(try_fast_i32)
                {
                    let take = match call.args.get(2) {
                        None | Some(FastArg::Omitted) => Some(None),
                        Some(FastArg::ByVal(len_op)) => self
                            .borrow_packed_op(top, cur_prog, code, *len_op)
                            .and_then(try_fast_i32)
                            .map(Some),
                        Some(FastArg::ByRef(_)) => None,
                    };
                    if let Some(take_opt) = take {
                        let res = calcvba_lib::mid_units(units, start, take_opt)
                            .map_err(|e| Vm3Error::Fault(Fault::from_lib(e)))?;
                        if let Some(dst) = call.dst {
                            self.write_packed_place(top, cur_prog, dst, res)?;
                        }
                        return Ok(());
                    }
                }
            }
            NativeImplId::InStr if matches!(call.args.len(), 2..=4) => {
                let has_start = call.args.len() >= 3;
                let base = usize::from(has_start);
                let start_ok = if has_start {
                    match call.args[0] {
                        FastArg::ByVal(op) => self
                            .borrow_packed_op(top, cur_prog, code, op)
                            .and_then(try_fast_i32),
                        _ => None,
                    }
                } else {
                    Some(1)
                };
                let binary_ok = match call.args.get(base + 2) {
                    None | Some(FastArg::Omitted) => true,
                    Some(FastArg::ByVal(op)) => {
                        self.borrow_packed_op(top, cur_prog, code, *op)
                            .and_then(try_fast_i32)
                            == Some(0)
                    }
                    Some(FastArg::ByRef(_)) => false,
                };
                if binary_ok
                    && let Some(start) = start_ok
                    && let FastArg::ByVal(hay_op) = call.args[base]
                    && let FastArg::ByVal(ndl_op) = call.args[base + 1]
                    && let (Some(hay), Some(needle)) = (
                        self.borrow_packed_str_units(top, cur_prog, code, hay_op),
                        self.borrow_packed_str_units(top, cur_prog, code, ndl_op),
                    )
                {
                    let pos = calcvba_lib::instr_binary_units(start, hay, needle)
                        .map_err(|e| Vm3Error::Fault(Fault::from_lib(e)))?;
                    if let Some(dst) = call.dst {
                        self.write_packed_i32(top, cur_prog, dst, pos)?;
                    }
                    return Ok(());
                }
            }
            NativeImplId::Replace if matches!(call.args.len(), 3..=6) => {
                let start_ok = match call.args.get(3) {
                    None | Some(FastArg::Omitted) => Some(1),
                    Some(FastArg::ByVal(op)) => self
                        .borrow_packed_op(top, cur_prog, code, *op)
                        .and_then(try_fast_i32),
                    Some(FastArg::ByRef(_)) => None,
                };
                let count_ok = match call.args.get(4) {
                    None | Some(FastArg::Omitted) => Some(-1),
                    Some(FastArg::ByVal(op)) => self
                        .borrow_packed_op(top, cur_prog, code, *op)
                        .and_then(try_fast_i32),
                    Some(FastArg::ByRef(_)) => None,
                };
                let binary_ok = match call.args.get(5) {
                    None | Some(FastArg::Omitted) => true,
                    Some(FastArg::ByVal(op)) => {
                        self.borrow_packed_op(top, cur_prog, code, *op)
                            .and_then(try_fast_i32)
                            == Some(0)
                    }
                    Some(FastArg::ByRef(_)) => false,
                };
                if binary_ok
                    && let (Some(start), Some(count)) = (start_ok, count_ok)
                    && let FastArg::ByVal(s_op) = call.args[0]
                    && let FastArg::ByVal(f_op) = call.args[1]
                    && let FastArg::ByVal(w_op) = call.args[2]
                    && let (Some(s), Some(find), Some(with)) = (
                        self.borrow_packed_str_units(top, cur_prog, code, s_op),
                        self.borrow_packed_str_units(top, cur_prog, code, f_op),
                        self.borrow_packed_str_units(top, cur_prog, code, w_op),
                    )
                {
                    let res = calcvba_lib::replace_binary_units(s, find, with, start, count)
                        .map_err(|e| Vm3Error::Fault(Fault::from_lib(e)))?;
                    if let Some(dst) = call.dst {
                        self.write_packed_place(top, cur_prog, dst, res)?;
                    }
                    return Ok(());
                }
            }
            NativeImplId::Round if matches!(call.args.len(), 1 | 2) => {
                if let FastArg::ByVal(v_op) = call.args[0]
                    && let Some(v_ref) = self.borrow_packed_op(top, cur_prog, code, v_op)
                    && v_ref.raw_vt() == 5
                    && let Some(x) = v_ref.as_f64()
                {
                    let digits_opt = match call.args.get(1) {
                        None | Some(FastArg::Omitted) => Some(0),
                        Some(FastArg::ByVal(d_op)) => self
                            .borrow_packed_op(top, cur_prog, code, *d_op)
                            .and_then(try_fast_i32),
                        Some(FastArg::ByRef(_)) => None,
                    };
                    if let Some(digits) = digits_opt {
                        let r = calcvba_lib::round_f64_digits(x, digits)
                            .map_err(|e| Vm3Error::Fault(Fault::from_lib(e)))?;
                        if let Some(dst) = call.dst {
                            self.write_packed_f64(top, cur_prog, dst, r)?;
                        }
                        return Ok(());
                    }
                }
            }
            NativeImplId::DateAdd if call.args.len() == 3 && self.exec.lib.calendar == 0 => {
                if let FastArg::ByVal(i_op) = call.args[0]
                    && let FastArg::ByVal(n_op) = call.args[1]
                    && let FastArg::ByVal(d_op) = call.args[2]
                    && let Some(interval_units) =
                        self.borrow_packed_str_units(top, cur_prog, code, i_op)
                    && let Some(num) = self
                        .borrow_packed_op(top, cur_prog, code, n_op)
                        .and_then(try_fast_f64)
                    && let Some(d_ref) = self.borrow_packed_op(top, cur_prog, code, d_op)
                    && d_ref.raw_vt() == 7
                    && let Some(serial) = d_ref.as_date_f64()
                {
                    let out = calcvba_lib::date_add_gregorian_units(interval_units, num, serial)
                        .map_err(|e| Vm3Error::Fault(Fault::from_lib(e)))?;
                    if let Some(dst) = call.dst {
                        self.write_packed_date_f64(top, cur_prog, dst, out)?;
                    }
                    return Ok(());
                }
            }
            _ => {}
        }

        let mut stack_buf: StackVariantBuf<8> = StackVariantBuf::new();
        let mut heap_buf: Vec<Variant> = Vec::new();
        let mut any_object = false;
        let argc = call.args.len();
        let argv: &mut [Variant] = if argc <= 8 {
            for arg in &call.args {
                let v = match *arg {
                    FastArg::ByVal(op) => self.eval_packed_op(top, cur_prog, code, op)?,
                    FastArg::ByRef(place) => self.eval_packed_place(top, cur_prog, place)?,
                    FastArg::Omitted => Variant::empty(),
                };
                if v.raw_vt() == 9 || v.raw_vt() == 13 {
                    any_object = true;
                }
                stack_buf.push(v);
            }
            stack_buf.as_mut_slice()
        } else {
            heap_buf.reserve_exact(argc);
            for arg in &call.args {
                let v = match *arg {
                    FastArg::ByVal(op) => self.eval_packed_op(top, cur_prog, code, op)?,
                    FastArg::ByRef(place) => self.eval_packed_place(top, cur_prog, place)?,
                    FastArg::Omitted => Variant::empty(),
                };
                if v.raw_vt() == 9 || v.raw_vt() == 13 {
                    any_object = true;
                }
                heap_buf.push(v);
            }
            heap_buf.as_mut_slice()
        };

        let result = if any_object && let Some(value) = self.vba_object_len(call.id, argv)? {
            value
        } else {
            if any_object {
                self.native_object_values(call.id, argv)?;
            }
            if call.pure_no_events {
                if call.string_typed_alias && call.id != NativeImplId::Format {
                    calcvba_lib::validate_string_alias_args(call.id, argv)
                        .map_err(|err| Vm3Error::Fault(Fault::from_lib(err)))?;
                }
                match calcvba_lib::invoke_context_free(call.id, argv, self.exec.host) {
                    Ok(v) => v,
                    Err(calcvba_lib::ContextFreeInvokeError::Library(err)) => {
                        return Err(Vm3Error::Fault(Fault::from_lib(err)));
                    }
                    Err(calcvba_lib::ContextFreeInvokeError::ContextRequired) => {
                        return Err(Vm3Error::Malformed(format!(
                            "contextual built-in {:?} escaped VM3 dispatch classification",
                            call.id
                        )));
                    }
                }
            } else {
                self.invoke_native_lib_with_policy(call.id, argv, call.string_typed_alias)?
            }
        };
        if let Some(dst) = call.dst {
            self.write_packed_place(top, cur_prog, dst, result)?;
        }
        Ok(())
    }

    fn exec_call_com_early(
        &mut self,
        top: usize,
        cur_prog: usize,
        cur_func: FuncId,
        code: &CompiledFunc,
        call_idx: u32,
    ) -> Result<bool, Vm3Error> {
        let call = &code.com_early_calls[call_idx as usize];
        let borrowed_info = self
            .borrow_packed_op(top, cur_prog, code, call.recv)
            .and_then(borrow_project_instance_info);
        if let Some((obj_bundle, route_key, descriptor)) = borrowed_info {
            if route_key == VBA_COLLECTION_ROUTE_KEY
                && let Some(pargs) = call.positional_args.as_deref()
                && pargs.len() <= 4
                && let Some(meta) = self.exec.programs[cur_prog].program.com_method(call.method)
                && let Some(method) = Self::vba_collection_fast_method(&meta.name)
            {
                return self
                    .exec_collection_fast_borrowed(
                        top, cur_prog, code, call.recv, call.dst, method, pargs,
                    )
                    .map(|()| false);
            }
            if route_key >= 0
                && descriptor.project_identity.is_some()
                && let Some(pargs) = call.positional_args.as_deref()
            {
                let desc_ptr = descriptor as *const RuntimeClassDescriptor as usize;
                let member = if call.ic_descriptor.get() == desc_ptr
                    && let Some(m) = call.ic_member.get()
                {
                    m
                } else {
                    let meta = self.exec.programs[cur_prog]
                        .program
                        .com_method(call.method)
                        .ok_or_else(|| {
                            Vm3Error::Malformed(format!(
                                "ComCallEarly: unresolved method ref {:?}",
                                call.method
                            ))
                        })?;
                    let runtime_kind =
                        calcvba_rt_abi::runtime_invoke_kind_from_typelib(call.invoke_kind);
                    let m = calcvba_rt_abi::resolve_project_member_for_call(
                        &mut self.exec.dispatch_plan_cache,
                        descriptor,
                        Some(meta),
                        &meta.name,
                        runtime_kind,
                        pargs.is_empty(),
                        obj_bundle != cur_prog,
                    )
                    .map_err(Vm3Error::Fault)?;
                    call.ic_descriptor.set(desc_ptr);
                    call.ic_member.set(Some(m));
                    m
                };
                if let calcvba_runtime::RuntimeMemberBacking::Proc(proc_idx) = member.backing
                    && let Some(pushed) = self.try_push_project_proc_fast(
                        top,
                        cur_prog,
                        code,
                        call.recv,
                        call.dst,
                        obj_bundle,
                        FuncId(proc_idx),
                        pargs,
                    )?
                {
                    return Ok(pushed);
                }
            }
        }
        let inst = &self.exec.programs[cur_prog].program.funcs[cur_func.0].blocks
            [call.inst.block as usize]
            .instrs[call.inst.inst as usize];
        self.exec(inst)?;
        Ok(self.frames.len() - 1 != top)
    }

    fn exec_call_com_late(
        &mut self,
        top: usize,
        cur_prog: usize,
        cur_func: FuncId,
        code: &CompiledFunc,
        call_idx: u32,
    ) -> Result<bool, Vm3Error> {
        let call = &code.com_late_calls[call_idx as usize];
        let borrowed_info = self
            .borrow_packed_op(top, cur_prog, code, call.recv)
            .and_then(borrow_project_instance_info);
        if let Some((obj_bundle, route_key, descriptor)) = borrowed_info {
            if route_key == VBA_COLLECTION_ROUTE_KEY
                && let Some(pargs) = call.positional_args.as_deref()
                && pargs.len() <= 4
            {
                let method = if call.default_member {
                    (!pargs.is_empty()).then_some(CollectionMethod::Item)
                } else {
                    Self::vba_collection_fast_method(&call.name)
                };
                if let Some(method) = method {
                    return self
                        .exec_collection_fast_borrowed(
                            top, cur_prog, code, call.recv, call.dst, method, pargs,
                        )
                        .map(|()| false);
                }
            }
            if route_key >= 0
                && descriptor.project_identity.is_some()
                && let Some(pargs) = call.positional_args.as_deref()
            {
                let desc_ptr = descriptor as *const RuntimeClassDescriptor as usize;
                let member = if call.ic_descriptor.get() == desc_ptr
                    && let Some(m) = call.ic_member.get()
                {
                    m
                } else {
                    let runtime_kind =
                        calcvba_rt_abi::runtime_invoke_kind_from_typelib(call.invoke_kind);
                    let m = if call.default_member {
                        resolve_project_default_member(
                            descriptor,
                            runtime_kind,
                            pargs.is_empty(),
                            obj_bundle != cur_prog,
                        )
                        .map_err(Vm3Error::Fault)?
                    } else {
                        let effective_public_only = call.public_only || obj_bundle != cur_prog;
                        calcvba_rt_abi::resolve_project_member_for_call(
                            &mut self.exec.dispatch_plan_cache,
                            descriptor,
                            None,
                            &call.name,
                            runtime_kind,
                            pargs.is_empty(),
                            effective_public_only,
                        )
                        .map_err(Vm3Error::Fault)?
                    };
                    call.ic_descriptor.set(desc_ptr);
                    call.ic_member.set(Some(m));
                    m
                };
                if let calcvba_runtime::RuntimeMemberBacking::Proc(proc_idx) = member.backing
                    && let Some(pushed) = self.try_push_project_proc_fast(
                        top,
                        cur_prog,
                        code,
                        call.recv,
                        call.dst,
                        obj_bundle,
                        FuncId(proc_idx),
                        pargs,
                    )?
                {
                    return Ok(pushed);
                }
            }
        }
        let inst = &self.exec.programs[cur_prog].program.funcs[cur_func.0].blocks
            [call.inst.block as usize]
            .instrs[call.inst.inst as usize];
        self.exec(inst)?;
        Ok(self.frames.len() - 1 != top)
    }

    #[inline]
    fn vba_collection_fast_method(name: &str) -> Option<CollectionMethod> {
        let member = name.trim().trim_start_matches('[').trim_end_matches(']');
        if member.eq_ignore_ascii_case("Add") {
            Some(CollectionMethod::Add)
        } else if member.eq_ignore_ascii_case("Item") {
            Some(CollectionMethod::Item)
        } else if member.eq_ignore_ascii_case("Count") {
            Some(CollectionMethod::Count)
        } else if member.eq_ignore_ascii_case("Remove") {
            Some(CollectionMethod::Remove)
        } else {
            None
        }
    }

    #[allow(clippy::too_many_arguments)]
    fn exec_collection_fast_borrowed(
        &mut self,
        top: usize,
        cur_prog: usize,
        code: &CompiledFunc,
        recv: PackedOp,
        dst: Option<PackedPlace>,
        method: CollectionMethod,
        pargs: &[FastCallArg],
    ) -> Result<(), Vm3Error> {
        if method == CollectionMethod::Add
            && let [
                FastCallArg::Operand(item),
                FastCallArg::Operand(key),
                tail @ ..,
            ] = pargs
            && tail.iter().all(|arg| matches!(arg, FastCallArg::Omitted))
            && let (Some(recv), Some(item), Some(key)) = (
                self.borrow_packed_op(top, cur_prog, code, recv),
                self.borrow_packed_op(top, cur_prog, code, *item),
                self.borrow_packed_op(top, cur_prog, code, *key),
            )
        {
            borrow_with_native_collection(recv, |data| {
                add_collection_item(data, Some(item), Some(key), None, None)
            })
            .ok_or_else(|| Vm3Error::Fault(Fault::new(424, "Object required")))?
            .map_err(|error| Vm3Error::Fault(Fault::standard(error.vba_number())))?;
            if let Some(dst) = dst {
                self.write_packed_place(top, cur_prog, dst, Variant::empty())?;
            }
            return Ok(());
        }
        // Zero-clone fast path for `c.Count` and `c(key)` / `c.Item(key)` / `c.Remove(key)`.
        if pargs.is_empty()
            && let Some(recv_ref) = self.borrow_packed_op(top, cur_prog, code, recv)
        {
            let ret = borrow_with_native_collection(recv_ref, |data| {
                dispatch_collection(method, data, &[])
            })
            .ok_or_else(|| Vm3Error::Fault(Fault::new(424, "Object required")))?
            .map_err(|error| Vm3Error::Fault(Fault::standard(error.vba_number())))?;
            if let Some(dst) = dst {
                self.write_packed_place(top, cur_prog, dst, ret)?;
            }
            return Ok(());
        }
        if pargs.len() == 1
            && matches!(method, CollectionMethod::Item | CollectionMethod::Remove)
            && let FastCallArg::Operand(arg_op) = pargs[0]
            && let (Some(recv_ref), Some(arg_ref)) = (
                self.borrow_packed_op(top, cur_prog, code, recv),
                self.borrow_packed_op(top, cur_prog, code, arg_op),
            )
        {
            let ret = borrow_with_native_collection(recv_ref, |data| {
                dispatch_collection(method, data, std::slice::from_ref(arg_ref))
            })
            .ok_or_else(|| Vm3Error::Fault(Fault::new(424, "Object required")))?
            .map_err(|error| Vm3Error::Fault(Fault::standard(error.vba_number())))?;
            if let Some(dst) = dst {
                self.write_packed_place(top, cur_prog, dst, ret)?;
            }
            return Ok(());
        }
        let mut argv: StackVariantBuf<4> = StackVariantBuf::new();
        for arg in pargs {
            let v = match *arg {
                FastCallArg::Operand(op) => self.eval_packed_op(top, cur_prog, code, op)?,
                FastCallArg::ByRef(place) => self.eval_packed_place(top, cur_prog, place)?,
                FastCallArg::Omitted => Variant::from_error_code(MISSING_ARG),
                FastCallArg::Const => {
                    return Err(Vm3Error::Malformed(
                        "a Const argument in a Collection method call".into(),
                    ));
                }
            };
            argv.push(v);
        }
        let recv_ref = self
            .borrow_packed_op(top, cur_prog, code, recv)
            .ok_or_else(|| Vm3Error::Fault(Fault::new(424, "Object required")))?;
        let ret = borrow_with_native_collection(recv_ref, |data| {
            dispatch_collection(method, data, argv.as_mut_slice())
        })
        .ok_or_else(|| Vm3Error::Fault(Fault::new(424, "Object required")))?
        .map_err(|error| Vm3Error::Fault(Fault::standard(error.vba_number())))?;
        if let Some(dst) = dst {
            self.write_packed_place(top, cur_prog, dst, ret)?;
        }
        Ok(())
    }

    #[allow(clippy::too_many_arguments)]
    fn try_push_project_proc_fast(
        &mut self,
        top: usize,
        cur_prog: usize,
        code: &CompiledFunc,
        recv: PackedOp,
        dst: Option<PackedPlace>,
        target_prog: usize,
        proc: FuncId,
        pargs: &[FastCallArg],
    ) -> Result<Option<bool>, Vm3Error> {
        let callee_prog = self.exec.programs[target_prog].program;
        let Some(callee) = callee_prog.funcs.get(proc.0) else {
            return Ok(None);
        };
        if pargs.len() + 1 > callee.param_count {
            return Ok(None);
        }
        if callee
            .locals
            .get(callee.param_count.saturating_sub(1))
            .and_then(|l| l.param.as_ref())
            .is_some_and(|p| p.variadic)
        {
            return Ok(None);
        }
        if pargs.iter().any(|a| matches!(a, FastCallArg::Const)) {
            return Ok(None);
        }
        self.guard_call_depth()?;
        let dst_loc = dst.map(|p| self.resolve_packed(top, cur_prog, p));
        let me = self.eval_packed_op(top, cur_prog, code, recv)?;
        let base = self.frames.len();
        let mut frame =
            self.new_frame_pooled(target_prog, proc, pargs.len() + 1, dst_loc, Some(me))?;
        let mut pending_param_array_aliases: Vec<(Loc, Vec<Option<Loc>>)> = Vec::new();
        for (i, arg) in pargs.iter().enumerate() {
            let li = i + 1;
            let formal = &callee.locals[li];
            let step: Result<(), Vm3Error> = match *arg {
                FastCallArg::Operand(op) => {
                    if self.param_array_aliases.is_empty()
                        && let Some(src_ref) = self.borrow_packed_op(top, cur_prog, code, op)
                        && write_fast_coerced_project_scalar(src_ref, formal, &mut frame.slots[li])
                    {
                        Ok(())
                    } else {
                        let pa = if !self.param_array_aliases.is_empty()
                            && let Some(place) = op.as_place()
                        {
                            self.param_array_aliases
                                .get(&self.resolve_packed(top, cur_prog, place))
                                .cloned()
                        } else {
                            None
                        };
                        match self.eval_packed_op(top, cur_prog, code, op) {
                            Ok(value) => match fast_coerce_project_arg(value, formal) {
                                Ok(coerced) => {
                                    frame.slots[li] = coerced;
                                    if let Some(aliases) = pa {
                                        pending_param_array_aliases
                                            .push((Loc::Local(base, li), aliases));
                                    }
                                    Ok(())
                                }
                                Err(e) => Err(e),
                            },
                            Err(e) => Err(e),
                        }
                    }
                }
                FastCallArg::ByRef(place) if formal.param.as_ref().is_some_and(|p| !p.by_ref) => {
                    if self.param_array_aliases.is_empty()
                        && let Some(src_ref) =
                            self.borrow_packed_op(top, cur_prog, code, PackedOp::from_place(place))
                        && write_fast_coerced_project_scalar(src_ref, formal, &mut frame.slots[li])
                    {
                        Ok(())
                    } else {
                        let pa = if !self.param_array_aliases.is_empty() {
                            self.param_array_aliases
                                .get(&self.resolve_packed(top, cur_prog, place))
                                .cloned()
                        } else {
                            None
                        };
                        match self.eval_packed_place(top, cur_prog, place) {
                            Ok(value) => match fast_coerce_project_arg(value, formal) {
                                Ok(coerced) => {
                                    frame.slots[li] = coerced;
                                    if let Some(aliases) = pa {
                                        pending_param_array_aliases
                                            .push((Loc::Local(base, li), aliases));
                                    }
                                    Ok(())
                                }
                                Err(e) => Err(e),
                            },
                            Err(e) => Err(e),
                        }
                    }
                }
                FastCallArg::ByRef(place) => {
                    let loc = self.resolve_packed(top, cur_prog, place);
                    let ox_place = place.to_ox(code.num_locals);
                    match self.byref_reference(&ox_place, Some(&formal.ty)) {
                        Ok(Some(reference)) => {
                            frame.slots[li] = reference;
                            Ok(())
                        }
                        Ok(None) => {
                            frame.set_alias(li, loc);
                            Ok(())
                        }
                        Err(e) => Err(e),
                    }
                }
                FastCallArg::Omitted => match omitted_param_value(formal) {
                    Some(Ok(v)) => {
                        frame.slots[li] = v;
                        Ok(())
                    }
                    Some(Err(e)) => Err(e),
                    None => {
                        frame.slots[li] = Variant::from_error_code(MISSING_ARG);
                        Ok(())
                    }
                },
                FastCallArg::Const => unreachable!("checked above"),
            };
            if let Err(e) = step {
                self.recycle_frame(frame);
                return Err(e);
            }
        }
        let saved_err = self.exec.err_engine.enter_activation();
        frame.saved_error_mode = saved_err.error_mode;
        frame.saved_active_error = saved_err.active_error;
        self.frames.push(frame);
        for (loc, aliases) in pending_param_array_aliases {
            self.param_array_aliases.insert(loc, aliases);
        }
        Ok(Some(true))
    }

    /// Jump frame `top` to the start of `block`.
    fn goto(&mut self, top: usize, block: BlockId) {
        let (prog, func) = {
            let fr = &self.frames[top];
            (fr.prog, fr.func)
        };
        let code = self.get_or_compile(prog, func);
        let pc = code.block_start_pc.get(block.0).copied().unwrap_or(0) as usize;
        let fr = &mut self.frames[top];
        fr.block = block;
        fr.pc = pc;
        fr.at_block_start = true;
    }

    /// Route an in-flight fault to the current frame's block fault pad (intra-frame);
    /// the pad's `FaultDispatch` then consults the error mode.
    fn route_fault(&mut self, fault: Fault) -> Result<(), Vm3Error> {
        let top = self.frames.len() - 1;
        // Re-target `cur` to the faulting frame's program: `propagate_fault` may have popped to a
        // caller in a DIFFERENT program before re-routing here, so the pad lookup + `Err.Source`
        // (raise) must use that frame's program (mirrors vm2's route_fault `cur` restore).
        self.cur = self.frames[top].prog;
        let (func, block) = (self.frames[top].func, self.frames[top].block);
        let pad = self.cur_program().funcs[func.0].blocks[block.0]
            .fault_target
            .ok_or_else(|| {
                Vm3Error::Malformed("fallible instruction in a block with no fault_target".into())
            })?;
        self.raise(fault);
        self.goto(top, pad);
        Ok(())
    }

    /// Propagate an unhandled fault out of the current frame: pop it (restoring the
    /// caller's error mode) and re-route at the caller's call site, or — at the base
    /// frame — surface it as the run's result.
    fn propagate_fault(&mut self, fault: Fault, base: usize) -> Result<(), Vm3Error> {
        if self.frames.len() <= base + 1 {
            return Err(Vm3Error::Fault(fault));
        }
        let callee = self.frames.pop().expect("frame to unwind");
        let drain_on_fault = callee.drain_on_fault;
        self.restore_err_from_frame(&callee);
        self.recycle_frame(callee);
        self.prune_param_array_aliases_from_depth(self.frames.len());
        if drain_on_fault {
            self.maybe_drain();
        }
        // The caller's `CallProc` faulted: route to *its* block's fault pad.
        self.route_fault(fault)
    }

    /// Pop a returning callee: restore the caller's error mode and copy out the
    /// function's return value (true aliasing already propagated ByRef writes live).
    fn do_return(&mut self) -> Result<(), Vm3Error> {
        let mut callee = self.frames.pop().expect("returning frame");
        self.restore_err_from_frame(&callee);
        let write_res = if let Some(loc) = callee.dst {
            if let Some(rl) = callee.return_local
                && rl.0 < callee.num_locals
            {
                let ret_slot = &mut callee.slots[rl.0];
                if self.param_array_aliases.is_empty()
                    && Variant::raw_vt_is_trivial_copy(ret_slot.raw_vt())
                    && let Some(dst_slot) = self.read_loc_mut_fast(loc)
                    && {
                        let old_vt = dst_slot.raw_vt();
                        old_vt < 8 || old_vt == 11
                    }
                {
                    dst_slot.copy_trivial_from(ret_slot);
                    Ok(())
                } else {
                    let v = std::mem::replace(ret_slot, Variant::empty());
                    self.write_loc(loc, v)
                }
            } else if callee.drain_on_fault {
                self.write_loc(loc, Variant::empty())
            } else {
                Ok(())
            }
        } else {
            Ok(())
        };
        // Proc epilogue: release the callee frame (and its held objects) and then
        // run any parked `Class_Terminate`s — vm2's epilogue drain timing.
        self.recycle_frame(callee);
        if !self.param_array_aliases.is_empty() || !self.as_new_slots.is_empty() {
            self.prune_param_array_aliases_from_depth(self.frames.len());
        }
        if self.exec.object_owner.has_pending() {
            self.maybe_drain();
        }
        write_res
    }

    fn restore_err_from_frame(&mut self, frame: &Frame) {
        self.exec.err_engine.restore(SavedErrState {
            error_mode: frame.saved_error_mode,
            active_error: frame.saved_active_error,
        });
    }

    /// Populate `Err` from a raised fault and stash it for the landing pad. Number and
    /// Description come from the fault; `Err.Source` is the fault's explicit source if it
    /// carries one (`Err.Raise … Source`), else the **project name** — the VBA default
    /// for any error generated within the project, matching the Excel/VBA 7.1 oracle
    /// (`Err.Source = "VBAProject"`; see `docs/VBA_ERROR_MODEL_ORACLE_FINDINGS.md`).
    fn raise(&mut self, fault: Fault) {
        self.exec
            .err_engine
            .raise(fault, self.cur_program().unit_name.clone());
    }

    /// Raise a vm3-internal run-time error (a code with its default message) by routing
    /// it through the current statement's fault pad — so an active `On Error` can catch
    /// it (used for error 20 "Resume without error" and error 3 "Return without GoSub").
    fn raise_runtime_error(&mut self, code: i32) -> Result<(), Vm3Error> {
        self.route_fault(Fault {
            code,
            message: default_error_message(code).into(),
            source: None,
            help_file: None,
            help_context: None,
        })
    }

    /// Execute one straight-line instruction against the top frame.
    fn exec(&mut self, inst: &OxInst) -> Result<(), Vm3Error> {
        match inst {
            OxInst::AsNew { place, binding } => {
                let loc = self.resolve(place);
                self.as_new_slots.insert(loc, binding.clone());
            }
            OxInst::Assign { dst, value } => {
                let v = self.operand(value)?;
                self.store_assigned(dst, v)?;
            }
            OxInst::Box { dst, src, .. } => {
                let v = self.operand(src)?;
                self.store_assigned(dst, v)?;
            }
            OxInst::Unbox {
                dst,
                src,
                to,
                checked,
            } => {
                let v = self.operand(src)?;
                if *checked && !variant_matches_ox_ty(&v, to) {
                    return Err(Vm3Error::Fault(Fault::new(13, "type mismatch")));
                }
                self.store_assigned(dst, v)?;
            }
            // `dst := (current != original)` — VBA-`Variant`-equality change detection that
            // guards a compound `ByRef` copy-out (mirrors vm2's `Op::VariantChanged`, which
            // compares with `Variant`'s `PartialEq`).
            OxInst::VariantChanged {
                dst,
                current,
                original,
            } => {
                let changed = variant_changed(&self.operand(current)?, &self.operand(original)?);
                self.store(dst, Variant::from_bool(changed))?;
            }
            OxInst::Arith {
                dst,
                op,
                lhs,
                rhs,
                mode,
            } => {
                let left = self.operand(lhs)?;
                let right = self.operand(rhs)?;
                let l = self.arithmetic_operand(left)?;
                let r = self.arithmetic_operand(right)?;
                let out = match op {
                    ArithOp::Add => arith::add(&l, &r, *mode),
                    ArithOp::Sub => arith::sub(&l, &r, *mode),
                    ArithOp::Mul => arith::mul(&l, &r, *mode),
                    ArithOp::IntDiv => arith::int_div(&l, &r, *mode),
                    ArithOp::Mod => arith::modulo(&l, &r, *mode),
                };
                self.store_arith(dst, out)?;
            }
            OxInst::Div { dst, lhs, rhs } => {
                let left = self.operand(lhs)?;
                let right = self.operand(rhs)?;
                let l = self.arithmetic_operand(left)?;
                let r = self.arithmetic_operand(right)?;
                self.store_arith(dst, arith::div(&l, &r))?;
            }
            OxInst::Pow { dst, lhs, rhs } => {
                let left = self.operand(lhs)?;
                let right = self.operand(rhs)?;
                let l = self.arithmetic_operand(left)?;
                let r = self.arithmetic_operand(right)?;
                self.store_arith(dst, arith::pow(&l, &r))?;
            }
            OxInst::Neg { dst, src, mode } => {
                let value = self.operand(src)?;
                let v = self.arithmetic_value(value)?;
                self.store_arith(dst, arith::neg(&v, *mode))?;
            }
            OxInst::Concat { dst, lhs, rhs } => {
                let left = self.operand(lhs)?;
                let right = self.operand(rhs)?;
                let l = self.arithmetic_value(left)?;
                let r = self.arithmetic_value(right)?;
                self.store_arith(dst, arith::concat(&l, &r))?;
            }
            OxInst::Compare {
                dst,
                op,
                lhs,
                rhs,
                mode,
                coercion,
            } => {
                let left = self.operand(lhs)?;
                let right = self.operand(rhs)?;
                let l = self.arithmetic_value(left)?;
                let r = self.arithmetic_value(right)?;
                self.store_arith(
                    dst,
                    arith::compare_declared(&l, &r, *mode, cmp_op(*op), *coercion),
                )?;
            }
            OxInst::Logical { dst, op, lhs, rhs } => {
                let left = self.operand(lhs)?;
                let right = self.operand(rhs)?;
                let l = self.arithmetic_value(left)?;
                let r = self.arithmetic_value(right)?;
                let out = match op {
                    LogicalOp::And => arith::and(&l, &r),
                    LogicalOp::Or => arith::or(&l, &r),
                    LogicalOp::Xor => arith::xor(&l, &r),
                    LogicalOp::Eqv => arith::eqv(&l, &r),
                    LogicalOp::Imp => arith::imp(&l, &r),
                };
                self.store_arith(dst, out)?;
            }
            OxInst::Not { dst, src } => {
                let value = self.operand(src)?;
                let v = self.arithmetic_value(value)?;
                self.store_arith(dst, arith::not(&v))?;
            }
            // Reduce a condition to a Boolean by VBA truthiness (the elaboration emits
            // this before a conditional `Branch`); the `is_truthy` rule + error code are
            // exactly what vm2's `JumpIfZero` uses.
            OxInst::Truthy { dst, src } => {
                let value = self.operand(src)?;
                let v = self.arithmetic_value(value)?;
                let out = arith::is_truthy(&v).map(Variant::from_bool);
                self.store_arith(dst, out)?;
            }
            OxInst::Coerce { dst, src, target } => {
                let value = self.operand(src)?;
                // A conversion reads an object's default member (`n = w`, `LongArg(w)`); the
                // widen to Variant keeps the object.
                let v = if matches!(target, OxCoerceTarget::ImplicitVariant) {
                    value
                } else {
                    self.arithmetic_value(value)?
                };
                let out = match target {
                    OxCoerceTarget::Numeric(t) => arith::coerce_numeric(&v, *t),
                    OxCoerceTarget::Str => arith::coerce_string(&v),
                    OxCoerceTarget::FixedStr(n) => arith::coerce_fixed_string(&v, *n as usize),
                    // A widen-to-Variant carries no value change.
                    OxCoerceTarget::ImplicitVariant => Ok(v),
                };
                self.store_arith(dst, out)?;
            }
            // A compiled VBA procedure call (intra-unit). The `AddressOf`-reference
            // `CallProcRef` is M3-7.
            OxInst::CallProc { dst, proc, args } => {
                self.call_proc_in(self.cur, *dst, *proc, args)?
            }
            // A cross-bundle call. vm3 links only the synthetic `VBA` library bundle today,
            // so this resolves to a native library function (`Strings.Left`, `Math.Abs`, …)
            // run through the same `invoke_native_lib` bridge as `CallNative { Builtin }`. A
            // reference to another VBA *project* needs a multi-`OxProgram` linker (deferred —
            // surfaced as an explicit `Unimplemented`, never a silent skip).
            OxInst::CallExtern { dst, import, args } => self.call_extern(*dst, *import, args)?,
            // A base-library built-in is classified into the disjoint context-free or
            // contextual `calcvba-lib` dispatcher; `Declare Lib` marshalling is M3.
            OxInst::CallNative { dst, callee, args } => match callee {
                OxNativeCallee::Builtin(id) => {
                    let mut argv = self.native_args(args)?;
                    let result = match self.vba_object_len(*id, &argv)? {
                        Some(value) => value,
                        None => {
                            self.native_object_values(*id, &mut argv)?;
                            self.invoke_native_lib(*id, &argv)?
                        }
                    };
                    if let Some(dst) = dst {
                        self.store(dst, result)?;
                    }
                }
                // A `Declare Lib` external call marshals through the host's dynamic-link HAL —
                // the identical `invoke_descriptor_variants` contract vm2 drives — capturing
                // `Err.LastDllError` and applying ByRef write-backs (M3-7).
                OxNativeCallee::Declare { descriptor_id } => {
                    let result = self.declare_call(*descriptor_id, args)?;
                    if let Some(dst) = dst {
                        self.store(dst, result)?;
                    }
                }
            },
            // `AddressOf <proc>` materializes a procedure reference value (a `ProcRef`
            // Variant carrying the proc index); it is later invoked through `CallProcRef`
            // or marshaled to a native callback slot by a `Declare` (M3-7).
            OxInst::LoadProcRef { dst, proc } => {
                self.store(dst, Variant::from_proc_ref(proc.0))?;
            }
            // Call through a runtime-resolved procedure reference (the `AddressOf` value):
            // resolve the index, validate it, then dispatch through the standard call
            // machinery (ByRef aliasing, frame push, return copy-out — identical to `CallProc`).
            OxInst::CallProcRef { dst, target, args } => {
                let proc = self
                    .operand(target)?
                    .as_proc_ref()
                    .filter(|&p| p < self.cur_program().funcs.len())
                    .ok_or_else(|| {
                        Vm3Error::Fault(Fault::new(490, "invalid procedure reference"))
                    })?;
                let dst = dst.as_ref().copied();
                self.call_proc_in(self.cur, dst, FuncId(proc), args)?;
            }
            // `On Error` sets the activation's handler policy and — per MS-VBAL §5.4.4.1
            // (doc rule R5) — unconditionally resets the `Err` object. (The active-error
            // latch is cleared only by `Resume`/`Exit`, not here.)
            OxInst::SetErrorHandler(handler) => {
                self.exec.err_engine.set_error_handler(handler);
            }
            // Read an `Err` property.
            OxInst::ErrFieldGet { dst, field } => {
                let v = self.exec.err_engine.field_value(*field);
                self.store(dst, v)?;
            }
            OxInst::ErlGet { dst } => {
                self.store(dst, Variant::from_i32(self.exec.err_engine.erl_line))?;
            }
            // Write an `Err` property. `Err.LastDllError` is read-only; accepted user
            // assignments to it should have been rejected by the binder.
            OxInst::ErrFieldSet { field, src } => {
                let value = self.operand(src)?;
                if *field == ErrField::LastDllError {
                    return Err(Vm3Error::Malformed("Err.LastDllError is read-only".into()));
                }
                self.exec
                    .err_engine
                    .set_field(*field, &value)
                    .map_err(Vm3Error::Fault)?;
            }
            // `Err.Clear` → reset the `Err` object.
            OxInst::ClearErr => self.exec.err_engine.clear_err(),
            OxInst::ClearErrOnExit => self.exec.err_engine.clear_err_on_exit(),

            // ── Pointer helpers ──────────────────────────────────────────────────────
            // `VarPtr`/`StrPtr`/`ObjPtr` give the real address of the storage (or the BSTR /
            // interface pointer it holds), located by `calcvba_runtime::address` exactly as
            // the JIT does. An r-value operand is a statement temporary.
            OxInst::AddrOf {
                dst,
                kind,
                base,
                path,
            } => {
                let pointer = self.address_of(*kind, base, path)?;
                self.store(dst, Variant::from_i64(pointer))?;
            }
            // A ByRef argument that is an array element or UDT field: a reference to it, so
            // the callee reads and writes the storage itself (the JIT's
            // `rt_jit_make_ref_to_slot` does the same).
            OxInst::MakeRef { dst, base, path } => {
                // `v(i)` passed ByRef where the Variant `v` holds an object: the default
                // member's value is passed, as for any property result.
                if let (OxAddrBase::Place(place), [OxAddrStep::Index(indices)]) =
                    (base, path.as_slice())
                {
                    let loc = self.resolve(place);
                    let holds_object = self
                        .read_loc_mut(loc)?
                        .is_some_and(|cell| cell.holds_object_through_ref());
                    if holds_object {
                        let recv = self.read(place)?;
                        return self.index_value_into(recv, indices, dst);
                    }
                }
                let steps = self.address_steps(path)?;
                let root = self.address_root(base)?;
                let fault = |number: i32| {
                    Vm3Error::Fault(Fault::new(number, default_error_message(number)))
                };
                // SAFETY: `root` is a live, stable engine cell; no Rust reference to it is held.
                let (location, locks) =
                    unsafe { calcvba_runtime::address::locate_locking(root, &steps) }
                        .map_err(fault)?;
                // SAFETY: the storage lives at least for the statement the reference is for.
                let reference = unsafe { location.reference(true) }.ok_or_else(|| {
                    Vm3Error::Malformed("MakeRef over storage with no reference form".into())
                })?;
                // The destination is a fresh statement temporary: replace, never write through.
                let loc = self.resolve(dst);
                let Loc::Temp(frame_index, temp) = loc else {
                    return Err(Vm3Error::Malformed(format!("MakeRef destination {loc:?}")));
                };
                let slot = self
                    .read_loc_mut(loc)?
                    .ok_or_else(|| Vm3Error::Malformed(format!("MakeRef destination {loc:?}")))?;
                *slot = reference;
                let frame = &mut self.frames[frame_index];
                frame.ref_locks.retain(|(held, _)| *held != temp);
                frame
                    .ref_locks
                    .extend(locks.into_iter().map(|lock| (temp, lock)));
            }
            OxInst::ForEachEnd { iter } => {
                if let OxPlace::Temp(iter_temp) = iter {
                    let frame = self.frames.last_mut().expect("executing frame");
                    if let Some(at) = frame
                        .loop_locks
                        .iter()
                        .rposition(|(held, _)| *held == iter_temp.0)
                    {
                        frame.loop_locks.remove(at);
                    }
                }
            }
            OxInst::ReleaseRef { src } => {
                if let Loc::Temp(frame_index, temp) = self.resolve(src) {
                    self.frames[frame_index]
                        .ref_locks
                        .retain(|(held, _)| *held != temp);
                }
            }
            // A statement boundary drives finalization timing: run any parked
            // `Class_Terminate`s released by the previous statement (the error model takes its
            // `Resume` seeds from `FaultDispatch`, not from here).
            OxInst::StmtBoundary {
                clear_temps_from, ..
            } => {
                // VBA statement-granular timing: run parked `Class_Terminate`s, then dispatch any
                // inbound host (COM) events (mirrors vm2's statement-boundary drain + pump).
                self.clear_statement_temps(*clear_temps_from);
                self.maybe_drain();
            }
            OxInst::SetLineNumber { line } => {
                if let Some(frame) = self.frames.last_mut() {
                    frame.current_line = *line;
                }
            }

            // `Let`/`Set` legality check (M3-4).
            OxInst::ValidateAssignment {
                src,
                intent,
                target_kind,
                target_name,
                target_type_name,
            } => {
                self.validate_assignment(src, *intent, *target_kind, target_name, target_type_name)?
            }

            // ── Arrays / For Each (M3-2) ─────────────────────────────────────────────
            OxInst::ArrayLiteral {
                dst,
                values,
                aliases,
                lower_bound,
            } => {
                let elems = values
                    .iter()
                    .map(|v| self.operand(v))
                    .collect::<Result<Vec<_>, _>>()?;
                let alias_locs = aliases
                    .iter()
                    .map(|alias| alias.as_ref().map(|place| self.resolve(place)))
                    .collect::<Vec<_>>();
                // `Array()` is based at the module's `Option Base` (0 or 1); a
                // `ParamArray` always at 0. A non-zero base needs explicit bounds.
                let array = SafeArray::array_literal(*lower_bound, elems)
                    .map_err(|error| Vm3Error::Fault(Fault::new(7, error)))?;
                self.store(dst, Variant::from_safearray(array))?;
                if alias_locs.iter().any(Option::is_some) {
                    let dst_loc = self.resolve(dst);
                    self.param_array_aliases.insert(dst_loc, alias_locs);
                }
            }
            OxInst::ArrayAppend { dst, array, item } => {
                let mut elems = match self.operand(array)?.as_safearray() {
                    Some(arr) => arr.variant_elements().unwrap_or_default(),
                    None => Vec::new(),
                };
                elems.push(self.operand(item)?);
                self.store(
                    dst,
                    Variant::from_safearray(SafeArray::from_variants(elems)),
                )?;
            }
            OxInst::ArrayRedim {
                dst,
                upper_bounds,
                lower_bounds,
                element,
                preserve,
                fixed,
                explicit_element,
            } => self.array_redim(
                dst,
                upper_bounds,
                lower_bounds,
                element,
                (*preserve, *fixed, *explicit_element),
            )?,
            OxInst::ArrayGet {
                dst,
                array,
                indices,
            } => {
                // Fast O(1) path: a place-resident array reads one element directly
                // through the SAFEARRAY descriptor (no whole-array clone) → O(N) loops.
                if let Some(value) = self.array_get_fast(array, indices)? {
                    self.store(dst, value)?;
                    return Ok(());
                }
                // `x(i…)` where `x` is a bare `Variant`/`As Object` resolves at run time.
                let recv = self.operand(array)?;
                self.index_value_into(recv, indices, dst)?;
            }
            OxInst::ArraySet {
                array,
                indices,
                value,
            } => {
                let v = self.operand(value)?;
                // Fast O(1) path: a place-resident array mutates one element in place
                // through the SAFEARRAY descriptor (no whole-array clone-and-write-back).
                // The element is written through the resolved (alias-resolved) location,
                // so a ByRef-aliased array sees the change.
                if self.array_set_fast(array, indices, &v)? {
                    return Ok(());
                }
                // General path: a non-place receiver — materialize, set the element, and
                // write the mutated array back to its place.
                let recv = self.read(array)?;
                let arr_v = self.set_index_in_value(recv, indices, &v)?;
                self.store(array, arr_v)?;
            }
            OxInst::FieldArrayGet {
                dst,
                object,
                field,
                indices,
            } => {
                let recv = self.operand(object)?;
                let instance = variant_to_object(&recv)?;
                // Fast O(1): the field holds an array → read one element through the
                // descriptor IN PLACE (no whole-array clone per access). `with_project_field`
                // borrows the field Variant; a non-array field yields `None` → fall back.
                let fast = instance
                    .with_project_field(*field, |stored| {
                        let (arr, (bounds, len)) =
                            stored.and_then(|a| a.safearray_bounds_len().map(|bl| (a, bl)))?;
                        Some((|| -> Result<Variant, Vm3Error> {
                            let flat = self.flat_index(indices, &bounds)?;
                            if flat >= len {
                                return Err(Vm3Error::Fault(Fault::new(
                                    9,
                                    default_error_message(9),
                                )));
                            }
                            if let Some(result) = arr.safearray_i32_element(flat) {
                                match result {
                                    Ok(Some(value)) => return Ok(Variant::from_i32(value)),
                                    Ok(None) => {}
                                    Err(err) => {
                                        return Err(Vm3Error::Fault(Fault::new(13, err)));
                                    }
                                }
                            }
                            arr.safearray_element(flat)
                                .expect("safearray_bounds_len proved this is an array")
                                .map_err(|e| Vm3Error::Fault(Fault::new(13, e)))
                        })())
                    })
                    .flatten();
                match fast {
                    Some(result) => {
                        let value = result?;
                        self.store(dst, value)?;
                    }
                    None => {
                        // Field is absent or not an array (e.g. an object whose default
                        // member is indexed): materialize (cheap for an object ref) and
                        // index generically.
                        let field_val = self.read_project_field_as_new(&instance, *field)?;
                        self.index_value_into(field_val, indices, dst)?;
                    }
                }
            }
            OxInst::FieldArraySet {
                object,
                field,
                indices,
                value,
            } => {
                let v = self.operand(value)?;
                let recv = self.operand(object)?;
                let instance = variant_to_object(&recv)?;
                // Fast O(1): the field holds an array → mutate one element IN PLACE.
                let fast = instance
                    .with_project_field_mut(*field, |arr| {
                        let (bounds, len) = arr.safearray_bounds_len()?;
                        Some((|| -> Result<(), Vm3Error> {
                            let flat = self.flat_index(indices, &bounds)?;
                            if flat >= len {
                                return Err(Vm3Error::Fault(Fault::new(
                                    9,
                                    default_error_message(9),
                                )));
                            }
                            if let Some(value_i32) = v.as_i32()
                                && let Some(result) = arr.set_safearray_i32_element(flat, value_i32)
                            {
                                match result {
                                    Ok(true) => return Ok(()),
                                    Ok(false) => {}
                                    Err(err) => {
                                        return Err(Vm3Error::Fault(Fault::new(13, err)));
                                    }
                                }
                            }
                            arr.set_safearray_element(flat, &v)
                                .map_err(|e| Vm3Error::Fault(Fault::new(13, e)))
                        })())
                    })
                    .flatten();
                match fast {
                    Some(result) => result?,
                    None => {
                        // Not an array field: materialize, set the element, write back.
                        let field_val = self.read_project_field_as_new(&instance, *field)?;
                        let updated = self.set_index_in_value(field_val, indices, &v)?;
                        instance.project_field_set(*field, updated);
                    }
                }
            }
            OxInst::ArrayErase { array, element } => {
                let location = self.resolve(array);
                if self
                    .read_loc_ref(location)?
                    .is_some_and(Variant::holds_locked_array)
                {
                    return Err(Vm3Error::Fault(Fault::new(
                        10,
                        calcvba_rt_abi::default_error_message(10),
                    )));
                }
                // VBA `Erase` is reset-vs-deallocate decided by the array's *own* storage
                // class, carried on the runtime SAFEARRAY's `FADF_FIXEDSIZE` bit (set at
                // allocation, travelling with copies) — so the dispatch is purely runtime,
                // exactly like real VBA:
                //   • fixed-size (`Dim a(1 To 3)`): reinitialize every element to its type
                //     default and keep the array allocated (reads after `Erase` succeed);
                //   • dynamic (`Dim a()` + `ReDim`): free the storage (becomes
                //     uninitialized; `UBound` raises until re-`ReDim`'d).
                let cur = self.read(array)?;
                let was_array = cur.vtype() == VarType::ArrayVariant;
                let erased_element_vartype = cur
                    .array_element_vartype()
                    .unwrap_or_else(|| safearray_vartype_for_element(&ArrayElementType::Variant));
                let reset = cur.as_safearray().filter(|a| a.is_fixed_size()).map(|arr| {
                    // Rebuild a fresh array of the SAME bounds + element type + fixed flag,
                    // default-initialized — i.e. a `ReDim`-to-current-bounds, which already
                    // default-inits correctly for every element type (scalars, String → "",
                    // Variant → Empty, UDT → recursively zeroed record).
                    //
                    // Pick the reset element type. The bind-site `element` is the authoritative
                    // DECLARED type for a directly-typed array — including a UDT fixed-array
                    // field, whose materialized value is a VT_VARIANT SAFEARRAY that has lost the
                    // declared element type. A bind-site `Variant` is the ambiguous case: either a
                    // genuine `Variant` array, or a typed array materialized into a `Variant` slot
                    // (which keeps its real element vartype on the value but erases the bind-site
                    // type). Only then is the array's OWN runtime vartype the truth.
                    let bounds = arr.bounds().unwrap_or_default();
                    let count = arr.len();
                    if let Some(rec_layout) = cur.record_array_layout() {
                        let elems: Vec<Variant> = (0..count)
                            .map(|_| {
                                calcvba_bundle::array_runtime::default_record_array_element(
                                    &rec_layout,
                                )
                            })
                            .collect::<Result<Vec<_>, String>>()?;
                        return calcvba_bundle::array_runtime::redim_record_safearray_from_elements(
                            bounds, rec_layout, elems, true,
                        );
                    }
                    let et = match element {
                        ArrayElementType::Variant => {
                            array_element_type_for_vartype(arr.element_vartype())
                        }
                        other => other.clone(),
                    };
                    if Self::element_supports_zeroed_redim(&et) {
                        return SafeArray::from_zeroed_typed_scalars_nd(
                            bounds,
                            safearray_vartype_for_element(&et),
                        )
                        .map(|array| array.with_fixed_size(true));
                    }
                    let elems: Vec<Variant> = (0..count)
                        .map(|_| default_array_element(&et))
                        .collect::<Result<Vec<_>, String>>()?;
                    redim_safearray_from_elements(bounds, &et, elems, true)
                });
                match reset {
                    Some(built) => {
                        let array_value = built.map_err(|e| Vm3Error::Fault(Fault::new(13, e)))?;
                        let loc = self.resolve(array);
                        let current = self
                            .read_loc_mut(loc)?
                            .ok_or_else(|| Vm3Error::Malformed("missing fixed array".into()))?;
                        for index in 0..array_value.len() {
                            let value = array_value
                                .variant_element(index)
                                .map_err(|error| Vm3Error::Fault(Fault::new(13, error)))?;
                            current
                                .set_safearray_element(index, &value)
                                .map_err(|error| Vm3Error::Fault(Fault::new(13, error)))?;
                        }
                    }
                    None if was_array => {
                        let value = if let Some(layout) = cur.record_array_layout() {
                            Variant::unallocated_record_array(layout)
                        } else {
                            Variant::unallocated_array(erased_element_vartype)
                        };
                        self.store(array, value)?
                    }
                    // `Erase` of a Variant that is not an array is a type mismatch.
                    None => {
                        return Err(Vm3Error::Fault(Fault::new(
                            13,
                            calcvba_rt_abi::default_error_message(13),
                        )));
                    }
                }
            }
            OxInst::Bound {
                dst,
                which,
                array,
                dimension,
            } => {
                let arr = self.array_of(array)?;
                let bounds = arr
                    .bounds()
                    .ok_or_else(|| Vm3Error::Fault(Fault::new(9, "array has no bounds")))?;
                let bound = &bounds[self.array_bound_index(dimension.as_ref(), &bounds)?];
                let value = match which {
                    BoundWhich::Lower => bound.lower,
                    BoundWhich::Upper => bound.lower + bound.count as i32 - 1,
                };
                self.store(dst, Variant::from_i32(value))?;
            }
            OxInst::ForEachInit { iter, source } => {
                let src = self.operand(source)?;
                let array_source = if src.safearray_bounds_len().is_some() {
                    match source {
                        OxOperand::Use(place @ (OxPlace::Local(_) | OxPlace::Global(_))) => {
                            Some(self.resolve(place))
                        }
                        // A reference to a dynamic-array field (`For Each x In t.d`).
                        OxOperand::Use(place @ OxPlace::Temp(_))
                            if self
                                .read_loc_ref(self.resolve(place))?
                                .is_some_and(|cell| cell.byref_target().is_some()) =>
                        {
                            Some(self.resolve(place))
                        }
                        _ => None,
                    }
                } else {
                    None
                };
                // VBA locks the array for the loop (`cLocks` + 1).
                if let (Some(source), OxPlace::Temp(iter_temp)) = (array_source, iter)
                    && let Some(lock) = self.read_loc_ref(source)?.and_then(Variant::lock_array)
                {
                    let frame = self.frames.last_mut().expect("executing frame");
                    frame.loop_locks.push((iter_temp.0, lock));
                }
                let (elements, collection_cursor, enumerator) = match self
                    .foreach_elements(src, 0)?
                {
                    ForEachItems::Snapshot(elements) => (elements, None, None),
                    ForEachItems::CollectionCursor(cursor) => (Vec::new(), Some(cursor), None),
                    ForEachItems::Enumerator(enumerator) => (Vec::new(), None, Some(enumerator)),
                };
                let key = self.resolve(iter);
                self.for_each.insert(
                    key,
                    ForEachState {
                        array_source,
                        elements,
                        position: 0,
                        collection_cursor,
                        enumerator,
                    },
                );
            }
            OxInst::ForEachNext {
                iter,
                item,
                has_value,
            } => {
                let key = self.resolve(iter);
                // Built-in `VBA.Collection` cursor: direct in-process read with no COM
                // event scope or trait-object boxing.
                let coll_next = self.for_each.get(&key).and_then(|state| {
                    state
                        .collection_cursor
                        .as_ref()
                        .map(CollectionCursor::next_item)
                });
                let next = if let Some(item_opt) = coll_next {
                    item_opt
                } else {
                    // The enumerator leaves the iterator table for the call: an event handler running
                    // inside it may start its own loops.
                    let enumerator = self
                        .for_each
                        .get_mut(&key)
                        .and_then(|state| state.enumerator.take());
                    if let Some(mut enumerator) = enumerator {
                        let next = {
                            let _events = self.event_scope();
                            enumerator.next_item()
                        };
                        if let Some(state) = self.for_each.get_mut(&key) {
                            state.enumerator = Some(enumerator);
                        }
                        self.take_event_abort()?;
                        next.map_err(|fault| {
                            Vm3Error::Fault(Fault {
                                code: fault.number,
                                message: fault
                                    .description
                                    .unwrap_or_else(|| default_error_message(fault.number).into()),
                                source: fault.source,
                                help_file: fault.help_file,
                                help_context: fault.help_context.map(|v| v as i32),
                            })
                        })?
                    } else if let Some(state) = self.for_each.get(&key) {
                        if state.position >= state.elements.len() {
                            None
                        } else if let Some(source) = state.array_source {
                            self.read_loc_ref(source)?
                                .and_then(|array| {
                                    array.safearray_enumeration_element(state.position)
                                })
                                .transpose()
                                .map_err(|error| Vm3Error::Fault(Fault::new(9, error)))?
                        } else {
                            state.elements.get(state.position).cloned()
                        }
                    } else {
                        None
                    }
                };
                if next.is_some() {
                    self.for_each
                        .get_mut(&key)
                        .expect("active iterator")
                        .position += 1;
                }
                match next {
                    Some(value) => {
                        self.store(item, value)?;
                        self.store(has_value, Variant::from_bool(true))?;
                    }
                    None => {
                        self.for_each.remove(&key);
                        self.store(has_value, Variant::from_bool(false))?;
                    }
                }
            }

            // ── Records / UDT (M3-3) — value aggregates with native VBA layout ────────
            OxInst::NewRecord { dst, fields } => {
                let anon = vba_record_layout_for_fields(fields)
                    .map_err(|e| Vm3Error::Fault(Fault::new(13, e)))?;
                let frame = self.frames.last().expect("NewRecord frame");
                let program = self.exec.programs[frame.prog].program;
                let function = &program.funcs[frame.func.0];
                let dst_ty = match dst {
                    OxPlace::Local(id) => &function.locals[id.0].ty,
                    OxPlace::Temp(id) => &function.temps[id.0],
                    OxPlace::Global(id) => &program.globals[id.0].ty,
                };
                let layout = if let OxTy::Record(id) = dst_ty
                    && let Ok(named) =
                        calcvba_rt_abi::vba_record_layout_for_program_record(program, *id)
                    && named.matches_structural(&anon)
                {
                    named
                } else {
                    anon
                };
                let record = VbaRecord::new_default(layout)
                    .map_err(|e| Vm3Error::Fault(Fault::new(13, e)))?;
                self.store(dst, Variant::from_vba_record(record))?;
            }
            OxInst::RecordGet { dst, record, index } => {
                let value = if let OxOperand::Use(place) = record {
                    let loc = self.resolve(place);
                    if (self.as_new_slots.is_empty() || !self.as_new_slots.contains_key(&loc))
                        && let Some(source) = self.read_loc_ref(loc)?
                        && source.raw_vt() == VarType::Record as u16
                    {
                        source
                            .read_record_field_variant(*index)
                            .map_err(|e| Vm3Error::Fault(Fault::new(13, e)))?
                    } else {
                        let source = self.operand(record)?;
                        source
                            .read_record_field_variant(*index)
                            .map_err(|e| Vm3Error::Fault(Fault::new(13, e)))?
                    }
                } else {
                    let source = self.operand(record)?;
                    source
                        .read_record_field_variant(*index)
                        .map_err(|e| Vm3Error::Fault(Fault::new(13, e)))?
                };
                self.store(dst, value)?;
            }
            OxInst::RecordArrayGet {
                dst,
                record,
                index,
                indices,
            } => {
                let value = self.record_array_get(record, *index, indices)?;
                self.store(dst, value)?;
            }
            OxInst::RecordSet {
                record,
                index,
                value,
            } => {
                let v = self.operand(value)?;
                let loc = self.resolve(record);
                if let Some(target) = self.read_loc_mut(loc)?
                    && target.raw_vt() == VarType::Record as u16
                {
                    target
                        .write_record_field_variant(*index, &v)
                        .map_err(|e| Vm3Error::Fault(Fault::new(13, e)))?;
                } else {
                    let mut target = self.read(record)?;
                    target
                        .write_record_field_variant(*index, &v)
                        .map_err(|e| Vm3Error::Fault(Fault::new(13, e)))?;
                    self.store(record, target)?;
                }
            }
            OxInst::RecordLSet { record, value } => {
                let source = self.operand(value)?;
                let mut target = self.read(record)?;
                target
                    .lset_record_from(&source)
                    .map_err(|e| Vm3Error::Fault(Fault::new(13, e)))?;
                self.store(record, target)?;
            }
            OxInst::RecordArraySet {
                record,
                index,
                indices,
                value,
            } => {
                let v = self.operand(value)?;
                self.record_array_set(record, *index, indices, &v)?;
            }

            // ── Objects / lifecycle / type identity (M3-5) ───────────────────────────
            OxInst::NewObject { dst, class } => {
                let object = self.new_project_instance(class.0)?;
                self.store(dst, object)?;
            }
            OxInst::NewExtern { dst, import } => {
                let object = self.new_extern_instance(*import)?;
                self.store(dst, object)?;
            }
            OxInst::Predeclared { dst, class } => {
                let object = self.predeclared_instance(class.0)?;
                self.store(dst, object)?;
            }
            OxInst::PredeclaredExtern { dst, import } => {
                let object = self.predeclared_extern_instance(*import)?;
                self.store(dst, object)?;
            }
            OxInst::PredeclaredSet { class, value } => {
                let object = self.operand(value)?;
                self.set_predeclared_instance(self.cur, class.0, object)?;
            }
            OxInst::PredeclaredExternSet { import, value } => {
                let object = self.operand(value)?;
                self.set_predeclared_extern_instance(*import, object)?;
            }
            OxInst::FieldGet { dst, object, field } => {
                let recv = self.operand(object)?;
                let instance = variant_to_object(&recv)?;
                let value = self.read_project_field_as_new(&instance, *field)?;
                self.store(dst, value)?;
            }
            OxInst::FieldSet {
                object,
                field,
                value,
            } => {
                let v = self.operand(value)?;
                let recv = self.operand(object)?;
                let instance = variant_to_object(&recv)?;
                instance.project_field_set(*field, v);
            }
            OxInst::CompareObjectIs { dst, lhs, rhs } => {
                let a = object_identity_for_is(&self.operand(lhs)?)?;
                let b = object_identity_for_is(&self.operand(rhs)?)?;
                self.store(dst, Variant::from_bool(a == b))?;
            }
            OxInst::TypeOfIs {
                dst,
                object,
                type_name,
            } => {
                let matches = self.type_of_is(object, type_name)?;
                self.store(dst, Variant::from_bool(matches))?;
            }

            // ── Events / WithEvents (M3-6) ───────────────────────────────────────────
            OxInst::WithEventsGet {
                dst,
                owner,
                binding,
            } => {
                let owner_value = self.operand(owner)?;
                let owner_ref = variant_to_object(&owner_value)?;
                let key = withevents_key(&owner_ref, *binding as i64);
                let value = self
                    .exec
                    .events
                    .withevents
                    .get(&key)
                    .map(|b| b.source.clone())
                    .unwrap_or_else(Variant::nothing);
                self.store(dst, value)?;
            }
            OxInst::WithEventsSet {
                dst,
                owner,
                binding,
                value,
            } => {
                let owner_value = self.operand(owner)?;
                let owner_ref = variant_to_object(&owner_value)?;
                let key = withevents_key(&owner_ref, *binding as i64);
                let v = self.operand(value)?;
                // Replacing a binding tears down its old host (COM) subscriptions first.
                self.unsubscribe_com_key(key);
                if is_nothing(&v) {
                    self.exec.events.withevents.remove(&key);
                } else {
                    // A COM/foreign source is wired through the host's connection points (the
                    // shared, live-tested HAL `subscribe_event`); a project source dispatches
                    // internally via `RaiseEvent` (no host subscription). Mirrors vm2's
                    // WithEventsSet.
                    if let Some(source) = v.as_object_ref()
                        && !source.is_project_instance()
                    {
                        self.advise_com_events(key, *binding, &owner_value, &source)?;
                    }
                    let order = self.exec.events.next_withevents_order;
                    self.exec.events.next_withevents_order =
                        self.exec.events.next_withevents_order.wrapping_add(1);
                    self.exec.events.withevents.insert(
                        key,
                        EventBinding {
                            owner: owner_value,
                            source: v.clone(),
                            order,
                        },
                    );
                }
                self.store(dst, v)?;
            }
            OxInst::WithEventsClearOwner { dst, owner } => {
                let owner_ref = variant_to_object(&self.operand(owner)?)?;
                let owner_raw = owner_ref.raw();
                self.unsubscribe_com_owner(owner_raw);
                self.exec
                    .events
                    .withevents
                    .retain(|key, _| withevents_owner_raw(*key) != owner_raw);
                self.store(dst, Variant::from_i32(0))?;
            }
            OxInst::WithEventsFirstOwner {
                dst,
                source,
                binding,
            } => {
                let source = self.operand(source)?;
                let mut owners: Vec<(u64, ObjectRef)> = Vec::new();
                if !is_nothing(&source) {
                    for (key, binding_data) in &self.exec.events.withevents {
                        if withevents_binding(*key) == (*binding as i64 & 0xFFFF_FFFF)
                            && object_identity(&binding_data.source) == object_identity(&source)
                            && let Some(owner) = binding_data.owner.as_object_ref()
                        {
                            owners.push((binding_data.order, owner));
                        }
                    }
                }
                owners.sort_unstable_by_key(|(order, _)| *order);
                let owners: Vec<ObjectRef> = owners.into_iter().map(|(_, owner)| owner).collect();
                match owners.first().cloned() {
                    Some(first) => {
                        self.exec.events.withevents_iters.push((owners, 1));
                        self.store(dst, Variant::from_object_ref(first))?;
                    }
                    None => self.store(dst, Variant::from_i32(0))?,
                }
            }
            OxInst::WithEventsNextOwner { dst } => {
                let next =
                    self.exec
                        .events
                        .withevents_iters
                        .last_mut()
                        .and_then(|(owners, pos)| {
                            let value = owners.get(*pos).cloned();
                            if value.is_some() {
                                *pos += 1;
                            }
                            value
                        });
                match next {
                    Some(owner) => self.store(dst, Variant::from_object_ref(owner))?,
                    None => {
                        self.exec.events.withevents_iters.pop();
                        self.store(dst, Variant::from_i32(0))?;
                    }
                }
            }
            OxInst::RaiseEvent {
                source,
                event,
                args,
            } => {
                let source_object = variant_to_object(&self.operand(source)?)?;
                let source_id = source_object.is_identity();
                let event_id = *event;
                // Collect subscribers whose binding holds this source and routes this event,
                // then run each handler in VBA subscription order with the sink as `me` and
                // the event args; an unhandled error propagates to the raiser.
                let mut targets: Vec<(u64, Variant, usize, usize)> = Vec::new();
                for (key, binding) in &self.exec.events.withevents {
                    if object_identity(&binding.source) != source_id {
                        continue;
                    }
                    let token = withevents_binding(*key) as i32;
                    // A sink's event routes + handler proc live in the SINK OWNER's program, which
                    // may differ from the raiser's `cur` for a cross-project `WithEvents` (project
                    // A sinks project B's source). Look the route up — and run the handler — there.
                    let owner_bundle = binding
                        .owner
                        .as_object_ref()
                        .map(|o| o.bundle_id() as usize)
                        .unwrap_or(self.cur);
                    if let Some(&handler) = self.exec.programs[owner_bundle]
                        .event_routes
                        .get(&(token, event_id))
                    {
                        targets.push((binding.order, binding.owner.clone(), handler, owner_bundle));
                    }
                }
                targets.sort_by_key(|(order, ..)| *order);
                for (_, sink, handler, owner_bundle) in targets {
                    self.run_proc_with_me(owner_bundle, FuncId(handler), sink, args)?;
                }
                // After internal WithEvents fan-out, deliver to advised COM connection-point
                // sinks directly on the source object (dc-yqq0.27), keeping `self.event_scope()`
                // active so sinks can call back into the object re-entrantly on this VM3 run.
                if source_object.has_connection_point_sinks() {
                    let values = self.extern_args(args)?;
                    let outcome = {
                        let _events = self.event_scope();
                        source_object.fire_connection_point_event(event_id, &values)
                    };
                    self.take_event_abort()?;
                    outcome.map_err(|msg| Vm3Error::Fault(Fault::new(5, msg)))?;
                }
                // After the internal WithEvents fan-out, deliver to the host event sink (W7):
                // the COM server forwards the event to its connection-point clients. Take the
                // sink out across the call so it does not alias `&mut self` (a host sink calls
                // back into the host, not the VM), then restore it.
                if let Some(sink) = self.reentrant_project_event_sink.clone() {
                    let values = self.extern_args(args)?;
                    let idle = Self::idle(self.exec.host);
                    let suspended = std::mem::replace(self, idle);
                    let (resumed, outcome) =
                        sink(suspended, source_object.clone(), event_id, values);
                    *self = resumed;
                    outcome.map_err(|msg| Vm3Error::Fault(Fault::new(5, msg)))?;
                } else if self.exec.events.project_event_sink.is_some() {
                    let values = self.extern_args(args)?;
                    let mut sink = self
                        .exec
                        .events
                        .project_event_sink
                        .take()
                        .expect("sink present");
                    let outcome = {
                        let _events = self.event_scope();
                        sink(source_object.clone(), event_id, values)
                    };
                    self.exec.events.project_event_sink = Some(sink);
                    self.take_event_abort()?;
                    outcome.map_err(|msg| Vm3Error::Fault(Fault::new(5, msg)))?;
                }
            }

            // ── Early-bound, descriptor-typed COM dispatch (M3-9) ────────────────────
            // Preserve the compiler's IID, slot and wire signature through HAL.
            // Native early calls never rediscover metadata or fall back to Invoke.
            // A receiver that is a project instance (a typed call through an `Implements`
            // interface) dispatches internally by the member's name.
            OxInst::ComCallEarly {
                dst,
                method,
                invoke_kind,
                recv,
                args,
            } => {
                let recv_v = self.operand(recv)?;
                let object = variant_to_object(&recv_v)?;
                let descriptor = self.cur_program().com_method(*method).ok_or_else(|| {
                    Vm3Error::Malformed(format!("ComCallEarly: unresolved method ref {method:?}"))
                })?;
                let ret = if object.is_project_instance() {
                    self.dispatch_project_method_resolved(
                        object,
                        recv_v,
                        Some(descriptor),
                        &descriptor.name,
                        *invoke_kind,
                        args,
                        false,
                    )?
                } else {
                    if descriptor.invoke_kind != *invoke_kind {
                        return Err(Vm3Error::Malformed("early COM accessor mismatch".into()));
                    }
                    self.dispatch_early_com_method(object, descriptor, args)?
                };
                if let Some(dst) = dst {
                    self.store(dst, ret)?;
                }
            }

            // ── Late-bound member dispatch (M3-6: project instances; M3-8: COM) ──────
            OxInst::ComCallLate {
                dst,
                recv,
                name,
                default_member,
                invoke_kind,
                args,
            } => {
                let recv_v = self.operand(recv)?;
                let ret = if *default_member {
                    self.dispatch_default_member(recv_v, *invoke_kind, args)?
                } else {
                    self.dispatch_member_by_name(
                        recv_v,
                        name,
                        *invoke_kind,
                        args,
                        !self.project_typed_receiver(recv),
                    )?
                };
                if let Some(dst) = dst {
                    self.store(dst, ret)?;
                }
            }
            // `CallByName(obj, "Member", vbMethod|vbGet|vbLet|vbSet, args…)` — the genuinely
            // dynamic by-name dispatch. The runtime `calltype` integer selects the accessor.
            OxInst::CallByName {
                dst,
                object,
                name,
                calltype,
                args,
            } => {
                let recv_v = self.operand(object)?;
                let member_name = arith::as_string(&self.operand(name)?);
                let ct = arith::int(&self.operand(calltype)?).map_err(arith_fault)?;
                let invoke_kind = match ct {
                    1 => TypeLibMemberInvokeKind::Method,
                    2 => TypeLibMemberInvokeKind::PropertyGet,
                    4 => TypeLibMemberInvokeKind::PropertyPut,
                    8 => TypeLibMemberInvokeKind::PropertyPutRef,
                    _ => {
                        return Err(Vm3Error::Fault(Fault::new(438, default_error_message(438))));
                    }
                };
                if recv_v.vtype() == VarType::Object && is_nothing(&recv_v) {
                    return Err(Vm3Error::Fault(Fault::new(5, default_error_message(5))));
                }
                if invoke_kind != TypeLibMemberInvokeKind::Method
                    && recv_v.as_object_ref().is_some_and(|object| {
                        !calcvba_lib::objects::is_object(&object)
                            && (object.native_collection_snapshot().is_some()
                                || object.is_project_instance()
                                    && object.route_key() == VBA_COLLECTION_ROUTE_KEY)
                    })
                {
                    return Err(Vm3Error::Fault(Fault::new(
                        438,
                        "Object doesn't support this property or method",
                    )));
                }
                let err_object = recv_v.as_object_ref();
                let ret = self
                    .dispatch_member_by_name(recv_v, &member_name, invoke_kind, args, true)
                    .map_err(|err| match (err, err_object.as_ref()) {
                        (Vm3Error::Fault(fault), Some(obj)) => {
                            Vm3Error::Fault(calcvba_rt_abi::normalize_call_by_name_err_fault(
                                obj,
                                &member_name,
                                fault,
                            ))
                        }
                        (err, _) => err,
                    })?;
                if let Some(dst) = dst {
                    self.store(dst, ret)?;
                }
            }

            // Everything else is a later milestone (cross-bundle calls / objects / COM /
            // arrays / records M3) — explicit, never a silent no-op.
            other => {
                return Err(Vm3Error::Unimplemented {
                    what: inst_kind(other),
                });
            }
        }
        Ok(())
    }

    /// Call a compiled VBA procedure: evaluate the arguments in the *caller* and push a
    /// callee frame (ByVal copies the value in; ByRef true-aliases the caller's backing
    /// location so writes propagate live; an omitted optional gets the `MISSING_ARG`
    /// sentinel), then hand control to the dispatch loop. The return value is copied out
    /// when the frame returns (see `do_return`). Mirrors vm2's `call_proc`.
    fn call_proc_in(
        &mut self,
        target_prog: usize,
        dst: Option<OxPlace>,
        proc: FuncId,
        args: &[OxArg],
    ) -> Result<(), Vm3Error> {
        // The callee runs against `target_prog` (its own program for a cross-project call); the
        // arguments + result dst are resolved below in the CALLER's context (`cur`), so a ByRef
        // alias / the dst stay bound to the caller's program (see `Loc::Global` tagging).
        let program = self.exec.programs[target_prog].program;
        let callee = program
            .funcs
            .get(proc.0)
            .ok_or_else(|| Vm3Error::Malformed(format!("call to unknown proc {}", proc.0)))?;
        // Resolve the destination + ByRef backings in the caller, before pushing.
        let dst_loc = dst.map(|p| self.resolve(&p));
        if self.cache.typed_entries[target_prog][proc.0].is_some()
            && args.len() == callee.param_count
            && args.iter().all(|arg| matches!(arg, OxArg::ByVal(_)))
        {
            self.guard_call_depth()?;
            let mut values =
                std::array::from_fn::<_, { calcvba_rt_abi::host_entry::MAX_TYPED_ARGS }, _>(|_| {
                    Variant::empty()
                });
            for (i, arg) in args.iter().enumerate() {
                let OxArg::ByVal(operand) = arg else {
                    unreachable!("checked ByVal")
                };
                values[i] = self.operand(operand)?;
            }
            if let Some(result) =
                self.invoke_scalar_values(target_prog, proc, &values[..args.len()])
            {
                let value = result?;
                if let Some(dst) = dst_loc {
                    self.write_loc(dst, value)?;
                }
                return Ok(());
            }
        }
        let mut frame = self.new_frame_pooled(target_prog, proc, args.len(), dst_loc, None)?;
        let frame_index = self.frames.len();
        let mut pending_param_array_aliases = Vec::new();
        for (i, arg) in args.iter().enumerate() {
            match arg {
                OxArg::ByVal(op) => {
                    let param_array_aliases = self.param_array_aliases_for_operand(op);
                    let v = self.operand(op)?;
                    if let Some(slot) = frame.local_mut(i) {
                        *slot = v;
                    }
                    if let Some(param_array_aliases) = param_array_aliases {
                        pending_param_array_aliases
                            .push((Loc::Local(frame_index, i), param_array_aliases));
                    }
                }
                OxArg::ByRef(place) => {
                    let param_ty = callee.locals.get(i).map(|local| &local.ty);
                    match self.byref_reference(place, param_ty)? {
                        Some(reference) => frame.slots[i] = reference,
                        None => {
                            frame.set_alias(i, self.resolve(place));
                        }
                    }
                }
                OxArg::Omitted => {
                    if let Some(slot) = frame.local_mut(i) {
                        *slot = match callee.locals.get(i).and_then(omitted_param_value) {
                            Some(value) => value?,
                            None => Variant::from_error_code(MISSING_ARG),
                        };
                    }
                }
            }
        }

        self.guard_call_depth()?;
        // Push the callee and hand control to it; each procedure starts with no active
        // handler, and the caller's mode is restored from the frame when it returns. The
        // dispatch loop runs the callee and `do_return`/`propagate_fault` pops it — there
        // is no native recursion here, so the call depth is heap-bounded.
        let saved_err = self.exec.err_engine.enter_activation();
        frame.saved_error_mode = saved_err.error_mode;
        frame.saved_active_error = saved_err.active_error;
        self.frames.push(frame);
        for (loc, aliases) in pending_param_array_aliases {
            self.param_array_aliases.insert(loc, aliases);
        }
        Ok(())
    }

    /// A cross-bundle call (`OxInst::CallExtern`). Resolve `import` to a native library
    /// function and run it through the same `invoke_native_lib` bridge a `CallNative`
    /// builtin uses — no frame is pushed (a `NativeBody::Library` body has no VM frame; its
    /// arguments are positional ByVal values), mirroring vm2's `call_extern` short-circuit
    /// and keeping a library function bit-identical however it is routed.
    fn call_extern(
        &mut self,
        dst: Option<OxPlace>,
        import: ImportId,
        args: &[OxArg],
    ) -> Result<(), Vm3Error> {
        let imp = self.cur_program().imports.get(import.0).ok_or_else(|| {
            Vm3Error::Malformed(format!("CallExtern names unknown import {}", import.0))
        })?;
        // A cross-PROJECT reference (a non-`VBA` unit) is a VBA-bodied proc in another loaded
        // program: resolve it by unit name + export token and call INTO that program — the
        // callee frame carries `prog = B` so it runs against B's globals, while the result dst
        // and any ByRef args are resolved in this caller's program (see `Loc::Global` tagging).
        // The synthetic `VBA` unit is the built-in native-library path below.
        if !imp.unit.eq_ignore_ascii_case("VBA") {
            let b = self.program_index_by_unit(&imp.unit).ok_or_else(|| {
                Vm3Error::Malformed(format!("unresolved reference to unit '{}'", imp.unit))
            })?;
            let proc = self.exec.programs[b]
                .program
                .exports
                .iter()
                .find(|e| e.token.matches(&imp.token))
                .ok_or_else(|| {
                    Vm3Error::Malformed(format!(
                        "unit '{}' has no export matching the CallExtern import",
                        imp.unit
                    ))
                })
                .and_then(|export| match export.target {
                    calcvba_bundle::ExportTarget::Proc(p) => Ok(p),
                    _ => Err(Vm3Error::Malformed(
                        "a cross-project CallExtern resolved to a non-procedure export".into(),
                    )),
                })?;
            return self.call_proc_in(b, dst, FuncId(proc), args);
        }
        let (id, string_typed_alias) = self.resolve_library_import(import)?;
        let args = calcvba_rt_abi::library::public_arguments(id, args);
        let mut argv = self.extern_args(&args)?;
        let result = match self.vba_object_len(id, &argv)? {
            Some(value) => value,
            None => {
                self.native_object_values(id, &mut argv)?;
                self.invoke_native_lib_with_policy(id, &argv, string_typed_alias)?
            }
        };
        if let Some(dst) = dst {
            self.store(&dst, result)?;
        }
        Ok(())
    }

    /// Resolve a cross-bundle `import` to the native library function it names.
    ///
    /// vm3 links the synthetic `VBA` library bundle
    /// (`calcvba_bundle::vba_library_bundle`) here: the home of every built-in
    /// function (`Strings.Left`, `Math.Abs`, the
    /// `DateTime`/`Conversion`/`Information`/`FileSystem` members, ...), which
    /// the binder lowers to a `CallExtern` rather than a `CallNative`. Non-`VBA`
    /// project references are handled by [`Self::call_extern`] before this helper
    /// is reached.
    fn resolve_library_import(&self, import: ImportId) -> Result<(NativeImplId, bool), Vm3Error> {
        let imp = self.cur_program().imports.get(import.0).ok_or_else(|| {
            Vm3Error::Malformed(format!("CallExtern names unknown import {}", import.0))
        })?;
        debug_assert!(imp.unit.eq_ignore_ascii_case("VBA"));
        let lib = calcvba_bundle::vba_library_bundle();
        let export = lib
            .exports
            .iter()
            .find(|e| e.token.matches(&imp.token))
            .ok_or_else(|| {
                Vm3Error::Malformed(format!(
                    "the VBA library bundle has no export matching import {}",
                    import.0
                ))
            })?;
        let calcvba_bundle::ExportTarget::Proc(proc) = export.target else {
            return Err(Vm3Error::Malformed(
                "a VBA library import resolved to a non-procedure export".into(),
            ));
        };
        match lib.procedures.get(proc).and_then(|p| p.native) {
            Some(calcvba_bundle::NativeBody::Library(id)) => {
                let string_typed_alias = match &imp.token {
                    calcvba_bundle::ExportToken::ModuleFunc { member, .. } => {
                        id.is_string_typed_library_alias(member)
                    }
                    _ => false,
                };
                Ok((id, string_typed_alias))
            }
            Some(calcvba_bundle::NativeBody::Method(_)) => Err(Vm3Error::Malformed(
                "a native object method is not callable via CallExtern".into(),
            )),
            None => Err(Vm3Error::Malformed(
                "a VBA library export has no native library body".into(),
            )),
        }
    }

    /// Marshal a cross-bundle library call's arguments to plain values: a native library
    /// body reads positional values (a ByRef argument by its *value*), and an omitted
    /// optional is `Empty` — matching vm2's `extern_native_args`.
    fn extern_args(&mut self, args: &[OxArg]) -> Result<Vec<Variant>, Vm3Error> {
        marshal_ox_args(
            args,
            |arg| match arg {
                MarshalArgRef::Operand(op) => self.operand(op),
                MarshalArgRef::ByRef(place) => self.read(place),
            },
            Variant::empty,
        )
    }

    /// Build SAFEARRAY bounds from `ReDim` upper/lower-bound operands. Each
    /// bound is a VBA `Long`, so a value outside `Long` range raises Overflow
    /// (6) instead of silently wrapping through `as i32`; `upper < lower` →
    /// subscript out of range (9); and a single dimension above `u32::MAX`
    /// elements → out of memory (7). These guard the bound *shape* only — the
    /// actual element store is allocated fallibly (the zeroed-scalar path and
    /// [`try_build_default_elements`]), so an over-large array raises catchable
    /// error 7 rather than aborting the host on an infallible allocation.
    fn build_bounds(
        &mut self,
        upper_bounds: &[OxOperand],
        lower_bounds: &[OxOperand],
    ) -> Result<Vec<SafeArrayBound>, Vm3Error> {
        let mut bounds = Vec::with_capacity(upper_bounds.len());
        for (i, upper_op) in upper_bounds.iter().enumerate() {
            let lower = if let Some(lower_op) = lower_bounds.get(i) {
                let lower_v = self.operand(lower_op)?;
                array_subscript(&lower_v).map_err(arith_fault)?
            } else {
                0
            };
            let upper_v = self.operand(upper_op)?;
            let upper = array_subscript(&upper_v).map_err(arith_fault)?;
            if upper < lower {
                return Err(Vm3Error::Fault(Fault::standard(9)));
            }
            let span = i64::from(upper) - i64::from(lower) + 1;
            if span > i64::from(u32::MAX) {
                return Err(Vm3Error::Fault(Fault::new(
                    7,
                    format!("array dimension too large ({span} elements)"),
                )));
            }
            bounds.push(SafeArrayBound {
                count: span as u32,
                lower,
            });
        }
        Ok(bounds)
    }

    /// Storage index from VBA (absolute) subscript operands, first dimension fastest as VBA
    /// stores arrays ([`calcvba_runtime::safe_array::element_index`]), bounds-checked →
    /// subscript out of range (9).
    fn flat_index(
        &mut self,
        indices: &[OxOperand],
        bounds: &[SafeArrayBound],
    ) -> Result<usize, Vm3Error> {
        if indices.len() != bounds.len() {
            return Err(Vm3Error::Fault(Fault::standard(9)));
        }
        if let ([index_op], [bound]) = (indices, bounds) {
            let index_v = self.operand(index_op)?;
            let raw = array_subscript(&index_v).map_err(arith_fault)?;
            let offset = i64::from(raw) - i64::from(bound.lower);
            if offset < 0 || offset >= i64::from(bound.count) {
                return Err(Vm3Error::Fault(Fault::standard(9)));
            }
            return Ok(offset as usize);
        }
        let mut flat = 0usize;
        let mut stride = 1usize;
        for (index_op, bound) in indices.iter().zip(bounds) {
            let index_v = self.operand(index_op)?;
            let raw = array_subscript(&index_v).map_err(arith_fault)?;
            let offset = i64::from(raw) - i64::from(bound.lower);
            if offset < 0 || offset >= i64::from(bound.count) {
                return Err(Vm3Error::Fault(Fault::standard(9)));
            }
            flat += offset as usize * stride;
            stride *= bound.count as usize;
        }
        Ok(flat)
    }

    /// Fast O(1) `arr(i…)` read: when the receiver is a place-resident array, read
    /// the single element directly through the SAFEARRAY descriptor without cloning
    /// the (possibly large) backing store — so an array loop is O(N), not O(N²).
    /// Returns `Ok(None)` when the receiver is not a place-resident array (a `Const`,
    /// an unwritten `Temp`, an object default-member receiver, …), so the caller runs
    /// the general path.
    fn array_get_fast(
        &mut self,
        array: &OxOperand,
        indices: &[OxOperand],
    ) -> Result<Option<Variant>, Vm3Error> {
        let OxOperand::Use(place) = array else {
            return Ok(None);
        };
        let loc = self.resolve(place);
        let Some(arr) = self.read_loc_ref(loc)? else {
            return Ok(None);
        };
        // Only an in-place array takes the fast path; objects/other receivers fall back.
        let Some((bounds, len)) = arr.safearray_bounds_len() else {
            return Ok(None);
        };
        let flat = self.flat_index(indices, &bounds)?;
        if flat >= len {
            return Err(Vm3Error::Fault(Fault::standard(9)));
        }
        let Some(arr) = self.read_loc_ref(loc)? else {
            return Ok(None);
        };
        if let Some(result) = arr.safearray_i32_element(flat) {
            match result {
                Ok(Some(value)) => return Ok(Some(Variant::from_i32(value))),
                Ok(None) => {}
                Err(err) => return Err(Vm3Error::Fault(Fault::new(13, err))),
            }
        }
        let value = arr
            .safearray_element(flat)
            .expect("safearray_bounds_len already proved this is an array")
            .map_err(|e| Vm3Error::Fault(Fault::new(13, e)))?;
        Ok(Some(value))
    }

    /// Fast O(1) `arr(i…) = v` write: when the receiver is a place-resident array,
    /// mutate the single element through the SAFEARRAY descriptor in place — no
    /// per-write deep-clone-and-write-back of the whole array. A ByRef-aliased array
    /// is mutated through its resolved (aliased) location, so callers see the change.
    /// Returns `Ok(false)` when the receiver is not a place-resident array.
    fn array_set_fast(
        &mut self,
        array: &OxPlace,
        indices: &[OxOperand],
        value: &Variant,
    ) -> Result<bool, Vm3Error> {
        let loc = self.resolve(array);
        // Bounds + flat index under an immutable borrow first (reads index operands),
        // then take the mutable borrow to write the element in place.
        let (bounds, len) = match self.read_loc_ref(loc)? {
            Some(arr) => match arr.safearray_bounds_len() {
                Some(bl) => bl,
                None => return Ok(false),
            },
            None => return Ok(false),
        };
        let flat = self.flat_index(indices, &bounds)?;
        if flat >= len {
            return Err(Vm3Error::Fault(Fault::standard(9)));
        }
        if self.mirror_param_array_element_write(loc, flat, value)? {
            return Ok(true);
        }
        let arr = self
            .read_loc_mut(loc)?
            .expect("location borrowed immutably just above is still present");
        if let Some(value_i32) = value.as_i32()
            && let Some(result) = arr.set_safearray_i32_element(flat, value_i32)
        {
            match result {
                Ok(true) => {
                    return Ok(true);
                }
                Ok(false) => {}
                Err(err) => return Err(Vm3Error::Fault(Fault::new(13, err))),
            }
        }
        arr.set_safearray_element(flat, value)
            .map_err(|e| Vm3Error::Fault(Fault::new(13, e)))?;
        Ok(true)
    }

    fn param_array_aliases_for_operand(&self, op: &OxOperand) -> Option<Vec<Option<Loc>>> {
        match op {
            OxOperand::Use(place) => self.param_array_aliases.get(&self.resolve(place)).cloned(),
            _ => None,
        }
    }

    fn mirror_param_array_element_write(
        &mut self,
        array_loc: Loc,
        flat: usize,
        value: &Variant,
    ) -> Result<bool, Vm3Error> {
        let Some(aliases) = self.param_array_aliases.get(&array_loc).cloned() else {
            return Ok(false);
        };
        let Some(Some(target)) = aliases.get(flat).copied() else {
            return Ok(false);
        };
        let ty = match target {
            Loc::Global(p, g) => self.exec.programs[p].program.globals[g].ty.clone(),
            Loc::Local(f, l) => {
                let frame = &self.frames[f];
                self.exec.programs[frame.prog].program.funcs[frame.func.0].locals[l]
                    .ty
                    .clone()
            }
            Loc::Temp(f, t) => {
                let frame = &self.frames[f];
                self.exec.programs[frame.prog].program.funcs[frame.func.0].temps[t].clone()
            }
        };
        let (converted, element) =
            calcvba_rt_abi::param_array_write_values(value, &ty).map_err(arith_fault)?;
        self.write_loc(target, converted)?;

        if let Some(array) = self.read_loc_mut(array_loc)? {
            calcvba_rt_abi::write_param_array_aliases(array, &aliases, target, &element)
                .map_err(Vm3Error::Fault)?;
        }
        Ok(true)
    }

    fn prune_param_array_aliases_from_depth(&mut self, depth: usize) {
        if !self.param_array_aliases.is_empty() {
            self.param_array_aliases.retain(|loc, aliases| {
                loc.frame_index().is_none_or(|frame| frame < depth)
                    && aliases
                        .iter()
                        .flatten()
                        .all(|alias| alias.frame_index().is_none_or(|frame| frame < depth))
            });
        }
        if !self.as_new_slots.is_empty() {
            self.as_new_slots
                .retain(|loc, _| loc.frame_index().is_none_or(|frame| frame < depth));
        }
    }

    fn record_array_get(
        &mut self,
        record: &OxOperand,
        index: usize,
        indices: &[OxOperand],
    ) -> Result<Variant, Vm3Error> {
        let OxOperand::Use(place) = record else {
            return Err(Vm3Error::Fault(Fault::new(
                13,
                "record array field requires a place",
            )));
        };
        let loc = self.resolve(place);
        let Some(record) = self.read_loc_ref(loc)? else {
            return Err(Vm3Error::Fault(Fault::new(13, "expected Record Variant")));
        };
        let Some((bounds, len)) = record
            .record_array_field_bounds_len(index)
            .map_err(|e| Vm3Error::Fault(Fault::new(13, e)))?
        else {
            return Err(Vm3Error::Fault(Fault::new(13, "Expected array")));
        };
        let flat = self.flat_index(indices, &bounds)?;
        if flat >= len {
            return Err(Vm3Error::Fault(Fault::standard(9)));
        }
        let Some(record) = self.read_loc_ref(loc)? else {
            return Err(Vm3Error::Fault(Fault::new(13, "expected Record Variant")));
        };
        record
            .record_array_field_element(index, flat)
            .map_err(|e| Vm3Error::Fault(Fault::new(13, e)))?
            .ok_or_else(|| Vm3Error::Fault(Fault::new(13, "Expected array")))
    }

    fn record_array_set(
        &mut self,
        record: &OxPlace,
        index: usize,
        indices: &[OxOperand],
        value: &Variant,
    ) -> Result<(), Vm3Error> {
        let loc = self.resolve(record);
        let Some(record) = self.read_loc_ref(loc)? else {
            return Err(Vm3Error::Fault(Fault::new(13, "expected Record Variant")));
        };
        let Some((bounds, len)) = record
            .record_array_field_bounds_len(index)
            .map_err(|e| Vm3Error::Fault(Fault::new(13, e)))?
        else {
            return Err(Vm3Error::Fault(Fault::new(13, "Expected array")));
        };
        let flat = self.flat_index(indices, &bounds)?;
        if flat >= len {
            return Err(Vm3Error::Fault(Fault::standard(9)));
        }
        let record = self
            .read_loc_mut(loc)?
            .expect("location borrowed immutably just above is still present");
        record
            .set_record_array_field_element(index, flat, value)
            .map_err(|e| Vm3Error::Fault(Fault::new(13, e)))?
            .ok_or_else(|| Vm3Error::Fault(Fault::new(13, "Expected array")))
    }

    fn is_unallocated_array_value(value: &Variant) -> bool {
        value.vtype() == VarType::ArrayVariant && value.safearray_bounds_len().is_none()
    }

    /// The array (SAFEARRAY) value of an operand, else type mismatch (13).
    /// A null array marker is still an array value in VBA, but has no descriptor/bounds;
    /// indexing and bounds queries surface subscript out of range (9).
    fn array_of(&mut self, op: &OxOperand) -> Result<SafeArray, Vm3Error> {
        let value = self.operand(op)?;
        if Self::is_unallocated_array_value(&value) {
            return Err(Vm3Error::Fault(Fault::new(9, "array has no bounds")));
        }
        value
            .as_safearray()
            .ok_or_else(|| Vm3Error::Fault(Fault::new(13, "expected an array")))
    }

    /// Index a received receiver VALUE `recv` and store the element into `dst`. The
    /// receiver is either an array (element read) or a built-in `Collection` object
    /// whose default member `Item` is indexed (`c(i)`). Shared by the `ArrayGet`
    /// non-fast path and the `FieldArrayGet` non-array-field fallback.
    fn index_value_into(
        &mut self,
        recv: Variant,
        indices: &[OxOperand],
        dst: &OxPlace,
    ) -> Result<(), Vm3Error> {
        let arr = match recv.as_safearray() {
            Some(arr) => arr,
            None if Self::is_unallocated_array_value(&recv) => {
                return Err(Vm3Error::Fault(Fault::new(9, "array has no bounds")));
            }
            None => {
                if let Some(obj) = recv.as_object_ref() {
                    let argv = self.operands_to_values(indices)?;
                    let value = self.dispatch_default_member_values(
                        Variant::from_object_ref(obj),
                        TypeLibMemberInvokeKind::PropertyGet,
                        argv,
                    )?;
                    self.store(dst, value)?;
                    return Ok(());
                }
                return Err(Vm3Error::Fault(Fault::new(13, "expected an array")));
            }
        };
        let bounds = arr
            .bounds()
            .ok_or_else(|| Vm3Error::Fault(Fault::new(9, "array has no bounds")))?;
        let flat = self.flat_index(indices, &bounds)?;
        if flat >= arr.len() {
            return Err(Vm3Error::Fault(Fault::standard(9)));
        }
        if let Some(result) = recv.safearray_i32_element(flat) {
            match result {
                Ok(Some(value)) => return self.store(dst, Variant::from_i32(value)),
                Ok(None) => {}
                Err(err) => return Err(Vm3Error::Fault(Fault::new(13, err))),
            }
        }
        let value = arr
            .variant_element(flat)
            .map_err(|e| Vm3Error::Fault(Fault::new(13, e)))?;
        self.store(dst, value)
    }

    /// Set one element of a received array VALUE `recv` and return the mutated array
    /// for the caller to write back where `recv` came from. Errors if `recv` is an
    /// object (default-member assignment unsupported) or not an array. Shared by the
    /// `ArraySet` non-fast path and the `FieldArraySet` non-array-field fallback.
    fn set_index_in_value(
        &mut self,
        mut recv: Variant,
        indices: &[OxOperand],
        value: &Variant,
    ) -> Result<Variant, Vm3Error> {
        let arr = match recv.as_safearray() {
            Some(arr) => arr,
            None if Self::is_unallocated_array_value(&recv) => {
                return Err(Vm3Error::Fault(Fault::new(9, "array has no bounds")));
            }
            None => {
                if let Some(obj) = recv.as_object_ref() {
                    let mut argv = self.operands_to_values(indices)?;
                    argv.push(value.clone());
                    let invoke_kind = if value.vtype() == VarType::Object {
                        TypeLibMemberInvokeKind::PropertyPutRef
                    } else {
                        TypeLibMemberInvokeKind::PropertyPut
                    };
                    self.dispatch_default_member_values(
                        Variant::from_object_ref(obj),
                        invoke_kind,
                        argv,
                    )?;
                    return Ok(recv);
                }
                return Err(Vm3Error::Fault(Fault::new(13, "expected an array")));
            }
        };
        let bounds = arr
            .bounds()
            .ok_or_else(|| Vm3Error::Fault(Fault::new(9, "array has no bounds")))?;
        let flat = self.flat_index(indices, &bounds)?;
        if flat >= arr.len() {
            return Err(Vm3Error::Fault(Fault::standard(9)));
        }
        if let Some(value_i32) = value.as_i32()
            && let Some(result) = recv.set_safearray_i32_element(flat, value_i32)
        {
            match result {
                Ok(true) => return Ok(recv),
                Ok(false) => {}
                Err(err) => return Err(Vm3Error::Fault(Fault::new(13, err))),
            }
        }
        recv.set_safearray_element(flat, value)
            .map_err(|e| Vm3Error::Fault(Fault::new(13, e)))?;
        Ok(recv)
    }

    /// The 0-based dimension index for `LBound`/`UBound` from an optional dimension operand
    /// (default dimension 1), validated against the array's rank → subscript out of range (9).
    fn array_bound_index(
        &mut self,
        dimension: Option<&OxOperand>,
        bounds: &[SafeArrayBound],
    ) -> Result<usize, Vm3Error> {
        let dim = match dimension {
            Some(op) => {
                let v = self.operand(op)?;
                arith::int(&v).map_err(arith_fault)?
            }
            None => 1,
        };
        if dim < 1 {
            return Err(Vm3Error::Fault(Fault::standard(9)));
        }
        let index = (dim - 1) as usize;
        if index >= bounds.len() {
            return Err(Vm3Error::Fault(Fault::standard(9)));
        }
        Ok(index)
    }

    /// `ReDim [Preserve]`: build the new SAFEARRAY shaped by `upper_bounds`/`lower_bounds`,
    /// seeding each element with the declared element type's typed default — except, when
    /// `preserve`, keeping each still-in-range existing element (so a UDT array's populated
    /// records survive and only the grown tail is freshly default-seeded). The element storage
    /// matches the declared element type (typed scalars / native records, not normalized to
    /// VT_VARIANT).
    fn array_redim(
        &mut self,
        dst: &OxPlace,
        upper_bounds: &[OxOperand],
        lower_bounds: &[OxOperand],
        element: &ArrayElementType,
        policy: (bool, bool, bool),
    ) -> Result<(), Vm3Error> {
        let (preserve, fixed, explicit_element) = policy;
        let frame = self.frames.last().expect("array instruction frame");
        let program = self.exec.programs[frame.prog].program;
        let function = &program.funcs[frame.func.0];
        let ty = match dst {
            OxPlace::Local(id) => &function.locals[id.0].ty,
            OxPlace::Temp(id) => &function.temps[id.0],
            OxPlace::Global(id) => &program.globals[id.0].ty,
        };
        let inherit_current = !explicit_element && matches!(ty, OxTy::Variant);
        // A plain `ReDim` of a Variant gets a new descriptor; an array variable, or `Preserve`,
        // keeps it (Excel, probe redim7c).
        let keep_descriptor = preserve || !matches!(ty, OxTy::Variant);
        let cur_value = self.read(dst)?;
        let selected = calcvba_bundle::array_runtime::resolve_redim_element(
            &cur_value,
            element,
            inherit_current,
            preserve,
        )
        .map_err(|code| Vm3Error::Fault(Fault::new(code, "ReDim element type mismatch")))?;
        let element = &selected;
        let nominal_record_layout = if let ArrayElementType::Record(fields) = element {
            let anon = vba_record_layout_for_fields(fields).ok();
            let from_decl = match ty {
                OxTy::Array(elem, _) => match elem.as_ref() {
                    OxTy::Record(id) => {
                        calcvba_rt_abi::vba_record_layout_for_program_record(program, *id).ok()
                    }
                    _ => None,
                },
                _ => None,
            };
            from_decl
                .or_else(|| cur_value.record_array_layout())
                .filter(|layout| anon.as_ref().is_none_or(|a| layout.matches_structural(a)))
        } else {
            None
        };
        let build_defaults = |count: usize| -> Result<Vec<Variant>, Vm3Error> {
            if let Some(ref rec_layout) = nominal_record_layout {
                let mut out = Vec::new();
                out.try_reserve_exact(count)
                    .map_err(|_| Vm3Error::Fault(Fault::new(7, "Out of memory")))?;
                for _ in 0..count {
                    out.push(
                        calcvba_bundle::array_runtime::default_record_array_element(rec_layout)
                            .map_err(|e| Vm3Error::Fault(Fault::new(13, e)))?,
                    );
                }
                Ok(out)
            } else {
                try_build_default_elements(element, count)
            }
        };
        // A user `ReDim` (lowered with `fixed = false`) of an already-fixed-size array is
        // illegal in VBA — `Dim a(1 To 3) : ReDim a(...)` is the compile error "Array already
        // dimensioned". vm3 carries fixed-ness only on the runtime value (the symbol model
        // can't tell a fixed top-level `Dim` from a dynamic one), so the faithful surfacing is
        // the runtime analog, error 10 "This array is fixed or temporarily locked". The
        // fixed-`Dim`/UDT-field allocation itself is `fixed = true` against an uninitialized
        // slot, so it is never caught here.
        if !fixed && cur_value.as_safearray().is_some_and(|a| a.is_fixed_size()) {
            return Err(Vm3Error::Fault(Fault::new(
                10,
                "This array is fixed or temporarily locked",
            )));
        }
        let bounds = self.build_bounds(upper_bounds, lower_bounds)?;
        // Total element count, guarding the cross-dimension product (`build_bounds` already
        // guards each single dimension's span) → out of memory (7) rather than a usize wrap
        // that would feed a bogus allocation.
        let mut count = 1usize;
        for b in &bounds {
            count = count
                .checked_mul(b.count as usize)
                .ok_or_else(|| Vm3Error::Fault(Fault::new(7, "array too large to allocate")))?;
        }
        if !preserve && Self::element_supports_zeroed_redim(element) {
            let array = SafeArray::from_zeroed_typed_scalars_nd(
                bounds,
                safearray_vartype_for_element(element),
            )
            .map(|array| array.with_fixed_size(fixed))
            .map_err(|_| Vm3Error::Fault(Fault::new(7, "array allocation failed")))?;
            return self.store_redimmed(dst, array, keep_descriptor);
        }
        let elems: Vec<Variant> = if preserve {
            match cur_value.as_safearray().and_then(|a| {
                a.bounds()
                    .map(|b| (b, a.variant_elements().unwrap_or_default()))
            }) {
                // `ReDim Preserve` over an existing allocation: VBA permits changing ONLY the
                // last dimension's upper bound. Changing the rank, any earlier dimension, or
                // the last dimension's lower bound is subscript out of range (9). Surviving
                // elements are preserved BY COORDINATE (not by flat position), so growing a
                // rank-2 array keeps `a(i, j)` in place.
                Some((old_bounds, old_elems)) => {
                    let rank = bounds.len();
                    let illegal = old_bounds.len() != rank
                        || (0..rank.saturating_sub(1)).any(|i| old_bounds[i] != bounds[i])
                        || match (old_bounds.last(), bounds.last()) {
                            (Some(o), Some(n)) => o.lower != n.lower,
                            _ => false,
                        };
                    if illegal {
                        return Err(Vm3Error::Fault(Fault::standard(9)));
                    }
                    let mut out = build_defaults(count)?;
                    for (old_flat, value) in old_elems.into_iter().enumerate() {
                        if let Some(new_flat) = calcvba_runtime::safe_array::preserved_element_index(
                            old_flat,
                            &old_bounds,
                            &bounds,
                        ) {
                            out[new_flat] = value;
                        }
                    }
                    out
                }
                // `ReDim Preserve` of a not-yet-allocated dynamic array just allocates it.
                None => build_defaults(count)?,
            }
        } else {
            build_defaults(count)?
        };
        let array = if let Some(rec_layout) = nominal_record_layout {
            calcvba_bundle::array_runtime::redim_record_safearray_from_elements(
                bounds, rec_layout, elems, fixed,
            )
        } else {
            redim_safearray_from_elements(bounds, element, elems, fixed)
        }
        .map_err(|e| Vm3Error::Fault(Fault::new(13, e)))?;
        self.store_redimmed(dst, array, keep_descriptor)
    }

    /// Store a `ReDim` result. With `keep_descriptor` the new storage moves into the array the
    /// place already holds (or refers to), so its `SAFEARRAY*` stays where it is, as VBA's
    /// `SafeArrayRedim`; otherwise, or when there is no matching descriptor, the place gets a
    /// new array value.
    fn store_redimmed(
        &mut self,
        dst: &OxPlace,
        array: SafeArray,
        keep_descriptor: bool,
    ) -> Result<(), Vm3Error> {
        let loc = self.resolve(dst);
        let locked = |number: i32| {
            Vm3Error::Fault(Fault::new(
                number,
                calcvba_rt_abi::default_error_message(number),
            ))
        };
        if self
            .read_loc_ref(loc)?
            .is_some_and(Variant::holds_locked_array)
        {
            return Err(locked(10));
        }
        let array = if keep_descriptor {
            match self.read_loc_mut(loc)? {
                Some(cell) => match cell.redim_in_place(array) {
                    Ok(old) => {
                        drop(old);
                        return Ok(());
                    }
                    Err(array) => array,
                },
                None => array,
            }
        } else {
            array
        };
        self.store(dst, Variant::from_safearray(array))
    }

    fn element_supports_zeroed_redim(element: &ArrayElementType) -> bool {
        matches!(
            element,
            ArrayElementType::Integer
                | ArrayElementType::Long
                | ArrayElementType::LongLong
                | ArrayElementType::Byte
                | ArrayElementType::Single
                | ArrayElementType::Double
                | ArrayElementType::Currency
                | ArrayElementType::Date
                | ArrayElementType::Boolean
        )
    }

    /// `Let`/`Set` legality check (mirrors vm2). `Set` requires an object source (else
    /// "Object required" 424); `Let` into an `Object` target requires `Set` when the source
    /// is an object (91) and an object source otherwise (424). The strict `Set` *type* check
    /// (error 13 — a project-instance source must be the target's declared class/interface)
    /// needs the project class tables and lands with the object model (M3-5); until a project
    /// instance can exist, every Set-of-object is a COM/`Nothing` value, which vm2 also
    /// passes — so falling through to `Ok` here is behaviorally exact.
    fn validate_assignment(
        &mut self,
        src: &OxOperand,
        intent: AssignmentIntent,
        target_kind: AssignmentTargetKind,
        target_name: &str,
        target_type_name: &str,
    ) -> Result<(), Vm3Error> {
        use AssignmentIntent as Intent;
        use AssignmentTargetKind as Kind;
        let value = self.operand(src)?;
        // `Nothing` is already a null-object `Variant` (VarType::Object), so the
        // object test is just the type — the old `|| is_nothing(&value)` also
        // matched Empty/Null and numeric 0 (is_nothing's stale scalar-zero
        // sentinel), which let `Set o = 0` silently store a scalar instead of
        // raising "Object required" (424).
        let is_object = matches!(value.vtype(), VarType::Object);
        match intent {
            // A non-object Variant `Set` into a Variant is a Type mismatch; into an object
            // variable (or from a non-Variant) it is "Object required" (Excel).
            Intent::Set if !is_object => Err(Vm3Error::Fault(
                if matches!(src, OxOperand::Use(_))
                    && self.operand_is_variant(src)
                    && target_kind == Kind::Variant
                {
                    Fault::new(13, "Type mismatch")
                } else {
                    Fault::new(424, "Object required")
                },
            )),
            Intent::Let
                if target_kind == Kind::Variant
                    && value.vtype() == VarType::Object
                    && is_nothing(&value) =>
            {
                Err(Vm3Error::Fault(Fault::new(
                    91,
                    "Object variable or With block variable not set",
                )))
            }
            Intent::Let if target_kind == Kind::Object && is_object => Err(Vm3Error::Fault(
                Fault::new(91, format!("Object variable requires Set: {target_name}")),
            )),
            Intent::Let if target_kind == Kind::Object => {
                Err(Vm3Error::Fault(Fault::new(424, "Object required")))
            }
            Intent::Set if value.vtype() == VarType::Object && is_nothing(&value) => Ok(()),
            // Strict `Set` type check (error 13): a project class/interface target takes only
            // an instance of that class (or implementing it), a `Collection` target only a
            // Collection; `Object`/`Variant` and COM interface targets are not checked here.
            Intent::Set if value.vtype() == VarType::Object && !target_type_name.is_empty() => {
                let obj = variant_to_object(&value)?;
                let source = obj
                    .is_project_instance()
                    .then(|| self.exec.programs.get(obj.bundle_id() as usize))
                    .flatten()
                    .map(|lp| lp.program);
                if calcvba_rt_abi::object_types::set_type_mismatch(
                    &obj,
                    target_type_name,
                    source,
                    self.exec.programs.iter().map(|lp| lp.program),
                ) {
                    return Err(Vm3Error::Fault(Fault::new(
                        13,
                        format!("Type mismatch: cannot be assigned to `{target_type_name}`"),
                    )));
                }
                Ok(())
            }
            _ => Ok(()),
        }
    }

    /// Allocate a fresh project-class instance (`New <Class>`): a refcounted IUnknown with a
    /// unique identity, then run its `Class_Initialize` (if any). Mirrors vm2's `Op::NewObject`
    /// (bundle id is always 0 — vm3 runs one program). The instance's `has_terminate` flag is
    /// what later parks it for `Class_Terminate` when its last reference drops.
    fn new_project_instance(&mut self, class_idx: usize) -> Result<Variant, Vm3Error> {
        let descriptor = *self.exec.programs[self.cur]
            .class_descriptors
            .get(class_idx)
            .ok_or_else(|| Vm3Error::Malformed(format!("unknown class {class_idx}")))?;
        let program = self.cur_program();
        let class = program
            .classes
            .get(class_idx)
            .ok_or_else(|| Vm3Error::Malformed(format!("unknown class {class_idx}")))?;
        let has_terminate = class.terminate.is_some();
        let initialize = class.initialize;
        let instance_id = self.exec.next_instance_id;
        self.exec.next_instance_id += 1;
        let object = ObjectRef::from_project_instance(
            instance_id,
            class_idx as i32,
            self.cur as i32,
            has_terminate,
            descriptor,
        );
        let value = Variant::from_object_ref(object.clone());
        if let Some(init) = initialize {
            let result = self.run_proc_with_me(
                self.cur,
                init,
                Variant::from_object_ref(object.clone()),
                &[],
            );
            if result.is_err() {
                object.mark_initialization_failed();
            }
            result?;
        }
        Ok(value)
    }

    /// A `VB_PredeclaredId` auto-instance (`Class1.Foo` with no explicit `New`): allocate the
    /// singleton + run `Class_Initialize` on first access, then reuse the cached instance.
    /// Mirrors vm2's `predeclared_instance`.
    fn predeclared_instance(&mut self, class_idx: usize) -> Result<Variant, Vm3Error> {
        if let Some(existing) = self.exec.programs[self.cur]
            .predeclared_singletons
            .get(&class_idx)
        {
            return Ok(existing.clone());
        }
        let descriptor = *self.exec.programs[self.cur]
            .class_descriptors
            .get(class_idx)
            .ok_or_else(|| Vm3Error::Malformed(format!("unknown class {class_idx}")))?;
        let program = self.cur_program();
        let class = program
            .classes
            .get(class_idx)
            .ok_or_else(|| Vm3Error::Malformed(format!("unknown class {class_idx}")))?;
        let has_terminate = class.terminate.is_some();
        let initialize = class.initialize;
        let instance_id = self.exec.next_instance_id;
        self.exec.next_instance_id += 1;
        let object = ObjectRef::from_project_instance(
            instance_id,
            class_idx as i32,
            self.cur as i32,
            has_terminate,
            descriptor,
        );
        calcvba_rt_abi::attach_host_document(self.exec.host, class, &object)
            .map_err(Vm3Error::Fault)?;
        let value = Variant::from_object_ref(object);
        self.exec.programs[self.cur]
            .predeclared_singletons
            .insert(class_idx, value.clone());
        if let Some(init) = initialize
            && let Err(err) = self.run_proc_with_me(self.cur, init, value.clone(), &[])
        {
            if let Some(object) = value.as_object_ref() {
                object.mark_initialization_failed();
            }
            self.exec.programs[self.cur]
                .predeclared_singletons
                .remove(&class_idx);
            return Err(err);
        }
        Ok(value)
    }

    /// Assign the storage slot behind a `VB_PredeclaredId` class name. Real VBA treats
    /// `Set ClassName = Nothing` as clearing the default-instance slot, and a later
    /// `ClassName.Member` access creates a fresh default instance. Assigning a non-Nothing
    /// object replaces the slot with that object; the RHS has already been evaluated by the
    /// caller, preserving VBA's `New`-then-release-old ordering.
    fn set_predeclared_instance(
        &mut self,
        program_index: usize,
        class_idx: usize,
        value: Variant,
    ) -> Result<(), Vm3Error> {
        let program = self
            .exec
            .programs
            .get_mut(program_index)
            .ok_or_else(|| Vm3Error::Malformed(format!("unknown program {program_index}")))?;
        if class_idx >= program.program.classes.len() {
            return Err(Vm3Error::Malformed(format!("unknown class {class_idx}")));
        }
        if is_nothing(&value) {
            program.predeclared_singletons.remove(&class_idx);
        } else {
            program.predeclared_singletons.insert(class_idx, value);
        }
        Ok(())
    }

    /// Run a procedure to completion **synchronously** with `me` as its hidden first local —
    /// the lifecycle/event entry point (`Class_Initialize`/`Class_Terminate`/event handlers).
    ///
    /// vm3's dispatch is an explicit loop, so this pushes the callee frame and drives a NESTED
    /// `run_loop(base)` that returns when this frame returns. `run_loop` breaks at
    /// `frames.len() == base + 1` on a normal `Return` *without* popping, and leaves the
    /// faulting frame in place on an uncaught fault — so afterwards we restore the caller's
    /// error state (saved on the frame we pushed) and truncate back to `base` either way.
    fn run_proc_with_me(
        &mut self,
        target_prog: usize,
        proc: FuncId,
        me: Variant,
        args: &[OxArg],
    ) -> Result<Variant, Vm3Error> {
        self.guard_call_depth()?;
        // Preserve the caller's program across the nested run_loop: it leaves `cur` at the
        // callee's program (run_loop re-derives cur per iteration), so the caller's post-call
        // resolves (store dst / extern_args) would otherwise use the wrong program. The caller
        // sets `cur` to the receiver's program (object dispatch / terminate) before calling.
        let saved_cur = self.cur;
        let base = self.frames.len();
        let mut frame = self.new_frame_in(target_prog, proc)?;
        if let Some(slot) = frame.local_mut(0) {
            *slot = me;
        }
        let mut pending_param_array_aliases = Vec::new();
        let program = self.exec.programs[target_prog].program;
        let callee = &program.funcs[proc.0];
        // Actuals are evaluated against the caller before the new frame is pushed.
        for (i, arg) in args.iter().enumerate() {
            let li = i + 1;
            let formal = callee.locals.get(li);
            let place_value;
            let actual = match arg {
                OxArg::ByVal(op) => op,
                OxArg::ByRef(place)
                    if formal
                        .and_then(|local| local.param.as_ref())
                        .is_some_and(|p| !p.by_ref) =>
                {
                    place_value = OxOperand::Use(*place);
                    &place_value
                }
                OxArg::ByRef(place) => {
                    let loc = self.resolve(place);
                    match self.byref_reference(place, formal.map(|local| &local.ty))? {
                        Some(reference) => frame.slots[li] = reference,
                        None => {
                            frame.set_alias(li, loc);
                        }
                    }
                    continue;
                }
                OxArg::Omitted => {
                    let seeded = formal.and_then(omitted_param_value).transpose()?;
                    if let Some(slot) = frame.local_mut(li) {
                        *slot = seeded.unwrap_or_else(|| Variant::from_error_code(MISSING_ARG));
                    }
                    continue;
                }
            };
            let aliases = self.param_array_aliases_for_operand(actual);
            let value = self.operand(actual)?;
            if let (Some(slot), Some(formal)) = (frame.local_mut(li), formal) {
                *slot =
                    calcvba_rt_abi::coerce_project_argument(&value, formal).map_err(arith_fault)?;
            }
            if let Some(aliases) = aliases {
                pending_param_array_aliases.push((Loc::Local(base, li), aliases));
            }
        }
        let saved_err = self.exec.err_engine.enter_activation();
        frame.saved_error_mode = saved_err.error_mode;
        frame.saved_active_error = saved_err.active_error;
        self.frames.push(frame);
        for (loc, aliases) in pending_param_array_aliases {
            self.param_array_aliases.insert(loc, aliases);
        }
        let result = self.run_loop(base);
        // The function result is the pushed frame's return local (the nested `run_loop` broke
        // at the frame's `Return` without copying it out). Capture it before unwinding.
        let ret = self
            .frames
            .get(base)
            .and_then(|fr| fr.return_local.and_then(|rl| fr.local(rl.0).cloned()))
            .unwrap_or_else(Variant::empty);
        let saved_err = self.frames.get(base).map(|fr| SavedErrState {
            error_mode: fr.saved_error_mode,
            active_error: fr.saved_active_error,
        });
        if let Some(saved_err) = saved_err {
            self.exec.err_engine.restore(saved_err);
        }
        self.truncate_frames_with_withevents_cleanup(base);
        self.prune_param_array_aliases_from_depth(self.frames.len());
        self.cur = saved_cur;
        self.maybe_drain();
        result.map(|()| ret)
    }

    /// Like [`Vm3::run_proc_with_me`] but seeds the arguments as already-evaluated **values**
    /// (a host-supplied `Vec<Variant>`) directly into `locals[1..]`, rather than resolving
    /// `OxArg` operands against a caller frame. Required for the host session API: an
    /// `OxArg::ByVal` carries only a scalar `OxConst`, so an object / array / arbitrary-Variant
    /// argument coming from the host cannot be expressed as an operand — it must be injected by
    /// value. `me` goes to local 0; all args are by-value (the host owns any ByRef itself).
    /// Returns the proc's result; with `suppress`, a `Fault` is swallowed to `Empty` (used by
    /// lifecycle/event callbacks). Wired by W3 (Collection dispatch), W5 (events) and W7
    /// (the host create/invoke session API).
    fn run_proc_with_values(
        &mut self,
        target_prog: usize,
        proc: FuncId,
        me: Variant,
        args: Vec<Variant>,
        suppress: bool,
    ) -> Result<Variant, Vm3Error> {
        self.run_proc_with_values_captured(target_prog, proc, me, args, suppress, None)
    }

    fn run_proc_with_values_captured(
        &mut self,
        target_prog: usize,
        proc: FuncId,
        me: Variant,
        args: Vec<Variant>,
        suppress: bool,
        capture: Option<(&mut [Variant], &[usize])>,
    ) -> Result<Variant, Vm3Error> {
        self.guard_call_depth()?;
        // Preserve the caller's program across the nested run_loop: it leaves `cur` at the
        // callee's program (run_loop re-derives cur per iteration), so the caller's post-call
        // resolves (store dst / extern_args) would otherwise use the wrong program. The caller
        // sets `cur` to the receiver's program (object dispatch / terminate) before calling.
        let saved_cur = self.cur;
        // Suppressed finalizers/callbacks must not overwrite the caller-visible Err object.
        let suppressed_err_engine = suppress.then(|| self.exec.err_engine.clone());
        let base = self.frames.len();
        let mut frame = self.new_frame_in(target_prog, proc)?;
        if let Some(slot) = frame.local_mut(0) {
            *slot = me;
        }
        // Host args follow `me` at locals 1.. as direct values — no caller-operand resolution,
        // no ByRef aliasing (the host marshals by value).
        for (i, mut v) in args.into_iter().enumerate() {
            if let Some(formal) = self.exec.programs[target_prog].program.funcs[proc.0]
                .locals
                .get(i + 1)
            {
                calcvba_rt_abi::validate_project_argument(&v, formal).map_err(arith_fault)?;
            }
            // A host passes an omitted argument as Missing; an Optional parameter then starts
            // as a call that leaves it out starts it (its default, Missing, or its type's zero).
            if v.vtype() == VarType::Error
                && v.as_error_code() == Some(MISSING_ARG)
                && let Some(local) = self.exec.programs[target_prog].program.funcs[proc.0]
                    .locals
                    .get(i + 1)
                && let Some(omitted) = omitted_param_value(local)
            {
                v = omitted?;
            }
            if let Some(slot) = frame.local_mut(i + 1) {
                *slot = v;
            }
        }
        if let Some((values, aliases)) = capture.as_ref() {
            let func = &self.exec.programs[target_prog].program.funcs[proc.0];
            if values.len() != aliases.len() || values.len() + 1 > func.param_count {
                return Err(Vm3Error::Malformed(
                    "host argument capture shape differs".into(),
                ));
            }
            for (index, &first) in aliases.iter().enumerate() {
                if first > index || aliases[first] != first {
                    return Err(Vm3Error::Malformed(
                        "host alias must name its first argument".into(),
                    ));
                }
                let by_ref = |i: usize| func.locals[i + 1].param.as_ref().is_some_and(|p| p.by_ref);
                if first < index && by_ref(index) && by_ref(first) {
                    if func.locals[index + 1].ty != func.locals[first + 1].ty {
                        return Err(Vm3Error::Malformed("host alias types differ".into()));
                    }
                    frame.set_alias(index + 1, Loc::Local(base, first + 1));
                }
            }
        }
        let saved_err = self.exec.err_engine.enter_activation();
        frame.saved_error_mode = saved_err.error_mode;
        frame.saved_active_error = saved_err.active_error;
        self.frames.push(frame);
        let result = self.run_loop(base);
        // Capture the result from the pushed frame's return local before unwinding (the nested
        // `run_loop` broke at the frame's `Return` without copying it out) — same as
        // `run_proc_with_me`.
        let ret = self
            .frames
            .get(base)
            .and_then(|fr| fr.return_local.and_then(|rl| fr.local(rl.0).cloned()))
            .unwrap_or_else(Variant::empty);
        if let Some((values, _)) = capture
            && let Some(frame) = self.frames.get(base)
        {
            for (index, value) in values.iter_mut().enumerate() {
                let formal =
                    &self.exec.programs[target_prog].program.funcs[proc.0].locals[index + 1];
                if !formal.param.as_ref().is_some_and(|param| param.by_ref) {
                    continue;
                }
                let slot = match frame.alias_of(index + 1) {
                    Some(Loc::Local(_, slot)) => slot,
                    _ => index + 1,
                };
                if let Some(updated) = frame.local(slot) {
                    *value = updated.clone();
                }
            }
        }
        let saved_err = self.frames.get(base).map(|fr| SavedErrState {
            error_mode: fr.saved_error_mode,
            active_error: fr.saved_active_error,
        });
        if let Some(saved_err) = saved_err {
            self.exec.err_engine.restore(saved_err);
        }
        self.truncate_frames_with_withevents_cleanup(base);
        self.prune_param_array_aliases_from_depth(self.frames.len());
        self.cur = saved_cur;
        self.maybe_drain();
        if let Some(saved) = suppressed_err_engine {
            self.exec.err_engine = saved;
        }
        match result {
            Ok(()) => Ok(ret),
            Err(Vm3Error::Fault(_)) if suppress => Ok(Variant::empty()),
            Err(e) => Err(e),
        }
    }

    fn maybe_drain(&mut self) {
        calcvba_rt_abi::drain_terminations(
            self,
            |vm| &mut vm.exec,
            |vm, bundle, proc, me| {
                let _ = vm.run_proc_with_me(bundle, proc, me, &[]);
            },
        );
    }

    /// Advises one sink for a WithEvents variable on a COM source (design doc, step 8c). An
    /// object without the declared type's connection point raises VBA error 459.
    fn advise_com_events(
        &mut self,
        key: i64,
        binding_token: i32,
        owner: &Variant,
        source: &ObjectRef,
    ) -> Result<(), Vm3Error> {
        let Some(interface) = self.exec.com_event_interface(self.cur, binding_token) else {
            return Ok(());
        };
        let host = self.exec.host;
        let connection = host
            .com()
            .advise_events(source, &interface, key as u64)
            .map_err(|error| Vm3Error::Fault(Fault::from_hal(error)))?;
        self.exec.insert_com_connection(
            key,
            calcvba_rt_abi::ComEventConnection {
                connection,
                owner: owner.clone(),
                program: self.cur,
                binding: binding_token,
            },
        );
        Ok(())
    }

    /// Unadvises a WithEvents key's sink (rebinding or `Set x = Nothing`).
    fn unsubscribe_com_key(&mut self, key: i64) {
        let connections = self.exec.take_com_connections(|held| held == key);
        drop(connections);
    }

    /// Unadvises every sink owned by `owner_raw` (owner cleared or terminated).
    fn unsubscribe_com_owner(&mut self, owner_raw: i32) {
        let connections = self
            .exec
            .take_com_connections(|key| withevents_owner_raw(key) == owner_raw);
        drop(connections);
    }

    /// `TypeOf <object> Is <Type>`: for a project instance, match the bare type name against
    /// the instance's class name or any `Implements`ed interface; for a foreign/COM object,
    /// delegate to the host (unreachable until `CreateObject` lands in M3-8, but mirrors vm2).
    fn type_of_is(&mut self, object: &OxOperand, type_name: &str) -> Result<bool, Vm3Error> {
        let v = self.operand(object)?;
        // `TypeOf` on a non-object (`Empty`, `Null`, a number) is False, and so is it on
        // `Nothing` — except for an `As Object` variable holding Nothing, which raises 91
        // (Excel: `Set o = Nothing: TypeOf o Is Collection`).
        if v.vtype() != VarType::Object {
            return Ok(false);
        }
        if is_nothing(&v) {
            return if self.operand_is_untyped_object(object) {
                Err(Vm3Error::Fault(Fault::new(
                    91,
                    "Object variable or With block variable not set",
                )))
            } else {
                Ok(false)
            };
        }
        let obj = variant_to_object(&v)?;
        calcvba_rt_abi::object_types::matches_type(
            self.exec.host,
            &obj,
            type_name,
            self.exec.programs[self.cur]
                .program
                .named_interface_iid(type_name),
        )
        .map_err(Vm3Error::Fault)
    }

    /// The class name of a project instance (else the host's COM name) — the
    /// `TypeName`-of-object resolution mirroring vm2's `object_type_name`.
    fn object_type_name(&self, object: &ObjectRef) -> Option<String> {
        // A built-in `Collection` (reserved sentinel route key) reports `TypeName` "Collection".
        if object.route_key() == VBA_COLLECTION_ROUTE_KEY {
            return Some("Collection".to_string());
        }
        if calcvba_lib::objects::is_object(object) {
            return Some(object.class_descriptor().type_name().to_string());
        }
        if object.is_project_instance() {
            return self
                .exec
                .programs
                .get(object.bundle_id() as usize)
                .and_then(|lp| lp.program.classes.get(object.route_key() as usize))
                .map(|c| c.type_name().to_owned());
        }
        self.exec
            .host
            .com()
            .object_type_name(object.clone())
            .ok()
            .flatten()
    }

    fn project_call_args(
        &self,
        target_prog: usize,
        proc: FuncId,
        args: &[OxCallArg],
    ) -> Result<Vec<OxArg>, Vm3Error> {
        let callee = self
            .exec
            .programs
            .get(target_prog)
            .and_then(|lp| lp.program.funcs.get(proc.0))
            .ok_or_else(|| Vm3Error::Malformed(format!("call to unknown proc {}", proc.0)))?;
        let actuals = args
            .iter()
            .map(|arg| {
                let lowered = match arg {
                    OxCallArg::ByRef(place)
                    | OxCallArg::Named {
                        by_ref: Some(place),
                        ..
                    } => OxArg::ByRef(*place),
                    OxCallArg::Operand(op) | OxCallArg::Named { value: op, .. } => {
                        OxArg::ByVal(op.clone())
                    }
                    OxCallArg::Omitted => OxArg::Omitted,
                    OxCallArg::Const(_) => {
                        return Err(Vm3Error::Malformed(
                            "a Const argument in a project method call".into(),
                        ));
                    }
                };
                let name = match arg {
                    OxCallArg::Named { name, .. } => Some(name.as_str()),
                    _ => None,
                };
                Ok((lowered, name))
            })
            .collect::<Result<Vec<_>, Vm3Error>>()?;
        Ok(calcvba_rt_abi::order_project_arguments(callee, actuals)
            .map_err(Vm3Error::Fault)?
            .into_iter()
            .map(|arg| arg.unwrap_or(OxArg::Omitted))
            .collect())
    }

    fn dispatch_project_default_member(
        &mut self,
        object: ObjectRef,
        me: Variant,
        invoke_kind: TypeLibMemberInvokeKind,
        args: &[OxCallArg],
    ) -> Result<Variant, Vm3Error> {
        let runtime_kind = calcvba_rt_abi::runtime_invoke_kind_from_typelib(invoke_kind);
        let obj_bundle = object.bundle_id() as usize;
        let member = resolve_project_default_member(
            object.class_descriptor(),
            runtime_kind,
            args.is_empty(),
            obj_bundle != self.cur,
        )
        .map_err(Vm3Error::Fault)?;
        self.dispatch_resolved_project_member(object, me, obj_bundle, member, runtime_kind, args)
    }

    fn dispatch_project_default_member_values(
        &mut self,
        object: ObjectRef,
        me: Variant,
        invoke_kind: TypeLibMemberInvokeKind,
        args: Vec<Variant>,
    ) -> Result<Variant, Vm3Error> {
        let runtime_kind = calcvba_rt_abi::runtime_invoke_kind_from_typelib(invoke_kind);
        let obj_bundle = object.bundle_id() as usize;
        let member = resolve_project_default_member(
            object.class_descriptor(),
            runtime_kind,
            args.is_empty(),
            obj_bundle != self.cur,
        )
        .map_err(Vm3Error::Fault)?;
        match member.backing {
            calcvba_runtime::RuntimeMemberBacking::Proc(proc) => {
                self.run_proc_with_values(obj_bundle, FuncId(proc), me, args, false)
            }
            calcvba_runtime::RuntimeMemberBacking::Field(field_token) => self
                .dispatch_project_field_member_values(
                    &object,
                    member,
                    field_token,
                    runtime_kind,
                    args,
                ),
            calcvba_runtime::RuntimeMemberBacking::None => {
                Err(Vm3Error::Fault(Fault::new(438, default_error_message(438))))
            }
        }
    }

    /// Late-bound dispatch on a project instance: resolve through `RuntimeDispatchPlanCache`
    /// and dispatch by descriptor backing (slot/DISPID), never scanning `OxClass::methods`.
    fn dispatch_project_method(
        &mut self,
        object: ObjectRef,
        me: Variant,
        name: &str,
        invoke_kind: TypeLibMemberInvokeKind,
        args: &[OxCallArg],
        public_only: bool,
    ) -> Result<Variant, Vm3Error> {
        self.dispatch_project_method_resolved(
            object,
            me,
            None,
            name,
            invoke_kind,
            args,
            public_only,
        )
    }

    #[allow(clippy::too_many_arguments)]
    fn dispatch_project_method_resolved(
        &mut self,
        object: ObjectRef,
        me: Variant,
        early_member: Option<&calcvba_com::TypeLibMemberMetadata>,
        name: &str,
        invoke_kind: TypeLibMemberInvokeKind,
        args: &[OxCallArg],
        public_only: bool,
    ) -> Result<Variant, Vm3Error> {
        // A built-in `Collection` carries the reserved sentinel route key; its members dispatch
        // to the shared keyed-Collection logic over the box-owned `CollectionData`, not a
        // program class (W3).
        if calcvba_lib::objects::is_object(&object) {
            return self.dispatch_library_object(
                &object,
                name,
                project_member_kind(invoke_kind),
                args,
            );
        }
        if object.route_key() == VBA_COLLECTION_ROUTE_KEY {
            return self.dispatch_collection_method(&object, name, args);
        }
        let runtime_kind = calcvba_rt_abi::runtime_invoke_kind_from_typelib(invoke_kind);
        let obj_bundle = object.bundle_id() as usize;
        let effective_public_only =
            (public_only && early_member.is_none()) || obj_bundle != self.cur;
        let member = calcvba_rt_abi::resolve_project_member_for_call(
            &mut self.exec.dispatch_plan_cache,
            object.class_descriptor(),
            early_member,
            name,
            runtime_kind,
            args.is_empty(),
            effective_public_only,
        )
        .map_err(Vm3Error::Fault)?;
        self.dispatch_resolved_project_member(object, me, obj_bundle, member, runtime_kind, args)
    }

    fn dispatch_resolved_project_member(
        &mut self,
        object: ObjectRef,
        me: Variant,
        obj_bundle: usize,
        member: &'static calcvba_runtime::RuntimeMemberDescriptor,
        requested_kind: calcvba_runtime::RuntimeMemberInvokeKind,
        args: &[OxCallArg],
    ) -> Result<Variant, Vm3Error> {
        match member.backing {
            calcvba_runtime::RuntimeMemberBacking::Proc(proc) => {
                let proc = FuncId(proc);
                // ByRef args alias the caller's place (write-back); named arguments reorder by
                // the callee parameter names before the frame is seeded.
                let proc_args = self.project_call_args(obj_bundle, proc, args)?;
                // Run the method body in the object's program (target_prog = obj_bundle).
                self.run_proc_with_me(obj_bundle, proc, me, &proc_args)
            }
            calcvba_runtime::RuntimeMemberBacking::Field(field_token) => self
                .dispatch_project_field_member(&object, member, field_token, requested_kind, args),
            calcvba_runtime::RuntimeMemberBacking::None => {
                Err(Vm3Error::Fault(Fault::new(438, default_error_message(438))))
            }
        }
    }

    fn dispatch_project_field_member(
        &mut self,
        object: &ObjectRef,
        member: &'static calcvba_runtime::RuntimeMemberDescriptor,
        field_token: i32,
        requested_kind: calcvba_runtime::RuntimeMemberInvokeKind,
        args: &[OxCallArg],
    ) -> Result<Variant, Vm3Error> {
        use calcvba_runtime::{RuntimeMemberInvokeKind, RuntimeValueType};
        match requested_kind {
            RuntimeMemberInvokeKind::PropertyGet | RuntimeMemberInvokeKind::Method => {
                let raw = self.read_project_field_as_new(object, field_token)?;
                if args.is_empty() {
                    let target_ty = member.return_type.unwrap_or(RuntimeValueType::Variant);
                    calcvba_rt_abi::coerce_value_for_runtime_type(&raw, target_ty)
                        .map_err(Vm3Error::Fault)
                } else if raw.vtype() == VarType::Object {
                    self.dispatch_default_member(raw, TypeLibMemberInvokeKind::PropertyGet, args)
                } else {
                    Err(Vm3Error::Fault(Fault::new(450, default_error_message(450))))
                }
            }
            RuntimeMemberInvokeKind::PropertyLet => {
                let [arg] = args else {
                    return Err(Vm3Error::Fault(Fault::new(450, default_error_message(450))));
                };
                let value = self.eval_call_arg_value(arg)?;
                let value = if value.vtype() == VarType::Object {
                    self.arithmetic_value(value)?
                } else {
                    value
                };
                let target_ty = member
                    .params
                    .last()
                    .map_or(RuntimeValueType::Variant, |param| param.value_type);
                let coerced = calcvba_rt_abi::coerce_value_for_runtime_type(&value, target_ty)
                    .map_err(Vm3Error::Fault)?;
                object.project_field_set(field_token, coerced);
                Ok(Variant::empty())
            }
            RuntimeMemberInvokeKind::PropertySet => {
                let [arg] = args else {
                    return Err(Vm3Error::Fault(Fault::new(450, default_error_message(450))));
                };
                let value = self.eval_call_arg_value(arg)?;
                if value.vtype() != VarType::Object {
                    return Err(Vm3Error::Fault(Fault::new(424, default_error_message(424))));
                }
                object.project_field_set(field_token, value);
                Ok(Variant::empty())
            }
        }
    }

    fn dispatch_project_field_member_values(
        &mut self,
        object: &ObjectRef,
        member: &'static calcvba_runtime::RuntimeMemberDescriptor,
        field_token: i32,
        requested_kind: calcvba_runtime::RuntimeMemberInvokeKind,
        args: Vec<Variant>,
    ) -> Result<Variant, Vm3Error> {
        use calcvba_runtime::{RuntimeMemberInvokeKind, RuntimeValueType};
        match requested_kind {
            RuntimeMemberInvokeKind::PropertyGet | RuntimeMemberInvokeKind::Method => {
                let raw = self.read_project_field_as_new(object, field_token)?;
                if args.is_empty() {
                    let target_ty = member.return_type.unwrap_or(RuntimeValueType::Variant);
                    calcvba_rt_abi::coerce_value_for_runtime_type(&raw, target_ty)
                        .map_err(Vm3Error::Fault)
                } else if raw.vtype() == VarType::Object {
                    self.dispatch_default_member_values(
                        raw,
                        TypeLibMemberInvokeKind::PropertyGet,
                        args,
                    )
                } else {
                    Err(Vm3Error::Fault(Fault::new(450, default_error_message(450))))
                }
            }
            RuntimeMemberInvokeKind::PropertyLet => {
                let [value] = args.as_slice() else {
                    return Err(Vm3Error::Fault(Fault::new(450, default_error_message(450))));
                };
                let value = if value.vtype() == VarType::Object {
                    self.arithmetic_value(value.clone())?
                } else {
                    value.clone()
                };
                let target_ty = member
                    .params
                    .last()
                    .map_or(RuntimeValueType::Variant, |param| param.value_type);
                let coerced = calcvba_rt_abi::coerce_value_for_runtime_type(&value, target_ty)
                    .map_err(Vm3Error::Fault)?;
                object.project_field_set(field_token, coerced);
                Ok(Variant::empty())
            }
            RuntimeMemberInvokeKind::PropertySet => {
                let [value] = args.as_slice() else {
                    return Err(Vm3Error::Fault(Fault::new(450, default_error_message(450))));
                };
                if value.vtype() != VarType::Object {
                    return Err(Vm3Error::Fault(Fault::new(424, default_error_message(424))));
                }
                object.project_field_set(field_token, value.clone());
                Ok(Variant::empty())
            }
        }
    }

    fn eval_call_arg_value(&mut self, arg: &OxCallArg) -> Result<Variant, Vm3Error> {
        match arg {
            OxCallArg::Operand(op) | OxCallArg::Named { value: op, .. } => self.operand(op),
            OxCallArg::ByRef(place) => self.read(place),
            OxCallArg::Const(n) => Ok(Variant::from_i32(*n)),
            OxCallArg::Omitted => Err(Vm3Error::Fault(Fault::new(449, default_error_message(449)))),
        }
    }

    /// Mint a built-in library-class instance for `New <LibClass>` (`OxInst::NewExtern`). Only
    /// the `VBA.Collection` class is wired today: a native-backed object whose state rides the
    /// object box (`native_state`), so no `Class_Initialize` runs and a reserved sentinel route
    /// key marks it for the Collection dispatch leg.
    fn new_extern_instance(&mut self, import: ImportId) -> Result<Variant, Vm3Error> {
        let imp = self.cur_program().imports.get(import.0).ok_or_else(|| {
            Vm3Error::Malformed(format!("NewExtern names unknown import {}", import.0))
        })?;
        // A cross-PROJECT class `New OtherProj.Class`: resolve it by unit name + class token to a
        // class in the referenced program and mint a project instance THERE — its bundle_id,
        // leaked descriptor, and Class_Initialize, exactly as a local `New` would in that
        // program. The synthetic `VBA` unit is the built-in library path (Collection) below.
        if !imp.unit.eq_ignore_ascii_case("VBA") {
            // Mint in the referenced program (cur=B) so new_project_instance stamps B's
            // bundle_id + descriptor and runs Class_Initialize in B; restore the caller's
            // program after (run_proc_with_me already preserves cur across the init body).
            let (b, class_idx) = self.resolve_cross_project_class(import)?;
            let saved_cur = self.cur;
            self.cur = b;
            let result = self.new_project_instance(class_idx);
            self.cur = saved_cur;
            return result;
        }
        // Built-in library class (Collection): native-backed, reserved sentinel route key.
        let descriptor = self.resolve_extern_class(import)?;
        let instance_id = self.exec.next_instance_id;
        self.exec.next_instance_id += 1;
        let object = ObjectRef::from_project_instance(
            instance_id,
            VBA_COLLECTION_ROUTE_KEY,
            self.cur as i32,
            false,
            descriptor,
        );
        Ok(Variant::from_object_ref(object))
    }

    /// Resolve a cross-project class import (`New`/predeclared of `OtherProj.Class`) to the
    /// referenced program index + the class index within it (by unit name + class export token).
    fn resolve_cross_project_class(&self, import: ImportId) -> Result<(usize, usize), Vm3Error> {
        let imp = self.cur_program().imports.get(import.0).ok_or_else(|| {
            Vm3Error::Malformed(format!("NewExtern names unknown import {}", import.0))
        })?;
        let b = self.program_index_by_unit(&imp.unit).ok_or_else(|| {
            Vm3Error::Malformed(format!("unresolved reference to unit '{}'", imp.unit))
        })?;
        let class_idx = self.exec.programs[b]
            .program
            .exports
            .iter()
            .find(|e| e.token.matches(&imp.token))
            .ok_or_else(|| {
                Vm3Error::Malformed(format!(
                    "unit '{}' has no export matching the class import",
                    imp.unit
                ))
            })
            .and_then(|export| match export.target {
                calcvba_bundle::ExportTarget::Class(c) => Ok(c),
                _ => Err(Vm3Error::Malformed(
                    "a cross-project class import resolved to a non-class export".into(),
                )),
            })?;
        Ok((b, class_idx))
    }

    /// A cross-project `VB_PredeclaredId` singleton (`OtherProj.Class1` with no `New`): the
    /// `PredeclaredExtern` analogue of [`Self::new_extern_instance`] — get-or-create the
    /// singleton in the referenced program (cached there, with that program's bundle_id).
    fn predeclared_extern_instance(&mut self, import: ImportId) -> Result<Variant, Vm3Error> {
        let (b, class_idx) = self.resolve_cross_project_class(import)?;
        let saved_cur = self.cur;
        self.cur = b;
        let result = self.predeclared_instance(class_idx);
        self.cur = saved_cur;
        result
    }

    fn set_predeclared_extern_instance(
        &mut self,
        import: ImportId,
        value: Variant,
    ) -> Result<(), Vm3Error> {
        let (program_index, class_idx) = self.resolve_cross_project_class(import)?;
        self.set_predeclared_instance(program_index, class_idx, value)
    }

    /// Resolve a `New <LibClass>` import to its runtime QI descriptor. The `ExportToken::Class`
    /// analogue of [`Self::resolve_library_import`]. Only `VBA.Collection` is wired today;
    /// another VBA project needs the multi-`OxProgram` linker (deferred), and any other library
    /// class is a clean `Unimplemented` — never a silent mis-run.
    fn resolve_extern_class(
        &self,
        import: ImportId,
    ) -> Result<&'static RuntimeClassDescriptor, Vm3Error> {
        let imp = self.cur_program().imports.get(import.0).ok_or_else(|| {
            Vm3Error::Malformed(format!("NewExtern names unknown import {}", import.0))
        })?;
        if !imp.unit.eq_ignore_ascii_case("VBA") {
            return Err(Vm3Error::Unimplemented {
                what: "cross-project OxProgram link",
            });
        }
        let lib = calcvba_bundle::vba_library_bundle();
        let export = lib
            .exports
            .iter()
            .find(|e| e.token.matches(&imp.token))
            .ok_or_else(|| {
                Vm3Error::Malformed(format!(
                    "the VBA library bundle has no export matching NewExtern import {}",
                    import.0
                ))
            })?;
        let calcvba_bundle::ExportTarget::Class(class_idx) = export.target else {
            return Err(Vm3Error::Malformed(
                "a VBA library NewExtern resolved to a non-class export".into(),
            ));
        };
        let class = lib.classes.get(class_idx).ok_or_else(|| {
            Vm3Error::Malformed(format!("VBA library class {class_idx} out of range"))
        })?;
        if !class.name.eq_ignore_ascii_case("Collection") {
            return Err(Vm3Error::Unimplemented {
                what: "built-in library class other than Collection",
            });
        }
        Ok(&VBA_COLLECTION_DESCRIPTOR)
    }

    /// Dispatch a member call on a built-in `Collection` receiver to the shared keyed logic
    /// (`calcvba_eval::collection::dispatch_collection`) over the box-owned `CollectionData`. The
    /// member name resolves to its `NativeMethodId` via the VBA bundle's Collection class.
    fn dispatch_collection_method(
        &mut self,
        object: &ObjectRef,
        name: &str,
        args: &[OxCallArg],
    ) -> Result<Variant, Vm3Error> {
        let native = Self::vba_collection_native_method(name).ok_or_else(|| {
            Vm3Error::Fault(Fault::new(
                438,
                format!("Collection doesn't support '{name}'"),
            ))
        })?;
        let method = match native {
            calcvba_bundle::NativeMethodId::CollectionAdd => CollectionMethod::Add,
            calcvba_bundle::NativeMethodId::CollectionItem => CollectionMethod::Item,
            calcvba_bundle::NativeMethodId::CollectionCount => CollectionMethod::Count,
            calcvba_bundle::NativeMethodId::CollectionRemove => CollectionMethod::Remove,
            calcvba_bundle::NativeMethodId::CollectionNewEnum => {
                return self.collection_new_enum_object(object);
            }
        };
        let has_named = args.iter().any(|a| matches!(a, OxCallArg::Named { .. }));
        if !has_named && args.len() <= 4 {
            let mut argv = [
                Variant::empty(),
                Variant::empty(),
                Variant::empty(),
                Variant::empty(),
            ];
            for (dst, a) in argv.iter_mut().zip(args.iter()) {
                *dst = match a {
                    OxCallArg::Operand(op) => self.operand(op)?,
                    OxCallArg::ByRef(place) => self.read(place)?,
                    OxCallArg::Omitted => Variant::from_error_code(MISSING_ARG),
                    OxCallArg::Named { value, .. } => self.operand(value)?,
                    OxCallArg::Const(_) => {
                        return Err(Vm3Error::Malformed(
                            "a Const argument in a Collection method call".into(),
                        ));
                    }
                };
            }
            return object
                .with_native_collection(|data| {
                    dispatch_collection(method, data, &argv[..args.len()])
                })
                .ok_or_else(|| Vm3Error::Fault(Fault::new(424, "Object required")))?
                .map_err(|error| Vm3Error::Fault(Fault::standard(error.vba_number())));
        }
        let names: Vec<_> = args
            .iter()
            .map(|arg| match arg {
                OxCallArg::Named { name, .. } => Some(name.clone()),
                _ => None,
            })
            .collect();
        let argv = self.collection_args(args)?;
        let argv = calcvba_eval::collection::order_collection_args(method, &argv, &names)
            .map_err(|number| Vm3Error::Fault(Fault::new(number, default_error_message(number))))?;
        object
            .with_native_collection(|data| dispatch_collection(method, data, &argv))
            .ok_or_else(|| Vm3Error::Fault(Fault::new(424, "Object required")))?
            .map_err(|error| Vm3Error::Fault(Fault::standard(error.vba_number())))
    }

    fn dispatch_library_object(
        &mut self,
        object: &ObjectRef,
        name: &str,
        kind: ProjectMemberKind,
        args: &[OxCallArg],
    ) -> Result<Variant, Vm3Error> {
        let names: Vec<_> = args
            .iter()
            .map(|arg| {
                if let OxCallArg::Named { name, .. } = arg {
                    Some(name.clone())
                } else {
                    None
                }
            })
            .collect();
        let args = self.collection_args(args)?;
        calcvba_rt_abi::invoke_library_object(&mut self.exec, object, name, kind, &args, &names)
            .map_err(Vm3Error::Fault)
    }

    /// Marshal a Collection member's call args to by-value `Variant`s (omitted → `MISSING_ARG`,
    /// the sentinel the shared dispatcher recognises). The built-in Collection never writes
    /// back, so a ByRef arg is read by value.
    fn collection_args(&mut self, args: &[OxCallArg]) -> Result<Vec<Variant>, Vm3Error> {
        args.iter()
            .map(|a| match a {
                OxCallArg::Operand(op) => self.operand(op),
                OxCallArg::ByRef(place) => self.read(place),
                OxCallArg::Named { value, .. } => self.operand(value),
                OxCallArg::Omitted => Ok(Variant::from_error_code(MISSING_ARG)),
                OxCallArg::Const(_) => Err(Vm3Error::Malformed(
                    "a Const argument in a Collection method call".into(),
                )),
            })
            .collect()
    }

    /// Evaluate a list of index operands to by-value `Variant`s (a Collection default-member
    /// `c(i)` call's indices are plain operands, not `OxCallArg`s).
    fn operands_to_values(&mut self, ops: &[OxOperand]) -> Result<Vec<Variant>, Vm3Error> {
        ops.iter().map(|op| self.operand(op)).collect()
    }

    /// Resolve a `Collection` member name to its `NativeMethodId` (direct case-insensitive
    /// match for the five fixed `VBA.Collection` members).
    fn vba_collection_native_method(member: &str) -> Option<calcvba_bundle::NativeMethodId> {
        let member = member.trim().trim_start_matches('[').trim_end_matches(']');
        if member.eq_ignore_ascii_case("Add") {
            Some(calcvba_bundle::NativeMethodId::CollectionAdd)
        } else if member.eq_ignore_ascii_case("Item") {
            Some(calcvba_bundle::NativeMethodId::CollectionItem)
        } else if member.eq_ignore_ascii_case("Count") {
            Some(calcvba_bundle::NativeMethodId::CollectionCount)
        } else if member.eq_ignore_ascii_case("Remove") {
            Some(calcvba_bundle::NativeMethodId::CollectionRemove)
        } else if member.eq_ignore_ascii_case("_NewEnum") {
            Some(calcvba_bundle::NativeMethodId::CollectionNewEnum)
        } else {
            None
        }
    }

    fn collection_new_enum_object(&mut self, object: &ObjectRef) -> Result<Variant, Vm3Error> {
        let snapshot = object
            .with_native_collection(|data| data.clone())
            .ok_or_else(|| Vm3Error::Fault(Fault::new(424, "Object required")))?;
        let instance_id = self.exec.next_instance_id;
        self.exec.next_instance_id += 1;
        let enumerator = ObjectRef::from_project_instance(
            instance_id,
            VBA_COLLECTION_ROUTE_KEY,
            object.bundle_id(),
            false,
            &VBA_COLLECTION_DESCRIPTOR,
        );
        enumerator
            .with_native_collection(|data| *data = snapshot)
            .ok_or_else(|| Vm3Error::Fault(Fault::new(424, "Object required")))?;
        Ok(Variant::from_object_ref(enumerator))
    }

    fn foreach_elements(&mut self, src: Variant, depth: usize) -> Result<ForEachItems, Vm3Error> {
        if depth > 8 {
            return Err(Vm3Error::Fault(Fault::new(
                438,
                "Object doesn't support this property or method",
            )));
        }
        if let Some(arr) = src.as_safearray() {
            return Ok(ForEachItems::Snapshot(
                arr.enumeration_elements().unwrap_or_default(),
            ));
        }
        let Some(obj) = src.as_object_ref() else {
            return Err(Vm3Error::Fault(Fault::new(13, "Type mismatch")));
        };
        if depth == 0
            && obj.is_compat_object()
            && obj.route_key() == VBA_COLLECTION_ROUTE_KEY
            && let Some(cursor) = CollectionCursor::new(&obj)
        {
            return Ok(ForEachItems::CollectionCursor(cursor));
        }
        if let Some(values) = obj.native_collection_snapshot() {
            return Ok(ForEachItems::Snapshot(values));
        }
        if obj.is_compat_object() && obj.route_key() == VBA_COLLECTION_ROUTE_KEY {
            return Ok(ForEachItems::Snapshot(Vec::new()));
        }
        if obj.is_project_instance() {
            return self.project_class_enumerator_elements(obj, depth + 1);
        }
        let host = self.exec.host;
        let enumerator = {
            let _events = self.event_scope();
            host.com().object_enumerator(&obj)
        };
        self.take_event_abort()?;
        enumerator
            .map(ForEachItems::Enumerator)
            .map_err(Fault::from_hal)
            .map_err(Vm3Error::Fault)
    }

    fn project_class_enumerator_elements(
        &mut self,
        object: ObjectRef,
        depth: usize,
    ) -> Result<ForEachItems, Vm3Error> {
        let obj_bundle = object.bundle_id() as usize;
        let member = object
            .class_descriptor()
            .resolve_enumerator_member()
            .ok_or_else(|| {
                Vm3Error::Fault(Fault::new(
                    438,
                    "Object doesn't support this property or method",
                ))
            })?;
        let result = match member.backing {
            calcvba_runtime::RuntimeMemberBacking::Proc(proc) => self.run_proc_with_me(
                obj_bundle,
                FuncId(proc),
                Variant::from_object_ref(object),
                &[],
            )?,
            calcvba_runtime::RuntimeMemberBacking::Field(field_token) => self
                .dispatch_project_field_member_values(
                    &object,
                    member,
                    field_token,
                    calcvba_runtime::RuntimeMemberInvokeKind::PropertyGet,
                    Vec::new(),
                )?,
            calcvba_runtime::RuntimeMemberBacking::None => {
                return Err(Vm3Error::Fault(Fault::new(
                    438,
                    "Object doesn't support this property or method",
                )));
            }
        };
        self.foreach_elements(result, depth)
    }

    /// Resolve an arithmetic value while retaining the caller's owning object.
    /// An arithmetic operand's value: [`Self::arithmetic_value`], with `Nothing` a type
    /// mismatch (13), as VBA reports `z + 1` (a conversion or comparison reports 91).
    fn arithmetic_operand(&mut self, value: Variant) -> Result<Variant, Vm3Error> {
        if value.vtype() == VarType::Object && is_nothing(&value) {
            return Err(Vm3Error::Fault(Fault::new(13, default_error_message(13))));
        }
        self.arithmetic_value(value)
    }

    fn arithmetic_value(&mut self, value: Variant) -> Result<Variant, Vm3Error> {
        read_default_value(
            value,
            |value| {
                self.dispatch_default_member_values(
                    value,
                    TypeLibMemberInvokeKind::PropertyGet,
                    Vec::new(),
                )
            },
            Vm3Error::Fault,
        )
    }

    fn dispatch_default_member(
        &mut self,
        recv_v: Variant,
        invoke_kind: TypeLibMemberInvokeKind,
        args: &[OxCallArg],
    ) -> Result<Variant, Vm3Error> {
        let object = variant_to_object(&recv_v)?;
        match object_member_route(&object) {
            ObjectMemberRoute::Library => {
                self.dispatch_library_object(&object, "", project_member_kind(invoke_kind), args)
            }
            ObjectMemberRoute::Collection => {
                let name = collection_member_name("", args.is_empty()).map_err(Vm3Error::Fault)?;
                self.dispatch_collection_method(&object, name, args)
            }
            ObjectMemberRoute::Project => {
                self.dispatch_project_default_member(object, recv_v, invoke_kind, args)
            }
            ObjectMemberRoute::Foreign => self.dispatch_com_method(
                object,
                calcvba_com::call::ComMember::Id(0),
                invoke_kind,
                args,
            ),
        }
    }

    fn dispatch_default_member_values(
        &mut self,
        recv_v: Variant,
        invoke_kind: TypeLibMemberInvokeKind,
        args: Vec<Variant>,
    ) -> Result<Variant, Vm3Error> {
        let object = variant_to_object(&recv_v)?;
        match object_member_route(&object) {
            ObjectMemberRoute::Library => calcvba_rt_abi::invoke_library_object(
                &mut self.exec,
                &object,
                "",
                project_member_kind(invoke_kind),
                &args,
                &[],
            )
            .map_err(Vm3Error::Fault),
            ObjectMemberRoute::Collection => {
                let name = collection_member_name("", args.is_empty()).map_err(Vm3Error::Fault)?;
                self.dispatch_collection_values(&object, name, args)
            }
            ObjectMemberRoute::Project => {
                self.dispatch_project_default_member_values(object, recv_v, invoke_kind, args)
            }
            ObjectMemberRoute::Foreign => self.dispatch_com_method_values(
                object,
                DynamicMemberSelector::DefaultMember,
                invoke_kind,
                args,
            ),
        }
    }

    fn dispatch_com_method_values(
        &mut self,
        object: ObjectRef,
        member: DynamicMemberSelector,
        invoke_kind: TypeLibMemberInvokeKind,
        args: Vec<Variant>,
    ) -> Result<Variant, Vm3Error> {
        let call_args: Vec<_> = args.iter().map(calcvba_com::call::ComArg::value).collect();
        let host = self.exec.host;
        let result = {
            let _events = self.event_scope();
            host.com().invoke_object(
                &object,
                com_member(&member),
                com_call_kind(invoke_kind),
                &call_args,
            )
        };
        self.take_event_abort()?;
        result.map_err(|e| Vm3Error::Fault(Fault::from_hal(e)))
    }

    fn dispatch_member_by_name(
        &mut self,
        recv_v: Variant,
        name: &str,
        invoke_kind: TypeLibMemberInvokeKind,
        args: &[OxCallArg],
        public_only: bool,
    ) -> Result<Variant, Vm3Error> {
        let object = variant_to_object(&recv_v)?;
        if object.is_project_instance() {
            self.dispatch_project_method(object, recv_v, name, invoke_kind, args, public_only)
        } else {
            self.dispatch_com_method(
                object,
                calcvba_com::call::ComMember::Name(name),
                invoke_kind,
                args,
            )
        }
    }

    /// COM dispatch on a genuine COM/foreign receiver: route the name or default
    /// member through the host's COM facet (the identical
    /// `dispatch_invoke_dynamic_variant_with_writebacks` contract vm2 drives — the host's
    /// `PreferVtable` strategy picks the transport, the same way for both VMs), and copy any
    /// `[out]`/`[in,out]` write-back to its ByRef call-site place. A dispatch fault carries
    /// the rich VBA `Err.Number` recovered from the HRESULT/EXCEPINFO (via [`Fault::from_hal`],
    /// never the flatten-to-5 default — the must-fix the milestone calls out).
    fn dispatch_com_method(
        &mut self,
        object: ObjectRef,
        member: calcvba_com::call::ComMember<'_>,
        invoke_kind: TypeLibMemberInvokeKind,
        args: &[OxCallArg],
    ) -> Result<Variant, Vm3Error> {
        // The argument cells: ByVal values (borrowed by the call), and for a ByRef place the
        // `VT_BYREF` reference to its storage, which the callee writes through.
        let mut cells = Vec::with_capacity(args.len());
        for arg in args {
            cells.push(match arg {
                OxCallArg::ByRef(place)
                | OxCallArg::Named {
                    by_ref: Some(place),
                    ..
                } => Some(self.com_reference(place)?),
                OxCallArg::Operand(op) => Some(self.operand(op)?),
                OxCallArg::Named { value, .. } => Some(self.operand(value)?),
                OxCallArg::Omitted => None,
                OxCallArg::Const(n) => Some(Variant::from_i32(*n)),
            });
        }
        let call_args = com_call_args(args, &cells);
        let host = self.exec.host;
        let result = {
            let _events = self.event_scope();
            host.com()
                .invoke_object(&object, member, com_call_kind(invoke_kind), &call_args)
        };
        self.take_event_abort()?;
        result.map_err(|e| Vm3Error::Fault(Fault::from_hal(e)))
    }

    /// Registers this VM as the thread's event executor for one host call (design doc, step
    /// 8c): a COM event fired during the call runs its WithEvents handler here, synchronously.
    /// The pointer is taken from `&mut self` right before the call; the caller makes the host
    /// call with no borrow of the VM held across it and drops the scope before using `self`.
    fn event_scope(&mut self) -> calcvba_runtime::events::ExecutorScope {
        let executor: &mut dyn calcvba_runtime::events::EventExecutor = self;
        // SAFETY: see above; the scope ends before this VM is used again.
        unsafe { calcvba_runtime::events::ExecutorScope::enter(NonNull::from(executor)) }
    }

    /// The error a WithEvents handler left unhandled during the last host call, as the
    /// uncatchable end of the run.
    fn take_event_abort(&mut self) -> Result<(), Vm3Error> {
        match self.exec.events.pending_abort.take() {
            Some(calcvba_rt_abi::RunAbort::Fault(fault)) => Err(Vm3Error::EventAbort(fault)),
            Some(calcvba_rt_abi::RunAbort::End) => Err(Vm3Error::Halted),
            None => Ok(()),
        }
    }

    /// The `VT_BYREF` cell a COM call gets for a ByRef place, as VBA builds it: a slot that
    /// already holds a reference (an element passed ByRef) is passed on, a typed variable is
    /// `VT_BYREF | vt` to its storage, and a Variant variable is `VT_BYREF | VT_VARIANT` to its
    /// cell.
    fn com_reference(&mut self, place: &OxPlace) -> Result<Variant, Vm3Error> {
        let typed = self.place_ty(place).is_byref_variant_reference();
        let loc = self.resolve(place);
        let cell = self
            .read_loc_mut(loc)?
            .ok_or_else(|| Vm3Error::Malformed(format!("ByRef COM argument {loc:?}")))?;
        if cell.byref_target().is_some() {
            return Ok(cell.clone_reference());
        }
        let cell: *mut Variant = cell;
        // SAFETY: frame slots and globals are stable storage that outlives the call; no Rust
        // reference to the cell is held across it.
        Ok(unsafe {
            if typed {
                Variant::by_ref((*cell).core().raw_vt(), Variant::payload_ptr(cell))
            } else {
                Variant::by_ref(0x0C, cell.cast())
            }
        })
    }

    fn dispatch_early_com_method(
        &mut self,
        object: ObjectRef,
        member: &calcvba_com::TypeLibMemberMetadata,
        args: &[OxCallArg],
    ) -> Result<Variant, Vm3Error> {
        let mut cells = Vec::with_capacity(args.len());
        for arg in args {
            cells.push(match arg {
                OxCallArg::ByRef(place)
                | OxCallArg::Named {
                    by_ref: Some(place),
                    ..
                } => Some(self.com_reference(place)?),
                OxCallArg::Operand(op) => Some(self.operand(op)?),
                OxCallArg::Named { value, .. } => Some(self.operand(value)?),
                OxCallArg::Omitted => None,
                OxCallArg::Const(n) => Some(Variant::from_i32(*n)),
            });
        }
        let call_args = com_call_args(args, &cells);
        let host = self.exec.host;
        let result = {
            let _events = self.event_scope();
            host.com()
                .invoke_early_bound_object(&object, member, &call_args)
        };
        self.take_event_abort()?;
        result.map_err(|error| Vm3Error::Fault(Fault::from_hal(error)))
    }

    /// Invoke a base-library built-in. Most builtins use the shared context-free
    /// `calcvba-lib` dispatcher, but a few are host/bundle-aware and the pure body would
    /// return a generically-wrong value: `TypeName` of an object yields the literal
    /// `"Object"` from the pure body, so resolve the real class/COM name here, where the
    /// host COM facet is in reach — never let the generic `"Object"` leak as a
    /// silently-wrong result.
    ///
    /// **This method is the intended `builtin_invoke` boundary of the future
    /// `RuntimeImports` ABI** (plan: M4). The Cranelift JIT does not re-implement builtins
    /// or this object-name special-case — it lowers `CallNative` to a `builtin_invoke`
    /// `extern "C"` shim that recovers `&mut Vm3` from its `ctx` and calls *this* method,
    /// so the interpreter and compiled code share one implementation and cannot drift.
    /// Keep its shape `(ctx, id, &[Variant]) -> Result<Variant, _>` ABI-friendly.
    ///
    /// Current VM3 hosts receive only shared HAL facets, never a VM/session callback.
    /// `StandardHostServices::DoEvents` pumps/marks queued work and returns; VM3 polls
    /// that queue from `StmtBoundary`. A custom host that stores and dereferences a
    /// `Vm3` raw pointer supplies separate unsafe authority not granted by
    /// `HostServices`; same-VM synchronous re-entry needs the future stable session
    /// root rather than pretending this safe API supports it.
    fn invoke_native_lib(
        &mut self,
        id: NativeImplId,
        argv: &[Variant],
    ) -> Result<Variant, Vm3Error> {
        self.invoke_native_lib_with_policy(id, argv, false)
    }

    fn invoke_native_lib_with_policy(
        &mut self,
        id: NativeImplId,
        argv: &[Variant],
        string_typed_alias: bool,
    ) -> Result<Variant, Vm3Error> {
        // Format$ accepts a Null value (captured empty result); other $ aliases reject Null.
        if string_typed_alias && id != NativeImplId::Format {
            calcvba_lib::validate_string_alias_args(id, argv)
                .map_err(|err| Vm3Error::Fault(Fault::from_lib(err)))?;
        }
        if id == NativeImplId::ErrorText && argv.is_empty() {
            return Ok(Variant::from_string(&self.exec.err_engine.err.description));
        }
        if let Some(class) =
            calcvba_rt_abi::object_types::builtin_dictionary_class(self.cur_program(), id, argv)
        {
            return self.new_project_instance(class);
        }
        if id == NativeImplId::TypeName
            && let Some(object) = argv.first().and_then(|a| a.as_object_ref())
            && let Some(name) = self.object_type_name(&object)
        {
            // A project instance resolves its class name from the program's class table; a COM
            // object is named by the host. Only if neither names it do we fall through to the
            // pure body (which yields the generic "Object"), exactly as vm2 does.
            return Ok(Variant::from_string(name));
        }
        if calcvba_lib::is_contextual(id) {
            let argv = calcvba_lib::prepare_contextual_args(id, argv, self.exec.host)
                .map_err(|e| Vm3Error::Fault(Fault::from_lib(e)))?;
            return calcvba_lib::invoke_contextual(id, &argv, &mut self.exec.lib)
                .map_err(|e| Vm3Error::Fault(Fault::from_lib(e)));
        }

        let host = self.exec.host;
        let result = {
            let _events = self.event_scope();
            calcvba_lib::invoke_context_free(id, argv, host)
        };
        self.take_event_abort()?;
        match result {
            Ok(value) => Ok(value),
            Err(calcvba_lib::ContextFreeInvokeError::Library(err)) => {
                Err(Vm3Error::Fault(Fault::from_lib(err)))
            }
            Err(calcvba_lib::ContextFreeInvokeError::ContextRequired) => Err(Vm3Error::Malformed(
                format!("contextual built-in {id:?} escaped VM3 dispatch classification"),
            )),
        }
    }

    /// Marshal a native built-in's arguments to plain values — a built-in reads the
    /// *value* of a ByRef argument (matching vm2's `native_args`), and an omitted one is
    /// `Empty`.
    /// `Len`/`LenB` of a VBA class object answer its default value itself, not that value's
    /// length (VBA: `Len(w)` of an object whose default is 42 is the Integer 42; probe
    /// temp/variant-layout-probe/ops13). `None` for any other call.
    fn vba_object_len(
        &mut self,
        id: NativeImplId,
        argv: &[Variant],
    ) -> Result<Option<Variant>, Vm3Error> {
        if !matches!(id, NativeImplId::Len | NativeImplId::LenB) {
            return Ok(None);
        }
        match argv.first() {
            Some(arg)
                if !is_nothing(arg)
                    && arg.as_object_ref().is_some_and(|o| o.is_project_instance()) =>
            {
                Ok(Some(self.arithmetic_value(arg.clone())?))
            }
            _ => Ok(None),
        }
    }

    /// A library function reads an object argument's default member, except those that take
    /// the object itself ([`NativeImplId::keeps_object_args`]); `VarType` reads it when there is
    /// one (an object without a default member is `vbObject`).
    fn native_object_values(
        &mut self,
        id: NativeImplId,
        argv: &mut [Variant],
    ) -> Result<(), Vm3Error> {
        let var_type = id == NativeImplId::VarType;
        if id.keeps_object_args() && !var_type {
            return Ok(());
        }
        for arg in argv.iter_mut() {
            if arg.vtype() != VarType::Object {
                continue;
            }
            if is_nothing(arg) {
                if var_type {
                    continue;
                }
                return Err(Vm3Error::Fault(Fault::new(91, default_error_message(91))));
            }
            if var_type && !project_has_default_get(arg) {
                continue;
            }
            *arg = self.arithmetic_value(arg.clone())?;
        }
        Ok(())
    }

    fn native_args(&mut self, args: &[OxCallArg]) -> Result<Vec<Variant>, Vm3Error> {
        marshal_ox_call_args(
            args,
            |arg| match arg {
                MarshalArgRef::Operand(op) => self.operand(op),
                MarshalArgRef::ByRef(place) => self.read(place),
            },
            Variant::empty,
        )
    }

    /// A `Declare Lib` call through the host's dynamic-link HAL (design doc, step 9): ByVal
    /// arguments are values, each ByRef place (and a String variable passed to a ByVal String
    /// parameter, which the binder binds ByRef for its copy-back) is the `VT_BYREF` reference a
    /// COM call gets, so the callee writes the caller's storage. Captures `Err.LastDllError`.
    fn declare_call(
        &mut self,
        descriptor_id: u32,
        args: &[OxCallArg],
    ) -> Result<Variant, Vm3Error> {
        // `cur_program()` returns `&'h OxProgram` (a reference independent of the `&mut self`
        // borrow), so the descriptor borrow does not conflict with the mutable calls below.
        let program = self.cur_program();
        let descriptor = program
            .external_calls
            .iter()
            .find(|d| d.descriptor_id == descriptor_id)
            .ok_or_else(|| {
                Vm3Error::Fault(Fault::new(
                    5,
                    format!("unknown Declare descriptor {descriptor_id}"),
                ))
            })?;
        let mut cells = Vec::with_capacity(args.len());
        for arg in args {
            cells.push(match arg {
                OxCallArg::ByRef(place)
                | OxCallArg::Named {
                    by_ref: Some(place),
                    ..
                } => Some(self.com_reference(place)?),
                OxCallArg::Operand(op) => Some(self.operand(op)?),
                OxCallArg::Named { value, .. } => Some(self.operand(value)?),
                OxCallArg::Omitted => None,
                OxCallArg::Const(n) => Some(Variant::from_i32(*n)),
            });
        }

        let param_type_strings: Vec<String> = descriptor
            .param_types
            .iter()
            .map(|pt| format!("{pt:?}"))
            .collect();
        let return_record_layout = descriptor
            .return_record_fields
            .as_ref()
            .map(|fields| calcvba_bundle::vba_record_layout_for_fields(fields))
            .transpose()
            .map_err(|message| Vm3Error::Fault(Fault::new(5, message)))?;
        let view = DynLinkDescriptorView {
            return_record_layout,
            descriptor_id: descriptor.descriptor_id,
            declared_name: &descriptor.declared_name,
            library: &descriptor.library,
            alias: &descriptor.alias,
            ordinal_alias: descriptor.ordinal_alias,
            symbol: descriptor.symbol,
            marshal_lane: &descriptor.marshal_lane,
            calling_convention: &descriptor.calling_convention,
            selection_policy: &descriptor.selection_policy,
            param_count: descriptor.param_count,
            param_types: &param_type_strings,
            param_by_ref: &descriptor.param_by_ref,
            return_type: descriptor
                .return_type
                .as_ref()
                .map(|rt| Cow::Owned(format!("{rt:?}"))),
        };

        let _callback_regs =
            self.prepare_native_callback_args(descriptor, &param_type_strings, &mut cells)?;
        let call_args = com_call_args(args, &cells);
        let host = self.exec.host;
        let invoke = {
            let _events = self.event_scope();
            host.dynlink().invoke_declare(&view, &call_args)
        };
        self.take_event_abort()?;
        // VBA refreshes `Err.LastDllError` after every `Declare` call (the OS last-error the
        // HAL captured immediately after the native call); non-native lanes report 0.
        self.exec.err_engine.last_dll_error = self.exec.host.dynlink().last_dll_error();
        invoke.map_err(|err| Vm3Error::Fault(Fault::from_hal(err)))
    }

    fn prepare_native_callback_args(
        &mut self,
        descriptor: &calcvba_bundle::ExternalCallDescriptor,
        param_type_strings: &[String],
        cells: &mut [Option<Variant>],
    ) -> Result<Vec<CallbackRegistration>, Vm3Error> {
        let mut registrations = Vec::new();
        for (index, value) in cells.iter_mut().enumerate() {
            let Some(value) = value.as_mut() else {
                continue;
            };
            let Some(proc_token) = value.as_proc_ref() else {
                continue;
            };
            let is_long_ptr = param_type_strings.get(index).map(String::as_str) == Some("LongPtr");
            let by_ref = descriptor.param_by_ref.get(index).copied().unwrap_or(false);
            if !is_long_ptr || by_ref {
                return Err(Vm3Error::Unimplemented {
                    what: "AddressOf proc passed to an unsupported Declare parameter",
                });
            }
            if !Self::is_synchronous_native_callback_descriptor(descriptor) {
                return Err(Vm3Error::Unimplemented {
                    what: "AddressOf proc passed to a non-synchronous Declare callback parameter",
                });
            }
            let owner = self as *mut Self as usize;
            let executor = NonNull::from(&mut *self);
            // SAFETY: the returned registration is kept alive until the native Declare
            // call returns, and this VM remains the same-thread callback executor for
            // the bounded synchronous CallWindowProc callback shape.
            let registration = unsafe { register_callback(owner, proc_token, executor) }.map_err(
                |err| match err {
                    calcvba_runtime::CallbackThunkError::Exhausted => {
                        Vm3Error::Fault(Fault::new(7, err.to_string()))
                    }
                    calcvba_runtime::CallbackThunkError::UnsupportedPlatform => {
                        Vm3Error::Unimplemented {
                            what: "AddressOf native callback thunks on this platform",
                        }
                    }
                },
            )?;
            *value = Variant::from_i64(registration.address() as i64);
            registrations.push(registration);
        }
        Ok(registrations)
    }

    fn is_synchronous_native_callback_descriptor(
        descriptor: &calcvba_bundle::ExternalCallDescriptor,
    ) -> bool {
        let library = descriptor.library.to_ascii_lowercase();
        let symbol = descriptor.alias.to_ascii_lowercase();
        matches!(library.as_str(), "user32" | "user32.dll")
            && matches!(
                symbol.as_str(),
                "callwindowprocw" | "callwindowproca" | "callwindowproc"
            )
    }

    fn run_native_callback_proc(&mut self, proc_token: usize, raw_args: &[isize]) -> isize {
        let proc = FuncId(proc_token);
        let target_prog = self.cur;
        let Some(callee) = self
            .exec
            .programs
            .get(target_prog)
            .and_then(|program| program.program.funcs.get(proc.0))
        else {
            return 0;
        };
        let args: Vec<Variant> = callee
            .locals
            .iter()
            .take(callee.param_count)
            .zip(raw_args.iter().copied())
            .map(|(local, raw)| Self::callback_arg_for_type(raw, &local.ty))
            .collect();
        match <Self as ProcInvoker>::invoke_callback_proc_with_values(
            self,
            target_prog,
            proc,
            args,
            true,
        ) {
            Ok(ret) => Self::callback_return_to_isize(&ret),
            // `End` in a callback ends the whole run once the Declare that called back returns.
            Err(Vm3Error::Halted) => {
                self.exec
                    .events
                    .pending_abort
                    .get_or_insert(calcvba_rt_abi::RunAbort::End);
                0
            }
            Err(_) => 0,
        }
    }

    fn callback_arg_for_type(raw: isize, ty: &OxTy) -> Variant {
        match ty {
            OxTy::Byte => Variant::from_u8(raw as u8),
            OxTy::Integer => Variant::from_i16(raw as i16),
            OxTy::Long => Variant::from_i32(raw as i32),
            OxTy::LongLong => Variant::from_i64(raw as i64),
            OxTy::Bool => Variant::from_bool(raw != 0),
            _ => Variant::from_i64(raw as i64),
        }
    }

    fn callback_return_to_isize(value: &Variant) -> isize {
        if let Some(value) = value.as_i64() {
            value as isize
        } else if let Some(value) = value.as_i32() {
            value as isize
        } else if let Some(value) = value.as_i16() {
            value as isize
        } else if let Some(value) = value.as_u8() {
            value as isize
        } else if let Some(value) = value.as_bool() {
            if value { -1 } else { 0 }
        } else {
            0
        }
    }

    fn run_callback_proc_with_values(
        &mut self,
        target_prog: usize,
        proc: FuncId,
        args: Vec<Variant>,
        suppress: bool,
    ) -> Result<Variant, Vm3Error> {
        self.guard_call_depth()?;
        let saved_cur = self.cur;
        // Suppressed native callback handlers must not overwrite the caller-visible Err object.
        let suppressed_err_engine = suppress.then(|| self.exec.err_engine.clone());
        let base = self.frames.len();
        let mut frame = self.new_frame_in(target_prog, proc)?;
        for (i, v) in args.into_iter().enumerate() {
            if let Some(slot) = frame.local_mut(i) {
                *slot = v;
            }
        }
        let saved_err = self.exec.err_engine.enter_activation();
        frame.saved_error_mode = saved_err.error_mode;
        frame.saved_active_error = saved_err.active_error;
        self.frames.push(frame);
        let result = self.run_loop(base);
        let ret = self
            .frames
            .get(base)
            .and_then(|fr| fr.return_local.and_then(|rl| fr.local(rl.0).cloned()))
            .unwrap_or_else(Variant::empty);
        let saved_err = self.frames.get(base).map(|fr| SavedErrState {
            error_mode: fr.saved_error_mode,
            active_error: fr.saved_active_error,
        });
        if let Some(saved_err) = saved_err {
            self.exec.err_engine.restore(saved_err);
        }
        self.truncate_frames_with_withevents_cleanup(base);
        self.prune_param_array_aliases_from_depth(self.frames.len());
        self.cur = saved_cur;
        self.maybe_drain();
        if let Some(saved) = suppressed_err_engine {
            self.exec.err_engine = saved;
        }
        match result {
            Ok(()) => Ok(ret),
            Err(Vm3Error::Fault(_)) if suppress => Ok(Variant::empty()),
            Err(e) => Err(e),
        }
    }

    /// Bound runaway recursion at vm2's frame ceiling, raising error 28 ("Out of stack
    /// space") as a fault, not a panic. The dispatch loop holds frames on the heap (no
    /// native recursion), so the same ceiling vm2 uses is reachable without overflow.
    fn guard_call_depth(&self) -> Result<(), Vm3Error> {
        const MAX_FRAMES: usize = 50_000;
        if self.frames.len() >= MAX_FRAMES {
            return Err(Vm3Error::Fault(Fault {
                code: 28,
                message: "Out of stack space".into(),
                source: None,
                help_file: None,
                help_context: None,
            }));
        }
        Ok(())
    }

    #[inline]
    fn clear_statement_temps(&mut self, first_temp: usize) {
        let Some(frame_index) = self.frames.len().checked_sub(1) else {
            return;
        };
        if let Some(frame) = self.frames.get_mut(frame_index) {
            let first_slot = frame.num_locals + first_temp;
            if first_slot < frame.temp_hwm {
                let end = frame.temp_hwm.min(frame.slots.len());
                for slot in &mut frame.slots[first_slot..end] {
                    if slot.raw_vt() != 0 {
                        slot.set_empty();
                    }
                }
                frame.temp_hwm = first_slot;
            }
            if !frame.pointer_strings.is_empty() {
                frame.pointer_strings.clear();
            }
            if !frame.ref_locks.is_empty() {
                frame.ref_locks.retain(|(temp, _)| *temp < first_temp);
            }
        }
        if !self.for_each.is_empty() {
            self.for_each
                .retain(|loc, _| !is_cleared_temp(*loc, frame_index, first_temp));
        }
        if !self.param_array_aliases.is_empty() {
            self.param_array_aliases.retain(|loc, aliases| {
                !is_cleared_temp(*loc, frame_index, first_temp)
                    && aliases
                        .iter()
                        .flatten()
                        .all(|alias| !is_cleared_temp(*alias, frame_index, first_temp))
            });
        }
    }

    /// Store the result of a fallible kernel op, raising its fault on error.
    fn store_arith(
        &mut self,
        dst: &OxPlace,
        out: Result<Variant, ArithError>,
    ) -> Result<(), Vm3Error> {
        let v = out.map_err(|e| Vm3Error::Fault(Fault::from_arith(e)))?;
        let non_finite_double =
            v.vtype() == VarType::Double && v.as_f64().is_some_and(|f| !f.is_finite());
        self.store(dst, v)?;
        if non_finite_double {
            return Err(Vm3Error::Fault(Fault::from_arith(ArithError::overflow())));
        }
        Ok(())
    }

    /// Resolve an [`OxPlace`] against the top frame to a concrete frame-stack [`Loc`],
    /// following a ByRef parameter's alias to its caller-side backing.
    #[inline]
    fn resolve(&self, place: &OxPlace) -> Loc {
        let top = self.frames.len() - 1;
        match place {
            OxPlace::Global(g) => Loc::Global(self.cur, g.0),
            OxPlace::Local(l) => self.frames[top]
                .alias_of(l.0)
                .unwrap_or(Loc::Local(top, l.0)),
            OxPlace::Temp(t) => Loc::Temp(top, t.0),
        }
    }

    /// Read a resolved location. Local/Global are dense, program-sized tables, so an
    /// out-of-range index is a structural defect (`Malformed`), never a silent default;
    /// `Temp` absence is the SSA write-before-read contract (sparse map → `Empty`).
    #[inline]
    fn read_loc(&self, loc: Loc) -> Result<Variant, Vm3Error> {
        match loc {
            Loc::Global(p, g) => self.exec.programs[p]
                .globals
                .get(g)
                .cloned()
                .ok_or_else(|| Vm3Error::Malformed(format!("global [{p}][{g}] out of range"))),
            Loc::Local(fi, li) => self
                .frames
                .get(fi)
                .and_then(|f| f.local(li))
                .cloned()
                .ok_or_else(|| Vm3Error::Malformed(format!("local [{fi}][{li}] out of range"))),
            Loc::Temp(fi, ti) => Ok(self
                .frames
                .get(fi)
                .and_then(|f| f.temp(ti))
                .cloned()
                .unwrap_or_else(Variant::empty)),
        }
    }

    /// Write a resolved location (same dense/sparse contract as [`Self::read_loc`]).
    #[inline]
    fn write_loc(&mut self, loc: Loc, v: Variant) -> Result<(), Vm3Error> {
        let slot = match loc {
            Loc::Global(p, g) => self.exec.programs[p]
                .globals
                .get_mut(g)
                .ok_or_else(|| Vm3Error::Malformed(format!("global [{p}][{g}] out of range")))?,
            Loc::Local(fi, li) => self
                .frames
                .get_mut(fi)
                .and_then(|f| f.local_mut(li))
                .ok_or_else(|| Vm3Error::Malformed(format!("local [{fi}][{li}] out of range")))?,
            Loc::Temp(fi, ti) => self
                .frames
                .get_mut(fi)
                .and_then(|f| f.temp_mut(ti))
                .ok_or_else(|| Vm3Error::Malformed(format!("temp [{fi}][{ti}] out of range")))?,
        };
        let old_vt = slot.raw_vt();
        let hi = old_vt & 0xF000;
        if hi == 0x2000 || hi == 0x4000 || hi == 0x6000 {
            if slot.holds_locked_array() {
                return Err(Vm3Error::Fault(Fault::new(
                    10,
                    calcvba_rt_abi::default_error_message(10),
                )));
            }
            if let Some((vt, _)) = slot.byref_target() {
                let coerced = arith::coerce_for_byref_target(&v, vt).map_err(arith_fault)?;
                let cell = self.read_loc_ref(loc)?.expect("slot existed above");
                // SAFETY: a reference cell's target is the caller's live variable, not borrowed here.
                let res = unsafe { cell.write_through(coerced) }
                    .map_err(|e| Vm3Error::Fault(Fault::new(13, e)));
                if !self.param_array_aliases.is_empty() {
                    self.param_array_aliases.remove(&loc);
                }
                return res;
            }
        }
        let old = std::mem::replace(slot, v);
        if !self.param_array_aliases.is_empty() {
            self.param_array_aliases.remove(&loc);
        }
        if !self.exec.events.withevents.is_empty() && (old_vt == 9 || old_vt == 13) {
            self.clear_withevents_owners_before_releasing_values(std::iter::once(&old));
        }
        drop(old);
        Ok(())
    }

    fn clear_withevents_owners_in_frame_before_drop(&mut self, frame: &Frame) {
        if self.exec.events.withevents.is_empty() {
            return;
        }
        self.clear_withevents_owners_before_releasing_values(frame.slots.iter());
    }

    fn truncate_frames_with_withevents_cleanup(&mut self, len: usize) {
        while self.frames.len() > len {
            let frame = self.frames.pop().expect("frame length checked before pop");
            self.recycle_frame(frame);
        }
    }

    fn clear_withevents_owners_before_releasing_values<'a>(
        &mut self,
        values: impl IntoIterator<Item = &'a Variant>,
    ) {
        if self.exec.events.withevents.is_empty() {
            return;
        }
        let mut candidates: HashMap<i32, (ObjectRef, u32)> = HashMap::new();
        for value in values {
            let Some(owner) = value.as_object_ref() else {
                continue;
            };
            if !owner.is_project_instance() {
                continue;
            }
            let owner_raw = owner.raw();
            match candidates.entry(owner_raw) {
                std::collections::hash_map::Entry::Vacant(entry) => {
                    entry.insert((owner, 1));
                }
                std::collections::hash_map::Entry::Occupied(mut entry) => {
                    entry.get_mut().1 += 1;
                }
            }
        }

        for (owner_raw, (owner, releasing_refs)) in candidates {
            let event_binding_refs = self
                .exec
                .events
                .withevents
                .keys()
                .filter(|key| withevents_owner_raw(**key) == owner_raw)
                .count() as u32;
            if event_binding_refs == 0 {
                continue;
            }
            let com_sink_refs = self
                .exec
                .events
                .com_connections
                .values()
                .filter(|sink| object_identity(&sink.owner) == owner_raw as u32 as usize)
                .count() as u32;
            let retained_event_refs = event_binding_refs + com_sink_refs;
            // `owner` is one temporary AddRef taken while inspecting the values. If every
            // remaining non-temporary reference is either in `values` or in the event fabric,
            // break the event-owner cycle before dropping `values` so the object can terminate.
            if owner.strong_count() == retained_event_refs + releasing_refs + 1 {
                self.unsubscribe_com_owner(owner_raw);
                self.exec
                    .events
                    .withevents
                    .retain(|key, _| withevents_owner_raw(*key) != owner_raw);
            }
        }
    }

    /// The address `AddrOf` yields: the storage `base` + `path` names, located by
    /// `calcvba_runtime::address` (the JIT's `rt_jit_addr_of_to_slot` does the same).
    fn address_of(
        &mut self,
        kind: PtrKind,
        base: &OxAddrBase,
        path: &[OxAddrStep],
    ) -> Result<i64, Vm3Error> {
        use calcvba_runtime::address;
        let steps = self.address_steps(path)?;
        let root = self.address_root(base)?;
        let fault =
            |number: i32| Vm3Error::Fault(Fault::new(number, default_error_message(number)));
        // SAFETY: `root` is a live, stable engine cell (a frame slot, a global or a boxed
        // class field) and no Rust reference to it is held across this call.
        let location = unsafe { address::locate(root, &steps) }.map_err(fault)?;
        self.address_from(kind, location)
    }

    /// The subscripts of an `AddrOf`/`MakeRef` path, evaluated as element access evaluates
    /// them.
    fn address_steps(
        &mut self,
        path: &[OxAddrStep],
    ) -> Result<Vec<calcvba_runtime::address::AddressStep>, Vm3Error> {
        use calcvba_runtime::address::AddressStep;
        let mut steps = Vec::with_capacity(path.len());
        for step in path {
            steps.push(match step {
                OxAddrStep::Index(indices) => {
                    let mut subscripts = Vec::with_capacity(indices.len());
                    for index in indices {
                        let value = self.operand(index)?;
                        let subscript = array_subscript(&value).map_err(arith_fault)?;
                        subscripts.push(i64::from(subscript));
                    }
                    AddressStep::Index(subscripts)
                }
                OxAddrStep::Field(index) => AddressStep::Field(*index),
            });
        }
        Ok(steps)
    }

    /// The engine cell an `AddrOf`/`MakeRef` path starts from.
    fn address_root(&mut self, base: &OxAddrBase) -> Result<*mut Variant, Vm3Error> {
        Ok(match base {
            OxAddrBase::Place(place) => {
                let loc = self.resolve(place);
                self.read_loc_mut(loc)?
                    .map(|cell| cell as *mut Variant)
                    .ok_or_else(|| Vm3Error::Malformed(format!("unaddressable {loc:?}")))?
            }
            OxAddrBase::ObjectField { object, field } => {
                let object = self.operand(object)?;
                let object = object
                    .as_object_ref()
                    .ok_or_else(|| Vm3Error::Fault(Fault::new(91, "Object variable not set")))?;
                object
                    .project_field_cell(*field)
                    .ok_or_else(|| Vm3Error::Fault(Fault::new(13, "type mismatch")))?
            }
        })
    }

    /// The pointer a helper of `kind` returns for a located piece of storage.
    fn address_from(
        &mut self,
        kind: PtrKind,
        location: calcvba_runtime::address::Location,
    ) -> Result<i64, Vm3Error> {
        use calcvba_runtime::address::{AddressKind, Location};
        let fault =
            |number: i32| Vm3Error::Fault(Fault::new(number, default_error_message(number)));
        let address = match kind {
            PtrKind::Var | PtrKind::VarString => {
                // SAFETY: the location was just found live.
                unsafe { location.address(AddressKind::Value) }
            }
            PtrKind::VarVariant => {
                // SAFETY: the location was just found live.
                unsafe { location.address(AddressKind::Cell) }
            }
            PtrKind::Str => {
                // SAFETY: the location was just found live.
                match unsafe { location.address(AddressKind::Bstr) } {
                    Some(address) => Some(address),
                    // VBA converts a non-String operand to a temporary String for the
                    // statement and passes that.
                    None => {
                        let Location::Cell(cell) = location else {
                            return Err(fault(13));
                        };
                        // SAFETY: a live engine cell, read before any other access.
                        let text = arith::as_string(unsafe { &*cell });
                        let converted = Variant::from_string(text);
                        let bstr = converted.raw_bstr_ptr().expose_provenance();
                        if let Some(frame) = self.frames.last_mut() {
                            frame.pointer_strings.push(converted);
                        }
                        Some(bstr)
                    }
                }
            }
            PtrKind::Obj => {
                // SAFETY: the location was just found live.
                if let Some(pointer) = unsafe { location.object_pointer() } {
                    Some(pointer)
                } else if let Location::Cell(cell) = location {
                    // SAFETY: a live engine cell, cloned before any other access.
                    let value = unsafe { (*cell).clone() };
                    return calcvba_rt_abi::object_variant_pointer(self.exec.host, &value)
                        .map_err(|e| Vm3Error::Fault(Fault::from_string(e)));
                } else {
                    return Err(fault(13));
                }
            }
        };
        address.map(|a| a as i64).ok_or_else(|| fault(13))
    }

    /// Borrow a resolved location's `Variant` in place (no clone) — the constant-time
    /// counterpart to [`Self::read_loc`], used by the array fast paths to read an
    /// element without deep-cloning the whole backing store. A `Temp` that has not
    /// been written (SSA write-before-read) has no stable slot to borrow and yields
    /// `None`, so the caller falls back to the cloning path.
    #[inline]
    fn read_loc_ref(&self, loc: Loc) -> Result<Option<&Variant>, Vm3Error> {
        match loc {
            Loc::Global(p, g) => self.exec.programs[p]
                .globals
                .get(g)
                .map(Some)
                .ok_or_else(|| Vm3Error::Malformed(format!("global [{p}][{g}] out of range"))),
            Loc::Local(fi, li) => self
                .frames
                .get(fi)
                .and_then(|f| f.local(li))
                .map(Some)
                .ok_or_else(|| Vm3Error::Malformed(format!("local [{fi}][{li}] out of range"))),
            Loc::Temp(fi, ti) => Ok(self.frames.get(fi).and_then(|f| f.temp(ti))),
        }
    }

    /// Mutably borrow a resolved location's `Variant` in place (no clone / writeback)
    /// — the constant-time counterpart to [`Self::write_loc`], used by the array
    /// write fast path to mutate one element through the SAFEARRAY descriptor. A
    /// `Temp` absent from the sparse map yields `None` (caller falls back).
    #[inline]
    fn read_loc_mut(&mut self, loc: Loc) -> Result<Option<&mut Variant>, Vm3Error> {
        match loc {
            Loc::Global(p, g) => self.exec.programs[p]
                .globals
                .get_mut(g)
                .map(Some)
                .ok_or_else(|| Vm3Error::Malformed(format!("global [{p}][{g}] out of range"))),
            Loc::Local(fi, li) => self
                .frames
                .get_mut(fi)
                .and_then(|f| f.local_mut(li))
                .map(Some)
                .ok_or_else(|| Vm3Error::Malformed(format!("local [{fi}][{li}] out of range"))),
            Loc::Temp(fi, ti) => Ok(self.frames.get_mut(fi).and_then(|f| f.temp_mut(ti))),
        }
    }

    /// The `VT_BYREF` cell a ByRef Variant parameter gets for `place`, as VBA binds it: a typed
    /// variable (scalar, String, object or typed array) is referred to by its storage (Excel:
    /// a Long is `&H4003` pointing at the Long), while a Variant variable is the caller's own
    /// cell (`None`: the parameter aliases it).
    fn byref_reference(
        &mut self,
        place: &OxPlace,
        param_ty: Option<&OxTy>,
    ) -> Result<Option<Variant>, Vm3Error> {
        let declared = self.place_ty(place);
        if param_ty != Some(&OxTy::Variant) || !declared.is_byref_variant_reference() {
            return Ok(None);
        }
        let loc = self.resolve(place);
        // As New must be resolved before exposing the typed variable's payload.
        if self.as_new_slots.contains_key(&loc) {
            self.read_loc_as_new(loc)?;
        }
        let Some(cell) = self.read_loc_mut(loc)? else {
            return Ok(None);
        };
        // SAFETY: the caller's frame or global slot outlives the callee and keeps its type.
        Ok(unsafe { calcvba_rt_abi::byref_variant_payload(&OxTy::Variant, &declared, cell) })
    }

    /// The declared type of `place` in the current frame.
    fn place_ty(&self, place: &OxPlace) -> OxTy {
        let frame = self.frames.last().expect("executing frame");
        let program = self.exec.programs[frame.prog].program;
        let function = &program.funcs[frame.func.0];
        match place {
            OxPlace::Local(id) => function.locals[id.0].ty.clone(),
            OxPlace::Temp(id) => function.temps[id.0].clone(),
            OxPlace::Global(id) => program.globals[id.0].ty.clone(),
        }
    }

    /// Store an assigned value: an array copied into a variable takes that variable's storage
    /// kind (`Variant::adopt_assignment_target_fixedness`), as the JIT's `rt_jit_store_variant`.
    fn store_assigned(&mut self, place: &OxPlace, mut v: Variant) -> Result<(), Vm3Error> {
        let loc = self.resolve(place);
        if (v.raw_vt() & 0xF000) == 0x2000 && !matches!(loc, Loc::Temp(..)) {
            let target = match loc {
                Loc::Global(p, g) => self.exec.programs.get(p).and_then(|p| p.globals.get(g)),
                Loc::Local(fi, li) => self.frames.get(fi).and_then(|f| f.local(li)),
                Loc::Temp(..) => None,
            };
            if let Some(target) = target {
                v.adopt_assignment_target_fixedness(target);
            }
        }
        self.write_loc(loc, v)
    }

    #[inline]
    fn store(&mut self, place: &OxPlace, v: Variant) -> Result<(), Vm3Error> {
        let loc = self.resolve(place);
        self.write_loc(loc, v)
    }

    #[inline]
    fn read(&mut self, place: &OxPlace) -> Result<Variant, Vm3Error> {
        let loc = self.resolve(place);
        self.read_loc_as_new(loc)
    }

    #[inline]
    fn read_loc_as_new(&mut self, loc: Loc) -> Result<Variant, Vm3Error> {
        let value = self.read_loc(loc)?;
        if self.as_new_slots.is_empty() {
            return Ok(value);
        }
        let Some(binding) = self.as_new_slots.get(&loc).cloned() else {
            return Ok(value);
        };
        if !is_nothing(&value) {
            return Ok(value);
        }
        let object = self.instantiate_as_new(binding)?;
        self.write_loc(loc, object.clone())?;
        Ok(object)
    }

    fn instantiate_as_new(&mut self, binding: OxAsNew) -> Result<Variant, Vm3Error> {
        self.instantiate_as_new_in_bundle(self.cur, binding)
    }

    fn instantiate_as_new_in_bundle(
        &mut self,
        bundle: usize,
        binding: OxAsNew,
    ) -> Result<Variant, Vm3Error> {
        if bundle >= self.exec.programs.len() {
            return Err(Vm3Error::Malformed(format!(
                "unknown As New owner bundle {bundle}"
            )));
        }
        let saved = self.cur;
        self.cur = bundle;
        let result = match binding {
            OxAsNew::ProjectClass { class } => self.new_project_instance(class.0),
            OxAsNew::ExternClass { import } => self.new_extern_instance(import),
            OxAsNew::ComClass { prog_id } => {
                self.invoke_native_lib(NativeImplId::CreateObject, &[Variant::from_string(prog_id)])
            }
        };
        self.cur = saved;
        result
    }

    fn class_field_as_new_binding(&self, object: &ObjectRef, field: i32) -> Option<OxAsNew> {
        if !object.is_project_instance() || object.route_key() == VBA_COLLECTION_ROUTE_KEY {
            return None;
        }
        self.exec
            .programs
            .get(object.bundle_id() as usize)?
            .program
            .classes
            .get(object.route_key() as usize)?
            .as_new_fields
            .iter()
            .find(|candidate| candidate.field == field)
            .map(|candidate| candidate.binding.clone())
    }

    fn read_project_field_as_new(
        &mut self,
        instance: &ObjectRef,
        field: i32,
    ) -> Result<Variant, Vm3Error> {
        let value = instance
            .project_field_get(field)
            .unwrap_or_else(Variant::empty);
        let Some(binding) = self.class_field_as_new_binding(instance, field) else {
            return Ok(value);
        };
        if !is_nothing(&value) {
            return Ok(value);
        }
        let object = self.instantiate_as_new_in_bundle(instance.bundle_id() as usize, binding)?;
        instance.project_field_set(field, object.clone());
        Ok(object)
    }

    fn operand(&mut self, op: &OxOperand) -> Result<Variant, Vm3Error> {
        match op {
            OxOperand::Const(c) => Ok(ox_const_variant(c)),
            OxOperand::Use(p) => self.read(p),
        }
    }
}

impl calcvba_runtime::events::EventExecutor for Vm3<'_> {
    /// Runs the WithEvents handler for a COM event, synchronously inside the host call that
    /// delivered it (design doc, step 8c). An error the handler leaves unhandled ends the run
    /// once that call returns (`take_event_abort`) and goes back to the source as the event's
    /// failure.
    fn deliver_event(
        &mut self,
        key: u64,
        dispid: i32,
        args: &[Variant],
    ) -> Result<bool, calcvba_runtime::events::EventFault> {
        let call = self.exec.com_event_call(key, dispid, args);
        let Some(call) = call else {
            return Ok(false);
        };
        let handler_result = self.run_proc_with_values(
            call.program,
            FuncId(call.handler),
            call.me,
            call.args,
            false,
        );
        match handler_result {
            Ok(_) => Ok(true),
            // `End` in the handler ends the whole run once the host call returns; the source
            // sees the event as handled (Excel, probe events8end).
            Err(Vm3Error::Halted) => {
                self.exec
                    .events
                    .pending_abort
                    .get_or_insert(calcvba_rt_abi::RunAbort::End);
                Ok(true)
            }
            Err(error) => {
                let fault = match error {
                    Vm3Error::Fault(fault) | Vm3Error::EventAbort(fault) => fault,
                    other => Fault::new(51, other.to_string()),
                };
                let reported = calcvba_runtime::events::EventFault {
                    number: fault.code,
                    description: fault.message.clone(),
                    source: fault.source.clone().unwrap_or_default(),
                };
                self.exec
                    .events
                    .pending_abort
                    .get_or_insert(calcvba_rt_abi::RunAbort::Fault(fault));
                Err(reported)
            }
        }
    }

    fn invoke_member(
        &mut self,
        object: ObjectRef,
        member: &'static calcvba_runtime::RuntimeMemberDescriptor,
        invoke_kind: calcvba_runtime::RuntimeMemberInvokeKind,
        args: &mut [Variant],
        aliases: &[usize],
    ) -> Result<Variant, calcvba_runtime::events::MemberFault> {
        match self.invoke_member_descriptor_captured(object, member, invoke_kind, args, aliases) {
            Ok(value) => Ok(value),
            Err(Vm3Error::Halted) => {
                self.exec
                    .events
                    .pending_abort
                    .get_or_insert(calcvba_rt_abi::RunAbort::End);
                Ok(Variant::empty())
            }
            Err(error) => {
                let fault = match error {
                    Vm3Error::Fault(fault) | Vm3Error::EventAbort(fault) => fault,
                    other => Fault::new(51, other.to_string()),
                };
                self.exec.err_engine.pending_fault = None;
                Err(calcvba_runtime::events::MemberFault {
                    number: fault.code,
                    description: fault.message,
                    source: fault.source.unwrap_or_default(),
                    help_file: fault.help_file.unwrap_or_default(),
                    help_context: fault.help_context.unwrap_or(0),
                })
            }
        }
    }
}

impl CallbackExecutor for Vm3<'_> {
    fn invoke_callback(&mut self, proc_token: usize, args: &[isize]) -> isize {
        self.run_native_callback_proc(proc_token, args)
    }
}

impl<'h> ProcInvoker for Vm3<'h> {
    type Error = Vm3Error;

    fn invoke_proc_with_values(
        &mut self,
        target_prog: usize,
        proc: FuncId,
        me: Variant,
        args: Vec<Variant>,
        suppress: bool,
    ) -> Result<Variant, Self::Error> {
        self.run_proc_with_values(target_prog, proc, me, args, suppress)
    }

    fn invoke_callback_proc_with_values(
        &mut self,
        target_prog: usize,
        proc: FuncId,
        args: Vec<Variant>,
        suppress: bool,
    ) -> Result<Variant, Self::Error> {
        self.run_callback_proc_with_values(target_prog, proc, args, suppress)
    }

    fn maybe_drain(&mut self) {
        Vm3::maybe_drain(self);
    }
}

/// The default VBA message for a run-time error code, used as `Err.Description` when a
/// raised error has no explicit Description.
fn default_error_message(code: i32) -> String {
    calcvba_runtime::default_error_message(code).to_string()
}

fn cmp_op(op: CmpOp) -> arith::CmpOp {
    match op {
        CmpOp::Eq => arith::CmpOp::Eq,
        CmpOp::Ne => arith::CmpOp::Ne,
        CmpOp::Lt => arith::CmpOp::Lt,
        CmpOp::Le => arith::CmpOp::Le,
        CmpOp::Gt => arith::CmpOp::Gt,
        CmpOp::Ge => arith::CmpOp::Ge,
    }
}

/// Wrap an `arith` coercion error as a routed vm3 fault (it carries its own VBA code).
fn arith_fault(e: ArithError) -> Vm3Error {
    Vm3Error::Fault(Fault::from_arith(e))
}

#[inline]
fn try_fast_i32(v: &Variant) -> Option<i32> {
    match v.raw_vt() {
        3 => v.as_i32(),
        2 => v.as_i16().map(i32::from),
        0 => Some(0),
        _ => None,
    }
}

#[inline]
fn try_fast_f64(v: &Variant) -> Option<f64> {
    match v.raw_vt() {
        5 => v.as_f64(),
        3 => v.as_i32().map(f64::from),
        2 => v.as_i16().map(f64::from),
        0 => Some(0.0),
        _ => None,
    }
}

#[inline]
fn fast_i32_binop(a: i32, b: i32, op: ArithOp) -> Result<i32, Vm3Error> {
    match op {
        ArithOp::Add => a
            .checked_add(b)
            .ok_or_else(|| arith_fault(ArithError::overflow())),
        ArithOp::Sub => a
            .checked_sub(b)
            .ok_or_else(|| arith_fault(ArithError::overflow())),
        ArithOp::Mul => a
            .checked_mul(b)
            .ok_or_else(|| arith_fault(ArithError::overflow())),
        ArithOp::IntDiv => {
            if b == 0 {
                return Err(arith_fault(ArithError::div_by_zero()));
            }
            a.checked_div(b)
                .ok_or_else(|| arith_fault(ArithError::overflow()))
        }
        ArithOp::Mod => {
            if b > 0 {
                Ok(a % b)
            } else if b == 0 {
                Err(arith_fault(ArithError::div_by_zero()))
            } else {
                Ok(if a == i32::MIN && b == -1 { 0 } else { a % b })
            }
        }
    }
}

#[inline]
fn eval_arith_op(
    l: &Variant,
    r: &Variant,
    op: ArithOp,
    mode: NumericMode,
) -> Result<Variant, ArithError> {
    match op {
        ArithOp::Add => arith::add(l, r, mode),
        ArithOp::Sub => arith::sub(l, r, mode),
        ArithOp::Mul => arith::mul(l, r, mode),
        ArithOp::IntDiv => arith::int_div(l, r, mode),
        ArithOp::Mod => arith::modulo(l, r, mode),
    }
}

#[inline]
fn try_fast_cmp(lv: &Variant, rv: &Variant, op: CmpOp) -> Option<bool> {
    let lvt = lv.raw_vt();
    let rvt = rv.raw_vt();
    if matches!(lvt, 0 | 2 | 3) && matches!(rvt, 0 | 2 | 3) {
        let a = try_fast_i32(lv)?;
        let b = try_fast_i32(rv)?;
        return Some(match op {
            CmpOp::Eq => a == b,
            CmpOp::Ne => a != b,
            CmpOp::Lt => a < b,
            CmpOp::Le => a <= b,
            CmpOp::Gt => a > b,
            CmpOp::Ge => a >= b,
        });
    }
    if matches!(lvt, 0 | 2 | 3 | 5) && matches!(rvt, 0 | 2 | 3 | 5) {
        let a = try_fast_f64(lv)?;
        let b = try_fast_f64(rv)?;
        if a.is_nan() || b.is_nan() {
            return None;
        }
        return Some(match op {
            CmpOp::Eq => a == b,
            CmpOp::Ne => a != b,
            CmpOp::Lt => a < b,
            CmpOp::Le => a <= b,
            CmpOp::Gt => a > b,
            CmpOp::Ge => a >= b,
        });
    }
    None
}

#[inline]
fn is_possibly_nothing_for_as_new(v: &Variant) -> bool {
    match v.raw_vt() {
        0 | 1 => true,
        9 | 13 => borrow_project_instance_info(v).is_none(),
        2 => v.as_i16() == Some(0),
        3 => v.as_i32() == Some(0),
        _ => false,
    }
}

#[inline]
fn format_i32_utf16(n: i32, buf: &mut [u16; 12]) -> &[u16] {
    let neg = n < 0;
    let mut mag = n.unsigned_abs();
    let mut pos = buf.len();
    loop {
        pos -= 1;
        buf[pos] = u16::from(b'0' + (mag % 10) as u8);
        mag /= 10;
        if mag == 0 {
            break;
        }
    }
    if neg {
        pos -= 1;
        buf[pos] = u16::from(b'-');
    }
    &buf[pos..]
}

enum FastArrScalar {
    I32(i32),
    F64(f64),
}

/// Fast-path parser for plain ASCII signed integer strings (`^[+-]?[0-9]{1,15}$`)
/// into exact IEEE-754 `f64` (used by `ArithWidening` for cases like `v * "2"`).
#[inline]
fn try_parse_ascii_int_f64(units: &[u16]) -> Option<f64> {
    if units.is_empty() || units.len() > 16 {
        return None;
    }
    let (neg, digits) = match units[0] {
        0x002D => (true, &units[1..]),
        0x002B => (false, &units[1..]),
        _ => (false, units),
    };
    if digits.is_empty() || digits.len() > 15 {
        return None;
    }
    let mut acc: i64 = 0;
    for &u in digits {
        if !(0x0030..=0x0039).contains(&u) {
            return None;
        }
        acc = acc * 10 + i64::from(u - 0x0030);
    }
    let signed = if neg { -acc } else { acc };
    Some(signed as f64)
}

#[inline]
fn write_fast_coerced_project_scalar(
    src: &Variant,
    formal: &calcvba_oxir::OxLocal,
    dst: &mut Variant,
) -> bool {
    let vt = src.raw_vt();
    match (&formal.ty, vt) {
        (OxTy::Variant, _) if Variant::raw_vt_is_trivial_copy(vt) => dst.copy_trivial_from(src),
        (OxTy::Long, 3) => {
            dst.set_i32(src.as_i32().unwrap());
            true
        }
        (OxTy::Double, 5) if src.as_f64().is_some_and(f64::is_finite) => {
            dst.set_f64(src.as_f64().unwrap());
            true
        }
        (OxTy::Double, 3) => {
            dst.set_f64(f64::from(src.as_i32().unwrap()));
            true
        }
        (OxTy::Double, 2) => {
            dst.set_f64(f64::from(src.as_i16().unwrap()));
            true
        }
        (OxTy::Bool, 11) => {
            dst.set_bool(src.as_bool().unwrap());
            true
        }
        (OxTy::Integer, 2) => dst.copy_trivial_from(src),
        _ => false,
    }
}

#[inline]
fn fast_coerce_project_arg(
    value: Variant,
    formal: &calcvba_oxir::OxLocal,
) -> Result<Variant, Vm3Error> {
    let vt = value.raw_vt();
    match (&formal.ty, vt) {
        (OxTy::Variant, _) => Ok(value),
        (OxTy::Long, 3) => Ok(value),
        (OxTy::Double, 5) if value.as_f64().is_some_and(f64::is_finite) => Ok(value),
        (OxTy::Double, 3) => Ok(Variant::from_f64(f64::from(value.as_i32().unwrap()))),
        (OxTy::Double, 2) => Ok(Variant::from_f64(f64::from(value.as_i16().unwrap()))),
        (OxTy::Str, 8) => Ok(value),
        (OxTy::Bool, 11) => Ok(value),
        (OxTy::Integer, 2) => Ok(value),
        _ => calcvba_rt_abi::coerce_project_argument(&value, formal).map_err(arith_fault),
    }
}

/// Backing-store ceiling for an element-wise (Variant/String/Object/record)
/// array. Each such element is a `Variant` slot built individually, so a
/// Long-range element count can demand tens of gigabytes and either abort the
/// host on an infallible allocation or, on an over-committing OS where a huge
/// reservation "succeeds", stall/OOM-kill it while the elements are built. VBA
/// itself raises Out of memory (7) for such arrays on ordinary machines, so we
/// reject clearly-pathological sizes up front. The limit is deliberately far
/// above any realistic in-memory VBA array (~350M `Variant` slots).
const MAX_ELEMENTWISE_ARRAY_BYTES: u64 = 8 * 1024 * 1024 * 1024;

/// Build `count` default elements for a (re)dimensioned array. A
/// `map(..).collect()` reserves capacity infallibly, so a large `ReDim` of a
/// Variant/String/Object/record element (e.g. `ReDim v(0 To 2000000000)`) would
/// abort the whole host process. A pathological size is rejected up front, and
/// the remaining allocation is reserved fallibly, so both surface as a catchable
/// Out of memory (7), matching VBA. (Scalar numeric elements take the separate
/// zeroed-SAFEARRAY path, which is a single fallible buffer allocation.)
fn try_build_default_elements(
    element: &ArrayElementType,
    count: usize,
) -> Result<Vec<Variant>, Vm3Error> {
    let bytes = (count as u64).saturating_mul(std::mem::size_of::<Variant>() as u64);
    if bytes > MAX_ELEMENTWISE_ARRAY_BYTES {
        return Err(Vm3Error::Fault(Fault::new(7, "Out of memory")));
    }
    let mut out: Vec<Variant> = Vec::new();
    out.try_reserve_exact(count)
        .map_err(|_| Vm3Error::Fault(Fault::new(7, "Out of memory")))?;
    for _ in 0..count {
        out.push(default_array_element(element).map_err(|e| Vm3Error::Fault(Fault::new(13, e)))?);
    }
    Ok(out)
}

/// VBA `Nothing`/empty test (mirrors vm2): a null object reference, `Empty`/`Null`, or a
/// numeric zero (the literal-0-as-Nothing representation). Used by `Let`/`Set` validation.
fn is_nothing(value: &Variant) -> bool {
    match value.vtype() {
        VarType::Object => value
            .as_object_ref()
            .is_none_or(|object| object.is_nothing_placeholder()),
        VarType::Empty | VarType::Null => true,
        _ => value.as_i16() == Some(0) || value.as_i32() == Some(0),
    }
}

/// The raw identity (an `i32`) of an object value, or 0 for a non-object/`Nothing`.
fn object_identity(value: &Variant) -> usize {
    value.as_object_ref().map(|o| o.is_identity()).unwrap_or(0)
}

/// VBA's `Is` operator accepts only object references at run time. A Variant
/// that currently holds a scalar compiles, but evaluating the comparison raises
/// error 424 instead of treating the scalar as the null object identity.
fn object_identity_for_is(value: &Variant) -> Result<usize, Vm3Error> {
    if value.vtype() == VarType::Object {
        Ok(object_identity(value))
    } else {
        Err(Vm3Error::Fault(Fault::new(424, "Object required")))
    }
}

/// The `withevents` map key: the sink owner's identity in the high 32 bits, the binding token
/// in the low 32 (mirrors vm2). One sink can hold several `WithEvents` sources, one per token.
fn withevents_key(owner: &ObjectRef, binding: i64) -> i64 {
    (i64::from(owner.raw()) << 32) | (binding & 0xFFFF_FFFF)
}
/// The sink owner identity recovered from a `withevents` key.
fn withevents_owner_raw(key: i64) -> i32 {
    (key >> 32) as i32
}
/// The binding token recovered from a `withevents` key.
fn withevents_binding(key: i64) -> i64 {
    key & 0xFFFF_FFFF
}

/// The project accessor kind a late-bound call's COM invoke-kind selects (the inverse of the
/// elaboration's `invoke_kind_from_member_kind`).
fn project_member_kind(k: TypeLibMemberInvokeKind) -> ProjectMemberKind {
    match k {
        TypeLibMemberInvokeKind::PropertyGet => ProjectMemberKind::PropertyGet,
        TypeLibMemberInvokeKind::PropertyPut => ProjectMemberKind::PropertyLet,
        TypeLibMemberInvokeKind::PropertyPutRef => ProjectMemberKind::PropertySet,
        TypeLibMemberInvokeKind::Method => ProjectMemberKind::Method,
    }
}

/// The dynamic COM dispatch kind (the late-bound `IDispatch::Invoke` flag hint) a call-site
/// accessor selects. Mirrors vm2's `member_kind_to_dynamic` composed with the invoke-kind→
/// member-kind mapping: `Put`→`Let`, `PutRef`→`Set`.
/// The member of a COM object call.
fn com_member(member: &DynamicMemberSelector) -> calcvba_com::call::ComMember<'_> {
    match member {
        DynamicMemberSelector::Name(name) | DynamicMemberSelector::TokenNamed { name, .. } => {
            calcvba_com::call::ComMember::Name(name)
        }
        DynamicMemberSelector::DefaultMember => calcvba_com::call::ComMember::Id(0),
        DynamicMemberSelector::Token(token) => calcvba_com::call::ComMember::Id(*token),
    }
}

/// What a COM object call does with its member.
fn com_call_kind(kind: TypeLibMemberInvokeKind) -> calcvba_com::call::ComCallKind {
    match kind {
        TypeLibMemberInvokeKind::Method => calcvba_com::call::ComCallKind::Method,
        TypeLibMemberInvokeKind::PropertyGet => calcvba_com::call::ComCallKind::Get,
        TypeLibMemberInvokeKind::PropertyPut => calcvba_com::call::ComCallKind::Let,
        TypeLibMemberInvokeKind::PropertyPutRef => calcvba_com::call::ComCallKind::Set,
    }
}

/// Whether `arg` passes a place ByRef: a positional `ByRef`, or a late-bound named argument
/// over an l-value (Excel passes `o.Bump n:=lk` ByRef, as `o.Bump lk`).
fn arg_is_by_ref(arg: &OxCallArg) -> bool {
    matches!(
        arg,
        OxCallArg::ByRef(_)
            | OxCallArg::Named {
                by_ref: Some(_),
                ..
            }
    )
}

/// The argument list of a COM object call over the cells built for `args` (`None` for an
/// omitted argument; a ByRef place's cell is its reference).
fn com_call_args<'a>(
    args: &'a [OxCallArg],
    cells: &'a [Option<Variant>],
) -> Vec<calcvba_com::call::ComArg<'a>> {
    args.iter()
        .zip(cells)
        .map(|(arg, cell)| calcvba_com::call::ComArg {
            value: match (arg, cell) {
                (arg, Some(reference)) if arg_is_by_ref(arg) => {
                    calcvba_com::call::ComArgValue::Reference(reference)
                }
                (_, Some(value)) => calcvba_com::call::ComArgValue::Value(value),
                (_, None) => calcvba_com::call::ComArgValue::Missing,
            },
            name: match arg {
                OxCallArg::Named { name, .. } => Some(name.as_str()),
                _ => None,
            },
        })
        .collect()
}

fn variant_to_object(value: &Variant) -> Result<ObjectRef, Vm3Error> {
    object_receiver(value).map_err(Vm3Error::Fault)
}

/// The value an `Optional` parameter a call leaves out starts with: its declared default, else
/// Missing for a Variant, else its type's zero (`Optional n As Long` is 0, as a declared local
/// is). `None` for a parameter that is not optional (or a `ParamArray`), and for an optional
/// object, record or array, which keep their slot's initial value.
pub fn omitted_param_value(local: &calcvba_oxir::OxLocal) -> Option<Result<Variant, Vm3Error>> {
    let info = local.param.as_ref()?;
    if !info.optional || info.variadic {
        return None;
    }
    if let Some(default) = &info.default {
        return Some(param_default_value(default, &local.ty));
    }
    if local.ty == OxTy::Variant {
        return Some(Ok(Variant::from_error_code(MISSING_ARG)));
    }
    scalar_zero(&local.ty).map(Ok)
}

/// An omitted `Optional` parameter's declared default, coerced to the parameter type as the
/// binder coerces the default it fills in (`Optional n As Long = 5` is Long 5; a Variant
/// parameter keeps the constant's own type, `= 3` is Integer 3).
fn param_default_value(default: &OxConst, ty: &OxTy) -> Result<Variant, Vm3Error> {
    let value = ox_const_variant(default);
    match ty {
        OxTy::Str => arith::coerce_string(&value),
        OxTy::FixedStr(len) => arith::coerce_fixed_string(&value, *len as usize),
        ty => match calcvba_oxir::passes::numeric_target(ty) {
            Some(target) => arith::coerce_numeric(&value, target),
            None => Ok(value),
        },
    }
    .map_err(arith_fault)
}

fn variant_matches_ox_ty(value: &Variant, ty: &OxTy) -> bool {
    match ty {
        OxTy::Variant => true,
        OxTy::Bool => value.vtype() == VarType::Boolean,
        OxTy::Byte => value.vtype() == VarType::Byte,
        OxTy::Integer => value.vtype() == VarType::Integer,
        OxTy::Long => value.vtype() == VarType::Long,
        OxTy::LongLong => value.vtype() == VarType::LongLong,
        OxTy::Single => value.vtype() == VarType::Single,
        OxTy::Double => value.vtype() == VarType::Double,
        OxTy::Currency => value.vtype() == VarType::Currency,
        OxTy::Date => value.vtype() == VarType::Date,
        OxTy::Decimal => value.vtype() == VarType::Decimal,
        OxTy::Str | OxTy::FixedStr(_) => value.vtype() == VarType::String,
        OxTy::Object(_) => value.vtype() == VarType::Object,
        OxTy::Record(_) => value.vtype() == VarType::Record,
        OxTy::Array(_, _) => value.vtype() == VarType::ArrayVariant,
        OxTy::ProcRef => value.vtype() == VarType::ProcRef,
    }
}

/// A short label for an instruction kind (for the `Unimplemented` message).
fn inst_kind(inst: &OxInst) -> &'static str {
    match inst {
        OxInst::CallProc { .. } => "CallProc",
        OxInst::CallProcRef { .. } => "CallProcRef",
        OxInst::CallExtern { .. } => "CallExtern",
        OxInst::CallNative { .. } => "CallNative (builtin/Declare)",
        OxInst::CallByName { .. } => "CallByName",
        OxInst::ComCallEarly { .. } => "ComCallEarly",
        OxInst::ComCallLate { .. } => "ComCallLate",
        OxInst::Box { .. } => "Box",
        OxInst::Unbox { .. } => "Unbox",
        OxInst::ValidateAssignment { .. } => "ValidateAssignment",
        OxInst::LoadProcRef { .. } => "LoadProcRef",
        OxInst::CompareObjectIs { .. } => "CompareObjectIs",
        OxInst::TypeOfIs { .. } => "TypeOfIs",
        OxInst::NewObject { .. } | OxInst::NewExtern { .. } => "New",
        OxInst::Predeclared { .. }
        | OxInst::PredeclaredExtern { .. }
        | OxInst::PredeclaredSet { .. }
        | OxInst::PredeclaredExternSet { .. } => "Predeclared",
        OxInst::NewRecord { .. } => "NewRecord",
        OxInst::FieldGet { .. } | OxInst::FieldSet { .. } => "object field access",
        OxInst::RecordGet { .. }
        | OxInst::RecordSet { .. }
        | OxInst::RecordLSet { .. }
        | OxInst::RecordArrayGet { .. }
        | OxInst::RecordArraySet { .. } => "record field access",
        OxInst::ArrayLiteral { .. }
        | OxInst::ArrayAppend { .. }
        | OxInst::ArrayRedim { .. }
        | OxInst::ArrayGet { .. }
        | OxInst::ArraySet { .. }
        | OxInst::ArrayErase { .. }
        | OxInst::Bound { .. } => "array op",
        OxInst::ForEachInit { .. } | OxInst::ForEachNext { .. } => "For Each",
        OxInst::WithEventsGet { .. }
        | OxInst::WithEventsSet { .. }
        | OxInst::WithEventsClearOwner { .. }
        | OxInst::WithEventsFirstOwner { .. }
        | OxInst::WithEventsNextOwner { .. } => "WithEvents",
        OxInst::RaiseEvent { .. } => "RaiseEvent",
        OxInst::AddrOf { .. } => "pointer helper",
        OxInst::MakeRef { .. } | OxInst::ReleaseRef { .. } => "ByRef reference",
        OxInst::ForEachEnd { .. } => "For Each",
        OxInst::ErrFieldGet { .. } => "Err field read",
        OxInst::ErlGet { .. } => "Erl read",
        OxInst::ErrFieldSet { .. } => "Err field write",
        OxInst::SetErrorHandler(_) => "On Error",
        OxInst::SetLineNumber { .. } => "line number label",
        OxInst::AddRef { .. } | OxInst::Release { .. } => "refcount effect",
        OxInst::DrainTerminations => "DrainTerminations",
        // The handled instructions never reach here.
        _ => "instruction",
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::ffi::c_void;
    use std::sync::{
        Arc, Mutex,
        atomic::{AtomicBool, AtomicU32, Ordering},
    };

    use calcvba_bundle::coreir::{
        ClassId as CoreClassId, CoreArg, CoreAsNew, CoreBinOp, CoreCallee, CoreClass,
        CoreClassAsNewField, CoreClassField, CoreClassMethod, CoreConst, CoreGlobal, CoreLocal,
        CoreParam, CorePlace, CoreProc, CoreProgram, CoreStmt, CoreValue, ErrorOp, ExitKind,
        LabelId, LocalId as CoreLocalId, ProcId,
    };
    use calcvba_bundle::{
        AssignmentIntent, AssignmentTargetKind, BuiltinType, DeclareParamType,
        ExternalCallDescriptor, NativeImplId, NumericCoerceTarget, NumericMode, ProcedureKind,
        ProjectMemberKind, StringCompareMode, VarTypeRef,
    };
    use calcvba_com::{ComCallbackPayload, ComCallbackToken, ComMemberToken, ComSubscriptionToken};
    use calcvba_hal::HostPolicy;
    use calcvba_hal::adapters::null::NullHostServices;
    use calcvba_hal::error::{HalError, HalResult};
    use calcvba_hal::model::{CapabilityId, HalDescriptor, HalProfileId};
    use calcvba_hal::traits::{
        ComHal, ConsoleHal, DiagnosticsHal, DynamicLinkHal, EventPumpHal, FileSystemHal,
        ProcessEnvHal, TimeLocaleHal, TypeLibCacheScope, TypeLibMetadataBlob,
        TypeLibResolveRequest, TypeLibResolvedIdentity, UiInteractionHal,
    };
    use calcvba_oxir::program::{OxFunc, OxLocal, OxParamInfo};
    use calcvba_oxir::ty::OxTy;
    use calcvba_runtime::DynLinkSymbol;
    use calcvba_runtime::object_ref::{
        Guid, IUnknownVtbl, RUNTIME_E_NOINTERFACE, RawRuntimeIUnknown,
    };
    use calcvba_runtime::variant::VarType;

    #[test]
    fn subscript_to_long_overflows_beyond_long_range() {
        // In-range Long values pass through unchanged.
        assert_eq!(array_subscript(&Variant::from_i64(0)).unwrap(), 0);
        assert_eq!(
            array_subscript(&Variant::from_i64(i64::from(i32::MAX))).unwrap(),
            i32::MAX
        );
        assert_eq!(
            array_subscript(&Variant::from_i64(i64::from(i32::MIN))).unwrap(),
            i32::MIN
        );
        // Anything outside Long range is Overflow (6), never a silent wrap.
        for v in [
            i64::from(i32::MAX) + 1,
            4_294_967_296, // 2^32 used to wrap to 0
            4_294_967_299, // 2^32 + 3 used to wrap to 3
            i64::from(i32::MIN) - 1,
            i64::MAX,
            i64::MIN,
        ] {
            match array_subscript(&Variant::from_i64(v)) {
                Err(e) => assert_eq!(e.code, 6, "value {v} should overflow"),
                other => panic!("expected Overflow(6) for {v}, got {other:?}"),
            }
        }
    }

    #[test]
    fn try_build_default_elements_reports_oom_not_abort() {
        // A count whose backing store cannot be allocated must surface a
        // catchable Out of memory (7) rather than an infallible allocation that
        // aborts the host. `usize::MAX` overflows the byte size deterministically
        // on every platform without attempting a real large allocation.
        match try_build_default_elements(&ArrayElementType::Variant, usize::MAX) {
            Err(Vm3Error::Fault(f)) => assert_eq!(f.code, 7),
            other => panic!("expected Out of memory (7), got {other:?}"),
        }
        // A small count still succeeds and yields the requested number of defaults.
        let ok = try_build_default_elements(&ArrayElementType::Long, 4).unwrap();
        assert_eq!(ok.len(), 4);
    }

    /// Bind-free: hand-build a `CoreProgram`, elaborate it to OxIR, run it on vm3, and
    /// read back a snapshot slot.
    fn run_core(prog: &CoreProgram) -> Vm3<'_> {
        // Leak the elaborated program so the returned VM can borrow it for the test.
        let oxp: &'static OxProgram = Box::leak(Box::new(
            calcvba_oxir::elaborate::elaborate(prog).expect("elaborate"),
        ));
        let host: &'static NullHostServices =
            Box::leak(Box::new(NullHostServices::new(HostPolicy::default())));
        Vm3::run(oxp, host).expect("vm3 run")
    }

    fn assign(place: CorePlace, value: CoreValue) -> CoreStmt {
        CoreStmt::Assign {
            place,
            value,
            intent: AssignmentIntent::Let,
            target_kind: AssignmentTargetKind::Scalar,
            target_name: String::new(),
            target_type_name: String::new(),
        }
    }

    fn local(name: &str, ty: VarTypeRef) -> CoreLocal {
        CoreLocal {
            name: name.into(),
            ty,
            array_element: None,
        }
    }

    fn main_proc(locals: Vec<CoreLocal>, body: Vec<CoreStmt>) -> CoreProgram {
        CoreProgram {
            long_ptr_width: Default::default(),
            procs: vec![CoreProc {
                name: "Main".into(),
                kind: ProcedureKind::Sub,
                params: Vec::new(),
                locals,
                return_local: None,
                label_lines: Vec::new(),
                body,
            }],
            entry: Some(ProcId(0)),
            unit_name: "T".into(),
            ..Default::default()
        }
    }

    /// A multi-proc program whose `procs[0]` is the entry (`Main`).
    fn procs_program(procs: Vec<CoreProc>) -> CoreProgram {
        CoreProgram {
            long_ptr_width: Default::default(),
            procs,
            entry: Some(ProcId(0)),
            unit_name: "T".into(),
            ..Default::default()
        }
    }

    fn proc(
        name: &str,
        kind: ProcedureKind,
        params: Vec<CoreParam>,
        locals: Vec<CoreLocal>,
        return_local: Option<CoreLocalId>,
        body: Vec<CoreStmt>,
    ) -> CoreProc {
        CoreProc {
            name: name.into(),
            kind,
            params,
            locals,
            return_local,
            label_lines: Vec::new(),
            body,
        }
    }

    fn long_param(name: &str) -> CoreParam {
        CoreParam {
            default: None,
            name: name.into(),
            ty: VarTypeRef::Builtin(BuiltinType::Long),
            optional: false,
            by_ref: true,
            variadic: false,
        }
    }

    fn byval_long_param(name: &str) -> CoreParam {
        CoreParam {
            default: None,
            name: name.into(),
            ty: VarTypeRef::Builtin(BuiltinType::Long),
            optional: false,
            by_ref: false,
            variadic: false,
        }
    }

    fn optional_long_param(name: &str) -> CoreParam {
        CoreParam {
            default: None,
            name: name.into(),
            ty: VarTypeRef::Builtin(BuiltinType::Long),
            optional: true,
            by_ref: true,
            variadic: false,
        }
    }

    fn ox_param_local(name: &str) -> OxLocal {
        OxLocal {
            name: name.into(),
            ty: OxTy::Long,
            array_element: None,
            param: Some(OxParamInfo {
                optional: false,
                by_ref: false,
                variadic: false,
                default: None,
            }),
            escaped: false,
        }
    }

    fn ox_func_with_params(name: &str, locals: Vec<OxLocal>, param_count: usize) -> OxFunc {
        OxFunc {
            name: name.into(),
            kind: ProcedureKind::Sub,
            locals,
            temps: Vec::new(),
            param_count,
            return_local: None,
            blocks: Vec::new(),
            entry: BlockId(0),
        }
    }

    #[test]
    fn runtime_member_params_skips_only_explicit_hidden_me_receiver() {
        let with_me = ox_func_with_params(
            "Touch",
            vec![ox_param_local("Me"), ox_param_local("value")],
            2,
        );
        let params = calcvba_rt_abi::runtime_member_params(&with_me);
        assert_eq!(params.len(), 1);
        assert_eq!(params[0].name, "value");

        let without_me = ox_func_with_params("Touch", vec![ox_param_local("value")], 1);
        let params = calcvba_rt_abi::runtime_member_params(&without_me);
        assert_eq!(params.len(), 1);
        assert_eq!(params[0].name, "value");

        let mut non_param_me = ox_param_local("Me");
        non_param_me.param = None;
        let params = calcvba_rt_abi::runtime_member_params(&ox_func_with_params(
            "Touch",
            vec![non_param_me],
            1,
        ));
        assert_eq!(params.len(), 1);
        assert_eq!(params[0].name, "Me");
    }

    #[test]
    fn project_member_named_args_skip_only_explicit_hidden_me_receiver() {
        let twice = proc(
            "Twice",
            ProcedureKind::Function,
            vec![long_param("me"), long_param("x")],
            vec![local("Twice", VarTypeRef::Builtin(BuiltinType::Long))],
            Some(CoreLocalId(2)),
            Vec::new(),
        );
        let main = proc(
            "Main",
            ProcedureKind::Sub,
            Vec::new(),
            Vec::new(),
            None,
            Vec::new(),
        );
        let prog = CoreProgram {
            long_ptr_width: Default::default(),
            procs: vec![main, twice],
            entry: Some(ProcId(0)),
            unit_name: "T".into(),
            classes: vec![CoreClass {
                name: "Widget".into(),
                runtime_type_name: None,
                predeclared: false,
                initialize: None,
                terminate: None,
                fields: Vec::new(),
                methods: vec![CoreClassMethod {
                    externally_visible: true,
                    name: "Twice".into(),
                    kind: ProjectMemberKind::Method,
                    proc: ProcId(1),
                    dispid: None,
                    vtable_slot: None,
                    is_default_member: false,
                    is_enumerator_member: false,
                }],
                events: Vec::new(),
                as_new_fields: Vec::new(),
                implements: Vec::new(),
            }],
            ..Default::default()
        };
        let oxp: &'static OxProgram = Box::leak(Box::new(
            calcvba_oxir::elaborate::elaborate(&prog).expect("elaborate"),
        ));
        let host: &'static NullHostServices =
            Box::leak(Box::new(NullHostServices::new(HostPolicy::default())));
        let vm = Vm3::activate(oxp, host).expect("activate");
        let args = vm
            .project_call_args(
                0,
                FuncId(1),
                &[OxCallArg::Named {
                    name: "x".into(),
                    value: OxOperand::Const(OxConst::I32(21)),
                    by_ref: None,
                }],
            )
            .expect("named arg should map to visible parameter after Me");
        assert_eq!(args.len(), 1);
        assert!(matches!(
            &args[0],
            OxArg::ByVal(OxOperand::Const(OxConst::I32(21)))
        ));
    }

    /// `CorePlace::Local(i)`.
    fn lc(i: usize) -> CorePlace {
        CorePlace::Local(CoreLocalId(i))
    }
    /// `CoreValue::Load(Local(i))`.
    fn load(i: usize) -> CoreValue {
        CoreValue::Load(lc(i))
    }
    /// A `Checked(Long)` addition (the regime the binder picks for `Long` operands).
    fn long_add(l: CoreValue, r: CoreValue) -> CoreValue {
        CoreValue::Binary {
            op: CoreBinOp::Add,
            lhs: Box::new(l),
            rhs: Box::new(r),
            mode: StringCompareMode::Binary,
            num: NumericMode::Checked(NumericCoerceTarget::Long),
        }
    }

    #[test]
    fn checked_long_arithmetic_matches() {
        // Sub Main(): n = (10 + 5) * 2  →  n is Long 30.
        let n = || CorePlace::Local(CoreLocalId(0));
        let long = NumericMode::Checked(NumericCoerceTarget::Long);
        let bin = |op, l, r| CoreValue::Binary {
            op,
            lhs: Box::new(l),
            rhs: Box::new(r),
            mode: StringCompareMode::Binary,
            num: long,
        };
        let expr = bin(
            CoreBinOp::Mul,
            bin(
                CoreBinOp::Add,
                CoreValue::Const(CoreConst::I32(10)),
                CoreValue::Const(CoreConst::I32(5)),
            ),
            CoreValue::Const(CoreConst::I32(2)),
        );
        let prog = main_proc(
            vec![local("n", VarTypeRef::Builtin(BuiltinType::Long))],
            vec![assign(n(), expr)],
        );
        let vm = run_core(&prog);
        let v = vm.slot(0).expect("slot n"); // global_count is 0, so slot 0 = Main local 0
        assert_eq!(v.vtype(), VarType::Long);
        assert_eq!(v.as_i32(), Some(30));
    }

    #[test]
    fn string_concat_matches() {
        // Sub Main(): s = "ab" & "cd"  →  s is "abcd".
        let s = || CorePlace::Local(CoreLocalId(0));
        let expr = CoreValue::Binary {
            op: CoreBinOp::Concat,
            lhs: Box::new(CoreValue::Const(CoreConst::Str("ab".into()))),
            rhs: Box::new(CoreValue::Const(CoreConst::Str("cd".into()))),
            mode: StringCompareMode::Binary,
            num: NumericMode::Widening,
        };
        let prog = main_proc(
            vec![local("s", VarTypeRef::Builtin(BuiltinType::String))],
            vec![assign(s(), expr)],
        );
        let vm = run_core(&prog);
        let v = vm.slot(0).expect("slot s");
        let s = calcvba_runtime::variant_to_vba_string(&v)
            .map(|b| b.as_str())
            .unwrap_or_default();
        assert_eq!(s, "abcd");
    }

    #[test]
    fn comparison_into_boolean_matches() {
        // Sub Main(): x = 3.5 : b = (x > 1)  →  b is True.
        let x = || CorePlace::Local(CoreLocalId(0));
        let b = || CorePlace::Local(CoreLocalId(1));
        let cmp = CoreValue::Binary {
            op: CoreBinOp::Gt,
            lhs: Box::new(CoreValue::Load(x())),
            rhs: Box::new(CoreValue::Const(CoreConst::I32(1))),
            mode: StringCompareMode::Binary,
            num: NumericMode::Widening,
        };
        let prog = main_proc(
            vec![
                local("x", VarTypeRef::Variant),
                local("b", VarTypeRef::Builtin(BuiltinType::Boolean)),
            ],
            vec![
                assign(x(), CoreValue::Const(CoreConst::F64(3.5f64.to_bits()))),
                assign(b(), cmp),
            ],
        );
        let vm = run_core(&prog);
        let bv = vm.slot(1).expect("slot b");
        assert_eq!(bv.as_bool(), Some(true));
    }

    #[test]
    fn module_globals_lead_the_snapshot() {
        // A module global `g As Long` then `Sub Main(): g = 7`.
        let g = CorePlace::Global(calcvba_bundle::coreir::GlobalId(0));
        let mut prog = main_proc(
            Vec::new(),
            vec![assign(g, CoreValue::Const(CoreConst::I32(7)))],
        );
        prog.globals = vec![CoreGlobal {
            name: "g".into(),
            ty: VarTypeRef::Builtin(BuiltinType::Long),
            array_element: None,
        }];
        let vm = run_core(&prog);
        // Slot 0 is the global (globals lead), and there are no Main locals after it.
        assert_eq!(vm.slot(0).and_then(|v| v.as_i32()), Some(7));
    }

    #[test]
    fn if_else_control_flow_executes() {
        // `If <c> Then n = 5 Else n = 9` — a non-Boolean condition flows through the
        // elaboration's Truthy coercion and the Branch terminator.
        use calcvba_bundle::coreir::CoreIfArm;
        let run_if = |cond: i32| -> Option<i32> {
            let n = || CorePlace::Local(CoreLocalId(0));
            let prog = main_proc(
                vec![local("n", VarTypeRef::Builtin(BuiltinType::Long))],
                vec![CoreStmt::If {
                    arms: vec![CoreIfArm {
                        condition: CoreValue::Const(CoreConst::I32(cond)),
                        body: vec![assign(n(), CoreValue::Const(CoreConst::I32(5)))],
                    }],
                    else_body: vec![assign(n(), CoreValue::Const(CoreConst::I32(9)))],
                }],
            );
            run_core(&prog).slot(0).and_then(|v| v.as_i32())
        };
        assert_eq!(
            run_if(1),
            Some(5),
            "a truthy condition takes the Then branch"
        );
        assert_eq!(
            run_if(0),
            Some(9),
            "a falsy condition takes the Else branch"
        );
    }

    #[test]
    fn entry_falls_back_to_the_first_proc_when_no_main() {
        // No `Sub Main` (CoreProgram.entry == None): vm3 must still run the only proc,
        // matching vm2's select_entry fallback (else nothing runs and `g` stays Empty).
        let g = CorePlace::Global(calcvba_bundle::coreir::GlobalId(0));
        let prog = CoreProgram {
            long_ptr_width: Default::default(),
            procs: vec![CoreProc {
                name: "Helper".into(), // deliberately not "Main"
                kind: ProcedureKind::Sub,
                params: Vec::new(),
                locals: Vec::new(),
                return_local: None,
                label_lines: Vec::new(),
                body: vec![assign(g, CoreValue::Const(CoreConst::I32(42)))],
            }],
            globals: vec![CoreGlobal {
                name: "g".into(),
                ty: VarTypeRef::Builtin(BuiltinType::Long),
                array_element: None,
            }],
            entry: None,
            unit_name: "T".into(),
            ..Default::default()
        };
        let vm = run_core(&prog);
        assert_eq!(vm.slot(0).and_then(|v| v.as_i32()), Some(42));
    }

    #[test]
    fn unassigned_boolean_condition_is_false_not_an_error() {
        // `Dim b As Boolean : If b Then n = 5 Else n = 9` — b is unassigned, so Empty at
        // runtime (not a Boolean tag). vm2's is_truthy(Empty) = False, so the Else branch
        // (n = 9); vm3 must match, not error on the non-Boolean tag. (Regression guard:
        // before the Truthy-always fix this returned Malformed.)
        use calcvba_bundle::coreir::CoreIfArm;
        let b = || CorePlace::Local(CoreLocalId(0)); // Dim b As Boolean (unassigned)
        let n = || CorePlace::Local(CoreLocalId(1)); // Dim n As Long
        let prog = main_proc(
            vec![
                local("b", VarTypeRef::Builtin(BuiltinType::Boolean)),
                local("n", VarTypeRef::Builtin(BuiltinType::Long)),
            ],
            vec![CoreStmt::If {
                arms: vec![CoreIfArm {
                    condition: CoreValue::Load(b()),
                    body: vec![assign(n(), CoreValue::Const(CoreConst::I32(5)))],
                }],
                else_body: vec![assign(n(), CoreValue::Const(CoreConst::I32(9)))],
            }],
        );
        let vm = run_core(&prog);
        assert_eq!(
            vm.slot(1).and_then(|v| v.as_i32()),
            Some(9),
            "unassigned Boolean is False -> Else branch, not an error"
        );
    }

    #[test]
    fn select_case_runs_including_a_null_selector() {
        // A Select matching a case, and one whose selector is Null (no case matches ->
        // Case Else) — the Null selector must fall through, not error (vm2 parity).
        use calcvba_bundle::coreir::{CaseClause, CoreCaseBlock};
        let run_select = |sel: CoreValue, sel_ty: VarTypeRef| -> Option<i32> {
            let s = || CorePlace::Local(CoreLocalId(0));
            let x = || CorePlace::Local(CoreLocalId(1));
            let prog = main_proc(
                vec![
                    local("s", sel_ty),
                    local("x", VarTypeRef::Builtin(BuiltinType::Long)),
                ],
                vec![
                    assign(s(), sel),
                    CoreStmt::Select {
                        selector: CoreValue::Load(s()),
                        cases: vec![CoreCaseBlock {
                            clauses: vec![CaseClause::Value(CoreValue::Const(CoreConst::I32(1)))],
                            body: vec![assign(x(), CoreValue::Const(CoreConst::I32(5)))],
                        }],
                        case_else: vec![assign(x(), CoreValue::Const(CoreConst::I32(9)))],
                        compare_mode: calcvba_bundle::StringCompareMode::Binary,
                    },
                ],
            );
            run_core(&prog).slot(1).and_then(|v| v.as_i32())
        };
        // Selector 1 matches `Case 1` -> x = 5.
        assert_eq!(
            run_select(
                CoreValue::Const(CoreConst::I32(1)),
                VarTypeRef::Builtin(BuiltinType::Long)
            ),
            Some(5)
        );
        // A Null selector matches nothing (is_truthy(Null) = False) -> Case Else, x = 9.
        assert_eq!(
            run_select(CoreValue::Const(CoreConst::Null), VarTypeRef::Variant),
            Some(9),
            "a Null Select selector falls through to Case Else, not an error"
        );
    }

    #[test]
    fn for_loop_accumulates() {
        // `For i = 1 To 3 : s = s + i : Next` -> s = 6 (exercises the For counter-test
        // Branch + its Truthy coercion end-to-end).
        let i = || CorePlace::Local(CoreLocalId(0));
        let s = || CorePlace::Local(CoreLocalId(1));
        let prog = main_proc(
            vec![
                local("i", VarTypeRef::Builtin(BuiltinType::Long)),
                local("s", VarTypeRef::Builtin(BuiltinType::Long)),
            ],
            vec![
                assign(s(), CoreValue::Const(CoreConst::I32(0))),
                CoreStmt::ForRange {
                    var: i(),
                    start: CoreValue::Const(CoreConst::I32(1)),
                    end: CoreValue::Const(CoreConst::I32(3)),
                    step: None,
                    body: vec![assign(
                        s(),
                        CoreValue::Binary {
                            op: CoreBinOp::Add,
                            lhs: Box::new(CoreValue::Load(s())),
                            rhs: Box::new(CoreValue::Load(i())),
                            mode: StringCompareMode::Binary,
                            num: NumericMode::Widening,
                        },
                    )],
                },
            ],
        );
        let vm = run_core(&prog);
        assert_eq!(vm.slot(1).and_then(|v| v.as_i32()), Some(6));
    }

    #[test]
    fn call_proc_returns_a_function_value() {
        // Function Add(a As Long, b As Long) As Long : Add = a + b
        // Sub Main() : n = Add(10, 5)   ->  n = 15
        let add = proc(
            "Add",
            ProcedureKind::Function,
            vec![long_param("a"), long_param("b")], // LocalId 0, 1
            vec![local("Add", VarTypeRef::Builtin(BuiltinType::Long))], // the return local, LocalId 2
            Some(CoreLocalId(2)),
            vec![assign(lc(2), long_add(load(0), load(1)))],
        );
        let main = proc(
            "Main",
            ProcedureKind::Sub,
            Vec::new(),
            vec![local("n", VarTypeRef::Builtin(BuiltinType::Long))], // LocalId 0
            None,
            vec![assign(
                lc(0),
                CoreValue::Call {
                    callee: CoreCallee::VbaProc { proc: ProcId(1) },
                    args: vec![
                        CoreArg::ByVal(CoreValue::Const(CoreConst::I32(10))),
                        CoreArg::ByVal(CoreValue::Const(CoreConst::I32(5))),
                    ],
                },
            )],
        );
        let prog = procs_program(vec![main, add]);
        let vm = run_core(&prog);
        assert_eq!(vm.slot(0).and_then(|v| v.as_i32()), Some(15));
    }

    #[test]
    fn link_over_a_single_program_runs_like_run() {
        // W2: Vm3::link over a single program activates the same image Vm3::run does; a
        // multi-program link now activates every program (cross-project EXECUTION — calls, fault
        // unwind, unresolved-reference rejection — is covered in tests/cross_program.rs).
        let prog = main_proc(
            vec![local("n", VarTypeRef::Builtin(BuiltinType::Long))],
            vec![assign(lc(0), CoreValue::Const(CoreConst::I32(42)))],
        );
        let oxp: &'static OxProgram = Box::leak(Box::new(
            calcvba_oxir::elaborate::elaborate(&prog).expect("elaborate"),
        ));
        let host: &'static NullHostServices =
            Box::leak(Box::new(NullHostServices::new(HostPolicy::default())));
        let mut vm = Vm3::link(&[oxp], host).expect("link single program");
        vm.run_entry().expect("run_entry");
        assert_eq!(vm.slot(0).and_then(|v| v.as_i32()), Some(42));

        // A multi-program link activates every program and runs the entry (the last program);
        // two independent copies link cleanly and the entry's `n` slot is 42.
        let mut multi = Vm3::link(&[oxp, oxp], host).expect("multi-program link");
        multi.run_entry().expect("run_entry");
        assert_eq!(multi.slot(0).and_then(|v| v.as_i32()), Some(42));
    }

    #[test]
    fn host_session_creates_a_class_and_invokes_a_member() {
        // W7: the host session API — create a project-class instance by name (activate without
        // running Main), then invoke a member with pre-marshaled value args. Class Widget with
        // Function Twice(x) = x + x; Widget.Twice(21) -> 42.
        let twice = proc(
            "Twice",
            ProcedureKind::Function,
            vec![long_param("me"), long_param("x")], // me LocalId 0, x LocalId 1
            vec![local("Twice", VarTypeRef::Builtin(BuiltinType::Long))], // return local, LocalId 2
            Some(CoreLocalId(2)),
            vec![assign(lc(2), long_add(load(1), load(1)))],
        );
        let main = proc(
            "Main",
            ProcedureKind::Sub,
            Vec::new(),
            Vec::new(),
            None,
            Vec::new(),
        );
        let prog = CoreProgram {
            long_ptr_width: Default::default(),
            procs: vec![main, twice],
            entry: Some(ProcId(0)),
            unit_name: "T".into(),
            classes: vec![CoreClass {
                name: "Widget".into(),
                runtime_type_name: None,
                predeclared: false,
                initialize: None,
                terminate: None,
                fields: Vec::new(),
                methods: vec![CoreClassMethod {
                    externally_visible: true,
                    name: "Twice".into(),
                    kind: ProjectMemberKind::Method,
                    proc: ProcId(1),
                    dispid: None,
                    vtable_slot: None,
                    is_default_member: false,
                    is_enumerator_member: false,
                }],
                events: Vec::new(),
                as_new_fields: Vec::new(),
                implements: Vec::new(),
            }],
            ..Default::default()
        };
        let oxp: &'static OxProgram = Box::leak(Box::new(
            calcvba_oxir::elaborate::elaborate(&prog).expect("elaborate"),
        ));
        let host: &'static NullHostServices =
            Box::leak(Box::new(NullHostServices::new(HostPolicy::default())));
        let mut vm = Vm3::activate(oxp, host).expect("activate");
        let widget = vm.create_project_instance("Widget").expect("create Widget");
        let obj = widget.as_object_ref().expect("Widget is an object");
        let r = vm
            .invoke_member_values(
                obj,
                "Twice",
                Some(ProjectMemberKind::Method),
                vec![Variant::from_i32(21)],
            )
            .expect("invoke Twice");
        assert_eq!(
            r.as_i32(),
            Some(42),
            "create + invoke a member with value args"
        );
        // An unknown class is "can't create object" (429).
        assert!(matches!(
            vm.create_project_instance("Nope"),
            Err(Vm3Error::Fault(_))
        ));
    }

    #[test]
    fn project_instance_runtime_descriptor_carries_dispatch_members() {
        let get_value = proc(
            "GetValue",
            ProcedureKind::PropertyGet,
            vec![long_param("me")],
            vec![local("GetValue", VarTypeRef::Builtin(BuiltinType::Long))],
            Some(CoreLocalId(1)),
            vec![assign(lc(1), CoreValue::Const(CoreConst::I32(7)))],
        );
        let let_value = proc(
            "LetValue",
            ProcedureKind::PropertyLet,
            vec![long_param("me"), byval_long_param("newValue")],
            Vec::new(),
            None,
            Vec::new(),
        );
        let reset = proc(
            "Reset",
            ProcedureKind::Sub,
            vec![long_param("me"), optional_long_param("count")],
            Vec::new(),
            None,
            Vec::new(),
        );
        let new_enum = proc(
            "NewEnum",
            ProcedureKind::PropertyGet,
            vec![long_param("me")],
            vec![local("NewEnum", VarTypeRef::Object("Object".into()))],
            Some(CoreLocalId(1)),
            Vec::new(),
        );
        let initialize = proc(
            "Class_Initialize",
            ProcedureKind::Sub,
            vec![long_param("me")],
            Vec::new(),
            None,
            Vec::new(),
        );
        let terminate = proc(
            "Class_Terminate",
            ProcedureKind::Sub,
            vec![long_param("me")],
            Vec::new(),
            None,
            Vec::new(),
        );
        let main = proc(
            "Main",
            ProcedureKind::Sub,
            Vec::new(),
            Vec::new(),
            None,
            Vec::new(),
        );
        let prog = CoreProgram {
            long_ptr_width: Default::default(),
            procs: vec![
                main, get_value, let_value, reset, initialize, terminate, new_enum,
            ],
            entry: Some(ProcId(0)),
            unit_name: "T".into(),
            classes: vec![CoreClass {
                name: "Widget".into(),
                runtime_type_name: None,
                predeclared: true,
                initialize: Some(ProcId(4)),
                terminate: Some(ProcId(5)),
                fields: vec![
                    CoreClassField {
                        externally_visible: false,
                        name: "mCount".into(),
                        token: 11,
                        ty: VarTypeRef::Builtin(BuiltinType::Long),
                        array_element: None,
                        dispid: None,
                        get_vtable_slot: None,
                        put_vtable_slot: None,
                        putref_vtable_slot: None,
                    },
                    CoreClassField {
                        externally_visible: false,
                        name: "Items".into(),
                        token: 12,
                        ty: VarTypeRef::Array(Box::new(VarTypeRef::Builtin(BuiltinType::Long))),
                        array_element: Some(ArrayElementType::Long),
                        dispid: None,
                        get_vtable_slot: None,
                        put_vtable_slot: None,
                        putref_vtable_slot: None,
                    },
                    CoreClassField {
                        externally_visible: true,
                        name: "Child".into(),
                        token: 13,
                        ty: VarTypeRef::Object("Widget".into()),
                        array_element: None,
                        dispid: Some(0x4003_0000),
                        get_vtable_slot: Some(9),
                        put_vtable_slot: Some(10),
                        putref_vtable_slot: Some(11),
                    },
                ],
                methods: vec![
                    CoreClassMethod {
                        externally_visible: true,
                        name: "Value".into(),
                        kind: ProjectMemberKind::PropertyGet,
                        proc: ProcId(1),
                        dispid: Some(0),
                        vtable_slot: Some(7),
                        is_default_member: true,
                        is_enumerator_member: false,
                    },
                    CoreClassMethod {
                        externally_visible: true,
                        name: "Value".into(),
                        kind: ProjectMemberKind::PropertyLet,
                        proc: ProcId(2),
                        dispid: Some(0),
                        vtable_slot: Some(8),
                        is_default_member: true,
                        is_enumerator_member: false,
                    },
                    CoreClassMethod {
                        externally_visible: true,
                        name: "Reset".into(),
                        kind: ProjectMemberKind::Method,
                        proc: ProcId(3),
                        dispid: None,
                        vtable_slot: None,
                        is_default_member: false,
                        is_enumerator_member: false,
                    },
                    CoreClassMethod {
                        externally_visible: true,
                        name: "NewEnum".into(),
                        kind: ProjectMemberKind::PropertyGet,
                        proc: ProcId(6),
                        dispid: Some(-4),
                        vtable_slot: Some(12),
                        is_default_member: false,
                        is_enumerator_member: true,
                    },
                ],
                events: Vec::new(),
                as_new_fields: vec![CoreClassAsNewField {
                    field: 13,
                    binding: CoreAsNew::ProjectClass {
                        class: CoreClassId(0),
                    },
                }],
                implements: vec!["IWidget".into(), "IDisposable".into()],
            }],
            ..Default::default()
        };
        let oxp: &'static OxProgram = Box::leak(Box::new(
            calcvba_oxir::elaborate::elaborate(&prog).expect("elaborate"),
        ));
        let host: &'static NullHostServices =
            Box::leak(Box::new(NullHostServices::new(HostPolicy::default())));
        let mut vm = Vm3::activate(oxp, host).expect("activate");
        let widget = vm.create_project_instance("Widget").expect("create Widget");
        let obj = widget.as_object_ref().expect("Widget object");
        let class_descriptor = obj.class_descriptor();

        assert_eq!(class_descriptor.name, "Widget");
        assert_eq!(
            class_descriptor.project_identity,
            Some(RuntimeProjectClassIdentity {
                unit_name: "T",
                class_index: 0,
            })
        );
        assert!(class_descriptor.predeclared);
        assert!(class_descriptor.lifecycle.has_initialize);
        assert!(class_descriptor.lifecycle.has_terminate);
        assert_eq!(class_descriptor.implements, &["IWidget", "IDisposable"]);
        assert_eq!(class_descriptor.fields.len(), 3);
        assert_eq!(class_descriptor.fields[0].name, "mCount");
        assert_eq!(class_descriptor.fields[0].token, 11);
        assert_eq!(
            class_descriptor.fields[0].value_type,
            RuntimeValueType::Long
        );
        assert_eq!(class_descriptor.fields[0].array_element_type, None);
        assert_eq!(class_descriptor.fields[1].name, "Items");
        assert_eq!(class_descriptor.fields[1].token, 12);
        assert_eq!(
            class_descriptor.fields[1].value_type,
            RuntimeValueType::Array
        );
        assert_eq!(
            class_descriptor.fields[1].array_element_type,
            Some(RuntimeValueType::Long)
        );
        assert_eq!(class_descriptor.fields[2].name, "Child");
        assert_eq!(class_descriptor.fields[2].token, 13);
        assert_eq!(
            class_descriptor.fields[2].value_type,
            RuntimeValueType::Object
        );
        assert_eq!(class_descriptor.fields[2].array_element_type, None);
        assert_eq!(class_descriptor.as_new_fields.len(), 1);
        assert_eq!(class_descriptor.as_new_fields[0].field_token, 13);
        assert_eq!(
            class_descriptor.as_new_fields[0].activation,
            RuntimeClassActivationDescriptor::ProjectClass { class_index: 0 }
        );

        assert!(
            obj.query_interface_descriptor(RuntimeInterfaceId::IUnknown)
                .is_some()
        );
        assert_eq!(class_descriptor.validate_com_export_shape(), Ok(()));
        assert_eq!(
            class_descriptor.default_interface_iid,
            Some(calcvba_runtime::Guid::deterministic("T", "IWidget"))
        );
        let dispatch = obj
            .query_interface_descriptor(RuntimeInterfaceId::IDispatch)
            .expect("project instances should advertise IDispatch");
        assert_eq!(dispatch.name, "IDispatch");
        assert!(dispatch.dual_dispatch);
        // Sorted by vtable slot: Value Get (7), Value Let (8), Child Get (9), Child Let
        // [restricted] (10), Child Set (11), NewEnum (12), Reset (13 — filled by the shared
        // descriptor builder).
        assert_eq!(dispatch.members.len(), 7);

        let get = &dispatch.members[0];
        assert_eq!(get.name, "Value");
        assert_eq!(get.dispatch_id, 0);
        assert_eq!(get.vtable_slot, Some(7));
        assert_eq!(get.invoke_kind, RuntimeMemberInvokeKind::PropertyGet);
        assert!(get.is_default_member);
        assert!(!get.is_enumerator_member);
        assert_eq!(get.arity, 0);
        assert_eq!(get.params, &[]);
        assert_eq!(get.return_type, Some(RuntimeValueType::Long));

        let put = &dispatch.members[1];
        assert_eq!(put.name, "Value");
        assert_eq!(put.dispatch_id, 0);
        assert_eq!(put.vtable_slot, Some(8));
        assert_eq!(put.invoke_kind, RuntimeMemberInvokeKind::PropertyLet);
        assert!(put.is_default_member);
        assert!(!put.is_enumerator_member);
        assert_eq!(put.arity, 1);
        assert_eq!(put.params[0].name, "newValue");
        assert_eq!(put.params[0].value_type, RuntimeValueType::Long);
        assert!(!put.params[0].by_ref);
        assert_eq!(put.return_type, None);

        let child_get = &dispatch.members[2];
        assert_eq!(child_get.name, "Child");
        assert_eq!(child_get.dispatch_id, 0x4003_0000);
        assert_eq!(child_get.vtable_slot, Some(9));
        assert_eq!(child_get.invoke_kind, RuntimeMemberInvokeKind::PropertyGet);
        assert_eq!(child_get.backing, RuntimeMemberBacking::Field(13));

        let child_let = &dispatch.members[3];
        assert_eq!(child_let.name, "Child");
        assert_eq!(child_let.dispatch_id, 0x4003_0000);
        assert_eq!(child_let.vtable_slot, Some(10));
        assert_eq!(child_let.invoke_kind, RuntimeMemberInvokeKind::PropertyLet);
        assert_eq!(
            child_let.func_flags,
            calcvba_runtime::com::FUNCFLAG_FRESTRICTED
        );
        assert_eq!(child_let.backing, RuntimeMemberBacking::Field(13));

        let child_set = &dispatch.members[4];
        assert_eq!(child_set.name, "Child");
        assert_eq!(child_set.dispatch_id, 0x4003_0000);
        assert_eq!(child_set.vtable_slot, Some(11));
        assert_eq!(child_set.invoke_kind, RuntimeMemberInvokeKind::PropertySet);
        assert_eq!(child_set.func_flags, 0);
        assert_eq!(child_set.backing, RuntimeMemberBacking::Field(13));

        let new_enum = &dispatch.members[5];
        assert_eq!(new_enum.name, "NewEnum");
        assert_eq!(new_enum.dispatch_id, -4);
        assert_eq!(new_enum.vtable_slot, Some(12));
        assert_eq!(new_enum.invoke_kind, RuntimeMemberInvokeKind::PropertyGet);
        assert!(!new_enum.is_default_member);
        assert!(new_enum.is_enumerator_member);
        assert_eq!(new_enum.arity, 0);
        assert_eq!(new_enum.return_type, Some(RuntimeValueType::Object));

        let method = &dispatch.members[6];
        assert_eq!(method.name, "Reset");
        assert_eq!(method.dispatch_id, 0x6003_0000);
        assert_eq!(method.vtable_slot, Some(13));
        assert_eq!(method.invoke_kind, RuntimeMemberInvokeKind::Method);
        assert!(!method.is_default_member);
        assert!(!method.is_enumerator_member);
        assert_eq!(method.arity, 1);
        assert_eq!(method.params[0].name, "count");
        assert_eq!(method.params[0].value_type, RuntimeValueType::Long);
        assert!(method.params[0].by_ref);
        assert!(method.params[0].optional);
        assert!(!method.params[0].param_array);
        assert_eq!(method.return_type, None);

        drop(obj);
        drop(widget);
        vm.maybe_drain();
    }

    #[test]
    fn project_instance_runtime_descriptor_carries_implemented_interfaces() {
        let iface_get_size = proc(
            "Size",
            ProcedureKind::PropertyGet,
            vec![long_param("me")],
            vec![local("Size", VarTypeRef::Builtin(BuiltinType::Long))],
            Some(CoreLocalId(1)),
            Vec::new(),
        );
        let iface_let_size = proc(
            "Size",
            ProcedureKind::PropertyLet,
            vec![long_param("me"), byval_long_param("value")],
            Vec::new(),
            None,
            Vec::new(),
        );
        let impl_get_size = proc(
            "IShape_Size",
            ProcedureKind::PropertyGet,
            vec![long_param("me")],
            vec![local("IShape_Size", VarTypeRef::Builtin(BuiltinType::Long))],
            Some(CoreLocalId(1)),
            Vec::new(),
        );
        let impl_let_size = proc(
            "IShape_Size",
            ProcedureKind::PropertyLet,
            vec![long_param("me"), byval_long_param("newSize")],
            Vec::new(),
            None,
            Vec::new(),
        );
        let main = proc(
            "Main",
            ProcedureKind::Sub,
            Vec::new(),
            Vec::new(),
            None,
            Vec::new(),
        );
        let prog = CoreProgram {
            long_ptr_width: Default::default(),
            procs: vec![
                main,
                iface_get_size,
                iface_let_size,
                impl_get_size,
                impl_let_size,
            ],
            entry: Some(ProcId(0)),
            unit_name: "T".into(),
            classes: vec![
                CoreClass {
                    name: "CBox".into(),
                    runtime_type_name: None,
                    predeclared: false,
                    initialize: None,
                    terminate: None,
                    fields: Vec::new(),
                    methods: vec![
                        CoreClassMethod {
                            externally_visible: true,
                            name: "IShape_Size".into(),
                            kind: ProjectMemberKind::PropertyGet,
                            proc: ProcId(3),
                            dispid: None,
                            vtable_slot: None,
                            is_default_member: false,
                            is_enumerator_member: false,
                        },
                        CoreClassMethod {
                            externally_visible: true,
                            name: "IShape_Size".into(),
                            kind: ProjectMemberKind::PropertyLet,
                            proc: ProcId(4),
                            dispid: None,
                            vtable_slot: None,
                            is_default_member: false,
                            is_enumerator_member: false,
                        },
                    ],
                    events: Vec::new(),
                    as_new_fields: Vec::new(),
                    implements: vec!["IShape".into()],
                },
                CoreClass {
                    name: "IShape".into(),
                    runtime_type_name: None,
                    predeclared: false,
                    initialize: None,
                    terminate: None,
                    fields: Vec::new(),
                    methods: vec![
                        CoreClassMethod {
                            externally_visible: true,
                            name: "Size".into(),
                            kind: ProjectMemberKind::PropertyGet,
                            proc: ProcId(1),
                            dispid: Some(0),
                            vtable_slot: Some(11),
                            is_default_member: true,
                            is_enumerator_member: false,
                        },
                        CoreClassMethod {
                            externally_visible: true,
                            name: "Size".into(),
                            kind: ProjectMemberKind::PropertyLet,
                            proc: ProcId(2),
                            dispid: Some(0),
                            vtable_slot: Some(12),
                            is_default_member: true,
                            is_enumerator_member: false,
                        },
                    ],
                    events: Vec::new(),
                    as_new_fields: Vec::new(),
                    implements: Vec::new(),
                },
            ],
            ..Default::default()
        };
        let oxp: &'static OxProgram = Box::leak(Box::new(
            calcvba_oxir::elaborate::elaborate(&prog).expect("elaborate"),
        ));
        let host: &'static NullHostServices =
            Box::leak(Box::new(NullHostServices::new(HostPolicy::default())));
        let mut vm = Vm3::activate(oxp, host).expect("activate");
        let cbox = vm.create_project_instance("CBox").expect("create CBox");
        let obj = cbox.as_object_ref().expect("CBox object");
        let class_descriptor = obj.class_descriptor();

        assert_eq!(class_descriptor.implements, &["IShape"]);
        assert_eq!(
            class_descriptor.project_identity,
            Some(RuntimeProjectClassIdentity {
                unit_name: "T",
                class_index: 0,
            })
        );
        assert_eq!(class_descriptor.interfaces.len(), 3);
        assert!(
            obj.query_interface_descriptor(RuntimeInterfaceId::IUnknown)
                .is_some()
        );
        assert!(
            obj.query_interface_descriptor(RuntimeInterfaceId::IDispatch)
                .is_some()
        );

        let shape = class_descriptor
            .interfaces
            .iter()
            .find(|interface| interface.name == "IShape")
            .expect("implemented project interface descriptor");
        assert_eq!(shape.id, RuntimeInterfaceId::Unsupported);
        assert_eq!(shape.identity.id, RuntimeInterfaceId::Unsupported);
        assert_eq!(shape.identity.name, "T.IShape");
        assert_eq!(shape.identity.kind, RuntimeInterfaceKind::Custom);
        assert!(!shape.dual_dispatch);
        assert_eq!(shape.members.len(), 2);
        assert_eq!(
            obj.query_interface_descriptor_by_guid(shape.identity.guid)
                .map(|descriptor| descriptor.name),
            Some("IShape")
        );

        let get = &shape.members[0];
        assert_eq!(get.name, "Size");
        assert_eq!(get.dispatch_id, 0);
        assert_eq!(get.vtable_slot, Some(11));
        assert_eq!(get.invoke_kind, RuntimeMemberInvokeKind::PropertyGet);
        assert!(get.is_default_member);
        assert_eq!(get.arity, 0);
        assert_eq!(get.return_type, Some(RuntimeValueType::Long));

        let put = &shape.members[1];
        assert_eq!(put.name, "Size");
        assert_eq!(put.dispatch_id, 0);
        assert_eq!(put.vtable_slot, Some(12));
        assert_eq!(put.invoke_kind, RuntimeMemberInvokeKind::PropertyLet);
        assert!(put.is_default_member);
        assert_eq!(put.arity, 1);
        assert_eq!(put.params[0].name, "newSize");
        assert_eq!(put.params[0].value_type, RuntimeValueType::Long);
        assert!(!put.params[0].by_ref);
        assert_eq!(put.return_type, None);

        drop(obj);
        drop(cbox);
        vm.maybe_drain();
    }

    #[test]
    fn run_proc_with_values_seeds_host_value_args_across_repeated_invokes() {
        // W0: a host session activates a program ONCE (no Main run), then invokes a function
        // directly with already-evaluated Variant args (the path object/array args need, since
        // an OxArg::ByVal carries only a scalar OxConst). Invoking twice against the one
        // activation proves the session lifecycle: activate() is separate from run_entry(), so
        // the descriptor-leak + termination-reset happen once, not per invoke.
        // run_proc_with_values mirrors run_proc_with_me's frame layout: local 0 is the
        // receiver `me` (as for a class method / event handler), and the value args land at
        // locals 1.. — so the callable is shaped like a method whose first slot is `me`.
        let add = proc(
            "Add",
            ProcedureKind::Function,
            vec![long_param("me"), long_param("a"), long_param("b")], // LocalId 0=me, 1=a, 2=b
            vec![local("Add", VarTypeRef::Builtin(BuiltinType::Long))], // return local, LocalId 3
            Some(CoreLocalId(3)),
            vec![assign(lc(3), long_add(load(1), load(2)))],
        );
        // Main is a no-op entry (procs_program requires procs[0] to be the entry); we never run it.
        let main = proc(
            "Main",
            ProcedureKind::Sub,
            Vec::new(),
            Vec::new(),
            None,
            Vec::new(),
        );
        let prog = procs_program(vec![main, add]);
        let oxp: &'static OxProgram = Box::leak(Box::new(
            calcvba_oxir::elaborate::elaborate(&prog).expect("elaborate"),
        ));
        let host: &'static NullHostServices =
            Box::leak(Box::new(NullHostServices::new(HostPolicy::default())));
        let mut vm = Vm3::activate(oxp, host).expect("activate"); // NOT run() — Main never executes

        let first = vm
            .run_proc_with_values(
                0,
                FuncId(1),
                Variant::empty(),
                vec![Variant::from_i32(20), Variant::from_i32(22)],
                false,
            )
            .expect("invoke Add #1");
        assert_eq!(
            first.as_i32(),
            Some(42),
            "value-seeded args reach the proc and return their sum"
        );

        let second = vm
            .run_proc_with_values(
                0,
                FuncId(1),
                Variant::empty(),
                vec![Variant::from_i32(1), Variant::from_i32(2)],
                false,
            )
            .expect("invoke Add #2");
        assert_eq!(
            second.as_i32(),
            Some(3),
            "a second invoke against the same activation works"
        );
    }

    #[test]
    fn call_proc_byref_mutates_the_caller() {
        // Sub Bump(ByRef x As Long) : x = x + 1
        // Sub Main() : v = 41 : Bump(v)   ->  v = 42 (true aliasing through the frame stack)
        let bump = proc(
            "Bump",
            ProcedureKind::Sub,
            vec![long_param("x")], // LocalId 0
            Vec::new(),
            None,
            vec![assign(
                lc(0),
                long_add(load(0), CoreValue::Const(CoreConst::I32(1))),
            )],
        );
        let main = proc(
            "Main",
            ProcedureKind::Sub,
            Vec::new(),
            vec![local("v", VarTypeRef::Builtin(BuiltinType::Long))], // LocalId 0
            None,
            vec![
                assign(lc(0), CoreValue::Const(CoreConst::I32(41))),
                CoreStmt::Eval(CoreValue::Call {
                    callee: CoreCallee::VbaProc { proc: ProcId(1) },
                    args: vec![CoreArg::ByRef(lc(0))],
                }),
            ],
        );
        let prog = procs_program(vec![main, bump]);
        let vm = run_core(&prog);
        assert_eq!(
            vm.slot(0).and_then(|v| v.as_i32()),
            Some(42),
            "a ByRef write must propagate to the caller's backing slot"
        );
    }

    #[test]
    fn call_proc_byval_does_not_mutate_the_caller() {
        // The same Bump, but Main passes `v` ByVal -> the callee mutates a copy, v stays 41.
        let bump = proc(
            "Bump",
            ProcedureKind::Sub,
            vec![long_param("x")],
            Vec::new(),
            None,
            vec![assign(
                lc(0),
                long_add(load(0), CoreValue::Const(CoreConst::I32(1))),
            )],
        );
        let main = proc(
            "Main",
            ProcedureKind::Sub,
            Vec::new(),
            vec![local("v", VarTypeRef::Builtin(BuiltinType::Long))],
            None,
            vec![
                assign(lc(0), CoreValue::Const(CoreConst::I32(41))),
                CoreStmt::Eval(CoreValue::Call {
                    callee: CoreCallee::VbaProc { proc: ProcId(1) },
                    args: vec![CoreArg::ByVal(load(0))],
                }),
            ],
        );
        let prog = procs_program(vec![main, bump]);
        let vm = run_core(&prog);
        assert_eq!(
            vm.slot(0).and_then(|v| v.as_i32()),
            Some(41),
            "a ByVal copy must NOT propagate back to the caller"
        );
    }

    #[test]
    fn call_native_builtin_invokes_the_shared_library() {
        // Sub Main() : n = Len("abc")   ->  n = 3 (through the context-free dispatcher).
        let main = proc(
            "Main",
            ProcedureKind::Sub,
            Vec::new(),
            vec![local("n", VarTypeRef::Builtin(BuiltinType::Long))],
            None,
            vec![assign(
                lc(0),
                CoreValue::Call {
                    callee: CoreCallee::Native(NativeImplId::Len),
                    args: vec![CoreArg::ByVal(CoreValue::Const(CoreConst::Str(
                        "abc".into(),
                    )))],
                },
            )],
        );
        let prog = procs_program(vec![main]);
        let vm = run_core(&prog);
        assert_eq!(vm.slot(0).and_then(|v| v.as_i32()), Some(3));
    }

    /// A bounded host probe for the current VM3 event-pump contract. The host
    /// receives no VM/session callback; `DoEvents` returns first, and VM3 polls
    /// queued COM callbacks only from its following `StmtBoundary`.
    struct QueuedDoEventsHost {
        inner: NullHostServices,
        in_do_events: AtomicBool,
        callback_ready: AtomicBool,
        order: Mutex<Vec<&'static str>>,
    }

    impl QueuedDoEventsHost {
        fn new() -> Self {
            Self {
                inner: NullHostServices::new(HostPolicy::default()),
                in_do_events: AtomicBool::new(false),
                callback_ready: AtomicBool::new(false),
                order: Mutex::new(Vec::new()),
            }
        }

        fn push(&self, event: &'static str) {
            self.order.lock().expect("order lock").push(event);
        }
    }

    impl HostServices for QueuedDoEventsHost {
        fn profile(&self) -> HalProfileId {
            self.inner.profile()
        }
        fn descriptor(&self) -> HalDescriptor {
            self.inner.descriptor()
        }
        fn policy(&self) -> &HostPolicy {
            self.inner.policy()
        }
        fn console(&self) -> &dyn ConsoleHal {
            self.inner.console()
        }
        fn ui(&self) -> &dyn UiInteractionHal {
            self.inner.ui()
        }
        fn events(&self) -> &dyn EventPumpHal {
            self
        }
        fn fs(&self) -> &dyn FileSystemHal {
            self.inner.fs()
        }
        fn process(&self) -> &dyn ProcessEnvHal {
            self.inner.process()
        }
        fn com(&self) -> &dyn ComHal {
            self
        }
        fn time_locale(&self) -> &dyn TimeLocaleHal {
            self.inner.time_locale()
        }
        fn dynlink(&self) -> &dyn DynamicLinkHal {
            self.inner.dynlink()
        }
        fn diag(&self) -> &dyn DiagnosticsHal {
            self.inner.diag()
        }
    }

    impl EventPumpHal for QueuedDoEventsHost {
        fn do_events_variant(&self) -> HalResult<Variant> {
            assert!(
                !self.in_do_events.swap(true, Ordering::AcqRel),
                "DoEvents must not recursively enter this host probe"
            );
            self.push("do-events-enter");
            // StandardHostServices performs a bounded OS pump and marks a queued
            // COM callback here. It has no safe VM3 callback authority.
            self.callback_ready.store(true, Ordering::Release);
            self.push("do-events-exit");
            self.in_do_events.store(false, Ordering::Release);
            Ok(Variant::from_i32(0))
        }
    }

    impl ComHal for QueuedDoEventsHost {
        fn describe_object(
            &self,
            object: ObjectRef,
        ) -> HalResult<Option<calcvba_com::ComObjectDescriptor>> {
            self.inner.com().describe_object(object)
        }

        fn subscribe_event(
            &self,
            object: ObjectRef,
            event: ComMemberToken,
        ) -> HalResult<ComSubscriptionToken> {
            self.inner.com().subscribe_event(object, event)
        }

        fn poll_event_callback(&self) -> HalResult<Option<ComCallbackPayload>> {
            assert!(
                !self.in_do_events.load(Ordering::Acquire),
                "VM3 must not poll or dispatch a callback inside the host DoEvents call"
            );
            self.push("stmt-boundary-poll");
            if self.callback_ready.swap(false, Ordering::AcqRel) {
                return Ok(Some(ComCallbackPayload {
                    callback: ComCallbackToken::new(60_001),
                    subscription: ComSubscriptionToken::new(40_001),
                    object: ObjectRef::from_compat_identity(20_001),
                    event: ComMemberToken::new(1),
                    args: Vec::new(),
                }));
            }
            Ok(None)
        }

        fn event_callback_subscription(
            &self,
            callback: ComCallbackToken,
        ) -> HalResult<ComSubscriptionToken> {
            self.inner.com().event_callback_subscription(callback)
        }

        fn event_callback_arity(&self, callback: ComCallbackToken) -> HalResult<usize> {
            self.inner.com().event_callback_arity(callback)
        }

        fn release_event_callback_variant(&self, callback: ComCallbackToken) -> HalResult<Variant> {
            assert_eq!(callback, ComCallbackToken::new(60_001));
            self.push("callback-release");
            Ok(Variant::from_i32(1))
        }

        fn resolve_typelib_reference(
            &self,
            request: &TypeLibResolveRequest,
        ) -> HalResult<TypeLibResolvedIdentity> {
            self.inner.com().resolve_typelib_reference(request)
        }

        fn load_typelib_metadata(
            &self,
            identity: &TypeLibResolvedIdentity,
        ) -> HalResult<TypeLibMetadataBlob> {
            self.inner.com().load_typelib_metadata(identity)
        }

        fn invalidate_typelib_cache(
            &self,
            scope: TypeLibCacheScope,
            reference_name: Option<&str>,
        ) -> HalResult<Variant> {
            self.inner
                .com()
                .invalidate_typelib_cache(scope, reference_name)
        }
    }

    #[test]
    fn do_events_returns_before_vm3_polls_callbacks_and_rng_context_remains_usable() {
        let main = proc(
            "Main",
            ProcedureKind::Sub,
            Vec::new(),
            vec![
                local("before", VarTypeRef::Builtin(BuiltinType::Single)),
                local("after", VarTypeRef::Builtin(BuiltinType::Single)),
            ],
            None,
            vec![
                assign(
                    lc(0),
                    CoreValue::Call {
                        callee: CoreCallee::Native(NativeImplId::Rnd),
                        args: Vec::new(),
                    },
                ),
                CoreStmt::Eval(CoreValue::Call {
                    callee: CoreCallee::Native(NativeImplId::DoEvents),
                    args: Vec::new(),
                }),
                assign(
                    lc(1),
                    CoreValue::Call {
                        callee: CoreCallee::Native(NativeImplId::Rnd),
                        args: Vec::new(),
                    },
                ),
            ],
        );
        let oxp = calcvba_oxir::elaborate::elaborate(&procs_program(vec![main]))
            .expect("elaborate DoEvents probe");
        let host = QueuedDoEventsHost::new();
        let vm = Vm3::run(&oxp, &host).expect("VM3 DoEvents probe");

        assert!(vm.slot(0).and_then(|value| value.as_f32()).is_some());
        assert!(vm.slot(1).and_then(|value| value.as_f32()).is_some());

        let order = host.order.lock().expect("order lock");
        let enter = order
            .iter()
            .position(|event| *event == "do-events-enter")
            .expect("DoEvents entered host");
        assert_eq!(order.get(enter + 1), Some(&"do-events-exit"));
        // COM events run synchronously inside the host call that delivers them (design doc,
        // step 8c), so VM3 never polls a queue afterwards.
        assert!(
            !order[enter + 2..].contains(&"stmt-boundary-poll"),
            "VM3 must not poll queued callbacks: {order:?}"
        );
    }

    // ── Native interop (M3-7): Declare Lib + pointer helpers + AddressOf ──────────
    //
    // The corpus differential gate exercises the `Declare`/`StrPtr`/`VarPtr` *reachable*
    // ops end-to-end against the deterministic host (vm3 matches vm2). These unit tests
    // pin the marshalling under a *controlled* mock dynlink — proving the ByRef write-back
    // targets the caller slot, `Err.LastDllError` refreshes, and a pointer pin round-trips
    // — independent of whatever a real OS `Declare` would do.

    /// A host whose dynlink echoes each argument incremented by 100 (a fake native
    /// "Bump") and reports a fixed `last_dll_error`, to exercise per-call-site `Declare`
    /// ByRef write-back and the `Err.LastDllError` refresh. Mirrors vm2's test host.
    struct MockDynlinkHost {
        inner: NullHostServices,
        last_dll_error: i32,
    }
    impl MockDynlinkHost {
        fn new(last_dll_error: i32) -> Self {
            Self {
                inner: NullHostServices::new(HostPolicy::default()),
                last_dll_error,
            }
        }
    }
    impl HostServices for MockDynlinkHost {
        fn profile(&self) -> HalProfileId {
            self.inner.profile()
        }
        fn descriptor(&self) -> HalDescriptor {
            self.inner.descriptor()
        }
        fn policy(&self) -> &HostPolicy {
            self.inner.policy()
        }
        fn console(&self) -> &dyn ConsoleHal {
            self.inner.console()
        }
        fn ui(&self) -> &dyn UiInteractionHal {
            self.inner.ui()
        }
        fn events(&self) -> &dyn EventPumpHal {
            self.inner.events()
        }
        fn fs(&self) -> &dyn FileSystemHal {
            self.inner.fs()
        }
        fn process(&self) -> &dyn ProcessEnvHal {
            self.inner.process()
        }
        fn com(&self) -> &dyn ComHal {
            self.inner.com()
        }
        fn time_locale(&self) -> &dyn TimeLocaleHal {
            self.inner.time_locale()
        }
        fn dynlink(&self) -> &dyn DynamicLinkHal {
            self
        }
        fn diag(&self) -> &dyn DiagnosticsHal {
            self.inner.diag()
        }
    }
    impl DynamicLinkHal for MockDynlinkHost {
        fn invoke_descriptor_variants(
            &self,
            _descriptor: &DynLinkDescriptorView<'_>,
            args: &[Variant],
        ) -> HalResult<(Variant, Vec<Variant>)> {
            let wb = args
                .iter()
                .map(|a| {
                    a.as_i32()
                        .map(|n| Variant::from_i32(n + 100))
                        .unwrap_or_else(|| a.clone())
                })
                .collect();
            Ok((Variant::empty(), wb))
        }
        fn last_dll_error(&self) -> i32 {
            self.last_dll_error
        }
    }

    struct MockComEnumerateFailHost {
        inner: NullHostServices,
    }

    impl MockComEnumerateFailHost {
        fn new() -> Self {
            Self {
                inner: NullHostServices::new(HostPolicy::default()),
            }
        }
    }

    impl HostServices for MockComEnumerateFailHost {
        fn profile(&self) -> HalProfileId {
            self.inner.profile()
        }
        fn descriptor(&self) -> HalDescriptor {
            self.inner.descriptor()
        }
        fn policy(&self) -> &HostPolicy {
            self.inner.policy()
        }
        fn console(&self) -> &dyn ConsoleHal {
            self.inner.console()
        }
        fn ui(&self) -> &dyn UiInteractionHal {
            self.inner.ui()
        }
        fn events(&self) -> &dyn EventPumpHal {
            self.inner.events()
        }
        fn fs(&self) -> &dyn FileSystemHal {
            self.inner.fs()
        }
        fn process(&self) -> &dyn ProcessEnvHal {
            self.inner.process()
        }
        fn com(&self) -> &dyn ComHal {
            self
        }
        fn time_locale(&self) -> &dyn TimeLocaleHal {
            self.inner.time_locale()
        }
        fn dynlink(&self) -> &dyn DynamicLinkHal {
            self.inner.dynlink()
        }
        fn diag(&self) -> &dyn DiagnosticsHal {
            self.inner.diag()
        }
    }

    impl ComHal for MockComEnumerateFailHost {
        fn create_object_variant(&self, _prog_id: Variant) -> HalResult<Variant> {
            Ok(Variant::from_object_ref(fake_foreign_object()))
        }

        fn describe_object(
            &self,
            _object: ObjectRef,
        ) -> HalResult<Option<calcvba_com::ComObjectDescriptor>> {
            Ok(None)
        }

        fn enumerate_object(&self, _object: ObjectRef) -> HalResult<Vec<Variant>> {
            Err(HalError::adapter_fault(
                HalProfileId::Windows,
                CapabilityId::ComActivationDispatch,
                "enumerate_object",
                "mock IEnumVARIANT failure",
            )
            .with_host_error_code(438))
        }

        fn subscribe_event(
            &self,
            object: ObjectRef,
            event: ComMemberToken,
        ) -> HalResult<ComSubscriptionToken> {
            self.inner.com().subscribe_event(object, event)
        }

        fn poll_event_callback(&self) -> HalResult<Option<ComCallbackPayload>> {
            self.inner.com().poll_event_callback()
        }

        fn event_callback_subscription(
            &self,
            callback: ComCallbackToken,
        ) -> HalResult<ComSubscriptionToken> {
            self.inner.com().event_callback_subscription(callback)
        }

        fn event_callback_arity(&self, callback: ComCallbackToken) -> HalResult<usize> {
            self.inner.com().event_callback_arity(callback)
        }

        fn resolve_typelib_reference(
            &self,
            request: &TypeLibResolveRequest,
        ) -> HalResult<TypeLibResolvedIdentity> {
            self.inner.com().resolve_typelib_reference(request)
        }

        fn load_typelib_metadata(
            &self,
            identity: &TypeLibResolvedIdentity,
        ) -> HalResult<TypeLibMetadataBlob> {
            self.inner.com().load_typelib_metadata(identity)
        }

        fn invalidate_typelib_cache(
            &self,
            scope: TypeLibCacheScope,
            reference_name: Option<&str>,
        ) -> HalResult<Variant> {
            self.inner
                .com()
                .invalidate_typelib_cache(scope, reference_name)
        }
    }

    #[repr(C)]
    struct FakeForeignObject {
        unknown: RawRuntimeIUnknown,
        ref_count: AtomicU32,
        drop_count: Arc<AtomicU32>,
    }

    impl Drop for FakeForeignObject {
        fn drop(&mut self) {
            self.drop_count.fetch_add(1, Ordering::AcqRel);
        }
    }

    static FAKE_FOREIGN_VTBL: IUnknownVtbl = IUnknownVtbl {
        query_interface: fake_foreign_query_interface,
        add_ref: fake_foreign_add_ref,
        release: fake_foreign_release,
    };

    unsafe extern "system" fn fake_foreign_query_interface(
        _this: *mut c_void,
        _iid: *const Guid,
        ppv: *mut *mut c_void,
    ) -> i32 {
        if !ppv.is_null() {
            // SAFETY: COM QueryInterface receives a caller-owned out pointer.
            unsafe {
                *ppv = std::ptr::null_mut();
            }
        }
        RUNTIME_E_NOINTERFACE
    }

    unsafe extern "system" fn fake_foreign_add_ref(this: *mut c_void) -> u32 {
        let object = this.cast::<FakeForeignObject>();
        // SAFETY: the fake vtable is installed only on FakeForeignObject boxes.
        unsafe { (*object).ref_count.fetch_add(1, Ordering::AcqRel) + 1 }
    }

    unsafe extern "system" fn fake_foreign_release(this: *mut c_void) -> u32 {
        let object = this.cast::<FakeForeignObject>();
        // SAFETY: the fake vtable is installed only on FakeForeignObject boxes.
        let previous = unsafe { (*object).ref_count.fetch_sub(1, Ordering::AcqRel) };
        let remaining = previous.saturating_sub(1);
        if remaining == 0 {
            // SAFETY: refcount reached zero for the box allocated in fake_foreign_object.
            unsafe {
                drop(Box::from_raw(object));
            }
        }
        remaining
    }

    fn fake_foreign_object_with_drop_count(drop_count: Arc<AtomicU32>) -> ObjectRef {
        let boxed = Box::new(FakeForeignObject {
            unknown: RawRuntimeIUnknown {
                vtbl: &FAKE_FOREIGN_VTBL,
            },
            ref_count: AtomicU32::new(1),
            drop_count,
        });
        let raw = Box::into_raw(boxed);
        // SAFETY: `FakeForeignObject` is `repr(C)` with `unknown` first. Casting the
        // complete allocation pointer gives the field's exact address while retaining
        // the full-box provenance that AddRef/Release need to access the adjacent
        // refcount and eventually reconstruct the Box. The initial reference transfers
        // to the returned ObjectRef.
        unsafe {
            ObjectRef::from_raw_iunknown_owned(raw.cast::<RawRuntimeIUnknown>())
                .expect("fake object pointer is non-null")
        }
    }

    fn fake_foreign_object() -> ObjectRef {
        fake_foreign_object_with_drop_count(Arc::new(AtomicU32::new(0)))
    }

    unsafe fn fake_foreign_object_ref_count(raw: *mut RawRuntimeIUnknown) -> u32 {
        let object = raw.cast::<FakeForeignObject>();
        // SAFETY: the caller supplies the live complete-allocation-derived interface
        // pointer minted by `fake_foreign_object_with_drop_count`.
        unsafe { (*object).ref_count.load(Ordering::Acquire) }
    }

    #[test]
    fn fake_foreign_object_addref_release_preserves_complete_allocation_provenance() {
        let drop_count = Arc::new(AtomicU32::new(0));
        let object = fake_foreign_object_with_drop_count(Arc::clone(&drop_count));
        let raw = object.raw_iunknown();

        // SAFETY: `object` owns the initial live reference to the fake allocation.
        assert_eq!(unsafe { fake_foreign_object_ref_count(raw) }, 1);
        let clone = object.clone();
        // SAFETY: the original and clone keep the same fake allocation live.
        assert_eq!(unsafe { fake_foreign_object_ref_count(raw) }, 2);
        // SAFETY: `raw` is a borrowed live IUnknown; this constructor takes one
        // additional reference and returns the matching owned ObjectRef.
        let retained = unsafe { ObjectRef::from_raw_iunknown_addref(raw) }
            .expect("borrowed fake object pointer is non-null");
        // SAFETY: three ObjectRef owners now keep the fake allocation live.
        assert_eq!(unsafe { fake_foreign_object_ref_count(raw) }, 3);

        drop(object);
        // SAFETY: `clone` and `retained` still keep the allocation live.
        assert_eq!(unsafe { fake_foreign_object_ref_count(raw) }, 2);
        assert_eq!(drop_count.load(Ordering::Acquire), 0);
        drop(clone);
        // SAFETY: `retained` still owns the final live reference.
        assert_eq!(unsafe { fake_foreign_object_ref_count(raw) }, 1);
        assert_eq!(drop_count.load(Ordering::Acquire), 0);
        drop(retained);
        assert_eq!(drop_count.load(Ordering::Acquire), 1);
        assert_eq!(Arc::strong_count(&drop_count), 1);
    }

    /// Elaborate a hand-built `CoreProgram` and run it on vm3 with a chosen host.
    fn run_core_with_host<'h>(prog: &CoreProgram, host: &'h dyn HostServices) -> Vm3<'h> {
        let oxp: &'static OxProgram = Box::leak(Box::new(
            calcvba_oxir::elaborate::elaborate(prog).expect("elaborate"),
        ));
        Vm3::run(oxp, host).expect("vm3 run")
    }

    /// A one-arg `Declare` descriptor (`descriptor_id` 0) whose single parameter is `ty`,
    /// `ByRef` per `by_ref`. The deterministic mock host ignores most of these fields.
    fn declare_descriptor(ty: DeclareParamType, by_ref: bool) -> ExternalCallDescriptor {
        ExternalCallDescriptor {
            descriptor_id: 0,
            declared_name: "Bump".into(),
            library: "t".into(),
            alias: "Bump".into(),
            ordinal_alias: false,
            symbol: DynLinkSymbol::new(0),
            marshal_lane: "m0-deterministic".into(),
            calling_convention: "platform-default".into(),
            selection_policy: "case-insensitive-canonical".into(),
            param_count: 1,
            param_types: vec![ty],
            param_by_ref: vec![by_ref],
            return_type: None,
            return_record_fields: None,
            return_record_name: None,
        }
    }

    #[test]
    fn foreach_over_foreign_object_surfaces_enumeration_failure() {
        let obj = || CorePlace::Local(CoreLocalId(0));
        let item = || CorePlace::Local(CoreLocalId(1));
        let touched = || CorePlace::Local(CoreLocalId(2));
        let err_number = || CorePlace::Local(CoreLocalId(3));
        let prog = main_proc(
            vec![
                local("obj", VarTypeRef::Variant),
                local("item", VarTypeRef::Variant),
                local("touched", VarTypeRef::Builtin(BuiltinType::Long)),
                local("err_number", VarTypeRef::Builtin(BuiltinType::Long)),
            ],
            vec![
                assign(
                    obj(),
                    CoreValue::Call {
                        callee: CoreCallee::Native(NativeImplId::CreateObject),
                        args: vec![CoreArg::ByVal(CoreValue::Const(CoreConst::Str(
                            "Probe.Object".into(),
                        )))],
                    },
                ),
                assign(touched(), CoreValue::Const(CoreConst::I32(0))),
                CoreStmt::Error(ErrorOp::OnErrorResumeNext),
                CoreStmt::ForEach {
                    item: item(),
                    source: CoreValue::Load(obj()),
                    body: vec![assign(touched(), CoreValue::Const(CoreConst::I32(1)))],
                },
                assign(err_number(), CoreValue::ErrField(ErrField::Number)),
            ],
        );
        let host = MockComEnumerateFailHost::new();
        let vm = run_core_with_host(&prog, &host);
        assert_eq!(
            vm.slot(2).and_then(|v| v.as_i32()),
            Some(0),
            "the failed For Each body must not run"
        );
        assert_eq!(
            vm.slot(3).and_then(|v| v.as_i32()),
            Some(438),
            "host enumeration failure must populate Err.Number, not become an empty loop"
        );
    }

    #[test]
    fn declare_byref_writes_back_to_the_call_site_slot() {
        // Declare Sub Bump Lib "t" (ByRef n As Long); r = 5; Bump r  ->  r = 105 (the mock
        // dynlink echoes +100). Proves the write-back targets the caller's `ByRef` slot.
        let mut prog = main_proc(
            vec![local("r", VarTypeRef::Builtin(BuiltinType::Long))],
            vec![
                assign(lc(0), CoreValue::Const(CoreConst::I32(5))),
                CoreStmt::Eval(CoreValue::Call {
                    callee: CoreCallee::Declare { descriptor_id: 0 },
                    args: vec![CoreArg::ByRef(lc(0))],
                }),
            ],
        );
        prog.external_calls = vec![declare_descriptor(DeclareParamType::Long, true)];
        let host = MockDynlinkHost::new(0);
        let vm = run_core_with_host(&prog, &host);
        assert_eq!(
            vm.slot(0).and_then(|v| v.as_i32()),
            Some(105),
            "a Declare ByRef arg must write the marshaled-back value to the caller slot"
        );
    }

    #[test]
    fn a_declare_refreshes_err_last_dll_error() {
        // After a Declare call, Err.LastDllError reads the OS last-error the HAL captured.
        let mut prog = main_proc(
            vec![local("e", VarTypeRef::Builtin(BuiltinType::Long))],
            vec![
                CoreStmt::Eval(CoreValue::Call {
                    callee: CoreCallee::Declare { descriptor_id: 0 },
                    args: vec![CoreArg::ByVal(CoreValue::Const(CoreConst::I32(1)))],
                }),
                assign(lc(0), CoreValue::ErrField(ErrField::LastDllError)),
            ],
        );
        prog.external_calls = vec![declare_descriptor(DeclareParamType::Long, false)];
        let host = MockDynlinkHost::new(2026);
        let vm = run_core_with_host(&prog, &host);
        assert_eq!(
            vm.slot(0).and_then(|v| v.as_i32()),
            Some(2026),
            "Err.LastDllError must reflect the dynlink HAL's last-error after a Declare"
        );
    }

    #[test]
    fn strptr_is_the_strings_own_bstr_and_a_declare_leaves_it_intact() {
        // s = "hi"; p1 = StrPtr(s); p2 = StrPtr(s); q = VarPtr(s); Poke StrPtr(s). StrPtr is
        // the variable's own BSTR (the same each time, as in VBA), VarPtr the slot holding it,
        // and passing the pointer to a Declare leaves `s` unchanged.
        let ptr = |kind| CoreValue::Ptr {
            kind,
            operand: calcvba_bundle::coreir::CorePtrOperand::Place(lc(0)),
        };
        let mut prog = main_proc(
            vec![
                local("s", VarTypeRef::Builtin(BuiltinType::String)),
                local("p1", VarTypeRef::Builtin(BuiltinType::LongPtr)),
                local("p2", VarTypeRef::Builtin(BuiltinType::LongPtr)),
                local("q", VarTypeRef::Builtin(BuiltinType::LongPtr)),
            ],
            vec![
                assign(lc(0), CoreValue::Const(CoreConst::Str("hi".into()))),
                assign(lc(1), ptr(PtrKind::Str)),
                assign(lc(2), ptr(PtrKind::Str)),
                assign(lc(3), ptr(PtrKind::VarString)),
                CoreStmt::Eval(CoreValue::Call {
                    callee: CoreCallee::Declare { descriptor_id: 0 },
                    args: vec![CoreArg::ByVal(ptr(PtrKind::Str))],
                }),
            ],
        );
        prog.external_calls = vec![declare_descriptor(DeclareParamType::LongPtr, false)];
        let host = MockDynlinkHost::new(0);
        let vm = run_core_with_host(&prog, &host);
        let s = calcvba_runtime::variant_to_vba_string(&vm.slot(0).expect("s"))
            .map(|b| b.as_str().to_string())
            .unwrap_or_default();
        assert_eq!(s, "hi");
        let p1 = vm.slot(1).and_then(|v| v.as_i64()).expect("p1");
        let p2 = vm.slot(2).and_then(|v| v.as_i64()).expect("p2");
        let q = vm.slot(3).and_then(|v| v.as_i64()).expect("q");
        assert!(
            p1 != 0 && p1 == p2,
            "StrPtr is the variable's own BSTR: {p1:#x} {p2:#x}"
        );
        assert!(
            q != 0 && q != p1,
            "VarPtr(s) is the slot, not the BSTR: {q:#x}"
        );
    }

    #[test]
    fn call_proc_ref_dispatches_through_address_of() {
        // A procedure reference (AddressOf, materialized by LoadProcRef) is called through
        // CallProcRef: Double(21) = 42 via a runtime-resolved proc index. These ops are
        // latent in the front-end (no producer yet) but consumed by the JIT, so the program
        // is built directly in OxIR. Mirrors vm2's `call_proc_ref_dispatches_through_address_of`.
        fn ox_local(name: &str, ty: OxTy, param: Option<OxParamInfo>) -> OxLocal {
            OxLocal {
                name: name.into(),
                ty,
                array_element: None,
                param,
                escaped: false,
            }
        }
        let main = OxFunc {
            name: "Main".into(),
            kind: ProcedureKind::Sub,
            locals: vec![
                ox_local("arg", OxTy::Long, None),  // Local 0
                ox_local("f", OxTy::ProcRef, None), // Local 1 (the AddressOf value)
                ox_local("n", OxTy::Long, None),    // Local 2 (result, snapshot slot 2)
            ],
            temps: Vec::new(),
            param_count: 0,
            return_local: None,
            blocks: vec![OxBlock {
                id: BlockId(0),
                instrs: vec![
                    OxInst::Assign {
                        dst: OxPlace::Local(LocalId(0)),
                        value: OxOperand::Const(OxConst::I32(21)),
                    },
                    OxInst::LoadProcRef {
                        dst: OxPlace::Local(LocalId(1)),
                        proc: FuncId(1),
                    },
                    OxInst::CallProcRef {
                        dst: Some(OxPlace::Local(LocalId(2))),
                        target: OxOperand::local(LocalId(1)),
                        args: vec![OxArg::ByVal(OxOperand::local(LocalId(0)))],
                    },
                ],
                fault_target: None,
                terminator: OxTerminator::Return,
            }],
            entry: BlockId(0),
        };
        let double = OxFunc {
            name: "Double".into(),
            kind: ProcedureKind::Function,
            locals: vec![
                ox_local(
                    "x",
                    OxTy::Long,
                    Some(OxParamInfo {
                        optional: false,
                        by_ref: false,
                        variadic: false,
                        default: None,
                    }),
                ), // Local 0 (param)
                ox_local("Double", OxTy::Long, None), // Local 1 (return)
            ],
            temps: Vec::new(),
            param_count: 1,
            return_local: Some(LocalId(1)),
            blocks: vec![OxBlock {
                id: BlockId(0),
                instrs: vec![OxInst::Arith {
                    dst: OxPlace::Local(LocalId(1)),
                    op: ArithOp::Add,
                    lhs: OxOperand::local(LocalId(0)),
                    rhs: OxOperand::local(LocalId(0)),
                    mode: NumericMode::Widening,
                }],
                fault_target: None,
                terminator: OxTerminator::Return,
            }],
            entry: BlockId(0),
        };
        let program = OxProgram {
            funcs: vec![main, double],
            entry: Some(FuncId(0)),
            unit_name: "T".into(),
            ..Default::default()
        };
        let oxp: &'static OxProgram = Box::leak(Box::new(program));
        let host: &'static NullHostServices =
            Box::leak(Box::new(NullHostServices::new(HostPolicy::default())));
        let vm = Vm3::run(oxp, host).expect("vm3 run");
        assert_eq!(
            vm.slot(2).and_then(|v| v.as_i32()),
            Some(42),
            "CallProcRef must dispatch through the AddressOf proc reference"
        );
    }

    #[test]
    fn call_native_typename_routes_through_the_veneer() {
        // Sub Main() : n = TypeName("hi")  ->  "String". Proves CallNative now goes
        // through the `invoke_native_lib` veneer (which only intercepts an *object*
        // argument, mirroring vm2): a non-object argument skips the interception and the
        // pure library body still answers correctly.
        let main = proc(
            "Main",
            ProcedureKind::Sub,
            Vec::new(),
            vec![local("n", VarTypeRef::Variant)],
            None,
            vec![assign(
                lc(0),
                CoreValue::Call {
                    callee: CoreCallee::Native(NativeImplId::TypeName),
                    args: vec![CoreArg::ByVal(CoreValue::Const(CoreConst::Str(
                        "hi".into(),
                    )))],
                },
            )],
        );
        let prog = procs_program(vec![main]);
        let vm = run_core(&prog);
        let s = calcvba_runtime::variant_to_vba_string(&vm.slot(0).expect("n"))
            .map(|b| b.as_str())
            .unwrap_or_default();
        assert_eq!(s, "String");
    }

    #[test]
    fn recursion_is_bounded_not_a_stack_overflow() {
        // Sub Spin() : Spin()  — unbounded self-recursion must surface as VBA error 28
        // ("Out of stack space"), not a native stack overflow / panic.
        let spin = proc(
            "Spin",
            ProcedureKind::Sub,
            Vec::new(),
            Vec::new(),
            None,
            vec![CoreStmt::Eval(CoreValue::Call {
                callee: CoreCallee::VbaProc { proc: ProcId(0) },
                args: Vec::new(),
            })],
        );
        let oxp: &'static OxProgram = Box::leak(Box::new(
            calcvba_oxir::elaborate::elaborate(&procs_program(vec![spin])).expect("elaborate"),
        ));
        let host: &'static NullHostServices =
            Box::leak(Box::new(NullHostServices::new(HostPolicy::default())));
        match Vm3::run(oxp, host) {
            Err(Vm3Error::Fault(f)) => assert_eq!(f.code, 28, "deep recursion is error 28"),
            Err(other) => panic!("expected an Out-of-stack fault, got error: {other}"),
            Ok(_) => panic!("expected an Out-of-stack fault, but the run completed"),
        }
    }

    // ── M2-c: error / Resume / Err / GoSub model ────────────────────────────────

    /// Run a single-proc program and expect it to end with an uncaught fault of `code`.
    fn run_expecting_fault(prog: &CoreProgram, code: i32) {
        let oxp: &'static OxProgram = Box::leak(Box::new(
            calcvba_oxir::elaborate::elaborate(prog).expect("elaborate"),
        ));
        let host: &'static NullHostServices =
            Box::leak(Box::new(NullHostServices::new(HostPolicy::default())));
        match Vm3::run(oxp, host) {
            Err(Vm3Error::Fault(f)) => assert_eq!(f.code, code, "expected uncaught error {code}"),
            Err(other) => panic!("expected uncaught error {code}, got error: {other}"),
            Ok(_) => panic!("expected uncaught error {code}, but the run completed"),
        }
    }

    /// `1 / 0` — a division-by-zero (error 11) expression.
    fn div_by_zero() -> CoreValue {
        CoreValue::Binary {
            op: CoreBinOp::Div,
            lhs: Box::new(CoreValue::Const(CoreConst::I32(1))),
            rhs: Box::new(CoreValue::Const(CoreConst::I32(0))),
            mode: StringCompareMode::Binary,
            num: NumericMode::Widening,
        }
    }

    #[test]
    fn on_error_resume_next_continues_and_reads_err() {
        // On Error Resume Next : n = 1/0 : n = Err.Number  ->  n = 11 (Resume Next skips
        // past the faulting statement; Err carries the division-by-zero code).
        let main = proc(
            "Main",
            ProcedureKind::Sub,
            Vec::new(),
            vec![local("n", VarTypeRef::Builtin(BuiltinType::Long))],
            None,
            vec![
                CoreStmt::Error(ErrorOp::OnErrorResumeNext),
                assign(lc(0), div_by_zero()),
                assign(lc(0), CoreValue::ErrField(ErrField::Number)),
            ],
        );
        let prog = procs_program(vec![main]);
        let vm = run_core(&prog);
        assert_eq!(vm.slot(0).and_then(|v| v.as_i32()), Some(11));
    }

    #[test]
    fn err_raise_is_caught_by_on_error_goto() {
        // On Error GoTo H : Err.Raise 5 : Exit Sub : H: n = Err.Number  ->  n = 5
        // (proves Err.Raise routes through the statement pad so On Error catches it).
        let main = proc(
            "Main",
            ProcedureKind::Sub,
            Vec::new(),
            vec![local("n", VarTypeRef::Builtin(BuiltinType::Long))],
            None,
            vec![
                CoreStmt::Error(ErrorOp::OnErrorGotoLabel(LabelId(0))),
                CoreStmt::Error(ErrorOp::Raise {
                    number: CoreValue::Const(CoreConst::I32(5)),
                    source: None,
                    description: None,
                    help_file: None,
                    help_context: None,
                    inherit: true,
                }),
                CoreStmt::Exit(ExitKind::Proc),
                CoreStmt::Label(LabelId(0)),
                assign(lc(0), CoreValue::ErrField(ErrField::Number)),
            ],
        );
        let prog = procs_program(vec![main]);
        let vm = run_core(&prog);
        assert_eq!(vm.slot(0).and_then(|v| v.as_i32()), Some(5));
    }

    #[test]
    fn resume_re_runs_the_faulting_statement() {
        // On Error GoTo H : k = k+1 : m = 1/(k-1) [faults when k=1] : n = 42 : Exit Sub
        // H: k = k+1 : Resume   ->  the division statement re-runs with k=2 (1/1 ok), so
        // control reaches `n = 42`. `Resume` re-enters the *faulting* statement, and the
        // handler's k-bump prevents an infinite re-fault. (`n` is a clean Long literal so
        // the assertion is unaffected by the Double division result, which lands in `m`.)
        let k = || lc(2);
        let div = CoreValue::Binary {
            op: CoreBinOp::Div,
            lhs: Box::new(CoreValue::Const(CoreConst::I32(1))),
            rhs: Box::new(CoreValue::Binary {
                op: CoreBinOp::Sub,
                lhs: Box::new(load(2)),
                rhs: Box::new(CoreValue::Const(CoreConst::I32(1))),
                mode: StringCompareMode::Binary,
                num: NumericMode::Widening,
            }),
            mode: StringCompareMode::Binary,
            num: NumericMode::Widening,
        };
        let main = proc(
            "Main",
            ProcedureKind::Sub,
            Vec::new(),
            vec![
                local("n", VarTypeRef::Builtin(BuiltinType::Long)), // 0
                local("m", VarTypeRef::Variant),                    // 1
                local("k", VarTypeRef::Builtin(BuiltinType::Long)), // 2
            ],
            None,
            vec![
                CoreStmt::Error(ErrorOp::OnErrorGotoLabel(LabelId(0))),
                assign(k(), long_add(load(2), CoreValue::Const(CoreConst::I32(1)))),
                assign(lc(1), div), // faulting statement: 1/(k-1)
                assign(lc(0), CoreValue::Const(CoreConst::I32(42))),
                CoreStmt::Exit(ExitKind::Proc),
                CoreStmt::Label(LabelId(0)),
                assign(k(), long_add(load(2), CoreValue::Const(CoreConst::I32(1)))),
                CoreStmt::Error(ErrorOp::Resume),
            ],
        );
        let prog = procs_program(vec![main]);
        let vm = run_core(&prog);
        assert_eq!(vm.slot(0).and_then(|v| v.as_i32()), Some(42));
    }

    #[test]
    fn resume_without_active_error_raises_20() {
        // Resume Next with no active error -> runtime error 20 (Resume without error).
        let main = proc(
            "Main",
            ProcedureKind::Sub,
            Vec::new(),
            Vec::new(),
            None,
            vec![CoreStmt::Error(ErrorOp::ResumeNext)],
        );
        run_expecting_fault(&procs_program(vec![main]), 20);
    }

    #[test]
    fn return_without_gosub_raises_3() {
        // A bare Return (no GoSub on the stack) -> runtime error 3 (Return without GoSub).
        let main = proc(
            "Main",
            ProcedureKind::Sub,
            Vec::new(),
            Vec::new(),
            None,
            vec![CoreStmt::GoSubReturn],
        );
        run_expecting_fault(&procs_program(vec![main]), 3);
    }

    #[test]
    fn gosub_resumption_list_is_lifo() {
        // GoSub A : Exit Sub : A: GoSub B : s = s & "A" : Return : B: s = s & "B" : Return
        // -> s = "BA": B's Return pops the inner ret, A's Return pops the outer (LIFO).
        let s = || lc(0);
        let cat = |suffix: &str| {
            assign(
                s(),
                CoreValue::Binary {
                    op: CoreBinOp::Concat,
                    lhs: Box::new(load(0)),
                    rhs: Box::new(CoreValue::Const(CoreConst::Str(suffix.into()))),
                    mode: StringCompareMode::Binary,
                    num: NumericMode::Widening,
                },
            )
        };
        let main = proc(
            "Main",
            ProcedureKind::Sub,
            Vec::new(),
            vec![local("s", VarTypeRef::Builtin(BuiltinType::String))],
            None,
            vec![
                CoreStmt::GoSub(LabelId(0)),
                CoreStmt::Exit(ExitKind::Proc),
                CoreStmt::Label(LabelId(0)), // A
                CoreStmt::GoSub(LabelId(1)),
                cat("A"),
                CoreStmt::GoSubReturn,
                CoreStmt::Label(LabelId(1)), // B
                cat("B"),
                CoreStmt::GoSubReturn,
            ],
        );
        let prog = procs_program(vec![main]);
        let vm = run_core(&prog);
        let s = calcvba_runtime::variant_to_vba_string(&vm.slot(0).expect("s"))
            .map(|b| b.as_str())
            .unwrap_or_default();
        assert_eq!(s, "BA");
    }
}
pub mod typed;

/// The value a declared scalar of type `ty` starts with (`Long` 0, `String` ""), for a slot
/// that no assignment has reached; `None` for Variant, object, record and array slots.
fn scalar_zero(ty: &OxTy) -> Option<Variant> {
    Some(match ty {
        OxTy::Bool => Variant::from_bool(false),
        OxTy::Byte => Variant::from_u8(0),
        OxTy::Integer => Variant::from_i16(0),
        OxTy::Long => Variant::from_i32(0),
        OxTy::LongLong => Variant::from_i64(0),
        OxTy::Single => Variant::from_f32(0.0),
        OxTy::Double => Variant::from_f64(0.0),
        OxTy::Currency => Variant::from_currency_scaled_i64(0),
        OxTy::Date => Variant::from_date_f64(0.0),
        OxTy::Str => Variant::from_string(""),
        _ => return None,
    })
}
