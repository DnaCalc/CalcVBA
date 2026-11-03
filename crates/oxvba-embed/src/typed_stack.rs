//! Retained call/ByRef/error mechanics shared by the two typed backends.
//! Each backend still executes its original OxFunc arithmetic and host calls.
use calcvba_host_api::{ScalarBits, TypedContext, Value, ValueType as ScalarType};
use oxvba_oxir::{
    BlockId, ErrField, FuncId, OxConst, OxInst, OxOperand, OxPlace, OxProgram, OxTerminator,
};
use oxvba_rt_abi::host_entry::LinkedTarget;
use oxvba_rt_abi::host_entry::activation::{Activation, CONTROL, SUSPENDED};
use oxvba_rt_abi::{ErrEngine, Fault, FaultAction, ResumeTarget, SavedErrState};
use oxvba_vm3::typed::TypedInterpreter;

struct Frame {
    // Box keeps ancestor alias targets stable as the stack grows.
    activation: Box<Activation<'static>>,
    procedure: FuncId,
    program: usize,
    destination: Option<OxPlace>,
    saved_error: SavedErrState,
    line: i32,
    finished: bool,
    initialized: Option<(OxPlace, Value)>,
    termination: Option<i32>,
}
pub(crate) struct TypedStack {
    frames: Vec<Frame>,
    error: *mut ErrEngine,
    stores: Vec<(*mut oxvba_runtime::Variant, usize)>,
    classes: Option<oxvba_rt_abi::TypedClassContext>,
    terminations: std::collections::VecDeque<(i32, i32, i32)>,
    owns_drain: bool,
    writeback_slots: Vec<usize>,
    writebacks: Vec<(usize, Value)>,
}
impl TypedStack {
    pub(crate) fn error(&mut self, number: i32, source: &str) -> crate::Error {
        if number == -1 {
            return crate::Error::BudgetExceeded;
        }
        let err = &self.err().err;
        crate::Error::Runtime {
            number,
            description: if err.number == number {
                err.description.clone()
            } else {
                oxvba_rt_abi::default_error_message(number)
            },
            vba_source: if err.number == number {
                err.source.clone()
            } else {
                source.into()
            },
        }
    }
    /// # Safety
    /// error and the activation's canonical globals stay live under the same
    /// session lease, on this thread, until the stack drops. No other execution
    /// may use that state except synchronous guarded callback reentry.
    pub(crate) unsafe fn new(
        activation: Activation<'static>,
        program: usize,
        procedure: FuncId,
        stores: Vec<(*mut oxvba_runtime::Variant, usize)>,
        error: *mut ErrEngine,
        classes: Option<oxvba_rt_abi::TypedClassContext>,
    ) -> Self {
        // SAFETY: caller retains unique session execution authority.
        let saved_error = unsafe { (&mut *error).enter_activation() };
        Self {
            frames: vec![Frame {
                activation: Box::new(activation),
                procedure,
                program,
                destination: None,
                saved_error,
                line: 0,
                finished: false,
                initialized: None,
                termination: None,
            }],
            error,
            stores,
            classes,
            terminations: Default::default(),
            owns_drain: false,
            writeback_slots: Vec::new(),
            writebacks: Vec::new(),
        }
    }
    /// Enable explicit synchronous copyback of root ByRef scalar parameters.
    /// Structural/object aliases require a separate caller-owned cell contract.
    pub(crate) fn enable_writebacks(&mut self, count: usize) -> Result<(), crate::Error> {
        let layout = self.frames[0].activation.layout();
        for (index, (&ty, &by_ref)) in layout
            .parameters
            .iter()
            .zip(&layout.by_ref_parameters)
            .take(count)
            .enumerate()
        {
            if !by_ref {
                continue;
            }
            if !matches!(
                ty,
                ScalarType::Long | ScalarType::Double | ScalarType::Boolean | ScalarType::String
            ) {
                return Err(crate::Error::UnsupportedSignature(
                    "synchronous host copyback supports primitive ByRef parameters only".into(),
                ));
            }
            self.writeback_slots.push(index);
        }
        Ok(())
    }
    fn snapshot_writebacks(&mut self) -> Result<(), i32> {
        let root = &mut self.frames[0].activation;
        let mut values = Vec::with_capacity(self.writeback_slots.len());
        for &index in &self.writeback_slots {
            let ty = root.layout().parameters[index];
            let bits = root.read(
                &OxOperand::Use(OxPlace::Local(oxvba_oxir::LocalId(index))),
                ty,
            )?;
            // SAFETY: the retained root owns the exact admitted primitive carrier.
            values.push((index, unsafe {
                calcvba_host_api::value::retain_bits(ty, bits)
            }?));
        }
        self.writebacks = values;
        Ok(())
    }
    pub(crate) fn take_writebacks(&mut self) -> Result<Vec<(usize, Value)>, i32> {
        if !self.frames[0].finished {
            self.snapshot_writebacks()?;
        }
        Ok(std::mem::take(&mut self.writebacks))
    }
    pub(crate) fn result(&mut self) -> Value {
        self.frames[0].activation.take_result()
    }
    pub(crate) fn raw_result(&self) -> ScalarBits {
        let frame = &self.frames[0].activation;
        if frame.layout().has_result {
            frame.raw().slots[frame.layout().return_slot]
        } else {
            ScalarBits::ZERO
        }
    }
    fn err(&mut self) -> &mut ErrEngine {
        // SAFETY: the stack's session lease/guards exclude competing access;
        // this short borrow never spans a backend entry or host callback.
        unsafe { &mut *self.error }
    }
    fn pop(&mut self) -> Frame {
        let frame = self.frames.pop().expect("retained frame");
        self.err().restore(frame.saved_error);
        frame
    }
    fn goto(&mut self, programs: &[&OxProgram], target: BlockId) -> Result<(), i32> {
        let frame = self.frames.last_mut().ok_or(5)?;
        let func = &programs[frame.program].funcs[frame.procedure.0];
        let mut pc = 0u32;
        for block in &func.blocks {
            if block.id == target {
                frame.activation.set_pc(pc);
                return Ok(());
            }
            pc = pc
                .checked_add(u32::try_from(block.instrs.len()).map_err(|_| 5)? + 2)
                .ok_or(5)?;
        }
        Err(5)
    }
    fn position(&self, programs: &[&OxProgram]) -> Result<(usize, usize), i32> {
        let frame = self.frames.last().ok_or(5)?;
        let mut start = 0u32;
        for (i, block) in programs[frame.program].funcs[frame.procedure.0]
            .blocks
            .iter()
            .enumerate()
        {
            let end = start
                .checked_add(u32::try_from(block.instrs.len()).map_err(|_| 5)? + 2)
                .ok_or(5)?;
            if frame.activation.pc() < end {
                return Ok((i, (frame.activation.pc() - start) as usize));
            }
            start = end;
        }
        Err(5)
    }
    fn route(&mut self, programs: &[&OxProgram], fault: Fault) -> Result<(), i32> {
        let (block, _) = self.position(programs)?;
        let frame = self.frames.last().ok_or(5)?;
        let program = programs[frame.program];
        let pad = program.funcs[frame.procedure.0].blocks[block]
            .fault_target
            .ok_or(fault.code)?;
        self.err().raise(fault, program.unit_name.clone());
        self.goto(programs, pad)
    }
    fn operand(&mut self, operand: &OxOperand) -> Result<oxvba_runtime::Variant, i32> {
        let frame = &mut self.frames.last_mut().ok_or(5)?.activation;
        let ty = frame.layout().operand_type(operand).map_err(|_| 5)?;
        let bits = frame.read(operand, ty)?;
        // SAFETY: the executing frame retains the exact admitted operand carrier.
        Ok(oxvba_rt_abi::host_entry::value::to_variant(unsafe {
            calcvba_host_api::value::retain_bits(ty, bits)
        }?))
    }
    fn project_object(&mut self, operand: &OxOperand) -> Result<oxvba_runtime::ObjectRef, i32> {
        let value = self.operand(operand)?;
        if value.vtype() != oxvba_runtime::VarType::Object {
            return Err(424);
        }
        let object = value.as_object_ref().ok_or(91)?;
        if !object.is_project_instance()
            || object.object_owner_id()
                != Some(oxvba_runtime::object_owner::RuntimeObjectOwner::current().id())
        {
            return Err(13);
        }
        Ok(object)
    }
    fn push_me(
        &mut self,
        entries: &[Vec<Result<TypedInterpreter<'_>, String>>],
        unit: usize,
        proc: FuncId,
        object: &oxvba_runtime::ObjectRef,
    ) -> Result<(), i32> {
        if self.frames.len() >= 64 {
            return Err(28);
        }
        let typed = entries
            .get(unit)
            .and_then(|p| p.get(proc.0))
            .and_then(|e| e.as_ref().ok())
            .ok_or(5)?;
        let me = oxvba_runtime::Variant::from_object_ref(object.clone());
        let arg = match typed.layout().parameters.first() {
            Some(ScalarType::Variant) => calcvba_host_api::ValueSlot::Variant(&me),
            Some(ScalarType::HostObject) => calcvba_host_api::ValueSlot::RuntimeObject(object),
            _ => return Err(13),
        };
        // SAFETY: this stack retains the canonical session stores and owns the child.
        let activation = unsafe {
            Activation::new_values(
                typed.layout(),
                &[arg],
                self.stores[unit].0,
                self.stores[unit].1,
                typed.entry_pc(),
            )
        }?;
        let saved_error = self.err().enter_activation();
        self.frames.push(Frame {
            activation: Box::new(activation),
            procedure: proc,
            program: unit,
            destination: None,
            saved_error,
            line: 0,
            finished: false,
            initialized: None,
            termination: None,
        });
        Ok(())
    }
    fn stop_drain(&mut self) {
        if self.owns_drain {
            if let Some(classes) = &self.classes {
                // SAFETY: the canonical draining flag stays leased by this stack.
                unsafe {
                    *classes.draining = false;
                }
            }
            self.owns_drain = false;
        }
    }
    fn drive_drain(
        &mut self,
        programs: &[&OxProgram],
        entries: &[Vec<Result<TypedInterpreter<'_>, String>>],
    ) -> Result<bool, i32> {
        let Some(classes) = &self.classes else {
            return Ok(false);
        };
        // SAFETY: short access to the canonical session flag; no state borrow spans a finalizer.
        if unsafe { *classes.draining } && !self.owns_drain
            || self.frames.iter().any(|frame| frame.termination.is_some())
        {
            return Ok(false);
        }
        if !self.owns_drain {
            // SAFETY: this stack now owns the finalizer drain until completion/drop.
            unsafe {
                *classes.draining = true;
            }
            self.owns_drain = true;
        }
        loop {
            if self.terminations.is_empty() {
                self.terminations
                    .extend(oxvba_runtime::take_pending_terminations());
            }
            let Some((id, unit, class)) = self.terminations.pop_front() else {
                self.stop_drain();
                break;
            };
            if let Some(proc) = programs
                .get(unit as usize)
                .and_then(|p| p.classes.get(class as usize))
                .and_then(|c| c.terminate)
            {
                let object = oxvba_runtime::retained_parked_termination_object(id).ok_or(5)?;
                if let Err(code) = self.push_me(entries, unit as usize, proc, &object) {
                    drop(object);
                    oxvba_runtime::finish_pending_termination(id);
                    return Err(code);
                }
                self.frames.last_mut().ok_or(5)?.termination = Some(id);
                return Ok(true);
            }
            oxvba_runtime::finish_pending_termination(id);
        }

        Ok(false)
    }
    pub(crate) fn resume(
        &mut self,
        programs: &[&OxProgram],
        entries: &[Vec<Result<TypedInterpreter<'_>, String>>],
        context: &mut TypedContext,
        mut operations: u64,
        mut execute: impl FnMut(
            usize,
            FuncId,
            &mut Activation<'_>,
            &mut TypedContext,
            u64,
        ) -> Result<i32, i32>,
    ) -> Result<i32, i32> {
        while operations != 0 {
            let frame = self.frames.last_mut().ok_or(5)?;
            let result = if frame.finished {
                Ok(0)
            } else {
                let result = execute(
                    frame.program,
                    frame.procedure,
                    &mut frame.activation,
                    context,
                    operations,
                );
                operations = frame.activation.raw().remaining_operations;
                result
            };
            // Ancestors remain parked while descendants hold their aliases.
            // Global alias reads retain scratch on the executing frame itself.
            self.frames.last_mut().ok_or(5)?.activation.clear_scratch();
            match result {
                Ok(SUSPENDED) => return Ok(SUSPENDED),
                Ok(0) => {
                    if self.frames.len() == 1 && !self.frames[0].finished {
                        self.snapshot_writebacks()?;
                    }
                    let frame = self.frames.last_mut().ok_or(5)?;
                    if !frame.finished {
                        frame.activation.release_locals_except_result();
                        frame.finished = true;
                    }
                    if self.drive_drain(programs, entries)? {
                        continue;
                    }
                    if self.frames.len() == 1 {
                        return Ok(0);
                    }
                    let mut callee = self.pop();
                    if let Some(id) = callee.termination {
                        drop(callee);
                        oxvba_runtime::finish_pending_termination(id);
                        continue;
                    }
                    if let Some((dst, value)) = callee.initialized.take() {
                        self.frames
                            .last_mut()
                            .ok_or(5)?
                            .activation
                            .write_owned_value(dst, value)?;
                    }
                    if let Some(dst) = callee.destination {
                        let mut value = callee.activation.take_result();
                        let parent = &self.frames.last().ok_or(5)?.activation;
                        let target = parent
                            .layout()
                            .operand_type(&OxOperand::Use(dst))
                            .map_err(|_| 5)?;
                        if target != value.ty() {
                            value = oxvba_rt_abi::host_entry::value::from_variant(
                                target,
                                oxvba_rt_abi::host_entry::value::to_variant(value),
                            )?;
                        }
                        self.frames
                            .last_mut()
                            .ok_or(5)?
                            .activation
                            .write_owned_value(dst, value)?;
                    }
                    let parent = &mut self.frames.last_mut().ok_or(5)?.activation;
                    parent.set_pc(parent.pc() + 1);
                }
                Ok(CONTROL) => {
                    if let Err(number) = self.control(programs, entries, context) {
                        self.route(
                            programs,
                            Fault::new(number, oxvba_rt_abi::default_error_message(number)),
                        )?;
                    }
                }
                Err(-1) => return Err(-1),
                Err(number) => {
                    let fault = calcvba_host_api::take_typed_fault(context, number).map_or_else(
                        || Fault::new(number, oxvba_rt_abi::default_error_message(number)),
                        |fault| Fault {
                            code: fault.number,
                            message: fault.description,
                            source: fault.source,
                            help_file: fault.help_file,
                            help_context: fault.help_context,
                        },
                    );
                    self.route(programs, fault)?;
                }
                _ => return Err(5),
            }
        }
        Ok(SUSPENDED)
    }
    fn control(
        &mut self,
        programs: &[&OxProgram],
        entries: &[Vec<Result<TypedInterpreter<'_>, String>>],
        context: &mut TypedContext,
    ) -> Result<(), i32> {
        let (block_id, offset) = self.position(programs)?;
        let current = self.frames.last().ok_or(5)?.program;
        let program = programs[current];
        let procedure = self.frames.last().ok_or(5)?.procedure;
        let block = &program.funcs[procedure.0].blocks[block_id];
        if let Some(inst) = offset.checked_sub(1).and_then(|i| block.instrs.get(i)) {
            match inst {
                OxInst::CallProc { args, dst, .. } | OxInst::CallExtern { args, dst, .. } => {
                    let (unit, proc) = match inst {
                        OxInst::CallProc { proc, .. } => (current, *proc),
                        OxInst::CallExtern { import, .. } => {
                            let parent = entries[current][procedure.0].as_ref().map_err(|_| 5)?;
                            let target = parent.layout().extern_calls.get(import).ok_or(5)?;
                            if let LinkedTarget::Builtin(
                                native @ (oxvba_bundle::NativeImplId::IsMissing
                                | oxvba_bundle::NativeImplId::IsArray
                                | oxvba_bundle::NativeImplId::VarType
                                | oxvba_bundle::NativeImplId::CStr),
                            ) = target.target
                            {
                                let operand = match &args[0] {
                                    oxvba_oxir::OxArg::ByVal(op) => op.clone(),
                                    oxvba_oxir::OxArg::ByRef(p) => OxOperand::Use(*p),
                                    _ => return Err(13),
                                };
                                let bits = self
                                    .frames
                                    .last_mut()
                                    .ok_or(5)?
                                    .activation
                                    .read(&operand, ScalarType::Variant)?;
                                // SAFETY: the executing frame retains this exact Variant cell.
                                let Value::Variant(value) = (unsafe {
                                    calcvba_host_api::value::retain_bits(ScalarType::Variant, bits)
                                })?
                                else {
                                    return Err(13);
                                };
                                let result = oxvba_rt_abi::host_entry::value::invoke_unary_builtin(
                                    native,
                                    value,
                                    entries[current][procedure.0]
                                        .as_ref()
                                        .map_err(|_| 5)?
                                        .layout(),
                                    context,
                                )?;
                                let frame = &mut self.frames.last_mut().ok_or(5)?.activation;
                                if let Some(dst) = dst {
                                    frame.write_owned_value(*dst, Value::Variant(result))?;
                                }
                                frame.set_pc(frame.pc() + 1);
                                return Ok(());
                            }
                            let LinkedTarget::Procedure { program, procedure } = target.target
                            else {
                                return Err(5);
                            };
                            (program, procedure)
                        }
                        _ => return Err(5),
                    };
                    if self.frames.len() >= 64 {
                        return Err(28);
                    }
                    let typed = entries
                        .get(unit)
                        .and_then(|p| p.get(proc.0))
                        .and_then(|e| e.as_ref().ok())
                        .ok_or(5)?;
                    // SAFETY: all ancestor activations are boxed at stable addresses;
                    // children are popped before parents on return, fault and cancel.
                    let activation = unsafe {
                        Activation::child_with_globals(
                            &mut self.frames.last_mut().ok_or(5)?.activation,
                            typed.layout(),
                            args,
                            typed.entry_pc(),
                            self.stores[unit].0,
                            self.stores[unit].1,
                        )
                    }?;
                    let saved_error = self.err().enter_activation();
                    self.frames.push(Frame {
                        activation: Box::new(activation),
                        procedure: proc,
                        program: unit,
                        destination: *dst,
                        saved_error,
                        line: 0,
                        finished: false,
                        initialized: None,
                        termination: None,
                    });
                    return Ok(());
                }
                OxInst::ComCallLate {
                    dst, recv, args, ..
                } => {
                    let (proc, class) = self
                        .frames
                        .last()
                        .ok_or(5)?
                        .activation
                        .layout()
                        .class_calls
                        .iter()
                        .find(|(site, _, _)| site == inst)
                        .map(|(_, p, c)| (*p, *c))
                        .ok_or(5)?;
                    let object = self.project_object(recv)?;
                    if object.bundle_id() != current as i32 || object.route_key() != class.0 as i32
                    {
                        return Err(13);
                    }
                    let typed = entries[current][proc.0].as_ref().map_err(|_| 5)?;
                    let mut operands = Vec::with_capacity(args.len() + 1);
                    operands.push(oxvba_oxir::OxArg::ByVal(recv.clone()));
                    for (index, arg) in args.iter().enumerate() {
                        operands.push(match arg {
                            oxvba_oxir::OxCallArg::Operand(op) => {
                                oxvba_oxir::OxArg::ByVal(op.clone())
                            }
                            oxvba_oxir::OxCallArg::ByRef(place)
                                if typed.layout().by_ref_parameters.get(index + 1)
                                    == Some(&true) =>
                            {
                                oxvba_oxir::OxArg::ByRef(*place)
                            }
                            oxvba_oxir::OxCallArg::ByRef(place) => {
                                oxvba_oxir::OxArg::ByVal(OxOperand::Use(*place))
                            }
                            oxvba_oxir::OxCallArg::Omitted => oxvba_oxir::OxArg::Omitted,
                            _ => return Err(13),
                        });
                    }
                    if self.frames.len() >= 64 {
                        return Err(28);
                    }
                    // SAFETY: boxed ancestors and the unit store remain leased until this child drops.
                    let activation = unsafe {
                        Activation::child_with_globals(
                            &mut self.frames.last_mut().ok_or(5)?.activation,
                            typed.layout(),
                            &operands,
                            typed.entry_pc(),
                            self.stores[current].0,
                            self.stores[current].1,
                        )
                    }?;
                    let saved_error = self.err().enter_activation();
                    self.frames.push(Frame {
                        activation: Box::new(activation),
                        procedure: proc,
                        program: current,
                        destination: *dst,
                        saved_error,
                        line: 0,
                        finished: false,
                        initialized: None,
                        termination: None,
                    });
                    return Ok(());
                }
                OxInst::NewObject { dst, class } => {
                    let plan = program.classes.get(class.0).ok_or(5)?;
                    let classes = self.classes.as_ref().ok_or(5)?;
                    // SAFETY: the typed stack holds the originating session execution lease.
                    let object =
                        unsafe { classes.allocate(current, class.0, plan.terminate.is_some()) }?;
                    if let Some(initialize) = plan.initialize {
                        self.push_me(entries, current, initialize, &object)?;
                        self.frames.last_mut().ok_or(5)?.initialized =
                            Some((*dst, Value::RuntimeObject(object)));
                        return Ok(());
                    }
                    self.frames
                        .last_mut()
                        .ok_or(5)?
                        .activation
                        .write_owned_value(*dst, Value::RuntimeObject(object))?;
                }
                OxInst::FieldGet { dst, object, field } => {
                    let object = self.project_object(object)?;
                    let class = programs
                        .get(object.bundle_id() as usize)
                        .and_then(|p| p.classes.get(object.route_key() as usize))
                        .ok_or(13)?;
                    let declaration = class.fields.iter().find(|f| f.token == *field).ok_or(438)?;
                    let mut value = object.project_field_get(*field).ok_or(438)?;
                    if value.vtype() == oxvba_runtime::VarType::Empty {
                        value = match declaration.ty {
                            oxvba_oxir::OxTy::Long => oxvba_runtime::Variant::from_i32(0),
                            oxvba_oxir::OxTy::Double => oxvba_runtime::Variant::from_f64(0.0),
                            oxvba_oxir::OxTy::Bool => oxvba_runtime::Variant::from_bool(false),
                            oxvba_oxir::OxTy::Str => oxvba_runtime::Variant::from_string(""),
                            oxvba_oxir::OxTy::Object(_) => oxvba_runtime::Variant::nothing(),
                            oxvba_oxir::OxTy::Variant => value,
                            _ => return Err(13),
                        };
                    }
                    let frame = &mut self.frames.last_mut().ok_or(5)?.activation;
                    let ty = frame
                        .layout()
                        .operand_type(&OxOperand::Use(*dst))
                        .map_err(|_| 5)?;
                    frame.write_owned_value(
                        *dst,
                        oxvba_rt_abi::host_entry::value::from_variant(ty, value)?,
                    )?;
                }
                OxInst::FieldSet {
                    object,
                    field,
                    value,
                } => {
                    let object = self.project_object(object)?;
                    let value = self.operand(value)?;
                    if !object.project_field_set(*field, value) {
                        return Err(438);
                    }
                }
                OxInst::StmtBoundary { .. } | OxInst::DrainTerminations => {
                    if let OxInst::StmtBoundary {
                        clear_temps_from, ..
                    } = inst
                    {
                        self.frames
                            .last_mut()
                            .ok_or(5)?
                            .activation
                            .clear_temps(*clear_temps_from);
                    }
                    if self.drive_drain(programs, entries)? {
                        return Ok(());
                    }
                }
                OxInst::SetErrorHandler(handler) => self.err().set_error_handler(handler),
                OxInst::ClearErr => self.err().clear_err(),
                OxInst::SetLineNumber { line } => self.frames.last_mut().ok_or(5)?.line = *line,
                OxInst::ErrFieldGet { dst, field } => {
                    let err = self.err();
                    let value = match field {
                        ErrField::Number => Value::Long(err.err.number),
                        ErrField::HelpContext => Value::Long(err.err.help_context),
                        ErrField::LastDllError => Value::Long(err.last_dll_error),
                        ErrField::Description => Value::String(err.err.description.as_str().into()),
                        ErrField::Source => Value::String(err.err.source.as_str().into()),
                        ErrField::HelpFile => Value::String(err.err.help_file.as_str().into()),
                    };
                    self.frames
                        .last_mut()
                        .ok_or(5)?
                        .activation
                        .write_owned_value(*dst, value)?;
                }
                OxInst::ErlGet { dst } => {
                    let value = self.err().erl_line;
                    // SAFETY: Erl's destination is an admitted Long.
                    unsafe {
                        self.frames
                            .last_mut()
                            .ok_or(5)?
                            .activation
                            .write(*dst, ScalarBits::long(value))
                    }?;
                }
                _ => return Err(5),
            }
            let frame = &mut self.frames.last_mut().ok_or(5)?.activation;
            frame.set_pc(frame.pc() + 1);
            return Ok(());
        }
        match &block.terminator {
            OxTerminator::FaultDispatch {
                resume,
                resume_next,
            } => {
                let line = self.frames.last().ok_or(5)?.line;
                match self
                    .err()
                    .dispatch_fault(*resume, *resume_next, line)
                    .map_err(|_| 5)?
                {
                    FaultAction::Handle(target) | FaultAction::ResumeNext(target) => {
                        self.goto(programs, target)
                    }
                    FaultAction::Propagate(fault) => {
                        if self.frames.len() == 1 {
                            return Err(fault.code);
                        }
                        let frame = self.pop();
                        if let Some(id) = frame.termination {
                            drop(frame);
                            oxvba_runtime::finish_pending_termination(id);
                            return Ok(());
                        }
                        drop(frame);
                        self.route(programs, fault)
                    }
                }
            }
            OxTerminator::Resume | OxTerminator::ResumeNext | OxTerminator::ResumeLabel(_) => {
                let target = match block.terminator {
                    OxTerminator::Resume => ResumeTarget::Same,
                    OxTerminator::ResumeNext => ResumeTarget::Next,
                    OxTerminator::ResumeLabel(label) => ResumeTarget::Label(label),
                    _ => unreachable!(),
                };
                match self.err().resume(target) {
                    Ok(target) => self.goto(programs, target),
                    Err(fault) => self.route(programs, fault),
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
                let frame = &mut self.frames.last_mut().ok_or(5)?.activation;
                let number = frame.read(number, ScalarType::Long)?.as_long();
                let help = help_context
                    .as_ref()
                    .map(|op| frame.read(op, ScalarType::Long).map(ScalarBits::as_long))
                    .transpose()?;
                let text = |op: &Option<OxOperand>| match op {
                    Some(OxOperand::Const(OxConst::Str(s))) => Some(s.clone()),
                    _ => None,
                };
                let err = self.err();
                let inherit = *inherit && err.err.inherit_fields;
                let message = text(description).unwrap_or_else(|| {
                    if inherit {
                        err.err.description.clone()
                    } else {
                        oxvba_rt_abi::default_error_message(number)
                    }
                });
                let fault = Fault {
                    code: number,
                    message,
                    source: text(source).or_else(|| inherit.then(|| err.err.source.clone())),
                    help_file: text(help_file)
                        .or_else(|| inherit.then(|| err.err.help_file.clone())),
                    help_context: help.or_else(|| inherit.then_some(err.err.help_context)),
                };
                self.route(programs, fault)
            }
            _ => Err(5),
        }
    }
}
impl Drop for TypedStack {
    fn drop(&mut self) {
        let _owner_scope = self.classes.as_ref().map(|classes| classes.owner.enter());
        // Descendants release all borrowed aliases before their ancestors.
        while !self.frames.is_empty() {
            let frame = self.pop();
            let termination = frame.termination;
            drop(frame);
            if let Some(id) = termination {
                oxvba_runtime::finish_pending_termination(id);
            }
        }
        for (id, _, _) in self.terminations.drain(..) {
            oxvba_runtime::finish_pending_termination(id);
        }
        self.stop_drain();
    }
}
