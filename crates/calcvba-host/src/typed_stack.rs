//! Retained call/ByRef/error mechanics shared by the two typed backends.
//! Each backend still executes its original OxFunc arithmetic and host calls.
use calcvba_host_api::{ScalarBits, TypedContext, Value, ValueType as ScalarType};
use calcvba_oxir::{
    BlockId, ErrField, FuncId, OxConst, OxInst, OxOperand, OxPlace, OxProgram, OxTerminator,
};
use calcvba_rt_abi::host_entry::LinkedTarget;
use calcvba_rt_abi::host_entry::activation::{Activation, CONTROL, SUSPENDED};
use calcvba_rt_abi::{ErrEngine, Fault, FaultAction, ResumeTarget, SavedErrState};
use calcvba_vm3::typed::TypedInterpreter;

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
    stores: Vec<(*mut calcvba_runtime::Variant, usize)>,
    classes: Option<calcvba_rt_abi::TypedClassContext>,
    terminations: std::collections::VecDeque<(i32, i32, i32)>,
    owns_drain: bool,
    writeback_slots: Vec<usize>,
    writebacks: Vec<(usize, Value)>,
    automatic: std::rc::Rc<crate::events::AutomaticEvents>,
}
impl TypedStack {
    fn abort_event_fault(
        &mut self,
        programs: &[&OxProgram],
        context: &TypedContext,
    ) -> Result<(), i32> {
        if let Some(fault) = calcvba_host_api::take_unhandled_event_fault(context) {
            let number = fault.number;
            let unit = programs[self.frames.last().ok_or(5)?.program]
                .unit_name
                .clone();
            self.err().raise(
                Fault {
                    code: number,
                    message: fault.description,
                    source: fault.source,
                    help_file: fault.help_file,
                    help_context: fault.help_context,
                },
                unit,
            );
            return Err(number);
        }
        Ok(())
    }
    fn dynamic_field(
        &mut self,
        program: &OxProgram,
        inst: &OxInst,
        context: &mut TypedContext,
    ) -> Result<bool, i32> {
        let OxInst::ComCallLate {
            dst,
            recv,
            name,
            default_member: false,
            invoke_kind,
            args,
        } = inst
        else {
            return Ok(false);
        };
        let reference = self.operand(recv)?;
        let Some(object) = reference
            .as_object_ref()
            .filter(|o| o.is_project_instance())
        else {
            return Ok(false);
        };
        let current = self.frames.last().ok_or(5)?.program;
        if object.bundle_id() != current as i32
            || object.object_owner_id()
                != Some(calcvba_runtime::object_owner::RuntimeObjectOwner::current().id())
        {
            return Err(13);
        }
        let class = program.classes.get(object.route_key() as usize).ok_or(13)?;
        let Some(field) = class
            .fields
            .iter()
            .find(|field| field.externally_visible && field.name.eq_ignore_ascii_case(name))
        else {
            return Ok(false);
        };
        let carrier = match field.ty {
            calcvba_oxir::OxTy::Long => ScalarType::Long,
            calcvba_oxir::OxTy::Double => ScalarType::Double,
            calcvba_oxir::OxTy::Bool => ScalarType::Boolean,
            calcvba_oxir::OxTy::Str => ScalarType::String,
            calcvba_oxir::OxTy::Object(_) => ScalarType::HostObject,
            calcvba_oxir::OxTy::Variant => ScalarType::Variant,
            _ => return Err(13),
        };
        match invoke_kind {
            calcvba_com::TypeLibMemberInvokeKind::PropertyGet
            | calcvba_com::TypeLibMemberInvokeKind::Method
                if args.is_empty() =>
            {
                let value = object.project_field_get(field.token).unwrap_or_else(|| {
                    if carrier == ScalarType::HostObject {
                        calcvba_runtime::Variant::nothing()
                    } else {
                        calcvba_runtime::Variant::empty()
                    }
                });
                let value = calcvba_rt_abi::host_entry::value::to_variant(
                    calcvba_rt_abi::host_entry::value::from_variant(carrier, value)?,
                );
                if let Some(dst) = dst {
                    let frame = &mut self.frames.last_mut().ok_or(5)?.activation;
                    let ty = frame
                        .layout()
                        .operand_type(&OxOperand::Use(*dst))
                        .map_err(|_| 5)?;
                    frame.write_owned_value(
                        *dst,
                        calcvba_rt_abi::host_entry::value::from_variant(ty, value)?,
                    )?;
                }
            }
            calcvba_com::TypeLibMemberInvokeKind::PropertyPut
            | calcvba_com::TypeLibMemberInvokeKind::PropertyPutRef
                if args.len() == 1 =>
            {
                let mut value = match &args[0] {
                    calcvba_oxir::OxCallArg::Operand(op) => self.operand(op)?,
                    calcvba_oxir::OxCallArg::ByRef(p) => self.operand(&OxOperand::Use(*p))?,
                    _ => return Err(450),
                };
                if *invoke_kind == calcvba_com::TypeLibMemberInvokeKind::PropertyPut {
                    value = calcvba_rt_abi::host_entry::value::resolve_default(
                        value,
                        self.frames.last().ok_or(5)?.activation.layout(),
                        context,
                    )?;
                } else if !matches!(carrier, ScalarType::HostObject | ScalarType::Variant)
                    || value.vtype() != calcvba_runtime::VarType::Object
                {
                    return Err(13);
                }
                let value = calcvba_rt_abi::host_entry::value::to_variant(
                    calcvba_rt_abi::host_entry::value::from_variant(carrier, value)?,
                );
                if !object.project_field_set(field.token, value) {
                    return Err(438);
                }
            }
            _ => return Err(450),
        }
        let frame = &mut self.frames.last_mut().ok_or(5)?.activation;
        frame.set_pc(frame.pc() + 1);
        Ok(true)
    }
    fn invoke_builtin(
        &self,
        native: calcvba_bundle::NativeImplId,
        mut values: Vec<calcvba_runtime::Variant>,
        layout: &calcvba_rt_abi::host_entry::ScalarLayout,
        context: &mut TypedContext,
        host: &dyn calcvba_hal::HostServices,
    ) -> Result<calcvba_runtime::Variant, i32> {
        use calcvba_bundle::NativeImplId as N;
        if matches!(native, N::IsMissing | N::IsArray | N::VarType | N::CStr) {
            return calcvba_rt_abi::host_entry::value::invoke_unary_builtin(
                native,
                values.into_iter().next().ok_or(450)?,
                layout,
                context,
            );
        }
        // These functions inspect object identity/type. Resolving a default getter here
        // would inspect its result instead, or fail for objects without a default member.
        if !matches!(native, N::TypeName | N::IsObject) {
            for value in &mut values {
                *value = calcvba_rt_abi::host_entry::value::resolve_default(
                    value.clone(),
                    layout,
                    context,
                )?;
            }
        }
        let result = calcvba_host_api::with_native_call(context, || {
            calcvba_lib::invoke_context_free(native, &values, host)
        })
        .map_err(|error| match error {
            calcvba_lib::ContextFreeInvokeError::Library(error) => error.code,
            calcvba_lib::ContextFreeInvokeError::ContextRequired => 5,
        })?;
        if native == N::CreateObject {
            let object = result.as_object_ref().ok_or(13)?;
            let object = host.com().prepare_dynamic_object(object).map_err(|_| 429)?;
            return Ok(calcvba_runtime::Variant::from_object_ref(
                object.object_ref(),
            ));
        }
        Ok(result)
    }
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
                calcvba_rt_abi::default_error_message(number)
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
        stores: Vec<(*mut calcvba_runtime::Variant, usize)>,
        error: *mut ErrEngine,
        classes: Option<calcvba_rt_abi::TypedClassContext>,
        automatic: std::rc::Rc<crate::events::AutomaticEvents>,
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
            automatic,
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
                &OxOperand::Use(OxPlace::Local(calcvba_oxir::LocalId(index))),
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
        if let Some(id) = frame.termination {
            self.automatic.clear_owner_id(frame.program, id);
        }
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
    fn operand(&mut self, operand: &OxOperand) -> Result<calcvba_runtime::Variant, i32> {
        let frame = &mut self.frames.last_mut().ok_or(5)?.activation;
        let ty = frame.layout().operand_type(operand).map_err(|_| 5)?;
        let bits = frame.read(operand, ty)?;
        // SAFETY: the executing frame retains the exact admitted operand carrier.
        Ok(calcvba_rt_abi::host_entry::value::to_variant(unsafe {
            calcvba_host_api::value::retain_bits(ty, bits)
        }?))
    }
    fn project_object(&mut self, operand: &OxOperand) -> Result<calcvba_runtime::ObjectRef, i32> {
        let value = self.operand(operand)?;
        if value.vtype() != calcvba_runtime::VarType::Object {
            return Err(424);
        }
        let object = value.as_object_ref().ok_or(91)?;
        if !object.is_project_instance()
            || object.object_owner_id()
                != Some(calcvba_runtime::object_owner::RuntimeObjectOwner::current().id())
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
        object: &calcvba_runtime::ObjectRef,
    ) -> Result<(), i32> {
        if self.frames.len() >= 64 {
            return Err(28);
        }
        let typed = entries
            .get(unit)
            .and_then(|p| p.get(proc.0))
            .and_then(|e| e.as_ref().ok())
            .ok_or(5)?;
        let me = calcvba_runtime::Variant::from_object_ref(object.clone());
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
        self.automatic.collect();
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
                    .extend(calcvba_runtime::take_pending_terminations());
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
                let object = calcvba_runtime::retained_parked_termination_object(id).ok_or(5)?;
                if let Err(code) = self.push_me(entries, unit as usize, proc, &object) {
                    drop(object);
                    self.automatic.clear_owner_id(unit as usize, id);
                    calcvba_runtime::finish_pending_termination(id);
                    return Err(code);
                }
                self.frames.last_mut().ok_or(5)?.termination = Some(id);
                return Ok(true);
            }
            self.automatic.clear_owner_id(unit as usize, id);
            calcvba_runtime::finish_pending_termination(id);
        }

        Ok(false)
    }
    pub(crate) fn resume(
        &mut self,
        programs: &[&OxProgram],
        entries: &[Vec<Result<TypedInterpreter<'_>, String>>],
        context: &mut TypedContext,
        mut operations: u64,
        host: &dyn calcvba_hal::HostServices,
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
            self.abort_event_fault(programs, context)?;
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
                        calcvba_runtime::finish_pending_termination(id);
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
                            value = calcvba_rt_abi::host_entry::value::from_variant(
                                target,
                                calcvba_rt_abi::host_entry::value::to_variant(value),
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
                    let result = self.control(programs, entries, context, host);
                    self.abort_event_fault(programs, context)?;
                    if let Err(number) = result {
                        if std::env::var_os("OXVBA_TYPED_TRACE").is_some() {
                            eprintln!(
                                "control fault {number}: {:?}",
                                self.position(programs).ok().and_then(|(b, o)| self
                                    .frames
                                    .last()
                                    .map(|f| (
                                        programs[f.program].funcs[f.procedure.0].blocks[b]
                                            .instrs
                                            .get(o.saturating_sub(1)),
                                        b,
                                        o
                                    )))
                            );
                        }
                        let fault = calcvba_host_api::take_typed_fault(context, number)
                            .map_or_else(
                                || {
                                    Fault::new(
                                        number,
                                        calcvba_rt_abi::default_error_message(number),
                                    )
                                },
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
                }
                Err(-1) => return Err(-1),
                Err(number) => {
                    if std::env::var_os("OXVBA_TYPED_TRACE").is_some() {
                        eprintln!(
                            "value fault {number}: {:?}",
                            self.position(programs)
                                .ok()
                                .and_then(|(b, o)| self.frames.last().map(|f| (
                                    programs[f.program].funcs[f.procedure.0].blocks[b]
                                        .instrs
                                        .get(o.saturating_sub(1)),
                                    b,
                                    o
                                )))
                        );
                    }
                    let fault = calcvba_host_api::take_typed_fault(context, number).map_or_else(
                        || Fault::new(number, calcvba_rt_abi::default_error_message(number)),
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
        host: &dyn calcvba_hal::HostServices,
    ) -> Result<(), i32> {
        let (block_id, offset) = self.position(programs)?;
        let current = self.frames.last().ok_or(5)?.program;
        let program = programs[current];
        let procedure = self.frames.last().ok_or(5)?.procedure;
        let block = &program.funcs[procedure.0].blocks[block_id];
        if let Some(inst) = offset.checked_sub(1).and_then(|i| block.instrs.get(i)) {
            match inst {
                OxInst::NewExtern { dst, .. } => {
                    let object = calcvba_rt_abi::host_entry::collection::new_collection();
                    self.frames
                        .last_mut()
                        .ok_or(5)?
                        .activation
                        .write_owned_value(*dst, Value::RuntimeObject(object))?;
                }
                OxInst::CallNative {
                    dst,
                    callee: calcvba_oxir::OxNativeCallee::Builtin(id),
                    args,
                } => {
                    let mut values = Vec::with_capacity(args.len());
                    for arg in args {
                        values.push(match arg {
                            calcvba_oxir::OxCallArg::Operand(op) => self.operand(op)?,
                            calcvba_oxir::OxCallArg::ByRef(p) => {
                                self.operand(&OxOperand::Use(*p))?
                            }
                            calcvba_oxir::OxCallArg::Const(v) => {
                                calcvba_runtime::Variant::from_i32(*v)
                            }
                            calcvba_oxir::OxCallArg::Omitted => {
                                calcvba_runtime::Variant::from_error_code(0x80020004u32 as i32)
                            }
                            _ => return Err(13),
                        });
                    }
                    let value = self.invoke_builtin(
                        *id,
                        values,
                        entries[current][procedure.0]
                            .as_ref()
                            .map_err(|_| 5)?
                            .layout(),
                        context,
                        host,
                    )?;
                    let frame = &mut self.frames.last_mut().ok_or(5)?.activation;
                    if let Some(dst) = dst {
                        let ty = frame
                            .layout()
                            .operand_type(&OxOperand::Use(*dst))
                            .map_err(|_| 5)?;
                        frame.write_owned_value(
                            *dst,
                            calcvba_rt_abi::host_entry::value::from_variant(ty, value)?,
                        )?;
                    }
                    frame.set_pc(frame.pc() + 1);
                    return Ok(());
                }
                OxInst::Predeclared { dst, class } => {
                    let classes = self.classes.as_ref().ok_or(5)?;
                    // SAFETY: the canonical singleton map is leased by this stack.
                    let value = unsafe { classes.predeclared(current, class.0) }.ok_or(91)?;
                    self.frames
                        .last_mut()
                        .ok_or(5)?
                        .activation
                        .write_owned_value(
                            *dst,
                            Value::RuntimeObject(value.as_object_ref().ok_or(91)?),
                        )?;
                }
                OxInst::WithEventsGet {
                    dst,
                    owner,
                    binding,
                } => {
                    let owner = self.project_object(owner)?;
                    let value = self.automatic.get(current, &owner, *binding);
                    let frame = &mut self.frames.last_mut().ok_or(5)?.activation;
                    let ty = frame
                        .layout()
                        .operand_type(&OxOperand::Use(*dst))
                        .map_err(|_| 5)?;
                    frame.write_owned_value(
                        *dst,
                        calcvba_rt_abi::host_entry::value::from_variant(ty, value)?,
                    )?;
                }
                OxInst::WithEventsSet {
                    dst,
                    owner,
                    binding,
                    value,
                } => {
                    let owner = self.project_object(owner)?;
                    let value = self.operand(value)?;
                    self.automatic.set(
                        current,
                        owner,
                        *binding,
                        value.clone(),
                        program,
                        &entries[current],
                    )?;
                    let frame = &mut self.frames.last_mut().ok_or(5)?.activation;
                    let ty = frame
                        .layout()
                        .operand_type(&OxOperand::Use(*dst))
                        .map_err(|_| 5)?;
                    frame.write_owned_value(
                        *dst,
                        calcvba_rt_abi::host_entry::value::from_variant(ty, value)?,
                    )?;
                }
                OxInst::WithEventsClearOwner { dst, owner } => {
                    let owner = self.project_object(owner)?;
                    self.automatic.clear_owner(current, &owner);
                    let frame = &mut self.frames.last_mut().ok_or(5)?.activation;
                    let ty = frame
                        .layout()
                        .operand_type(&OxOperand::Use(*dst))
                        .map_err(|_| 5)?;
                    frame.write_owned_value(
                        *dst,
                        calcvba_rt_abi::host_entry::value::from_variant(
                            ty,
                            calcvba_runtime::Variant::from_i32(0),
                        )?,
                    )?;
                }
                OxInst::CallProc { args, dst, .. } | OxInst::CallExtern { args, dst, .. } => {
                    let (unit, proc) = match inst {
                        OxInst::CallProc { proc, .. } => (current, *proc),
                        OxInst::CallExtern { import, .. } => {
                            let parent = entries[current][procedure.0].as_ref().map_err(|_| 5)?;
                            let target = parent.layout().extern_calls.get(import).ok_or(5)?;
                            if let LinkedTarget::Builtin(native) = target.target {
                                let mut values = Vec::with_capacity(args.len());
                                for arg in args {
                                    values.push(match arg {
                                        calcvba_oxir::OxArg::ByVal(op) => self.operand(op)?,
                                        calcvba_oxir::OxArg::ByRef(p) => {
                                            self.operand(&OxOperand::Use(*p))?
                                        }
                                        calcvba_oxir::OxArg::Omitted => {
                                            calcvba_runtime::Variant::from_error_code(
                                                0x80020004u32 as i32,
                                            )
                                        }
                                    });
                                }
                                if let calcvba_bundle::ExportToken::ModuleFunc { member, .. } =
                                    &program.imports[import.0].token
                                    && native.is_string_typed_library_alias(member)
                                    && values
                                        .iter()
                                        .any(|v| v.vtype() == calcvba_runtime::VarType::Null)
                                {
                                    return Err(94);
                                }
                                let result = self.invoke_builtin(
                                    native,
                                    values,
                                    entries[current][procedure.0]
                                        .as_ref()
                                        .map_err(|_| 5)?
                                        .layout(),
                                    context,
                                    host,
                                )?;
                                let frame = &mut self.frames.last_mut().ok_or(5)?.activation;
                                if let Some(dst) = dst {
                                    let ty = frame
                                        .layout()
                                        .operand_type(&OxOperand::Use(*dst))
                                        .map_err(|_| 5)?;
                                    frame.write_owned_value(
                                        *dst,
                                        calcvba_rt_abi::host_entry::value::from_variant(
                                            ty, result,
                                        )?,
                                    )?;
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
                    if self.dynamic_field(program, inst, context)? {
                        return Ok(());
                    }
                    let reference = self.operand(recv)?;
                    let object = reference.as_object_ref();
                    let route = object
                        .as_ref()
                        .filter(|o| o.is_project_instance())
                        .and_then(|object| {
                            self.frames
                                .last()?
                                .activation
                                .layout()
                                .class_calls
                                .iter()
                                .find(|(site, _, class)| {
                                    site == inst && class.0 as i32 == object.route_key()
                                })
                                .map(|(_, p, c)| (*p, *c))
                        });
                    if route.is_none() {
                        let frame = &mut self.frames.last_mut().ok_or(5)?.activation;
                        // SAFETY: this is the original admitted instruction and its
                        // uniquely executing retained frame/context. Native dispatch
                        // executes once, through the same shared member helper.
                        let status = unsafe {
                            calcvba_rt_abi::host_entry::value::evaluate(
                                frame.raw_mut(),
                                inst,
                                context,
                            )
                        };
                        if status != 0 {
                            return Err(status);
                        }
                        frame.set_pc(frame.pc() + 1);
                        return Ok(());
                    }
                    let object = self.project_object(recv)?;
                    let (proc, class) = route.ok_or(438)?;
                    if object.bundle_id() != current as i32 || object.route_key() != class.0 as i32
                    {
                        return Err(13);
                    }
                    let typed = entries[current][proc.0].as_ref().map_err(|_| 5)?;
                    let mut operands = Vec::with_capacity(args.len() + 1);
                    operands.push(calcvba_oxir::OxArg::ByVal(recv.clone()));
                    for (index, arg) in args.iter().enumerate() {
                        operands.push(match arg {
                            calcvba_oxir::OxCallArg::Operand(op) => {
                                calcvba_oxir::OxArg::ByVal(op.clone())
                            }
                            calcvba_oxir::OxCallArg::ByRef(place)
                                if typed.layout().by_ref_parameters.get(index + 1)
                                    == Some(&true) =>
                            {
                                calcvba_oxir::OxArg::ByRef(*place)
                            }
                            calcvba_oxir::OxCallArg::ByRef(place) => {
                                calcvba_oxir::OxArg::ByVal(OxOperand::Use(*place))
                            }
                            calcvba_oxir::OxCallArg::Omitted => calcvba_oxir::OxArg::Omitted,
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
                    let mut value = object
                        .project_field_get(*field)
                        .unwrap_or_else(calcvba_runtime::Variant::empty);
                    if value.vtype() == calcvba_runtime::VarType::Empty {
                        value = match declaration.ty {
                            calcvba_oxir::OxTy::Long => calcvba_runtime::Variant::from_i32(0),
                            calcvba_oxir::OxTy::Double => calcvba_runtime::Variant::from_f64(0.0),
                            calcvba_oxir::OxTy::Bool => calcvba_runtime::Variant::from_bool(false),
                            calcvba_oxir::OxTy::Str => calcvba_runtime::Variant::from_string(""),
                            calcvba_oxir::OxTy::Object(_) => calcvba_runtime::Variant::nothing(),
                            calcvba_oxir::OxTy::Variant => value,
                            calcvba_oxir::OxTy::Array(_, _) => {
                                calcvba_runtime::Variant::unallocated_array(
                                    calcvba_bundle::safearray_vartype_for_element(
                                        declaration.array_element.as_ref().ok_or(13)?,
                                    ),
                                )
                            }
                            calcvba_oxir::OxTy::Record(id) => {
                                let fields = programs[object.bundle_id() as usize]
                                    .record_layouts
                                    .get(id.0)
                                    .ok_or(13)?;
                                let layout = calcvba_bundle::vba_record_layout_for_fields(fields)
                                    .map_err(|_| 13)?;
                                calcvba_runtime::Variant::from_vba_record(
                                    calcvba_runtime::vba_record::VbaRecord::new_default(layout)
                                        .map_err(|_| 7)?,
                                )
                            }
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
                        calcvba_rt_abi::host_entry::value::from_variant(ty, value)?,
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
                OxInst::FieldArrayGet {
                    dst,
                    object,
                    field,
                    indices,
                } => {
                    let object = self.project_object(object)?;
                    let array = object
                        .project_field_get(*field)
                        .unwrap_or_else(calcvba_runtime::Variant::empty);
                    let (bounds, _) = array.safearray_bounds_len().ok_or(9)?;
                    let indices = indices
                        .iter()
                        .map(|op| calcvba_eval::arith::int(&self.operand(op)?).map_err(|e| e.code))
                        .collect::<Result<Vec<_>, _>>()?;
                    let index = calcvba_rt_abi::host_entry::array_flat_index(&bounds, &indices)?;
                    let value = array.safearray_element(index).ok_or(9)?.map_err(|_| 9)?;
                    let frame = &mut self.frames.last_mut().ok_or(5)?.activation;
                    let ty = frame
                        .layout()
                        .operand_type(&OxOperand::Use(*dst))
                        .map_err(|_| 5)?;
                    frame.write_owned_value(
                        *dst,
                        calcvba_rt_abi::host_entry::value::from_variant(ty, value)?,
                    )?;
                }
                OxInst::FieldArraySet {
                    object,
                    field,
                    indices,
                    value,
                } => {
                    let object = self.project_object(object)?;
                    let mut array = object
                        .project_field_get(*field)
                        .unwrap_or_else(calcvba_runtime::Variant::empty);
                    let (bounds, _) = array.safearray_bounds_len().ok_or(9)?;
                    let indices = indices
                        .iter()
                        .map(|op| calcvba_eval::arith::int(&self.operand(op)?).map_err(|e| e.code))
                        .collect::<Result<Vec<_>, _>>()?;
                    let index = calcvba_rt_abi::host_entry::array_flat_index(&bounds, &indices)?;
                    array
                        .set_safearray_element(index, &self.operand(value)?)
                        .map_err(|_| 13)?;
                    if !object.project_field_set(*field, array) {
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
                        if let Some((_, Value::RuntimeObject(object))) = &frame.initialized {
                            object.mark_initialization_failed();
                        }
                        if let Some(id) = frame.termination {
                            drop(frame);
                            calcvba_runtime::finish_pending_termination(id);
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
                        calcvba_rt_abi::default_error_message(number)
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
                calcvba_runtime::finish_pending_termination(id);
            }
        }
        for (id, unit, _) in self.terminations.drain(..) {
            self.automatic.clear_owner_id(unit as usize, id);
            calcvba_runtime::finish_pending_termination(id);
        }
        self.stop_drain();
        self.automatic.collect();
    }
}
