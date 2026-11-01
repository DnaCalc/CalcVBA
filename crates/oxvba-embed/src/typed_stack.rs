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
}
pub(crate) struct TypedStack {
    frames: Vec<Frame>,
    error: *mut ErrEngine,
    stores: Vec<(*mut oxvba_runtime::Variant, usize)>,
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
            }],
            error,
            stores,
        }
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
            let result = execute(
                frame.program,
                frame.procedure,
                &mut frame.activation,
                context,
                operations,
            );
            operations = frame.activation.raw().remaining_operations;
            // Ancestors remain parked while descendants hold their aliases.
            // Global alias reads retain scratch on the executing frame itself.
            self.frames.last_mut().ok_or(5)?.activation.clear_scratch();
            match result {
                Ok(SUSPENDED) => return Ok(SUSPENDED),
                Ok(0) if self.frames.len() == 1 => return Ok(0),
                Ok(0) => {
                    let mut callee = self.pop();
                    if let Some(dst) = callee.destination {
                        let value = callee.activation.take_result();
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
                    if let Err(number) = self.control(programs, entries) {
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
                            if target.target
                                == LinkedTarget::Builtin(oxvba_bundle::NativeImplId::IsMissing)
                            {
                                let frame = &mut self.frames.last_mut().ok_or(5)?.activation;
                                let operand = match &args[0] {
                                    oxvba_oxir::OxArg::ByVal(op) => op.clone(),
                                    oxvba_oxir::OxArg::ByRef(p) => OxOperand::Use(*p),
                                    _ => return Err(13),
                                };
                                let bits = frame.read(&operand, ScalarType::Variant)?;
                                // SAFETY: the executing frame retains this exact Variant cell.
                                let Value::Variant(value) = (unsafe {
                                    calcvba_host_api::value::retain_bits(ScalarType::Variant, bits)
                                })?
                                else {
                                    return Err(13);
                                };
                                let missing = value.vtype() == oxvba_runtime::VarType::Error
                                    && value.as_error_code() == Some(0x8002_0004u32 as i32);
                                if let Some(dst) = dst {
                                    frame.write_owned_value(
                                        *dst,
                                        Value::Variant(oxvba_runtime::Variant::from_bool(missing)),
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
                    });
                    return Ok(());
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
                        drop(self.pop());
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
        // Descendants release all borrowed aliases before their ancestors.
        while !self.frames.is_empty() {
            drop(self.pop());
        }
    }
}
