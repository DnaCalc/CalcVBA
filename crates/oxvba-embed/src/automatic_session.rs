//! Owning event entry and host document identity binding, shared by VM3 and JIT.
use super::*;
use calcvba_host_api::{self as api, Value};
fn number(error: Error) -> i32 {
    match error {
        Error::Runtime { number, .. } => number,
        Error::BudgetExceeded => -1,
        _ => 5,
    }
}
impl OwnedSession {
    pub(crate) fn ensure_event_dispatcher(&self) {
        self.cell.with_dependent(|_, session| {
            *session.automatic.session.borrow_mut() = Rc::downgrade(&self.cell);
        });
    }
    /// Bind a host document's identity. The compiler selects all event handlers;
    /// the host supplies no event name, DISPID, callback or handler registration.
    pub fn attach_document(&self, module: &str, source: &api::HostObject) -> Result<(), Error> {
        self.ensure_event_dispatcher();
        self.cell
            .with_dependent(|_, session| session.attach_document(module, source))
    }
}
impl Session<'_> {
    pub(crate) fn validate_reachable(&self, unit: usize, procedure: FuncId) -> Result<(), Error> {
        let mut pending = vec![(unit, procedure)];
        let mut visited = std::collections::HashSet::new();
        while let Some((unit, id)) = pending.pop() {
            if !visited.insert((unit, id)) {
                continue;
            }
            let entry = self
                .owning_typed
                .get(unit)
                .and_then(|entries| entries.get(id.0))
                .ok_or_else(|| Error::UnsupportedSignature("invalid owning callee".into()))?
                .as_ref()
                .map_err(|e| Error::UnsupportedSignature(e.clone()))?;
            pending.extend(entry.layout().calls.keys().map(|id| (unit, *id)));
            pending.extend(entry.layout().extern_calls.values().filter_map(
                |call| match call.target {
                    LinkedTarget::Procedure { program, procedure } => Some((program, procedure)),
                    LinkedTarget::Builtin(_) => None,
                },
            ));
            for inst in self.programs[unit].funcs[id.0]
                .blocks
                .iter()
                .flat_map(|block| &block.instrs)
            {
                if let oxvba_oxir::OxInst::WithEventsSet { binding, .. } = inst {
                    pending.extend(
                        self.programs[unit]
                            .event_routes
                            .iter()
                            .filter(|route| route.binding == *binding)
                            .map(|route| (unit, FuncId(route.handler))),
                    );
                }
            }
        }
        Ok(())
    }
    pub(crate) fn new_internal_stack(
        &self,
        unit: usize,
        procedure: FuncId,
        args: &[api::ValueSlot<'_>],
    ) -> Result<typed_stack::TypedStack, Error> {
        let _object_scope = self.object_owner.borrow().enter();
        let typed = self
            .owning_typed
            .get(unit)
            .and_then(|entries| entries.get(procedure.0))
            .ok_or(Error::ArgumentMismatch)?
            .as_ref()
            .map_err(|e| Error::UnsupportedSignature(e.clone()))?;
        let root = self.module.reference_programs.len();
        let (stores, error, classes) = match &self.executable {
            Executable::Interpreter(state) => {
                let mut state = state.try_borrow_mut().map_err(|_| Error::Reentry)?;
                (
                    (0..=root)
                        .map(|unit| state.typed_globals_raw_at(unit))
                        .collect::<Vec<_>>(),
                    state.typed_error_raw(),
                    state.typed_class_context(),
                )
            }
            #[cfg(feature = "jit")]
            Executable::Jit(_, state) => {
                let mut state = state.try_borrow_mut().map_err(|_| Error::Reentry)?;
                (
                    (0..=root)
                        .map(|unit| state.typed_globals_raw_at(unit))
                        .collect::<Vec<_>>(),
                    state.typed_error_raw(),
                    state.typed_class_context(),
                )
            }
        };
        // SAFETY: this session retains the code, Err, canonical stores and class
        // context; all are accessed only under the owning same-thread lease.
        let activation = unsafe {
            oxvba_rt_abi::host_entry::activation::Activation::new_values(
                typed.layout(),
                args,
                stores[unit].0,
                stores[unit].1,
                typed.entry_pc(),
            )
        }
        .map_err(|e| self.typed_error(e))?;
        // SAFETY: identical session execution lease; stack drops before the cell.
        Ok(unsafe {
            typed_stack::TypedStack::new(
                activation,
                unit,
                procedure,
                stores,
                error,
                classes,
                self.automatic.clone(),
            )
        })
    }
    pub(crate) fn call_event(
        &self,
        unit: usize,
        procedure: FuncId,
        owner: &oxvba_runtime::ObjectRef,
        args: &mut [Value],
        context: &mut api::TypedContext,
    ) -> Result<(), i32> {
        let _activation = self.enter_typed(context).map_err(number)?;
        self.validate_reachable(unit, procedure).map_err(number)?;
        let me = match self.owning_typed[unit][procedure.0]
            .as_ref()
            .map_err(|_| 5)?
            .layout()
            .parameters
            .first()
        {
            Some(api::ValueType::Variant) => {
                Value::Variant(Variant::from_object_ref(owner.clone()))
            }
            Some(api::ValueType::HostObject) => Value::RuntimeObject(owner.clone()),
            _ => return Err(13),
        };
        let mut inputs = vec![me.borrow_slot().ok_or(13)?];
        inputs.extend(
            args.iter()
                .map(|v| v.borrow_slot().ok_or(13))
                .collect::<Result<Vec<_>, _>>()?,
        );
        let mut stack = self
            .new_internal_stack(unit, procedure, &inputs)
            .map_err(number)?;
        stack.enable_writebacks(inputs.len()).map_err(number)?;
        drop(inputs);
        let result = self.execute_frame(&mut stack, context, u64::MAX);
        for (index, value) in stack.take_writebacks()? {
            if index == 0 || index > args.len() {
                return Err(13);
            }
            args[index - 1] = value;
        }
        if result.map_err(number)? != 0 {
            return Err(5);
        }
        Ok(())
    }
    pub(crate) fn attach_document(
        &self,
        module: &str,
        source: &api::HostObject,
    ) -> Result<(), Error> {
        if self.typed_depth.get() != 0 || self.suspended.get() != 0 {
            return Err(Error::Reentry);
        }
        let _executing = self
            .executing
            .try_borrow_mut()
            .map_err(|_| Error::Reentry)?;
        let _object_scope = self.object_owner.borrow().enter();
        let document = self
            .module
            .documents
            .iter()
            .find(|d| d.name.eq_ignore_ascii_case(module))
            .ok_or_else(|| Error::UnknownProcedure(module.into()))?;
        let unit = self.module.reference_programs.len();
        for route in self.programs[unit]
            .event_routes
            .iter()
            .filter(|route| route.document_class == Some(document.class))
        {
            self.validate_reachable(unit, FuncId(route.handler))?;
        }
        let iid = self
            .module
            .program
            .com_interfaces
            .iter()
            .find_map(|i| match i {
                oxvba_oxir::ComInterface::Com(m)
                    if m.name.eq_ignore_ascii_case(&document.host_type) =>
                {
                    m.iid.map(provider::interface_id)
                }
                _ => None,
            })
            .ok_or_else(|| {
                Error::UnsupportedSignature("document host type has no retained interface".into())
            })?;
        let source = source
            .resolve_interface(iid)
            .map_err(|_| Error::ArgumentMismatch)?;
        if self
            .automatic
            .documents
            .borrow()
            .get(&(unit, document.class))
            .is_some_and(|old| !old.same_identity(&source))
        {
            return Err(Error::ArgumentMismatch);
        }
        let classes = match &self.executable {
            Executable::Interpreter(state) => state.borrow_mut().typed_class_context(),
            #[cfg(feature = "jit")]
            Executable::Jit(_, state) => state.borrow_mut().typed_class_context(),
        }
        .ok_or_else(|| Error::Compile("missing document class".into()))?;
        // SAFETY: canonical singleton maps and allocator remain leased by the session.
        let owner =
            unsafe { classes.predeclared(unit, document.class) }.and_then(|v| v.as_object_ref());
        let owner = match owner {
            Some(owner) => owner,
            // SAFETY: allocate exactly the compiler's document class in this owner.
            None => unsafe { classes.allocate(unit, document.class, false) }
                .map_err(|n| self.typed_error(n))?,
        };
        owner.project_field_set(
            oxvba_bundle::HOST_DOCUMENT_FIELD,
            Variant::from_object_ref(source.object_ref()),
        );
        // SAFETY: same retained canonical singleton map; no borrow crosses subscription.
        unsafe {
            classes.set_predeclared(
                unit,
                document.class,
                Variant::from_object_ref(owner.clone()),
            )
        }
        .map_err(|n| self.typed_error(n))?;
        let binding = i32::MIN
            .checked_add(i32::try_from(document.class).map_err(|_| Error::ArgumentMismatch)?)
            .ok_or(Error::ArgumentMismatch)?;
        self.automatic
            .set(
                unit,
                owner,
                binding,
                Variant::from_object_ref(source.object_ref()),
                self.programs[unit],
                &self.owning_typed[unit],
            )
            .map_err(|n| self.typed_error(n))?;
        self.automatic
            .documents
            .borrow_mut()
            .insert((unit, document.class), source);
        Ok(())
    }
}
