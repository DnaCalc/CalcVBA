use super::*;
use calcvba_host_api::{self as api, InvocationProgress, ValueProgress};
use oxvba_rt_abi::host_entry::activation::SUSPENDED;

impl OwnedSession {
    pub fn begin_invocation(
        &self,
        bound: &BoundProcedure,
        args: &[api::Slot<'_>],
        frame: api::InvocationFrame,
        max_blocks: u64,
    ) -> Result<Box<dyn api::Invocation>, Error> {
        let activation = self.cell.with_dependent(|_, session| {
            session.validate(bound)?;
            let result = session.owning_typed[bound.id.0]
                .as_ref()
                .map_err(|e| Error::UnsupportedSignature(e.clone()))?
                .layout()
                .result;
            if !matches!(
                result.legacy(),
                Some(ScalarType::Long | ScalarType::Double | ScalarType::HostObject)
            ) {
                return Err(Error::UnsupportedSignature(
                    "this invocation result requires begin_values".into(),
                ));
            }
            if session.suspended.get() != 0 {
                return Err(Error::Reentry);
            }
            let _exclusive = session
                .executing
                .try_borrow_mut()
                .map_err(|_| Error::Reentry)?;
            let activation = session.new_activation(bound, args)?;
            session.suspended.set(1);
            Ok(activation)
        })?;
        Ok(Box::new(Continuation {
            cell: Some(self.cell.clone()),
            bound: bound.clone(),
            activation: Some(activation),
            frame: Rc::new(frame),
            context: TypedContext {
                remaining_blocks: max_blocks,
            },
        }))
    }
}
struct Continuation {
    // Activation drops before the cell, even if the host cancels by dropping it.
    cell: Option<Rc<cache::SessionCell>>,
    bound: BoundProcedure,
    activation: Option<typed_stack::TypedStack>,
    frame: Rc<api::InvocationFrame>,
    context: TypedContext,
}
impl Continuation {
    fn finish(&mut self) {
        drop(self.activation.take());
        if let Some(cell) = self.cell.take() {
            cell.with_dependent(|_, session| session.suspended.set(0));
        }
    }
}
impl Drop for Continuation {
    fn drop(&mut self) {
        self.finish();
    }
}
impl Continuation {
    fn frame(&self) -> &api::InvocationFrame {
        &self.frame
    }
    fn remaining_blocks(&self) -> u64 {
        self.context.remaining_blocks
    }
    fn resume_value(&mut self, operations: u64) -> Result<ValueProgress, api::Error> {
        let cell = self
            .cell
            .as_ref()
            .ok_or(api::Error::InvocationFinished)?
            .clone();
        let _view = api::InvocationScope::enter(self.frame.clone());
        let result = cell.with_dependent(|_, session| {
            session.resume_frame(
                &self.bound,
                self.activation.as_mut().expect("live continuation"),
                &mut self.context,
                operations,
                true,
            )
        });
        match result {
            Ok(SUSPENDED) => Ok(ValueProgress::Suspended),
            Ok(0) => {
                let value = self
                    .activation
                    .as_mut()
                    .expect("live continuation")
                    .result();
                self.finish();
                Ok(ValueProgress::Returned(value))
            }
            Err(error) => {
                self.finish();
                Err(provider::error(error))
            }
            Ok(_) => {
                self.finish();
                Err(api::Error::Backend("invalid continuation status".into()))
            }
        }
    }
}

impl api::Invocation for Continuation {
    fn frame(&self) -> &api::InvocationFrame {
        self.frame()
    }
    fn remaining_blocks(&self) -> u64 {
        self.remaining_blocks()
    }
    fn resume(&mut self, operations: u64) -> Result<InvocationProgress, api::Error> {
        match self.resume_value(operations)? {
            ValueProgress::Suspended => Ok(InvocationProgress::Suspended),
            ValueProgress::Returned(value) => {
                Ok(InvocationProgress::Returned(value.into_legacy()?))
            }
        }
    }
}
impl api::ValueInvocation for Continuation {
    fn frame(&self) -> &api::InvocationFrame {
        self.frame()
    }
    fn remaining_blocks(&self) -> u64 {
        self.remaining_blocks()
    }
    fn resume(&mut self, operations: u64) -> Result<ValueProgress, api::Error> {
        self.resume_value(operations)
    }
}
impl OwnedSession {
    pub fn begin_values(
        &self,
        bound: &BoundProcedure,
        args: &[api::ValueSlot<'_>],
        frame: api::InvocationFrame,
        max_blocks: u64,
    ) -> Result<Box<dyn api::ValueInvocation>, Error> {
        let activation = self.cell.with_dependent(|_, session| {
            session.validate(bound)?;
            if session.suspended.get() != 0 {
                return Err(Error::Reentry);
            }
            let _exclusive = session
                .executing
                .try_borrow_mut()
                .map_err(|_| Error::Reentry)?;
            let activation = session.new_value_activation(bound, args)?;
            session.suspended.set(1);
            Ok(activation)
        })?;
        Ok(Box::new(Continuation {
            cell: Some(self.cell.clone()),
            bound: bound.clone(),
            activation: Some(activation),
            frame: Rc::new(frame),
            context: TypedContext {
                remaining_blocks: max_blocks,
            },
        }))
    }
    pub fn call_values(
        &self,
        bound: &BoundProcedure,
        args: &[api::ValueSlot<'_>],
        context: &mut TypedContext,
    ) -> Result<api::Value, Error> {
        self.cell
            .with_dependent(|_, session| session.call_values(bound, args, context))
    }
    pub fn bind_values(
        &self,
        bound: &BoundProcedure,
        arguments: &[api::ValueType],
        result: api::ValueType,
    ) -> Result<(), Error> {
        self.cell
            .with_dependent(|_, session| session.bind_values(bound, arguments, result))
    }
}
