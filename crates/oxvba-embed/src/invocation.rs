use super::*;
use calcvba_host_api::{self as api, InvocationProgress};
use oxvba_rt_abi::host_entry::activation::{Activation, SUSPENDED};

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
    activation: Option<Activation<'static>>,
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
impl api::Invocation for Continuation {
    fn frame(&self) -> &api::InvocationFrame {
        &self.frame
    }
    fn remaining_blocks(&self) -> u64 {
        self.context.remaining_blocks
    }
    fn resume(&mut self, operations: u64) -> Result<InvocationProgress, api::Error> {
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
            Ok(SUSPENDED) => Ok(InvocationProgress::Suspended),
            Ok(0) => {
                let value = self
                    .activation
                    .as_ref()
                    .expect("live continuation")
                    .result();
                self.finish();
                Ok(InvocationProgress::Returned(value))
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
