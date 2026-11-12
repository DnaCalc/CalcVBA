//! Scoped authorization at native COM boundaries, including compiler type-library loads.
use std::{cell::RefCell, rc::Rc};

type Authorize = dyn Fn(&'static str, &str) -> Result<(), String>;
thread_local! {
    static SCOPES: RefCell<Vec<Rc<Authorize>>> = const { RefCell::new(Vec::new()) };
}
/// Same-thread, nested policy scope. Outer restrictions also apply during reentry.
/// Dropping the scope restores the previous policy, including while unwinding.
#[must_use]
pub struct ComPolicyScope {
    authorize: Option<Rc<Authorize>>,
    parent: Option<Box<ComPolicyScope>>,
}
impl ComPolicyScope {
    pub fn unrestricted() -> Self {
        Self {
            authorize: None,
            parent: None,
        }
    }
    pub fn enter(authorize: impl Fn(&'static str, &str) -> Result<(), String> + 'static) -> Self {
        let authorize: Rc<Authorize> = Rc::new(authorize);
        SCOPES.with(|scopes| scopes.borrow_mut().push(authorize.clone()));
        Self {
            authorize: Some(authorize),
            parent: None,
        }
    }
}
impl ComPolicyScope {
    pub fn with_parent(mut self, parent: Self) -> Self {
        self.parent = Some(Box::new(parent));
        self
    }
}
impl Drop for ComPolicyScope {
    fn drop(&mut self) {
        if let Some(authorize) = &self.authorize {
            SCOPES.with(|scopes| {
                scopes
                    .borrow_mut()
                    .retain(|entry| !Rc::ptr_eq(entry, authorize))
            });
        }
    }
}
/// Native calls use this before any registry, file, QI, member or enumerator access.
pub fn authorize_native_com(op: &'static str, detail: &str) -> Result<(), String> {
    let scopes = SCOPES.with(|scopes| scopes.borrow().clone());
    for authorize in scopes {
        authorize(op, detail)?;
    }
    Ok(())
}
