//! The shape of a call on a COM object, shared by every platform: the HAL's object calls take
//! these, and on Windows `crate::automation` performs them over the runtime cell (step 8 of the
//! VARIANT-layout programme).
use calcvba_runtime::Variant;

/// Which member a call names.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ComMember<'a> {
    /// Late bound: `GetIDsOfNames` on each call, as VBA does.
    Name(&'a str),
    /// A known DISPID: `DISPID_VALUE` (0) for the default member, or one the compiler bound.
    Id(i32),
}

/// What a call does with the member.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ComCallKind {
    /// A call statement or a call in an expression; a property read is the same `Invoke`.
    Method,
    /// `x = obj.Member`.
    Get,
    /// `obj.Member = value` (the value is the last argument).
    Let,
    /// `Set obj.Member = object` (the object is the last argument).
    Set,
}

/// One argument cell.
#[derive(Debug, Clone, Copy)]
pub enum ComArgValue<'a> {
    /// A ByVal argument: the cell is borrowed; the callee reads it and must not clear it.
    Value(&'a Variant),
    /// A ByRef argument: a `VT_BYREF` cell referring to the caller's storage (step 7).
    Reference(&'a Variant),
    /// An omitted optional argument (`VT_ERROR` with `DISP_E_PARAMNOTFOUND`).
    Missing,
}

/// One argument, positional or named.
#[derive(Debug, Clone, Copy)]
pub struct ComArg<'a> {
    pub value: ComArgValue<'a>,
    pub name: Option<&'a str>,
}

impl<'a> ComArg<'a> {
    pub fn value(value: &'a Variant) -> Self {
        Self {
            value: ComArgValue::Value(value),
            name: None,
        }
    }

    pub fn reference(reference: &'a Variant) -> Self {
        Self {
            value: ComArgValue::Reference(reference),
            name: None,
        }
    }

    pub fn missing() -> Self {
        Self {
            value: ComArgValue::Missing,
            name: None,
        }
    }
}

/// A failed call, as VBA's `Err` sees it.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ComFault {
    pub number: i32,
    pub description: Option<String>,
    pub source: Option<String>,
    pub help_file: Option<String>,
    pub help_context: Option<u32>,
    pub hresult: Option<i32>,
}

impl ComFault {
    /// A VBA error with its default message.
    pub fn number(number: i32) -> Self {
        Self {
            number,
            description: None,
            source: None,
            help_file: None,
            help_context: None,
            hresult: None,
        }
    }
}

/// A `For Each` over an object, one item at a time.
pub trait ComEnumerator {
    /// The next item, or `None` at the end.
    fn next_item(&mut self) -> Result<Option<Variant>, ComFault>;
}
