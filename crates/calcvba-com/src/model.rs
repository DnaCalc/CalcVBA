use calcvba_runtime::ObjectRefExt;
pub const DISPATCH_INVOKE_MISSING_ARG_TOKEN: i32 = i32::MIN + 2_048;

use calcvba_runtime::{ObjectRef, RuntimeByRefSlot, Variant};

macro_rules! define_token {
    ($name:ident) => {
        #[repr(transparent)]
        #[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
        pub struct $name(i32);

        impl $name {
            pub const fn new(raw: i32) -> Self {
                Self(raw)
            }

            pub const fn raw(self) -> i32 {
                self.0
            }
        }

        impl From<i32> for $name {
            fn from(value: i32) -> Self {
                Self::new(value)
            }
        }

        impl From<$name> for i32 {
            fn from(value: $name) -> Self {
                value.raw()
            }
        }

        impl core::fmt::Display for $name {
            fn fmt(&self, f: &mut core::fmt::Formatter<'_>) -> core::fmt::Result {
                write!(f, "{}", self.raw())
            }
        }
    };
}

define_token!(ComObjectToken);
define_token!(ComMemberToken);
define_token!(ComSubscriptionToken);
define_token!(ComCallbackToken);

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ComObjectTransportKind {
    Projection,
    NativeDispatch,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ComObjectDescriptor {
    pub object: ObjectRef,
    pub prog_id_name: String,
    pub transport: ComObjectTransportKind,
    pub supports_events: bool,
    pub known_member_tokens: Vec<ComMemberToken>,
    pub known_event_tokens: Vec<ComMemberToken>,
    pub default_member_token: Option<ComMemberToken>,
    pub default_member_name: Option<String>,
    pub typelib_cache_key: Option<String>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ComInvokeKind {
    Method,
    PropertyGet,
    PropertyPut,
    PropertyPutRef,
}

/// An argument value for a COM invocation: the runtime cell itself.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ComInvokeValue {
    value: Variant,
}

impl ComInvokeValue {
    pub fn from_variant(value: Variant) -> Self {
        Self { value }
    }

    /// Compatibility projection into scalar i32 transport.
    pub fn to_runtime_token(&self) -> Result<i32, String> {
        crate::compat::variant_to_runtime_token(&self.value)
    }

    /// Compatibility projection into legacy dispatch-token transport.
    pub fn to_legacy_dispatch_token(&self) -> Result<i32, String> {
        crate::compat::variant_to_legacy_dispatch_token(&self.value)
    }

    /// The argument's cell.
    pub fn variant(&self) -> &Variant {
        &self.value
    }

    pub fn into_variant(self) -> Variant {
        self.value
    }
}

impl From<Variant> for ComInvokeValue {
    fn from(value: Variant) -> Self {
        Self::from_variant(value)
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ComInvokeArg {
    pub value: Option<ComInvokeValue>,
    pub name: Option<String>,
    pub by_ref: Option<RuntimeByRefSlot>,
}

impl ComInvokeArg {
    /// Compatibility positional argument from scalar i32 transport.
    ///
    /// New value-model call sites should prefer [`Self::positional_value`].
    pub fn positional(value: i32) -> Self {
        Self {
            value: Some(Variant::from_i32(value).into()),
            name: None,
            by_ref: None,
        }
    }

    /// Compatibility named argument from scalar i32 transport.
    ///
    /// New value-model call sites should prefer [`Self::named_value`].
    pub fn named(value: i32, name: impl Into<String>) -> Self {
        Self {
            value: Some(Variant::from_i32(value).into()),
            name: Some(name.into()),
            by_ref: None,
        }
    }

    /// Retained value-model positional argument.
    pub fn positional_value(value: Variant) -> Self {
        Self {
            value: Some(value.into()),
            name: None,
            by_ref: None,
        }
    }

    /// Retained value-model positional ByRef argument.
    pub fn positional_by_ref(value: Variant, slot: RuntimeByRefSlot) -> Self {
        Self {
            value: Some(value.into()),
            name: None,
            by_ref: Some(slot),
        }
    }

    /// Retained value-model named argument.
    pub fn named_value(value: Variant, name: impl Into<String>) -> Self {
        Self {
            value: Some(value.into()),
            name: Some(name.into()),
            by_ref: None,
        }
    }

    pub fn omitted() -> Self {
        Self {
            value: None,
            name: None,
            by_ref: None,
        }
    }

    pub fn omitted_named(name: impl Into<String>) -> Self {
        Self {
            value: None,
            name: Some(name.into()),
            by_ref: None,
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ComInvokeRequest {
    pub object: ObjectRef,
    pub member: ComMemberToken,
    pub args: Vec<ComInvokeArg>,
    pub invoke_kind_hint: Option<ComInvokeKind>,
}

impl ComInvokeRequest {
    pub fn new(object: ObjectRef, member: ComMemberToken, args: Vec<ComInvokeArg>) -> Self {
        Self {
            object,
            member,
            args,
            invoke_kind_hint: None,
        }
    }

    /// Compatibility request constructor from legacy object/member/argument
    /// tokens.
    pub fn legacy(object: i32, member: i32, arg: i32) -> Self {
        let args = if arg == DISPATCH_INVOKE_MISSING_ARG_TOKEN {
            Vec::new()
        } else {
            vec![ComInvokeArg::positional(arg)]
        };
        Self::new(
            ObjectRef::from_compat_identity(object),
            ComMemberToken::new(member),
            args,
        )
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ComCallbackValue {
    value: Variant,
}

impl ComCallbackValue {
    pub fn from_variant(value: Variant) -> Self {
        Self { value }
    }

    /// The callback argument's cell.
    pub fn variant(&self) -> &Variant {
        &self.value
    }

    pub fn into_variant(self) -> Variant {
        self.value
    }
}

impl From<Variant> for ComCallbackValue {
    fn from(value: Variant) -> Self {
        Self::from_variant(value)
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ComCallbackPayload {
    pub callback: ComCallbackToken,
    pub subscription: ComSubscriptionToken,
    pub object: ObjectRef,
    pub event: ComMemberToken,
    pub args: Vec<ComCallbackValue>,
}

#[cfg(test)]
mod tests {
    use super::{ComCallbackValue, ComInvokeArg, ComInvokeValue, ComMemberToken};
    use calcvba_runtime::{Decimal96, VarType, Variant, bstr::BStr, safe_array::SafeArray};

    #[test]
    fn invoke_value_array_projects_to_length_token() {
        let value = ComInvokeValue::from_variant(Variant::from_safearray(SafeArray::vector(4)));
        assert_eq!(value.to_runtime_token().expect("array token"), 4);
        assert_eq!(
            value
                .to_legacy_dispatch_token()
                .expect("legacy dispatch token"),
            20_004
        );
    }

    #[test]
    fn invoke_value_token_projection() {
        let token = |value: Variant| ComInvokeValue::from_variant(value).to_runtime_token();
        assert_eq!(token(Variant::from_i16(-3)), Ok(-3));
        assert_eq!(token(Variant::from_i64(42)), Ok(42));
        assert!(token(Variant::from_i64(5_000_000_000)).is_err());
        assert!(token(Variant::from_string(BStr::from("ABC"))).is_err());
        assert_eq!(token(Variant::from_string(BStr::from(" 12 "))), Ok(12));
        assert_eq!(
            token(Variant::from_decimal96(Decimal96::from_parts(
                123, 0, 0, 0, false
            ))),
            Ok(123)
        );
        assert_eq!(
            token(Variant::from_safearray(SafeArray::from_variants(vec![
                Variant::from_i32(4),
                Variant::from_string(BStr::from("A")),
            ]))),
            Ok(2)
        );
    }

    #[test]
    fn com_invoke_arg_keeps_the_cell() {
        let arg = ComInvokeArg::positional_value(Variant::from_i16(7));
        let value = arg.value.expect("invoke arg should retain a value");
        assert_eq!(value.variant().vtype(), VarType::Integer);
        assert_eq!(value.variant().as_i16(), Some(7));
    }

    #[test]
    fn com_callback_value_keeps_the_cell() {
        let value = ComCallbackValue::from_variant(Variant::from_f32(1.5));
        assert_eq!(value.variant().vtype(), VarType::Single);
    }

    #[test]
    fn com_member_token_roundtrips_raw_dispid() {
        let token = ComMemberToken::new(11);
        assert_eq!(token.raw(), 11);
        assert_eq!(i32::from(token), 11);
    }
}
