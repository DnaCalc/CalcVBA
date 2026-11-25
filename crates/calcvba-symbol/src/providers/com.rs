//! COM typelib provider — resolves members and events of a resolved
//! `TypeLibMetadataBlob` against a typed receiver. It returns one fully-typed
//! `ComMember` route (dispid + vtable slot + name) regardless of receiver, so the
//! binder picks early vs late by the receiver's static type.

use calcvba_bundle::ProjectMemberKind;
use calcvba_com::{TypeLibMemberInvokeKind, TypeLibMemberMetadata, TypeLibMetadataBlob};

use crate::binding::{Binding, DispatchRoute, member_kind_from_invoke};
use crate::model::fold_identifier;
use crate::provider::Provider;
use crate::signature::VarTypeRef;

pub struct ComTypeLibProvider {
    blob: TypeLibMetadataBlob,
    /// Folded names this blob answers to (its reference name, requested coclass,
    /// activation ProgID, and library-level coclasses).
    type_names: Vec<String>,
    /// Folded names whose members are scoped by this blob's member list. A
    /// library-wide blob knows every coclass type name, but its flat member list
    /// must not make one coclass see another coclass's members.
    member_type_names: Vec<String>,
    interface_iid: Option<calcvba_com::ComInterfaceIid>,
}

/// The `DISPID_EVALUATE` member that evaluates a foreign name `[text]` (Excel's `_Evaluate`).
const DISPID_EVALUATE: i32 = -5;
impl ComTypeLibProvider {
    fn resolve_constant(&self, parts: &[&str]) -> Option<Binding> {
        use calcvba_bundle::coreir::CoreConst as C;
        use calcvba_com::ComDefaultValue as V;
        let name = *parts.last()?;
        let candidates = self
            .blob
            .constants
            .iter()
            .filter(|constant| {
                constant.name.eq_ignore_ascii_case(name)
                    && match parts {
                        [_] => true,
                        [owner, _] => {
                            owner.eq_ignore_ascii_case(&self.blob.identity.reference_name)
                                || owner.eq_ignore_ascii_case(&constant.owner)
                        }
                        [library, owner, _] => {
                            library.eq_ignore_ascii_case(&self.blob.identity.reference_name)
                                && owner.eq_ignore_ascii_case(&constant.owner)
                        }
                        _ => false,
                    }
            })
            .collect::<Vec<_>>();
        let [constant] = candidates.as_slice() else {
            return None;
        };
        let value = match &constant.value {
            V::Empty => C::Empty,
            V::Null => C::Null,
            V::Bool(value) => C::Bool(*value),
            // Enum constants have their declared Long type, even for small values.
            V::Int(value) => i32::try_from(*value).map(C::I32).unwrap_or(C::I64(*value)),
            V::FloatBits(value) => C::F64(*value),
            V::CurrencyScaled(value) => C::Currency(*value),
            V::DateBits(value) => C::Date(*value),
            V::Str(value) => C::Str(value.clone()),
            V::ErrorCode(_) => return None,
        };
        Some(Binding::new(None, DispatchRoute::ConstValue(value)))
    }
    /// Folded names a declaration may use for this library's enum types (`xlcalculation`,
    /// `excel.xlcalculation`). Such a type is a Long, as in VBA.
    pub(crate) fn enum_type_aliases(&self) -> Vec<String> {
        let reference = fold_identifier(&self.blob.identity.reference_name);
        self.blob
            .enum_type_names
            .iter()
            .flat_map(|name| {
                let folded = fold_identifier(name);
                [format!("{reference}.{folded}"), folded]
            })
            .collect()
    }
    pub fn new(blob: TypeLibMetadataBlob) -> Self {
        let iid = blob
            .members
            .first()
            .and_then(|m| m.interface_iid)
            .filter(|iid| blob.members.iter().all(|m| m.interface_iid == Some(*iid)));
        Self::with_interface_iid(blob, iid)
    }
    pub fn with_interface_iid(
        mut blob: TypeLibMetadataBlob,
        interface_iid: Option<calcvba_com::ComInterfaceIid>,
    ) -> Self {
        // An application object's instance is the host's (Excel hands its `Application` in for
        // `Global`), and only the members' DISPIDs are its contract: its interface need not
        // answer `QueryInterface`. Its members are therefore called through `IDispatch` by
        // DISPID, never through a vtable.
        let app_object = blob
            .identity
            .requested_coclass
            .as_deref()
            .is_some_and(|coclass| {
                blob.app_objects
                    .iter()
                    .any(|object| object.coclass.eq_ignore_ascii_case(coclass))
            });
        if app_object {
            for member in &mut blob.members {
                member.is_dual = false;
            }
        }
        let reference = fold_identifier(&blob.identity.reference_name);
        let mut type_names = vec![reference.clone()];
        let mut member_type_names = vec![reference.clone()];
        if let Some(coclass) = &blob.identity.requested_coclass {
            let folded = fold_identifier(coclass);
            type_names.push(folded.clone());
            type_names.push(format!("{reference}.{folded}"));
            member_type_names.push(folded.clone());
            member_type_names.push(format!("{reference}.{folded}"));
        }
        if let Some(prog_id) = &blob.activation_prog_id {
            let folded = fold_identifier(prog_id);
            type_names.push(folded.clone());
            member_type_names.push(folded);
        }
        for coclass in &blob.coclass_names {
            let folded = fold_identifier(coclass);
            type_names.push(folded.clone());
            type_names.push(format!("{reference}.{folded}"));
        }
        type_names.sort();
        type_names.dedup();
        member_type_names.sort();
        member_type_names.dedup();
        Self {
            blob,
            type_names,
            member_type_names,
            interface_iid,
        }
    }

    fn owns(&self, type_name: &str) -> bool {
        let folded = fold_identifier(type_name);
        self.type_names.contains(&folded)
    }

    fn owns_members(&self, type_name: &str) -> bool {
        let folded = fold_identifier(type_name);
        self.member_type_names.contains(&folded)
    }

    /// Whether this typelib's library reference defines a coclass `<reference>.<coclass>`
    /// (e.g. `Excel.Application` for the `Excel` library). Used to resolve a `WithEvents`
    /// EVENT SOURCE by coclass name even when the blob is a library-level reference that
    /// names no single coclass. Scoped to the event-source path: member resolution stays
    /// keyed on [`owns`] so a library reference does not lump every coclass's members.
    fn owns_coclass_source(&self, type_name: &str) -> bool {
        if self.blob.identity.requested_coclass.is_some() {
            return self.owns_members(type_name);
        }
        if self.owns(type_name) {
            return true;
        }
        let folded = fold_identifier(type_name);
        let reference = fold_identifier(&self.blob.identity.reference_name);
        self.blob
            .coclass_names
            .iter()
            .any(|coclass| folded == format!("{reference}.{}", fold_identifier(coclass)))
    }

    /// The coclass name a `WithEvents` receiver type maps to (e.g. `Application` for
    /// `Excel.Application`), used to scope event enumeration to that coclass. `None` when no
    /// specific coclass can be derived — the caller then keeps the whole event set.
    fn matched_coclass(&self, type_name: &str) -> Option<String> {
        let folded = fold_identifier(type_name);
        let reference = fold_identifier(&self.blob.identity.reference_name);
        if let Some(suffix) = folded.strip_prefix(&format!("{reference}.")) {
            return Some(suffix.to_string());
        }
        if let Some(coclass) = &self.blob.identity.requested_coclass {
            return Some(fold_identifier(coclass));
        }
        if self
            .blob
            .coclass_names
            .iter()
            .any(|coclass| fold_identifier(coclass) == folded)
        {
            return Some(folded);
        }
        None
    }

    pub fn activation_prog_id(&self) -> Option<&str> {
        self.blob.activation_prog_id.as_deref()
    }

    /// The application object this blob is scoped to, when its requested coclass is one.
    fn app_object(&self) -> Option<&calcvba_com::TypeLibAppObject> {
        let coclass = self.blob.identity.requested_coclass.as_deref()?;
        self.blob
            .app_objects
            .iter()
            .find(|object| object.coclass.eq_ignore_ascii_case(coclass))
    }

    /// `name` as a member of this blob's application object (`Range` of Excel's `Global`).
    pub fn resolve_app_object_member(&self, name: &str) -> Option<Binding> {
        let object = self.app_object()?;
        self.blob
            .members
            .iter()
            .any(|member| member.name.eq_ignore_ascii_case(name))
            .then(|| {
                Binding::new(
                    None,
                    DispatchRoute::AppObjectMember {
                        type_name: format!(
                            "{}.{}",
                            self.blob.identity.reference_name, object.coclass
                        ),
                        class: object.clsid.clone(),
                    },
                )
            })
    }

    pub fn resolve_object_root(&self, name: &str) -> Option<Binding> {
        if let Some(binding) = self.resolve_app_object_member(name) {
            return Some(binding);
        }
        if !self.owns(name) {
            return None;
        }
        Some(Binding::new(
            None,
            DispatchRoute::ComObjectRoot {
                type_name: name.to_string(),
                prog_id: self
                    .coclass_prog_id_for_name(name)
                    .or_else(|| self.activation_prog_id().map(str::to_string)),
            },
        ))
    }

    fn coclass_prog_id_for_name(&self, name: &str) -> Option<String> {
        let folded = fold_identifier(name);
        let reference = fold_identifier(&self.blob.identity.reference_name);
        self.blob.coclass_names.iter().find_map(|coclass| {
            let folded_coclass = fold_identifier(coclass);
            if folded == folded_coclass || folded == format!("{reference}.{folded_coclass}") {
                Some(
                    self.blob
                        .coclass_prog_ids
                        .iter()
                        .find(|(name, _)| name.eq_ignore_ascii_case(coclass))
                        .map(|(_, prog_id)| prog_id.clone())
                        .unwrap_or_else(|| {
                            format!("{}.{}", self.blob.identity.reference_name, coclass)
                        }),
                )
            } else {
                None
            }
        })
    }

    fn member_by_name(
        &self,
        name: &str,
        want: Option<ProjectMemberKind>,
    ) -> Option<&TypeLibMemberMetadata> {
        match want {
            Some(kind) => self.blob.members.iter().find(|member| {
                member.name.eq_ignore_ascii_case(name)
                    && member_kind_from_invoke(member.invoke_kind) == kind
            }),
            None => self
                .blob
                .members
                .iter()
                .filter(|member| member.name.eq_ignore_ascii_case(name))
                .filter_map(|member| read_member_rank(member).map(|rank| (rank, member)))
                .min_by_key(|(rank, _)| *rank)
                .map(|(_, member)| member),
        }
    }

    fn default_member(&self, want: Option<ProjectMemberKind>) -> Option<&TypeLibMemberMetadata> {
        match want {
            Some(kind) => self.blob.members.iter().find(|member| {
                member.is_default_member && member_kind_from_invoke(member.invoke_kind) == kind
            }),
            None => self
                .blob
                .members
                .iter()
                .filter(|member| member.is_default_member)
                .filter_map(|member| read_member_rank(member).map(|rank| (rank, member)))
                .min_by_key(|(rank, _)| *rank)
                .map(|(_, member)| member),
        }
    }

    fn scoped_source_events(
        &self,
        type_name: &str,
    ) -> Option<Vec<&calcvba_com::TypeLibEventMetadata>> {
        if !self.owns_coclass_source(type_name) {
            return None;
        }
        let want = self.matched_coclass(type_name);
        Some(
            self.blob
                .events
                .iter()
                .filter(|event| match (&want, &event.coclass) {
                    (Some(want), Some(have)) => fold_identifier(have) == *want,
                    _ => true,
                })
                .collect(),
        )
    }
}

impl Provider for ComTypeLibProvider {
    fn resolve(&self, name: &str) -> Option<Binding> {
        self.resolve_constant(&[name])
            .or_else(|| self.resolve_app_object_member(name))
    }
    fn resolve_qualified(&self, parts: &[&str]) -> Option<Binding> {
        self.resolve_constant(parts).or_else(|| match parts {
            // `Excel.Range`: the library name qualifies its application object's members.
            [library, member]
                if library.eq_ignore_ascii_case(&self.blob.identity.reference_name) =>
            {
                self.resolve_app_object_member(member)
            }
            _ => None,
        })
    }
    fn is_extensible_object_type(&self, type_name: &str) -> bool {
        self.blob.extensible && self.owns_members(type_name)
    }
    fn declares_evaluate_member(&self, type_name: &str) -> bool {
        self.owns_members(type_name)
            && self
                .blob
                .members
                .iter()
                .any(|member| member.token == DISPID_EVALUATE)
    }
    fn is_library_name(&self, name: &str) -> bool {
        self.blob.identity.reference_name.eq_ignore_ascii_case(name)
    }
    fn foreign_evaluator(&self) -> Option<(String, String)> {
        let object = self.app_object()?;
        self.blob
            .members
            .iter()
            .any(|member| member.token == DISPID_EVALUATE)
            .then(|| {
                (
                    format!("{}.{}", self.blob.identity.reference_name, object.coclass),
                    object.clsid.clone(),
                )
            })
    }
    fn has_ambiguous_unqualified_name(&self, name: &str) -> bool {
        self.blob
            .constants
            .iter()
            .filter(|c| c.name.eq_ignore_ascii_case(name))
            .take(2)
            .count()
            > 1
    }
    fn source_event_metadata(
        &self,
        recv: &VarTypeRef,
    ) -> Option<Vec<calcvba_com::TypeLibEventMetadata>> {
        let VarTypeRef::Object(name) = recv else {
            return None;
        };
        self.scoped_source_events(name)
            .map(|events| events.into_iter().cloned().collect())
    }
    fn object_interface(&self, type_name: &str) -> Option<calcvba_com::TypeLibInterfaceMetadata> {
        if !self.owns_members(type_name) {
            return None;
        }
        let iid = self.interface_iid?;
        Some(calcvba_com::TypeLibInterfaceMetadata {
            name: type_name.to_string(),
            iid: Some(iid),
            members: self.blob.members.clone(),
        })
    }
    fn resolve_member(
        &self,
        recv: &VarTypeRef,
        name: &str,
        want: Option<ProjectMemberKind>,
    ) -> Option<Binding> {
        let VarTypeRef::Object(type_name) = recv else {
            return None;
        };
        if self.owns_members(type_name)
            && let Some(member) = self.member_by_name(name, want)
        {
            return Some(com_member_binding(member, type_name));
        }

        if want.is_none()
            && let Some(events) = self.scoped_source_events(type_name)
            && let Some(event) = events
                .into_iter()
                .find(|event| event.name.eq_ignore_ascii_case(name))
        {
            return Some(Binding::new(
                None,
                DispatchRoute::ComEvent {
                    event_name: event.name.clone(),
                    token: event.token,
                    dispatch_path: event.dispatch_path,
                    dispatch_member_id: event.dispatch_member_id,
                    connection_point_iid: event.connection_point_iid.clone(),
                    callback_arity: event.callback_arity,
                },
            ));
        }

        None
    }

    fn resolve_default_member(&self, recv: &VarTypeRef) -> Option<Binding> {
        let VarTypeRef::Object(type_name) = recv else {
            return None;
        };
        if !self.owns_members(type_name) {
            return None;
        }
        self.default_member(None)
            .map(|member| com_member_binding(member, type_name))
    }

    fn resolve_default_member_kind(
        &self,
        recv: &VarTypeRef,
        want: Option<ProjectMemberKind>,
    ) -> Option<Binding> {
        let VarTypeRef::Object(type_name) = recv else {
            return None;
        };
        if !self.owns_members(type_name) {
            return None;
        }
        self.default_member(want)
            .map(|member| com_member_binding(member, type_name))
    }

    fn resolve_coclass(&self, name: &str) -> Option<String> {
        if !self.owns(name) {
            return None;
        }
        self.coclass_prog_id_for_name(name)
            .or_else(|| self.activation_prog_id().map(str::to_string))
    }

    fn is_known_object_type(&self, type_name: &str) -> bool {
        self.owns(type_name)
    }

    fn source_events(&self, recv: &VarTypeRef) -> Option<Vec<(String, i32)>> {
        let VarTypeRef::Object(type_name) = recv else {
            return None;
        };
        // The event `token` is the value the runtime keys subscriptions on (it is what
        // `subscribe_event` is passed), so the route built from this list routes a delivered
        // callback for that dispid back to the handler.
        //
        // A library-wide blob holds EVERY coclass's source events; scope to the coclass this
        // WithEvents field is typed as (e.g. `Excel.Worksheet` -> only Worksheet's events)
        // so the binder stops synthesizing phantom routes for unrelated coclasses (which the
        // runtime would then re-recover per subscribe and reject). Events with no recovered
        // coclass tag, or a coclass we cannot derive, are kept because the metadata gives no
        // receiver-specific basis for excluding them.
        self.scoped_source_events(type_name).map(|events| {
            events
                .into_iter()
                .map(|event| (fold_identifier(&event.name), event.token))
                .collect()
        })
    }
}

fn read_member_rank(member: &TypeLibMemberMetadata) -> Option<u8> {
    match member.invoke_kind {
        TypeLibMemberInvokeKind::PropertyGet => Some(0),
        TypeLibMemberInvokeKind::Method => Some(1),
        TypeLibMemberInvokeKind::PropertyPut | TypeLibMemberInvokeKind::PropertyPutRef => None,
    }
}

fn com_member_binding(member: &calcvba_com::TypeLibMemberMetadata, type_name: &str) -> Binding {
    Binding {
        symbol: None,
        is_default: member.is_default_member,
        route: DispatchRoute::ComMember {
            member_name: member.name.clone(),
            dispid: member.token,
            vtable_slot: member.vtable_slot,
            invoke_kind: member.invoke_kind,
            member_kind: member_kind_from_invoke(member.invoke_kind),
            is_default_member: member.is_default_member,
            param_by_ref: member
                .parameter_types
                .iter()
                .map(|t| t.is_by_ref())
                .collect(),
            // The declared receiver COM type — the de-erased early-bound call's
            // typed-interface grouping key — and the full canonical descriptor.
            interface_name: type_name.to_string(),
            member: Box::new(member.clone()),
        },
    }
}
