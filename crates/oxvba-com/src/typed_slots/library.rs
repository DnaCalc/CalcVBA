//! Resolve native result identities once, before publishing a member table.
use super::*;
use crate::TypeLibInterfaceMetadata;

#[derive(Clone)]
pub(super) struct NativeMemberPlan {
    pub member: TypeLibMemberMetadata,
    pub result_interface: Option<usize>,
    pub parameter_interfaces: Vec<Option<usize>>,
}
pub(super) struct NativeLibrary {
    pub metadata: Vec<TypeLibInterfaceMetadata>,
    pub plans: Vec<Vec<NativeMemberPlan>>,
}
impl NativeLibrary {
    pub fn new(
        mut metadata: Vec<TypeLibInterfaceMetadata>,
        root: &[TypeLibMemberMetadata],
    ) -> Result<(Self, usize), String> {
        let iid = root
            .first()
            .and_then(|m| m.interface_iid)
            .ok_or("native root IID is missing")?;
        let root_index = metadata
            .iter()
            .position(|i| i.iid == Some(iid))
            .unwrap_or_else(|| {
                metadata.push(TypeLibInterfaceMetadata {
                    name: "TypedRoot".into(),
                    iid: Some(iid),
                    members: root.to_vec(),
                });
                metadata.len() - 1
            });
        let plans = metadata
            .iter()
            .map(|interface| {
                interface
                    .members
                    .iter()
                    .filter_map(|member| {
                        if !plan::validate(member) {
                            return None;
                        }
                        let result_interface =
                            if let Some(W::InterfacePointer { name }) = &member.return_wire_type {
                                Some(resolve(&metadata, name)?)
                            } else {
                                None
                            };
                        let parameter_interfaces: Vec<_> =
                            member
                                .parameter_types
                                .iter()
                                .enumerate()
                                .map(|(i, _)| {
                                    member.parameter_iids.get(i).copied().flatten().and_then(
                                        |iid| {
                                            metadata.iter().position(|item| {
                                                item.iid == Some(iid) && dispatch_interface(item)
                                            })
                                        },
                                    )
                                })
                                .collect();
                        for (i, ty) in member.parameter_types.iter().enumerate() {
                            if *ty == P::ByRefObject
                                && member
                                    .parameter_iids
                                    .get(i)
                                    .is_some_and(|iid| iid.is_some())
                                && parameter_interfaces[i].is_none()
                                && member.parameter_iids[i]
                                    != Some(crate::ComInterfaceIid {
                                        data1: 0x00020400,
                                        data2: 0,
                                        data3: 0,
                                        data4: [0xc0, 0, 0, 0, 0, 0, 0, 0x46],
                                    })
                            {
                                return None;
                            }
                        }
                        Some(NativeMemberPlan {
                            member: member.clone(),
                            result_interface,
                            parameter_interfaces,
                        })
                    })
                    .collect()
            })
            .collect();
        Ok((Self { metadata, plans }, root_index))
    }
}
fn resolve(metadata: &[TypeLibInterfaceMetadata], name: &str) -> Option<usize> {
    let exact = metadata
        .iter()
        .position(|i| i.name.eq_ignore_ascii_case(name));
    let index = if let Some(exact) = exact {
        exact
    } else {
        let mut matches = metadata.iter().enumerate().filter(|(_, i)| {
            i.name
                .rsplit('.')
                .next()
                .is_some_and(|bare| bare.eq_ignore_ascii_case(name))
        });
        let (index, first) = matches.next()?;
        if matches.any(|(_, i)| i.iid != first.iid) {
            return None;
        }
        index
    };
    let target = &metadata[index];
    // This family exposes dual interfaces. An arbitrary IUnknown pointer must
    // never be re-labelled as IDispatch merely because both have three slots.
    target.iid?;
    dispatch_interface(target).then_some(index)
}
fn dispatch_interface(target: &TypeLibInterfaceMetadata) -> bool {
    target.members.iter().any(|m| {
        plan::validate(m)
            || (m.source_typekind == Some(SourceTypeKind::Dispatch)
                && !m.is_dual
                && m.vtable_slot.is_none()
                && m.vtable_slot_bound == Some(7))
    })
}
