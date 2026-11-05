//! Native project objects. Lifecycle entries are compiled VBA functions; no
//! source instructions or procedure tokens are evaluated by these helpers.
use calcvba_rt_abi::native_export::Context;
use calcvba_runtime::{ObjectRef, object_ref::*};
use std::{collections::HashMap, ffi::c_void};

type Lifecycle = unsafe extern "C" fn(*mut Context, *mut RawRuntimeIUnknown, *mut u8) -> i32;

/// # Safety
/// Out owns one initialized Object cell.
#[unsafe(no_mangle)]
pub unsafe extern "C" fn cvb_collection_new(out: *mut u8) -> i32 {
    super::values::boundary(|| {
        // SAFETY: the returned object transfers through a retaining native copy.
        unsafe {
            let object = calcvba_rt_abi::host_entry::collection::new_collection();
            let pointer = object.raw_iunknown();
            super::values::copy(out, std::ptr::from_ref(&pointer).cast(), 9, 9)
        }
    })
}

/// # Safety
/// Cell is one initialized owning Object cell on the native calling thread.
#[unsafe(no_mangle)]
pub unsafe extern "C" fn cvb_collection_ensure(cell: *mut u8) -> i32 {
    // SAFETY: the retained Object cell contains either Nothing or a live owner.
    unsafe {
        if cell.cast::<usize>().read() == 0 {
            cvb_collection_new(cell)
        } else {
            0
        }
    }
}

/// # Safety
/// Cell has the declared kind, name is immutable UTF-8 image data and out, when
/// nonnull, is a writable native Boolean. Assignment mode validates Set.
#[unsafe(no_mangle)]
pub unsafe extern "C" fn cvb_object_type(
    context: *mut Context,
    cell: *const u8,
    kind: u32,
    name: *const u8,
    len: usize,
    assignment: u32,
    out: *mut i16,
) -> i32 {
    super::values::boundary(|| {
        // SAFETY: all arguments are retained by the generated calling frame.
        unsafe {
            let value = super::services::read(context, cell, kind)?;
            let name = std::str::from_utf8_unchecked(std::slice::from_raw_parts(name, len));
            let bare = name.rsplit('.').next().unwrap_or(name);
            let mut iid = [0u8; 16];
            std::ptr::copy_nonoverlapping(name.as_ptr().add(len), iid.as_mut_ptr(), 16);
            let iid = (iid != [0; 16]).then_some(iid);
            let matches = if let Some(object) = value.as_object_ref() {
                if bare.eq_ignore_ascii_case("Object") {
                    true
                } else {
                    super::services::foreign_type_matches(context, object, name, iid)?
                }
            } else {
                assignment != 0 && value.vtype() == calcvba_runtime::VarType::Object
            };
            if assignment != 0 && !matches {
                return Err(13);
            }
            if !out.is_null() {
                out.write(if matches { -1 } else { 0 });
            }
        }
        Ok(())
    })
}

pub(super) struct Classes {
    owner: calcvba_runtime::object_owner::RuntimeObjectOwner,
    next: i32,
    definitions: HashMap<i32, Definition>,
}
struct Definition {
    descriptor: &'static RuntimeClassDescriptor,
    fields: Vec<(i32, u32)>,
    terminate: Option<Lifecycle>,
}
impl Default for Classes {
    fn default() -> Self {
        Self {
            owner: calcvba_runtime::object_owner::RuntimeObjectOwner::fresh()
                .expect("native object owner identity"),
            next: 1,
            definitions: HashMap::new(),
        }
    }
}

unsafe fn definition(metadata: *const u8, terminate: usize) -> Definition {
    // SAFETY: the compiler emits this immutable, length-prefixed class schema.
    unsafe {
        let mut at = metadata;
        let mut word = || {
            let value = at.cast::<u32>().read_unaligned();
            at = at.add(4);
            value
        };
        let _class = word();
        let name_len = word() as usize;
        let implements_count = word();
        let field_count = word();
        let name = std::str::from_utf8_unchecked(std::slice::from_raw_parts(at, name_len));
        at = at.add(name_len);
        let mut implements = Vec::new();
        for _ in 0..implements_count {
            let len = at.cast::<u32>().read_unaligned() as usize;
            at = at.add(4);
            let name = std::str::from_utf8_unchecked(std::slice::from_raw_parts(at, len));
            at = at.add(len);
            implements.push(name);
        }
        let mut fields = Vec::new();
        for _ in 0..field_count {
            let token = at.cast::<i32>().read_unaligned();
            let kind = at.add(4).cast::<u32>().read_unaligned();
            at = at.add(8);
            fields.push((token, kind));
        }
        // Like the VM/JIT descriptor pool, immutable class identities live for
        // the process lifetime. Strings borrow the loaded image's schema only
        // while copying into that pool; no pointer into the DLL escapes it.
        let name = Box::leak(name.to_owned().into_boxed_str());
        let implements = Box::leak(
            implements
                .into_iter()
                .map(|name| &*Box::leak(name.to_owned().into_boxed_str()))
                .collect::<Vec<_>>()
                .into_boxed_slice(),
        );
        Definition {
            descriptor: Box::leak(Box::new(RuntimeClassDescriptor {
                name,
                project_identity: None,
                predeclared: false,
                lifecycle: RuntimeClassLifecycleDescriptor {
                    has_initialize: false,
                    has_terminate: terminate != 0,
                },
                fields: &[],
                as_new_fields: &[],
                implements,
                interfaces: &[RUNTIME_IUNKNOWN_INTERFACE_DESCRIPTOR],
            })),
            fields,
            terminate: (terminate != 0).then(|| std::mem::transmute::<usize, Lifecycle>(terminate)),
        }
    }
}

/// # Safety
/// Metadata is a compiler-emitted class schema; lifecycle addresses use the
/// native VBA ABI. Out owns an initialized Object cell on the calling thread.
#[unsafe(no_mangle)]
pub unsafe extern "C" fn cvb_class_new(
    context: *mut Context,
    metadata: *const u8,
    initialize: usize,
    terminate: usize,
    out: *mut u8,
) -> i32 {
    super::values::boundary(|| {
        // SAFETY: all borrows of runtime state end before entering compiled VBA.
        unsafe {
            let class = metadata.cast::<i32>().read_unaligned();
            let state = super::services::runtime(context);
            let (owner, descriptor, fields, id) = {
                let classes = &mut (*state).classes;
                let definition = classes
                    .definitions
                    .entry(class)
                    .or_insert_with(|| definition(metadata, terminate));
                let id = classes.next;
                classes.next = id.checked_add(1).ok_or(7)?;
                (
                    classes.owner.clone(),
                    definition.descriptor,
                    definition.fields.clone(),
                    id,
                )
            };
            let _scope = owner.enter();
            let object = ObjectRef::from_project_instance(id, class, 0, terminate != 0, descriptor);
            for (field, kind) in fields {
                let zero = [0u64; 3];
                let value = super::services::read(context, zero.as_ptr().cast(), kind)?;
                if !object.project_field_set(field, value) {
                    return Err(438);
                }
            }
            if initialize != 0 {
                let initialize: Lifecycle = std::mem::transmute(initialize);
                let mut result = [0u64; 3];
                let number = initialize(context, object.raw_iunknown(), result.as_mut_ptr().cast());
                if number != 0 {
                    object.mark_initialization_failed();
                    return Err(number);
                }
            }
            let pointer = object.into_raw_iunknown();
            let result = super::values::copy(out, std::ptr::from_ref(&pointer).cast(), 9, 9);
            drop(ObjectRef::from_raw_iunknown_owned(pointer));
            result
        }
    })
}

/// # Safety
/// Same contract as cvb_class_new; only a null owning Object cell is allocated.
#[unsafe(no_mangle)]
pub unsafe extern "C" fn cvb_as_new(
    context: *mut Context,
    metadata: *const u8,
    initialize: usize,
    terminate: usize,
    cell: *mut u8,
) -> i32 {
    // SAFETY: caller retains this initialized Object cell throughout the call.
    unsafe {
        if cell.cast::<*mut c_void>().read().is_null() {
            cvb_class_new(context, metadata, initialize, terminate, cell)
        } else {
            0
        }
    }
}

/// # Safety
/// Object is a retained COM-shaped reference and out is a writable i32.
#[unsafe(no_mangle)]
pub unsafe extern "C" fn cvb_class_route(object: *mut RawRuntimeIUnknown, out: *mut i32) -> i32 {
    super::values::boundary(|| {
        // SAFETY: caller retains the object through this temporary reference.
        unsafe {
            let object = ObjectRef::from_raw_iunknown_addref(object).ok_or(91)?;
            out.write(if object.is_project_instance() {
                object.route_key()
            } else {
                -1
            });
        }
        Ok(())
    })
}

/// # Safety
/// Object is retained; cell has exactly kind's initialized native carrier.
#[unsafe(no_mangle)]
pub unsafe extern "C" fn cvb_class_field(
    context: *mut Context,
    object: *mut RawRuntimeIUnknown,
    field: i32,
    cell: *mut u8,
    kind: u32,
    set: u32,
) -> i32 {
    super::values::boundary(|| {
        // SAFETY: retained independent values; no field borrow spans a callback.
        unsafe {
            let object = ObjectRef::from_raw_iunknown_addref(object).ok_or(91)?;
            if !object.is_project_instance() {
                return Err(13);
            }
            if set != 0 {
                let value = super::services::read(context, cell, kind)?;
                if !object.project_field_set(field, value) {
                    return Err(438);
                }
                Ok(())
            } else {
                let value = object.project_field_get(field).ok_or(438)?;
                super::services::write(context, cell, kind, &value)
            }
        }
    })
}

/// # Safety
/// Context belongs to the native image's owner and compiled lifecycle entries
/// remain loaded. No mutable runtime borrow may span the compiled callback.
#[unsafe(no_mangle)]
pub unsafe extern "C" fn cvb_drain_terminations(context: *mut Context) -> i32 {
    super::values::boundary(|| {
        // SAFETY: parked objects and their callbacks stay owned during execution.
        unsafe {
            let state = super::services::runtime(context);
            let owner = (*state).classes.owner.clone();
            let _scope = owner.enter();
            loop {
                let pending = calcvba_runtime::take_pending_terminations();
                if pending.is_empty() {
                    break;
                }
                for (id, _, class) in pending {
                    if let Some(object) = calcvba_runtime::retained_parked_termination_object(id) {
                        let terminate = (*state)
                            .classes
                            .definitions
                            .get(&class)
                            .and_then(|d| d.terminate);
                        if let Some(terminate) = terminate {
                            let mut result = [0u64; 3];
                            let _ = terminate(
                                context,
                                object.raw_iunknown(),
                                result.as_mut_ptr().cast(),
                            );
                        }
                        drop(object);
                    }
                    calcvba_runtime::finish_pending_termination(id);
                }
            }
        }
        Ok(())
    })
}
