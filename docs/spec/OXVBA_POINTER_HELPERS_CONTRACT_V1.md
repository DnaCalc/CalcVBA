# Pointer helpers

StrPtr, VarPtr and ObjPtr are native interop operations, not portable integer
identities for arbitrary engine values. Their supported routes retain the
underlying storage for the native call and fail explicitly for unsupported
objects or layouts.

StrPtr exposes the UTF-16 payload of an owned BSTR on Windows. VarPtr of a
String refers to its native BSTR-pointer cell, which differs from the string's
character payload; VarPtr of a Variant refers to its native VARIANT cell.
ObjPtr must preserve the supported object's native identity and reference
lifetime. A synthetic object token is not a native COM pointer.

`oxvba-runtime/src/pointer_helpers.rs` owns retained storage/pins and
`oxvba-com` consumes them at the FFI boundary. Native parameter cells and
pinned owners outlive every synchronous use and copyback. Raw addresses must
not escape their supported lifetime or move to another runtime thread.
General pointer/native callback signatures and physical carriers remain
incomplete. The pointer, Declare and controlled native tests describe the
supported cases on both backends.
