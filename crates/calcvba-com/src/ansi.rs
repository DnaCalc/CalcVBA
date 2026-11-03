//! ANSI (system code page) ↔ Unicode conversion for byte-exact VBA string I/O.
//!
//! The codec itself now lives in `calcvba-runtime` (the value substrate shared by
//! both the `Declare` ANSI string marshalling here and the `Chr`/`Asc` built-ins
//! in `calcvba-lib`). This module re-exports it so existing `calcvba_com::ansi::…`
//! call sites keep working unchanged.

pub use calcvba_runtime::ansi::{ansi_decode, ansi_encode};
