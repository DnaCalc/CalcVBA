//! Automation value layouts shared by CalcVBA's runtime, its JIT and native lanes, and other
//! in-process libraries (CalcDb, DNA Calc) that exchange values with it without conversion.
//!
//! The platform's Automation ABI is the specification: a `VARIANT` is 24 bytes on 64-bit
//! targets and 16 bytes where pointers are 32-bit (wasm32 included), because its union holds
//! `BRECORD { pvRecord, pRecInfo }`. See `docs/design/runtime-variant-automation-layout-2026-11-13.md`.
#![no_std]

pub mod abi;
pub mod com;
