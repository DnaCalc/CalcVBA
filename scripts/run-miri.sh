#!/bin/bash
# Run cargo miri on unsafe-heavy crates.
# Excludes Windows-specific COM FFI paths (Windows APIs are incompatible with miri) and the
# libffi/dlopen calls of the Declare bridge.
# calcvba-automation holds the portable VARIANT, BSTR, SAFEARRAY and record values
# themselves, so its tests check the platform layout without COM.
# proptest needs the working directory (isolation off) and runs few cases here.
export MIRIFLAGS="${MIRIFLAGS:--Zmiri-disable-isolation}"
export PROPTEST_CASES="${PROPTEST_CASES:-4}"
#
# Prerequisites:
#   rustup +nightly component add miri
#
# Usage:
#   bash scripts/run-miri.sh

set -euo pipefail

# The main runs keep strict accounting of integer-to-pointer casts (none expected); the
# `exposed_wire` tests decode exposed addresses on purpose and run on their own with
# permissive provenance.
echo "=== Running miri on calcvba-automation ==="
cargo +nightly miri test -p calcvba-automation -- --skip exposed_wire 2>&1 || {
    echo "FAIL: calcvba-automation miri"
    exit 1
}

echo "=== Running miri on calcvba-runtime ==="
cargo +nightly miri test -p calcvba-runtime -- --skip exposed_wire 2>&1 || {
    echo "FAIL: calcvba-runtime miri"
    exit 1
}

echo "=== Running miri on the wire decoder (exposed provenance) ==="
MIRIFLAGS="$MIRIFLAGS -Zmiri-permissive-provenance" \
    cargo +nightly miri test -p calcvba-automation -p calcvba-runtime -- exposed_wire 2>&1 || {
    echo "FAIL: wire decoder miri"
    exit 1
}

# calcvba-com: its portable library tests (typelib metadata, bindings, policy). Skipped: the
# windows_ffi_bridge tests, which call the C library through dlopen and libffi (foreign calls
# Miri cannot run). Leaks are ignored because typelib runtime descriptors are leaked to
# 'static on purpose (Box::leak), which Miri reports when nothing global holds them.
echo "=== Running miri on calcvba-com (library tests) ==="
MIRIFLAGS="$MIRIFLAGS -Zmiri-ignore-leaks" \
    cargo +nightly miri test -p calcvba-com --lib -- --skip windows_ffi_bridge::tests 2>&1 || {
    echo "FAIL: calcvba-com miri"
    exit 1
}

echo "=== Running miri on calcvba-jit ==="
cargo +nightly miri test -p calcvba-jit 2>&1 || {
    echo "FAIL: calcvba-jit miri"
    exit 1
}

echo "=== All miri checks passed ==="
