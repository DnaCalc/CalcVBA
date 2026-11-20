#!/bin/bash
# Run cargo miri on unsafe-heavy crates.
# Excludes Windows-specific COM FFI paths (Windows APIs are incompatible with miri).
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

echo "=== Running miri on calcvba-automation ==="
cargo +nightly miri test -p calcvba-automation 2>&1 || {
    echo "FAIL: calcvba-automation miri"
    exit 1
}

echo "=== Running miri on calcvba-runtime ==="
cargo +nightly miri test -p calcvba-runtime 2>&1 || {
    echo "FAIL: calcvba-runtime miri"
    exit 1
}

echo "=== Running miri on calcvba-jit ==="
cargo +nightly miri test -p calcvba-jit 2>&1 || {
    echo "FAIL: calcvba-jit miri"
    exit 1
}

echo "=== All miri checks passed ==="
