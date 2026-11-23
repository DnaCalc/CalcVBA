$ErrorActionPreference = "Stop"
$PSNativeCommandUseErrorActionPreference = $true

Push-Location (Join-Path $PSScriptRoot "..")
try {
    cargo fmt --all --check
    if ($IsLinux) {
        $windowsTarget = "x86_64-pc-windows-msvc"
        $windowsLibDir = rustc --print target-libdir --target $windowsTarget
        if (Test-Path -LiteralPath $windowsLibDir -PathType Container) {
            cargo check -p calcvba-automation --tests --target $windowsTarget
        }
    }
    cargo clippy --workspace --all-targets -- -D warnings
    # These four optional COM fixtures register type libraries. Ordinary checks
    # must leave COM registrations alone, on developer machines and in CI.
    cargo test --workspace -- --skip windows_vtable::tests::record_safearray_parameter_lowers_with_record_element_vartype --skip windows_vtable::tests::descriptor_backed_empty_record_safearray_parameter_lowers_with_record_element_vartype --skip windows_vtable::tests::record_return_allocates_descriptor_backed_record_cell --skip windows_vtable::tests::typed_byref_record_safearray_returns_record_writeback
}
finally {
    Pop-Location
}
