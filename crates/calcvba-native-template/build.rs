fn main() {
    if std::env::var("CARGO_CFG_TARGET_OS").as_deref() == Ok("windows") {
        // Developer link only: enough header space for generated sections.
        // No custom /ENTRY; retain the platform CRT/Rust DLL startup.
        println!("cargo:rustc-cdylib-link-arg=/FILEALIGN:4096");
        println!("cargo:rustc-cdylib-link-arg=/GUARD:NO");
        println!("cargo:rustc-cdylib-link-arg=/INCREMENTAL:NO");
    }
}
