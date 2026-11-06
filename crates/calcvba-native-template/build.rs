fn main() {
    if std::env::var("CARGO_CFG_TARGET_OS").as_deref() == Ok("windows") {
        // Developer link only: enough header space for generated sections.
        // No custom /ENTRY; retain the platform CRT/Rust DLL startup.
        println!("cargo:rustc-cdylib-link-arg=/FILEALIGN:4096");
        println!("cargo:rustc-cdylib-link-arg=/GUARD:NO");
        println!("cargo:rustc-cdylib-link-arg=/INCREMENTAL:NO");
        for flag in ["/FILEALIGN:4096", "/GUARD:NO", "/INCREMENTAL:NO"] {
            println!("cargo:rustc-link-arg-bin=calcvba-native-exe-template={flag}");
        }
        let root = std::path::PathBuf::from(std::env::var_os("CARGO_MANIFEST_DIR").unwrap());
        let contract = root.join("../../assets/native/helpers-v1.txt");
        println!("cargo:rerun-if-changed={}", contract.display());
        let text = std::fs::read_to_string(contract).expect("native helper contract");
        let names: Vec<_> = text
            .lines()
            .skip(1)
            .map(|line| line.split('|').next().unwrap())
            .collect();
        let mut source = String::from("pub fn retain_native_helpers() { unsafe extern \"C\" {\n");
        for name in &names {
            println!("cargo:rustc-link-arg-bin=calcvba-native-exe-template=/EXPORT:{name}");
            source.push_str(&format!("fn {name}();\n"));
        }
        println!(
            "cargo:rustc-link-arg-bin=calcvba-native-exe-template=/EXPORT:cvb_template_abi,DATA"
        );
        source.push_str("}\nstd::hint::black_box([\n");
        for name in &names {
            source.push_str(&format!("{name} as *const (),\n"));
        }
        source.push_str("]);\nstd::hint::black_box(std::ptr::addr_of!(CVB_TEMPLATE_ABI));\n}\n");
        let output = std::path::PathBuf::from(std::env::var_os("OUT_DIR").unwrap());
        std::fs::write(output.join("native_keepalive.rs"), source)
            .expect("native helper retention source");
    }
}
