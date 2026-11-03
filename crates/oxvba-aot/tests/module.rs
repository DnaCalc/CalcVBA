#![cfg(all(windows, target_arch = "x86_64"))]
use cranelift_codegen::ir::{AbiParam, InstBuilder, types};
use cranelift_frontend::{FunctionBuilder, FunctionBuilderContext};
use cranelift_module::{DataDescription, Linkage, Module};
use oxvba_aot::PeModule;
use oxvba_pe::{HELPER_ABI, Template};

#[test]
fn module_relocates_direct_calls_helpers_constants_and_windows_unwind() {
    let bytes = include_bytes!("../../../assets/native/calcvba-native-x64.dll");
    let template = Template::parse(bytes, HELPER_ABI).unwrap();
    let mut module = PeModule::new().unwrap();
    let mut sig = module.make_signature();
    sig.params.push(AbiParam::new(types::I32));
    sig.returns.push(AbiParam::new(types::I32));
    let helper = module.declare_helper("cvb_identity_i32", 1, &sig).unwrap();
    let inner = module
        .declare_function("inner", Linkage::Local, &sig)
        .unwrap();
    let outer = module
        .declare_function("outer", Linkage::Export, &sig)
        .unwrap();
    let constant = module
        .declare_data("constant", Linkage::Local, false, false)
        .unwrap();
    let mut data = DataDescription::new();
    data.define(5i32.to_le_bytes().to_vec().into_boxed_slice());
    module.define_data(constant, &data).unwrap();
    for id in [inner, outer] {
        let mut context = module.make_context();
        context.func.signature = sig.clone();
        let mut frontend = FunctionBuilderContext::new();
        {
            let mut b = FunctionBuilder::new(&mut context.func, &mut frontend);
            let entry = b.create_block();
            b.switch_to_block(entry);
            b.append_block_params_for_function_params(entry);
            let arg = b.block_params(entry)[0];
            let (target, arg) = if id == inner {
                let global = module.declare_data_in_func(constant, b.func);
                let address = b.ins().global_value(types::I64, global);
                let five = b.ins().load(
                    types::I32,
                    cranelift_codegen::ir::MemFlagsData::new(),
                    address,
                    0,
                );
                (helper, b.ins().iadd(arg, five))
            } else {
                (inner, arg)
            };
            let target = module.declare_func_in_func(target, b.func);
            let call = b.ins().call(target, &[arg]);
            let result = b.inst_results(call)[0];
            b.ins().return_(&[result]);
            b.seal_all_blocks();
            b.finalize();
        }
        module.define_function(id, &mut context).unwrap();
    }
    let program = module
        .finish(&template, &[("Compute".into(), None, outer)])
        .unwrap();
    assert_eq!(program.unwind.len(), 2);
    let patched = template.patch("module.dll", &program).unwrap();
    let image = oxvba_pe::Image::parse(&patched.bytes).unwrap();
    assert_eq!(
        image.runtime_functions().unwrap().len(),
        oxvba_pe::Image::parse(bytes)
            .unwrap()
            .runtime_functions()
            .unwrap()
            .len()
            + 2
    );
    // RUNTIME_FUNCTIONs locate each generated function. Both functions must
    // contain E8 rel32, with outer targeting inner and inner the helper RVA.
    let text = template.text_rva().unwrap();
    let helper_rva = template.manifest.resolve("cvb_identity_i32", 1).unwrap();
    let targets: Vec<_> = program
        .unwind
        .iter()
        .map(|f| {
            let code = &program.text[f.begin as usize..f.end as usize];
            let at = code
                .iter()
                .position(|b| *b == 0xe8)
                .expect("direct call opcode");
            let displacement = i32::from_le_bytes(code[at + 1..at + 5].try_into().unwrap());
            (i64::from(text) + i64::from(f.begin) + at as i64 + 5 + i64::from(displacement)) as u32
        })
        .collect();
    assert_eq!(targets, vec![helper_rva, text + program.unwind[0].begin]);
    let directory = std::env::temp_dir().join(format!("calcvba-pe-module-{}", std::process::id()));
    std::fs::create_dir_all(&directory).unwrap();
    let path = directory.join("module.dll");
    std::fs::write(&path, &patched.bytes).unwrap();
    use windows_sys::Win32::{
        Foundation::FreeLibrary,
        System::LibraryLoader::{GetProcAddress, LoadLibraryW},
    };
    let path16: Vec<_> = path
        .to_str()
        .unwrap()
        .encode_utf16()
        .chain(Some(0))
        .collect();
    // SAFETY: complete generated DLL, NUL-terminated path; code retained until
    // FreeLibrary. Compute's signature was declared and compiled just above.
    unsafe {
        let dll = LoadLibraryW(path16.as_ptr());
        assert!(!dll.is_null());
        let compute: extern "C" fn(i32) -> i32 =
            std::mem::transmute(GetProcAddress(dll, c"Compute".as_ptr().cast()).unwrap());
        assert_eq!(compute(37), 42);
        FreeLibrary(dll);
    }
    std::fs::remove_file(path).unwrap();
    std::fs::remove_dir(directory).unwrap();
}

#[test]
fn refuses_an_unversioned_or_mismatched_helper_before_patching() {
    let bytes = include_bytes!("../../../assets/native/calcvba-native-x64.dll");
    let template = Template::parse(bytes, HELPER_ABI).unwrap();
    for version in [None, Some(2)] {
        let mut module = PeModule::new().unwrap();
        let mut sig = module.make_signature();
        sig.params.push(AbiParam::new(types::I32));
        sig.returns.push(AbiParam::new(types::I32));
        let helper = if let Some(version) = version {
            module
                .declare_helper("cvb_identity_i32", version, &sig)
                .unwrap()
        } else {
            module
                .declare_function("cvb_identity_i32", Linkage::Import, &sig)
                .unwrap()
        };
        let entry = module
            .declare_function("entry", Linkage::Export, &sig)
            .unwrap();
        let mut context = module.make_context();
        context.func.signature = sig;
        let mut frontend = FunctionBuilderContext::new();
        {
            let mut b = FunctionBuilder::new(&mut context.func, &mut frontend);
            let block = b.create_block();
            b.switch_to_block(block);
            b.append_block_params_for_function_params(block);
            let value = b.block_params(block)[0];
            let helper = module.declare_func_in_func(helper, b.func);
            let call = b.ins().call(helper, &[value]);
            let result = b.inst_results(call)[0];
            b.ins().return_(&[result]);
            b.seal_all_blocks();
            b.finalize();
        }
        module.define_function(entry, &mut context).unwrap();
        assert!(
            module
                .finish(&template, &[("Entry".into(), None, entry)])
                .is_err()
        );
    }
}
