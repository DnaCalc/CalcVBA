use calcvba_pe::{
    Export, HELPER_ABI, HelperSpec, Image, Program, Template, Unwind, finish_template,
};
use object::{Object, ObjectSection};

fn p16(b: &mut [u8], at: usize, v: u16) {
    b[at..at + 2].copy_from_slice(&v.to_le_bytes());
}
fn p32(b: &mut [u8], at: usize, v: u32) {
    b[at..at + 4].copy_from_slice(&v.to_le_bytes());
}
fn p64(b: &mut [u8], at: usize, v: u64) {
    b[at..at + 8].copy_from_slice(&v.to_le_bytes());
}
fn section(b: &mut [u8], at: usize, name: &[u8], rva: u32, raw: u32, flags: u32) {
    b[at..at + name.len()].copy_from_slice(name);
    for (off, val) in [(8, 512), (12, rva), (16, 512), (20, raw), (36, flags)] {
        p32(b, at + off, val);
    }
}
/// Minimal developer-linked shape, intentionally independent of the writer.
fn linked_fixture() -> Vec<u8> {
    let mut b = vec![0; 3072];
    b[..2].copy_from_slice(b"MZ");
    p32(&mut b, 0x3c, 128);
    b[128..132].copy_from_slice(b"PE\0\0");
    p16(&mut b, 132, 0x8664);
    p16(&mut b, 134, 4);
    p16(&mut b, 148, 240);
    p16(&mut b, 150, 0x2022);
    let o = 152;
    p16(&mut b, o, 0x20b);
    p64(&mut b, o + 24, 0x180000000);
    p32(&mut b, o + 32, 4096);
    p32(&mut b, o + 36, 512);
    p32(&mut b, o + 56, 20480);
    p32(&mut b, o + 60, 1024);
    p32(&mut b, o + 108, 16);
    p16(&mut b, o + 68, 3);
    p16(&mut b, o + 70, 0x160);
    p32(&mut b, o + 112, 8192);
    p32(&mut b, o + 116, 128);
    p32(&mut b, o + 112 + 5 * 8, 16384);
    p32(&mut b, o + 116 + 5 * 8, 12);
    section(&mut b, 392, b".text", 4096, 1024, 0x60000020);
    section(&mut b, 432, b".rdata", 8192, 1536, 0x40000040);
    section(&mut b, 472, b".data", 12288, 2048, 0xc0000040);
    section(&mut b, 512, b".reloc", 16384, 2560, 0x42000040);
    b[1024..1027].copy_from_slice(&[0x89, 0xc8, 0xc3]); // mov eax,ecx; ret
    let e = 1536;
    p32(&mut b, e + 12, 8192 + 64);
    p32(&mut b, e + 16, 1);
    p32(&mut b, e + 20, 2);
    p32(&mut b, e + 24, 2);
    p32(&mut b, e + 28, 8192 + 40);
    p32(&mut b, e + 32, 8192 + 48);
    p32(&mut b, e + 36, 8192 + 56);
    p32(&mut b, e + 40, 4096);
    p32(&mut b, e + 44, 8192 + 160);
    p32(&mut b, e + 48, 8192 + 80);
    p32(&mut b, e + 52, 8192 + 100);
    p16(&mut b, e + 58, 1);
    b[e + 64..e + 69].copy_from_slice(b"t.dll");
    b[e + 80..e + 96].copy_from_slice(b"cvb_identity_i32");
    b[e + 100..e + 116].copy_from_slice(b"cvb_template_abi");
    p64(&mut b, e + 160, HELPER_ABI);
    p64(&mut b, 2048, 0x180001000);
    p32(&mut b, 2560, 12288);
    p32(&mut b, 2564, 12);
    p16(&mut b, 2568, 0xa000);
    b
}
fn finished() -> Vec<u8> {
    finish_template(
        &linked_fixture(),
        &[HelperSpec {
            name: "cvb_identity_i32",
            version: 1,
        }],
    )
    .unwrap()
}
fn program() -> Program {
    Program {
        text: vec![
            0xb8, 42, 0, 0, 0, 0xc3, 0x90, 0x90, 0x48, 0x83, 0xec, 0x28, 0xff, 0xd1, 0x48, 0x83,
            0xc4, 0x28, 0xc3,
        ],
        exports: vec![
            Export {
                name: "Zulu".into(),
                ordinal: Some(7),
                offset: 0,
            },
            Export {
                name: "Alpha".into(),
                ordinal: Some(2),
                offset: 8,
            },
        ],
        unwind: vec![Unwind {
            begin: 8,
            end: 19,
            info: vec![1, 4, 1, 0, 4, 0x42, 0, 0],
        }],
        image: b"test image".to_vec(),
        ..Program::default()
    }
}
#[test]
fn independent_reader_checks_exports_sections_and_exception_table() {
    let bytes = finished();
    let template = Template::parse(&bytes, HELPER_ABI).unwrap();
    assert_eq!(
        template.manifest.resolve("cvb_identity_i32", 1).unwrap(),
        4096
    );
    assert!(
        Image::parse(&bytes)
            .unwrap()
            .named_exports()
            .unwrap()
            .is_empty()
    );
    let patched = template.patch("test.dll", &program()).unwrap();
    let file = object::File::parse(patched.bytes.as_slice()).unwrap();
    let exports = file.exports().unwrap();
    assert_eq!(
        exports
            .iter()
            .map(|e| std::str::from_utf8(e.name()).unwrap())
            .collect::<Vec<_>>(),
        vec!["Alpha", "Zulu"]
    );
    assert_eq!(
        exports[0].address(),
        0x180000000 + u64::from(template.text_rva().unwrap()) + 8
    );
    assert!(
        file.section_by_name(".cvbtext").unwrap().flags()
            == object::SectionFlags::Coff {
                characteristics: 0x60000020
            }
    );
    assert_eq!(
        &file.section_by_name(".cvbimg").unwrap().data().unwrap()[..10],
        b"test image"
    );
    let image = Image::parse(&patched.bytes).unwrap();
    let runtime = image.runtime_functions().unwrap();
    assert_eq!(runtime.len(), 1);
    assert_eq!(runtime[0][0], template.text_rva().unwrap() + 8);
    assert_eq!(
        image.at_rva(runtime[0][2], 8).unwrap(),
        &[1, 4, 1, 0, 4, 0x42, 0, 0]
    );
    assert!(image.relocations().unwrap().contains(&12288));
    assert_eq!(image.sections.last().unwrap().name, ".reloc");
    // Existing absolute pointer and all original executable code stay fixed.
    assert_eq!(
        image.at_rva(12288, 8).unwrap(),
        &0x180001000u64.to_le_bytes()
    );
    assert_eq!(image.at_rva(4096, 3).unwrap(), &[0x89, 0xc8, 0xc3]);
}
#[test]
fn preserves_template_unwind_and_adds_generated_dir64() {
    let mut bytes = linked_fixture();
    p32(&mut bytes, 152 + 112 + 3 * 8, 8192 + 128);
    p32(&mut bytes, 152 + 116 + 3 * 8, 12);
    p32(&mut bytes, 1536 + 128, 4096);
    p32(&mut bytes, 1536 + 132, 4099);
    p32(&mut bytes, 1536 + 136, 8192 + 144);
    bytes[1536 + 144..1536 + 148].copy_from_slice(&[1, 0, 0, 0]);
    let bytes = finish_template(
        &bytes,
        &[HelperSpec {
            name: "cvb_identity_i32",
            version: 1,
        }],
    )
    .unwrap();
    let template = Template::parse(&bytes, HELPER_ABI).unwrap();
    let mut p = program();
    p.data = 0x180001000u64.to_le_bytes().to_vec();
    p.data_relocations = vec![0];
    let patched = template.patch("a.dll", &p).unwrap();
    let image = Image::parse(&patched.bytes).unwrap();
    assert_eq!(
        image.runtime_functions().unwrap()[0],
        [4096, 4099, 8192 + 144]
    );
    assert_eq!(image.runtime_functions().unwrap().len(), 2);
    assert_eq!(image.relocations().unwrap().len(), 2);
}
#[test]
fn refuses_malformed_and_incompatible_inputs() {
    let bytes = finished();
    assert!(Template::parse(&bytes, HELPER_ABI + 1).is_err());
    let template = Template::parse(&bytes, HELPER_ABI).unwrap();
    assert!(template.manifest.resolve("cvb_identity_i32", 2).is_err());
    let mut p = program();
    p.exports[1].name = p.exports[0].name.clone();
    assert!(template.patch("a.dll", &p).is_err());
    let mut p = program();
    p.exports[1].ordinal = Some(7);
    assert!(template.patch("a.dll", &p).is_err());
    let mut p = program();
    p.exports[0].offset = 100;
    assert!(template.patch("a.dll", &p).is_err());
    let mut p = program();
    p.unwind[0].end = 100;
    assert!(template.patch("a.dll", &p).is_err());
    let mut p = program();
    p.unwind[0].info.pop();
    assert!(template.patch("a.dll", &p).is_err());
    let mut p = program();
    p.data_relocations.push(0);
    assert!(template.patch("a.dll", &p).is_err());
    for at in [0, 128, 152, 200, 500, bytes.len() - 1] {
        assert!(Image::parse(&bytes[..at]).is_err());
    }
    let mut bad = linked_fixture();
    p16(&mut bad, 152 + 70, 0x4160);
    assert!(finish_template(&bad, &[]).is_err());
    let mut bad = linked_fixture();
    p32(&mut bad, 152 + 112 + 4 * 8, 1024);
    p32(&mut bad, 152 + 116 + 4 * 8, 8);
    assert!(finish_template(&bad, &[]).is_err());
    let mut bad = linked_fixture();
    bad[2569] = 0x30;
    assert!(finish_template(&bad, &[]).is_err());
    let mut bad = linked_fixture();
    bad[1536 + 160] ^= 1;
    assert!(finish_template(&bad, &[]).is_err());
    // /GUARD:NO CRTs can retain only the inactive instrumentation marker.
    let mut crt = linked_fixture();
    p32(&mut crt, 152 + 112 + 10 * 8, 8192 + 256);
    p32(&mut crt, 152 + 116 + 10 * 8, 160);
    p32(&mut crt, 1536 + 256 + 144, 0x100);
    assert!(finish_template(&crt, &[]).is_ok());
    p64(&mut crt, 1536 + 256 + 128, 0x180001000);
    assert!(finish_template(&crt, &[]).is_err());
}

#[cfg(windows)]
#[test]
fn named_exports_load_rebase_and_unwind_through_nonleaf_generated_frame() {
    use std::sync::atomic::{AtomicUsize, Ordering};
    use windows_sys::Win32::{
        Foundation::{FreeLibrary, GetLastError},
        System::{
            Diagnostics::Debug::{
                CheckSumMappedFile, RtlCaptureStackBackTrace, RtlLookupFunctionEntry,
            },
            LibraryLoader::{GetProcAddress, LoadLibraryW},
            Memory::{MEM_RELEASE, MEM_RESERVE, PAGE_NOACCESS, VirtualAlloc, VirtualFree},
        },
    };
    static FRAME: AtomicUsize = AtomicUsize::new(0);
    static OBSERVED: AtomicUsize = AtomicUsize::new(0);
    extern "C" fn capture() -> i32 {
        let mut frames = [std::ptr::null_mut(); 64];
        // SAFETY: frames has capacity for the requested 64 stack addresses.
        let count =
            unsafe { RtlCaptureStackBackTrace(0, 64, frames.as_mut_ptr(), std::ptr::null_mut()) };
        let begin = FRAME.load(Ordering::Relaxed);
        OBSERVED.store(
            usize::from(
                frames[..usize::from(count)]
                    .iter()
                    .any(|p| (*p as usize) >= begin && (*p as usize) < begin + 11),
            ),
            Ordering::Relaxed,
        );
        42
    }
    for bytes in [
        finished(),
        include_bytes!("../../../assets/native/calcvba-native-x64.dll").to_vec(),
    ] {
        let template = Template::parse(&bytes, HELPER_ABI).unwrap();
        let preferred = Image::parse(&bytes).unwrap().preferred_base;
        let mut p = program();
        p.data = (preferred + u64::from(template.text_rva().unwrap()))
            .to_le_bytes()
            .to_vec();
        p.data_relocations.push(0);
        let patched = template.patch("load.dll", &p).unwrap();
        let mut header_checksum = 0;
        let mut calculated_checksum = 0;
        // SAFETY: this complete owned byte slice remains live for the call.
        let checksum_header = unsafe {
            CheckSumMappedFile(
                patched.bytes.as_ptr().cast(),
                patched.bytes.len() as u32,
                &mut header_checksum,
                &mut calculated_checksum,
            )
        };
        assert!(!checksum_header.is_null());
        assert_eq!(
            header_checksum, calculated_checksum,
            "Windows checksum differs from the patcher"
        );
        let independent = object::File::parse(patched.bytes.as_slice()).unwrap();
        assert_eq!(independent.exports().unwrap().len(), 2);
        let patched_image = Image::parse(&patched.bytes).unwrap();
        let data = patched_image
            .sections
            .iter()
            .find(|s| s.name == ".cvbdata")
            .unwrap()
            .rva;
        let dir = std::env::temp_dir().join(format!("calcvba-pe-{}", std::process::id()));
        std::fs::create_dir_all(&dir).unwrap();
        let path = dir.join("load.dll");
        std::fs::write(&path, &patched.bytes).unwrap();
        let wide: Vec<u16> = path
            .as_os_str()
            .to_string_lossy()
            .encode_utf16()
            .chain(Some(0))
            .collect();
        // SAFETY: reserve an otherwise free preferred image range to force ASLR.
        let reservation = unsafe {
            VirtualAlloc(
                preferred as usize as *const _,
                0x100000,
                MEM_RESERVE,
                PAGE_NOACCESS,
            )
        };
        // SAFETY: path is NUL-terminated; this fixture contains only owned code.
        let module = unsafe { LoadLibraryW(wide.as_ptr()) };
        // SAFETY: GetLastError has no pointer preconditions.
        assert!(!module.is_null(), "LoadLibraryW: {}", unsafe {
            GetLastError()
        });
        assert_ne!(module as u64, preferred);
        // SAFETY: exact named signatures emitted above, retained until FreeLibrary.
        unsafe {
            let return42: extern "C" fn() -> i32 =
                std::mem::transmute(GetProcAddress(module, c"Zulu".as_ptr().cast()).unwrap());
            assert_eq!(return42(), 42);
            let walk: extern "C" fn(extern "C" fn() -> i32) -> i32 =
                std::mem::transmute(GetProcAddress(module, c"Alpha".as_ptr().cast()).unwrap());
            let address = walk as usize;
            FRAME.store(address, Ordering::Relaxed);
            let mut base = 0u64;
            let entry = RtlLookupFunctionEntry(address as u64, &mut base, std::ptr::null_mut());
            assert!(!entry.is_null());
            assert_eq!(base, module as u64);
            assert_eq!(walk(capture), 42);
            assert_eq!(
                OBSERVED.load(Ordering::Relaxed),
                1,
                "stack walk did not traverse the generated frame"
            );
            let pointer = (module as usize + data as usize) as *const usize;
            assert_eq!(
                *pointer,
                module as usize + template.text_rva().unwrap() as usize,
                "base relocation was not applied"
            );
            FreeLibrary(module);
            if !reservation.is_null() {
                VirtualFree(reservation, 0, MEM_RELEASE);
            }
        }
        std::fs::remove_file(path).unwrap();
        std::fs::remove_dir(dir).unwrap();
    }
}
