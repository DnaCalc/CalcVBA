//! Attach only to the HWND/PID recorded by the owned Excel oracle supervisor.
pub struct ComScope;
impl Drop for ComScope {
    fn drop(&mut self) {
        // SAFETY: balances this test thread's successful STA initialization.
        unsafe {
            windows_sys::Win32::System::Com::CoUninitialize();
        }
    }
}
pub fn attach() -> (ComScope, *mut std::ffi::c_void) {
    use windows_sys::Win32::{
        System::Com::{COINIT_APARTMENTTHREADED, CoInitializeEx},
        UI::{
            Accessibility::AccessibleObjectFromWindow,
            WindowsAndMessaging::{EnumChildWindows, GetClassNameW, GetWindowThreadProcessId},
        },
    };
    // SAFETY: initialize the current test thread; no pointer argument is retained.
    let initialized = unsafe { CoInitializeEx(std::ptr::null(), COINIT_APARTMENTTHREADED as u32) };
    assert!(
        initialized >= 0,
        "COM initialization failed: {initialized:08x}"
    );
    let scope = ComScope;
    let expected_pid: u32 = std::env::var("OXVBA_OWNED_EXCEL_PID")
        .unwrap()
        .parse()
        .unwrap();
    let hwnd_number: i32 = std::env::var("OXVBA_OWNED_EXCEL_HWND")
        .unwrap()
        .parse()
        .unwrap();
    assert_ne!(expected_pid, 0);
    assert_ne!(expected_pid, 3144);
    assert_ne!(hwnd_number, 0);
    let hwnd = hwnd_number as u32 as usize as windows_sys::Win32::Foundation::HWND;
    let mut actual = 0;
    // SAFETY: read only; actual is one aligned writable process-ID cell.
    assert_ne!(unsafe { GetWindowThreadProcessId(hwnd, &mut actual) }, 0);
    assert_eq!(actual, expected_pid);
    struct Search {
        process: u32,
        window: windows_sys::Win32::Foundation::HWND,
    }
    unsafe extern "system" fn find(
        hwnd: windows_sys::Win32::Foundation::HWND,
        parameter: isize,
    ) -> i32 {
        // SAFETY: synchronous EnumChildWindows retains this exact live Search.
        let search = unsafe { &mut *(parameter as *mut Search) };
        let mut process = 0;
        let mut class = [0u16; 64];
        // SAFETY: read-only window queries, initialized aligned output buffers.
        let len = unsafe {
            GetWindowThreadProcessId(hwnd, &mut process);
            GetClassNameW(hwnd, class.as_mut_ptr(), class.len() as i32)
        };
        if process == search.process
            && len == 6
            && class[..6] == [69, 88, 67, 69, 76, 55]
            && search.window.is_null()
        {
            search.window = hwnd;
        }
        1
    }
    let mut search = Search {
        process: expected_pid,
        window: std::ptr::null_mut(),
    };
    // SAFETY: enumerate only children of the checked owner window; Search
    // remains a unique live stack object until the callback walk returns.
    unsafe {
        EnumChildWindows(hwnd, Some(find), (&mut search as *mut Search) as isize);
    }
    assert!(!search.window.is_null(), "owned EXCEL7 child must be live");
    let iid = windows_sys::core::GUID {
        data1: 0x00020400,
        data2: 0,
        data3: 0,
        data4: [0xc0, 0, 0, 0, 0, 0, 0, 0x46],
    };
    let mut pointer = std::ptr::null_mut();
    // SAFETY: query the enumerated owned EXCEL7 public NativeOM; IID/output
    // cells remain live. A successful result transfers one IDispatch ref.
    let hr = unsafe { AccessibleObjectFromWindow(search.window, 0xfffffff0, &iid, &mut pointer) };
    assert!(
        hr >= 0 && !pointer.is_null(),
        "owned NativeOM failed: {hr:08x}"
    );
    (scope, pointer)
}
