//! VBA's Windows settings are values under the current user's VBA settings key.
use std::{io, ptr};
use windows_sys::Win32::{
    Foundation::{
        ERROR_FILE_NOT_FOUND, ERROR_MORE_DATA, ERROR_NO_MORE_ITEMS, ERROR_PATH_NOT_FOUND,
        ERROR_SUCCESS,
    },
    System::Registry::{
        HKEY, HKEY_CURRENT_USER, KEY_ENUMERATE_SUB_KEYS, KEY_QUERY_VALUE, KEY_SET_VALUE,
        REG_OPTION_NON_VOLATILE, REG_SZ, RRF_RT_REG_SZ, RegCloseKey, RegCreateKeyExW,
        RegDeleteKeyW, RegDeleteTreeW, RegDeleteValueW, RegEnumKeyExW, RegEnumValueW, RegGetValueW,
        RegOpenKeyExW, RegSetValueExW,
    },
};

struct Key(HKEY);

impl Drop for Key {
    fn drop(&mut self) {
        // SAFETY: this owned handle was returned by a successful open or create.
        unsafe { RegCloseKey(self.0) };
    }
}

fn wide(text: &str) -> Vec<u16> {
    text.encode_utf16().chain([0]).collect()
}

fn setting_path(app: &str, section: Option<&str>) -> io::Result<Vec<u16>> {
    let valid = |name: &str| !name.contains('\0') && name.split('\\').all(|part| !part.is_empty());
    if !valid(app) || section.is_some_and(|section| !valid(section)) {
        return Err(io::Error::new(
            io::ErrorKind::InvalidInput,
            "invalid VBA settings key path",
        ));
    }
    let mut path = format!("Software\\VB and VBA Program Settings\\{app}");
    if let Some(section) = section {
        path.push('\\');
        path.push_str(section);
    }
    Ok(wide(&path))
}

fn status(status: u32) -> io::Result<()> {
    if status == ERROR_SUCCESS {
        Ok(())
    } else {
        Err(io::Error::from_raw_os_error(status as i32))
    }
}

fn open(path: &[u16], access: u32) -> io::Result<Option<Key>> {
    let mut key = ptr::null_mut();
    // SAFETY: the subkey is terminated and Windows writes the handle into `key`.
    let result = unsafe { RegOpenKeyExW(HKEY_CURRENT_USER, path.as_ptr(), 0, access, &mut key) };
    match result {
        ERROR_SUCCESS => Ok(Some(Key(key))),
        ERROR_FILE_NOT_FOUND | ERROR_PATH_NOT_FOUND => Ok(None),
        other => Err(io::Error::from_raw_os_error(other as i32)),
    }
}

fn read_value(key: &Key, name: &[u16]) -> io::Result<Option<String>> {
    let mut bytes = 0u32;
    loop {
        let mut data = vec![0u16; (bytes as usize).div_ceil(2)];
        let buffer = if data.is_empty() {
            ptr::null_mut()
        } else {
            data.as_mut_ptr().cast()
        };
        // SAFETY: the key is live and name terminated. A null buffer queries the
        // size; otherwise data has room for the bytes supplied in `bytes`.
        let result = unsafe {
            RegGetValueW(
                key.0,
                ptr::null(),
                name.as_ptr(),
                RRF_RT_REG_SZ,
                ptr::null_mut(),
                buffer,
                &mut bytes,
            )
        };
        match result {
            ERROR_FILE_NOT_FOUND => return Ok(None),
            ERROR_MORE_DATA => continue,
            ERROR_SUCCESS if data.is_empty() && bytes != 0 => continue,
            ERROR_SUCCESS => {
                let units = &data[..bytes as usize / 2];
                return Ok(Some(String::from_utf16_lossy(
                    units.strip_suffix(&[0]).unwrap_or(units),
                )));
            }
            other => return Err(io::Error::from_raw_os_error(other as i32)),
        }
    }
}

pub(super) fn get(app: &str, section: &str, name: &str) -> io::Result<Option<String>> {
    let Some(key) = open(&setting_path(app, Some(section))?, KEY_QUERY_VALUE)? else {
        return Ok(None);
    };
    read_value(&key, &wide(name))
}

pub(super) fn all(app: &str, section: &str) -> io::Result<Vec<(String, String)>> {
    let Some(key) = open(&setting_path(app, Some(section))?, KEY_QUERY_VALUE)? else {
        return Ok(Vec::new());
    };
    let mut entries = Vec::new();
    let mut name = vec![0u16; 16_384];
    for index in 0.. {
        let mut len = name.len() as u32;
        // SAFETY: the writable name buffer has `len` units; optional pointers are null.
        let result = unsafe {
            RegEnumValueW(
                key.0,
                index,
                name.as_mut_ptr(),
                &mut len,
                ptr::null(),
                ptr::null_mut(),
                ptr::null_mut(),
                ptr::null_mut(),
            )
        };
        if result == ERROR_NO_MORE_ITEMS {
            break;
        }
        status(result)?;
        if let Some(value) = read_value(&key, &name[..=len as usize])? {
            entries.push((String::from_utf16_lossy(&name[..len as usize]), value));
        }
    }
    Ok(entries)
}

pub(super) fn save(app: &str, section: &str, name: &str, value: &str) -> io::Result<()> {
    if name.is_empty() || name.contains('\0') {
        return Err(io::Error::new(
            io::ErrorKind::InvalidInput,
            "invalid VBA setting name",
        ));
    }
    let path = setting_path(app, Some(section))?;
    let mut key = ptr::null_mut();
    // SAFETY: path is terminated and the output pointer is a live HKEY slot.
    status(unsafe {
        RegCreateKeyExW(
            HKEY_CURRENT_USER,
            path.as_ptr(),
            0,
            ptr::null(),
            REG_OPTION_NON_VOLATILE,
            KEY_SET_VALUE,
            ptr::null(),
            &mut key,
            ptr::null_mut(),
        )
    })?;
    let key = Key(key);
    let name = wide(name);
    let data = wide(value);
    let bytes = u32::try_from(std::mem::size_of_val(data.as_slice()))
        .map_err(|_| io::Error::new(io::ErrorKind::InvalidData, "setting exceeds registry size"))?;
    // SAFETY: the data allocation is live for `bytes` bytes and is terminated UTF-16.
    status(unsafe { RegSetValueExW(key.0, name.as_ptr(), 0, REG_SZ, data.as_ptr().cast(), bytes) })
}

pub(super) fn delete(app: &str, section: Option<&str>, value: Option<&str>) -> io::Result<()> {
    let path = setting_path(app, section)?;
    if let Some(value) = value {
        if section.is_none() {
            return Err(io::Error::from(io::ErrorKind::InvalidInput));
        }
        let key =
            open(&path, KEY_SET_VALUE)?.ok_or_else(|| io::Error::from(io::ErrorKind::NotFound))?;
        let value = wide(value);
        // SAFETY: the value name is terminated and the owning section key is live.
        status(unsafe { RegDeleteValueW(key.0, value.as_ptr()) })
    } else if section.is_some() {
        // SAFETY: the path is a nonempty owned string below the VBA settings key.
        status(unsafe { RegDeleteTreeW(HKEY_CURRENT_USER, path.as_ptr()) })
    } else {
        delete_app(&path)
    }
}

fn delete_app(path: &[u16]) -> io::Result<()> {
    let app = open(path, KEY_ENUMERATE_SUB_KEYS)?
        .ok_or_else(|| io::Error::from(io::ErrorKind::NotFound))?;
    // A registry key name is at most 255 UTF-16 units, plus the terminator.
    let mut name = [0u16; 256];
    loop {
        let mut len = name.len() as u32;
        // SAFETY: the app handle and writable name and length buffers are live.
        let result = unsafe {
            RegEnumKeyExW(
                app.0,
                0,
                name.as_mut_ptr(),
                &mut len,
                ptr::null(),
                ptr::null_mut(),
                ptr::null_mut(),
                ptr::null_mut(),
            )
        };
        if result == ERROR_NO_MORE_ITEMS {
            break;
        }
        status(result)?;
        // Excel stops at a section with subkeys after deleting the preceding sections.
        // RegDeleteKeyW deletes one leaf, so the next section will be at index zero.
        // SAFETY: the app is live and the successful enumeration returned a terminated name.
        status(unsafe { RegDeleteKeyW(app.0, name.as_ptr()) })?;
    }
    drop(app);
    // SAFETY: path is terminated and identifies the now-empty application under HKCU.
    status(unsafe { RegDeleteKeyW(HKEY_CURRENT_USER, path.as_ptr()) })
}
