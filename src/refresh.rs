use std::{
    ffi::OsStr, iter::once, os::windows::prelude::OsStrExt, ptr::null_mut, thread::sleep,
    time::Duration,
};

use windows::Win32::{
    Foundation::{CloseHandle, BOOL, HWND, LPARAM, WPARAM},
    System::Threading::{
        OpenProcess, QueryFullProcessImageNameW, PROCESS_NAME_FORMAT, PROCESS_QUERY_INFORMATION,
        PROCESS_VM_READ,
    },
    UI::WindowsAndMessaging::{
        EnumWindows, GetWindowThreadProcessId, SendMessageTimeoutW, SMTO_NORMAL,
        SMTO_NOTIMEOUTIFNOTHUNG, WM_SETTINGCHANGE, WM_THEMECHANGED,
    },
};

use crate::regkey::{RegistryKey, RegistryPermission};

fn os_str(value: &str) -> Vec<u16> {
    OsStr::new(value).encode_wide().chain(once(0)).collect()
}

pub fn get_process_name(hwnd: HWND) -> Option<String> {
    unsafe {
        let mut process_id = 0;

        if GetWindowThreadProcessId(hwnd, &mut process_id as *mut u32) == 0 {
            return None;
        }

        let Ok(process) = OpenProcess(
            PROCESS_QUERY_INFORMATION | PROCESS_VM_READ,
            false,
            process_id,
        ) else {
            return None;
        };

        let mut process_name = [0u16; 512];
        let mut length = process_name.len() as u32 - 1;

        let has_image_name = QueryFullProcessImageNameW(
            process,
            PROCESS_NAME_FORMAT(0),
            windows::core::PWSTR(&mut process_name as *mut u16),
            &mut length,
        )
        .as_bool();
        CloseHandle(process);

        if has_image_name {
            String::from_utf16(&process_name).ok()
        } else {
            None
        }
    }
}

unsafe extern "system" fn refresh_window_callback(hwnd: HWND, _: LPARAM) -> BOOL {
    // these processes are known to require refreshes
    let whitelist = ["explorer.exe"];

    let Some(process_name) = get_process_name(hwnd) else {
        return BOOL(1);
    };

    if whitelist.iter().any(|&w| process_name.contains(w)) {
        SendMessageTimeoutW(
            hwnd,
            WM_SETTINGCHANGE,
            WPARAM(0),
            LPARAM(os_str("ImmersiveColorSet").as_ptr() as isize),
            SMTO_NORMAL | SMTO_NOTIMEOUTIFNOTHUNG,
            200,
            null_mut(),
        );

        SendMessageTimeoutW(
            hwnd,
            WM_THEMECHANGED,
            WPARAM(0),
            LPARAM(0),
            SMTO_NORMAL | SMTO_NOTIMEOUTIFNOTHUNG,
            200,
            null_mut(),
        );
    }

    BOOL(1)
}

pub fn refresh_windows() {
    let key = RegistryKey::open_or_create(
        &RegistryKey::HKCU,
        "SOFTWARE\\Microsoft\\Windows\\DWM",
        RegistryPermission::ReadWrite,
    );

    // update accent color as a way to trigger apps that might listen to it
    let accent = key.get_dword("AccentColor");
    key.set_dword("AccentColor", accent + 1);
    sleep(Duration::from_millis(10));
    key.set_dword("AccentColor", accent);

    // refresh the windows
    unsafe {
        EnumWindows(Some(refresh_window_callback), LPARAM(1));
    }
}
