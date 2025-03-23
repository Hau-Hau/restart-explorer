use std::{
    ptr::null_mut,
    sync::{Arc, Mutex},
    thread,
    time::{Duration, Instant},
};

use windows::{
    Win32::{
        Foundation::{ERROR_TIMEOUT, HWND, S_FALSE, WIN32_ERROR},
        System::Variant::VT_DISPATCH,
        UI::{
            Shell::IShellBrowser,
            WindowsAndMessaging::{FLASHW_STOP, FLASHWINFO, FlashWindowEx, GW_HWNDNEXT},
        },
    },
    core::{IUnknown, Interface, VARIANT},
};

use crate::infrastructure::windows_os::{
    enum_variant::EnumVariant, shell_windows::ShellWindows, windows_api::WindowApi,
};

use super::shell_view::get_path_from_shell_view;

pub fn get_topmost_window<TWindowApi: WindowApi>(hwnd: &HWND, window_api: &TWindowApi) -> HWND {
    let mut topmost_hwnd = *hwnd;
    loop {
        let handle = window_api.get_parent(topmost_hwnd);
        match handle {
            Ok(x) => topmost_hwnd = x,
            Err(_) => break,
        }
    }
    topmost_hwnd
}

pub fn wait_for_window_stable<TWindowApi: WindowApi>(
    location: &str,
    timeout: Duration,
    already_open_explorer_windows: &Arc<Mutex<Vec<isize>>>,
    window_api: &TWindowApi,
) -> Result<isize, windows::core::Error> {
    let start = Instant::now();
    let mut id = 0;
    while id == 0 {
        if start.elapsed() > timeout {
            return Err(windows::core::Error::from(WIN32_ERROR(ERROR_TIMEOUT.0)));
        }

        let shell_windows = window_api.create_shell_windows()?;
        let enum_variant = shell_windows.new_enum_variant()?;
        loop {
            let mut fetched = 0;
            let mut var = [VARIANT::default(); 1];
            let hr = enum_variant.next(&mut var, &mut fetched);
            if hr == S_FALSE || fetched == 0 {
                thread::sleep(Duration::from_millis(100));
                break;
            }

            if unsafe { var[0].as_raw().Anonymous.Anonymous.vt } != VT_DISPATCH.0 {
                thread::sleep(Duration::from_millis(100));
                continue;
            }

            let browser: IShellBrowser = unsafe {
                windows::Win32::UI::Shell::IUnknown_QueryService(
                    IUnknown::from_raw_borrowed(
                        &var[0].as_raw().Anonymous.Anonymous.Anonymous.pdispVal,
                    ),
                    &windows::Win32::UI::Shell::SID_STopLevelBrowser,
                )
            }?;

            let shell_view = unsafe { browser.QueryActiveShellView() }?;
            let path = match get_path_from_shell_view(&shell_view) {
                Ok(path) => path,
                Err(_) => {
                    thread::sleep(Duration::from_millis(100));
                    continue;
                }
            };

            if path != location {
                thread::sleep(Duration::from_millis(100));
                continue;
            }

            let hwnd = unsafe { shell_view.GetWindow()? };
            let topmost_parent = get_topmost_window(&hwnd, window_api);

            let temp_id = topmost_parent.0 as isize;

            let already_open_ids = {
                let guard = already_open_explorer_windows.lock().unwrap();
                guard.clone() // Make a temporary copy to check against
            };

            if already_open_ids.contains(&temp_id) {
                thread::sleep(Duration::from_millis(100));
                continue;
            }

            id = temp_id;
        }
        if id != 0 {
            break;
        }
    }

    Ok(id)
}

pub fn get_window_z_index<TWindowApi: WindowApi>(
    hwnd: HWND,
    window_api: &TWindowApi,
) -> windows_core::Result<i32> {
    let mut z_index = 0;
    let mut current = window_api.get_top_window(HWND::default())?;

    while current.0 != null_mut() {
        if current == hwnd {
            break;
        }
        z_index += 1;
        current = window_api.get_window(current, GW_HWNDNEXT)?;
    }

    Ok(z_index)
}

pub fn stop_window_flashing(hwnd: HWND) {
    unsafe {
        let flash_info = FLASHWINFO {
            cbSize: std::mem::size_of::<FLASHWINFO>() as u32,
            hwnd,
            dwFlags: FLASHW_STOP,
            uCount: 0,
            dwTimeout: 0,
        };
        let _ = FlashWindowEx(&flash_info);
    }
}
