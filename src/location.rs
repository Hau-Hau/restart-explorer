use std::{
    ffi::c_void,
    time::{Duration, Instant},
};

use windows::{
    core::{w, IUnknown, Interface, PCWSTR, VARIANT},
    Win32::{
        Foundation::{ERROR_TIMEOUT, RECT, S_FALSE, WIN32_ERROR},
        System::{
            Com::{CoCreateInstance, CoTaskMemFree, CLSCTX_LOCAL_SERVER},
            Ole::IEnumVARIANT,
            Variant::VT_DISPATCH,
        },
        UI::{
            Shell::{
                IPersistIDList, IShellBrowser, IShellItem, IShellWindows, SHCreateItemFromIDList,
                ShellExecuteW, ShellWindows, SIGDN_DESKTOPABSOLUTEPARSING,
            },
            WindowsAndMessaging::{
                GetParent, GetWindowRect, SetWindowPos, SWP_SHOWWINDOW, SW_SHOW,
            },
        },
    },
};

use crate::models::window::Window;

fn wait_for_explorer_window_stable(
    location: &str,
    timeout: Duration,
    already_open_explorer_windows: Vec<isize>,
) -> Result<isize, windows::core::Error> {
    let start = Instant::now();
    let mut id = 0;
    while id == 0 {
        if start.elapsed() > timeout {
            return Err(windows::core::Error::from(WIN32_ERROR(ERROR_TIMEOUT.0)));
        }

        let windows: IShellWindows =
            unsafe { CoCreateInstance(&ShellWindows, None, CLSCTX_LOCAL_SERVER) }?;
        let unk_enum = unsafe { windows._NewEnum() }?;
        let enum_variant = unk_enum.cast::<IEnumVARIANT>()?;
        loop {
            let mut fetched = 0;
            let mut var = [VARIANT::default(); 1];
            let hr = unsafe { enum_variant.Next(&mut var, &mut fetched) };
            if hr == S_FALSE || fetched == 0 {
                std::thread::sleep(Duration::from_millis(100));
                break;
            }

            if unsafe { var[0].as_raw().Anonymous.Anonymous.vt } != VT_DISPATCH.0 {
                std::thread::sleep(Duration::from_millis(100));
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
            let persist_id_list: IPersistIDList = shell_view.cast()?;
            let id_list = unsafe { persist_id_list.GetIDList() }?;

            let item = unsafe { SHCreateItemFromIDList::<IShellItem>(id_list) }?;
            let ptr = unsafe { item.GetDisplayName(SIGDN_DESKTOPABSOLUTEPARSING) }?;

            let path: Vec<u16> = unsafe {
                let mut len = 0;
                while (*ptr.0.add(len)) != 0 {
                    len += 1;
                }
                std::slice::from_raw_parts(ptr.0, len + 1)
            }
            .to_vec();

            unsafe { CoTaskMemFree(Option::Some(ptr.0 as _)) };
            unsafe { CoTaskMemFree(Option::Some(id_list as _)) };

            if String::from_utf16_lossy(&path) != location {
                std::thread::sleep(Duration::from_millis(100));
                continue;
            }

            let hwnd = unsafe { shell_view.GetWindow()? };
            let mut topmost_parent = hwnd;
            loop {
                let handle = unsafe { GetParent(topmost_parent) };
                match handle {
                    Ok(x) => topmost_parent = x,
                    Err(_) => break,
                }
                topmost_parent = handle.unwrap();
            }

            let temp_id = topmost_parent.0 as isize;
            if already_open_explorer_windows.contains(&temp_id) {
                std::thread::sleep(Duration::from_millis(100));
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

pub fn open_location(window: &Window, already_open_explorer_windows: Vec<isize>) -> Option<isize> {
    let location_utf16: Vec<u16> = window
        .location
        .encode_utf16()
        .chain(std::iter::once(0))
        .collect();
    unsafe {
        ShellExecuteW(
            None,
            w!("open"),
            w!("explorer.exe"),
            PCWSTR(location_utf16.as_ptr()),
            None,
            SW_SHOW,
        )
    };

    if let Ok(id) = wait_for_explorer_window_stable(
        &window.location,
        Duration::from_secs(10),
        already_open_explorer_windows,
    ) {
        let _ = adjust_window_position(&window.location, window.rect, id);
        return Some(id);
    }

    None
}

fn adjust_window_position(
    location: &str,
    rect: RECT,
    id: isize,
) -> Result<(), windows::core::Error> {
    let windows: IShellWindows =
        unsafe { CoCreateInstance(&ShellWindows, None, CLSCTX_LOCAL_SERVER) }?;

    let unk_enum = unsafe { windows._NewEnum() }?;
    let enum_variant = unk_enum.cast::<IEnumVARIANT>()?;

    loop {
        let mut fetched = 0;
        let mut var = [VARIANT::default(); 1];
        let hr = unsafe { enum_variant.Next(&mut var, &mut fetched) };
        if hr == S_FALSE || fetched == 0 {
            break;
        }

        if unsafe { var[0].as_raw().Anonymous.Anonymous.vt } != VT_DISPATCH.0 {
            continue;
        }

        let result = try_set_position(
            unsafe { var[0].as_raw().Anonymous.Anonymous.Anonymous.pdispVal },
            location,
            rect,
            id,
        );

        if let Ok(is_position_set) = result {
            if is_position_set {
                break;
            }
        }
    }

    Ok(())
}

fn try_set_position(
    unk: *mut c_void,
    location: &str,
    rect: RECT,
    id: isize,
) -> Result<bool, windows::core::Error> {
    let browser: IShellBrowser = unsafe {
        windows::Win32::UI::Shell::IUnknown_QueryService(
            IUnknown::from_raw_borrowed(&unk),
            &windows::Win32::UI::Shell::SID_STopLevelBrowser,
        )
    }?;

    let shell_view = unsafe { browser.QueryActiveShellView() }?;
    let persist_id_list: IPersistIDList = shell_view.cast()?;
    let id_list = unsafe { persist_id_list.GetIDList() }?;

    let item = unsafe { SHCreateItemFromIDList::<IShellItem>(id_list) }?;
    let ptr = unsafe { item.GetDisplayName(SIGDN_DESKTOPABSOLUTEPARSING) }?;
    let path: Vec<u16> = unsafe {
        let mut len = 0;
        while (*ptr.0.add(len)) != 0 {
            len += 1;
        }
        std::slice::from_raw_parts(ptr.0, len + 1)
    }
    .to_vec();

    unsafe { CoTaskMemFree(Option::Some(ptr.0 as _)) };
    unsafe { CoTaskMemFree(Option::Some(id_list as _)) };

    if String::from_utf16_lossy(&path) != location {
        return Ok(false);
    }

    let hwnd = unsafe { shell_view.GetWindow()? };
    let mut topmost_parent = hwnd;
    loop {
        let handle = unsafe { GetParent(topmost_parent) };
        match handle {
            Ok(x) => topmost_parent = x,
            Err(_) => break,
        }
        topmost_parent = handle.unwrap();
    }

    if (topmost_parent.0 as isize) != id {
        return Ok(false);
    }

    unsafe {
        SetWindowPos(
            topmost_parent,
            None,
            rect.left,
            rect.top,
            rect.right - rect.left,
            rect.bottom - rect.top,
            SWP_SHOWWINDOW,
        )
    }?;

    Ok(true)
}

pub fn get_explorer_windows() -> Vec<Window> {
    let mut windows = vec![];

    let shell_windows: windows::core::Result<IShellWindows> =
        unsafe { CoCreateInstance(&ShellWindows, None, CLSCTX_LOCAL_SERVER) };

    let windows_enum = match shell_windows {
        Ok(shell_windows) => {
            let unk_enum = unsafe { shell_windows._NewEnum() };
            match unk_enum {
                Ok(unk_enum) => unk_enum.cast::<IEnumVARIANT>(),
                Err(_) => return windows, // Return empty Vec on error
            }
        }
        Err(_) => return windows, // Return empty Vec on error
    };

    let enum_variant = match windows_enum {
        Ok(enum_variant) => enum_variant,
        Err(_) => return windows, // Return empty Vec on error
    };

    loop {
        let mut fetched = 0;
        let mut var = [VARIANT::default(); 1];
        let hr = unsafe { enum_variant.Next(&mut var, &mut fetched) };
        if hr == S_FALSE || fetched == 0 {
            break;
        }

        if unsafe { var[0].as_raw().Anonymous.Anonymous.vt } != VT_DISPATCH.0 {
            continue;
        }

        match get_window_from_view(unsafe {
            var[0].as_raw().Anonymous.Anonymous.Anonymous.pdispVal
        }) {
            Ok(window) => windows.push(window),
            Err(_) => continue, // Skip this window on error
        }
    }

    windows
}

fn get_window_from_view(unk: *mut c_void) -> Result<Window, windows::core::Error> {
    let browser: IShellBrowser = unsafe {
        windows::Win32::UI::Shell::IUnknown_QueryService(
            IUnknown::from_raw_borrowed(&unk),
            &windows::Win32::UI::Shell::SID_STopLevelBrowser,
        )
    }?;

    let shell_view = unsafe { browser.QueryActiveShellView() }?;
    let persist_id_list: IPersistIDList = shell_view.cast()?;
    let id_list = unsafe { persist_id_list.GetIDList() }?;

    let item = unsafe { SHCreateItemFromIDList::<IShellItem>(id_list) }?;
    let ptr = unsafe { item.GetDisplayName(SIGDN_DESKTOPABSOLUTEPARSING) }?;

    let path: Vec<u16> = unsafe {
        let mut len = 0;
        while (*ptr.0.add(len)) != 0 {
            len += 1;
        }
        std::slice::from_raw_parts(ptr.0, len + 1)
    }
    .to_vec();

    unsafe { CoTaskMemFree(Option::Some(ptr.0 as _)) };
    unsafe { CoTaskMemFree(Option::Some(id_list as _)) };

    let hwnd = unsafe { shell_view.GetWindow()? };
    let mut topmost_parent = hwnd;
    loop {
        let handle = unsafe { GetParent(topmost_parent) };
        match handle {
            Ok(x) => topmost_parent = x,
            Err(_) => break,
        }
        topmost_parent = handle.unwrap();
    }

    let mut rect = RECT::default();
    unsafe {
        let _ = GetWindowRect(topmost_parent, &mut rect);
    }

    Ok(Window {
        location: String::from_utf16_lossy(&path),
        rect,
    })
}
