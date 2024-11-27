use std::ffi::c_void;

use windows::{
    core::{w, IUnknown, Interface, Result, PCWSTR, VARIANT},
    Win32::{
        Foundation::S_FALSE,
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
            WindowsAndMessaging::SW_SHOWMINIMIZED,
        },
    },
};

pub fn open_location(location: &str) {
    let location_utf16: Vec<u16> = location.encode_utf16().chain(std::iter::once(0)).collect();
    unsafe {
        ShellExecuteW(
            None,
            w!("open"),
            PCWSTR(location_utf16.as_ptr()),
            None,
            None,
            SW_SHOWMINIMIZED,
        )
    };
}

pub fn get_explorer_locations() -> Result<Vec<String>> {
    let windows: IShellWindows =
        unsafe { CoCreateInstance(&ShellWindows, None, CLSCTX_LOCAL_SERVER) }?;

    let unk_enum = unsafe { windows._NewEnum() }?;
    let enum_variant = unk_enum.cast::<IEnumVARIANT>()?;

    let mut locations = vec![];
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

        let location = get_location_from_view(unsafe {
            var[0].as_raw().Anonymous.Anonymous.Anonymous.pdispVal
        })?;

        let location = String::from_utf16_lossy(&location);
        locations.push(location);
    }

    Ok(locations)
}

fn get_location_from_view(unk: *mut c_void) -> Result<Vec<u16>> {
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

    // Copy UTF-16 string to `Vec<u16>` (including NUL terminator)
    let mut path = Vec::new();
    let mut p = ptr.0 as *const u16;
    loop {
        let ch = unsafe { *p };
        path.push(ch);
        if ch == 0 {
            break;
        }
        p = unsafe { p.add(1) };
    }

    // Cleanup
    unsafe { CoTaskMemFree(Option::Some(ptr.0 as _)) };
    unsafe { CoTaskMemFree(Option::Some(id_list as _)) };

    Ok(path)
}
