use std::{
    borrow::{Borrow, BorrowMut},
    ffi::{c_void, CString, OsStr, OsString},
    fmt::Debug,
    mem::MaybeUninit,
    os::windows::ffi::OsStringExt,
    process,
    ptr::{self, addr_of_mut},
};

use windows::{
    core::{Error, IUnknown, Interface, Param, Result, PROPVARIANT, VARIANT},
    Win32::{
        Foundation::{CloseHandle, HANDLE, HWND, S_FALSE, VARIANT_BOOL},
        System::{
            Com::{
                CoCreateInstance, CoInitializeEx, CoTaskMemFree, CLSCTX_LOCAL_SERVER,
                COINIT_APARTMENTTHREADED, COINIT_DISABLE_OLE1DDE,
            },
            Diagnostics::ToolHelp::{
                CreateToolhelp32Snapshot, Process32First, Process32Next,
                CREATE_TOOLHELP_SNAPSHOT_FLAGS, PROCESSENTRY32, TH32CS_SNAPALL,
            },
            Ole::{IEnumVARIANT, SAFEARR_VARIANT},
            Threading::{OpenProcess, TerminateProcess, PROCESS_TERMINATE},
            Variant::VT_DISPATCH,
        },
        UI::Shell::{
            IPersistIDList, IShellBrowser, IShellItem, IShellWindows, IUnknown_QueryService,
            SHCreateItemFromIDList, SID_STopLevelBrowser, ShellWindows,
            SIGDN_DESKTOPABSOLUTEPARSING,
        },
    },
};

fn main() -> Result<()> {
    unsafe {
        CoInitializeEx(
            Option::Some(ptr::null()),
            COINIT_APARTMENTTHREADED | COINIT_DISABLE_OLE1DDE,
        )
    };

    let windows: IShellWindows =
        unsafe { CoCreateInstance(&ShellWindows, None, CLSCTX_LOCAL_SERVER) }?;
    let locations = get_explorer_locations(&windows)?;
    Ok(())
}

fn get_explorer_locations(windows: &IShellWindows) -> Result<Vec<String>> {
    let unk_enum = unsafe { windows._NewEnum() }?;
    let enum_variant = unk_enum.cast::<IEnumVARIANT>()?;

    let mut locations = vec![];
    loop {
        let mut fetched = 0;
        let mut var = [VARIANT::default(); 1];
        let hr = unsafe { enum_variant.Next(&mut var, &mut fetched) };
        // No more windows?
        if hr == S_FALSE || fetched == 0 {
            break;
        }

        // Not an IDispatch interface?
        if unsafe { var[0].clone().as_raw().Anonymous.Anonymous.vt } != VT_DISPATCH.0 as _ {
            continue;
        }

        // let y = unsafe {
        //     std::mem::transmute::<*mut std::ffi::c_void, IUnknown>(
        //         var[0]
        //             .clone()
        //             .as_raw()
        //             .Anonymous
        //             .Anonymous
        //             .Anonymous
        //             .pdispVal,
        //     )
        // };
        // let x = unsafe {
        //     IUnknown::from_raw(
        //         var[0]
        //             .clone()
        //             .as_raw()
        //             .Anonymous
        //             .Anonymous
        //             .Anonymous
        //             .pdispVal,
        //     )
        // };

        // Get the information
        let mut hwnd = Default::default();
        let location = get_browser_info(
            unsafe {
                var[0]
                    .clone()
                    .as_raw()
                    .Anonymous
                    .Anonymous
                    .Anonymous
                    .pdispVal
            },
            &mut hwnd,
        )?;

        // Convert UTF-16 to UTF-8 for display
        let location = String::from_utf16_lossy(&location);
        locations.push(location);
    }

    Ok(locations)
}

// fn get_browser_info(unk: *mut c_void) -> Result<Vec<u16>> {
//     let shell_browser: IShellBrowser =
//         unsafe { IUnknown_QueryService(&IUnknown::from_raw(unk), &SID_STopLevelBrowser) }?;
//     // *hwnd = unsafe { shell_browser.GetWindow() }?;

//     let output = get_location_from_view(&shell_browser);
//     unsafe { CoTaskMemFree(Option::Some(unk)) };
//     output
// }

fn get_browser_info(unk: *mut c_void, hwnd: &mut HWND) -> Result<Vec<u16>> {
    let shell_browser: IShellBrowser = unsafe {
        IUnknown_QueryService(
            &std::mem::transmute::<*mut std::ffi::c_void, IUnknown>(unk),
            &SID_STopLevelBrowser,
        )
    }?;
    *hwnd = unsafe { shell_browser.GetWindow() }?;
    get_location_from_view(&shell_browser)
}

fn get_location_from_view(browser: &IShellBrowser) -> Result<Vec<u16>> {
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
