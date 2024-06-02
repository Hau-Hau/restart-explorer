#![windows_subsystem = "windows"]

use std::{
    ffi::{c_char, c_void, CStr},
    ptr,
};

use windows::{
    core::{ComInterface, IUnknown, Interface, Result, PCSTR},
    s,
    Win32::{
        Foundation::{CloseHandle, BOOL, HWND, S_FALSE},
        System::{
            Com::{
                CoCreateInstance, CoInitializeEx, CoTaskMemFree, CLSCTX_LOCAL_SERVER,
                COINIT_APARTMENTTHREADED, COINIT_DISABLE_OLE1DDE, VARIANT, VT_DISPATCH,
            },
            Diagnostics::ToolHelp::{
                CreateToolhelp32Snapshot, Process32First, Process32Next, PROCESSENTRY32,
                TH32CS_SNAPALL,
            },
            Ole::IEnumVARIANT,
            Threading::{OpenProcess, TerminateProcess, PROCESS_TERMINATE},
        },
        UI::{
            Shell::{
                IPersistIDList, IShellBrowser, IShellItem, IShellWindows, IUnknown_QueryService,
                SHCreateItemFromIDList, SID_STopLevelBrowser, ShellExecuteA, ShellWindows,
                SIGDN_DESKTOPABSOLUTEPARSING,
            },
            WindowsAndMessaging::{SW_NORMAL, SW_SHOWMINIMIZED},
        },
    },
};

fn main() -> Result<()> {
    unsafe {
        CoInitializeEx(
            Some(ptr::null()),
            COINIT_APARTMENTTHREADED | COINIT_DISABLE_OLE1DDE,
        )
    }?;

    let locations = get_explorer_locations()?;
    kill_process_by_name("explorer.exe");
    start_process("explorer.exe");
    for location in locations {
        open_location(&location);
    }
    Ok(())
}

fn open_location(location: &str) {
    unsafe {
        ShellExecuteA(
            None,
            s!("open"),
            PCSTR::from_raw(format!("{}\0", location).as_mut_ptr()),
            None,
            None,
            SW_SHOWMINIMIZED,
        )
    };
}

fn start_process(process_name: &str) {
    unsafe {
        ShellExecuteA(
            None,
            None,
            PCSTR::from_raw(format!("{}\0", process_name).as_mut_ptr()),
            None,
            None,
            SW_NORMAL,
        )
    };
}

fn kill_process_by_name(process_name: &str) {
    let snapshot =
        unsafe { CreateToolhelp32Snapshot(TH32CS_SNAPALL, 0).expect("Failed to create snapshot") };
    let mut entry = PROCESSENTRY32 {
        dwSize: std::mem::size_of::<PROCESSENTRY32>() as u32,
        ..Default::default()
    };
    let mut has_next_process = bool::from(unsafe { Process32First(snapshot, &mut entry) });
    while has_next_process {
        let current_process_name = unsafe {
            CStr::from_ptr(entry.szExeFile.as_ptr() as *const c_char)
                .to_str()
                .unwrap()
        };
        if current_process_name == process_name {
            kill_process(entry.th32ProcessID);
        }

        has_next_process = bool::from(unsafe { Process32Next(snapshot, &mut entry) });
    }

    unsafe { CloseHandle(snapshot) };
}

pub fn kill_process(process_id: u32) {
    unsafe {
        let process_handle = OpenProcess(PROCESS_TERMINATE, BOOL::from(false), process_id);
        if let Ok(process) = process_handle {
            TerminateProcess(process, 1);
            CloseHandle(process);
        }
    }
}

// Below code is based on https://stackoverflow.com/questions/73311644/get-path-to-selected-files-in-active-explorer-window

fn get_explorer_locations() -> Result<Vec<String>> {
    let windows: IShellWindows =
        unsafe { CoCreateInstance(&ShellWindows, None, CLSCTX_LOCAL_SERVER) }?;
    let unk_enum = unsafe { windows._NewEnum() }?;
    let enum_variant = unk_enum.cast::<IEnumVARIANT>()?;

    let mut locations = vec![];
    loop {
        let mut fetched = 0;
        let mut var: [VARIANT; 1] = [VARIANT::default(); 1];
        let hr = unsafe { enum_variant.Next(&mut var, &mut fetched) };
        if hr == S_FALSE || fetched == 0 {
            break;
        }
        if unsafe { var[0].Anonymous.Anonymous.vt } != VT_DISPATCH as _ {
            continue;
        }

        let mut hwnd = Default::default();
        let location = get_browser_info(
            unsafe {
                var[0]
                    .Anonymous
                    .Anonymous
                    .Anonymous
                    .pdispVal
                    .as_ref()
                    .unwrap()
                    .as_raw()
            },
            &mut hwnd,
        )?;

        let location = String::from_utf16_lossy(&location);
        locations.push(location);
    }

    Ok(locations)
}

fn get_browser_info(unk: *mut c_void, hwnd: &mut HWND) -> Result<Vec<u16>> {
    let shell_browser: IShellBrowser =
        unsafe { IUnknown_QueryService(&IUnknown::from_raw(unk), &SID_STopLevelBrowser) }?;
    *hwnd = unsafe { shell_browser.GetWindow() }?;

    get_location_from_view(&shell_browser)
}

fn get_location_from_view(browser: &IShellBrowser) -> Result<Vec<u16>> {
    let shell_view = unsafe { browser.QueryActiveShellView() }?;
    let persist_id_list: IPersistIDList = shell_view.cast()?;
    let id_list = unsafe { persist_id_list.GetIDList() }?;

    let item = unsafe { SHCreateItemFromIDList::<IShellItem>(id_list) }?;
    let ptr = unsafe { item.GetDisplayName(SIGDN_DESKTOPABSOLUTEPARSING) }?;

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

    unsafe { CoTaskMemFree(Some(ptr.0 as _)) };
    unsafe { CoTaskMemFree(Some(id_list as _)) };

    Ok(path)
}
