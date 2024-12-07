use std::time::Duration;

use core::operations::explorer::wait_for_explorer_stable;
use core::operations::location::{get_explorer_windows, open_location};
use core::operations::process::{kill_process_by_name, start_process};
use infrastructure::windows_os::windows_api::Win32WindowApi;
use windows::Win32::System::Com::{
    CoInitializeEx, COINIT_APARTMENTTHREADED, COINIT_DISABLE_OLE1DDE,
};

pub mod core;
pub mod data;
pub mod infrastructure;

fn main() {
    unsafe {
        let _ = CoInitializeEx(None, COINIT_APARTMENTTHREADED | COINIT_DISABLE_OLE1DDE);
    };

    let window_api = Win32WindowApi;

    let windows = get_explorer_windows(&window_api);

    kill_process_by_name("explorer.exe");
    start_process("explorer.exe");

    let mut already_open_explorer_windows: Vec<isize> = vec![];
    if let Ok(_) = wait_for_explorer_stable(Duration::from_secs(10)) {
        for window in windows {
            if let Some(id) = open_location(&window, &already_open_explorer_windows, &window_api) {
                already_open_explorer_windows.push(id);
            }
        }
    }
}
