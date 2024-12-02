use std::time::Duration;

use explorer::wait_for_explorer_stable;
use location::{get_explorer_windows, open_location};
use process::{kill_process_by_name, start_process};
use windows::Win32::System::Com::{
    CoInitializeEx, COINIT_APARTMENTTHREADED, COINIT_DISABLE_OLE1DDE,
};

pub mod explorer;
pub mod location;
pub mod models;
pub mod process;

fn main() {
    unsafe {
        let _ = CoInitializeEx(None, COINIT_APARTMENTTHREADED | COINIT_DISABLE_OLE1DDE);
    };

    let windows = get_explorer_windows();
    kill_process_by_name("explorer.exe");
    start_process("explorer.exe");

    let mut already_open_explorer_windows: Vec<isize> = vec![];
    if let Ok(_) = wait_for_explorer_stable(Duration::from_secs(10)) {
        for window in windows {
            if let Some(id) = open_location(&window, already_open_explorer_windows.clone()) {
                already_open_explorer_windows.push(id);
            }
        }
    }
}
