use std::sync::{Arc, Mutex};
use std::time::Duration;

use core::operations::explorer::wait_for_explorer_stable;
use core::operations::location::{get_explorer_windows, open_location};
use core::operations::process::{kill_process_by_name, start_process};
use infrastructure::windows_os::windows_api::Win32WindowApi;
use tokio::task;
use windows::Win32::System::Com::{
    COINIT_APARTMENTTHREADED, COINIT_DISABLE_OLE1DDE, CoInitializeEx,
};

pub mod core;
pub mod data;
pub mod infrastructure;

#[tokio::main]
async fn main() {
    unsafe {
        let _ = CoInitializeEx(None, COINIT_APARTMENTTHREADED | COINIT_DISABLE_OLE1DDE);
    };

    let window_api = Win32WindowApi;

    let windows = get_explorer_windows(&window_api);

    kill_process_by_name("explorer.exe");
    start_process("explorer.exe");

    let already_open_explorer_windows = Arc::new(Mutex::new(Vec::<isize>::new()));
    if let Ok(_) = wait_for_explorer_stable(Duration::from_secs(10)).await {
        let mut tasks = Vec::new();

        for window in windows {
            let windows_mutex = Arc::clone(&already_open_explorer_windows);

            // working when synchronous
            // open_location(&window, &windows_mutex, &Win32WindowApi);
            let task = tokio::spawn(async move {
                // not working when asynchronous
                open_location(&window, &windows_mutex, &Win32WindowApi);
            });

            tasks.push(task);
        }

        for task in tasks {
            let _ = task.await;
        }
    }
}
