use std::time::{Duration, Instant};

use tokio::time::sleep;
use windows::{
    Win32::{
        Foundation::{ERROR_TIMEOUT, WIN32_ERROR},
        UI::WindowsAndMessaging::{FindWindowW, IsWindowVisible},
    },
    core::w,
};

pub async fn wait_for_explorer_stable(timeout: Duration) -> Result<(), windows::core::Error> {
    let start = Instant::now();

    loop {
        if start.elapsed() > timeout {
            return Err(windows::core::Error::from(WIN32_ERROR(ERROR_TIMEOUT.0)));
        }

        let (progman_visible, explorer_visible) = tokio::task::spawn_blocking(|| unsafe {
            let progman_window = FindWindowW(w!("Progman"), None);
            let progman_visible = if progman_window.is_err() {
                false
            } else {
                IsWindowVisible(progman_window.unwrap()).as_bool()
            };

            let explorer_window = FindWindowW(w!("Shell_TrayWnd"), None);
            let explorer_visible = if explorer_window.is_err() {
                false
            } else {
                IsWindowVisible(explorer_window.unwrap()).as_bool()
            };

            (progman_visible, explorer_visible)
        })
        .await
        .unwrap_or((false, false));

        if !progman_visible || !explorer_visible {
            sleep(Duration::from_millis(100)).await;
            continue;
        }

        break;
    }

    Ok(())
}
