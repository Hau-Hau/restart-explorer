use std::borrow::Borrow;

use restart_explorer::infrastructure::windows_os::{
    shell_windows::ShellWindows, windows_api::WindowApi,
};
use windows::Win32::{Foundation::HWND, UI::WindowsAndMessaging::GET_WINDOW_CMD};

use super::shell_windows_mock::MockShellWindows;

pub struct MockWindowApi {
    pub window: HWND,
    pub top_window: HWND,
    pub parent: HWND,
    pub shell_windows: MockShellWindows,
}

impl WindowApi for MockWindowApi {
    fn get_top_window(&self, _hwnd: HWND) -> windows::core::Result<HWND> {
        Ok(self.window)
    }

    fn get_window(&self, _hwnd: HWND, _command: GET_WINDOW_CMD) -> windows::core::Result<HWND> {
        Ok(self.top_window)
    }

    fn get_parent(&self, _hwnd: HWND) -> windows::core::Result<HWND> {
        Ok(self.parent)
    }

    fn create_shell_windows(&self) -> windows::core::Result<impl ShellWindows> {
        Ok(self.shell_windows.clone())
    }
}
