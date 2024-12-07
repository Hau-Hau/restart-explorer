use windows::Win32::{
    Foundation::HWND,
    UI::WindowsAndMessaging::{GetParent, GetTopWindow, GetWindow, GET_WINDOW_CMD},
};

pub trait WindowApi {
    fn get_top_window(&self, hwnd: HWND) -> windows::core::Result<HWND>;
    fn get_window(&self, hwnd: HWND, command: GET_WINDOW_CMD) -> windows::core::Result<HWND>;
    fn get_parent(&self, hwnd: HWND) -> windows::core::Result<HWND>;
}

pub struct Win32WindowApi;

impl WindowApi for Win32WindowApi {
    fn get_top_window(&self, hwnd: HWND) -> windows::core::Result<HWND> {
        unsafe { Ok(GetTopWindow(hwnd)?) }
    }

    fn get_window(&self, hwnd: HWND, command: GET_WINDOW_CMD) -> windows::core::Result<HWND> {
        unsafe { Ok(GetWindow(hwnd, command)?) }
    }

    fn get_parent(&self, hwnd: HWND) -> windows::core::Result<HWND> {
        unsafe { Ok(GetParent(hwnd)?) }
    }
}
