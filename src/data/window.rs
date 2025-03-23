use windows::Win32::Foundation::RECT;

#[derive(Clone)]
pub struct Window {
    pub location: String,
    pub rect: RECT,
    pub is_minimized: bool,
    pub zindex: i32,
}
