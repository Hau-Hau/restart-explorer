use location::{get_explorer_locations, open_location};
use process::{kill_process_by_name, start_process};
use windows::Win32::System::Com::{
    CoInitializeEx, COINIT_APARTMENTTHREADED, COINIT_DISABLE_OLE1DDE,
};

pub mod location;
pub mod process;

fn main() {
    unsafe {
        let _ = CoInitializeEx(None, COINIT_APARTMENTTHREADED | COINIT_DISABLE_OLE1DDE);
    };

    let Ok(locations) = get_explorer_locations() else {
        return;
    };

    kill_process_by_name("explorer.exe");
    start_process("explorer.exe");
    for location in locations {
        open_location(&location);
    }
}
