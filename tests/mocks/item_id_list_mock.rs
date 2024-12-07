use windows::Win32::UI::Shell::Common::{self, SHITEMID};

pub fn mock_item_id_list() -> *mut Common::ITEMIDLIST {
    use std::alloc::{alloc, Layout};
    use std::ptr;

    let shitemid = SHITEMID {
        cb: 4,
        abID: [0x42],
    };

    let item_id_list = Common::ITEMIDLIST { mkid: shitemid };
    let layout = Layout::new::<Common::ITEMIDLIST>();
    unsafe {
        let ptr = alloc(layout) as *mut Common::ITEMIDLIST;
        if ptr.is_null() {
            panic!("Failed to allocate memory for mock ITEMIDLIST");
        }

        ptr::write(ptr, item_id_list);
        ptr
    }
}
