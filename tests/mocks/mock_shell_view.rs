use restart_explorer::infrastructure::windows_os::{
    persist_id_list::PersistIDList, shell_view::ShellView,
};
use windows::Win32::Foundation::HWND;

use super::mock_persist_id_list::MockPersistIDList;

pub struct MockShellView {
    pub persist_id_list: MockPersistIDList,
}

impl ShellView for MockShellView {
    fn get_window(&self) -> windows::core::Result<HWND> {
        Ok(HWND::default())
    }

    fn as_persist_id_list(&self) -> windows::core::Result<impl PersistIDList> {
        Ok(&self.persist_id_list)
    }
}
