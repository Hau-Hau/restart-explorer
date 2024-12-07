use restart_explorer::infrastructure::windows_os::persist_id_list::PersistIDList;
use restart_explorer::infrastructure::windows_os::shell_item::ShellItem;
use windows::Win32::UI::Shell::Common;

use super::mock_shell_item::MockShellItem;

pub struct MockPersistIDList {
    pub id_list: *mut Common::ITEMIDLIST,
    pub shell_item: MockShellItem,
}

impl PersistIDList for &MockPersistIDList {
    fn get_id_list(&self) -> windows_core::Result<*mut Common::ITEMIDLIST> {
        Ok(self.id_list)
    }

    fn id_list_into_shell_item(
        &self,
        _id_list: *mut Common::ITEMIDLIST,
    ) -> windows_core::Result<impl ShellItem> {
        Ok(self.shell_item.clone())
    }
}
