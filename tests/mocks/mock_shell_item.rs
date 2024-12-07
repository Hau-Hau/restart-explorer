use restart_explorer::infrastructure::windows_os::shell_item::ShellItem;

#[derive(Clone, Default)]
pub struct MockShellItem {
    pub path: String,
}

impl ShellItem for MockShellItem {
    fn get_display_name(&self) -> windows_core::Result<String> {
        Ok(self.path.clone())
    }
}
