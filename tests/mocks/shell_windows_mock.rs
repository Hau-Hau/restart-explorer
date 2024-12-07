use restart_explorer::infrastructure::windows_os::{
    enum_variant::EnumVariant, shell_windows::ShellWindows,
};

use super::enum_variant::MockEnumVariant;

#[derive(Clone, Default)]
pub struct MockShellWindows {
    pub enum_variant: MockEnumVariant,
}

impl ShellWindows for MockShellWindows {
    fn new_enum_variant(self) -> windows_core::Result<impl EnumVariant> {
        Ok(self.enum_variant)
    }
}
