use restart_explorer::infrastructure::windows_os::enum_variant::EnumVariant;
use windows_core::HRESULT;

#[derive(Clone, Default)]
pub struct MockEnumVariant {
    pub expected_hresult: i32,
}

impl EnumVariant for MockEnumVariant {
    fn next(
        &self,
        _rgvar: &mut [windows_core::VARIANT],
        _pceltfetched: *mut u32,
    ) -> windows_core::HRESULT {
        HRESULT(self.expected_hresult)
    }
}
