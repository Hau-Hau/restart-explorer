#[cfg(test)]
mod tests {
    use std::ffi::c_void;

    use crate::mocks::item_id_list_mock::mock_item_id_list;
    use crate::mocks::mock_persist_id_list::MockPersistIDList;
    use crate::mocks::mock_shell_item::MockShellItem;
    use crate::mocks::mock_shell_view::MockShellView;
    use crate::mocks::window_api_mock::MockWindowApi;

    use restart_explorer::core::operations::shell_view::get_path_from_shell_view;
    use restart_explorer::core::operations::window::get_window_z_index;
    use windows::Win32::Foundation::HWND;

    #[test]
    fn test_get_path_from_shell_view_success() {
        let mock_shell_view = MockShellView {
            persist_id_list: MockPersistIDList {
                id_list: mock_item_id_list(),
                shell_item: MockShellItem {
                    path: String::from("C:\\MockPath"),
                },
            },
        };

        let result = get_path_from_shell_view(&mock_shell_view);

        assert!(result.is_ok());
        assert_eq!(result.unwrap(), "C:\\MockPath");
    }

    #[test]
    fn test_window_at_top() {
        let api = MockWindowApi {
            parent: HWND(1 as *mut c_void),
            top_window: HWND::default(),
            window: HWND(2 as *mut c_void),
        };
        let z_index = get_window_z_index(HWND::default(), &api).unwrap();
        assert_eq!(z_index, 1, "Window at the top should have Z-index 1");
    }
}
