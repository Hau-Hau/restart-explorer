#[cfg(test)]
mod tests {
    use crate::mocks::item_id_list_mock::mock_item_id_list;
    use crate::mocks::mock_persist_id_list::MockPersistIDList;
    use crate::mocks::mock_shell_item::MockShellItem;
    use crate::mocks::mock_shell_view::MockShellView;

    use restart_explorer::core::operations::shell_view::get_path_from_shell_view;

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
}
