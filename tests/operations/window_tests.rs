#[cfg(test)]
mod tests {
    use std::{ffi::c_void, time::Duration};

    use crate::mocks::{
        enum_variant::MockEnumVariant, shell_windows_mock::MockShellWindows,
        window_api_mock::MockWindowApi,
    };

    use restart_explorer::core::operations::window::{get_window_z_index, wait_for_window_stable};
    use windows::Win32::Foundation::HWND;

    #[test]
    fn test_wait_for_window_stable_timeout() {
        let window_api = MockWindowApi {
            parent: HWND(1 as *mut c_void),
            top_window: HWND::default(),
            window: HWND(2 as *mut c_void),
            shell_windows: MockShellWindows::default(),
        };
        let location = "C:\\TestPath";
        let timeout = Duration::from_secs(0);
        let already_open_explorer_windows: Vec<isize> = vec![];

        let result = wait_for_window_stable(
            location,
            timeout,
            &already_open_explorer_windows,
            &window_api,
        );

        assert!(result.is_err());
    }

    #[test]
    fn test_wait_for_window_stable_found() {
        let window_api = MockWindowApi {
            parent: HWND(1 as *mut c_void),
            top_window: HWND::default(),
            window: HWND(2 as *mut c_void),
            shell_windows: MockShellWindows {
                enum_variant: MockEnumVariant {
                    expected_hresult: 0,
                },
            },
        };
        let location = "C:\\TestPath";
        let timeout = Duration::from_secs(5);
        let already_open_explorer_windows: Vec<isize> = vec![];

        let result = wait_for_window_stable(
            location,
            timeout,
            &already_open_explorer_windows,
            &window_api,
        );

        assert!(result.is_ok());
        let window_id = result.unwrap();
        assert_ne!(window_id, 0);
    }

    #[test]
    fn test_window_at_top() {
        let window_api = MockWindowApi {
            parent: HWND(1 as *mut c_void),
            top_window: HWND::default(),
            window: HWND(2 as *mut c_void),
            shell_windows: MockShellWindows::default(),
        };
        let z_index = get_window_z_index(HWND::default(), &window_api).unwrap();
        assert_eq!(z_index, 1, "Window at the top should have Z-index 1");
    }
}
