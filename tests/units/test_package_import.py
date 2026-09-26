def test_package_imports_without_loading_pywin32() -> None:
    import email_downloader

    assert email_downloader.__version__ == "0.1.0"
