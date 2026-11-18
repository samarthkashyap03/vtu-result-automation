"""
Tests for main.py — process_results() is tested with mocked GUI, Scraper,
and excel_io so no real browser, display, or file I/O is required.
"""
import sys
import os

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'src'))

import pytest
from unittest.mock import MagicMock, patch, call


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _make_gui(captcha="WXYZ"):
    gui = MagicMock()
    gui.get_inputs.return_value = {
        "driver_path": "/fake/chromedriver",
        "usn_file": "/fake/usns.xlsx",
        "website": "http://fake.vtu.ac.in",
        "save_path": "/fake/results.xls",
        "start_row": "1",
        "end_row": "2",
    }
    gui.get_captcha_input.return_value = captcha
    return gui


def _make_scraper(usn="1AT21CS001", name="Alice", subjects=None):
    if subjects is None:
        subjects = [{"name": "MATH", "ia": "20", "see": "60", "total": "80", "res": "P"}]
    scraper = MagicMock()
    scraper.setup_driver.return_value = None
    scraper.locate_page_elements.return_value = True
    scraper.submit_and_switch_to_result.return_value = True
    scraper.scrape_student_info.return_value = (usn, name)
    scraper.scrape_subjects.return_value = subjects
    return scraper


# ---------------------------------------------------------------------------
# process_results — row validation
# ---------------------------------------------------------------------------

class TestProcessResultsRowValidation:

    @patch("main.time.sleep")
    @patch("main.load_input_workbook", return_value=(None, None))
    def test_invalid_start_row_shows_error(self, mock_load, mock_sleep):
        from main import process_results
        gui = _make_gui()
        gui.get_inputs.return_value = {
            **gui.get_inputs(),
            "start_row": "abc",
            "end_row": "5",
        }
        process_results(gui, gui.get_inputs())
        gui.show_error.assert_called_once()
        error_msg = gui.show_error.call_args[0][0]
        assert "row" in error_msg.lower() or "number" in error_msg.lower()

    @patch("main.time.sleep")
    @patch("main.load_input_workbook", return_value=(None, None))
    def test_invalid_end_row_shows_error(self, mock_load, mock_sleep):
        from main import process_results
        gui = _make_gui()
        gui.get_inputs.return_value = {
            **gui.get_inputs(),
            "start_row": "1",
            "end_row": "xyz",
        }
        process_results(gui, gui.get_inputs())
        gui.show_error.assert_called_once()


# ---------------------------------------------------------------------------
# process_results — workbook load failure
# ---------------------------------------------------------------------------

class TestProcessResultsWorkbookFailure:

    @patch("main.time.sleep")
    @patch("main.load_input_workbook", return_value=(None, None))
    def test_shows_error_when_workbook_fails_to_load(self, mock_load, mock_sleep):
        from main import process_results
        gui = _make_gui()
        process_results(gui, gui.get_inputs())
        gui.show_error.assert_called_once()


# ---------------------------------------------------------------------------
# process_results — scraper driver failure
# ---------------------------------------------------------------------------

class TestProcessResultsDriverFailure:

    @patch("main.time.sleep")
    @patch("main.save_workbook")
    @patch("main.write_subject_data")
    @patch("main.write_student_info")
    @patch("main.write_headers")
    @patch("main.create_output_workbook")
    @patch("main.load_input_workbook")
    @patch("main.ResultScraper")
    def test_shows_error_when_driver_setup_fails(
        self, MockScraper, mock_load, mock_create, mock_headers,
        mock_write_info, mock_write_sub, mock_save, mock_sleep
    ):
        from main import process_results

        mock_load.return_value = (MagicMock(), MagicMock())
        mock_create.return_value = (MagicMock(), MagicMock(), MagicMock())

        scraper_instance = MagicMock()
        scraper_instance.setup_driver.side_effect = Exception("Chrome not found")
        MockScraper.return_value = scraper_instance

        gui = _make_gui()
        process_results(gui, gui.get_inputs())
        gui.show_error.assert_called_once()


# ---------------------------------------------------------------------------
# process_results — captcha cancelled
# ---------------------------------------------------------------------------

class TestProcessResultsCaptchaCancelled:

    @patch("main.time.sleep")
    @patch("main.save_workbook")
    @patch("main.write_headers")
    @patch("main.create_output_workbook")
    @patch("main.load_input_workbook")
    @patch("main.ResultScraper")
    def test_shows_warning_when_captcha_cancelled(
        self, MockScraper, mock_load, mock_create,
        mock_headers, mock_save, mock_sleep
    ):
        from main import process_results

        mock_load.return_value = (MagicMock(), MagicMock())
        mock_create.return_value = (MagicMock(), MagicMock(), MagicMock())
        MockScraper.return_value = _make_scraper()

        gui = _make_gui(captcha=None)   # simulate user cancelling the captcha dialog
        process_results(gui, gui.get_inputs())
        gui.show_warning.assert_called_once()


# ---------------------------------------------------------------------------
# process_results — happy path (two successful USNs)
# ---------------------------------------------------------------------------

class TestProcessResultsHappyPath:

    @patch("main.time.sleep")
    @patch("main.read_usn")
    @patch("main.save_workbook")
    @patch("main.write_subject_data")
    @patch("main.write_student_info")
    @patch("main.write_headers")
    @patch("main.create_output_workbook")
    @patch("main.load_input_workbook")
    @patch("main.ResultScraper")
    def test_saves_workbook_on_success(
        self, MockScraper, mock_load, mock_create, mock_headers,
        mock_write_info, mock_write_sub, mock_save, mock_read_usn, mock_sleep
    ):
        from main import process_results

        mock_load.return_value = (MagicMock(), MagicMock())
        mock_create.return_value = (MagicMock(), MagicMock(), MagicMock())
        mock_read_usn.side_effect = ["1AT21CS001", "1AT21CS002"]
        MockScraper.return_value = _make_scraper()

        gui = _make_gui()
        process_results(gui, gui.get_inputs())
        mock_save.assert_called_once()

    @patch("main.time.sleep")
    @patch("main.read_usn")
    @patch("main.save_workbook")
    @patch("main.write_subject_data")
    @patch("main.write_student_info")
    @patch("main.write_headers")
    @patch("main.create_output_workbook")
    @patch("main.load_input_workbook")
    @patch("main.ResultScraper")
    def test_shows_success_info_dialog(
        self, MockScraper, mock_load, mock_create, mock_headers,
        mock_write_info, mock_write_sub, mock_save, mock_read_usn, mock_sleep
    ):
        from main import process_results

        mock_load.return_value = (MagicMock(), MagicMock())
        mock_create.return_value = (MagicMock(), MagicMock(), MagicMock())
        mock_read_usn.side_effect = ["1AT21CS001", "1AT21CS002"]
        MockScraper.return_value = _make_scraper()

        gui = _make_gui()
        process_results(gui, gui.get_inputs())
        gui.show_info.assert_called_once()

    @patch("main.time.sleep")
    @patch("main.read_usn")
    @patch("main.save_workbook")
    @patch("main.write_subject_data")
    @patch("main.write_student_info")
    @patch("main.write_headers")
    @patch("main.create_output_workbook")
    @patch("main.load_input_workbook")
    @patch("main.ResultScraper")
    def test_write_student_info_called_per_usn(
        self, MockScraper, mock_load, mock_create, mock_headers,
        mock_write_info, mock_write_sub, mock_save, mock_read_usn, mock_sleep
    ):
        from main import process_results

        mock_load.return_value = (MagicMock(), MagicMock())
        mock_create.return_value = (MagicMock(), MagicMock(), MagicMock())
        mock_read_usn.side_effect = ["1AT21CS001", "1AT21CS002"]
        MockScraper.return_value = _make_scraper()

        gui = _make_gui()
        process_results(gui, gui.get_inputs())
        assert mock_write_info.call_count == 2

    @patch("main.time.sleep")
    @patch("main.read_usn")
    @patch("main.save_workbook")
    @patch("main.write_subject_data")
    @patch("main.write_student_info")
    @patch("main.write_headers")
    @patch("main.create_output_workbook")
    @patch("main.load_input_workbook")
    @patch("main.ResultScraper")
    def test_cleanup_always_called(
        self, MockScraper, mock_load, mock_create, mock_headers,
        mock_write_info, mock_write_sub, mock_save, mock_read_usn, mock_sleep
    ):
        from main import process_results

        mock_load.return_value = (MagicMock(), MagicMock())
        mock_create.return_value = (MagicMock(), MagicMock(), MagicMock())
        mock_read_usn.side_effect = ["1AT21CS001", "1AT21CS002"]
        scraper = _make_scraper()
        MockScraper.return_value = scraper

        gui = _make_gui()
        process_results(gui, gui.get_inputs())
        scraper.cleanup.assert_called_once()
