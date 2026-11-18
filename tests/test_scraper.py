"""
Tests for scraper.py — all Selenium / Chrome interactions are mocked,
so the tests run without a real browser installed.
"""
import sys
import os

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'src'))

import pytest
from unittest.mock import MagicMock, patch, PropertyMock
from selenium.common.exceptions import NoSuchElementException, TimeoutException

from scraper import ResultScraper
from config import MAX_SUBJECTS


# ---------------------------------------------------------------------------
# Fixtures
# ---------------------------------------------------------------------------

@pytest.fixture
def scraper():
    """Return a ResultScraper with no live driver."""
    return ResultScraper(driver_path="/fake/chromedriver", website_url="http://fake.vtu.ac.in")


@pytest.fixture
def scraper_with_mock_driver(scraper):
    """Attach a fully mocked Selenium driver to the scraper instance."""
    mock_driver = MagicMock()
    mock_driver.window_handles = ["main_window", "result_window"]
    mock_driver.current_window_handle = "main_window"

    scraper.driver = mock_driver
    scraper.actions = MagicMock()
    scraper.wait = MagicMock()
    scraper.main_window = "main_window"
    scraper.usn_box = MagicMock()
    scraper.captcha_box = MagicMock()
    scraper.submit_btn = MagicMock()
    return scraper


# ---------------------------------------------------------------------------
# __init__
# ---------------------------------------------------------------------------

class TestResultScraperInit:

    def test_stores_driver_path(self, scraper):
        assert scraper.driver_path == "/fake/chromedriver"

    def test_stores_website_url(self, scraper):
        assert scraper.website_url == "http://fake.vtu.ac.in"

    def test_initial_driver_is_none(self, scraper):
        assert scraper.driver is None

    def test_initial_page_elements_are_none(self, scraper):
        assert scraper.usn_box is None
        assert scraper.captcha_box is None
        assert scraper.submit_btn is None


# ---------------------------------------------------------------------------
# locate_page_elements
# ---------------------------------------------------------------------------

class TestLocatePageElements:

    def test_returns_true_when_all_elements_found(self, scraper_with_mock_driver):
        s = scraper_with_mock_driver
        s.wait.until.return_value = MagicMock()
        assert s.locate_page_elements() is True

    def test_returns_false_on_timeout(self, scraper_with_mock_driver):
        s = scraper_with_mock_driver
        s.wait.until.side_effect = TimeoutException("timed out")
        assert s.locate_page_elements() is False

    def test_returns_false_when_element_not_found(self, scraper_with_mock_driver):
        s = scraper_with_mock_driver
        s.wait.until.side_effect = NoSuchElementException("not found")
        assert s.locate_page_elements() is False


# ---------------------------------------------------------------------------
# enter_usn_and_captcha
# ---------------------------------------------------------------------------

class TestEnterUsnAndCaptcha:

    def test_clears_and_fills_usn_box(self, scraper_with_mock_driver):
        s = scraper_with_mock_driver
        with patch("scraper.time.sleep"):
            s.enter_usn_and_captcha("1AT21CS001", "ABCD")
        s.usn_box.clear.assert_called_once()
        s.usn_box.send_keys.assert_called_once_with("1AT21CS001")

    def test_clears_and_fills_captcha_box(self, scraper_with_mock_driver):
        s = scraper_with_mock_driver
        with patch("scraper.time.sleep"):
            s.enter_usn_and_captcha("1AT21CS001", "ABCD")
        s.captcha_box.clear.assert_called_once()
        s.captcha_box.send_keys.assert_called_once_with("ABCD")


# ---------------------------------------------------------------------------
# submit_and_switch_to_result
# ---------------------------------------------------------------------------

class TestSubmitAndSwitchToResult:

    def test_returns_true_when_result_window_opens(self, scraper_with_mock_driver):
        s = scraper_with_mock_driver
        s.wait.until.return_value = True  # number_of_windows_to_be condition
        assert s.submit_and_switch_to_result() is True

    def test_switches_to_result_window(self, scraper_with_mock_driver):
        s = scraper_with_mock_driver
        s.wait.until.return_value = True
        s.submit_and_switch_to_result()
        s.driver.switch_to.window.assert_called_with("result_window")

    def test_returns_false_on_timeout(self, scraper_with_mock_driver):
        s = scraper_with_mock_driver
        s.wait.until.side_effect = TimeoutException("window timeout")
        assert s.submit_and_switch_to_result() is False

    def test_clears_usn_box_on_timeout(self, scraper_with_mock_driver):
        s = scraper_with_mock_driver
        s.wait.until.side_effect = TimeoutException("window timeout")
        s.submit_and_switch_to_result()
        s.usn_box.clear.assert_called()


# ---------------------------------------------------------------------------
# scrape_student_info
# ---------------------------------------------------------------------------

class TestScrapeStudentInfo:

    def test_returns_usn_and_name(self, scraper_with_mock_driver):
        s = scraper_with_mock_driver
        usn_el = MagicMock()
        usn_el.text = "1AT21CS001"
        name_el = MagicMock()
        name_el.text = "Alice Kumar"

        s.wait.until.return_value = usn_el
        s.driver.find_element.return_value = name_el

        usn, name = s.scrape_student_info()
        assert usn == "1AT21CS001"
        assert name == "Alice Kumar"

    def test_returns_none_none_on_timeout(self, scraper_with_mock_driver):
        s = scraper_with_mock_driver
        s.wait.until.side_effect = TimeoutException("timed out")
        usn, name = s.scrape_student_info()
        assert usn is None
        assert name is None

    def test_returns_none_none_when_element_missing(self, scraper_with_mock_driver):
        s = scraper_with_mock_driver
        s.wait.until.side_effect = NoSuchElementException("not found")
        usn, name = s.scrape_student_info()
        assert usn is None
        assert name is None


# ---------------------------------------------------------------------------
# scrape_subjects
# ---------------------------------------------------------------------------

class TestScrapeSubjects:

    def _make_text_element(self, text):
        el = MagicMock()
        el.text = text
        return el

    def test_returns_empty_list_when_no_subjects(self, scraper_with_mock_driver):
        s = scraper_with_mock_driver
        s.driver.find_element.side_effect = NoSuchElementException("not found")
        subjects = s.scrape_subjects()
        assert subjects == []

    def test_returns_one_subject_then_stops(self, scraper_with_mock_driver):
        s = scraper_with_mock_driver

        # First five find_element calls return valid data; the next raises
        elements = [
            self._make_text_element("MATH"),   # name
            self._make_text_element("20"),     # ia
            self._make_text_element("60"),     # see
            self._make_text_element("80"),     # total
            self._make_text_element("P"),      # res
        ]
        call_count = {"n": 0}

        def side_effect(*args, **kwargs):
            if call_count["n"] < len(elements):
                el = elements[call_count["n"]]
                call_count["n"] += 1
                return el
            raise NoSuchElementException("no more subjects")

        s.driver.find_element.side_effect = side_effect
        subjects = s.scrape_subjects()
        assert len(subjects) == 1
        assert subjects[0]["name"] == "MATH"
        assert subjects[0]["ia"] == "20"
        assert subjects[0]["res"] == "P"

    def test_stops_at_max_subjects(self, scraper_with_mock_driver):
        s = scraper_with_mock_driver
        # Always return a valid element — should stop at MAX_SUBJECTS
        s.driver.find_element.return_value = self._make_text_element("X")
        subjects = s.scrape_subjects()
        assert len(subjects) == MAX_SUBJECTS


# ---------------------------------------------------------------------------
# cleanup
# ---------------------------------------------------------------------------

class TestCleanup:

    def test_calls_driver_quit(self, scraper_with_mock_driver):
        s = scraper_with_mock_driver
        s.cleanup()
        s.driver.quit.assert_called_once()

    def test_does_not_raise_when_driver_is_none(self, scraper):
        """cleanup() must be safe to call before setup_driver()."""
        scraper.driver = None
        scraper.cleanup()   # Should not raise

    def test_does_not_raise_when_quit_throws(self, scraper_with_mock_driver):
        s = scraper_with_mock_driver
        s.driver.quit.side_effect = Exception("WebDriver crashed")
        s.cleanup()         # Should swallow the exception and not propagate
