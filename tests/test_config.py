"""
Tests for config.py — verifies all constants are present, correctly typed,
and have sensible values.
"""
import sys
import os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'src'))

import config


class TestXPathSelectors:
    """XPath selector constants must be non-empty strings."""

    def test_xpath_usn_input_is_string(self):
        assert isinstance(config.XPATH_USN_INPUT, str)
        assert len(config.XPATH_USN_INPUT) > 0

    def test_xpath_captcha_input_is_string(self):
        assert isinstance(config.XPATH_CAPTCHA_INPUT, str)
        assert len(config.XPATH_CAPTCHA_INPUT) > 0

    def test_xpath_submit_button_is_string(self):
        assert isinstance(config.XPATH_SUBMIT_BUTTON, str)
        assert len(config.XPATH_SUBMIT_BUTTON) > 0

    def test_xpath_student_usn_is_string(self):
        assert isinstance(config.XPATH_STUDENT_USN, str)
        assert len(config.XPATH_STUDENT_USN) > 0

    def test_xpath_student_name_is_string(self):
        assert isinstance(config.XPATH_STUDENT_NAME, str)
        assert len(config.XPATH_STUDENT_NAME) > 0

    def test_xpath_subject_base_is_formattable(self):
        """XPATH_SUBJECT_BASE must have one format placeholder {}"""
        assert isinstance(config.XPATH_SUBJECT_BASE, str)
        result = config.XPATH_SUBJECT_BASE.format(1)
        assert "{}" not in result, "Placeholder was not consumed after .format()"


class TestExcelConfig:
    """Excel layout constants must be non-negative integers."""

    def test_header_row_is_non_negative_int(self):
        assert isinstance(config.EXCEL_HEADER_ROW, int)
        assert config.EXCEL_HEADER_ROW >= 0

    def test_subheader_row_is_non_negative_int(self):
        assert isinstance(config.EXCEL_SUBHEADER_ROW, int)
        assert config.EXCEL_SUBHEADER_ROW >= 0

    def test_subheader_row_after_header_row(self):
        assert config.EXCEL_SUBHEADER_ROW > config.EXCEL_HEADER_ROW

    def test_usn_column_is_non_negative_int(self):
        assert isinstance(config.EXCEL_USN_COLUMN, int)
        assert config.EXCEL_USN_COLUMN >= 0

    def test_name_column_is_non_negative_int(self):
        assert isinstance(config.EXCEL_NAME_COLUMN, int)
        assert config.EXCEL_NAME_COLUMN >= 0

    def test_subjects_start_column_after_name_column(self):
        assert config.EXCEL_SUBJECTS_START_COLUMN > config.EXCEL_NAME_COLUMN


class TestSubjectIndices:
    """Subject index constants must be positive integers and logically ordered."""

    def test_all_indices_are_positive_ints(self):
        for attr in [
            "SUBJECT_NAME_INDEX", "SUBJECT_IA_INDEX",
            "SUBJECT_SEE_INDEX", "SUBJECT_TOTAL_INDEX", "SUBJECT_RESULT_INDEX"
        ]:
            value = getattr(config, attr)
            assert isinstance(value, int) and value > 0, f"{attr} must be a positive int"

    def test_indices_are_unique(self):
        indices = [
            config.SUBJECT_NAME_INDEX,
            config.SUBJECT_IA_INDEX,
            config.SUBJECT_SEE_INDEX,
            config.SUBJECT_TOTAL_INDEX,
            config.SUBJECT_RESULT_INDEX,
        ]
        assert len(indices) == len(set(indices)), "Subject indices must all be unique"


class TestWaitConstants:
    """Wait time constants must be non-negative numbers."""

    def test_wait_after_startup_is_non_negative(self):
        assert isinstance(config.WAIT_AFTER_STARTUP, (int, float))
        assert config.WAIT_AFTER_STARTUP >= 0

    def test_wait_before_captcha_is_non_negative(self):
        assert isinstance(config.WAIT_BEFORE_CAPTCHA, (int, float))
        assert config.WAIT_BEFORE_CAPTCHA >= 0

    def test_wait_after_input_is_non_negative(self):
        assert isinstance(config.WAIT_AFTER_INPUT, (int, float))
        assert config.WAIT_AFTER_INPUT >= 0

    def test_wait_timeout_is_positive(self):
        assert isinstance(config.WAIT_TIMEOUT, (int, float))
        assert config.WAIT_TIMEOUT > 0, "WebDriver timeout must be > 0 seconds"

    def test_max_subjects_is_positive_int(self):
        assert isinstance(config.MAX_SUBJECTS, int)
        assert config.MAX_SUBJECTS > 0
