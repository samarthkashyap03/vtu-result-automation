"""
Tests for logger.py — verifies setup_logger returns a correctly configured logger.
"""
import sys
import os
import logging
import tempfile

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'src'))

from logger import setup_logger


class TestSetupLogger:

    def test_returns_logger_instance(self):
        logger = setup_logger("test_instance")
        assert isinstance(logger, logging.Logger)

    def test_logger_has_correct_name(self):
        logger = setup_logger("test_name_check")
        assert logger.name == "test_name_check"

    def test_logger_default_level_is_info(self):
        logger = setup_logger("test_level")
        assert logger.level == logging.INFO

    def test_logger_custom_level(self):
        logger = setup_logger("test_custom_level", level=logging.DEBUG)
        assert logger.level == logging.DEBUG

    def test_logger_has_handlers(self):
        logger = setup_logger("test_handlers")
        assert len(logger.handlers) > 0

    def test_logger_has_file_and_console_handler(self):
        with tempfile.NamedTemporaryFile(suffix=".log", delete=False) as f:
            log_path = f.name

        logger = setup_logger("test_handler_types", log_file=log_path)
        handler_types = {type(h) for h in logger.handlers}

        assert logging.FileHandler in handler_types, "Expected a FileHandler"
        assert logging.StreamHandler in handler_types, "Expected a StreamHandler"

    def test_no_duplicate_handlers_on_repeated_calls(self):
        """Calling setup_logger twice with the same name must not add extra handlers."""
        logger_name = "test_no_duplicates"
        logger1 = setup_logger(logger_name)
        handler_count = len(logger1.handlers)

        logger2 = setup_logger(logger_name)
        assert len(logger2.handlers) == handler_count, "Duplicate handlers were added"

    def test_logger_writes_to_log_file(self):
        with tempfile.NamedTemporaryFile(suffix=".log", delete=False, mode='w') as f:
            log_path = f.name

        logger = setup_logger("test_writes_to_file", log_file=log_path)
        logger.info("hello from test")

        with open(log_path, "r") as f:
            contents = f.read()

        assert "hello from test" in contents
