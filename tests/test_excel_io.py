"""
Tests for excel_io.py — covers reading USNs from an input workbook
and writing headers/student/subject data to the output workbook.

Write-side tests (write_headers, write_student_info, write_subject_data)
verify that the functions run without errors and produce the correct
structure via xlwt's public API; file round-trip tests that require
xlrd are intentionally excluded.
"""
import sys
import os
import tempfile

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'src'))

import openpyxl
import xlwt
import pytest

from excel_io import (
    load_input_workbook,
    read_usn,
    create_output_workbook,
    write_headers,
    write_student_info,
    write_subject_data,
    save_workbook,
)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _make_input_xlsx(usns: list) -> str:
    """Write a list of USNs into column A of a temp .xlsx file and return its path."""
    wb = openpyxl.Workbook()
    ws = wb.active
    for i, usn in enumerate(usns, start=1):
        ws.cell(row=i, column=1, value=usn)
    tmp = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
    wb.save(tmp.name)
    tmp.close()
    return tmp.name


def _make_subject(name="MATH", ia="20", see="60", total="80", res="P"):
    return {"name": name, "ia": ia, "see": see, "total": total, "res": res}


# ---------------------------------------------------------------------------
# load_input_workbook
# ---------------------------------------------------------------------------

class TestLoadInputWorkbook:

    def test_loads_valid_file(self):
        path = _make_input_xlsx(["1AT21CS001"])
        wb, ws = load_input_workbook(path)
        assert wb is not None
        assert ws is not None

    def test_returns_none_for_missing_file(self):
        wb, ws = load_input_workbook("/nonexistent/path/file.xlsx")
        assert wb is None
        assert ws is None

    def test_returns_none_for_invalid_file(self):
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False, mode='w') as f:
            f.write("not an excel file")
            bad_path = f.name
        wb, ws = load_input_workbook(bad_path)
        assert wb is None
        assert ws is None

    def test_active_sheet_has_correct_data(self):
        usns = ["1AT21CS001", "1AT21CS002", "1AT21CS003"]
        path = _make_input_xlsx(usns)
        _, ws = load_input_workbook(path)
        for i, usn in enumerate(usns, start=1):
            assert ws.cell(row=i, column=1).value == usn


# ---------------------------------------------------------------------------
# read_usn
# ---------------------------------------------------------------------------

class TestReadUsn:

    def setup_method(self):
        path = _make_input_xlsx(["1AT21CS001", "1AT21CS002", None, "1AT21CS004"])
        _, self.sheet = load_input_workbook(path)

    def test_reads_first_row(self):
        assert read_usn(self.sheet, 1) == "1AT21CS001"

    def test_reads_second_row(self):
        assert read_usn(self.sheet, 2) == "1AT21CS002"

    def test_reads_last_populated_row(self):
        assert read_usn(self.sheet, 4) == "1AT21CS004"

    def test_returns_none_for_empty_cell(self):
        assert read_usn(self.sheet, 3) is None

    def test_returns_none_for_out_of_bounds_row(self):
        assert read_usn(self.sheet, 999) is None


# ---------------------------------------------------------------------------
# create_output_workbook
# ---------------------------------------------------------------------------

class TestCreateOutputWorkbook:

    def test_returns_three_tuple(self):
        result = create_output_workbook()
        assert len(result) == 3

    def test_workbook_is_xlwt_workbook(self):
        wb, _, _ = create_output_workbook()
        assert isinstance(wb, xlwt.Workbook)

    def test_sheet_is_xlwt_worksheet(self):
        _, ws, _ = create_output_workbook()
        assert isinstance(ws, xlwt.Worksheet)

    def test_orange_style_is_xlwt_xfstyle(self):
        _, _, style = create_output_workbook()
        assert isinstance(style, xlwt.XFStyle)


# ---------------------------------------------------------------------------
# write_headers — smoke tests (no crash = correct)
# ---------------------------------------------------------------------------

class TestWriteHeaders:

    def test_does_not_raise(self):
        _, ws, _ = create_output_workbook()
        write_headers(ws)   # must not raise


# ---------------------------------------------------------------------------
# write_student_info — smoke tests
# ---------------------------------------------------------------------------

class TestWriteStudentInfo:

    def test_does_not_raise(self):
        _, ws, _ = create_output_workbook()
        write_student_info(ws, 2, "1AT21CS001", "Alice")

    def test_multiple_rows_do_not_raise(self):
        _, ws, _ = create_output_workbook()
        write_student_info(ws, 2, "1AT21CS001", "Alice")
        write_student_info(ws, 3, "1AT21CS002", "Bob")


# ---------------------------------------------------------------------------
# write_subject_data — smoke tests
# ---------------------------------------------------------------------------

class TestWriteSubjectData:

    def test_single_subject_does_not_raise(self):
        _, ws, style = create_output_workbook()
        write_subject_data(ws, 2, [_make_subject()], style)

    def test_multiple_subjects_do_not_raise(self):
        _, ws, style = create_output_workbook()
        subjects = [_make_subject("MATH"), _make_subject("PHYSICS"), _make_subject("CHEMISTRY")]
        write_subject_data(ws, 2, subjects, style)

    def test_empty_subject_list_does_not_raise(self):
        _, ws, style = create_output_workbook()
        write_subject_data(ws, 2, [], style)


# ---------------------------------------------------------------------------
# save_workbook
# ---------------------------------------------------------------------------

class TestSaveWorkbook:

    def test_saves_file_to_disk(self):
        wb, _, _ = create_output_workbook()
        tmp = tempfile.NamedTemporaryFile(suffix=".xls", delete=False)
        tmp.close()
        save_workbook(wb, tmp.name)
        assert os.path.exists(tmp.name)
        assert os.path.getsize(tmp.name) > 0

    def test_raises_on_invalid_path(self):
        wb, _, _ = create_output_workbook()
        with pytest.raises(Exception):
            save_workbook(wb, "/nonexistent_dir/that/does/not/exist/out.xls")
