"""
Test yearbook-choice: from raw input Excel run validation, produce run-summary
and errors, then compare to expected files in this folder. No GUI step.

Put input.xlsx, expected_run_summary.xlsx, and expected_errors.xlsx in this folder.
"""
import os
import sys
import unittest
import pandas as pd

TESTS_DIR = os.path.dirname(os.path.abspath(__file__))
CODE_DIR = os.path.join(os.path.dirname(TESTS_DIR), "code-yearbook-choice")
if CODE_DIR not in sys.path:
    sys.path.insert(0, CODE_DIR)


def _normalize_df_for_compare(df):
    if df is None or df.empty:
        return df
    df = df.copy()
    if "timestamp" in df.columns:
        df = df.drop(columns=["timestamp"])
    return df.sort_index(axis=1).reset_index(drop=True)


def _safe_str(val):
    """Convert to str; treat NaN/None as empty."""
    if pd.isna(val) or val is None:
        return ""
    return str(val).strip()


def _safe_selection(val):
    """Selection value: a/b/c/d or default 'd'."""
    if pd.isna(val) or val is None or str(val).strip() == "":
        return "d"
    s = str(val).lower().strip()
    return s if s in ("a", "b", "c", "d") else "d"


class TestYearbookChoice(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.input_path = os.path.join(TESTS_DIR, "input.xlsx")
        cls.expected_summary = os.path.join(TESTS_DIR, "expected_run_summary.xlsx")
        cls.expected_errors = os.path.join(TESTS_DIR, "expected_errors.xlsx")

    def test_run_summary_and_errors(self):
        if not os.path.exists(self.input_path):
            self.skipTest("Put input.xlsx in yearbook-choice/tests/")
        if not os.path.exists(self.expected_summary):
            self.skipTest("Put expected_run_summary.xlsx in yearbook-choice/tests/")
        if not os.path.exists(self.expected_errors):
            self.skipTest("Put expected_errors.xlsx in yearbook-choice/tests/")

        from validate_data import validate_data_from_path, write_error_report

        cleaned, errors, cols = validate_data_from_path(self.input_path)
        id_col = cols["student_id"]
        ln_col = cols["last_name"]
        sel_col = cols["selection"]

        summary_rows = [
            {
                "Student ID": _safe_str(row.get(id_col)),
                "Last Name": _safe_str(row.get(ln_col)),
                "Yearbook Selection": _safe_selection(row.get(sel_col)),
            }
            for row in cleaned
        ]
        got_summary = _normalize_df_for_compare(pd.DataFrame(summary_rows))

        # Generate got_errors.xlsx using the same code as the app
        got_errors_path = os.path.join(TESTS_DIR, "got_errors.xlsx")
        write_error_report(errors, got_errors_path)
        got_errors = _normalize_df_for_compare(pd.read_excel(got_errors_path))
        exp_summary = _normalize_df_for_compare(pd.read_excel(self.expected_summary))
        exp_errors = _normalize_df_for_compare(pd.read_excel(self.expected_errors))

        # Keep unsorted copies for writing (same row order as generated; column order will match expected)
        got_summary_for_file = got_summary.copy()
        got_errors_for_file = got_errors.copy()

        for col in ["Student ID"]:
            if col in exp_summary.columns and col in got_summary.columns:
                exp_summary = exp_summary.copy()
                exp_summary[col] = exp_summary[col].astype(str)
                got_summary = got_summary.copy()
                got_summary[col] = got_summary[col].astype(str)
        def _student_id_col(df):
            for c in df.columns:
                if c and "student" in str(c).lower() and "id" in str(c).lower():
                    return c
            return None
        sid_col_exp = _student_id_col(exp_errors)
        sid_col_got = _student_id_col(got_errors)
        if sid_col_exp:
            exp_errors = exp_errors.copy()
            exp_errors[sid_col_exp] = exp_errors[sid_col_exp].astype(str)
        if sid_col_got:
            got_errors = got_errors.copy()
            got_errors[sid_col_got] = got_errors[sid_col_got].astype(str)

        # Sort by key columns so row order does not affect comparison
        sort_cols_summary = [c for c in ["Student ID", "Last Name"] if c in got_summary.columns]
        if sort_cols_summary and not got_summary.empty:
            got_summary = got_summary.sort_values(sort_cols_summary).reset_index(drop=True)
        if sort_cols_summary and not exp_summary.empty:
            exp_summary = exp_summary.sort_values(sort_cols_summary).reset_index(drop=True)
        sort_cols_err_got = [c for c in (([sid_col_got] if sid_col_got else []) + ["error_reason"]) if c in got_errors.columns]
        sort_cols_err_exp = [c for c in (([sid_col_exp] if sid_col_exp else []) + ["error_reason"]) if c in exp_errors.columns]
        if sort_cols_err_got and not got_errors.empty:
            got_errors = got_errors.sort_values(sort_cols_err_got).reset_index(drop=True)
        if sort_cols_err_exp and not exp_errors.empty:
            exp_errors = exp_errors.sort_values(sort_cols_err_exp).reset_index(drop=True)

        # Write outputs with same column order as expected (for easy comparison)
        summary_col_order = [c for c in exp_summary.columns if c in got_summary_for_file.columns]
        if summary_col_order:
            got_summary_for_file[summary_col_order].to_excel(os.path.join(TESTS_DIR, "got_run_summary.xlsx"), index=False)
        else:
            got_summary_for_file.to_excel(os.path.join(TESTS_DIR, "got_run_summary.xlsx"), index=False)
        errors_col_order = [c for c in exp_errors.columns if c in got_errors_for_file.columns]
        if errors_col_order:
            got_errors_for_file[errors_col_order].to_excel(got_errors_path, index=False)

        pd.testing.assert_frame_equal(got_summary, exp_summary, check_like=True)
        pd.testing.assert_frame_equal(got_errors, exp_errors, check_like=True)


if __name__ == "__main__":
    unittest.main()
