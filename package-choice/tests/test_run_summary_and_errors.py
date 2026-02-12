"""
Test package-choice: from input Excel, produce run-summary and errors Excel,
then compare to expected files in this folder. No GUI/coordinates step.

Put input.xlsx, expected_run_summary.xlsx, expected_errors.xlsx in this folder.
"""
import os
import sys
import unittest
import pandas as pd

# This folder (package-choice/tests/)
TESTS_DIR = os.path.dirname(os.path.abspath(__file__))
CODE_DIR = os.path.join(os.path.dirname(TESTS_DIR), "code-package-choice")
if CODE_DIR not in sys.path:
    sys.path.insert(0, CODE_DIR)


def _build_run_summary(students):
    """Same logic as main.py: build verif_data from students."""
    verif_data = []
    for s in students:
        sid = s["id"]
        lname = s.get("last_name", "")
        for grp in s.get("choices_groups", []):
            quick_pkg = grp["standard_string"]
            cd_value = touchup_value = class_pix = class_pix_no_pkg = ""
            for item in grp["others"]:
                if "lost order" in item["raw_product"].lower() or "invalid" in item["raw_product"].lower():
                    continue
                target_box = item.get("target_box", "")
                code = item["code"]
                if target_box == "cd_box":
                    cd_value = code
                elif target_box == "touchup":
                    touchup_value = "Pending"
                elif target_box == "class_pkg_box":
                    class_pix = f"{class_pix}, {code}".lstrip(", ") if class_pix else code
                elif target_box == "class_pix_no_pkg_box":
                    class_pix_no_pkg = f"{class_pix_no_pkg}, {code}".lstrip(", ") if class_pix_no_pkg else code
            verif_data.append({
                "Student ID": sid,
                "Last Name": lname,
                "Photo Choice": grp["photo_choice"] if grp["photo_choice"] else "(NONE)",
                "Quick Package Entry": quick_pkg,
                "CD": cd_value,
                "Touchup": touchup_value,
                "Class Pix": class_pix,
                "Class Pix No Pkg": class_pix_no_pkg,
            })
    return verif_data


def _collect_errors(students):
    """Same as main: log all data-load errors into a list of dicts."""
    errors = []
    for s in students:
        for err in s.get("errors", []):
            errors.append({
                "student_id": s["id"],
                "first_name": s.get("first_name", ""),
                "last_name": s.get("last_name", ""),
                "product_raw": err.get("raw_product", ""),
                "error_reason": err.get("reason", ""),
                "timestamp": "",
            })
    return errors


def _normalize_df_for_compare(df):
    if df is None or df.empty:
        return df
    df = df.copy()
    if "timestamp" in df.columns:
        df = df.drop(columns=["timestamp"])
    return df.sort_index(axis=1).reset_index(drop=True)


def _normalize_dtypes(df):
    """Convert string dtypes to object so got/expected match (Excel read can yield either)."""
    if df is None or df.empty:
        return df
    df = df.copy()
    for c in df.columns:
        if pd.api.types.is_string_dtype(df[c]):
            df[c] = df[c].astype(object)
    return df


def _normalize_empty_and_nan(df):
    """Treat NaN and empty string as equal: normalize to empty string for object columns."""
    if df is None or df.empty:
        return df
    df = df.copy()
    for c in df.columns:
        if df[c].dtype == object or pd.api.types.is_string_dtype(df[c]):
            df[c] = df[c].fillna("")
            df[c] = df[c].astype(str).replace("nan", "")
    return df


def _student_id_col(df):
    """Column name that identifies student (student_id or Student ID etc)."""
    for c in df.columns:
        if c and "student" in str(c).lower() and "id" in str(c).lower():
            return c
    return None


class TestPackageChoice(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.input_path = os.path.join(TESTS_DIR, "input.xlsx")
        cls.expected_summary = os.path.join(TESTS_DIR, "expected_run_summary.xlsx")
        cls.expected_errors = os.path.join(TESTS_DIR, "expected_errors.xlsx")

    def test_run_summary_and_errors(self):
        if not os.path.exists(self.input_path):
            self.skipTest("Put input.xlsx in package-choice/tests/")
        if not os.path.exists(self.expected_summary):
            self.skipTest("Put expected_run_summary.xlsx in package-choice/tests/")
        if not os.path.exists(self.expected_errors):
            self.skipTest("Put expected_errors.xlsx in package-choice/tests/")

        import data_handler_package

        students = data_handler_package.load_and_process_data(self.input_path)
        self.assertTrue(students, "load_and_process_data should return at least one student")

        summary_rows = _build_run_summary(students)
        error_rows = _collect_errors(students)
        err_cols = ["student_id", "first_name", "last_name", "product_raw", "error_reason", "timestamp"]
        error_df = pd.DataFrame(error_rows, columns=err_cols) if error_rows else pd.DataFrame(columns=err_cols)

        got_summary = _normalize_df_for_compare(pd.DataFrame(summary_rows))
        got_errors = _normalize_df_for_compare(error_df.copy())

        # Write got_errors.xlsx to tests folder (same as yearbook: output files for comparison)
        got_errors_path = os.path.join(TESTS_DIR, "got_errors.xlsx")
        got_errors.to_excel(got_errors_path, index=False)
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

        # Normalize dtypes (e.g. StringDtype vs object from Excel)
        got_summary = _normalize_dtypes(got_summary)
        exp_summary = _normalize_dtypes(exp_summary)
        got_errors = _normalize_dtypes(got_errors)
        exp_errors = _normalize_dtypes(exp_errors)
        # Treat NaN and "" as equal (got uses "", Excel reads blanks as NaN)
        got_summary = _normalize_empty_and_nan(got_summary)
        exp_summary = _normalize_empty_and_nan(exp_summary)

        pd.testing.assert_frame_equal(got_summary, exp_summary, check_like=True)
        # Compare errors only on student id + error_reason (other columns may differ)
        err_key_got = [c for c in (["error_reason"] + ([sid_col_got] if sid_col_got else [])) if c in got_errors.columns]
        err_key_exp = [c for c in (["error_reason"] + ([sid_col_exp] if sid_col_exp else [])) if c in exp_errors.columns]
        if err_key_got and err_key_exp:
            got_sub = got_errors[err_key_got].copy()
            exp_sub = exp_errors[err_key_exp].copy()
            if sid_col_got and sid_col_exp and sid_col_got != sid_col_exp:
                exp_sub = exp_sub.rename(columns={sid_col_exp: sid_col_got})
            sort_cols = [c for c in [sid_col_got or sid_col_exp, "error_reason"] if c in got_sub.columns and c in exp_sub.columns]
            if sort_cols:
                got_sub = got_sub.sort_values(sort_cols).reset_index(drop=True)
                exp_sub = exp_sub.sort_values(sort_cols).reset_index(drop=True)
            pd.testing.assert_frame_equal(got_sub, exp_sub, check_like=True)
        else:
            pd.testing.assert_frame_equal(got_errors, exp_errors, check_like=True)


if __name__ == "__main__":
    unittest.main()
