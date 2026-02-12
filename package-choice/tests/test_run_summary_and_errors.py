"""
Test package-choice: from input Excel, produce run-summary and errors Excel,
then compare to expected files in this folder. No GUI/coordinates step.

Put input.xlsx, expected_run_summary.xlsx, expected_errors.xlsx in this folder.
"""
import os
import sys
import tempfile
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

        with tempfile.TemporaryDirectory() as tmp:
            out_summary = os.path.join(tmp, "run_summary.xlsx")
            out_errors = os.path.join(tmp, "errors.xlsx")
            pd.DataFrame(summary_rows).to_excel(out_summary, index=False)
            error_df.to_excel(out_errors, index=False)

            got_summary = _normalize_df_for_compare(pd.read_excel(out_summary))
            got_errors = _normalize_df_for_compare(pd.read_excel(out_errors))
            exp_summary = _normalize_df_for_compare(pd.read_excel(self.expected_summary))
            exp_errors = _normalize_df_for_compare(pd.read_excel(self.expected_errors))

        pd.testing.assert_frame_equal(got_summary, exp_summary, check_like=True)
        pd.testing.assert_frame_equal(got_errors, exp_errors, check_like=True)


if __name__ == "__main__":
    unittest.main()
