# Package-choice tests

Tests the **data-processing only** (no coordinates, no app). Compares produced run-summary and errors Excel to your expected files.

## Setup

Put your sample files in **this folder** (`package-choice/tests/`):

- **input.xlsx** — sample input Excel (same format as you use for the automation)
- **expected_run_summary.xlsx** — the run-summary output you consider correct
- **expected_errors.xlsx** — the errors output you consider correct (can be empty with just headers)

## Run

From **this folder** (`package-choice/tests/`):

```bash
python test_run_summary_and_errors.py
```

Or from repo root:

```bash
python -m unittest package-choice.tests.test_run_summary_and_errors -v
```

(If you get import errors from repo root, run from `package-choice/tests/` instead.)
