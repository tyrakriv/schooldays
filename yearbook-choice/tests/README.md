# Yearbook-choice tests

Tests the **data-processing only** (validation + run-summary, no GUI). Compares produced run-summary and errors Excel to your expected files.

## Setup

Put your sample files in **this folder** (`yearbook-choice/tests/`):

- **input.xlsx** — raw input Excel (same format as Step 1 validation input: Student ID, Yearbook Date, Yearbook Selection, Student Last Name, etc.)
- **expected_run_summary.xlsx** — the run-summary output you consider correct (Student ID, Last Name, Yearbook Selection)
- **expected_errors.xlsx** — the validation errors output you consider correct (Student ID, Last Name, Yearbook Selection, error_reason; can be empty with just headers)

## Run

From **this folder** (`yearbook-choice/tests/`):

```bash
python test_run_summary.py
```
