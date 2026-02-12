# School Days Automation

Two automations live here:

- **yearbook-choice** — Yearbook photo choice entry  
- **package-choice** — Package choice entry  

Each folder has its own run script. Use the one you need.

---

## Yearbook choice

1. Put **one** Excel file in the **yearbook-choice** folder (same folder as the run script).
2. Double-click **Run_Yearbook_Choice.bat** (Windows) inside `yearbook-choice`.
3. Do Step 1 (validation), then Step 2 (click where the program asks), then Step 3 (automation runs).
4. To stop: move the mouse to a corner of the screen.

---

## Package choice

1. Put **one** Excel file in the **package-choice** folder (same folder as the run script).
2. Double-click **Run_Package_Choice.bat** (Windows) or **Run_Package_Choice.command** (Mac) inside `package-choice`.
3. Do Step 1 (validation), then Step 2 (click where the program asks), then Step 3 (automation runs).
4. To stop: move the mouse to a corner of the screen.

---

## Output files

All of these are created in a **reports** folder inside the automation folder you ran (e.g. `yearbook-choice/reports` or `package-choice/reports`).

| File | Meaning |
|------|--------|
| **yearbook-choice-errors-**…**.xlsx** | List of errors (validation or during run). Check this if something went wrong. |
| **yearbook-choice-completed-**…**.xlsx** | List of students that were finished successfully. |
| **package-choice-errors-**…**.xlsx** | List of errors (data problems or during run). Check this if something went wrong. |
| **package-choice-completed-**…**.xlsx** | List of students that were finished successfully. |
| **package-choice-run-summary-**…**.xlsx** | Summary of what the automation is entering (packages, choices, etc.) for the run. |

All report files are Excel (`.xlsx`).

---

## When the code is updated

If the code is updated, get the latest version like this:

1. Open **Git Bash**.
2. Go to this project folder (the one that contains `yearbook-choice` and `package-choice`).
3. Run: **`git pull`**

That’s it. You’ll then have the latest instructions and fixes.
