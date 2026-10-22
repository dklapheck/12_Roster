# 12_Roster
Automation. Don't have multiple tabs sync.

## Setup instructions for a copied roster

1. Open **Extensions → Apps Script → Config.gs**. Replace `const ID` with the
   copied roster's spreadsheet ID: the text between `/d/` and `/edit` in its URL.
2. Optional: change `const WORKBOOK_NAME = '12RosterORN';` to your roster's name.
   This name appears in messages; the ID controls which spreadsheet the scripts use.
3. Put the separate Student Data Imports workbook URL in **Info!A2**. Leave
   **Info!A3** for the automatic refresh timestamp. **Info!C1:C5** holds these
   setup instructions.
4. Save the script, run `installRosterTools`, and authorize it if prompted.
   Reload the spreadsheet, then use **Update → Pull Data → Preview source changes**
   before refreshing.

For an existing roster moved to this layout, sync the updated `Config.gs`,
`RefreshStudentData.gs`, and `MasterSave.gs` together before running a preview or
refresh. They read the imports URL from A2 and write the refresh timestamp to A3.

The student refresh and new-student import populate Master's **Learning Coach**
from Upstream's `LC_COACH_NAME` and **LC Email** from `EMAIL_LEARNING_COACH`.
Fields are matched by header name, independent of column position. Existing
students with blank source values retain their current values, and the refresh
report flags blank/missing source fields. These two fields follow the normal
refresh flow; Welcome Email, SCC Completion, and ISMA keep their separate review.

After installing this update, run **Update → Pull Data → Preview source changes**
to check the Learning Coach and LC Email changes, then **Refresh students,
contacts and grades** to save them. Install both `RefreshStudentData.gs` and
`AddStudents.gs` in the bound Apps Script project.

Master's former **Grade Screenshot** column is now named **% Passing**. The
refresh and new-student import fill it from Upstream's `PCT_CLASSES_PASSING`,
matching students by Student Number and the column by header name. The source
uses a 0–100 scale: `42.9` means 42.9% and `0` is a valid value. Existing values
are retained and flagged when the source value/header is blank or missing.
Preview shows proposed changes without saving them. Install the updated
`RefreshStudentData.gs` and `AddStudents.gs`, then run the refresh to populate
the column. The live Master header has already been changed to `% Passing`.
Grade-report photo imports still target the separate **Grades** tab.
