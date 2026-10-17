# 12_Roster
Automation. Don't have multiple tabs sync.

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
