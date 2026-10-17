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
