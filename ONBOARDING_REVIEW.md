# Onboarding changes during refresh

Previously, Class Roster populated **Welcome Email**, **SCC Completion** and
**ISMA** only when a student was first added. Existing students' values were
left alone without comparison. These fields now have a separate review step.

1. Update the Class Roster export in the imports workbook identified by Info!A2.
2. Run **Update → Pull Data → Refresh students, contacts and grades**.
3. Near the top of the refresh report, compare **Current Master** with
   **Imported Class Roster**. Use **Accept all**, **Reject all**, or select a
   subset and use **Accept selected** / **Reject selected**. Each checkbox is
   one student's field, so fields can be reviewed independently.

Other refresh fields, contacts, grades, and automatic new-student additions
continue through their existing workflow. The three onboarding fields on
existing students remain unchanged until accepted, including when the imported
value is blank. Accepting a displayed `(blank)` deliberately clears that field.
Rejecting or closing the dialog keeps the current value; the same difference
can be reviewed again on a later refresh. **Preview source changes** compares
the fields without saving, caching a review, or offering accept/reject actions.

The report also lists Master students missing from **Upstream**, **Class
Roster**, **ProRoster**, and **CourseGrades**, using names and Student Numbers.
Missing exports are distinguished from missing students in the review notes.
An unavailable source cannot be compared. A toast points to the report, and
the existing removal workflow stays separate: these notices never remove
students. Upstream-missing students still on Class Roster or with future start
dates remain in the “not in Upstream yet” group. AddDrop errors no longer hide
the missing student's details.

Reviews expire after six hours of inactivity and are private to the editor who
ran the refresh. Acceptance rechecks the student ID, current Master value,
formula protection, source value and duplicate IDs under the shared save lock.
If a selected item changed in the meantime, the whole selected batch is refused
before any writes. Run the refresh again to obtain a new review. Values are
saved through `saveToMaster`, retaining its rollback behavior on write failure.
Source workbooks are read-only.

## Installation and verification

Include **OnboardingReview.gs** along with the updated **RefreshStudentData.gs**
and **RemoveStudents.gs** in the existing bound Apps Script project. Publishing
the repository alone does not update that live project. With a local clasp
checkout already linked to that project, pull the merged branch and run
`clasp push`, or copy those three files through the Apps Script editor.

Run `node --test tests/*.test.cjs` for synthetic regression checks. In Sheets,
first run **Preview source changes**, then exercise accept/reject on a small
known difference. Tests do not execute the live Apps Script service.
