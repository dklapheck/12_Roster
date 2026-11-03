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
   Reload the spreadsheet, then use **Update → Refresh Roster Data → Preview source changes**
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

After installing this update, run **Update → Refresh Roster Data → Preview source changes**
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

## Archive and log ECC conversations and attempts

Install `EccBulk.gs` and `EccBulkDialog.html` along with the updated `LogECC.gs`,
`EccPowerSchool.gs`, and `Menus.gs` in the bound Apps Script project, then reload
the spreadsheet. This PR does not deploy the bound script automatically.

**Update → Archive and Log All ECC** reviews all Master students, including hidden
and filtered rows. Every conversation with content in ECC Overall, ECC Classes,
ECC Grades, or ECC Social is checked by default. ECC Follow-up is included in
conversation notes, but does not qualify a student by itself. Every nonblank
New Attempt Student Contact is offered as a separately checked attempt.

The review uses the existing ECC Date and ECC Attempt Date without changing
them. Dates before today in the spreadsheet's timezone are flagged. Correct
missing/invalid dates on Master and reopen the review, or uncheck those entries.
Search narrows the display only; Select all and Clear all apply to the full list.
The selection count includes checked entries hidden by search.

Click **Archive and Show Links**. This saves the checked entries into ECC Contact
History, newest date first, and displays a separate PowerSchool link for each
entry. No PowerSchool tab opens automatically. A conversation and attempt for
the same student remain two separate links.
The notes and dates remain on Master. Review and submit each log in
PowerSchool, then manually delete the source notes when done. Unchanged notes
are offered and archived again on a new run; there is no duplicate-note filter.
Individual ECC archive actions also archive repeated notes.

PowerSchool student selection may be shared between tabs. Close all other
PowerSchool student pages, including Demographics, before opening a link.
Open only one link at a time. Review the student and fields, submit manually,
verify the saved log under that student's name and number, then close the
PowerSchool tab before opening the next link. The dialog does not enforce
tab closure or verify saved logs. Its opened-link count and checkmarks only
record clicks, not successful submissions. If a log cannot be found after
Submit, investigate before retrying or moving to the next entry.

Reopening a displayed link does not archive again. An expired review must be
reopened; changes to reviewed notes or
dates cancel the selected batch before saving. Sorts are supported because
students are matched by Student Number. Script saves share the roster's lock
and rollback path; direct manual sheet edits are not blocked by that lock.

Archive and Log All ECC, Log selected student's conversation, Log selected
student's attempt, and Open Student Demographics are at the top of Update and
repeated under **ECC & Demographics**. Select one student's row on Master before
using either individual logging shortcut. Each archives the existing dated
notes into ECC Contact History through `saveToMaster()`, then shows the existing
PowerSchool handoff dialog. Current notes and dates stay on Master. Close all
other PowerSchool student tabs before opening the handoff, review the student
and fields, and submit manually. Retrying the archive action archives the notes
again; it does not detect whether the PowerSchool log was submitted.

The remaining submenus are Call / SCC, Refresh Roster Data, Grade Report Photos,
WIG, and Maintenance. Install the updated `Menus.gs` and reload the spreadsheet
to see the new quick-access items. The existing `LogECC.gs` and
`EccPowerSchool.gs` functions are reused without changing the extension handoff.

## WIG snapshots available Monday

Install the updated `WigSnapshots.gs`. **Update → WIG → Preview this week's
snapshot** and **Save this week's snapshot** target the Tuesday date belonging
to the current Monday–Sunday week, using the spreadsheet's timezone. For
example, Monday, November 2, 2026 targets the existing November 3 row, just as
Tuesday, November 3 does. Sunday, November 1 still targets October 27. Keep the
Tuesday labels in WIG column A; no date migration is needed.

Snapshots remain manual. Preview changes no cells. Save still requires exactly
one matching dated row and preserves a week that already contains snapshot
data or formulas, including when a Monday snapshot is run again on Tuesday.
Missing rows are reported rather than added automatically.

GitHub changes do not deploy automatically. Compare the two changed `.gs`
files with the bound Apps Script project before uploading them, retaining any
live-only edits. This update does not change PowerSchool settings or its URL
payload format.

Browser checks after syncing and reloading:

- Preview WIG on Monday: the report should name the next day's Tuesday row and
  leave cells unchanged. Save an empty week when ready, then preview/save again
  to confirm the existing snapshot is preserved.
- Select one Master student with valid ECC dates and notes. Try the conversation
  shortcut, then the attempt shortcut separately. Confirm the correct entry is
  archived with its existing date and that the PowerSchool dialog opens.
- Keep only one PowerSchool student tab open. Verify the student number, date,
  workflow settings, and one-line notes before submitting manually, then verify
  the saved log and close that tab before testing the other shortcut.
