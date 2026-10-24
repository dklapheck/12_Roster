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

Click **Archive and Open PowerSchool Tabs** to save the checked entries into
ECC Contact History, newest date first, and open one PowerSchool tab per entry.
A conversation and attempt for the same student create two history entries and
two tabs. The notes and dates remain on Master. Review and submit each log in
PowerSchool, then manually delete the source notes when done. Unchanged notes
are offered and archived again on a new run; there is no duplicate-note filter.
Individual ECC archive actions also archive repeated notes.

The dialog reserves blank tabs during the button click and sends them to
PowerSchool only after the archive succeeds. If the browser blocks tabs, use
**Open remaining tabs** or the individual entry links. Reopening links does not
archive again. An expired review must be reopened; changes to reviewed notes or
dates cancel the selected batch before saving. Sorts are supported because
students are matched by Student Number. Script saves share the roster's lock
and rollback path; direct manual sheet edits are not blocked by that lock.

Archive and Log All ECC and Open Student Demographics are both at the top of
Update and repeated under **ECC & Demographics**. The remaining submenus are
Call / SCC, Refresh Roster Data, Grade Report Photos, WIG, and Maintenance.
