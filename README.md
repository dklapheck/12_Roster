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

## Log successful parent contacts (Parent Communication)

Master's **ECC Attempt Date** column is now called **Contact Date**. It dates
both ECC attempts and parent contacts.

Setup, once:

1. On **Info**, add a row to the PowerSchoolSettings table directly under its
   last row (the table ends at the first blank Workflow cell). Put
   `Parent Contact` in **Workflow**. Copy the three date-field cells and any
   Opening Log Text from the ECC rows. Leave **Log Type value** and **Subtype
   value** blank to choose them in PowerSchool each time, or fill them with the
   values PowerSchool Helper 3.7.1 shows after you choose them by hand
   (its **Copy** buttons).
2. Install the updated `LogECC.gs`, `EccBulk.gs`, `EccBulkDialog.html`,
   `EccPowerSchool.gs` and `Menus.gs`, then run `installRosterTools` once. It
   renames the **ECC Attempt Date** header to **Contact Date** (only that
   header cell) and checks the new settings row. Reload the spreadsheet.
3. Install PowerSchool Helper 3.6.0 or later.

To log one student: type the note in **Parent Communication**, put the date in
**Contact Date**, select the student's row on Master, and choose **Update → ECC
& Demographics → Log selected student's parent contact**. The note is added to
**ECC Contact History** as `-- 11/19/2026 [Parent Contact]`, placed by date, and
PowerSchool opens a Parent Contact log for you to review and submit.

**Archive and Log All ECC** also lists every nonblank Parent Communication as a
separately checked Parent Contact entry, with its own PowerSchool link.
Parent Communication stays on Master after logging, like the ECC notes; clear it
yourself when done.

## 5+ Attempt Log

A **streak** is every `[Attempt]` in ECC Contact History since the last
successful contact (`[Conversation]`, `[Parent Contact]`, or an older
`-- 9/14/2026: ...` entry). When a streak reaches 5 attempts, the student is due
**one** 5+ Attempt Log. Only a successful contact starts a new streak.

| ECC Contact History | Color | Offered for logging? |
|---|---|---|
| Fewer than 5 attempts in the streak | none | No |
| 5+ attempts, not logged yet | red | Yes, once |
| Streak has a 5+ Attempt Log (more attempts keep it yellow) | yellow | No |
| A successful contact is newer than all of those | none | No (counting starts over) |

**Archive and Log All ECC** lists every due student as a checked
**5+ attempt log**. To log one student, select the row and choose **Update →
ECC & Demographics → Log selected student's 5+ attempt log**. No Master cell is
used: the script writes the entry from the history and dates it today:

```
-- 11/20/2026 [5+ Attempt Log]
6 unsuccessful student/parent contact attempts in a row since the last successful contact on 10/15/2026:
11/5 email to student and LC; 11/13 phone call/voicemail to LC; 11/16 phone call/voicemail to EC; 11/18 ParentSquare message to LC; 11/19 email to student and LC; 11/20 ParentSquare message to LC.
Contacted: student, LC and EC.
No response received.
```

**Who each attempt went to.** The log adds "to student", "to LC", "to LG" or
"to EC" after each attempt, and a closing **Contacted:** line. The script reads
this from the note: the codes **LC**, **LG** and **EC** anywhere (capitals);
"learning coach", "legal guardian" or "emergency contact" on a label line
(the first line, or a line ending with ":" such as `Email to LC Tameika
Moore (Mother):`); and greetings (`Hi Tameika,` matches Master's Learning
Coach, `Hello Scholar,` or the student's first name means the student). A
relationship such as `[Relative]` or `(Mother)` alone is not counted, so start
those notes with the code, for example `EC Voicemail for Danielle Hayes:`.
Each ECC attempt sent to PowerSchool also starts with the same `To: LC` line
when the script can tell; ECC Contact History keeps the note as written.

Each attempt gets one label from its note's first line (ParentSquare, email,
text, phone/voicemail, absent from live class, checked in that order), or the
start of the note when none fits. PowerSchool gets the same text on one line.
If the same batch also archives an attempt for that student, the log includes
it; if it archives a conversation or parent contact, the batch stops and asks
you to uncheck the log.

The log uses the **5+ Attempts** row of PowerSchoolSettings on Info (its own
Log Type and Subtype, or both blank to choose them in PowerSchool each time).
The row must sit directly under the other workflow rows with no blank row
between. PowerSchool Helper fills it like a Parent Contact log and keeps the
subtype's template text. With PowerSchool Helper 3.7.1 or later, its status
button says "5+ Attempts", and when Log Type and Subtype are blank it shows the
values you chose with Copy buttons, so you can paste them into this row.

**Colors and the Attempt Streak column.** The script fills a Master column
headed **Attempt Streak** (added at the right end the first time colors are
refreshed, if it is missing): `Due: 6 attempts`, `Logged 11/20/2026`,
`3 attempts`, or blank. Two conditional formatting rules color ECC Contact
History and Attempt Streak from that column: red for `Due:`, yellow for
`Logged`. The script puts them first, so they show over other rules such as
an every-other-row grey; your other rules are kept in their order. Because the
rules read the row's own value, the colors stay with the student when you sort.
Don't type in Attempt Streak or put a formula there; the script rewrites it.
It updates after every archive; after editing a history cell by hand, run
**Update → Maintenance → Refresh 5+ attempt colors**.

Install: add `AttemptStreak.gs`, update `LogECC.gs`, `EccBulk.gs`,
`EccBulkDialog.html` and `Menus.gs`, run `installRosterTools` (it checks the
5+ Attempts row), reload the sheet, then run **Refresh 5+ attempt colors** once.
Updating from the first version only needs the new `AttemptStreak.gs`; the
first refresh adds Attempt Streak and clears the old red/yellow cell fills.

## Archive and log ECC conversations and attempts

Install `EccBulk.gs` and `EccBulkDialog.html` along with the updated `LogECC.gs`,
`EccPowerSchool.gs`, and `Menus.gs` in the bound Apps Script project, then reload
the spreadsheet. This PR does not deploy the bound script automatically.

**Update → Archive and Log All ECC** reviews all Master students, including hidden
and filtered rows. Every conversation with content in ECC Overall, ECC Classes,
ECC Grades, or ECC Social is checked by default. ECC Follow-up is included in
conversation notes, but does not qualify a student by itself. Every nonblank
New Attempt Student Contact is offered as a separately checked attempt.

The review uses the existing ECC Date and Contact Date without changing
them. Dates before today in the spreadsheet's timezone are flagged. Correct
missing/invalid dates on Master and reopen the review, or uncheck those entries.
Search narrows the display only; Select all and Clear all apply to the full list.
The selection count includes checked entries hidden by search.

Click **Archive and Show Links** to save the checked entries into ECC Contact
History, newest date first, and display one PowerSchool link per entry. Nothing
opens automatically. A conversation, attempt and parent contact for the same
student remain separate entries. The notes and dates remain on Master.
Unchanged notes are offered and archived again on a new run; there is no
duplicate-note filter. Individual archive actions also archive repeated notes.

PowerSchool student selection may be shared between tabs. Close other student
pages, including Demographics, and open one log link at a time. Review and
submit the log, verify it appears under the intended student's name and number,
then close that PowerSchool tab before opening the next link.

The dialog marks clicked links as **opened** and counts link openings; it
cannot confirm that PowerSchool saved a log or enforce closing other tabs. If
a log cannot be found after Submit, investigate before retrying or opening
another student. Delete source notes manually after verifying the saved logs.

Clicking a displayed link again does not archive again. An expired review must be
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

The **List of Students Chronically Absent** shows each student's name, number,
and days needed to become compliant. Last-activity dates are omitted from both
the list and its attendance note; lesson-attempt counts remain in the note.
Sync `WigSnapshots.gs` to apply this to future previews and snapshots. Existing
snapshots are not rewritten by Save. After syncing, preview an empty week and
check that the list and attendance note omit activity dates while retaining
days needed and available lesson-attempt counts.

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
