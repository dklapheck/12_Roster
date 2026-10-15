# Grade changes and history

`GradeTracking.gs` adds two tabs in 12RosterORN:

- **Grade Changes** shows the last source comparison, previous and current course grades, change in percentage points, and student/course flags. Any student with a course drop greater than 10 points has all their course rows highlighted and is sorted first.
- **Grade History** keeps full course snapshots, including source date/import timestamp, capture time, previous grade, delta, and status. Keep its headers and saved rows intact.

85% to 75% is not flagged; 85% to 74.99% is flagged. This is an absolute percentage-point difference, not a relative percent decrease. Extra credit above 100 is supported. Blank grades, new courses and new terms do not count as drops. Courses are matched by student number, full course name and term, independently of visible course slots. All courses in the source for current Master students are tracked, including overflow courses without a visible Grades slot. A course-name change starts a new course comparison.

## Install

Copy or push all three changed `.gs` files into the existing Apps Script project: **GradeTracking.gs**, **GradesImport.gs**, and **RefreshStudentData.gs**. Do not replace the project with only these files; retain the other roster scripts. GitHub changes alone do not deploy to Apps Script. No new trigger, button assignment or authorization scope is required.

Run the existing **Update → Pull Data → Refresh students, contacts and grades** command. Preview does not write history. The normal saved refresh captures history only after Grades successfully finishes; successful imports with no changed Grades cells also update tracking. History and Grades share the document lock. A report failure is surfaced in the refresh notes and can be retried.

On first use, the existing Grades values are saved before the import overwrites them, labeled **Existing Grades (date unknown)**. This bootstrap assumes same-name existing courses belong to the source's current term because Grades does not store term identifiers. New terms after that have separate keys. Missing source courses are marked **No longer in source**, not treated as zero. Students absent from CourseGrades keep their Grades values under existing importer behavior; tracking marks their courses as missing in source and does not compute drops.

The initial report prepared on October 15, 2026 compares the saved Grades values with the October 15 source import without refreshing Master or Grades. Those two snapshots are already saved in the tabs. The next refresh with the same source reuses the comparison.

## When a snapshot is saved

The source's `IMPORT_DATE` / `IMPORT_DATETIME`, student/course membership, names and numeric grades are compared with the last snapshot. If any changes, a full new snapshot is appended. Repeating an identical refresh does not duplicate history or clear the last comparison's alerts. A later import produces a new comparison, even if its grades are identical. Grade changes within the same import timestamp are also captured. No tab per refresh is created.

`CourseGrades!A3` is not the trigger: it was October 9 while the source row import dates were October 15 during this implementation. Apps Script refresh explicitly captures after its successful import; an `onEdit` trigger is unnecessary.

The report is maintained by the script. Keep personal notes on a different tab. Do not rename either tracking tab without updating the script. History starts with the saved roster baseline; grades from older overwritten imports cannot be reconstructed automatically.

## Verification

Run `node --test tests/grade-tracking.test.cjs`. Synthetic fixtures cover the strict threshold, zero and extra credit, matching across row/slot moves, flags for all courses of an affected student, unchanged-source retries and recovery of report comparisons, new terms/courses, missing grades, overflow courses, and duplicate keys. Native Apps Script execution still needs verification after installation.
