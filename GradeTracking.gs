/** Grade history v1.0. All writes use the caller's CONFIG document lock.
 * Grades are points on a 0..100 scale (extra credit may exceed 100).
 * Snapshots change when source import metadata or roster/course/grade data changes.
 * Same-source retries keep the previous comparison and its alerts.
 */
const GRADE_TRACKING = (() => {
  const HISTORY = 'Grade History', REPORT = 'Grade Changes';
  const HEADERS = ['Snapshot ID', 'Captured At', 'Source Date', 'Source Import',
    'Student Number', 'Student Name', 'Course', 'Period', 'Grade',
    'Previous Grade', 'Change (pp)', 'Status'];
  const REPORT_HEADERS = ['Student flag', 'Student name', 'Student Number', 'Course',
    'Previous %', 'Current %', 'Change (pp)', 'Course flag', 'Status', 'Previous data', 'Current data'];
  const key = c => JSON.stringify([c.id, c.course, c.period || '']);
  const numeric = v => typeof v === 'number' && Number.isFinite(v);
  function grade(v) {
    if (numeric(v)) return v;
    const s = CONFIG.text(v).replace(/%$/, '');
    return /^-?\d+(\.\d+)?$/.test(s) && Number.isFinite(Number(s)) ? Number(s) : '';
  }
  function sorted(courses) { return courses.slice().sort((a, b) => key(a).localeCompare(key(b))); }
  function signature(courses, meta) {
    return JSON.stringify([meta.date || '', meta.imported || '', sorted(courses).map(c => [key(c), c.name, c.grade])]);
  }
  function sourceCourses(master, source) {
    const result = [];
    master.forEach(s => (source.byId.get(s.id) || []).forEach(c => result.push({
      id: s.id, name: CONFIG.text(s.values[0]) || [s.values[3], s.values[2]].filter(Boolean).join(' '),
      course: c.name, period: c.period || '', grade: grade(c.grade)})));
    if (!result.length) throw new Error('No roster course grades available; history was preserved.');
    return sorted(result);
  }
  /** Bootstrap from Grades before it is rewritten. Slot positions are never keys. */
  function baseline(headers, values, master, source) {
    const layout = GRADES_IMPORT.slots(headers), people = new Map(master.map(s => [s.id, s]));
    const courses = [];
    values.forEach(row => {
      const id = CONFIG.studentId(row[1]), person = people.get(id);
      if (!person) return;
      layout.all.forEach(slot => {
        const course = CONFIG.text(row[slot.course]);
        if (!course) return;
        const matching = (source.byId.get(id) || []).find(c => c.name === course);
        courses.push({id, name: CONFIG.text(person.values[0]) || [person.values[3], person.values[2]].filter(Boolean).join(' '),
          course, period: matching ? matching.period || '' : '', grade: grade(row[slot.pct])});
      });
    });
    return sorted(courses);
  }
  function latest(rows) {
    if (!rows.length) return [];
    const id = rows[rows.length - 1][0];
    return rows.filter(r => r[0] === id);
  }
  function unpack(rows) {
    return rows.filter(r => r[11] !== 'No longer in source').map(r => ({
      id: CONFIG.studentId(r[4]), name: r[5], course: r[6], period: r[7], grade: grade(r[8])}));
  }
  /** Pure comparison, including missing/new courses and numeric boundary handling. */
  function plan(history, courses, meta, captured, snapshotId) {
    const prior = latest(history), previous = unpack(prior);
    const oldMeta = prior.length ? {date: prior[0][2], imported: prior[0][3]} : {};
    if (prior.length && signature(previous, oldMeta) === signature(courses, meta)) {
      return {rows: prior, append: [], unchanged: true};
    }
    const byKey = new Map(previous.map(c => [key(c), c])), seen = new Set();
    const rows = courses.map(c => {
      const k = key(c);
      if (seen.has(k)) throw new Error('Duplicate student/course/period in grade snapshot.');
      seen.add(k);
      const p = byKey.get(k), before = p ? p.grade : '';
      const delta = numeric(before) && numeric(c.grade) ? Math.round((c.grade - before) * 1e8) / 1e8 : '';
      const status = !prior.length ? 'Baseline' : !p ? 'New course' : !numeric(c.grade) ? 'No current grade' :
        !numeric(before) ? 'No previous numeric grade' : delta < -10 ? 'RED FLAG' : delta < 0 ? 'Decreased' : delta > 0 ? 'Increased' : 'Unchanged';
      return [snapshotId, captured, meta.date || '', meta.imported || '', c.id, c.name,
        c.course, c.period || '', c.grade, before, delta, status];
    });
    previous.forEach(p => {
      if (!seen.has(key(p))) rows.push([snapshotId, captured, meta.date || '', meta.imported || '', p.id,
        p.name, p.course, p.period || '', '', p.grade, '', 'No longer in source']);
    });
    return {rows, append: rows, unchanged: false};
  }
  function reportRows(rows, history) {
    const flagged = new Set(rows.filter(r => r[11] === 'RED FLAG').map(r => String(r[4])));
    const prior = history.length ? latest(history.filter(r => r[0] !== rows[0][0])) : [];
    const previousDate = prior.length ? prior[0][2] || prior[0][3] : '';
    return rows.map(r => [flagged.has(String(r[4])) ? 'RED FLAG' : '', r[5], r[4], r[6],
      r[9], r[8], r[10], r[11] === 'RED FLAG' ? 'RED FLAG' : '', r[11], previousDate,
      r[2] || r[3]]).sort((a, b) => Number(b[0] === 'RED FLAG') - Number(a[0] === 'RED FLAG') ||
        String(a[1]).localeCompare(String(b[1])) || (numeric(a[6]) ? a[6] : 0) - (numeric(b[6]) ? b[6] : 0) || String(a[3]).localeCompare(String(b[3])));
  }
  /** Master course/percentage cells to underline for the latest comparison.
   * headers and values use zero-based arrays; returned rows/columns are 1-based. */
  function underlineTargets(headers, values, rows) {
    const layout = GRADES_IMPORT.findSlots(headers, 0);
    if (!layout) return {all: [], decreased: []};
    const idColumn = headers.indexOf('Student Number');
    if (idColumn < 0) throw new Error('Master is missing Student Number.');
    const down = new Set(rows.filter(r => ['Decreased', 'RED FLAG'].includes(r[11]))
      .map(r => JSON.stringify([String(r[4]), CONFIG.text(r[6])])));
    const all = [], decreased = [];
    values.forEach((row, i) => {
      const id = CONFIG.studentId(row[idColumn]);
      if (!id) return;
      layout.all.forEach(slot => {
        const course = CONFIG.text(row[slot.course]);
        const target = {id, row: i + 2, course: slot.course + 1, pct: slot.pct + 1};
        all.push(target);
        if (course && down.has(JSON.stringify([id, course]))) decreased.push(target);
      });
    });
    return {all, decreased};
  }
  /** Underline means the course decreased in the latest saved comparison.
   * Identical-source retries keep the comparison rows, so they keep the underline. */
  function underlineMasterDeclines(ss, rows) {
    const sheet = ss.getSheetByName('Master');
    if (!sheet) throw new Error('Master tab is missing.');
    const info = CONFIG.headerColumns(sheet), height = Math.max(sheet.getLastRow() - 1, 0);
    if (!height) return {coursesUnderlined: 0, studentsUnderlined: 0};
    const targets = underlineTargets(info.headers, sheet.getRange(2, 1, height, info.width).getValues(), rows);
    const a1 = t => CONFIG.columnLetter(t.course) + t.row + ':' + CONFIG.columnLetter(t.pct) + t.row;
    if (targets.all.length) sheet.getRangeList(targets.all.map(a1)).setFontLine('none');
    if (targets.decreased.length) sheet.getRangeList(targets.decreased.map(a1)).setFontLine('underline');
    return {coursesUnderlined: targets.decreased.length,
      studentsUnderlined: new Set(targets.decreased.map(t => t.id)).size};
  }
  function historyRows(ss) {
    const sheet = ss.getSheetByName(HISTORY);
    if (!sheet || !sheet.getLastRow()) return [];
    const actual = sheet.getRange(1, 1, 1, HEADERS.length).getValues()[0];
    if (!HEADERS.every((h, i) => actual[i] === h)) throw new Error('Grade History headers were changed; restore them before refreshing.');
    return sheet.getLastRow() > 1 ? sheet.getRange(2, 1, sheet.getLastRow() - 1, HEADERS.length).getValues() : [];
  }
  function size(sheet, height, width) {
    if (height > sheet.getMaxRows()) sheet.insertRowsAfter(sheet.getMaxRows(), height - sheet.getMaxRows());
    if (width > sheet.getMaxColumns()) sheet.insertColumnsAfter(sheet.getMaxColumns(), width - sheet.getMaxColumns());
  }
  function append(ss, rows) {
    if (!rows.length) return;
    const sheet = ss.getSheetByName(HISTORY) || ss.insertSheet(HISTORY);
    const start = Math.max(2, sheet.getLastRow() + 1);
    size(sheet, start + rows.length - 1, HEADERS.length);
    if (!sheet.getLastRow()) {
      sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setBackground('#eeeeee').setFontWeight('bold').setWrap(true);
      sheet.setColumnWidth(1, 285); sheet.setColumnWidths(2, 3, 220);
      sheet.setColumnWidth(5, 120); sheet.setColumnWidth(6, 180);
      sheet.setColumnWidth(7, 340); sheet.setColumnWidth(8, 290);
      sheet.setColumnWidths(9, 3, 115); sheet.setColumnWidth(12, 180);
      sheet.setRowHeight(1, 36);
    }
    const range = sheet.getRange(start, 1, rows.length, HEADERS.length);
    range.setValues(CONFIG.prepareValues(range, rows));
    sheet.setFrozenRows(1);
    sheet.getRange(start, 9, rows.length, 2).setNumberFormat('0.00"%"');
    sheet.getRange(start, 11, rows.length, 1).setNumberFormat('+0.00;-0.00;0.00');
    if (sheet.getFilter()) sheet.getFilter().remove();
    sheet.getRange(1, 1, start + rows.length - 1, HEADERS.length).createFilter();
    SpreadsheetApp.flush();
  }
  function render(ss, rows, history) {
    const sheet = ss.getSheetByName(REPORT) || ss.insertSheet(REPORT);
    // Full-width title cells must not cross a frozen-column boundary.
    sheet.setFrozenColumns(0);
    const values = reportRows(rows, history), width = REPORT_HEADERS.length;
    size(sheet, Math.max(5, values.length + 4), width);
    if (sheet.getLastRow() >= 4) {
      const actual = sheet.getRange(4, 1, 1, width).getValues()[0];
      if (actual.some(Boolean) && !REPORT_HEADERS.every((h, i) => actual[i] === h)) throw new Error('Grade Changes headers were changed; existing report preserved.');
    }
    if (sheet.getFilter()) sheet.getFilter().remove();
    if (sheet.getLastRow() > 4) sheet.getRange(5, 1, sheet.getLastRow() - 4, width).clearContent();
    const flagged = new Set(rows.filter(r => r[11] === 'RED FLAG').map(r => r[4]));
    sheet.getRange(1, 1, 1, width).merge().setValue('Grade Changes | ' + flagged.size + ' students flagged | drops greater than 10 percentage points').setFontWeight('bold');
    sheet.getRange(2, 1, 1, width).merge().setValue('Source: ' + (rows[0][2] || rows[0][3]) + ' | Captured: ' + rows[0][1] + ' | Same-source refreshes keep alerts. Initial baseline date is unknown.').setWrap(true);
    sheet.getRange(4, 1, 1, width).setValues([REPORT_HEADERS]).setBackground('#eeeeee').setFontWeight('bold').setWrap(true);
    if (values.length) {
      const range = sheet.getRange(5, 1, values.length, width);
      range.setValues(CONFIG.prepareValues(range, values)).setVerticalAlignment('top');
      sheet.getRange(5, 5, values.length, 2).setNumberFormat('0.00"%"');
      sheet.getRange(5, 7, values.length, 1).setNumberFormat('+0.00;-0.00;0.00');
      sheet.getRange(5, 4, values.length, 1).setWrap(true);
      sheet.getRange(4, 1, values.length + 1, width).createFilter();
      sheet.setConditionalFormatRules([SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('=$A5="RED FLAG"').setBackground('#fce8e6').setFontColor('#a50e0e').setRanges([range]).build()]);
    }
    sheet.setFrozenRows(4);
    sheet.setColumnWidth(1, 110); sheet.setColumnWidth(2, 180); sheet.setColumnWidth(3, 120);
    sheet.setColumnWidth(4, 340); sheet.setColumnWidths(5, 4, 115);
    sheet.setColumnWidth(9, 180); sheet.setColumnWidths(10, 2, 190);
    sheet.setRowHeight(1, 30); sheet.setRowHeight(2, 42); sheet.setRowHeight(4, 36);
    SpreadsheetApp.flush();
    return {studentsFlagged: flagged.size, coursesFlagged: rows.filter(r => r[11] === 'RED FLAG').length};
  }
  /** Called BEFORE Grades writes. Saves an initial baseline even if the import later fails. */
  function seed(ss, headers, values, master, source) {
    if (historyRows(ss).length) return;
    const courses = baseline(headers, values, master, source);
    if (!courses.length) return;
    const captured = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd HH:mm:ss z');
    append(ss, plan([], courses, {date: '', imported: 'Existing Grades (date unknown)'}, captured, Utilities.getUuid()).append);
  }
  /** Called only AFTER a successful Grades import, including successful no-op imports. */
  function capture(ss, master, source) {
    const history = historyRows(ss), courses = sourceCourses(master, source);
    const meta = {date: source.importDate || '', imported: source.importStamp || ''};
    const captured = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd HH:mm:ss z');
    const result = plan(history, courses, meta, captured, Utilities.getUuid());
    append(ss, result.append);
    const counts = render(ss, result.rows, history);
    const underlined = underlineMasterDeclines(ss, result.rows);
    SpreadsheetApp.flush();
    return Object.assign({unchanged: result.unchanged}, counts, underlined);
  }
  return Object.freeze({HEADERS, REPORT_HEADERS, grade, baseline, sourceCourses, plan, reportRows,
    underlineTargets, underlineMasterDeclines, seed, capture});
})();
