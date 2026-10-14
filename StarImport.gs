/** STAR import v1.0 (STAR Import tab -> BOY percentiles and Latest STAR date/score).
 * Requires Config.gs. Called by RefreshStudentData.gs (Student refresh menu);
 * it has no menu of its own. Source: the STAR Import tab of the imports
 * workbook in Info!H2 (read only), pasted straight from the Renaissance export.
 *
 * Columns are found by header name, so the export's column order may change.
 * Used: Student ID, Application Name (Star Math / Star Reading), Completed Date
 * (UTC text such as 2026-08-28T00:29:15.471Z), Score (unified scale score),
 * Percentile Rank and Season/Screening Window Name (BOY / MOY / EOY).
 * Rows with no Application Name are students who have not tested; skipped.
 *
 * Completed Date is converted to the roster's time zone, so a test finished at
 * 5:29 PM Pacific on Aug 27 is dated Aug 27 (the export's UTC text says Aug 28).
 * Upstream's STAR_2TO12_* dates follow the same rule.
 */
const HR10STARIMPORT = (() => {
  const SOURCE_TAB = 'STAR Import';
  const H = {
    id: 'Student ID', app: 'Application Name', completed: 'Completed Date',
    score: 'Score', percentile: 'Percentile Rank', window: 'Season/Screening Window Name'
  };
  const REQUIRED = [H.id, H.app, H.completed, H.score];
  const text = v => HR10CFG.text(v);
  const blank = v => HR10CFG.blank(v);

  function subjectOf(app) {
    const a = text(app).toLowerCase();
    return a.includes('math') ? 'Math' : a.includes('reading') ? 'Reading' : '';
  }

  /** Instant of a Completed Date cell: ISO text from the export, or a Sheets date. */
  function instant(v) {
    if (v instanceof Date) return isNaN(v.getTime()) ? null : v;
    const s = text(v);
    if (!s) return null;
    const d = new Date(s);
    return isNaN(d.getTime()) ? null : d;
  }

  /** Whole number 0-1400ish for scores, 1-99 for percentiles; NaN otherwise. */
  function number(v) {
    if (typeof v === 'number') return v;
    const s = text(v);
    return /^\d+(\.\d+)?$/.test(s) ? Number(s) : NaN;
  }

  /** Parse the tab. Returns {byId: Map(studentNumber -> {Math: [tests], Reading: [tests]}), issues}.
   * Each test: {at (Date instant), ymd ('yyyy-MM-dd' in tz), score, percentile, window, row}.
   * Throws only if the tab or a required header is missing. */
  function parse(sheet, tz) {
    if (!sheet) throw new Error('the imports workbook has no ' + SOURCE_TAB + ' tab');
    const t = HR10CFG.table(sheet, REQUIRED);
    const col = h => t.columns.get(h);
    const byId = new Map(), issues = [];
    t.rows.forEach((row, i) => {
      const n = i + 2;
      const get = h => (col(h) === undefined ? '' : row[col(h)]);
      if (row.every(blank)) return;
      let key;
      try { key = HR10CFG.studentId(get(H.id)); }
      catch (e) { issues.push(SOURCE_TAB + ' row ' + n + ': ' + e.message + ' Row skipped.'); return; }
      if (!key) { issues.push(SOURCE_TAB + ' row ' + n + ' has no Student ID; skipped.'); return; }
      if (blank(get(H.app))) return; // listed but not tested yet
      const subject = subjectOf(get(H.app));
      if (!subject) return; // other Star apps (e.g. Early Literacy) are not roster columns
      const at = instant(get(H.completed));
      const score = number(get(H.score));
      if (!at || !Number.isFinite(score) || score < 0) {
        issues.push(SOURCE_TAB + ' row ' + n + ' (' + key + ' ' + subject + '): missing or invalid Completed Date/Score; skipped.');
        return;
      }
      const p = number(get(H.percentile));
      const test = {at, ymd: Utilities.formatDate(at, tz, 'yyyy-MM-dd'), score,
        percentile: Number.isFinite(p) ? p : null, window: text(get(H.window)).toUpperCase(), row: n};
      if (!byId.has(key)) byId.set(key, {Math: [], Reading: []});
      byId.get(key)[subject].push(test);
    });
    byId.forEach(s => ['Math', 'Reading'].forEach(k => s[k].sort((a, b) => a.at - b.at)));
    return {byId, issues};
  }

  /** Read STAR Import from the imports workbook. */
  function sourceFrom(importsBook, tz) {
    return parse(importsBook.getSheetByName(SOURCE_TAB), tz);
  }

  /** Most recent test for one student/subject, or null. */
  function latest(source, key, subject) {
    const tests = source && source.byId.get(key);
    const list = tests ? tests[subject] : [];
    return list.length ? list[list.length - 1] : null;
  }

  /** The student's first BOY test for one subject (the screening score to keep), or null. */
  function boy(source, key, subject) {
    const tests = source && source.byId.get(key);
    return (tests ? tests[subject] : []).find(t => t.window === 'BOY') || null;
  }

  return Object.freeze({SOURCE_TAB, parse, sourceFrom, latest, boy});
})();
