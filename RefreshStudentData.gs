/** 10RosterHR source refresh v1.7 (adds new students first). Replace the old RefreshStudentData file.
 * @NotOnlyCurrentDoc
 * Requires Config.gs, RosterSync v1.3, RemoveStudents.gs, GradesImport.gs, StarImport.gs and AddStudents.gs v1.1.
 * v1.7: after a saved refresh, Call Due colors are reapplied (CallDueColors.gs).
 * v1.7: Upstream GE_NEXT_CALL_DUE fills Call Due (a blank source keeps the cell).
 * v1.7: contacts are also filled for Master students not in Upstream yet, and the
 * Contacts Student Number is written as a number (text did not match on the Call tab).
 * v1.7: step 1 adds every student who is on Class Roster or active in Upstream but
 * not on Master (AddStudents.gs; Preview only lists them). Master students who are
 * not in Upstream but are still on Class Roster, or whose START DATE is in the
 * future, are listed as "not in Upstream yet" and are not offered for removal.
 * v1.6: STAR Import tab -> BOY percentiles and Latest STAR date/score.
 * STAR: BOY Reading/Math Percentile are filled from the STAR Import tab's BOY
 * test when blank (a value already in Master is kept). Latest Reading/Math
 * Date/Score take the newest test from STAR Import or Upstream STAR_2TO12_*.
 * If STAR Import is missing or unreadable, Upstream alone is used, as before.
 * Course grades from the CourseGrades tab are part of the Master refresh
 * (synced to SCC, ECC and STAR). The Grades tab is written after Master is
 * saved, under its own lock; a Grades tab problem is reported in the dialog and
 * never undoes the Master refresh. Menus are built by Menus.gs. Info!H2 holds the separate imports workbook URL.
 * Source is READ ONLY. Planner performs no writes; shared helper saves everything.
 */
const HR10REFRESH = (() => {
  const MAP = [
    ['STUDENT_LAST_NAME', 'LAST NAME'], ['STUDENT_FIRST_NAME', 'FIRST NAME'],
    ['GRADE', 'GRADE'], ['ENROLL_DATE_PS', 'START DATE'], ['SCHOOL', 'SCHOOL'],
    ['SPED504PLAN', 'SPED'], ['EMAIL_O365', 'Student Email'],
    ['COUNSELOR', 'Counselor'], ['COUNSELOR_EMAIL', 'Counselor Email'],
    ['GE_NEXT_CALL_DUE', 'Call Due']
  ];
  const DATE_FIELDS = ['START DATE', 'Call Due'];
  // Blank in Upstream is normal for these (e.g. no GE call scheduled); keep the cell, no review note.
  const QUIET_BLANK = ['Call Due'];
  const CONTACT_HEADERS = ['Student Number', 'Student Source Name', 'Gender', 'Student Phone',
    'Contact Name', 'Relationship', 'Source Phone', 'Source Email', 'Effective Phone',
    'Effective Email', 'Learning Coach', 'Contact Key'];
  const text = v => HR10CFG.text(v);
  const blank = v => HR10CFG.blank(v);
  const same = (a, b) => HR10CFG.same(a, b);
  const id = v => HR10CFG.studentId(v);
  const book = () => HR10CFG.workbook();
  const table = (sheet, required) => HR10CFG.table(sheet, required);
  function source(ss) {
    const upstream = HR10CFG.sourceWorkbook(ss);
    const u = table(upstream.getSheetByName('Upstream'), ['STUDENT_NUMBER', 'CURRENT_ACTIVE_STUDENT']);
    const p = table(upstream.getSheetByName('ProRoster'),
      ['Student Number', 'Name', 'Gender', 'Phone', 'Contact Name', 'Contact Phone', 'Contact Email']);
    const students = new Map();
    u.rows.forEach((row, offset) => {
      const key = id(row[u.columns.get('STUDENT_NUMBER')]);
      if (!key) {
        if (row.some(v => !blank(v))) throw new Error('Upstream row ' + (offset + 2) + ' has no STUDENT_NUMBER.');
        return;
      }
      if (students.has(key)) throw new Error('Upstream has duplicate STUDENT_NUMBER ' + key);
      const active = ['Y', 'YES', 'TRUE', '1'].includes(text(row[u.columns.get('CURRENT_ACTIVE_STUDENT')]).toUpperCase());
      students.set(key, {row, active});
    });
    let grades = null, gradesError = '';
    try { grades = HR10GRADEIMPORT.sourceFrom(upstream); }
    catch (e) { gradesError = e.message; }
    let star = null, starError = '';
    try { star = HR10STARIMPORT.sourceFrom(upstream, ss.getSpreadsheetTimeZone()); }
    catch (e) { starError = e.message; }
    return {u, p, students, book: upstream, grades, gradesError, star, starError};
  }
  function date(v, label) {
    if (v instanceof Date && !isNaN(v.getTime())) return new Date(v.getTime());
    const s = text(v);
    let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/), y, mo, d;
    if (m) { y = +m[1]; mo = +m[2]; d = +m[3]; }
    else {
      m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
      if (m) { y = +m[3]; mo = +m[1]; d = +m[2]; }
    }
    if (m) {
      const check = new Date(Date.UTC(y, mo - 1, d));
      if (check.getUTCFullYear() === y && check.getUTCMonth() === mo - 1 && check.getUTCDate() === d) {
        return Utilities.parseDate(y + '-' + mo + '-' + d, book().getSpreadsheetTimeZone(), 'yyyy-M-d');
      }
    }
    throw new Error('Invalid date for ' + label + '. Use a Sheets date or YYYY-MM-DD/MM/DD/YYYY.');
  }
  function nameKey(v) {
    return text(v).replace(/\([^)]*\)/g, '').toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim().split(/\s+/).sort().join(' ');
  }
  function phoneKey(v) { return text(v).replace(/\([^)]*\)/g, '').replace(/\D/g, ''); }
  function emailKey(v) { return text(v).replace(/\([^)]*\)/g, '').trim().toLowerCase(); }
  function roleKey(v) { return text(v).toLowerCase().replace(/\s+/g, ' '); }
  function planContacts(ss, masterIds, data, report) {
    const t = table(ss.getSheetByName('Contacts'), CONTACT_HEADERS);
    const old = new Map(), fallback = new Map();
    const append = (map, key, value) => {
      if (!map.has(key)) map.set(key, []);
      map.get(key).push(value);
    };
    t.rows.forEach((row, i) => {
      if (row.every(blank)) return;
      const key = id(row[t.columns.get('Student Number')]);
      if (!key) throw new Error('Contacts row ' + (i + 2) + ' has no Student Number.');
      const ck = text(row[t.columns.get('Contact Key')]);
      if (ck && !ck.startsWith(key + '|')) throw new Error('Contacts row ' + (i + 2) + ': Contact Key does not match Student Number.');
      const fk = JSON.stringify([key, nameKey(row[t.columns.get('Contact Name')])]);
      const entry = {row, number: i + 2, role: roleKey(row[t.columns.get('Relationship')]),
        phone: phoneKey(row[t.columns.get('Source Phone')]), email: emailKey(row[t.columns.get('Source Email')])};
      // The old import can use one Contact Key for several phone/email rows.
      // Keep every row; uniqueness is established when matching a source record.
      if (ck) append(old, ck, entry);
      append(fallback, fk, entry);
    });
    const proposed = [], seenRecords = new Set(), sourceCounts = new Map(), matched = new Set();
    let key = '', studentName = '', gender = '', phone = '';
    data.p.rows.forEach((row, i) => {
      const get = h => row[data.p.columns.get(h)];
      const next = id(get('Student Number'));
      if (next) { key = next; studentName = ''; gender = ''; phone = ''; }
      if (!key) {
        if (row.some(v => !blank(v))) throw new Error('ProRoster row ' + (i + 2) + ' starts without a Student Number.');
        return;
      }
      if (!blank(get('Name'))) studentName = get('Name');
      if (!blank(get('Gender'))) gender = get('Gender');
      if (!blank(get('Phone'))) phone = get('Phone');
      // Master students not in Upstream yet (new enrollments) get contacts too; only
      // students Upstream marks inactive are skipped.
      const upstreamStudent = data.students.get(key);
      if (!masterIds.has(key) || (upstreamStudent && !upstreamStudent.active) || blank(get('Contact Name'))) return;
      const contact = text(get('Contact Name'));
      const contactKey = key + '|' + contact.replace(/[^A-Za-z0-9]+/g, '_').replace(/^_+|_+$/g, '');
      const fk = JSON.stringify([key, nameKey(contact)]);
      const recordKey = JSON.stringify([contactKey, phoneKey(get('Contact Phone')), emailKey(get('Contact Email'))]);
      if (seenRecords.has(recordKey)) {
        report.issues.push('ProRoster row ' + (i + 2) + ' for ' + key + ': repeated contact/channel record skipped. Existing contacts retained.');
        return;
      }
      seenRecords.add(recordKey);
      sourceCounts.set(contactKey, (sourceCounts.get(contactKey) || 0) + 1);
      const coachColumn = data.u.columns.get('LC_COACH_NAME');
      const coach = coachColumn === undefined || !upstreamStudent ? '' : upstreamStudent.row[coachColumn];
      const relationship = (contact.match(/\(([^)]*)\)/) || [])[1] || '';
      const values = {'Student Number': key, 'Student Source Name': studentName, 'Gender': gender,
        'Student Phone': phone, 'Contact Name': contact, 'Relationship': relationship,
        'Source Phone': get('Contact Phone'), 'Source Email': get('Contact Email'),
        'Learning Coach': blank(coach) ? '' : nameKey(contact) === nameKey(coach), 'Contact Key': contactKey};
      proposed.push({key, values, fk, sourceRow: i + 2});
    });
    const writes = [];
    let nextRow = t.rows.length + 2;
    proposed.forEach(item => {
      const {key, values, fk} = item;
      const contactKey = values['Contact Key'];
      const exact = old.get(contactKey) || [];
      const role = roleKey(values.Relationship);
      const candidates = exact.length ? exact : (fallback.get(fk) || []).filter(e => !role || e.role === role);
      const phone = phoneKey(values['Source Phone']), email = emailKey(values['Source Email']);
      const scored = candidates.map(entry => ({entry, score: (phone && phone === entry.phone ? 1 : 0) +
        (email && email === entry.email ? 1 : 0)})).sort((a, b) => b.score - a.score);
      let existing;
      if (scored.length && scored[0].score > 0 && (scored.length === 1 || scored[0].score > scored[1].score)) {
        existing = scored[0].entry;
      } else if (candidates.length === 1 && sourceCounts.get(contactKey) === 1) {
        existing = candidates[0];
      }
      if (candidates.length && !existing || existing && matched.has(existing.number)) {
        report.issues.push('Student ' + key + ', source ProRoster row ' + item.sourceRow + ' (' + values['Contact Name'] +
          '): contact match is uncertain against Contacts rows ' + candidates.map(e => e.number).join(', ') +
          '. Source contact skipped; existing rows/overrides retained. Review phone/email details.');
        return;
      }
      item.existing = existing;
      if (existing) matched.add(existing.number);
      const effective = (sourceHeader, effectiveHeader) => {
        const previousSource = existing ? existing.row[t.columns.get(sourceHeader)] : '';
        const previousEffective = existing ? existing.row[t.columns.get(effectiveHeader)] : '';
        // Preserve manual overrides, including an intentionally cleared effective value.
        if (existing && !same(previousSource, previousEffective)) return previousEffective;
        return blank(values[sourceHeader]) ? previousSource : values[sourceHeader];
      };
      values['Effective Phone'] = effective('Source Phone', 'Effective Phone');
      values['Effective Email'] = effective('Source Email', 'Effective Email');
      if (existing) CONTACT_HEADERS.forEach(h => {
        if (!['Effective Phone', 'Effective Email'].includes(h) && blank(values[h])) values[h] = existing.row[t.columns.get(h)];
      });
      const row = item.existing ? item.existing.number : nextRow++;
      const changes = [];
      const cells = [];
      CONTACT_HEADERS.forEach(h => {
        const before = item.existing ? item.existing.row[t.columns.get(h)] : '';
        let after = item.values[h];
        // Student Number is always written as a number, like Master. The Call tab's
        // FILTER(Contacts!A:A = student number) does not match a number stored as text.
        // Rows saved earlier as text are converted on the next refresh.
        if (h === 'Student Number') after = Number(item.key);
        if (!same(before, after)) {
          cells.push({column: t.columns.get(h) + 1, value: after});
          changes.push({field: h, before, after});
        }
      });
      cells.sort((a, b) => a.column - b.column);
      for (let i = 0; i < cells.length;) {
        const first = cells[i], values = [first.value];
        i++;
        while (i < cells.length && cells[i].column === first.column + values.length) values.push(cells[i++].value);
        writes.push({sheet: 'Contacts', row, column: first.column, values: [values]});
      }
      if (changes.length) report.contacts.push({studentNumber: item.key, row, added: !item.existing, changes});
    });
    t.rows.forEach((row, i) => {
      const key = id(row[t.columns.get('Student Number')]);
      if (key && masterIds.has(key) && !matched.has(i + 2)) {
        report.issues.push('Existing Contacts row ' + (i + 2) + ' for ' + key + ' was not matched in active source contacts; retained.');
      }
    });
    return writes;
  }
  function build(ss, master, data) {
    const report = {changes: [], contacts: [], issues: [], sourceOnly: [], missing: [], inactive: []};
    const rows = master.sheet.getRange(2, 1, master.lastRow - 1, master.headers.length).getValues();
    const formulas = master.sheet.getRange(2, 1, master.lastRow - 1, master.headers.length).getFormulas();
    const fields = [];
    const mirrors = ['SCC', 'ECC', 'STAR'].map(name => {
      const t = table(ss.getSheetByName(name), ['Student Number']);
      const byId = new Map(t.rows.map((row, i) => [id(row[t.columns.get('Student Number')]), i]));
      const formulas = t.rows.length ? t.sheet.getRange(2, 1, t.rows.length, t.width).getFormulas() : [];
      return {name, ...t, byId, formulas};
    });
    MAP.forEach(([src, dest]) => {
      if (!data.u.columns.has(src)) report.issues.push('Upstream is missing ' + src + '; ' + dest + ' retained.');
    });
    const gradeLayout = HR10GRADEIMPORT.findSlots(master.headers, 0);
    if (data.gradesError) report.issues.push('Course grades not refreshed: ' + data.gradesError);
    else if (!gradeLayout) report.issues.push('Master has no course columns (e.g. MTH and MTH%); course grades not copied to Master.');
    if (data.starError) report.issues.push('STAR Import not used; Upstream STAR columns only: ' + data.starError);
    else data.star.issues.forEach(x => report.issues.push(x));
    function add(key, rowNumber, header, value) {
      const c = master.columns.get(header);
      if (!c) throw new Error('Master is missing ' + header);
      if (formulas[rowNumber - 2][c - 1]) throw new Error('Master has a formula in source field ' + header);
      const before = rows[rowNumber - 2][c - 1];
      if (same(before, value)) return;
      // A pending edit in a mirror must be resolved before source refresh overwrites it.
      mirrors.forEach(t => {
        const r = t.byId.get(key), c = t.columns.get(header);
        if (t.formulas[r][c] || !same(t.rows[r][c], before)) {
          throw new Error(t.name + ' differs from Master for student ' + key + ', ' + header + '. Resolve/synchronize this field first.');
        }
      });
      fields.push({studentNumber: key, header, value});
      report.changes.push({studentNumber: key, field: header, before, after: value});
    }
    for (const [key, r] of master.rows) {
      const student = data.students.get(key);
      if (!student) { report.missing.push(key); continue; }
      if (!student.active) { report.inactive.push(key); continue; }
      MAP.forEach(([src, dest]) => {
        if (!data.u.columns.has(src)) return;
        let value = student.row[data.u.columns.get(src)];
        if (dest === 'SPED') {
          // SPED = Yes if the student has an IEP (STUDENT_WITH_DISABILITIES) or a 504 plan (SPED504PLAN).
          const yes = v => ['Y', 'YES', 'TRUE', '1'].includes(text(v).toUpperCase());
          const swdCol = data.u.columns.get('STUDENT_WITH_DISABILITIES');
          const swd = swdCol === undefined ? '' : student.row[swdCol];
          if (yes(swd) || yes(value)) value = 'Yes';
          else if (blank(value) && !blank(swd)) value = 'No';
        }
        if (blank(value)) {
          if (!QUIET_BLANK.includes(dest)) report.issues.push(key + ': blank source ' + src + '; existing value retained.');
          return;
        }
        if (DATE_FIELDS.includes(dest)) value = date(value, key + ' ' + dest);
        add(key, r, dest, value);
      });
      for (const subject of ['Reading', 'Math']) {
        // BOY percentile: only the STAR export has it. Filled once; never overwrites Master.
        const pctHeader = 'BOY ' + subject + ' Percentile';
        const boy = HR10STARIMPORT.boy(data.star, key, subject);
        if (boy && boy.percentile !== null && master.columns.has(pctHeader)) {
          const before = rows[r - 2][master.columns.get(pctHeader) - 1];
          if (blank(before)) add(key, r, pctHeader, boy.percentile);
          else if (!same(before, boy.percentile)) report.issues.push(key + ': ' + pctHeader + ' is ' + before +
            '; STAR Import BOY test says ' + boy.percentile + '. Existing value retained.');
        }
        // Latest date/score: newest test from STAR Import or Upstream (STAR Import wins a tie).
        const candidates = [];
        const latest = HR10STARIMPORT.latest(data.star, key, subject);
        if (latest) candidates.push({date: date(latest.ymd, key + ' STAR Import ' + subject + ' date'), score: latest.score});
        const prefix = 'STAR_2TO12_' + subject.toUpperCase();
        const dh = prefix + '_DATE', sh = prefix + '_UNIFIED_SCORE';
        let upstreamNote = '';
        if (!data.u.columns.has(dh) || !data.u.columns.has(sh)) upstreamNote = 'Upstream ' + subject + ' date/score headers missing';
        else {
          const dv = student.row[data.u.columns.get(dh)], sv = student.row[data.u.columns.get(sh)];
          if (blank(dv) || blank(sv)) upstreamNote = 'incomplete Upstream ' + subject + ' date/score';
          else {
            const score = typeof sv === 'number' ? sv : /^\d+(\.\d+)?$/.test(text(sv)) ? Number(text(sv)) : NaN;
            if (!Number.isFinite(score) || score < 0) throw new Error('Invalid ' + subject + ' score for ' + key);
            candidates.push({date: date(dv, key + ' ' + dh), score});
          }
        }
        if (!candidates.length) {
          report.issues.push(key + ': no ' + subject + ' test in STAR Import' + (upstreamNote ? ' and ' + upstreamNote : '') + '; pair retained.');
          continue;
        }
        const pick = candidates.reduce((a, b) => b.date.getTime() > a.date.getTime() ? b : a);
        const targetDate = 'Latest ' + subject + ' Date', targetScore = 'Latest ' + subject + ' Score';
        const previous = rows[r - 2][master.columns.get(targetDate) - 1];
        if (!blank(previous) && pick.date.getTime() < date(previous, key + ' existing ' + targetDate).getTime()) {
          report.issues.push(key + ': older source ' + subject + ' assessment retained existing pair.'); continue;
        }
        add(key, r, targetDate, pick.date); add(key, r, targetScore, pick.score);
      }
      // Course names and Canvas percentages; students without CourseGrades rows keep theirs.
      const gradeFields = data.grades && HR10GRADEIMPORT.masterFields(gradeLayout, data.grades.byId.get(key));
      if (gradeFields) gradeFields.forEach(f => add(key, r, f.header, f.value));
    }
    data.students.forEach((v, key) => { if (v.active && !master.rows.has(key)) report.sourceOnly.push(key); });
    const relatedWrites = planContacts(ss, master.rows, data, report);
    return {fields, relatedWrites, report};
  }
  function display(report, title) {
    const esc = v => HR10CFG.escapeHtml(v instanceof Date ? Utilities.formatDate(v, book().getSpreadsheetTimeZone(), 'yyyy-MM-dd') : v);
    let html = '<div style="font:14px Arial">';
    const a = report.newStudents;
    if (a) {
      const listOf = xs => esc(xs.join('; ') || 'None');
      html += '<h3>New students</h3><p>' + (a.error ? '<b>Not added:</b> ' + esc(a.error) :
        (report.preview ? 'Would be added to Master, SCC, ECC and STAR: ' + listOf(a.wouldAdd) : 'Added to Master, SCC, ECC and STAR: ' + listOf(a.added))) +
        (a.skipped.length ? '<br>Not added because they are on the Removed Students tab: ' + listOf(a.skipped) +
          ' (use show10AddStudents from the Apps Script editor to add one back).' : '') + '</p>';
      if (report.preview && a.wouldAdd.length) html += '<p><small>Preview: their other fields, contacts and grades are filled when you run the refresh.</small></p>';
    }
    html += '<p>' + report.changes.length + ' Master field changes; ' + report.contacts.length + ' changed/added contact rows.</p>';
    html += '<p>Source-only active students (not added): ' + esc(report.sourceOnly.join(', ') || 'None') + '<br>Master students missing in source (retained): ' + esc(report.missing.join(', ') || 'None') + '<br>Inactive source students (retained): ' + esc(report.inactive.join(', ') || 'None') + '</p>';
    html += '<table border="1" cellpadding="5" style="border-collapse:collapse"><tr><th>Student / location</th><th>Field</th><th>Before</th><th>After</th></tr>';
    report.changes.forEach(x => { html += '<tr><td>' + esc(x.studentNumber) + '</td><td>' + esc(x.field) + '</td><td>' + esc(x.before) + '</td><td>' + esc(x.after) + '</td></tr>'; });
    report.contacts.forEach(x => x.changes.forEach(c => {
      html += '<tr><td>' + esc(x.studentNumber + ' / Contacts row ' + x.row + (x.added ? ' (new)' : '')) + '</td><td>' + esc(c.field) + '</td><td>' + esc(c.before) + '</td><td>' + esc(c.after) + '</td></tr>';
    }));
    html += '</table>';
    const g = report.grades;
    if (g) {
      html += '<h3>Grades tab</h3><p>' + esc(g.studentsUpdated) + ' students with grade changes; ' + esc(g.added.length) +
        ' rows added' + (g.importDate ? '; CourseGrades import date ' + esc(g.importDate) : '') + '.</p>';
      if (g.changes.length) {
        html += '<table border="1" cellpadding="5" style="border-collapse:collapse"><tr><th>Student / row</th><th>Column</th><th>Before</th><th>After</th></tr>';
        g.changes.forEach(x => { html += '<tr><td>' + esc(x.studentNumber + ' / Grades row ' + x.row) + '</td><td>' + esc(x.field) + '</td><td>' + esc(x.before) + '</td><td>' + esc(x.after) + '</td></tr>'; });
        html += '</table>';
      }
    }
    if (report.notYet && report.notYet.length) {
      html += '<h3>On Master, not in Upstream yet</h3><p>Still on Class Roster or starting later, so they are kept and not offered for removal: ' +
        esc(report.notYet.map(s => s.key + ' ' + s.name + (s.start ? ' (starts ' + s.start + ')' : '')).join('; ')) + '</p>';
    }
    html += HR10REMOVE.section(report.missingDetails || []);
    html += '<h3>Missing information and review notes</h3><ul>' + report.issues.map(x => '<li>' + esc(x) + '</li>').join('') + '</ul></div>';
    SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(950).setHeight(650), title);
  }
  function run(preview) {
    const ss = book();
    let report, data, newStudents;
    // Step 1: add new students (own lock; the Master refresh below takes it next).
    try { newStudents = HR10ADD.addAll(ss, preview); }
    catch (e) {
      newStudents = {added: [], wouldAdd: [], skipped: [], notes: [], classRosterIds: new Set(), error: e.message};
      if (!preview) HR10CFG.log('ERROR', 'AddStudents', '', e.message, HR10CFG.errorDetails(e));
    }
    try {
      data = source(ss);
      update10RosterFields(master => {
        const result = build(ss, master, data);
        report = result.report;
        result.relatedWrites.push({sheet: 'Info', row: 3, column: 8,
          values: [[Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd HH:mm:ss z')]]});
        return {fields: result.fields, relatedWrites: result.relatedWrites, dryRun: preview, bulk: true};
      });
    } catch (e) {
      if (!preview) HR10CFG.log('ERROR', 'RefreshStudentData', '', e.message, HR10CFG.errorDetails(e));
      throw e;
    }
    report.preview = preview;
    report.newStudents = newStudents;
    newStudents.notes.forEach(x => report.issues.push('New students: ' + x));
    if (!preview) HR10CFG.log('SUCCESS', 'RefreshStudentData', '',
      report.changes.length + ' Master fields; ' + report.contacts.length + ' contact rows refreshed.', '');
    try {
      if (!data.grades) throw new Error('see the Course grades note above');
      report.grades = HR10GRADEIMPORT.run(ss, data.grades, preview);
      report.grades.issues.forEach(x => report.issues.push('Grades: ' + x));
      if (!preview) HR10CFG.log('SUCCESS', 'GradesImport', '', report.grades.changes.length + ' Grades cells changed.', '');
    } catch (e) {
      report.issues.push('Grades import stopped' + (preview ? '' : '; check the Grades tab') + ': ' + e.message);
      if (!preview) HR10CFG.log('ERROR', 'GradesImport', '', e.message, HR10CFG.errorDetails(e));
    }
    if (!preview && typeof HR10DUECOLORS !== 'undefined') {
      try { HR10DUECOLORS.apply(true); }
      catch (e) { report.issues.push('Call Due colors not reapplied: ' + e.message); }
    }
    try { report.missingDetails = HR10REMOVE.details(ss, report.missing, data.book); }
    catch (e) { report.issues.push('AddDrop details unavailable: ' + e.message); report.missingDetails = []; }
    // Not in Upstream yet (still on Class Roster, or a future start): keep, no Remove checkbox.
    const today = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd');
    const waiting = s => newStudents.classRosterIds.has(s.key) || (/^\d{4}-\d{2}-\d{2}$/.test(s.start) && s.start > today);
    report.notYet = report.missingDetails.filter(waiting);
    report.missingDetails = report.missingDetails.filter(s => !waiting(s));
    console.log(JSON.stringify(report));
    try { display(report, preview ? 'Student refresh preview — no data saved' : 'Student refresh saved'); }
    catch (e) { console.error('Refresh report could not be displayed: ' + e.message); }
    return report;
  }
  return Object.freeze({run});
})();
function preview10StudentRefresh() { return HR10REFRESH.run(true); }
function refreshStudentData() { return HR10REFRESH.run(false); }
function refreshStudentDataAndCounselors() { return refreshStudentData(); }
