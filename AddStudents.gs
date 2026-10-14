/** 10RosterHR: add new students to Master, SCC, ECC and STAR (v1.1). Requires Config.gs.
 * @NotOnlyCurrentDoc
 * v1.1: runs as the first step of Pull Data > Refresh students (RefreshStudentData v1.7), which
 * adds every new student automatically. Students listed on the Removed Students
 * tab are never re-added automatically; the refresh report names them.
 * show10AddStudents() still opens a pick-and-confirm dialog if you want it.
 *
 * Finds students who are in the imports workbook (Info!H2) but not on Master:
 *  - every row of the Class Roster tab, and
 *  - every active Upstream student (CURRENT_ACTIVE_STUDENT = Y).
 * A dialog lists them with a checkbox each; nothing changes until you confirm.
 *
 * Each added student gets one new row at the bottom of all four tabs (each tab
 * keeps its own column order). Filled in:
 *  - every Class Roster column whose header matches a roster header
 *    (LAST NAME, FIRST NAME, GRADE, START DATE, SCHOOL, Welcome Email, ISMA, ...),
 *  - Upstream values when the student is already there (names, grade, start date,
 *    school, SPED, Student Email, Counselor), which win over Class Roster,
 *  - Preferred Name = "First Last", Name and Subject Line formulas.
 * A Class Roster grade of 11 is written as "11th Grade", Upstream's format.
 * The new row copies formatting and dropdowns from the row above it.
 * Everything else stays blank. Run Refresh existing students afterwards to fill
 * contacts and course grades (students not yet in Upstream get those once a
 * newer Upstream export includes them).
 * If any tab fails, rows already added by this run are deleted again.
 */
const HR10ADD = (() => {
  const ID = 'Student Number';
  const DERIVED = ['Name', 'Subject Line'];
  // Upstream header -> roster header. Same fields the refresh keeps current.
  const UPSTREAM = [['STUDENT_LAST_NAME', 'LAST NAME'], ['STUDENT_FIRST_NAME', 'FIRST NAME'], ['GRADE', 'GRADE'],
    ['ENROLL_DATE_PS', 'START DATE'], ['SCHOOL', 'SCHOOL'], ['EMAIL_O365', 'Student Email'], ['COUNSELOR', 'Counselor'],
    ['GE_NEXT_CALL_DUE', 'Call Due']];
  const text = v => HR10CFG.text(v);
  const blank = v => HR10CFG.blank(v);
  const yes = v => ['Y', 'YES', 'TRUE', '1'].includes(text(v).toUpperCase());
  const safeId = v => {
    try { return HR10CFG.studentId(typeof v === 'number' ? Math.round(v) : v); } catch (e) { return null; }
  };

  /** 10 -> "10th Grade" (Upstream's wording). Text is kept as it is. */
  function gradeText(v) {
    const n = typeof v === 'number' ? v : /^\d{1,2}$/.test(text(v)) ? Number(text(v)) : NaN;
    if (!Number.isInteger(n)) return v;
    const suffix = n % 100 >= 11 && n % 100 <= 13 ? 'th' : ({1: 'st', 2: 'nd', 3: 'rd'}[n % 10] || 'th');
    return n + suffix + ' Grade';
  }

  /** Students in Class Roster or active in Upstream who are not on Master. */
  function candidates(ss) {
    const master = HR10CFG.table(ss.getSheetByName('Master'), [ID]);
    const masterHeaders = new Set(master.headers.filter(Boolean));
    const onMaster = new Set(master.rows.map(r => safeId(r[master.columns.get(ID)])).filter(Boolean));
    const source = HR10CFG.sourceWorkbook(ss);
    const found = new Map(), notes = [], classRosterIds = new Set();
    // Students removed earlier with the refresh report's Remove button.
    const removedSheet = ss.getSheetByName('Removed Students');
    const removed = new Set();
    if (removedSheet) {
      const rs = HR10CFG.table(removedSheet, []);
      if (rs.columns.has(ID)) rs.rows.forEach(r => { const k = safeId(text(r[rs.columns.get(ID)]).replace(/^'/, '')); if (k) removed.add(k); });
    }
    const entry = key => {
      if (!found.has(key)) found.set(key, {key, values: {}, sources: []});
      return found.get(key);
    };

    const crSheet = source.getSheetByName('Class Roster');
    if (!crSheet) notes.push('The imports workbook has no Class Roster tab; only Upstream was checked.');
    else {
      const cr = HR10CFG.table(crSheet, [ID, 'LAST NAME', 'FIRST NAME']);
      cr.rows.forEach((row, i) => {
        const raw = row[cr.columns.get(ID)];
        if (blank(raw)) return;
        const key = safeId(raw);
        if (!key) { notes.push('Class Roster row ' + (i + 2) + ': Student Number "' + text(raw) + '" is not a number; skipped.'); return; }
        classRosterIds.add(key);
        if (onMaster.has(key)) return;
        const e = entry(key);
        if (!e.sources.includes('Class Roster')) e.sources.push('Class Roster');
        cr.headers.forEach((h, c) => {
          if (h && h !== ID && !DERIVED.includes(h) && masterHeaders.has(h) && !blank(row[c])) e.values[h] = row[c];
        });
      });
    }

    const upSheet = source.getSheetByName('Upstream');
    if (!upSheet) notes.push('The imports workbook has no Upstream tab; only Class Roster was checked.');
    else {
      const u = HR10CFG.table(upSheet, ['STUDENT_NUMBER', 'CURRENT_ACTIVE_STUDENT']);
      const get = (row, h) => (u.columns.has(h) ? row[u.columns.get(h)] : '');
      u.rows.forEach(row => {
        const key = safeId(get(row, 'STUDENT_NUMBER'));
        if (!key || onMaster.has(key) || !yes(get(row, 'CURRENT_ACTIVE_STUDENT'))) return;
        const e = entry(key);
        e.sources.push('Upstream');
        UPSTREAM.forEach(([src, dest]) => {
          const v = get(row, src);
          if (masterHeaders.has(dest) && !blank(v)) e.values[dest] = v;
        });
        if (masterHeaders.has('SPED')) {
          const swd = get(row, 'STUDENT_WITH_DISABILITIES'), plan = get(row, 'SPED504PLAN');
          if (yes(swd) || yes(plan)) e.values.SPED = 'Yes';
          else if (!blank(swd) || !blank(plan)) e.values.SPED = 'No';
        }
      });
    }

    const list = [...found.values()].map(e => {
      if (!blank(e.values.GRADE)) e.values.GRADE = gradeText(e.values.GRADE);
      const first = text(e.values['FIRST NAME']), last = text(e.values['LAST NAME']);
      if (masterHeaders.has('Preferred Name') && blank(e.values['Preferred Name']) && (first || last)) {
        e.values['Preferred Name'] = (first + ' ' + last).trim();
      }
      return {key: e.key, first, last, values: e.values, sources: e.sources, removed: removed.has(e.key),
        grade: text(e.values.GRADE), start: e.values['START DATE'], school: text(e.values.SCHOOL)};
    });
    list.sort((a, b) => a.last.localeCompare(b.last) || a.first.localeCompare(b.first));
    return {list, notes, classRosterIds};
  }

  function formula(columns, row, header) {
    const cell = h => HR10CFG.columnLetter(columns.get(h)) + row;
    return header === 'Name'
      ? '=CONCATENATE(' + cell('FIRST NAME') + '," ",' + cell('LAST NAME') + ')'
      : '=LEFT(' + cell('FIRST NAME') + ',1)&" "&' + cell('LAST NAME') + '&", "&' + cell(ID) + '&", "&' + cell('SCHOOL');
  }

  /** All four tabs: headers present, same students. Returns per-tab info. */
  function tabs(ss) {
    const result = HR10CFG.ROSTER_TABS.map(name => {
      const sheet = ss.getSheetByName(name);
      if (!sheet) throw new Error(name + ' tab is missing.');
      const info = HR10CFG.headerColumns(sheet);
      [ID, 'FIRST NAME', 'LAST NAME', 'SCHOOL', ...DERIVED].forEach(h => {
        if (!info.columns.has(h)) throw new Error(name + ': missing header ' + h + '.');
      });
      const last = sheet.getLastRow();
      const ids = last > 1 ? sheet.getRange(2, info.columns.get(ID), last - 1, 1).getValues().map(r => safeId(r[0])).filter(Boolean) : [];
      return {name, sheet, ...info, last, ids: new Set(ids)};
    });
    const master = result[0];
    result.slice(1).forEach(t => {
      if (t.ids.size !== master.ids.size || [...master.ids].some(k => !t.ids.has(k))) {
        throw new Error(t.name + ': student list differs from Master. Match the four tabs before adding students.');
      }
    });
    return result;
  }

  function add(ids) {
    const ss = HR10CFG.workbook();
    const wanted = [...new Set((ids || []).map(v => HR10CFG.studentId(v)).filter(Boolean))];
    if (!wanted.length) throw new Error('No students were selected.');
    return HR10CFG.withLock(() => {
      const {list} = candidates(ss);
      const byKey = new Map(list.map(c => [c.key, c]));
      const gone = wanted.filter(k => !byKey.has(k));
      if (gone.length) throw new Error('Already on Master or no longer in the imports workbook: ' + gone.join(', ') + '. Nothing was added.');
      const students = wanted.map(k => byKey.get(k));
      const all = tabs(ss);
      const done = [];
      try {
        all.forEach(t => {
          const start = Math.max(t.last, 1) + 1, height = students.length;
          if (start + height - 1 > t.sheet.getMaxRows()) t.sheet.insertRowsAfter(t.sheet.getMaxRows(), start + height - 1 - t.sheet.getMaxRows());
          const range = t.sheet.getRange(start, 1, height, t.width);
          done.push({t, start, height});
          if (t.last >= 2) {
            const above = t.sheet.getRange(t.last, 1, 1, t.width);
            for (let r = 0; r < height; r++) {
              const target = t.sheet.getRange(start + r, 1, 1, t.width);
              above.copyTo(target, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
              above.copyTo(target, SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false);
            }
          }
          const values = [], mask = [];
          students.forEach((s, r) => {
            const row = Array(t.width).fill(''), isFormula = Array(t.width).fill(false);
            t.headers.forEach((h, c) => {
              if (h === ID) row[c] = Number(s.key);
              else if (DERIVED.includes(h)) { row[c] = formula(t.columns, start + r, h); isFormula[c] = true; }
              else if (h && Object.prototype.hasOwnProperty.call(s.values, h)) row[c] = s.values[h];
            });
            values.push(row); mask.push(isFormula);
          });
          range.setValues(HR10CFG.prepareValues(range, values, mask));
        });
        SpreadsheetApp.flush();
      } catch (error) {
        const failed = [];
        done.reverse().forEach(d => {
          try { d.t.sheet.deleteRows(d.start, d.height); } catch (e) { failed.push(d.t.name); }
        });
        throw new Error('Students not added: ' + error.message +
          (failed.length ? '. Check the bottom of ' + failed.join(', ') + ' for partly added rows.' : '. No rows were left behind.'));
      }
      const names = students.map(s => s.key + ' ' + (s.first + ' ' + s.last).trim());
      HR10CFG.log('SUCCESS', 'AddStudents', wanted.join(','), 'Added to Master, SCC, ECC and STAR', names.join('; '));
      return {added: names};
    }, 'Roster tools are busy. Please retry adding students.');
  }

  function dialog() {
    const ss = HR10CFG.workbook();
    const {list, notes} = candidates(ss);
    const esc = HR10CFG.escapeHtml;
    const day = v => v instanceof Date ? Utilities.formatDate(v, ss.getSpreadsheetTimeZone(), 'M/d/yyyy') : text(v);
    let html = '<div style="font:14px Arial">';
    if (!list.length) {
      html += '<p>Every student in Class Roster and every active Upstream student is already on Master.</p>';
    } else {
      html += '<p>These students are in the imports workbook but not on Master. Checked students are added to the bottom of ' +
        '<b>Master, SCC, ECC and STAR</b>.</p><table border="1" cellpadding="5" style="border-collapse:collapse">' +
        '<tr><th>Add</th><th>Student</th><th>Grade / start / school</th><th>Found in</th><th>Other fields</th></tr>';
      list.forEach(s => {
        const other = Object.keys(s.values).filter(h => !['LAST NAME', 'FIRST NAME', 'Preferred Name', 'GRADE', 'START DATE', 'SCHOOL'].includes(h))
          .map(h => esc(h) + ': ' + esc(day(s.values[h]))).join('<br>');
        html += '<tr><td style="text-align:center"><input type="checkbox" class="add" value="' + esc(s.key) + '"' + (s.removed ? '' : ' checked') + '></td>' +
          '<td>' + esc(s.key) + '<br>' + esc((s.first + ' ' + s.last).trim()) + (s.removed ? '<br><small>Removed earlier (Removed Students tab)</small>' : '') + '</td>' +
          '<td>' + esc(s.grade) + '<br>' + esc(day(s.start)) + '<br>' + esc(s.school) + '</td>' +
          '<td>' + esc(s.sources.join(', ')) + (s.sources.includes('Upstream') ? '' : '<br><small>Not in Upstream yet: email, counselor, contacts and grades come later.</small>') + '</td>' +
          '<td><small>' + (other || '') + '</small></td></tr>';
      });
      html += '</table><p><button id="addBtn" onclick="addClick()">Add checked students</button> <span id="addMsg"></span></p>' +
        '<script>let armed=false;function msg(t){document.getElementById("addMsg").textContent=t;}' +
        'function addClick(){const ids=[...document.querySelectorAll("input.add:checked:not(:disabled)")].map(x=>x.value);' +
        'const b=document.getElementById("addBtn");if(!ids.length){msg("Check at least one student.");return;}' +
        'if(!armed){armed=true;b.textContent="Click again to add "+ids.length+" student(s)";return;}' +
        'b.disabled=true;msg("Adding…");google.script.run.withSuccessHandler(r=>{msg("Added: "+r.added.join("; ")+' +
        '". Next: Refresh existing students, contacts and grades.");' +
        'document.querySelectorAll("input.add:checked").forEach(x=>{x.disabled=true;});b.textContent="Add checked students";b.disabled=false;armed=false;})' +
        '.withFailureHandler(e=>{msg(e.message);b.textContent="Add checked students";b.disabled=false;armed=false;})' +
        '.add10NewStudents(ids);}</script>';
    }
    if (notes.length) html += '<h3>Notes</h3><ul>' + notes.map(n => '<li>' + esc(n) + '</li>').join('') + '</ul>';
    html += '</div>';
    SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(900).setHeight(550), 'Add new students');
  }

  const label = s => s.key + ' ' + (s.first + ' ' + s.last).trim();

  /** Refresh step 1. preview: list only. Otherwise adds every candidate
   * except students on the Removed Students tab. */
  function addAll(ss, preview) {
    const c = candidates(ss);
    const toAdd = c.list.filter(s => !s.removed);
    const skipped = c.list.filter(s => s.removed).map(label);
    const added = preview || !toAdd.length ? [] : add(toAdd.map(s => s.key)).added;
    return {added, wouldAdd: preview ? toAdd.map(label) : [], skipped, notes: c.notes, classRosterIds: c.classRosterIds};
  }

  return Object.freeze({candidates, add, addAll, dialog, gradeText});
})();

function show10AddStudents() { return HR10ADD.dialog(); }
function add10NewStudents(ids) { return HR10ADD.add(ids); }
