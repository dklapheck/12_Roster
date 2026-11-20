/** 5+ Attempt Log v1.2 for 12RosterORN. Requires Config.gs, MasterSave.gs and LogECC.gs.
 *
 * A streak is every [Attempt] entry in ECC Contact History since the last
 * successful contact ([Conversation], [Parent Contact], or an older
 * "-- 9/14/2026: ..." entry). History is newest entry first.
 *
 * Once a streak reaches 5 attempts, the student is due ONE 5+ Attempt Log.
 * Archive and Log All ECC offers it (checked) and Update > ECC & Demographics >
 * Log selected student's 5+ attempt log logs one student. The entry is
 * written by the script from the history, dated today:
 *     -- 11/20/2026 [5+ Attempt Log]
 * It never counts as an attempt and is never offered again in the same streak.
 * Only a successful contact starts a new streak.
 *
 * v1.2: each attempt says who it went to (student, LC, LG, EC), read from the
 * note. LogECC also starts each attempt's PowerSchool note with "To: ...".
 * v1.1: colors are conditional formatting rules, placed first so they show
 * over other rules such as an every-other-row grey. They read the Attempt
 * Streak column on Master, which this script fills in (and adds once, at the
 * right end, if it is missing). Values move with the row when you sort:
 *   "Due: 6 attempts"    - 5+ attempts, not logged yet -> red
 *   "Logged 11/20/2026"  - the streak has a 5+ Attempt Log -> yellow
 *   "3 attempts"         - fewer than 5 -> no color
 *   blank                - no attempts since the last successful contact
 * The same rules color ECC Contact History and Attempt Streak. Refresh after
 * editing a history cell by hand (Update > Maintenance > Refresh 5+ attempt colors).
 */
const ATTEMPT_STREAK = (() => {
  const LIMIT = 5, TYPE = '5+ Attempt Log';
  const RED = '#EA9999', YELLOW = '#FFE599';
  const HEADER = /^-- (\d{1,2})\/(\d{1,2})\/(\d{2,4})\b(.*)$/;
  const clean = value => String(value == null ? '' : value).replace(/\r\n?/g, '\n').trim();

  /** History text -> [{date, type, note}], newest first (the order written). */
  function entries(history) {
    const list = [];
    clean(history).split('\n').forEach(line => {
      const match = line.match(HEADER);
      if (match) {
        const tag = match[4].match(/^\s*\[([^\]]+)\]\s*(.*)$/);
        const rest = tag ? tag[2] : match[4].replace(/^\s*:?\s*/, '');
        list.push({date: match[1] + '/' + match[2] + '/' + match[3],
          type: tag ? tag[1].trim() : '', lines: rest.trim() ? [rest.trim()] : []});
      } else if (list.length) list[list.length - 1].lines.push(line);
    });
    return list.map(e => ({date: e.date, type: e.type, note: e.lines.join('\n').trim()}));
  }

  /** {count, attempts (oldest first), logged, logDate, lastSuccess, due, color} */
  function status(history) {
    const attempts = [];
    let logged = false, logDate = '', lastSuccess = '';
    for (const entry of entries(history)) {
      if (entry.type === TYPE) { if (!logged) { logged = true; logDate = entry.date; } continue; }
      if (entry.type !== 'Attempt') { lastSuccess = entry.date; break; }
      attempts.push(entry);
    }
    attempts.reverse();
    const count = attempts.length, due = count >= LIMIT && !logged;
    return {count, attempts, logged, logDate, lastSuccess, due,
      color: logged ? 'yellow' : due ? 'red' : ''};
  }

  // ---- summary ------------------------------------------------------------
  // One label per attempt, from the note's first line, first match wins, so a
  // "ParentSquare: missed check-in" message is a ParentSquare message.
  const METHODS = [
    ['ParentSquare message', /parent\s*square/i],
    ['email', /\be-?mail|^(hello|hi)\b|^missed check-?in.*reschedule/i],
    ['text message', /\btext(ed)?\b|\bsms\b/i],
    ['phone call/voicemail', /voice\s*mail|\bvm\b|\bcall(ed)?\b|\bphone\b/i],
    ['absent from live class', /unresponsive|did not respond|no.?show|left without|did not (come|join|attend)|didn.t (come|join|attend)|missed (homeroom|owl|special|two|check)/i]
  ];
  const SHORT = 50;
  function methodOf(note) {
    const first = clean(note).split('\n')[0].replace(/\s+/g, ' ').trim();
    const hit = METHODS.find(m => m[1].test(first.slice(0, 120)));
    if (hit) return hit[0];
    if (!first) return 'attempt';
    return first.length > SHORT ? first.slice(0, SHORT).replace(/\s+\S*$/, '') + '…' : first;
  }
  const shortDate = date => date.replace(/\/\d{2,4}$/, '');

  // ---- who an attempt went to -------------------------------------------
  // Codes anywhere in the note: LC, LG, EC (capitals, whole word). Words only
  // on label lines ("Email to LC ...:", "Voicemail for ...:", the first line)
  // and greeting lines ("Hi Tameika,"), so body text such as "check with your
  // student" or "ask your learning coach" does not count.
  const WHO = ['student', 'LC', 'LG', 'EC'];
  const firstName = value => String(value == null ? '' : value).trim().split(/[\s,]+/)[0] || '';
  const escape = value => value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  /** names: {lc: Learning Coach, student: Preferred Name}, both optional. */
  function recipients(note, names) {
    const lines = clean(note).split('\n').map(line => line.trim()).filter(Boolean);
    const found = new Set();
    const text = lines.join('\n');
    if (/\bLC\b/.test(text)) found.add('LC');
    if (/\bLG\b/.test(text)) found.add('LG');
    if (/\bEC\b/.test(text)) found.add('EC');
    const lc = firstName(names && names.lc), student = firstName(names && names.student);
    lines.forEach((line, i) => {
      const greeting = line.match(/^(?:hi|hello|dear|good (?:morning|afternoon|evening))\b[\s,]*([^,!.:]*)/i);
      if (greeting) {
        const who = greeting[1].trim();
        if (/learning coach/i.test(who) || (lc && new RegExp('^' + escape(lc) + '\\b', 'i').test(who))) found.add('LC');
        else if (/guardian/i.test(who)) found.add('LG');
        else if (/scholar|student/i.test(who) || (student && new RegExp('^' + escape(student) + '\\b', 'i').test(who))) found.add('student');
        return;
      }
      if (i > 0 && !/:\s*$/.test(line)) return;   // label lines end with ":" (or are the first line)
      const label = line.slice(0, 160);
      if (/learning coach/i.test(label)) found.add('LC');
      if (/legal guardian|\bguardian\b/i.test(label)) found.add('LG');
      if (/emergency contact/i.test(label)) found.add('EC');
      if (/\bto (?:the )?student\b|\bstudent\b.*\b(?:email|text|call|voicemail)|\bscholar\b/i.test(label)) found.add('student');
    });
    return WHO.filter(code => found.has(code));
  }
  const listWho = codes => codes.length < 3 ? codes.join(' and ') :
    codes.slice(0, -1).join(', ') + ' and ' + codes[codes.length - 1];

  /** The entry's note lines (no header). names: see recipients(). */
  function summary(state, names) {
    const since = state.lastSuccess ? ' since the last successful contact on ' + state.lastSuccess : '';
    const everyone = new Set();
    const items = state.attempts.map(a => {
      const who = recipients(a.note, names);
      who.forEach(code => everyone.add(code));
      return shortDate(a.date) + ' ' + methodOf(a.note) + (who.length ? ' to ' + listWho(who) : '');
    });
    const lines = [
      state.count + ' unsuccessful student/parent contact attempts in a row' + since + ':',
      items.join('; ') + '.'
    ];
    if (everyone.size) lines.push('Contacted: ' + listWho(WHO.filter(code => everyone.has(code))) + '.');
    lines.push('No response received.');
    return lines.join('\n');
  }

  /** Why a student is not due, or '' when due. */
  function notDueReason(state) {
    if (state.logged) return 'this streak already has a 5+ Attempt Log (' + state.logDate +
      '). A successful contact starts a new streak.';
    if (state.count < LIMIT) return 'only ' + state.count + ' attempt' + (state.count === 1 ? '' : 's') +
      ' since the last successful contact (needs ' + LIMIT + ').';
    return '';
  }

  // ---- colors -------------------------------------------------------------
  const COLUMN = 'Attempt Streak';
  const CF_MARK = 'N("HR10STREAK")=0';   // always TRUE; identifies this script's rules

  /** The Attempt Streak text for one history cell. */
  function label(state) {
    if (state.logged) return 'Logged ' + state.logDate;
    if (state.due) return 'Due: ' + state.count + ' attempts';
    return state.count ? state.count + ' attempt' + (state.count === 1 ? '' : 's') : '';
  }

  /** Adds the Attempt Streak header after the last column if Master lacks it. */
  function ensureColumn(sheet) {
    const found = CONFIG.headerColumns(sheet).columns.get(COLUMN);
    if (found) return {column: found, added: false};
    const last = sheet.getLastColumn();
    if (sheet.getMaxColumns() <= last) sheet.insertColumnAfter(last);
    sheet.getRange(1, last + 1).setValue(COLUMN);
    CONFIG.log('INFO', 'Setup', '', 'Added the Master column ' + COLUMN + '.');
    return {column: last + 1, added: true};
  }

  function ours(rule) {
    const condition = rule.getBooleanCondition && rule.getBooleanCondition();
    if (!condition || condition.getCriteriaType() !== SpreadsheetApp.BooleanCriteria.CUSTOM_FORMULA) return false;
    return String(condition.getCriteriaValues()[0] || '').replace(/\s+/g, '').indexOf('N("HR10STREAK")') >= 0;
  }
  /** Rule formulas, relative to row 2 of the Attempt Streak column. */
  function formulas(letter) {
    const cell = '$' + letter + '2';
    return [
      {formula: '=AND(' + CF_MARK + ',LEFT(' + cell + ',4)="Due:")', background: RED},
      {formula: '=AND(' + CF_MARK + ',LEFT(' + cell + ',7)="Logged ")', background: YELLOW}
    ];
  }
  /** Puts this script's rules first; every other rule is kept in its order. */
  function applyRules(sheet, historyColumn, statusColumn) {
    const height = Math.max(sheet.getMaxRows() - 1, 1);
    const ranges = [sheet.getRange(2, historyColumn, height, 1), sheet.getRange(2, statusColumn, height, 1)];
    const all = sheet.getConditionalFormatRules();
    const kept = all.filter(rule => !ours(rule));
    const fresh = formulas(CONFIG.columnLetter(statusColumn)).map(x => SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(x.formula).setBackground(x.background).setRanges(ranges).build());
    sheet.setConditionalFormatRules(fresh.concat(kept));
  }

  /** v1.0 filled ECC Contact History directly; clear only those two fills. */
  function clearOldFills(range) {
    const old = [RED, YELLOW].map(c => c.toLowerCase());
    const fills = range.getBackgrounds();
    if (!fills.some(row => old.includes(String(row[0]).toLowerCase()))) return;
    range.setBackgrounds(fills.map(row => [old.includes(String(row[0]).toLowerCase()) ? null : row[0]]));
  }

  /** Fills Attempt Streak for every student and reapplies the color rules. */
  function recolor(quiet) {
    const ss = CONFIG.workbook();
    const result = CONFIG.withLock(() => {
      const sheet = ss.getSheetByName(CONFIG.MASTER);
      if (!sheet) throw new Error('The Master tab is missing.');
      const added = ensureColumn(sheet).added;
      const master = MASTER_SAVE.readMaster(ss);
      const history = master.columns.get(ECC_LOG.FIELDS.history), statusColumn = master.columns.get(COLUMN);
      if (!history) throw new Error('Missing LogECC header: ' + ECC_LOG.FIELDS.history);
      let red = 0, yellow = 0;
      const height = master.lastRow - 1;
      if (height >= 1) {
        const target = sheet.getRange(2, statusColumn, height, 1);
        if (target.getFormulas().some(row => row[0])) {
          throw new Error('Attempt Streak on Master must not contain formulas; the script fills it in.');
        }
        const current = target.getValues();
        const values = sheet.getRange(2, history, height, 1).getValues().map((row, i) => {
          if (!master.rowKeys.has(i + 2)) return [current[i][0]];
          const state = status(row[0]);
          if (state.color === 'red') red++;
          if (state.color === 'yellow') yellow++;
          return [label(state)];
        });
        if (JSON.stringify(values) !== JSON.stringify(current)) target.setValues(values);
        clearOldFills(sheet.getRange(2, history, height, 1));
      }
      applyRules(sheet, history, statusColumn);
      return {red, yellow, added};
    }, 'Roster tools are busy. Please retry.');
    if (!quiet) ss.toast(result.red + ' red (5+ attempts, not logged) · ' + result.yellow +
      ' yellow (logged this streak).' + (result.added ? ' Added the Attempt Streak column on Master.' : ''),
      '5+ attempt colors', 8);
    return result;
  }
  /** Used after archives: a color problem never undoes a saved archive. */
  function recolorQuietly() {
    try { return recolor(true); }
    catch (error) { CONFIG.log('ERROR', '5+ Attempt Colors', '', error.message, CONFIG.errorDetails(error)); return null; }
  }

  return Object.freeze({entries, status, summary, recipients, listWho, notDueReason, methodOf, label, formulas, ours,
    recolor, recolorQuietly, LIMIT, TYPE, RED, YELLOW, COLUMN});
})();

function refreshFiveAttemptColors() { return ATTEMPT_STREAK.recolor(false); }
