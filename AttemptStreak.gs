/** 5+ Attempt Log v1.0 for 12RosterORN. Requires Config.gs, MasterSave.gs and LogECC.gs.
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
 * ECC Contact History cell colors (this script owns that column's fill):
 *   red    - 5+ attempts in the streak, not logged yet
 *   yellow - the streak has a 5+ Attempt Log (more attempts keep it yellow)
 *   none   - otherwise
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

  /** The entry's note lines (no header). */
  function summary(state) {
    const since = state.lastSuccess ? ' since the last successful contact on ' + state.lastSuccess : '';
    return [
      state.count + ' unsuccessful student/parent contact attempts in a row' + since + ':',
      state.attempts.map(a => shortDate(a.date) + ' ' + methodOf(a.note)).join('; ') + '.',
      'No response received.'
    ].join('\n');
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
  /** Recolors every ECC Contact History cell on Master. Writes fills only. */
  function recolor(quiet) {
    const ss = CONFIG.workbook();
    const result = CONFIG.withLock(() => {
      const master = MASTER_SAVE.readMaster(ss);
      const column = master.columns.get(ECC_LOG.FIELDS.history);
      if (!column) throw new Error('Missing LogECC header: ' + ECC_LOG.FIELDS.history);
      const height = master.lastRow - 1;
      if (height < 1) return {red: 0, yellow: 0};
      const range = master.sheet.getRange(2, column, height, 1);
      let red = 0, yellow = 0;
      const fills = range.getValues().map((row, i) => {
        if (!master.rowKeys.has(i + 2)) return [null];
        const color = status(row[0]).color;
        if (color === 'red') { red++; return [RED]; }
        if (color === 'yellow') { yellow++; return [YELLOW]; }
        return [null];
      });
      range.setBackgrounds(fills);
      return {red, yellow};
    }, 'Roster tools are busy. Please retry.');
    if (!quiet) ss.toast(result.red + ' red (5+ attempts, not logged) · ' + result.yellow +
      ' yellow (logged this streak).', '5+ attempt colors', 8);
    return result;
  }
  /** Used after archives: a color problem never undoes a saved archive. */
  function recolorQuietly() {
    try { return recolor(true); }
    catch (error) { CONFIG.log('ERROR', '5+ Attempt Colors', '', error.message, CONFIG.errorDetails(error)); return null; }
  }

  return Object.freeze({entries, status, summary, notDueReason, methodOf, recolor, recolorQuietly,
    LIMIT, TYPE, RED, YELLOW});
})();

function refreshFiveAttemptColors() { return ATTEMPT_STREAK.recolor(false); }
