/** LogECC v4.2 for 12RosterORN. Requires Config.gs, MasterSave.gs, EccPowerSchool.gs.
 * Menus are built by Menus.gs.
 * v4.2: 5+ Attempt Log entries (needs AttemptStreak.gs). The script writes the
 * note from ECC Contact History and dates it today; no Master cell is used.
 * Archiving recolors ECC Contact History.
 * v4.1: Parent Contact entries (successful parent contact from Parent
 * Communication) and the ECC Attempt Date column renamed to Contact Date,
 * which dates both attempts and parent contacts. installRosterTools renames
 * the old header once.
 * v3.0: Master only. Select the student's row on Master. The one-time Old
 * Attempts merge was removed (that column is gone).
 *
 * ECC Contact History is one timeline, newest entry first. Conversations and
 * attempts both go there, each starting with a header line:
 *     -- 10/5/2026 [Conversation]        -- 10/2/2026 [Attempt]
 * Parent contacts use the same timeline:   -- 10/6/2026 [Parent Contact]
 * So do 5+ attempt logs:                   -- 11/20/2026 [5+ Attempt Log]
 * A new entry is placed by its own date (ECC Date or Contact Date), so an
 * older attempt archived late still lands in the right spot.
 *
 * PowerSchool gets its own one-line version of the note (see EccPowerSchool).
 * Current notes and New Attempt Student Contact stay visible after archiving.
 * All saves go through saveToMaster(planner).
 */
const ECC_LOG = (() => {
  const FIELDS = {
    date: 'ECC Date', history: 'ECC Contact History',
    notes: [
      // [header, label in history, label in PowerSchool (defaults to history label)]
      // An empty label writes the note with no "Label: " prefix.
      ['ECC Overall', ''],
      ['ECC Classes', 'Classes'],
      ['ECC Grades', 'Grades'],
      ['ECC Social', 'Socially'],
      ['ECC Follow-up', 'Actions / Follow up', 'Follow-up']
    ],
    contactDate: 'Contact Date', attemptNote: 'New Attempt Student Contact',
    parentNote: 'Parent Communication'
  };
  const OLD_CONTACT_DATE = 'ECC Attempt Date';
  // Entry type -> PowerSchoolSettings workflow row and extension handoff type.
  const TYPES = Object.freeze({
    'Conversation': {workflow: 'ECC Conversation', handoff: 'ecc'},
    'Attempt': {workflow: 'ECC Attempt', handoff: 'ecc'},
    'Parent Contact': {workflow: 'Parent Contact', handoff: 'parent'},
    // The 'parent' handoff accepts configured or manually chosen Type/Subtype
    // and keeps any PowerSchool template text, so it suits this log too.
    '5+ Attempt Log': {workflow: '5+ Attempts', handoff: 'parent'}
  });
  const FIVE = '5+ Attempt Log';
  function typeInfo(type) {
    if (!Object.prototype.hasOwnProperty.call(TYPES, type)) throw new Error('Unknown ECC entry type.');
    return TYPES[type];
  }
  const CELL_LIMIT = 50000, CELL_WARNING = 45000;
  // "-- 9/14/2026 [Conversation]" and the older "-- 9/14/2026: Good..." both start entries.
  const HEADER_LINE = /^-- (\d{1,2})\/(\d{1,2})\/(\d{2,4})\b/;
  const text = value => CONFIG.text(value);

  // ---- dates and history text -------------------------------------------
  function dateNumber(match) {
    let year = Number(match[3]);
    if (match[3].length === 2) year += 2000;
    const month = Number(match[1]), day = Number(match[2]);
    if (month < 1 || month > 12 || day < 1 || day > 31 || year < 2000 || year > 2100) return null;
    return year * 10000 + month * 100 + day;
  }
  /** Puts entry above the first dated entry on or before its date (newest first). */
  function insertByDate(history, entry, entryDate) {
    const existing = String(history == null ? '' : history).replace(/\r\n?/g, '\n').trim();
    if (!existing) return entry;
    let offset = 0, insertAt = -1, anyHeader = false;
    for (const line of existing.split('\n')) {
      const match = line.match(HEADER_LINE), date = match && dateNumber(match);
      if (date) {
        anyHeader = true;
        if (date <= entryDate) { insertAt = offset; break; }
      }
      offset += line.length + 1;
    }
    if (insertAt === -1) return anyHeader ? existing + '\n\n' + entry : entry + '\n\n' + existing;
    const above = existing.slice(0, insertAt).trimEnd();
    return (above ? above + '\n\n' : '') + entry + '\n\n' + existing.slice(insertAt);
  }

  // ---- reading -----------------------------------------------------------
  function selectedStudent(ss) {
    const range = SpreadsheetApp.getActiveRange();
    if (!range || range.getSheet().getName() !== CONFIG.MASTER ||
        range.getNumRows() !== 1 || range.getRow() <= 1) {
      throw new Error('Select one student row on Master.');
    }
    const sheet = range.getSheet();
    if (sheet.getParent().getId() !== ss.getId()) throw new Error('Wrong workbook selection.');
    const key = CONFIG.studentOnRow(sheet, range.getRow());
    return {key, sheet, row: range.getRow()};
  }
  /** Whole row in two calls; get(header) refuses formulas in editable fields. */
  function rowReader(sheet, row, columns) {
    const width = Math.max(...columns.values());
    const range = sheet.getRange(row, 1, 1, width);
    const values = range.getValues()[0], formulas = range.getFormulas()[0];
    return valueReader(values, formulas, columns);
  }
  function valueReader(values, formulas, columns) {
    return {
      get(header) {
        const column = columns.get(header);
        if (!column) throw new Error('Master is missing the field: ' + header);
        if (formulas[column - 1]) throw new Error('Use a value in the editable field: ' + header + ' (Master).');
        return values[column - 1];
      }
    };
  }

  /** Shared formatter for individual and bulk logging. Dates are never changed. */
  function readEntry(central, type, ss) {
    typeInfo(type);
    if (type === FIVE) return fiveAttemptEntry(central, ss);
    const attempt = type === 'Attempt', parent = type === 'Parent Contact';
    const dateHeader = attempt || parent ? FIELDS.contactDate : FIELDS.date;
    const date = central.get(dateHeader);
    if (!(date instanceof Date) || isNaN(date.getTime())) throw new Error('Enter a valid date in ' + dateHeader + '.');
    const label = Utilities.formatDate(date, ss.getSpreadsheetTimeZone(), 'M/d/yyyy');
    const lines = ['-- ' + label + ' [' + type + ']'], powerSchoolLines = [];
    if (attempt || parent) {
      const header = attempt ? FIELDS.attemptNote : FIELDS.parentNote;
      const note = text(central.get(header));
      if (!note) throw new Error('Enter ' + header + ' first.');
      lines.push(note); powerSchoolLines.push(note);
    } else {
      FIELDS.notes.forEach(([header, category, powerSchoolLabel]) => {
        const note = text(central.get(header));
        if (!note) return;
        const powerSchool = powerSchoolLabel || category;
        lines.push(category ? category + ': ' + note : note);
        powerSchoolLines.push(powerSchool ? powerSchool + ': ' + note : note);
      });
      if (lines.length === 1) throw new Error('Enter at least one ECC note.');
    }
    return {type, date: label, entryDate: dateNumber(label.match(/^(\d+)\/(\d+)\/(\d+)$/)),
      note: lines.join('\n'), powerSchoolNote: powerSchoolLines.join('\n')};
  }

  /** Written from ECC Contact History and dated today. Reads no other field. */
  function fiveAttemptEntry(central, ss) {
    if (typeof ATTEMPT_STREAK === 'undefined') throw new Error('Add AttemptStreak.gs to log 5+ attempts.');
    const state = ATTEMPT_STREAK.status(central.get(FIELDS.history));
    const reason = ATTEMPT_STREAK.notDueReason(state);
    if (reason) throw new Error('No 5+ Attempt Log is due: ' + reason);
    const label = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'M/d/yyyy');
    const note = ATTEMPT_STREAK.summary(state);
    return {type: FIVE, date: label, entryDate: dateNumber(label.match(/^(\d+)\/(\d+)\/(\d+)$/)),
      note: '-- ' + label + ' [' + FIVE + ']\n' + note, powerSchoolNote: note};
  }

  /** Repeated entries are intentional; teachers clear source notes manually. */
  function append(history, entry, key) {
    const updated = insertByDate(history, entry.note, entry.entryDate);
    if (updated.length > CELL_LIMIT) {
      throw new Error('ECC Contact History for ' + key + ' would pass the 50,000-character cell limit. Move older entries elsewhere first.');
    }
    return {value: updated, nearLimit: updated.length > CELL_WARNING};
  }

  // ---- archive -----------------------------------------------------------
  /** ECC Contact History colors (AttemptStreak.gs). Never undoes a save. */
  function recolor() {
    if (typeof ATTEMPT_STREAK !== 'undefined') ATTEMPT_STREAK.recolorQuietly();
  }
  function archive(type, openPowerSchool) {
    const ss = SpreadsheetApp.getActiveSpreadsheet(), clock = CONFIG.timer();
    let payload, archived = false, studentNumber = '';
    const action = openPowerSchool ? 'ECC Archive & Handoff' : 'ECC Archive';
    try {
      const selection = selectedStudent(ss), key = selection.key;
      studentNumber = key;
      // Validate the optional handoff before changing any roster data.
      const info = typeInfo(type);
      let settings;
      if (openPowerSchool) {
        if (typeof POWERSCHOOL === 'undefined') throw new Error('Add EccPowerSchool.gs to use PowerSchool handoffs.');
        settings = POWERSCHOOL.settings(ss, info.workflow);
      }
      clock.lap('prepare');
      saveToMaster(master => {
        const masterRow = master.rows.get(key);
        if (!masterRow) throw new Error('Student ' + key + ' is not on Master.');
        const central = rowReader(master.sheet, masterRow, master.columns);
        const entry = readEntry(central, type, ss);
        const updated = append(central.get(FIELDS.history), entry, key);
        payload = {studentNumber: key, date: entry.date, outcome: info.workflow,
          note: openPowerSchool ? POWERSCHOOL.withOpening(settings, entry.powerSchoolNote) : entry.note,
          nearLimit: updated.nearLimit};
        return [{studentNumber: key, header: FIELDS.history, value: updated.value}];
      });
      archived = true;
      clock.lap('save');
      // Open PowerSchool first so nothing else delays the dialog.
      if (openPowerSchool) {
        POWERSCHOOL.handoff(info.handoff, {studentNumber: payload.studentNumber, date: payload.date,
          note: payload.note, outcome: payload.outcome}, settings);
        clock.lap('dialog');
      }
      recolor();
      ss.toast((type === 'Parent Contact' ? 'Parent contact' : type === FIVE ? '5+ Attempt Log' : 'ECC entry') +
        ' saved to ECC Contact History.' +
        (payload.nearLimit ? ' History is close to the cell size limit.' : ''), 'ECC Tools', 6);
      CONFIG.log('INFO', action, key, 'Entry archived to Master.', JSON.stringify({type, ms: clock.done()}));
      return payload;
    } catch (error) {
      if (archived) recolor();
      const message = (archived ? 'Entry is saved; PowerSchool did not open. ' : 'ECC archive failed. ') + error.message;
      CONFIG.log('ERROR', action, studentNumber, message, CONFIG.errorDetails(error));
      ss.toast(message, 'ECC needs attention', 12);
      throw error;
    }
  }

  /** Checks Master has every LogECC field. Writes nothing. */
  function validate() {
    saveToMaster(master => {
      if (!master.columns.has(FIELDS.contactDate) && master.columns.has(OLD_CONTACT_DATE)) {
        throw new Error('Rename the Master header ' + OLD_CONTACT_DATE + ' to ' + FIELDS.contactDate +
          ' (installRosterTools does this for you).');
      }
      [FIELDS.date, FIELDS.history, FIELDS.contactDate, FIELDS.attemptNote, FIELDS.parentNote,
        ...FIELDS.notes.map(pair => pair[0])].forEach(header => {
        if (!master.columns.has(header)) throw new Error('Missing LogECC header: ' + header);
      });
      return [];
    });
  }
  /** One-time header rename: ECC Attempt Date -> Contact Date. Changes only that
   * header cell; returns true if it renamed. Safe to run again. */
  function renameContactDateHeader() {
    const ss = CONFIG.workbook();
    return CONFIG.withLock(() => {
      const sheet = ss.getSheetByName(CONFIG.MASTER);
      const columns = CONFIG.headerColumns(sheet).columns;
      if (columns.has(FIELDS.contactDate) || !columns.has(OLD_CONTACT_DATE)) return false;
      sheet.getRange(1, columns.get(OLD_CONTACT_DATE)).setValue(FIELDS.contactDate);
      CONFIG.log('INFO', 'Setup', '', 'Renamed Master header ' + OLD_CONTACT_DATE + ' to ' + FIELDS.contactDate + '.');
      return true;
    });
  }
  return Object.freeze({archive, validate, readEntry, append, valueReader, renameContactDateHeader,
    typeInfo, recolor, TYPES, FIELDS, FIVE});
})();

function archiveEccConversation() { return ECC_LOG.archive('Conversation', false); }
function archiveEccConversationAndOpenPowerSchool() { return ECC_LOG.archive('Conversation', true); }
function archiveEccAttempt() { return ECC_LOG.archive('Attempt', false); }
function archiveEccAttemptAndOpenPowerSchool() { return ECC_LOG.archive('Attempt', true); }
function archiveParentContact() { return ECC_LOG.archive('Parent Contact', false); }
function archiveParentContactAndOpenPowerSchool() { return ECC_LOG.archive('Parent Contact', true); }
function archiveFiveAttemptLogAndOpenPowerSchool() { return ECC_LOG.archive('5+ Attempt Log', true); }
