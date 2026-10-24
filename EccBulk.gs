/** Review all Master ECC notes, archive selected entries, and prepare PS tabs.
 * Requires LogECC.gs, EccPowerSchool.gs, Config.gs and MasterSave.gs.
 * Reviews use per-editor tokens. Only IDs are accepted from the dialog; source
 * dates/notes are checked again by Student Number before saving. Reopening the
 * review intentionally offers unchanged notes again. No source cells are cleared.
 */
const ECC_BULK = (() => {
  const TTL = 21600, PREFIX = 'ecc-bulk:';
  const headersFor = type => {
    const F = ECC_LOG.FIELDS;
    return type === 'Attempt' ? [F.attemptDate, F.attemptNote] : [F.date, ...F.notes.map(pair => pair[0])];
  };

  function signature(values, formulas, columns, type) {
    const json = JSON.stringify(headersFor(type).map(header => {
      const index = columns.get(header) - 1;
      const value = values[index];
      return [header, value instanceof Date ? {date: value.getTime()} : value, formulas[index]];
    }));
    return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, json, Utilities.Charset.UTF_8)
      .map(byte => ('0' + ((byte + 256) % 256).toString(16)).slice(-2)).join('');
  }

  /** Includes filtered/hidden students. Two reads for all values and formulas. */
  function scan(master, ss) {
    const F = ECC_LOG.FIELDS;
    const required = [F.date, F.history, F.attemptDate, F.attemptNote, ...F.notes.map(pair => pair[0])];
    required.forEach(h => { if (!master.columns.has(h)) throw new Error('Missing LogECC header: ' + h); });
    const width = master.headers.length;
    const range = master.sheet.getRange(2, 1, master.lastRow - 1, width);
    const values = range.getValues(), formulas = range.getFormulas();
    const today = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'yyyyMMdd');
    const records = new Map();
    master.rows.forEach((row, key) => {
      const cells = values[row - 2], mask = formulas[row - 2];
      const raw = h => cells[master.columns.get(h) - 1];
      const label = CONFIG.text(raw('Preferred Name')) ||
        [CONFIG.text(raw('FIRST NAME')), CONFIG.text(raw('LAST NAME'))].filter(Boolean).join(' ') || key;
      const reader = ECC_LOG.valueReader(cells, mask, master.columns);
      ['Conversation', 'Attempt'].forEach(type => {
        const triggers = type === 'Attempt' ? [F.attemptNote] : F.notes.slice(0, 4).map(pair => pair[0]);
        if (!triggers.some(h => CONFIG.text(raw(h)) || mask[master.columns.get(h) - 1])) return;
        const id = key + '|' + type;
        const record = {id, studentNumber: key, label, type,
          signature: signature(cells, mask, master.columns, type), reader};
        try {
          record.entry = ECC_LOG.readEntry(reader, type, ss);
          record.oldDate = String(record.entry.entryDate) < today;
          // Reject formula history destinations; combined size is checked on save.
          reader.get(F.history);
        } catch (error) { record.error = error.message; }
        records.set(id, record);
      });
    });
    return records;
  }

  function put(token, review) {
    const json = JSON.stringify(review);
    // A cache item is limited to 100 KB; fingerprints keep note bodies out of it.
    if (json.length * 3 > 95000) throw new Error('This ECC review is too large. Archive individual students instead.');
    CacheService.getUserCache().put(PREFIX + token, json, TTL);
  }

  function get(token) {
    if (typeof token !== 'string' || !/^[A-Za-z0-9-]{1,100}$/.test(token)) throw new Error('Invalid ECC review.');
    const json = CacheService.getUserCache().get(PREFIX + token);
    if (!json) throw new Error('This ECC review expired. Close it and choose Archive and Log All ECC again.');
    const review = JSON.parse(json);
    if (review.workbook !== CONFIG.workbook().getId()) throw new Error('This review belongs to another workbook.');
    return review;
  }

  function review() {
    const ss = CONFIG.workbook();
    let records;
    saveToMaster(master => { records = scan(master, ss); return []; });
    const token = Utilities.getUuid();
    put(token, {workbook: ss.getId(), state: 'pending', entries: [...records.values()].map(r => ({
      id: r.id, signature: r.signature, requestId: Utilities.getUuid()
    }))});
    return {token, entries: [...records.values()].map(r => ({
      id: r.id, studentNumber: r.studentNumber, label: r.label, type: r.type,
      date: r.entry ? r.entry.date : '', oldDate: Boolean(r.oldDate),
      note: r.entry ? r.entry.note : '', error: r.error || ''
    }))};
  }

  function archive(token, ids) {
    const F = ECC_LOG.FIELDS;
    if (!Array.isArray(ids) || !ids.length || ids.some(id => typeof id !== 'string') ||
        new Set(ids).size !== ids.length) throw new Error('Select at least one ECC entry.');
    const ss = CONFIG.workbook();
    let state, result, claimed = false;
    try {
      saveToMaster(master => {
        state = get(token);
        if (state.state === 'saving') throw new Error('This ECC batch is still saving. Wait before trying again.');
        const wanted = new Map(state.entries.map(entry => [entry.id, entry]));
        if (ids.some(id => !wanted.has(id))) throw new Error('Invalid ECC selection. Reopen the review.');
        if (state.state === 'saved' && JSON.stringify([...ids].sort()) !== JSON.stringify([...state.selected].sort())) {
          throw new Error('This review has already been saved. Reopen it to archive another selection.');
        }
        const current = scan(master, ss), histories = new Map(), settings = state.settings || {}, links = [];
        // Validate the full selection and build every URL before returning writes.
        ids.forEach(id => {
          const record = current.get(id), reviewed = wanted.get(id);
          if (!record || record.signature !== reviewed.signature) {
            throw new Error('ECC dates or notes changed since this review (' + id + '). Reopen the review.');
          }
          if (record.error) throw new Error(record.label + ': ' + record.error);
          const type = record.type, entry = record.entry, key = record.studentNumber;
          if (!settings[type]) settings[type] = POWERSCHOOL.settings(ss, 'ECC ' + type);
          let nearLimit = false;
          if (state.state !== 'saved') {
            const history = histories.has(key) ? histories.get(key) : record.reader.get(F.history);
            const updated = ECC_LOG.append(history, entry, key);
            histories.set(key, updated.value); nearLimit = updated.nearLimit;
          }
          const url = POWERSCHOOL.url('ecc', {
            studentNumber: key, date: entry.date, outcome: 'ECC ' + type,
            note: POWERSCHOOL.withOpening(settings[type], entry.powerSchoolNote)
          }, settings[type], reviewed.requestId);
          links.push({id, label: record.label, type, date: entry.date, url, nearLimit});
        });
        result = {links, students: new Set(links.map(link => link.id.split('|')[0])).size};
        // Retrying the same dialog after a lost response returns the same links,
        // without archiving that request twice. A NEW review always logs again.
        if (state.state === 'saved') return [];
        state.state = 'saving'; state.selected = ids; state.settings = settings;
        put(token, state); claimed = true;
        return {bulk: true, fields: [...histories].map(([studentNumber, value]) => ({
          studentNumber, header: F.history, value
        }))};
      });
    } catch (error) {
      if (claimed) { state.state = 'pending'; put(token, state); }
      throw error;
    }
    if (claimed) { state.state = 'saved'; put(token, state); }
    CONFIG.log('INFO', 'Bulk ECC archive', '', 'Archived selection; PowerSchool submission is manual.',
      JSON.stringify({students: result.students, entries: result.links.length}));
    return result;
  }

  function open() {
    CONFIG.workbook();
    const output = HtmlService.createHtmlOutputFromFile('EccBulkDialog').setWidth(960).setHeight(640);
    SpreadsheetApp.getUi().showModalDialog(output, 'Archive and Log All ECC');
  }
  return Object.freeze({review, archive, open});
})();

function archiveAndLogAllEcc() { return ECC_BULK.open(); }
function getEccBulkReview() { return ECC_BULK.review(); }
function archiveSelectedEcc(token, ids) { return ECC_BULK.archive(token, ids); }
