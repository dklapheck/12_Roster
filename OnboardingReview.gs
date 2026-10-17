/** Review Class Roster onboarding changes without overwriting local work.
 * Requires Config.gs and MasterSave.gs. Imported data is always read-only.
 * Review tokens are private to the editor, expire after six hours, and carry
 * the server-side values; the dialog submits only change IDs and a decision.
 */
const ONBOARDING_REVIEW = (() => {
  const FIELDS = ['Welcome Email', 'SCC Completion', 'ISMA'];
  const ID = 'Student Number';
  const TTL = 21600;
  const pack = v => v instanceof Date ? {date: v.getTime()} : {value: v};
  const unpack = v => Object.prototype.hasOwnProperty.call(v, 'date') ? new Date(v.date) : v.value;
  const equal = (a, b) => CONFIG.same(a, b) ||
    (typeof a === 'string' && typeof b === 'string' && CONFIG.text(a) === CONFIG.text(b));

  function sourceFrom(book) {
    const result = {byId: new Map(), issues: [], available: false};
    try {
      const t = CONFIG.table(book.getSheetByName('Class Roster'), [ID]);
      result.columns = t.columns;
      result.available = true;
      const duplicate = new Set();
      t.rows.forEach((row, i) => {
        let key;
        try { key = CONFIG.studentId(row[t.columns.get(ID)]); }
        catch (e) { result.issues.push('Class Roster row ' + (i + 2) + ': ' + e.message); return; }
        if (!key) return;
        if (result.byId.has(key)) duplicate.add(key);
        result.byId.set(key, row);
      });
      // Keep duplicate IDs visible for membership, but never choose a status row.
      result.duplicate = duplicate;
      duplicate.forEach(key => result.issues.push('Class Roster has duplicate Student Number ' + key + '; onboarding fields retained.'));
      FIELDS.forEach(h => {
        if (!t.columns.has(h)) result.issues.push('Class Roster is missing ' + h + '; existing values retained.');
      });
    } catch (e) { result.issues.push('Class Roster onboarding review unavailable: ' + e.message); }
    return result;
  }

  function studentLabel(master, rows, key, rowNumber) {
    const row = rows[rowNumber - 2];
    const get = h => master.columns.has(h) ? CONFIG.text(row[master.columns.get(h) - 1]) : '';
    const name = get('Preferred Name') || (get('FIRST NAME') + ' ' + get('LAST NAME')).trim();
    return key + (name ? ' — ' + name : '');
  }

  /** Pure comparison: existing students matched by Student Number, never row. */
  function plan(master, rows, formulas, source) {
    const changes = [], issues = source.issues.slice(), missing = [];
    if (!source.available) return {changes, issues, missing};
    FIELDS.forEach(h => {
      if (!master.columns.has(h)) issues.push('Master is missing ' + h + '; field not reviewed.');
    });
    for (const [key, rowNumber] of master.rows) {
      const label = studentLabel(master, rows, key, rowNumber);
      if (!source.byId.has(key)) { missing.push(label); continue; }
      if (source.duplicate.has(key)) continue;
      FIELDS.forEach(header => {
        if (!master.columns.has(header) || !source.columns.has(header)) return;
        const column = master.columns.get(header) - 1;
        const before = rows[rowNumber - 2][column];
        const after = source.byId.get(key)[source.columns.get(header)];
        if (equal(before, after)) return;
        if (formulas[rowNumber - 2][column]) {
          issues.push(label + ': ' + header + ' has a formula on Master; retained.'); return;
        }
        changes.push({id: key + '|' + header, studentNumber: key, label, header, before: pack(before), after: pack(after)});
      });
    }
    return {changes, issues, missing};
  }

  function storeReview(changes) {
    const token = Utilities.getUuid();
    const json = JSON.stringify({workbook: CONFIG.workbook().getId(), changes});
    // Cache entries have a 100 KB limit. Fail closed instead of truncating a review.
    if (json.length * 3 > 95000) throw new Error('Too many onboarding changes for one dialog. Use a smaller Class Roster export.');
    CacheService.getUserCache().put('onboarding:' + token, json, TTL);
    return token;
  }

  function decide(token, ids, decision) {
    if (typeof token !== 'string' || !/^[A-Za-z0-9-]{1,100}$/.test(token) ||
        !Array.isArray(ids) || !ids.length || !['accept', 'reject'].includes(decision)) {
      throw new Error('Select at least one onboarding change and a valid decision.');
    }
    const ss = CONFIG.workbook();
    let resolved, pending;
    // Re-read local cells and the source within the shared save lock. Validate
    // every selected change before returning any writes: one stale cell cancels
    // the entire selection. This also protects against replay and moved rows.
    saveToMaster(master => {
      const json = CacheService.getUserCache().get('onboarding:' + token);
      if (!json) throw new Error('This review expired. Run Refresh students, contacts and grades again.');
      const review = JSON.parse(json);
      if (review.workbook !== ss.getId()) throw new Error('This review belongs to another workbook.');
      const byId = new Map(review.changes.map(c => [c.id, c]));
      const wanted = [...new Set(ids)];
      if (wanted.some(id => typeof id !== 'string' || !byId.has(id))) throw new Error('This selection is no longer pending. Refresh the review.');
      resolved = wanted.map(id => byId.get(id));
      pending = review.changes.filter(c => !wanted.includes(c.id));
      if (decision === 'reject') return [];
      const source = sourceFrom(CONFIG.sourceWorkbook(ss));
      const fields = resolved.map(c => {
        const row = master.rows.get(c.studentNumber), column = master.columns.get(c.header);
        const src = source.byId.get(c.studentNumber);
        if (!row || !column || !FIELDS.includes(c.header) || !source.available || !src ||
            source.duplicate.has(c.studentNumber) || !source.columns.has(c.header)) {
          throw new Error(c.label + ': student or field is no longer available. Nothing accepted; refresh the review.');
        }
        const cell = master.sheet.getRange(row, column);
        if (cell.getFormula() || !equal(cell.getValue(), unpack(c.before)) ||
            !equal(src[source.columns.get(c.header)], unpack(c.after))) {
          throw new Error(c.label + ': ' + c.header + ' changed since this review. Nothing accepted; refresh the review.');
        }
        return {studentNumber: c.studentNumber, header: c.header, value: unpack(c.after)};
      });
      return {fields, bulk: true};
    });
    const cache = CacheService.getUserCache();
    if (pending.length) cache.put('onboarding:' + token, JSON.stringify({workbook: ss.getId(), changes: pending}), TTL);
    else cache.remove('onboarding:' + token);
    // Rejections keep Master unchanged. A later refresh compares the source
    // again, so rejected differences can be reconsidered rather than hidden.
    CONFIG.log('SUCCESS', 'OnboardingReview', '', decision + ': ' + resolved.length + ' onboarding fields.', '');
    return {ids: resolved.map(c => c.id), decision, count: resolved.length};
  }

  function section(review, preview) {
    const esc = CONFIG.escapeHtml;
    let html = '<h3>Welcome Email, SCC Completion and ISMA</h3>';
    if (!review.changes.length) return html + '<p>No pending field changes from Class Roster.</p>';
    html += '<p>' + review.changes.length + ' differences from Class Roster. Existing Master values stay in place until accepted. ' +
      'Reject keeps your current value; a later refresh will offer the difference again. Blank imported values are shown explicitly.</p>';
    let token = '';
    if (!preview) {
      try { token = storeReview(review.changes); }
      catch (e) { html += '<p><b>Review controls unavailable:</b> ' + esc(e.message) + ' No onboarding values changed.</p>'; }
    } else html += '<p>Preview only: run Refresh students, contacts and grades to accept or reject these changes.</p>';
    if (token) html += '<p><button onclick="onboardSelect(true)">Select all</button> <button onclick="onboardSelect(false)">Select none</button></p>';
    html += '<table border="1" cellpadding="5" style="border-collapse:collapse"><tr>' +
      (token ? '<th>Select</th>' : '') + '<th>Student</th><th>Field</th><th>Current Master</th><th>Imported Class Roster</th></tr>';
    const show = p => {
      const v = unpack(p);
      return CONFIG.blank(v) ? '<i>(blank)</i>' : esc(v instanceof Date ?
        Utilities.formatDate(v, CONFIG.workbook().getSpreadsheetTimeZone(), 'yyyy-MM-dd') : v);
    };
    review.changes.forEach(c => {
      html += '<tr>' + (token ? '<td><input type="checkbox" class="onboard" value="' + esc(c.id) + '" checked></td>' : '') +
        '<td>' + esc(c.label) + '</td><td>' + esc(c.header) + '</td><td>' + show(c.before) + '</td><td>' + show(c.after) + '</td></tr>';
    });
    html += '</table>';
    if (!token) return html;
    html += '<p><button class="onboardAction" onclick="onboardDecide(\'accept\',false)">Accept selected</button> ' +
      '<button class="onboardAction" onclick="onboardDecide(\'reject\',false)">Reject selected</button> ' +
      '<button class="onboardAction" onclick="onboardDecide(\'accept\',true)">Accept all</button> ' +
      '<button class="onboardAction" onclick="onboardDecide(\'reject\',true)">Reject all</button> <span id="onboardMsg"></span></p>' +
      '<script>function onboardSelect(on){document.querySelectorAll("input.onboard:not(:disabled)").forEach(x=>x.checked=on);}' +
      'function onboardDecide(decision,all){const boxes=[...document.querySelectorAll("input.onboard:not(:disabled)")];' +
      'const ids=boxes.filter(x=>all||x.checked).map(x=>x.value);const msg=document.getElementById("onboardMsg");' +
      'if(!ids.length){msg.textContent="No pending changes selected.";return;}' +
      'document.querySelectorAll(".onboardAction,input.onboard").forEach(x=>x.disabled=true);msg.textContent="Saving decision…";' +
      'google.script.run.withSuccessHandler(r=>{boxes.forEach(x=>{if(r.ids.includes(x.value)){' +
      'x.dataset.resolved="true";x.checked=false;x.closest("tr").style.opacity="0.5";}});' +
      'msg.textContent=(r.decision==="accept"?"Accepted ":"Rejected ")+r.count+" change(s).";onboardEnable();})' +
      '.withFailureHandler(e=>{msg.textContent=e.message;onboardEnable();})' +
      '.reviewOnboardingChanges(' + JSON.stringify(token) + ',ids,decision);}' +
      'function onboardEnable(){document.querySelectorAll(".onboardAction").forEach(x=>x.disabled=false);' +
      'document.querySelectorAll("input.onboard").forEach(x=>x.disabled=x.dataset.resolved==="true");}</script>';
    return html;
  }

  return Object.freeze({sourceFrom, plan, studentLabel, section, decide});
})();

/** Called by the refresh report; only selected onboarding fields can be saved. */
function reviewOnboardingChanges(token, ids, decision) { return ONBOARDING_REVIEW.decide(token, ids, decision); }
