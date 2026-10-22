/** Master saves v2.0 for 12RosterORN. Requires Config.gs.
 * Replaces RosterSync.gs: SCC, ECC and STAR were removed, so every roster
 * save writes to Master only. There is no edit trigger to install.
 *
 * Every tool saves through saveToMaster(planner):
 *  1. takes the shared document lock (two saves can never overlap),
 *  2. reads Master's headers and Student Numbers,
 *  3. calls planner(master), which reads what it needs and returns the changes,
 *  4. writes those cells to Master, then releases the lock.
 * Reading and writing happen under one lock, so nothing can change in between.
 *
 * master (what the planner receives):
 *   {sheet, headers, columns (header -> column number), idColumn, lastRow,
 *    rows (Student Number -> row number), rowKeys (row number -> Student Number)}
 *
 * The planner returns either
 *   [{studentNumber, header, value}, ...]
 * or {fields: [...], relatedWrites: [...], bulk: true, dryRun: true}
 *   relatedWrites: [{sheet: 'Contacts' (row 2+) or 'Info' (A3 only), row, column, values: [[...]]}]
 *   bulk: joins neighbouring rows into fewer writes (Pull Data).
 *   dryRun: plan only; nothing is written (Preview).
 * Return [] to read Master under the lock without saving anything.
 * The planner must not write cells itself or take the lock again.
 *
 * Student Number, Name and Subject Line are never written here. Name and
 * Subject Line are formulas on Master and recalculate on their own.
 * Text that Sheets would reinterpret (3/4, 00123, TRUE, =x) is saved as text.
 * If one write fails, the cells already written by this save are put back.
 */
const MASTER_SAVE = (() => {
  const ID_HEADER = 'Student Number';
  const PROTECTED = ['Student Number', 'Name', 'Subject Line'];

  function readMaster(ss) {
    const sheet = ss.getSheetByName(CONFIG.MASTER);
    if (!sheet) throw new Error('The Master tab is missing.');
    const {headers, columns} = CONFIG.headerColumns(sheet, true);
    if (!columns.has(ID_HEADER)) throw new Error('Master: missing header ' + ID_HEADER + '.');
    const idColumn = columns.get(ID_HEADER);
    const lastRow = Math.max(sheet.getLastRow(), 1);
    const ids = lastRow > 1 ? sheet.getRange(2, idColumn, lastRow - 1, 1).getValues() : [];
    const rows = new Map(), rowKeys = new Map();
    ids.forEach((entry, i) => {
      const key = CONFIG.studentId(entry[0]);
      if (!key) return;
      if (rows.has(key)) throw new Error('Master: duplicate Student Number ' + key + '.');
      rows.set(key, i + 2);
      rowKeys.set(i + 2, key);
    });
    if (!rows.size) throw new Error('Master has no student records.');
    return {sheet, headers, columns, idColumn, lastRow, rows, rowKeys};
  }

  function validValue(value) {
    if (value instanceof Date) return !isNaN(value.getTime());
    if (typeof value === 'number') return Number.isFinite(value);
    return typeof value === 'string' || typeof value === 'boolean';
  }

  /** Checks the planner's list; returns [{row, column, value}]. */
  function checkFields(master, planned) {
    const seen = new Set();
    return planned.map(item => {
      const key = CONFIG.studentId(item.studentNumber), header = CONFIG.normalize(item.header);
      const row = master.rows.get(key);
      if (!row) throw new Error('Student Number is not on Master: ' + key);
      if (!master.columns.has(header) || PROTECTED.includes(header)) throw new Error('Cannot save this Master field: ' + header);
      const token = key + '|' + header;
      if (seen.has(token)) throw new Error('The same field is listed twice for ' + key + ': ' + header);
      seen.add(token);
      if (!validValue(item.value)) throw new Error('Use a value or an empty string for ' + header + ' (' + key + ').');
      return {row, column: master.columns.get(header), value: item.value};
    });
  }

  /** Neighbouring cells in one row become one block. */
  function rowBlocks(cells) {
    const byRow = new Map(), blocks = [];
    cells.forEach(cell => {
      if (!byRow.has(cell.row)) byRow.set(cell.row, []);
      byRow.get(cell.row).push(cell);
    });
    byRow.forEach((list, row) => {
      list.sort((a, b) => a.column - b.column);
      for (let i = 0; i < list.length;) {
        const column = list[i].column, values = [];
        while (i < list.length && list[i].column === column + values.length) values.push(list[i++].value);
        blocks.push({row, column, values: [values]});
      }
    });
    return blocks;
  }

  /** bulk: blocks in consecutive rows that cover the same columns become one write. */
  function stackBlocks(blocks) {
    const groups = new Map(), result = [];
    blocks.forEach(block => {
      const key = block.column + ':' + block.values[0].length;
      if (!groups.has(key)) groups.set(key, []);
      groups.get(key).push(block);
    });
    groups.forEach(group => {
      group.sort((a, b) => a.row - b.row);
      for (let i = 0; i < group.length;) {
        const first = {row: group[i].row, column: group[i].column, values: group[i].values.slice()};
        i++;
        while (i < group.length && group[i].row === first.row + first.values.length) first.values.push(...group[i++].values);
        result.push(first);
      }
    });
    return result;
  }

  /** A write plus what was there before, so it can be put back. */
  function operation(range, values) {
    const before = range.getValues(), formulas = range.getFormulas();
    return {range, values,
      before: before.map((row, r) => row.map((value, c) => formulas[r][c] || value)),
      beforeFormulas: formulas.map(row => row.map(Boolean))};
  }

  function relatedOperations(ss, related) {
    const occupied = new Set(), ops = [];
    let cells = 0;
    related.forEach(item => {
      const sheet = ss.getSheetByName(item.sheet), values = item.values;
      if (!sheet || !Array.isArray(values) || !values.length || !Array.isArray(values[0]) || !values[0].length ||
          !Number.isInteger(item.row) || !Number.isInteger(item.column)) throw new Error('Invalid related write.');
      const height = values.length, width = values[0].length;
      if (!(item.sheet === 'Contacts' && item.row >= 2 && item.column >= 1 ||
            item.sheet === 'Info' && item.row === 3 && item.column === 1 && height === 1 && width === 1)) {
        throw new Error('Related writes are limited to Contacts data and Info!A3.');
      }
      if (item.row + height - 1 > sheet.getMaxRows() || item.column + width - 1 > sheet.getMaxColumns()) {
        throw new Error('Add enough empty rows/columns to ' + item.sheet + ' first.');
      }
      values.forEach((row, r) => {
        if (!Array.isArray(row) || row.length !== width) throw new Error('A related write has uneven rows.');
        row.forEach((value, c) => {
          if (!validValue(value)) throw new Error('Invalid value in a related write to ' + item.sheet + '.');
          const token = item.sheet + '|' + (item.row + r) + '|' + (item.column + c);
          if (occupied.has(token)) throw new Error('Related writes overlap on ' + item.sheet + '.');
          occupied.add(token);
        });
      });
      const range = sheet.getRange(item.row, item.column, height, width);
      if (range.getFormulas().some(row => row.some(Boolean))) throw new Error('A related write would overwrite a formula on ' + item.sheet + '.');
      ops.push(operation(range, values));
      cells += height * width;
    });
    return {ops, cells};
  }

  function write(ops) {
    ops.forEach(op => { op.ready = CONFIG.prepareValues(op.range, op.values); });
    const done = [];
    try {
      ops.forEach(op => { done.push(op); op.range.setValues(op.ready); });
      SpreadsheetApp.flush();
    } catch (error) {
      const failures = [];
      done.reverse().forEach(op => {
        try { op.range.setValues(CONFIG.prepareValues(op.range, op.before, op.beforeFormulas)); }
        catch (undoError) { failures.push(String(undoError)); }
      });
      try { SpreadsheetApp.flush(); } catch (flushError) { failures.push(String(flushError)); }
      throw new Error('Save failed: ' + error.message + (failures.length ?
        '. Some cells could not be put back; review Master and Contacts.' : '. Cells saved by this action were put back.'));
    }
  }

  function save(planner) {
    if (typeof planner !== 'function') throw new Error('saveToMaster needs a planner function.');
    const ss = CONFIG.workbook();
    return CONFIG.withLock(() => {
      const master = readMaster(ss);
      const packet = planner(master);
      const planned = Array.isArray(packet) ? packet : packet && packet.fields;
      const related = Array.isArray(packet) ? [] : packet && packet.relatedWrites || [];
      if (!Array.isArray(planned) || !Array.isArray(related)) throw new Error('The planner must return an array or {fields, relatedWrites}.');
      const cells = checkFields(master, planned);
      const dryRun = !Array.isArray(packet) && packet.dryRun === true;
      const bulk = !Array.isArray(packet) && packet.bulk === true;
      const students = new Set(cells.map(cell => cell.row)).size;
      let relatedCells = 0;
      if (!dryRun && (cells.length || related.length)) {
        const blocks = bulk ? stackBlocks(rowBlocks(cells)) : rowBlocks(cells);
        const ops = blocks.map(b => operation(master.sheet.getRange(b.row, b.column, b.values.length, b.values[0].length), b.values));
        const extra = relatedOperations(ss, related);
        relatedCells = extra.cells;
        write(ops.concat(extra.ops));
      } else if (dryRun) {
        relatedCells = related.reduce((n, item) => n + (item.values || []).length * ((item.values || [[]])[0] || []).length, 0);
      }
      return {cells: cells.length, students, relatedCells, dryRun};
    }, 'Roster tools are busy. Please retry.');
  }

  return Object.freeze({save, readMaster});
})();

/** The one save path for every roster tool. See the notes at the top of this file. */
function saveToMaster(planner) { return MASTER_SAVE.save(planner); }
