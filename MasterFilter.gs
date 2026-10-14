/** Extend Master's filter v2.0 for 12RosterORN. Requires Config.gs; menu in Menus.gs.
 * Replaces RosterColumns.gs. With Master as the only roster tab, add a column by
 * inserting it on Master directly; no script is needed.
 * Sheets only grows a filter for columns inserted inside it. If you add columns
 * to the right of Master's filter, Pull Data > Extend Master filter to all columns
 * rebuilds the filter to reach the last header, keeping every column's conditions.
 */
const MASTER_FILTER = (() => {
  /** Rebuilds the sheet's basic filter so it reaches lastColumn. No filter: nothing to do. */
  function extendFilter(sheet, lastColumn) {
    const filter = sheet.getFilter();
    if (!filter) return false;
    const range = filter.getRange();
    const first = range.getColumn(), end = range.getLastColumn();
    if (end >= lastColumn) return false;
    const criteria = [];
    for (let c = first; c <= end; c++) {
      const value = filter.getColumnFilterCriteria(c);
      if (value) criteria.push([c, value]);
    }
    filter.remove();
    const rebuilt = sheet.getRange(range.getRow(), first, range.getNumRows(), lastColumn - first + 1).createFilter();
    criteria.forEach(([c, value]) => rebuilt.setColumnFilterCriteria(c, value));
    return true;
  }

  function extendMaster() {
    const ss = CONFIG.workbook();
    const sheet = ss.getSheetByName(CONFIG.MASTER);
    if (!sheet) throw new Error('Master tab is missing.');
    const changed = CONFIG.withLock(() => extendFilter(sheet, CONFIG.headerColumns(sheet).headers.length),
      'Roster tools are busy. Please retry.');
    CONFIG.log('SUCCESS', 'ExtendMasterFilter', '', changed ? 'Extended.' : 'Already covered.', '');
    ss.toast(changed ? 'The filter on Master now covers every column.' :
      'The filter already covers every column (or Master has no filter).', 'Master filter', 8);
    return changed;
  }

  return Object.freeze({extendFilter, extendMaster});
})();

function extendMasterFilter() { return MASTER_FILTER.extendMaster(); }
