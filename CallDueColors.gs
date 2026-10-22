/** 12RosterORN Call Due colors v3.0. Master only (the ECC tab was removed).
 * Requires Config.gs. Menu: Update > ECC Tools > Refresh Call Due colors.
 * Pull Data > Refresh students also reapplies these rules at the end of each run.
 *
 * Colors each date column on Master independently:
 *   ECC Date: use 20 weekdays after that cell's date (WORKDAY).
 *   Call Due: use that cell's date directly.
 * One column's date never controls the other column's color. A missing or
 * invalid date is maroon in that column only, when a Student Number is present.
 * WORKDAY skips weekends, but no holiday or school-break calendar is supplied.
 *
 *   Missing/invalid date        -> maroon, white text
 *   Overdue (deadline passed)   -> red, white text
 *   Due today or within 2 weekdays -> hot pink
 *   Due in 3-5 weekdays          -> orange
 *   Due in 6-10 weekdays         -> yellow
 *   Due in 11+ weekdays          -> green
 *
 * The rules live in the sheet and recalculate every day on their own (TODAY()).
 * Re-running removes every rule this script (or the old EccDateColors script)
 * made, wherever it now sits, then adds a fresh set. Other rules are untouched.
 */
const CALL_DUE_COLORS = (() => {
  const TABS = ['Master'];   // a literal: CONFIG may not be loaded yet when this file loads
  const MARK = 'N("HR10DUE")=0';   // always TRUE; identifies this script's rules (keep: existing rules use it)
  const WEEKDAYS = 20;
  const BANDS = [
    // [name, test on DUE, background, font]
    ['overdue', due => due + '<TODAY()', '#CC0000', '#FFFFFF'],
    ['pink', due => due + '<=WORKDAY(TODAY(),2)', '#FF66C4', null],
    ['orange', due => due + '<=WORKDAY(TODAY(),5)', '#F6B26B', null],
    ['yellow', due => due + '<=WORKDAY(TODAY(),10)', '#FFE599', null],
    ['green', () => 'TRUE', '#93C47D', null]
  ];

  /** Independent rules for each column, relative to row 2. */
  function formulas(eccLetter, dueLetter, idLetter) {
    const s = '$' + idLetter + '2', list = [];
    [{column: eccLetter, weekdays: WEEKDAYS}, {column: dueLetter, weekdays: 0}].forEach(({column, weekdays}) => {
      const date = '$' + column + '2', has = 'ISNUMBER(' + date + ')';
      const deadline = weekdays ? 'WORKDAY(' + date + ',' + weekdays + ')' : date;
      // Keep comparisons safe for invalid dates; has excludes them from the bands.
      const due = 'IF(' + has + ',' + deadline + ',0)';
      list.push({column, name: 'missing', formula: '=AND(' + MARK + ',' + s + '<>"",NOT(' + has + '))', background: '#7B1113', font: '#FFFFFF'});
      BANDS.forEach(([name, test, background, font]) => {
        list.push({column, name, formula: '=AND(' + MARK + ',' + has + ',' + test(due) + ')', background, font});
      });
    });
    return list;
  }

  /** Old EccDateColors rules: =AND(ISNUMBER($AD2),...) or =AND($AD2="",$B2<>""). */
  const OLD = /^=AND\((?:ISNUMBER\(\$[A-Z]+\d+\)|\$[A-Z]+\d+="",\$[A-Z]+\d+<>"")/i;
  function ours(rule) {
    const condition = rule.getBooleanCondition();
    if (!condition || condition.getCriteriaType() !== SpreadsheetApp.BooleanCriteria.CUSTOM_FORMULA) return false;
    const f = String(condition.getCriteriaValues()[0] || '').replace(/\s+/g, '');
    if (f.indexOf('N("HR10DUE")') >= 0) return true;
    // Old rules are recognised only when they mention TODAY() or the blank-cell
    // test, so grade rules like =AND(ISNUMBER(AW2),AW2<60) are never removed.
    return OLD.test(f) && (/TODAY\(\)/i.test(f) || /^=AND\(\$[A-Z]+\d+="",/i.test(f));
  }

  function applyTo(sheet) {
    const cols = CONFIG.headerColumns(sheet).columns;
    ['ECC Date', 'Call Due', 'Student Number'].forEach(h => {
      if (!cols.has(h)) throw new Error(sheet.getName() + ': missing header ' + h + '.');
    });
    const ecc = cols.get('ECC Date'), due = cols.get('Call Due');
    const L = CONFIG.columnLetter;
    const height = Math.max(sheet.getMaxRows() - 1, 1);
    const ranges = new Map([[L(ecc), sheet.getRange(2, ecc, height, 1)],
      [L(due), sheet.getRange(2, due, height, 1)]]);
    const all = sheet.getConditionalFormatRules();
    const kept = all.filter(r => !ours(r));
    const fresh = formulas(L(ecc), L(due), L(cols.get('Student Number'))).map(x => {
      const b = SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(x.formula)
        .setBackground(x.background).setRanges([ranges.get(x.column)]);
      if (x.font) b.setFontColor(x.font);
      return b.build();
    });
    // Ours go first so they win over any older rule on the same cells.
    sheet.setConditionalFormatRules(fresh.concat(kept));
    return {tab: sheet.getName(), replaced: all.length - kept.length};
  }

  function apply(quiet) {
    const ss = CONFIG.workbook();
    const results = TABS.map(name => {
      const sheet = ss.getSheetByName(name);
      if (!sheet) throw new Error(name + ' tab not found.');
      return applyTo(sheet);
    });
    if (!quiet) ss.toast('Call Due colors applied to Master.', 'ECC Tools', 5);
    return results;
  }

  return Object.freeze({apply, formulas, ours, TABS});
})();

function applyCallDueColors() { return CALL_DUE_COLORS.apply(false); }
