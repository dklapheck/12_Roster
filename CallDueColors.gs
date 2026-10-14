/** 10RosterHR Call Due colors v1.0 (replaces EccDateColors.gs; delete that file).
 * Requires Config.gs. Menu: Update > ECC Tools > Refresh Call Due colors.
 * Pull Data > Refresh students also reapplies these rules at the end of each run.
 *
 * Colors the ECC Date and Call Due cells on the ECC tab and Master by when the
 * student's next ECC is due:
 *   due date = ECC Date + 20 weekdays (WORKDAY), or Upstream's Call Due when
 *   there is no ECC Date yet. (ECC calls are due every 20 school days; weekdays
 *   are used to stay on the safe side, since holidays only push the real date later.)
 *
 *   No ECC Date and no Call Due  -> maroon, white text
 *   Overdue (due before today)   -> red, white text
 *   Due within 2 weekdays        -> hot pink
 *   Due in 3-5 weekdays          -> orange
 *   Due in 6-10 weekdays         -> yellow
 *   Due in 11+ weekdays          -> green
 *
 * The rules live in the sheet and recalculate every day on their own (TODAY()).
 * Re-running removes every rule this script (or the old EccDateColors script)
 * made, wherever it now sits, then adds a fresh set. Other rules are untouched.
 */
const HR10DUECOLORS = (() => {
  const TABS = ['ECC', 'Master'];
  const MARK = 'N("HR10DUE")=0';   // always TRUE; identifies this script's rules
  const WEEKDAYS = 20;
  const BANDS = [
    // [name, test on DUE, background, font]
    ['overdue', due => due + '<TODAY()', '#CC0000', '#FFFFFF'],
    ['pink', due => due + '<=WORKDAY(TODAY(),2)', '#FF66C4', null],
    ['orange', due => due + '<=WORKDAY(TODAY(),5)', '#F6B26B', null],
    ['yellow', due => due + '<=WORKDAY(TODAY(),10)', '#FFE599', null],
    ['green', () => 'TRUE', '#93C47D', null]
  ];

  /** Conditional-format formulas for one tab, relative to row 2. */
  function formulas(eccLetter, dueLetter, idLetter) {
    const e = '$' + eccLetter + '2', c = '$' + dueLetter + '2', s = '$' + idLetter + '2';
    const has = 'OR(ISNUMBER(' + e + '),ISNUMBER(' + c + '))';
    const due = 'IF(ISNUMBER(' + e + '),WORKDAY(' + e + ',' + WEEKDAYS + '),' + c + ')';
    const list = [{name: 'missing', formula: '=AND(' + MARK + ',' + s + '<>"",NOT(' + has + '))', background: '#7B1113', font: '#FFFFFF'}];
    BANDS.forEach(([name, test, background, font]) => {
      list.push({name, formula: '=AND(' + MARK + ',' + has + ',' + test(due) + ')', background, font});
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
    const cols = HR10CFG.headerColumns(sheet).columns;
    ['ECC Date', 'Call Due', 'Student Number'].forEach(h => {
      if (!cols.has(h)) throw new Error(sheet.getName() + ': missing header ' + h + '.');
    });
    const ecc = cols.get('ECC Date'), due = cols.get('Call Due');
    const L = HR10CFG.columnLetter;
    const height = Math.max(sheet.getMaxRows() - 1, 1);
    const ranges = [sheet.getRange(2, ecc, height, 1), sheet.getRange(2, due, height, 1)];
    const all = sheet.getConditionalFormatRules();
    const kept = all.filter(r => !ours(r));
    const fresh = formulas(L(ecc), L(due), L(cols.get('Student Number'))).map(x => {
      const b = SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(x.formula)
        .setBackground(x.background).setRanges(ranges);
      if (x.font) b.setFontColor(x.font);
      return b.build();
    });
    // Ours go first so they win over any older rule on the same cells.
    sheet.setConditionalFormatRules(fresh.concat(kept));
    return {tab: sheet.getName(), replaced: all.length - kept.length};
  }

  function apply(quiet) {
    const ss = HR10CFG.workbook();
    const results = TABS.map(name => {
      const sheet = ss.getSheetByName(name);
      if (!sheet) throw new Error(name + ' tab not found.');
      return applyTo(sheet);
    });
    if (!quiet) ss.toast('Call Due colors applied to ' + TABS.join(' and ') + '.', 'ECC Tools', 5);
    return results;
  }

  return Object.freeze({apply, formulas, ours, TABS});
})();

function applyCallDueColors() { return HR10DUECOLORS.apply(false); }
/** Old menu name, kept so an existing button or menu still works. */
function applyEccDateColors() { return applyCallDueColors(); }
