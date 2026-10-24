/** All 12RosterORN menus (v4.0). Requires Config.gs.
 * v3.0: Master only. Check roster sync and Add columns were removed (there is
 * nothing to sync, and a column is added by inserting it on Master).
 * Functions were renamed; OldNames.gs keeps the old names working until any
 * buttons are reassigned.
 * onOpen is a simple trigger: it runs for every editor who opens the file, with
 * no install step and no per-person trigger. It only builds menus.
 * This must be the only onOpen function in the project.
 */
function onOpen() { buildMenus_(); }

function buildMenus_() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu('Update')
    // Frequently used quick-access actions
    .addItem('Archive and Log All ECC', 'archiveAndLogAllEcc')
    .addItem('Open Student Demographics', 'openStudentDemographics')
    .addSeparator()
    .addSubMenu(
      ui.createMenu('ECC & Demographics')
        .addItem('Archive and Log All ECC', 'archiveAndLogAllEcc')
        .addItem('Log selected student\u2019s conversation', 'archiveEccConversationAndOpenPowerSchool')
        .addItem('Log selected student\u2019s attempt', 'archiveEccAttemptAndOpenPowerSchool')
        .addItem('Open Student Demographics', 'openStudentDemographics')
    )
    .addSubMenu(
      ui.createMenu('Call / SCC')
        .addItem('Save Call to Master', 'saveCallToMaster')
        .addItem('Log SCC success / failed attempt in PowerSchool', 'logSccInPowerSchool')
        .addItem('Save Call & open PowerSchool', 'saveCallAndOpenPowerSchool')
        .addItem('Reset Call form', 'resetCallForm')
    )
    .addSubMenu(
      ui.createMenu('Refresh Roster Data')
        .addItem('Preview source changes', 'previewStudentRefresh')
        .addItem('Refresh students, contacts and grades', 'refreshStudentData')
    )
    .addSubMenu(
      ui.createMenu('Grade Report Photos')
        .addItem('Preview grade report photos', 'previewGradeReportPhotos')
        .addItem('Import grade report photos', 'importGradeReportPhotos')
    )
    .addSubMenu(
      ui.createMenu('WIG')
        .addItem('Preview this week\u2019s snapshot', 'previewWigSnapshot')
        .addItem('Save this week\u2019s snapshot', 'saveWigSnapshot')
        .addItem('Reformat student lists', 'recolorWigStudentRepeats')
    )
    .addSubMenu(
      ui.createMenu('Maintenance')
        .addItem('Refresh Call Due colors', 'applyCallDueColors')
        .addItem('Extend Master filter to all columns', 'extendMasterFilter')
        .addItem('Check setup', 'checkRosterSetup')
    )
    .addToUi();
}

/** One setup for every tool: checks Master's headers, the Call named ranges and
 * the PowerSchool settings, then shows the menus. Changes no student data.
 * Safe to run again. */
function installRosterTools() {
  const ss = CONFIG.workbook();
  saveToMaster(master => { CALL_FORM.validate(master); return []; });
  ECC_LOG.validate();
  if (!ss.getSheetByName('Call') || !ss.getSheetByName('Info')) throw new Error('Call and Info tabs must exist.');
  if (typeof POWERSCHOOL === 'undefined') throw new Error('Add EccPowerSchool.gs before installing.');
  ['SCC Success', 'SCC Attempt', 'ECC Conversation', 'ECC Attempt'].forEach(w => POWERSCHOOL.settings(ss, w));
  CALL_FORM.setupNames();
  buildMenus_();
  ss.toast('Menus ready. Student data was not changed.', 'Roster tools', 8);
}

/** Validate existing setup without creating names or changing any cells. */
function checkRosterSetup() {
  const ss = CONFIG.workbook();
  saveToMaster(master => { CALL_FORM.validate(master); return []; });
  ECC_LOG.validate();
  if (!ss.getSheetByName('Call') || !ss.getSheetByName('Info')) throw new Error('Call and Info tabs must exist.');
  ['SCC Success', 'SCC Attempt', 'ECC Conversation', 'ECC Attempt'].forEach(w => POWERSCHOOL.settings(ss, w));
  ss.toast('Roster headers and PowerSchool settings checked.', 'Roster tools', 8);
}
