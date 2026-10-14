/** Old function names (12RosterORN v3.0). Temporary.
 * Functions were renamed when SCC, ECC and STAR were removed. A drawing button
 * on a tab (for example on Call) still runs the old name it was assigned, so
 * each old name below simply runs the new function.
 *
 * TO RETIRE THIS FILE: on each button, click it, open its three-dot menu,
 * choose Assign script, and type the new name shown on the right. When no
 * button uses an old name, delete this file.
 */
function saveSccToRoster() { return saveCallToMaster(); }
function archiveCurrentSCCNote() { return saveCallToMaster(); }
function archiveSccAndOpenPowerSchool() { return saveCallAndOpenPowerSchool(); }
// The original button with this name only logged in PowerSchool (it did not save).
function saveSccAndOpenPowerSchool() { return logSccInPowerSchool(); }
function resetCallEntry() { return resetCallForm(); }
function setup10CallNamedRanges() { return setupCallNamedRanges(); }

function archiveCurrentECCNote() { return archiveEccConversation(); }
function logCurrentEccRowAndOpenPowerSchool() { return archiveEccConversationAndOpenPowerSchool(); }
function archiveAndOpenPowerSchoolECC() { return archiveEccConversationAndOpenPowerSchool(); }
function archiveCurrentECCAttempt() { return archiveEccAttempt(); }
function logCurrentEccAttemptAndOpenPowerSchool() { return archiveEccAttemptAndOpenPowerSchool(); }

function applyEccDateColors() { return applyCallDueColors(); }
function preview10StudentRefresh() { return previewStudentRefresh(); }
function refreshStudentDataAndCounselors() { return refreshStudentData(); }
function show10AddStudents() { return showAddStudentsDialog(); }
function extend10RosterFilters() { return extendMasterFilter(); }
function install10RosterTools() { return installRosterTools(); }
function preview10GradeReportPhotos() { return previewGradeReportPhotos(); }
function import10GradeReportPhotos() { return importGradeReportPhotos(); }
