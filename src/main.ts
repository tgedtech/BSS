/**
 * Conway BSS: Robust Behavior System Script (Production-Ready)
 * 
 * - Per-team roster sync with execution-safe, incremental operation and last-sync tracking in Document Properties.
 * - Team tab auto-creation, infraction/event/email logging, robust UI/menu controls, developer/user mode.
 * - No duplicate event log entries; Source Value updated only if changed.
 * - "All Student Data" is only set for new entries.
 * - Bulletproof JSON parsing in scanAndSendAlerts.
 * - Per-user email requirement + attribution: teachers must confirm email, infractions stamped via cell note.
 * - Week protections: only the current week block (1st–5th) is editable; past/future weeks are locked. On‑edit guard reverts accidental edits outside the active block.
 * 
 * @author Fieldstone
 */

/* ------------------- GLOBAL CONFIG & UTILS ------------------- */
/** Returns true if the script is running in a context where UI is available */
function canUseUi(): boolean {
  try {
    SpreadsheetApp.getUi();
    return true;
  } catch (e) {
    return false;
  }
}

const FIRST_STUDENT_ROW = 3;
const TEAM_TEMPLATE_NAME = 'TeamTemplate';
const EXCLUDE_SHEETS = [
  'Variables',
  'Templates',
  'StudentList',
  'EmailLog',
  'DebugLog',
  'EventLog',
  TEAM_TEMPLATE_NAME,
];
const INFRACTION_SEQUENCE = ['1st', '2nd', '3rd', '4th', '5th'] as const;
const DEV_MENU_OPTIONS = [
  { name: '▶️ 1. Restore Formulas & Merge Tags (BSS Setup)', functionName: 'bssSetup' },
  { name: '▶️ 2. Log New Infractions to Event Log', functionName: 'processNewEvents' },
  { name: '▶️ 3. Send Alerts for 5th Infraction (Active Week Only)', functionName: 'scanAndSendAlerts' },
  { name: '▶️ 4. Sync Rosters (Incremental, per team)', functionName: 'syncRosterForAllTeams' },
  { separator: true },
  { name: '🔒 Refresh Week Locks (Only Current Week Editable)', functionName: 'refreshWeekLocksForAllTeams' },
  { name: '🧹 Reset All Team Tabs (EXCEPT Template)', functionName: 'resetTeamTabs' },
  { name: '🎨 Update Formatting from TeamTemplate', functionName: 'updateTeamFormatting' },
  { name: '🪵 View Debug Log Sheet', functionName: 'openDebugLogSheet' },
  { name: 'Sync Next Available Team', functionName: 'syncNextTeamForToday' },
  { name: 'Import Students from Source', functionName: 'importNeededStudentColumns' },
  { name: 'Build Student List from Import', functionName: 'buildStudentListFromImport' },
  { separator: true },
  { name: '❓ About Conway BSS...', functionName: 'showAboutDialog' },
  { separator: true },
  { name: 'Switch to User Mode', functionName: 'enableUserMode' },
];
const USER_MENU_OPTIONS = [{ name: 'Switch to Developer Mode', functionName: 'enableDeveloperMode' }];

function replaceAllTokens(source: string, search: string, replacement: string): string {
  return source.split(search).join(replacement);
}

/* -------- Per-user email (required) -------- */
function getSavedUserEmail_(): string | null {
  return PropertiesService.getUserProperties().getProperty('bssUserEmail') || null;
}
function saveUserEmail_(email: string): void {
  if (!email || !/@/.test(email)) throw new Error('Invalid email');
  PropertiesService.getUserProperties().setProperty('bssUserEmail', email.trim());
}
function guessUserEmail_(): string {
  try {
    const e = Session.getActiveUser().getEmail();
    return e && /@/.test(e) ? e : '';
  } catch (_) {
    return '';
  }
}
function promptForEmailSetup_(): string | null {
  if (!canUseUi()) return null;
  const ui = SpreadsheetApp.getUi();
  const guess = guessUserEmail_();
  const msg = guess
    ? 'We’ll use this to attribute infractions you enter.\n\nConfirm or change your email:'
    : 'We’ll use this to attribute infractions you enter.\n\nType your email:';
  const resp = ui.prompt(
    'Confirm your email',
    msg + (guess ? `\n\nDetected: ${guess}` : ''),
    ui.ButtonSet.OK_CANCEL
  );
  if (resp.getSelectedButton() !== ui.Button.OK) return null;
  const typed = resp.getResponseText().trim();
  const email = typed || guess;
  if (!/@/.test(email)) {
    ui.alert('Please enter a valid email address.');
    return promptForEmailSetup_();
  }
  saveUserEmail_(email);
  ui.alert('Thanks! Your email is set to: ' + email);
  return email;
}

/* ------------------- MENU & MODE LOGIC ------------------- */

function onOpen(): void {
  if (!canUseUi()) return;
  const props = PropertiesService.getDocumentProperties();
  const devMode = props.getProperty('developerMode') !== 'false'; // Default: true if not set

  const menu = SpreadsheetApp.getUi().createMenu('Conway BSS');
  const hasEmail = !!getSavedUserEmail_();

  if (!hasEmail && !devMode) {
    // Gate: only allow email setup until confirmed
    menu.addItem('Confirm My Email…', 'promptForEmailSetup_');
  } else {
    const options = devMode ? DEV_MENU_OPTIONS : USER_MENU_OPTIONS;
    options.forEach((opt) => {
      if ((opt as { separator?: boolean }).separator) menu.addSeparator();
      else menu.addItem(opt.name, opt.functionName);
    });
    menu.addSeparator();
    menu.addItem('Confirm My Email…', 'promptForEmailSetup_');
  }

  menu.addToUi();
  if (!devMode) hideAllExceptTeamTabs();

  // Gentle auto‑prompt if missing (non-blocking)
  if (!hasEmail && !devMode) {
    promptForEmailSetup_();
  }
}

function enableUserMode(): void {
  if (!canUseUi()) return;
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Switch to User Mode',
    'This will hide all sheets except team tabs and limit the menu to a single option for enabling Developer Mode. Continue?',
    ui.ButtonSet.YES_NO
  );
  if (response !== ui.Button.YES) return;
  PropertiesService.getDocumentProperties().setProperty('developerMode', 'false');
  hideAllExceptTeamTabs();
  SpreadsheetApp.getUi().alert('User Mode enabled! Only team tabs are visible and the menu is simplified.');
  onOpen();
}

function enableDeveloperMode(): void {
  if (!canUseUi()) return;
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Enable Developer Mode',
    'Developer mode will unhide all sheets and re-enable all BSS menu options.  You can break the application in developer mode! Continue?',
    ui.ButtonSet.YES_NO
  );
  if (response !== ui.Button.YES) return;
  PropertiesService.getDocumentProperties().setProperty('developerMode', 'true');
  showAllSheets();
  SpreadsheetApp.getUi().alert('Developer Mode enabled! All sheets and menu options are now available.');
  onOpen();
}

function hideAllExceptTeamTabs(): void {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const teams = getTeamNamesFromVariables();
  ss.getSheets().forEach((sheet) => {
    if (teams.includes(sheet.getName())) sheet.showSheet();
    else sheet.hideSheet();
  });
}

function showAllSheets(): void {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.getSheets().forEach((sheet) => sheet.showSheet());
}

/* ------------------- TEAM NAME LOOKUP ------------------- */

function getTeamNamesFromVariables(): string[] {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const variablesSheet = ss.getSheetByName('Variables');
  if (!variablesSheet) return [];
  const teams: string[] = [];
  let col = 2;
  let row = 3;
  while (true) {
    const value = variablesSheet.getRange(row, col).getValue();
    if (!value) break;
    teams.push(value.toString().trim());
    row++;
    if (row > 100) break;
  }
  return teams;
}

/* ------------------- WEEK PROTECTION HELPERS ------------------- */

type WeekColumnBlock = {
  startCol: number;
  endCol: number;
  label: string;
};

function getWeekColumnBlocks(sheet: GoogleAppsScript.Spreadsheet.Sheet): WeekColumnBlock[] {
  const width = sheet.getLastColumn();
  if (width < 5) return [];
  const headerData = sheet.getRange(1, 1, 2, width).getValues();
  const topRow = headerData[0];
  const secondRow = headerData[1];
  const blocks: WeekColumnBlock[] = [];

  for (let c = 0; c < secondRow.length; c++) {
    const subHeader = (secondRow[c] || '').toString().trim().toLowerCase();
    if (subHeader === '5th') {
      const endCol = c + 1;
      const startCol = Math.max(1, endCol - 4);
      const label = (topRow[c] || '').toString().trim();
      blocks.push({ startCol, endCol, label });
    }
  }

  return blocks;
}

function hidePastWeeksForSheet(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  activeBlock: { startCol: number; endCol: number } | null
): void {
  const weekBlocks = getWeekColumnBlocks(sheet);
  if (!weekBlocks.length) return;

  weekBlocks.forEach((block) => {
    sheet.showColumns(block.startCol, block.endCol - block.startCol + 1);
  });

  if (!activeBlock) return;

  weekBlocks
    .filter((block) => block.endCol < activeBlock.startCol)
    .forEach((block) => {
      sheet.hideColumns(block.startCol, block.endCol - block.startCol + 1);
    });
}

// Find the current week's 5-column block (1st..5th) for a given team sheet.
// Returns {startCol, endCol} in 1-based indexing, or null if not found.
function getActiveWeekBlockCols(sheet: GoogleAppsScript.Spreadsheet.Sheet):
  | { startCol: number; endCol: number }
  | null {
  const width = sheet.getLastColumn();
  if (width < 5) return null;
  const data = sheet.getRange(1, 1, 3, width).getValues(); // rows 1-3 headers
  const activeWeekStr = getActiveWeekForSheet(sheet); // e.g., "Week of Aug 11, 2025"
  if (!activeWeekStr) return null;

  const weekDate = activeWeekStr.replace(/^Week of\s+/i, '').trim();

  for (let c = 0; c < data[1].length; c++) {
    const subHeader = (data[1][c] || '').toString().trim(); // "1st..5th"
    const topHeader = (data[0][c] || '').toString().trim(); // "Aug 11, 2025"
    if (subHeader === '5th' && topHeader === weekDate) {
      const endCol = c + 1; // convert to 1-based
      const startCol = Math.max(1, endCol - 4);
      return { startCol, endCol };
    }
  }
  return null;
}

// Remove protections created by this script on a sheet
function removeOldWeekLocks_(sheet: GoogleAppsScript.Spreadsheet.Sheet): void {
  const tag = '[BSS] Lock past weeks';
  const protections = []
    .concat(sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET))
    .concat(sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE));
  protections.forEach((p) => {
    try {
      if (p && p.getDescription && p.getDescription() === tag) p.remove();
    } catch (_) {
      // ignore
    }
  });
}

// Apply protection so only the current week's block (student rows) is editable.
// strict=true => hard lock; false => warning-only.
function lockPastWeeksForSheet(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  strict = true
): void {
  if (EXCLUDE_SHEETS.includes(sheet.getName())) return;

  const cols = getActiveWeekBlockCols(sheet);
  if (!cols) {
    hidePastWeeksForSheet(sheet, null);
    debugLog(`[${sheet.getName()}] No active week block detected. Skipping protection.`);
    return;
  }

  removeOldWeekLocks_(sheet);

  const lastRow = sheet.getMaxRows();
  const editableRange = sheet.getRange(
    FIRST_STUDENT_ROW + 1,
    cols.startCol,
    Math.max(0, lastRow - FIRST_STUDENT_ROW),
    cols.endCol - cols.startCol + 1
  );

  const prot = sheet.protect();
  prot.setDescription('[BSS] Lock past weeks');
  prot.setWarningOnly(!strict);

  try {
    // Keep the owner; remove other editors so protection applies globally
    const me = Session.getEffectiveUser();
    prot.removeEditors(
      prot.getEditors().filter((ed) => ed.getEmail && ed.getEmail() !== me.getEmail())
    );
  } catch (_) {
    // ignore
  }

  prot.setUnprotectedRanges([editableRange]);

  debugLog(
    `[${sheet.getName()}] Week locks applied. Editable columns ${cols.startCol}-${cols.endCol} for student rows.`
  );

  hidePastWeeksForSheet(sheet, cols);
}

// Refresh week locks across all team tabs
function refreshWeekLocksForAllTeams(strict = true): void {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.getSheets().forEach((sheet) => {
    if (!EXCLUDE_SHEETS.includes(sheet.getName())) {
      lockPastWeeksForSheet(sheet, strict);
    }
  });
  debugLog(`Week locks refreshed for all team tabs. Strict=${strict}`);
}

/* ------------------- ROSTER SYNC (PER TEAM, INCREMENTAL) ------------------- */

/**
 * Syncs a single team's tab, preserves infractions, persists last sync date in Document Properties.
 * Can be called independently or in sequence for all teams (see syncRosterForAllTeams).
 */
function syncRosterForTeam(teamName: string): void {
  const props = PropertiesService.getDocumentProperties();
  let syncMap: Record<string, string> = {};
  const timers: Record<string, number> = {};
  function t(label: string): void {
    timers[label] = Date.now();
    debugLog(`[${new Date().toISOString()}][${teamName}] [Timer] ${label}...`);
  }
  function tElapsed(label: string, prev: string): void {
    const elapsed = ((Date.now() - timers[prev]) / 1000).toFixed(2);
    debugLog(`[${new Date().toISOString()}][${teamName}] [Timer] ${label}: +${elapsed}s`);
    timers[label] = Date.now();
  }

  try {
    syncMap = JSON.parse(props.getProperty('teamSyncMap') || '{}');
  } catch (e) {
    debugLog(`[${new Date().toISOString()}][${teamName}] ERROR parsing syncMap: ${e}`);
    syncMap = {};
  }
  const todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  if (syncMap[teamName] === todayStr) {
    debugLog(`[${new Date().toISOString()}][${teamName}] Already synced today; skipping.`);
    return;
  }

  debugLog(`[${new Date().toISOString()}][${teamName}] Starting sync...`);
  t('START');

  try {
    // Spreadsheet & Template
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const template = ss.getSheetByName(TEAM_TEMPLATE_NAME);
    if (!template) {
      debugLog(`[${new Date().toISOString()}][${teamName}] ERROR: TeamTemplate not found!`);
      if (canUseUi()) SpreadsheetApp.getUi().alert('Missing TeamTemplate tab!');
      return;
    }
    tElapsed('Template loaded', 'START');

    // Student Data
    const studentSheet = ss.getSheetByName('StudentList');
    const students = studentSheet ? studentSheet.getDataRange().getValues() : [];
    const studentHeaders = students[0] || [];
    tElapsed('StudentList loaded', 'Template loaded');

    // Event Log
    const eventSheet = ss.getSheetByName('EventLog');
    const events = eventSheet ? eventSheet.getDataRange().getValues() : [];
    const eventHeaders = events[0] || [];
    tElapsed('EventLog loaded', 'StudentList loaded');

    // Template headers
    const templateLastCol = template.getLastColumn();
    const templateHeaders = template.getRange(1, 1, 3, templateLastCol).getValues();
    tElapsed('Template headers pulled', 'EventLog loaded');

    // Build infraction map
    const infractionMap: Record<string, Set<string>> = {};
    if (events.length > 1) {
      const colDSID = eventHeaders.indexOf('DSID');
      const colInf = eventHeaders.indexOf('Infraction');
      const colWeek = eventHeaders.indexOf('Week');
      const colSourceValue = eventHeaders.indexOf('Source Value');
      for (let i = 1; i < events.length; i++) {
        const row = events[i];
        if (!row[colDSID] || !row[colInf] || !row[colWeek]) continue;
        const weekFormatted = formatWeekDate(row[colWeek]);
        const key = [row[colDSID], row[colInf], weekFormatted].join('|');
        if (!infractionMap[key]) infractionMap[key] = new Set();
        infractionMap[key].add((row[colSourceValue] || 'X') as string);
      }
    }
    tElapsed('Infraction map built', 'Template headers pulled');

    // Sheet creation/copy
    let sheet = ss.getSheetByName(teamName);
    if (!sheet) {
      debugLog(`[${new Date().toISOString()}][${teamName}] Sheet not found, copying template...`);
      sheet = template.copyTo(ss).setName(teamName);
      debugLog(`[${new Date().toISOString()}][${teamName}] New tab created from template.`);
    }
    template.hideSheet();
    tElapsed('Sheet prepared', 'Infraction map built');

    // Clear & header write
    const maxRows = Math.max(sheet.getMaxRows(), 100 + FIRST_STUDENT_ROW);
    const maxCols = Math.max(sheet.getMaxColumns(), templateLastCol);
    sheet.getRange(1, 1, maxRows, maxCols).clearContent();
    sheet.getRange(1, 1, 3, templateLastCol).setValues(templateHeaders);
    copySheetFormattingSafe(template, sheet, templateLastCol, 3);
    tElapsed('Sheet cleared & headers copied', 'Sheet prepared');

    // Filter/sort team students
    const teamStudents = students
      .filter((row, idx) => idx > 0 && row[studentHeaders.indexOf('Team')] === teamName)
      .sort((a, b) => {
        const lastA = (a[studentHeaders.indexOf('Last Name')] || '').toString().toLowerCase();
        const lastB = (b[studentHeaders.indexOf('Last Name')] || '').toString().toLowerCase();
        return lastA.localeCompare(lastB);
      });
    tElapsed('Students filtered/sorted', 'Sheet cleared & headers copied');

    // Build allRows (main logic)
    const allRows: (string | number | Date)[][] = [];
    for (let r = 0; r < teamStudents.length; r++) {
      const student = teamStudents[r];
      const dataRow: (string | number | Date)[] = new Array(templateLastCol).fill('');
      for (let c = 0; c < templateHeaders[2].length; c++) {
        const h = templateHeaders[2][c];
        if (!h) continue;
        const idx = studentHeaders.indexOf(h);
        if (idx !== -1) dataRow[c] = student[idx];
      }
      const dsid = student[studentHeaders.indexOf('DSID')];
      for (let col = 0; col < templateHeaders[2].length; col++) {
        const infractionLabel = (templateHeaders[1][col] || '').toString().trim();
        const normalizedLabel = infractionLabel.toLowerCase();
        if (INFRACTION_SEQUENCE.includes(normalizedLabel as typeof INFRACTION_SEQUENCE[number])) {
          const weekRaw = getWeekForInfractionCol(templateHeaders, col);
          const weekFormatted = formatWeekDate(weekRaw);
          const key = [dsid, infractionLabel, weekFormatted].join('|');
          if (infractionMap[key] && infractionMap[key].size > 0) {
            dataRow[col] = Array.from(infractionMap[key]).join('; ');
          }
        }
      }
      allRows.push(dataRow);
    }
    tElapsed('Roster rows built', 'Students filtered/sorted');

    // Write all rows to sheet
    if (allRows.length > 0) {
      sheet.getRange(FIRST_STUDENT_ROW + 1, 1, allRows.length, templateLastCol).setValues(allRows);
      tElapsed('Roster rows written', 'Roster rows built');
    } else {
      debugLog(`[${new Date().toISOString()}][${teamName}] No students found for roster, skipped data write.`);
    }

    const lastDataRow = FIRST_STUDENT_ROW + allRows.length;
    if (sheet.getLastRow() > lastDataRow) {
      sheet
        .getRange(lastDataRow + 1, 1, sheet.getLastRow() - lastDataRow, sheet.getLastColumn())
        .clearContent();
      debugLog(`[${new Date().toISOString()}][${teamName}] Cleared extra content below last student.`);
    }
    tElapsed('Cleared extra content', 'Roster rows written');

    sheet.getRange('D1').setValue(teamName);

    // Success! Record sync
    syncMap[teamName] = todayStr;
    props.setProperty('teamSyncMap', JSON.stringify(syncMap));
    tElapsed('SyncMap updated', 'Cleared extra content');

    const totalElapsed = ((Date.now() - timers['START']) / 1000).toFixed(2);
    debugLog(
      `[${new Date().toISOString()}][${teamName}] SUCCESS: Roster synced and timestamp updated in ${totalElapsed}s.`
    );
  } catch (err: unknown) {
    const message = err && typeof err === 'object' && 'message' in err ? (err as Error).message : String(err);
    debugLog(`[${new Date().toISOString()}][${teamName}] ERROR during sync: ${message}`);
  }

  pruneDebugLogSheet(1000); // Keep last 1000 entries
}

/**
 * Syncs rosters for ALL teams, one at a time (sequential).
 * Each call will skip teams already synced today.
 * Designed to be called repeatedly (e.g., from time-driven triggers) to avoid execution timeouts.
 */
function syncRosterForAllTeams(): void {
  const teams = getTeamNamesFromVariables();
  teams.forEach((team) => {
    syncRosterForTeam(team);
  });
  // After syncs, refresh locks so only the active week is editable
  refreshWeekLocksForAllTeams(true);
  if (canUseUi())
    SpreadsheetApp.getUi().alert('Roster sync complete for all teams (see DebugLog for per-team results).');
}

/**
 * (Admin Utility) Resets last sync status for all teams.
 * Use this to force full resync for all teams (e.g., after major data changes).
 */
function resetTeamSyncStatus(): void {
  PropertiesService.getDocumentProperties().deleteProperty('teamSyncMap');
  debugLog('All team sync timestamps reset.');
  if (canUseUi()) SpreadsheetApp.getUi().alert('All team sync timestamps have been reset.');
}

/* --------- OTHER CORE & UTILITY FUNCTIONS: UNCHANGED ----------- */

function copySheetFormattingSafe(
  template: GoogleAppsScript.Spreadsheet.Sheet,
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  colCount: number,
  rowCount: number
): void {
  const fromRange = template.getRange(1, 1, rowCount, colCount);
  const toRange = sheet.getRange(1, 1, rowCount, colCount);
  fromRange.copyTo(toRange, { formatOnly: true });
  for (let c = 1; c <= colCount; c++) {
    sheet.setColumnWidth(c, template.getColumnWidth(c));
  }
  for (let r = 1; r <= rowCount; r++) {
    sheet.setRowHeight(r, template.getRowHeight(r));
  }
}

/* --------- REQUIRE EMAIL + WEEK GUARD + TAG INFRACTIONS ON EDIT ----------- */
function onEdit(e: GoogleAppsScript.Events.SheetsOnEdit): void {
  try {
    if (!e || !e.range || !e.source) return;
    const sheet = e.range.getSheet();
    const sheetName = sheet.getName();
    if (EXCLUDE_SHEETS && EXCLUDE_SHEETS.indexOf(sheetName) >= 0) return;

    const row = e.range.getRow();
    const col = e.range.getColumn();

    // Only guard student rows (headers remain editable by admins)
    if (row <= FIRST_STUDENT_ROW) return;

    // --- Week guard: only allow edits inside the active week's 5-column block ---
    const cols = getActiveWeekBlockCols(sheet);
    if (cols) {
      const inActiveWeek = col >= cols.startCol && col <= cols.endCol;
      if (!inActiveWeek) {
        // Revert edit outside active week
        if (typeof e.oldValue !== 'undefined') {
          e.range.setValue(e.oldValue);
        } else {
          e.range.clearContent();
        }
        if (canUseUi()) {
          SpreadsheetApp.getActive().toast(
            'Edits are locked for past/future weeks. Please use the current week\'s columns only.',
            'BSS: Edit Blocked',
            5
          );
        }
        debugLog(`[${sheetName}] Blocked edit outside current week at R${row}C${col}.`);
        return;
      }
    }

    // Identify infraction columns by row 2 labels
    const label = (sheet.getRange(2, col).getValue() || '').toString().trim().toLowerCase();
    let labelIndex = INFRACTION_SEQUENCE.findIndex((entry) => entry === label);
    const isInfractionCol = labelIndex >= 0;
    if (!isInfractionCol) return; // Allow non-infraction edits within the active week

    const newVal = typeof e.value === 'undefined' ? sheet.getRange(row, col).getValue() : e.value;
    if (!newVal || newVal === '') return;

    const originalRange = e.range;
    let workingRange = originalRange;
    let workingColumn = col;
    let workingLabelIndex = labelIndex;

    let weekBlock = getWeekColumnBlocks(sheet).find(
      (block) => col >= block.startCol && col <= block.endCol
    );
    if (!weekBlock && cols) {
      weekBlock = { startCol: cols.startCol, endCol: cols.endCol, label: 'active' };
    }

    if (weekBlock) {
      const firstMissingIndex = INFRACTION_SEQUENCE.findIndex((_, idx) => {
        const requiredValue = sheet.getRange(row, weekBlock.startCol + idx).getValue();
        return requiredValue === '' || requiredValue === null;
      });

      if (firstMissingIndex !== -1 && labelIndex > firstMissingIndex) {
        const targetColumn = weekBlock.startCol + firstMissingIndex;
        const targetRange = sheet.getRange(row, targetColumn);
        if (typeof e.oldValue !== 'undefined') {
          originalRange.setValue(e.oldValue);
        } else {
          originalRange.clearContent();
        }
        targetRange.setValue(newVal);
        if (canUseUi()) {
          SpreadsheetApp.getActive().toast(
            `Value moved to the ${INFRACTION_SEQUENCE[firstMissingIndex]} infraction column for this week.`,
            'BSS: Sequence Adjusted',
            5
          );
        }
        debugLog(
          `[${sheetName}] Auto-moved entry at R${row}C${col} to ${INFRACTION_SEQUENCE[firstMissingIndex]} column (C${targetColumn}).`
        );
        workingRange = targetRange;
        workingColumn = targetColumn;
        workingLabelIndex = firstMissingIndex;
        labelIndex = firstMissingIndex;
      }

      if (workingLabelIndex > 0) {
        for (let i = 0; i < workingLabelIndex; i++) {
          const requiredCol = weekBlock.startCol + i;
          const requiredValue = sheet.getRange(row, requiredCol).getValue();
          if (requiredValue === '' || requiredValue === null) {
            if (workingRange === originalRange) {
              if (typeof e.oldValue !== 'undefined') {
                workingRange.setValue(e.oldValue);
              } else {
                workingRange.clearContent();
              }
            } else {
              workingRange.clearContent();
            }
            if (canUseUi()) {
              SpreadsheetApp.getActive().toast(
                `Enter the ${INFRACTION_SEQUENCE[i]} infraction before recording the ${
                  INFRACTION_SEQUENCE[workingLabelIndex]
                }.`,
                'BSS: Sequence Required',
                5
              );
            }
            debugLog(
              `[${sheetName}] Blocked ${INFRACTION_SEQUENCE[workingLabelIndex]} entry at R${row}C${workingColumn}: missing ${INFRACTION_SEQUENCE[i]}.`
            );
            return;
          }
        }
      }
    }

    // Require previously-confirmed email (do NOT prompt here—simple trigger)
    const email = getSavedUserEmail_();
    if (!email) {
      if (workingRange === originalRange) {
        if (typeof e.oldValue !== 'undefined') {
          workingRange.setValue(e.oldValue);
        } else {
          workingRange.clearContent();
        }
      } else {
        workingRange.clearContent();
      }
      originalRange.setNote('Edit blocked: confirm your email via menu “Conway BSS → Confirm My Email…”.');
      return;
    }

    // Stamp note with author + timestamp (most recent on top)
    const stamp = 'Entered by: ' + email + ' | ' + new Date().toLocaleString();
    const prevNote = workingRange.getNote() || '';
    const note = stamp + (prevNote ? '\n' + prevNote : '');
    workingRange.setNote(note);
  } catch (err) {
    try {
      Logger.log('onEdit error: ' + err);
    } catch (_) {
      // ignore
    }
  }
}

function processNewEvents(): void {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let eventSheet = ss.getSheetByName('EventLog');
  // Create EventLog if not found
  if (!eventSheet) {
    eventSheet = ss.insertSheet('EventLog');
    eventSheet.appendRow([
      'Timestamp',
      'DSID',
      'Last Name',
      'First Name',
      'Grade',
      'Team',
      'Infraction',
      'Week',
      'Entered By',
      'Sheet Name',
      'Source Value',
      'All Student Data',
      'Alerted',
    ]);
  }
  const eventHeaders = eventSheet.getDataRange().getValues()[0];
  const eventCol: Record<string, number> = {};
  eventHeaders.forEach((h, i) => {
    eventCol[h] = i;
  });

  const logged = eventSheet.getDataRange().getValues();
  // Map: DSID|Infraction|Week -> rowIndex (1-based for sheet)
  const eventMap: Record<string, number> = {};
  for (let i = 1; i < logged.length; i++) {
    const row = logged[i];
    const weekFormatted = formatWeekDate(row[eventCol['Week']]);
    const key = [row[eventCol['DSID']], row[eventCol['Infraction']], weekFormatted].join('|');
    eventMap[key] = i + 1; // +1 because of header
  }

  const studentSheet = ss.getSheetByName('StudentList');
  const students = studentSheet ? studentSheet.getDataRange().getValues() : [];
  const studentHeaders = students[0] || [];
  let eventsLogged = 0;
  const newEvents: (string | number | Date)[][] = [];
  const sheets = ss.getSheets();

  sheets.forEach((sheet) => {
    if (EXCLUDE_SHEETS.includes(sheet.getName())) return;
    const data = sheet.getDataRange().getValues();
    if (data.length < FIRST_STUDENT_ROW + 1) return;
    const sheetHeaders = data[2];
    const dsidIdx = sheetHeaders.indexOf('DSID');
    if (dsidIdx === -1) return;

    const infractionCols: number[] = [];
    for (let c = 0; c < sheetHeaders.length; c++) {
      const labelValue = (data[1][c] || '').toString().trim().toLowerCase();
      if (INFRACTION_SEQUENCE.includes(labelValue as typeof INFRACTION_SEQUENCE[number])) {
        infractionCols.push(c);
      }
    }

    for (let r = FIRST_STUDENT_ROW; r < data.length; r++) {
      const studentRow = data[r];
      const dsid = studentRow[dsidIdx];
      if (!dsid) continue;

      const studentData: Record<string, unknown> = {};
      let foundStudentList = false;
      for (let i = 1; i < students.length; i++) {
        if (students[i][studentHeaders.indexOf('DSID')] === dsid) {
          studentHeaders.forEach((key, idx) => {
            studentData[key] = students[i][idx];
          });
          foundStudentList = true;
          break;
        }
      }
      if (!foundStudentList) {
        sheetHeaders.forEach((key, idx) => {
          studentData[key] = studentRow[idx];
        });
      }

      for (const infCol of infractionCols) {
        const value = studentRow[infCol];
        if (!value || value === '') continue;
        const infractionNum = data[1][infCol];

        // Week key
        const weekRaw = getWeekForInfractionCol(data, infCol);
        const week = formatWeekDate(weekRaw);
        const key = [dsid, infractionNum, week].join('|');

        // Pull author from cell note (first line)
        const cell = sheet.getRange(r + 1, infCol + 1); // convert to 1-based
        const note = cell.getNote() || '';
        let enteredBy = 'unknown';
        if (note) {
          const firstLine = note.split('\n')[0];
          const m = firstLine.match(/Entered by:\s*([^|]+)\s*\|/i);
          if (m && m[1] && /@/.test(m[1].trim())) enteredBy = m[1].trim();
        }

        if (eventMap[key]) {
          // Only update Source Value if changed, NEVER overwrite All Student Data
          const rowIdx = eventMap[key];
          const rowData = eventSheet
            .getRange(rowIdx, 1, 1, eventHeaders.length)
            .getValues()[0];
          const existingValue = rowData[eventCol['Source Value']];
          if (existingValue !== value) {
            eventSheet.getRange(rowIdx, eventCol['Source Value'] + 1).setValue(value);
          }
          // Do not duplicate
          continue;
        }

        newEvents.push([
          new Date(),
          dsid,
          studentData['Last Name'],
          studentData['First Name'],
          studentData['Grade'],
          studentData['Team'],
          infractionNum,
          week, // formatted
          enteredBy,
          sheet.getName(),
          value,
          JSON.stringify(studentData),
          '',
        ]);
        eventMap[key] = eventSheet.getLastRow() + newEvents.length;
        eventsLogged++;
      }
    }
  });

  if (newEvents.length > 0) {
    eventSheet
      .getRange(eventSheet.getLastRow() + 1, 1, newEvents.length, eventHeaders.length)
      .setValues(newEvents);
  }
  if (typeof canUseUi !== 'undefined' && canUseUi())
    SpreadsheetApp.getUi().alert(`Event log complete. ${eventsLogged} new events recorded.`);
}

function scanAndSendAlerts(): void {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const eventSheet = ss.getSheetByName('EventLog');
  const logSheet = ss.getSheetByName('EmailLog');
  const templateSheet = ss.getSheetByName('Templates');
  const variablesSheet = ss.getSheetByName('Variables');
  if (!eventSheet) {
    debugLog('No EventLog sheet found, aborting scanAndSendAlerts.');
    return;
  }
  const events = eventSheet.getDataRange().getValues();
  const headers = events[0];
  let alertsSent = 0;
  debugLog(`Starting scanAndSendAlerts. Total events: ${events.length - 1}`);

  // --- Helper: normalize week string to consistent 'Aug 11, 2025' format ---
  function normalizeWeekStr(weekStr: unknown): string {
    if (!weekStr) return '';
    if (weekStr instanceof Date) {
      return Utilities.formatDate(weekStr, Session.getScriptTimeZone(), 'MMM d, yyyy');
    }
    const str = weekStr.toString();
    const match = str.match(/([A-Za-z]{3} \d{1,2}, \d{4})/);
    if (match) return match[1];
    const date = new Date(str);
    if (!isNaN(date.getTime())) {
      return Utilities.formatDate(date, Session.getScriptTimeZone(), 'MMM d, yyyy');
    }
    return str;
  }

  // --- Get active week for each team tab, normalized ---
  const teamActiveWeeks: Record<string, string> = {};
  SpreadsheetApp.getActiveSpreadsheet()
    .getSheets()
    .forEach((sheet) => {
      if (!EXCLUDE_SHEETS.includes(sheet.getName())) {
        const activeWeekRaw = getActiveWeekForSheet(sheet);
        const activeWeek = normalizeWeekStr(activeWeekRaw);
        if (activeWeek) teamActiveWeeks[sheet.getName()] = activeWeek;
      }
    });

  for (let i = 1; i < events.length; i++) {
    const event = events[i];
    const dsid = event[headers.indexOf('DSID')];
    const infraction = event[headers.indexOf('Infraction')];
    const weekRaw = event[headers.indexOf('Week')];
    const week = normalizeWeekStr(weekRaw);
    const sheetName = event[headers.indexOf('Sheet Name')];
    const alerted = event[headers.indexOf('Alerted')];

    debugLog(
      `Row ${i + 1} | DSID: ${dsid} | Infraction: ${infraction} | Week: ${weekRaw} (normalized: ${week}) | Sheet: ${sheetName} | Alerted: ${!!alerted}`
    );

    if (infraction !== '5th') continue;
    if (alerted) continue;

    const activeWeek = teamActiveWeeks[sheetName];
    if (!activeWeek || week !== activeWeek) {
      debugLog(`Row ${i + 1}: SKIP, activeWeek='${activeWeek}', eventWeek='${week}'`);
      continue;
    }

    const repeatCount = events.filter(
      (ev) => ev[headers.indexOf('DSID')] === dsid && ev[headers.indexOf('Infraction')] === '5th'
    ).length;
    const totalInfractions = events.filter((ev) => ev[headers.indexOf('DSID')] === dsid).length;
    const templateName = '5th Infraction Notice';
    if (!templateSheet) {
      debugLog('scanAndSendAlerts: Templates sheet not found.');
      continue;
    }
    const template = getTemplateByName(templateSheet, templateName);

    let emailSubject = template.subject;
    let emailBody = template.body;
    let student: Record<string, unknown> = {};
    try {
      const studentDataRaw = event[headers.indexOf('All Student Data')];
      if (typeof studentDataRaw === 'string' && studentDataRaw.trim().length > 0) {
        try {
          student = JSON.parse(studentDataRaw);
        } catch (e) {
          debugLog(`Row ${i + 1}: ERROR parsing student data: ${e}`);
          student = {};
        }
      } else if (typeof studentDataRaw === 'object' && studentDataRaw !== null) {
        student = studentDataRaw as Record<string, unknown>;
      } else {
        student = {};
      }
    } catch (e) {
      debugLog(`Row ${i + 1}: ERROR parsing student data: ${e}`);
      student = {};
    }
    const mergeFields: Record<string, unknown> = {
      DSID: dsid,
      'Last Name': event[headers.indexOf('Last Name')],
      'First Name': event[headers.indexOf('First Name')],
      Grade: event[headers.indexOf('Grade')],
      Team: event[headers.indexOf('Team')],
      Infraction: infraction,
      Week: week,
      'Sheet Name': sheetName,
      'Entered By': event[headers.indexOf('Entered By')],
      'Repeat Count': repeatCount,
      'Total Infractions': totalInfractions,
      ...student,
    };
    const recipientEmails = variablesSheet ? getAlertRecipients(variablesSheet, mergeFields['Team'] as string) : [];
    mergeFields['Recipients'] = recipientEmails.join(', ');
    (template.mergeFields.length ? template.mergeFields : Object.keys(mergeFields)).forEach((k) => {
      const value = mergeFields[k];
      const replacement = value === undefined || value === null ? '' : String(value);
      emailSubject = replaceAllTokens(emailSubject, `{{${k}}}`, replacement);
      emailBody = replaceAllTokens(emailBody, `{{${k}}}`, replacement);
    });
    const recipientsNotice = `<b>This alert has been sent to:</b> <br>${recipientEmails.join(', ')}<br><br>`;
    emailBody = recipientsNotice + emailBody;
    const prefillLink = buildPrefillLink(mergeFields);
    emailBody += `<br><b>Counselor follow-up required:</b> <a href="${prefillLink}" target="_blank">Click here to confirm student meeting</a><br>
 <span style="color:#B22222;"><b>Do NOT enter counseling notes or sensitive info here—this is for process tracking only.</b></span><br>`;
    if (recipientEmails.length > 0) {
      try {
        MailApp.sendEmail({
          to: recipientEmails.join(','),
          subject: emailSubject,
          htmlBody: emailBody,
          noReply: true,
        });
        debugLog(
          `Row ${i + 1}: Email sent to: ${recipientEmails.join(', ')} | Subject: ${emailSubject}`
        );
        alertsSent++;
      } catch (e) {
        debugLog(`Row ${i + 1}: ERROR sending email: ${e}`);
      }
    }
    eventSheet.getRange(i + 1, headers.indexOf('Alerted') + 1).setValue(new Date());
    if (logSheet) {
      const logHeaders = logSheet.getDataRange().getValues()[0];
      const summary =
        `ALERT: ${mergeFields['First Name']} ${mergeFields['Last Name']} (Grade ${mergeFields['Grade']}, ${mergeFields['Team']}) | ` +
        `5th infraction for ${mergeFields['Week']} | Repeats: ${repeatCount}, Total: ${totalInfractions} | ` +
        `Recipients: ${recipientEmails.join(', ')}`;
      const logEntry: (string | number | Date)[] = [];
      for (let j = 0; j < logHeaders.length; j++) {
        switch (logHeaders[j]) {
          case 'Timestamp':
            logEntry.push(new Date());
            break;
          case 'DSID':
            logEntry.push(dsid);
            break;
          case 'Last Name':
            logEntry.push(event[headers.indexOf('Last Name')]);
            break;
          case 'First Name':
            logEntry.push(event[headers.indexOf('First Name')]);
            break;
          case 'Grade':
            logEntry.push(event[headers.indexOf('Grade')]);
            break;
          case 'Team':
            logEntry.push(event[headers.indexOf('Team')]);
            break;
          case 'Infraction':
            logEntry.push(infraction);
            break;
          case 'Week':
            logEntry.push(week);
            break;
          case 'Sheet Name':
            logEntry.push(sheetName);
            break;
          case 'Entered By':
            logEntry.push(event[headers.indexOf('Entered By')]);
            break;
          case 'Body':
            logEntry.push(summary);
            break;
          case 'Recipients':
            logEntry.push(recipientEmails.join(', '));
            break;
          default:
            logEntry.push('');
            break;
        }
      }
      logSheet.appendRow(logEntry);
      debugLog(`Row ${i + 1}: Logged in EmailLog.`);
    }
  }
  debugLog(`scanAndSendAlerts: Completed. ${alertsSent} alerts sent.`);
  if (typeof canUseUi !== 'undefined' && canUseUi())
    SpreadsheetApp.getUi().alert(`Scan complete. ${alertsSent} new alerts sent.`);
}

function getWeekForInfractionCol(headers: unknown[][], col: number): unknown {
  const infractionLabel = (headers[1][col] || '').toString().trim();
  if (infractionLabel === '5th') {
    const weekDate = (headers[0][col] || '').toString().trim();
    return weekDate ? 'Week of ' + weekDate : 'Unknown Week';
  } else {
    for (let c = col + 1; c < headers[1].length; c++) {
      const label = (headers[1][c] || '').toString().trim();
      if (label === '5th') {
        const weekDate = (headers[0][c] || '').toString().trim();
        return weekDate ? 'Week of ' + weekDate : 'Unknown Week';
      }
    }
  }
  return 'Unknown Week';
}

function getActiveWeekForSheet(sheet: GoogleAppsScript.Spreadsheet.Sheet): string | null {
  const data = sheet.getDataRange().getValues();
  for (let c = 0; c < data[0].length; c++) {
    if ((data[0][c] || '').toString().trim().toLowerCase() === 'active') {
      for (let cc = c + 1; cc < data[1].length; cc++) {
        if ((data[1][cc] || '').toString().trim() === '5th') {
          const weekDate = (data[0][cc] || '').toString().trim();
          return weekDate ? 'Week of ' + weekDate : null;
        }
      }
    }
  }
  return null;
}

function bssSetup(): void {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  const templateSheet = ss.getSheetByName('Templates');
  if (!templateSheet) {
    debugLog('bssSetup: Templates sheet not found.');
    return;
  }
  const templateData = templateSheet.getDataRange().getValues();
  const tagSet = new Set<string>([
    'DSID',
    'SSID',
    'Last Name',
    'First Name',
    'Grade',
    'Team',
    'Gender',
    'Ethnicity',
    'ML',
    'ECE',
    'Week',
    'Repeat Count',
    'Total Infractions',
    'Sheet Name',
    'Recipients',
  ]);
  for (let col = 1; col < templateData[0].length; col++) {
    for (let r = 1; r <= 2; r++) {
      const cell = templateData[r][col];
      if (cell) {
        const matches = Array.from(cell.toString().matchAll(/{{(.*?)}}/g));
        for (const m of matches) tagSet.add(m[1].trim());
      }
    }
  }
  templateSheet.getRange(4, 1).setValue('Merge Fields');
  for (let col = 2; col <= templateData[0].length; col++) {
    templateSheet.getRange(4, col).setValue(Array.from(tagSet).join(', '));
  }
  sheets.forEach((sheet) => {
    if (EXCLUDE_SHEETS.includes(sheet.getName())) return;
    const data = sheet.getDataRange().getValues();
    if (data.length < FIRST_STUDENT_ROW + 1) return;
    const headers = data[2];
    let idCol = headers.indexOf('DSID');
    if (idCol === -1) idCol = headers.indexOf('Last Name');
    if (idCol === -1) return;
    let lastRow = data.length;
    for (let r = FIRST_STUDENT_ROW; r < data.length; r++) {
      if (!data[r][idCol] || data[r][idCol] === '') {
        lastRow = r;
        break;
      }
    }
    for (let r = FIRST_STUDENT_ROW + 1; r <= lastRow; r++) {
      const formula = `=IF(INDEX($${r}:$${r}, 1, MATCH("Active", $1:$1, 0)+4)<>"", "Needs Email", "")`;
      sheet.getRange(r, 1).setFormula(formula);
    }
  });
  if (canUseUi()) SpreadsheetApp.getUi().alert('BSS Setup complete: Formulas and merge tags updated.');
}

function debugLog(message: string): void {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let logSheet = ss.getSheetByName('DebugLog');
  if (!logSheet) logSheet = ss.insertSheet('DebugLog');
  logSheet.appendRow([new Date(), message]);
  Logger.log(message);
}

function getAlertRecipients(
  variablesSheet: GoogleAppsScript.Spreadsheet.Sheet,
  teamName: string
): string[] {
  const values = variablesSheet.getDataRange().getValues();
  let teamLeaderEmail = '';
  for (let i = 1; i < values.length; i++) {
    const rowTeam = values[i][1] ? values[i][1].toString().trim().toLowerCase() : '';
    if (rowTeam !== '' && rowTeam === teamName.toString().trim().toLowerCase()) {
      teamLeaderEmail = values[i][2] ? values[i][2].toString().trim() : '';
      break;
    }
  }
  debugLog(`getAlertRecipients: Team: "${teamName}" - Team Leader: "${teamLeaderEmail}"`);
  const primaryResponderEmails: string[] = [];
  let srHeaderRow = -1;
  let srEmailCol = -1;
  let srRoleCol = -1;
  for (let r = 0; r < values.length; r++) {
    for (let c = 0; c < values[r].length; c++) {
      if (values[r][c] && values[r][c].toString().trim() === 'Email') {
        srHeaderRow = r;
        srEmailCol = c;
      }
      if (values[r][c] && values[r][c].toString().trim() === 'Role') {
        srRoleCol = c;
      }
    }
    if (srHeaderRow !== -1 && srEmailCol !== -1 && srRoleCol !== -1) break;
  }
  if (srHeaderRow !== -1 && srEmailCol !== -1 && srRoleCol !== -1) {
    for (let r = srHeaderRow + 1; r < values.length; r++) {
      const role = values[r][srRoleCol] ? values[r][srRoleCol].toString().trim() : '';
      const email = values[r][srEmailCol] ? values[r][srEmailCol].toString().trim() : '';
      if (role === 'Primary Responder' && email.includes('@')) primaryResponderEmails.push(email);
    }
  }
  debugLog(`getAlertRecipients: Primary Responders: ${primaryResponderEmails.join(', ')}`);
  const allEmails = primaryResponderEmails.slice();
  if (teamLeaderEmail && !allEmails.includes(teamLeaderEmail)) {
    allEmails.push(teamLeaderEmail);
  }
  debugLog(`getAlertRecipients: ALL recipients: ${allEmails.join(', ')}`);
  return allEmails.filter((e, idx) => e && allEmails.indexOf(e) === idx);
}

function stripHtml(html: string): string {
  if (!html) return '';
  return html
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<[^>]*>/g, '')
    .replace(/&nbsp;/g, ' ')
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'");
}

function getTemplateByName(
  templateSheet: GoogleAppsScript.Spreadsheet.Sheet,
  templateName: string
): { subject: string; body: string; mergeFields: string[] } {
  const values = templateSheet.getDataRange().getValues();
  const colIndex = values[0].indexOf(templateName);
  if (colIndex === -1) throw new Error('Template not found: ' + templateName);

  const fieldToRow: Record<string, number> = {};
  for (let r = 0; r < values.length; r++) {
    fieldToRow[values[r][0]] = r;
  }
  let mergeFields: string[] = [];
  if (fieldToRow['Merge Fields'] !== undefined) {
    mergeFields = (values[fieldToRow['Merge Fields']][colIndex] || '')
      .toString()
      .split(',')
      .map((s) => s.replace(/[{}]/g, '').trim())
      .filter((s) => s.length > 0);
  }
  return {
    subject: values[fieldToRow['Subject']][colIndex],
    body: values[fieldToRow['Email Body']][colIndex],
    mergeFields,
  };
}

function openDebugLogSheet(): void {
  if (!canUseUi()) return;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let logSheet = ss.getSheetByName('DebugLog');
  if (!logSheet) logSheet = ss.insertSheet('DebugLog');
  ss.setActiveSheet(logSheet);
}

function showAboutDialog(): void {
  if (!canUseUi()) return;
  const message =
    'Conway Behavior Support System\n' +
    '\n' +
    'Track infractions, sync rosters, and notify teams with guardrails that keep weekly data accurate.\n' +
    '\n' +
    'Need help or have feedback? Contact your Conway administrator or the FS support team.\n' +
    '\n' +
    '© 2025 FS';

  SpreadsheetApp.getUi().alert('About Conway BSS', message, SpreadsheetApp.getUi().ButtonSet.OK);
}

function buildPrefillLink(mergeFields: Record<string, unknown>): string {
  return (
    'https://docs.google.com/forms/d/e/1FAIpQLScIuVFw1ZAI32ewW0BomMHK7bltqkKPyK5Wy37KOr46A0_8gA/viewform?usp=pp_url' +
    '&entry.1580995247=' +
    encodeURIComponent(`${mergeFields['First Name']} ${mergeFields['Last Name']}`) +
    '&entry.806184538=' +
    encodeURIComponent(String(mergeFields['DSID'] ?? '')) +
    '&entry.778985363=' +
    encodeURIComponent(String(mergeFields['Grade'] ?? '')) +
    '&entry.2029453213=' +
    encodeURIComponent(String(mergeFields['Team'] ?? '')) +
    '&entry.1429204745=' +
    encodeURIComponent(String(mergeFields['Week'] ?? ''))
  );
}

function formatWeekDate(dateVal: unknown): string {
  let d: Date;
  if (dateVal instanceof Date) {
    d = dateVal;
  } else if (typeof dateVal === 'string' && dateVal.trim() !== '') {
    d = new Date(dateVal);
    if (isNaN(d.getTime())) return dateVal;
  } else {
    return '';
  }
  const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  return `${months[d.getMonth()]} ${d.getDate()}, ${d.getFullYear()}`;
}

function syncNextTeamForToday(): void {
  const props = PropertiesService.getDocumentProperties();
  let syncMap: Record<string, string> = {};
  try {
    syncMap = JSON.parse(props.getProperty('teamSyncMap') || '{}');
  } catch (e) {
    syncMap = {};
  }
  const todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const teams = getTeamNamesFromVariables();
  for (const team of teams) {
    if (syncMap[team] !== todayStr) {
      debugLog(`[${new Date().toISOString()}][${team}] Selected as next team to sync.`);
      syncRosterForTeam(team); // Only sync the FIRST unsynced team this run
      return;
    }
  }
  debugLog('All teams have been synced for today.');
}

function pruneDebugLogSheet(maxRows = 1000): void {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName('DebugLog');
  if (!logSheet) return;
  const lastRow = logSheet.getLastRow();
  if (lastRow > maxRows + 1) {
    logSheet.deleteRows(2, lastRow - maxRows - 1);
    Logger.log('DebugLog pruned to last ' + maxRows + ' entries.');
  }
}

/** Import Only Needed Student Columns for BSS */
function importNeededStudentColumns(): void {
  const sourceSheetId = '1G41kcEBMTtfpLnzfyVYMZ_DxwwWQis2p_J_c0vqhJKI';
  const sourceTabName = 'Raw';
  const destTabName = 'ImportList';
  const neededHeaders = [
    'SSID',
    'StudentID',
    'First name',
    'Last name',
    'Gender',
    'Race',
    'CourseName',
    'GP',
    'Period',
    'TeacherName',
    'ELL',
    'Team',
    'Grade',
    'activeEnrollment.specialEdStatus',
  ];

  // === OPEN SHEETS ===
  let sourceSheet: GoogleAppsScript.Spreadsheet.Sheet | null;
  let destSheet: GoogleAppsScript.Spreadsheet.Sheet;
  try {
    sourceSheet = SpreadsheetApp.openById(sourceSheetId).getSheetByName(sourceTabName);
    if (!sourceSheet) throw new Error(`Source tab "${sourceTabName}" not found!`);
  } catch (e) {
    throw new Error(`Cannot open source: ${e}`);
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  destSheet = ss.getSheetByName(destTabName) || ss.insertSheet(destTabName);
  destSheet.clearContents();

  // === GET DATA ===
  const data = sourceSheet.getDataRange().getValues();
  const headers = data[0];

  // Build header → index mapping
  const headerMap: Record<string, number> = {};
  headers.forEach((h, idx) => {
    headerMap[h] = idx;
  });

  // Find indices of only the needed columns
  const colIndices = neededHeaders.map((h) => {
    if (headerMap[h] === undefined) throw new Error(`Header "${h}" not found in source!`);
    return headerMap[h];
  });

  // Build new data array: headers row + all data rows with only the needed columns
  const out: unknown[][] = [];
  out.push(neededHeaders);
  for (let r = 1; r < data.length; r++) {
    out.push(colIndices.map((idx) => data[r][idx]));
  }

  destSheet.getRange(1, 1, out.length, out[0].length).setValues(out);

  Logger.log(`Imported ${out.length - 1} students × ${out[0].length} columns to "${destTabName}".`);
}

function normalizeName(name: string): string {
  if (!name) return '';
  const base = name.split(',')[0].trim();
  if (!base) return '';
  return base
    .toLowerCase()
    .replace(/\b\w/g, (l) => l.toUpperCase());
}

function buildStudentListFromImport(): void {
  const sourceTabName = 'ImportList';
  const destTabName = 'StudentList';
  const debugLogTab = 'DebugLog';
  const teamMapRange = 'v_TeamsImport';
  const teamOutRange = 'v_TeamsOutput';
  const raceMapRange = 'v_raceCode';
  const raceOutRange = 'v_raceDisplay';
  const mlMapRange = 'v_mlImport';
  const mlOutRange = 'v_mlInterpretation';
  const spedMapRange = 'v_SPEDCode';
  const spedOutRange = 'v_SPEDDisplay';

  const destHeaders = [
    'DSID',
    'SSID',
    'Last Name',
    'First Name',
    'Advisory',
    'Grade',
    'Team',
    'Gender',
    'Race',
    'ML',
    'ECE',
  ];

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const importSheet = ss.getSheetByName(sourceTabName);
  if (!importSheet) {
    debugLog('buildStudentListFromImport: ImportList sheet not found.');
    return;
  }
  const destSheet = ss.getSheetByName(destTabName) || ss.insertSheet(destTabName);
  const debugSheet = ss.getSheetByName(debugLogTab) || ss.insertSheet(debugLogTab);

  function logException(msg: string, obj: unknown): void {
    debugSheet.appendRow([new Date(), msg, JSON.stringify(obj)]);
  }

  const data = importSheet.getDataRange().getValues();
  const headers = data[0];
  const col: Record<string, number> = {};
  headers.forEach((h, i) => {
    col[h.trim()] = i;
  });

  function mappingDict(range1: string, range2: string): Record<string, string> {
    const mapVals = ss.getRangeByName(range1).getValues().flat();
    const outVals = ss.getRangeByName(range2).getValues().flat();
    const map: Record<string, string> = {};
    for (let i = 0; i < mapVals.length; i++) map[mapVals[i]] = outVals[i];
    return map;
  }
  const teamMap = mappingDict(teamMapRange, teamOutRange);
  const raceMap = mappingDict(raceMapRange, raceOutRange);
  const mlMap = mappingDict(mlMapRange, mlOutRange);
  const spedMap = mappingDict(spedMapRange, spedOutRange);

  // Build lookup: {StudentID: [all rows for this student, GP==1]}
  const byStudentID: Record<string, any[][]> = {};
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[col['GP']] !== '1') continue;
    const studentID = String(row[col['StudentID']]);
    if (!byStudentID[studentID]) byStudentID[studentID] = [];
    byStudentID[studentID].push(row);
  }

  const out: (string | number)[][] = [destHeaders];
  let perfectMatch = 0;
  let fallback = 0;
  let assignedNoTeam = 0;
  let skipped = 0;
  const skippedList: string[] = [];

  for (const studentID in byStudentID) {
    const rows = byStudentID[studentID];
    // 1. Try to find Period == "A" and non-blank Team and a real teacher
    let target = rows.find(
      (row) =>
        row[col['Period']] === 'A' &&
        row[col['Team']] &&
        row[col['Team']].toString().trim() &&
        row[col['TeacherName']] &&
        !row[col['TeacherName']].toString().toLowerCase().includes('vacancy')
    );
    let matchType: 'perfect' | 'fallback' | 'no_team' = 'perfect';

    // 2. Try any period A with non-vacancy teacher (still "perfect")
    if (!target) {
      const candidates = rows.filter(
        (row) =>
          row[col['Period']] === 'A' &&
          row[col['Team']] &&
          row[col['Team']].toString().trim()
      );
      const notVacancy = candidates.find(
        (row) => row[col['TeacherName']] && !row[col['TeacherName']].toString().toLowerCase().includes('vacancy')
      );
      if (notVacancy) {
        target = notVacancy;
        matchType = 'perfect';
      }
    }

    // 3. Fallback: first row with non-blank Team (any period)
    if (!target) {
      target = rows.find((row) => row[col['Team']] && row[col['Team']].toString().trim());
      matchType = 'fallback';
    }

    // 4. If still not found, assign "No Team"
    if (!target) {
      target = rows[0];
      matchType = 'no_team';
    }

    if (!target) {
      skipped++;
      skippedList.push(studentID);
      logException('No usable row found for student (should not occur)', {
        StudentID: studentID,
        AllPeriods: rows.map((r) => r[col['Period']]).join(','),
      });
      continue;
    }

    const teamCode = String(target[col['Team']] ?? '');
    const raceCode = String(target[col['Race']] ?? '');
    const mlCode = String(target[col['ELL']] ?? '');
    const spedCode = String(target[col['activeEnrollment.specialEdStatus']] ?? '');
    const advisoryName = normalizeName(String(target[col['TeacherName']] ?? ''));
    const teamDisplay = matchType === 'no_team' ? 'No Team' : teamMap[teamCode] || teamCode;

    const outputRow = [
      target[col['StudentID']],
      target[col['SSID']],
      target[col['Last name']],
      target[col['First name']],
      advisoryName,
      target[col['Grade']],
      teamDisplay,
      target[col['Gender']],
      raceMap[raceCode] || raceCode,
      mlMap[mlCode] || mlCode,
      spedMap[spedCode] || spedCode,
    ];
    out.push(outputRow);

    if (matchType === 'perfect') {
      perfectMatch++;
    } else if (matchType === 'fallback') {
      fallback++;
      logException('Used fallback row for student (no period A with team/teacher, used first with team)', {
        StudentID: studentID,
        SSID: target[col['SSID']],
        UsedPeriod: target[col['Period']],
        Team: target[col['Team']],
        Advisory: target[col['TeacherName']],
        AllTeams: rows.map((r) => r[col['Team']]).join(','),
        AllPeriods: rows.map((r) => r[col['Period']]).join(','),
      });
    } else if (matchType === 'no_team') {
      assignedNoTeam++;
      logException('Assigned No Team for student (never found team)', {
        StudentID: studentID,
        SSID: target[col['SSID']],
        Advisory: target[col['TeacherName']],
        AllTeams: rows.map((r) => r[col['Team']]).join(','),
        AllPeriods: rows.map((r) => r[col['Period']]).join(','),
      });
    }
  }

  destSheet.clearContents();
  destSheet.getRange(1, 1, out.length, out[0].length).setValues(out);

  debugSheet.appendRow([
    new Date(),
    'SUMMARY',
    `Processed: ${Object.keys(byStudentID).length}, Perfect: ${perfectMatch}, Fallback: ${fallback}, AssignedNoTeam: ${assignedNoTeam}, Skipped: ${skipped}`,
  ]);
  if (skippedList.length) {
    debugSheet.appendRow([new Date(), 'SKIPPED_IDS', skippedList.join(',')]);
  }
  Logger.log(
    `StudentList build complete. Total students: ${Object.keys(byStudentID).length}, Perfect: ${perfectMatch}, Fallback: ${fallback}, AssignedNoTeam: ${assignedNoTeam}, Skipped: ${skipped}`
  );
}

function sendNightlyStatusEmail(): void {
  const EMAIL = 'tyler.gibson@jefferson.kyschools.us';
  const SUBJECT = 'BSS Nightly Status Report';

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName('DebugLog');
  if (!logSheet) {
    Logger.log('No DebugLog sheet found, skipping status email.');
    return;
  }

  const logData = logSheet.getDataRange().getValues();
  const now = new Date();
  const cutoff = new Date(now.getTime() - 48 * 60 * 60 * 1000); // 48 hours ago

  const scriptRuns: string[] = [];
  const errorLines: string[] = [];
  const summaryLines: string[] = [];

  for (let i = logData.length - 1; i >= 0; i--) {
    const row = logData[i];
    const timestamp = row[0];
    const type = row[1];
    const detail = row[2];
    const ts = timestamp instanceof Date ? timestamp : new Date(timestamp);
    if (ts < cutoff) break;

    if (typeof type === 'string') {
      if (
        type.match(/import/i) ||
        type.match(/StudentList/i) ||
        type.match(/Roster sync/i) ||
        type.match(/SUCCESS/i) ||
        type.match(/Starting sync/i)
      ) {
        scriptRuns.push(`[${ts.toLocaleString()}] ${type} ${detail || ''}`);
      }
      if (type.match(/error/i)) {
        errorLines.push(`[${ts.toLocaleString()}] ${type} ${detail || ''}`);
      }
      if (type.match(/summary/i) || type.match(/complete/i)) {
        summaryLines.push(`[${ts.toLocaleString()}] ${type} ${detail || ''}`);
      }
    }
  }

  let body = `<b>BSS Nightly Status Report (${now.toLocaleString()})</b><br><br>`;
  body += `<b>Script Runs:</b><ul>${scriptRuns.map((s) => `<li>${s}</li>`).join('')}</ul>`;
  body += `<b>Summaries:</b><ul>${summaryLines.map((s) => `<li>${s}</li>`).join('')}</ul>`;
  body += `<b>Errors/Exceptions:</b><ul>${errorLines
    .map((s) => `<li style="color:#b22222;">${s}</li>`)
    .join('')}</ul>`;

  if (!scriptRuns.length && !summaryLines.length && !errorLines.length) {
    body += `<i>No log entries found for the past 48 hours. Check if triggers are firing.</i>`;
  }

  MailApp.sendEmail({
    to: EMAIL,
    subject: SUBJECT,
    htmlBody: body,
  });
}

/* -------------- END OF FILE -------------- */
