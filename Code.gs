/**
 * RACK SHEET — League Backend (Google Apps Script)
 * ============================================================================
 * WHAT THIS FILE DOES
 *  1. Sets up the sheet tabs the app reads from (run setupSheets() once).
 *  2. Serves read-only JSON (players / teams / schedule / standings / matches)
 *     at a Web App URL your frontend can fetch from.
 *
 * The Google Sheet is the single source of truth: whatever you type into the
 * tabs is exactly what the webpage shows. Nothing else writes to the board —
 * there is no Google Form, no automatic rating math, and no approval step.
 *
 * SETUP — do this once:
 *  A. Paste this file into an Apps Script project.
 *  B. Run setupSheets from the editor (it will ask for permissions).
 *  C. Fill in the Roster, Teams, Schedule (and Matches) tabs by hand.
 *  D. Deploy > New deployment > Web app, execute as "Me", access "Anyone" —
 *     copy the URL into the frontend so it can fetch the data.
 * ============================================================================
 */

// ---------------------------------------------------------------------------
// CONFIG
// ---------------------------------------------------------------------------
const SPREADSHEET_ID = '1uckbx5NbiRDHADsfb-k2IDiiVk8CjGYIWWYiK_T7yZc';

// Sheet tab names — change here if you rename tabs
const SHEETS = {
  ROSTER: 'Roster',
  TEAMS: 'Teams',
  SCHEDULE: 'Schedule',
  MATCHES: 'Matches',
};

// ---------------------------------------------------------------------------
// ONE-TIME SETUP
// ---------------------------------------------------------------------------
function setupSheets() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  ensureSheet_(ss, SHEETS.ROSTER,
    ['PlayerName', 'PlayerNumber', 'Team', 'CurrentRating', 'Wins', 'Losses']);

  ensureSheet_(ss, SHEETS.TEAMS,
    ['TeamName', 'Captain', 'TotalPoints']);

  ensureSheet_(ss, SHEETS.SCHEDULE,
    ['Week', 'Date', 'Day', 'AwayTeam', 'HomeTeam', 'Location', 'TableNumber']);

  ensureSheet_(ss, SHEETS.MATCHES,
    ['Week', 'Date', 'HomeTeam', 'AwayTeam', 'GamesHome', 'GamesAway',
     'TeamPointsHome', 'TeamPointsAway', 'Comments']);

  notify_('Sheet tabs are set up. Fill in Roster, Teams, Schedule, and Matches.');
}

// Shows a pop-up when the script is bound to a sheet; falls back to the
// execution log in a standalone project (where getUi() isn't available).
function notify_(message) {
  try {
    SpreadsheetApp.getUi().alert(message);
  } catch (err) {
    Logger.log(message);
  }
}

function ensureSheet_(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// ---------------------------------------------------------------------------
// GENERIC SHEET READER — returns array of objects keyed by header row
// ---------------------------------------------------------------------------
function readSheet_(name) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(name);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const values = sheet.getDataRange().getValues();
  const headers = values.shift();
  return values.map(row => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    return obj;
  });
}

// ---------------------------------------------------------------------------
// WEB APP API — read-only JSON for the frontend
// Deploy this project as a Web App, then fetch:
//   {url}?resource=players
//   {url}?resource=teams
//   {url}?resource=schedule
//   {url}?resource=standings
//   {url}?resource=matches
// ---------------------------------------------------------------------------
function doGet(e) {
  const resource = (e.parameter.resource || '').toLowerCase();
  let data;

  switch (resource) {
    case 'players':
      data = readSheet_(SHEETS.ROSTER);
      break;
    case 'teams':
      data = readSheet_(SHEETS.TEAMS);
      break;
    case 'schedule':
      data = readSheet_(SHEETS.SCHEDULE);
      break;
    case 'standings':
      data = readSheet_(SHEETS.TEAMS).sort((a, b) => Number(b.TotalPoints) - Number(a.TotalPoints));
      break;
    case 'matches':
      data = readSheet_(SHEETS.MATCHES);
      break;
    default:
      return ContentService.createTextOutput(JSON.stringify({ error: 'Unknown resource. Use players, teams, schedule, standings, or matches.' }))
        .setMimeType(ContentService.MimeType.JSON);
  }

  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ---------------------------------------------------------------------------
// MENU — adds "League Admin" to the Sheet's menu bar for one-click setup
// ---------------------------------------------------------------------------
function onOpen() {
  // Only works when the script is bound to the spreadsheet; ignored otherwise.
  try {
    SpreadsheetApp.getUi()
      .createMenu('League Admin')
      .addItem('Set up sheet tabs', 'setupSheets')
      .addToUi();
  } catch (err) {
    Logger.log('onOpen menu skipped (standalone script): ' + err);
  }
}
