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
  GAMES: 'Games',
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

  // One row per individual game (rack): player vs player. Fill this in for
  // per-player stats like head-to-head records and win rates.
  ensureSheet_(ss, SHEETS.GAMES,
    ['Week', 'Date', 'HomePlayer', 'AwayPlayer', 'Winner', 'WinType']);

  installFormulas_(ss);

  notify_('Sheet tabs are set up. Fill in Roster (names/rating), Teams, Schedule, Matches, and Games. Wins/Losses and TotalPoints auto-fill from what you enter.');
}

// ---------------------------------------------------------------------------
// AUTO-POPULATION — derive columns from what you type, all inside the sheet.
//   - Roster Wins   = games in the Games tab that the player won
//   - Roster Losses = games the player appeared in but did not win
//   - Teams TotalPoints = team points summed from the Matches tab
// These are live formulas: type into Games/Matches and these columns update
// themselves. Don't hand-edit the Wins, Losses, or TotalPoints columns.
// Run this once (it's also called by setupSheets); safe to re-run.
// ---------------------------------------------------------------------------
function installFormulas_(ss) {
  ss = ss || SpreadsheetApp.openById(SPREADSHEET_ID);

  const roster = ss.getSheetByName(SHEETS.ROSTER);
  if (roster) {
    // Wins (col E): count of Games where this player is the Winner.
    roster.getRange('E2').setFormula(
      '=ARRAYFORMULA(IF(A2:A="","",COUNTIF(Games!E:E,A2:A)))');
    // Losses (col F): games the player appeared in (home or away) minus wins.
    roster.getRange('F2').setFormula(
      '=ARRAYFORMULA(IF(A2:A="","",COUNTIF(Games!C:C,A2:A)+COUNTIF(Games!D:D,A2:A)-COUNTIF(Games!E:E,A2:A)))');
  }

  const teams = ss.getSheetByName(SHEETS.TEAMS);
  if (teams) {
    // TotalPoints (col C): team points from Matches, home + away.
    teams.getRange('C2').setFormula(
      '=ARRAYFORMULA(IF(A2:A="","",SUMIF(Matches!C:C,A2:A,Matches!G:G)+SUMIF(Matches!D:D,A2:A,Matches!H:H)))');
  }
}

function installFormulas() {
  installFormulas_();
  notify_('Auto-fill formulas installed. Wins/Losses come from the Games tab; TotalPoints comes from the Matches tab.');
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
//   {url}?resource=games
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
    case 'games':
      data = readSheet_(SHEETS.GAMES);
      break;
    default:
      return ContentService.createTextOutput(JSON.stringify({ error: 'Unknown resource. Use players, teams, schedule, standings, matches, or games.' }))
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
      .addItem('Install auto-fill formulas', 'installFormulas')
      .addToUi();
  } catch (err) {
    Logger.log('onOpen menu skipped (standalone script): ' + err);
  }
}
