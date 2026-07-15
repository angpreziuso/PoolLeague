/**
 * RACK SHEET — League Backend (Google Apps Script, bound to your Google Sheet)
 * ============================================================================
 * WHAT THIS FILE DOES
 *  1. Sets up all the sheet tabs the app needs (run setupSheets() once).
 *  2. Keeps the Google Form's player dropdowns in sync with your Roster tab,
 *     so scorekeepers pick names instead of typing them (run syncRosterToForm()
 *     any time the roster changes, or let the installable trigger do it).
 *  3. Parses each form submission into per-game rows, runs the rating
 *     algorithm, flags anything that needs a human (early 9, golden break,
 *     break-and-run, 3-foul win), and drops a PENDING match into the Matches
 *     tab.
 *  4. Lets an admin approve a pending match (checkbox in the sheet), which
 *     locks in the rating changes and updates team points/standings.
 *  5. Serves read-only JSON (players / teams / schedule / standings) at a
 *     Web App URL your frontend can fetch from.
 *
 * SETUP — do this once:
 *  A. Extensions > Apps Script on your Google Sheet, paste this file in as Code.gs.
 *  B. Run setupSheets from the Apps Script editor (it will ask for permissions).
 *  C. Create your Google Form, then set FORM_ID below to its form ID
 *     (the long string in the form's edit URL).
 *  D. In the Form, create the questions listed in FORM_QUESTION_TITLES exactly
 *     as named there (title text is how this script finds them).
 *  E. Run syncRosterToForm once to populate the player dropdowns.
 *  F. Deploy > New deployment > Web app, execute as "Me", access "Anyone" —
 *     copy the URL into the frontend's API_BASE constant.
 *  G. Form > ⋮ > Select response destination > this spreadsheet, so
 *     onFormSubmit can fire.
 *  H. Triggers (clock icon, left sidebar) > Add Trigger:
 *       - onFormSubmitInstalled, from Form, on form submit
 *       - onEditInstalled, from spreadsheet, on edit
 * ============================================================================
 */

// ---------------------------------------------------------------------------
// CONFIG — fill these in for your league
// ---------------------------------------------------------------------------
const FORM_ID = 'PASTE_YOUR_GOOGLE_FORM_ID_HERE';
const MAX_GAMES = 9;

// Expected-innings table by rating (per league rules)
const EXPECTED_INNINGS = {10:1, 9:2, 8:3, 7:4, 6:5, 5:6, 4:7, 3:8, 2:9, 1:10};

// Win types that are auto-scored (added on top of the innings-based delta)
const WIN_TYPE_BONUS = {
  'Hill-Hill': 0.1,
  'Normal': 0.2,
  'Sweep': 0.4,
};

// Win types that ALWAYS require manual admin review instead of an
// automatic rating change — the script logs the game but leaves the
// rating delta at 0 and raises an AdminAlerts row.
const MANUAL_REVIEW_WIN_TYPES = ['Early 9', 'Golden Break', 'Break and Run', '3-Foul Rule'];

// Sheet tab names — change here if you rename tabs
const SHEETS = {
  ROSTER: 'Roster',
  TEAMS: 'Teams',
  SCHEDULE: 'Schedule',
  RAW_RESPONSES: 'Form Responses 1',
  GAMES: 'Games',
  MATCHES: 'Matches',
  ADMIN_ALERTS: 'AdminAlerts',
  RATING_LOG: 'RatingLog',
};

// ---------------------------------------------------------------------------
// ONE-TIME SETUP
// ---------------------------------------------------------------------------
function setupSheets() {
  const ss = SpreadsheetApp.openById('1uckbx5NbiRDHADsfb-k2IDiiVk8CjGYIWWYiK_T7yZc');

  ensureSheet_(ss, SHEETS.ROSTER,
    ['PlayerName', 'PlayerNumber', 'Team', 'CurrentRating', 'Wins', 'Losses']);

  ensureSheet_(ss, SHEETS.TEAMS,
    ['TeamName', 'Captain', 'TotalPoints']);

  ensureSheet_(ss, SHEETS.SCHEDULE,
    ['Week', 'Date', 'Day', 'AwayTeam', 'HomeTeam', 'Location', 'TableNumber']);

  ensureSheet_(ss, SHEETS.GAMES, [
    'MatchId', 'GameNumber', 'HomePlayer', 'AwayPlayer', 'MatchLength', 'StartTime',
    'LagWinner', 'BallsOnBreak', 'MissCount', 'DefensiveShots', 'TimeOuts',
    'GameWinner', 'WinType', 'ExpectedInnings', 'ActualInnings', 'RatingDelta', 'NeedsReview'
  ]);

  ensureSheet_(ss, SHEETS.MATCHES, [
    'MatchId', 'Timestamp', 'Week', 'HomeTeam', 'AwayTeam', 'GamesHome', 'GamesAway',
    'TeamPointsHome', 'TeamPointsAway', 'Scorekeeper', 'Comments', 'Approved'
  ]);

  ensureSheet_(ss, SHEETS.ADMIN_ALERTS, [
    'MatchId', 'GameNumber', 'Player', 'AlertType', 'Timestamp', 'Resolved', 'AdminNotes'
  ]);

  ensureSheet_(ss, SHEETS.RATING_LOG, [
    'Timestamp', 'Player', 'MatchId', 'GameNumber', 'OldRating', 'Delta', 'NewRating', 'Reason'
  ]);

  SpreadsheetApp.getUi().alert('Sheet tabs are set up. Fill in Roster, Teams, and Schedule next.');
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
// KEEP FORM DROPDOWNS IN SYNC WITH THE ROSTER
// (Answers the "load names like a multiple choice box from the roster" ask)
// ---------------------------------------------------------------------------

// The exact question titles this script looks for in your Form.
// Create one Home/Away player question PER GAME (Game 1 Home Player ... Game 9 Away Player),
// plus a Scorekeeper question. Title text must match exactly.
function FORM_QUESTION_TITLES_() {
  const titles = ['Scorekeeper Name', 'Home Team', 'Away Team'];
  for (let g = 1; g <= MAX_GAMES; g++) {
    titles.push(`Game ${g} Home Player`, `Game ${g} Away Player`, `Game ${g} Lag Winner`, `Game ${g} Game Winner`);
  }
  return titles;
}

function syncRosterToForm() {
  if (FORM_ID === 'PASTE_YOUR_GOOGLE_FORM_ID_HERE') {
    SpreadsheetApp.getUi().alert('Set FORM_ID at the top of Code.gs first.');
    return;
  }
  const form = FormApp.openById(FORM_ID);
  const roster = getRosterNames_();
  const teamNames = getTeamNames_();

  form.getItems().forEach(item => {
    const title = item.getTitle();
    // Player-pickers: any question ending in "Player" or "Lag Winner" / "Game Winner"
    // gets the full roster as choices. In practice you'll want the
    // Home/Away player questions scoped to that match's two rosters — the
    // simplest reliable version (used here) offers every player;
    // scorekeepers pick from a short, correctly-spelled list either way.
    if (/Player$|Lag Winner$|Game Winner$|Scorekeeper Name$/.test(title)) {
      setChoices_(item, roster);
    }
    if (title === 'Home Team' || title === 'Away Team') {
      setChoices_(item, teamNames);
    }
  });

  SpreadsheetApp.getUi().alert('Form dropdowns synced to roster (' + roster.length + ' players).');
}

function setChoices_(item, choices) {
  const type = item.getType();
  if (type === FormApp.ItemType.LIST) {
    item.asListItem().setChoiceValues(choices);
  } else if (type === FormApp.ItemType.MULTIPLE_CHOICE) {
    item.asMultipleChoiceItem().setChoiceValues(choices);
  }
}

function getRosterNames_() {
  const rows = readSheet_(SHEETS.ROSTER);
  return rows.map(r => r.PlayerName).filter(Boolean);
}

function getTeamNames_() {
  const rows = readSheet_(SHEETS.TEAMS);
  return rows.map(r => r.TeamName).filter(Boolean);
}

// Optional: run this on a nightly time-driven trigger so the form never
// drifts out of sync when you edit the roster.
function onEditInstalled(e) {
  if (!e || !e.range) return;
  const sheetName = e.range.getSheet().getName();
  if (sheetName === SHEETS.ROSTER) {
    syncRosterToForm();
  }
  if (sheetName === SHEETS.MATCHES) {
    handleApprovalEdit_(e);
  }
}

// ---------------------------------------------------------------------------
// FORM SUBMISSION -> GAMES + MATCH (rating math happens here)
// ---------------------------------------------------------------------------
function onFormSubmitInstalled(e) {
  const responses = e.namedValues; // { "Question Title": ["answer"] }
  const get = (title) => (responses[title] && responses[title][0]) || '';

  const matchId = Utilities.getUuid();
  const scorekeeper = get('Scorekeeper Name');
  const homeTeam = get('Home Team');
  const awayTeam = get('Away Team');
  const comments = get('Comments');

  const gamesSheet = SpreadsheetApp.openById('1uckbx5NbiRDHADsfb-k2IDiiVk8CjGYIWWYiK_T7yZc').getSheetByName(SHEETS.GAMES);
  let homeGamesWon = 0, awayGamesWon = 0;

  for (let g = 1; g <= MAX_GAMES; g++) {
    const homePlayer = get(`Game ${g} Home Player`);
    const awayPlayer = get(`Game ${g} Away Player`);
    if (!homePlayer && !awayPlayer) continue; // game not played (match ended early)

    const matchLength = get(`Game ${g} Match Length`);
    const startTime = get(`Game ${g} Start Time`);
    const lagWinner = get(`Game ${g} Lag Winner`);
    const ballsOnBreak = Number(get(`Game ${g} Balls On Break`)) || 0;
    const missCount = Number(get(`Game ${g} Miss Count`)) || 0;
    const defensiveShots = Number(get(`Game ${g} Defensive Shot Count`)) || 0;
    const timeOuts = Number(get(`Game ${g} Time Out Count`)) || 0;
    const gameWinner = get(`Game ${g} Game Winner`);
    const winType = get(`Game ${g} How Win Was Achieved`);
    // Innings played this game = misses + game-ending shot, approximated as
    // missCount + 1 (each miss is a lost inning; the final successful
    // inning closes it out). Adjust here if your scoresheet counts innings directly.
    const actualInnings = missCount + 1;

    if (gameWinner === homePlayer) homeGamesWon++;
    if (gameWinner === awayPlayer) awayGamesWon++;

    const loser = gameWinner === homePlayer ? awayPlayer : homePlayer;
    const ratingResult = computeRatingDelta_(gameWinner, loser, actualInnings, winType);

    gamesSheet.appendRow([
      matchId, g, homePlayer, awayPlayer, matchLength, startTime, lagWinner,
      ballsOnBreak, missCount, defensiveShots, timeOuts, gameWinner, winType,
      ratingResult.expectedInnings, actualInnings, ratingResult.winnerDelta, ratingResult.needsReview
    ]);

    if (ratingResult.needsReview) {
      SpreadsheetApp.openById('1uckbx5NbiRDHADsfb-k2IDiiVk8CjGYIWWYiK_T7yZc').getSheetByName(SHEETS.ADMIN_ALERTS).appendRow([
        matchId, g, gameWinner, winType, new Date(), false, ''
      ]);
    }
  }

  // Team points (0-25) — plug in your league's actual points table here;
  // this default just scales game-win margin. Replace with your real rule.
  const teamPointsHome = Math.round((homeGamesWon / (homeGamesWon + awayGamesWon || 1)) * 25);
  const teamPointsAway = 25 - teamPointsHome;

  SpreadsheetApp.openById('1uckbx5NbiRDHADsfb-k2IDiiVk8CjGYIWWYiK_T7yZc').getSheetByName(SHEETS.MATCHES).appendRow([
    matchId, new Date(), '', homeTeam, awayTeam, homeGamesWon, awayGamesWon,
    teamPointsHome, teamPointsAway, scorekeeper, comments, false // Approved = false until admin checks it
  ]);
}

/**
 * Rating algorithm.
 *   - Base delta from innings-vs-expected (per league rule):
 *       win in FEWER innings than expected  -> +0.1
 *       win in MORE innings than expected   -> -0.1
 *       loss in FEWER innings than expected -> no change
 *       loss in MORE innings than expected  -> -0.1
 *   - On top of a win, the win-type bonus is added (Hill-Hill +0.1,
 *     Normal +0.2, Sweep +0.4) — ASSUMPTION: these stack with the base
 *     delta above rather than replace it. Adjust WIN_TYPE_BONUS logic
 *     below if your league intends them as alternatives instead.
 *   - Early 9 / Golden Break / Break and Run / 3-Foul Rule wins are NEVER
 *     auto-scored — delta is forced to 0 and an AdminAlerts row is raised.
 */
function computeRatingDelta_(winnerName, loserName, actualInnings, winType) {
  const roster = getRosterMap_();
  const winner = roster[winnerName];
  const expectedInnings = winner ? EXPECTED_INNINGS[winner.CurrentRating] || 10 : 10;

  if (MANUAL_REVIEW_WIN_TYPES.indexOf(winType) !== -1) {
    return { winnerDelta: 0, expectedInnings, needsReview: true };
  }

  let delta = actualInnings < expectedInnings ? 0.1
            : actualInnings > expectedInnings ? -0.1
            : 0;

  if (WIN_TYPE_BONUS[winType] !== undefined) {
    delta += WIN_TYPE_BONUS[winType];
  }

  return { winnerDelta: round1_(delta), expectedInnings, needsReview: false };
}

function round1_(n) { return Math.round(n * 10) / 10; }

function getRosterMap_() {
  const rows = readSheet_(SHEETS.ROSTER);
  const map = {};
  rows.forEach(r => { map[r.PlayerName] = r; });
  return map;
}

// ---------------------------------------------------------------------------
// ADMIN APPROVAL — check the "Approved" box on a Matches row to finalize it
// ---------------------------------------------------------------------------
function handleApprovalEdit_(e) {
  const sheet = e.range.getSheet();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const approvedCol = headers.indexOf('Approved') + 1;
  if (e.range.getColumn() !== approvedCol) return;
  if (e.value !== 'TRUE') return;

  const row = e.range.getRow();
  const matchId = sheet.getRange(row, headers.indexOf('MatchId') + 1).getValue();
  finalizeMatch_(matchId);
}

function finalizeMatch_(matchId) {
  const gamesSheet = SpreadsheetApp.openById('1uckbx5NbiRDHADsfb-k2IDiiVk8CjGYIWWYiK_T7yZc').getSheetByName(SHEETS.GAMES);
  const rosterSheet = SpreadsheetApp.openById('1uckbx5NbiRDHADsfb-k2IDiiVk8CjGYIWWYiK_T7yZc').getSheetByName(SHEETS.ROSTER);
  const rows = readSheet_(SHEETS.GAMES).filter(r => r.MatchId === matchId);
  const rosterRows = readSheet_(SHEETS.ROSTER);
  const rosterHeaders = rosterSheet.getRange(1, 1, 1, rosterSheet.getLastColumn()).getValues()[0];
  const nameCol = rosterHeaders.indexOf('PlayerName');
  const ratingCol = rosterHeaders.indexOf('CurrentRating');

  rows.forEach(g => {
    if (g.NeedsReview === true || g.NeedsReview === 'TRUE') return; // wait for admin to set the delta manually
    const delta = Number(g.RatingDelta) || 0;
    if (!delta) return;
    for (let i = 0; i < rosterRows.length; i++) {
      if (rosterRows[i].PlayerName === g.GameWinner) {
        const sheetRow = i + 2; // header + 1-index
        const oldRating = Number(rosterSheet.getRange(sheetRow, ratingCol + 1).getValue());
        const newRating = round1_(oldRating + delta);
        rosterSheet.getRange(sheetRow, ratingCol + 1).setValue(newRating);
        SpreadsheetApp.openById('1uckbx5NbiRDHADsfb-k2IDiiVk8CjGYIWWYiK_T7yZc').getSheetByName(SHEETS.RATING_LOG).appendRow([
          new Date(), g.GameWinner, matchId, g.GameNumber, oldRating, delta, newRating, g.WinType
        ]);
      }
    }
  });

  recalculateStandings_();
}

function recalculateStandings_() {
  const teamsSheet = SpreadsheetApp.openById('1uckbx5NbiRDHADsfb-k2IDiiVk8CjGYIWWYiK_T7yZc').getSheetByName(SHEETS.TEAMS);
  const matches = readSheet_(SHEETS.MATCHES).filter(m => m.Approved === true || m.Approved === 'TRUE');
  const teamRows = readSheet_(SHEETS.TEAMS);
  const totals = {};
  teamRows.forEach(t => totals[t.TeamName] = 0);
  matches.forEach(m => {
    totals[m.HomeTeam] = (totals[m.HomeTeam] || 0) + Number(m.TeamPointsHome || 0);
    totals[m.AwayTeam] = (totals[m.AwayTeam] || 0) + Number(m.TeamPointsAway || 0);
  });
  const headers = teamsSheet.getRange(1, 1, 1, teamsSheet.getLastColumn()).getValues()[0];
  const ptsCol = headers.indexOf('TotalPoints') + 1;
  teamRows.forEach((t, i) => {
    teamsSheet.getRange(i + 2, ptsCol).setValue(totals[t.TeamName] || 0);
  });
}

// ---------------------------------------------------------------------------
// GENERIC SHEET READER — returns array of objects keyed by header row
// ---------------------------------------------------------------------------
function readSheet_(name) {
  const sheet = SpreadsheetApp.openById('1uckbx5NbiRDHADsfb-k2IDiiVk8CjGYIWWYiK_T7yZc').getSheetByName(name);
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
      // only publish approved matches — pending ones stay internal until an admin checks the box
      data = readSheet_(SHEETS.MATCHES).filter(m => m.Approved === true || m.Approved === 'TRUE');
      break;
    default:
      return ContentService.createTextOutput(JSON.stringify({ error: 'Unknown resource. Use players, teams, schedule, standings, or matches.' }))
        .setMimeType(ContentService.MimeType.JSON);
  }

  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ---------------------------------------------------------------------------
// MENU — adds "League Admin" to the Sheet's menu bar for one-click actions
// ---------------------------------------------------------------------------
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('League Admin')
    .addItem('1. Set up sheet tabs', 'setupSheets')
    .addItem('2. Sync roster to Form dropdowns', 'syncRosterToForm')
    .addItem('Recalculate standings', 'recalculateStandings_')
    .addToUi();
}
