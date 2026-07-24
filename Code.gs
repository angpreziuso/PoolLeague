/**
 * POOL LEAGUE — League Backend (Google Apps Script)
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

// ---------------------------------------------------------------------------
// THE RULE BOOK
// Official World Pool-Billiard Association rules, available to the frontend
// at: {web-app-url}?resource=rules
// ---------------------------------------------------------------------------
const RULE_BOOK = {
  title: 'The Rule Book',
  url: 'https://wpapool.com/rules/',
};

// ---------------------------------------------------------------------------
// NEW PLAYER SIGNUP
// Run createPlayerSignupForm() once to create a shareable Google Form. Form
// responses are saved to this league spreadsheet; the signup link is retained
// in the script's document properties and can be retrieved later with
// getPlayerSignupFormUrl().
// ---------------------------------------------------------------------------
const PLAYER_SIGNUP_FORM_PROPERTY = 'PLAYER_SIGNUP_FORM_URL';

function createPlayerSignupForm() {
  const properties = PropertiesService.getDocumentProperties();
  const existingUrl = properties.getProperty(PLAYER_SIGNUP_FORM_PROPERTY);
  if (existingUrl) {
    notify_('New Player Signup form already exists: ' + existingUrl);
    return existingUrl;
  }

  const form = FormApp.create('New Player Signup')
    .setDescription('Welcome to Shark Club. Please provide contact info below to receive event updates.')
    .setConfirmationMessage('Thanks for signing up for Shark Club. We will be in touch with league information.');

  const emailValidation = FormApp.createTextValidation()
    .requireTextIsEmail()
    .setHelpText('Please enter a valid email address.')
    .build();

  form.addTextItem()
    .setTitle('Email')
    .setHelpText('Please provide your primary email address that we will use for your account, to view scores and send you additional league info.')
    .setValidation(emailValidation)
    .setRequired(true);
  form.addTextItem()
    .setTitle('Full Name (as you would like to be called)')
    .setRequired(true);
  form.addTextItem()
    .setTitle('Cellphone Number')
    .setRequired(true);
  form.addMultipleChoiceItem()
    .setTitle('Are you currently ranked in any league?')
    .setChoiceValues([
      'Have never played in any league',
      'Yes — I will provide league name(s), ranking, and details below',
      'Yes, but it was a long time ago and my level is not accurate — I will provide details below',
      'Other',
    ])
    .setRequired(true);
  form.addParagraphTextItem()
    .setTitle('League name(s), ranking, or additional details')
    .setRequired(false);

  form.setDestination(FormApp.DestinationType.SPREADSHEET, SPREADSHEET_ID);
  const publishedUrl = form.getPublishedUrl();
  properties.setProperty(PLAYER_SIGNUP_FORM_PROPERTY, publishedUrl);
  notify_('New Player Signup form created: ' + publishedUrl);
  return publishedUrl;
}

function getPlayerSignupFormUrl() {
  const url = PropertiesService.getDocumentProperties()
    .getProperty(PLAYER_SIGNUP_FORM_PROPERTY) || '';
  notify_(url
    ? 'New Player Signup form: ' + url
    : 'No signup form exists yet. Choose “Create New Player Signup form” first.');
  return url;
}

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

  // Columns you TYPE: PlayerName, PlayerNumber, Team, CurrentRating.
  // Columns that AUTO-FILL: Wins, Losses, GamesPlayed, WinRate, Streak, Badges.
  ensureSheet_(ss, SHEETS.ROSTER,
    ['PlayerName', 'PlayerNumber', 'Team', 'CurrentRating',
     'Wins', 'Losses', 'GamesPlayed', 'WinRate', 'Streak', 'Badges']);

  // You TYPE: TeamName, Captain. AUTO-FILL: TotalPoints, Wins, Losses, Rank, GamesBack.
  ensureSheet_(ss, SHEETS.TEAMS,
    ['TeamName', 'Captain', 'TotalPoints', 'Wins', 'Losses', 'Rank', 'GamesBack']);

  // You TYPE everything except Result, which auto-fills from the Matches tab.
  ensureSheet_(ss, SHEETS.SCHEDULE,
    ['Week', 'Date', 'Day', 'AwayTeam', 'HomeTeam', 'Location', 'TableNumber', 'Result']);

  // You TYPE: Week, Date, HomeTeam, AwayTeam, TeamPointsHome, TeamPointsAway, Comments.
  // AUTO-FILL: GamesHome, GamesAway (counted from the Games tab).
  ensureSheet_(ss, SHEETS.MATCHES,
    ['Week', 'Date', 'HomeTeam', 'AwayTeam', 'GamesHome', 'GamesAway',
     'TeamPointsHome', 'TeamPointsAway', 'Comments']);

  // One row per individual game (rack): player vs player.
  // You TYPE: Week, Date, HomePlayer, AwayPlayer, Winner, WinType.
  // AUTO-FILL: HomePlayerTeam, AwayPlayerTeam, WinnerTeam (looked up from Roster).
  ensureSheet_(ss, SHEETS.GAMES,
    ['Week', 'Date', 'HomePlayer', 'AwayPlayer', 'Winner', 'WinType',
     'HomePlayerTeam', 'AwayPlayerTeam', 'WinnerTeam']);

  installFormulas_(ss);

  notify_('Sheet tabs are set up. Fill in Roster (names/rating), Teams, Schedule, Matches, and Games. Wins/Losses and TotalPoints auto-fill from what you enter.');
}

// ---------------------------------------------------------------------------
// AUTO-POPULATION — derive columns from what you type, all inside the sheet.
// Everything below is a live spreadsheet formula: nothing external writes to
// the board, it all recalculates from what you enter. Don't hand-edit any of
// the auto-filled columns (their formulas would be overwritten).
//
//  ROSTER   Wins, Losses, GamesPlayed, WinRate, Streak, Badges  (from Games)
//  TEAMS    TotalPoints (from Matches points), Wins, Losses, Rank, GamesBack
//  SCHEDULE Result       (from the matching Matches row)
//  MATCHES  GamesHome, GamesAway  (counted from the Games tab)
//  GAMES    HomePlayerTeam, AwayPlayerTeam, WinnerTeam  (looked up from Roster)
//
// Run this once (it's also called by setupSheets); safe to re-run.
// ---------------------------------------------------------------------------
const FORMULA_ROWS = 500; // how far down the per-row formulas are filled

// Set a per-row formula in row 2 and fill it down (relative refs adjust).
function fillDown_(sheet, col, formula) {
  sheet.getRange(col + '2').setFormula(formula);
  sheet.getRange(col + '2').copyTo(sheet.getRange(col + '3:' + col + FORMULA_ROWS));
}

function installFormulas_(ss) {
  ss = ss || SpreadsheetApp.openById(SPREADSHEET_ID);

  const roster = ss.getSheetByName(SHEETS.ROSTER);
  if (roster) {
    // Wins (E): games this player won.
    roster.getRange('E2').setFormula(
      '=ARRAYFORMULA(IF(A2:A="","",COUNTIF(Games!E:E,A2:A)))');
    // Losses (F): games the player appeared in (home or away) minus wins.
    roster.getRange('F2').setFormula(
      '=ARRAYFORMULA(IF(A2:A="","",COUNTIF(Games!C:C,A2:A)+COUNTIF(Games!D:D,A2:A)-COUNTIF(Games!E:E,A2:A)))');
    // GamesPlayed (G) [2]: wins + losses.
    roster.getRange('G2').setFormula(
      '=ARRAYFORMULA(IF(A2:A="","",E2:E+F2:F))');
    // WinRate % (H) [1]: wins / games played, rounded to a whole percent.
    roster.getRange('H2').setFormula(
      '=ARRAYFORMULA(IF(A2:A="","",IF(G2:G=0,0,ROUND(100*E2:E/G2:G))))');
    // Streak (I) [8]: current run like "W3" / "L2" from the player's games in
    // row order (assumes games are entered top-to-bottom in time order).
    fillDown_(roster, 'I',
      '=IF(A2="","",IFERROR(' +
        'LET(res,TEXTJOIN("",TRUE,IF((Games!$C$2:$C=A2)+(Games!$D$2:$D=A2)>0,IF(Games!$E$2:$E=A2,"W","L"),"")),' +
        'IF(res="","-",' +
        'LET(lastc,RIGHT(res,1),n,LEN(res),' +
        'diffpos,IFERROR(MAX(FILTER(SEQUENCE(n),MID(res,SEQUENCE(n),1)<>lastc)),0),' +
        'lastc&(n-diffpos)))),"-"))');
    // Badges (J) [9]: simple auto-awards derived from the stats above.
    fillDown_(roster, 'J',
      '=IF(A2="","",TRIM(' +
        'IF(AND(H2>=70,G2>=5),"Sharpshooter  ","")&' +
        'IF(G2>=20,"Veteran  ","")&' +
        'IF(AND(LEFT(I2,1)="W",IFERROR(VALUE(MID(I2,2,5)),0)>=3),"Hot Streak","")))');
  }

  const teams = ss.getSheetByName(SHEETS.TEAMS);
  if (teams) {
    // TotalPoints (C): team points from Matches, home + away.
    teams.getRange('C2').setFormula(
      '=ARRAYFORMULA(IF(A2:A="","",SUMIF(Matches!C:C,A2:A,Matches!G:G)+SUMIF(Matches!D:D,A2:A,Matches!H:H)))');
    // Team Wins (D) [3]: matches won as home or away (more games than opponent).
    fillDown_(teams, 'D',
      '=IF(A2="","",SUMPRODUCT((Matches!$C$2:$C=A2)*(Matches!$E$2:$E>Matches!$F$2:$F))' +
      '+SUMPRODUCT((Matches!$D$2:$D=A2)*(Matches!$F$2:$F>Matches!$E$2:$E)))');
    // Team Losses (E) [3]: matches lost as home or away.
    fillDown_(teams, 'E',
      '=IF(A2="","",SUMPRODUCT((Matches!$C$2:$C=A2)*(Matches!$E$2:$E<Matches!$F$2:$F))' +
      '+SUMPRODUCT((Matches!$D$2:$D=A2)*(Matches!$F$2:$F<Matches!$E$2:$E)))');
    // Rank (F) [4]: standings position by TotalPoints (1 = most points).
    fillDown_(teams, 'F',
      '=IF(A2="","",IFERROR(RANK(C2,FILTER($C$2:$C$' + FORMULA_ROWS + ',$A$2:$A$' + FORMULA_ROWS + '<>"")),""))');
    // GamesBack (G) [4]: points behind the league leader.
    fillDown_(teams, 'G',
      '=IF(A2="","",IFERROR(MAX(FILTER($C$2:$C$' + FORMULA_ROWS + ',$A$2:$A$' + FORMULA_ROWS + '<>""))-C2,""))');
  }

  const schedule = ss.getSheetByName(SHEETS.SCHEDULE);
  if (schedule) {
    // Result (H) [7]: the score for this scheduled match, pulled from Matches
    // by Week + HomeTeam + AwayTeam. Shows "Not played" until a match is recorded.
    fillDown_(schedule, 'H',
      '=IF(E2="","",IFERROR(' +
        'LET(gh,SUMIFS(Matches!$E:$E,Matches!$A:$A,A2,Matches!$C:$C,E2,Matches!$D:$D,D2),' +
        'ga,SUMIFS(Matches!$F:$F,Matches!$A:$A,A2,Matches!$C:$C,E2,Matches!$D:$D,D2),' +
        'cnt,COUNTIFS(Matches!$A:$A,A2,Matches!$C:$C,E2,Matches!$D:$D,D2),' +
        'IF(cnt=0,"Not played",E2&" "&gh&"-"&ga&" "&D2)),"Not played"))');
  }

  const matches = ss.getSheetByName(SHEETS.MATCHES);
  if (matches) {
    // GamesHome (E) [6]: games in this match won by a home-team player.
    fillDown_(matches, 'E',
      '=IF(C2="","",SUMPRODUCT((Games!$A$2:$A=A2)*(Games!$I$2:$I=C2)*' +
      '(((Games!$G$2:$G=C2)*(Games!$H$2:$H=D2))+((Games!$G$2:$G=D2)*(Games!$H$2:$H=C2)))))');
    // GamesAway (F) [6]: games in this match won by an away-team player.
    fillDown_(matches, 'F',
      '=IF(C2="","",SUMPRODUCT((Games!$A$2:$A=A2)*(Games!$I$2:$I=D2)*' +
      '(((Games!$G$2:$G=C2)*(Games!$H$2:$H=D2))+((Games!$G$2:$G=D2)*(Games!$H$2:$H=C2)))))');
  }

  const games = ss.getSheetByName(SHEETS.GAMES);
  if (games) {
    // HomePlayerTeam (G) [5]: looked up from the Roster by player name.
    games.getRange('G2').setFormula(
      '=ARRAYFORMULA(IF(C2:C="","",IFERROR(VLOOKUP(C2:C,Roster!$A:$C,3,FALSE),"")))');
    // AwayPlayerTeam (H) [5].
    games.getRange('H2').setFormula(
      '=ARRAYFORMULA(IF(D2:D="","",IFERROR(VLOOKUP(D2:D,Roster!$A:$C,3,FALSE),"")))');
    // WinnerTeam (I): team of the game winner (used to tally match scores).
    games.getRange('I2').setFormula(
      '=ARRAYFORMULA(IF(E2:E="","",IFERROR(VLOOKUP(E2:E,Roster!$A:$C,3,FALSE),"")))');
  }
}

function installFormulas() {
  installFormulas_();
  notify_('Auto-fill formulas installed. Just type into the plain columns — the rest fills itself.');
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
//   {url}?resource=rules
//   {url}?resource=signup
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
    case 'rules':
      data = RULE_BOOK;
      break;
    case 'signup':
      data = {
        title: 'New Player Signup',
        url: PropertiesService.getDocumentProperties()
          .getProperty(PLAYER_SIGNUP_FORM_PROPERTY) || '',
      };
      break;
    default:
      return ContentService.createTextOutput(JSON.stringify({ error: 'Unknown resource. Use players, teams, schedule, standings, matches, games, rules, or signup.' }))
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
      .addSeparator()
      .addItem('Create New Player Signup form', 'createPlayerSignupForm')
      .addItem('Show New Player Signup link', 'getPlayerSignupFormUrl')
      .addToUi();
  } catch (err) {
    Logger.log('onOpen menu skipped (standalone script): ' + err);
  }
}
