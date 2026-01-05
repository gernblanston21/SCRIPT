/**
 * MSSG server.js — CAPS leagues everywhere (MLB/NBA/NHL)
 * - No npm deps (uses global fetch from Node 18+)
 * - .env loader (no package) for FD_API_KEY and SR_API_KEYs
 * - Pollers:
 * • Events list per league every 2 minutes (creates/removes game folders) // UPDATED
 * • FAST TICK (5s): watchlist games OR games marked InPlay
 * • SLOW TICK (120s): all other active events
 * - Roster Fetching:
 * • Fetches Sportradar rosters ONCE per game and caches them with unique names (e.g., homeRosterBOS.json).
 * • Builds a master roster from FanDuel and cached Sportradar files to enrich props data.
 * - Pause/Resume:
 * • If cache\STOP.flag exists -> PAUSED
 * • On entering PAUSED, blank odds/lines + props (keep team names)
 * • While PAUSED, skip all outbound API calls until STOP.flag is removed
 * - Log pruning:
 * • Delete logs older than 7 days
 * • Trim any log >5MB to last 1MB
 */

'use strict';

process.on('unhandledRejection', (e) => { console.error('[unhandledRejection]', e?.stack || e); process.exit(1); });
process.on('uncaughtException',  (e) => { console.error('[uncaughtException]',  e?.stack || e); process.exit(1); });

const fs   = require('fs');
const fsp  = require('fs/promises');
const path = require('path');

// === Paths & constants =======================================================
const BASE_DIR    = process.env.BASE_DIR || __dirname;
const CONFIG_DIR  = path.join(BASE_DIR, 'config');
const CACHE_DIR   = path.join(BASE_DIR, 'cache');
const LOG_DIR     = path.join(CACHE_DIR, 'logs');

const STOP_FLAG   = path.join(CACHE_DIR, 'STOP.flag');
let   PAUSED      = false;

const LEAGUES         = ['MLB', 'NHL', 'NBA']; // ALL CAPS everywhere
const WATCHLIST_FILE  = path.join(CONFIG_DIR, 'watchlist.json');

const WATCHLIST_MS    = 1_000;         // 1s
const EVENTS_POLL_MS  = 2 * 60_000;    // 2m <--- UPDATED FROM 10m
const TICK_FAST_MS    = 5_000;         // 5s (InPlay or on watchlist)
const TICK_SLOW_MS    = 120_000;       // 2m (others)

// === .env loader (no dependencies) ==========================================
function loadDotEnv(file) {
  const out = {};
  try {
    const raw = fs.readFileSync(file, 'utf8');
    for (const line of raw.split(/\r?\n/)) {
      const trimmed = line.trim();
      if (!trimmed || trimmed.startsWith('#')) continue;
      const eq = trimmed.indexOf('=');
      if (eq < 0) continue;
      const key = trimmed.slice(0, eq).trim();
      let val = trimmed.slice(eq + 1).trim();
      if ((val.startsWith('"') && val.endsWith('"')) || (val.startsWith("'") && val.endsWith("'"))) {
        val = val.slice(1, -1);
      }
      out[key] = val;
    }
  } catch {
    // ignore if missing
  }
  return out;
}
const LOCAL_ENV = loadDotEnv(path.join(BASE_DIR, '.env'));

// API Keys & Endpoints
const API_BASE       = 'https://partnerships-api.fanduel.com';
const ID_PROVIDER    = 'FANDUEL_SPORTSBOOK';
const FD_API_KEY     = process.env.FD_API_KEY || LOCAL_ENV.FD_API_KEY || '';

const SR_API_BASE    = 'https://api.sportradar.com';
const SR_API_KEY_NBA = process.env.SR_API_KEY_NBA || LOCAL_ENV.SR_API_KEY_NBA || '';
const SR_API_KEY_MLB = process.env.SR_API_KEY_MLB || LOCAL_ENV.SR_API_KEY_MLB || '';
const SR_API_KEY_NHL = process.env.SR_API_KEY_NHL || LOCAL_ENV.SR_API_KEY_NHL || '';

function getSportradarApiKey(league) {
  if (league === 'NBA') return SR_API_KEY_NBA;
  if (league === 'MLB') return SR_API_KEY_MLB;
  if (league === 'NHL') return SR_API_KEY_NHL;
  return '';
}

function apiHeaders() {
  return FD_API_KEY ? { 'x-api-key': FD_API_KEY } : {};
}


// === HTTP helper (no deps) ===================================================
async function j(url, options = {}) {
    const finalOptions = {
        ...options,
        headers: { ...(options.headers || {}), ...apiHeaders() }
    };
    const r = await fetch(url, finalOptions);
    if (!r.ok) throw new Error(`[http ${r.status}] ${url}`);
    return r.json();
}

// === FS helpers ==============================================================
async function ensureDir(d) { await fsp.mkdir(d, { recursive: true }); }
async function writeJSON(file, obj) {
  await ensureDir(path.dirname(file));
  await fsp.writeFile(file, JSON.stringify(obj, null, 2), 'utf8');
}
async function readJSON(file, fallback) {
  try { return JSON.parse(await fsp.readFile(file, 'utf8')); }
  catch { return fallback; }
}
async function existsPath(p) {
  try { await fsp.access(p, fs.constants.F_OK); return true; } catch { return false; }
}

// === Team mappings ===========================================================
// MLB
const MLB_TRI    = {"ARIZONA DIAMONDBACKS":"ARI","ATLANTA BRAVES":"ATL","BALTIMORE ORIOLES":"BAL","BOSTON RED SOX":"BOS","CHICAGO WHITE SOX":"CWS","CHICAGO CUBS":"CHC","CINCINNATI REDS":"CIN","CLEVELAND GUARDIANS":"CLE","COLORADO ROCKIES":"COL","DETROIT TIGERS":"DET","HOUSTON ASTROS":"HOU","KANSAS CITY ROYALS":"KC","LOS ANGELES ANGELS":"LAA","LOS ANGELES DODGERS":"LAD","MIAMI MARLINS":"MIA","MILWAUKEE BREWERS":"MIL","MINNESOTA TWINS":"MIN","NEW YORK METS":"NYM","NEW YORK YANKEES":"NYY","ATHLETICS":"ATH","PHILADELPHIA PHILLIES":"PHI","PITTSBURGH PIRATES":"PIT","SAN DIEGO PADRES":"SD","SAN FRANCISCO GIANTS":"SF","SEATTLE MARINERS":"SEA","ST. LOUIS CARDINALS":"STL","TAMPA BAY RAYS":"TB","TEXAS RANGERS":"TEX","TORONTO BLUE JAYS":"TOR","WASHINGTON NATIONALS":"WSH"};
const MLB_NICK   = {"ARIZONA DIAMONDBACKS":"DIAMONDBACKS","ATLANTA BRAVES":"BRAVES","BALTIMORE ORIOLES":"ORIOLES","BOSTON RED SOX":"RED SOX","CHICAGO WHITE SOX":"WHITE SOX","CHICAGO CUBS":"CUBS","CINCINNATI REDS":"REDS","CLEVELAND GUARDIANS":"GUARDIANS","COLORADO ROCKIES":"ROCKIES","DETROIT TIGERS":"TIGERS","HOUSTON ASTROS":"ASTROS","KANSAS CITY ROYALS":"ROYALS","LOS ANGELES ANGELS":"ANGELS","LOS ANGELES DODGERS":"DODGERS","MIAMI MARLINS":"MARLINS","MILWAUKEE BREWERS":"BREWERS","MINNESOTA TWINS":"TWINS","NEW YORK METS":"METS","NEW YORK YANKEES":"YANKEES","ATHLETICS":"ATHLETICS","PHILADELPHIA PHILLIES":"PHILLIES","PITTSBURGH PIRATES":"PIRATES","SAN DIEGO PADRES":"PADRES","SAN FRANCISCO GIANTS":"GIANTS","SEATTLE MARINERS":"MARINERS","ST. LOUIS CARDINALS":"CARDINALS","TAMPA BAY RAYS":"RAYS","TEXAS RANGERS":"RANGERS","TORONTO BLUE JAYS":"BLUE JAYS","WASHINGTON NATIONALS":"NATIONALS"};
const MLB_MARKET = {"ARIZONA DIAMONDBACKS":"ARIZONA","ATLANTA BRAVES":"ATLANTA","BALTIMORE ORIOLES":"BALTIMORE","BOSTON RED SOX":"BOSTON","CHICAGO WHITE SOX":"CHICAGO","CHICAGO CUBS":"CHICAGO","CINCINNATI REDS":"CINCINNATI","CLEVELAND GUARDIANS":"CLEVELAND","COLORADO ROCKIES":"COLORADO","DETROIT TIGERS":"DETROIT","HOUSTON ASTROS":"HOUSTON","KANSAS CITY ROYALS":"KANSAS CITY","LOS ANGELES ANGELS":"LOS ANGELES","LOS ANGELES DODGERS":"LOS ANGELES","MIAMI MARLINS":"MIAMI","MILWAUKEE BREWERS":"MILWAUKEE","MINNESOTA TWINS":"MINNESOTA","NEW YORK METS":"NEW YORK","NEW YORK YANKEES":"NEW YORK","ATHLETICS":"","PHILADELPHIA PHILLIES":"PHILADELPHIA","PITTSBURGH PIRATES":"PITTSBURGH","SAN DIEGO PADRES":"SAN DIEGO","SAN FRANCISCO GIANTS":"SAN FRANCISCO","SEATTLE MARINERS":"SEATTLE","ST. LOUIS CARDINALS":"ST. LOUIS","TAMPA BAY RAYS":"TAMPA BAY","TEXAS RANGERS":"TEXAS","TORONTO BLUE JAYS":"TORONTO","WASHINGTON NATIONALS":"WASHINGTON"};
// NHL
const NHL_TRI    = {"ANAHEIM DUCKS":"ANA","ARIZONA COYOTES":"ARI","BOSTON BRUINS":"BOS","BUFFALO SABRES":"BUF","CALGARY FLAMES":"CGY","CAROLINA HURRICANES":"CAR","CHICAGO BLACKHAWKS":"CHI","COLORADO AVALANCHE":"COL","COLUMBUS BLUE JACKETS":"CBJ","DALLAS STARS":"DAL","DETROIT RED WINGS":"DET","EDMONTON OILERS":"EDM","FLORIDA PANTHERS":"FLA","LOS ANGELES KINGS":"LA","MINNESOTA WILD":"MIN","MONTREAL CANADIENS":"MTL","NASHVILLE PREDATORS":"NSH","NEW JERSEY DEVILS":"NJ","NEW YORK ISLANDERS":"NYI","NEW YORK RANGERS":"NYR","OTTAWA SENATORS":"OTT","PHILADELPHIA FLYERS":"PHI","PITTSBURGH PENGUINS":"PIT","SAN JOSE SHARKS":"SJS","SEATTLE KRAKEN":"SEA","ST. LOUIS BLUES":"STL","TAMPA BAY LIGHTNING":"TB","TORONTO MAPLE LEAFS":"TOR","VANCOUVER CANUCKS":"VAN","VEGAS GOLDEN KNIGHTS":"VGK","UTAH MAMMOTH":"UTAH","WASHINGTON CAPITALS":"WSH","WINNIPEG JETS":"WPG"};
const NHL_NICK   = {"ANAHEIM DUCKS":"DUCKS","ARIZONA COYOTES":"COYOTES","BOSTON BRUINS":"BRUINS","BUFFALO SABRES":"SABRES","CALGARY FLAMES":"FLAMES","CAROLINA HURRICANES":"HURRICANES","CHICAGO BLACKHAWKS":"BLACKHAWKS","COLORADO AVALANCHE":"AVALANCHE","COLUMBUS BLUE JACKETS":"BLUE JACKETS","DALLAS STARS":"STARS","DETROIT RED WINGS":"RED WINGS","EDMONTON OILERS":"OILERS","FLORIDA PANTHERS":"PANTHERS","LOS ANGELES KINGS":"KINGS","MINNESOTA WILD":"WILD","MONTREAL CANADIENS":"CANADIENS","NASHVILLE PREDATORS":"PREDATORS","NEW JERSEY DEVILS":"DEVILS","NEW YORK ISLANDERS":"ISLANDERS","NEW YORK RANGERS":"RANGERS","OTTAWA SENATORS":"SENATORS","PHILADELPHIA FLYERS":"FLYERS","PITTSBURGH PENGUINS":"PENGUINS","SAN JOSE SHARKS":"SHARKS","SEATTLE KRAKEN":"KRAKEN","ST. LOUIS BLUES":"BLUES","TAMPA BAY LIGHTNING":"LIGHTNING","TORONTO MAPLE LEAFS":"MAPLE LEAFS","VANCOUVER CANUCKS":"CANUCKS","VEGAS GOLDEN KNIGHTS":"GOLDEN KNIGHTS","UTAH MAMMOTH":"MAMMOTH","WASHINGTON CAPITALS":"CAPITALS","WINNIPEG JETS":"JETS"};
const NHL_MARKET = {"ANAHEIM DUCKS":"ANAHEIM","ARIZONA COYOTES":"ARIZONA","BOSTON BRUINS":"BOSTON","BUFFALO SABRES":"BUFFALO","CALGARY FLAMES":"CALGARY","CAROLINA HURRICANES":"CAROLINA","CHICAGO BLACKHAWKS":"CHICAGO","COLORADO AVALANCHE":"COLORADO","COLUMBUS BLUE JACKETS":"COLUMBUS","DALLAS STARS":"DALLAS","DETROIT RED WINGS":"DETROIT","EDMONTON OILERS":"EDMONTON","FLORIDA PANTHERS":"FLORIDA","LOS ANGELES KINGS":"LOS ANGELES","MINNESOTA WILD":"MINNESOTA","MONTREAL CANADIENS":"MONTREAL","NASHVILLE PREDATORS":"NASHVILLE","NEW JERSEY DEVILS":"NEW JERSEY","NEW YORK ISLANDERS":"NEW YORK","NEW YORK RANGERS":"NEW YORK","OTTAWA SENATORS":"OTTAWA","PHILADELPHIA FLYERS":"PHILADELPHIA","PITTSBURGH PENGUINS":"PITTSBURGH","SAN JOSE SHARKS":"SAN JOSE","SEATTLE KRAKEN":"SEATTLE","ST. LOUIS BLUES":"ST. LOUIS","TAMPA BAY LIGHTNING":"TAMPA BAY","TORONTO MAPLE LEAFS":"TORONTO","VANCOUVER CANUCKS":"VANCOUVER","VEGAS GOLDEN KNIGHTS":"VEGAS","UTAH MAMMOTH":"UTAH","WASHINGTON CAPITALS":"WASHINGTON","WINNIPEG JETS":"WINNIPEG"};
// NBA
const NBA_TRI    = {"ATLANTA HAWKS":"ATL","BOSTON CELTICS":"BOS","BROOKLYN NETS":"BKN","CHARLOTTE HORNETS":"CHA","CHICAGO BULLS":"CHI","CLEVELAND CAVALIERS":"CLE","DALLAS MAVERICKS":"DAL","DENVER NUGGETS":"DEN","DETROIT PISTONS":"DET","GOLDEN STATE WARRIORS":"GS","HOUSTON ROCKETS":"HOU","INDIANA PACERS":"IND","LOS ANGELES CLIPPERS":"LAC","LOS ANGELES LAKERS":"LAL","MEMPHIS GRIZZLIES":"MEM","MIAMI HEAT":"MIA","MILWAUKEE BUCKS":"MIL","MINNESOTA TIMBERWOLVES":"MIN","NEW ORLEANS PELICANS":"NO","NEW YORK KNICKS":"NYK","OKLAHOMA CITY THUNDER":"OKC","ORLANDO MAGIC":"ORL","PHILADELPHIA 76ERS":"PHI","PHOENIX SUNS":"PHX","PORTLAND TRAIL BLAZERS":"POR","SACRAMENTO KINGS":"SAC","SAN ANTONIO SPURS":"SA","TORONTO RAPTORS":"TOR","UTAH JAZZ":"UTAH","WASHINGTON WIZARDS":"WAS"};
const NBA_NICK   = {"ATLANTA HAWKS":"HAWKS","BOSTON CELTICS":"CELTICS","BROOKLYN NETS":"NETS","CHARLOTTE HORNETS":"HORNETS","CHICAGO BULLS":"BULLS","CLEVELAND CAVALIERS":"CAVALIERS","DALLAS MAVERICKS":"MAVERICKS","DENVER NUGGETS":"NUGGETS","DETROIT PISTONS":"PISTONS","GOLDEN STATE WARRIORS":"WARRIORS","HOUSTON ROCKETS":"ROCKETS","INDIANA PACERS":"PACERS","LOS ANGELES CLIPPERS":"CLIPPERS","LOS ANGELES LAKERS":"LAKERS","MEMPHIS GRIZZLIES":"GRIZZLIES","MIAMI HEAT":"HEAT","MILWAUKEE BUCKS":"BUCKS","MINNESOTA TIMBERWOLVES":"TIMBERWOLVES","NEW ORLEANS PELICANS":"PELICANS","NEW YORK KNICKS":"KNICKS","OKLAHOMA CITY THUNDER":"THUNDER","ORLANDO MAGIC":"MAGIC","PHILADELPHIA 76ERS":"76ERS","PHOENIX SUNS":"SUNS","PORTLAND TRAIL BLAZERS":"TRAIL BLAZERS","SACRAMENTO KINGS":"KINGS","SAN ANTONIO SPURS":"SPURS","TORONTO RAPTORS":"RAPTORS","UTAH JAZZ":"JAZZ","WASHINGTON WIZARDS":"WIZARDS"};
const NBA_MARKET = {"ATLANTA HAWKS":"ATLANTA","BOSTON CELTICS":"BOSTON","BROOKLYN NETS":"BROOKLYN","CHARLOTTE HORNETS":"CHARLOTTE","CHICAGO BULLS":"CHICAGO","CLEVELAND CAVALIERS":"CLEVELAND","DALLAS MAVERICKS":"DALLAS","DENVER NUGGETS":"DENVER","DETROIT PISTONS":"DETROIT","GOLDEN STATE WARRIORS":"GOLDEN STATE","HOUSTON ROCKETS":"HOUSTON","INDIANA PACERS":"INDIANA","LOS ANGELES CLIPPERS":"LOS ANGELES","LOS ANGELES LAKERS":"LOS ANGELES","MEMPHIS GRIZZLIES":"MEMPHIS","MIAMI HEAT":"MIAMI","MILWAUKEE BUCKS":"MILWAUKEE","MINNESOTA TIMBERWOLVES":"MINNESOTA","NEW ORLEANS PELICANS":"NEW ORLEANS","NEW YORK KNICKS":"NEW YORK","OKLAHOMA CITY THUNDER":"OKLAHOMA CITY","ORLANDO MAGIC":"ORLANDO","PHILADELPHIA 76ERS":"PHILADELPHIA","PHOENIX SUNS":"PHOENIX","PORTLAND TRAIL BLAZERS":"PORTLAND","SACRAMENTO KINGS":"SACRAMENTO","SAN ANTONIO SPURS":"SAN ANTONIO","TORONTO RAPTORS":"TORONTO","UTAH JAZZ":"UTAH","WASHINGTON WIZARDS":"WASHINGTON"};

// --- Fanduel Name (Full CAPS) to Sportradar UUID (The Core Lookup Tables) ---
const MLB_SR_ID = {
  "ARIZONA DIAMONDBACKS":   "25507be1-6a68-4267-bd82-e097d94b359b",
  "ATLANTA BRAVES":         "12079497-e414-450a-8bf2-29f91de646bf",
  "BALTIMORE ORIOLES":      "75729d34-bca7-4a0f-b3df-6f26c6ad3719",
  "BOSTON RED SOX":         "93941372-eb4c-4c40-aced-fe3267174393",
  "CHICAGO WHITE SOX":      "47f490cd-2f58-4ef7-9dfd-2ad6ba6c1ae8",
  "CHICAGO CUBS":           "55714da8-fcaf-4574-8443-59bfb511a524",
  "CINCINNATI REDS":        "c874a065-c115-4e7d-b0f0-235584fb0e6f",
  "CLEVELAND GUARDIANS":    "80715d0d-0d2a-450f-a970-1b9a3b18c7e7",
  "COLORADO ROCKIES":       "29dd9a87-5bcc-4774-80c3-7f50d985068b",
  "DETROIT TIGERS":         "575c19b7-4052-41c2-9f0a-1c5813d02f99",
  "HOUSTON ASTROS":         "eb21dadd-8f10-4095-8bf3-dfb3b779f107",
  "KANSAS CITY ROYALS":     "833a51a9-0d84-410f-bd77-da08c3e5e26e",
  "LOS ANGELES ANGELS":     "4f735188-37c8-473d-ae32-1f7e34ccf892",
  "LOS ANGELES DODGERS":    "ef64da7f-cfaf-4300-87b0-9313386b977c",
  "MIAMI MARLINS":          "03556285-bdbb-4576-a06d-42f71f46ddc5",
  "MILWAUKEE BREWERS":      "dcfd5266-00ce-442c-bc09-264cd20cf455",
  "MINNESOTA TWINS":        "aa34e0ed-f342-4ec6-b774-c79b47b60e2d",
  "NEW YORK METS":          "f246a5e5-afdb-479c-9aaa-c68beeda7af6",
  "NEW YORK YANKEES":       "a09ec676-f887-43dc-bbb3-cf4bbaee9a18",
  "ATHLETICS":              "27a59d3b-ff7c-48ea-b016-4798f560f5e1",
  "PHILADELPHIA PHILLIES":  "2142e1ba-3b40-445c-b8bb-f1f8b1054220",
  "PITTSBURGH PIRATES":     "481dfe7e-5dab-46ab-a49f-9dcc2b6e2cfd",
  "SAN DIEGO PADRES":       "d52d5339-cbdd-43f3-9dfa-a42fd588b9a3",
  "SAN FRANCISCO GIANTS":   "a7723160-10b7-4277-a309-d8dd95a8ae65",
  "SEATTLE MARINERS":       "43a39081-52b4-4f93-ad29-da7f329ea960",
  "ST. LOUIS CARDINALS":    "44671792-dc02-4fdd-a5ad-f5f17edaa9d7",
  "TAMPA BAY RAYS":         "bdc11650-6f74-49c4-875e-778aeb7632d9",
  "TEXAS RANGERS":          "d99f919b-1534-4516-8e8a-9cd106c6d8cd",
  "TORONTO BLUE JAYS":      "1d678440-b4b1-4954-9b39-70afb3ebbcfa",
  "WASHINGTON NATIONALS":   "d89bed32-3aee-4407-99e3-4103641b999a",
};
const NBA_SR_ID = {
  "ATLANTA HAWKS":          "583ecb8f-fb46-11e1-82cb-f4ce4684ea4c",
  "BOSTON CELTICS":         "583eccfa-fb46-11e1-82cb-f4ce4684ea4c",
  "BROOKLYN NETS":          "583ec9d6-fb46-11e1-82cb-f4ce4684ea4c",
  "CHARLOTTE HORNETS":      "583ec97e-fb46-11e1-82cb-f4ce4684ea4c",
  "CHICAGO BULLS":          "583ec5fd-fb46-11e1-82cb-f4ce4684ea4c",
  "CLEVELAND CAVALIERS":    "583ec773-fb46-11e1-82cb-f4ce4684ea4c",
  "DALLAS MAVERICKS":       "583ecf50-fb46-11e1-82cb-f4ce4684ea4c",
  "DENVER NUGGETS":         "583ed102-fb46-11e1-82cb-f4ce4684ea4c",
  "DETROIT PISTONS":        "583ec928-fb46-11e1-82cb-f4ce4684ea4c",
  "GOLDEN STATE WARRIORS":  "583ec825-fb46-11e1-82cb-f4ce4684ea4c",
  "HOUSTON ROCKETS":        "583ecb3a-fb46-11e1-82cb-f4ce4684ea4c",
  "INDIANA PACERS":         "583ec7cd-fb46-11e1-82cb-f4ce4684ea4c",
  "LOS ANGELES CLIPPERS":   "583ecdfb-fb46-11e1-82cb-f4ce4684ea4c",
  "LOS ANGELES LAKERS":     "583ecae2-fb46-11e1-82cb-f4ce4684ea4c",
  "MEMPHIS GRIZZLIES":      "583eca88-fb46-11e1-82cb-f4ce4684ea4c",
  "MIAMI HEAT":             "583ecea6-fb46-11e1-82cb-f4ce4684ea4c",
  "MILWAUKEE BUCKS":        "583ecefd-fb46-11e1-82cb-f4ce4684ea4c",
  "MINNESOTA TIMBERWOLVES": "583eca2f-fb46-11e1-82cb-f4ce4684ea4c",
  "NEW ORLEANS PELICANS":   "583ecc9a-fb46-11e1-82cb-f4ce4684ea4c",
  "NEW YORK KNICKS":        "583ec70e-fb46-11e1-82cb-f4ce4684ea4c",
  "OKLAHOMA CITY THUNDER":  "583ecfff-fb46-11e1-82cb-f4ce4684ea4c",
  "ORLANDO MAGIC":          "583ed157-fb46-11e1-82cb-f4ce4684ea4c",
  "PHILADELPHIA 76ERS":     "583ec87d-fb46-11e1-82cb-f4ce4684ea4c",
  "PHOENIX SUNS":           "583ecfa8-fb46-11e1-82cb-f4ce4684ea4c",
  "PORTLAND TRAIL BLAZERS": "583ed056-fb46-11e1-82cb-f4ce4684ea4c",
  "SACRAMENTO KINGS":       "583ed0ac-fb46-11e1-82cb-f4ce4684ea4c",
  "SAN ANTONIO SPURS":      "583ecd4f-fb46-11e1-82cb-f4ce4684ea4c",
  "TORONTO RAPTORS":        "583ecda6-fb46-11e1-82cb-f4ce4684ea4c",
  "UTAH JAZZ":              "583ece50-fb46-11e1-82cb-f4ce4684ea4c",
  "WASHINGTON WIZARDS":     "583ec8d4-fb46-11e1-82cb-f4ce4684ea4c",
};
const NHL_SR_ID = {
  "ANAHEIM DUCKS":          "441862de-0f24-11e2-8525-18a905767e44",
  "ARIZONA COYOTES":        "4418464d-0f24-11e2-8525-18a905767e44",
  "BOSTON BRUINS":          "4416ba1a-0f24-11e2-8525-18a905767e44",
  "BUFFALO SABRES":         "4416d559-0f24-11e2-8525-18a905767e44",
  "CALGARY FLAMES":         "44159241-0f24-11e2-8525-18a905767e44",
  "CAROLINA HURRICANES":    "44182a9d-0f24-11e2-8525-18a905767e44",
  "CHICAGO BLACKHAWKS":     "4416272f-0f24-11e2-8525-18a905767e44",
  "COLORADO AVALANCHE":     "4415ce44-0f24-11e2-8525-18a905767e44",
  "COLUMBUS BLUE JACKETS":  "44167db4-0f24-11e2-8525-18a905767e44",
  "DALLAS STARS":           "44157522-0f24-11e2-8525-18a905767e44",
  "DETROIT RED WINGS":      "44169bb9-0f24-11e2-8525-18a905767e44",
  "EDMONTON OILERS":        "4415ea6c-0f24-11e2-8525-18a905767e44",
  "FLORIDA PANTHERS":       "4418464d-0f24-11e2-8525-18a905767e44",
  "LOS ANGELES KINGS":      "44151f7a-0f24-11e2-8525-18a905767e44",
  "MINNESOTA WILD":         "4416091c-0f24-11e2-8525-18a905767e44",
  "MONTREAL CANADIENS":     "441713b7-0f24-11e2-8525-18a905767e44",
  "NASHVILLE PREDATORS":    "441643b7-0f24-11e2-8525-18a905767e44",
  "NEW JERSEY DEVILS":      "44174b0c-0f24-11e2-8525-18a905767e44",
  "NEW YORK ISLANDERS":     "441766b9-0f24-11e2-8525-18a905767e44",
  "NEW YORK RANGERS":       "441781b9-0f24-11e2-8525-18a905767e44",
  "OTTAWA SENATORS":        "4416f5e2-0f24-11e2-8525-18a905767e44",
  "PHILADELPHIA FLYERS":    "44179d47-0f24-11e2-8525-18a905767e44",
  "PITTSBURGH PENGUINS":    "4417b7d7-0f24-11e2-8525-18a905767e44",
  "SAN JOSE SHARKS":        "44155909-0f24-11e2-8525-18a905767e44",
  "SEATTLE KRAKEN":         "1fb48e65-9688-4084-8868-02173525c3e1",
  "ST. LOUIS BLUES":        "441660ea-0f24-11e2-8525-18a905767e44",
  "TAMPA BAY LIGHTNING":    "4417d3cb-0f24-11e2-8525-18a905767e44",
  "TORONTO MAPLE LEAFS":    "441730a9-0f24-11e2-8525-18a905767e44",
  "VANCOUVER CANUCKS":      "4415b0a7-0f24-11e2-8525-18a905767e44",
  "VEGAS GOLDEN KNIGHTS":   "42376e1c-6da8-461e-9443-cfcf0a9fcc4d",
  "UTAH MAMMOTH":           "715a1dba-4e9f-4158-8346-3473b6e3557f",
  "WASHINGTON CAPITALS":    "4417eede-0f24-11e2-8525-18a905767e44",
  "WINNIPEG JETS":          "44180e55-0f24-11e2-8525-18a905767e44",
};

function maps(league) {
  if (league === 'MLB') return { TRI: MLB_TRI, NICK: MLB_NICK, MARKET: MLB_MARKET, SR_ID: MLB_SR_ID };
  if (league === 'NHL') return { TRI: NHL_TRI, NICK: NHL_NICK, MARKET: NHL_MARKET, SR_ID: NHL_SR_ID };
  return { TRI: NBA_TRI, NICK: NBA_NICK, MARKET: NBA_MARKET, SR_ID: NBA_SR_ID };
}

// === Normalizers & helpers ===================================================
const CORE_EMPTY = {
  OddsTeamAway:'N/A', OddsTeamHome:'N/A',
  OddsAwayMarket:'N/A', OddsHomeMarket:'N/A',
  OddsAwayNickname:'N/A', OddsHomeNickname:'N/A',
  OddsAwayTri:'N/A', OddsHomeTri:'N/A',
  MoneylineAwayOdds:'N/A', MoneylineHomeOdds:'N/A',
  SpreadAwayHandicap:'N/A', SpreadAwayOdds:'N/A',
  SpreadHomeHandicap:'N/A', SpreadHomeOdds:'N/A',
  OverUnderTotalLine:'N/A', OverUnderTotalLineDuplicate:'N/A',
  OverOdds:'N/A', UnderOdds:'N/A'
};
const american = (v) => (typeof v === 'string' && v.trim() !== '') ? v : 'N/A';

function stripAccents(str) {
  if (typeof str !== 'string') return '';
  return str.normalize("NFD").replace(/[\u0300-\u036f]/g, "");
}

/**
 * Calculates the Levenshtein distance between two strings.
 * Lower distance means more similar strings.
 * @param {string} s1 First string
 * @param {string} s2 Second string
 * @returns {number} The edit distance
 */
function levenshteinDistance(s1, s2) {
  s1 = s1.toLowerCase();
  s2 = s2.toLowerCase();

  const costs = [];
  for (let i = 0; i <= s1.length; i++) {
    let lastValue = i;
    for (let j = 0; j <= s2.length; j++) {
      if (i === 0) {
        costs[j] = j;
      } else {
        if (j > 0) {
          let newValue = costs[j - 1];
          if (s1.charAt(i - 1) !== s2.charAt(j - 1)) {
            newValue = Math.min(Math.min(newValue, lastValue), costs[j]) + 1;
          }
          costs[j - 1] = lastValue;
          lastValue = newValue;
        }
      }
    }
    if (i > 0) {
      costs[s2.length] = lastValue;
    }
  }
  return costs[s2.length];
}

async function buildMasterRoster(league, eventId, teams, detail) {
    const rosterMap = new Map(); // Key: NORMALIZED name, Value: { details object }
    const gameCacheDir = path.join(CACHE_DIR, league, String(eventId));

    // Helper to add a player to the map, avoiding duplicates
    const addPlayer = (playerData) => {
        if (!playerData || !playerData.fullName) return;
        const normalizedName = stripAccents(playerData.fullName).toUpperCase();
        if (!rosterMap.has(normalizedName)) {
            rosterMap.set(normalizedName, playerData);
        }
    };

    // 1. Parse and add players from cached Sportradar files (HIGHEST PRIORITY SOURCE)
    const parseAndAddSportradarPlayers = async (filePath, triCode) => {
        const rosterData = await readJSON(filePath, null);
        if (!rosterData || !rosterData.players) return;

        for (const player of rosterData.players) {
            const fullName = player.full_name;
            if (fullName) {
                addPlayer({
                    fullName: fullName,
                    firstName: player.first_name || '',
                    lastName: player.last_name || '',
                    triCode: triCode,
                    position: player.primary_position || 'N/A',
                    jersey: String(player.jersey_number || 'N/A'),
                });
            }
        }
    };

    if (teams.awayTri !== 'N/A') {
        await parseAndAddSportradarPlayers(path.join(gameCacheDir, `awayRoster${teams.awayTri}.json`), teams.awayTri);
    }
    if (teams.homeTri !== 'N/A') {
        await parseAndAddSportradarPlayers(path.join(gameCacheDir, `homeRoster${teams.homeTri}.json`), teams.homeTri);
    }

    // 2. Add players from FanDuel's /roster endpoint (fills in gaps)
    try {
        const data = await j(`${API_BASE}/v1/leagues/${league}/events/${eventId}/roster`);
        for (const team of data?.teams || []) {
            const triCode = team?.team?.triCode || 'N/A';
            for (const player of team?.players || []) {
                addPlayer({
                    fullName: `${player.firstName} ${player.lastName}`.trim(),
                    firstName: player.firstName,
                    lastName: player.lastName,
                    triCode: triCode,
                    position: player.position || 'N/A',
                    jersey: String(player.jerseyNumber || 'N/A'),
                });
            }
        }
    } catch (e) {
        console.error(`[roster] Failed to fetch FanDuel roster for ${league}:${eventId}`, e.message);
    }

    return rosterMap;
}


async function fetchAndCacheSportradarRosters(league, eventId, teams) {
    const teamMaps = maps(league);
    const apiKey = getSportradarApiKey(league);
    if (!apiKey) return;

    const awayTeamId = teamMaps.SR_ID[teams.awayFull.toUpperCase()];
    const homeTeamId = teamMaps.SR_ID[teams.homeFull.toUpperCase()];

    const gameCacheDir = path.join(CACHE_DIR, league, String(eventId));

    const getTeamRoster = async (teamId, teamTri, side) => {
        if (!teamId || !teamTri || teamTri === 'N/A') return;
        const fileName = `${side}Roster${teamTri}.json`;
        const filePath = path.join(gameCacheDir, fileName);
        if (await existsPath(filePath)) {
            return;
        }

        try {
            const version = league === 'NHL' ? 'v3' : 'v7';
            const url = `${SR_API_BASE}/${league.toLowerCase()}/production/${version}/en/teams/${teamId}/profile.json?api_key=${apiKey}`;
            const rosterData = await (await fetch(url)).json();
            await writeJSON(filePath, rosterData);
            console.log(`[roster] ✅ Wrote ${fileName} for ${league}:${eventId}`);
        } catch (e) {
            console.error(`[roster] ❌ Failed to fetch Sportradar roster for team ${teamId}:`, e.message);
        }
    };

    await Promise.all([
        getTeamRoster(awayTeamId, teams.awayTri, 'away'),
        getTeamRoster(homeTeamId, teams.homeTri, 'home')
    ]);
}


function normalizeRunner(name, rosterData, teamMaps) {
    // --- Existing initial setup ---
    const originalName = name; // Keep the original name for the final result
    const defaultResult = {
        original: originalName,
        fullName: '',
        firstName: '', lastName: '', jersey: '', position: '', triCode: ''
    };

    // --- NEW: Strip common suffixes BEFORE normalization ---
    // Use regex to remove common suffixes like " Over", " Under", " Yes", " No" at the end of the string
    // Allows for optional space before the suffix. Handles variations like "O 1.5", "U 2.5" by removing O/U prefix later
    let cleanedName = name.replace(/\s+(?:OVER|UNDER|YES|NO)$/i, '').trim();

    // If the name itself is just Over/Under/Yes/No (or similar generic), return default early.
    // Handles cases where the market name might be a player, but selection name is just "Over".
     if (/^(?:OVER|UNDER|YES|NO)$/i.test(cleanedName)) {
         // Return default for now; the logic in splitProps will handle assigning market entity later if appropriate.
         return defaultResult;
     }

    // --- Normalize the CLEANED name ---
    const nameUpper = stripAccents(cleanedName).toUpperCase();
    // --- END NEW ---

    // --- 1. Generic Check (Using nameUpper - the cleaned/normalized name) ---
    const genericTerms = [
        // No need for OVER/UNDER here as they are handled above or should be associated with a player
         'TIE', 'ANY OTHER', 'HITS RECORDED',
        'RUNS OR MORE', 'RUNS', 'ALTERNATE TOTAL', 'WINNING MARGIN',
        'CORRECT SCORE', 'INNING RESULT', 'INNING HITS', 'INNING CORRECT'
    ];
    const genericRegex = new RegExp(`\\b(${genericTerms.join('|')})\\b`, 'i');
    if (genericRegex.test(nameUpper) || /^\d+-\d+$/.test(nameUpper) || /^\d+\+?\s/.test(nameUpper) ) {
        return defaultResult;
    }

    // --- 2. Player Check using Levenshtein (Using nameUpper) ---
    let bestMatch = null;
    let lowestDistance = Infinity;
    const MAX_ACCEPTABLE_DISTANCE = 4; // Allow up to 4 edits (Increased from 2)

    for (const [playerKeyUpper, playerData] of rosterData.entries()) {
        // Check 2a: Exact Match (Fastest)
        if (nameUpper === playerKeyUpper) {
            bestMatch = playerData;
            lowestDistance = 0;
            break; // Found perfect match, stop searching
        }

        // Check 2b: Calculate Levenshtein Distance
        const distance = levenshteinDistance(nameUpper, playerKeyUpper);

        // Check 2c: Heuristic for Length and Distance
        const lengthDifference = Math.abs(nameUpper.length - playerKeyUpper.length);
        if (distance <= MAX_ACCEPTABLE_DISTANCE && distance < lowestDistance) {
            if (distance === 0 || lengthDifference <= distance + 1) { // Allow length diff up to distance + 1
                 lowestDistance = distance;
                 bestMatch = playerData;
            }
        }
         // Optional: Handle simple initial matching if Levenshtein doesn't catch it
         // (Could be added here if needed, comparing first initial + last name)
    }

    if (bestMatch) {
        return {
            original: originalName, // Return the ORIGINAL name here
            fullName: bestMatch.fullName,
            firstName: bestMatch.firstName,
            lastName: bestMatch.lastName,
            jersey: bestMatch.jersey,
            position: bestMatch.position,
            triCode: bestMatch.triCode
        };
    }
    // --- End Player Check ---


    // --- 3. Team Check (Using nameUpper) ---
    const sortedTeamKeys = Object.keys(teamMaps.TRI).sort((a, b) => b.length - a.length);
    for (const teamKey of sortedTeamKeys) {
        const teamRegex = new RegExp(`\\b${teamKey.toUpperCase()}\\b`);
        if (teamRegex.test(nameUpper)) {
            return {
                original: originalName, // Return the ORIGINAL name here
                fullName: teamKey,
                firstName: teamMaps.MARKET[teamKey] || '',
                lastName: teamMaps.NICK[teamKey] || '',
                jersey: '',
                position: '',
                triCode: teamMaps.TRI[teamKey] || ''
            };
        }
    }

    // --- 4. Final Fallback ---
    // console.warn(`[normalizeRunner] No match found for: "${originalName}" (Cleaned: "${cleanedName}", Normalized: "${nameUpper}")`);
    return defaultResult; // Contains the original name
}

function pickCore(markets) {
  const out = { ...CORE_EMPTY };
  if (!Array.isArray(markets)) return out;

  const ml = markets.find(m => m.marketType === 'MONEY_LINE');
  if (ml?.selections) {
    const a = ml.selections.find(s => (s.resultType||'').toUpperCase() === 'AWAY');
    const h = ml.selections.find(s => (s.resultType||'').toUpperCase() === 'HOME');
    if (a) out.MoneylineAwayOdds = american(a.odds);
    if (h) out.MoneylineHomeOdds = american(h.odds);
  }

  const spread =
    markets.find(m => ['MATCH_HANDICAP_(2-WAY)','SPREAD','RUN_LINE','PUCK_LINE','MATCH_HANDICAP'].includes(m.marketType)) ||
    markets.find(m => /RUN LINE|PUCK LINE|HANDICAP/i.test(m.marketName||''));
  if (spread?.selections) {
    const a = spread.selections.find(s => (s.resultType||'').toUpperCase() === 'AWAY');
    const h = spread.selections.find(s => (s.resultType||'').toUpperCase() === 'HOME');
    if (a) { out.SpreadAwayHandicap = (a.formattedHandicap ?? a.handicap ?? 'N/A')+''; out.SpreadAwayOdds = american(a.odds); }
    if (h) { out.SpreadHomeHandicap = (h.formattedHandicap ?? h.handicap ?? 'N/A')+''; out.SpreadHomeOdds = american(h.odds); }
  }

  const total =
    markets.find(m => ['TOTAL_POINTS_(OVER/UNDER)','TOTAL','TOTAL_GOALS','TOTAL_RUNS','TOTAL_POINTS'].includes(m.marketType)) ||
    markets.find(m => /TOTAL/i.test(m.marketName||''));
  if (total?.selections) {
    const o = total.selections.find(s => (s.resultType||'').toUpperCase() === 'OVER'  || /over/i.test(s.runnerName||''));
    const u = total.selections.find(s => (s.resultType||'').toUpperCase() === 'UNDER' || /under/i.test(s.runnerName||''));
    const line = o?.formattedHandicap ?? u?.formattedHandicap ?? o?.handicap ?? u?.handicap ?? 'N/A';
    const norm = (String(line)).replace(/^O\s*|^U\s*/i, '').trim() || 'N/A';
    out.OverUnderTotalLine = norm;
    out.OverUnderTotalLineDuplicate = norm;
    if (o) out.OverOdds = american(o.odds);
    if (u) out.UnderOdds = american(u.odds);
  }
  return out;
}

function deriveTeams(league, markets) {
  const ml = (markets||[]).find(m => m.marketType === 'MONEY_LINE');
  let away = '', home = '';
  if (ml?.selections) {
    const a = ml.selections.find(s => (s.resultType||'').toUpperCase() === 'AWAY');
    const h = ml.selections.find(s => (s.resultType||'').toUpperCase() === 'HOME');
    away = (a?.runnerName || '').toUpperCase();
    home = (h?.runnerName || '').toUpperCase();
  }
  const m = maps(league);
  return {
    awayFull: away || 'N/A', homeFull: home || 'N/A',
    awayTri: m.TRI[away] || 'N/A', homeTri: m.TRI[home] || 'N/A',
    awayNick: m.NICK[away] || 'N/A', homeNick: m.NICK[home] || 'N/A',
    awayMarket: m.MARKET[away] ?? 'N/A',
    homeMarket: m.MARKET[home] ?? 'N/A'
  };
}

function splitProps(markets, masterRoster, league) {
    const byM = {}, byP = {};
    const teamMaps = maps(league);

    // More specific set of core market types to exclude from props
    const CORE_MARKET_TYPES = new Set([
        'MONEY_LINE',
        'MATCH_HANDICAP_(2-WAY)',
        'SPREAD',
        'RUN_LINE',
        'PUCK_LINE',
        'MATCH_HANDICAP', // Keep consistent with pickCore
        'TOTAL_POINTS_(OVER/UNDER)',
        'TOTAL',
        'TOTAL_GOALS',
        'TOTAL_RUNS',
        'TOTAL_POINTS'
    ]);

    for (const m of markets || []) {
        const marketTypeUpper = (m.marketType || '').toUpperCase();
        const marketNameLower = (m.marketName || '').toLowerCase();

        // --- REVISED CORE CHECK ---
        // 1. Skip if the marketType EXACTLY matches one of the core types
        let isCore = CORE_MARKET_TYPES.has(marketTypeUpper);

        // 2. Skip if pickCore might identify it by NAME and it wasn't caught by TYPE above.
        // This prevents double processing of markets sometimes identified by name in pickCore.
        if (!isCore) {
             // Check spread by name regex (only if not already matched by known spread types)
             if (/(?:^|\s)(run line|puck line|handicap)(?:$|\s)/i.test(marketNameLower) &&
                 !['MATCH_HANDICAP_(2-WAY)','SPREAD','RUN_LINE','PUCK_LINE','MATCH_HANDICAP'].includes(marketTypeUpper)) {
                 // Check if it's NOT an alternate or period-specific market name
                 if (!/alternate|first|1st|2nd|3rd|4th|5th|6th|7th|8th|9th|inning|half|quarter|period/i.test(marketNameLower)) {
                    isCore = true;
                 }
             }
             // Check total by name regex (only if not already matched by known total types)
             else if (/(?:^|\s)total(?:$|\s)/i.test(marketNameLower) &&
                 !['TOTAL_POINTS_(OVER/UNDER)','TOTAL','TOTAL_GOALS','TOTAL_RUNS','TOTAL_POINTS'].includes(marketTypeUpper)) {
                  // Check if it's NOT an alternate or period-specific market name
                 if (!/alternate|first|1st|2nd|3rd|4th|5th|6th|7th|8th|9th|inning|half|quarter|period/i.test(marketNameLower)) {
                    isCore = true;
                 }
             }
        }

        if (isCore) {
             // console.log(`Skipping core market: ${m.marketName} (Type: ${marketTypeUpper})`); // Optional debug log
             continue;
        }
        // --- END REVISED CORE CHECK ---


        const key = m.marketName || m.marketType || 'Other';
        if (!byM[key]) byM[key] = [];

        // Try to identify a player/team from the market name itself
        const marketEntity = normalizeRunner(m.marketName, masterRoster, teamMaps);

        for (const s of m.selections || []) {
            // Normalize the runner name from the selection
            let runnerDetails = normalizeRunner(s.runnerName, masterRoster, teamMaps);

            // If the selection runner is generic ('Over', 'Under', 'Yes', 'No', specific scores like '1-0')
            // AND the market name identified a specific player (not a team), use the market's player details.
            // Check runnerDetails.original specifically for generic terms after cleaning in normalizeRunner
            if (/^(?:OVER|UNDER|YES|NO)$/i.test(runnerDetails.original) && marketEntity.fullName && !teamMaps.TRI[marketEntity.fullName.toUpperCase()]) {
                 runnerDetails = { ...marketEntity, original: s.runnerName }; // Keep original runnerName from API here
            }

            // --- REVISED Line Extraction Logic (with 0.0 vs 0 distinction) ---
            let lineValue = 'N/A'; // Default to N/A
            const rawNumericHandicap = s.handicap; // Store original numeric value

            if (s.formattedHandicap && s.formattedHandicap !== 'N/A' && s.formattedHandicap !== '') {
                 // Remove potential Over/Under prefixes from formattedHandicap
                 lineValue = String(s.formattedHandicap).replace(/^[OU]\s*/i, '').trim();
            // --- MODIFIED CONDITION BELOW ---
            } else if (rawNumericHandicap !== undefined && rawNumericHandicap !== null && rawNumericHandicap !== 0.0) { // Exclude specifically 0.0, but allow integer 0
                 lineValue = String(rawNumericHandicap);
            // --- END MODIFIED CONDITION ---
            }
            // Additional check for runnerName containing Over/Under line, often used when handicap fields are missing/zero
            // Make sure the runnerDetails.original (which might be just 'Over' or 'Under') wasn't matched above
            else if (!/^(?:OVER|UNDER|YES|NO)$/i.test(runnerDetails.original) && /^(O|U)\s?(\d+(\.\d+)?)$/i.test(s.runnerName || '')) {
                 const match = (s.runnerName || '').match(/^(O|U)\s?(\d+(\.\d+)?)$/i);
                 if (match && match[2]) {
                     lineValue = match[2];
                 }
            }
             // Ensure lineValue is 'N/A' if it's empty after processing
            if (lineValue === '') lineValue = 'N/A';
            // --- END REVISED Line Extraction ---


            const propData = {
                selectionId: s.selectionId,
                runnerDetails: runnerDetails, // Contains original API runnerName
                odds: s.odds ?? 'N/A',
                line: lineValue // Use the determined line value ("N/A" if not applicable/extractable)
            };

            byM[key].push(propData);

            // Populate propsByPlayer, OMITTING TEAMS and GENERIC selections
            // We use runnerDetails.fullName which comes from roster/team match (not the potentially generic original name)
            if (runnerDetails.fullName) {
                const playerKey = runnerDetails.fullName;
                // Check if the identified full name is a key in the TRI map (i.e., it's a team name)
                const isTeam = !!teamMaps.TRI[playerKey.toUpperCase()];

                if (!isTeam) { // Only add if it's NOT a team
                    if (!byP[playerKey]) byP[playerKey] = [];
                    byP[playerKey].push({
                        market: key,
                        ...propData
                    });
                }
            }
        }
    }
    return { propsByMarket: byM, propsByPlayer: byP };
}


// === State & Control =========================================================
let watchlist = [];
const eventsMeta     = new Map();
const eventsActive   = new Map();
const pregameCaptured= new Set();
const kKey = (league, id) => `${league}:${id}`;

const ODDS_KEYS = ['MoneylineAwayOdds','MoneylineHomeOdds','SpreadAwayHandicap','SpreadAwayOdds','SpreadHomeHandicap','SpreadHomeOdds','OverUnderTotalLine','OverUnderTotalLineDuplicate','OverOdds','UnderOdds'];
const TEAM_KEYS = ['OddsTeamAway','OddsTeamHome','OddsAwayMarket','OddsHomeMarket','OddsAwayNickname','OddsHomeNickname','OddsAwayTri','OddsHomeTri'];

function blankOddsKeepTeams(coreObj) {
  const out = { ...CORE_EMPTY };
  for (const k of TEAM_KEYS) {
    if (coreObj && Object.prototype.hasOwnProperty.call(coreObj, k)) out[k] = coreObj[k];
  }
  for (const k of ODDS_KEYS) out[k] = 'N/A';
  return out;
}

async function scrubOddsFile(file) {
  try {
    const cur = await readJSON(file, null);
    const next = blankOddsKeepTeams(cur || {});
    await writeJSON(file, next);
  } catch {
    await writeJSON(file, CORE_EMPTY);
  }
}

async function blankAllOddsAndProps() {
  for (const league of LEAGUES) {
    const lDir = path.join(CACHE_DIR, league);
    await ensureDir(lDir);
    const items = await fsp.readdir(lDir, { withFileTypes: true });
    for (const d of items) {
      if (!d.isDirectory() || !/^\d+$/.test(d.name)) continue;
      const idDir = path.join(lDir, d.name);
      await scrubOddsFile(path.join(idDir, 'liveCore.json'));
      const stamp = { _updatedAt: new Date().toISOString(), status: 'UNAVAILABLE' };
      await writeJSON(path.join(idDir, 'propsByMarket.json'), stamp);
      await writeJSON(path.join(idDir, 'propsByPlayer.json'), stamp);
    }
  }
}

async function isPausedFile() { return await existsPath(STOP_FLAG); }

async function applyPauseState() {
  const wantPause = await isPausedFile();
  if (wantPause && !PAUSED) {
    PAUSED = true;
    await blankAllOddsAndProps();
    console.log('[pause] STOP.flag detected -> blanked odds/props (teams preserved)');
  } else if (!wantPause && PAUSED) {
    PAUSED = false;
    console.log('[pause] STOP.flag removed -> resuming');
  }
}

// === Pollers ================================================================
async function pollWatchlist() {
  const arr = (await readJSON(WATCHLIST_FILE, []))
    .map(x => ({ league: String(x.league||'').toUpperCase(), eventId: Number(x.eventId) }))
    .filter(x => x.league && x.eventId);
  if (JSON.stringify(arr) !== JSON.stringify(watchlist)) {
    watchlist = arr;
    for (const w of watchlist) {
      const d = path.join(CACHE_DIR, w.league, String(w.eventId));
      await ensureDir(d);
      const preFile = path.join(d, 'pregameCore.json');
      if (!(await existsPath(preFile))) await writeJSON(preFile, CORE_EMPTY);
    }
    console.log('[watchlist]', watchlist);
  }
}

async function sweepStaleLeagueDirs(league, activeIds) {
  const lDir = path.join(CACHE_DIR, league);
  await ensureDir(lDir);
  const items = await fsp.readdir(lDir, { withFileTypes: true });
  for (const d of items) {
    if (d.isDirectory() && /^\d+$/.test(d.name)) {
      const id = Number(d.name);
      if (!activeIds.has(id)) {
        try {
          await fsp.rm(path.join(lDir, d.name), { recursive: true, force: true });
          pregameCaptured.delete(kKey(league, id));
          console.log('[cleanup]', league, id);
        } catch (e) {
          console.error('[cleanup error]', league, id, e.message);
        }
      }
    }
  }
}

const evListURL = (league) => `${API_BASE}/v1/leagues/${league}/events`;

async function updateEvents() {
  if (PAUSED) return;
  for (const league of LEAGUES) {
    let list = null;
    try {
      list = await j(evListURL(league));
      await writeJSON(path.join(CACHE_DIR, league, 'events.json'), list);

      const active = new Set();
      const byId = new Map();

      for (const ev of list || []) {
        active.add(ev.eventId);
        byId.set(ev.eventId, ev);
        const dir = path.join(CACHE_DIR, league, String(ev.eventId));
        await ensureDir(dir);
        const preFile = path.join(dir, 'pregameCore.json');
        if (!(await existsPath(preFile))) await writeJSON(preFile, CORE_EMPTY);
      }

      eventsActive.set(league, active);
      for (const id of active) {
        const ev = byId.get(id);
        const inPlay = !!(ev?.markets||[]).some(m => m.inPlay);
        eventsMeta.set(kKey(league, id), { openDate: ev?.openDate || null, inPlay });
      }
      await sweepStaleLeagueDirs(league, active);
      console.log('[events]', league, list?.length ?? 0, 'active');
    } catch (e) {
      console.error('[events]', league, e.message);
      await writeJSON(path.join(CACHE_DIR, league, 'events.json'), []);
      await blankAllOddsAndProps();
    }
  }
}

async function writeAllForDetail(league, id, detail, rosterData, doPulse=true, doProps=true) {
  const dir = path.join(CACHE_DIR, league, String(id));
  const inPlay = (detail.markets||[]).some(m => m.inPlay);
  const open   = detail.openDate || eventsMeta.get(kKey(league, id))?.openDate || null;
  eventsMeta.set(kKey(league, id), { openDate: open, inPlay });

  const core = pickCore(detail.markets||[]);
  const tm   = deriveTeams(league, detail.markets||[]);
  Object.assign(core, {
    OddsTeamAway: tm.awayFull, OddsTeamHome: tm.homeFull,
    OddsAwayTri: tm.awayTri, OddsHomeTri: tm.homeTri,
    OddsAwayNickname: tm.awayNick, OddsHomeNickname: tm.homeNick,
    OddsAwayMarket: tm.awayMarket, OddsHomeMarket: tm.homeMarket
  });
  await writeJSON(path.join(dir, 'liveCore.json'), core);

  if (doProps) {
    try {
      const { propsByMarket, propsByPlayer } = splitProps(detail.markets || [], rosterData, league);
      await writeJSON(path.join(dir, 'propsByMarket.json'), { _updatedAt: new Date().toISOString(), ...propsByMarket });
      await writeJSON(path.join(dir, 'propsByPlayer.json'), { _updatedAt: new Date().toISOString(), ...propsByPlayer });
    } catch (e) {
      console.error('[props write error]', league, id, e.message);
      const errObj = { _updatedAt: new Date().toISOString(), status: 'ERROR' };
      await writeJSON(path.join(dir, 'propsByMarket.json'), errObj);
      await writeJSON(path.join(dir, 'propsByPlayer.json'), errObj);
    }
  }

  if (doPulse) {
    try {
      const pulse = await j(`${API_BASE}/v1/leagues/${league}/events/${id}/pulse?idProvider=${ID_PROVIDER}`);
      await writeJSON(path.join(dir, 'pulse.json'), pulse);
    } catch { /* ignore pulse errors */ }
  }
}

async function tryPregame(league, id) {
  const key = kKey(league, id);
  if (pregameCaptured.has(key)) return;

  const dir = path.join(CACHE_DIR, league, String(id));
  if (!(await existsPath(dir))) return;

  const preFile = path.join(dir, 'pregameCore.json');
  if (!(await existsPath(preFile))) await writeJSON(preFile, CORE_EMPTY);

  try {
    const pre = await j(`${API_BASE}/v1/leagues/${league}/events/${id}/pre-game?idProvider=${ID_PROVIDER}`);
    const preCore = pickCore(pre.markets||[]);
    const preTm   = deriveTeams(league, pre.markets||[]);
    Object.assign(preCore, {
      OddsTeamAway: preTm.awayFull, OddsTeamHome: preTm.homeFull,
      OddsAwayTri: preTm.awayTri, OddsHomeTri: preTm.homeTri,
      OddsAwayNickname: preTm.awayNick, OddsHomeNickname: preTm.homeNick,
      OddsAwayMarket: preTm.awayMarket, OddsHomeMarket: preTm.homeMarket
    });
    await writeJSON(preFile, preCore);
    pregameCaptured.add(key);
    console.log('[pregame captured]', key);
  } catch {}
}

function partitionWork() {
  const fast = [], slow = [];
  const watchSet = new Set(watchlist.map(w => kKey(w.league, w.eventId)));
  for (const league of LEAGUES) {
    const active = eventsActive.get(league);
    if (!active) continue;
    for (const id of active) {
      const key = kKey(league, id);
      const meta = eventsMeta.get(key) || {};
      const isInPlay = !!meta.inPlay;
      if (isInPlay || watchSet.has(key)) fast.push({ league, eventId: id });
      else slow.push({ league, eventId: id });
    }
  }
  return { fast, slow };
}

async function tickFast() {
  if (PAUSED) return;
  const { fast } = partitionWork();
  for (const w of fast) {
    try {
      const detail = await j(`${API_BASE}/v1/leagues/${w.league}/events/${w.eventId}?marketCategories=PLAYER_PROPS&marketCategories=GAME_PROPS&marketCategories=CORE&idProvider=${ID_PROVIDER}`);
      const teams = deriveTeams(w.league, detail.markets || []);
      await fetchAndCacheSportradarRosters(w.league, w.eventId, teams);
      const rosterData = await buildMasterRoster(w.league, w.eventId, teams, detail);
      await writeAllForDetail(w.league, w.eventId, detail, rosterData, true, true);
      await tryPregame(w.league, w.eventId);
    } catch (e) {
      console.error('[fast]', w.league, w.eventId, e.message);
      const dir = path.join(CACHE_DIR, w.league, String(w.eventId));
      await scrubOddsFile(path.join(dir, 'liveCore.json'));
      const stamp = { _updatedAt: new Date().toISOString(), status: 'UNAVAILABLE' };
      await writeJSON(path.join(dir, 'propsByMarket.json'), stamp);
      await writeJSON(path.join(dir, 'propsByPlayer.json'), stamp);
    }
  }
}

async function tickSlow() {
  if (PAUSED) return;
  const { slow } = partitionWork();
  for (const w of slow) {
    try {
      const detail = await j(`${API_BASE}/v1/leagues/${w.league}/events/${w.eventId}?marketCategories=PLAYER_PROPS&marketCategories=GAME_PROPS&marketCategories=CORE&idProvider=${ID_PROVIDER}`);
      const teams = deriveTeams(w.league, detail.markets || []);
      await fetchAndCacheSportradarRosters(w.league, w.eventId, teams);
      const rosterData = await buildMasterRoster(w.league, w.eventId, teams, detail);
      await writeAllForDetail(w.league, w.eventId, detail, rosterData, false, false);
      await tryPregame(w.league, w.eventId);
    } catch (e) {
      console.error('[slow]', w.league, w.eventId, e.message);
      const dir = path.join(CACHE_DIR, w.league, String(w.eventId));
      await scrubOddsFile(path.join(dir, 'liveCore.json'));
      const stamp = { _updatedAt: new Date().toISOString(), status: 'UNAVAILABLE' };
      await writeJSON(path.join(dir, 'propsByMarket.json'), stamp);
      await writeJSON(path.join(dir, 'propsByPlayer.json'), stamp);
    }
  }
}

async function pruneLogs() {
  try {
    await ensureDir(LOG_DIR);
    const files = await fsp.readdir(LOG_DIR);
    const now = Date.now();
    for (const f of files) {
      if (!/\.log$/i.test(f)) continue;
      const full = path.join(LOG_DIR, f);
      const st = await fsp.stat(full);
      if (now - st.mtimeMs > 7 * 24 * 60 * 60 * 1000) {
        await fsp.unlink(full).catch(()=>{});
        continue;
      }
      if (st.size > 5 * 1024 * 1024) {
        try {
          const fd = await fsp.open(full, 'r');
          const keep = 1 * 1024 * 1024;
          const start = Math.max(0, st.size - keep);
          const buf = Buffer.alloc(Math.min(keep, st.size));
          await fd.read(buf, 0, buf.length, start);
          await fd.close();
          await fsp.writeFile(full, buf);
        } catch {}
      }
    }
  } catch {}
}

// === Bootstrap ===============================================================
(async () => {
  await ensureDir(CACHE_DIR);
  await ensureDir(LOG_DIR);
  for (const lg of LEAGUES) await ensureDir(path.join(CACHE_DIR, lg));
  await ensureDir(CONFIG_DIR);

  console.log('[MSSG] base:', BASE_DIR);
  if (!FD_API_KEY) {
    console.warn('[MSSG] FD_API_KEY is empty. Place it in .env (FD_API_KEY=your_key). Requests will fail/blank until set.');
  }

  await applyPauseState();
  await pollWatchlist();
  await updateEvents();
  await tickFast();
  await tickSlow();

  setInterval(() => applyPauseState().catch(()=>{}), 1000);
  setInterval(() => pollWatchlist().catch(()=>{}), WATCHLIST_MS);
  setInterval(() => updateEvents().catch(()=>{}), EVENTS_POLL_MS);
  setInterval(() => tickFast().catch(()=>{}),     TICK_FAST_MS);
  setInterval(() => tickSlow().catch(()=>{}),     TICK_SLOW_MS);
  setInterval(() => pruneLogs().catch(()=>{}),    60_000);
})();