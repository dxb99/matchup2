const SHEET_PLAYERS = 'Players';
const SHEET_CURRENT = 'CurrentMatchup';
const SHEET_HISTORY = 'History';
const SHEET_MATCH_IDS = 'MatchIDs';
const SHEET_RATINGS_VOTES = 'RatingsVotes';
const SHEET_RATING_CODES = 'RatingCodes';
const SHEET_PLAYER_RATINGS = 'PlayerRatings';
const RATINGS_TEST_MODE = false;
const RATINGS_TIME_ZONE = 'Asia/Dubai';
const RATINGS_TIME_ZONE_LABEL = 'GST';

function isValidAdmin(password){
  return password === "UT4L!FE";
}

function doGet() {
  return ContentService
    .createTextOutput(JSON.stringify({ ok: true }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {

  const data = JSON.parse(e.postData.contents);
  const action = data.action;

  if(action==="getInitialData") return json(getInitialData());
  if(action==="saveMatchupDirect") return json(saveMatchup(data));
  if(action==="getHistory") return json(getHistory(data));
  if(action==="clearHistory") return json(clearHistory(data));
  if(action==="getPlayersAdmin") return json(getPlayersAdmin());
  if(action==="savePlayersAdmin") return json(savePlayersAdmin(data));
  if(action==="addRatingPlayer") return json(addRatingPlayer(data));
  if(action==="deactivateRatingPlayer") return json(deactivateRatingPlayer(data));
  if(action==="getSessionMaps") return json(getSessionMaps());
  if(action==="getCustomSession") return json(getCustomSession());
  if(action==="saveCustomSession") return json(saveCustomSession(data));
  if(action==="clearCustomSession") return json(clearCustomSession(data));
  if(action==="saveGlobalMapMatchMaker") return json(saveGlobalMapMatchMaker(data));
  if(action==="generateSessionMaps") return json(runGenerateSessionMaps(data));
  if(action==="clearSessionMaps") return json(runClearSessionMaps(data));
  if(action==="saveSessionProgress") return json(runSaveSessionProgress(data));
  if(action==="deleteSessionMap") return json(deleteSessionMap(data));
  if(action==="copySessionMaps") return json(copySessionMaps());
  if(action==="getRatingStatus") return json(getRatingStatus(data));
  if(action==="setupRatingSheets") return json(setupRatingSheets(data));
  if(action==="setManualVotingWindow") return json(setManualVotingWindow(data));
  if(action==="requestRatingCode") return json(requestRatingCode(data));
  if(action==="submitRatings") return json(submitRatings(data));
  if(action==="applyLatestRatingsToPlayers") return json(applyLatestRatingsToPlayers(data));
  if(action==="verifyAdminPassword") return json(verifyAdminPassword(data));
  
  return json({ok:false,error:"Unknown action"});
}

function json(obj){
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function getSheet(name){
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
}

function getOrCreateMID(redTeam, blueTeam){

  const sheet = getSheet(SHEET_MATCH_IDS);

  // Sort each team
const redSorted = redTeam.slice().sort();
const blueSorted = blueTeam.slice().sort();

// Create two possible keys (to handle swapped teams)
const key1 = redSorted.join("|") + "||" + blueSorted.join("|");
const key2 = blueSorted.join("|") + "||" + redSorted.join("|");

// Always store the same direction (alphabetically smaller one)
const key = key1 < key2 ? key1 : key2;

  const data = sheet.getDataRange().getValues();

  // 🔥 Check existing
  for(let i = 1; i < data.length; i++){
    if(data[i][0] === key){
      return data[i][1].toString().replace("MID_", "");
    }
  }

  // 🔥 Find HIGHEST existing MID
let lastMID = 0;

for(let i = 1; i < data.length; i++){

  const storedMID = data[i][1];

  if(!storedMID) continue;

  const num = parseInt(
    storedMID.toString().replace("MID_", "")
  );

  if(num > lastMID){
    lastMID = num;
  }

}

  const nextId = lastMID + 1;
  const formattedMID = String(nextId).padStart(4, "0");

  // 🔥 Force column to TEXT
  sheet.getRange("B:B").setNumberFormat("@");

  // Find next row
const nextRow = sheet.getLastRow() + 1;

// Write KEY
sheet.getRange(nextRow, 1).setValue(key);

// 🔥 Write MID as TRUE STRING (THIS IS THE FIX)
sheet.getRange(nextRow, 2).setValue("MID_" + formattedMID);

  return formattedMID;

}

function getPlayers(){

  const sheet=getSheet(SHEET_PLAYERS);
  const rows=sheet.getDataRange().getValues();

  let players=[];

  for(let i=1;i<rows.length;i++){

    let name=rows[i][0];
    let skill=rows[i][1];
    let active=rows[i][2];
    let matchMaker=rows[i][3];

    if(active===false) continue;

    players.push({
      name:name,
      skill:Number(skill),
      matchMaker: matchMaker !== false
    });

  }

  return players;

}

function getInitialData(){

  return {
  ok:true,
  players:getPlayers(),
  currentMatchup:getCurrentMatchup(),
  mapList:getMapList(),
  mapMatchMaker:getSavedMapMatchMaker()
}

}

function getCurrentMatchup(){

  const sheet=getSheet(SHEET_CURRENT);

  if(sheet.getLastRow()<2) return null;

  const row=sheet.getRange(2,1,1,9).getValues()[0];

  return {

  selectedAt:row[0],
  matchMaker:row[1],
  redTeam:row[2].split(", "),
  blueTeam:row[3].split(", "),
  redSkill:row[4],
  blueSkill:row[5],
  skillGap:row[6],
  expiresAt:row[7],
  MID: row[8]

};

}

function saveMatchup(data){

  const matchMaker=data.matchMaker;
  const red=data.redTeam;
  const blue=data.blueTeam;

  const MID = getOrCreateMID(red, blue);

  const players=getPlayers();

  let redSkill=0;
  let blueSkill=0;

  players.forEach(p=>{
    if(red.includes(p.name)) redSkill+=p.skill;
    if(blue.includes(p.name)) blueSkill+=p.skill;
  });

  const gap=Math.abs(redSkill-blueSkill);

  const now=new Date();
  const expiry=new Date(now.getTime()+7200000); // 2 hours

  const sheet=getSheet(SHEET_CURRENT);

  sheet.clearContents();

  sheet.appendRow([
  "SelectedAt","MatchMaker","RedTeam","BlueTeam","RedSkill","BlueSkill","SkillGap","ExpiresAt","MID"
]);

sheet.appendRow([
  now,matchMaker,red.join(", "),blue.join(", "),redSkill,blueSkill,gap,expiry,"MID_" + MID
]);

getSheet(SHEET_HISTORY).appendRow([
  now,matchMaker,red.join(", "),blue.join(", "),redSkill,blueSkill,gap,"MID_" + MID
]);

  return {ok:true};

}

function getHistory(){

  const sheet=getSheet(SHEET_HISTORY);

  if(sheet.getLastRow()<2) return {ok:true,history:[]};

  const rows=sheet.getRange(2,1,sheet.getLastRow()-1,8).getValues();

  let history=rows.map(r=>({

  selectedAt:r[0],
  matchMaker:r[1],
  redTeam:r[2],
  blueTeam:r[3],
  redSkill:r[4],
  blueSkill:r[5],
  skillGap:r[6],
  MID: r[7]

}));

  history.sort((a,b)=>new Date(b.selectedAt)-new Date(a.selectedAt));

  return {ok:true,history:history};

}

function clearHistory(data){

  const pass=data.password;

  if(!isValidAdmin(pass)) return {ok:false,error:"Wrong password"};

  const sheet=getSheet(SHEET_HISTORY);

  sheet.clearContents();

  sheet.appendRow([
  "SelectedAt","MatchMaker","RedTeam","BlueTeam","RedSkill","BlueSkill","SkillGap","MID"
]);

// also clear current matchup

const current=getSheet(SHEET_CURRENT);

current.clearContents();

current.appendRow([
  "SelectedAt","MatchMaker","RedTeam","BlueTeam","RedSkill","BlueSkill","SkillGap","ExpiresAt","MID"
]);

return {ok:true,message:"History cleared"};

}

function getPlayersAdmin(){

  return {ok:true,players:getPlayers()};

}

function savePlayersAdmin(data){

  const pass=data.password;

  if(!isValidAdmin(pass)) return {ok:false,error:"Wrong password"};

  const sheet=getSheet(SHEET_PLAYERS);

  sheet.clearContents();

  sheet.appendRow(["Name","Skill","Active","MatchMaker"]);

  data.players.forEach(p=>{
  sheet.appendRow([p.name,p.skill,true,p.matchMaker !== false]);
  });

  const savedMapMaker = getSavedMapMatchMaker();
  const savedMapMakerStillEligible = data.players.some(p =>
    p.name === savedMapMaker && p.matchMaker !== false
  );

  if(savedMapMaker && !savedMapMakerStillEligible){
    PropertiesService
      .getScriptProperties()
      .deleteProperty("MAP_MATCH_MAKER");
  }

  return {ok:true,message:"Players saved"};

}

// 🔥 GET MAP LIST FROM SHEET
function addRatingPlayer(data){

  const pass = data.password;

  if(!isValidAdmin(pass)) return {ok:false,error:"Wrong password"};

  const name = data.name ? data.name.toString().trim() : "";
  const skill = Number(data.skill);

  if(!name) return {ok:false,error:"Player name is required"};
  if(isNaN(skill)) return {ok:false,error:"Skill must be a number"};

  const sheet = getSheet(SHEET_PLAYERS);
  const values = sheet.getDataRange().getValues();

  for(let i = 1; i < values.length; i++){

    const existingName = values[i][0] ? values[i][0].toString().trim() : "";

    if(existingName.toLowerCase() === name.toLowerCase()){
      sheet.getRange(i + 1, 1).setValue(name);
      sheet.getRange(i + 1, 2).setValue(skill);
      sheet.getRange(i + 1, 3).setValue(true);
      sheet.getRange(i + 1, 4).setValue(true);

      return {
        ok:true,
        message:"Player reactivated",
        players:getPlayers()
      };
    }

  }

  sheet.appendRow([name, skill, true, true]);

  return {
    ok:true,
    message:"Player added",
    players:getPlayers()
  };

}

function deactivateRatingPlayer(data){

  const pass = data.password;

  if(!isValidAdmin(pass)) return {ok:false,error:"Wrong password"};

  const name = data.name ? data.name.toString().trim() : "";

  if(!name) return {ok:false,error:"Player name is required"};

  const sheet = getSheet(SHEET_PLAYERS);
  const values = sheet.getDataRange().getValues();

  for(let i = 1; i < values.length; i++){

    const existingName = values[i][0] ? values[i][0].toString().trim() : "";

    if(existingName.toLowerCase() === name.toLowerCase()){
      sheet.getRange(i + 1, 3).setValue(false);

      if(getSavedMapMatchMaker().toLowerCase() === name.toLowerCase()){
        PropertiesService
          .getScriptProperties()
          .deleteProperty("MAP_MATCH_MAKER");
      }

      return {
        ok:true,
        message:"Player deactivated",
        players:getPlayers()
      };
    }

  }

  return {
    ok:false,
    error:"Player not found"
  };

}

function getMapList(){

  const sheet = getSheet("MapList");

  const data = sheet.getDataRange().getValues();

  let elimination = [];
  let blitz = [];
  let ctf = [];

  // 🔥 START FROM ROW 2 (skip headers)
  for(let i = 1; i < data.length; i++){

    const el = data[i][0];
    const bl = data[i][1];
    const ct = data[i][2];

    if(el) elimination.push(el);
    if(bl) blitz.push(bl);
    if(ct) ctf.push(ct);

  }

  return {
    elimination,
    blitz,
    ctf
  };

}

// 🔥 GET CURRENT POINTERS FROM SETTINGS
function getMapPointers(){

  const sheet = getSheet("MapSettings");

  const data = sheet.getDataRange().getValues();

  let pointers = {};

  for(let i = 1; i < data.length; i++){

    const mode = data[i][0];
    const index = Number(data[i][1]);

    // 🔥 ONLY ACCEPT VALID MODES
    if(
      mode === "Elimination" ||
      mode === "Blitz" ||
      mode === "CTF"
    ){
      pointers[mode.toLowerCase()] = isNaN(index) ? 0 : index;
    }

  }

  return pointers;

}

// 🔥 GENERATE SESSION MAPS
function generateSessionMaps(){

  const maps = getMapList();
  const pointers = getMapPointers();

  const sheet = getSheet("MapList");

  const EL_COUNT = 2;
  const BL_COUNT = 2;
  const CTF_COUNT = 5;

  const elMaps = getNextMaps(maps.elimination, pointers.elimination, EL_COUNT);
  const blMaps = getNextMaps(maps.blitz, pointers.blitz, BL_COUNT);
  const ctMaps = getNextMaps(maps.ctf, pointers.ctf, CTF_COUNT);

// 🔥 CLEAR OLD SESSION LIST (COLUMN D)
sheet.getRange(2,4,EL_COUNT,1).clearContent();
sheet.getRange(5,4,BL_COUNT,1).clearContent();
sheet.getRange(8,4,CTF_COUNT,1).clearContent();

// 🔥 WRITE NEW SESSION LIST (COLUMN D)
sheet.getRange(2,4,elMaps.length,1).setValues(elMaps.map(m => [m]));
sheet.getRange(5,4,blMaps.length,1).setValues(blMaps.map(m => [m]));
sheet.getRange(8,4,ctMaps.length,1).setValues(ctMaps.map(m => [m]));

}

// 🔥 SAVE SESSION PROGRESS
function saveSessionProgress(){

  const maps = getMapList();
  const sheet = getSheet("MapList");

  const EL_COUNT = 2;
  const BL_COUNT = 2;
  const CTF_COUNT = 5;

  // 🔥 READ CURRENT SESSION LIST (COLUMN D)
  const elList = sheet.getRange(2,4,EL_COUNT,1).getValues().flat().filter(x => x);
  const blList = sheet.getRange(5,4,BL_COUNT,1).getValues().flat().filter(x => x);
  const ctList = sheet.getRange(8,4,CTF_COUNT,1).getValues().flat().filter(x => x);

  // 🔥 FIND LAST PLAYED MAP
  const lastEl = elList.length ? elList[elList.length - 1] : null;
  const lastBl = blList.length ? blList[blList.length - 1] : null;
  const lastCt = ctList.length ? ctList[ctList.length - 1] : null;

  let pointers = {
    elimination: null,
    blitz: null,
    ctf: null
  };

  if(lastEl){
    const idx = maps.elimination.indexOf(lastEl);
    pointers.elimination = idx >= 0 ? idx + 1 : maps.elimination.length;
  } else {
    pointers.elimination = maps.elimination.length;
  }

  if(lastBl){
    const idx = maps.blitz.indexOf(lastBl);
    pointers.blitz = idx >= 0 ? idx + 1 : maps.blitz.length;
  } else {
    pointers.blitz = maps.blitz.length;
  }

  if(lastCt){
    const idx = maps.ctf.indexOf(lastCt);
    pointers.ctf = idx >= 0 ? idx + 1 : maps.ctf.length;
  } else {
    pointers.ctf = maps.ctf.length;
  }

  saveMapPointers(pointers);

}

// 🔥 SAVE POINTERS BACK TO SHEET
function saveMapPointers(pointers){

  const sheet = getSheet("MapSettings");

  const data = sheet.getDataRange().getValues();

  for(let i = 1; i < data.length; i++){

    const mode = data[i][0];

    if(mode === "Elimination"){
      sheet.getRange(i+1,2).setValue(pointers.elimination);
    }

    if(mode === "Blitz"){
      sheet.getRange(i+1,2).setValue(pointers.blitz);
    }

    if(mode === "CTF"){
      sheet.getRange(i+1,2).setValue(pointers.ctf);
    }

  }

}

// 🔥 GET NEXT MAPS WITH WRAP-AROUND
function getNextMaps(list, startIndex, count){

  if(!list || list.length === 0) return [];

  let result = [];

  for(let i = 0; i < count; i++){
    const index = (startIndex + i) % list.length;
    result.push(list[index]);
  }

  return result;

}

// 🔥 GET CURRENT SESSION MAPS FROM COLUMN D
function getSessionMaps(){

  const sheet = getSheet("MapList");

  const EL_COUNT = 2;
  const BL_COUNT = 2;
  const CTF_COUNT = 5;

  const elimination = sheet.getRange(2,4,EL_COUNT,1).getValues().flat();
  const blitz = sheet.getRange(5,4,BL_COUNT,1).getValues().flat();
  const ctf = sheet.getRange(8,4,CTF_COUNT,1).getValues().flat();

  return {
    ok: true,
    elimination: elimination,
    blitz: blitz,
    ctf: ctf
  };

}

// 🔥 GENERATE SESSION MAPS WRAPPER FOR API
function runGenerateSessionMaps(data){

  const pass = data.password;

  if(!isValidAdmin(pass)){
    return { ok:false, error:"Wrong password" };
  }

  generateSessionMaps();

  return getSessionMaps();

}

function clearSessionMaps(){

  const sheet = getSheet("MapList");

  const EL_COUNT = 2;
  const BL_COUNT = 2;
  const CTF_COUNT = 5;

  sheet.getRange(2,4,EL_COUNT,1).clearContent();
  sheet.getRange(5,4,BL_COUNT,1).clearContent();
  sheet.getRange(8,4,CTF_COUNT,1).clearContent();

}

function runClearSessionMaps(data){

  const pass = data.password;

  if(!isValidAdmin(pass)){
    return { ok:false, error:"Wrong password" };
  }

  clearSessionMaps();

  return getSessionMaps();

}

// 🔥 SAVE SESSION PROGRESS WRAPPER FOR API
function runSaveSessionProgress(data){

  const pass = data.password;

  if(!isValidAdmin(pass)){
    return { ok:false, error:"Wrong password" };
  }

  saveSessionProgress();

  return {
    ok: true,
    message: "Session progress saved"
  };

}

// 🔥 DELETE ONE SESSION MAP BY MODE + SLOT
function deleteSessionMap(data){

  const pass = data.password;

if(!isValidAdmin(pass)){
  return { ok:false, error:"Wrong password" };
}

  const mode = data.mode;
  const slot = Number(data.slot);

  const sheet = getSheet("MapList");

  let startRow = null;
  let maxSlots = null;

  if(mode === "elimination"){
    startRow = 2;
    maxSlots = 2;
  }

  if(mode === "blitz"){
    startRow = 5;
    maxSlots = 2;
  }

  if(mode === "ctf"){
    startRow = 8;
    maxSlots = 5;
  }

  if(startRow === null){
    return { ok:false, error:"Invalid mode" };
  }

  if(!slot || slot < 1 || slot > maxSlots){
    return { ok:false, error:"Invalid slot" };
  }

// 🔥 SHIFT MAPS UP INSTEAD OF LEAVING GAPS

const range = sheet.getRange(startRow, 4, maxSlots, 1);
let values = range.getValues().flat();

// remove selected slot (0-based)
values.splice(slot - 1, 1);

// push empty to end
values.push("");

// write back cleaned list
range.setValues(values.map(v => [v]));  

  return getSessionMaps();

}

// 🔥 COPY SESSION MAPS FOR WEB APP
function copySessionMaps(){

  const session = getSessionMaps();

  let output = "";

  const elList = session.elimination.filter(x => x);
  const blList = session.blitz.filter(x => x);
  const ctList = session.ctf.filter(x => x);

  if(elList.length){
    output += "ELIMINATION:\\n";
    elList.forEach(m => output += m + "\\n");
    output += "\\n";
  }

  if(blList.length){
    output += "BLITZ:\\n";
    blList.forEach(m => output += m + "\\n");
    output += "\\n";
  }

  if(ctList.length){
    output += "CTF:\\n";
    ctList.forEach(m => output += m + "\\n");
  }

  return {
    ok: true,
    text: output.trim()
  };

}

function getDefaultCustomSessionData(){
  return {
    elimination: [],
    blitz: [],
    ctf: []
  };
}

function getSavedCustomSessionData(){

  const raw = PropertiesService
    .getScriptProperties()
    .getProperty("CUSTOM_SESSION_DATA");

  if(!raw){
    return getDefaultCustomSessionData();
  }

  try{
    const parsed = JSON.parse(raw);

    return {
      elimination: Array.isArray(parsed.elimination) ? parsed.elimination : [],
      blitz: Array.isArray(parsed.blitz) ? parsed.blitz : [],
      ctf: Array.isArray(parsed.ctf) ? parsed.ctf : []
    };
  }catch(err){
    return getDefaultCustomSessionData();
  }

}

function isCustomSessionActive(){
  return PropertiesService
    .getScriptProperties()
    .getProperty("CUSTOM_SESSION_ACTIVE") === "true";
}

function getCustomSession(){
  return {
    ok: true,
    active: isCustomSessionActive(),
    session: getSavedCustomSessionData()
  };
}

function saveCustomSession(data){

  const pass = data.password;

  if(!isValidAdmin(pass)){
    return { ok:false, error:"Wrong password" };
  }

  const session = data.session || {};

  const cleaned = {
    elimination: Array.isArray(session.elimination) ? session.elimination.filter(Boolean).slice(0, 2) : [],
    blitz: Array.isArray(session.blitz) ? session.blitz.filter(Boolean).slice(0, 2) : [],
    ctf: Array.isArray(session.ctf) ? session.ctf.filter(Boolean).slice(0, 5) : []
  };

  PropertiesService
    .getScriptProperties()
    .setProperty("CUSTOM_SESSION_DATA", JSON.stringify(cleaned));

  PropertiesService
    .getScriptProperties()
    .setProperty("CUSTOM_SESSION_ACTIVE", "true");

  return {
    ok: true,
    active: true,
    session: cleaned
  };

}

function clearCustomSession(data){

  const pass = data.password;

  if(!isValidAdmin(pass)){
    return { ok:false, error:"Wrong password" };
  }

  PropertiesService
    .getScriptProperties()
    .deleteProperty("CUSTOM_SESSION_DATA");

  PropertiesService
    .getScriptProperties()
    .setProperty("CUSTOM_SESSION_ACTIVE", "false");

  return {
    ok: true,
    active: false,
    session: getDefaultCustomSessionData()
  };

}

function getSavedMapMatchMaker(){
  return PropertiesService.getScriptProperties().getProperty("MAP_MATCH_MAKER") || "";
}

function saveGlobalMapMatchMaker(data){

  const matchMaker = (data.matchMaker || "").toString().trim();

  PropertiesService
    .getScriptProperties()
    .setProperty("MAP_MATCH_MAKER", matchMaker);

  return {
    ok: true,
    matchMaker: matchMaker
  };

}

// 🔐 VERIFY ADMIN PASSWORD (SAFE CHECK ONLY)
function getOrCreateSheet(name, headers){

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);

  if(!sheet){
    sheet = ss.insertSheet(name);
  }

  if(headers && headers.length){
    const firstRow = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
    const hasHeaders = firstRow.some(value => value !== "" && value !== null);

    if(!hasHeaders){
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.setFrozenRows(1);
    }
  }

  return sheet;

}

function getRatingSheetHeaders(){
  return {
    votes: [
      "Timestamp",
      "CycleId",
      "Rater",
      "RatedPlayer",
      "BlitzRating",
      "CTFRating",
      "FinalRating"
    ],
    codes: [
      "Code",
      "CycleId",
      "PlayerName",
      "Email",
      "CreatedAt",
      "ExpiresAt",
      "Used",
      "UsedAt"
    ],
    playerRatings: [
      "CycleId",
      "Player",
      "FinalSkill",
      "BlitzMedian",
      "CTFMedian",
      "VoteCount",
      "AppliedAt"
    ]
  };
}

function setupRatingSheets(data){

  const pass = data.password;

  if(!isValidAdmin(pass)){
    return { ok:false, error:"Wrong password" };
  }

  const headers = getRatingSheetHeaders();

  getOrCreateSheet(SHEET_RATINGS_VOTES, headers.votes);
  getOrCreateSheet(SHEET_RATING_CODES, headers.codes);
  getOrCreateSheet(SHEET_PLAYER_RATINGS, headers.playerRatings);

  return {
    ok: true,
    message: "Rating sheets are ready",
    sheets: [
      SHEET_RATINGS_VOTES,
      SHEET_RATING_CODES,
      SHEET_PLAYER_RATINGS
    ]
  };

}

function getDubaiDate(year, monthIndex, day, hour, minute, second){

  return new Date(Date.UTC(year, monthIndex, day, hour - 4, minute || 0, second || 0));

}

function getDubaiYear(dateValue){

  return Number(Utilities.formatDate(dateValue, RATINGS_TIME_ZONE, "yyyy"));

}

function getManualVotingOverride(now){

  const raw = PropertiesService
    .getScriptProperties()
    .getProperty("MANUAL_VOTING_OPEN_UNTIL");

  if(!raw) return null;

  const expiresAt = new Date(raw);

  if(isNaN(expiresAt.getTime()) || expiresAt <= now){
    PropertiesService
      .getScriptProperties()
      .deleteProperty("MANUAL_VOTING_OPEN_UNTIL");

    return null;
  }

  return expiresAt;

}

function getRatingCycleInfo(dateValue){

  const now = dateValue ? new Date(dateValue) : new Date();
  const year = getDubaiYear(now);

  const manualExpiresAt = !dateValue ? getManualVotingOverride(now) : null;

  if(manualExpiresAt){
    return {
      isOpen: true,
      manualOverride: true,
      cycleId: year + "-MANUAL",
      name: "Manual voting",
      message: "Manual voting is open",
      opensAt: now,
      closesAt: manualExpiresAt,
      appliesAt: manualExpiresAt,
      timeZoneLabel: RATINGS_TIME_ZONE_LABEL,
      daysUntilOpen: 0,
      daysUntilClose: Math.max(0, Math.ceil((manualExpiresAt - now) / 86400000))
    };
  }

  if(RATINGS_TEST_MODE && !dateValue){
    return {
      isOpen: true,
      manualOverride: false,
      cycleId: year + "-TEST",
      name: "Test voting",
      message: "Voting is open",
      opensAt: now,
      closesAt: new Date(now.getTime() + 14 * 24 * 60 * 60 * 1000),
      appliesAt: new Date(now.getTime() + 15 * 24 * 60 * 60 * 1000),
      timeZoneLabel: RATINGS_TIME_ZONE_LABEL,
      daysUntilOpen: 0,
      daysUntilClose: 14
    };
  }

  const windows = [
    {
      cycleId: year + "-SPRING",
      name: "Spring voting",
      opens: getDubaiDate(year, 3, 17, 0, 0, 0),
      closes: getDubaiDate(year, 3, 30, 23, 59, 59),
      applies: getDubaiDate(year, 4, 1, 0, 0, 0)
    },
    {
      cycleId: year + "-FALL",
      name: "Fall voting",
      opens: getDubaiDate(year, 10, 17, 0, 0, 0),
      closes: getDubaiDate(year, 10, 30, 23, 59, 59),
      applies: getDubaiDate(year, 11, 1, 0, 0, 0)
    },
    {
      cycleId: (year + 1) + "-SPRING",
      name: "Next spring voting",
      opens: getDubaiDate(year + 1, 3, 17, 0, 0, 0),
      closes: getDubaiDate(year + 1, 3, 30, 23, 59, 59),
      applies: getDubaiDate(year + 1, 4, 1, 0, 0, 0)
    }
  ];

  const active = windows.find(window => now >= window.opens && now <= window.closes);

  if(active){
    return {
      isOpen: true,
      manualOverride: false,
      cycleId: active.cycleId,
      name: active.name,
      message: "Voting is open",
      opensAt: active.opens,
      closesAt: active.closes,
      appliesAt: active.applies,
      timeZoneLabel: RATINGS_TIME_ZONE_LABEL,
      daysUntilOpen: 0,
      daysUntilClose: Math.max(0, Math.ceil((active.closes - now) / 86400000))
    };
  }

  const next = windows.find(window => now < window.opens);
  const daysUntilOpen = Math.ceil((next.opens - now) / 86400000);

  return {
    isOpen: false,
    manualOverride: false,
    cycleId: next.cycleId,
    name: next.name,
    message: "Voting is closed",
    opensAt: next.opens,
    closesAt: next.closes,
    appliesAt: next.applies,
    timeZoneLabel: RATINGS_TIME_ZONE_LABEL,
    daysUntilOpen: daysUntilOpen,
    daysUntilClose: null
  };

}

function hasRaterVotedInCycle(rater, cycleId){

  if(!rater || !cycleId) return false;

  const sheet = getSheet(SHEET_RATINGS_VOTES);

  if(!sheet || sheet.getLastRow() < 2) return false;

  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();

  return rows.some(row => row[1] === cycleId && row[2] === rater);

}

function getActiveRatingCode(rater, cycleId){

  const sheet = getSheet(SHEET_RATING_CODES);

  if(!sheet || sheet.getLastRow() < 2) return null;

  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 8).getValues();
  const now = new Date();

  for(let i = rows.length - 1; i >= 0; i--){

    const row = rows[i];

    if(
      row[1] === cycleId &&
      row[2] === rater &&
      row[6] !== true &&
      new Date(row[5]) > now
    ){
      return {
        rowNumber: i + 2,
        code: row[0],
        email: row[3],
        expiresAt: row[5]
      };
    }

  }

  return null;

}

function generateRatingCode(){
  return String(Math.floor(100000 + Math.random() * 900000));
}

function requestRatingCode(data){

  const status = getRatingCycleInfo();

  if(!status.isOpen){
    return {
      ok: false,
      error: "Voting is currently closed"
    };
  }

  const rater = data && data.rater ? data.rater.toString().trim() : "";
  const email = data && data.email ? data.email.toString().trim() : "";
  const activePlayers = getActivePlayerNameSet();

  if(!rater || !activePlayers[rater]){
    return {
      ok: false,
      error: "Select a valid rater"
    };
  }

  if(!email || email.indexOf("@") === -1){
    return {
      ok: false,
      error: "Enter a valid email address"
    };
  }

  if(hasRaterVotedInCycle(rater, status.cycleId)){
    return {
      ok: false,
      error: "You have already voted in this cycle"
    };
  }

  const headers = getRatingSheetHeaders();
  const codeSheet = getOrCreateSheet(SHEET_RATING_CODES, headers.codes);
  const existingCode = getActiveRatingCode(rater, status.cycleId);

  if(existingCode){
    return {
      ok: false,
      error: "An active code was already sent. Use that code or wait 1 hour for it to expire.",
      expiresAt: existingCode.expiresAt
    };
  }

  const now = new Date();
  const expiresAt = new Date(now.getTime() + 60 * 60 * 1000);
  const code = generateRatingCode();

  codeSheet.appendRow([
    code,
    status.cycleId,
    rater,
    email,
    now,
    expiresAt,
    false,
    ""
  ]);

  MailApp.sendEmail({
    to: email,
    subject: "DXB99 rating vote code",
    body:
      "Your DXB99 rating vote code is: " + code + "\n\n" +
      "This code expires in 1 hour.\n\n" +
      "Cycle: " + status.cycleId
  });

  return {
    ok: true,
    cycleId: status.cycleId,
    rater: rater,
    email: email,
    expiresAt: expiresAt
  };

}

function validateRatingCode(rater, cycleId, code){

  if(!code){
    return {
      ok: false,
      error: "Enter your voting code"
    };
  }

  const activeCode = getActiveRatingCode(rater, cycleId);

  if(!activeCode){
    return {
      ok: false,
      error: "No active voting code found. Request a new code."
    };
  }

  if(activeCode.code.toString() !== code.toString().trim()){
    return {
      ok: false,
      error: "Invalid voting code"
    };
  }

  return {
    ok: true,
    rowNumber: activeCode.rowNumber
  };

}

function markRatingCodeUsed(rowNumber){

  const sheet = getSheet(SHEET_RATING_CODES);

  if(!sheet || !rowNumber) return;

  sheet.getRange(rowNumber, 7).setValue(true);
  sheet.getRange(rowNumber, 8).setValue(new Date());

}

function getRatingStatus(data){

  const status = getRatingCycleInfo();
  const rater = data && data.rater ? data.rater.toString().trim() : "";
  const hasVoted = hasRaterVotedInCycle(rater, status.cycleId);
  const latestRatings = getLatestPlayerRatings();

  return {
    ok: true,
    isOpen: status.isOpen,
    manualOverride: !!status.manualOverride,
    cycleId: status.cycleId,
    name: status.name,
    message: status.message,
    opensAt: status.opensAt,
    closesAt: status.closesAt,
    appliesAt: status.appliesAt,
    timeZoneLabel: status.timeZoneLabel,
    daysUntilOpen: status.daysUntilOpen,
    daysUntilClose: status.daysUntilClose,
    rater: rater,
    hasVoted: hasVoted,
    latestRatings: latestRatings
  };

}

function setManualVotingWindow(data){

  const pass = data.password;

  if(!isValidAdmin(pass)){
    return { ok:false, error:"Wrong password" };
  }

  const enabled = data.enabled === true;
  const props = PropertiesService.getScriptProperties();

  if(enabled){
    const now = new Date();
    const expiresAt = new Date(now.getTime() + 2 * 60 * 60 * 1000);

    props.setProperty("MANUAL_VOTING_OPEN_UNTIL", expiresAt.toISOString());

    return {
      ok: true,
      enabled: true,
      expiresAt: expiresAt,
      status: getRatingCycleInfo()
    };
  }

  props.deleteProperty("MANUAL_VOTING_OPEN_UNTIL");

  return {
    ok: true,
    enabled: false,
    status: getRatingCycleInfo()
  };

}

function getLatestPlayerRatings(){

  const sheet = getSheet(SHEET_PLAYER_RATINGS);

  if(!sheet || sheet.getLastRow() < 2) return [];

  const headers = getRatingSheetHeaders();
  const rows = sheet
    .getRange(2, 1, sheet.getLastRow() - 1, headers.playerRatings.length)
    .getValues();

  let latestCycle = "";
  let latestDate = null;

  rows.forEach(row => {
    const cycleId = row[0];
    const appliedAt = row[6] ? new Date(row[6]) : null;

    if(!cycleId) return;

    if(!latestCycle){
      latestCycle = cycleId;
      latestDate = appliedAt;
      return;
    }

    if(appliedAt && (!latestDate || appliedAt > latestDate)){
      latestCycle = cycleId;
      latestDate = appliedAt;
    }
  });

  if(!latestCycle) return [];

  return rows
    .filter(row => row[0] === latestCycle)
    .map(row => ({
      cycleId: row[0],
      player: row[1],
      finalSkill: row[2],
      blitzMedian: row[3],
      ctfMedian: row[4],
      voteCount: row[5],
      appliedAt: row[6]
    }));

}

function getActivePlayerNameSet(){

  const players = getPlayers();
  const names = {};

  players.forEach(player => {
    names[player.name] = true;
  });

  return names;

}

function normalizeRatingValue(value){

  const numberValue = Number(value);

  if(!Number.isInteger(numberValue)) return null;
  if(numberValue < 0 || numberValue > 10) return null;

  return numberValue;

}

function getMedian(values){

  const cleaned = values
    .map(value => Number(value))
    .filter(value => !isNaN(value))
    .sort((a, b) => a - b);

  if(cleaned.length === 0) return "";

  const middle = Math.floor(cleaned.length / 2);

  if(cleaned.length % 2 === 1){
    return cleaned[middle];
  }

  return (cleaned[middle - 1] + cleaned[middle]) / 2;

}

function recalculatePlayerRatings(cycleId){

  const headers = getRatingSheetHeaders();
  const votesSheet = getOrCreateSheet(SHEET_RATINGS_VOTES, headers.votes);
  const ratingsSheet = getOrCreateSheet(SHEET_PLAYER_RATINGS, headers.playerRatings);

  if(votesSheet.getLastRow() < 2){
    return [];
  }

  const voteRows = votesSheet
    .getRange(2, 1, votesSheet.getLastRow() - 1, headers.votes.length)
    .getValues()
    .filter(row => row[1] === cycleId);

  const grouped = {};

  voteRows.forEach(row => {
    const ratedPlayer = row[3];
    const blitz = Number(row[4]);
    const ctf = Number(row[5]);

    if(!ratedPlayer) return;

    if(!grouped[ratedPlayer]){
      grouped[ratedPlayer] = {
        blitz: [],
        ctf: [],
        combined: [],
        raters: {}
      };
    }

    grouped[ratedPlayer].blitz.push(blitz);
    grouped[ratedPlayer].ctf.push(ctf);
    grouped[ratedPlayer].combined.push(blitz, ctf);
    grouped[ratedPlayer].raters[row[2]] = true;
  });

  const now = new Date();

  const output = Object.keys(grouped)
    .sort()
    .map(player => {
      const blitzMedian = getMedian(grouped[player].blitz);
      const ctfMedian = getMedian(grouped[player].ctf);
      const finalMedian = getMedian(grouped[player].combined);
      const finalSkill = finalMedian === "" ? "" : Math.round(finalMedian);

      return [
        cycleId,
        player,
        finalSkill,
        blitzMedian,
        ctfMedian,
        Object.keys(grouped[player].raters).length,
        now
      ];
    });

  if(ratingsSheet.getLastRow() > 1){
    const existingRows = ratingsSheet
      .getRange(2, 1, ratingsSheet.getLastRow() - 1, headers.playerRatings.length)
      .getValues()
      .filter(row => row[0] !== cycleId);

    ratingsSheet
      .getRange(2, 1, ratingsSheet.getLastRow() - 1, headers.playerRatings.length)
      .clearContent();

    const rowsToWrite = existingRows.concat(output);

    if(rowsToWrite.length){
      ratingsSheet
        .getRange(2, 1, rowsToWrite.length, headers.playerRatings.length)
        .setValues(rowsToWrite);
    }
  }else if(output.length){
    ratingsSheet
      .getRange(2, 1, output.length, headers.playerRatings.length)
      .setValues(output);
  }

  return output;

}

function submitRatings(data){

  const status = getRatingCycleInfo();

  if(!status.isOpen){
    return {
      ok: false,
      error: "Voting is currently closed"
    };
  }

  const rater = data && data.rater ? data.rater.toString().trim() : "";
  const ratings = data && Array.isArray(data.ratings) ? data.ratings : [];
  const code = data && data.code ? data.code.toString().trim() : "";
  const activePlayers = getActivePlayerNameSet();

  if(!rater || !activePlayers[rater]){
    return {
      ok: false,
      error: "Select a valid rater"
    };
  }

  if(hasRaterVotedInCycle(rater, status.cycleId)){
    return {
      ok: false,
      error: "You have already voted in this cycle"
    };
  }

  const codeCheck = validateRatingCode(rater, status.cycleId, code);

  if(!codeCheck.ok){
    return codeCheck;
  }

  if(ratings.length === 0){
    return {
      ok: false,
      error: "No ratings submitted"
    };
  }

  const cleanedRatings = [];
  const seenPlayers = {};

  ratings.forEach(rating => {
    const ratedPlayer = rating && rating.ratedPlayer
      ? rating.ratedPlayer.toString().trim()
      : "";

    const blitz = normalizeRatingValue(rating ? rating.blitz : null);
    const ctf = normalizeRatingValue(rating ? rating.ctf : null);

    if(!ratedPlayer) return;
    if(ratedPlayer === rater) return;
    if(!activePlayers[ratedPlayer]) return;
    if(seenPlayers[ratedPlayer]) return;
    if(blitz === null || ctf === null) return;

    seenPlayers[ratedPlayer] = true;

    cleanedRatings.push({
      ratedPlayer: ratedPlayer,
      blitz: blitz,
      ctf: ctf,
      finalRating: (blitz + ctf) / 2
    });
  });

  if(cleanedRatings.length === 0){
    return {
      ok: false,
      error: "No valid ratings submitted"
    };
  }

  const headers = getRatingSheetHeaders();
  const votesSheet = getOrCreateSheet(SHEET_RATINGS_VOTES, headers.votes);
  getOrCreateSheet(SHEET_RATING_CODES, headers.codes);
  getOrCreateSheet(SHEET_PLAYER_RATINGS, headers.playerRatings);

  const now = new Date();
  const rows = cleanedRatings.map(rating => [
    now,
    status.cycleId,
    rater,
    rating.ratedPlayer,
    rating.blitz,
    rating.ctf,
    rating.finalRating
  ]);

  votesSheet
    .getRange(votesSheet.getLastRow() + 1, 1, rows.length, headers.votes.length)
    .setValues(rows);

  const recalculated = recalculatePlayerRatings(status.cycleId);

  markRatingCodeUsed(codeCheck.rowNumber);

  return {
    ok: true,
    cycleId: status.cycleId,
    rater: rater,
    submittedCount: cleanedRatings.length,
    playerRatingsUpdated: recalculated.length
  };

}

function getLatestRatingCycleId(){

  const sheet = getSheet(SHEET_PLAYER_RATINGS);

  if(!sheet || sheet.getLastRow() < 2) return "";

  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 7).getValues();
  let latestCycle = "";
  let latestDate = null;

  rows.forEach(row => {
    const cycleId = row[0];
    const appliedAt = row[6] ? new Date(row[6]) : null;

    if(!cycleId) return;

    if(!latestCycle){
      latestCycle = cycleId;
      latestDate = appliedAt;
      return;
    }

    if(appliedAt && (!latestDate || appliedAt > latestDate)){
      latestCycle = cycleId;
      latestDate = appliedAt;
    }
  });

  return latestCycle;

}

function sanitizeSheetNamePart(value){

  return value
    .toString()
    .trim()
    .replace(/[^A-Za-z0-9_]/g, "_")
    .substring(0, 40);

}

function makePlayersBackup(cycleId){

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const playersSheet = getSheet(SHEET_PLAYERS);
  const safeCycleId = sanitizeSheetNamePart(cycleId || "UNKNOWN");
  const baseName = "PlayersBackup_" + safeCycleId;
  let backupName = baseName;

  if(ss.getSheetByName(backupName)){
    const timestamp = Utilities.formatDate(
      new Date(),
      Session.getScriptTimeZone(),
      "yyyyMMdd_HHmmss"
    );

    backupName = baseName + "_" + timestamp;
  }

  if(backupName.length > 99){
    const timestamp = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    "yyyyMMdd_HHmmss"
  );
    backupName = "PlayersBackup_" + timestamp;
  }

  const backup = playersSheet.copyTo(ss);

  backup.setName(backupName);

  return backupName;

}

function applyLatestRatingsToPlayers(data){

  const pass = data.password;

  if(!isValidAdmin(pass)) return {ok:false,error:"Wrong password"};

  const ratingsSheet = getSheet(SHEET_PLAYER_RATINGS);

  if(!ratingsSheet || ratingsSheet.getLastRow() < 2){
    return {
      ok:false,
      error:"No calculated player ratings found"
    };
  }

  const cycleId = data.cycleId
    ? data.cycleId.toString().trim()
    : getLatestRatingCycleId();

  if(!cycleId){
    return {
      ok:false,
      error:"No rating cycle found"
    };
  }

  const ratingRows = ratingsSheet
    .getRange(2, 1, ratingsSheet.getLastRow() - 1, 7)
    .getValues()
    .filter(row => row[0] === cycleId);

  if(ratingRows.length === 0){
    return {
      ok:false,
      error:"No ratings found for cycle " + cycleId
    };
  }

  const skillsByPlayer = {};

  ratingRows.forEach(row => {
    const player = row[1] ? row[1].toString().trim() : "";
    const skill = Number(row[2]);

    if(player && !isNaN(skill)){
      skillsByPlayer[player.toLowerCase()] = skill;
    }
  });

  const playersSheet = getSheet(SHEET_PLAYERS);
  const players = playersSheet.getDataRange().getValues();
  const backupSheet = makePlayersBackup(cycleId);
  let updatedCount = 0;

  for(let i = 1; i < players.length; i++){
    const playerName = players[i][0] ? players[i][0].toString().trim() : "";
    const key = playerName.toLowerCase();

    if(skillsByPlayer.hasOwnProperty(key)){
      playersSheet.getRange(i + 1, 2).setValue(skillsByPlayer[key]);
      updatedCount++;
    }
  }

  return {
    ok:true,
    cycleId:cycleId,
    updatedCount:updatedCount,
    backupSheet:backupSheet,
    players:getPlayers()
  };

}

function verifyAdminPassword(data){

  const pass = data.password;

  if(!isValidAdmin(pass)){
    return { ok:false, error:"Wrong password" };
  }

  return { ok:true };

}
