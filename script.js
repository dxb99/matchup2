const API_URL = "https://script.google.com/macros/s/AKfycbwhOi84b1_5S95OM2fANahXJi_8e4ARzikvKvFv9aU6Z0d-AM1w50MzplebCP0fr2utAg/exec";

let allPlayers = [];
let adminLoaded = false;
let countdownTimer = null;
let lastMatchTimestamp = null;
let lastGeneratedMatchups = [];
let selectedMatchKey = null;
let matchHistory = [];
let lastSelectedPlayers = [];
let lastSelectedMatchMaker = "";
let currentMatchKeyFromServer = null;
let blitzEnabled = false;
let currentHistorySort = {
  key: "date",
  direction: "desc"
};
let globalMapMatchMaker = "";
let adminHasUnsavedChanges = false;
let currentAdminSort = {
  key: "name",
  direction: "asc"
};
let customSessionActive = false;
let customSessionData = {
  elimination: [],
  blitz: [],
  ctf: []
};
let currentSessionMaps = {
  elimination: [],
  blitz: [],
  ctf: []
};
let currentRatingStatus = null;
const RATINGS_TIME_ZONE = "Asia/Dubai";
const RATINGS_TIME_ZONE_LABEL = "GST";
let armedMatchKey = null; // 🔥 tracks first click before confirm

/* 🔥 GLOBAL MODAL SYSTEM */

function showModal(message, type = "alert", confirmText = "Confirm", cancelText = "Cancel", withInput = false, inputType = "password", inputPlaceholder = "Enter password"){

  return new Promise((resolve)=>{

    const modal = document.getElementById("customModal");
    const msg = document.getElementById("modalMessage");
    const confirmBtn = document.getElementById("modalConfirm");
    const cancelBtn = document.getElementById("modalCancel");
    const input = document.getElementById("modalInput");

    msg.innerText = message;

    confirmBtn.innerHTML = confirmText === "Confirm" ? "✓" : confirmText;
    cancelBtn.innerHTML = cancelText === "Cancel" ? "✕" : cancelText;

    input.style.display = withInput ? "block" : "none";
    input.type = inputType;
    input.placeholder = inputPlaceholder;
    input.value = "";

    modal.style.display = "flex";

    if(type === "alert"){
      cancelBtn.style.display = "none";
    }else{
      cancelBtn.style.display = "inline-flex";
    }

    const cleanup = () => {
      modal.style.display = "none";
      confirmBtn.onclick = null;
      cancelBtn.onclick = null;
    };

    confirmBtn.onclick = () => {
      const value = withInput ? input.value : true;
      cleanup();
      resolve(value);
    };

    cancelBtn.onclick = () => {
      cleanup();
      resolve(null);
    };

  });

}

/* 🔓 ADMIN UNLOCK SYSTEM */

async function getAdminPassword(){

  let stored = sessionStorage.getItem("adminPass");

  // ✅ If already unlocked → reuse
  if(stored){
    return stored;
  }

  while(true){

let pass = await showModal(
  "Enter Admin Password",
  "confirm",
  "Confirm",
  "Cancel",
  true
);    

// ❌ user cancelled
if(!pass) return null;

const test = await api({
  action:"verifyAdminPassword",
  password: pass
});

if(test && test.ok){
  sessionStorage.setItem("adminPass", pass);
  updateAdminBar();
  return pass;
}

await showModal("Wrong password. Try again.", "alert");
  }
}

// 🔥 lock function (for later button)
function clearAdminSession(){
  sessionStorage.removeItem("adminPass");
  updateAdminBar();
}

function updateAdminBar(){

  const status = document.getElementById("adminStatus");
  const lockBtn = document.getElementById("adminLockBtn");

  if(!status || !lockBtn) return;

  const pass = sessionStorage.getItem("adminPass");

if(pass){
  status.textContent = "🔓 ADMIN MODE ACTIVE";
  lockBtn.style.display = "inline-flex";

  document.body.classList.remove("admin-locked");
  document.body.classList.add("admin-unlocked");

}else{
  status.textContent = "🔒 LOCKED";
  lockBtn.style.display = "none";

  document.body.classList.remove("admin-unlocked");
  document.body.classList.add("admin-locked");
}
  
// 🔥 disable session buttons when locked
const generateBtn = document.getElementById("generateSessionMapsBtn");
const saveBtn = document.getElementById("saveSessionProgressBtn");
const clearSessionBtn = document.getElementById("clearSessionMapsBtn");
const buildCustomBtn = document.getElementById("buildCustomSessionBtn");
const mapMakerSelect = document.getElementById("mapMatchMakerSelect");

const protectedSessionButtons = [
  generateBtn,
  saveBtn,
  clearSessionBtn,
  buildCustomBtn
].filter(Boolean);

if(protectedSessionButtons.length){

  protectedSessionButtons.forEach(btn => {

    if(pass){

      btn.classList.remove("disabled");
      btn.removeAttribute("data-tooltip");

    }else{

      btn.classList.add("disabled");
      btn.setAttribute("data-tooltip", "🔒 Admin mode required");

    }

  });

}

if(mapMakerSelect){

  mapMakerSelect.disabled = false;

  if(pass){

    mapMakerSelect.classList.remove("disabled");

  }else{

    mapMakerSelect.classList.add("disabled");

  }

}

// 🔥 CLICK LOCK STATUS TO UNLOCK
status.onclick = async () => {

  // only allow unlock when locked
  if(sessionStorage.getItem("adminPass")) return;

  const pass = await getAdminPassword();
  if(!pass) return;

  sessionStorage.setItem("adminPass", pass);
  updateAdminBar();

};
  
}

function isAdminUnlocked(){
  return !!sessionStorage.getItem("adminPass");
}

window.addEventListener("load", async () => {

sessionStorage.removeItem("selectedGeneratorMatchMaker");
sessionStorage.removeItem("selectedPlayers");
sessionStorage.removeItem("adminPass");

  try {

    await loadInitialData();
    setupRatingsTab();
    renderRatingsPreview();
    await refreshRatingStatus();

/* 🔥 HIDE BLITZ ON LOAD */

const blitzContainer = document.querySelector(".blitzToggle");
if(blitzContainer){
  blitzContainer.style.display = "none";
}

    document.getElementById("loadingScreen").style.display = "none";
    document.getElementById("app").classList.remove("hidden");
    updateAdminBar();

window.scrollTo(0, 0);

    document.querySelectorAll('input[name="gapFilter"]').forEach(radio => {
      radio.addEventListener("change", applyGapFilter);
    });

const blitzToggle = document.getElementById("blitzToggle");

if(blitzToggle){

blitzToggle.addEventListener("change", () => {

  blitzEnabled = blitzToggle.checked;

  updateGapCounts(); /* 🔥 ADD THIS */
  applyGapFilter();

});

}

    setupMapListButtons();
    await loadSessionMaps();

    document.getElementById("adminLockBtn").onclick = clearAdminSession;
    
    startMatchAutoRefresh();

  } catch (err) {

    console.error(err);
    await showModal("Startup error. Open console (F12).", "alert");

  }

});

async function api(data){

  const res = await fetch(API_URL,{
    method:"POST",
    headers:{
      "Content-Type":"text/plain;charset=utf-8"
    },
    body:JSON.stringify(data)
  });

  return await res.json();

}

function showBusy(message = "LOADING"){

  const overlay = document.getElementById("savingOverlay");

  if(!overlay) return;

  const text = overlay.querySelector(".generatingText");

  if(text){
    text.innerHTML = `${message}<span class="dots"></span>`;
  }

  overlay.style.display = "flex";

}

function hideBusy(){

  const overlay = document.getElementById("savingOverlay");

  if(!overlay) return;

  overlay.style.display = "none";

  const text = overlay.querySelector(".generatingText");

  if(text){
    text.innerHTML = "SAVING<span class=\"dots\"></span>";
  }

}

async function loadInitialData(){

const data = await api({action:"getInitialData"});

if(!data.ok){
  throw new Error("Failed loading data");
}

allPlayers = data.players || [];
globalMapMatchMaker = data.mapMatchMaker || "";

populatePlayers(allPlayers);

if(lastGeneratedMatchups.length === 0){

  document.querySelectorAll('input[name="gapFilter"]').forEach(r=>{
    r.disabled = true;
    r.parentElement.classList.add("disabled");
  });

}else{

  // Re-apply correct radio states based on matchups
  updateGapCounts();

  // Re-render filtered matchups
  applyGapFilter();

}

renderMatchup(data.currentMatchup);

/* LOAD MATCH HISTORY */

const historyData = await api({action:"getHistory"});
if(historyData.ok){
  matchHistory = historyData.history || [];
}

}

function populatePlayers(players){

  window.allPlayers = [...players].sort((a,b)=>a.name.localeCompare(b.name));

  renderPlayers(window.allPlayers);

  document.getElementById("playerSort").onchange = function(){

    let type = this.value;

    let sorted = [...window.allPlayers];

    if(type === "alpha"){

      sorted.sort((a,b)=>a.name.localeCompare(b.name));

    }else{

      sorted.sort((a,b)=>{

        if(b.skill !== a.skill) return b.skill - a.skill;

        return a.name.localeCompare(b.name);

      });

    }

    renderPlayers(sorted);

  };

}

function renderMatchup(match){

const el=document.getElementById("matchupContent");
const countdown=document.getElementById("matchCountdown");
const upcomingSessionMaker = document.getElementById("upcomingSessionMaker");

// 🔥 RESET server key if no matchup
currentMatchKeyFromServer = null;
  
if(!match){

  if(countdownTimer){
    clearInterval(countdownTimer);
    countdownTimer = null;
  }

  if(upcomingSessionMaker) upcomingSessionMaker.style.display = "";

  el.innerHTML=`

  <div class="matchCard">
    <div class="matchHeader">
      NO CURRENT MATCHUP
    </div>

    <button id="getStartedBtn" class="getStartedBtn">
      CLICK TO GET STARTED
    </button>

  </div>

  `;

  countdown.innerHTML = "MATCHUP EXPIRES IN --:--:--";

/* 🔥 GET STARTED BUTTON CLICK */
setTimeout(() => {
  const btn = document.getElementById("getStartedBtn");
  if(btn){
    btn.onclick = () => {
      const generatorBtn = document.querySelector('.tabButton[onclick*="generatorTab"]');
      showTab("generatorTab", generatorBtn);
    };
  }
}, 0);

return;
  return;

}

el.innerHTML=`

<div class="matchCard">

<div class="matchHeader">
  Match Maker: <strong>${match.matchMaker}</strong>

<span class="midTag">
  ${match.MID ? "MID_" + String(match.MID).replace("MID_","").padStart(4,"0") : "----"}
</span>
</div>

  <div class="teamsRow">

    <div class="team red">
  <div class="teamTitle">
    RED TEAM <span class="teamBadge">${match.redSkill}</span>
  </div>
      <div class="teamPlayers">
        ${match.redTeam.map(name => {

  const player = allPlayers.find(p => p.name === name);

  return `
    <div class="playerRow">
      ${name}
      <span class="skillMedal">${player ? player.skill : ""}</span>
    </div>
  `;

}).join("")}
      </div>
    </div>

    <div class="vs">VS</div>

    <div class="team blue">
  <div class="teamTitle">
    BLUE TEAM <span class="teamBadge">${match.blueSkill}</span>
  </div>
      <div class="teamPlayers">
        ${match.blueTeam.map(name => {

  const player = allPlayers.find(p => p.name === name);

  return `
    <div class="playerRow">
      ${name}
      <span class="skillMedal">${player ? player.skill : ""}</span>
    </div>
  `;

}).join("")}
      </div>
    </div>

  </div>

  <div class="matchFooter">
    <span class="diff diff-${match.skillGap}">
  Difference: ${match.skillGap}
</span>
  </div>

</div>

`;

const expiry = new Date(match.expiresAt);
const now = new Date();

if(expiry <= now){

  // 🔥 CLEAR SERVER MATCH KEY (THIS FIXES YOUR ISSUE)
  currentMatchKeyFromServer = null;

  if(upcomingSessionMaker) upcomingSessionMaker.style.display = "";

  el.innerHTML=`

  <div class="matchCard">
    <div class="matchHeader">
      NO CURRENT MATCHUP
    </div>

    <button id="getStartedBtn" class="getStartedBtn">
      CLICK TO GET STARTED
    </button>

  </div>

  `;

  countdown.innerHTML = "MATCHUP EXPIRES IN --:--:--";

  /* 🔥 GET STARTED BUTTON CLICK */
  setTimeout(() => {
    const btn = document.getElementById("getStartedBtn");
    if(btn){
      btn.onclick = () => {
        const generatorBtn = document.querySelector('.tabButton[onclick*="generatorTab"]');
        showTab("generatorTab", generatorBtn);
      };
    }
  }, 0);

  return;

}

// 🔥 BUILD KEY FROM SERVER MATCH (ONLY IF NOT EXPIRED)
const redKey = match.redTeam.slice().sort().join("|");
const blueKey = match.blueTeam.slice().sort().join("|");

currentMatchKeyFromServer = redKey + "-" + blueKey;

if(upcomingSessionMaker) upcomingSessionMaker.style.display = "none";

if(match.selectedAt !== lastMatchTimestamp){

  lastMatchTimestamp = match.selectedAt;

  startCountdown(expiry);

}

}

document.getElementById("generateButton").onclick = generateMatchups;

async function generateMatchups(){

  const selectedPlayers=[];

  const maker = document.getElementById("matchMakerSelect").value;

if(!maker){
  showModal("Select Match Maker first.", "alert");
  return;
}

  document.querySelectorAll("#playersCheckboxes input:checked").forEach(x=>{
    selectedPlayers.push(x.value);
  });
  
  if(selectedPlayers.length < 2){
    showModal("Select at least 2 players.", "alert");
    return;
  }

  document.getElementById("generatingOverlay").style.display = "flex";

  const gap = document.querySelector('input[name="gapFilter"]:checked').value;

const matchups = generateMatchupsLocal(selectedPlayers, gap);

/* 🔥 CONTROL BLITZ VISIBILITY AFTER GENERATE */

const blitzToggle = document.getElementById("blitzToggle");
const blitzContainer = document.querySelector(".blitzToggle");

if(blitzToggle && blitzContainer){

  if(selectedPlayers.length % 2 !== 0){

    /* SHOW with animation */
    blitzContainer.style.display = "flex";

requestAnimationFrame(() => {
  blitzContainer.classList.add("show");
});

  }else{

    /* HIDE */
    blitzContainer.classList.remove("show");

    setTimeout(()=>{
      blitzContainer.style.display = "none";
    },300);

    blitzToggle.checked = false;
    blitzEnabled = false;

  }

}

/* Sort matchups by skill gap */

matchups.sort((a,b)=>a.skillGap - b.skillGap);

lastGeneratedMatchups = matchups;
lastSelectedPlayers = selectedPlayers.slice();

/* Force overlay to stay visible for 1 seconds */

  setTimeout(() => {

  document.querySelectorAll('input[name="gapFilter"]').forEach(r=>{
  r.disabled = false;
  r.parentElement.classList.remove("disabled");
});

  updateGapCounts();

  applyGapFilter();

  document.getElementById("generatingOverlay").style.display = "none";

}, 1000);

}

function renderGeneratedMatchups(matchups){

  const container=document.getElementById("generatedMatchups");

  armedMatchKey = null; // 🔥 reset when rendering new matchups

  container.innerHTML="";

  matchups.forEach(m=>{

    const div=document.createElement("div");

    div.className="matchOption";

div.innerHTML=`

<div class="matchCompact">

<div class="teamLine">

<span class="redTeam"><strong><span class="skillMedal">${m.redSkill}</span> RED TEAM :</strong></span>

<span class="teamPlayers">
${m.redTeam.map(p=>p.name).join(", ")}
</span>

</div>

<div class="teamLine">

<span class="blueTeam"><strong><span class="skillMedal">${m.blueSkill}</span> BLUE TEAM :</strong></span>

<span class="teamPlayers">
${m.blueTeam.map(p=>p.name).join(", ")}
</span>

</div>

<div class="badges">

<span class="badge gap-${m.skillGap}">
Difference ${m.skillGap}
</span>

<span class="badge picks">
Picked ${m.pickCount} ${m.pickCount === 1 ? "time" : "times"}
</span>

<button class="selectMatch">CLICK TO SELECT</button>

</div>

</div>

`;
    
const btn = div.querySelector(".selectMatch");

const redKey = m.redTeam.map(p=>p.name).sort().join("|");
const blueKey = m.blueTeam.map(p=>p.name).sort().join("|");

const key = redKey + "-" + blueKey;

const isServerSelected = currentMatchKeyFromServer === key;

if(isServerSelected){
  div.classList.add("selectedCard");
  btn.classList.add("selected");
  btn.innerText = "SELECTED";

  btn.style.cursor = "not-allowed";
  btn.disabled = true;
}

btn.onclick = () => {

  // 🔥 BLOCK IF THIS IS CURRENT ACTIVE MATCH
  if(currentMatchKeyFromServer === key){
    return;
  }

  const maker = document.getElementById("matchMakerSelect").value;

  if(!maker){
    showModal("Select Match Maker first.", "alert");
    return;
  }

// 🔥 CUSTOM MODAL CONFIRM
showModal("Are you sure you want to select this matchup?", "confirm")
.then(confirmSelection => {

  if(!confirmSelection){
    return;
  }

  // 🔥 SAVE DIRECTLY
  selectMatchup(m, key, btn, div);

});
};

// 🔥 MAKE ENTIRE CARD CLICKABLE (SAME AS BUTTON)
div.onclick = (e) => {

  // prevent double trigger if button itself clicked
  if(e.target.classList.contains("selectMatch")) return;

  btn.click(); // trigger same logic as button

};
    
container.appendChild(div);

  });

}

async function selectMatchup(match, key, btn, div){

  const maker=document.getElementById("matchMakerSelect").value;

  if(!maker){
    showModal("Select Match Maker first.", "alert");
    return;
  }

document.getElementById("savingMatchOverlay").style.display = "flex";

const data = await api({

  action:"saveMatchupDirect",

  matchMaker:maker,

  redTeam:match.redTeam.map(p=>p.name),

  blueTeam:match.blueTeam.map(p=>p.name)

});

  if(!data.ok){

  document.getElementById("savingMatchOverlay").style.display = "none";

  showModal(data.error, "alert");

  return;

}

// 🔥 ONLY mark selected AFTER SUCCESS
currentMatchKeyFromServer = key; // 🔥 FORCE SYNC IMMEDIATELY

document.querySelectorAll(".matchOption").forEach(card=>{
  card.classList.remove("armedCard");
  card.classList.remove("selectedCard");
});

document.querySelectorAll(".selectMatch").forEach(b=>{
  b.classList.remove("selected");
  b.classList.remove("confirming"); // 🔥 ADD
  b.innerText = "CLICK TO SELECT";
  b.disabled = false; // 🔥 reset disabled state
  b.style.cursor = "pointer";
});

div.classList.add("selectedCard");
btn.classList.add("selected");
btn.innerText = "SELECTED";
btn.disabled = true;
btn.style.cursor = "not-allowed";

/* CHANGE OVERLAY TEXT TO SAVED */

const overlay = document.getElementById("savingMatchOverlay");

overlay.querySelector(".generatingText").innerHTML = "SAVED ✓";

/* WAIT 1 SECOND THEN REDIRECT */

setTimeout(async () => {

  overlay.style.display = "none";

  // 🔥 REFRESH HISTORY
  const historyData = await api({ action: "getHistory" });
  if(historyData.ok){
    matchHistory = historyData.history || [];
  }

  // 🔥 UPDATE GENERATED MATCHUPS (so counts stay accurate)
  const updatedMatchups = generateMatchupsLocal(lastSelectedPlayers, "all");

  lastGeneratedMatchups = updatedMatchups;
  updateGapCounts();
  applyGapFilter();

  // 🔥 NOW GO TO MATCHUP TAB
  const matchupBtn = document.querySelector('.tabButton[onclick*="matchupTab"]');
  showTab("matchupTab", matchupBtn);

  // 🔥 REFRESH CURRENT MATCH DISPLAY
  await loadInitialData();

}, 1000);

}

function startCountdown(expiry){

  if(countdownTimer){
  clearInterval(countdownTimer);
}

  const el=document.getElementById("matchCountdown");

  el.innerHTML = "MATCHUP EXPIRES IN --:--:--";

  countdownTimer = setInterval(()=>{

    const now=new Date();

    const diff=expiry-now;

    if(diff<=0){

  clearInterval(countdownTimer);
  countdownTimer = null;

  el.innerHTML = "MATCHUP EXPIRES IN --:--:--";

  return;

}

    const hours=Math.floor(diff/3600000);
    const mins=Math.floor((diff%3600000)/60000);
    const secs=Math.floor((diff%60000)/1000);

    el.innerHTML=`MATCHUP EXPIRES IN ${hours}:${mins}:${secs}`;

  },1000);

}

function renderPlayers(players){

  const maker = document.getElementById("matchMakerSelect");
  const mapMaker = document.getElementById("mapMatchMakerSelect");
  const list = document.getElementById("playersCheckboxes");

  maker.innerHTML="";
  if(mapMaker) mapMaker.innerHTML="";
  list.innerHTML="";

  // 🔥 ADD PLACEHOLDER (DEFAULT BLANK OPTION)

const placeholder = document.createElement("option");
placeholder.value = "";
placeholder.textContent = "Select Match Maker";
placeholder.disabled = true;
placeholder.selected = true;

maker.appendChild(placeholder);

if(mapMaker){
  mapMaker.appendChild(placeholder.cloneNode(true));
}

const matchMakerPlayers = players.filter(p => p.matchMaker !== false);
  
players.forEach(p=>{

  const div=document.createElement("div");

const savedPlayers = JSON.parse(sessionStorage.getItem("selectedPlayers") || "null");

const isChecked = !savedPlayers || savedPlayers.includes(p.name);

div.innerHTML=`
  <label>
  <input type="checkbox" ${isChecked ? "checked" : ""} value="${p.name}">
  ${p.name}
  <span class="skillMedal">${p.skill}</span>
  </label>
  `;

  div.querySelector("input").addEventListener("change", () => {

    updateSelectedPlayerCount();

    const currentPlayers = Array.from(
      document.querySelectorAll("#playersCheckboxes input:checked")
    ).map(x => x.value).sort();

    const previousPlayers = [...lastSelectedPlayers].sort();

    const isSame =
      currentPlayers.length === previousPlayers.length &&
      currentPlayers.every((v,i)=>v === previousPlayers[i]);

    if(!isSame){
      resetGeneratedMatchups();
    }

/* 🔥 SAVE PLAYER SELECTION */

sessionStorage.setItem(
  "selectedPlayers",
  JSON.stringify(currentPlayers)
);

  });

  list.appendChild(div);

});

matchMakerPlayers.forEach(p=>{

  const opt=document.createElement("option");

  opt.value=p.name;
  opt.innerText=p.name;

  maker.appendChild(opt);

  if(mapMaker){
    const opt2 = opt.cloneNode(true);
    mapMaker.appendChild(opt2);
  }

});

/* 🔥 ADD THIS BLOCK */

const savedGeneratorMaker = sessionStorage.getItem("selectedGeneratorMatchMaker");

if(savedGeneratorMaker){
  maker.value = savedGeneratorMaker;
}

if(mapMaker && globalMapMatchMaker){
  mapMaker.value = globalMapMatchMaker;
}

/* 🔥 AND THIS BLOCK */

maker.onchange = function(){

  sessionStorage.setItem("selectedGeneratorMatchMaker", this.value);

  resetGeneratedMatchups();
  lastSelectedPlayers = [];
  selectedMatchKey = null;

};

if(mapMaker){

  mapMaker.onmousedown = function(e){

    if(isAdminUnlocked()) return;

    e.preventDefault();
    showModal("Unlock admin mode first.", "alert");

  };

  mapMaker.ontouchstart = function(e){

    if(isAdminUnlocked()) return;

    e.preventDefault();
    showModal("Unlock admin mode first.", "alert");

  };

  mapMaker.onchange = async function(){

    if(!isAdminUnlocked()){
      showModal("Unlock admin mode first.", "alert");
      return;
    }

    const selectedName = this.value;

    const res = await api({
      action:"saveGlobalMapMatchMaker",
      matchMaker: selectedName
    });

    if(!res || !res.ok){
      showModal("Could not save map list match maker", "alert");
      return;
    }

    globalMapMatchMaker = res.matchMaker || "";

    renderUpcomingSessionCard(getActiveSessionMaps());

  };

}

updateSelectedPlayerCount();
renderRatingsPreview();

}

function setupRatingsTab(){

  const submitBtn = document.getElementById("submitRatingsBtn");
  const addBtn = document.getElementById("ratingsAddPlayerBtn");
  const removeBtn = document.getElementById("ratingsRemovePlayerBtn");
  const requestCodeBtn = document.getElementById("requestRatingCodeBtn");
  const setupSheetsBtn = document.getElementById("setupRatingSheetsBtn");
  const manualVotingBtn = document.getElementById("manualVotingToggleBtn");
  const applyRatingsBtn = document.getElementById("applyRatingsBtn");
  const infoBtn = document.getElementById("ratingsInfoBtn");
  const infoTooltip = document.getElementById("ratingsInfoTooltip");
  const infoCloseBtn = document.getElementById("ratingsInfoCloseBtn");

  if(submitBtn){
    submitBtn.onclick = () => {
      submitRatings();
    };
  }

  if(addBtn){
    addBtn.onclick = addRatingsPlayer;
  }

  if(removeBtn){
    removeBtn.onclick = removeRatingsPlayer;
  }

  if(requestCodeBtn){
    requestCodeBtn.onclick = requestRatingCode;
  }

  if(setupSheetsBtn){
    setupSheetsBtn.onclick = setupRatingSheets;
  }

  if(manualVotingBtn){
    manualVotingBtn.onclick = toggleManualVotingWindow;
  }

  if(applyRatingsBtn){
    applyRatingsBtn.onclick = applyRatingsToPlayers;
  }

  if(infoBtn && infoTooltip){
    infoBtn.onclick = (e) => {
      e.stopPropagation();
      infoTooltip.classList.toggle("show");
    };

    infoTooltip.onclick = (e) => {
      if(e.target === infoTooltip){
        infoTooltip.classList.remove("show");
      }
    };

    if(infoCloseBtn){
      infoCloseBtn.onclick = (e) => {
        e.stopPropagation();
        infoTooltip.classList.remove("show");
      };
    }
  }

  const raterSelect = document.getElementById("ratingsRaterSelect");

  if(raterSelect){
    raterSelect.onchange = async () => {
      renderRatingsPreview();
      await refreshRatingStatus(raterSelect.value);
    };
  }

}

function getRatingWindowStatus(now = new Date()){

  const year = now.getFullYear();

  const windows = [
    {
      name: "Spring voting",
      opens: new Date(year, 3, 17, 0, 0, 0),
      closes: new Date(year, 3, 30, 23, 59, 59),
      applies: new Date(year, 4, 1, 0, 0, 0)
    },
    {
      name: "Fall voting",
      opens: new Date(year, 10, 17, 0, 0, 0),
      closes: new Date(year, 10, 30, 23, 59, 59),
      applies: new Date(year, 11, 1, 0, 0, 0)
    },
    {
      name: "Next spring voting",
      opens: new Date(year + 1, 3, 17, 0, 0, 0),
      closes: new Date(year + 1, 3, 30, 23, 59, 59),
      applies: new Date(year + 1, 4, 1, 0, 0, 0)
    }
  ];

  const active = windows.find(w => now >= w.opens && now <= w.closes);

  if(active){
    return {
      isOpen: true,
      title: "Voting is open",
      detail: `Closes ${active.closes.toLocaleDateString()}. New ratings apply ${active.applies.toLocaleDateString()}.`
    };
  }

  const next = windows.find(w => now < w.opens);
  const days = Math.ceil((next.opens - now) / 86400000);

  return {
    isOpen: false,
    title: `Voting opens in ${days} ${days === 1 ? "day" : "days"}`,
    detail: `${next.name}: ${next.opens.toLocaleDateString()} - ${next.closes.toLocaleDateString()}`
  };

}

function formatRatingDate(dateValue){

  if(!dateValue) return "";

  const date = new Date(dateValue);

  if(Number.isNaN(date.getTime())) return "";

  return new Intl.DateTimeFormat("en-US", {
    timeZone: RATINGS_TIME_ZONE,
    month: "numeric",
    day: "numeric",
    year: "numeric"
  }).format(date);

}

function formatRatingDateTime(dateValue, label = RATINGS_TIME_ZONE_LABEL){

  const dateText = formatRatingDate(dateValue);

  if(!dateText) return "";

  return `${dateText} ${label}`;

}

function formatBackendRatingStatus(status){

  if(!status || !status.ok){
    return getRatingWindowStatus();
  }

  const zoneLabel = status.timeZoneLabel || RATINGS_TIME_ZONE_LABEL;
  const openDate = formatRatingDateTime(status.opensAt, zoneLabel);
  const closeDate = formatRatingDateTime(status.closesAt, zoneLabel);
  const applyDate = formatRatingDateTime(status.appliesAt, zoneLabel);

  if(status.manualOverride){
    return {
      isOpen: true,
      title: "Manual voting is open",
      detail: `Admin override closes ${closeDate}.`
    };
  }

  if(status.hasVoted){
    return {
      isOpen: status.isOpen,
      title: "You already voted this cycle",
      detail: applyDate ? `New ratings apply ${applyDate}.` : "Your vote has been recorded."
    };
  }

  if(status.isOpen){
    return {
      isOpen: true,
      title: "Voting is open",
      detail: `Closes ${closeDate}. New ratings apply ${applyDate}.`
    };
  }

  const days = Number(status.daysUntilOpen || 0);

  return {
    isOpen: false,
    title: `Voting opens in ${days} ${days === 1 ? "day" : "days"}`,
    detail: `${status.name || "Next voting"}: ${openDate} - ${closeDate}`
  };

}

function formatRatingNumber(value){

  if(value === "" || value === null || typeof value === "undefined") return "-";

  const numeric = Number(value);

  if(Number.isNaN(numeric)) return "-";

  return Number.isInteger(numeric) ? String(numeric) : numeric.toFixed(1);

}

function updateManualVotingButton(){

  const btn = document.getElementById("manualVotingToggleBtn");

  if(!btn) return;

  const manualOpen = currentRatingStatus && currentRatingStatus.manualOverride;

  btn.textContent = manualOpen ? "END 2H OVERRIDE" : "UNLOCK VOTING 2H";
  btn.classList.toggle("btn-orange", !!manualOpen);
  btn.classList.toggle("btn-blue", !manualOpen);

}

async function refreshRatingStatus(rater = ""){

  try{

    const res = await api({
      action:"getRatingStatus",
      rater:rater
    });

    if(res && res.ok){
      currentRatingStatus = res;
    }else{
      currentRatingStatus = null;
    }

  }catch(err){

    currentRatingStatus = null;

  }

  renderRatingsPreview();
  updateManualVotingButton();

}

function getRatingsCodeFormValues(){

  const raterSelect = document.getElementById("ratingsRaterSelect");
  const emailInput = document.getElementById("ratingsEmailInput");
  const codeInput = document.getElementById("ratingsCodeInput");

  return {
    rater: raterSelect ? raterSelect.value : "",
    email: emailInput ? emailInput.value.trim() : "",
    code: codeInput ? codeInput.value.trim() : ""
  };

}

async function requestRatingCode(){

  const values = getRatingsCodeFormValues();
  const statusEl = document.getElementById("ratingsCodeStatus");

  if(!values.rater){
    showModal("Select your name before requesting a code.", "alert");
    return;
  }

  if(!values.email){
    showModal("Enter your email address before requesting a code.", "alert");
    return;
  }

  showBusy("SENDING CODE");

  try{

    const res = await api({
      action:"requestRatingCode",
      rater:values.rater,
      email:values.email
    });

    if(!res || !res.ok){
      showModal((res && res.error) || "Could not request voting code.", "alert");
      return;
    }

    if(statusEl){
      statusEl.textContent = `Code sent to ${values.email}. It expires in 1 hour.`;
    }

    showModal("Voting code sent. Check your email.", "alert");

  }catch(err){

    if(statusEl){
      statusEl.textContent = "Code request is not connected in this deployed version yet.";
    }

    showModal("Code request is not connected in this deployed version yet.", "alert");

  }finally{

    hideBusy();

  }

}

async function setupRatingSheets(){

  if(!isAdminUnlocked()){
    showModal("Unlock admin mode first.", "alert");
    return;
  }

  const pass = await getAdminPassword();
  if(!pass) return;

  try{

    const res = await api({
      action:"setupRatingSheets",
      password:pass
    });

    if(!res || !res.ok){
      showModal((res && res.error) || "Could not setup rating sheets.", "alert");
      return;
    }

    sessionStorage.setItem("adminPass", pass);
    updateAdminBar();

    showModal("Rating sheets are ready.", "alert");

  }catch(err){

    showModal("Rating sheet setup is not connected in this deployed version yet.", "alert");

  }

}

async function toggleManualVotingWindow(){

  if(!isAdminUnlocked()){
    showModal("Unlock admin mode first.", "alert");
    return;
  }

  const pass = await getAdminPassword();
  if(!pass) return;

  const manualOpen = currentRatingStatus && currentRatingStatus.manualOverride;
  const enabled = !manualOpen;
  const message = enabled
    ? "Unlock voting for 2 hours?"
    : "End the 2-hour voting override?";

  const confirmed = await showModal(message, "confirm");
  if(!confirmed) return;

  try{

    const res = await api({
      action:"setManualVotingWindow",
      password:pass,
      enabled:enabled
    });

    if(!res || !res.ok){
      showModal((res && res.error) || "Could not update manual voting.", "alert");
      return;
    }

    sessionStorage.setItem("adminPass", pass);
    updateAdminBar();
    await refreshRatingStatus(document.getElementById("ratingsRaterSelect")?.value || "");

    showModal(enabled ? "Voting unlocked for 2 hours." : "Manual override ended.", "alert");

  }catch(err){

    showModal("Manual voting control is not connected in this deployed version yet.", "alert");

  }

}

async function applyRatingsToPlayers(){

  if(!isAdminUnlocked()){
    showModal("Unlock admin mode first.", "alert");
    return;
  }

  const confirmed = await showModal(
    "Apply latest rating skills to the Players sheet? A backup will be created first.",
    "confirm"
  );

  if(!confirmed) return;

  const pass = await getAdminPassword();
  if(!pass) return;

  try{

    const res = await api({
      action:"applyLatestRatingsToPlayers",
      password:pass
    });

    if(!res || !res.ok){
      showModal((res && res.error) || "Could not apply ratings.", "alert");
      return;
    }

    sessionStorage.setItem("adminPass", pass);
    updateAdminBar();
    syncPlayersFromResponse(res.players);

    showModal(
      `Ratings applied. ${res.updatedCount || 0} players updated. Backup: ${res.backupSheet || "created"}.`,
      "alert"
    );

  }catch(err){

    showModal("Apply ratings is not connected in this deployed version yet.", "alert");

  }

}

function syncPlayersFromResponse(players){

  if(!Array.isArray(players)) return;

  allPlayers = players;
  populatePlayers(allPlayers);
  renderRatingsPreview();

}

async function addRatingsPlayer(){

  if(!isAdminUnlocked()){
    showModal("Unlock admin mode first.", "alert");
    return;
  }

  const pass = await getAdminPassword();
  if(!pass) return;

  const name = await showModal(
    "Enter player name",
    "confirm",
    "Confirm",
    "Cancel",
    true,
    "text",
    "Player name"
  );

  if(!name) return;

  const skillValue = await showModal(
    "Enter starting skill number",
    "confirm",
    "Confirm",
    "Cancel",
    true,
    "number",
    "0-10"
  );

  if(skillValue === null) return;

  const skill = parseInt(skillValue, 10);

  if(Number.isNaN(skill)){
    showModal("Skill must be a number.", "alert");
    return;
  }

  try{

    const res = await api({
      action:"addRatingPlayer",
      password:pass,
      name:name,
      skill:skill
    });

    if(!res || !res.ok){
      showModal((res && res.error) || "Could not add player.", "alert");
      return;
    }

    sessionStorage.setItem("adminPass", pass);
    updateAdminBar();
    syncPlayersFromResponse(res.players);

    showModal("Player added to ratings.", "alert");

  }catch(err){

    showModal("Add player is not connected in this deployed version yet.", "alert");

  }

}

async function removeRatingsPlayer(){

  if(!isAdminUnlocked()){
    showModal("Unlock admin mode first.", "alert");
    return;
  }

  const pass = await getAdminPassword();
  if(!pass) return;

  const name = await showModal(
    "Enter player name to remove from active ratings",
    "confirm",
    "Confirm",
    "Cancel",
    true,
    "text",
    "Player name"
  );

  if(!name) return;

  try{

    const res = await api({
      action:"deactivateRatingPlayer",
      password:pass,
      name:name
    });

    if(!res || !res.ok){
      showModal((res && res.error) || "Could not remove player.", "alert");
      return;
    }

    sessionStorage.setItem("adminPass", pass);
    updateAdminBar();
    syncPlayersFromResponse(res.players);

    showModal("Player removed from active ratings.", "alert");

  }catch(err){

    showModal("Remove player is not connected in this deployed version yet.", "alert");

  }

}

function renderRatingsPreview(){

  const raterSelect = document.getElementById("ratingsRaterSelect");
  const rows = document.getElementById("ratingsRows");
  const statusText = document.getElementById("ratingsStatusText");
  const statusDate = document.getElementById("ratingsStatusDate");
  const tableShell = document.querySelector(".ratingsTableShell");
  const tableHeader = document.querySelector(".ratingsTableHeader");
  const raterCard = document.querySelector(".ratingsRaterCard");
  const submitWrap = document.querySelector(".ratingsSubmitWrap");

  if(!raterSelect || !rows) return;

  const players = Array.isArray(allPlayers) ? [...allPlayers] : [];
  const status = currentRatingStatus
    ? formatBackendRatingStatus(currentRatingStatus)
    : getRatingWindowStatus();
  const isVotingOpen = currentRatingStatus
    ? !!currentRatingStatus.isOpen
    : !!status.isOpen;

  if(statusText) statusText.textContent = status.title;
  if(statusDate) statusDate.textContent = status.detail;
  if(tableShell) tableShell.classList.toggle("ratingsReadOnlyMode", !isVotingOpen);
  if(raterCard) raterCard.style.display = isVotingOpen ? "" : "none";
  if(submitWrap) submitWrap.style.display = isVotingOpen ? "" : "none";
  if(tableHeader){
    tableHeader.innerHTML = isVotingOpen
      ? `
        <div>PLAYER</div>
        <div class="ratingsBlitzHeader">BLITZ <span>0-10</span></div>
        <div class="ratingsCtfHeader">CTF <span>0-10</span></div>
      `
      : `
        <div>PLAYER</div>
        <div>CURRENT SKILL</div>
        <div class="ratingsBlitzHeader">BLITZ <span>MEDIAN</span></div>
        <div class="ratingsCtfHeader">CTF <span>MEDIAN</span></div>
        <div>VOTES</div>
      `;
  }

  const selectedRater = raterSelect.value;

  raterSelect.innerHTML = "";

  const placeholder = document.createElement("option");
  placeholder.value = "";
  placeholder.textContent = "Select your name";
  placeholder.disabled = true;
  placeholder.selected = true;
  raterSelect.appendChild(placeholder);

  players
    .slice()
    .sort((a,b)=>a.name.localeCompare(b.name))
    .forEach(player => {
      const option = document.createElement("option");
      option.value = player.name;
      option.textContent = player.name;
      raterSelect.appendChild(option);
    });

  if(selectedRater){
    raterSelect.value = selectedRater;
  }

  rows.innerHTML = "";

  if(players.length === 0){
    rows.innerHTML = `<div class="ratingsEmpty">No players loaded yet.</div>`;
    return;
  }

  if(!isVotingOpen){
    renderRatingsReadOnlyRows(players);
    return;
  }

  renderRatingsVotingRows(players, selectedRater);

}

function renderRatingsReadOnlyRows(players){

  const rows = document.getElementById("ratingsRows");
  if(!rows) return;

  const latestRows = currentRatingStatus && Array.isArray(currentRatingStatus.latestRatings)
    ? currentRatingStatus.latestRatings
    : [];

  const latestByPlayer = {};

  latestRows.forEach(rating => {
    if(rating && rating.player){
      latestByPlayer[rating.player] = rating;
    }
  });

  rows.innerHTML = "";

  players
    .slice()
    .sort((a,b)=>a.name.localeCompare(b.name))
    .forEach((player, index) => {
      const latest = latestByPlayer[player.name] || {};
      const row = document.createElement("div");
      row.className = "ratingsRow ratingsReadOnlyRow";
      row.setAttribute("data-player", player.name);

      row.innerHTML = `
        <div class="ratingsPlayerCell">
          <span class="ratingsPlayerNumber">${index + 1}</span>
          <span class="ratingsPlayerName">${player.name}</span>
        </div>

        <div class="ratingsReadonlyValue">
          <span class="ratingsReadonlyLabel">Current Skill</span>
          <strong>${formatRatingNumber(player.skill)}</strong>
        </div>

        <div class="ratingsReadonlyValue ratingsReadonlyBlitz">
          <span class="ratingsReadonlyLabel">Blitz Median</span>
          <strong>${formatRatingNumber(latest.blitzMedian)}</strong>
        </div>

        <div class="ratingsReadonlyValue ratingsReadonlyCtf">
          <span class="ratingsReadonlyLabel">CTF Median</span>
          <strong>${formatRatingNumber(latest.ctfMedian)}</strong>
        </div>

        <div class="ratingsReadonlyValue">
          <span class="ratingsReadonlyLabel">Votes</span>
          <strong>${formatRatingNumber(latest.voteCount)}</strong>
        </div>
      `;

      rows.appendChild(row);
    });

}

function renderRatingsVotingRows(players, selectedRater){

  const rows = document.getElementById("ratingsRows");
  if(!rows) return;

  rows.innerHTML = "";

  players
    .slice()
    .sort((a,b)=>a.name.localeCompare(b.name))
    .forEach((player, index) => {
      const row = document.createElement("div");
      row.className = "ratingsRow";
      row.setAttribute("data-player", player.name);

      const isSelf = selectedRater && selectedRater === player.name;

      if(isSelf){
        row.classList.add("ratingsSelfRow");
      }

      row.innerHTML = `
        <div class="ratingsPlayerCell">
          <span class="ratingsPlayerNumber">${index + 1}</span>
          <span class="ratingsPlayerName">${player.name}</span>
          ${isSelf ? `<span class="ratingsSelfBadge">YOU</span>` : ""}
        </div>

        <div class="ratingsSliderCell">
          <input class="ratingsValueBox ratingsBlitzValue" type="number" min="0" max="10" step="1" value="5">
          <span class="ratingsRangeMin">0</span>
          <input class="ratingsSlider ratingsBlitzSlider" type="range" min="0" max="10" step="1" value="5">
          <span class="ratingsRangeMax">10</span>
        </div>

        <div class="ratingsSliderCell">
          <input class="ratingsValueBox ratingsCtfValue" type="number" min="0" max="10" step="1" value="5">
          <span class="ratingsRangeMin">0</span>
          <input class="ratingsSlider ratingsCtfSlider" type="range" min="0" max="10" step="1" value="5">
          <span class="ratingsRangeMax">10</span>
        </div>
      `;

      row.querySelectorAll(".ratingsSliderCell").forEach(cell => {
        const box = cell.querySelector(".ratingsValueBox");
        const slider = cell.querySelector(".ratingsSlider");

        if(isSelf){
          box.disabled = true;
          slider.disabled = true;
        }

        slider.addEventListener("input", () => {
          box.value = slider.value;
        });

        box.addEventListener("input", () => {
          let nextValue = parseInt(box.value, 10);

          if(Number.isNaN(nextValue)) nextValue = 0;
          nextValue = Math.max(0, Math.min(10, nextValue));

          box.value = nextValue;
          slider.value = nextValue;
        });
      });

      rows.appendChild(row);
    });

}

function collectRatingsFormData(){

  const raterSelect = document.getElementById("ratingsRaterSelect");
  const rows = Array.from(document.querySelectorAll("#ratingsRows .ratingsRow"));

  if(!raterSelect || !raterSelect.value){
    return {
      ok: false,
      error: "Select your name before submitting ratings."
    };
  }

  const ratings = [];

  rows.forEach(row => {
    const player = row.getAttribute("data-player");

    if(!player || player === raterSelect.value) return;

    const blitz = row.querySelector(".ratingsBlitzSlider");
    const ctf = row.querySelector(".ratingsCtfSlider");

    ratings.push({
      ratedPlayer: player,
      blitz: Number(blitz ? blitz.value : 0),
      ctf: Number(ctf ? ctf.value : 0)
    });
  });

  if(ratings.length === 0){
    return {
      ok: false,
      error: "There are no other players to rate."
    };
  }

  const invalid = ratings.some(r =>
    !Number.isInteger(r.blitz) ||
    !Number.isInteger(r.ctf) ||
    r.blitz < 0 ||
    r.blitz > 10 ||
    r.ctf < 0 ||
    r.ctf > 10
  );

  if(invalid){
    return {
      ok: false,
      error: "Ratings must be whole numbers from 0 to 10."
    };
  }

  return {
    ok: true,
    rater: raterSelect.value,
    ratings: ratings
  };

}

function submitRatingsPreview(formData){

  const data = formData || collectRatingsFormData();

  if(!data.ok){
    showModal(data.error, "alert");
    return;
  }

  const ratings = data.ratings;
  const combinedScores = ratings.flatMap(r => [r.blitz, r.ctf]);
  const average = combinedScores.reduce((sum, value) => sum + value, 0) / combinedScores.length;

  sessionStorage.setItem("ratingsPreview", JSON.stringify({
    rater: data.rater,
    ratings: ratings,
    previewAverage: Math.round(average * 10) / 10,
    savedAt: new Date().toISOString()
  }));

  showModal(
    `Offline preview saved for ${data.rater}. ${ratings.length} players rated. Average score: ${Math.round(average * 10) / 10}.`,
    "alert"
  );

}

async function submitRatings(){

  const formData = collectRatingsFormData();
  const codeValues = getRatingsCodeFormValues();

  if(!formData.ok){
    showModal(formData.error, "alert");
    return;
  }

  if(!codeValues.code){
    showModal("Enter your one-hour voting code before submitting.", "alert");
    return;
  }

  showBusy("SUBMITTING RATINGS");

  try{

    const res = await api({
      action:"submitRatings",
      rater:formData.rater,
      code:codeValues.code,
      ratings:formData.ratings
    });

    if(!res || !res.ok){
      showModal((res && res.error) || "Could not submit ratings.", "alert");
      return;
    }

    sessionStorage.removeItem("ratingsPreview");

    await refreshRatingStatus(formData.rater);

    showModal(
      `Ratings submitted successfully. ${res.submittedCount || formData.ratings.length} players rated.`,
      "alert"
    );

  }catch(err){

    submitRatingsPreview(formData);

  }finally{

    hideBusy();

  }

}

function markAdminDirty(isDirty = true){
  adminHasUnsavedChanges = isDirty;

  const status = document.getElementById("adminDirtyStatus");
  if(!status) return;

  status.style.display = isDirty ? "inline-block" : "none";
}

function createSkillAdjuster(value){

  const wrapper = document.createElement("div");
  wrapper.className = "skillAdjuster";

  const downBtn = document.createElement("button");
  downBtn.className = "skillAdjustBtn";
  downBtn.type = "button";
  downBtn.textContent = "▼";

  const cell = document.createElement("div");
  cell.className = "skillValueCell";
  cell.contentEditable = "true";
  cell.textContent = value;

  const upBtn = document.createElement("button");
  upBtn.className = "skillAdjustBtn";
  upBtn.type = "button";
  upBtn.textContent = "▲";

  const clampAndSet = (nextValue) => {
    const numeric = parseInt(nextValue, 10);
    cell.textContent = Number.isNaN(numeric) ? "0" : String(numeric);
    markAdminDirty(true);
  };

  downBtn.onclick = () => {
    const current = parseInt(cell.textContent.trim(), 10) || 0;
    cell.textContent = String(current - 1);
    markAdminDirty(true);
  };

  upBtn.onclick = () => {
    const current = parseInt(cell.textContent.trim(), 10) || 0;
    cell.textContent = String(current + 1);
    markAdminDirty(true);
  };

  cell.addEventListener("input", () => {
    markAdminDirty(true);
  });

  cell.addEventListener("blur", () => {
    clampAndSet(cell.textContent.trim());
  });

  wrapper.appendChild(downBtn);
  wrapper.appendChild(cell);
  wrapper.appendChild(upBtn);

  return wrapper;
}

async function openAdminTab(btn){

  showTab("adminTab", btn);

  /* SHOW LOADING OVERLAY */

  document.getElementById("historyLoadingOverlay").style.display = "flex";

  const data = await api({
    action:"getPlayersAdmin"
  });

  /* HIDE LOADING OVERLAY */

  document.getElementById("historyLoadingOverlay").style.display = "none";

  if(!data.ok){
    showModal("Failed loading players", "alert");
    return;
  }

  const table = document.querySelector("#adminTable tbody");

  table.innerHTML="";

  data.players.forEach(p=>{

    const row=document.createElement("tr");

    row.innerHTML=`

    <td contenteditable="true">${p.name}</td>

    <td class="skillCell"></td>

    <td><input type="checkbox" class="matchMakerEligible" ${p.matchMaker !== false ? "checked" : ""}></td>

    <td><button class="btn btn-red remove">REMOVE</button></td>

    `;

    row.cells[0].addEventListener("input", () => {
      markAdminDirty(true);
    });

    row.querySelector(".skillCell").appendChild(createSkillAdjuster(p.skill));

    row.querySelector(".matchMakerEligible").addEventListener("change", () => {
      markAdminDirty(true);
    });

    row.querySelector(".remove").onclick=()=>{

      row.remove();
      updatePlayerCount();
      markAdminDirty(true);

    };

    table.appendChild(row);

  });

  updatePlayerCount();
  setupAdminSorting();
  sortAdminTableRows();
  markAdminDirty(false);

}

function updatePlayerCount(){

  const rows = document.querySelectorAll("#adminTable tbody tr").length;

  document.getElementById("playerCount").innerText = "Players: " + rows;

}

function sortAdminTableRows(){

  const tbody = document.querySelector("#adminTable tbody");
  if(!tbody || !currentAdminSort.key) return;

  const rows = Array.from(tbody.querySelectorAll("tr"));

  rows.sort((a, b) => {

    let valA = "";
    let valB = "";

    if(currentAdminSort.key === "name"){
      valA = a.cells[0].innerText.trim().toLowerCase();
      valB = b.cells[0].innerText.trim().toLowerCase();
    }

    if(currentAdminSort.key === "skill"){
      valA = parseInt(a.querySelector(".skillValueCell")?.innerText.trim() || "0", 10);
      valB = parseInt(b.querySelector(".skillValueCell")?.innerText.trim() || "0", 10);
    }

    if(valA < valB) return currentAdminSort.direction === "asc" ? -1 : 1;
    if(valA > valB) return currentAdminSort.direction === "asc" ? 1 : -1;
    return 0;

  });

  rows.forEach(row => tbody.appendChild(row));

}

function updateAdminSortIndicators(){

  const headers = document.querySelectorAll("#adminTable th[data-sort]");

  headers.forEach(th => {
    const key = th.dataset.sort;
    const base = key === "name" ? "Name" : "Skill";

    if(currentAdminSort.key === key){
      th.innerText = base + (currentAdminSort.direction === "asc" ? " ▲" : " ▼");
    }else{
      th.innerText = base + " ▴▾";
    }
  });

}

function setupAdminSorting(){

  const headers = document.querySelectorAll("#adminTable th[data-sort]");

  headers.forEach(th => {
    th.onclick = () => {
      const key = th.dataset.sort;

      if(currentAdminSort.key === key){
        currentAdminSort.direction =
          currentAdminSort.direction === "asc" ? "desc" : "asc";
      }else{
        currentAdminSort.key = key;
        currentAdminSort.direction = "asc";
      }

      sortAdminTableRows();
      updateAdminSortIndicators();
    };
  });

  updateAdminSortIndicators();

}

function addAdminPlayerRow(){

  const table = document.querySelector("#adminTable tbody");

  const row = document.createElement("tr");

  row.innerHTML = `

  <td contenteditable="true"></td>

  <td class="skillCell"></td>

  <td><input type="checkbox" class="matchMakerEligible" checked></td>

  <td><button class="btn btn-red remove">REMOVE</button></td>

  `;

  row.cells[0].addEventListener("input", () => {
    markAdminDirty(true);
  });

  row.querySelector(".skillCell").appendChild(createSkillAdjuster(0));

  row.querySelector(".matchMakerEligible").addEventListener("change", () => {
    markAdminDirty(true);
  });

  row.querySelector(".remove").onclick = () => {

    row.remove();
    updatePlayerCount();
    markAdminDirty(true);

  };

  table.appendChild(row);

  updatePlayerCount();
  markAdminDirty(true);

}

document.getElementById("addPlayer").onclick = addAdminPlayerRow;

document.getElementById("savePlayers").onclick = savePlayers;

async function savePlayers(){

const pass = await getAdminPassword();
if(!pass) return;

  document.getElementById("savingOverlay").style.display = "flex";

  const players = [];

  document.querySelectorAll("#adminTable tbody tr").forEach(row=>{

    const name = row.cells[0].innerText.trim();
    const skillCell = row.querySelector(".skillValueCell");
    const skill = parseInt(skillCell ? skillCell.innerText.trim() : "0", 10);
    const matchMakerEligible = row.querySelector(".matchMakerEligible");

    if(!name) return;

    players.push({
      name:name,
      skill:Number.isNaN(skill) ? 0 : skill,
      matchMaker: matchMakerEligible ? matchMakerEligible.checked : true
    });

  });

  const data = await api({

    action:"savePlayersAdmin",

    password:pass,

    players:players

  });

  document.getElementById("savingOverlay").style.display = "none";

if(!data.ok){

  showModal(data.error, "alert");
  return;

}

// 🔥 ADD THIS
sessionStorage.setItem("adminPass", pass);
updateAdminBar();

markAdminDirty(false);

showModal("Players saved successfully", "alert");

openAdminTab();

}

async function openHistoryTab(btn){

  showTab("historyTab", btn);

  document.getElementById("historyLoadingOverlay").style.display = "flex";

  const data = await api({
    action:"getHistory"
  });

  if(!data.ok){

    showModal("Could not load history", "alert");
    return;

  }

  renderHistory(data.history);

  setupHistorySorting();
  updateSortIndicators();

  document.getElementById("historyLoadingOverlay").style.display = "none";

}

function renderHistory(history){

const tbody = document.getElementById("historyTableBody");

tbody.innerHTML = "";

if(!history || history.length === 0){
  tbody.innerHTML = `<tr><td colspan="5">No match history yet.</td></tr>`;
  return;
}

/* 🔥 COUNT MATCHUP FREQUENCY */

const counts = {};

history.forEach(h => {

  const red = h.redTeam.split(", ").sort().join(",");
  const blue = h.blueTeam.split(", ").sort().join(",");

  const key1 = red + "|" + blue;
  const key2 = blue + "|" + red;

  if(counts[key1] || counts[key2]){
    counts[key1] = (counts[key1] || counts[key2]) + 1;
  }else{
    counts[key1] = 1;
  }

});

/* 🔥 DEFAULT SORT (NEWEST FIRST) */

/* 🔥 APPLY CURRENT SORT */

history.sort((a,b)=>{

  let valA, valB;

  switch(currentHistorySort.key){

    case "date":
      valA = new Date(a.selectedAt);
      valB = new Date(b.selectedAt);
      break;

    case "mid":
      valA = parseInt(a.MID || 0);
      valB = parseInt(b.MID || 0);
      break;

    case "maker":
      valA = a.matchMaker.toLowerCase();
      valB = b.matchMaker.toLowerCase();
      break;

    case "picked":

      const getCount = (m) => {
        const r = m.redTeam.split(", ").sort().join(",");
        const b = m.blueTeam.split(", ").sort().join(",");
        return counts[r+"|"+b] || counts[b+"|"+r] || 0;
      };

      valA = getCount(a);
      valB = getCount(b);
      break;

    case "gap":
      valA = a.skillGap;
      valB = b.skillGap;
      break;

    default:
      valA = 0;
      valB = 0;
  }

  if(valA < valB) return currentHistorySort.direction === "asc" ? -1 : 1;
  if(valA > valB) return currentHistorySort.direction === "asc" ? 1 : -1;
  return 0;

});

history.forEach(match => {

  const row = document.createElement("tr");

  const key1 = match.redTeam.split(", ").sort().join(",") + "|" + match.blueTeam.split(", ").sort().join(",");
  const key2 = match.blueTeam.split(", ").sort().join(",") + "|" + match.redTeam.split(", ").sort().join(",");

  const count = counts[key1] || counts[key2] || 0;

row.innerHTML = `
  <td>
    <span class="expandIcon">▶</span>
    ${match.selectedAtDisplay || formatDate(match.selectedAt)}
  </td>
  <td>${match.MID ? "MID_" + String(match.MID).replace("MID_","").padStart(4,"0") : "----"}</td>
  <td>${match.matchMaker}</td>
  <td>${count}</td>
  <td>${match.skillGap}</td>
`;

  /* 🔥 DETAIL ROW */

  const detailRow = document.createElement("tr");

  detailRow.className = "historyDetailRow";
  detailRow.style.display = "none";

detailRow.innerHTML = `
  <td colspan="5">

    <div class="matchCard">

      <div class="teamsRow">

        <div class="team red">
          <div class="teamTitle">
            RED TEAM <span class="teamBadge">${match.redSkill}</span>
          </div>

          <div class="teamPlayers">
${match.redTeam
  .split(", ")
  .sort((a,b)=>a.localeCompare(b))
  .map(name => {

  const player = allPlayers.find(p => p.name === name);

  return `
    <div class="playerRow">
      ${name}
      <span class="skillMedal">${player ? player.skill : ""}</span>
    </div>
  `;

}).join("")}
          </div>
        </div>

        <div class="vs">VS</div>

        <div class="team blue">
          <div class="teamTitle">
            BLUE TEAM <span class="teamBadge">${match.blueSkill}</span>
          </div>

          <div class="teamPlayers">
${match.blueTeam
  .split(", ")
  .sort((a,b)=>a.localeCompare(b))
  .map(name => {

  const player = allPlayers.find(p => p.name === name);

  return `
    <div class="playerRow">
      ${name}
      <span class="skillMedal">${player ? player.skill : ""}</span>
    </div>
  `;

}).join("")}
          </div>
        </div>

      </div>

      <div class="matchFooter">
        <span class="diff diff-${match.skillGap}">
          Difference: ${match.skillGap}
        </span>
      </div>

    </div>

  </td>
`;

  /* 🔥 CLICK TO TOGGLE */

row.onclick = () => {

  const isOpen = detailRow.style.display === "table-row";

  detailRow.style.display = isOpen ? "none" : "table-row";

  const icon = row.querySelector(".expandIcon");

  if(icon){
    icon.innerText = isOpen ? "▶" : "▼";
  }

};

  tbody.appendChild(row);
  tbody.appendChild(detailRow);

});
}

function setupHistorySorting(){

  const headers = document.querySelectorAll("#historyTable th");

  headers.forEach(th => {

    th.onclick = () => {

      const key = th.dataset.sort;

      if(!key) return;

      // 🔥 TOGGLE DIRECTION
      if(currentHistorySort.key === key){
        currentHistorySort.direction =
          currentHistorySort.direction === "asc" ? "desc" : "asc";
      }else{
        currentHistorySort.key = key;
        currentHistorySort.direction = "asc";
      }

      // 🔥 RE-RENDER WITH SORT
      renderHistory([...matchHistory]);

      updateSortIndicators();

    };

  });

}

function updateSortIndicators(){

  const headers = document.querySelectorAll("#historyTable th");

  headers.forEach(th => {

    const key = th.dataset.sort;

    if(!key) return;

th.innerHTML = th.innerText
  .replace(" ↑", "")
  .replace(" ↓", "")
  .replace(" ⇅", "")
  .replace(" ▴▾", "")
  .replace(" ▲", "")
  .replace(" ▼", "");

if(key === currentHistorySort.key){

  const arrow = currentHistorySort.direction === "asc" ? " ▲" : " ▼";

  th.innerHTML = th.innerText + arrow;

} else {

  th.innerHTML = th.innerText + " ▴▾";

}

  });

}

function formatDate(date){

  const d = new Date(date);

  if(Number.isNaN(d.getTime())) return "";

  const dateText = new Intl.DateTimeFormat("en-US", {
    timeZone: RATINGS_TIME_ZONE,
    month: "numeric",
    day: "numeric",
    year: "numeric",
    hour: "numeric",
    minute: "2-digit",
    second: "2-digit",
    hour12: true
  }).format(d);

  return `${dateText} ${RATINGS_TIME_ZONE_LABEL}`;

}

document.getElementById("clearHistoryBtn").onclick = clearHistory;

async function clearHistory(){

const pass = await getAdminPassword();
if(!pass) return;

  /* SHOW CLEARING OVERLAY */

  document.getElementById("clearHistoryOverlay").style.display = "flex";

  const data = await api({

    action:"clearHistory",

    password:pass

  });

if(!data.ok){

  document.getElementById("clearHistoryOverlay").style.display = "none";

  showModal(data.error, "alert");
  return;

}

// 🔥 ADD THIS
sessionStorage.setItem("adminPass", pass);
updateAdminBar();

/* CHANGE OVERLAY TEXT TO CLEARED */

const overlay = document.getElementById("clearHistoryOverlay");

overlay.querySelector("div:last-child").innerHTML = "CLEARED & SESSION RESET ✓";

/* WAIT THEN RESET UI */

setTimeout(async () => {

  overlay.style.display = "none";

  document.getElementById("historyTableBody").innerHTML = `
  <tr><td colspan="5">No match history yet.</td></tr>
`;

  /* 🔥 RESET GENERATOR STATE */
  resetGeneratedMatchups();
  lastSelectedPlayers = [];
  selectedMatchKey = null;

  /* 🔥 RESET MATCH MAKER */
  document.getElementById("matchMakerSelect").selectedIndex = 0;
  sessionStorage.removeItem("selectedMatchMaker");

  /* 🔥 RESET PLAYER CHECKBOXES (all checked) */
  document.querySelectorAll("#playersCheckboxes input").forEach(cb=>{
    cb.checked = true;
  });

  updateSelectedPlayerCount();

  /* 🔥 FORCE MATCHUP REFRESH */
  await loadInitialData();

  /* RESET TEXT BACK */

  overlay.querySelector("div:last-child").innerHTML = "CLEARING HISTORY<span class='dots'></span>";

}, 1000);

}

function startMatchAutoRefresh(){

  setInterval(async ()=>{

    try{

      const data = await api({
        action:"getInitialData"
      });

      if(data.ok){

        renderMatchup(data.currentMatchup);

      }

    }catch(e){

      console.log("Auto refresh error");

    }

  },10000);

}

function generateMatchupsLocal(selectedPlayers, filterGap){

  const players = allPlayers.filter(p => selectedPlayers.includes(p.name));

  const size = Math.floor(players.length / 2);

const combos = getCombinationsLocal(players, size);

const results = [];

const seen = new Set();

combos.forEach(red => {

const blue = players.filter(p => !red.includes(p));

const redSkill = red.reduce((s,p)=>s+p.skill,0);
const blueSkill = blue.reduce((s,p)=>s+p.skill,0);

const gap = Math.abs(redSkill - blueSkill);

/* Hide matchups with skill gap greater than 4 */

if(gap > 4) return;

/* Prevent mirrored duplicates */

const redNames = red.map(p=>p.name).sort().join(",");
const blueNames = blue.map(p=>p.name).sort().join(",");

const key1 = redNames + "|" + blueNames;
const key2 = blueNames + "|" + redNames;

if(seen.has(key1) || seen.has(key2)) return;

seen.add(key1);

/* CALCULATE PICK COUNT FROM HISTORY */

const redNamesSorted = red.map(p=>p.name).sort().join(",");
const blueNamesSorted = blue.map(p=>p.name).sort().join(",");

let pickCount = 0;

matchHistory.forEach(h=>{

  const hRed = h.redTeam.split(", ").sort().join(",");
  const hBlue = h.blueTeam.split(", ").sort().join(",");

  if(
    (hRed === redNamesSorted && hBlue === blueNamesSorted) ||
    (hRed === blueNamesSorted && hBlue === redNamesSorted)
  ){
    pickCount++;
  }

});

const redSorted = [...red].sort((a,b)=>a.name.localeCompare(b.name));
const blueSorted = [...blue].sort((a,b)=>a.name.localeCompare(b.name));

results.push({
  redTeam:redSorted,
  blueTeam:blueSorted,
  redSkill:redSkill,
  blueSkill:blueSkill,
  skillGap:gap,
  pickCount:pickCount
});

  });

  return results;

}

function getCombinationsLocal(arr,size){

  const result = [];

  function helper(start,combo){

    if(combo.length === size){
      result.push([...combo]);
      return;
    }

    for(let i=start;i<arr.length;i++){

      combo.push(arr[i]);
      helper(i+1,combo);
      combo.pop();

    }

  }

  helper(0,[]);
  return result;

}

function applyGapFilter(){

  const filter = document.querySelector('input[name="gapFilter"]:checked').value;

  let filtered = lastGeneratedMatchups;

  if(filter !== "all"){

    const gapValue = Number(filter);

    filtered = lastGeneratedMatchups.filter(m => m.skillGap === gapValue);

  }

/* 🔥 BLITZ FILTER (ONLY SHOW ADVANTAGED SMALL TEAM) */

if(blitzEnabled){

  filtered = filtered.filter(m => {

    const small =
      m.redTeam.length < m.blueTeam.length ? m.redTeam : m.blueTeam;

    const large =
      m.redTeam.length > m.blueTeam.length ? m.redTeam : m.blueTeam;

    /* if equal teams, ignore */
    if(m.redTeam.length === m.blueTeam.length) return false;

    const smallSkill = small.reduce((s,p)=>s+p.skill,0);
    const largeSkill = large.reduce((s,p)=>s+p.skill,0);

    return smallSkill > largeSkill;

  });

}

/* Restore normal sorting when BLITZ is OFF */

if(!blitzEnabled){

  filtered.sort((a,b)=>a.skillGap - b.skillGap);

}  
  
  renderGeneratedMatchups(filtered);

}

function updateSelectedPlayerCount(){

  const count = document.querySelectorAll("#playersCheckboxes input:checked").length;

  document.getElementById("selectedPlayerCount").innerText = count;

}

function updateGapCounts(){

  const radios = document.querySelectorAll('input[name="gapFilter"]');

// 🔥 USE FILTERED MATCHUPS IF BLITZ IS ENABLED
let source = lastGeneratedMatchups;

if(blitzEnabled){

  source = lastGeneratedMatchups.filter(m => {

    const small =
      m.redTeam.length < m.blueTeam.length ? m.redTeam : m.blueTeam;

    const large =
      m.redTeam.length > m.blueTeam.length ? m.redTeam : m.blueTeam;

    if(m.redTeam.length === m.blueTeam.length) return false;

    const smallSkill = small.reduce((s,p)=>s+p.skill,0);
    const largeSkill = large.reduce((s,p)=>s+p.skill,0);

    return smallSkill > largeSkill;

  });

}

const counts = {
  all: source.length,
  0: 0,
  1: 0,
  2: 0,
  3: 0,
  4: 0
};

source.forEach(m=>{
  if(counts.hasOwnProperty(m.skillGap)){
    counts[m.skillGap]++;
  }
});

  radios.forEach(radio=>{

    const value = radio.value;

    const label = radio.parentElement;

if(value === "all"){

  label.childNodes[1].nodeValue = ` All options [${counts.all}]`;

  const isDisabled = counts.all === 0;

  radio.disabled = isDisabled;

  if(isDisabled){
    label.classList.add("disabled");
  }else{
    label.classList.remove("disabled");
  }

}else{

  label.childNodes[1].nodeValue = ` Difference ${value} [${counts[value]}]`;

  const isDisabled = counts[value] === 0;

  radio.disabled = isDisabled;

  if(isDisabled){
    label.classList.add("disabled");
  }else{
    label.classList.remove("disabled");
  }

}

  });

}

function resetGeneratedMatchups(){

  // Clear UI
  document.getElementById("generatedMatchups").innerHTML = "";

  // Reset stored data
  lastGeneratedMatchups = [];
  selectedMatchKey = null;

  // Disable radio buttons again
  document.querySelectorAll('input[name="gapFilter"]').forEach(r=>{
    r.disabled = true;
    r.checked = r.value === "all"; // reset to default
    r.parentElement.classList.add("disabled");
  });

/* 🔥 RESET + HIDE BLITZ */

const blitzToggle = document.getElementById("blitzToggle");
const blitzContainer = document.querySelector(".blitzToggle");

if(blitzToggle && blitzContainer){

  blitzToggle.checked = false;
  blitzEnabled = false;

  blitzContainer.classList.remove("show");

  setTimeout(()=>{
    blitzContainer.style.display = "none";
  },300);

}

}

function normalizeSessionData(data = {}){
  return {
    elimination: Array.isArray(data.elimination) ? data.elimination.filter(Boolean) : [],
    blitz: Array.isArray(data.blitz) ? data.blitz.filter(Boolean) : [],
    ctf: Array.isArray(data.ctf) ? data.ctf.filter(Boolean) : []
  };
}

function getActiveSessionMaps(){
  return customSessionActive
    ? normalizeSessionData(customSessionData)
    : normalizeSessionData(currentSessionMaps);
}

function updateCustomSessionButtons(){

  const buildBtn = document.getElementById("buildCustomSessionBtn");
  const saveBtn = document.getElementById("saveCustomSessionBtn");
  const clearBtn = document.getElementById("clearCustomSessionBtn");
  const generateBtn = document.getElementById("generateSessionMapsBtn");
  const progressBtn = document.getElementById("saveSessionProgressBtn");
  const badge = document.getElementById("customSessionBadge");

  if(!buildBtn || !saveBtn || !clearBtn || !generateBtn || !progressBtn) return;

  if(customSessionActive){
    buildBtn.style.display = "none";
    saveBtn.style.display = "inline-flex";
    clearBtn.style.display = "inline-flex";
    generateBtn.style.display = "none";
    progressBtn.style.display = "none";
    if(badge) badge.classList.add("active");
  }else{
    buildBtn.style.display = "inline-flex";
    saveBtn.style.display = "none";
    clearBtn.style.display = "none";
    generateBtn.style.display = "inline-flex";
    progressBtn.style.display = "inline-flex";
    if(badge) badge.classList.remove("active");
  }

}

function renderAllSessionViews(){
  const activeSession = getActiveSessionMaps();

  renderSessionMaps(activeSession);
  renderUpcomingSessionCard(activeSession);
  updateCustomMapHighlights();

  setTimeout(()=>{
    handleSessionHighlightUpdate();
  }, 50);
}

function getCustomSessionLimit(mode){
  if(mode === "elimination") return 2;
  if(mode === "blitz") return 2;
  if(mode === "ctf") return 5;
  return 0;
}

function isMapSelectedInCustomSession(mode, mapName){
  const list = customSessionData[mode] || [];
  return list.includes(mapName);
}

async function toggleCustomSessionMap(mode, mapName){

  if(!customSessionActive){
    return;
  }

  const current = [...(customSessionData[mode] || [])];
  const existingIndex = current.indexOf(mapName);

  if(existingIndex !== -1){
    current.splice(existingIndex, 1);
  }else{
    const limit = getCustomSessionLimit(mode);

    if(current.length >= limit){
      const modeLabel =
        mode === "ctf" ? "CTF" :
        mode.charAt(0).toUpperCase() + mode.slice(1);

      await showModal(`${modeLabel} already has ${limit} maps selected.`, "alert");
      return;
    }

    current.push(mapName);
  }

  customSessionData = {
    ...customSessionData,
    [mode]: current
  };

  renderAllSessionViews();

}

function updateCustomMapHighlights(){

  document.querySelectorAll(".mapMasterRow").forEach(row => {
    const mode = row.getAttribute("data-mode");
    const mapName = row.getAttribute("data-map-name");

    if(!mode || !mapName) return;

    row.classList.toggle("customMapSelectable", customSessionActive);
    row.classList.toggle(
      "customMapSelected",
      customSessionActive && isMapSelectedInCustomSession(mode, mapName)
    );

    row.classList.remove("customSelectedElimination", "customSelectedBlitz", "customSelectedCtf");

    if(customSessionActive && isMapSelectedInCustomSession(mode, mapName)){
      if(mode === "elimination") row.classList.add("customSelectedElimination");
      if(mode === "blitz") row.classList.add("customSelectedBlitz");
      if(mode === "ctf") row.classList.add("customSelectedCtf");
    }
  });

}

async function loadCustomSessionState(){

  const res = await api({
    action:"getCustomSession"
  });

  if(!res || !res.ok){
    console.log("Failed loading custom session");
    customSessionActive = false;
    customSessionData = normalizeSessionData();
    return;
  }

  customSessionActive = false;
  customSessionData = normalizeSessionData(res.session);

}

// 🔥 LOAD CURRENT SESSION MAPS
async function loadSessionMaps(){

const overlay = document.getElementById("mapListLoadingOverlay");

if(overlay){
  overlay.style.display = "flex";
}

  const sessionData = await api({
    action:"getSessionMaps"
  });

  if(!sessionData.ok){
    console.log("Failed loading session maps");
    return;
  }

  await loadCustomSessionState();
  currentSessionMaps = normalizeSessionData(sessionData);

  const initialData = await api({
  action:"getInitialData"
});

if(initialData.ok && initialData.mapList){
  renderMasterMapList(initialData.mapList);
}

updateCustomSessionButtons();
renderAllSessionViews();

// 🔥 HIDE LOADER
if(overlay){
  overlay.style.display = "none";
}

}

// 🔥 RENDER SESSION MAPS
function renderSessionMaps(data){

  // Keep legacy hidden lists updated for existing highlight logic
  renderModeSessionList("eliminationSessionList", data.elimination || [], "elimination");
  renderModeSessionList("blitzSessionList", data.blitz || [], "blitz");
  renderModeSessionList("ctfSessionList", data.ctf || [], "ctf");

  // Render new visible single-card layout
  renderUnifiedSessionMaps(data);

}

function renderUpcomingSessionCard(data){

  buildCopySessionCard(data, globalMapMatchMaker, {
    makerId: "upcomingSessionMaker",
    bodyId: "upcomingSessionBody"
  });

  renderMatchMakerRotationStrip(globalMapMatchMaker);

}

function getShortMatchMakerName(name){

  const clean = (name || "").trim();

  if(!clean) return "--";

  return clean.substring(0, 2).toUpperCase();

}

function renderMatchMakerRotationStrip(currentMaker){

  const makerEl = document.getElementById("upcomingSessionMaker");

  if(!makerEl) return;

  const players = Array.isArray(allPlayers)
    ? allPlayers
      .filter(player => player.matchMaker !== false)
      .slice()
      .sort((a,b)=>a.name.localeCompare(b.name))
    : [];

  makerEl.innerHTML = "";

  if(players.length === 0){
    makerEl.textContent = "Match Maker: Not selected";
    return;
  }

  const label = document.createElement("span");
  label.className = "matchMakerInlineLabel";
  label.textContent = "Match Maker:";

  const strip = document.createElement("div");
  strip.className = "matchMakerRotationStrip";

  makerEl.appendChild(label);
  makerEl.appendChild(strip);

  const currentIndex = players.findIndex(player => player.name === currentMaker);

  players.forEach((player, index) => {

    const item = document.createElement("div");
    item.className = "matchMakerRotationItem";

    if(index === currentIndex){
      item.classList.add("current");
    }

    item.innerHTML = `
      <span class="rotationFullName">${player.name}</span>
      <span class="rotationShortName">${getShortMatchMakerName(player.name)}</span>
    `;

    strip.appendChild(item);

  });

}

// 🔥 RENDER SESSION MAPS
function renderUnifiedSessionMaps(data){

  const container = document.getElementById("sessionMapsUnifiedRows");

  if(!container) return;

  container.innerHTML = "";

  const sections = [
    {
      label: "Elimination",
      mode: "elimination",
      headerClass: "eliminationHeader",
      maps: data.elimination || []
    },
    {
      label: "Blitz",
      mode: "blitz",
      headerClass: "blitzHeader",
      maps: data.blitz || []
    },
    {
      label: "CTF",
      mode: "ctf",
      headerClass: "ctfHeader",
      maps: data.ctf || []
    }
  ];

  sections.forEach((section, sectionIndex) => {

    const header = document.createElement("div");
    header.className = `sessionUnifiedHeader ${section.headerClass}`;

    if(sectionIndex === 0){
      header.classList.add("firstHeader");
    }

    header.textContent = section.label;
    container.appendChild(header);

    section.maps.forEach((mapName, index) => {

      if(!mapName) return;

     const row = document.createElement("div");
     row.className = "mapMasterRow sessionUnifiedRow";
      
     const masterContainer = document.getElementById(section.mode + "MasterList");

let masterIndex = "";

if(masterContainer){
  const masterRows = Array.from(masterContainer.querySelectorAll(".mapMasterRow"));

  const foundIndex = masterRows.findIndex(r => r.innerText.trim() === mapName);

  if(foundIndex !== -1){
    masterIndex = foundIndex + 1;
  }
}

row.setAttribute("data-index", masterIndex);

     row.innerHTML = `
       <span class="sessionUnifiedName">${mapName}</span>
       <button class="mapDeleteMini">✕</button>
     `;

row.querySelector(".mapDeleteMini").onclick = async () => {

  if(!isAdminUnlocked()){
  showModal("Unlock admin mode first.", "alert");
  return;
}

  const pass = await getAdminPassword();
  if(!pass) return;

  const res = await api({
    action:"deleteSessionMap",
    mode: section.mode,
    slot: index + 1,
    password: pass
  });

  if(!res.ok){
    showModal(res.error || "Delete failed", "alert");
    return;
  }

  // 🔥 ADD THIS
  sessionStorage.setItem("adminPass", pass);
  updateAdminBar();

  renderSessionMaps(res);

  setTimeout(()=>{
    handleSessionHighlightUpdate();
  }, 50);

};
      
      container.appendChild(row);

    });

  });

}

// 🔥 RENDER ONE MODE

function renderModeSessionList(containerId, maps, mode){

  const container = document.getElementById(containerId);

  if(!container) return;

  container.innerHTML = "";

  // 🔥 Create ONE compact card
  const card = document.createElement("div");
  card.className = "mapSessionCompactCard";

  maps.forEach((mapName, index) => {

    if(!mapName) return;

    const row = document.createElement("div");
    row.className = "mapSessionCompactRow";

    const masterContainer = document.getElementById(mode + "MasterList");

let masterIndex = "";

if(masterContainer){
  const masterRows = Array.from(masterContainer.querySelectorAll(".mapMasterRow"));
  const foundIndex = masterRows.findIndex(row => row.innerText.trim() === mapName);
  
  if(foundIndex !== -1){
    masterIndex = foundIndex + 1;
  }
}
    
    row.innerHTML = `
      <span class="mapSessionName" data-index="${masterIndex}">
      ${mapName}
      </span>
      <button class="mapDeleteMini">✕</button>
    `;

row.querySelector(".mapDeleteMini").onclick = async () => {

  if(!isAdminUnlocked()){
  showModal("Unlock admin mode first.", "alert");
  return;
}

  const pass = await getAdminPassword();
  if(!pass) return;

  const res = await api({
    action:"deleteSessionMap",
    mode: mode,
    slot: index + 1,
    password: pass
  });

  if(!res.ok){
    showModal(res.error || "Delete failed", "alert");
    return;
  }

  // 🔥 ADD THIS
  sessionStorage.setItem("adminPass", pass);
  updateAdminBar();

  renderSessionMaps(res);

  setTimeout(()=>{
    handleSessionHighlightUpdate();
  }, 50);

};
    
    card.appendChild(row);

  });

  // If empty
  if(card.children.length === 0){
    card.innerHTML = `<div class="mapSessionEmpty">—</div>`;
  }

  container.appendChild(card);

}

// 🔥 RENDER FULL MASTER MAP LIST
function renderMasterMapList(mapList){

  renderMasterModeList("eliminationMasterList", mapList.elimination || [], "elimination");
  renderMasterModeList("blitzMasterList", mapList.blitz || [], "blitz");
  renderMasterModeList("ctfMasterList", mapList.ctf || [], "ctf");

}

// 🔥 RENDER ONE MASTER MODE COLUMN
function renderMasterModeList(containerId, maps, mode){

  const container = document.getElementById(containerId);

  if(!container) return;

  container.innerHTML = "";

  maps.forEach((mapName, index) => {

    const row = document.createElement("div");
    row.className = "mapMasterRow";
    row.textContent = mapName;
    row.setAttribute("data-index", index + 1);
    row.setAttribute("data-mode", mode);
    row.setAttribute("data-map-name", mapName);

    row.onclick = async () => {
      await toggleCustomSessionMap(mode, mapName);
    };

    container.appendChild(row);

  });

}

// 🔥 BUTTON ACTIONS
function setupMapListButtons(){

  const generateBtn = document.getElementById("generateSessionMapsBtn");
  const saveBtn = document.getElementById("saveSessionProgressBtn");
  const clearSessionBtn = document.getElementById("clearSessionMapsBtn");
  const buildCustomBtn = document.getElementById("buildCustomSessionBtn");
  const saveCustomBtn = document.getElementById("saveCustomSessionBtn");
  const clearCustomBtn = document.getElementById("clearCustomSessionBtn");
  const copyBtn = document.getElementById("copySessionMapsBtn");

if(generateBtn){
  generateBtn.onclick = async () => {

    if(!isAdminUnlocked()){
  showModal("Unlock admin mode first.", "alert");
  return;
}

const pass = await getAdminPassword();
if(!pass) return;

    const res = await api({
      action:"generateSessionMaps",
      password: pass
    });

    if(!res.ok){
      showModal(res.error || "Generate failed", "alert");
      return;
    }

    sessionStorage.setItem("adminPass", pass);
    updateAdminBar();

renderSessionMaps(res);

/* 🔥 RE-RUN HIGHLIGHT AFTER GENERATE */
setTimeout(()=>{
  handleSessionHighlightUpdate();
}, 50);

  };
}

if(saveBtn){
saveBtn.onclick = async () => {

  if(!isAdminUnlocked()){
  showModal("Unlock admin mode first.", "alert");
  return;
}

  const pass = await getAdminPassword();
  if(!pass) return;

  const res = await api({
    action:"saveSessionProgress",
    password: pass
  });

  if(!res.ok){
    showModal(res.error || "Save failed", "alert");
    return;
  }

  // 🔥 ADD THIS
  sessionStorage.setItem("adminPass", pass);
  updateAdminBar();

  showModal("Session progress saved", "alert");

  handleSessionHighlightUpdate();

};
}

if(clearSessionBtn){
  clearSessionBtn.onclick = async () => {

    if(!isAdminUnlocked()){
      showModal("Unlock admin mode first.", "alert");
      return;
    }

    if(customSessionActive){

      customSessionData = normalizeSessionData({
        elimination: [],
        blitz: [],
        ctf: []
      });

      renderAllSessionViews();
      showModal("Custom session maps cleared", "alert");
      return;

    }

    const pass = await getAdminPassword();
    if(!pass) return;

    const res = await api({
      action:"clearSessionMaps",
      password: pass
    });

    if(!res || !res.ok){
      showModal(res.error || "Clear session maps failed", "alert");
      return;
    }

    sessionStorage.setItem("adminPass", pass);
    updateAdminBar();

    currentSessionMaps = normalizeSessionData(res);
    renderAllSessionViews();

    showModal("Session maps cleared", "alert");

  };
}

if(buildCustomBtn){
  buildCustomBtn.onclick = async () => {

    if(!isAdminUnlocked()){
      showModal("Unlock admin mode first.", "alert");
      return;
    }

    customSessionActive = true;
    customSessionData = normalizeSessionData(currentSessionMaps);

    updateCustomSessionButtons();
    renderAllSessionViews();

    showModal("Custom session mode enabled", "alert");

  };
}

if(saveCustomBtn){
  saveCustomBtn.onclick = async () => {

    if(!isAdminUnlocked()){
      showModal("Unlock admin mode first.", "alert");
      return;
    }

    const pass = await getAdminPassword();
    if(!pass) return;

    const res = await api({
      action:"saveCustomSession",
      password: pass,
      session: customSessionData
    });

    if(!res || !res.ok){
      showModal(res.error || "Save custom session failed", "alert");
      return;
    }

    sessionStorage.setItem("adminPass", pass);
    updateAdminBar();

    customSessionActive = !!res.active;
    customSessionData = normalizeSessionData(res.session);

    updateCustomSessionButtons();
    renderAllSessionViews();

    showModal("Custom session saved", "alert");

  };
}

if(clearCustomBtn){
  clearCustomBtn.onclick = async () => {

    if(!isAdminUnlocked()){
      showModal("Unlock admin mode first.", "alert");
      return;
    }

    const pass = await getAdminPassword();
    if(!pass) return;

    const res = await api({
      action:"clearCustomSession",
      password: pass
    });

    if(!res || !res.ok){
      showModal(res.error || "Clear custom session failed", "alert");
      return;
    }

    sessionStorage.setItem("adminPass", pass);
    updateAdminBar();

    customSessionActive = false;
    customSessionData = normalizeSessionData(res.session);

    updateCustomSessionButtons();
    renderAllSessionViews();

    showModal("Exited custom session", "alert");

  };
}
  
if(copyBtn){
  copyBtn.onclick = async () => {

    const matchMaker = globalMapMatchMaker;

    const sessionData = getActiveSessionMaps();

    const copyCard = document.getElementById("copySessionCard");

    if(!copyCard){
      showModal("Copy card not found", "alert");
      return;
    }

    buildCopySessionCard(sessionData, matchMaker);

    copyCard.classList.add("show");

    try{

      await new Promise(resolve => requestAnimationFrame(() => resolve()));

      const canvas = await html2canvas(copyCard, {
        backgroundColor: null,
        scale: 4,
        useCORS: true
      });

      const blob = await new Promise(resolve => canvas.toBlob(resolve, "image/png"));

      if(!blob){
        throw new Error("Canvas export failed");
      }

      await navigator.clipboard.write([
        new ClipboardItem({ "image/png": blob })
      ]);

      showModal("Session maps copied as image ✅", "alert");

    }catch(err){

      console.error(err);
      showModal("Copy image failed. Your browser may not support it.", "alert");

    }finally{

      copyCard.classList.remove("show");

    }

  };
}


}

function buildCopySessionCard(data, matchMaker, options = {}){

  const makerId = options.makerId || "copySessionMaker";
  const bodyId = options.bodyId || "copySessionBody";

  const makerEl = document.getElementById(makerId);
  const bodyEl = document.getElementById(bodyId);

  if(!makerEl || !bodyEl) return;

  makerEl.textContent = "Match Maker: " + (matchMaker || "Not selected");

  const sections = [
    {
      label: "Elimination",
      className: "eliminationHeader",
      maps: (data.elimination || []).filter(Boolean)
    },
    {
      label: "Blitz",
      className: "blitzHeader",
      maps: (data.blitz || []).filter(Boolean)
    },
    {
      label: "CTF",
      className: "ctfHeader",
      maps: (data.ctf || []).filter(Boolean)
    }
  ];

  bodyEl.innerHTML = sections.map(section => {
    const rows = section.maps.length
      ? section.maps.map(mapName => `
        <div class="copySessionRow">
          <span class="copySessionName">${mapName}</span>
        </div>
      `).join("")
      : `
        <div class="copySessionEmpty">-</div>
      `;

    return `
      <div class="copySessionSection">
        <div class="copySessionSectionHeader ${section.className}">${section.label}</div>
        <div class="copySessionRows">${rows}</div>
      </div>
    `;
  }).join("");

}

function handleSessionHighlightUpdate(){

  processMode("elimination", "eliminationSessionList", "eliminationMasterList");
  processMode("blitz", "blitzSessionList", "blitzMasterList");
  processMode("ctf", "ctfSessionList", "ctfMasterList");

}

function processMode(mode, sessionId, masterId){

  const sessionContainer = document.getElementById(sessionId);
  const masterContainer = document.getElementById(masterId);

  if(!sessionContainer || !masterContainer) return;

  // 🔥 get session maps
const sessionMaps = Array.from(
  sessionContainer.querySelectorAll(".mapSessionName")
).map(el => el.innerText.trim());

/* 🔥 ALWAYS CLEAR OLD HIGHLIGHTS FIRST */
masterContainer.querySelectorAll(".mapMasterRow").forEach(row=>{
  row.classList.remove("lastPlayedMap");
});

/* 🔥 IF NO SESSION MAPS → STOP HERE (NO HIGHLIGHT) */
if(sessionMaps.length === 0) return;

  const firstMap = sessionMaps[0];

  // 🔥 get full map list
  const masterMaps = Array.from(
    masterContainer.querySelectorAll(".mapMasterRow")
  ).map(el => el.innerText.trim());

  const index = masterMaps.indexOf(firstMap);

  if(index === -1){
    return;
  }

  // 🔥 get previous (with wrap-around)
  const prevIndex = (index - 1 + masterMaps.length) % masterMaps.length;

  const prevMap = masterMaps[prevIndex];

  // 🔥 REMOVE OLD HIGHLIGHTS
masterContainer.querySelectorAll(".mapMasterRow").forEach(row=>{
  row.classList.remove("lastPlayedMap");
});

// 🔥 APPLY NEW HIGHLIGHT
const rows = masterContainer.querySelectorAll(".mapMasterRow");

rows.forEach(row => {
  if(row.innerText.trim() === prevMap){
    row.classList.add("lastPlayedMap");
  }
});

}
