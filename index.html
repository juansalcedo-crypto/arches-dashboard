<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width, initial-scale=1.0"/>
<title>Arches Dashboard — PM View</title>
<link href="https://fonts.googleapis.com/css2?family=DM+Mono:wght@400;500;600&family=Syne:wght@700;800;900&display=swap" rel="stylesheet"/>
<script src="https://cdn.jsdelivr.net/npm/@supabase/supabase-js@2"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
<style>
:root{--bg:#0D0F1A;--sur:#111320;--bdr:#1E2240;--org:#F7A200;--od:#E07000;--grn:#16A34A;--red:#DC2626;--blu:#3B82F6;}
*{box-sizing:border-box;margin:0;padding:0;}
body{background:var(--bg);color:#fff;font-family:'DM Mono',monospace;min-height:100vh;}
::-webkit-scrollbar{width:4px}::-webkit-scrollbar-track{background:var(--bg)}::-webkit-scrollbar-thumb{background:var(--bdr);border-radius:2px}
@keyframes pulse{0%,100%{opacity:1}50%{opacity:0.5}}
@keyframes fadeIn{from{opacity:0;transform:translateY(8px)}to{opacity:1;transform:translateY(0)}}
@keyframes live{0%,100%{box-shadow:0 0 0 0 rgba(22,163,74,0)}50%{box-shadow:0 0 0 4px rgba(22,163,74,0.3)}}

/* LOGIN */
#login{position:fixed;inset:0;background:var(--bg);display:flex;align-items:center;justify-content:center;z-index:200}
.lbox{background:var(--sur);border:1px solid var(--bdr);border-radius:20px;padding:40px;max-width:420px;width:90%;box-shadow:0 20px 60px rgba(0,0,0,0.6)}

/* HEADER */
#hdr{background:var(--sur);border-bottom:1px solid var(--bdr);padding:0 20px;height:56px;display:flex;align-items:center;justify-content:space-between;position:sticky;top:0;z-index:50;box-shadow:0 2px 20px rgba(0,0,0,0.4)}

/* TABS */
.tab-bar{display:flex;gap:4px;background:rgba(255,255,255,0.04);border-radius:10px;padding:4px}
.tab{padding:6px 14px;border-radius:7px;cursor:pointer;font-size:11px;font-weight:600;color:rgba(255,255,255,0.4);transition:all 0.2s;border:none;background:transparent;font-family:'DM Mono',monospace}
.tab.active{background:var(--org);color:#fff;box-shadow:0 2px 8px rgba(247,162,0,0.4)}
.tab:hover:not(.active){color:rgba(255,255,255,0.7);background:rgba(255,255,255,0.06)}

/* CONTENT */
#content{padding:20px;max-width:1400px;margin:0 auto}
.page{display:none;animation:fadeIn 0.3s ease}
.page.active{display:block}

/* CARDS */
.card{background:var(--sur);border:1px solid var(--bdr);border-radius:12px;overflow:hidden;margin-bottom:14px}
.card-head{padding:12px 16px;border-bottom:1px solid var(--bdr);display:flex;align-items:center;justify-content:space-between;background:linear-gradient(135deg,rgba(247,162,0,0.06),transparent)}
.card-title{font-family:'Syne',sans-serif;font-size:12px;font-weight:800;color:#fff;letter-spacing:0.3px}

/* PERSON ROW */
.person-row{display:grid;grid-template-columns:180px 80px 1fr 120px 100px;align-items:center;gap:12px;padding:10px 16px;border-bottom:1px solid rgba(255,255,255,0.04);transition:background 0.15s}
.person-row:hover{background:rgba(247,162,0,0.04)}
.person-row:last-child{border-bottom:none}
.pname{font-size:11px;font-weight:600;color:#fff}
.pscore{font-family:'Syne',sans-serif;font-size:14px;font-weight:900}
.pbar-wrap{height:6px;background:rgba(255,255,255,0.06);border-radius:3px;overflow:hidden}
.pbar-fill{height:100%;border-radius:3px;transition:width 0.5s}
.pteam{font-size:9px;padding:2px 7px;border-radius:10px;font-weight:700}
.live-dot{width:7px;height:7px;border-radius:50%;background:var(--grn);animation:live 2s infinite;display:inline-block;margin-right:4px}

/* STAT BOXES */
.stats-grid{display:grid;grid-template-columns:repeat(4,1fr);gap:12px;margin-bottom:16px}
.stat-card{background:var(--sur);border:1px solid var(--bdr);border-radius:12px;padding:16px;text-align:center}
.stat-val{font-family:'Syne',sans-serif;font-size:28px;font-weight:900;color:var(--org);line-height:1}
.stat-lbl{font-size:10px;color:rgba(255,255,255,0.4);margin-top:4px;letter-spacing:0.8px;text-transform:uppercase}

/* ACTION TABLE */
.action-table{width:100%;border-collapse:collapse}
.action-table th{font-size:9px;color:rgba(255,255,255,0.35);letter-spacing:1px;text-transform:uppercase;padding:8px 12px;text-align:left;background:rgba(255,255,255,0.02);border-bottom:1px solid var(--bdr)}
.action-table td{padding:8px 12px;border-bottom:1px solid rgba(255,255,255,0.04);font-size:11px}
.action-table tr:last-child td{border-bottom:none}
.action-table tr:hover td{background:rgba(247,162,0,0.03)}

/* BUTTONS */
.btn{padding:7px 14px;border-radius:8px;border:none;cursor:pointer;font-family:'DM Mono',monospace;font-size:11px;font-weight:700;transition:all 0.15s;display:inline-flex;align-items:center;gap:5px}
.btn-org{background:linear-gradient(135deg,var(--org),var(--od));color:#fff;box-shadow:0 3px 10px rgba(247,162,0,0.4)}
.btn-org:hover{transform:translateY(-1px)}
.btn-ghost{background:rgba(255,255,255,0.06);border:1px solid rgba(255,255,255,0.1);color:rgba(255,255,255,0.6)}
.btn-ghost:hover{background:rgba(255,255,255,0.1);color:#fff}
.btn-red{background:rgba(220,38,38,0.15);border:1px solid rgba(220,38,38,0.3);color:#FF6B6B}
.btn-grn{background:rgba(22,163,74,0.15);border:1px solid rgba(22,163,74,0.3);color:#4ADE80}
.btn-sm{padding:4px 10px;font-size:10px;border-radius:6px}

/* FORM INPUTS */
.finp{background:rgba(255,255,255,0.05);border:1px solid rgba(255,255,255,0.1);border-radius:8px;padding:8px 12px;font-family:'DM Mono',monospace;font-size:12px;color:#fff;outline:none;width:100%;transition:border-color 0.2s}
.finp:focus{border-color:var(--org);box-shadow:0 0 0 3px rgba(247,162,0,0.1)}
.finp::placeholder{color:rgba(255,255,255,0.2)}
.fsel{background:rgba(255,255,255,0.05);border:1px solid rgba(255,255,255,0.1);border-radius:8px;padding:8px 12px;font-family:'DM Mono',monospace;font-size:12px;color:#fff;outline:none;cursor:pointer}
.fsel option{background:#111320}
.flbl{font-size:10px;color:rgba(255,255,255,0.4);letter-spacing:1px;text-transform:uppercase;margin-bottom:6px;display:block}

/* MODAL */
.modal-overlay{position:fixed;inset:0;background:rgba(0,0,0,0.7);z-index:300;display:none;align-items:center;justify-content:center}
.modal-box{background:var(--sur);border:1px solid var(--bdr);border-radius:16px;padding:28px;max-width:480px;width:90%;box-shadow:0 20px 60px rgba(0,0,0,0.8);max-height:80vh;overflow-y:auto}
.modal-title{font-family:'Syne',sans-serif;font-size:15px;font-weight:900;color:var(--org);margin-bottom:16px}

/* PONDERADO */
.pond-grid{display:grid;grid-template-columns:repeat(3,1fr);gap:10px;margin-bottom:16px}
.pond-card{background:rgba(255,255,255,0.03);border:1px solid rgba(255,255,255,0.06);border-radius:10px;padding:12px;text-align:center}
.pond-val{font-family:'Syne',sans-serif;font-size:22px;font-weight:900;line-height:1}
.pond-lbl{font-size:9px;color:rgba(255,255,255,0.35);margin-top:3px;letter-spacing:0.5px}

/* DATE FILTER */
.date-filter{display:flex;gap:8px;align-items:center;margin-bottom:16px;flex-wrap:wrap}

/* TOAST */
#toast{position:fixed;bottom:20px;left:50%;transform:translateX(-50%) translateY(20px);border-radius:10px;padding:10px 20px;font-size:12px;font-weight:600;opacity:0;transition:all 0.3s;pointer-events:none;z-index:999;white-space:nowrap}
#toast.show{opacity:1;transform:translateX(-50%) translateY(0)}
</style>
</head>
<body>

<!-- LOGIN -->
<div id="login">
  <div class="lbox">
    <div style="display:flex;align-items:center;gap:12px;margin-bottom:28px">
      <svg width="36" height="24" viewBox="0 0 200 130" fill="none">
        <path d="M100 10 C60 10 30 50 20 90 L50 90 C55 65 75 45 100 45 C125 45 145 65 150 90 L180 90 C170 50 140 10 100 10Z" fill="#F7A200"/>
        <path d="M20 90 C15 70 5 55 0 90 L25 90Z" fill="#E07000"/>
        <path d="M180 90 C185 70 195 55 200 90 L175 90Z" fill="#E07000"/>
        <path d="M50 90 C48 75 42 62 35 55 L15 90Z" fill="#F7A200" opacity="0.7"/>
        <path d="M150 90 C152 75 158 62 165 55 L185 90Z" fill="#F7A200" opacity="0.7"/>
        <path d="M85 45 C85 25 95 15 100 10 C105 15 115 25 115 45Z" fill="#E07000"/>
      </svg>
      <div>
        <div style="font-family:'Syne',sans-serif;font-size:18px;font-weight:900;color:#fff">ARCHES DASHBOARD</div>
        <div style="font-size:9px;color:var(--org);letter-spacing:2px;margin-top:2px">MANAGER VIEW</div>
      </div>
    </div>
    <label class="flbl">Your Name</label>
    <input class="finp" id="login-name" placeholder="e.g. Luis" style="margin-bottom:12px"/>
    <label class="flbl">Password</label>
    <input class="finp" id="login-pass" type="password" placeholder="••••••••" style="margin-bottom:20px"/>
    <button class="btn btn-org" style="width:100%;justify-content:center;padding:12px" onclick="doLogin()">🚀 Enter Dashboard</button>
    <div id="login-err" style="color:#FF6B6B;font-size:11px;margin-top:10px;text-align:center;display:none"></div>
  </div>
</div>

<!-- HEADER -->
<div id="hdr" style="display:none">
  <div style="display:flex;align-items:center;gap:12px">
    <svg width="26" height="17" viewBox="0 0 200 130" fill="none">
      <path d="M100 10 C60 10 30 50 20 90 L50 90 C55 65 75 45 100 45 C125 45 145 65 150 90 L180 90 C170 50 140 10 100 10Z" fill="#F7A200"/>
      <path d="M20 90 C15 70 5 55 0 90 L25 90Z" fill="#E07000"/>
      <path d="M180 90 C185 70 195 55 200 90 L175 90Z" fill="#E07000"/>
      <path d="M50 90 C48 75 42 62 35 55 L15 90Z" fill="#F7A200" opacity="0.7"/>
      <path d="M150 90 C152 75 158 62 165 55 L185 90Z" fill="#F7A200" opacity="0.7"/>
      <path d="M85 45 C85 25 95 15 100 10 C105 15 115 25 115 45Z" fill="#E07000"/>
    </svg>
    <div>
      <div style="font-family:'Syne',sans-serif;font-size:13px;font-weight:900">ARCHES DASHBOARD</div>
      <div style="font-size:8px;color:var(--org);letter-spacing:2px" id="hdr-date"></div>
    </div>
  </div>
  <div class="tab-bar">
    <button class="tab active" onclick="showPage('overview')">📊 Overview</button>
    <button class="tab" onclick="showPage('performance')">📈 Performance</button>
    <button class="tab" onclick="showPage('history')">🗓 History</button>
    <button class="tab" onclick="showPage('projects')">📋 Projects</button>
    <button class="tab" onclick="showPage('settings')">⚙️ Settings</button>
  </div>
  <div style="display:flex;align-items:center;gap:8px">
    <div style="display:flex;align-items:center;gap:5px;font-size:10px;color:rgba(255,255,255,0.4)">
      <span class="live-dot"></span>LIVE
    </div>
    <span id="hdr-manager" style="font-size:11px;color:rgba(255,255,255,0.5)"></span>
    <button class="btn btn-org" onclick="exportExcel()">⬇ Export Excel</button>
    <button class="btn btn-ghost btn-sm" onclick="doLogout()">↩ Logout</button>
  </div>
</div>

<!-- CONTENT -->
<div id="content" style="display:none">

  <!-- OVERVIEW -->
  <div id="page-overview" class="page active">
    <div class="stats-grid" id="overview-stats"></div>
    <div style="display:grid;grid-template-columns:1fr 1fr;gap:14px">
      <div>
        <div class="card">
          <div class="card-head">
            <span class="card-title">🟠 Team Luis</span>
            <span id="luis-pct" style="font-family:'Syne',sans-serif;font-size:13px;font-weight:900;color:var(--org)">—</span>
          </div>
          <div id="luis-members"></div>
        </div>
      </div>
      <div>
        <div class="card">
          <div class="card-head">
            <span class="card-title">🔵 Team Taiki</span>
            <span id="taiki-pct" style="font-family:'Syne',sans-serif;font-size:13px;font-weight:900;color:#3B82F6">—</span>
          </div>
          <div id="taiki-members"></div>
        </div>
      </div>
    </div>
    <!-- Recent activity -->
    <div class="card" style="margin-top:14px">
      <div class="card-head"><span class="card-title">⚡ Recent Activity</span><button class="btn btn-ghost btn-sm" onclick="loadOverview()">↺ Refresh</button></div>
      <div id="recent-activity" style="max-height:300px;overflow-y:auto"></div>
    </div>
  </div>

  <!-- PONDERADOS -->
  <div id="page-performance" class="page">
    <div class="date-filter">
      <select class="fsel" id="perf-period" onchange="loadPerformance()" style="width:140px">
        <option value="today">Today</option>
        <option value="week">This Week</option>
        <option value="month">This Month</option>
      </select>
      <select class="fsel" id="perf-team" onchange="loadPerformance()" style="width:130px">
        <option value="all">All Teams</option>
        <option value="luis">Team Luis</option>
        <option value="taiki">Team Taiki</option>
      </select>
    </div>
    <div id="performance-content"></div>
  </div>

  <!-- HISTORICO -->
  <div id="page-history" class="page">
    <div class="date-filter">
      <input type="date" class="finp" id="hist-from" style="width:150px"/>
      <span style="color:rgba(255,255,255,0.3);font-size:11px">to</span>
      <input type="date" class="finp" id="hist-to" style="width:150px"/>
      <select class="fsel" id="hist-person" style="width:160px"><option value="">All People</option></select>
      <button class="btn btn-org btn-sm" onclick="loadHistory()">🔍 Search</button>
    </div>
    <div class="card">
      <div class="card-head"><span class="card-title">📅 Daily Log</span><span id="hist-count" style="font-size:10px;color:rgba(255,255,255,0.4)"></span></div>
      <div style="overflow-x:auto"><table class="action-table" id="hist-table"><thead><tr><th>Date</th><th>Person</th><th>Team</th><th>Project</th><th>P</th><th>DB</th><th>LI</th><th>Exp</th><th>Calls</th><th>Conn</th><th>InMail</th><th>Emails</th><th>SMS</th><th>Co.Call</th><th>Prop</th><th>Int.S</th><th>Int.D</th><th>Other</th><th>Done</th><th>%</th><th>Actions</th></tr></thead><tbody id="hist-body"></tbody></table></div>
    </div>
  </div>

  <!-- PROYECTOS -->
  <div id="page-projects" class="page">
    <div style="display:grid;grid-template-columns:1fr 1fr;gap:14px">
      <!-- Assign project -->
      <div class="card">
        <div class="card-head"><span class="card-title">📋 Assign Project</span></div>
        <div style="padding:16px;display:flex;flex-direction:column;gap:10px">
          <div><label class="flbl">Person</label><select class="fsel" id="ap-person" style="width:100%"><option value="">Select person…</option></select></div>
          <div><label class="flbl">Project Name</label><input class="finp" id="ap-project" placeholder="e.g. advancy_us_specialty_plant"/></div>
          <div><label class="flbl">Priority</label>
            <select class="fsel" id="ap-prio" style="width:100%">
              <option value="P1">P1 — Highest</option>
              <option value="P2" selected>P2 — Medium</option>
              <option value="P3">P3 — Normal</option>
            </select>
          </div>
          <div><label class="flbl">Date</label><input type="date" class="finp" id="ap-date"/></div>
          <button class="btn btn-org" onclick="assignProject()" style="align-self:flex-start">✅ Assign</button>
        </div>
      </div>
      <!-- Assigned today -->
      <div class="card">
        <div class="card-head"><span class="card-title">📌 Assigned Today</span><button class="btn btn-ghost btn-sm" onclick="loadAssigned()">↺</button></div>
        <div id="assigned-list" style="max-height:400px;overflow-y:auto"></div>
      </div>
    </div>
    <!-- Comments -->
    <div class="card" style="margin-top:14px">
      <div class="card-head"><span class="card-title">💬 Comments & Feedback</span></div>
      <div style="padding:16px;display:grid;grid-template-columns:1fr 1fr 2fr auto;gap:10px;align-items:end;border-bottom:1px solid var(--bdr)">
        <div><label class="flbl">Person</label><select class="fsel" id="cm-person" style="width:100%"><option value="">Select…</option></select></div>
        <div><label class="flbl">Date</label><input type="date" class="finp" id="cm-date"/></div>
        <div><label class="flbl">Comment</label><input class="finp" id="cm-text" placeholder="e.g. Great performance today!"/></div>
        <button class="btn btn-org" onclick="addComment()">Send</button>
      </div>
      <div id="comments-list" style="max-height:300px;overflow-y:auto"></div>
    </div>
  </div>

  <!-- SETTINGS -->
  <div id="page-settings" class="page">
    <div style="display:grid;grid-template-columns:1fr 1fr;gap:14px">
      <!-- Action targets -->
      <div class="card">
        <div class="card-head"><span class="card-title">🎯 Action Targets</span><button class="btn btn-org btn-sm" onclick="saveTargets()">💾 Save</button></div>
        <div id="targets-form" style="padding:12px;display:flex;flex-direction:column;gap:6px"></div>
      </div>
      <!-- Team management -->
      <div>
        <div class="card" style="margin-bottom:14px">
          <div class="card-head"><span class="card-title">👥 Team Management</span></div>
          <div style="padding:16px;display:flex;flex-direction:column;gap:10px">
            <div style="display:flex;gap:8px">
              <input class="finp" id="new-member-name" placeholder="New member name…"/>
              <select class="fsel" id="new-member-team"><option value="luis">Team Luis</option><option value="taiki">Team Taiki</option></select>
              <button class="btn btn-grn btn-sm" onclick="addMember()" style="white-space:nowrap">+ Add</button>
            </div>
            <div id="members-list" style="max-height:300px;overflow-y:auto"></div>
          </div>
        </div>
        <!-- Managers -->
        <div class="card">
          <div class="card-head"><span class="card-title">🔐 Dashboard Access</span></div>
          <div style="padding:16px">
            <div style="font-size:10px;color:rgba(255,255,255,0.4);margin-bottom:10px">Managers with access:</div>
            <div id="managers-list"></div>
            <div style="display:flex;gap:8px;margin-top:10px">
              <input class="finp" id="new-mgr-name" placeholder="Manager name…"/>
              <input class="finp" id="new-mgr-pass" placeholder="Password…" type="password"/>
              <button class="btn btn-grn btn-sm" onclick="addManager()" style="white-space:nowrap">+ Add</button>
            </div>
          </div>
        </div>
      </div>
    </div>
  </div>

</div>

<!-- MODALS -->
<div class="modal-overlay" id="detail-modal">
  <div class="modal-box">
    <div class="modal-title" id="detail-title">Person Detail</div>
    <div id="detail-content"></div>
    <button class="btn btn-ghost" style="margin-top:16px" onclick="closeModal('detail-modal')">Close</button>
  </div>
</div>

<div id="toast"></div>

<script>
const SUPA_URL = "https://watgvezsgppjvusotuge.supabase.co";
const SUPA_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6IndhdGd2ZXpzZ3BwanZ1c290dWdlIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODA2ODExMzgsImV4cCI6MjA5NjI1NzEzOH0.U1vPzCzCPKHozbdvIV6n2_67UqB3etUSFc9yKoyE0SU";
const { createClient } = supabase;
const db = createClient(SUPA_URL, SUPA_KEY);

// Default managers (stored in localStorage for simplicity)
const DEFAULT_MANAGERS = [
  {name:"Luis", pass:"Arches2026"},
  {name:"Taiki", pass:"Arches2026"},
  {name:"Admin", pass:"Arches2026"}
];

let currentManager = null;
let allMembers = [];
let actionTargets = [];
let realtimeSub = null;

// ── AUTH ──────────────────────────────────────────────────────────────────────
function getManagers(){
  const stored = localStorage.getItem("arches_managers");
  return stored ? JSON.parse(stored) : DEFAULT_MANAGERS;
}

function doLogin(){
  const name = document.getElementById("login-name").value.trim();
  const pass = document.getElementById("login-pass").value;
  const managers = getManagers();
  const mgr = managers.find(m=>m.name.toLowerCase()===name.toLowerCase()&&m.pass===pass);
  if(!mgr){ document.getElementById("login-err").style.display="block"; document.getElementById("login-err").textContent="❌ Invalid name or password"; return; }
  currentManager = mgr.name;
  localStorage.setItem("arches_session", JSON.stringify({name:mgr.name, ts:Date.now()}));
  document.getElementById("login").style.display="none";
  document.getElementById("hdr").style.display="flex";
  document.getElementById("content").style.display="block";
  document.getElementById("hdr-manager").textContent=mgr.name;
  document.getElementById("hdr-date").textContent=new Date().toLocaleDateString("en-US",{weekday:"short",month:"short",day:"numeric"}).toUpperCase();
  init();
}

function doLogout(){
  localStorage.removeItem("arches_session");
  currentManager=null;
  document.getElementById("login").style.display="flex";
  document.getElementById("hdr").style.display="none";
  document.getElementById("content").style.display="none";
  if(realtimeSub) db.removeChannel(realtimeSub);
}

function checkSession(){
  const s = localStorage.getItem("arches_session");
  if(!s) return;
  const sess = JSON.parse(s);
  if(Date.now()-sess.ts > 8*60*60*1000){ localStorage.removeItem("arches_session"); return; }
  currentManager = sess.name;
  document.getElementById("login").style.display="none";
  document.getElementById("hdr").style.display="flex";
  document.getElementById("content").style.display="block";
  document.getElementById("hdr-manager").textContent=sess.name;
  document.getElementById("hdr-date").textContent=new Date().toLocaleDateString("en-US",{weekday:"short",month:"short",day:"numeric"}).toUpperCase();
  init();
}

// ── INIT ──────────────────────────────────────────────────────────────────────
async function init(){
  await Promise.all([loadMembers(), loadTargets()]);
  populatePersonSelects();
  loadOverview();
  loadAssigned();
  loadComments();
  loadTargetsForm();
  loadMembersList();
  loadManagersList();
  setupRealtime();
  // Set default dates
  const today = new Date().toISOString().split("T")[0];
  document.getElementById("ap-date").value = today;
  document.getElementById("cm-date").value = today;
  document.getElementById("hist-from").value = today;
  document.getElementById("hist-to").value = today;
  // Populate hist person
  const sel = document.getElementById("hist-person");
  allMembers.forEach(m=>{ const o=document.createElement("option"); o.value=m.name; o.textContent=m.name; sel.appendChild(o); });
}

async function loadMembers(){
  const {data} = await db.from("members").select("*").eq("active",true).order("name");
  allMembers = data||[];
}

async function loadTargets(){
  const {data} = await db.from("action_targets").select("*").order("sort_order");
  actionTargets = data||[];
}

function populatePersonSelects(){
  ["ap-person","cm-person"].forEach(id=>{
    const sel=document.getElementById(id);
    sel.innerHTML="<option value=''>Select person…</option>";
    allMembers.forEach(m=>{ const o=document.createElement("option"); o.value=m.name; o.textContent=`${m.name} (${m.team_id})`; sel.appendChild(o); });
  });
}

// ── REALTIME ──────────────────────────────────────────────────────────────────
function setupRealtime(){
  realtimeSub = db.channel("dashboard-live")
    .on("postgres_changes",{event:"INSERT",schema:"public",table:"daily_logs"},()=>{ loadOverview(); showToast("📊 New data received!","info"); })
    .on("postgres_changes",{event:"INSERT",schema:"public",table:"comments"},()=>loadComments())
    .subscribe();
}

// ── OVERVIEW ──────────────────────────────────────────────────────────────────
async function loadOverview(){
  const today = new Date().toISOString().split("T")[0];
  const {data:logs} = await db.from("daily_logs").select("*").eq("date",today);
  const rows = logs||[];

  // Stats
  const people = [...new Set(rows.map(r=>r.person))];
  const totalActions = rows.reduce((s,r)=>s+r.actions_done,0);
  const avgPct = people.length>0 ? Math.round(rows.reduce((s,r)=>s+r.pct,0)/rows.length) : 0;
  const totalEmails = rows.reduce((s,r)=>s+r.emails,0);

  document.getElementById("overview-stats").innerHTML=`
    <div class="stat-card"><div class="stat-val">${people.length}</div><div class="stat-lbl">Active Today</div></div>
    <div class="stat-card"><div class="stat-val">${avgPct}%</div><div class="stat-lbl">Avg Complete</div></div>
    <div class="stat-card"><div class="stat-val">${totalActions}</div><div class="stat-lbl">Total Actions</div></div>
    <div class="stat-card"><div class="stat-val">${totalEmails}</div><div class="stat-lbl">Emails Sent</div></div>`;

  // Per team
  renderTeamMembers("luis", rows, "#F7A200", "luis-members", "luis-pct");
  renderTeamMembers("taiki", rows, "#3B82F6", "taiki-members", "taiki-pct");

  // Recent
  const recent = [...rows].sort((a,b)=>new Date(b.created_at)-new Date(a.created_at)).slice(0,20);
  const ra = document.getElementById("recent-activity");
  if(recent.length===0){ ra.innerHTML=`<div style="padding:16px;font-size:11px;color:rgba(255,255,255,0.2);text-align:center">No data yet today — waiting for team EODs</div>`; return; }
  ra.innerHTML = recent.map(r=>`
    <div style="padding:9px 16px;border-bottom:1px solid rgba(255,255,255,0.04);display:grid;grid-template-columns:140px 100px 1fr 60px 80px;align-items:center;gap:10px;font-size:10px">
      <span style="font-weight:600;color:#fff">${r.person}</span>
      <span style="color:rgba(255,255,255,0.4)">${r.project.length>14?r.project.substring(0,12)+"…":r.project}</span>
      <div style="height:4px;background:rgba(255,255,255,0.06);border-radius:2px;overflow:hidden"><div style="height:100%;width:${r.pct}%;background:${r.pct>=80?"#16A34A":r.pct>=50?"#EAB308":"#F7A200"};border-radius:2px"></div></div>
      <span style="font-family:'Syne',sans-serif;font-weight:900;color:${r.pct>=80?"#4ADE80":r.pct>=50?"#EAB308":"#F7A200"}">${r.pct}%</span>
      <button class="btn btn-ghost btn-sm" onclick="showPersonDetail('${r.person}')">Detail</button>
    </div>`).join("");
}

function renderTeamMembers(teamId, rows, color, elId, pctId){
  const teamMembers = allMembers.filter(m=>m.team_id===teamId);
  const el = document.getElementById(elId);
  if(teamMembers.length===0){ el.innerHTML=`<div style="padding:14px;font-size:10px;color:rgba(255,255,255,0.2)">No members</div>`; return; }

  const memberRows = teamMembers.map(m=>{
    const mLogs = rows.filter(r=>r.person===m.name);
    const avgPct = mLogs.length>0 ? Math.round(mLogs.reduce((s,r)=>s+r.pct,0)/mLogs.length) : 0;
    const totalOut = mLogs.reduce((s,r)=>s+r.emails+r.cold_call+r.connections+r.inmail,0);
    const hasData = mLogs.length>0;
    return {name:m.name, pct:avgPct, total:totalOut, hasData, projects:mLogs.length};
  });

  const teamAvg = memberRows.filter(m=>m.hasData).length>0
    ? Math.round(memberRows.filter(m=>m.hasData).reduce((s,m)=>s+m.pct,0)/memberRows.filter(m=>m.hasData).length) : 0;
  document.getElementById(pctId).textContent = teamAvg+"%";

  el.innerHTML = memberRows.map(m=>`
    <div class="person-row" onclick="showPersonDetail('${m.name}')" style="cursor:pointer">
      <div>
        <div class="pname">${m.name}</div>
        <div style="font-size:9px;color:rgba(255,255,255,0.3);margin-top:1px">${m.hasData?m.projects+" project(s)":"No data today"}</div>
      </div>
      <div class="pscore" style="color:${m.pct>=80?"#4ADE80":m.pct>=50?"#EAB308":m.hasData?"#FF6B6B":"rgba(255,255,255,0.2)"}">${m.hasData?m.pct+"%":"—"}</div>
      <div class="pbar-wrap"><div class="pbar-fill" style="width:${m.pct}%;background:${m.pct>=80?"#16A34A":m.pct>=50?"#EAB308":color}"></div></div>
      <div style="font-size:10px;color:rgba(255,255,255,0.4)">${m.hasData?m.total+" outreach":"—"}</div>
      <div style="display:flex;gap:4px">
        ${m.hasData?`<span style="width:8px;height:8px;border-radius:50%;background:#4ADE80;box-shadow:0 0 6px #4ADE80;display:inline-block"></span>`:`<span style="width:8px;height:8px;border-radius:50%;background:rgba(255,255,255,0.1);display:inline-block"></span>`}
      </div>
    </div>`).join("");
}

// ── PERSON DETAIL ─────────────────────────────────────────────────────────────
async function showPersonDetail(person){
  const today = new Date().toISOString().split("T")[0];
  const {data:logs} = await db.from("daily_logs").select("*").eq("date",today).eq("person",person);
  const {data:comms} = await db.from("comments").select("*").eq("person",person).order("created_at",{ascending:false}).limit(5);
  const rows = logs||[];

  const totalOut = rows.reduce((s,r)=>s+r.emails+r.cold_call+r.connections+r.inmail,0);
  const avgPct = rows.length>0 ? Math.round(rows.reduce((s,r)=>s+r.pct,0)/rows.length) : 0;

  document.getElementById("detail-title").textContent=`📊 ${person} — Today`;
  document.getElementById("detail-content").innerHTML=`
    <div class="pond-grid" style="grid-template-columns:repeat(3,1fr)">
      <div class="pond-card"><div class="pond-val" style="color:var(--org)">${avgPct}%</div><div class="pond-lbl">Avg Complete</div></div>
      <div class="pond-card"><div class="pond-val" style="color:#4ADE80">${rows.length}</div><div class="pond-lbl">Projects</div></div>
      <div class="pond-card"><div class="pond-val" style="color:#3B82F6">${totalOut}</div><div class="pond-lbl">Total Outreach</div></div>
    </div>
    ${rows.map(r=>`
      <div style="background:rgba(255,255,255,0.03);border:1px solid rgba(255,255,255,0.06);border-radius:8px;padding:10px;margin-bottom:8px">
        <div style="font-family:'Syne',sans-serif;font-size:11px;font-weight:700;color:#fff;margin-bottom:6px">[${r.priority}] ${r.project} — ${r.pct}%</div>
        <div style="display:grid;grid-template-columns:repeat(4,1fr);gap:4px">
          ${actionTargets.map(a=>`<div style="font-size:9px;color:rgba(255,255,255,0.4)">${a.emoji} ${r[a.id]||0}</div>`).join("")}
        </div>
      </div>`).join("")}
    ${comms&&comms.length>0?`<div style="margin-top:12px"><div style="font-size:10px;color:rgba(255,255,255,0.4);margin-bottom:6px">RECENT COMMENTS</div>${comms.map(c=>`<div style="background:rgba(247,162,0,0.08);border:1px solid rgba(247,162,0,0.2);border-radius:7px;padding:8px;margin-bottom:5px;font-size:10px"><span style="color:var(--org);font-weight:600">${c.manager}:</span> ${c.comment} <span style="color:rgba(255,255,255,0.25);float:right">${new Date(c.created_at).toLocaleDateString()}</span></div>`).join("")}</div>`:""}
    <div style="margin-top:12px;display:flex;gap:8px">
      <input class="finp" id="quick-comment" placeholder="Add comment…" style="flex:1"/>
      <button class="btn btn-org btn-sm" onclick="quickComment('${person}')">Send</button>
    </div>`;
  document.getElementById("detail-modal").style.display="flex";
}

async function quickComment(person){
  const text = document.getElementById("quick-comment").value.trim();
  if(!text) return;
  const today = new Date().toISOString().split("T")[0];
  await db.from("comments").insert({date:today, person, manager:currentManager, comment:text});
  showToast("💬 Comment added!","success");
  document.getElementById("quick-comment").value="";
  showPersonDetail(person);
}

// ── PONDERADOS ────────────────────────────────────────────────────────────────
async function loadPerformance(){
  const period = document.getElementById("perf-period").value;
  const team = document.getElementById("perf-team").value;
  const today = new Date();
  let from = new Date();
  if(period==="week") from.setDate(today.getDate()-7);
  else if(period==="month") from.setDate(today.getDate()-30);

  let q = db.from("daily_logs").select("*").gte("date",from.toISOString().split("T")[0]).lte("date",today.toISOString().split("T")[0]);
  if(team!=="all") q = q.eq("team_id",team);
  const {data:logs} = await q;
  const rows = logs||[];

  // Group by person
  const byPerson = {};
  rows.forEach(r=>{
    if(!byPerson[r.person]) byPerson[r.person]={name:r.person,team:r.team_id,logs:[],totals:{}};
    byPerson[r.person].logs.push(r);
    actionTargets.forEach(a=>{ byPerson[r.person].totals[a.id]=(byPerson[r.person].totals[a.id]||0)+(r[a.id]||0); });
  });

  const people = Object.values(byPerson).sort((a,b)=>{
    const aAvg=a.logs.reduce((s,r)=>s+r.pct,0)/a.logs.length;
    const bAvg=b.logs.reduce((s,r)=>s+r.pct,0)/b.logs.length;
    return bAvg-aAvg;
  });

  const el = document.getElementById("performance-content");
  if(people.length===0){ el.innerHTML=`<div style="padding:30px;text-align:center;color:rgba(255,255,255,0.3);font-size:12px">No data for selected period</div>`; return; }

  el.innerHTML = `
    <div class="card">
      <div class="card-head"><span class="card-title">📈 Performance — ${period==="today"?"Today":period==="week"?"Last 7 Days":"Last 30 Days"}</span></div>
      <table class="action-table">
        <thead><tr><th>#</th><th>Person</th><th>Team</th><th>Days</th><th>Avg %</th><th>Score</th>${actionTargets.map(a=>`<th>${a.emoji}</th>`).join("")}<th>Actions</th></tr></thead>
        <tbody>
          ${people.map((p,i)=>{
            const avgPct=Math.round(p.logs.reduce((s,r)=>s+r.pct,0)/p.logs.length);
            const days=[...new Set(p.logs.map(r=>r.date))].length;
            const color=avgPct>=80?"#4ADE80":avgPct>=50?"#EAB308":"#FF6B6B";
            const medal=i===0?"🥇":i===1?"🥈":i===2?"🥉":"";
            return `<tr>
              <td style="color:rgba(255,255,255,0.3)">${medal||i+1}</td>
              <td style="font-weight:600;cursor:pointer;color:var(--org)" onclick="showPersonDetail('${p.name}')">${p.name}</td>
              <td><span style="font-size:9px;padding:2px 6px;border-radius:8px;background:rgba(255,255,255,0.06)">${p.team}</span></td>
              <td style="color:rgba(255,255,255,0.5)">${days}</td>
              <td style="font-family:'Syne',sans-serif;font-weight:900;color:${color}">${avgPct}%</td>
              <td><div style="width:80px;height:4px;background:rgba(255,255,255,0.08);border-radius:2px;overflow:hidden"><div style="height:100%;width:${avgPct}%;background:${color};border-radius:2px"></div></div></td>
              ${actionTargets.map(a=>`<td style="color:rgba(255,255,255,0.6)">${p.totals[a.id]||0}</td>`).join("")}
              <td><button class="btn btn-ghost btn-sm" onclick="showPersonDetail('${p.name}')">Detail</button></td>
            </tr>`;
          }).join("")}
        </tbody>
      </table>
    </div>`;
}

// ── HISTORICO ─────────────────────────────────────────────────────────────────
async function loadHistory(){
  const from = document.getElementById("hist-from").value;
  const to = document.getElementById("hist-to").value;
  const person = document.getElementById("hist-person").value;

  let q = db.from("daily_logs").select("*").gte("date",from).lte("date",to).order("date",{ascending:false});
  if(person) q = q.eq("person",person);
  const {data:logs} = await q.limit(500);
  const rows = logs||[];

  document.getElementById("hist-count").textContent=`${rows.length} records`;
  const tbody = document.getElementById("hist-body");
  tbody.innerHTML = rows.map(r=>`
    <tr>
      <td>${r.date}</td>
      <td style="font-weight:600;color:#fff">${r.person}</td>
      <td><span style="font-size:9px;padding:2px 6px;border-radius:8px;background:rgba(255,255,255,0.06)">${r.team_id||"—"}</span></td>
      <td style="font-size:10px;color:rgba(255,255,255,0.6)">${r.project}</td>
      <td><span style="font-size:9px;padding:2px 5px;border-radius:5px;background:rgba(247,162,0,0.15);color:var(--org)">${r.priority}</span></td>
      <td>${r.db_search}</td><td>${r.li_search}</td><td>${r.expansion}</td><td>${r.cold_call}</td>
      <td>${r.connections}</td><td>${r.inmail}</td><td>${r.emails}</td><td>${r.sms}</td>
      <td>${r.company_call}</td><td>${r.proposals}</td><td>${r.int_sched}</td><td>${r.int_done}</td><td>${r.other}</td>
      <td>${r.actions_done}</td>
      <td style="font-family:'Syne',sans-serif;font-weight:900;color:${r.pct>=80?"#4ADE80":r.pct>=50?"#EAB308":"#FF6B6B"}">${r.pct}%</td>
      <td><button class="btn btn-red btn-sm" onclick="deleteLog(${r.id})">🗑</button></td>
    </tr>`).join("");
}

async function deleteLog(id){
  if(!confirm("Delete this record?")) return;
  await db.from("daily_logs").delete().eq("id",id);
  showToast("🗑 Record deleted","info");
  loadHistory();
}

// ── PROJECTS ──────────────────────────────────────────────────────────────────
async function assignProject(){
  const person = document.getElementById("ap-person").value;
  const project = document.getElementById("ap-project").value.trim();
  const prio = document.getElementById("ap-prio").value;
  const date = document.getElementById("ap-date").value;
  if(!person||!project||!date){ showToast("⚠️ Fill all fields","error"); return; }
  await db.from("assigned_projects").insert({person,project,priority:prio,date,assigned_by:currentManager});
  showToast(`✅ Project assigned to ${person}!`,"success");
  document.getElementById("ap-project").value="";
  loadAssigned();
}

async function loadAssigned(){
  const today = new Date().toISOString().split("T")[0];
  const {data} = await db.from("assigned_projects").select("*").eq("date",today).order("created_at",{ascending:false});
  const rows = data||[];
  const el = document.getElementById("assigned-list");
  if(rows.length===0){ el.innerHTML=`<div style="padding:14px;font-size:11px;color:rgba(255,255,255,0.2)">No assignments today</div>`; return; }
  el.innerHTML = rows.map(r=>`
    <div style="padding:9px 14px;border-bottom:1px solid rgba(255,255,255,0.04);display:flex;align-items:center;justify-content:space-between;gap:8px">
      <div>
        <div style="font-size:11px;font-weight:600;color:#fff">${r.person}</div>
        <div style="font-size:10px;color:rgba(255,255,255,0.4);margin-top:1px">${r.project}</div>
      </div>
      <div style="display:flex;align-items:center;gap:6px">
        <span style="font-size:9px;padding:2px 6px;border-radius:5px;background:rgba(247,162,0,0.15);color:var(--org)">${r.priority}</span>
        <span style="font-size:9px;color:rgba(255,255,255,0.3)">by ${r.assigned_by}</span>
        <button class="btn btn-red btn-sm" onclick="deleteAssigned(${r.id})">🗑</button>
      </div>
    </div>`).join("");
}

async function deleteAssigned(id){ await db.from("assigned_projects").delete().eq("id",id); loadAssigned(); }

async function addComment(){
  const person = document.getElementById("cm-person").value;
  const date = document.getElementById("cm-date").value;
  const text = document.getElementById("cm-text").value.trim();
  if(!person||!text){ showToast("⚠️ Fill all fields","error"); return; }
  await db.from("comments").insert({date,person,manager:currentManager,comment:text});
  showToast("💬 Comment added!","success");
  document.getElementById("cm-text").value="";
  loadComments();
}

async function loadComments(){
  const {data} = await db.from("comments").select("*").order("created_at",{ascending:false}).limit(50);
  const rows = data||[];
  const el = document.getElementById("comments-list");
  if(rows.length===0){ el.innerHTML=`<div style="padding:14px;font-size:11px;color:rgba(255,255,255,0.2)">No comments yet</div>`; return; }
  el.innerHTML = rows.map(r=>`
    <div style="padding:10px 16px;border-bottom:1px solid rgba(255,255,255,0.04);display:flex;align-items:flex-start;gap:10px">
      <div style="flex:1">
        <div style="font-size:11px"><span style="color:var(--org);font-weight:700">${r.manager}</span> → <span style="color:#fff;font-weight:600">${r.person}</span> <span style="color:rgba(255,255,255,0.3);font-size:9px">${r.date}</span></div>
        <div style="font-size:11px;color:rgba(255,255,255,0.6);margin-top:3px">${r.comment}</div>
      </div>
      <button class="btn btn-red btn-sm" onclick="deleteComment(${r.id})">🗑</button>
    </div>`).join("");
}

async function deleteComment(id){ await db.from("comments").delete().eq("id",id); loadComments(); }

// ── SETTINGS ──────────────────────────────────────────────────────────────────
function loadTargetsForm(){
  const el = document.getElementById("targets-form");
  el.innerHTML = actionTargets.map(a=>`
    <div style="display:flex;align-items:center;gap:8px">
      <span style="font-size:13px;width:20px">${a.emoji}</span>
      <span style="font-size:11px;color:rgba(255,255,255,0.6);flex:1">${a.name}</span>
      <input type="number" min="0" value="${a.target}" id="tgt-${a.id}" style="width:60px;background:rgba(255,255,255,0.06);border:1px solid rgba(255,255,255,0.1);border-radius:6px;padding:4px 8px;font-family:'DM Mono',monospace;font-size:12px;color:var(--org);text-align:center;outline:none"/>
    </div>`).join("");
}

async function saveTargets(){
  for(const a of actionTargets){
    const val = parseInt(document.getElementById(`tgt-${a.id}`)?.value)||a.target;
    await db.from("action_targets").update({target:val}).eq("id",a.id);
  }
  await loadTargets();
  showToast("✅ Targets saved!","success");
}

async function loadMembersList(){
  const el = document.getElementById("members-list");
  const {data} = await db.from("members").select("*").order("team_id").order("name");
  const members = data||[];
  el.innerHTML = members.map(m=>`
    <div style="display:flex;align-items:center;gap:8px;padding:6px 0;border-bottom:1px solid rgba(255,255,255,0.04)">
      <span style="font-size:11px;flex:1;color:#fff">${m.name}</span>
      <select onchange="moveMember(${m.id},this.value)" style="background:rgba(255,255,255,0.05);border:1px solid rgba(255,255,255,0.1);border-radius:5px;padding:3px 6px;font-family:'DM Mono',monospace;font-size:10px;color:#fff">
        <option value="luis" ${m.team_id==="luis"?"selected":""}>🟠 Luis</option>
        <option value="taiki" ${m.team_id==="taiki"?"selected":""}>🔵 Taiki</option>
      </select>
      <button class="btn btn-red btn-sm" onclick="removeMember(${m.id})">🗑</button>
    </div>`).join("");
}

async function addMember(){
  const name = document.getElementById("new-member-name").value.trim();
  const team = document.getElementById("new-member-team").value;
  if(!name){ showToast("⚠️ Enter name","error"); return; }
  await db.from("members").insert({name,team_id:team});
  showToast(`✅ ${name} added to Team ${team}!`,"success");
  document.getElementById("new-member-name").value="";
  await loadMembers();
  populatePersonSelects();
  loadMembersList();
}

async function moveMember(id, teamId){
  await db.from("members").update({team_id:teamId}).eq("id",id);
  showToast("✅ Member moved!","success");
  await loadMembers();
  populatePersonSelects();
}

async function removeMember(id){
  if(!confirm("Remove this member?")) return;
  await db.from("members").update({active:false}).eq("id",id);
  showToast("Member removed","info");
  await loadMembers();
  populatePersonSelects();
  loadMembersList();
}

function loadManagersList(){
  const managers = getManagers();
  const el = document.getElementById("managers-list");
  el.innerHTML = managers.map((m,i)=>`
    <div style="display:flex;align-items:center;gap:8px;padding:5px 0;border-bottom:1px solid rgba(255,255,255,0.04)">
      <span style="font-size:11px;flex:1;color:#fff">${m.name}</span>
      <span style="font-size:9px;color:rgba(255,255,255,0.3)">••••••••</span>
      ${i>0?`<button class="btn btn-red btn-sm" onclick="removeManager(${i})">🗑</button>`:"<span style='font-size:9px;color:rgba(255,255,255,0.2)'>owner</span>"}
    </div>`).join("");
}

function addManager(){
  const name = document.getElementById("new-mgr-name").value.trim();
  const pass = document.getElementById("new-mgr-pass").value;
  if(!name||!pass){ showToast("⚠️ Fill both fields","error"); return; }
  const managers = getManagers();
  managers.push({name,pass});
  localStorage.setItem("arches_managers",JSON.stringify(managers));
  showToast(`✅ ${name} added as manager!`,"success");
  document.getElementById("new-mgr-name").value="";
  document.getElementById("new-mgr-pass").value="";
  loadManagersList();
}

function removeManager(i){
  const managers = getManagers();
  managers.splice(i,1);
  localStorage.setItem("arches_managers",JSON.stringify(managers));
  loadManagersList();
}

// ── EXPORT EXCEL ──────────────────────────────────────────────────────────────
async function exportExcel(){
  const from = document.getElementById("hist-from")?.value || new Date().toISOString().split("T")[0];
  const to = document.getElementById("hist-to")?.value || new Date().toISOString().split("T")[0];
  const {data:logs} = await db.from("daily_logs").select("*").gte("date",from).lte("date",to).order("date",{ascending:false});
  const rows = logs||[];
  if(rows.length===0){ showToast("⚠️ No data to export","error"); return; }

  const headers = ["Date","Person","Team","Project","Priority","DB Search","LI Search","Expansion","Cold Call","Connections","InMail","Emails","SMS","Co.Call","Proposals","Int.Sched","Int.Done","Other","Actions Done","% Complete"];
  const data = rows.map(r=>[r.date,r.person,r.team_id||"",r.project,r.priority,r.db_search,r.li_search,r.expansion,r.cold_call,r.connections,r.inmail,r.emails,r.sms,r.company_call,r.proposals,r.int_sched,r.int_done,r.other,r.actions_done,r.pct+"%"]);

  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet([headers,...data]);
  // Style headers
  ws["!cols"] = headers.map((_,i)=>({wch:i<4?20:10}));
  XLSX.utils.book_append_sheet(wb,"Daily Log",ws);

  // Performance sheet
  const byPerson = {};
  rows.forEach(r=>{
    if(!byPerson[r.person]) byPerson[r.person]={name:r.person,team:r.team_id,days:new Set(),totals:{},pcts:[]};
    byPerson[r.person].days.add(r.date);
    byPerson[r.person].pcts.push(r.pct);
    ["db_search","li_search","expansion","cold_call","connections","inmail","emails","sms","company_call","proposals","int_sched","int_done","other"].forEach(k=>{
      byPerson[r.person].totals[k]=(byPerson[r.person].totals[k]||0)+(r[k]||0);
    });
  });
  const pHeaders = ["Person","Team","Days Active","Avg %","DB Search","LI Search","Expansion","Cold Call","Connections","InMail","Emails","SMS","Co.Call","Proposals","Int.Sched","Int.Done","Other"];
  const pData = Object.values(byPerson).map(p=>[p.name,p.team,p.days.size,Math.round(p.pcts.reduce((s,x)=>s+x,0)/p.pcts.length)+"%",...["db_search","li_search","expansion","cold_call","connections","inmail","emails","sms","company_call","proposals","int_sched","int_done","other"].map(k=>p.totals[k]||0)]);
  const ws2 = XLSX.utils.aoa_to_sheet([pHeaders,...pData]);
  XLSX.utils.book_append_sheet(wb,"Performance",ws2);

  XLSX.writeFile(wb,`arches-report-${from}-to-${to}.xlsx`);
  showToast("⬇ Excel exported!","success");
}

// ── UI HELPERS ────────────────────────────────────────────────────────────────
function showPage(page){
  document.querySelectorAll(".page").forEach(p=>p.classList.remove("active"));
  document.querySelectorAll(".tab").forEach(t=>t.classList.remove("active"));
  document.getElementById(`page-${page}`).classList.add("active");
  event.target.classList.add("active");
  if(page==="performance") loadPerformance();
  if(page==="history") loadHistory();
  if(page==="settings"){ loadTargetsForm(); loadMembersList(); loadManagersList(); }
}

function closeModal(id){ document.getElementById(id).style.display="none"; }

function showToast(msg,type="info"){
  const t=document.getElementById("toast");
  t.style.background=type==="success"?"linear-gradient(135deg,#16A34A,#15803D)":type==="error"?"linear-gradient(135deg,#DC2626,#B91C1C)":"linear-gradient(135deg,#1E2240,#111320)";
  t.style.color="#fff"; t.style.boxShadow="0 6px 24px rgba(0,0,0,0.5)";
  t.textContent=msg; t.classList.add("show");
  setTimeout(()=>t.classList.remove("show"),4000);
}

// ── START ─────────────────────────────────────────────────────────────────────
checkSession();
</script>
</body>
</html>
