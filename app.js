/* PortCo Learning Hub – dark & gold edition.
   Reads data.xlsx (same folder) and redraws the dashboard. Update the Excel file,
   commit it to GitHub, and the dashboard refreshes automatically. */

const CONFIG = {
  dataFile: "data.xlsx",
  autoRefreshSeconds: 60,
  dueSoonDays: 7,
  statusOrder: ["Script Received", "In Development", "In Review", "Portco Feedback", "Delivered", "On Hold"]
};
/* Colour sets for each theme (charts can't read CSS variables, so they live here too) */
const THEMES = {
  dark:  { gold: "#E3C88B", gold2: "#B8964A", gold3: "#7A6332", red: "#E26D5F", amber: "#E7A94F", green: "#7CC49A", grey: "#5E5E5E",
           track: "#262626", muted: "#8C8C8C", line: "#2B2B2B", card: "#1B1B1B", tip: "#111111",
           status: ["#F3E2B8", "#E3C88B", "#C9A660", "#9C7E43", "#6E5A30", "#4A4A4A"] },
  light: { gold: "#2E5A8A", gold2: "#1F3A5F", gold3: "#7FA7D9", red: "#D64545", amber: "#E8A33D", green: "#2E9E6B", grey: "#9AA3B2",
           track: "#EEF2F8", muted: "#6B7385", line: "#E3E7EF", card: "#FFFFFF", tip: "#1F3A5F",
           status: ["#9AA3B2", "#3B7DD8", "#E8A33D", "#8E6CD8", "#2E9E6B", "#C9CED8"] }
};
let theme = document.documentElement.getAttribute("data-theme") === "light" ? "light" : "dark";
let C = THEMES[theme];
const HEALTH_COLOR = { gold: C.gold, green: C.green, amber: C.amber, red: C.red, grey: C.grey };
const HEALTH_LABEL = { gold: "Delivered", green: "On track", amber: "At risk", red: "Overdue", grey: "On hold" };

let DATA = { portcos: [], projects: [], comms: [], files: [] };
let charts = {}, lastSignature = "", usingLocalFile = false;

const $ = (id) => document.getElementById(id);
const today = () => { const d = new Date(); d.setHours(0, 0, 0, 0); return d; };

/* ---------- Theme ---------- */
function applyChartDefaults() {
  if (!window.Chart) return;
  Chart.defaults.color = C.muted;
  Chart.defaults.borderColor = C.line;
  Chart.defaults.font.family = "Manrope, Segoe UI, sans-serif";
  Chart.defaults.plugins.tooltip.backgroundColor = C.tip;
  Chart.defaults.plugins.tooltip.borderColor = C.gold3;
  Chart.defaults.plugins.tooltip.borderWidth = 1;
}
function setTheme(t) {
  theme = t; C = THEMES[t];
  document.documentElement.setAttribute("data-theme", t);
  try { localStorage.setItem("plh-theme", t); } catch (e) {}
  document.querySelectorAll("#themeToggle span").forEach(s => s.classList.toggle("on", s.dataset.t === t));
  applyChartDefaults();
  const s = $("syncStatus").textContent;
  if (s && !s.startsWith("Loading")) setSync(s.replace(/^[●⚠] /, ""), !s.startsWith("⚠"));
  if (DATA.projects.length) render();
}

/* ---------- Load Excel ---------- */
async function loadFromRepo(force = false) {
  if (usingLocalFile && !force) return;
  try {
    const res = await fetch(`${CONFIG.dataFile}?v=${Date.now()}`, { cache: "no-store" });
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    const buf = await res.arrayBuffer();
    const sig = await signature(buf);
    if (sig === lastSignature && !force) { setSync("Up to date", true); return; }
    lastSignature = sig; usingLocalFile = false;
    parseWorkbook(buf, CONFIG.dataFile);
  } catch (err) {
    console.warn(err);
    setSync(location.protocol === "file:" ? "Use Load Excel to preview" : "Can't read data.xlsx", false);
  }
}
async function signature(buf) {
  if (window.crypto && crypto.subtle) {
    const h = await crypto.subtle.digest("SHA-1", buf);
    return Array.from(new Uint8Array(h)).map(b => b.toString(16).padStart(2, "0")).join("");
  }
  return String(buf.byteLength);
}
function loadFromFile(file) {
  const r = new FileReader();
  r.onload = (e) => { usingLocalFile = true; parseWorkbook(e.target.result, file.name); };
  r.readAsArrayBuffer(file);
}
function parseWorkbook(buf, name) {
  const wb = XLSX.read(buf, { type: "array", cellDates: true });
  const sheet = (n) => {
    const key = wb.SheetNames.find(s => s.toLowerCase() === n.toLowerCase());
    return key ? XLSX.utils.sheet_to_json(wb.Sheets[key], { defval: "" }) : [];
  };
  DATA.portcos = sheet("Portcos").filter(r => r.PortcoName);
  DATA.projects = sheet("Projects").filter(r => r.ProjectID || r.ProjectName).map(r => ({
    ...r, ScriptReceived: toDate(r.ScriptReceived), StartDate: toDate(r.StartDate), DueDate: toDate(r.DueDate),
    Progress: Math.max(0, Math.min(100, Number(r.Progress) || 0)), FeedbackRounds: Number(r.FeedbackRounds) || 0
  }));
  DATA.comms = sheet("Communications").filter(r => r.Date || r.Subject).map(r => ({ ...r, Date: toDate(r.Date) }));
  DATA.files = sheet("Files").filter(r => r.FileName).map(r => ({ ...r, UploadedOn: toDate(r.UploadedOn) }));
  const known = new Set(DATA.portcos.map(p => p.PortcoName));
  DATA.projects.forEach(p => { if (p.PortcoName && !known.has(p.PortcoName)) { DATA.portcos.push({ PortcoName: p.PortcoName }); known.add(p.PortcoName); } });

  $("sourceName").textContent = name;
  $("lastRefresh").textContent = new Date().toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" });
  setSync(usingLocalFile ? "Showing uploaded file" : "Live · synced with Excel", true);
  buildFilterOptions();
  render();
}
function toDate(v) {
  if (!v) return null;
  if (v instanceof Date) return isNaN(v) ? null : v;
  if (typeof v === "number") return new Date(Math.round((v - 25569) * 86400 * 1000));
  const d = new Date(v); return isNaN(d) ? null : d;
}
const fmt = (d) => d ? d.toLocaleDateString("en-GB", { day: "2-digit", month: "short", year: "numeric" }) : "–";
function setSync(text, ok) { const el = $("syncStatus"); el.textContent = (ok ? "● " : "⚠ ") + text; el.style.color = ok ? C.gold : C.amber; }

/* ---------- Health ---------- */
const isClosed = (p) => ["delivered", "on hold"].includes(String(p.Status).toLowerCase());
function health(p) {
  const s = String(p.Status).toLowerCase();
  if (s === "delivered") return "gold";
  if (s === "on hold") return "grey";
  if (!p.DueDate) return "green";
  const days = (p.DueDate - today()) / 86400000;
  if (days < 0) return "red";
  if (days <= CONFIG.dueSoonDays && p.Progress < 70) return "amber";
  return "green";
}

/* ---------- Filters ---------- */
const unique = (a) => [...new Set(a.filter(Boolean))];
function buildFilterOptions() {
  fill("fPortco", unique(DATA.portcos.map(p => p.PortcoName)), "ALL PORTCOS");
  fill("fStatus", unique([...CONFIG.statusOrder.filter(s => DATA.projects.some(p => p.Status === s)), ...DATA.projects.map(p => p.Status)]), "All statuses");
  fill("fOwner", unique(DATA.projects.map(p => p.IDOwner)), "All owners");
}
function fill(id, values, allLabel) {
  const el = $(id), cur = el.value;
  el.innerHTML = `<option value="">${allLabel}</option>` + values.map(v => `<option value="${esc(v)}">${esc(id === "fPortco" ? String(v).toUpperCase() : v)}</option>`).join("");
  if (values.includes(cur)) el.value = cur;
}
const F = () => ({ portco: $("fPortco").value, status: $("fStatus").value, owner: $("fOwner").value, q: $("fSearch").value.trim().toLowerCase() });
function filteredProjects() {
  const f = F();
  return DATA.projects.filter(p => (!f.portco || p.PortcoName === f.portco) && (!f.status || p.Status === f.status) &&
    (!f.owner || p.IDOwner === f.owner) && (!f.q || `${p.ProjectName} ${p.ProjectID}`.toLowerCase().includes(f.q)));
}

/* ---------- Render ---------- */
function render() {
  const projects = filteredProjects();
  const ids = new Set(projects.map(p => p.ProjectID));
  const f = F(), narrowed = f.status || f.owner || f.q;
  const comms = DATA.comms.filter(c => (!f.portco || c.PortcoName === f.portco) && (!narrowed || ids.has(c.ProjectID)));
  const files = DATA.files.filter(x => (!f.portco || x.PortcoName === f.portco) && (!narrowed || ids.has(x.ProjectID)));
  renderKPIs(projects, comms, files);
  renderStatus(projects);
  renderPortcoBars(projects);
  renderDelivery(projects);
  renderFunnel(projects);
  renderComms(comms);
  renderOwner(projects);
  renderPortcoCards();
  renderProjectTable(projects);
  renderCommsTable(comms);
  renderFilesTable(files);
}

function renderKPIs(projects, comms, files) {
  const active = projects.filter(p => !isClosed(p)).length;
  const overdue = projects.filter(p => health(p) === "red").length;
  const atRisk = projects.filter(p => health(p) === "amber").length;
  const avgFb = projects.length ? (projects.reduce((a, p) => a + p.FeedbackRounds, 0) / projects.length).toFixed(1) : "0";
  const sent = comms.filter(c => c.Direction === "Sent").length;
  const k = [
    ["PORTCOS", unique(projects.map(p => p.PortcoName)).length, "with projects", "gold"],
    ["TOTAL PROJECTS", projects.length, `${active} active`, ""],
    ["AT RISK", atRisk, `due ≤ ${CONFIG.dueSoonDays} days`, "amber"],
    ["OVERDUE", overdue, "past due date", "red"],
    ["COMMUNICATIONS", comms.length, `${sent} sent · ${comms.length - sent} received`, ""],
    ["FILES STORED", files.length, "scripts → final content", ""],
    ["AVG FEEDBACK ROUNDS", avgFb, "per project", "gold"]
  ];
  $("kpis").innerHTML = k.map(([l, v, n, c]) => `<div class="kpi ${c}"><div class="l">${l}</div><div class="v">${v}</div><div class="n">${n}</div></div>`).join("");
}

function renderStatus(projects) {
  const st = unique([...CONFIG.statusOrder, ...projects.map(p => p.Status)]).filter(s => projects.some(p => p.Status === s));
  const colors = st.map(s => C.status[CONFIG.statusOrder.indexOf(s)] || C.gold3);
  const counts = st.map(s => projects.filter(p => p.Status === s).length);
  drawChart("statusChart", "doughnut", { labels: st, datasets: [{ data: counts, backgroundColor: colors, borderColor: C.card, borderWidth: 4, spacing: 2 }] },
    { cutout: "80%", plugins: { legend: { display: false } } });
  $("statusCenter").textContent = projects.length;
  $("statusLegend").innerHTML = st.map((s, i) => `<li><span class="sw" style="background:${colors[i]}"></span>${esc(s)} – ${counts[i]}</li>`).join("") || `<li class="empty">No projects</li>`;
}

function renderPortcoBars(projects) {
  const pcs = unique(projects.map(p => p.PortcoName));
  const active = pcs.map(pc => projects.filter(p => p.PortcoName === pc && !isClosed(p)).length);
  const done = pcs.map(pc => projects.filter(p => p.PortcoName === pc && isClosed(p)).length);
  drawChart("portcoChart", "bar", {
    labels: pcs.map(shorten),
    datasets: [
      { label: "Active", data: active, backgroundColor: C.gold, borderRadius: 20, barThickness: 10 },
      { label: "Delivered / on hold", data: done, backgroundColor: C.gold3, borderRadius: 20, barThickness: 10 }
    ]
  }, { indexAxis: "y", scales: { x: { stacked: true, ticks: { precision: 0 }, grid: { color: C.line } }, y: { stacked: true, grid: { display: false } } },
       plugins: { legend: { position: "bottom", labels: { boxWidth: 10, boxHeight: 10, useBorderRadius: true, borderRadius: 3 } } } });
  $("activeTotal").textContent = active.reduce((a, b) => a + b, 0);
}

function renderDelivery(projects) {
  const delivered = projects.filter(p => health(p) === "gold").length;
  const pct = projects.length ? Math.round(delivered / projects.length * 100) : 0;
  drawChart("deliveryChart", "doughnut", {
    datasets: [{ data: [pct, 100 - pct], backgroundColor: [C.gold, C.track], borderWidth: 0, borderRadius: [20, 0] }]
  }, { cutout: "82%", rotation: 0, plugins: { legend: { display: false }, tooltip: { enabled: false } } });
  $("deliveryCenter").textContent = pct + "%";
  $("deliveredOf").textContent = `${delivered} of ${projects.length}`;
}

function renderFunnel(projects) {
  const stages = CONFIG.statusOrder.filter(s => s !== "On Hold");
  const total = projects.length || 1;
  const counts = stages.map(s => projects.filter(p => p.Status === s).length);
  const max = Math.max(1, ...counts);
  const lead = counts.indexOf(max);
  $("funnel").innerHTML = stages.map((s, i) => {
    const w = Math.round(counts[i] / max * 100);
    const avgDays = avgAge(projects.filter(p => p.Status === s));
    return `<div class="f-row">
      <div class="f-stage ${i === lead && counts[i] ? "lead" : ""}"><div class="fill" style="width:${i === lead ? 100 : 0}%;"></div>
        ${i !== lead ? `<div class="fill" style="width:${w}%;opacity:.18"></div>` : ""}
        <div class="t">${esc(s)}</div><div class="s">${avgDays !== null ? `avg ${avgDays} days since script` : "–"}</div></div>
      <div class="f-cell">${counts[i]}</div>
      <div class="f-cell">${Math.round(counts[i] / total * 100)}%</div>
    </div>`;
  }).join("");
}
function avgAge(list) {
  const ds = list.map(p => p.ScriptReceived).filter(Boolean);
  if (!ds.length) return null;
  return Math.round(ds.reduce((a, d) => a + (today() - d) / 86400000, 0) / ds.length);
}

function renderComms(comms) {
  const by = {};
  comms.forEach(c => { if (!c.Date) return; const k = `${c.Date.getFullYear()}-${String(c.Date.getMonth() + 1).padStart(2, "0")}`; by[k] = by[k] || { s: 0, r: 0 }; c.Direction === "Received" ? by[k].r++ : by[k].s++; });
  const m = Object.keys(by).sort();
  const lbl = (k) => new Date(k + "-01").toLocaleDateString("en-GB", { month: "short", year: "2-digit" });
  drawChart("commsChart", "bar", {
    labels: m.map(lbl),
    datasets: [
      { label: "Sent", data: m.map(k => by[k].s), backgroundColor: C.gold, borderRadius: 6, maxBarThickness: 34 },
      { label: "Received", data: m.map(k => by[k].r), backgroundColor: C.gold3, borderRadius: 6, maxBarThickness: 34 }
    ]
  }, { scales: { y: { beginAtZero: true, ticks: { precision: 0 }, grid: { color: C.line } }, x: { grid: { display: false } } },
       plugins: { legend: { position: "bottom", labels: { boxWidth: 10, boxHeight: 10 } } } });
}

function renderOwner(projects) {
  const act = projects.filter(p => !isClosed(p));
  const o = unique(act.map(p => p.IDOwner));
  drawChart("ownerChart", "bar", {
    labels: o, datasets: [{ label: "Active projects", data: o.map(x => act.filter(p => p.IDOwner === x).length), backgroundColor: C.gold, borderRadius: 8, maxBarThickness: 60 }]
  }, { scales: { y: { beginAtZero: true, ticks: { precision: 0 }, grid: { color: C.line } }, x: { grid: { display: false } } }, plugins: { legend: { display: false } } });
}

function drawChart(id, type, data, options) {
  if (charts[id]) charts[id].destroy();
  charts[id] = new Chart($(id), { type, data, options: { responsive: true, maintainAspectRatio: false, animation: { duration: 600 }, ...options } });
}

function renderPortcoCards() {
  const sel = $("fPortco").value;
  $("portcoCards").innerHTML = DATA.portcos.map(pc => {
    const ps = DATA.projects.filter(p => p.PortcoName === pc.PortcoName);
    const active = ps.filter(p => !isClosed(p)).length;
    const overdue = ps.filter(p => health(p) === "red").length;
    const cm = DATA.comms.filter(c => c.PortcoName === pc.PortcoName);
    const last = cm.map(c => c.Date).filter(Boolean).sort((a, b) => b - a)[0];
    const avg = ps.length ? Math.round(ps.reduce((a, p) => a + p.Progress, 0) / ps.length) : 0;
    return `<div class="pc ${sel === pc.PortcoName ? "active" : ""}" data-pc="${esc(pc.PortcoName)}">
      <h3>${esc(pc.PortcoName)}</h3><div class="ind">${esc(pc.Industry || "")}</div>
      <div class="stats"><div><b>${ps.length}</b>Projects</div><div><b>${active}</b>Active</div>
        <div><b class="${overdue ? "alert" : ""}">${overdue}</b>Overdue</div><div><b>${cm.length}</b>Comms</div></div>
      <div class="bar"><span style="width:${avg}%"></span></div>
      <div class="meta">Avg progress ${avg}% · SPOC ${esc(pc.SPOC || "–")}<br>Last contact ${fmt(last)}</div>
    </div>`;
  }).join("") || `<div class="empty">No portcos found in the Excel file.</div>`;
  document.querySelectorAll(".pc").forEach(el => el.onclick = () => { $("fPortco").value = $("fPortco").value === el.dataset.pc ? "" : el.dataset.pc; render(); });
}

function renderProjectTable(projects) {
  const rank = { red: 0, amber: 1, green: 2, grey: 3, gold: 4 };
  const rows = [...projects].sort((a, b) => rank[health(a)] - rank[health(b)] || (a.DueDate || 0) - (b.DueDate || 0));
  $("projectTable").querySelector("tbody").innerHTML = rows.map(p => {
    const h = health(p);
    return `<tr><td><i class="dot ${h}"></i>${HEALTH_LABEL[h]}</td>
      <td><strong>${esc(p.ProjectName)}</strong><br><small>${esc(p.ProjectID)}</small></td>
      <td>${esc(p.PortcoName)}</td><td>${esc(p.ContentType)}</td><td>${esc(p.IDOwner)}</td>
      <td><span class="pill st">${esc(p.Status)}</span></td>
      <td><div class="prog"><div class="bar"><span style="width:${p.Progress}%"></span></div>${p.Progress}%</div></td>
      <td>${fmt(p.DueDate)}</td><td style="text-align:center">${p.FeedbackRounds}</td>
      <td><span class="pill ${esc(p.Priority)}">${esc(p.Priority || "–")}</span></td></tr>`;
  }).join("") || `<tr><td colspan="10" class="empty">No projects match the filters.</td></tr>`;
}
function renderCommsTable(comms) {
  $("commsTable").querySelector("tbody").innerHTML = [...comms].sort((a, b) => (b.Date || 0) - (a.Date || 0)).map(c => `<tr>
    <td>${fmt(c.Date)}</td><td>${esc(c.PortcoName)}</td><td>${esc(c.ProjectID)}</td><td>${esc(c.Type)}</td>
    <td><span class="pill ${c.Direction === "Sent" ? "st" : ""}">${esc(c.Direction)}</span></td><td>${esc(c.Subject)}</td></tr>`).join("")
    || `<tr><td colspan="6" class="empty">No communications logged.</td></tr>`;
}
function renderFilesTable(files) {
  $("filesTable").querySelector("tbody").innerHTML = [...files].sort((a, b) => (b.UploadedOn || 0) - (a.UploadedOn || 0)).map(f => `<tr>
    <td>${f.Link ? `<a href="${esc(f.Link)}" target="_blank" rel="noopener">${esc(f.FileName)}</a>` : esc(f.FileName)}</td>
    <td>${esc(f.PortcoName)}</td><td>${esc(f.FileType)}</td><td>${esc(f.Version)}</td><td>${fmt(f.UploadedOn)}</td></tr>`).join("")
    || `<tr><td colspan="5" class="empty">No files logged.</td></tr>`;
}

const shorten = (s) => String(s).length > 20 ? String(s).slice(0, 19) + "…" : s;
function esc(v) { return String(v ?? "").replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c])); }

/* ---------- Wire up ---------- */
["fPortco", "fStatus", "fOwner"].forEach(id => $(id).addEventListener("change", render));
$("fSearch").addEventListener("input", render);
$("clearBtn").onclick = () => { ["fPortco", "fStatus", "fOwner", "fSearch"].forEach(id => $(id).value = ""); render(); };
$("refreshBtn").onclick = () => { usingLocalFile = false; loadFromRepo(true); };
["fileInput", "fileInputTop"].forEach(id => $(id).onchange = (e) => { if (e.target.files[0]) loadFromFile(e.target.files[0]); e.target.value = ""; });
document.querySelectorAll("nav a").forEach(a => a.addEventListener("click", () => {
  document.querySelectorAll("nav a").forEach(x => x.classList.remove("active")); a.classList.add("active");
}));

$("themeToggle").onclick = () => setTheme(theme === "dark" ? "light" : "dark");
setTheme(theme);

loadFromRepo(true);
setInterval(() => loadFromRepo(false), CONFIG.autoRefreshSeconds * 1000);
