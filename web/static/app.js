"use strict";
const $ = (id) => document.getElementById(id);
let csrf = "", currentView = "dashboard", page = 1, total = 0, pageSize = 50;
let initialized = false, running = false, timer, controller, debounce, viewSequence = 0;
const number = new Intl.NumberFormat("hu-HU", {maximumFractionDigits: 2});
const titles = {master: "Részletes adatok", ktj: "Telephelyek / KTJ", source: "Cégadatok", missing: "Kimaradt / hibás lekérések"};
const labels = {
  CEGNEV: "Cég", CEGNEV_FORRAS: "Forrás szerinti cégnév", SOURCE_KUJ: "Forrás KÜJ", KUJ: "KÜJ",
  EV: "Év", ADATTIPUS: "Adattípus kód", ADATTIPUS_LABEL: "Adattípus", PRODUCT_GROUP: "Termékcsoport",
  HULLADEKKOD: "HAK kód", HULLADEK_MEGNEVEZES: "Hulladék megnevezése", HULLADEK_VESZELYES_10: "Veszélyes",
  OSSZES_MENNYISEG: "Mennyiség (kg)", MENNYISEG: "Mennyiség (kg)", SZEKHELY: "Székhely",
  ADOSZAM: "Adószám", SRC_EMAIL: "Email", SRC_TELEFON: "Telefon", SRC_WEB: "Web",
  SRC_SAJAT_TOKE: "Saját tőke", SRC_NETTO_ARBEV: "Nettó árbevétel", SRC_CEGFORMA: "Cégforma",
  SRC_ALAPITAS_DATUMA: "Alapítás", SRC_UT_LETSZAM: "Létszám", SRC_FO_TEV: "Fő tevékenység",
  SRC_VEZETOK_SZAMA: "Vezetők száma", SRC_TISZTSEGVISELOK: "Tisztségviselők", SRC_TELEPHELYEK_SZAMA: "Telephelyek száma"
};
function showError(error) {
  $("error").textContent = error.message || String(error);
  $("error").hidden = false;
}
async function api(url, options = {}) {
  const headers = {...options.headers};
  if (options.method && options.method !== "GET") headers["X-CSRF-Token"] = csrf;
  const response = await fetch(url, {...options, headers, cache: "no-store"});
  const data = await response.json();
  if (!response.ok) {
    if (response.status === 401) window.location.assign("/login");
    throw new Error(data.error || `HTTP ${response.status}`);
  }
  return data;
}
function params() {
  const result = new URLSearchParams(new FormData($("filters")));
  result.set("view", currentView === "dashboard" ? "master" : currentView);
  result.set("page", page);
  return result;
}
function updateExports() {
  const query = params();
  query.delete("view"); query.delete("page");
  document.querySelectorAll("[data-export]").forEach((link) => {
    link.href = `/api/export/${link.dataset.export}?${query}`;
  });
  const selected = ["companies", "years", "categories", "types"].map((id) =>
    Array.from($(id).selectedOptions, (option) => option.textContent).join(", ")).filter(Boolean);
  if ($("code").value.trim()) selected.push(`HAK: ${$("code").value.trim()}`);
  $("selection-summary").textContent = `Aktuális szűrés: ${selected.join(" · ") || "összes adat"}`;
}
function options(id, values, names = {}) {
  const selected = new Set(Array.from($(id).selectedOptions, (option) => option.value));
  $(id).replaceChildren(...values.map((value) => {
    const option = new Option(names[value] || String(value), String(value));
    option.selected = selected.has(String(value));
    return option;
  }));
}
async function refreshState(reset = false) {
  const state = await api("/api/state");
  csrf = state.csrf;
  $("filename").textContent = state.filename ? `${state.filename} · ${number.format(state.rows)} nyers adatsor · ${number.format(state.missing)} kimaradt / hiba` : "Tölts be egy forrásfájlt vagy korábbi mentést.";
  for (const id of ["companies", "years", "categories"]) options(id, state.options[id]);
  options("types", Object.keys(state.types), state.types);
  if (!initialized || (reset && state.job.status === "done")) {
    for (const id of ["companies", "years", "categories", "types"]) {
      Array.from($(id).options).forEach((option, index) => option.selected = id === "years" && index < 2);
    }
    $("code").value = "";
    page = 1;
  }
  initialized = true;
  const previous = $("snapshots").value;
  $("snapshots").replaceChildren(...state.snapshots.map(([id, date, file, note]) =>
    new Option(`#${id} · ${date.slice(0, 16).replace("T", " ")} · ${note || file}`, id)));
  if (state.snapshots.some((row) => String(row[0]) === previous)) $("snapshots").value = previous;
  displayJob(state.job);
  if (!state.rows) $("import-panel").open = true;
  else if (reset && state.job.status === "done") $("import-panel").open = false;
  return state;
}
function displayJob(job) {
  running = job.status === "running";
  $("status").textContent = job.message;
  $("progress").max = job.total;
  $("progress").value = job.step;
  document.querySelectorAll(".mutation").forEach((button) => button.disabled = running);
  if (job.status === "error") showError(new Error(job.message));
}
function chart(id, rows, key, filter) {
  const maximum = Math.max(1, ...rows.map((row) => Math.abs(row.OSSZES_MENNYISEG)));
  $(id).replaceChildren(...rows.map((row) => {
    const button = document.createElement("button");
    button.className = "bar";
    const label = document.createElement("span"); label.className = "bar-label";
    const text = document.createElement("span"); text.textContent = row[key] ?? "Ismeretlen";
    const value = document.createElement("b"); value.textContent = number.format(row.OSSZES_MENNYISEG / 1000);
    label.append(text, value);
    const track = document.createElement("span"); track.className = "bar-track";
    const fill = document.createElement("span"); fill.className = "bar-fill";
    fill.style.width = `${Math.abs(row.OSSZES_MENNYISEG) / maximum * 100}%`;
    track.append(fill); button.append(label, track);
    button.title = `${text.textContent}: ${value.textContent} t — szűrés`;
    button.addEventListener("click", () => {
      if (filter === "code") $("code").value = String(row[key]);
      else Array.from($(filter).options).forEach((option) => option.selected = option.value === String(row[key]));
      page = 1; refreshView().catch(showError);
    });
    return button;
  }));
}
function renderSummary(data) {
  $("kpi-quantity").textContent = number.format(data.quantity_kg / 1000);
  $("kpi-companies").textContent = number.format(data.companies);
  $("kpi-codes").textContent = number.format(data.codes);
  $("kpi-danger").textContent = `${number.format(data.dangerous_percent)}%`;
  document.querySelector('[data-export="matrix"]').setAttribute("aria-disabled", String(data.companies === 0));
  chart("yearly", data.yearly, "EV", "years");
  chart("category-chart", data.categories, "PRODUCT_GROUP", "categories");
  chart("company-chart", data.company_totals, "CEGNEV", "companies");
  chart("code-chart", data.waste_codes, "HULLADEKKOD", "code");
}
function renderTable(data) {
  const columns = data.columns.filter((column) => column !== "ADATTIPUS");
  const header = document.createElement("tr");
  columns.forEach((column) => {
    const cell = document.createElement("th");
    cell.scope = "col"; cell.textContent = labels[column] || column; header.append(cell);
  });
  $("table").tHead.replaceChildren(header);
  $("table").tBodies[0].replaceChildren(...data.rows.map((row) => {
    const tr = document.createElement("tr");
    columns.forEach((column) => {
      const cell = document.createElement("td");
      const value = row[column];
      cell.textContent = value == null ? "" : column === "HULLADEK_VESZELYES_10" ? (value === 1 ? "Igen" : "Nem")
        : ["OSSZES_MENNYISEG", "MENNYISEG"].includes(column) ? number.format(value) : String(value);
      cell.title = cell.textContent; tr.append(cell);
    });
    return tr;
  }));
  $("page").textContent = `${page} / ${Math.max(1, Math.ceil(total / pageSize))} oldal · ${number.format(total)} sor`;
  $("previous").disabled = page <= 1;
  $("next").disabled = page * pageSize >= total;
}
async function refreshView() {
  if (controller) controller.abort();
  controller = new AbortController();
  const sequence = ++viewSequence;
  updateExports();
  try {
    const data = await api(`/api/view?${params()}`, {signal: controller.signal});
    if (sequence !== viewSequence) return;
    page = data.page; total = data.total; pageSize = data.page_size;
    renderSummary(data.summary);
    const dashboard = currentView === "dashboard";
    $("dashboard").hidden = !dashboard || !data.summary.companies;
    $("empty").hidden = !dashboard || !!data.summary.companies;
    $("table-panel").hidden = dashboard;
    $("table-title").textContent = titles[currentView] || "";
    $("load-ktj").hidden = currentView !== "ktj";
    $("table-note").textContent = currentView === "missing" ? "A teljes munkamenet hibajegyzéke, a szűrőktől függetlenül."
      : currentView === "ktj" ? "A metaadatok külön tölthetők be, hálózati lekéréssel. A táblázat oldalanként 50 sort mutat."
      : currentView === "source" ? "Az importált cégadatok; csak a cégnévszűrő vonatkozik erre a nézetre."
      : "Oldalanként 50 sor. A termékcsoportok átfedhetnek; az összes mennyiséget a fenti mutató adja.";
    renderTable(data);
  } catch (error) {
    if (error.name !== "AbortError") throw error;
  }
}
async function poll(reset = false) {
  clearTimeout(timer);
  const state = await refreshState(reset);
  if (state.job.status === "running") {
    timer = setTimeout(() => poll(reset).catch(showError), 1000);
  } else {
    await refreshView();
  }
}
async function mutate(url, options = {}, reset = false) {
  $("error").hidden = true;
  displayJob({status: "running", step: 0, total: 1, message: "Művelet indítása..."});
  try {
    await api(url, {method: "POST", ...options});
    await poll(reset);
  } catch (error) {
    displayJob({status: "error", step: 0, total: 1, message: error.message});
    throw error;
  }
}
function handle(id, callback) {
  $(id).addEventListener("click", () => Promise.resolve().then(callback).catch(showError));
}
$("filters").addEventListener("submit", (event) => event.preventDefault());
$("filters").addEventListener("input", () => {
  clearTimeout(debounce);
  debounce = setTimeout(() => {page = 1; refreshView().catch(showError);}, 220);
});
handle("reset", () => {
  for (const id of ["companies", "years", "categories", "types"]) Array.from($(id).options).forEach((option) => option.selected = false);
  $("code").value = ""; page = 1; return refreshView();
});
document.querySelectorAll("[data-view]").forEach((button) => button.addEventListener("click", () => {
  currentView = button.dataset.view; page = 1;
  document.querySelectorAll("[data-view]").forEach((tab) => tab.setAttribute("aria-pressed", String(tab === button)));
  refreshView().catch(showError);
}));
$("import-form").addEventListener("submit", (event) => {
  event.preventDefault();
  const data = new FormData(event.currentTarget);
  data.set("use_cache", String(event.currentTarget.elements.use_cache.checked));
  data.set("force_refresh", String(event.currentTarget.elements.force_refresh.checked));
  mutate("/api/import", {body: data}, true).catch(showError);
});
$("legacy-import").addEventListener("submit", (event) => {
  event.preventDefault();
  mutate("/api/import", {body: new FormData(event.currentTarget)}, true).catch(showError);
});
document.querySelector('[data-export="matrix"]').addEventListener("click", (event) => {
  if (event.currentTarget.getAttribute("aria-disabled") === "true") {
    event.preventDefault();
    showError(new Error("Nincs a szűrésnek megfelelő adat. Tölts be adatokat vagy módosítsd a szűrést."));
  }
});
handle("previous", () => {page = Math.max(1, page - 1); return refreshView();});
handle("next", () => {page += 1; return refreshView();});
handle("load-ktj", () => mutate(`/api/ktj?${params()}`));
handle("save", () => mutate("/api/snapshots", {headers: {"Content-Type": "application/json"}, body: JSON.stringify({note: $("note").value})}));
handle("load", () => {
  if (!$("snapshots").value) throw new Error("Nincs kiválasztott mentés.");
  return mutate(`/api/snapshots/${$("snapshots").value}/load`, {}, true);
});
handle("delete", () => {
  if (!$("snapshots").value) throw new Error("Nincs kiválasztott mentés.");
  if (confirm("Végleg törlöd a kiválasztott mentést?")) return mutate(`/api/snapshots/${$("snapshots").value}`, {method: "DELETE"});
});
handle("clear-cache", () => {
  if (confirm("Törlöd az OKIR lekérési gyorsítótárat? A mentések megmaradnak.")) return mutate("/api/cache", {method: "DELETE"});
});
handle("logout", async () => {await api("/api/logout", {method: "POST"}); window.location.assign("/login");});
document.querySelector('a[href="#import-panel"]').addEventListener("click", () => $("import-panel").open = true);
window.addEventListener("offline", () => {
  $("connection").textContent = "Offline";
  showError(new Error("Nincs hálózati kapcsolat. A megjelenített adatok nem frissülnek; a lekérésekhez és exportokhoz elérhető szerver kell."));
});
window.addEventListener("online", () => {
  $("connection").textContent = "Online"; $("error").hidden = true;
  poll().catch(showError);
});
async function start() {
  if (window.matchMedia("(max-width: 760px)").matches) $("filter-panel").open = false;
  await refreshState();
  await refreshView();
  if (running) await poll();
  if ("serviceWorker" in navigator && window.isSecureContext) {
    try {await navigator.serviceWorker.register("/sw.js");}
    catch (error) {showError(new Error(`A főképernyős mód regisztrációja nem sikerült: ${error.message}`));}
  }
}
start().catch(showError);
