import * as XLSX from "xlsx";
import "./style.css";

const EXPECTED = [
  "Comida", "Categoría", "Temporada", "Ingredientes", "Calorías",
  "Proteína (g)", "Carbohidratos (g)", "Grasas (g)", "Notas", "Link a receta"
];
const DAYS = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"];
const MEALS = ["Desayuno", "Comida", "Cena"];
const COLORS = ["blue", "green", "purple", "yellow", "pink", "teal", "lavender"];
let recipes = [];
let menu = {};
let filters = { category: "", season: "", calories: "", protein: "", carbs: "", fat: "", notes: "", search: "" };
let activeView = "week";
let weekOffset = 0;
let selectedSlot = null;

const app = document.querySelector("#app");
const escapeHTML = value => String(value ?? "").replace(/[&<>"']/g, ch => ({
  "&":"&amp;", "<":"&lt;", ">":"&gt;", '"':"&quot;", "'":"&#39;"
}[ch]));
const normalize = value => String(value ?? "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").trim().toLowerCase();
const num = value => {
  if (typeof value === "number") return Number.isFinite(value) ? value : 0;
  const parsed = Number(String(value ?? "").replace(",", ".").replace(/[^\d.-]/g, ""));
  return Number.isFinite(parsed) ? parsed : 0;
};
const get = (recipe, key) => recipe[key] ?? "";
const ingredients = recipe => String(get(recipe, "Ingredientes")).split(/[,;|]/).map(s => s.trim()).filter(Boolean);
const recipeId = recipe => String(recipe.__id);
const slotKey = (day, meal) => `${day}|${meal}`;
const selectedRecipes = () => Object.values(menu).flat().filter(Boolean);

function parseRow(row, id) {
  const normalized = {};
  for (const [key, value] of Object.entries(row)) normalized[normalize(key)] = value;
  const find = names => {
    for (const name of names) if (normalized[normalize(name)] !== undefined) return normalized[normalize(name)];
    return "";
  };
  const item = {
    "Comida": find(["Comida", "Nombre", "Receta"]),
    "Categoría": find(["Categoría", "Categoria"]),
    "Temporada": find(["Temporada"]),
    "Ingredientes": find(["Ingredientes"]),
    "Calorías": num(find(["Calorías", "Calorias", "Kcal"])),
    "Proteína (g)": num(find(["Proteína (g)", "Proteina (g)", "Proteína", "Proteina"])),
    "Carbohidratos (g)": num(find(["Carbohidratos (g)", "Carbohidratos"])),
    "Grasas (g)": num(find(["Grasas (g)", "Grasas"])),
    "Notas": find(["Notas", "Notas (ej. vegano, sin gluten, etc.)", "Etiquetas"]),
    "Link a receta": find(["Link a receta", "Link receta", "URL", "Receta URL"]),
    __id: id
  };
  return item;
}

function sampleRecipes() {
  return [
    ["Avena con plátano y frutos secos","Desayuno","Todo el año","avena, plátano, nueces, leche",420,12,65,14,"vegetariano",""],
    ["Yogur con frutos rojos y granola","Desayuno","Verano","yogur, frutos rojos, granola",380,15,50,12,"vegetariano",""],
    ["Tostada de aguacate y huevo","Desayuno","Primavera","pan integral, aguacate, huevo",390,14,40,18,"vegetariano",""],
    ["Pollo con verduras a la plancha","Comida","Todo el año","pollo, calabacín, pimiento, aceite de oliva",520,35,45,18,"sin gluten",""],
    ["Salmón al horno con espárragos","Comida","Primavera","salmón, espárragos, limón",480,32,30,24,"sin gluten",""],
    ["Pasta integral con pesto y tomate cherry","Comida","Verano","pasta integral, pesto, tomate cherry",480,20,70,16,"vegetariano",""],
    ["Ensalada de garbanzos","Cena","Verano","garbanzos, tomate, pepino, cebolla",410,18,55,16,"vegano; sin lactosa",""],
    ["Crema de calabaza","Cena","Otoño","calabaza, zanahoria, cebolla",360,10,45,12,"vegano; sin gluten",""],
    ["Tortilla de verduras","Cena","Todo el año","huevo, espinacas, calabacín",360,12,40,14,"sin gluten; vegetariano",""]
  ].map((r, i) => parseRow(Object.fromEntries(EXPECTED.map((key, j) => [key, r[j]])), i + 1));
}

function allUnique(key) {
  return [...new Set(recipes.map(r => String(get(r, key)).trim()).filter(Boolean))]
    .sort((a,b) => a.localeCompare(b, "es"));
}
function filteredRecipes() {
  const q = normalize(filters.search);
  return recipes.filter(r => {
    const searchable = normalize([r["Comida"], r["Ingredientes"], r["Notas"], r["Categoría"], r["Temporada"]].join(" "));
    return (!filters.category || normalize(r["Categoría"]) === normalize(filters.category))
      && (!filters.season || normalize(r["Temporada"]) === normalize(filters.season))
      && (!filters.calories || num(r["Calorías"]) <= num(filters.calories))
      && (!filters.protein || num(r["Proteína (g)"]) >= num(filters.protein))
      && (!filters.carbs || num(r["Carbohidratos (g)"]) <= num(filters.carbs))
      && (!filters.fat || num(r["Grasas (g)"]) <= num(filters.fat))
      && (!filters.notes || normalize(r["Notas"]).includes(normalize(filters.notes)))
      && (!q || searchable.includes(q));
  });
}
function optionList(values, selected, label) {
  return `<option value="">${label}</option>` + values.map(v => `<option value="${escapeHTML(v)}" ${v === selected ? "selected" : ""}>${escapeHTML(v)}</option>`).join("");
}
function setFilter(key, value) {
  filters[key] = value;
  render();
}
function nutrition(recipe) {
  return `<div class="nutrition">
    <span>🔥 ${num(recipe["Calorías"])} kcal</span>
    <span>Proteína ${num(recipe["Proteína (g)"])} g</span>
    <span>Carb. ${num(recipe["Carbohidratos (g)"])} g</span>
    <span>Grasas ${num(recipe["Grasas (g)"])} g</span>
  </div>`;
}
function recipeCard(recipe, compact = false) {
  const tags = String(recipe["Notas"] || "").split(/[,;|]/).map(s => s.trim()).filter(Boolean).slice(0, 3);
  return `<article class="recipe-card ${compact ? "compact" : ""}">
    <div class="recipe-title">${escapeHTML(recipe["Comida"] || "Receta sin nombre")}</div>
    <div class="recipe-meta">${escapeHTML(recipe["Categoría"] || "Sin categoría")} · ${escapeHTML(recipe["Temporada"] || "Temporada no indicada")}</div>
    ${recipe["Ingredientes"] ? `<p class="ingredients">${escapeHTML(recipe["Ingredientes"])}</p>` : ""}
    ${nutrition(recipe)}
    ${tags.length ? `<div class="tags">${tags.map(t => `<span class="tag">${escapeHTML(t)}</span>`).join("")}</div>` : ""}
    <div class="card-actions">
      ${recipe["Link a receta"] && /^https?:\/\//i.test(String(recipe["Link a receta"])) ? `<a class="text-link" href="${escapeHTML(recipe["Link a receta"])}" target="_blank" rel="noopener noreferrer">Ver receta ↗</a>` : `<span class="muted small">Sin enlace</span>`}
      ${selectedSlot ? `<button class="btn btn-primary btn-small" data-add="${recipeId(recipe)}">Añadir aquí</button>` : ""}
    </div>
  </article>`;
}
function weekLabel() {
  const start = new Date();
  const day = (start.getDay() + 6) % 7;
  start.setDate(start.getDate() - day + weekOffset * 7);
  const end = new Date(start); end.setDate(start.getDate() + 6);
  const fmt = d => d.toLocaleDateString("es-ES", { day: "numeric", month: "short" }).replace(".", "");
  return `${fmt(start)} – ${fmt(end)} ${end.getFullYear()}`;
}
function slotMarkup(day, meal, dIndex, mIndex) {
  const key = slotKey(day, meal);
  const assigned = (menu[key] || []).map(id => recipes.find(r => recipeId(r) === String(id))).filter(Boolean);
  return `<section class="meal-slot">
    <div class="slot-heading"><span class="meal-dot dot-${mIndex}"></span>${meal}</div>
    ${assigned.map((r, i) => `<div class="assigned-recipe">
      <div class="assigned-name">${escapeHTML(r["Comida"])}</div>
      <div class="assigned-stats">${num(r["Calorías"])} kcal · ${num(r["Proteína (g)"])} g proteína</div>
      <button class="remove-btn" title="Quitar del menú" aria-label="Quitar ${escapeHTML(r["Comida"])}" data-remove="${escapeHTML(key)}" data-index="${i}">×</button>
    </div>`).join("")}
    <button class="add-slot" data-slot="${escapeHTML(key)}">＋ Añadir receta</button>
  </section>`;
}
function renderWeek() {
  return `<div class="week-grid">${DAYS.map((day, i) => `<section class="day-column day-${i}">
    <header class="day-heading"><strong>${day}</strong><span>${dayDate(i)}</span></header>
    ${MEALS.map((meal, m) => slotMarkup(day, meal, i, m)).join("")}
  </section>`).join("")}</div>`;
}
function dayDate(index) {
  const start = new Date();
  start.setDate(start.getDate() - ((start.getDay() + 6) % 7) + weekOffset * 7 + index);
  return start.toLocaleDateString("es-ES", { day: "numeric", month: "short" }).replace(".", "");
}
function renderRecipeList() {
  const list = filteredRecipes();
  return `<div class="recipe-results-head"><div><h2>${activeView === "week" ? "Añadir una receta" : "Todas las recetas"}</h2><p class="muted">${list.length} receta${list.length === 1 ? "" : "s"} encontrada${list.length === 1 ? "" : "s"}${selectedSlot ? ` · destino: ${escapeHTML(selectedSlot.replace("|", " / "))}` : ""}</p></div>
    ${selectedSlot ? `<button class="btn btn-quiet" id="cancel-add">Cancelar</button>` : ""}</div>
    ${list.length ? `<div class="recipe-list">${list.map(r => recipeCard(r)).join("")}</div>` : `<div class="empty-state"><span>🍲</span><h3>No hay recetas con estos filtros</h3><p>Prueba a quitar algún filtro o cambia el texto de búsqueda.</p><button class="btn btn-quiet" id="reset-filters">Limpiar filtros</button></div>`}`;
}
function average(key) {
  const list = selectedRecipes();
  return list.length ? Math.round(list.reduce((sum, r) => sum + num(r[key]), 0) / 7) : 0;
}
function renderSidebar() {
  const count = selectedRecipes().length;
  const totals = ["Calorías","Proteína (g)","Carbohidratos (g)","Grasas (g)"].map(key => {
    const total = selectedRecipes().reduce((sum, r) => sum + num(r[key]), 0);
    const unit = key === "Calorías" ? "kcal" : "g";
    const title = key === "Proteína (g)" ? "Proteína" : key === "Carbohidratos (g)" ? "Carbohidratos" : key === "Grasas (g)" ? "Grasas" : "Calorías";
    return `<div class="summary-row"><span>${title}</span><strong>${Math.round(total / 7)} ${unit}<small>/ día</small></strong></div>`;
  }).join("");
  const ingredientCounts = {};
  selectedRecipes().forEach(r => ingredients(r).forEach(ing => {
    const k = ing.trim().toLowerCase(); if (k) ingredientCounts[k] = (ingredientCounts[k] || 0) + 1;
  }));
  const topIngredients = Object.entries(ingredientCounts).sort((a,b) => b[1]-a[1]).slice(0,8);
  return `<aside class="sidebar">
    <section class="panel"><h3>Resumen nutricional</h3><p class="muted small">Media diaria del menú</p>${totals}</section>
    <section class="panel"><h3>Recetas del menú <span class="count-badge">${count}</span></h3>
      ${count ? selectedRecipes().slice(0,6).map(r => `<div class="side-recipe"><strong>${escapeHTML(r["Comida"])}</strong><span>${escapeHTML(r["Categoría"] || "Sin categoría")} · ${num(r["Calorías"])} kcal</span></div>`).join("") : `<p class="muted small">Todavía no has añadido recetas. Pulsa «Añadir receta» en un día.</p>`}
    </section>
    ${topIngredients.length ? `<section class="panel"><h3>Ingredientes destacados</h3><div class="tags">${topIngredients.map(([ing]) => `<button class="tag tag-button" data-ingredient="${escapeHTML(ing)}">${escapeHTML(ing)}</button>`).join("")}</div></section>` : ""}
    <div class="privacy-note"><strong>🔒 Tu archivo es privado</strong><p>El Excel se procesa en este navegador. No se sube a ningún servidor.</p></div>
  </aside>`;
}
function render() {
  const categories = allUnique("Categoría"), seasons = allUnique("Temporada");
  app.innerHTML = `<div class="app-shell">
    <aside class="left-nav">
      <div class="brand"><div class="brand-mark">♧</div><div><strong>Mis Recetas</strong><span>Planificador semanal</span></div></div>
      <nav>
        <button class="nav-item ${activeView === "week" ? "active" : ""}" data-view="week">▦ <span>Menú semanal</span></button>
        <button class="nav-item ${activeView === "recipes" ? "active" : ""}" data-view="recipes">⌕ <span>Todas las recetas</span></button>
        <button class="nav-item ${activeView === "about" ? "active" : ""}" data-view="about">ⓘ <span>Acerca de</span></button>
      </nav>
      <div class="upload-panel">
        <div class="upload-icon">⇧</div><strong>Sube tu archivo Excel</strong>
        <p>Arrastra un archivo .xlsx o .xls aquí, o selecciónalo desde tu equipo.</p>
        <label class="btn btn-primary upload-button" for="excel-file">Seleccionar archivo</label>
        <input id="excel-file" type="file" accept=".xlsx,.xls,.csv" />
        <span class="upload-status">${recipes.length ? `${recipes.length} recetas cargadas` : "Puedes probar con datos de ejemplo o cargar tu archivo."}</span>
        <button class="text-link sample-button" id="load-sample">Cargar datos de ejemplo</button>
      </div>
      <div class="nav-footer">Hecho para planificar mejor <span>🌱</span></div>
    </aside>
    <main class="main-content">
      <header class="topbar"><div><div class="eyebrow">ORGANIZA TU COCINA</div><h1>${activeView === "week" ? "Tu menú semanal" : activeView === "recipes" ? "Todas las recetas" : "Acerca de Mis Recetas"}</h1><p>${activeView === "week" ? "Encuentra tus platos favoritos y organiza la semana a tu manera." : activeView === "recipes" ? "Busca y filtra las recetas de tu Excel." : "Una forma sencilla de convertir tu Excel en un menú semanal."}</p></div>
        <div class="top-actions"><button class="btn btn-quiet" id="export-menu">↧ <span>Exportar menú</span></button><button class="btn btn-quiet" id="clear-menu">Vaciar menú</button></div>
      </header>
      ${activeView === "about" ? `<section class="panel about-panel"><h2>Cómo funciona</h2><ol><li>Sube un Excel con una sola hoja y las columnas de recetas.</li><li>Combina filtros por categoría, temporada, calorías, macros, notas o ingredientes.</li><li>En el menú semanal, pulsa «Añadir receta» en el día y la comida que quieras.</li><li>Exporta el menú a CSV para guardarlo o compartirlo.</li></ol><p>El archivo se procesa localmente en tu navegador. No se almacena en un servidor.</p></section>` : `
      <section class="panel filters-panel"><div class="section-title"><h2>⌘ Filtros</h2><button class="text-link" id="reset-filters">Restablecer</button></div>
        <div class="filters-grid">
          <label>Categoría<select id="filter-category">${optionList(categories, filters.category, "Todas")}</select></label>
          <label>Temporada<select id="filter-season">${optionList(seasons, filters.season, "Todas")}</select></label>
          <label>Calorías máx.<input id="filter-calories" type="number" min="0" placeholder="Sin límite" value="${escapeHTML(filters.calories)}" /></label>
          <label>Proteína mín. (g)<input id="filter-protein" type="number" min="0" placeholder="Sin mínimo" value="${escapeHTML(filters.protein)}" /></label>
          <label>Carbohidratos máx. (g)<input id="filter-carbs" type="number" min="0" placeholder="Sin límite" value="${escapeHTML(filters.carbs)}" /></label>
          <label>Grasas máx. (g)<input id="filter-fat" type="number" min="0" placeholder="Sin límite" value="${escapeHTML(filters.fat)}" /></label>
          <label class="search-label">Buscar receta o ingrediente<input id="filter-search" type="search" placeholder="Ej. pollo, garbanzos, pasta..." value="${escapeHTML(filters.search)}" /></label>
          <label>Notas / etiquetas<input id="filter-notes" type="search" placeholder="Ej. vegano, sin gluten..." value="${escapeHTML(filters.notes)}" /></label>
        </div>
      </section>
      <div class="content-layout"><section class="content-primary">
        ${activeView === "week" ? `<div class="week-toolbar"><div><h2>Tu semana</h2><p class="muted">Puedes añadir varias recetas a cada franja.</p></div><div class="week-nav"><button class="icon-btn" id="prev-week" aria-label="Semana anterior">‹</button><span>${weekLabel()}</span><button class="icon-btn" id="next-week" aria-label="Semana siguiente">›</button></div></div>${renderWeek()}` : ""}
        <div id="recipe-area">${selectedSlot || activeView === "recipes" ? renderRecipeList() : ""}</div>
      </section>${renderSidebar()}</div>`}
    </main>
  </div>`;
  bindEvents();
}
function resetFilters() {
  filters = { category:"", season:"", calories:"", protein:"", carbs:"", fat:"", notes:"", search:"" };
  render();
}
function bindEvents() {
  app.querySelectorAll("[data-view]").forEach(el => el.addEventListener("click", () => { activeView = el.dataset.view; selectedSlot = null; render(); }));
  app.querySelector("#filter-category")?.addEventListener("change", e => setFilter("category", e.target.value));
  app.querySelector("#filter-season")?.addEventListener("change", e => setFilter("season", e.target.value));
  [["calories","calories"],["protein","protein"],["carbs","carbs"],["fat","fat"],["notes","notes"],["search","search"]].forEach(([id,key]) => {
    app.querySelector(`#filter-${id}`)?.addEventListener(id === "notes" || id === "search" ? "input" : "change", e => {
      const cursor = e.target.selectionStart; filters[key] = e.target.value; render();
      const next = app.querySelector(`#filter-${id}`); next?.focus(); if (cursor !== null && next?.setSelectionRange) next.setSelectionRange(cursor, cursor);
    });
  });
  app.querySelectorAll("[data-slot]").forEach(btn => btn.addEventListener("click", () => { selectedSlot = btn.dataset.slot; activeView = "week"; render(); app.querySelector("#recipe-area")?.scrollIntoView({ behavior:"smooth", block:"start" }); }));
  app.querySelectorAll("[data-add]").forEach(btn => btn.addEventListener("click", () => {
    if (!selectedSlot) return;
    const id = btn.dataset.add;
    menu[selectedSlot] ||= [];
    menu[selectedSlot].push(id);
    selectedSlot = null; render();
  }));
  app.querySelectorAll("[data-remove]").forEach(btn => btn.addEventListener("click", () => {
    const key = btn.dataset.remove; const index = Number(btn.dataset.index);
    if (menu[key]) { menu[key].splice(index, 1); if (!menu[key].length) delete menu[key]; }
    render();
  }));
  app.querySelectorAll("[data-ingredient]").forEach(btn => btn.addEventListener("click", () => {
    filters.search = btn.dataset.ingredient; activeView = "recipes"; selectedSlot = null; render();
  }));
  app.querySelector("#cancel-add")?.addEventListener("click", () => { selectedSlot = null; render(); });
  app.querySelectorAll("#reset-filters").forEach(btn => btn.addEventListener("click", resetFilters));
  app.querySelector("#prev-week")?.addEventListener("click", () => { weekOffset--; render(); });
  app.querySelector("#next-week")?.addEventListener("click", () => { weekOffset++; render(); });
  app.querySelector("#clear-menu")?.addEventListener("click", () => {
    if (selectedRecipes().length && !confirm("¿Quieres quitar todas las recetas del menú semanal?")) return;
    menu = {}; render();
  });
  app.querySelector("#export-menu")?.addEventListener("click", exportMenu);
  app.querySelector("#load-sample")?.addEventListener("click", () => { recipes = sampleRecipes(); menu = {}; activeView = "week"; selectedSlot = null; render(); });
  const fileInput = app.querySelector("#excel-file");
  fileInput?.addEventListener("change", async e => { const file = e.target.files?.[0]; if (file) await loadFile(file); });
  const uploadPanel = app.querySelector(".upload-panel");
  uploadPanel?.addEventListener("dragover", e => { e.preventDefault(); uploadPanel.classList.add("dragging"); });
  uploadPanel?.addEventListener("dragleave", () => uploadPanel.classList.remove("dragging"));
  uploadPanel?.addEventListener("drop", async e => { e.preventDefault(); uploadPanel.classList.remove("dragging"); const file = e.dataTransfer.files?.[0]; if (file) await loadFile(file); });
}
async function loadFile(file) {
  if (!/\.(xlsx|xls|csv)$/i.test(file.name)) { alert("Selecciona un archivo .xlsx, .xls o .csv."); return; }
  try {
    const buffer = await file.arrayBuffer();
    const workbook = XLSX.read(buffer, { type:"array", cellDates:true });
    const sheetName = workbook.SheetNames[0];
    if (!sheetName) throw new Error("El archivo no contiene ninguna hoja.");
    const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { defval:"", raw:true });
    if (!rows.length) throw new Error("La primera hoja está vacía o no tiene encabezados.");
    const headers = Object.keys(rows[0]).map(normalize);
    const missing = EXPECTED.filter(h => !headers.includes(normalize(h)) && !(h === "Notas" && headers.includes("notas (ej. vegano, sin gluten, etc.)")));
    const required = ["Comida", "Categoría", "Temporada", "Ingredientes", "Calorías"];
    const missingRequired = required.filter(h => !headers.includes(normalize(h)));
    if (missingRequired.length) throw new Error(`Faltan columnas necesarias: ${missingRequired.join(", ")}. Revisa que la primera fila contenga los encabezados.`);
    const nextRecipes = rows.map((r, i) => parseRow(r, i + 1)).filter(r => String(r["Comida"]).trim());
    if (!nextRecipes.length) throw new Error("No se encontraron recetas con nombre en la columna «Comida».");
    recipes = nextRecipes; menu = {}; filters = { category:"", season:"", calories:"", protein:"", carbs:"", fat:"", notes:"", search:"" };
    activeView = "week"; selectedSlot = null; render();
    if (missing.length) alert(`Se han cargado ${recipes.length} recetas. Algunas columnas opcionales no se han encontrado: ${missing.join(", ")}. Comprueba los encabezados si esos datos no aparecen.`);
  } catch (err) {
    alert(`No se ha podido leer el Excel: ${err.message || err}`);
  }
}
function exportMenu() {
  const rows = [];
  DAYS.forEach(day => MEALS.forEach(meal => {
    (menu[slotKey(day, meal)] || []).forEach(id => {
      const r = recipes.find(x => recipeId(x) === String(id));
      if (r) rows.push({
        Día: day, Momento: meal, Comida: r["Comida"], Categoría: r["Categoría"], Temporada: r["Temporada"],
        Ingredientes: r["Ingredientes"], Calorías: r["Calorías"], "Proteína (g)": r["Proteína (g)"],
        "Carbohidratos (g)": r["Carbohidratos (g)"], "Grasas (g)": r["Grasas (g)"],
        Notas: r["Notas"], "Link a receta": r["Link a receta"]
      });
    });
  }));
  if (!rows.length) { alert("Añade al menos una receta al menú antes de exportar."); return; }
  const csv = "\ufeff" + [Object.keys(rows[0]).join(";"), ...rows.map(row => Object.values(row).map(v => `"${String(v ?? "").replace(/"/g, '""')}"`).join(";"))].join("\r\n");
  const url = URL.createObjectURL(new Blob([csv], { type:"text/csv;charset=utf-8;" }));
  const a = document.createElement("a"); a.href = url; a.download = "mi-menu-semanal.csv"; a.click(); URL.revokeObjectURL(url);
}
recipes = sampleRecipes();
render();
