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
const selectedRecipes = () =>
  Object.values(menu)
    .flat()
    .filter(Boolean)
    .map(id => recipes.find(recipe => recipeId(recipe) === String(id)))
    .filter(Boolean);
function debugShoppingList() {
  console.log("Menú guardado:", menu);
  console.log("Recetas cargadas:", recipes.length);
  console.log("Recetas seleccionadas:", selectedRecipes());

  console.log(
    "Ingredientes detectados:",
    selectedRecipes().map(recipe => ({
      nombre: recipe["Comida"],
      ingredientesOriginales: recipe["Ingredientes"],
      ingredientesSeparados: ingredients(recipe)
    }))
  );
}
const STORAGE_KEY = "mis-recetas-app-v1";
let checkedShoppingItems = new Set();
// Personas asignadas a cada receta dentro de cada franja del menú.
let recipeServings = {};

function servingKey(slot, id) {
  return `${slot}|${id}`;
}

function getServings(slot, id) {
  const value = Number(recipeServings[servingKey(slot, id)]);
  return Number.isFinite(value) && value >= 1 ? value : 2;
}

function saveState() {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify({
      recipes,
      menu,
      filters,
      activeView,
      weekOffset,
      checkedShoppingItems: [...checkedShoppingItems],
      recipeServings
    }));
  } catch (error) {
    console.warn("No se pudo guardar el estado localmente", error);
  }
}

function restoreState() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return false;

    const state = JSON.parse(raw);

    if (!Array.isArray(state.recipes) || !state.recipes.length) {
      return false;
    }

    recipes = state.recipes;
    menu = state.menu || {};
    filters = { ...filters, ...(state.filters || {}) };
    activeView = ["week", "recipes", "about", "shopping"]
      .includes(state.activeView) ? state.activeView : "week";
    weekOffset = Number.isFinite(state.weekOffset)
      ? state.weekOffset : 0;

    checkedShoppingItems = new Set(
      state.checkedShoppingItems || []
    );
    recipeServings = state.recipeServings || {};

    return true;
  } catch (error) {
    console.warn("No se pudo recuperar el estado guardado", error);
    return false;
  }
}

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


const GOOGLE_CLIENT_ID = import.meta.env.VITE_GOOGLE_CLIENT_ID;
const DRIVE_FILE_ID_KEY = "mis-recetas-drive-file-id";

const DRIVE_LAST_UPDATE_KEY = "mis-recetas-drive-last-update";

function getLastDriveUpdate() {
  const value = localStorage.getItem(DRIVE_LAST_UPDATE_KEY);

  if (!value) return "Todavía no has actualizado desde Drive";

  const date = new Date(value);

  return `Última actualización: ${date.toLocaleDateString("es-ES")} ${date.toLocaleTimeString("es-ES", {
    hour: "2-digit",
    minute: "2-digit"
  })}`;
}


let googleTokenClient = null;
let googleAccessToken = null;

function connectGoogleDrive() {
  if (!window.google?.accounts?.oauth2) {
    alert("Google todavía se está cargando. Espera unos segundos y vuelve a intentarlo.");
    return;
  }

  if (!GOOGLE_CLIENT_ID) {
    alert("Falta configurar VITE_GOOGLE_CLIENT_ID en .env.local.");
    return;
  }

  googleTokenClient = window.google.accounts.oauth2.initTokenClient({
    client_id: GOOGLE_CLIENT_ID,
    // Permiso de lectura para descargar el Excel privado.
    scope: "https://www.googleapis.com/auth/drive.readonly",
    callback: async (response) => {
      if (response.error) {
        console.error("Error de autorización de Google:", response);
        alert("No se ha podido autorizar el acceso a Google Drive.");
        return;
      }

      googleAccessToken = response.access_token;

      const savedFileId = localStorage.getItem(DRIVE_FILE_ID_KEY);
      const fileId = savedFileId || prompt(
        "Introduce el ID del archivo Excel de Google Drive:"
      );

      if (!fileId?.trim()) return;

      localStorage.setItem(DRIVE_FILE_ID_KEY, fileId.trim());

      await loadExcelFromDrive(fileId.trim());
    },
    error_callback: (error) => {
      console.error("Error al abrir la autorización de Google:", error);
      alert("No se ha podido abrir la ventana de autorización de Google.");
    }
  });

  googleTokenClient.requestAccessToken({
    prompt: localStorage.getItem(DRIVE_FILE_ID_KEY) ? "" : "consent"
  });

}


async function updateRecipesFromDrive() {
  const fileId = localStorage.getItem(DRIVE_FILE_ID_KEY);

  if (!fileId) {
    alert("Primero conecta Google Drive para seleccionar tu archivo.");
    connectGoogleDrive();
    return;
  }

  if (!googleAccessToken) {
    // Solicita autorización si no hay token disponible.
    connectGoogleDrive();
    return;
  }

  await loadExcelFromDrive(fileId);
}



async function loadExcelFromDrive(fileId) {
  if (!googleAccessToken) {
    alert("Conecta primero tu cuenta de Google.");
    return;
  }

  try {
    const exportUrl =
      `https://www.googleapis.com/drive/v3/files/${encodeURIComponent(fileId)}/export?mimeType=application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`;

    const response = await fetch(exportUrl, {
      method: "GET",
      cache: "no-store",
      headers: {
        Authorization: `Bearer ${googleAccessToken}`
      }
    });

    if (!response.ok) {
      if (response.status === 401) {
        googleAccessToken = null;
        connectGoogleDrive();
        return;
      }

      const details = await response.text();
      console.error("Error descargando el Excel:", details);

      throw new Error(
        `Google Drive ha devuelto el error ${response.status}: ${details}`
      );
    }

    const buffer = await response.arrayBuffer();
    const workbook = XLSX.read(buffer, { type: "array" });

    const firstSheet = workbook.Sheets[workbook.SheetNames[0]];

    if (!firstSheet) {
      throw new Error("No se ha encontrado ninguna hoja en Google Sheets.");
    }

    const rows = XLSX.utils.sheet_to_json(firstSheet, {
      defval: ""
    });

    if (!rows.length) {
      throw new Error("Google Sheets no contiene recetas.");
    }

    // ESTA ERA LA LÍNEA QUE FALTABA:
    // procesa las recetas nuevas y conserva el menú semanal.
    loadRecipesFromRows(rows);

    // Actualiza la fecha solo después de cargar correctamente.
    localStorage.setItem(
      DRIVE_LAST_UPDATE_KEY,
      new Date().toISOString()
    );

    // Actualiza la interfaz para mostrar la nueva fecha.
    render();

    alert(`Recetas actualizadas correctamente: ${recipes.length}`);
  } catch (error) {
    console.error("Error actualizando las recetas:", error);

    alert(
      error.message ||
      "No se ha podido actualizar las recetas desde Google Drive."
    );
  }
}



function loadRecipesFromRows(rows) {
  const nextRecipes = rows
    .map((row, index) => parseRow(row, index + 1))
    .filter(recipe => String(recipe["Comida"]).trim());

  if (!nextRecipes.length) {
    throw new Error("No se han encontrado recetas válidas en el Excel.");
  }

  // Actualiza las recetas, pero conserva el menú semanal.
  recipes = nextRecipes;

  // Mantiene las recetas seleccionadas siempre que sus IDs
  // sigan existiendo en el Excel.
  const validIds = new Set(recipes.map(recipe => recipeId(recipe)));

  for (const key of Object.keys(menu)) {
    menu[key] = (menu[key] || [])
      .filter(id => validIds.has(String(id)));

    if (!menu[key].length) {
      delete menu[key];
    }
  }

  saveState();
  render();
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

  const assigned = (menu[key] || [])
    .map(id => ({
      id: String(id),
      recipe: recipes.find(r => recipeId(r) === String(id))
    }))
    .filter(item => item.recipe);

  return `<section class="meal-slot">
    <div class="slot-heading">
      <span class="meal-dot dot-${mIndex}"></span>${meal}
    </div>

    ${assigned.map(({ id, recipe }, i) => {
      const servings = getServings(key, id);

      return `
        <div class="assigned-recipe">
          <div class="assigned-name">
            ${escapeHTML(recipe["Comida"])}
          </div>

          <div class="assigned-stats">
            ${num(recipe["Calorías"])} kcal ·
            ${num(recipe["Proteína (g)"])} g proteína
          </div>

          <label class="servings-control">
            <span class="servings-label">Personas</span>

            <div class="servings-stepper">
              <button
                type="button"
                class="servings-btn"
                data-servings-slot="${escapeHTML(key)}"
                data-servings-id="${escapeHTML(id)}"
                data-servings-change="-1"
                aria-label="Restar una persona"
              >−</button>

              <input
                type="number"
                min="1"
                max="99"
                step="1"
                value="${servings}"
                data-servings-slot="${escapeHTML(key)}"
                data-servings-id="${escapeHTML(id)}"
                aria-label="Número de personas"
              />

              <button
                type="button"
                class="servings-btn servings-btn-plus"
                data-servings-slot="${escapeHTML(key)}"
                data-servings-id="${escapeHTML(id)}"
                data-servings-change="1"
                aria-label="Añadir una persona"
              >+</button>
            </div>
          </label>

          <button
            class="remove-btn"
            title="Quitar del menú"
            aria-label="Quitar ${escapeHTML(recipe["Comida"])}"
            data-remove="${escapeHTML(key)}"
            data-index="${i}"
          >×</button>
        </div>
      `;
    }).join("")}

    <button class="add-slot" data-slot="${escapeHTML(key)}">
      ＋ Añadir receta
    </button>
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

function parseIngredient(raw, servings = 2) {
  const label = String(raw || "").trim().replace(/\s+/g, " ");
  if (!label) return null;

  // Ejemplos: "patata: 3 unidades", "nata: 15 ml".
  const match = label.match(
    /^(.+?)\s*:\s*(-?\d+(?:[.,]\d+)?)\s*([a-zA-ZáéíóúÁÉÍÓÚñÑ]+)?$/
  );

  if (!match) {
    return { label, amount: null, unit: "", count: 1 };
  }

  const name = match[1].trim();
  const perPerson = Number(match[2].replace(",", "."));
  let unit = (match[3] || "").toLowerCase();

  if (!name || !Number.isFinite(perPerson) || perPerson < 0) {
    return { label, amount: null, unit: "", count: 1 };
  }

  if (["kg", "kilo", "kilos", "kilogramo", "kilogramos"].includes(unit)) {
    return {
      label: name,
      amount: perPerson * 1000 * servings,
      unit: "g",
      count: 1
    };
  }

  if (["l", "lt", "litro", "litros"].includes(unit)) {
    return {
      label: name,
      amount: perPerson * 1000 * servings,
      unit: "ml",
      count: 1
    };
  }

  if (["gr", "grs", "gramo", "gramos"].includes(unit)) unit = "g";
  if (["mililitro", "mililitros", "cc"].includes(unit)) unit = "ml";

  if (["unidad", "unidades", "ud", "uds"].includes(unit)) {
    unit = "unidades";
  }

  return {
    label: name,
    amount: perPerson * servings,
    unit,
    count: 1
  };
}

function buildShoppingItems() {
  const counts = new Map();

  for (const [slot, ids] of Object.entries(menu)) {
    for (const id of ids || []) {
      const recipe = recipes.find(
        r => recipeId(r) === String(id)
      );

      if (!recipe) continue;

      const servings = getServings(slot, id);

      for (const raw of ingredients(recipe)) {
        const parsed = parseIngredient(raw, servings);
        if (!parsed) continue;

        const key = `${normalize(parsed.label)}|${parsed.unit}`;

        if (!counts.has(key)) {
          counts.set(key, {
            label: parsed.label,
            amount: parsed.amount,
            unit: parsed.unit,
            count: 1,
            hasAmount: parsed.amount !== null
          });
        } else {
          const item = counts.get(key);
          item.count++;

          if (parsed.amount !== null) {
            item.amount = (item.amount ?? 0) + parsed.amount;
            item.hasAmount = true;
          }
        }
      }
    }
  }

  return [...counts.entries()]
    .map(([key, item]) => {
      if (item.hasAmount && item.amount !== null) {
        if (item.unit === "g" && item.amount >= 1000) {
          item.amount /= 1000;
          item.unit = "kg";
        } else if (item.unit === "ml" && item.amount >= 1000) {
          item.amount /= 1000;
          item.unit = "l";
        }
      }

      return [key, item];
    })
    .sort((a, b) => a[1].label.localeCompare(b[1].label, "es"));
}

function formatQuantity(item) {
  if (!item.hasAmount || item.amount === null) return "";

  const amount = Number.isInteger(item.amount)
    ? String(item.amount)
    : String(Number(item.amount.toFixed(2)));

  return `${amount}${item.unit ? ` ${item.unit}` : ""}`;
}

function renderShoppingList() {
  const items = buildShoppingItems();
  const selected = selectedRecipes();

  const checkedCount = items.filter(
    ([key]) => checkedShoppingItems.has(key)
  ).length;

  return `
    <section class="panel shopping-panel">
      <div class="shopping-head">
        <div>
          <h2>Lista de la compra</h2>
          <p class="muted">
            ${items.length} ingredientes ·
            ${checkedCount} marcados como comprados
          </p>
        </div>

        <div class="shopping-actions">
          <button class="btn btn-quiet" id="uncheck-shopping">
            Desmarcar todos
          </button>
          <button class="btn btn-primary" id="export-shopping">
            Exportar lista
          </button>
        </div>
      </div>

      ${
        !selected.length
          ? `
            <div class="empty-state">
              <span>🧺</span>
              <h3>Tu lista está vacía</h3>
              <p>Añade recetas al menú semanal para ver sus ingredientes.</p>
              <button class="btn btn-primary" id="go-week">
                Ir al menú semanal
              </button>
            </div>
          `
          : items.length
            ? `
              <div class="shopping-list">
                ${items.map(([key, item]) => `
                  <label class="shopping-item ${
                    checkedShoppingItems.has(key) ? "is-checked" : ""
                  }">
                    <input
                      type="checkbox"
                      data-shopping-check="${escapeHTML(key)}"
                      ${checkedShoppingItems.has(key) ? "checked" : ""}
                    />

                    <span class="shopping-item-name">
                      ${escapeHTML(item.label)}
                    </span>

                    ${
                      item.hasAmount
                        ? `<strong class="ingredient-quantity">
                            ${escapeHTML(formatQuantity(item))}
                          </strong>`
                        : ""
                    }

                    <span class="ingredient-count">
                      ${item.count > 1
                        ? `En ${item.count} recetas`
                        : "1 receta"}
                    </span>
                  </label>
                `).join("")}
              </div>

              <p class="muted shopping-footnote">
                Las cantidades se calculan según las personas indicadas
                en cada receta. Usa formatos como «patata: 3 unidades»
                o «nata: 15 ml». Los ingredientes sin cantidad se
                agrupan por nombre.
              </p>
            `
            : `
              <div class="empty-state">
                <span>🥕</span>
                <h3>No se han detectado ingredientes</h3>
                <p>
                  Comprueba que la columna Ingredientes tenga valores
                  separados por comas o punto y coma.
                </p>
              </div>
            `
      }
    </section>
  `;
}

function render() {
  const categories = allUnique("Categoría"), seasons = allUnique("Temporada");
  app.innerHTML = `<div class="app-shell">
    <aside class="left-nav">
      <div class="brand"><div class="brand-mark">♧</div><div><strong>Mis Recetas</strong><span>Planificador semanal</span></div></div>
      <nav>
        <button class="nav-item ${activeView === "week" ? "active" : ""}" data-view="week">▦ <span>Menú semanal</span></button>
        <button class="nav-item ${activeView === "recipes" ? "active" : ""}" data-view="recipes">⌕ <span>Todas las recetas</span></button>
        <button class="nav-item ${activeView === "shopping" ? "active" : ""}" data-view="shopping">☷ <span>Lista de la compra</span></button>
        <button class="nav-item ${activeView === "about" ? "active" : ""}" data-view="about">ⓘ <span>Acerca de</span></button>
      </nav>
      <div class="upload-panel">
        <div class="upload-icon">⇧</div><strong>Sube tu archivo Excel</strong>
        <p>Arrastra un archivo .xlsx o .xls aquí, o selecciónalo desde tu equipo.</p>
        <label class="btn btn-primary upload-button" for="excel-file">Seleccionar archivo</label>
        <input id="excel-file" type="file" accept=".xlsx,.xls,.csv" />
        <button class="btn btn-primary upload-button" type="button" id="connect-drive">Conectar con Google Drive</button>
        <button class="btn btn-primary upload-button" type="button" id="update-recipes">↻ Actualizar recetas</button>
        <span class="upload-status">${recipes.length ? `${recipes.length} recetas cargadas` : "Puedes probar a cargar tu archivo."}</span>

<p id="drive-last-update">
  ${getLastDriveUpdate()}
</p>

      </div>
      <div class="nav-footer">Hecho para planificar mejor <span>🌱</span></div>
    </aside>
    <main class="main-content">
      <header class="topbar"><div><div class="eyebrow">ORGANIZA TU COCINA</div><h1>${activeView === "week"
  ? "Tu menú semanal"
  : activeView === "recipes"
    ? "Todas las recetas"
    : activeView === "shopping"
      ? "Lista de la compra"
      : "Acerca de Mis Recetas"}</h1><p>${activeView === "week"
  ? "Encuentra tus platos favoritos y organiza la semana a tu manera."
  : activeView === "recipes"
    ? "Busca y filtra las recetas de tu Excel."
    : activeView === "shopping"
      ? "Todos los ingredientes de las recetas elegidas para esta semana."
      : "Una forma sencilla de convertir tu Excel en un menú semanal."}</p></div>
        <div class="top-actions"><button class="btn btn-quiet" id="export-menu">↧ <span>Exportar menú</span></button><button class="btn btn-quiet" id="clear-menu">Vaciar menú</button></div>
      </header>
      ${activeView === "shopping" ? renderShoppingList() : activeView === "about" ? `<section class="panel about-panel"><h2>Cómo funciona</h2><ol><li>Sube un Excel con una sola hoja y las columnas de recetas.</li><li>Combina filtros por categoría, temporada, calorías, macros, notas o ingredientes.</li><li>En el menú semanal, pulsa «Añadir receta» en el día y la comida que quieras.</li><li>Exporta el menú a CSV para guardarlo o compartirlo.</li></ol><p>El archivo se procesa localmente en tu navegador. No se almacena en un servidor.</p></section>` : `
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
  saveState();
}
function resetFilters() {
  filters = { category:"", season:"", calories:"", protein:"", carbs:"", fat:"", notes:"", search:"" };
  render();
}
function bindEvents() {
  const connectDriveButton = app.querySelector("#connect-drive");
if (connectDriveButton) {
  connectDriveButton.addEventListener("click", connectGoogleDrive);
}

const updateRecipesButton = app.querySelector("#update-recipes");
if (updateRecipesButton) {
  updateRecipesButton.addEventListener("click", updateRecipesFromDrive);
}

  app.querySelectorAll("[data-view]").forEach(el => el.addEventListener("click", () => { activeView = el.dataset.view; selectedSlot = null; render(); }));
  app.querySelectorAll("[data-shopping-check]").forEach(el => {
  el.addEventListener("change", () => {
    const key = el.dataset.shoppingCheck;

    if (el.checked) {
      checkedShoppingItems.add(key);
    } else {
      checkedShoppingItems.delete(key);
    }

    saveState();
    render();
  });
});

app.querySelector("#uncheck-shopping")?.addEventListener("click", () => {
  checkedShoppingItems.clear();
  saveState();
  render();
});

app.querySelector("#go-week")?.addEventListener("click", () => {
  activeView = "week";
  render();
});

app.querySelector("#export-shopping")?.addEventListener("click", exportShoppingList);
app.querySelectorAll("[data-servings-slot]").forEach(input => {
  input.addEventListener("change", () => {
    const slot = input.dataset.servingsSlot;
    const id = input.dataset.servingsId;

    const value = Math.max(
      1,
      Math.min(99, Math.floor(Number(input.value) || 1))
    );

    recipeServings[servingKey(slot, id)] = value;

    saveState();
    render();
  });
});
app.querySelectorAll("[data-servings-change]").forEach(button => {
  button.addEventListener("click", () => {
    const slot = button.dataset.servingsSlot;
    const id = button.dataset.servingsId;
    const key = servingKey(slot, id);

    const input = app.querySelector(
      `input[data-servings-slot="${CSS.escape(slot)}"]` +
      `[data-servings-id="${CSS.escape(id)}"]`
    );

    if (!input) return;

    const current = Number(input.value) || 1;
    const change = Number(button.dataset.servingsChange);
    const next = Math.min(99, Math.max(1, current + change));

    recipeServings[key] = next;
    saveState();
    render();
  });
});
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

    if (menu[selectedSlot].includes(id)) {
      alert("Esta receta ya está añadida a esa franja del menú.");
      return;
    }

    menu[selectedSlot].push(id);
    selectedSlot = null;
    saveState();
    render();
  }));
  app.querySelectorAll("[data-remove]").forEach(btn => btn.addEventListener("click", () => {
    const key = btn.dataset.remove; const index = Number(btn.dataset.index);
    if (menu[key]) { menu[key].splice(index, 1); if (!menu[key].length) delete menu[key]; }
    saveState(); render();
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
    menu = {};checkedShoppingItems.clear();saveState(); render();
  });
  app.querySelector("#export-menu")?.addEventListener("click", exportMenu);
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
    /*const nextRecipes = rows.map((r, i) => parseRow(r, i + 1)).filter(r => String(r["Comida"]).trim());
    if (!nextRecipes.length) throw new Error("No se encontraron recetas con nombre en la columna «Comida».");
    recipes = nextRecipes; menu = {}; filters = { category:"", season:"", calories:"", protein:"", carbs:"", fat:"", notes:"", search:"" };
    activeView = "week"; selectedSlot = null; saveState(); render();*/
    loadRecipesFromRows(rows);
    if (missing.length) alert(`Se han cargado ${recipes.length} recetas. Algunas columnas opcionales no se han encontrado: ${missing.join(", ")}. Comprueba los encabezados si esos datos no aparecen.`);
  } catch (err) {
    alert(`No se ha podido leer el Excel: ${err.message || err}`);
  }
}
function exportShoppingList() {
  const items = buildShoppingItems();

  if (!items.length) {
    alert("No hay ingredientes que exportar. Añade recetas al menú semanal.");
    return;
  }

  const csv =
    "\ufeffIngrediente;Cantidad total;Aparece en recetas;Comprado\r\n" +
    items.map(([key, item]) =>
      `"${item.label.replace(/"/g, '""')}";"${formatQuantity(item)}";${item.count};${
        checkedShoppingItems.has(key) ? "Sí" : "No"
      }`
    ).join("\r\n");

  const url = URL.createObjectURL(
    new Blob([csv], { type: "text/csv;charset=utf-8;" })
  );

  const a = document.createElement("a");
  a.href = url;
  a.download = "lista-de-la-compra.csv";
  a.click();

  URL.revokeObjectURL(url);
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
if (!restoreState()) {
  recipes = sampleRecipes();
}

render();