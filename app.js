const CONFIG = window.MASOL_CONFIG || {};
const FLOW_URL = CONFIG.FLOW_URL || "";
const AUTO_SAVE_DELAY_MS = Number(CONFIG.AUTO_SAVE_DELAY_MS) || 1400;
const OPERATOR_KEY = "masol-consumos-planta-operator-v2";
const MONTHS = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"];

const GROUPS = [
  { id: "produccion", label: "Producción", detail: "Lecturas de proceso · lunes a domingo" },
  { id: "electrico", label: "Mantenimiento eléctrico", detail: "Contadores eléctricos · lunes a viernes" },
  { id: "servicios", label: "Mantenimiento y servicios", detail: "Contadores auxiliares" }
];

const READINGS = [
  reading(1, "produccion", "Contador General H2O", "B", "B", "Totalizador"),
  reading(2, "produccion", "Contador H2O SCI", "C", "C", "Totalizador"),
  reading(3, "produccion", "Contador agua de torre", "D", "D", "Totalizador"),
  reading(4, "produccion", "Contador Desc. Entrada / Contador Nuevo Entrada", "E", "E", "Totalizador"),
  reading(5, "produccion", "Contador Desc. Salida / Contador Nuevo Salida", "F", "F", "Totalizador"),
  reading(6, "produccion", "Contador Desmi Entrada", "G", "G", "Totalizador"),
  reading(7, "produccion", "Contador Desmi Salida", "H", "H", "Totalizador"),
  reading(8, "produccion", "Contador Reactor Biológico", "I", "I", "Totalizador"),
  reading(9, "produccion", "Contador Agua EDARI", "J", "J", "Totalizador"),
  reading(10, "produccion", "ERM analógico", "K", "L", "Totalizador"),
  reading(11, "produccion", "ERM volumen bruto Vmt", "L", "M", "Totalizador"),
  reading(12, "produccion", "ERM volumen corregido Vbt", "M", "N", "Totalizador"),
  reading(13, "produccion", "Cfv", "N", "O", "LecturaDirecta"),
  reading(14, "produccion", "Contador Gas Natural", "O", "P", "Totalizador"),
  reading(15, "produccion", "Contador 800H1", "P", "Q", "Totalizador"),
  reading(16, "produccion", "Contador Metanol", "Q", "R", "Totalizador"),
  reading(17, "produccion", "Nivel O2 (%)", "R", "S", "LecturaDirecta", { min: 0, max: 100 }),
  reading(40, "produccion", "Salida mar", "AO", "K", "Totalizador"),
  reading(43, "produccion", "Totalizador 390", "AR", "AT", "Totalizador"),
  reading(18, "electrico", "Lectura eléctrica Bio 90.LVB.05", "S", "T", "TotalizadorDiv10"),
  reading(19, "electrico", "P1", "T", "U", "Totalizador"),
  reading(20, "electrico", "P2", "U", "V", "Totalizador"),
  reading(21, "electrico", "P3", "V", "W", "Totalizador"),
  reading(22, "electrico", "P4", "W", "X", "Totalizador"),
  reading(23, "electrico", "P5", "X", "Y", "Totalizador"),
  reading(24, "electrico", "P6", "Y", "Z", "Totalizador"),
  reading(25, "electrico", "Trafo 1", "Z", "AB", "Totalizador"),
  reading(26, "electrico", "Trafo 2", "AA", "AC", "Totalizador"),
  reading(27, "electrico", "Trafo 3", "AB", "AD", "Totalizador"),
  reading(28, "servicios", "Contador Refino 90.LVB.06", "AC", "AE", "Totalizador"),
  reading(29, "servicios", "Contador B. Cubetos 90.LVB.07", "AD", "AG", "Totalizador"),
  reading(30, "servicios", "Contador Servicios 90.LVB.08", "AE", "AH", "Totalizador"),
  reading(31, "servicios", "Contador Edificios 90.LVB.09", "AF", "AI", "Totalizador"),
  reading(32, "servicios", "Contador Emergencia 90.LVB.10", "AG", "AJ", "Totalizador"),
  reading(33, "servicios", "Compresor A RUN", "AH", "AK", "Totalizador"),
  reading(34, "servicios", "Compresor B RUN", "AI", "AL", "Totalizador"),
  reading(35, "servicios", "Compresor C RUN", "AJ", "AM", "Totalizador"),
  reading(36, "servicios", "Compresor A marcha", "AK", "AN", "Totalizador"),
  reading(37, "servicios", "Compresor B marcha", "AL", "AO", "Totalizador"),
  reading(38, "servicios", "Compresor C marcha", "AM", "AP", "Totalizador"),
  reading(39, "servicios", "PALSHAL", "AN", "AR", "Totalizador"),
  reading(41, "servicios", "Ampliación 90.LVB.10 / Contador vehículos", "AP", "AS", "Totalizador"),
  reading(42, "servicios", "Contador EDARI", "AQ", "AQ", "Totalizador")
].sort((a, b) => a.Orden - b.Orden);

const form = document.querySelector("#readingsForm");
const sectionsContainer = document.querySelector("#readingsSections");
const dateInput = document.querySelector("#readingDate");
const operatorInput = document.querySelector("#operatorName");
const progressBar = document.querySelector("#progressBar");
const progressLabel = document.querySelector("#progressLabel");
const readyCount = document.querySelector("#readyCount");
const draftLabel = document.querySelector("#draftLabel");
const sheetDestination = document.querySelector("#sheetDestination");
const submitButton = document.querySelector("#submitButton");
const saveButton = document.querySelector("#saveButton");
const refreshButton = document.querySelector("#refreshButton");
const sharedState = document.querySelector("#sharedState");
const sharedTitle = document.querySelector("#sharedTitle");
const sharedDetail = document.querySelector("#sharedDetail");
const resultDialog = document.querySelector("#resultDialog");
const dialogIcon = document.querySelector("#dialogIcon");
const dialogTitle = document.querySelector("#dialogTitle");
const dialogMessage = document.querySelector("#dialogMessage");
const connectionState = document.querySelector("#connectionState");
const connectionLabel = document.querySelector("#connectionLabel");
const installButton = document.querySelector("#installButton");
const toast = document.querySelector("#toast");

let deferredInstallPrompt;
let saveTimer;
let saveInFlight = false;
let loadingDraft = false;
const dirtyColumns = new Set();

function reading(Orden, group, NombreCampo, ColumnaLectura, ColumnaTotales, TipoValidacion, limits = {}) {
  return { Orden, group, Bloque: group === "produccion" ? "PRODUCCION L a D" : "MANTENIMIENTO L a V", NombreCampo, ColumnaLectura, ColumnaTotales, TipoValidacion, ...limits };
}

function localIsoDate() {
  const now = new Date();
  return `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, "0")}-${String(now.getDate()).padStart(2, "0")}`;
}

function deviceType() {
  if (window.innerWidth < 620) return "Movil";
  if (window.innerWidth < 1024) return "Tableta";
  return "Escritorio";
}

function isConfigured() {
  return /^https:\/\//.test(FLOW_URL) && !FLOW_URL.includes("REEMPLAZAR_");
}

function renderSections() {
  sectionsContainer.innerHTML = GROUPS.map((group, index) => {
    const fields = READINGS.filter((item) => item.group === group.id);
    const fieldMarkup = fields.map((item) => `
      <label class="field reading-field" data-field="${item.ColumnaLectura}">
        <span class="label-line"><span>${item.NombreCampo}</span><span class="column-code" title="Columna Excel ${item.ColumnaLectura}">${item.ColumnaLectura}</span></span>
        <span class="reading-input-wrap">
          <input id="reading-${item.ColumnaLectura}" name="reading-${item.ColumnaLectura}" type="text" inputmode="decimal" autocomplete="off" aria-label="${item.NombreCampo}" data-column="${item.ColumnaLectura}" ${item.min !== undefined ? `data-min="${item.min}"` : "data-min=\"0\""} ${item.max !== undefined ? `data-max="${item.max}"` : ""} />
          <i class="input-status" data-lucide="check-circle-2" aria-hidden="true"></i>
        </span>
      </label>`).join("");
    return `<details class="reading-section" data-group="${group.id}" ${index === 0 ? "open" : ""}>
      <summary><span class="section-index">0${index + 1}</span><span class="section-title"><span><strong>${group.label}</strong><small>${group.detail}</small></span></span><span class="section-count" data-group-count="${group.id}">0 / ${fields.length}</span><i class="section-chevron" data-lucide="chevron-down" aria-hidden="true"></i></summary>
      <div class="fields-grid">${fieldMarkup}</div>
    </details>`;
  }).join("");
}

function parseNumber(value) {
  const normalized = String(value).trim().replace(",", ".");
  if (!normalized || !/^\d+(\.\d+)?$/.test(normalized)) return null;
  const number = Number(normalized);
  return Number.isFinite(number) ? number : null;
}

function validateInput(input) {
  const value = input.value.trim();
  input.classList.remove("invalid");
  input.removeAttribute("aria-invalid");
  if (!value) return true;
  const number = parseNumber(value);
  const min = input.dataset.min === undefined ? null : Number(input.dataset.min);
  const max = input.dataset.max === undefined ? null : Number(input.dataset.max);
  const valid = number !== null && (min === null || number >= min) && (max === null || number <= max);
  if (!valid) {
    input.classList.add("invalid");
    input.setAttribute("aria-invalid", "true");
  }
  return valid;
}

function updateProgress() {
  const inputs = [...document.querySelectorAll("[data-column]")];
  const count = inputs.filter((input) => input.value.trim()).length;
  inputs.forEach((input) => input.closest(".reading-field").classList.toggle("filled", Boolean(input.value.trim())));
  GROUPS.forEach((group) => {
    const groupInputs = [...document.querySelectorAll(`[data-group="${group.id}"] [data-column]`)];
    const groupFilled = groupInputs.filter((input) => input.value.trim()).length;
    document.querySelector(`[data-group-count="${group.id}"]`).textContent = `${groupFilled} / ${groupInputs.length}`;
  });
  progressBar.style.width = `${(count / READINGS.length) * 100}%`;
  progressLabel.textContent = `${count} de ${READINGS.length} lecturas`;
  readyCount.textContent = count === 1 ? "1 lectura disponible" : `${count} lecturas disponibles`;
  submitButton.disabled = count !== READINGS.length || loadingDraft || saveInFlight;
  submitButton.querySelector("span").textContent = count === READINGS.length ? "Enviar a aprobación" : `Faltan ${READINGS.length - count} lecturas`;
}

function updateDestination() {
  if (!dateInput.value) return void (sheetDestination.textContent = "Destino: hoja Mes Año");
  const [year, month] = dateInput.value.split("-");
  sheetDestination.textContent = `Destino: Datos-Planta-${MONTHS[Number(month) - 1]}-${year}.xlsx · hoja ${MONTHS[Number(month) - 1]} ${year}`;
}

function setSharedState(kind, title, detail) {
  sharedState.dataset.state = kind;
  sharedTitle.textContent = title;
  sharedDetail.textContent = detail;
}

async function callFlow(payload) {
  if (!isConfigured()) throw new Error("Falta configurar la URL del flujo en config.js.");
  const response = await fetch(FLOW_URL, { method: "POST", mode: "cors", headers: { "Content-Type": "application/json" }, body: JSON.stringify(payload) });
  const text = await response.text();
  let result = {};
  if (text) {
    try { result = JSON.parse(text); } catch { result = { mensaje: text }; }
  }
  if (!response.ok) throw new Error(result.mensaje || `El flujo respondió con el código ${response.status}.`);
  return result;
}

function normalizeServerReadings(result) {
  let readings = result.lecturas || result.readings || [];
  if (result.borradorJson) {
    try {
      const draft = typeof result.borradorJson === "string" ? JSON.parse(result.borradorJson) : result.borradorJson;
      readings = draft.lecturas || [];
    } catch {
      readings = [];
    }
  }
  if (typeof readings === "string") {
    try { readings = JSON.parse(readings); } catch { readings = []; }
  }
  return Array.isArray(readings) ? readings : [];
}

async function loadSharedDraft() {
  clearTimeout(saveTimer);
  dirtyColumns.clear();
  document.querySelectorAll("[data-column]").forEach((input) => {
    input.value = "";
    input.closest(".reading-field").classList.remove("shared");
    delete input.dataset.updatedBy;
  });
  updateProgress();
  if (!dateInput.value) return;
  if (!isConfigured()) {
    setSharedState("warning", "Flujo pendiente de configurar", "Añade la URL del flujo a config.js antes de registrar lecturas.");
    draftLabel.textContent = "Sin conexión con SharePoint";
    return;
  }
  loadingDraft = true;
  setFormDisabled(true);
  setSharedState("loading", "Cargando borrador compartido", "Consultando SharePoint para la fecha seleccionada.");
  try {
    const result = await callFlow({ accion: "cargar", fechaLectura: dateInput.value });
    const readings = normalizeServerReadings(result);
    readings.forEach((item) => {
      const column = item.ColumnaLectura || item.columna || item.Columna;
      const value = item.Valor ?? item.valor ?? item.Value;
      const input = document.querySelector(`[data-column="${column}"]`);
      if (!input || value === null || value === undefined || value === "") return;
      input.value = String(value);
      input.dataset.updatedBy = item.Operario || item.operario || "otro operario";
      input.closest(".reading-field").classList.add("shared");
    });
    const count = readings.length;
    setSharedState("ready", count ? "Borrador JSON actualizado" : "No hay lecturas guardadas", count ? `${count} lecturas recuperadas del archivo JSON. Puedes completar las restantes.` : "Empieza a introducir lecturas; Power Automate creará el archivo JSON.");
    draftLabel.textContent = count ? `${count} recuperadas de SharePoint` : "Sin cambios pendientes";
  } catch (error) {
    setSharedState("error", "No se pudo cargar SharePoint", error.message);
    showToast(error.message);
  } finally {
    loadingDraft = false;
    setFormDisabled(false);
    updateProgress();
  }
}

function queueSharedSave() {
  clearTimeout(saveTimer);
  draftLabel.textContent = "Cambios pendientes…";
  if (!operatorInput.value.trim()) return void (draftLabel.textContent = "Indica el operario para guardar");
  saveTimer = setTimeout(() => saveSharedDraft(false), AUTO_SAVE_DELAY_MS);
}

function collectReadings(onlyDirty = false) {
  const result = [];
  let firstInvalid = null;
  for (const item of READINGS) {
    if (onlyDirty && !dirtyColumns.has(item.ColumnaLectura)) continue;
    const input = document.querySelector(`[data-column="${item.ColumnaLectura}"]`);
    if (!input.value.trim()) continue;
    if (!validateInput(input)) { firstInvalid ||= input; continue; }
    result.push({ Orden: item.Orden, Bloque: item.Bloque, NombreCampo: item.NombreCampo, ColumnaLectura: item.ColumnaLectura, ColumnaTotales: item.ColumnaTotales, TipoValidacion: item.TipoValidacion, Valor: parseNumber(input.value) });
  }
  if (firstInvalid) {
    firstInvalid.closest("details").open = true;
    firstInvalid.focus();
    throw new Error("Revisa los valores marcados. Solo se admiten números positivos y decimales.");
  }
  return result;
}

async function saveSharedDraft(showConfirmation = true) {
  clearTimeout(saveTimer);
  if (saveInFlight) return false;
  if (!dateInput.value) throw new Error("Selecciona la fecha de lectura.");
  if (!operatorInput.value.trim()) { operatorInput.focus(); throw new Error("Introduce el nombre del operario antes de guardar."); }
  if (!dirtyColumns.size) {
    if (showConfirmation) showToast("No hay cambios nuevos que guardar.");
    return true;
  }
  const readings = collectReadings(false);
  saveInFlight = true;
  setButtonsBusy(true, "Guardando…");
  draftLabel.textContent = "Guardando en SharePoint…";
  try {
    const draft = buildDraft("Borrador", readings);
    const result = await callFlow({ accion: "guardar", fechaLectura: dateInput.value, borradorJson: JSON.stringify(draft) });
    dirtyColumns.clear();
    document.querySelectorAll("[data-column]").forEach((input) => {
      if (!input.value.trim()) return;
      input.closest(".reading-field").classList.add("shared");
      input.dataset.updatedBy = operatorInput.value.trim();
    });
    draftLabel.textContent = "Borrador guardado en SharePoint";
    setSharedState("ready", "Archivo JSON guardado", result.mensaje || `${readings.length} lecturas disponibles para el otro operario.`);
    if (showConfirmation) showToast("Borrador guardado en SharePoint.");
    return true;
  } catch (error) {
    draftLabel.textContent = "Error al guardar";
    setSharedState("error", "No se pudo guardar el borrador", error.message);
    showToast(error.message);
    return false;
  } finally {
    saveInFlight = false;
    setButtonsBusy(false);
    updateProgress();
  }
}

async function submitReadings(event) {
  event.preventDefault();
  try {
    if (!dateInput.value) throw new Error("Selecciona la fecha de lectura.");
    if (!operatorInput.value.trim()) throw new Error("Introduce el nombre del operario.");
    if (dirtyColumns.size && !(await saveSharedDraft(false))) return;
    const readings = collectReadings(false);
    if (readings.length !== READINGS.length) throw new Error(`Faltan ${READINGS.length - readings.length} lecturas antes del envío definitivo.`);
    setButtonsBusy(true, "Enviando…");
    const draft = buildDraft("Enviado", readings);
    const result = await callFlow({ accion: "enviar", fechaLectura: dateInput.value, borradorJson: JSON.stringify(draft) });
    showResult(true, "Lecturas enviadas", result.mensaje || "La solicitud se ha enviado a aprobación. El Excel se actualizará únicamente si se aprueba.");
    setSharedState("submitted", "Envío registrado", "Las lecturas quedan bloqueadas por el proceso de aprobación.");
  } catch (error) {
    showToast(error.message || "No se pudo enviar la solicitud.");
    if (error instanceof TypeError) showResult(false, "No se pudo conectar", "El navegador no pudo comunicarse con Power Automate. Revisa la URL, CORS y la conexión.");
  } finally {
    setButtonsBusy(false);
    updateProgress();
  }
}

function buildDraft(estado, readings) {
  return {
    version: 1,
    fechaLectura: dateInput.value,
    horaLectura: "08:00",
    estado,
    actualizadoPor: operatorInput.value.trim(),
    actualizadoEn: new Date().toISOString(),
    dispositivo: {
      tipo: deviceType(),
      anchoPantalla: window.innerWidth,
      altoPantalla: window.innerHeight,
      navegador: navigator.userAgent
    },
    lecturas: readings
  };
}

function setButtonsBusy(active, label) {
  saveButton.disabled = active;
  refreshButton.disabled = active;
  if (active) {
    submitButton.disabled = true;
    if (label) submitButton.querySelector("span").textContent = label;
  }
  saveButton.classList.toggle("loading", active);
}

function setFormDisabled(disabled) {
  document.querySelectorAll("[data-column]").forEach((input) => { input.disabled = disabled; });
  saveButton.disabled = disabled;
}

function showResult(success, title, message) {
  dialogIcon.classList.toggle("error", !success);
  dialogIcon.innerHTML = `<i data-lucide="${success ? "check" : "triangle-alert"}"></i>`;
  dialogTitle.textContent = title;
  dialogMessage.textContent = message;
  if (window.lucide) lucide.createIcons();
  resultDialog.showModal();
}

function showToast(message) {
  toast.textContent = message;
  toast.hidden = false;
  clearTimeout(showToast.timer);
  showToast.timer = setTimeout(() => { toast.hidden = true; }, 4500);
}

function updateConnectionState() {
  const online = navigator.onLine;
  connectionState.classList.toggle("offline", !online);
  connectionLabel.textContent = online ? "En línea" : "Sin conexión";
}

function restoreOperator() {
  try { operatorInput.value = localStorage.getItem(OPERATOR_KEY) || ""; } catch { /* El nombre se puede introducir manualmente. */ }
}

function initialize() {
  renderSections();
  restoreOperator();
  dateInput.value = localIsoDate();
  updateDestination();
  updateProgress();
  updateConnectionState();
  loadSharedDraft();

  form.addEventListener("input", (event) => {
    if (event.target.matches("[data-column]")) {
      validateInput(event.target);
      event.target.closest(".reading-field").classList.remove("shared");
      dirtyColumns.add(event.target.dataset.column);
      queueSharedSave();
    }
    if (event.target === operatorInput) {
      try { localStorage.setItem(OPERATOR_KEY, operatorInput.value.trim()); } catch { /* Sin efecto en las lecturas. */ }
      if (dirtyColumns.size) queueSharedSave();
    }
    updateProgress();
  });
  dateInput.addEventListener("change", () => { updateDestination(); loadSharedDraft(); });
  form.addEventListener("submit", submitReadings);
  saveButton.addEventListener("click", () => saveSharedDraft(true).catch((error) => showToast(error.message)));
  refreshButton.addEventListener("click", loadSharedDraft);
  document.querySelector("#closeDialog").addEventListener("click", () => resultDialog.close());
  window.addEventListener("online", updateConnectionState);
  window.addEventListener("offline", updateConnectionState);
  window.addEventListener("beforeinstallprompt", (event) => { event.preventDefault(); deferredInstallPrompt = event; installButton.hidden = false; });
  installButton.addEventListener("click", async () => { if (!deferredInstallPrompt) return; deferredInstallPrompt.prompt(); await deferredInstallPrompt.userChoice; deferredInstallPrompt = null; installButton.hidden = true; });
  if ("serviceWorker" in navigator) window.addEventListener("load", () => navigator.serviceWorker.register("./service-worker.js"));
  if (window.lucide) lucide.createIcons();
}

initialize();
