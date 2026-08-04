const FLOW_URL = "https://default65afa47b9e4e4ad28cfe30d4118f06.2e.environment.api.powerplatform.com:443/powerautomate/automations/direct/cu/07/workflows/2dead834c4b5407194a6caeddd6abd4c/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=SzGhuT8ppyDUsTHEHhP6VnoHkCzotgy6tzIs3k4sFEA";

const DRAFT_KEY = "masol-consumos-planta-draft-v1";
const OPERATOR_KEY = "masol-consumos-planta-operator-v1";
const MONTHS = [
  "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
  "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"
];

const GROUPS = [
  { id: "produccion", label: "Producción", detail: "Columnas B:R · lunes a domingo" },
  { id: "electrico", label: "Mantenimiento eléctrico", detail: "Columnas S:AB · lunes a viernes" },
  { id: "servicios", label: "Mantenimiento y servicios", detail: "Columnas AC:AR · contadores auxiliares" }
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
  reading(40, "servicios", "Salida mar", "AO", "K", "Totalizador"),
  reading(41, "servicios", "Ampliación 90.LVB.10 / Contador vehículos", "AP", "AS", "Totalizador"),
  reading(42, "servicios", "Contador EDARI", "AQ", "AQ", "Totalizador"),
  reading(43, "servicios", "Totalizador 390", "AR", "AT", "Totalizador")
];

const form = document.querySelector("#readingsForm");
const sectionsContainer = document.querySelector("#readingsSections");
const dateInput = document.querySelector("#readingDate");
const operatorInput = document.querySelector("#operatorName");
const operatorEmailInput = document.querySelector("#operatorEmail");
const progressBar = document.querySelector("#progressBar");
const progressLabel = document.querySelector("#progressLabel");
const readyCount = document.querySelector("#readyCount");
const draftLabel = document.querySelector("#draftLabel");
const sheetDestination = document.querySelector("#sheetDestination");
const submitButton = document.querySelector("#submitButton");
const resultDialog = document.querySelector("#resultDialog");
const dialogIcon = document.querySelector("#dialogIcon");
const dialogTitle = document.querySelector("#dialogTitle");
const dialogMessage = document.querySelector("#dialogMessage");
const connectionState = document.querySelector("#connectionState");
const connectionLabel = document.querySelector("#connectionLabel");
const installButton = document.querySelector("#installButton");
const toast = document.querySelector("#toast");

let deferredInstallPrompt;
let draftTimer;

function reading(Orden, group, NombreCampo, ColumnaLectura, ColumnaTotales, TipoValidacion, limits = {}) {
  return {
    Orden,
    group,
    Bloque: group === "produccion" ? "PRODUCCION L a D" : "MANTENIMIENTO L a V",
    NombreCampo,
    ColumnaLectura,
    ColumnaTotales,
    TipoValidacion,
    ...limits
  };
}

function localIsoDate() {
  const now = new Date();
  const year = now.getFullYear();
  const month = String(now.getMonth() + 1).padStart(2, "0");
  const day = String(now.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function deviceType() {
  const width = window.innerWidth;
  if (width < 620) return "Movil";
  if (width < 1024) return "Tableta";
  return "Escritorio";
}

function renderSections() {
  sectionsContainer.innerHTML = GROUPS.map((group, index) => {
    const fields = READINGS.filter((item) => item.group === group.id);
    const fieldMarkup = fields.map((item) => `
      <label class="field reading-field" data-field="${item.ColumnaLectura}">
        <span class="label-line">
          <span>${item.NombreCampo}</span>
          <span class="column-code" title="Columna Excel ${item.ColumnaLectura}">${item.ColumnaLectura}</span>
        </span>
        <span class="reading-input-wrap">
          <input
            id="reading-${item.ColumnaLectura}"
            name="reading-${item.ColumnaLectura}"
            type="text"
            inputmode="decimal"
            autocomplete="off"
            aria-label="${item.NombreCampo}"
            data-column="${item.ColumnaLectura}"
            ${item.min !== undefined ? `data-min="${item.min}"` : "data-min=\"0\""}
            ${item.max !== undefined ? `data-max="${item.max}"` : ""}
          />
          <i class="input-status" data-lucide="check-circle-2" aria-hidden="true"></i>
        </span>
      </label>
    `).join("");

    return `
      <details class="reading-section" data-group="${group.id}" ${index === 0 ? "open" : ""}>
        <summary>
          <span class="section-index">0${index + 1}</span>
          <span class="section-title">
            <span>
              <strong>${group.label}</strong>
              <small>${group.detail}</small>
            </span>
          </span>
          <span class="section-count" data-group-count="${group.id}">0 / ${fields.length}</span>
          <i class="section-chevron" data-lucide="chevron-down" aria-hidden="true"></i>
        </summary>
        <div class="fields-grid">${fieldMarkup}</div>
      </details>
    `;
  }).join("");
}

function parseNumber(value) {
  const normalized = String(value).trim().replace(",", ".");
  if (!normalized || !/^-?\d+(\.\d+)?$/.test(normalized)) return null;
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
  const filled = inputs.filter((input) => input.value.trim() !== "");
  const count = filled.length;
  const percentage = (count / READINGS.length) * 100;

  inputs.forEach((input) => {
    input.closest(".reading-field").classList.toggle("filled", input.value.trim() !== "");
  });

  GROUPS.forEach((group) => {
    const groupInputs = [...document.querySelectorAll(`[data-group="${group.id}"] [data-column]`)];
    const groupFilled = groupInputs.filter((input) => input.value.trim() !== "").length;
    document.querySelector(`[data-group-count="${group.id}"]`).textContent = `${groupFilled} / ${groupInputs.length}`;
  });

  progressBar.style.width = `${percentage}%`;
  progressLabel.textContent = `${count} de ${READINGS.length} lecturas`;
  readyCount.textContent = count === 1 ? "1 lectura preparada" : `${count} lecturas preparadas`;
}

function updateDestination() {
  if (!dateInput.value) {
    sheetDestination.textContent = "Destino: hoja Mes Año";
    return;
  }
  const [year, month] = dateInput.value.split("-");
  const monthName = MONTHS[Number(month) - 1];
  sheetDestination.textContent = `Destino: hoja Mes Año · ${monthName} ${year}`;
}

function queueDraftSave() {
  clearTimeout(draftTimer);
  draftLabel.textContent = "Guardando borrador…";
  draftTimer = setTimeout(saveDraft, 350);
}

function saveDraft() {
  const values = {};
  document.querySelectorAll("[data-column]").forEach((input) => {
    if (input.value.trim()) values[input.dataset.column] = input.value.trim();
  });

  const operator = {
    name: operatorInput.value.trim(),
    email: operatorEmailInput.value.trim().toLowerCase()
  };
  localStorage.setItem(OPERATOR_KEY, JSON.stringify(operator));
  localStorage.setItem(DRAFT_KEY, JSON.stringify({
    date: dateInput.value,
    operator: operator.name,
    operatorEmail: operator.email,
    values,
    savedAt: Date.now()
  }));
  draftLabel.textContent = "Borrador guardado";
}

function restoreDraft() {
  const raw = localStorage.getItem(DRAFT_KEY);
  const storedOperator = localStorage.getItem(OPERATOR_KEY);
  dateInput.value = localIsoDate();

  if (storedOperator) {
    try {
      const operator = JSON.parse(storedOperator);
      operatorInput.value = operator.name || "";
      operatorEmailInput.value = operator.email || "";
    } catch {
      localStorage.removeItem(OPERATOR_KEY);
    }
  }

  if (!raw) return;

  try {
    const draft = JSON.parse(raw);
    dateInput.value = draft.date || localIsoDate();
    operatorInput.value = draft.operator || operatorInput.value;
    operatorEmailInput.value = draft.operatorEmail || operatorEmailInput.value;
    Object.entries(draft.values || {}).forEach(([column, value]) => {
      const input = document.querySelector(`[data-column="${column}"]`);
      if (input) input.value = value;
    });
    draftLabel.textContent = "Borrador recuperado";
  } catch {
    localStorage.removeItem(DRAFT_KEY);
  }
}

function collectReadings() {
  const result = [];
  let firstInvalid = null;

  for (const item of READINGS) {
    const input = document.querySelector(`[data-column="${item.ColumnaLectura}"]`);
    if (!input.value.trim()) continue;

    if (!validateInput(input)) {
      firstInvalid ||= input;
      continue;
    }

    result.push({
      Orden: item.Orden,
      Bloque: item.Bloque,
      NombreCampo: item.NombreCampo,
      ColumnaLectura: item.ColumnaLectura,
      ColumnaTotales: item.ColumnaTotales,
      TipoValidacion: item.TipoValidacion,
      Valor: parseNumber(input.value)
    });
  }

  if (firstInvalid) {
    firstInvalid.closest("details").open = true;
    firstInvalid.focus();
    throw new Error("Revisa los valores marcados. Solo se admiten números positivos y decimales.");
  }

  return result;
}

async function submitReadings(event) {
  event.preventDefault();

  try {
    if (!dateInput.value) {
      dateInput.focus();
      throw new Error("Selecciona la fecha de lectura.");
    }
    if (!operatorInput.value.trim()) {
      operatorInput.focus();
      throw new Error("Introduce el nombre del operario.");
    }
    if (!operatorEmailInput.value.trim() || !operatorEmailInput.validity.valid) {
      operatorEmailInput.focus();
      throw new Error("Introduce un correo corporativo valido.");
    }

    const readings = collectReadings();
    if (readings.length === 0) {
      document.querySelector(".reading-section").open = true;
      document.querySelector("[data-column]").focus();
      throw new Error("Introduce al menos una lectura antes de enviar.");
    }

    setSubmitting(true);
    const payload = {
      fechaLectura: dateInput.value,
      horaLectura: "08:00",
      operario: operatorInput.value.trim(),
      operarioEmail: operatorEmailInput.value.trim().toLowerCase(),
      fechaHoraRegistro: new Date().toISOString(),
      dispositivo: {
        tipo: deviceType(),
        anchoPantalla: window.innerWidth,
        altoPantalla: window.innerHeight,
        navegador: navigator.userAgent
      },
      lecturasJson: JSON.stringify(readings)
    };

    const response = await fetch(FLOW_URL, {
      method: "POST",
      mode: "cors",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload)
    });

    const responseText = await response.text();
    let result = {};
    if (responseText) {
      try { result = JSON.parse(responseText); } catch { result = { mensaje: responseText }; }
    }

    if (!response.ok) {
      throw new Error(result.mensaje || `El flujo respondió con el código ${response.status}.`);
    }

    localStorage.removeItem(DRAFT_KEY);
    showResult(
      true,
      "Lecturas enviadas",
      result.mensaje || "La solicitud se ha enviado a aprobación."
    );
  } catch (error) {
    showToast(error.message || "No se pudo enviar la solicitud.");
    if (error instanceof TypeError) {
      showResult(false, "No se pudo conectar", "El navegador no pudo comunicarse con Power Automate. Revisa la conexión y la configuración CORS del flujo.");
    }
  } finally {
    setSubmitting(false);
  }
}

function setSubmitting(active) {
  submitButton.disabled = active;
  submitButton.querySelector("span").textContent = active ? "Enviando…" : "Enviar a aprobación";
  submitButton.querySelector("svg")?.setAttribute("data-lucide", active ? "loader-circle" : "send");
  submitButton.classList.toggle("loading", active);
  if (window.lucide) lucide.createIcons();
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

function initialize() {
  renderSections();
  restoreDraft();
  updateDestination();
  updateProgress();
  updateConnectionState();

  form.addEventListener("input", (event) => {
    if (event.target.matches("[data-column]")) validateInput(event.target);
    updateProgress();
    queueDraftSave();
  });
  dateInput.addEventListener("change", updateDestination);
  form.addEventListener("submit", submitReadings);
  document.querySelector("#closeDialog").addEventListener("click", () => resultDialog.close());
  window.addEventListener("online", updateConnectionState);
  window.addEventListener("offline", updateConnectionState);

  window.addEventListener("beforeinstallprompt", (event) => {
    event.preventDefault();
    deferredInstallPrompt = event;
    installButton.hidden = false;
  });
  installButton.addEventListener("click", async () => {
    if (!deferredInstallPrompt) return;
    deferredInstallPrompt.prompt();
    await deferredInstallPrompt.userChoice;
    deferredInstallPrompt = null;
    installButton.hidden = true;
  });

  if ("serviceWorker" in navigator) {
    window.addEventListener("load", () => navigator.serviceWorker.register("./service-worker.js"));
  }

  if (window.lucide) lucide.createIcons();
}

initialize();
