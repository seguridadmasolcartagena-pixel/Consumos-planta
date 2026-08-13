const CONFIG = window.MASOL_CONFIG || {};
const FLOW_URL = CONFIG.FLOW_URL || "";
const AUTO_SAVE_DELAY_MS = Number(CONFIG.AUTO_SAVE_DELAY_MS) || 1400;
const MAX_PREVIOUS_LOOKBACK_DAYS = 10;
const PHOTO_MAX_EDGE = 1600;
const PHOTO_JPEG_QUALITY = 0.82;
const OPERATOR_KEY = "masol-consumos-planta-operator-v2";
const MONTHS = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"];
const OPTIONAL_COLUMNS = new Set(["E", "F", "I", "N", "O", "AR"]);

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
const REQUIRED_READINGS = READINGS.filter((item) => !OPTIONAL_COLUMNS.has(item.ColumnaLectura));

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
const photoInput = document.querySelector("#counterPhotoInput");
const photoDialog = document.querySelector("#photoDialog");
const photoPreview = document.querySelector("#photoPreview");
const photoTitle = document.querySelector("#photoTitle");
const photoStatus = document.querySelector("#photoStatus");
const recognizedValue = document.querySelector("#recognizedValue");
const recognizedText = document.querySelector("#recognizedText");
const useRecognizedButton = document.querySelector("#useRecognizedButton");
const retryPhotoButton = document.querySelector("#retryPhotoButton");
const cancelPhotoButton = document.querySelector("#cancelPhotoButton");

let deferredInstallPrompt;
let saveTimer;
let saveInFlight = false;
let loadingDraft = false;
const dirtyColumns = new Set();
const previousReadings = new Map();
let previousReadingDate = "";
let activePhotoColumn = "";
let activeCameraButton = null;
let ocrWorkerPromise = null;

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
      <div class="field reading-field" data-field="${item.ColumnaLectura}">
        <label class="label-line" for="reading-${item.ColumnaLectura}"><span>${item.NombreCampo}${OPTIONAL_COLUMNS.has(item.ColumnaLectura) ? " · Opcional" : ""}</span><span class="column-code" title="Columna Excel ${item.ColumnaLectura}">${item.ColumnaLectura}</span></label>
        <span class="reading-input-wrap">
          <span class="input-shell">
            <input id="reading-${item.ColumnaLectura}" name="reading-${item.ColumnaLectura}" type="text" inputmode="decimal" autocomplete="off" aria-label="${item.NombreCampo}" data-column="${item.ColumnaLectura}" ${item.min !== undefined ? `data-min="${item.min}"` : "data-min=\"0\""} ${item.max !== undefined ? `data-max="${item.max}"` : ""} />
            <i class="input-status" data-lucide="check-circle-2" aria-hidden="true"></i>
          </span>
          <button class="camera-button" type="button" data-camera-column="${item.ColumnaLectura}" title="Fotografiar ${item.NombreCampo}" aria-label="Fotografiar ${item.NombreCampo}"><i data-lucide="camera"></i></button>
        </span>
        <small class="reading-author" data-author-for="${item.ColumnaLectura}" hidden></small>
      </div>`).join("");
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
  delete input.dataset.validationMessage;
  if (!value) return true;
  const number = parseNumber(value);
  const min = input.dataset.min === undefined ? null : Number(input.dataset.min);
  const max = input.dataset.max === undefined ? null : Number(input.dataset.max);
  const readingDefinition = READINGS.find((item) => item.ColumnaLectura === input.dataset.column);
  const previous = previousReadings.get(input.dataset.column);
  const validatesTotalizer = readingDefinition && readingDefinition.TipoValidacion !== "LecturaDirecta";
  const valid = number !== null
    && (min === null || number >= min)
    && (max === null || number <= max)
    && (!validatesTotalizer || previous === undefined || number >= previous);
  if (!valid) {
    input.classList.add("invalid");
    input.setAttribute("aria-invalid", "true");
    input.dataset.validationMessage = validatesTotalizer && previous !== undefined && number !== null && number < previous
      ? `${readingDefinition.NombreCampo}: ${number} no puede ser menor que la lectura anterior ${previous} (${formatSpanishDate(previousReadingDate)}).`
      : "Solo se admiten números positivos dentro del rango permitido.";
  }
  return valid;
}

function updateProgress() {
  const inputs = [...document.querySelectorAll("[data-column]")];
  const requiredInputs = inputs.filter((input) => !OPTIONAL_COLUMNS.has(input.dataset.column));
  const requiredCount = requiredInputs.filter((input) => input.value.trim()).length;
  const optionalCount = inputs.filter((input) => OPTIONAL_COLUMNS.has(input.dataset.column) && input.value.trim()).length;
  inputs.forEach((input) => input.closest(".reading-field").classList.toggle("filled", Boolean(input.value.trim())));
  GROUPS.forEach((group) => {
    const groupInputs = [...document.querySelectorAll(`[data-group="${group.id}"] [data-column]`)];
    const groupRequired = groupInputs.filter((input) => !OPTIONAL_COLUMNS.has(input.dataset.column));
    const groupFilled = groupRequired.filter((input) => input.value.trim()).length;
    document.querySelector(`[data-group-count="${group.id}"]`).textContent = `${groupFilled} / ${groupRequired.length} obligatorias`;
  });
  progressBar.style.width = `${(requiredCount / REQUIRED_READINGS.length) * 100}%`;
  progressLabel.textContent = `${requiredCount} de ${REQUIRED_READINGS.length} obligatorias · ${optionalCount} opcionales`;
  readyCount.textContent = `${requiredCount} lecturas obligatorias preparadas`;
  submitButton.disabled = requiredCount !== REQUIRED_READINGS.length || loadingDraft || saveInFlight;
  submitButton.querySelector("span").textContent = requiredCount === REQUIRED_READINGS.length ? "Enviar lecturas" : `Faltan ${REQUIRED_READINGS.length - requiredCount} obligatorias`;
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

function loadImage(file) {
  return new Promise((resolve, reject) => {
    const url = URL.createObjectURL(file);
    const image = new Image();
    image.onload = () => { URL.revokeObjectURL(url); resolve(image); };
    image.onerror = () => { URL.revokeObjectURL(url); reject(new Error("No se pudo abrir la fotografía.")); };
    image.src = url;
  });
}

function loadDataUrlImage(dataUrl) {
  return new Promise((resolve, reject) => {
    const image = new Image();
    image.onload = () => resolve(image);
    image.onerror = () => reject(new Error("No se pudo preparar la imagen para el reconocimiento."));
    image.src = dataUrl;
  });
}

function canvasToDataUrl(canvas) {
  return new Promise((resolve, reject) => {
    canvas.toBlob((blob) => {
      if (!blob) return reject(new Error("No se pudo preparar la fotografía."));
      const reader = new FileReader();
      reader.onload = () => resolve(String(reader.result));
      reader.onerror = () => reject(new Error("No se pudo leer la fotografía."));
      reader.readAsDataURL(blob);
    }, "image/jpeg", PHOTO_JPEG_QUALITY);
  });
}

async function preparePhoto(file) {
  if (!file || !file.type.startsWith("image/")) throw new Error("Selecciona una fotografía válida.");
  const image = await loadImage(file);
  const scale = Math.min(1, PHOTO_MAX_EDGE / Math.max(image.naturalWidth, image.naturalHeight));
  const width = Math.max(1, Math.round(image.naturalWidth * scale));
  const height = Math.max(1, Math.round(image.naturalHeight * scale));
  const canvas = document.createElement("canvas");
  canvas.width = width;
  canvas.height = height;
  const context = canvas.getContext("2d");
  context.fillStyle = "#ffffff";
  context.fillRect(0, 0, width, height);
  context.drawImage(image, 0, 0, width, height);
  const dataUrl = await canvasToDataUrl(canvas);
  return { dataUrl, base64: dataUrl.split(",")[1] };
}

function otsuThreshold(grays) {
  const histogram = new Uint32Array(256);
  for (const gray of grays) histogram[gray] += 1;
  const total = grays.length;
  let weightedTotal = 0;
  for (let value = 0; value < 256; value += 1) weightedTotal += value * histogram[value];
  let backgroundWeight = 0;
  let backgroundSum = 0;
  let bestVariance = -1;
  let threshold = 128;
  for (let value = 0; value < 256; value += 1) {
    backgroundWeight += histogram[value];
    if (!backgroundWeight) continue;
    const foregroundWeight = total - backgroundWeight;
    if (!foregroundWeight) break;
    backgroundSum += value * histogram[value];
    const backgroundMean = backgroundSum / backgroundWeight;
    const foregroundMean = (weightedTotal - backgroundSum) / foregroundWeight;
    const variance = backgroundWeight * foregroundWeight * (backgroundMean - foregroundMean) ** 2;
    if (variance > bestVariance) { bestVariance = variance; threshold = value; }
  }
  return threshold;
}

async function prepareOcrVariants(dataUrl) {
  const image = await loadDataUrlImage(dataUrl);
  const marginX = Math.round(image.naturalWidth * 0.04);
  const marginY = Math.round(image.naturalHeight * 0.04);
  const sourceWidth = Math.max(1, image.naturalWidth - marginX * 2);
  const sourceHeight = Math.max(1, image.naturalHeight - marginY * 2);
  const scale = Math.max(1, Math.min(2.5, 2400 / sourceWidth));
  const width = Math.max(1, Math.round(sourceWidth * scale));
  const height = Math.max(1, Math.round(sourceHeight * scale));
  const canvas = document.createElement("canvas");
  canvas.width = width;
  canvas.height = height;
  const context = canvas.getContext("2d", { willReadFrequently: true });
  context.drawImage(image, marginX, marginY, sourceWidth, sourceHeight, 0, 0, width, height);
  const imageData = context.getImageData(0, 0, width, height);
  const grays = new Uint8Array(width * height);
  for (let pixel = 0, offset = 0; offset < imageData.data.length; pixel += 1, offset += 4) {
    const gray = Math.round(imageData.data[offset] * 0.299 + imageData.data[offset + 1] * 0.587 + imageData.data[offset + 2] * 0.114);
    grays[pixel] = gray;
    const contrasted = Math.max(0, Math.min(255, (gray - 128) * 1.65 + 128));
    imageData.data[offset] = contrasted;
    imageData.data[offset + 1] = contrasted;
    imageData.data[offset + 2] = contrasted;
  }
  context.putImageData(imageData, 0, 0);
  const grayscale = canvas.toDataURL("image/png");
  const threshold = otsuThreshold(grays);
  for (let pixel = 0, offset = 0; offset < imageData.data.length; pixel += 1, offset += 4) {
    const binary = grays[pixel] > threshold ? 255 : 0;
    imageData.data[offset] = binary;
    imageData.data[offset + 1] = binary;
    imageData.data[offset + 2] = binary;
  }
  context.putImageData(imageData, 0, 0);
  const binary = canvas.toDataURL("image/png");
  return [binary, grayscale, dataUrl];
}

function normalizeOcrNumber(value) {
  let text = String(value ?? "").trim().replace(/\s/g, "");
  if (!text || !/^\d[\d.,]*$/.test(text)) return "";
  const lastComma = text.lastIndexOf(",");
  const lastDot = text.lastIndexOf(".");
  if (lastComma >= 0 && lastDot >= 0) {
    const decimalSeparator = lastComma > lastDot ? "," : ".";
    const thousandsSeparator = decimalSeparator === "," ? /\./g : /,/g;
    text = text.replace(thousandsSeparator, "").replace(decimalSeparator, ".");
  } else if (lastComma >= 0) {
    text = text.replace(/\./g, "").replace(",", ".");
  } else if ((text.match(/\./g) || []).length > 1) {
    const parts = text.split(".");
    text = `${parts.slice(0, -1).join("")}.${parts.at(-1)}`;
  }
  return parseNumber(text) === null ? "" : text;
}

function numericCandidates(text) {
  const matches = String(text || "").match(/\d[\d\t .,]*\d|\d/g) || [];
  return [...new Set(matches.map(normalizeOcrNumber).filter(Boolean))];
}

function chooseRecognizedValue(result, column) {
  const directValue = result.valor ?? result.Valor ?? result.lectura ?? result.numero;
  const normalizedDirect = normalizeOcrNumber(directValue);
  if (normalizedDirect) return normalizedDirect;
  const text = result.textoDetectado ?? result.texto ?? result.fullText ?? result.text ?? "";
  const input = document.querySelector(`[data-column="${column}"]`);
  const min = Number(input?.dataset.min ?? 0);
  const max = input?.dataset.max === undefined ? Number.POSITIVE_INFINITY : Number(input.dataset.max);
  const previous = previousReadings.get(column);
  const readingDefinition = READINGS.find((item) => item.ColumnaLectura === column);
  const totalizer = readingDefinition?.TipoValidacion !== "LecturaDirecta";
  const candidates = numericCandidates(text)
    .map((candidate) => ({ candidate, number: parseNumber(candidate) }))
    .filter((item) => item.number !== null && item.number >= min && item.number <= max)
    .filter((item) => !totalizer || previous === undefined || item.number >= previous);
  if (!candidates.length) return "";
  if (totalizer && previous !== undefined) {
    candidates.sort((a, b) => (a.number - previous) - (b.number - previous));
  } else {
    candidates.sort((a, b) => b.candidate.replace(".", "").length - a.candidate.replace(".", "").length);
  }
  return candidates[0].candidate;
}

function setRecognitionBusy(busy) {
  if (activeCameraButton) activeCameraButton.disabled = busy;
  retryPhotoButton.disabled = busy;
  useRecognizedButton.disabled = busy || parseNumber(recognizedValue.value) === null;
  photoDialog.classList.toggle("recognizing", busy);
}

function ocrProgressMessage(message) {
  const percent = Number.isFinite(message.progress) ? ` ${Math.round(message.progress * 100)} %` : "";
  const labels = {
    "loading tesseract core": "Cargando el motor OCR",
    "initializing tesseract": "Iniciando el motor OCR",
    "loading language traineddata": "Cargando el reconocimiento de dígitos",
    "initializing api": "Preparando el reconocimiento",
    "recognizing text": "Reconociendo la lectura"
  };
  return `${labels[message.status] || "Procesando la fotografía"}${percent}…`;
}

async function getOcrWorker() {
  if (!window.Tesseract) throw new Error("No se pudo cargar el motor OCR local. Comprueba la conexión y vuelve a abrir la aplicación.");
  if (!ocrWorkerPromise) {
    ocrWorkerPromise = window.Tesseract.createWorker("eng", 1, {
      logger: (message) => { if (photoDialog.open) photoStatus.textContent = ocrProgressMessage(message); }
    }).then(async (worker) => {
      await worker.setParameters({
        tessedit_char_whitelist: "0123456789.,",
        tessedit_pageseg_mode: window.Tesseract.PSM.SINGLE_LINE,
        preserve_interword_spaces: "1"
      });
      return worker;
    }).catch((error) => {
      ocrWorkerPromise = null;
      throw error;
    });
  }
  return ocrWorkerPromise;
}

async function recognizePhoto(file) {
  const readingDefinition = READINGS.find((item) => item.ColumnaLectura === activePhotoColumn);
  if (!readingDefinition) return;
  photoTitle.textContent = readingDefinition.NombreCampo;
  photoStatus.textContent = "Preparando fotografía…";
  recognizedValue.value = "";
  recognizedText.textContent = "";
  photoPreview.removeAttribute("src");
  if (!photoDialog.open) photoDialog.showModal();
  setRecognitionBusy(true);
  try {
    const photo = await preparePhoto(file);
    photoPreview.src = photo.dataUrl;
    photoStatus.textContent = "Cargando el reconocimiento local…";
    const worker = await getOcrWorker();
    const variants = await prepareOcrVariants(photo.dataUrl);
    const attempts = [];
    for (let index = 0; index < variants.length; index += 1) {
      photoStatus.textContent = `Analizando los dígitos · intento ${index + 1} de ${variants.length}…`;
      const recognition = await worker.recognize(variants[index], { rotateAuto: index === variants.length - 1 });
      const text = recognition.data.text || "";
      const candidates = numericCandidates(text);
      const longest = candidates.reduce((length, candidate) => Math.max(length, candidate.replace(".", "").length), 0);
      attempts.push({ text, confidence: Number(recognition.data.confidence) || 0, score: longest * 12 + (Number(recognition.data.confidence) || 0) });
      if (longest >= 4 && Number(recognition.data.confidence) >= 75) break;
    }
    attempts.sort((a, b) => b.score - a.score);
    const text = attempts[0]?.text || "";
    const result = { textoDetectado: text };
    const value = chooseRecognizedValue(result, activePhotoColumn);
    recognizedText.textContent = text || "El reconocimiento local no encontró dígitos legibles.";
    recognizedValue.value = value;
    photoStatus.textContent = value
      ? "Lectura detectada. Comprueba el número antes de usarlo."
      : "No se ha podido elegir una lectura. Escríbela manualmente o repite la foto.";
  } catch (error) {
    photoStatus.textContent = error.message || "No se pudo reconocer la fotografía.";
    recognizedText.textContent = "Puedes cerrar esta ventana y continuar introduciendo la lectura manualmente.";
  } finally {
    setRecognitionBusy(false);
  }
}

function startPhotoCapture(column, button) {
  activePhotoColumn = column;
  activeCameraButton = button;
  photoInput.value = "";
  photoInput.click();
}

function applyRecognizedValue() {
  const input = document.querySelector(`[data-column="${activePhotoColumn}"]`);
  const value = normalizeOcrNumber(recognizedValue.value);
  if (!input || !value) return void showToast("Revisa la lectura reconocida.");
  input.value = value;
  input.dispatchEvent(new Event("input", { bubbles: true }));
  input.focus();
  photoDialog.close();
  showToast("Lectura añadida desde la fotografía. Comprueba el valor antes de enviar.");
}

function normalizeServerReadings(result) {
  let readings = result.lecturas || result.readings || [];
  if (result.borradorJson) {
    try {
      const draft = typeof result.borradorJson === "string" ? JSON.parse(result.borradorJson) : result.borradorJson;
      readings = (draft.lecturas || []).map((item) => ({
        ...item,
        Operario: item.Operario || item.operario || draft.actualizadoPor || ""
      }));
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
      updateReadingAuthor(input);
    });
    await loadPreviousReadings();
    document.querySelectorAll("[data-column]").forEach((input) => validateInput(input));
    const count = readings.length;
    const previousDetail = previousReadingDate ? ` Referencia anterior: ${formatSpanishDate(previousReadingDate)}.` : "";
    setSharedState("ready", count ? "Borrador JSON actualizado" : "No hay lecturas guardadas", (count ? `${count} lecturas recuperadas del archivo JSON. Puedes completar las restantes.` : "Empieza a introducir lecturas; Power Automate creará el archivo JSON.") + previousDetail);
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

async function loadPreviousReadings() {
  previousReadings.clear();
  previousReadingDate = "";
  document.querySelectorAll("[data-column]").forEach((input) => {
    delete input.dataset.previousValue;
    input.removeAttribute("title");
  });

  for (let daysBack = 1; daysBack <= MAX_PREVIOUS_LOOKBACK_DAYS; daysBack += 1) {
    const candidateDate = shiftIsoDate(dateInput.value, -daysBack);
    let readings = [];
    for (const origen of ["procesados", "borradores"]) {
      const result = await callFlow({ accion: "cargar", fechaLectura: candidateDate, origen });
      readings = normalizeServerReadings(result);
      if (readings.length) break;
    }
    if (!readings.length) continue;

    previousReadingDate = candidateDate;
    readings.forEach((reading) => {
      const column = reading.ColumnaLectura || reading.columna || reading.Columna;
      const value = reading.Valor ?? reading.valor ?? reading.Value;
      if (!column || typeof Number(value) !== "number" || !Number.isFinite(Number(value))) return;
      previousReadings.set(column, Number(value));
      const input = document.querySelector(`[data-column="${column}"]`);
      if (input) {
        input.dataset.previousValue = String(value);
        input.title = `Lectura anterior: ${value} (${formatSpanishDate(candidateDate)})`;
      }
    });
    break;
  }
}

function shiftIsoDate(value, days) {
  const date = new Date(`${value}T00:00:00Z`);
  date.setUTCDate(date.getUTCDate() + days);
  return date.toISOString().slice(0, 10);
}

function formatSpanishDate(value) {
  const [year, month, day] = value.split("-");
  return `${day}/${month}/${year}`;
}

function updateReadingAuthor(input) {
  const author = input.dataset.updatedBy || "";
  const label = document.querySelector(`[data-author-for="${input.dataset.column}"]`);
  if (!label) return;
  label.hidden = !author || !input.value.trim();
  label.textContent = label.hidden ? "" : `Introducida por: ${author}`;
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
    const operario = dirtyColumns.has(item.ColumnaLectura)
      ? operatorInput.value.trim()
      : input.dataset.updatedBy || operatorInput.value.trim();
    result.push({ Orden: item.Orden, Bloque: item.Bloque, NombreCampo: item.NombreCampo, ColumnaLectura: item.ColumnaLectura, ColumnaTotales: item.ColumnaTotales, TipoValidacion: item.TipoValidacion, Valor: parseNumber(input.value), Operario: operario });
  }
  if (firstInvalid) {
    firstInvalid.closest("details").open = true;
    firstInvalid.focus();
    throw new Error(firstInvalid.dataset.validationMessage || "Revisa los valores marcados.");
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
    const changedColumns = new Set(dirtyColumns);
    const draft = buildDraft("Borrador", readings);
    const result = await callFlow({ accion: "guardar", fechaLectura: dateInput.value, borradorJson: JSON.stringify(draft) });
    dirtyColumns.clear();
    document.querySelectorAll("[data-column]").forEach((input) => {
      if (!input.value.trim()) return;
      input.closest(".reading-field").classList.add("shared");
      if (changedColumns.has(input.dataset.column)) input.dataset.updatedBy = operatorInput.value.trim();
      updateReadingAuthor(input);
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
    const receivedColumns = new Set(readings.map((reading) => reading.ColumnaLectura));
    const missingRequired = REQUIRED_READINGS.filter((reading) => !receivedColumns.has(reading.ColumnaLectura));
    if (missingRequired.length) throw new Error(`Faltan ${missingRequired.length} lecturas obligatorias antes del envío definitivo.`);
    setButtonsBusy(true, "Enviando…");
    const draft = buildDraft("Enviado", readings);
    const result = await callFlow({ accion: "enviar", fechaLectura: dateInput.value, borradorJson: JSON.stringify(draft) });
    showResult(true, "Lecturas enviadas", result.mensaje || "Power Automate ha recibido las lecturas y está actualizando el Excel mensual.");
    setSharedState("submitted", "Procesamiento iniciado", "Power Automate está escribiendo las lecturas en el Excel mensual.");
  } catch (error) {
    showToast(error.message || "No se pudo enviar la solicitud.");
    if (error instanceof TypeError) showResult(false, "No se pudo conectar", "El navegador no pudo comunicarse con Power Automate. Revisa la URL, CORS y la conexión.");
  } finally {
    setButtonsBusy(false);
    updateProgress();
  }
}

function buildDraft(estado, readings) {
  const operarios = [...new Set(readings.map((reading) => reading.Operario).filter(Boolean))];
  return {
    version: 1,
    fechaLectura: dateInput.value,
    horaLectura: "08:00",
    estado,
    actualizadoPor: operatorInput.value.trim(),
    operarios,
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
  document.querySelectorAll("[data-camera-column]").forEach((button) => { button.disabled = disabled; });
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
      event.target.dataset.updatedBy = operatorInput.value.trim();
      updateReadingAuthor(event.target);
      dirtyColumns.add(event.target.dataset.column);
      queueSharedSave();
    }
    if (event.target === operatorInput) {
      try { localStorage.setItem(OPERATOR_KEY, operatorInput.value.trim()); } catch { /* Sin efecto en las lecturas. */ }
      if (dirtyColumns.size) queueSharedSave();
    }
    updateProgress();
  });
  form.addEventListener("focusout", (event) => {
    if (!event.target.matches("[data-column]")) return;
    if (!validateInput(event.target) && event.target.dataset.validationMessage) {
      showToast(event.target.dataset.validationMessage);
    }
  });
  sectionsContainer.addEventListener("click", (event) => {
    const button = event.target.closest("[data-camera-column]");
    if (button) startPhotoCapture(button.dataset.cameraColumn, button);
  });
  photoInput.addEventListener("change", () => {
    const [file] = photoInput.files || [];
    if (file) recognizePhoto(file);
  });
  recognizedValue.addEventListener("input", () => {
    useRecognizedButton.disabled = normalizeOcrNumber(recognizedValue.value) === "";
  });
  useRecognizedButton.addEventListener("click", applyRecognizedValue);
  retryPhotoButton.addEventListener("click", () => { photoInput.value = ""; photoInput.click(); });
  cancelPhotoButton.addEventListener("click", () => photoDialog.close());
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
