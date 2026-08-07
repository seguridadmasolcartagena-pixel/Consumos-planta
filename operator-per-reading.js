/*
 * Mantiene la trazabilidad de operario por lectura dentro del JSON compartido.
 * Se carga después de app.js y sustituye únicamente la recopilación y guardado
 * del borrador, sin modificar el contrato actual con Power Automate.
 */

collectReadings = function collectReadingsWithOperator(onlyDirty = false) {
  const result = [];
  let firstInvalid = null;
  const currentOperator = operatorInput.value.trim();
  const now = new Date().toISOString();

  for (const item of READINGS) {
    if (onlyDirty && !dirtyColumns.has(item.ColumnaLectura)) continue;

    const input = document.querySelector(`[data-column="${item.ColumnaLectura}"]`);
    if (!input.value.trim()) continue;

    if (!validateInput(input)) {
      firstInvalid ||= input;
      continue;
    }

    const modifiedNow = dirtyColumns.has(item.ColumnaLectura);
    const operator = modifiedNow
      ? currentOperator
      : (input.dataset.updatedBy || "");

    result.push({
      Orden: item.Orden,
      Bloque: item.Bloque,
      NombreCampo: item.NombreCampo,
      ColumnaLectura: item.ColumnaLectura,
      ColumnaTotales: item.ColumnaTotales,
      TipoValidacion: item.TipoValidacion,
      Valor: parseNumber(input.value),
      Operario: operator,
      FechaRegistro: modifiedNow ? now : (input.dataset.updatedAt || "")
    });
  }

  if (firstInvalid) {
    firstInvalid.closest("details").open = true;
    firstInvalid.focus();
    throw new Error("Revisa los valores marcados. Solo se admiten números positivos y decimales.");
  }

  return result;
};

saveSharedDraft = async function saveSharedDraftWithOperator(showConfirmation = true) {
  clearTimeout(saveTimer);
  if (saveInFlight) return false;
  if (!dateInput.value) throw new Error("Selecciona la fecha de lectura.");
  if (!operatorInput.value.trim()) {
    operatorInput.focus();
    throw new Error("Introduce el nombre del operario antes de guardar.");
  }

  if (!dirtyColumns.size) {
    if (showConfirmation) showToast("No hay cambios nuevos que guardar.");
    return true;
  }

  const changedColumns = new Set(dirtyColumns);
  const savedAt = new Date().toISOString();
  const readings = collectReadings(false);

  saveInFlight = true;
  setButtonsBusy(true, "Guardando…");
  draftLabel.textContent = "Guardando en SharePoint…";

  try {
    const draft = buildDraft("Borrador", readings);
    const result = await callFlow({
      accion: "guardar",
      fechaLectura: dateInput.value,
      borradorJson: JSON.stringify(draft)
    });

    // Solo las medidas modificadas en este guardado cambian de operario.
    changedColumns.forEach((column) => {
      const input = document.querySelector(`[data-column="${column}"]`);
      if (!input || !input.value.trim()) return;
      input.dataset.updatedBy = operatorInput.value.trim();
      input.dataset.updatedAt = savedAt;
      input.closest(".reading-field").classList.add("shared");
    });

    dirtyColumns.clear();
    draftLabel.textContent = "Borrador guardado en SharePoint";
    setSharedState(
      "ready",
      "Archivo JSON guardado",
      result.mensaje || `${readings.length} lecturas disponibles para el otro operario.`
    );
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
};

// Conserva también la fecha de registro cuando el borrador viene de SharePoint.
const originalNormalizeServerReadings = normalizeServerReadings;
normalizeServerReadings = function normalizeServerReadingsWithOperator(result) {
  return originalNormalizeServerReadings(result).map((item) => ({
    ...item,
    Operario: item.Operario || item.operario || "",
    FechaRegistro: item.FechaRegistro || item.fechaRegistro || ""
  }));
};

// Al terminar una carga, copia la fecha de registro a cada input sin alterar
// el operario que ya establece loadSharedDraft.
const originalLoadSharedDraft = loadSharedDraft;
loadSharedDraft = async function loadSharedDraftWithOperatorMetadata() {
  await originalLoadSharedDraft();
  if (!dateInput.value || !isConfigured()) return;

  try {
    const result = await callFlow({ accion: "cargar", fechaLectura: dateInput.value });
    const readings = normalizeServerReadings(result);
    readings.forEach((item) => {
      const column = item.ColumnaLectura || item.columna || item.Columna;
      const input = document.querySelector(`[data-column="${column}"]`);
      if (!input) return;
      input.dataset.updatedAt = item.FechaRegistro || "";
    });
  } catch {
    // La carga principal ya muestra cualquier error al usuario.
  }
};
