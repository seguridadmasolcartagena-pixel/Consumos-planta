function main(workbook: ExcelScript.Workbook) {
  /*
   * Rellena los huecos de cada columna de medidas de forma independiente.
   * Usa la hoja que esté activa, por lo que su nombre puede ser "Agosto 2026",
   * "Septiembre 2026", etc. No modifica la hoja "Totales" ni ninguna otra hoja.
   *
   * Plantilla esperada:
   *   - Encabezado "DIA" o "DÍA" en A3.
   *   - Días/fechas en A4:A36.
   *   - Medidas en B:AR.
   */

  const FIRST_TEMPLATE_ROW = 4;
  const LAST_TEMPLATE_ROW = 36;
  const FIRST_MEASURE_COLUMN = 2; // B
  const LAST_MEASURE_COLUMN = 44; // AR
  const DECIMALS = 0;
  const CREATE_BACKUP = true;

  const sheet = workbook.getActiveWorksheet();

  // Evita que el script se ejecute por error en "Totales" u otra hoja.
  const header = sheet
    .getRange("A3")
    .getText()
    .replace(/\u00A0/g, " ")
    .trim()
    .toUpperCase();

  if (header !== "DIA" && header !== "DÍA") {
    throw new Error(
      `No se ha modificado nada. Activa primero la hoja mensual. ` +
      `En la hoja correcta, A3 debe contener "DIA" o "DÍA".`
    );
  }

  const isBlankText = (text: string): boolean => {
    return text.replace(/\u00A0/g, " ").trim() === "";
  };

  const isBlankValue = (
    value: string | number | boolean
  ): boolean => {
    return (
      value === "" ||
      (typeof value === "string" && value.trim() === "")
    );
  };

  const isNumeric = (
    value: string | number | boolean
  ): value is number => {
    return typeof value === "number" && Number.isFinite(value);
  };

  // Localiza las filas que realmente pertenecen al mes según la columna A.
  const templateRowCount = LAST_TEMPLATE_ROW - FIRST_TEMPLATE_ROW + 1;
  const dayTexts = sheet
    .getRangeByIndexes(
      FIRST_TEMPLATE_ROW - 1,
      0,
      templateRowCount,
      1
    )
    .getTexts();

  let firstDayOffset = -1;
  let lastDayOffset = -1;

  for (let r = 0; r < dayTexts.length; r++) {
    if (!isBlankText(dayTexts[r][0])) {
      if (firstDayOffset === -1) {
        firstDayOffset = r;
      }
      lastDayOffset = r;
    }
  }

  if (firstDayOffset === -1) {
    throw new Error(
      "No se ha modificado nada. Escribe primero los días o las fechas en la columna A."
    );
  }

  // No continúa si falta un día/fecha entre el primero y el último.
  for (let r = firstDayOffset; r <= lastDayOffset; r++) {
    if (isBlankText(dayTexts[r][0])) {
      const excelRow = FIRST_TEMPLATE_ROW + r;
      throw new Error(
        `No se ha modificado nada. Falta el día o la fecha en A${excelRow}.`
      );
    }
  }

  const firstDataRow = FIRST_TEMPLATE_ROW + firstDayOffset;
  const rowCount = lastDayOffset - firstDayOffset + 1;
  const columnCount =
    LAST_MEASURE_COLUMN - FIRST_MEASURE_COLUMN + 1;

  const dataRange = sheet.getRangeByIndexes(
    firstDataRow - 1,
    FIRST_MEASURE_COLUMN - 1,
    rowCount,
    columnCount
  );

  const values = dataRange.getValues();
  const formulas = dataRange.getFormulas();

  // Seguridad: la zona de lecturas debe contener valores manuales, no fórmulas.
  // Si encuentra una fórmula, se detiene sin cambiar ninguna celda.
  let formulaAddress = "";

  formulaSearch:
  for (let r = 0; r < rowCount; r++) {
    for (let c = 0; c < columnCount; c++) {
      const formulaOrValue = formulas[r][c];
      if (
        typeof formulaOrValue === "string" &&
        formulaOrValue.startsWith("=")
      ) {
        formulaAddress = sheet
          .getCell(
            firstDataRow - 1 + r,
            FIRST_MEASURE_COLUMN - 1 + c
          )
          .getAddress();
        break formulaSearch;
      }
    }
  }

  if (formulaAddress !== "") {
    throw new Error(
      `No se ha modificado nada. Hay una fórmula en la zona de lecturas ` +
      `(${formulaAddress}). Restaura o usa la plantilla reparada antes de ejecutar el script.`
    );
  }

  type Change = {
    rowOffset: number;
    columnOffset: number;
    value: number;
  };

  const changes: Change[] = [];
  const roundingFactor = Math.pow(10, DECIMALS);
  let skippedEdgeCells = 0;
  let skippedNonNumericCells = 0;

  // Cada columna busca sus propios tramos vacíos. El contenido de las demás
  // columnas de la misma fila no afecta a la interpolación.
  for (let c = 0; c < columnCount; c++) {
    let r = 0;

    while (r < rowCount) {
      if (!isBlankValue(values[r][c])) {
        r++;
        continue;
      }

      const blankStart = r;

      while (r < rowCount && isBlankValue(values[r][c])) {
        r++;
      }

      const previousRow = blankStart - 1;
      const nextRow = r;
      const blankCellCount = nextRow - blankStart;

      // Para repartir hacen falta una lectura anterior y otra posterior
      // dentro de esta misma columna.
      if (previousRow < 0 || nextRow >= rowCount) {
        skippedEdgeCells += blankCellCount;
        continue;
      }

      const previousValue = values[previousRow][c];
      const nextValue = values[nextRow][c];

      if (!isNumeric(previousValue) || !isNumeric(nextValue)) {
        skippedNonNumericCells += blankCellCount;
        continue;
      }

      const intervals = nextRow - previousRow;

      for (let rr = blankStart; rr < nextRow; rr++) {
        const position = rr - previousRow;
        const interpolated =
          previousValue +
          ((nextValue - previousValue) * position) / intervals;

        const rounded =
          Math.round(interpolated * roundingFactor) /
          roundingFactor;

        changes.push({
          rowOffset: rr,
          columnOffset: c,
          value: rounded
        });
      }
    }
  }

  if (changes.length === 0) {
    console.log(
      `No se ha modificado nada en "${sheet.getName()}". ` +
      `No hay celdas vacías situadas entre dos lecturas numéricas de la misma columna.`
    );
    return;
  }

  // Crea una copia completa de la hoja antes de escribir ningún valor.
  let backupName = "";

  if (CREATE_BACKUP) {
    backupName = createUniqueBackupName(
      workbook,
      sheet.getName()
    );

    const backup = sheet.copy(
      ExcelScript.WorksheetPositionType.after,
      sheet
    );

    backup.setName(backupName);
  }

  // Escribe únicamente en las celdas calculadas. No reescribe el rango completo.
  for (const change of changes) {
    sheet
      .getCell(
        firstDataRow - 1 + change.rowOffset,
        FIRST_MEASURE_COLUMN - 1 + change.columnOffset
      )
      .setValue(change.value);
  }

  sheet.activate();

  const backupMessage = CREATE_BACKUP
    ? ` Copia de seguridad creada: "${backupName}".`
    : "";

  const skippedEdgeMessage = skippedEdgeCells > 0
    ? ` Celdas no rellenadas por faltar una lectura anterior o posterior en su columna: ${skippedEdgeCells}.`
    : "";

  const skippedNonNumericMessage = skippedNonNumericCells > 0
    ? ` Celdas no rellenadas porque alguno de los extremos no es numérico: ${skippedNonNumericCells}.`
    : "";

  console.log(
    `Proceso terminado en "${sheet.getName()}". ` +
    `Celdas rellenadas: ${changes.length}.` +
    backupMessage +
    skippedEdgeMessage +
    skippedNonNumericMessage
  );
}

function createUniqueBackupName(
  workbook: ExcelScript.Workbook,
  sourceSheetName: string
): string {
  let baseName = `COPIA_${sourceSheetName}`
    .replace(/[:\\\/\?\*\[\]]/g, "_")
    .trim();

  if (baseName === "") {
    baseName = "COPIA_SEGURIDAD";
  }

  let candidate = baseName.substring(0, 31);
  let counter = 1;

  while (workbook.getWorksheet(candidate) !== undefined) {
    const suffix = `_${counter}`;
    candidate =
      baseName.substring(0, 31 - suffix.length) + suffix;
    counter++;
  }

  return candidate;
}
