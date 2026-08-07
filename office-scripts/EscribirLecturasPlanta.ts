interface Lectura {
  Orden: number;
  Bloque: string;
  NombreCampo: string;
  ColumnaLectura: string;
  ColumnaTotales: string;
  TipoValidacion: "Totalizador" | "TotalizadorDiv10" | "LecturaDirecta";
  Valor: number;
}

interface ResultadoScript {
  ok: boolean;
  hojaLecturas: string;
  filaLecturas: number;
  lecturasEscritas: number;
  codigo?: string;
  mensaje?: string;
  errores?: string[];
}

const MESES = [
  "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
  "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"
];

const TOTAL_LECTURAS = 43;
const COLUMNA_PRIMERA_LECTURA = 1; // B, índice 0-based.
const COLUMNA_ULTIMA_LECTURA = 43; // AR, índice 0-based.
const TOTAL_COLUMNAS_HASTA_AR = 44; // A:AR.

/**
 * Escribe las 43 lecturas en la fila de la fecha dentro de la hoja Mes Año.
 *
 * El script trabaja en memoria con A:AR para evitar cientos de accesos
 * individuales al libro, que pueden ralentizar o bloquear Power Automate.
 */
function main(
  workbook: ExcelScript.Workbook,
  fechaLectura: string,
  lecturasJson: string
): ResultadoScript {
  const fecha = parseIsoDate(fechaLectura);
  const hojaNombre = `${MESES[fecha.getUTCMonth()]} ${fecha.getUTCFullYear()}`;
  const hoja = workbook.getWorksheet(hojaNombre);

  if (!hoja) {
    throw new Error(`No existe la hoja '${hojaNombre}'.`);
  }

  const lecturas = parseLecturas(lecturasJson);
  if (lecturas.length !== TOTAL_LECTURAS) {
    throw new Error(`Se esperaban ${TOTAL_LECTURAS} lecturas y se recibieron ${lecturas.length}.`);
  }

  const used = hoja.getUsedRange(true);
  if (!used) {
    throw new Error(`La hoja '${hojaNombre}' está vacía.`);
  }

  const primeraFilaUsada = used.getRowIndex();
  const totalFilasUsadas = used.getRowCount();

  // Una sola lectura masiva del libro: columnas A:AR para todas las filas usadas.
  const valores = hoja
    .getRangeByIndexes(primeraFilaUsada, 0, totalFilasUsadas, TOTAL_COLUMNAS_HASTA_AR)
    .getValues();

  const filaRelativa = findDateRow(valores, fechaLectura);
  if (filaRelativa === -1) {
    throw new Error(`No se encontró ${fechaLectura} en la columna A de '${hojaNombre}'.`);
  }

  const filaExcel0 = primeraFilaUsada + filaRelativa;
  const errores: string[] = [];
  const valoresSalida: number[] = new Array<number>(TOTAL_LECTURAS);
  const columnasRecibidas = new Set<number>();

  for (const lectura of lecturas) {
    validateReading(lectura);

    const indiceColumna = columnLetterToIndex(lectura.ColumnaLectura);
    if (
      indiceColumna < COLUMNA_PRIMERA_LECTURA ||
      indiceColumna > COLUMNA_ULTIMA_LECTURA
    ) {
      throw new Error(`La columna ${lectura.ColumnaLectura} está fuera del rango B:AR.`);
    }

    if (columnasRecibidas.has(indiceColumna)) {
      throw new Error(`La columna ${lectura.ColumnaLectura} está repetida en las lecturas recibidas.`);
    }
    columnasRecibidas.add(indiceColumna);

    // B debe ocupar la posición 0, C la 1, ... AR la 42.
    valoresSalida[indiceColumna - COLUMNA_PRIMERA_LECTURA] = lectura.Valor;

    // Las lecturas directas no son totalizadores y no se comparan con días anteriores.
    if (lectura.TipoValidacion === "LecturaDirecta") {
      continue;
    }

    const anterior = findPreviousReadingInMemory(
      valores,
      filaRelativa,
      indiceColumna
    );

    if (anterior !== null && lectura.Valor < anterior) {
      errores.push(
        `${lectura.NombreCampo}: ${lectura.Valor} es menor que la lectura anterior ${anterior}.`
      );
    }
  }

  if (columnasRecibidas.size !== TOTAL_LECTURAS || valoresSalida.some((v) => v === undefined)) {
    throw new Error("No se han recibido exactamente las columnas B:AR una vez cada una.");
  }

  // Error de negocio controlado: el Office Script termina correctamente y
  // Power Automate puede responder a la app sin marcar la acción como fallida.
  if (errores.length > 0) {
    return {
      ok: false,
      hojaLecturas: hojaNombre,
      filaLecturas: filaExcel0 + 1,
      lecturasEscritas: 0,
      codigo: "LECTURA_MENOR_DIA_ANTERIOR",
      mensaje: "No se pueden introducir medidas menores a las del día anterior.",
      errores
    };
  }

  // Una sola escritura masiva de B:AR en vez de 43 setValue independientes.
  hoja
    .getRangeByIndexes(filaExcel0, COLUMNA_PRIMERA_LECTURA, 1, TOTAL_LECTURAS)
    .setValues([valoresSalida]);

  return {
    ok: true,
    hojaLecturas: hojaNombre,
    filaLecturas: filaExcel0 + 1,
    lecturasEscritas: TOTAL_LECTURAS,
    mensaje: "Lecturas escritas correctamente."
  };
}

function parseLecturas(json: string): Lectura[] {
  const parsed = JSON.parse(json) as Lectura[];
  return Array.isArray(parsed) ? parsed : [];
}

function validateReading(lectura: Lectura): void {
  if (!lectura.ColumnaLectura || !/^[A-Z]{1,2}$/.test(lectura.ColumnaLectura)) {
    throw new Error(`Columna de lectura no válida: ${lectura.ColumnaLectura}.`);
  }

  if (
    typeof lectura.Valor !== "number" ||
    !Number.isFinite(lectura.Valor) ||
    lectura.Valor < 0
  ) {
    throw new Error(`Valor no válido para ${lectura.NombreCampo}.`);
  }
}

function findDateRow(
  values: (string | number | boolean)[][],
  fechaLectura: string
): number {
  for (let row = 0; row < values.length; row += 1) {
    if (cellToIsoDate(values[row][0]) === fechaLectura) {
      return row;
    }
  }
  return -1;
}

function findPreviousReadingInMemory(
  values: (string | number | boolean)[][],
  currentRow: number,
  columnIndex: number
): number | null {
  for (let row = currentRow - 1; row >= 0; row -= 1) {
    const value = toNumber(values[row][columnIndex]);
    if (value !== null) {
      return value;
    }
  }
  return null;
}

function columnLetterToIndex(column: string): number {
  let result = 0;
  for (let i = 0; i < column.length; i += 1) {
    result = result * 26 + (column.charCodeAt(i) - 64);
  }
  return result - 1;
}

function toNumber(value: string | number | boolean): number | null {
  if (typeof value === "number" && Number.isFinite(value)) {
    return value;
  }

  if (typeof value !== "string" || !value.trim()) {
    return null;
  }

  const parsed = Number(value.trim().replace(",", "."));
  return Number.isFinite(parsed) ? parsed : null;
}

function cellToIsoDate(value: string | number | boolean): string | null {
  if (typeof value === "number" && Number.isFinite(value)) {
    const date = new Date(Math.round((value - 25569) * 86400000));
    return isoFromDate(date);
  }

  if (typeof value !== "string") {
    return null;
  }

  const text = value.trim();

  const iso = /^(\d{4})-(\d{1,2})-(\d{1,2})/.exec(text);
  if (iso) {
    return `${iso[1]}-${iso[2].padStart(2, "0")}-${iso[3].padStart(2, "0")}`;
  }

  const spanish = /^(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})$/.exec(text);
  if (spanish) {
    return `${spanish[3]}-${spanish[2].padStart(2, "0")}-${spanish[1].padStart(2, "0")}`;
  }

  return null;
}

function parseIsoDate(value: string): Date {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(value)) {
    throw new Error(`Fecha no válida: ${value}. Debe usar AAAA-MM-DD.`);
  }

  const date = new Date(`${value}T00:00:00Z`);
  if (Number.isNaN(date.getTime())) {
    throw new Error(`Fecha no válida: ${value}.`);
  }

  return date;
}

function isoFromDate(date: Date): string {
  return `${date.getUTCFullYear()}-${String(date.getUTCMonth() + 1).padStart(2, "0")}-${String(date.getUTCDate()).padStart(2, "0")}`;
}
