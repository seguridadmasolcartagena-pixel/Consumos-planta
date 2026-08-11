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
}

const MESES = [
  "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
  "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"
];

/** Escribe las 43 lecturas en la fila de la fecha dentro de la hoja Mes Año. */
function main(workbook: ExcelScript.Workbook, fechaLectura: string, lecturasJson: string): ResultadoScript {
  const fecha = parseIsoDate(fechaLectura);
  const hojaNombre = `${MESES[fecha.getUTCMonth()]} ${fecha.getUTCFullYear()}`;
  const hoja = workbook.getWorksheet(hojaNombre);

  if (!hoja) throw new Error(`No existe la hoja '${hojaNombre}'.`);

  const lecturas = parseLecturas(lecturasJson);
  if (lecturas.length !== 43) {
    throw new Error(`Se esperaban 43 lecturas y se recibieron ${lecturas.length}.`);
  }

  const used = hoja.getUsedRange(true);
  if (!used) throw new Error(`La hoja '${hojaNombre}' no contiene datos.`);

  const firstRow = used.getRowIndex();
  const firstColumn = used.getColumnIndex();
  const values = used.getValues();
  const dateOffset = -firstColumn;
  if (dateOffset < 0 || dateOffset >= used.getColumnCount()) {
    throw new Error(`La columna A no forma parte del rango usado de '${hojaNombre}'.`);
  }

  const rowOffset = values.findIndex((row) => cellToIsoDate(row[dateOffset]) === fechaLectura);
  if (rowOffset < 0) {
    throw new Error(`No se encontró ${fechaLectura} en la columna A de '${hojaNombre}'.`);
  }

  const fila = firstRow + rowOffset;

  const errores: string[] = [];
  for (const lectura of lecturas) {
    validateReading(lectura);
    if (lectura.TipoValidacion === "LecturaDirecta") continue;

    const columnIndex = columnToIndex(lectura.ColumnaLectura);
    const anterior = findPreviousReading(values, firstColumn, columnIndex, rowOffset);
    if (anterior !== null && lectura.Valor < anterior) {
      errores.push(`${lectura.NombreCampo}: ${lectura.Valor} es menor que la lectura anterior ${anterior}.`);
    }
  }

  if (errores.length) {
    throw new Error(`Validación rechazada:\n${errores.join("\n")}`);
  }

  const rowValues: (string | number | boolean)[][] = [Array<string | number | boolean>(43).fill("")];
  const usedColumns = new Set<number>();
  for (const lectura of lecturas) {
    const columnIndex = columnToIndex(lectura.ColumnaLectura);
    const outputOffset = columnIndex - 1;
    if (outputOffset < 0 || outputOffset >= 43) {
      throw new Error(`La columna ${lectura.ColumnaLectura} está fuera del rango B:AR.`);
    }
    if (usedColumns.has(columnIndex)) {
      throw new Error(`La columna ${lectura.ColumnaLectura} está repetida en el JSON.`);
    }
    usedColumns.add(columnIndex);
    rowValues[0][outputOffset] = lectura.Valor;
  }

  hoja.getRangeByIndexes(fila, 1, 1, 43).setValues(rowValues);

  return {
    ok: true,
    hojaLecturas: hojaNombre,
    filaLecturas: fila + 1,
    lecturasEscritas: lecturas.length
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
  if (typeof lectura.Valor !== "number" || !Number.isFinite(lectura.Valor) || lectura.Valor < 0) {
    throw new Error(`Valor no válido para ${lectura.NombreCampo}.`);
  }
}

function findPreviousReading(
  values: (string | number | boolean)[][],
  firstColumn: number,
  columnIndex: number,
  currentRowOffset: number
): number | null {
  const columnOffset = columnIndex - firstColumn;
  if (columnOffset < 0 || columnOffset >= values[0].length) return null;

  for (let row = currentRowOffset - 1; row >= 0; row -= 1) {
    const value = toNumber(values[row][columnOffset]);
    if (value !== null) return value;
  }
  return null;
}

function columnToIndex(column: string): number {
  let result = 0;
  for (const character of column) {
    result = result * 26 + character.charCodeAt(0) - 64;
  }
  return result - 1;
}

function toNumber(value: string | number | boolean): number | null {
  if (typeof value === "number" && Number.isFinite(value)) return value;
  if (typeof value !== "string" || !value.trim()) return null;
  const parsed = Number(value.trim().replace(",", "."));
  return Number.isFinite(parsed) ? parsed : null;
}

function cellToIsoDate(value: string | number | boolean): string | null {
  if (typeof value === "number" && Number.isFinite(value)) {
    const date = new Date(Math.round((value - 25569) * 86400000));
    return isoFromDate(date);
  }
  if (typeof value !== "string") return null;

  const text = value.trim();
  const iso = /^(\d{4})-(\d{1,2})-(\d{1,2})/.exec(text);
  if (iso) return `${iso[1]}-${iso[2].padStart(2, "0")}-${iso[3].padStart(2, "0")}`;

  const spanish = /^(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})$/.exec(text);
  if (spanish) return `${spanish[3]}-${spanish[2].padStart(2, "0")}-${spanish[1].padStart(2, "0")}`;
  return null;
}

function parseIsoDate(value: string): Date {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(value)) {
    throw new Error(`Fecha no válida: ${value}. Debe usar AAAA-MM-DD.`);
  }
  const date = new Date(`${value}T00:00:00Z`);
  if (Number.isNaN(date.getTime())) throw new Error(`Fecha no válida: ${value}.`);
  return date;
}

function isoFromDate(date: Date): string {
  return `${date.getUTCFullYear()}-${String(date.getUTCMonth() + 1).padStart(2, "0")}-${String(date.getUTCDate()).padStart(2, "0")}`;
}
