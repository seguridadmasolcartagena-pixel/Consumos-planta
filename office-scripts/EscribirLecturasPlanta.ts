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

  const fila = getDateRows(hoja).get(fechaLectura);
  if (fila === undefined) {
    throw new Error(`No se encontró ${fechaLectura} en la columna A de '${hojaNombre}'.`);
  }

  const errores: string[] = [];
  for (const lectura of lecturas) {
    validateReading(lectura);
    if (lectura.TipoValidacion === "LecturaDirecta") continue;

    const anterior = findPreviousReading(hoja, lectura.ColumnaLectura, fila);
    if (anterior !== null && lectura.Valor < anterior) {
      errores.push(`${lectura.NombreCampo}: ${lectura.Valor} es menor que la lectura anterior ${anterior}.`);
    }
  }

  if (errores.length) {
    return {
      ok: false,
      hojaLecturas: hojaNombre,
      filaLecturas: fila + 1,
      lecturasEscritas: 0,
      codigo: "LECTURA_MENOR_DIA_ANTERIOR",
      mensaje: "No se pueden introducir medidas menores a las del día anterior.",
      errores
    };
  }

  for (const lectura of lecturas) {
    hoja.getRange(`${lectura.ColumnaLectura}${fila + 1}`).setValue(lectura.Valor);
  }

  return {
    ok: true,
    hojaLecturas: hojaNombre,
    filaLecturas: fila + 1,
    lecturasEscritas: lecturas.length,
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
  if (typeof lectura.Valor !== "number" || !Number.isFinite(lectura.Valor) || lectura.Valor < 0) {
    throw new Error(`Valor no válido para ${lectura.NombreCampo}.`);
  }
}

function getDateRows(sheet: ExcelScript.Worksheet): Map<string, number> {
  const used = sheet.getUsedRange(true);
  if (!used) return new Map<string, number>();

  const firstRow = used.getRowIndex();
  const rowCount = used.getRowCount();
  const values = sheet.getRangeByIndexes(firstRow, 0, rowCount, 1).getValues();
  const result = new Map<string, number>();

  values.forEach((row, index) => {
    const key = cellToIsoDate(row[0]);
    if (key) result.set(key, firstRow + index);
  });
  return result;
}

function findPreviousReading(sheet: ExcelScript.Worksheet, column: string, currentRow: number): number | null {
  for (let row = currentRow - 1; row >= 0; row -= 1) {
    const value = toNumber(sheet.getRange(`${column}${row + 1}`).getValue());
    if (value !== null) return value;
  }
  return null;
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
