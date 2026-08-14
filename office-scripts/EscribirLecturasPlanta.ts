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

const COLUMNAS_PERMITIDAS = [
  "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N",
  "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "AA",
  "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ", "AK", "AL",
  "AM", "AN", "AO", "AP", "AQ", "AR"
];
const TOTAL_COLUMNAS_HASTA_AR = 44;
const MAX_FILAS_BUSQUEDA = 366;

/** Escribe las lecturas recibidas en la fila de la fecha dentro de la hoja Mes Año. */
function main(workbook: ExcelScript.Workbook, fechaLectura: string, lecturasJson: string): ResultadoScript {
  const fechaIso = normalizeDate(fechaLectura);
  const fecha = parseIsoDate(fechaIso);
  const hojaNombre = `${MESES[fecha.getUTCMonth()]} ${fecha.getUTCFullYear()}`;
  const hoja = workbook.getWorksheet(hojaNombre);
  if (!hoja) throw new Error(`No existe la hoja '${hojaNombre}'.`);

  const lecturas = parseLecturas(lecturasJson);
  if (lecturas.length < 1 || lecturas.length > 43) {
    throw new Error(`Se esperaban entre 1 y 43 lecturas y se recibieron ${lecturas.length}.`);
  }

  // Una hoja con formato aplicado a columnas completas puede tener un UsedRange enorme.
  // Limitamos la lectura a 366 filas, más que suficiente para una hoja mensual.
  const fechas = hoja.getRangeByIndexes(0, 0, MAX_FILAS_BUSQUEDA, 1).getValues();
  const filaExcel0 = findDateRow(fechas, fechaIso);
  if (filaExcel0 === -1) throw new Error(`No se encontró ${formatSpanishDate(fechaIso)} en la columna A de '${hojaNombre}'.`);

  const valores = hoja.getRangeByIndexes(0, 0, filaExcel0 + 1, TOTAL_COLUMNAS_HASTA_AR).getValues();
  const errores: string[] = [];
  const valoresPorColumna: { columna: number; valor: number }[] = [];
  const letrasRecibidas = new Set<string>();

  for (const lectura of lecturas) {
    validateReading(lectura);
    const indice = columnLetterToIndex(lectura.ColumnaLectura);
    if (indice < 1 || indice > 43) throw new Error(`La columna ${lectura.ColumnaLectura} está fuera del rango B:AR.`);
    if (letrasRecibidas.has(lectura.ColumnaLectura)) throw new Error(`La columna ${lectura.ColumnaLectura} está repetida.`);
    letrasRecibidas.add(lectura.ColumnaLectura);
    valoresPorColumna.push({ columna: indice, valor: lectura.Valor });

    if (lectura.TipoValidacion !== "LecturaDirecta") {
      const anterior = findPreviousReadingInMemory(valores, filaExcel0, indice);
      if (anterior !== null && lectura.Valor < anterior) {
        errores.push(`${lectura.NombreCampo}: ${lectura.Valor} es menor que la lectura anterior ${anterior}.`);
      }
    }
  }

  for (const columna of letrasRecibidas) {
    if (!COLUMNAS_PERMITIDAS.includes(columna)) {
      throw new Error(`La columna ${columna} no está permitida.`);
    }
  }

  if (errores.length) {
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

  // Escribe únicamente las columnas recibidas. Las medidas no recibidas no se modifican.
  const ordenadas = valoresPorColumna.sort((a, b) => a.columna - b.columna);
  let inicio = 0;
  while (inicio < ordenadas.length) {
    let fin = inicio;
    while (fin + 1 < ordenadas.length && ordenadas[fin + 1].columna === ordenadas[fin].columna + 1) fin += 1;
    const columnaInicial = ordenadas[inicio].columna;
    const valoresBloque = ordenadas.slice(inicio, fin + 1).map((entry) => entry.valor);
    hoja.getRangeByIndexes(filaExcel0, columnaInicial, 1, valoresBloque.length).setValues([valoresBloque]);
    inicio = fin + 1;
  }

  return {
    ok: true,
    hojaLecturas: hojaNombre,
    filaLecturas: filaExcel0 + 1,
    lecturasEscritas: lecturas.length,
    mensaje: "Lecturas escritas correctamente."
  };
}

function parseLecturas(json: string): Lectura[] {
  const parsed = JSON.parse(String(json).trim()) as Lectura[] | { lecturas?: Lectura[] };
  if (Array.isArray(parsed)) return parsed;
  return parsed && Array.isArray(parsed.lecturas) ? parsed.lecturas : [];
}

function validateReading(lectura: Lectura): void {
  if (!lectura.ColumnaLectura || !/^[A-Z]{1,2}$/.test(lectura.ColumnaLectura)) throw new Error(`Columna no válida: ${lectura.ColumnaLectura}.`);
  if (typeof lectura.Valor !== "number" || !Number.isFinite(lectura.Valor) || lectura.Valor < 0) throw new Error(`Valor no válido para ${lectura.NombreCampo}.`);
  if (lectura.ColumnaLectura === "R" && lectura.Valor > 100) {
    throw new Error(`Nivel O2 (%): ${lectura.Valor} está fuera del intervalo permitido de 0 a 100.`);
  }
}

function findDateRow(values: (string | number | boolean)[][], fecha: string): number {
  for (let row = 0; row < values.length; row += 1) if (cellToIsoDate(values[row][0]) === fecha) return row;
  return -1;
}

function findPreviousReadingInMemory(values: (string | number | boolean)[][], currentRow: number, columnIndex: number): number | null {
  for (let row = currentRow - 1; row >= 0; row -= 1) {
    const value = toNumber(values[row][columnIndex]);
    if (value !== null) return value;
  }
  return null;
}

function columnLetterToIndex(column: string): number {
  let result = 0;
  for (let index = 0; index < column.length; index += 1) result = result * 26 + column.charCodeAt(index) - 64;
  return result - 1;
}

function toNumber(value: string | number | boolean): number | null {
  if (typeof value === "number" && Number.isFinite(value)) return value;
  if (typeof value !== "string" || !value.trim()) return null;
  const parsed = Number(value.trim().replace(",", "."));
  return Number.isFinite(parsed) ? parsed : null;
}

function cellToIsoDate(value: string | number | boolean): string | null {
  if (typeof value === "number" && Number.isFinite(value)) return isoFromDate(new Date(Math.round((value - 25569) * 86400000)));
  if (typeof value !== "string") return null;
  const text = value.trim();
  const iso = /^(\d{4})-(\d{1,2})-(\d{1,2})/.exec(text);
  if (iso) return `${iso[1]}-${iso[2].padStart(2, "0")}-${iso[3].padStart(2, "0")}`;
  const spanish = /^(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})$/.exec(text);
  return spanish ? `${spanish[3]}-${spanish[2].padStart(2, "0")}-${spanish[1].padStart(2, "0")}` : null;
}

function parseIsoDate(value: string): Date {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(value)) throw new Error(`Fecha no válida: ${value}.`);
  const date = new Date(`${value}T00:00:00Z`);
  if (Number.isNaN(date.getTime())) throw new Error(`Fecha no válida: ${value}.`);
  return date;
}

function normalizeDate(value: string): string {
  const text = String(value).trim();
  const iso = /^(\d{4})-(\d{1,2})-(\d{1,2})$/.exec(text);
  const spanish = /^(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})$/.exec(text);
  const normalized = iso
    ? `${iso[1]}-${iso[2].padStart(2, "0")}-${iso[3].padStart(2, "0")}`
    : spanish
      ? `${spanish[3]}-${spanish[2].padStart(2, "0")}-${spanish[1].padStart(2, "0")}`
      : "";
  if (!normalized) throw new Error(`Fecha no válida: ${text}. Usa DD/MM/AAAA o AAAA-MM-DD.`);
  const date = parseIsoDate(normalized);
  if (isoFromDate(date) !== normalized) throw new Error(`La fecha ${text} no existe.`);
  return normalized;
}

function formatSpanishDate(value: string): string {
  const parts = value.split("-");
  return `${parts[2]}/${parts[1]}/${parts[0]}`;
}

function isoFromDate(date: Date): string {
  return `${date.getUTCFullYear()}-${String(date.getUTCMonth() + 1).padStart(2, "0")}-${String(date.getUTCDate()).padStart(2, "0")}`;
}
