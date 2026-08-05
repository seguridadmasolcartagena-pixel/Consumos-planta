interface Lectura {
  Orden: number;
  Bloque: string;
  NombreCampo: string;
  ColumnaLectura: string;
  ColumnaTotales: string;
  TipoValidacion: "Totalizador" | "TotalizadorDiv10" | "LecturaDirecta";
  Valor: number;
}

interface LecturaCalculada {
  lectura: Lectura;
  consumos: Array<{ row: number; valor: number }>;
  aviso?: string;
}

interface ResultadoScript {
  ok: boolean;
  hojaLecturas: string;
  filaLecturas: number;
  lecturasEscritas: number;
  consumosEscritos: number;
  avisos: string[];
}

const MESES = [
  "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
  "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"
];

/**
 * Escribe lecturas acumuladas en la hoja mensual y consumos diarios en Totales.
 * Si hay varios dias desde la lectura anterior, reparte el consumo por igual.
 */
function main(workbook: ExcelScript.Workbook, fechaLectura: string, lecturasJson: string): ResultadoScript {
  const fecha = parseIsoDate(fechaLectura);
  const hojaMensualNombre = `${MESES[fecha.getUTCMonth()]} ${fecha.getUTCFullYear()}`;
  const hojaMensual = workbook.getWorksheet(hojaMensualNombre);
  const hojaTotales = workbook.getWorksheet("Totales");

  if (!hojaMensual) throw new Error(`No existe la hoja mensual '${hojaMensualNombre}'.`);
  if (!hojaTotales) throw new Error("No existe la hoja 'Totales'.");

  const lecturas = parseLecturas(lecturasJson);
  if (lecturas.length === 0) throw new Error("El JSON no contiene lecturas.");

  const monthlyDates = getDateRows(hojaMensual);
  const totalsDates = getDateRows(hojaTotales);
  const currentMonthlyRow = monthlyDates.get(fechaLectura);
  const currentTotalsRow = totalsDates.get(fechaLectura);

  if (currentMonthlyRow === undefined) throw new Error(`No se encontró ${fechaLectura} en la columna A de '${hojaMensualNombre}'.`);
  if (currentTotalsRow === undefined) throw new Error(`No se encontró ${fechaLectura} en la columna A de 'Totales'.`);

  const errores: string[] = [];
  const calculadas: LecturaCalculada[] = [];

  // Primero se valida todo. No se modifica el libro si una lectura es incoherente.
  for (const lectura of lecturas) {
    validateReadingShape(lectura);

    if (lectura.TipoValidacion === "LecturaDirecta") {
      calculadas.push({ lectura, consumos: [{ row: currentTotalsRow, valor: lectura.Valor }] });
      continue;
    }

    const anterior = findPreviousReading(hojaMensual, lectura.ColumnaLectura, currentMonthlyRow);
    if (!anterior) {
      calculadas.push({
        lectura,
        consumos: [],
        aviso: `${lectura.NombreCampo}: lectura guardada sin consumo porque no hay una lectura anterior en la hoja mensual.`
      });
      continue;
    }

    if (lectura.Valor < anterior.valor) {
      errores.push(`${lectura.NombreCampo}: ${lectura.Valor} es menor que la lectura anterior ${anterior.valor} (${anterior.fecha}).`);
      continue;
    }

    const dias = daysBetween(anterior.fecha, fechaLectura);
    if (dias < 1) {
      errores.push(`${lectura.NombreCampo}: la fecha anterior ${anterior.fecha} no es anterior a ${fechaLectura}.`);
      continue;
    }

    const factor = lectura.TipoValidacion === "TotalizadorDiv10" ? 10 : 1;
    const consumoDiario = roundForExcel((lectura.Valor - anterior.valor) / dias / factor);
    const consumos: Array<{ row: number; valor: number }> = [];
    const fechasFaltantes: string[] = [];

    for (let offset = 1; offset <= dias; offset += 1) {
      const dateKey = addDays(anterior.fecha, offset);
      const totalRow = totalsDates.get(dateKey);
      if (totalRow === undefined) fechasFaltantes.push(dateKey);
      else consumos.push({ row: totalRow, valor: consumoDiario });
    }

    calculadas.push({
      lectura,
      consumos,
      aviso: fechasFaltantes.length ? `${lectura.NombreCampo}: no existen filas en Totales para ${fechasFaltantes.join(", ")}.` : undefined
    });
  }

  if (errores.length) throw new Error(`Validación rechazada:\n${errores.join("\n")}`);

  let consumosEscritos = 0;
  const avisos: string[] = [];
  for (const item of calculadas) {
    hojaMensual.getRange(`${item.lectura.ColumnaLectura}${currentMonthlyRow + 1}`).setValue(item.lectura.Valor);
    for (const consumo of item.consumos) {
      hojaTotales.getRange(`${item.lectura.ColumnaTotales}${consumo.row + 1}`).setValue(consumo.valor);
      consumosEscritos += 1;
    }
    if (item.aviso) avisos.push(item.aviso);
  }

  return {
    ok: true,
    hojaLecturas: hojaMensualNombre,
    filaLecturas: currentMonthlyRow + 1,
    lecturasEscritas: lecturas.length,
    consumosEscritos,
    avisos
  };
}

function parseLecturas(json: string): Lectura[] {
  const parsed = JSON.parse(json) as Lectura[];
  return Array.isArray(parsed) ? parsed : [];
}

function validateReadingShape(lectura: Lectura): void {
  if (!lectura.ColumnaLectura || !/^[A-Z]{1,2}$/.test(lectura.ColumnaLectura)) throw new Error(`Columna de lectura no válida: ${lectura.ColumnaLectura}.`);
  if (!lectura.ColumnaTotales || !/^[A-Z]{1,2}$/.test(lectura.ColumnaTotales)) throw new Error(`Columna de Totales no válida: ${lectura.ColumnaTotales}.`);
  if (typeof lectura.Valor !== "number" || !Number.isFinite(lectura.Valor) || lectura.Valor < 0) throw new Error(`Valor no válido para ${lectura.NombreCampo}.`);
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

function findPreviousReading(sheet: ExcelScript.Worksheet, column: string, currentRow: number): { valor: number; fecha: string } | null {
  for (let row = currentRow - 1; row >= 0; row -= 1) {
    const raw = sheet.getRange(`${column}${row + 1}`).getValue();
    const value = toNumber(raw);
    if (value === null) continue;
    const dateKey = cellToIsoDate(sheet.getRange(`A${row + 1}`).getValue());
    if (dateKey) return { valor: value, fecha: dateKey };
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
  const isoMatch = /^(\d{4})-(\d{1,2})-(\d{1,2})/.exec(text);
  if (isoMatch) return `${isoMatch[1]}-${isoMatch[2].padStart(2, "0")}-${isoMatch[3].padStart(2, "0")}`;
  const spanishMatch = /^(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})$/.exec(text);
  if (spanishMatch) return `${spanishMatch[3]}-${spanishMatch[2].padStart(2, "0")}-${spanishMatch[1].padStart(2, "0")}`;
  return null;
}

function parseIsoDate(value: string): Date {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(value)) throw new Error(`Fecha no válida: ${value}. Debe usar AAAA-MM-DD.`);
  const date = new Date(`${value}T00:00:00Z`);
  if (Number.isNaN(date.getTime())) throw new Error(`Fecha no válida: ${value}.`);
  return date;
}

function isoFromDate(date: Date): string {
  return `${date.getUTCFullYear()}-${String(date.getUTCMonth() + 1).padStart(2, "0")}-${String(date.getUTCDate()).padStart(2, "0")}`;
}

function daysBetween(fromIso: string, toIso: string): number {
  return Math.round((parseIsoDate(toIso).getTime() - parseIsoDate(fromIso).getTime()) / 86400000);
}

function addDays(iso: string, days: number): string {
  const date = parseIsoDate(iso);
  date.setUTCDate(date.getUTCDate() + days);
  return isoFromDate(date);
}

function roundForExcel(value: number): number {
  return Math.round((value + Number.EPSILON) * 1000000) / 1000000;
}
