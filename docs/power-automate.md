# Power Automate con borradores JSON en SharePoint

La aplicación no escribe directamente en SharePoint ni en Excel. Solo envía y recibe JSON mediante un flujo HTTP. Power Automate es responsable de crear el borrador, devolverlo al segundo dispositivo y ejecutar directamente el Office Script.

## 1. Crear las carpetas

Dentro de la biblioteca donde están los libros de consumos, crea:

```text
Consumos Planta/
├── Borradores/
└── Procesados/
```

Cada fecha tendrá un único archivo:

```text
Borrador_Lecturas_2026-08-05.json
```

No es necesario crear una lista de SharePoint.

## 2. Formato del archivo JSON

La aplicación genera el documento completo. Ejemplo reducido:

```json
{
  "version": 1,
  "fechaLectura": "2026-08-05",
  "horaLectura": "08:00",
  "estado": "Borrador",
  "actualizadoPor": "Operario A",
  "actualizadoEn": "2026-08-05T07:15:00.000Z",
  "dispositivo": {
    "tipo": "Movil",
    "anchoPantalla": 390,
    "altoPantalla": 844,
    "navegador": "..."
  },
  "lecturas": [
    {
      "Orden": 1,
      "Bloque": "PRODUCCION L a D",
      "NombreCampo": "Contador General H2O",
      "ColumnaLectura": "B",
      "ColumnaTotales": "B",
      "TipoValidacion": "Totalizador",
      "Valor": 123456.78
    }
  ]
}
```

## 3. Crear el flujo

Crea un flujo de nube instantáneo llamado `API_Lecturas_Planta_JSON` con el disparador `Cuando se recibe una solicitud HTTP`.

Esquema del disparador:

```json
{
  "type": "object",
  "properties": {
    "accion": { "type": "string" },
    "fechaLectura": { "type": "string" },
    "borradorJson": { "type": "string" }
  },
  "required": ["accion", "fechaLectura"]
}
```

Después del disparador, añade `Componer - Nombre JSON`:

```text
concat('Borrador_Lecturas_',triggerBody()?['fechaLectura'],'.json')
```

Añade un control `Cambiar` sobre:

```text
triggerBody()?['accion']
```

Crea los casos `cargar`, `guardar`, `enviar` y un caso predeterminado.

## 4. Caso cargar

1. Añade `SharePoint > Obtener archivos (solo propiedades)`.
2. Selecciona la biblioteca que contiene `Consumos Planta`.
3. Limita las entradas a la carpeta `Consumos Planta/Borradores`.
4. En `Consulta de filtro` usa:

   ```text
   FileLeafRef eq '@{outputs('Componer_-_Nombre_JSON')}'
   ```

5. Añade una condición:

   ```text
   length(body('Obtener_archivos_(solo_propiedades)')?['value'])
   ```

   es mayor que `0`.

6. Si existe, añade `Obtener contenido del archivo` usando este identificador:

   ```text
   first(body('Obtener_archivos_(solo_propiedades)')?['value'])?['Identifier']
   ```

7. Responde con código `200` y este cuerpo:

   ```json
   {
     "ok": true,
     "borradorJson": "@{base64ToString(body('Obtener_contenido_del_archivo')?['$content'])}"
   }
   ```

8. Si no existe, responde `200`:

   ```json
   { "ok": true, "borradorJson": "" }
   ```

## 5. Caso guardar

1. Añade `Analizar JSON - Borrador` usando `triggerBody()?['borradorJson']` como contenido. Puedes generar el esquema pegando el ejemplo anterior.
2. Comprueba que la fecha interna coincide con la fecha de la petición:

   ```text
   body('Analizar_JSON_-_Borrador')?['fechaLectura']
   ```

   debe ser igual a `triggerBody()?['fechaLectura']`.

3. Repite `Obtener archivos (solo propiedades)` con el mismo filtro del caso cargar.
4. Si el archivo existe, usa `SharePoint > Actualizar archivo`:

   - Identificador: el primer `Identifier` encontrado.
   - Contenido: `triggerBody()?['borradorJson']`.

5. Si no existe, usa `SharePoint > Crear archivo`:

   - Carpeta: `Consumos Planta/Borradores`.
   - Nombre: salida de `Componer - Nombre JSON`.
   - Contenido: `triggerBody()?['borradorJson']`.

6. Después de la condición responde `200`:

   ```json
   { "ok": true, "mensaje": "Archivo JSON guardado en SharePoint" }
   ```

La aplicación envía siempre el documento completo. El segundo operario primero carga el archivo existente, añade sus medidas y vuelve a guardar el JSON completo.

## 6. Caso enviar

1. Analiza `triggerBody()?['borradorJson']`.
2. Comprueba:

   ```text
   length(body('Analizar_JSON_-_Envio')?['lecturas'])
   ```

   igual a `43`.

3. Crea o actualiza el archivo JSON igual que en el caso guardar.
4. Responde inmediatamente con código `202`:

   ```json
   { "ok": true, "mensaje": "Lecturas recibidas. Se está actualizando el Excel" }
   ```

5. Después de la respuesta, localiza el Excel mensual y ejecuta el Office Script.
6. Después de escribir correctamente en Excel, mueve el JSON a `Consumos Planta/Procesados`.

## 7. Seleccionar el Excel mensual

Inicializa una variable de matriz `Meses`:

```json
["Enero","Febrero","Marzo","Abril","Mayo","Junio","Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre"]
```

Obtén el mes:

```text
variables('Meses')[sub(int(formatDateTime(triggerBody()?['fechaLectura'],'MM')),1)]
```

Construye el nombre según la nomenclatura real de la biblioteca. Para los archivos visibles actualmente, el ejemplo sería:

```text
concat('Datos-Planta-',outputs('Componer_-_Mes'),'-',formatDateTime(triggerBody()?['fechaLectura'],'yyyy'),'.xlsx')
```

Ejemplo:

```text
Datos-Planta-Agosto-2026.xlsx
```

Usa `Obtener metadatos del archivo mediante la ruta` y después `Excel Online (Empresa) > Ejecutar script`.

Parámetros del script:

| Parámetro | Valor |
|---|---|
| `fechaLectura` | `triggerBody()?['fechaLectura']` |
| `lecturasJson` | `string(body('Analizar_JSON_-_Envio')?['lecturas'])` |

El script está en `office-scripts/EscribirLecturasPlanta.ts`.

## 8. Conectar la aplicación

Guarda el flujo y copia la URL HTTP generada. Sustituye el marcador de `config.js`:

```js
window.MASOL_CONFIG = Object.freeze({
  FLOW_URL: "URL_NUEVA_DEL_FLUJO",
  AUTO_SAVE_DELAY_MS: 1400
});
```

Regenera la URL antigua antes de producción. Una URL firmada incluida en una aplicación web puede ser inspeccionada desde el navegador.

## 9. Prueba mínima

1. El operario A introduce dos lecturas.
2. Comprueba que se crea `Borrador_Lecturas_FECHA.json`.
3. El operario B abre la misma fecha y debe recuperar ambas lecturas.
4. B añade otra lectura y comprueba que el archivo contiene las tres.
5. Completa las 43 lecturas y envía.
6. Aprueba y confirma la escritura en el Excel y el traslado del JSON a `Procesados`.

## Límite conocido

El archivo completo se sustituye en cada guardado. Si dos operarios modifican exactamente la misma fecha al mismo tiempo, el último guardado puede imponerse al anterior. Para el trabajo secuencial descrito (primer operario y después segundo operario) el modelo es adecuado y mucho más sencillo que una lista por contador.
