# Configuración de SharePoint y Power Automate

Esta versión usa Power Automate como API. La aplicación nunca accede directamente a SharePoint ni al Excel.

## 1. Crear la lista de borradores

En el sitio de SharePoint de planta, crea una lista llamada `LecturasPlantaBorradores` con estas columnas:

| Columna | Tipo | Configuración |
|---|---|---|
| `Title` | Texto | Se usará la misma clave técnica |
| `Clave` | Texto | Valores únicos: Sí; ejemplo `2026-08-05|B` |
| `FechaClave` | Texto | Formato `AAAA-MM-DD`, indexada |
| `Columna` | Texto | Columna del Excel, por ejemplo `B` |
| `Valor` | Número | Permitir decimales |
| `Operario` | Texto | Nombre del último operario que cambió el valor |
| `FechaRegistro` | Fecha y hora | Incluir hora |
| `Dispositivo` | Texto | Móvil, tableta o escritorio |
| `Estado` | Elección | `Borrador`, `Enviado`, `Aprobado`, `Rechazado`, `Error` |

La clave única evita dos registros distintos para el mismo contador y día. Concede edición solo al grupo de operarios y responsables.

## 2. Crear el Office Script

1. Abre en Excel para la web uno de los libros mensuales.
2. Ve a `Automatizar > Nuevo script`.
3. Nómbralo `EscribirLecturasPlanta`.
4. Sustituye el contenido por [EscribirLecturasPlanta.ts](../office-scripts/EscribirLecturasPlanta.ts).
5. Guarda el script en una ubicación accesible para la cuenta de conexión del flujo.

El script busca la fecha en la columna A de la hoja mensual y de `Totales`. Para totalizadores calcula `lectura nueva - lectura anterior` y divide el resultado entre todos los días naturales transcurridos. Para `TotalizadorDiv10` aplica además el factor `/10`. Si falta una fila diaria en `Totales`, devuelve un aviso; si una lectura baja respecto a la anterior, detiene toda la escritura.

## 3. Crear el flujo HTTP

Crea un flujo de nube instantáneo llamado `API_Lecturas_Planta` con el disparador `Cuando se recibe una solicitud HTTP`.

Usa este esquema JSON en el disparador:

```json
{
  "type": "object",
  "properties": {
    "accion": { "type": "string" },
    "fechaLectura": { "type": "string" },
    "horaLectura": { "type": "string" },
    "operario": { "type": "string" },
    "fechaHoraRegistro": { "type": "string" },
    "dispositivo": {},
    "lecturasJson": { "type": "string" },
    "lecturas": {
      "type": "array",
      "items": {
        "type": "object",
        "properties": {
          "Orden": { "type": "integer" },
          "Bloque": { "type": "string" },
          "NombreCampo": { "type": "string" },
          "ColumnaLectura": { "type": "string" },
          "ColumnaTotales": { "type": "string" },
          "TipoValidacion": { "type": "string" },
          "Valor": { "type": "number" }
        },
        "required": ["Orden", "NombreCampo", "ColumnaLectura", "ColumnaTotales", "TipoValidacion", "Valor"]
      }
    }
  },
  "required": ["accion", "fechaLectura"]
}
```

Añade un control `Cambiar` usando `triggerBody()?['accion']`. Crea los casos `cargar`, `guardar`, `enviar` y un caso predeterminado que responda `400`.

### Caso `cargar`

1. Añade `SharePoint > Obtener elementos` sobre `LecturasPlantaBorradores`.
2. En `Consulta de filtro` usa:

   ```text
   FechaClave eq '@{triggerBody()?['fechaLectura']}'
   ```

3. Añade `Operaciones de datos > Seleccionar`; como origen usa `value` de Obtener elementos y mapea:

   | Salida | Valor SharePoint |
   |---|---|
   | `ColumnaLectura` | `item()?['Columna']` |
   | `Valor` | `item()?['Valor']` |
   | `Operario` | `item()?['Operario']` |
   | `Estado` | `item()?['Estado']?['Value']` |

4. Añade `Solicitud > Respuesta`, código `200`, cuerpo:

   ```json
   {
     "ok": true,
     "lecturas": "@{body('Seleccionar')}"
   }
   ```

### Caso `guardar`

1. Añade `Aplicar a cada` sobre `triggerBody()?['lecturas']`.
2. Dentro, crea `Componer - Clave`:

   ```text
   concat(triggerBody()?['fechaLectura'],'|',items('Aplicar_a_cada')?['ColumnaLectura'])
   ```

3. Añade `Obtener elementos`, filtro `Clave eq '@{outputs('Componer_-_Clave')}'` y recuento máximo `1`.
4. Añade una condición: `length(body('Obtener_elementos')?['value'])` es mayor que `0`.
5. Si existe, usa `Actualizar elemento` con el ID:

   ```text
   first(body('Obtener_elementos')?['value'])?['ID']
   ```

6. Si no existe, usa `Crear elemento`.
7. En ambos casos asigna `Title` y `Clave` a la clave compuesta, `FechaClave`, `Columna`, `Valor`, `Operario`, `FechaRegistro`, `Dispositivo` y `Estado = Borrador`.
8. Fuera del bucle responde `200`:

   ```json
   { "ok": true, "mensaje": "Borrador guardado en SharePoint" }
   ```

No actives la simultaneidad dentro de `Aplicar a cada`: así se evitan carreras al actualizar la misma clave.

### Caso `enviar`

1. Añade una condición `length(triggerBody()?['lecturas'])` igual a `43`. Si no se cumple, responde `400` y no continúes.
2. Responde inmediatamente con código `202`:

   ```json
   { "ok": true, "mensaje": "Lecturas enviadas a aprobación" }
   ```

3. Después de la respuesta, añade `Iniciar y esperar una aprobación`, tipo `Aprobar/Rechazar - El primero en responder`. Incluye fecha, operario y el JSON recibido en los detalles.
4. Añade una condición: `Outcome` igual a `Approve`.
5. En la rama de rechazo, obtén los elementos de la fecha y actualiza `Estado = Rechazado`.
6. En la rama aprobada, continúa con la selección del archivo y la ejecución del script.

## 4. Seleccionar el Excel mensual

Los archivos deben llamarse exactamente `Datos_Planta_Mes_Año.xlsx`, por ejemplo `Datos_Planta_Agosto_2026.xlsx`.

1. Inicializa una variable de matriz `Meses` con los doce meses en español, de enero a diciembre.
2. Crea `Componer - Mes`:

   ```text
   variables('Meses')[sub(int(formatDateTime(triggerBody()?['fechaLectura'],'MM')),1)]
   ```

3. Crea `Componer - Nombre archivo`:

   ```text
   concat('Datos_Planta_',outputs('Componer_-_Mes'),'_',formatDateTime(triggerBody()?['fechaLectura'],'yyyy'),'.xlsx')
   ```

4. Usa `SharePoint > Obtener metadatos del archivo mediante la ruta` con la carpeta real de consumos y el nombre anterior.
5. Añade `Excel Online (Empresa) > Ejecutar script`. Como archivo usa el `Identificador` obtenido por SharePoint.
6. Selecciona `EscribirLecturasPlanta` y pasa:

   | Parámetro | Valor |
   |---|---|
   | `fechaLectura` | `triggerBody()?['fechaLectura']` |
   | `lecturasJson` | `triggerBody()?['lecturasJson']` |

7. Si termina correctamente, actualiza los registros de la fecha a `Aprobado`. Configura la rama `Ejecutar después` de error para poner `Estado = Error` y avisar al responsable.

## 5. Respuestas CORS y conexión de la app

En las acciones `Respuesta`, añade:

```text
Access-Control-Allow-Origin: https://seguridadmasolcartagena-pixel.github.io
Access-Control-Allow-Headers: Content-Type
Access-Control-Allow-Methods: POST, OPTIONS
```

Guarda el flujo, copia la URL del disparador y sustitúyela en `config.js`. La URL firmada puede verse desde el navegador; para producción debe restringirse el acceso a usuarios corporativos mediante Microsoft Entra ID o un proxy autenticado. Regenera la URL antigua que estaba publicada en el repositorio.

## 6. Prueba de aceptación

1. Operario A abre una fecha, completa dos lecturas y espera `Borrador guardado en SharePoint`.
2. Operario B abre la misma fecha en otro dispositivo y pulsa actualizar; debe ver ambas lecturas en verde.
3. B completa otra lectura; A debe verla al actualizar.
4. Verifica que el botón definitivo solo se habilita con 43 lecturas.
5. Aprueba un envío de prueba y comprueba la hoja mensual y `Totales`.
6. Prueba un salto viernes-lunes: la diferencia debe quedar dividida entre sábado, domingo y lunes.
7. Prueba una lectura inferior a la anterior: el script debe rechazar toda la escritura.
