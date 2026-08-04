# Lecturas de planta

Aplicación web instalable para registrar las lecturas diarias de consumos de Masol Cartagena. Envía las lecturas a Power Automate, donde siguen el circuito de aprobación y escritura en el Excel mensual de SharePoint.

El formulario sigue el orden físico de la hoja `Mes Año` de `Plantilla_Consumos_Planta.xlsx`: Producción (`B:R`), Mantenimiento eléctrico (`S:AB`) y Mantenimiento/servicios (`AC:AR`). Las columnas de `Totales` se conservan en el mapa enviado al flujo.

## Publicación

El flujo de GitHub Actions publica automáticamente la rama `main` mediante GitHub Pages. En la configuración del repositorio, `Settings > Pages > Build and deployment` debe estar seleccionado `GitHub Actions`.

## Trazabilidad y dispositivos

Cada envio incluye el nombre y correo corporativo del operario, la fecha y hora real del registro y datos basicos del dispositivo (movil, tableta o escritorio, dimensiones y navegador). La identidad se recuerda localmente en cada dispositivo.

El formulario adapta sus columnas y controles a escritorio, tableta y movil. Para validar automaticamente la identidad corporativa sin permitir nombres manuales sera necesario integrar Microsoft Entra ID.

## Prueba

1. Seleccionar una fecha existente en la hoja mensual.
2. Introducir el nombre, el correo corporativo y una o mas lecturas controladas.
3. Enviar a aprobación.
4. Comprobar la ejecución en Power Automate antes de aprobarla.

## Configuración temporal

La URL del disparador HTTP está configurada en `app.js` para la fase de prueba. Antes de utilizar la aplicación con datos reales se debe regenerar la firma del flujo y proteger el disparador.


## Borradores compartidos por etapas

La aplicación trata cada fecha y la hora `08:00` como un formulario único. Cada operario puede añadir una parte de las lecturas y cerrar la aplicación. El siguiente operario recupera desde SharePoint las medidas ya guardadas y continúa con las pendientes.

El navegador conserva además una copia local para no perder datos durante una caída de red. Esa copia no se considera compartida hasta que la pantalla muestre `Borrador compartido guardado`.

El envío final permanece bloqueado hasta que las 43 lecturas sean numéricas y válidas. Antes de enviar, la aplicación sincroniza cualquier cambio pendiente. El backend debe aplicar idempotencia por `idFormulario` para que una misma fecha no se envíe dos veces.

### Contrato del endpoint HTTP

La aplicación sigue utilizando `FLOW_URL` como único endpoint, pero añade el campo `accion`:

| Acción | Finalidad | Respuesta esperada |
|---|---|---|
| `cargar_borrador` | Recuperar el formulario de una fecha | `borrador.lecturasJson`, `borrador.version` y `borrador.estado` |
| `guardar_borrador` | Fusionar solamente las columnas aportadas por el operario | Borrador fusionado y nueva `version` |
| `enviar_completo` | Validar las 43 lecturas y efectuar el único envío final | `mensaje` y estado final |

Las tres acciones reciben `idFormulario`, `fechaLectura` y `horaLectura`. El guardado y el envío incluyen también el operario, su correo, la fecha real del registro y `lecturasJson`.

El servicio debe guardar en SharePoint un registro por `idFormulario`, fusionar las lecturas por `ColumnaLectura`, conservar la autoría y fecha de cada aportación, y rechazar `enviar_completo` si falta alguna de las 43 columnas. La operación final debe marcar el registro como enviado antes de ejecutar acciones posteriores para impedir duplicados.
