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
