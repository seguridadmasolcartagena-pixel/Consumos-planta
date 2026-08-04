# Lecturas de planta

Aplicación web instalable para registrar las lecturas diarias de consumos de Masol Cartagena. Envía las lecturas a Power Automate, donde siguen el circuito de aprobación y escritura en el Excel mensual de SharePoint.

## Publicación

El flujo de GitHub Actions publica automáticamente la rama `main` mediante GitHub Pages. En la configuración del repositorio, `Settings > Pages > Build and deployment` debe estar seleccionado `GitHub Actions`.

## Prueba

1. Seleccionar una fecha existente en la hoja mensual.
2. Introducir el nombre del operario y una o más lecturas controladas.
3. Enviar a aprobación.
4. Comprobar la ejecución en Power Automate antes de aprobarla.

## Configuración temporal

La URL del disparador HTTP está configurada en `app.js` para la fase de prueba. Antes de utilizar la aplicación con datos reales se debe regenerar la firma del flujo y proteger el disparador.
