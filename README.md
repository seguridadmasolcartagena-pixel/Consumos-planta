# Lecturas de planta

PWA para registrar de forma compartida las lecturas diarias de consumos de Masol Cartagena. Dos operarios pueden completar la misma fecha desde dispositivos distintos: Power Automate guarda un archivo JSON por fecha en SharePoint y el segundo dispositivo recupera los valores ya introducidos.

Cada lectura dispone de un botón de cámara. La aplicación envía temporalmente la fotografía a Power Automate, AI Builder reconoce los dígitos y el operario revisa el valor antes de incorporarlo al formulario. La introducción manual permanece siempre disponible y las fotografías no se guardan en SharePoint.

El envío definitivo se habilita al completar las 37 lecturas obligatorias. Se pueden dejar vacíos `Contador Desc. Entrada`, `Contador Desc. Salida`, `Contador Reactor Biológico`, `Cfv`, `Contador Gas Natural` y `Totalizador 390`; el Office Script no modifica sus celdas cuando se omiten. Power Automate ejecuta directamente el script sobre `Datos-Planta-Mes-Año.xlsx` y escribe únicamente en la hoja `Mes Año`.

## Configuración

1. Sigue [docs/power-automate.md](docs/power-automate.md) para crear la lista, el flujo y el Office Script.
2. Copia la URL del disparador HTTP en `config.js`.
3. Publica la rama `main` con GitHub Pages.

La aplicación no guarda lecturas en el dispositivo. Solo recuerda el nombre del operario; las medidas viven en archivos JSON de SharePoint hasta que el flujo completa la escritura en Excel.

## Archivos principales

- `app.js`: captura manual o fotográfica, reconocimiento OCR, carga y guardado compartido, validación y envío JSON.
- `config.js`: URL del flujo y retardo del guardado automático.
- `office-scripts/EscribirLecturasPlanta.ts`: escritura y reparto de consumos.
- `docs/power-automate.md`: guía completa de implementación y pruebas.

## Seguridad

No reutilices la firma HTTP que estuvo publicada anteriormente. Regenera el disparador y limita el acceso a usuarios corporativos antes de usar datos reales. Una URL firmada dentro de una web estática no debe considerarse un secreto.
