# RPA-CBN

Automatización RPA para el procesamiento de cotizaciones en CBN.

## Stack principal

- Python
- Playwright
- Pandas
- OpenPyXL
- CustomTkinter
- PyInstaller

## Estructura principal

- `src/browser/` - sesión y autenticación del navegador
- `src/flows/` - pasos de automatización del proceso CBN
- `src/excel/` - lectura, validación y escritura de Excel
- `src/gui/` - interfaz gráfica
- `src/core/` - orquestación de la ejecución
- `src/utils/` - utilidades compartidas
- `scripts/` - herramientas auxiliares
- `assets/` - recursos visuales

## Nota

Los archivos de ejecución, perfiles del navegador, logs, datos operativos,
temporales y configuraciones locales no se almacenan en Git.
