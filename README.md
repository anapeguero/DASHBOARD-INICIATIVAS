# Procurement & CAPEX Projects — Grupo Ramos

Dashboard web estático preparado para GitHub Pages.

## Estructura

- `index.html` — aplicación.
- `style.css` — diseño.
- `script.js` — lógica, lectura de Excel y visualizaciones.
- `assets/` — logos.
- `data/projects.json` — catálogo de proyectos disponibles.
- `data/projects/mao.xlsx` — plantilla fuente de Mao.

## Páginas incluidas

1. Resumen Ejecutivo
2. Presupuesto & Compras
3. Seguimiento / Pipeline de Procurement
4. Cronograma
5. Importaciones
6. Pendientes
7. Pre-Liquidación

## Cómo subir a GitHub Pages

Sube **el contenido de esta carpeta**, no el ZIP:
`index.html`, `style.css`, `script.js`, las carpetas `assets` y `data`, y `README.md`.

Configura GitHub Pages usando la rama `main` y carpeta `(root)`.

## Diseño para varios proyectos

La aplicación no tiene a Mao “amarrado” al diseño. Mao es un proyecto configurado en `data/projects.json`.

Para agregar otro proyecto compatible:
1. Copia su Excel dentro de `data/projects/`.
2. Agrega una entrada a `data/projects.json`.
3. Mantén los nombres/estructura base de las hojas utilizadas por la plantilla.

Ejemplo:

```json
{
  "id": "hispanoamerica",
  "name": "Hispanoamérica",
  "displayName": "Proyecto Hispanoamérica",
  "file": "data/projects/hispanoamerica.xlsx",
  "brand": "Sirena",
  "status": "En ejecución"
}
```

## Hojas de Mao utilizadas

- `Plan OI`
- `BD Plan Detallado`
- `Cronograma Inverso`
- `Seguimiento Mao `
- `Puntos pendientes`
- `Pre-Liquidar`

La web usa SheetJS para leer el Excel directamente en el navegador y Chart.js para las visualizaciones.

## Consideración importante

Las fechas del archivo Mao incluyen registros históricos que, comparados con la fecha actual, pueden aparecer vencidos. La web no altera la fuente: muestra una alerta de calidad de datos para que el cronograma se actualice antes de usarlo como referencia de fechas futuras.
