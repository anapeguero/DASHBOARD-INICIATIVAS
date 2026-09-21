# Compras Estratégicas CAF — GitHub Pages

## Archivos que deben estar juntos en la raíz del repositorio
- `index.html`
- `styles.css`
- `app.js`
- `Data para plantilla modelo R6.xlsm`

## Cómo actualizar el dashboard
1. Mantén el mismo nombre del archivo Excel: `Data para plantilla modelo R6.xlsm`.
2. En GitHub, reemplaza únicamente ese Excel por la versión nueva.
3. No necesitas modificar `app.js`, `index.html` ni convertir el Excel a JavaScript.
4. Espera el despliegue de GitHub Pages y recarga la página.

## Fuentes dentro del Excel
- `Fechas`: cantidad estratégica de tiendas, formatos, AF y pipeline. Los textos como `X 30` o `10 tiendas` se interpretan como cantidades.
- `Plan `: tipo de proyecto, responsables, fechas y roadmap de remodelaciones.
- `Resumen estratégico`: objetivo, pilares, estrategias, acciones, valor para el negocio y métricas.
- `Seguimiento`: hoja de ruta estratégica y fechas objetivo.
- `Data`: CAPEX, agrupadores, supranúmeros y presupuesto.

## PDF
El botón **Descargar reporte PDF** genera el reporte completo. Si el navegador bloquea la librería de PDF, se abre automáticamente la impresión para usar **Guardar como PDF**.

> La página usa SheetJS y html2pdf desde CDN, por lo que el navegador necesita acceso a Internet al abrir el dashboard.
