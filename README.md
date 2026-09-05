# Procurement & CAPEX Projects — Grupo Ramos (v4)

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


## Cambios v2

- Nueva página **Inicio** que explica qué contiene el dashboard y cómo interpretarlo.
- **Cuenta regresiva** en Resumen Ejecutivo hacia la apertura de Mao: **15 de noviembre de 2026**.
- Nueva página **Flujo de Caja** alimentada por la hoja `Flujo de Caja`.
- **Recordatorios automáticos de pagos** generados desde `Per.P1` a `Per.P5` y `Pago 1` a `Pago 5` de `BD Plan Detallado`.
- Los reminders identifican el mes, monto, cantidad de pagos y si corresponde al mes actual o a próximos meses.
- Configuración de fechas por proyecto en `data/projects.json`:
  - `openingDate`
  - `cashflowStart`
- Esto permite replicar la misma lógica para futuros proyectos sin amarrar las fechas a Mao dentro del código.


## Cambios v3

- Se eliminó completamente la página **Cronograma** de la navegación.
- También se retiraron referencias visibles a Cronograma en Inicio, Project Health y Calidad de Datos.
- El gráfico principal del Resumen Ejecutivo ahora muestra exclusivamente las **partidas maestras 1, 2, 3 y 4** del `Plan OI`:
  1. Estudios Preliminares
  2. Obra Civil
  3. Sistema Eléctrico
  4. Sistema de Climatización
- Los valores del gráfico se leen directamente de las filas maestras de `Plan OI`, por lo que se actualizan cuando cambie la plantilla.


## Cambios v4 — Descuentos

- Se agregó visibilidad de los **descuentos negociados** registrados en `Seguimiento Mao`.
- Resumen Ejecutivo muestra:
  - Descuento acumulado.
  - Cantidad de órdenes con descuento.
  - Porcentaje de descuento sobre el subtotal de esas órdenes.
- Importaciones muestra un nuevo KPI **Descuentos obtenidos**.
- Nuevo gráfico **Top proveedores por descuento**.
- La tabla logística ahora incluye:
  - Subtotal antes de descuento.
  - Descuento.
  - Total después del descuento.
  - Valor final en DOP.
- Los descuentos en USD/EUR se convierten a DOP usando las tasas registradas en la propia hoja de seguimiento.
