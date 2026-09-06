# Dashboard de Seguimiento de Iniciativas — Grupo Ramos

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


## Cambios v5 — versión ejecutiva

- Se eliminó la pestaña **Importaciones**.
- Los ahorros se calculan exclusivamente desde **Plan OI**:
  - Plan detallado Aprobado = presupuesto original.
  - Plan Detallado en ejecución = presupuesto vigente.
  - Ahorro = aprobado − vigente.
- Resumen Ejecutivo reforzado con visuales de decisión:
  - concentración del presupuesto vigente por partidas 1–9;
  - aprobado vs vigente por partida;
  - ahorro / incremento por partida;
  - avance financiero ponderado por monto;
  - pipeline de compras medido por valor, no por cantidad de filas.
- Se corrigió el formato de ejes monetarios de gráficos horizontales.
- Flujo de Caja:
  - selector DOP / moneda original;
  - monto original y equivalente DOP por pago;
  - pagos anteriores al corte separados de los próximos;
  - los gráficos permanecen consolidados en DOP para comparabilidad.
- La plantilla indica que `Per.P1` inicia en **mayo de 2023**; se corrigió el calendario del proyecto Mao.
- Pre-Liquidación incluye una explicación ejecutiva sobre su propósito y uso.


## Cambios v6

- Se cambió el nombre principal de la solución a **Dashboard de Seguimiento de Iniciativas**.
- El logo principal visible en el encabezado ahora es **Sirena**.
- Se mantiene la arquitectura multiproyecto para que Mao sea el proyecto inicial y puedan incorporarse otras iniciativas posteriormente.


## Cambios v7

- Se corrigió el cargado de partidas para usar **únicamente las partidas maestras 1.00 a 9.00**.
- Se excluyen subpartidas como 1.01, 2.06, 5.03, 7.03, etc. de los tres gráficos ejecutivos del Resumen.
- Los gráficos de presupuesto quedan alineados con la lectura ejecutiva:
  - Presupuesto vigente por partida 1.00–9.00.
  - Aprobado vs vigente por partida 1.00–9.00.
  - Ahorro / incremento por partida 1.00–9.00.
- Flujo de Caja rediseñado con visuales ejecutivos:
  - necesidad de caja por mes;
  - programado hasta hoy vs por venir;
  - composición de compromisos por moneda;
  - flujo acumulado;
  - próximos meses que requieren preparación financiera.
- Los gráficos del flujo se consolidan en DOP para comparabilidad, mientras la tabla mantiene moneda original y equivalente en pesos.


## Cambios v8 — Integración Hispanoamérica

- Se agregó **Hispanoamérica** como segundo proyecto seleccionable.
- El selector superior ahora carga:
  - Mao
  - Hispanoamérica
- Cada proyecto utiliza su propio Excel dentro de `data/projects/`.
- Se agregó configuración por proyecto para soportar diferencias de estructura en `Plan OI`.
- Hispanoamérica usa:
  - **Alto Nivel Rev III** como presupuesto base de comparación.
  - **Plan Detallado** como presupuesto vigente.
  - Flujo de caja configurado de **enero 2025 a diciembre 2026**.
- La plantilla de Hispanoamérica indica apertura en **diciembre 2026**, pero no un día exacto; el dashboard no inventa una fecha y muestra que el día está pendiente.
- `Consideraciones` de Hispanoamérica se muestra como contexto del proyecto sin convertirlo artificialmente en pendientes con responsables.
- Si un proyecto no tiene hoja `Pre-Liquidar`, el estimador lo advierte expresamente.


## Cambio v9

- Fecha exacta de apertura de **Hispanoamérica: 15 de diciembre de 2026**. La cuenta regresiva ya utiliza esta fecha.
