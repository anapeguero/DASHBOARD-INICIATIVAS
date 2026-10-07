/* =========================================================
DASHBOARD DE SEGUIMIENTO DE INICIATIVAS
Grupo Ramos
========================================================= */

const state = {
projects: [],
project: null,
workbook: null,

// Fuentes principales
planOI: [],
budget: [],

cashflow: [],
issues: [],
considerations: [],

approvedBudget: 0,
currentBudget: 0,
orderedValue: 0,
savings: 0,

charts: {},
exchangeRate: 63
};

const $ = id => document.getElementById(id);

const COLORS = {
blue: "#064c8c",
blue2: "#0b6fb8",
cyan: "#28a9e0",
green: "#16a05d",
yellow: "#f4b323",
red: "#d94a4a",
gray: "#cbd7e2",
ink: "#18324b"
};

document.addEventListener("DOMContentLoaded", init);


/* =========================================================
INICIO
========================================================= */

async function init() {

setupNavigation();
setupBudgetFilters();
setupLiquidation();
setupCashflowCurrency();

await loadProjects();

}


/* =========================================================
CARGAR LISTA DE PROYECTOS
========================================================= */

async function loadProjects() {

try {

const response = await fetch(
"data/projects.json?v=" + Date.now()
);

if (!response.ok) {
throw new Error(
"No se pudo cargar data/projects.json"
);
}

const config = await response.json();

state.projects = Array.isArray(config.projects)
? config.projects
: [];

if (!state.projects.length) {
throw new Error(
"projects.json no contiene proyectos."
);
}

setupProjectSelector();

const saved =
localStorage.getItem("selectedProject");

const initial =
state.projects.find(
project => project.id === saved
) ||
state.projects[0];

$("projectSelect").value = initial.id;

await loadProject(initial);

} catch (error) {

console.error(error);

showLoader(false);

alert(
"No se pudo cargar la lista de iniciativas. Revisa data/projects.json."
);

}

}


/* =========================================================
SELECTOR DE PROYECTOS
========================================================= */

function setupProjectSelector() {

const select = $("projectSelect");

if (!select) return;

select.innerHTML = "";

state.projects.forEach(project => {

const option =
document.createElement("option");

option.value = project.id;
option.textContent =
project.name || project.id;

select.appendChild(option);

});

select.onchange = async event => {

const project =
state.projects.find(
item =>
item.id === event.target.value
);

if (!project) return;

localStorage.setItem(
"selectedProject",
project.id
);

await loadProject(project);

};

}


/* =========================================================
NAVEGACIÓN
========================================================= */

function setupNavigation() {

document
.querySelectorAll("#navTabs button")
.forEach(button => {

button.addEventListener(
"click",
() => {
openPage(
button.dataset.page
);
}
);

});

document
.querySelectorAll("[data-go-page]")
.forEach(button => {

button.addEventListener(
"click",
() => {
openPage(
button.dataset.goPage
);
}
);

});

}


function openPage(page) {

document
.querySelectorAll(".page")
.forEach(section => {
section.classList.remove(
"active"
);
});

document
.querySelectorAll(
"#navTabs button"
)
.forEach(button => {
button.classList.remove(
"active"
);
});

const target =
$(`page-${page}`);

if (target) {
target.classList.add(
"active"
);
}

const button =
document.querySelector(
`#navTabs button[data-page="${page}"]`
);

if (button) {
button.classList.add(
"active"
);
}

window.scrollTo({
top: 0,
behavior: "smooth"
});

}


/* =========================================================
CARGAR EXCEL DEL PROYECTO
========================================================= */

async function loadProject(project) {

showLoader(true);

state.project = project;

updateProjectLabels(project);

try {

const response =
await fetch(
project.file +
"?v=" +
Date.now()
);

if (!response.ok) {
throw new Error(
`No se encontró ${project.file}`
);
}

const buffer =
await response.arrayBuffer();

state.workbook =
XLSX.read(
buffer,
{
type: "array",
cellDates: true
}
);

parseWorkbook();

renderAll();

} catch (error) {

console.error(error);

state.planOI = [];
state.budget = [];
state.cashflow = [];
state.issues = [];
state.considerations = [];

renderEmptyState();

}

showLoader(false);

}


/* =========================================================
NOMBRES DEL PROYECTO
========================================================= */

function updateProjectLabels(project) {

const name =
project.name ||
project.id ||
"Proyecto";

const brand =
project.brand ||
"Sirena";

setText(
"projectBrand",
brand
);

setText(
"projectName",
name
);

setText(
"homeProjectName",
name
);

setText(
"overviewTitle",
`${brand} ${name}`
);

setText(
"countdownProject",
name
);

setText(
"footerProject",
`Proyecto ${name}`
);

/*
Si projects.json tiene openingDate,
usamos la fecha.

Si no tiene fecha, el dashboard
sigue funcionando normalmente.
*/

if (project.openingDate) {

const opening =
parseDate(
project.openingDate
);

if (opening) {

setText(
"homeOpeningDate",
formatDate(opening)
);

setText(
"countdownDateLabel",
formatLongDate(opening)
);

startCountdown(
project.openingDate
);

}

} else {

setText(
"homeOpeningDate",
"Por definir"
);

setText(
"countdownDateLabel",
"Por definir"
);

setText(
"cdDays",
"—"
);

setText(
"cdHours",
"—"
);

setText(
"cdMinutes",
"—"
);

setText(
"cdSeconds",
"—"
);

setText(
"countdownStatus",
"Fecha de apertura no definida"
);

}

}


/* =========================================================
LEER EXCEL

IMPORTANTE:

Plan OI
= resumen financiero / presupuesto / ahorro

BD Plan Detallado
= detalle de compras / OC / proveedores /
pagos / seguimiento
========================================================= */

function parseWorkbook() {

const wb =
state.workbook;

state.planOI = [];
state.budget = [];
state.cashflow = [];
state.issues = [];
state.considerations = [];

/* =========================
PLAN OI
========================= */

const planOISheet =
findExactSheet(
wb,
"Plan OI"
);

if (planOISheet) {

state.planOI =
smartSheetToObjects(
planOISheet
);

} else {

console.warn(
"No se encontró la hoja Plan OI"
);

}


/* =========================
BD PLAN DETALLADO
========================= */

const bdSheet =
findExactSheet(
wb,
"BD Plan Detallado"
);

if (bdSheet) {

state.budget =
smartSheetToObjects(
bdSheet
);

} else {

console.warn(
"No se encontró la hoja BD Plan Detallado"
);

}


/* =========================
OTRAS HOJAS OPCIONALES
========================= */

const issueSheet =
findSheet(
wb,
[
"Pendientes",
"Issues",
"Seguimiento"
]
);

if (issueSheet) {

state.issues =
smartSheetToObjects(
issueSheet
);

}


const considerationsSheet =
findSheet(
wb,
[
"Consideraciones",
"Notas"
]
);

if (considerationsSheet) {

state.considerations =
smartSheetToObjects(
considerationsSheet
);

}


/* =========================
NORMALIZACIÓN
========================= */

normalizePlanOI();
normalizeBudget();
normalizeIssues();

calculateTotals();

}


/* =========================================================
ENCONTRAR HOJAS
========================================================= */

function findExactSheet(
workbook,
requestedName
) {

const found =
workbook.SheetNames.find(
name =>
normalizeText(name) ===
normalizeText(requestedName)
);

return found
? workbook.Sheets[found]
: null;

}


function findSheet(
workbook,
candidates
) {

for (
const candidate
of candidates
) {

const exact =
workbook.SheetNames.find(
name =>
normalizeText(name) ===
normalizeText(candidate)
);

if (exact) {
return workbook.Sheets[
exact
];
}

}

for (
const candidate
of candidates
) {

const partial =
workbook.SheetNames.find(
name =>
normalizeText(name)
.includes(
normalizeText(
candidate
)
)
);

if (partial) {
return workbook.Sheets[
partial
];
}

}

return null;

}


/* =========================================================
CONVERTIR HOJA A OBJETOS
Detecta automáticamente la fila de encabezados
========================================================= */

function smartSheetToObjects(sheet) {

const rows =
XLSX.utils.sheet_to_json(
sheet,
{
header: 1,
defval: "",
raw: false
}
);

if (!rows.length) {
return [];
}

let headerIndex = 0;
let bestScore = -1;

const keywords = [
"partida",
"item",
"orden",
"proveedor",
"plan",
"total",
"ppto",
"valor",
"mon",
"pago",
"per.p"
];

const limit =
Math.min(
rows.length,
20
);

for (
let i = 0;
i < limit;
i++
) {

const row =
rows[i];

let score = 0;

row.forEach(cell => {

const text =
normalizeText(cell);

keywords.forEach(
keyword => {

if (
text.includes(
normalizeText(
keyword
)
)
) {
score++;
}

}
);

});

if (score > bestScore) {

bestScore = score;
headerIndex = i;

}

}

const headers =
makeUniqueHeaders(
rows[headerIndex]
);

const data = [];

for (
let i =
headerIndex + 1;

i < rows.length;

i++
) {

const row =
rows[i];

const object = {};

let hasData = false;

headers.forEach(
(header, index) => {

if (!header) return;

const value =
row[index] ?? "";

object[header] =
value;

if (
String(value)
.trim() !== ""
) {
hasData = true;
}

}
);

if (hasData) {
data.push(object);
}

}

return data;

}


function makeUniqueHeaders(row) {

const used = {};

return row.map(
(value, index) => {

let name =
String(
value || ""
).trim();

if (!name) {
name =
`Columna_${index + 1}`;
}

if (
used[name] !== undefined
) {

used[name]++;

name =
`${name}_${used[name]}`;

} else {

used[name] = 1;

}

return name;

}
);

}


/* =========================================================
NORMALIZAR PLAN OI
========================================================= */

function normalizePlanOI() {

state.planOI =
state.planOI
.map(
(row, index) => {

const partida =
valueFrom(
row,
[
"Partida",
"Partidas",
"Descripción",
"Descripcion",
"Concepto"
]
) ||
`Partida ${index + 1}`;


/*
PLAN OI puede tener encabezados
repetidos de Total USD/DOP.

Por eso buscamos primero
los nombres más específicos
y luego utilizamos detección
por posición/encabezado.
*/

const financial =
extractPlanOIFinancials(
row
);

return {

raw: row,

partida,

approved:
financial.approved,

current:
financial.current,

savings:
financial.savings

};

}
)
.filter(
row =>
row.approved ||
row.current ||
row.savings ||
(
row.partida &&
!row.partida.startsWith(
"Partida "
)
)
);

}


/* =========================================================
EXTRAER VALORES FINANCIEROS PLAN OI
========================================================= */

function extractPlanOIFinancials(row) {

const entries =
Object.entries(row);

let approved = 0;
let current = 0;
let savings = 0;


/* AHORRO */

savings =
numberFrom(
row,
[
"Ahorro según ejecución",
"Ahorro segun ejecucion",
"Ahorro",
"Savings"
]
);


/*
BUSCAR COLUMNAS DOP
*/

const dopValues = [];

entries.forEach(
([key, value]) => {

const normalized =
normalizeText(key);

if (
normalized.includes(
"totaldop"
) ||
normalized.includes(
"plandop"
) ||
normalized.includes(
"presupuestodop"
)
) {

const n =
toNumber(value);

if (
Number.isFinite(n)
) {
dopValues.push(n);
}

}

}
);


/*
Si existen dos columnas DOP:
primera = aprobado
segunda = vigente.

Si existe una sola, se utiliza
como vigente.
*/

if (
dopValues.length >= 2
) {

approved =
dopValues[0];

current =
dopValues[
dopValues.length - 1
];

} else if (
dopValues.length === 1
) {

current =
dopValues[0];

}


/*
Nombres específicos si existen
*/

const approvedNamed =
numberFrom(
row,
[
"Plan detallado Aprobado",
"Plan Detallado Aprobado",
"Presupuesto Aprobado",
"Aprobado DOP"
]
);

const currentNamed =
numberFrom(
row,
[
"Plan Detallado en ejecución",
"Plan Detallado en ejecucion",
"Presupuesto Vigente",
"Plan Vigente"
]
);


if (approvedNamed) {
approved =
approvedNamed;
}

if (currentNamed) {
current =
currentNamed;
}


/*
Si tenemos vigente + ahorro,
podemos reconstruir aprobado.
*/

if (
!approved &&
current &&
savings
) {

approved =
current + savings;

}


/*
Si tenemos aprobado + ahorro,
reconstruimos vigente.
*/

if (
approved &&
!current &&
savings
) {

current =
approved - savings;

}


return {
approved,
current,
savings
};

}


/* =========================================================
NORMALIZAR BD PLAN DETALLADO
========================================================= */

function normalizeBudget() {

state.budget =
state.budget
.map(
(row, index) => {

const partida =
valueFrom(
row,
[
"Partidas",
"Partida",
"Gran Partida"
]
) ||
"Sin partida";


const capexType =
valueFrom(
row,
[
"OI / Capex Tipo Ppto",
"OI Capex Tipo Ppto",
"Capex Tipo Ppto",
"OI"
]
);


const plan =
valueFrom(
row,
[
"Plan"
]
);


const item =
valueFrom(
row,
[
"Item",
"Ítem",
"Activo",
"Descripción",
"Descripcion"
]
) ||
`Ítem ${index + 1}`;


const po =
valueFrom(
row,
[
"Orden de compra",
"Orden de Compra",
"OC",
"PO"
]
);


const date =
valueFrom(
row,
[
"Fecha",
"Fecha OC"
]
);


const quantity =
numberFrom(
row,
[
"Cant.",
"Cant",
"Cantidad"
]
);


const currency =
valueFrom(
row,
[
"Mon",
"Moneda",
"Currency"
]
) ||
"DOP";


const rate =
numberFrom(
row,
[
"Tasa",
"Exchange Rate"
]
);


const unitPrice =
numberFrom(
row,
[
"Precio Und",
"Precio Unitario",
"Precio"
]
);


const netValue =
numberFrom(
row,
[
"Valor Neto [DOP]",
"Valor Neto DOP",
"Valor Neto"
]
);


const liquidation =
numberFrom(
row,
[
"Liquid.",
"Liquidación",
"Liquidacion"
]
);


const internalBudgetDOP =
numberFrom(
row,
[
"Ppto. Orden Interna [DOP]",
"Ppto Orden Interna DOP",
"Presupuesto Orden Interna DOP"
]
);


const internalBudgetUSD =
numberFrom(
row,
[
"Ppto. Orden Interna [USD]",
"Ppto Orden Interna USD",
"Presupuesto Orden Interna USD"
]
);


const comment =
valueFrom(
row,
[
"Comentario",
"Comentarios",
"Observación",
"Observacion"
]
);


const supplier =
valueFrom(
row,
[
"Proveedor",
"Supplier"
]
);


/*
Para el dashboard:

Presupuesto = Ppto Orden Interna DOP

Ejecutado/comprometido por OC =
Valor Neto DOP cuando existe OC
*/

const budget =
internalBudgetDOP ||
netValue;


const ordered =
po
? netValue
: 0;


return {

raw: row,

partida,
capexType,
plan,
item,
po,
date,
quantity,
currency,
rate,
unitPrice,
netValue,
liquidation,
internalBudgetDOP,
internalBudgetUSD,
comment,
supplier,

current:
budget,

ordered,

status:
determinePurchaseStatus(
po,
ordered,
supplier
)

};

}
)
.filter(
row =>
row.item ||
row.current ||
row.netValue ||
row.po
);

}


/* =========================================================
ESTADO DE COMPRA
========================================================= */

function determinePurchaseStatus(
po,
ordered,
supplier
) {

if (
po ||
ordered > 0
) {

return "Con OC";

}

if (!supplier) {

return "Sin proveedor";

}

return "Pendiente de compra";

}


/* =========================================================
NORMALIZAR PENDIENTES
========================================================= */

function normalizeIssues() {

state.issues =
state.issues.map(
(row, index) => ({

title:
valueFrom(
row,
[
"Pendiente",
"Tema",
"Issue",
"Actividad",
"Descripción",
"Descripcion"
]
) ||
`Pendiente ${index + 1}`,

owner:
valueFrom(
row,
[
"Responsable",
"Owner",
"Líder",
"Lider"
]
) ||
"Sin responsable",

comment:
valueFrom(
row,
[
"Comentario",
"Comentarios",
"Observación",
"Observacion"
]
) || "",

status:
valueFrom(
row,
[
"Estado",
"Status"
]
) ||
"Abierto"

})
);

}


/* =========================================================
TOTALES

PLAN OI = presupuesto ejecutivo
BD = órdenes de compra
========================================================= */

function calculateTotals() {

/*
Primero intentamos obtener presupuesto
desde Plan OI.
*/

let approved =
sum(
state.planOI.map(
row =>
row.approved
)
);

let current =
sum(
state.planOI.map(
row =>
row.current
)
);

let savings =
sum(
state.planOI.map(
row =>
row.savings
)
);


/*
Si Plan OI no devuelve vigente,
utilizamos BD Plan Detallado.
*/

if (!current) {

current =
sum(
state.budget.map(
row =>
row.internalBudgetDOP
)
);

}


/*
Si aprobado no está explícito
pero tenemos vigente y ahorro.
*/

if (
!approved &&
current
) {

approved =
current +
savings;

}


/*
Si ahorro no está explícito.
*/

if (
!savings &&
approved
) {

savings =
approved -
current;

}


/*
Valor con OC sale del BD.
*/

const ordered =
sum(
state.budget
.filter(
row =>
row.po
)
.map(
row =>
row.netValue
)
);


state.approvedBudget =
approved;

state.currentBudget =
current;

state.orderedValue =
ordered;

state.savings =
savings;

}


/* =========================================================
RENDER GENERAL
========================================================= */

function renderAll() {

renderOverview();
renderBudget();
renderCashflow();
renderTracking();
renderPending();
renderLiquidation();

}


/* =========================================================
RESUMEN EJECUTIVO
========================================================= */

function renderOverview() {

const approved =
state.approvedBudget;

const current =
state.currentBudget;

const ordered =
state.orderedValue;

const pending =
Math.max(
current - ordered,
0
);

const saving =
state.savings ||
(
approved -
current
);

const savingPct =
approved
? saving / approved
: 0;

const progress =
current
? ordered / current
: 0;


setText(
"kpiApprovedBudget",
money(approved)
);

setText(
"kpiBudget",
money(current)
);

setText(
"kpiOrdered",
money(ordered)
);

setText(
"kpiPendingBuy",
money(pending)
);

setText(
"kpiWeightedProgress",
percent(progress)
);

setText(
"kpiOpenIssues",
state.issues.length
);

setText(
"overviewDiscount",
money(saving)
);

setText(
"overviewDiscountPct",
percent(savingPct)
);


renderHealth(
progress,
pending
);

renderBudgetMix();
renderBudgetEvolution();
renderSavingsByPart();
renderPurchaseStatus();
renderTopSuppliers();
renderDataQuality();
renderAttentionList();

}


/* =========================================================
PROJECT HEALTH
========================================================= */

function renderHealth(
progress,
pending
) {

const container =
$("healthStrip");

if (!container) return;

const withPO =
state.budget.filter(
row =>
row.status ===
"Con OC"
).length;

const total =
state.budget.length;

const coverage =
total
? withPO / total
: 0;

const noSupplier =
state.budget.filter(
row =>
!row.supplier
).length;


const health = [

{
title:
"Avance financiero",

value:
percent(progress),

text:
progress >= .75
? "Buen nivel de colocación"
: "Existe presupuesto pendiente",

color:
progress >= .75
? "good"
: "warn"
},

{
title:
"Cobertura de compras",

value:
percent(coverage),

text:
`${withPO} de ${total} ítems con OC`,

color:
coverage >= .75
? "good"
: "warn"
},

{
title:
"Pendiente de compra",

value:
money(pending),

text:
"Presupuesto todavía no colocado",

color:
pending <=
state.currentBudget * .2
? "good"
: "warn"
},

{
title:
"Sin proveedor",

value:
String(noSupplier),

text:
"Ítems sin proveedor definido",

color:
noSupplier
? "bad"
: "good"
}

];


container.innerHTML =
health.map(
item => `

<div class="health-item">

<div class="row">

<b>
${escapeHtml(
item.title
)}
</b>

<i class="dot ${item.color}">
</i>

</div>

<strong>
${escapeHtml(
item.value
)}
</strong>

<p>
${escapeHtml(
item.text
)}
</p>

</div>

`
).join("");

}


/* =========================================================
GRÁFICA PRESUPUESTO POR PARTIDA
========================================================= */

function renderBudgetMix() {

const source =
state.planOI.some(
row => row.current
)
? state.planOI
: state.budget;


const grouped =
groupSum(
source,
"partida",
"current"
);


createChart(
"budgetMixChart",
"doughnut",
grouped.labels,
grouped.values,
{
plugins: {
legend: {
position: "bottom"
}
}
}
);

}


/* =========================================================
APROBADO VS VIGENTE
========================================================= */

function renderBudgetEvolution() {

const groups = {};

state.planOI.forEach(
row => {

if (!groups[row.partida]) {

groups[row.partida] = {
approved: 0,
current: 0
};

}

groups[
row.partida
].approved +=
row.approved;

groups[
row.partida
].current +=
row.current;

}
);


const entries =
Object.entries(groups)
.filter(
([, values]) =>
values.approved ||
values.current
)
.sort(
(a, b) =>
b[1].current -
a[1].current
)
.slice(0, 15);


createMultiChart(
"budgetEvolutionChart",
"bar",
entries.map(
x => x[0]
),
[

{
label:
"Aprobado",

data:
entries.map(
x =>
x[1].approved
),

backgroundColor:
COLORS.blue
},

{
label:
"Vigente",

data:
entries.map(
x =>
x[1].current
),

backgroundColor:
COLORS.cyan
}

]
);

}


/* =========================================================
AHORRO POR PARTIDA
========================================================= */

function renderSavingsByPart() {

const groups = {};

state.planOI.forEach(
row => {

const saving =
row.savings ||
(
row.approved -
row.current
);

groups[row.partida] =
(
groups[row.partida] ||
0
) +
saving;

}
);


const entries =
Object.entries(groups)
.filter(
([, value]) =>
value !== 0
)
.sort(
(a, b) =>
Math.abs(b[1]) -
Math.abs(a[1])
)
.slice(0, 15);


createChart(
"savingsByPartChart",
"bar",
entries.map(
x => x[0]
),
entries.map(
x => x[1]
),
{
indexAxis: "y",

plugins: {
legend: {
display: false
}
}
}
);

}


/* =========================================================
ESTADO DE COMPRAS
========================================================= */

function renderPurchaseStatus() {

const withPO =
state.budget.filter(
row =>
row.status ===
"Con OC"
).length;

const pending =
state.budget.length -
withPO;


createChart(
"purchaseStatusChart",
"doughnut",
[
"Con OC",
"Pendiente"
],
[
withPO,
pending
]
);

}


/* =========================================================
TOP PROVEEDORES
========================================================= */

function renderTopSuppliers() {

const container =
$("topSuppliers");

if (!container) return;

const groups = {};


state.budget.forEach(
row => {

if (!row.supplier) {
return;
}

groups[row.supplier] =
(
groups[
row.supplier
] ||
0
) +
(
row.netValue ||
row.current
);

}
);


const entries =
Object.entries(groups)
.sort(
(a, b) =>
b[1] - a[1]
)
.slice(0, 6);


container.innerHTML =
entries.length
? entries.map(
(
[supplier, value],
index
) => `

<div class="rank-row">

<span class="n">
${index + 1}
</span>

<b>
${escapeHtml(
supplier
)}
</b>

<span>
${money(value)}
</span>

</div>

`
).join("")
: emptyMessage(
"No hay proveedores disponibles."
);

}


/* =========================================================
CALIDAD DE DATOS
========================================================= */

function renderDataQuality() {

const container =
$("dataQuality");

if (!container) return;


const total =
state.budget.length ||
1;


const missingSupplier =
state.budget.filter(
row =>
!row.supplier
).length;


const missingPO =
state.budget.filter(
row =>
!row.po
).length;


const missingDate =
state.budget.filter(
row =>
!row.date
).length;


const items = [

{
label:
"Sin proveedor",

value:
missingSupplier
},

{
label:
"Sin OC",

value:
missingPO
},

{
label:
"Sin fecha",

value:
missingDate
}

];


container.innerHTML =
items.map(
item => `

<div class="quality">

<div class="top">

<b>
${item.label}
</b>

<strong>
${item.value}
</strong>

</div>

<small>

${percent(
item.value /
total
)}

de los ítems

</small>

</div>

`
).join("");

}


/* =========================================================
PRIORIDADES
========================================================= */

function renderAttentionList() {

const container =
$("attentionList");

if (!container) return;


const pending =
[...state.budget]
.filter(
row =>
row.status !==
"Con OC"
)
.sort(
(a, b) =>
b.current -
a.current
)
.slice(0, 6);


container.innerHTML =
pending.length
? pending.map(
row => `

<div class="attention">

<i class="dot warn">
</i>

<div>

<b>
${escapeHtml(
row.item
)}
</b>

<span>
${escapeHtml(
row.partida
)}
·
${money(
row.current
)}
</span>

</div>

</div>

`
).join("")
: emptyMessage(
"No se identifican compras pendientes."
);

}


/* =========================================================
FILTROS PRESUPUESTO
========================================================= */

function setupBudgetFilters() {

[
"budgetCategoryFilter",
"budgetStatusFilter",
"budgetSupplierFilter"
].forEach(id => {

const element =
$(id);

if (element) {

element.addEventListener(
"change",
renderBudgetTable
);

}

});


const reset =
$("budgetReset");

if (reset) {

reset.addEventListener(
"click",
() => {

$("budgetCategoryFilter").value =
"";

$("budgetStatusFilter").value =
"";

$("budgetSupplierFilter").value =
"";

renderBudgetTable();

}
);

}

}


/* =========================================================
PRESUPUESTO Y COMPRAS
========================================================= */

function renderBudget() {

populateBudgetFilters();
renderBudgetTable();


const category =
groupSum(
state.budget,
"partida",
"current"
);


createChart(
"categoryChart",
"bar",
category.labels,
category.values,
{
indexAxis: "y",

plugins: {
legend: {
display: false
}
}
}
);


const supplier =
groupSupplierValue();


createChart(
"supplierChart",
"bar",
supplier.labels,
supplier.values,
{
plugins: {
legend: {
display: false
}
}
}
);

}


function populateBudgetFilters() {

fillSelect(
"budgetCategoryFilter",
unique(
state.budget.map(
row =>
row.partida
)
)
);


fillSelect(
"budgetSupplierFilter",
unique(
state.budget
.map(
row =>
row.supplier
)
.filter(Boolean)
)
);

}


function fillSelect(
id,
values
) {

const select =
$(id);

if (!select) return;


const current =
select.value;


const first =
select.options[0]
? select.options[0]
.outerHTML
: '<option value="">Todos</option>';


select.innerHTML =
first;


values
.sort(
(a, b) =>
String(a)
.localeCompare(
String(b),
"es"
)
)
.forEach(
value => {

const option =
document.createElement(
"option"
);

option.value =
value;

option.textContent =
value;

select.appendChild(
option
);

}
);


if (
[...select.options]
.some(
option =>
option.value ===
current
)
) {

select.value =
current;

}

}


/* =========================================================
TABLA PRESUPUESTO
========================================================= */

function renderBudgetTable() {

const category =
$("budgetCategoryFilter")
?.value || "";

const status =
$("budgetStatusFilter")
?.value || "";

const supplier =
$("budgetSupplierFilter")
?.value || "";


let rows =
[...state.budget];


if (category) {

rows =
rows.filter(
row =>
row.partida ===
category
);

}


if (supplier) {

rows =
rows.filter(
row =>
row.supplier ===
supplier
);

}


if (status) {

if (
status ===
"Sin OC definida"
) {

rows =
rows.filter(
row =>
!row.po
);

} else {

rows =
rows.filter(
row =>
row.status ===
status
);

}

}


const budget =
sum(
rows.map(
row =>
row.current
)
);


const ordered =
sum(
rows
.filter(
row =>
row.po
)
.map(
row =>
row.netValue
)
);


setText(
"budgetFilteredTotal",
money(budget)
);

setText(
"budgetFilteredOrdered",
money(ordered)
);

setText(
"budgetFilteredPending",
money(
Math.max(
budget -
ordered,
0
)
)
);

setText(
"budgetFilteredItems",
rows.length
);

setText(
"budgetTableCount",
`${rows.length} registros`
);


const tbody =
$("budgetTableBody");

if (!tbody) return;


tbody.innerHTML =
rows.map(
row => `

<tr>

<td>
${escapeHtml(
row.partida
)}
</td>

<td>
${escapeHtml(
row.item
)}
</td>

<td>
${escapeHtml(
row.supplier ||
"—"
)}
</td>

<td>

<span class="status-pill ${
row.status ===
"Con OC"
? "good"
: row.status ===
"Sin proveedor"
? "bad"
: "warn"
}">

${escapeHtml(
row.po ||
row.status
)}

</span>

</td>

<td>
${escapeHtml(
formatPossibleDate(
row.date
)
)}
</td>

<td class="num money">
${money(
row.current
)}
</td>

<td>
${escapeHtml(
row.currency
)}
</td>

</tr>

`
).join("");

}


/* =========================================================
FLUJO DE CAJA
========================================================= */

function renderCashflow() {

const payments =
extractPayments();

state.cashflowPayments =
payments;

renderCashflowKPIs(
payments
);

renderPaymentReminders(
payments
);

renderCashflowCharts(
payments
);

renderPaymentTable(
payments
);

renderOverviewPaymentReminders();

}


/* =========================================================
EXTRAER PAGOS DE BD PLAN DETALLADO
========================================================= */

function extractPayments() {

const payments = [];


state.budget.forEach(
row => {

const raw =
row.raw || {};


/*
Soporta Pago 1 hasta Pago 10
por si en el futuro la plantilla
crece.
*/

for (
let i = 1;
i <= 10;
i++
) {

const period =
valueFrom(
raw,
[
`Per.P${i}`,
`Per P${i}`,
`Periodo ${i}`,
`Período ${i}`,
`Fecha Pago ${i}`
]
);


const amount =
numberFrom(
raw,
[
`Pago ${i}`,
`Pago${i}`,
`Payment ${i}`
]
);


if (!amount) {
continue;
}


const date =
parsePaymentPeriod(
period
);


const currency =
row.currency ||
"DOP";


payments.push({

date,

period,

amount,

currency,

dop:
convertToDOP(
amount,
currency,
row.rate
),

item:
row.item,

partida:
row.partida,

supplier:
row.supplier,

paymentNumber:
i

});

}

}
);


return payments.sort(
(a, b) =>
dateNumber(a.date) -
dateNumber(b.date)
);

}


/* =========================================================
INTERPRETAR PERÍODOS DE PAGO
========================================================= */

function parsePaymentPeriod(
value
) {

if (!value) {
return null;
}


const direct =
parseExcelDate(
value
);

if (direct) {
return direct;
}


const text =
String(value)
.trim()
.toLowerCase();


const months = {

ene: 0,
enero: 0,

feb: 1,
febrero: 1,

mar: 2,
marzo: 2,

abr: 3,
abril: 3,

may: 4,
mayo: 4,

jun: 5,
junio: 5,

jul: 6,
julio: 6,

ago: 7,
agosto: 7,

sep: 8,
sept: 8,
septiembre: 8,

oct: 9,
octubre: 9,

nov: 10,
noviembre: 10,

dic: 11,
diciembre: 11

};


for (
const [name, month]
of Object.entries(months)
) {

if (
text.includes(name)
) {

const yearMatch =
text.match(
/20\d{2}/
);

const year =
yearMatch
? Number(
yearMatch[0]
)
: new Date()
.getFullYear();

return new Date(
year,
month,
1
);

}

}


return null;

}


/* =========================================================
CONVERSIÓN A DOP
========================================================= */

function convertToDOP(
amount,
currency,
rowRate
) {

const c =
normalizeText(
currency
);


if (
c.includes("usd") ||
c.includes("dolar")
) {

const rate =
rowRate ||
state.exchangeRate;

return amount *
rate;

}


return amount;

}


/* =========================================================
KPI FLUJO DE CAJA
========================================================= */

function renderCashflowKPIs(
payments
) {

const now =
startOfToday();


const total =
sum(
payments.map(
payment =>
payment.dop
)
);


const past =
sum(
payments
.filter(
payment =>
payment.date &&
payment.date <
now
)
.map(
payment =>
payment.dop
)
);


const monthly =
groupPaymentsByMonth(
payments
);


const entries =
Object.entries(
monthly
).sort(
(a, b) =>
a[1].date -
b[1].date
);


const peak =
[...entries]
.sort(
(a, b) =>
b[1].amount -
a[1].amount
)[0];


const next =
entries.find(
([, value]) =>
value.date >=
new Date(
now.getFullYear(),
now.getMonth(),
1
)
);


const ninety =
new Date(now);

ninety.setDate(
ninety.getDate() +
90
);


const next90 =
sum(
payments
.filter(
payment =>
payment.date &&
payment.date >=
now &&
payment.date <=
ninety
)
.map(
payment =>
payment.dop
)
);


setText(
"cfTotal",
money(total)
);

setText(
"cfPaidToDate",
money(past)
);

setText(
"cfPeakMonth",
peak
? peak[0]
: "—"
);

setText(
"cfPeakAmount",
peak
? money(
peak[1].amount
)
: "—"
);

setText(
"cfNextMonth",
next
? next[0]
: "—"
);

setText(
"cfNextAmount",
next
? money(
next[1].amount
)
: "—"
);

setText(
"cf90Days",
money(next90)
);


renderUpcomingMonths(
entries
);

}


/* =========================================================
RECORDATORIOS DE PAGO
========================================================= */

function renderPaymentReminders(
payments
) {

const container =
$("paymentReminderList");

if (!container) return;


const now =
startOfToday();


const future =
payments
.filter(
payment =>
payment.date
)
.map(
payment => ({

...payment,

days:
Math.ceil(
(
payment.date -
now
) /
86400000
)

})
)
.filter(
payment =>
payment.days >=
-30
)
.sort(
(a, b) =>
a.days -
b.days
)
.slice(0, 8);


container.innerHTML =
future.length
? future.map(
payment => {

let cls =
"upcoming";

let label =
"";


if (
payment.days < 0
) {

cls =
"overdue";

label =
`${Math.abs(
payment.days
)} días vencido`;

} else if (
payment.days <= 30
) {

cls =
"current";

label =
`En ${payment.days} días`;

} else {

label =
formatDate(
payment.date
);

}


return `

<div class="payment-reminder ${cls}">

<span class="label">
${escapeHtml(
label
)}
</span>

<h3>
${escapeHtml(
payment.item
)}
</h3>

<b>
${formatOriginalMoney(
payment.amount,
payment.currency
)}
</b>

<p>
${escapeHtml(
payment.supplier ||
"Sin proveedor"
)}
· Pago
${payment.paymentNumber}
</p>

</div>

`;

}
).join("")
: emptyMessage(
"No hay pagos programados."
);

}


/* =========================================================
RECORDATORIOS RESUMEN
========================================================= */

function renderOverviewPaymentReminders() {

const container =
$("overviewPaymentReminders");

if (!container) return;


const payments =
state.cashflowPayments ||
[];


const now =
startOfToday();


const items =
payments
.filter(
payment =>
payment.date &&
payment.date >=
now
)
.sort(
(a, b) =>
a.date -
b.date
)
.slice(0, 4);


container.innerHTML =
items.length
? items.map(
payment => {

const days =
Math.ceil(
(
payment.date -
now
) /
86400000
);


return `

<div class="reminder-mini ${
days <= 30
? "current"
: ""
}">

<span>
${
days === 0
? "Hoy"
: `En ${days} días`
}
</span>

<b>
${money(
payment.dop
)}
</b>

<small>
${escapeHtml(
payment.item
)}
</small>

</div>

`;

}
).join("")
: emptyMessage(
"No hay próximos pagos identificados."
);

}


/* =========================================================
GRÁFICAS FLUJO DE CAJA
========================================================= */

function renderCashflowCharts(
payments
) {

const monthly =
groupPaymentsByMonth(
payments
);


const entries =
Object.entries(
monthly
).sort(
(a, b) =>
a[1].date -
b[1].date
);


const labels =
entries.map(
entry =>
entry[0]
);


const values =
entries.map(
entry =>
entry[1].amount
);


createChart(
"cashflowChart",
"bar",
labels,
values,
{
plugins: {
legend: {
display: false
}
}
}
);


const now =
startOfToday();


const paid =
sum(
payments
.filter(
p =>
p.date &&
p.date <
now
)
.map(
p =>
p.dop
)
);


const future =
sum(
payments
.filter(
p =>
!p.date ||
p.date >=
now
)
.map(
p =>
p.dop
)
);


createChart(
"cashflowStatusChart",
"doughnut",
[
"Calendario anterior",
"Por venir"
],
[
paid,
future
]
);


const currencyGroups = {};


payments.forEach(
payment => {

const currency =
payment.currency ||
"DOP";

currencyGroups[
currency
] =
(
currencyGroups[
currency
] ||
0
) +
payment.dop;

}
);


createChart(
"cashflowCurrencyChart",
"doughnut",
Object.keys(
currencyGroups
),
Object.values(
currencyGroups
)
);


let cumulative = 0;


const cumulativeValues =
values.map(
value => {

cumulative +=
value;

return cumulative;

}
);


createChart(
"cashflowCumulativeChart",
"line",
labels,
cumulativeValues,
{
plugins: {
legend: {
display: false
}
}
}
);

}


/* =========================================================
AGRUPAR PAGOS POR MES
========================================================= */

function groupPaymentsByMonth(
payments
) {

const groups = {};


payments.forEach(
payment => {

if (!payment.date) {
return;
}


const date =
new Date(
payment.date
.getFullYear(),

payment.date
.getMonth(),

1
);


const label =
date.toLocaleDateString(
"es-DO",
{
month:
"short",

year:
"numeric"
}
);


if (!groups[label]) {

groups[label] = {
amount: 0,
date
};

}


groups[label].amount +=
payment.dop;

}
);


return groups;

}


/* =========================================================
PRÓXIMOS MESES
========================================================= */

function renderUpcomingMonths(
entries
) {

const container =
$("cashflowUpcomingMonths");

if (!container) return;


const now =
new Date();


const currentMonth =
new Date(
now.getFullYear(),
now.getMonth(),
1
);


const future =
entries
.filter(
([, value]) =>
value.date >=
currentMonth
)
.slice(0, 6);


const max =
Math.max(
...future.map(
([, value]) =>
value.amount
),
0
);


container.innerHTML =
future.length
? future.map(
(
[label, value],
index
) => `

<div class="upcoming-month ${
index === 0
? "next"
: value.amount >=
max * .8
? "high"
: ""
}">

<div>

<span>
${
index === 0
? "Próximo compromiso"
: "Mes programado"
}
</span>

<b>
${escapeHtml(
label
)}
</b>

</div>

<div class="upcoming-amount">

<b>
${money(
value.amount
)}
</b>

<small>
Equiv. DOP
</small>

</div>

</div>

`
).join("")
: emptyMessage(
"No hay meses futuros con desembolsos."
);

}


/* =========================================================
TABLA DE PAGOS
========================================================= */

function renderPaymentTable(
payments
) {

const tbody =
$("paymentTable");

if (!tbody) return;


const now =
startOfToday();


setText(
"paymentTableCount",
`${payments.length} pagos`
);


tbody.innerHTML =
payments.map(
payment => {

const days =
payment.date
? Math.ceil(
(
payment.date -
now
) /
86400000
)
: null;


let status =
"Sin fecha";

let cls =
"info";


if (
days !== null
) {

if (
days < 0
) {

status =
"Fecha anterior";

cls =
"bad";

} else if (
days <= 30
) {

status =
"Próximo";

cls =
"warn";

} else {

status =
"Futuro";

cls =
"good";

}

}


return `

<tr>

<td>
${escapeHtml(
payment.date
? monthYear(
payment.date
)
: payment.period ||
"—"
)}
</td>

<td>
${escapeHtml(
payment.partida ||
"—"
)}
</td>

<td>
${escapeHtml(
payment.item ||
"—"
)}
</td>

<td>
${escapeHtml(
payment.supplier ||
"—"
)}
</td>

<td>
Pago
${payment.paymentNumber}
</td>

<td>
${escapeHtml(
payment.currency
)}
</td>

<td class="num">
${formatOriginalMoney(
payment.amount,
payment.currency
)}
</td>

<td class="num money">
${money(
payment.dop
)}
</td>

<td>

<span class="status-pill ${cls}">
${status}
</span>

</td>

</tr>

`;

}
).join("");

}


/* =========================================================
SELECTOR MONEDA
========================================================= */

function setupCashflowCurrency() {

const select =
$("cashflowCurrencyView");

if (!select) return;


select.addEventListener(
"change",
renderCashflow
);

}


/* =========================================================
SEGUIMIENTO
========================================================= */

function renderTracking() {

const withPO =
state.budget.filter(
row =>
row.status ===
"Con OC"
);


const pending =
state.budget.filter(
row =>
row.status !==
"Con OC"
);


const noSupplier =
state.budget.filter(
row =>
!row.supplier
);


const noDate =
state.budget.filter(
row =>
!row.date
);


setText(
"trackWithPO",
withPO.length
);

setText(
"trackPending",
pending.length
);

setText(
"trackNoSupplier",
noSupplier.length
);

setText(
"trackNoDate",
noDate.length
);


renderPipeline(
withPO.length,
pending.length,
noSupplier.length,
noDate.length
);


renderBottlenecks(
pending
);

renderSupplierPO(
withPO
);

renderTrackingTable(
pending
);

}


/* =========================================================
PIPELINE
========================================================= */

function renderPipeline(
withPO,
pending,
noSupplier,
noDate
) {

const container =
$("pipeline");

if (!container) return;


const total =
state.budget.length;


const stages = [

{
label:
"Necesidades",

value:
total,

detail:
"Ítems identificados"
},

{
label:
"Proveedor definido",

value:
total -
noSupplier,

detail:
"Con proveedor"
},

{
label:
"Orden colocada",

value:
withPO,

detail:
"Con OC"
},

{
label:
"Pendientes",

value:
pending,

detail:
`${noDate} sin fecha`
}

];


container.innerHTML =
stages.map(
stage => `

<div class="stage">

<span>
${stage.label}
</span>

<b>
${stage.value}
</b>

<small>
${stage.detail}
</small>

</div>

`
).join("");

}


/* =========================================================
CUELLOS DE BOTELLA
========================================================= */

function renderBottlenecks(
pending
) {

const top =
[...pending]
.sort(
(a, b) =>
b.current -
a.current
)
.slice(0, 10);


createChart(
"bottleneckChart",
"bar",
top.map(
row =>
row.item
),
top.map(
row =>
row.current
),
{
indexAxis: "y",

plugins: {
legend: {
display: false
}
}
}
);

}


/* =========================================================
OC POR PROVEEDOR
========================================================= */

function renderSupplierPO(
withPO
) {

const groups = {};


withPO.forEach(
row => {

const supplier =
row.supplier ||
"Sin proveedor";

groups[supplier] =
(
groups[
supplier
] ||
0
) + 1;

}
);


const entries =
Object.entries(groups)
.sort(
(a, b) =>
b[1] - a[1]
)
.slice(0, 10);


createChart(
"supplierPOChart",
"bar",
entries.map(
x => x[0]
),
entries.map(
x => x[1]
),
{
plugins: {
legend: {
display: false
}
}
}
);

}


/* =========================================================
TABLA SEGUIMIENTO
========================================================= */

function renderTrackingTable(
pending
) {

const tbody =
$("trackingTable");

if (!tbody) return;


setText(
"trackingCount",
`${pending.length} registros`
);


tbody.innerHTML =
pending
.sort(
(a, b) =>
b.current -
a.current
)
.map(
row => `

<tr>

<td>
${escapeHtml(
row.item
)}
</td>

<td>
${escapeHtml(
row.partida
)}
</td>

<td>
${escapeHtml(
row.supplier ||
"—"
)}
</td>

<td>

<span class="status-pill ${
row.supplier
? "warn"
: "bad"
}">

${escapeHtml(
row.status
)}

</span>

</td>

<td>
${escapeHtml(
row.comment ||
"—"
)}
</td>

<td class="num money">
${money(
row.current
)}
</td>

</tr>

`
).join("");

}


/* =========================================================
PENDIENTES
========================================================= */

function renderPending() {

/*
Si existe hoja Pendientes,
usamos esa.

Si no existe, generamos pendientes
desde BD Plan Detallado.
*/

let issues =
[...state.issues];


if (!issues.length) {

issues =
state.budget
.filter(
row =>
row.comment &&
row.status !==
"Con OC"
)
.map(
row => ({

title:
row.item,

owner:
row.supplier ||
"Sin responsable",

comment:
row.comment,

status:
row.status

})
);

}


const owners =
unique(
issues.map(
issue =>
issue.owner
)
);


const withComments =
issues.filter(
issue =>
issue.comment
);


setText(
"pendingTotal",
issues.length
);

setText(
"pendingOwners",
owners.length
);

setText(
"pendingComments",
withComments.length
);

setText(
"pendingNoComments",
issues.length -
withComments.length
);


renderOwnerChart(
issues
);

renderConsiderations();

renderIssueBoard(
issues
);

}


/* =========================================================
RESPONSABLES
========================================================= */

function renderOwnerChart(
issues
) {

const groups = {};


issues.forEach(
issue => {

groups[
issue.owner
] =
(
groups[
issue.owner
] ||
0
) + 1;

}
);


const entries =
Object.entries(groups)
.sort(
(a, b) =>
b[1] - a[1]
);


createChart(
"ownerChart",
"bar",
entries.map(
x => x[0]
),
entries.map(
x => x[1]
),
{
plugins: {
legend: {
display: false
}
}
}
);

}


/* =========================================================
CONSIDERACIONES
========================================================= */

function renderConsiderations() {

const container =
$("considerations");

if (!container) return;


let items =
state.considerations
.map(
row =>
valueFrom(
row,
[
"Consideración",
"Consideracion",
"Nota",
"Descripción",
"Descripcion"
]
)
)
.filter(Boolean);


if (!items.length) {

items = [

"Validar fechas de entrega de partidas críticas.",

"Confirmar proveedor y orden de compra antes del corte.",

"Actualizar comentarios de pendientes abiertos.",

"Revisar pagos próximos contra el flujo de caja.",

"Mantener BD Plan Detallado y Plan OI actualizados."

];

}


container.innerHTML =
items
.slice(0, 8)
.map(
item => `

<div class="consideration">

${escapeHtml(
item
)}

</div>

`
).join("");

}


/* =========================================================
TARJETAS PENDIENTES
========================================================= */

function renderIssueBoard(
issues
) {

const container =
$("issueBoard");

if (!container) return;


container.innerHTML =
issues.length
? issues.map(
issue => `

<article class="issue-card">

<b>
${escapeHtml(
issue.title
)}
</b>

<div class="owner">
${escapeHtml(
issue.owner
)}
</div>

<p>
${escapeHtml(
issue.comment ||
"Sin comentario registrado."
)}
</p>

</article>

`
).join("")
: `

<section class="card">

No hay pendientes
registrados.

</section>

`;

}


/* =========================================================
PRE-LIQUIDACIÓN
========================================================= */

function setupLiquidation() {

[
"liqAmount",
"liqExw",
"liqFob",
"liqContainers",
"liqRate"
].forEach(
id => {

const input =
$(id);

if (input) {

input.addEventListener(
"input",
renderLiquidation
);

}

}
);

}


function renderLiquidation() {

const container =
$("liquidationScenarios");

if (!container) return;


const amount =
numberInput(
"liqAmount",
10000
);

const exw =
numberInput(
"liqExw",
0
);

const fob =
numberInput(
"liqFob",
1000
);

const containers =
numberInput(
"liqContainers",
1
);

const rate =
numberInput(
"liqRate",
state.exchangeRate
);


state.exchangeRate =
rate;


const scenarios = [

{
country:
"China",

freight:
4200,

duty:
.20,

insurance:
.01,

other:
1200
},

{
country:
"Estados Unidos",

freight:
2200,

duty:
.20,

insurance:
.01,

other:
900
},

{
country:
"Europa",

freight:
3500,

duty:
.20,

insurance:
.01,

other:
1100
},

{
country:
"México",

freight:
2600,

duty:
.20,

insurance:
.01,

other:
950
},

{
country:
"Puerto Rico",

freight:
1600,

duty:
.20,

insurance:
.01,

other:
800
}

];


container.innerHTML =
scenarios.map(
scenario => {

const freight =
scenario.freight *
containers;


const base =
amount +
exw +
fob;


const insurance =
base *
scenario.insurance;


const cif =
base +
freight +
insurance;


const duty =
cif *
scenario.duty;


const taxable =
cif +
duty +
scenario.other;


const itbis =
taxable *
.18;


const total =
taxable +
itbis;


return `

<article class="scenario">

<h3>
${scenario.country}
</h3>

<div class="metric">

<span>
Compra
</span>

<b>
${usd(amount)}
</b>

</div>

<div class="metric">

<span>
Flete estimado
</span>

<b>
${usd(freight)}
</b>

</div>

<div class="metric">

<span>
Seguro
</span>

<b>
${usd(insurance)}
</b>

</div>

<div class="metric">

<span>
CIF
</span>

<b>
${usd(cif)}
</b>

</div>

<div class="metric">

<span>
Gravamen
</span>

<b>
${usd(duty)}
</b>

</div>

<div class="metric">

<span>
Otros gastos
</span>

<b>
${usd(
scenario.other
)}
</b>

</div>

<div class="metric">

<span>
ITBIS
</span>

<b>
${usd(itbis)}
</b>

</div>

<div class="total">

<span>
Costo estimado puesto RD
</span>

<b>
${usd(total)}
</b>

<small>
≈
${money(
total *
rate
)}
</small>

</div>

</article>

`;

}
).join("");

}


/* =========================================================
AGRUPACIONES
========================================================= */

function groupSupplierValue() {

const groups = {};


state.budget.forEach(
row => {

if (!row.supplier) {
return;
}


groups[
row.supplier
] =
(
groups[
row.supplier
] ||
0
) +
(
row.netValue ||
row.current
);

}
);


const entries =
Object.entries(groups)
.sort(
(a, b) =>
b[1] - a[1]
)
.slice(0, 12);


return {

labels:
entries.map(
x => x[0]
),

values:
entries.map(
x => x[1]
)

};

}


function groupSum(
rows,
labelKey,
valueKey
) {

const groups = {};


rows.forEach(
row => {

const label =
row[labelKey] ||
"Sin clasificar";


const value =
Number(
row[valueKey]
) || 0;


groups[label] =
(
groups[label] ||
0
) +
value;

}
);


const entries =
Object.entries(groups)
.filter(
([, value]) =>
value !== 0
)
.sort(
(a, b) =>
b[1] - a[1]
);


return {

labels:
entries.map(
x => x[0]
),

values:
entries.map(
x => x[1]
)

};

}


/* =========================================================
CHART.JS
========================================================= */

function createChart(
id,
type,
labels,
values,
options = {}
) {

const canvas =
$(id);

if (!canvas) return;


destroyChart(id);


const background =
type ===
"doughnut"

? [
COLORS.blue,
COLORS.cyan,
COLORS.green,
COLORS.yellow,
COLORS.red,
"#4388c7",
"#7bc7ed",
"#8b7dbb",
"#f6b65b"
]

: COLORS.cyan;


state.charts[id] =
new Chart(
canvas.getContext(
"2d"
),
{

type,

data: {

labels,

datasets: [
{

data:
values,

backgroundColor:
background,

borderColor:
type ===
"line"
? COLORS.blue
: undefined,

tension:
.3,

fill:
false

}
]

},

options:
mergeChartOptions(
options
)

}
);

}


function createMultiChart(
id,
type,
labels,
datasets,
options = {}
) {

const canvas =
$(id);

if (!canvas) return;


destroyChart(id);


state.charts[id] =
new Chart(
canvas.getContext(
"2d"
),
{

type,

data: {
labels,
datasets
},

options:
mergeChartOptions(
options
)

}
);

}


function mergeChartOptions(
custom
) {

const base = {

responsive:
true,

maintainAspectRatio:
false,

interaction: {

intersect:
false,

mode:
"index"

},

plugins: {

legend: {

labels: {

boxWidth:
10,

boxHeight:
10,

font: {
size: 10
}

}

},

tooltip: {

callbacks: {

label(context) {

const value =
Number(
context.raw
);


if (
Number.isFinite(
value
) &&
Math.abs(value) >=
1000
) {

return (
(
context.dataset
.label ||
""
) +
" " +
money(value)
);

}


return (
(
context.dataset
.label ||
""
) +
" " +
value
);

}

}

}

},

scales: {

x: {

ticks: {

font: {
size: 9
},

maxRotation:
45,

minRotation:
0

},

grid: {
display: false
}

},

y: {

ticks: {

font: {
size: 9
}

},

grid: {
color:
"#edf2f6"
}

}

}

};


return deepMerge(
base,
custom
);

}


function deepMerge(
target,
source
) {

const output = {
...target
};


Object.keys(
source || {}
).forEach(
key => {

if (
source[key] &&
typeof source[key] ===
"object" &&
!Array.isArray(
source[key]
)
) {

output[key] =
deepMerge(
target[key] ||
{},
source[key]
);

} else {

output[key] =
source[key];

}

}
);


return output;

}


function destroyChart(id) {

if (
state.charts[id]
) {

state.charts[id]
.destroy();

delete state.charts[
id
];

}

}


/* =========================================================
COUNTDOWN
========================================================= */

let countdownInterval =
null;


function startCountdown(
dateString
) {

if (
countdownInterval
) {

clearInterval(
countdownInterval
);

}


const target =
parseDate(
dateString
);


function update() {

if (!target) return;


const now =
new Date();


let diff =
target -
now;


const status =
$("countdownStatus");


if (
diff <= 0
) {

diff = 0;

if (status) {

status.textContent =
"Fecha de apertura alcanzada";

}

} else if (
status
) {

status.textContent =
"Hacia la apertura";

}


const days =
Math.floor(
diff /
86400000
);


const hours =
Math.floor(
(
diff %
86400000
) /
3600000
);


const minutes =
Math.floor(
(
diff %
3600000
) /
60000
);


const seconds =
Math.floor(
(
diff %
60000
) /
1000
);


setText(
"cdDays",
days
);

setText(
"cdHours",
String(hours)
.padStart(
2,
"0"
)
);

setText(
"cdMinutes",
String(minutes)
.padStart(
2,
"0"
)
);

setText(
"cdSeconds",
String(seconds)
.padStart(
2,
"0"
)
);

}


update();


countdownInterval =
setInterval(
update,
1000
);

}


/* =========================================================
ESTADO VACÍO
========================================================= */

function renderEmptyState() {

setText(
"overviewSubtitle",
"No fue posible cargar la plantilla de esta iniciativa."
);


[
"kpiApprovedBudget",
"kpiBudget",
"kpiOrdered",
"kpiPendingBuy",
"kpiWeightedProgress",
"kpiOpenIssues",
"overviewDiscount",
"overviewDiscountPct"
].forEach(
id =>
setText(
id,
"—"
)
);

}


/* =========================================================
UTILIDADES DE COLUMNAS
========================================================= */

function valueFrom(
row,
candidates
) {

const keys =
Object.keys(row);


for (
const candidate
of candidates
) {

const exact =
keys.find(
key =>
normalizeText(key) ===
normalizeText(
candidate
)
);


if (
exact &&
row[exact] !==
undefined &&
row[exact] !==
null &&
String(
row[exact]
).trim() !==
""
) {

return row[
exact
];

}

}


for (
const candidate
of candidates
) {

const partial =
keys.find(
key =>
normalizeText(key)
.includes(
normalizeText(
candidate
)
)
);


if (
partial &&
row[partial] !==
undefined &&
row[partial] !==
null &&
String(
row[partial]
).trim() !==
""
) {

return row[
partial
];

}

}


return "";

}


function numberFrom(
row,
candidates
) {

return toNumber(
valueFrom(
row,
candidates
)
);

}


/* =========================================================
CONVERSIÓN DE NÚMEROS
========================================================= */

function toNumber(value) {

if (
value === null ||
value === undefined ||
value === ""
) {

return 0;

}


if (
typeof value ===
"number"
) {

return value;

}


let text =
String(value)
.trim()
.replace(
/\s/g,
""
)
.replace(
/RD\$/gi,
""
)
.replace(
/US\$/gi,
""
)
.replace(
/USD/gi,
""
)
.replace(
/\$/g,
""
);


if (
text.includes(",") &&
text.includes(".")
) {

if (
text.lastIndexOf(",") >
text.lastIndexOf(".")
) {

text =
text
.replace(
/\./g,
""
)
.replace(
",",
"."
);

} else {

text =
text.replace(
/,/g,
""
);

}

} else if (
text.includes(",")
) {

const parts =
text.split(",");


if (
parts.length === 2 &&
parts[1].length <= 2
) {

text =
parts[0]
.replace(
/\./g,
""
) +
"." +
parts[1];

} else {

text =
text.replace(
/,/g,
""
);

}

}


text =
text.replace(
/[^0-9.-]/g,
""
);


const number =
Number(text);


return Number.isFinite(
number
)
? number
: 0;

}


/* =========================================================
NORMALIZAR TEXTO
========================================================= */

function normalizeText(value) {

return String(
value || ""
)
.normalize("NFD")
.replace(
/[\u0300-\u036f]/g,
""
)
.toLowerCase()
.replace(
/[^a-z0-9]/g,
""
);

}


/* =========================================================
FECHAS
========================================================= */

function parseExcelDate(
value
) {

if (!value) {
return null;
}


if (
value instanceof Date &&
!isNaN(value)
) {

return value;

}


if (
typeof value ===
"number"
) {

const parsed =
XLSX.SSF
.parse_date_code(
value
);


if (parsed) {

return new Date(
parsed.y,
parsed.m - 1,
parsed.d
);

}

}


const text =
String(value)
.trim();


if (!text) {
return null;
}


const match =
text.match(
/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/
);


if (match) {

let year =
Number(
match[3]
);


if (
year < 100
) {

year +=
2000;

}


return new Date(
year,
Number(
match[2]
) - 1,
Number(
match[1]
)
);

}


const direct =
new Date(text);


if (
!isNaN(direct)
) {

return direct;

}


return null;

}


function parseDate(value) {

if (!value) {
return null;
}


const date =
new Date(
`${value}T00:00:00`
);


return isNaN(date)
? null
: date;

}


function dateNumber(date) {

return date
? date.getTime()
: Number.MAX_SAFE_INTEGER;

}


function startOfToday() {

const now =
new Date();


return new Date(
now.getFullYear(),
now.getMonth(),
now.getDate()
);

}


function formatDate(date) {

if (!date) {
return "—";
}


return date
.toLocaleDateString(
"es-DO",
{
day:
"2-digit",

month:
"short",

year:
"numeric"
}
);

}


function formatLongDate(date) {

if (!date) {
return "—";
}


return date
.toLocaleDateString(
"es-DO",
{
day:
"numeric",

month:
"long",

year:
"numeric"
}
);

}


function formatPossibleDate(
value
) {

const date =
parseExcelDate(
value
);


return date
? formatDate(date)
: value ||
"—";

}


function monthYear(date) {

if (!date) {
return "—";
}


return date
.toLocaleDateString(
"es-DO",
{
month:
"short",

year:
"numeric"
}
);

}


/* =========================================================
FORMATOS
========================================================= */

function money(value) {

const number =
Number(value) ||
0;


if (
Math.abs(number) >=
1000000000
) {

return (
"RD$ " +
(
number /
1000000000
).toFixed(2) +
" B"
);

}


if (
Math.abs(number) >=
1000000
) {

return (
"RD$ " +
(
number /
1000000
).toFixed(2) +
" MM"
);

}


return new Intl
.NumberFormat(
"es-DO",
{
style:
"currency",

currency:
"DOP",

maximumFractionDigits:
0
}
)
.format(number);

}


function usd(value) {

return new Intl
.NumberFormat(
"en-US",
{
style:
"currency",

currency:
"USD",

maximumFractionDigits:
0
}
)
.format(
Number(value) ||
0
);

}


function formatOriginalMoney(
value,
currency
) {

const c =
normalizeText(
currency
);


if (
c.includes("usd") ||
c.includes("dolar")
) {

return usd(value);

}


return money(value);

}


function percent(value) {

return new Intl
.NumberFormat(
"es-DO",
{
style:
"percent",

maximumFractionDigits:
1
}
)
.format(
Number(value) ||
0
);

}


/* =========================================================
UTILIDADES
========================================================= */

function sum(values) {

return values.reduce(
(
total,
value
) =>
total +
(
Number(value) ||
0
),
0
);

}


function unique(values) {

return [
...new Set(
values.filter(
value =>
value !== null &&
value !== undefined &&
String(value)
.trim() !==
""
)
)
];

}


function setText(
id,
value
) {

const element =
$(id);


if (element) {

element.textContent =
value;

}

}


function numberInput(
id,
fallback
) {

const element =
$(id);


if (!element) {
return fallback;
}


const value =
Number(
element.value
);


return Number.isFinite(
value
)
? value
: fallback;

}


function showLoader(show) {

const loader =
$("loader");


if (!loader) {
return;
}


loader.classList.toggle(
"hidden",
!show
);

}


function escapeHtml(value) {

return String(
value === null ||
value === undefined
? ""
: value
)
.replace(
/&/g,
"&amp;"
)
.replace(
/</g,
"&lt;"
)
.replace(
/>/g,
"&gt;"
)
.replace(
/"/g,
"&quot;"
)
.replace(
/'/g,
"&#039;"
);

}


function emptyMessage(text) {

return `

<div class="muted">

${escapeHtml(text)}

</div>

`;

}
