const PROJECTS = [
{
id: "mao",
name: "Mao",
brand: "Sirena",
file: "data/projects/mao.xlsx",
openingDate: "2026-11-15"
},
{
id: "hispano",
name: "Hispano",
brand: "Sirena",
file: "data/projects/hispano.xlsx",
openingDate: "2026-12-15"
}
];

const state = {
project: null,
workbook: null,
budget: [],
cashflow: [],
issues: [],
considerations: [],
approvedBudget: 0,
currentBudget: 0,
orderedValue: 0,
charts: {},
exchangeRate: 57.85
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

async function init() {
setupNavigation();
setupProjectSelector();
setupBudgetFilters();
setupLiquidation();
setupCashflowCurrency();

const initial =
PROJECTS.find(p => p.id === localStorage.getItem("selectedProject")) ||
PROJECTS[0];

$("projectSelect").value = initial.id;

await loadProject(initial);
}

function setupNavigation() {
document.querySelectorAll("#navTabs button").forEach(button => {
button.addEventListener("click", () => {
openPage(button.dataset.page);
});
});

document.querySelectorAll("[data-go-page]").forEach(button => {
button.addEventListener("click", () => {
openPage(button.dataset.goPage);
});
});
}

function openPage(page) {
document.querySelectorAll(".page").forEach(section => {
section.classList.remove("active");
});

document.querySelectorAll("#navTabs button").forEach(button => {
button.classList.remove("active");
});

const target = $(`page-${page}`);
if (target) target.classList.add("active");

const navButton = document.querySelector(
`#navTabs button[data-page="${page}"]`
);

if (navButton) navButton.classList.add("active");

window.scrollTo({
top: 0,
behavior: "smooth"
});
}

function setupProjectSelector() {
const select = $("projectSelect");

PROJECTS.forEach(project => {
const option = document.createElement("option");
option.value = project.id;
option.textContent = `${project.brand} · ${project.name}`;
select.appendChild(option);
});

select.addEventListener("change", async event => {
const project = PROJECTS.find(
item => item.id === event.target.value
);

if (!project) return;

localStorage.setItem("selectedProject", project.id);
await loadProject(project);
});
}

async function loadProject(project) {
showLoader(true);

state.project = project;

updateProjectLabels(project);

try {
const response = await fetch(project.file);

if (!response.ok) {
throw new Error(`No se pudo cargar ${project.file}`);
}

const buffer = await response.arrayBuffer();

state.workbook = XLSX.read(buffer, {
type: "array",
cellDates: true
});

parseWorkbook();

renderAll();

} catch (error) {
console.error(error);

state.budget = [];
state.cashflow = [];
state.issues = [];
state.considerations = [];

renderEmptyState();
}

showLoader(false);
}

function updateProjectLabels(project) {
$("projectBrand").textContent = project.brand;
$("projectName").textContent = project.name;

$("homeProjectName").textContent = project.name;
$("overviewTitle").textContent =
`${project.brand} ${project.name}`;

$("countdownProject").textContent = project.name;
$("footerProject").textContent =
`Proyecto ${project.name}`;

const opening = parseDate(project.openingDate);

if (opening) {
$("homeOpeningDate").textContent =
formatDate(opening);

$("countdownDateLabel").textContent =
formatLongDate(opening);
}

startCountdown(project.openingDate);
}

function parseWorkbook() {
const wb = state.workbook;

state.budget = [];
state.cashflow = [];
state.issues = [];
state.considerations = [];

const planSheet =
findSheet(wb, [
"Plan OI",
"Plan",
"Plan Detallado",
"BD Plan Detallado"
]);

if (planSheet) {
state.budget = sheetToObjects(planSheet);
}

const cashSheet =
findSheet(wb, [
"Flujo de Caja",
"Flujo Caja",
"Cash Flow"
]);

if (cashSheet) {
state.cashflow = sheetToObjects(cashSheet);
}

const issueSheet =
findSheet(wb, [
"Pendientes",
"Issues",
"Seguimiento"
]);

if (issueSheet) {
state.issues = sheetToObjects(issueSheet);
}

const considerationsSheet =
findSheet(wb, [
"Consideraciones",
"Notas"
]);

if (considerationsSheet) {
state.considerations =
sheetToObjects(considerationsSheet);
}

normalizeBudget();
normalizeIssues();
calculateTotals();
}

function findSheet(workbook, candidates) {
const names = workbook.SheetNames;

for (const candidate of candidates) {
const found = names.find(
name =>
normalizeText(name) === normalizeText(candidate)
);

if (found) {
return workbook.Sheets[found];
}
}

for (const candidate of candidates) {
const found = names.find(
name =>
normalizeText(name).includes(
normalizeText(candidate)
)
);

if (found) {
return workbook.Sheets[found];
}
}

return null;
}

function sheetToObjects(sheet) {
return XLSX.utils.sheet_to_json(sheet, {
defval: "",
raw: false
});
}

function normalizeBudget() {
state.budget = state.budget
.map((row, index) => {
const partida =
valueFrom(row, [
"Partida",
"Gran Partida",
"Partida General",
"Agrupador",
"Supranumero",
"Supranúmero"
]) || "Sin partida";

const item =
valueFrom(row, [
"Item",
"Ítem",
"Descripcion",
"Descripción",
"Activo",
"Activo requerido",
"Concepto"
]) || `Ítem ${index + 1}`;

const supplier =
valueFrom(row, [
"Proveedor",
"Supplier"
]) || "";

const po =
valueFrom(row, [
"OC",
"Orden de Compra",
"PO"
]) || "";

const approved =
numberFrom(row, [
"Plan detallado Aprobado",
"Plan Detallado Aprobado",
"Presupuesto Aprobado",
"Total Aprobado RD$",
"Aprobado"
]);

const current =
numberFrom(row, [
"Plan Detallado en ejecución",
"Plan detallado en ejecucion",
"Plan Vigente",
"Presupuesto",
"Valor Total RD$",
"Monto"
]);

const ordered =
numberFrom(row, [
"Monto OC",
"Valor OC",
"Monto Orden",
"Comprometido"
]);

const currency =
valueFrom(row, [
"Moneda",
"Currency"
]) || "DOP";

const date =
valueFrom(row, [
"Fecha OC",
"Fecha",
"ETA",
"Fecha Entrega"
]);

const comment =
valueFrom(row, [
"Comentario",
"Comentarios",
"Observacion",
"Observación"
]);

return {
raw: row,
partida,
item,
supplier,
po,
approved,
current,
ordered,
currency,
date,
comment,
status: determinePurchaseStatus(
po,
ordered,
supplier
)
};
})
.filter(row =>
row.approved ||
row.current ||
row.ordered ||
row.item
);
}

function determinePurchaseStatus(po, ordered, supplier) {
if (po || ordered > 0) {
return "Con OC";
}

if (!supplier) {
return "Sin proveedor";
}

return "Pendiente de compra";
}

function normalizeIssues() {
state.issues = state.issues.map((row, index) => ({
title:
valueFrom(row, [
"Pendiente",
"Tema",
"Issue",
"Actividad",
"Descripción",
"Descripcion"
]) || `Pendiente ${index + 1}`,

owner:
valueFrom(row, [
"Responsable",
"Owner",
"Líder",
"Lider"
]) || "Sin responsable",

comment:
valueFrom(row, [
"Comentario",
"Comentarios",
"Observación",
"Observacion"
]) || "",

status:
valueFrom(row, [
"Estado",
"Status"
]) || "Abierto"
}));
}

function calculateTotals() {
state.approvedBudget = sum(
state.budget.map(row => row.approved)
);

state.currentBudget = sum(
state.budget.map(row => row.current)
);

state.orderedValue = sum(
state.budget.map(row => row.ordered)
);

if (!state.orderedValue) {
state.orderedValue = sum(
state.budget
.filter(row => row.po)
.map(row => row.current)
);
}
}

function renderAll() {
renderOverview();
renderBudget();
renderCashflow();
renderTracking();
renderPending();
renderLiquidation();
}

function renderOverview() {
const approved = state.approvedBudget;
const current = state.currentBudget;
const ordered = state.orderedValue;

const pending = Math.max(
current - ordered,
0
);

const saving =
approved > 0
? approved - current
: 0;

const savingPct =
approved > 0
? saving / approved
: 0;

const progress =
current > 0
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

renderHealth(progress, pending);
renderBudgetMix();
renderBudgetEvolution();
renderSavingsByPart();
renderPurchaseStatus();
renderTopSuppliers();
renderDataQuality();
renderAttentionList();
renderOverviewPaymentReminders();
}

function renderHealth(progress, pending) {
const container = $("healthStrip");

if (!container) return;

const withPO =
state.budget.filter(
row => row.status === "Con OC"
).length;

const total =
state.budget.length;

const coverage =
total
? withPO / total
: 0;

const noSupplier =
state.budget.filter(
row => !row.supplier
).length;

const health = [
{
title: "Ejecución presupuestaria",
value: percent(progress),
text:
progress >= .75
? "Buen nivel de colocación"
: "Todavía existe valor pendiente",
color:
progress >= .75
? "good"
: "warn"
},
{
title: "Cobertura de compras",
value: percent(coverage),
text:
`${withPO} de ${total} ítems con OC`,
color:
coverage >= .75
? "good"
: "warn"
},
{
title: "Pendiente de compra",
value: money(pending),
text:
"Valor todavía sin colocar",
color:
pending <= state.currentBudget * .2
? "good"
: "warn"
},
{
title: "Datos incompletos",
value: `${noSupplier}`,
text:
"Ítems sin proveedor definido",
color:
noSupplier
? "bad"
: "good"
}
];

container.innerHTML =
health.map(item => `
<div class="health-item">
<div class="row">
<b>${escapeHtml(item.title)}</b>
<i class="dot ${item.color}"></i>
</div>
<strong>${escapeHtml(item.value)}</strong>
<p>${escapeHtml(item.text)}</p>
</div>
`).join("");
}

function renderBudgetMix() {
const grouped =
groupSum(
state.budget,
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

function renderBudgetEvolution() {
const approved =
groupSum(
state.budget,
"partida",
"approved"
);

const currentMap =
toMap(
groupSum(
state.budget,
"partida",
"current"
)
);

const labels =
approved.labels;

const current =
labels.map(
label => currentMap[label] || 0
);

createMultiChart(
"budgetEvolutionChart",
"bar",
labels,
[
{
label: "Aprobado",
data: approved.values,
backgroundColor: COLORS.blue
},
{
label: "Vigente",
data: current,
backgroundColor: COLORS.cyan
}
]
);
}

function renderSavingsByPart() {
const groups = {};

state.budget.forEach(row => {
if (!groups[row.partida]) {
groups[row.partida] = {
approved: 0,
current: 0
};
}

groups[row.partida].approved +=
row.approved;

groups[row.partida].current +=
row.current;
});

const entries =
Object.entries(groups)
.map(([label, values]) => ({
label,
value:
values.approved -
values.current
}))
.sort(
(a, b) =>
Math.abs(b.value) -
Math.abs(a.value)
);

createChart(
"savingsByPartChart",
"bar",
entries.map(x => x.label),
entries.map(x => x.value),
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

function renderPurchaseStatus() {
const withPO =
state.budget.filter(
row => row.status === "Con OC"
).length;

const pending =
state.budget.filter(
row =>
row.status !== "Con OC"
).length;

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

function renderTopSuppliers() {
const container =
$("topSuppliers");

if (!container) return;

const groups = {};

state.budget.forEach(row => {
if (!row.supplier) return;

groups[row.supplier] =
(groups[row.supplier] || 0) +
(row.ordered || row.current);
});

const entries =
Object.entries(groups)
.sort((a, b) => b[1] - a[1])
.slice(0, 6);

container.innerHTML =
entries.length
? entries.map(
([supplier, value], index) => `
<div class="rank-row">
<span class="n">${index + 1}</span>
<b>${escapeHtml(supplier)}</b>
<span>${money(value)}</span>
</div>
`
).join("")
: emptyMessage("No hay proveedores disponibles.");
}

function renderDataQuality() {
const container =
$("dataQuality");

if (!container) return;

const total =
state.budget.length || 1;

const missingSupplier =
state.budget.filter(
row => !row.supplier
).length;

const missingPO =
state.budget.filter(
row => !row.po
).length;

const missingDate =
state.budget.filter(
row => !row.date
).length;

const items = [
{
label: "Sin proveedor",
value: missingSupplier
},
{
label: "Sin OC",
value: missingPO
},
{
label: "Sin fecha",
value: missingDate
}
];

container.innerHTML =
items.map(item => `
<div class="quality">
<div class="top">
<b>${item.label}</b>
<strong>${item.value}</strong>
</div>
<small>
${percent(item.value / total)}
de los ítems
</small>
</div>
`).join("");
}

function renderAttentionList() {
const container =
$("attentionList");

if (!container) return;

const pendingHigh =
[...state.budget]
.filter(
row =>
row.status !== "Con OC"
)
.sort(
(a, b) =>
b.current - a.current
)
.slice(0, 4);

const issues =
state.issues
.slice(0, 3)
.map(issue => ({
title: issue.title,
detail:
`${issue.owner}${
issue.comment
? " · " + issue.comment
: ""
}`,
color: "bad"
}));

const purchases =
pendingHigh.map(row => ({
title:
`Pendiente: ${row.item}`,
detail:
`${row.partida} · ${money(row.current)}`,
color: "warn"
}));

const items =
[...issues, ...purchases]
.slice(0, 6);

container.innerHTML =
items.length
? items.map(item => `
<div class="attention">
<i class="dot ${item.color}"></i>
<div>
<b>${escapeHtml(item.title)}</b>
<span>${escapeHtml(item.detail)}</span>
</div>
</div>
`).join("")
: emptyMessage(
"No se identifican alertas críticas."
);
}

function setupBudgetFilters() {
[
"budgetCategoryFilter",
"budgetStatusFilter",
"budgetSupplierFilter"
].forEach(id => {
const element = $(id);

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
$("budgetCategoryFilter").value = "";
$("budgetStatusFilter").value = "";
$("budgetSupplierFilter").value = "";

renderBudgetTable();
}
);
}
}

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
row => row.partida
)
)
);

fillSelect(
"budgetSupplierFilter",
unique(
state.budget
.map(row => row.supplier)
.filter(Boolean)
)
);
}

function fillSelect(id, values) {
const select = $(id);

if (!select) return;

const current =
select.value;

const first =
select.options[0]
? select.options[0].outerHTML
: '<option value="">Todos</option>';

select.innerHTML = first;

values
.sort((a, b) =>
String(a).localeCompare(
String(b),
"es"
)
)
.forEach(value => {
const option =
document.createElement("option");

option.value = value;
option.textContent = value;

select.appendChild(option);
});

if (
[...select.options].some(
option =>
option.value === current
)
) {
select.value = current;
}
}

function renderBudgetTable() {
const category =
$("budgetCategoryFilter")?.value || "";

const status =
$("budgetStatusFilter")?.value || "";

const supplier =
$("budgetSupplierFilter")?.value || "";

let rows =
[...state.budget];

if (category) {
rows = rows.filter(
row =>
row.partida === category
);
}

if (supplier) {
rows = rows.filter(
row =>
row.supplier === supplier
);
}

if (status) {
if (status === "Sin OC definida") {
rows = rows.filter(
row => !row.po
);
} else {
rows = rows.filter(
row =>
row.status === status
);
}
}

const budget =
sum(
rows.map(row => row.current)
);

const ordered =
sum(
rows.map(row => row.ordered)
) ||
sum(
rows
.filter(row => row.po)
.map(row => row.current)
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
budget - ordered,
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
rows.map(row => `
<tr>
<td>${escapeHtml(row.partida)}</td>
<td>${escapeHtml(row.item)}</td>
<td>${escapeHtml(row.supplier || "—")}</td>
<td>
<span class="status-pill ${
row.status === "Con OC"
? "good"
: row.status === "Sin proveedor"
? "bad"
: "warn"
}">
${escapeHtml(row.po || row.status)}
</span>
</td>
<td>${escapeHtml(formatPossibleDate(row.date))}</td>
<td class="num money">${money(row.current)}</td>
<td>${escapeHtml(row.currency)}</td>
</tr>
`).join("");
}

function renderCashflow() {
const payments =
extractPayments();

state.cashflowPayments =
payments;

renderCashflowKPIs(payments);
renderPaymentReminders(payments);
renderCashflowCharts(payments);
renderPaymentTable(payments);
renderOverviewPaymentReminders();
}

function extractPayments() {
const payments = [];

state.budget.forEach(row => {
const raw = row.raw || {};

for (let i = 1; i <= 5; i++) {
const period =
valueFrom(raw, [
`Per.P${i}`,
`Per P${i}`,
`Periodo ${i}`,
`Período ${i}`,
`Fecha Pago ${i}`
]);

const amount =
numberFrom(raw, [
`Pago ${i}`,
`Pago${i}`,
`Payment ${i}`
]);

if (!amount) continue;

const date =
parseExcelDate(period);

const currency =
valueFrom(raw, [
"Moneda",
"Currency"
]) ||
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
currency
),
item: row.item,
partida: row.partida,
supplier: row.supplier,
paymentNumber: i
});
}
});

if (!payments.length) {
state.cashflow.forEach(
(row, index) => {
const amount =
numberFrom(row, [
"Monto",
"Pago",
"Valor",
"Amount"
]);

if (!amount) return;

const dateValue =
valueFrom(row, [
"Fecha",
"Mes",
"Periodo",
"Período"
]);

const currency =
valueFrom(row, [
"Moneda",
"Currency"
]) || "DOP";

payments.push({
date:
parseExcelDate(dateValue),
period: dateValue,
amount,
currency,
dop:
convertToDOP(
amount,
currency
),
item:
valueFrom(row, [
"Item",
"Ítem",
"Concepto"
]) ||
`Pago ${index + 1}`,
partida:
valueFrom(row, [
"Partida"
]) || "",
supplier:
valueFrom(row, [
"Proveedor"
]) || "",
paymentNumber: 1
});
}
);
}

return payments.sort(
(a, b) =>
dateNumber(a.date) -
dateNumber(b.date)
);
}

function convertToDOP(amount, currency) {
const c =
normalizeText(currency);

if (
c.includes("usd") ||
c.includes("dolar")
) {
return amount *
state.exchangeRate;
}

return amount;
}

function renderCashflowKPIs(payments) {
const now =
startOfToday();

const total =
sum(
payments.map(
payment => payment.dop
)
);

const past =
sum(
payments
.filter(
payment =>
payment.date &&
payment.date < now
)
.map(
payment => payment.dop
)
);

const monthly =
groupPaymentsByMonth(payments);

const entries =
Object.entries(monthly)
.sort(
(a, b) =>
a[1].date - b[1].date
);

const peak =
[...entries].sort(
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
ninety.getDate() + 90
);

const next90 =
sum(
payments
.filter(
payment =>
payment.date &&
payment.date >= now &&
payment.date <= ninety
)
.map(
payment => payment.dop
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
? money(peak[1].amount)
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
? money(next[1].amount)
: "—"
);

setText(
"cf90Days",
money(next90)
);

renderUpcomingMonths(entries);
}

function renderPaymentReminders(payments) {
const container =
$("paymentReminderList");

if (!container) return;

const now =
startOfToday();

const future =
payments
.filter(payment => payment.date)
.map(payment => ({
...payment,
days:
Math.ceil(
(
payment.date -
now
) /
86400000
)
}))
.filter(
payment =>
payment.days >= -30
)
.sort(
(a, b) =>
a.days - b.days
)
.slice(0, 8);

container.innerHTML =
future.length
? future.map(payment => {
let cls = "upcoming";
let label = "";

if (payment.days < 0) {
cls = "overdue";
label =
`${Math.abs(payment.days)} días vencido`;
} else if (payment.days <= 30) {
cls = "current";
label =
`En ${payment.days} días`;
} else {
label =
formatDate(payment.date);
}

return `
<div class="payment-reminder ${cls}">
<span class="label">${escapeHtml(label)}</span>
<h3>${escapeHtml(payment.item)}</h3>
<b>${formatOriginalMoney(payment.amount, payment.currency)}</b>
<p>
${escapeHtml(payment.supplier || "Sin proveedor")}
· Pago ${payment.paymentNumber}
</p>
</div>
`;
}).join("")
: emptyMessage(
"No hay pagos programados."
);
}

function renderOverviewPaymentReminders() {
const container =
$("overviewPaymentReminders");

if (!container) return;

const payments =
state.cashflowPayments || [];

const now =
startOfToday();

const items =
payments
.filter(
payment =>
payment.date &&
payment.date >= now
)
.sort(
(a, b) =>
a.date - b.date
)
.slice(0, 4);

container.innerHTML =
items.length
? items.map(payment => {
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
<b>${money(payment.dop)}</b>
<small>
${escapeHtml(payment.item)}
</small>
</div>
`;
}).join("")
: emptyMessage(
"No hay próximos pagos identificados."
);
}

function renderCashflowCharts(payments) {
const monthly =
groupPaymentsByMonth(payments);

const entries =
Object.entries(monthly)
.sort(
(a, b) =>
a[1].date - b[1].date
);

const labels =
entries.map(
entry => entry[0]
);

const values =
entries.map(
entry => entry[1].amount
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
p.date < now
)
.map(p => p.dop)
);

const future =
sum(
payments
.filter(
p =>
!p.date ||
p.date >= now
)
.map(p => p.dop)
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

payments.forEach(payment => {
const currency =
payment.currency || "DOP";

currencyGroups[currency] =
(currencyGroups[currency] || 0) +
payment.dop;
});

createChart(
"cashflowCurrencyChart",
"doughnut",
Object.keys(currencyGroups),
Object.values(currencyGroups)
);

let cumulative = 0;

const cumulativeValues =
values.map(value => {
cumulative += value;
return cumulative;
});

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

function groupPaymentsByMonth(payments) {
const groups = {};

payments.forEach(payment => {
if (!payment.date) return;

const date =
new Date(
payment.date.getFullYear(),
payment.date.getMonth(),
1
);

const label =
date.toLocaleDateString(
"es-DO",
{
month: "short",
year: "numeric"
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
});

return groups;
}

function renderUpcomingMonths(entries) {
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
value.date >= currentMonth
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
([label, value], index) => `
<div class="upcoming-month ${
index === 0
? "next"
: value.amount >= max * .8
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
<b>${escapeHtml(label)}</b>
</div>

<div class="upcoming-amount">
<b>${money(value.amount)}</b>
<small>Equiv. DOP</small>
</div>
</div>
`
).join("")
: emptyMessage(
"No hay meses futuros con desembolsos."
);
}

function renderPaymentTable(payments) {
const tbody =
$("paymentTable");

if (!tbody) return;

const now =
startOfToday();

const rows =
[...payments]
.sort(
(a, b) =>
dateNumber(a.date) -
dateNumber(b.date)
);

setText(
"paymentTableCount",
`${rows.length} pagos`
);

tbody.innerHTML =
rows.map(payment => {
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

if (days !== null) {
if (days < 0) {
status =
"Fecha anterior";
cls =
"bad";
} else if (days <= 30) {
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
<td>${escapeHtml(
payment.date
? monthYear(payment.date)
: payment.period || "—"
)}</td>

<td>${escapeHtml(payment.partida || "—")}</td>

<td>${escapeHtml(payment.item || "—")}</td>

<td>${escapeHtml(payment.supplier || "—")}</td>

<td>Pago ${payment.paymentNumber}</td>

<td>
<span class="currency-pill">
${escapeHtml(payment.currency)}
</span>
</td>

<td class="num">
${formatOriginalMoney(
payment.amount,
payment.currency
)}
</td>

<td class="num money">
${money(payment.dop)}
</td>

<td>
<span class="status-pill ${cls}">
${status}
</span>
</td>
</tr>
`;
}).join("");
}

function setupCashflowCurrency() {
const select =
$("cashflowCurrencyView");

if (!select) return;

select.addEventListener(
"change",
() => {
renderCashflow();
}
);
}

function renderTracking() {
const withPO =
state.budget.filter(
row =>
row.status === "Con OC"
);

const pending =
state.budget.filter(
row =>
row.status !== "Con OC"
);

const noSupplier =
state.budget.filter(
row => !row.supplier
);

const noDate =
state.budget.filter(
row => !row.date
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

renderBottlenecks(pending);
renderSupplierPO(withPO);
renderTrackingTable(pending);
}

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
label: "Necesidades",
value: total,
detail: "Ítems identificados"
},
{
label: "Proveedor definido",
value:
total - noSupplier,
detail: "Con proveedor"
},
{
label: "Orden colocada",
value: withPO,
detail: "Con OC"
},
{
label: "Pendientes",
value: pending,
detail:
`${noDate} sin fecha`
}
];

container.innerHTML =
stages.map(stage => `
<div class="stage">
<span>${stage.label}</span>
<b>${stage.value}</b>
<small>${stage.detail}</small>
</div>
`).join("");
}

function renderBottlenecks(pending) {
const top =
[...pending]
.sort(
(a, b) =>
b.current - a.current
)
.slice(0, 10);

createChart(
"bottleneckChart",
"bar",
top.map(
row => row.item
),
top.map(
row => row.current
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

function renderSupplierPO(withPO) {
const groups = {};

withPO.forEach(row => {
const supplier =
row.supplier ||
"Sin proveedor";

groups[supplier] =
(groups[supplier] || 0) + 1;
});

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
entries.map(x => x[0]),
entries.map(x => x[1]),
{
plugins: {
legend: {
display: false
}
}
}
);
}

function renderTrackingTable(pending) {
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
b.current - a.current
)
.map(row => `
<tr>
<td>${escapeHtml(row.item)}</td>
<td>${escapeHtml(row.partida)}</td>
<td>${escapeHtml(row.supplier || "—")}</td>
<td>
<span class="status-pill ${
row.supplier
? "warn"
: "bad"
}">
${escapeHtml(row.status)}
</span>
</td>
<td>${escapeHtml(row.comment || "—")}</td>
<td class="num money">${money(row.current)}</td>
</tr>
`).join("");
}

function renderPending() {
const issues =
state.issues;

const owners =
unique(
issues.map(
issue => issue.owner
)
);

const withComments =
issues.filter(
issue => issue.comment
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

renderOwnerChart(issues);
renderConsiderations();
renderIssueBoard(issues);
}

function renderOwnerChart(issues) {
const groups = {};

issues.forEach(issue => {
groups[issue.owner] =
(groups[issue.owner] || 0) + 1;
});

const entries =
Object.entries(groups)
.sort(
(a, b) =>
b[1] - a[1]
);

createChart(
"ownerChart",
"bar",
entries.map(x => x[0]),
entries.map(x => x[1]),
{
plugins: {
legend: {
display: false
}
}
}
);
}

function renderConsiderations() {
const container =
$("considerations");

if (!container) return;

let items =
state.considerations.map(row =>
valueFrom(row, [
"Consideración",
"Consideracion",
"Nota",
"Descripción",
"Descripcion"
])
).filter(Boolean);

if (!items.length) {
items = [
"Validar fechas de entrega de partidas críticas.",
"Confirmar proveedor y orden de compra antes del corte.",
"Actualizar comentarios de pendientes abiertos.",
"Revisar pagos próximos contra el flujo de caja.",
"Mantener la plantilla actualizada para conservar la lectura ejecutiva."
];
}

container.innerHTML =
items
.slice(0, 8)
.map(item => `
<div class="consideration">
${escapeHtml(item)}
</div>
`).join("");
}

function renderIssueBoard(issues) {
const container =
$("issueBoard");

if (!container) return;

container.innerHTML =
issues.length
? issues.map(issue => `
<article class="issue-card">
<b>${escapeHtml(issue.title)}</b>
<div class="owner">
${escapeHtml(issue.owner)}
</div>
<p>
${escapeHtml(
issue.comment ||
"Sin comentario registrado."
)}
</p>
</article>
`).join("")
: `
<section class="card">
No hay pendientes registrados.
</section>
`;
}

function setupLiquidation() {
[
"liqAmount",
"liqExw",
"liqFob",
"liqContainers",
"liqRate"
].forEach(id => {
const input = $(id);

if (input) {
input.addEventListener(
"input",
renderLiquidation
);
}
});
}

function renderLiquidation() {
const container =
$("liquidationScenarios");

if (!container) return;

const amount =
numberInput("liqAmount", 10000);

const exw =
numberInput("liqExw", 0);

const fob =
numberInput("liqFob", 1000);

const containers =
numberInput("liqContainers", 1);

const rate =
numberInput("liqRate", 57.85);

state.exchangeRate =
rate;

const scenarios = [
{
country: "China",
freight: 4200,
duty: .20,
insurance: .01,
other: 1200
},
{
country: "Estados Unidos",
freight: 2200,
duty: .20,
insurance: .01,
other: 900
},
{
country: "Europa",
freight: 3500,
duty: .20,
insurance: .01,
other: 1100
},
{
country: "México",
freight: 2600,
duty: .20,
insurance: .01,
other: 950
},
{
country: "Puerto Rico",
freight: 1600,
duty: .20,
insurance: .01,
other: 800
}
];

container.innerHTML =
scenarios.map(scenario => {
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
taxable * .18;

const total =
taxable +
itbis;

return `
<article class="scenario">

<h3>${scenario.country}</h3>

<div class="metric">
<span>Compra</span>
<b>${usd(amount)}</b>
</div>

<div class="metric">
<span>Flete estimado</span>
<b>${usd(freight)}</b>
</div>

<div class="metric">
<span>Seguro</span>
<b>${usd(insurance)}</b>
</div>

<div class="metric">
<span>CIF</span>
<b>${usd(cif)}</b>
</div>

<div class="metric">
<span>Gravamen</span>
<b>${usd(duty)}</b>
</div>

<div class="metric">
<span>Otros gastos</span>
<b>${usd(scenario.other)}</b>
</div>

<div class="metric">
<span>ITBIS</span>
<b>${usd(itbis)}</b>
</div>

<div class="total">
<span>Costo estimado puesto RD</span>
<b>${usd(total)}</b>
<small>
≈ ${money(total * rate)}
</small>
</div>

</article>
`;
}).join("");
}

function groupSupplierValue() {
const groups = {};

state.budget.forEach(row => {
if (!row.supplier) return;

groups[row.supplier] =
(groups[row.supplier] || 0) +
(row.current || row.ordered);
});

const entries =
Object.entries(groups)
.sort(
(a, b) =>
b[1] - a[1]
)
.slice(0, 12);

return {
labels:
entries.map(x => x[0]),
values:
entries.map(x => x[1])
};
}

function groupSum(rows, labelKey, valueKey) {
const groups = {};

rows.forEach(row => {
const label =
row[labelKey] ||
"Sin clasificar";

const value =
Number(row[valueKey]) || 0;

groups[label] =
(groups[label] || 0) +
value;
});

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
entries.map(x => x[0]),
values:
entries.map(x => x[1])
};
}

function toMap(group) {
const map = {};

group.labels.forEach(
(label, index) => {
map[label] =
group.values[index];
}
);

return map;
}

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
type === "doughnut"
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
canvas.getContext("2d"),
{
type,
data: {
labels,
datasets: [
{
data: values,
backgroundColor: background,
borderColor:
type === "line"
? COLORS.blue
: undefined,
tension: .3,
fill:
type === "line"
? false
: undefined
}
]
},
options: mergeChartOptions(
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
canvas.getContext("2d"),
{
type,
data: {
labels,
datasets
},
options:
mergeChartOptions(options)
}
);
}

function mergeChartOptions(custom) {
const base = {
responsive: true,
maintainAspectRatio: false,
interaction: {
intersect: false,
mode: "index"
},
plugins: {
legend: {
labels: {
boxWidth: 10,
boxHeight: 10,
font: {
size: 10
}
}
},
tooltip: {
callbacks: {
label(context) {
const value =
Number(context.raw);

if (
Number.isFinite(value) &&
Math.abs(value) >= 1000
) {
return `${context.dataset.label || ""} ${money(value)}`;
}

return `${context.dataset.label || ""} ${value}`;
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
maxRotation: 45,
minRotation: 0
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
color: "#edf2f6"
}
}
}
};

return deepMerge(
base,
custom
);
}

function deepMerge(target, source) {
const output =
{ ...target };

Object.keys(source || {})
.forEach(key => {
if (
source[key] &&
typeof source[key] === "object" &&
!Array.isArray(source[key])
) {
output[key] =
deepMerge(
target[key] || {},
source[key]
);
} else {
output[key] =
source[key];
}
});

return output;
}

function destroyChart(id) {
if (state.charts[id]) {
state.charts[id].destroy();
delete state.charts[id];
}
}

let countdownInterval = null;

function startCountdown(dateString) {
if (countdownInterval) {
clearInterval(
countdownInterval
);
}

const target =
parseDate(dateString);

function update() {
if (!target) return;

const now =
new Date();

let diff =
target - now;

const status =
$("countdownStatus");

if (diff <= 0) {
diff = 0;

if (status) {
status.textContent =
"Fecha de apertura alcanzada";
}
} else if (status) {
status.textContent =
"Hacia la apertura";
}

const days =
Math.floor(
diff / 86400000
);

const hours =
Math.floor(
(
diff % 86400000
) /
3600000
);

const minutes =
Math.floor(
(
diff % 3600000
) /
60000
);

const seconds =
Math.floor(
(
diff % 60000
) /
1000
);

setText(
"cdDays",
days
);

setText(
"cdHours",
String(hours).padStart(2, "0")
);

setText(
"cdMinutes",
String(minutes).padStart(2, "0")
);

setText(
"cdSeconds",
String(seconds).padStart(2, "0")
);
}

update();

countdownInterval =
setInterval(
update,
1000
);
}

function renderEmptyState() {
setText(
"overviewSubtitle",
"No fue posible cargar la plantilla del proyecto. Verifica la ruta del archivo Excel."
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
id => setText(id, "—")
);
}

function valueFrom(row, candidates) {
const keys =
Object.keys(row);

for (const candidate of candidates) {
const exact =
keys.find(
key =>
normalizeText(key) ===
normalizeText(candidate)
);

if (
exact &&
row[exact] !== undefined &&
row[exact] !== null &&
String(row[exact]).trim() !== ""
) {
return row[exact];
}
}

for (const candidate of candidates) {
const partial =
keys.find(
key =>
normalizeText(key).includes(
normalizeText(candidate)
)
);

if (
partial &&
row[partial] !== undefined &&
row[partial] !== null &&
String(row[partial]).trim() !== ""
) {
return row[partial];
}
}

return "";
}

function numberFrom(row, candidates) {
return toNumber(
valueFrom(
row,
candidates
)
);
}

function toNumber(value) {
if (
value === null ||
value === undefined ||
value === ""
) {
return 0;
}

if (
typeof value === "number"
) {
return value;
}

let text =
String(value)
.trim()
.replace(/\s/g, "")
.replace(/RD\$/gi, "")
.replace(/US\$/gi, "")
.replace(/USD/gi, "")
.replace(/\$/g, "");

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
.replace(/\./g, "")
.replace(",", ".");
} else {
text =
text.replace(/,/g, "");
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
parts[0].replace(/\./g, "") +
"." +
parts[1];
} else {
text =
text.replace(/,/g, "");
}
}

text =
text.replace(
/[^0-9.-]/g,
""
);

const number =
Number(text);

return Number.isFinite(number)
? number
: 0;
}

function normalizeText(value) {
return String(value || "")
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

function parseExcelDate(value) {
if (!value) return null;

if (
value instanceof Date &&
!isNaN(value)
) {
return value;
}

if (
typeof value === "number"
) {
const parsed =
XLSX.SSF.parse_date_code(
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
String(value).trim();

if (!text) return null;

const direct =
new Date(text);

if (!isNaN(direct)) {
return direct;
}

const match =
text.match(
/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/
);

if (match) {
let year =
Number(match[3]);

if (year < 100) {
year += 2000;
}

return new Date(
year,
Number(match[2]) - 1,
Number(match[1])
);
}

return null;
}

function parseDate(value) {
if (!value) return null;

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
if (!date) return "—";

return date.toLocaleDateString(
"es-DO",
{
day: "2-digit",
month: "short",
year: "numeric"
}
);
}

function formatLongDate(date) {
if (!date) return "—";

return date.toLocaleDateString(
"es-DO",
{
day: "numeric",
month: "long",
year: "numeric"
}
);
}

function formatPossibleDate(value) {
const date =
parseExcelDate(value);

return date
? formatDate(date)
: value || "—";
}

function monthYear(date) {
if (!date) return "—";

return date.toLocaleDateString(
"es-DO",
{
month: "short",
year: "numeric"
}
);
}

function money(value) {
const number =
Number(value) || 0;

if (
Math.abs(number) >= 1000000000
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
Math.abs(number) >= 1000000
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

return new Intl.NumberFormat(
"es-DO",
{
style: "currency",
currency: "DOP",
maximumFractionDigits: 0
}
).format(number);
}

function usd(value) {
return new Intl.NumberFormat(
"en-US",
{
style: "currency",
currency: "USD",
maximumFractionDigits: 0
}
).format(
Number(value) || 0
);
}

function formatOriginalMoney(
value,
currency
) {
const c =
String(currency || "DOP")
.toUpperCase();

if (
c.includes("USD") ||
normalizeText(c).includes("dolar")
) {
return usd(value);
}

return money(value);
}

function percent(value) {
return new Intl.NumberFormat(
"es-DO",
{
style: "percent",
maximumFractionDigits: 1
}
).format(
Number(value) || 0
);
}

function sum(values) {
return values.reduce(
(total, value) =>
total +
(Number(value) || 0),
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
String(value).trim() !== ""
)
)
];
}

function setText(id, value) {
const element =
$(id);

if (element) {
element.textContent =
value;
}
}

function numberInput(id, fallback) {
const element =
$(id);

if (!element) return fallback;

const value =
Number(element.value);

return Number.isFinite(value)
? value
: fallback;
}

function showLoader(show) {
const loader =
$("loader");

if (!loader) return;

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
.replace(/&/g, "&amp;")
.replace(/</g, "&lt;")
.replace(/>/g, "&gt;")
.replace(/"/g, "&quot;")
.replace(/'/g, "&#039;");
}

function emptyMessage(text) {
return `
<div class="muted">
${escapeHtml(text)}
</div>
`;
}
