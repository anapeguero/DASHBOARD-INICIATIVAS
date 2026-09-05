
const $=id=>document.getElementById(id);
const charts={};
let projectConfig=null, WB=null, model=null;

const money=n=>{
  n=Number(n)||0;
  return new Intl.NumberFormat('es-DO',{style:'currency',currency:'DOP',maximumFractionDigits:0}).format(n);
};
const compactMoney=n=>{
  n=Number(n)||0;
  if(Math.abs(n)>=1e9)return `RD$ ${(n/1e9).toFixed(2)} B`;
  if(Math.abs(n)>=1e6)return `RD$ ${(n/1e6).toFixed(2)} MM`;
  if(Math.abs(n)>=1e3)return `RD$ ${(n/1e3).toFixed(0)} mil`;
  return money(n);
};
const pct=n=>`${(Number(n)||0).toFixed(1)}%`;
const esc=s=>String(s??'').replace(/[&<>"']/g,m=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#039;'}[m]));
const val=(r,i)=>r?.[i]??null;
const num=v=>{const n=Number(v);return Number.isFinite(n)?n:0};
const text=v=>String(v??'').trim();
const isOrder=v=>{
  const s=text(v);
  if(!s)return false;
  if(/pendiente|plan|no aplica/i.test(s))return false;
  return /\d{6,}/.test(s)||/^\d+$/.test(s);
};
const excelDate=v=>{
  const n=Number(v);
  if(!Number.isFinite(n)||n<20000||n>80000)return null;
  const d=new Date(Date.UTC(1899,11,30)+n*86400000);
  return d;
};
const fmtDate=v=>{
  const d=v instanceof Date?v:excelDate(v);
  return d?new Intl.DateTimeFormat('es-DO',{day:'2-digit',month:'short',year:'numeric',timeZone:'UTC'}).format(d):'—';
};
const daysFromToday=d=>{
  if(!(d instanceof Date))return null;
  const now=new Date(); const today=Date.UTC(now.getFullYear(),now.getMonth(),now.getDate());
  return Math.round((d.getTime()-today)/86400000);
};
const sheet=name=>WB.Sheets[name];
const rows=name=>sheet(name)?XLSX.utils.sheet_to_json(sheet(name),{header:1,raw:true,defval:null}):[];

function destroyChart(id){if(charts[id]){charts[id].destroy();delete charts[id]}}
function chart(id,type,labels,datasets,opts={}){
  destroyChart(id);
  const c=$(id); if(!c)return;
  charts[id]=new Chart(c,{type,data:{labels,datasets},options:{
    responsive:true,maintainAspectRatio:false,
    plugins:{legend:{display:opts.legend??true,position:'bottom',labels:{boxWidth:9,font:{size:9}}},tooltip:{callbacks:{label:ctx=>opts.money?`${ctx.dataset.label||''}: ${money(ctx.raw)}`:`${ctx.dataset.label||''}: ${ctx.raw}`}}},
    scales: opts.noScales?{}:{
      x:{grid:{display:false},ticks:{font:{size:9},maxRotation:0,autoSkip:true}},
      y:{beginAtZero:true,grid:{color:'#edf2f6'},ticks:{font:{size:9},callback:v=>opts.money?compactMoney(v):v}}
    },
    ...opts.extra
  }});
}

function groupSum(arr,keyFn,valFn){
  const m=new Map();
  arr.forEach(x=>{const k=keyFn(x)||'Sin definir';m.set(k,(m.get(k)||0)+(valFn(x)||0))});
  return [...m.entries()].map(([label,value])=>({label,value}));
}
function uniq(arr){return [...new Set(arr.filter(Boolean))]}

function buildModel(){
  const plan=rows('Plan OI');
  const bd=rows('BD Plan Detallado');
  const cron=rows('Cronograma Inverso');
  const seg=rows('Seguimiento Mao ')?.length?rows('Seguimiento Mao '):rows('Seguimiento Mao');
  const pen=rows('Puntos pendientes');
  const pre=rows('Pre-Liquidar');

  const totalBudget=num(val(plan[6],8))||num(val(plan[6],6));
  const summary=(plan.slice(9,13)||[]).map(r=>({name:text(val(r,2)),value:num(val(r,8))||num(val(r,6))})).filter(x=>x.name&&x.value>=0);

  // BD Plan Detallado: only lines with actual item description are procurement detail rows.
  const purchases=bd.slice(4).map((r,idx)=>({
    row:idx+5,pos:text(val(r,0)),category:text(val(r,2)),budgetType:text(val(r,3)),
    plan:num(val(r,4)),item:text(val(r,5)),po:text(val(r,6)),date:excelDate(val(r,7)),
    qty:num(val(r,8)),currency:text(val(r,9)),rate:num(val(r,10)),unitPrice:num(val(r,11)),
    net:num(val(r,12)),liquidation:num(val(r,13)),budget:num(val(r,14)),usd:num(val(r,15)),
    comment:text(val(r,16)),supplier:text(val(r,17))
  })).filter(r=>r.item && !/^total/i.test(r.item));

  purchases.forEach(r=>{
    if(/pendiente/i.test(r.po))r.status='Pendiente de compra';
    else if(isOrder(r.po))r.status='Con OC';
    else r.status='Sin OC definida';
  });

  const scheduled=cron.slice(14).map((r,idx)=>({
    row:idx+15,pos:text(val(r,0)),group:text(val(r,1)),asset:text(val(r,2)),
    design:excelDate(val(r,18)),poDate:excelDate(val(r,19)),arrival:excelDate(val(r,20)),delivery:excelDate(val(r,21))||excelDate(val(r,17)),
    status:text(val(r,22)),condition:text(val(r,23)),observation:text(val(r,24)),
    lead:num(val(r,25)),approved:num(val(r,26)),currency:text(val(r,27)),executed:num(val(r,28)),deviation:num(val(r,29)),
    actualPODate:excelDate(val(r,30)),po:text(val(r,31))
  })).filter(r=>r.asset && !/activo fijo/i.test(r.asset));

  const opening=excelDate(val(cron[7],2)) || excelDate(val(seg[5],1));

  const logistics=seg.slice(8).map((r,idx)=>({
    row:idx+9,initiative:text(val(r,0)),supplierCode:text(val(r,1)),supplier:text(val(r,2)),po:text(val(r,3)),
    orderDate:excelDate(val(r,4)),description:text(val(r,5)),origin:text(val(r,6)),purchaseType:text(val(r,7)),
    term:text(val(r,8)),currency:text(val(r,9)),subtotal:num(val(r,12)),tax:num(val(r,13)),total:num(val(r,14)),dop:num(val(r,15)),
    leadRaw:val(r,16),eta1:excelDate(val(r,17)),p1:num(val(r,18)),pay1:num(val(r,19)),p2:num(val(r,20)),pay2:num(val(r,21)),p3:num(val(r,22)),pay3:num(val(r,23)),
    balanceDate:excelDate(val(r,24)),eta:excelDate(val(r,25))||excelDate(val(r,17)),daysRaw:val(r,26),observation:text(val(r,27))
  })).filter(r=>r.supplier||r.po||r.description);

  const issues=pen.slice(1,13).map((r,idx)=>({row:idx+2,issue:text(val(r,0)),owner:text(val(r,1)),comment:text(val(r,2))})).filter(x=>x.issue);
  const considerations=pen.slice(14,22).map(r=>text(val(r,0))).filter(Boolean);

  const preParams={
    rate:num(val(pre[4],0))||57.85,
    defaultAmount:num(val(pre[1],2))||10000,
    defaultExw:num(val(pre[2],2)),
    defaultFob:num(val(pre[3],2))||1000,
    origins:[
      {name:'China',freightUSD:num(val(pre[9],1)),duty:.20,agent:num(val(pre[16],2)),local:num(val(pre[17],2)),customs:num(val(pre[18],2)),port:num(val(pre[19],2))},
      {name:'Europa',freightUSD:num(val(pre[9],4)),duty:.14,agent:num(val(pre[16],5)),local:num(val(pre[17],5)),customs:num(val(pre[18],5)),port:num(val(pre[19],5))},
      {name:'USA / México / Canadá',freightUSD:num(val(pre[9],7)),duty:0,agent:num(val(pre[16],8)),local:num(val(pre[17],8)),customs:num(val(pre[18],8)),port:num(val(pre[19],8))},
      {name:'América del Sur',freightUSD:num(val(pre[9],10)),duty:.20,agent:num(val(pre[16],11)),local:num(val(pre[17],11)),customs:num(val(pre[18],11)),port:num(val(pre[19],11))},
      {name:'España',freightUSD:num(val(pre[9],13)),duty:.14,agent:num(val(pre[16],14)),local:num(val(pre[17],14)),customs:num(val(pre[18],14)),port:num(val(pre[19],14))}
    ]
  };

  return {plan,bd,cron,seg,pen,pre,totalBudget,summary,purchases,scheduled,opening,logistics,issues,considerations,preParams};
}

function renderOverview(){
  const p=model.purchases;
  const total=model.totalBudget;
  const ordered=p.filter(x=>x.status==='Con OC');
  const pending=p.filter(x=>x.status==='Pendiente de compra');
  const orderedValue=ordered.reduce((s,x)=>s+x.budget,0);
  const pendingValue=pending.reduce((s,x)=>s+x.budget,0);
  const coverage=(ordered.length+pending.length)?ordered.length/(ordered.length+pending.length)*100:0;
  const suppliers=uniq(p.map(x=>x.supplier));

  $('kpiBudget').textContent=compactMoney(total);
  $('kpiOrdered').textContent=compactMoney(orderedValue);
  $('kpiPendingBuy').textContent=compactMoney(pendingValue);
  $('kpiCoverage').textContent=pct(coverage);
  $('kpiSuppliers').textContent=suppliers.length;
  $('kpiOpenIssues').textContent=model.issues.length;

  const noSupplier=p.filter(x=>!x.supplier).length;
  const noDate=p.filter(x=>!x.date).length;
  const schedWithDelivery=model.scheduled.filter(x=>x.delivery).length;
  const logNoEta=model.logistics.filter(x=>!x.eta).length;

  const health=[
    {name:'Presupuesto',state:'info',title:'Plan disponible',desc:`${money(total)} cargados en el plan detallado.`},
    {name:'Compras',state:coverage>=80?'good':coverage>=50?'warn':'bad',title:`${pct(coverage)} con OC`,desc:`${pending.length} partidas están marcadas como pendientes de compra.`},
    {name:'Cronograma',state:schedWithDelivery/model.scheduled.length>=.8?'good':'warn',title:`${schedWithDelivery}/${model.scheduled.length} con entrega`,desc:'Cobertura de fechas de entrega registradas en el cronograma inverso.'},
    {name:'Pendientes',state:model.issues.length<=3?'good':model.issues.length<=7?'warn':'bad',title:`${model.issues.length} abiertos`,desc:'Puntos pendientes identificados en la plantilla.'}
  ];
  $('healthStrip').innerHTML=health.map(h=>`<div class="health-item"><div class="row"><b>${h.name}</b><i class="dot ${h.state}"></i></div><strong>${h.title}</strong><p>${h.desc}</p></div>`).join('');

  const sm=model.summary.filter(x=>x.value>0);
  chart('budgetMixChart','doughnut',sm.map(x=>x.name),[{data:sm.map(x=>x.value),borderWidth:0}],{money:true,noScales:true,legend:true});

  const attention=[];
  if(pending.length)attention.push({c:'var(--red)',t:`${pending.length} partidas pendientes de compra`,d:`Representan aproximadamente ${compactMoney(pendingValue)} en el detalle.`});
  if(noSupplier)attention.push({c:'var(--yellow)',t:`${noSupplier} ítems sin proveedor definido`,d:'Completar proveedor mejora la trazabilidad de procurement.'});
  if(noDate)attention.push({c:'var(--yellow)',t:`${noDate} ítems sin fecha de orden/plan válida`,d:'Revisar fechas en BD Plan Detallado.'});
  if(model.issues.length)attention.push({c:'var(--cyan)',t:`${model.issues.length} puntos pendientes`,d:'Revisar responsables y comentarios antes de la próxima reunión.'});
  if(logNoEta)attention.push({c:'var(--yellow)',t:`${logNoEta} órdenes logísticas sin ETA válida`,d:'Evita depender de fechas no estructuradas.'});
  $('attentionList').innerHTML=attention.slice(0,5).map(a=>`<div class="attention"><i style="background:${a.c}"></i><div><b>${a.t}</b><span>${a.d}</span></div></div>`).join('');

  const statusData=[
    {label:'Con OC',value:ordered.length},
    {label:'Pendiente de compra',value:pending.length},
    {label:'Sin OC definida',value:p.filter(x=>x.status==='Sin OC definida').length}
  ];
  chart('purchaseStatusChart','doughnut',statusData.map(x=>x.label),[{data:statusData.map(x=>x.value),borderWidth:0}],{noScales:true});

  const sup=groupSum(p.filter(x=>x.supplier),x=>x.supplier,x=>x.budget).sort((a,b)=>b.value-a.value).slice(0,6);
  $('topSuppliers').innerHTML=sup.map((x,i)=>`<div class="rank-row"><span class="n">${i+1}</span><b title="${esc(x.label)}">${esc(x.label)}</b><span>${compactMoney(x.value)}</span></div>`).join('')||'<span class="muted">Sin proveedores identificados.</span>';

  const quality=[
    ['Ítems sin proveedor',noSupplier,'BD Plan Detallado'],
    ['Ítems sin fecha',noDate,'BD Plan Detallado'],
    ['Logística sin ETA',logNoEta,'Seguimiento Mao'],
    ['Cronograma sin entrega',model.scheduled.length-schedWithDelivery,'Cronograma Inverso']
  ];
  $('dataQuality').innerHTML=quality.map(x=>`<div class="quality"><div class="top"><b>${x[0]}</b><strong>${x[1]}</strong></div><small>${x[2]}</small></div>`).join('');
}

function setupBudgetFilters(){
  const cats=uniq(model.purchases.map(x=>x.category)).sort();
  const sups=uniq(model.purchases.map(x=>x.supplier)).sort();
  $('budgetCategoryFilter').innerHTML='<option value="">Todas</option>'+cats.map(x=>`<option>${esc(x)}</option>`).join('');
  $('budgetSupplierFilter').innerHTML='<option value="">Todos</option>'+sups.map(x=>`<option>${esc(x)}</option>`).join('');
}
function filteredPurchases(){
  const c=$('budgetCategoryFilter').value,s=$('budgetStatusFilter').value,p=$('budgetSupplierFilter').value;
  return model.purchases.filter(x=>(!c||x.category===c)&&(!s||x.status===s)&&(!p||x.supplier===p));
}
function renderBudget(){
  const arr=filteredPurchases(), ordered=arr.filter(x=>x.status==='Con OC'), pending=arr.filter(x=>x.status==='Pendiente de compra');
  $('budgetFilteredTotal').textContent=compactMoney(arr.reduce((s,x)=>s+x.budget,0));
  $('budgetFilteredOrdered').textContent=compactMoney(ordered.reduce((s,x)=>s+x.budget,0));
  $('budgetFilteredPending').textContent=compactMoney(pending.reduce((s,x)=>s+x.budget,0));
  $('budgetFilteredItems').textContent=arr.length;
  const cat=groupSum(arr,x=>x.category,x=>x.budget).sort((a,b)=>b.value-a.value).slice(0,12);
  chart('categoryChart','bar',cat.map(x=>x.label.replace(/^\d+(\.\d+)?\s*/,'').slice(0,28)),[{label:'Presupuesto',data:cat.map(x=>x.value)}],{money:true,legend:false,extra:{indexAxis:'y'}});
  const sup=groupSum(arr.filter(x=>x.supplier),x=>x.supplier,x=>x.budget).sort((a,b)=>b.value-a.value).slice(0,10);
  chart('supplierChart','bar',sup.map(x=>x.label.slice(0,24)),[{label:'Valor',data:sup.map(x=>x.value)}],{money:true,legend:false,extra:{indexAxis:'y'}});
  $('budgetTableCount').textContent=`${arr.length} registros`;
  $('budgetTableBody').innerHTML=arr.sort((a,b)=>b.budget-a.budget).map(x=>`<tr>
    <td>${esc(x.category)}</td><td><b>${esc(x.item)}</b><br><span class="muted">${esc(x.comment)}</span></td>
    <td>${esc(x.supplier)||'<span class="muted">Sin definir</span>'}</td>
    <td><span class="status-pill ${x.status==='Con OC'?'good':x.status==='Pendiente de compra'?'warn':'info'}">${esc(x.po||x.status)}</span></td>
    <td>${fmtDate(x.date)}</td><td class="num money">${money(x.budget)}</td><td>${esc(x.currency)}</td></tr>`).join('');
}

function renderTracking(){
  const p=model.purchases, withPO=p.filter(x=>x.status==='Con OC'), pending=p.filter(x=>x.status==='Pendiente de compra');
  const noSupplier=p.filter(x=>!x.supplier), noDate=p.filter(x=>!x.date);
  $('trackWithPO').textContent=withPO.length;$('trackPending').textContent=pending.length;$('trackNoSupplier').textContent=noSupplier.length;$('trackNoDate').textContent=noDate.length;
  const stages=[
    ['Partidas',p.length,'Ítems identificados'],
    ['Proveedor definido',p.filter(x=>x.supplier).length,'Con proveedor registrado'],
    ['OC colocada',withPO.length,'Orden identificada'],
    ['Seguimiento logístico',model.logistics.filter(x=>x.po).length,'Órdenes en hoja logística']
  ];
  $('pipeline').innerHTML=stages.map(x=>`<div class="stage"><span>${x[0]}</span><b>${x[1]}</b><small>${x[2]}</small></div>`).join('');
  const bot=pending.sort((a,b)=>b.budget-a.budget).slice(0,10);
  chart('bottleneckChart','bar',bot.map(x=>x.item.slice(0,28)),[{label:'Pendiente',data:bot.map(x=>x.budget)}],{money:true,legend:false,extra:{indexAxis:'y'}});
  const poSup=groupSum(withPO.filter(x=>x.supplier),x=>x.supplier,x=>1).sort((a,b)=>b.value-a.value).slice(0,10);
  chart('supplierPOChart','bar',poSup.map(x=>x.label.slice(0,24)),[{label:'Órdenes / líneas',data:poSup.map(x=>x.value)}],{legend:false,extra:{indexAxis:'y'}});
  const manage=p.filter(x=>x.status!=='Con OC'||!x.supplier||!x.date).sort((a,b)=>b.budget-a.budget);
  $('trackingCount').textContent=`${manage.length} registros`;
  $('trackingTable').innerHTML=manage.slice(0,200).map(x=>`<tr><td><b>${esc(x.item)}</b></td><td>${esc(x.category)}</td><td>${esc(x.supplier)||'—'}</td><td><span class="status-pill ${x.status==='Pendiente de compra'?'warn':'info'}">${esc(x.status)}</span></td><td>${esc(x.comment)||'—'}</td><td class="num money">${money(x.budget)}</td></tr>`).join('');
}

function renderSchedule(){
  const arr=model.scheduled;
  $('scheduleOpening').textContent=fmtDate(model.opening);
  $('scheduleItems').textContent=arr.length;
  const def=arr.filter(x=>x.delivery).length;$('scheduleDefined').textContent=def;
  const past=arr.filter(x=>x.delivery&&daysFromToday(x.delivery)<0).length;$('schedulePast').textContent=past;
  if(model.opening && daysFromToday(model.opening)<-30){
    $('scheduleDataWarning').innerHTML='<b>Atención de calidad de datos:</b> la fecha de apertura registrada en la plantilla está en el pasado respecto a la fecha actual. El dashboard conserva la fecha del archivo y marca vencimientos, pero conviene actualizar el cronograma antes de usarlo para decisiones.';
  } else $('scheduleDataWarning').innerHTML='';

  const milestones=arr.filter(x=>x.delivery).sort((a,b)=>a.delivery-b.delivery).slice(-12);
  $('timeline').innerHTML=milestones.map(x=>`<div class="milestone"><div class="date">${fmtDate(x.delivery)}</div><div class="track"><i></i></div><div><b>${esc(x.asset)}</b><p>${esc(x.group)} ${x.status?'· '+esc(x.status):''}</p></div></div>`).join('')||'<span class="muted">No hay fechas de entrega estructuradas.</span>';
  $('scheduleTableCount').textContent=`${arr.length} activos`;
  $('scheduleTable').innerHTML=arr.slice(0,250).map(x=>`<tr><td><b>${esc(x.asset)}</b></td><td>${esc(x.group)}</td><td>${fmtDate(x.design)}</td><td>${fmtDate(x.poDate||x.actualPODate)}</td><td>${fmtDate(x.arrival)}</td><td>${fmtDate(x.delivery)}</td><td>${esc(x.status)||'—'}</td><td>${esc(x.observation)||'—'}</td></tr>`).join('');
}

function renderImports(){
  const arr=model.logistics;
  $('impOrders').textContent=arr.length;
  $('impValue').textContent=compactMoney(arr.reduce((s,x)=>s+x.dop,0));
  const intl=arr.filter(x=>x.origin&&!/^(rd|rep|dominicana)/i.test(x.origin));
  $('impInternational').textContent=intl.length;
  const noEta=arr.filter(x=>!x.eta);$('impNoEta').textContent=noEta.length;
  const origins=groupSum(arr,x=>x.origin||'Sin origen',x=>x.dop).sort((a,b)=>b.value-a.value);
  chart('originChart','doughnut',origins.map(x=>x.label),[{data:origins.map(x=>x.value),borderWidth:0}],{money:true,noScales:true});
  const payments=[
    ['Pago 1',arr.reduce((s,x)=>s+x.pay1,0)],
    ['Pago 2',arr.reduce((s,x)=>s+x.pay2,0)],
    ['Pago 3',arr.reduce((s,x)=>s+x.pay3,0)]
  ];
  chart('paymentsChart','bar',payments.map(x=>x[0]),[{label:'Valor a pagar',data:payments.map(x=>x[1])}],{money:true,legend:false});
  $('importsTableCount').textContent=`${arr.length} órdenes / líneas`;
  $('importsTable').innerHTML=arr.map(x=>{
    const days=x.eta?daysFromToday(x.eta):null;
    const stat=days===null?'Sin ETA':days<0?`${Math.abs(days)} días vencida`:`${days} días`;
    const cls=days===null?'info':days<0?'bad':days<15?'warn':'good';
    return `<tr><td><b>${esc(x.supplier)}</b></td><td>${esc(x.po)}</td><td>${esc(x.description)}</td><td>${esc(x.origin)}</td><td>${esc(x.term)}</td><td>${fmtDate(x.orderDate)}</td><td>${fmtDate(x.eta)}</td><td><span class="status-pill ${cls}">${stat}</span></td><td class="num money">${money(x.dop)}</td></tr>`
  }).join('');
}

function renderPending(){
  const arr=model.issues;
  $('pendingTotal').textContent=arr.length;$('pendingOwners').textContent=uniq(arr.map(x=>x.owner)).length;$('pendingComments').textContent=arr.filter(x=>x.comment).length;$('pendingNoComments').textContent=arr.filter(x=>!x.comment).length;
  const owners=groupSum(arr,x=>x.owner||'Sin responsable',x=>1).sort((a,b)=>b.value-a.value);
  chart('ownerChart','bar',owners.map(x=>x.label),[{label:'Pendientes',data:owners.map(x=>x.value)}],{legend:false,extra:{indexAxis:'y'}});
  $('considerations').innerHTML=model.considerations.map(x=>`<div class="consideration">${esc(x)}</div>`).join('');
  $('issueBoard').innerHTML=arr.map(x=>`<article class="issue-card"><b>${esc(x.issue)}</b><div class="owner">${esc(x.owner)||'Sin responsable'}</div><p>${esc(x.comment)||'Sin comentario registrado.'}</p></article>`).join('');
}

function renderLiquidation(){
  const pp=model.preParams;
  $('liqRate').value=pp.rate||57.85;$('liqAmount').value=pp.defaultAmount||10000;$('liqExw').value=pp.defaultExw||0;$('liqFob').value=pp.defaultFob||1000;
  calcLiquidation();
}
function calcLiquidation(){
  const amount=num($('liqAmount').value),exw=num($('liqExw').value),fob=num($('liqFob').value),containers=Math.max(1,num($('liqContainers').value)),rate=num($('liqRate').value)||1;
  const originAmount=amount+exw+fob;
  $('liquidationScenarios').innerHTML=model.preParams.origins.map(o=>{
    const freight=o.freightUSD*rate*containers;
    const insurance=amount*.0025*rate;
    const cif=originAmount*rate+freight+insurance;
    const duty=cif*o.duty;
    const itbis=cif*.18;
    const misc=(o.agent+o.local+o.customs+o.port)*containers;
    const total=cif+duty+itbis+misc;
    const factor=(originAmount*rate)?total/(originAmount*rate):0;
    return `<article class="scenario"><h3>${esc(o.name)}</h3>
      <div class="metric"><span>Flete</span><b>${money(freight)}</b></div>
      <div class="metric"><span>Seguro</span><b>${money(insurance)}</b></div>
      <div class="metric"><span>CIF</span><b>${money(cif)}</b></div>
      <div class="metric"><span>Gravamen ${pct(o.duty*100)}</span><b>${money(duty)}</b></div>
      <div class="metric"><span>ITBIS</span><b>${money(itbis)}</b></div>
      <div class="metric"><span>Otros gastos</span><b>${money(misc)}</b></div>
      <div class="total"><span>Costo estimado</span><b>${money(total)}</b><small>Factor aprox. ${factor.toFixed(2)}x</small></div></article>`;
  }).join('');
}

function renderAll(){
  $('projectName').textContent=projectConfig.name;$('projectBrand').textContent=projectConfig.brand||'Proyecto';
  $('overviewTitle').textContent=projectConfig.displayName||projectConfig.name;
  $('footerProject').textContent=`Proyecto ${projectConfig.name}`;
  renderOverview();setupBudgetFilters();renderBudget();renderTracking();renderSchedule();renderImports();renderPending();renderLiquidation();
}
async function loadProject(id){
  const cfg=(window.projectList||[]).find(x=>x.id===id);if(!cfg)return;
  projectConfig=cfg;$('loader').classList.remove('hidden');
  try{
    const res=await fetch(cfg.file+'?v='+Date.now(),{cache:'no-store'});
    if(!res.ok)throw new Error('No se pudo abrir el Excel');
    const ab=await res.arrayBuffer();WB=XLSX.read(ab,{type:'array',cellDates:false});
    model=buildModel();renderAll();
  }catch(e){console.error(e);alert('Error cargando el proyecto: '+e.message)}
  finally{$('loader').classList.add('hidden')}
}
async function init(){
  const res=await fetch('data/projects.json?v='+Date.now(),{cache:'no-store'});const cfg=await res.json();window.projectList=cfg.projects||[];
  $('projectSelect').innerHTML=window.projectList.map(x=>`<option value="${esc(x.id)}">${esc(x.name)}</option>`).join('');
  $('projectSelect').addEventListener('change',e=>loadProject(e.target.value));
  if(window.projectList[0])loadProject(window.projectList[0].id);

  document.querySelectorAll('#navTabs button').forEach(b=>b.addEventListener('click',()=>{
    document.querySelectorAll('#navTabs button').forEach(x=>x.classList.remove('active'));b.classList.add('active');
    document.querySelectorAll('.page').forEach(x=>x.classList.remove('active'));$('page-'+b.dataset.page).classList.add('active');
    window.scrollTo({top:0,behavior:'smooth'});
  }));
  ['budgetCategoryFilter','budgetStatusFilter','budgetSupplierFilter'].forEach(id=>$(id).addEventListener('change',renderBudget));
  $('budgetReset').addEventListener('click',()=>{$('budgetCategoryFilter').value='';$('budgetStatusFilter').value='';$('budgetSupplierFilter').value='';renderBudget()});
  ['liqAmount','liqExw','liqFob','liqContainers','liqRate'].forEach(id=>$(id).addEventListener('input',calcLiquidation));
}
init();
