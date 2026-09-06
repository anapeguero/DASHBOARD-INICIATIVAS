
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
function axisMoney(v){
  const n=Number(v)||0, a=Math.abs(n);
  if(a>=1e9)return `RD$ ${(n/1e9).toFixed(a>=10e9?0:1)} B`;
  if(a>=1e6)return `RD$ ${(n/1e6).toFixed(a>=10e6?0:1)} MM`;
  if(a>=1e3)return `RD$ ${(n/1e3).toFixed(0)} mil`;
  return `RD$ ${Math.round(n).toLocaleString('es-DO')}`;
}
function chart(id,type,labels,datasets,opts={}){
  destroyChart(id);
  const c=$(id); if(!c)return;
  const horizontal=opts.extra?.indexAxis==='y';
  const moneyTicks={font:{size:9},callback:v=>axisMoney(v)};
  const categoryTicks={font:{size:9},maxRotation:0,autoSkip:true};
  const scales=opts.noScales?{}:(horizontal?{
    x:{beginAtZero:true,grid:{color:'#edf2f6'},ticks:opts.money?moneyTicks:{font:{size:9}}},
    y:{grid:{display:false},ticks:categoryTicks}
  }:{
    x:{grid:{display:false},ticks:categoryTicks},
    y:{beginAtZero:true,grid:{color:'#edf2f6'},ticks:opts.money?moneyTicks:{font:{size:9}}}
  });
  charts[id]=new Chart(c,{type,data:{labels,datasets},options:{
    responsive:true,maintainAspectRatio:false,
    plugins:{
      legend:{display:opts.legend??true,position:'bottom',labels:{boxWidth:9,font:{size:9}}},
      tooltip:{callbacks:{label:ctx=>opts.money?`${ctx.dataset.label||''}: ${money(ctx.raw)}`:`${ctx.dataset.label||''}: ${ctx.raw}`}}
    },
    scales,
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
  // Partidas maestras 1.00–9.00 únicamente. Se excluyen expresamente
  // subpartidas como 1.01, 2.06, 7.03, etc.
  const planRate=num(val(bd[1],9))||63;
  const part19=(plan.slice(20)||[])
    .map(r=>{
      const rawPart=num(val(r,0));
      const approvedUsd=num(val(r,4));
      const currentUsd=num(val(r,9));
      const currentDop=num(val(r,8));
      const approvedDop=approvedUsd*planRate;
      return {
        rawPart,
        part:rawPart,
        name:text(val(r,2)),
        approvedUsd,currentUsd,currentDop,approvedDop,
        savingsUsd:approvedUsd-currentUsd,
        savingsDop:approvedDop-currentDop
      };
    })
    .filter(x=>Number.isInteger(x.rawPart) && x.rawPart>=1 && x.rawPart<=9)
    .sort((a,b)=>a.part-b.part);

  const approvedBudgetUsd=num(val(plan[6],4));
  const currentBudgetUsd=num(val(plan[6],9));
  const approvedBudgetDop=approvedBudgetUsd*planRate;
  const currentBudgetDop=num(val(plan[6],8))||currentBudgetUsd*planRate;
  const projectSavingsDop=approvedBudgetDop-currentBudgetDop;


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

  const rateDollar=num(val(seg[5],13))||63;
  const rateEuro=num(val(seg[5],14))||76.5;
  const logistics=seg.slice(8).map((r,idx)=>{
    const currency=text(val(r,9));
    const grossSubtotal=num(val(r,10));
    const discount=num(val(r,11));
    const subtotal=num(val(r,12));
    const fx=/EUR/i.test(currency)?rateEuro:/USD/i.test(currency)?rateDollar:1;
    return {
      row:idx+9,initiative:text(val(r,0)),supplierCode:text(val(r,1)),supplier:text(val(r,2)),po:text(val(r,3)),
      orderDate:excelDate(val(r,4)),description:text(val(r,5)),origin:text(val(r,6)),purchaseType:text(val(r,7)),
      term:text(val(r,8)),currency,grossSubtotal,discount,discountDop:discount*fx,subtotal,
      tax:num(val(r,13)),total:num(val(r,14)),dop:num(val(r,15)),
      leadRaw:val(r,16),eta1:excelDate(val(r,17)),p1:num(val(r,18)),pay1:num(val(r,19)),p2:num(val(r,20)),pay2:num(val(r,21)),p3:num(val(r,22)),pay3:num(val(r,23)),
      balanceDate:excelDate(val(r,24)),eta:excelDate(val(r,25))||excelDate(val(r,17)),daysRaw:val(r,26),observation:text(val(r,27))
    };
  }).filter(r=>r.supplier||r.po||r.description);

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

  // Flujo de Caja: row 75 is the monthly total in the current project template.
  const cf=rows('Flujo de Caja');
  const cashflowStart = projectConfig?.cashflowStart ? new Date(projectConfig.cashflowStart+'T00:00:00') : new Date('2025-01-01T00:00:00');
  const cashflowMonths=[];
  if(cf.length>=75){
    const totalRow=cf[74]||[];
    for(let i=0;i<24;i++){
      const d=new Date(cashflowStart.getFullYear(),cashflowStart.getMonth()+i,1);
      cashflowMonths.push({
        period:i+1,
        label:new Intl.DateTimeFormat('es-DO',{month:'short',year:'numeric'}).format(d),
        fullLabel:new Intl.DateTimeFormat('es-DO',{month:'long',year:'numeric'}).format(d),
        date:d,
        amount:num(val(totalRow,3+i))
      });
    }
  }

  // Planned payments from BD Plan Detallado.
  const paymentDefs=[
    {periodCol:18,amountCol:20,name:'Pago 1'},
    {periodCol:21,amountCol:23,name:'Pago 2'},
    {periodCol:24,amountCol:26,name:'Pago 3'},
    {periodCol:27,amountCol:29,name:'Pago 4'},
    {periodCol:30,amountCol:32,name:'Pago 5'}
  ];
  const plannedPayments=[];
  bd.slice(4).forEach((r,idx)=>{
    const category=text(val(r,2)), item=text(val(r,5)), supplier=text(val(r,17));
    paymentDefs.forEach(def=>{
      const period=Math.round(num(val(r,def.periodCol)));
      const amount=num(val(r,def.amountCol));
      if(period>0 && amount>0){
        const d=new Date(cashflowStart.getFullYear(),cashflowStart.getMonth()+period-1,1);
        const currency=text(val(r,9))||'DOP';
        const rate=num(val(r,10))||1;
        plannedPayments.push({
          row:idx+5,period,date:d,monthLabel:new Intl.DateTimeFormat('es-DO',{month:'long',year:'numeric'}).format(d),
          category,item:item||category,supplier,payment:def.name,currency,rate,
          amountDop:amount,amountOriginal:rate?amount/rate:amount,amount
        });
      }
    });
  });

  return {plan,bd,cron,seg,pen,pre,totalBudget,summary,part19,approvedBudgetUsd,currentBudgetUsd,approvedBudgetDop,currentBudgetDop,projectSavingsDop,planRate,purchases,scheduled,opening,logistics,issues,considerations,preParams,cashflowMonths,plannedPayments};
}


let countdownTimer=null;
function openingDate(){
  if(projectConfig?.openingDate){
    const [y,m,d]=projectConfig.openingDate.split('-').map(Number);
    return new Date(y,m-1,d,0,0,0);
  }
  return model?.opening || null;
}
function renderHome(){
  const od=openingDate();
  $('homeProjectName').textContent=projectConfig?.displayName||projectConfig?.name||'Proyecto';
  $('homeOpeningDate').textContent=od?new Intl.DateTimeFormat('es-DO',{day:'2-digit',month:'short',year:'numeric'}).format(od):'Fecha no definida';
}
function startCountdown(){
  if(countdownTimer)clearInterval(countdownTimer);
  const target=openingDate();
  $('countdownProject').textContent=projectConfig?.name||'Proyecto';
  if(!target)return;
  $('countdownDateLabel').textContent=new Intl.DateTimeFormat('es-DO',{day:'numeric',month:'long',year:'numeric'}).format(target);
  const tick=()=>{
    const diff=target.getTime()-Date.now();
    if(diff<=0){
      $('cdDays').textContent='0';$('cdHours').textContent='0';$('cdMinutes').textContent='0';$('cdSeconds').textContent='0';
      $('countdownStatus').textContent='Fecha de apertura alcanzada';
      return;
    }
    const days=Math.floor(diff/86400000);
    const hours=Math.floor((diff%86400000)/3600000);
    const mins=Math.floor((diff%3600000)/60000);
    const secs=Math.floor((diff%60000)/1000);
    $('cdDays').textContent=days;$('cdHours').textContent=String(hours).padStart(2,'0');$('cdMinutes').textContent=String(mins).padStart(2,'0');$('cdSeconds').textContent=String(secs).padStart(2,'0');
    $('countdownStatus').textContent=days<=15?'Apertura muy próxima':days<=45?'Fase crítica hacia apertura':'Hacia la apertura';
  };
  tick();countdownTimer=setInterval(tick,1000);
}
function monthStart(d){return new Date(d.getFullYear(),d.getMonth(),1)}
function sameMonth(a,b){return a&&b&&a.getFullYear()===b.getFullYear()&&a.getMonth()===b.getMonth()}
function moneyCurrency(n,currency='DOP'){
  const cur=(currency||'DOP').toUpperCase();
  const symbol=cur==='USD'?'US$':cur==='EUR'?'€':'RD$';
  return `${symbol} ${new Intl.NumberFormat('es-DO',{maximumFractionDigits:2}).format(Number(n)||0)}`;
}
function cashflowView(){return $('cashflowCurrencyView')?.value||'DOP'}
function paymentDisplayAmount(p){
  return cashflowView()==='ORIGINAL'?p.amountOriginal:p.amountDop;
}
function paymentMonthSummary(){
  const m=new Map();
  model.plannedPayments.forEach(p=>{
    const key=`${p.date.getFullYear()}-${String(p.date.getMonth()+1).padStart(2,'0')}`;
    if(!m.has(key))m.set(key,{date:p.date,label:p.monthLabel,amountDop:0,count:0,currencies:new Map()});
    const x=m.get(key);x.amountDop+=p.amountDop;x.count++;
    x.currencies.set(p.currency,(x.currencies.get(p.currency)||0)+p.amountOriginal);
  });
  return [...m.values()].sort((a,b)=>a.date-b.date);
}
function paymentState(d){
  const now=monthStart(new Date()), dm=monthStart(d);
  if(dm<now)return 'overdue';
  if(sameMonth(dm,now))return 'current';
  return 'upcoming';
}
function paymentAlertText(d){
  const now=monthStart(new Date()), dm=monthStart(d);
  const months=(dm.getFullYear()-now.getFullYear())*12+(dm.getMonth()-now.getMonth());
  if(months<0)return 'Fecha programada ya transcurrida';
  if(months===0)return 'Pago programado este mes';
  if(months===1)return 'Pago programado el próximo mes';
  if(months<=3)return `Pago dentro de ${months} meses`;
  return 'Pago futuro programado';
}
function summaryAmount(x){
  if(cashflowView()==='DOP')return compactMoney(x.amountDop);
  return [...x.currencies.entries()].map(([c,v])=>moneyCurrency(v,c)).join(' + ');
}
function renderPaymentReminders(){
  const now=monthStart(new Date());
  const summaries=paymentMonthSummary().filter(x=>x.date>=now).slice(0,4);
  const fallback=paymentMonthSummary().slice(-4);
  const list=summaries.length?summaries:fallback;
  const markup=list.map(x=>`<article class="payment-reminder ${paymentState(x.date)}">
    <span class="label">${paymentAlertText(x.date)}</span>
    <h3>${esc(x.label)}</h3><b>${summaryAmount(x)}</b>
    <p>${x.count} pago${x.count===1?'':'s'} programado${x.count===1?'':'s'} en BD Plan Detallado.</p>
  </article>`).join('');
  $('paymentReminderList').innerHTML=markup||'<span class="muted">No hay pagos programados.</span>';
  $('overviewPaymentReminders').innerHTML=list.map(x=>`<div class="reminder-mini ${paymentState(x.date)}"><span>${paymentAlertText(x.date)}</span><b>${esc(x.label)}</b><small>${summaryAmount(x)} · ${x.count} pagos</small></div>`).join('');
}
function renderCashflow(){
  const summaries=paymentMonthSummary();
  const now=monthStart(new Date());
  const totalDop=summaries.reduce((s,x)=>s+x.amountDop,0);
  const scheduledPast=summaries.filter(x=>x.date<now);
  const future=summaries.filter(x=>x.date>=now);
  const paidDop=scheduledPast.reduce((s,x)=>s+x.amountDop,0);
  const futureDop=future.reduce((s,x)=>s+x.amountDop,0);
  const peak=[...summaries].sort((a,b)=>b.amountDop-a.amountDop)[0];
  const next=future[0];
  const horizon=new Date(now.getFullYear(),now.getMonth()+3,1);
  const next90=future.filter(x=>x.date<horizon).reduce((s,x)=>s+x.amountDop,0);

  if(cashflowView()==='DOP'){
    $('cfTotal').textContent=compactMoney(totalDop);
    $('cfPaidToDate').textContent=compactMoney(paidDop);
    $('cfPeakAmount').textContent=peak?compactMoney(peak.amountDop):'—';
    $('cfNextAmount').textContent=next?compactMoney(next.amountDop):'Sin pagos futuros';
    $('cf90Days').textContent=compactMoney(next90);
  }else{
    const allByCur=new Map(), paidByCur=new Map(), next90ByCur=new Map();
    model.plannedPayments.forEach(p=>{
      allByCur.set(p.currency,(allByCur.get(p.currency)||0)+p.amountOriginal);
      if(p.date<now)paidByCur.set(p.currency,(paidByCur.get(p.currency)||0)+p.amountOriginal);
      if(p.date>=now&&p.date<horizon)next90ByCur.set(p.currency,(next90ByCur.get(p.currency)||0)+p.amountOriginal);
    });
    const fmtMap=m=>[...m.entries()].map(([c,v])=>moneyCurrency(v,c)).join(' + ')||'—';
    $('cfTotal').textContent=fmtMap(allByCur);
    $('cfPaidToDate').textContent=fmtMap(paidByCur);
    $('cfPeakAmount').textContent=peak?summaryAmount(peak):'—';
    $('cfNextAmount').textContent=next?summaryAmount(next):'Sin pagos futuros';
    $('cf90Days').textContent=fmtMap(next90ByCur);
  }
  $('cfPeakMonth').textContent=peak?peak.label:'—';
  $('cfNextMonth').textContent=next?next.label:'—';

  // 1) Necesidad mensual: siempre consolidada en DOP para comparar monedas.
  chart('cashflowChart','bar',
    summaries.map(x=>x.label),
    [{label:'Necesidad de caja',data:summaries.map(x=>x.amountDop),borderWidth:0}],
    {money:true,legend:false}
  );

  // 2) Calendario: programado hasta hoy vs por venir.
  chart('cashflowStatusChart','doughnut',
    ['Programado hasta hoy','Por venir'],
    [{data:[paidDop,futureDop],borderWidth:0}],
    {money:true,noScales:true,legend:true}
  );

  // 3) Composición por moneda, convertida a DOP para que la proporción sea comparable.
  const byCurrency=new Map();
  model.plannedPayments.forEach(p=>{
    byCurrency.set(p.currency,(byCurrency.get(p.currency)||0)+p.amountDop);
  });
  const curEntries=[...byCurrency.entries()].sort((a,b)=>b[1]-a[1]);
  chart('cashflowCurrencyChart','doughnut',
    curEntries.map(x=>x[0]),
    [{data:curEntries.map(x=>x[1]),borderWidth:0}],
    {money:true,noScales:true,legend:true}
  );

  // 4) Acumulado.
  let cum=0;const cumulative=summaries.map(x=>cum+=x.amountDop);
  chart('cashflowCumulativeChart','line',
    summaries.map(x=>x.label),
    [{label:'Necesidad acumulada',data:cumulative,borderWidth:3,tension:.25,pointRadius:2}],
    {money:true,legend:false}
  );

  // 5) Próximos meses con mayor exigencia de caja.
  const upcoming=[...future].sort((a,b)=>a.date-b.date).slice(0,6);
  $('cashflowUpcomingMonths').innerHTML=upcoming.length?upcoming.map((x,i)=>{
    const totalFuture=futureDop||1;
    const share=x.amountDop/totalFuture*100;
    const state=i===0?'next':share>=20?'high':'normal';
    return `<div class="upcoming-month ${state}">
      <div><span>${i===0?'PRÓXIMO PAGO':'PAGO PROGRAMADO'}</span><b>${esc(x.label)}</b></div>
      <div class="upcoming-amount"><b>${compactMoney(x.amountDop)}</b><small>${pct(share)} del flujo futuro</small></div>
    </div>`;
  }).join(''):'<span class="muted">No hay pagos futuros programados.</span>';

  renderPaymentReminders();
  const pay=[...model.plannedPayments].sort((a,b)=>a.date-b.date||b.amountDop-a.amountDop);
  $('paymentTableCount').textContent=`${pay.length} pagos programados · gráficos comparables en DOP`;
  $('paymentTable').innerHTML=pay.map(x=>`<tr>
    <td><b>${esc(x.monthLabel)}</b></td><td>${esc(x.category)}</td><td>${esc(x.item)}</td><td>${esc(x.supplier)||'—'}</td>
    <td>${x.payment}</td><td><span class="currency-pill">${esc(x.currency)}</span></td>
    <td class="num money">${moneyCurrency(x.amountOriginal,x.currency)}</td>
    <td class="num money">${money(x.amountDop)}</td>
    <td><span class="status-pill ${paymentState(x.date)==='current'?'warn':paymentState(x.date)==='overdue'?'info':'good'}">${paymentAlertText(x.date)}</span></td>
  </tr>`).join('');
}
function renderOverview(){
  const p=model.purchases;
  const ordered=p.filter(x=>x.status==='Con OC');
  const pending=p.filter(x=>x.status==='Pendiente de compra');
  const purchasableValue=p.reduce((s,x)=>s+x.budget,0);
  const orderedValue=ordered.reduce((s,x)=>s+x.budget,0);
  const pendingValue=pending.reduce((s,x)=>s+x.budget,0);
  // Avance ponderado por dinero: una partida de mayor valor pesa más que una fila pequeña.
  const weightedProgress=purchasableValue?orderedValue/purchasableValue*100:0;
  const suppliers=uniq(p.map(x=>x.supplier));

  $('kpiApprovedBudget').textContent=compactMoney(model.approvedBudgetDop);
  $('kpiBudget').textContent=compactMoney(model.currentBudgetDop);
  $('kpiWeightedProgress').textContent=pct(weightedProgress);
  $('kpiOrdered').textContent=compactMoney(orderedValue);
  $('kpiPendingBuy').textContent=compactMoney(pendingValue);
  $('kpiCoverage').textContent=pct(weightedProgress);
  $('kpiSuppliers').textContent=suppliers.length;
  $('kpiOpenIssues').textContent=model.issues.length;

  $('overviewDiscount').textContent=compactMoney(model.projectSavingsDop);
  $('overviewDiscountOrders').textContent=`Aprobado ${compactMoney(model.approvedBudgetDop)} → vigente ${compactMoney(model.currentBudgetDop)}`;
  $('overviewDiscountPct').textContent=pct(model.approvedBudgetDop?model.projectSavingsDop/model.approvedBudgetDop*100:0);

  const noSupplier=p.filter(x=>!x.supplier).length;
  const noDate=p.filter(x=>!x.date).length;

  const health=[
    {name:'Presupuesto',state:model.projectSavingsDop>=0?'good':'warn',title:`${pct(model.approvedBudgetDop?model.projectSavingsDop/model.approvedBudgetDop*100:0)} optimizado`,desc:'Comparación del aprobado original contra el plan vigente.'},
    {name:'Avance financiero',state:weightedProgress>=80?'good':weightedProgress>=50?'warn':'bad',title:`${pct(weightedProgress)} ponderado`,desc:'El avance se pondera por valor monetario, no por cantidad de filas.'},
    {name:'Pendientes',state:model.issues.length<=3?'good':model.issues.length<=7?'warn':'bad',title:`${model.issues.length} abiertos`,desc:'Puntos pendientes identificados en la plantilla.'}
  ];
  $('healthStrip').innerHTML=health.map(h=>`<div class="health-item"><div class="row"><b>${h.name}</b><i class="dot ${h.state}"></i></div><strong>${h.title}</strong><p>${h.desc}</p></div>`).join('');

  const partLabels=model.part19.map(x=>`${Number(x.part).toFixed(2)} ${x.name.replace(/^\d+(?:\.\d+)?\s*/,'').trim()}`);
  chart('budgetMixChart','doughnut',partLabels,[{label:'Presupuesto vigente',data:model.part19.map(x=>x.currentDop),borderWidth:0}],{money:true,noScales:true,legend:true});

  chart('budgetEvolutionChart','bar',partLabels,[
    {label:'Aprobado original',data:model.part19.map(x=>x.approvedDop),borderWidth:0},
    {label:'Plan vigente',data:model.part19.map(x=>x.currentDop),borderWidth:0}
  ],{money:true,legend:true,extra:{indexAxis:'y'}});

  chart('savingsByPartChart','bar',partLabels,[{
    label:'Ahorro / (incremento)',data:model.part19.map(x=>x.savingsDop),borderWidth:0
  }],{money:true,legend:false,extra:{indexAxis:'y'}});

  const attention=[];
  if(pending.length)attention.push({c:'var(--red)',t:`${compactMoney(pendingValue)} pendientes de compra`,d:`Corresponden a ${pending.length} partidas todavía sin colocar.`});
  if(weightedProgress<70)attention.push({c:'var(--yellow)',t:`Avance financiero ponderado de ${pct(weightedProgress)}`,d:'Priorizar las partidas de mayor valor tiene más impacto que cerrar muchas partidas pequeñas.'});
  const negativeParts=model.part19.filter(x=>x.savingsDop<0).sort((a,b)=>a.savingsDop-b.savingsDop);
  if(negativeParts.length)attention.push({c:'var(--red)',t:`${negativeParts.length} partidas superan el aprobado`,d:`Mayor incremento: ${negativeParts[0].name} (${compactMoney(Math.abs(negativeParts[0].savingsDop))}).`});
  if(noSupplier)attention.push({c:'var(--yellow)',t:`${noSupplier} ítems sin proveedor definido`,d:'Completar proveedor mejora la trazabilidad de procurement.'});
  if(noDate)attention.push({c:'var(--yellow)',t:`${noDate} ítems sin fecha válida`,d:'Revisar fecha de orden o planificación en BD Plan Detallado.'});
  if(model.issues.length)attention.push({c:'var(--cyan)',t:`${model.issues.length} puntos pendientes`,d:'Revisar responsables y comentarios antes de la próxima reunión.'});
  $('attentionList').innerHTML=attention.slice(0,6).map(a=>`<div class="attention"><i style="background:${a.c}"></i><div><b>${a.t}</b><span>${a.d}</span></div></div>`).join('');

  const statusValue=[
    {label:'Con OC',value:orderedValue},
    {label:'Pendiente de compra',value:pendingValue},
    {label:'Sin OC definida',value:p.filter(x=>x.status==='Sin OC definida').reduce((s,x)=>s+x.budget,0)}
  ];
  chart('purchaseStatusChart','doughnut',statusValue.map(x=>x.label),[{data:statusValue.map(x=>x.value),borderWidth:0}],{money:true,noScales:true});

  const sup=groupSum(p.filter(x=>x.supplier),x=>x.supplier,x=>x.budget).sort((a,b)=>b.value-a.value).slice(0,6);
  $('topSuppliers').innerHTML=sup.map((x,i)=>`<div class="rank-row"><span class="n">${i+1}</span><b title="${esc(x.label)}">${esc(x.label)}</b><span>${compactMoney(x.value)}</span></div>`).join('')||'<span class="muted">Sin proveedores identificados.</span>';

  const quality=[
    ['Ítems sin proveedor',noSupplier,'BD Plan Detallado'],
    ['Ítems sin fecha',noDate,'BD Plan Detallado'],
    ['Pagos sin moneda',model.plannedPayments.filter(x=>!x.currency).length,'BD Plan Detallado']
  ];
  $('dataQuality').innerHTML=quality.map(x=>`<div class="quality"><div class="top"><b>${x[0]}</b><strong>${x[1]}</strong></div><small>${x[2]}</small></div>`).join('');
  startCountdown();
  renderPaymentReminders();
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
  renderHome();renderOverview();setupBudgetFilters();renderBudget();renderCashflow();renderTracking();renderPending();renderLiquidation();
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

  const goPage=page=>{
    const b=document.querySelector(`#navTabs button[data-page="${page}"]`);if(!b)return;
    document.querySelectorAll('#navTabs button').forEach(x=>x.classList.remove('active'));b.classList.add('active');
    document.querySelectorAll('.page').forEach(x=>x.classList.remove('active'));$('page-'+page).classList.add('active');
    window.scrollTo({top:0,behavior:'smooth'});
  };
  document.querySelectorAll('#navTabs button').forEach(b=>b.addEventListener('click',()=>goPage(b.dataset.page)));
  document.querySelectorAll('[data-go-page]').forEach(b=>b.addEventListener('click',()=>goPage(b.dataset.goPage)));
  ['budgetCategoryFilter','budgetStatusFilter','budgetSupplierFilter'].forEach(id=>$(id).addEventListener('change',renderBudget));
  $('cashflowCurrencyView').addEventListener('change',renderCashflow);
  $('budgetReset').addEventListener('click',()=>{$('budgetCategoryFilter').value='';$('budgetStatusFilter').value='';$('budgetSupplierFilter').value='';renderBudget()});
  ['liqAmount','liqExw','liqFob','liqContainers','liqRate'].forEach(id=>$(id).addEventListener('input',calcLiquidation));
}
init();
