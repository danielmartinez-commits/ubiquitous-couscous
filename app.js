const App = (() => {
  const state = { rows: [], brokers: [], clients: [], source: '', loadedAt: null, chart: null };
  const money = new Intl.NumberFormat('en-US',{style:'currency',currency:'USD',minimumFractionDigits:2});
  const number = new Intl.NumberFormat('en-US');
  const $ = id => document.getElementById(id);
  const norm = v => String(v ?? '').trim().replace(/\s+/g,' ').toLowerCase();
  const title = v => String(v ?? '').replace(/\s+/g,' ').trim().replace(/\b\w/g,c=>c.toUpperCase());
  const cents = v => Math.round(Number(v || 0) * 100);
  const fmt = v => money.format(Number(v || 0));
  const esc = v => String(v ?? '').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
  const toast = msg => { const t=$('toast'); t.textContent=msg; t.classList.add('show'); setTimeout(()=>t.classList.remove('show'),2600); };

  function parseDate(v){
    if(v instanceof Date && !isNaN(v)) return v;
    if(typeof v === 'number'){ const d = XLSX.SSF.parse_date_code(v); return d ? new Date(d.y,d.m-1,d.d) : null; }
    const s=String(v||'').trim(); if(!s) return null;
    let m=s.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})$/); if(m) return new Date(+m[3],+m[1]-1,+m[2]);
    const d=new Date(s); return isNaN(d)?null:d;
  }
  function dateLabel(d){ return d ? new Intl.DateTimeFormat('en-US',{month:'2-digit',day:'2-digit',year:'numeric'}).format(d) : '—'; }
  function ageDays(d){ return d ? Math.max(0,Math.floor((Date.now()-d.getTime())/86400000)) : null; }

  function findHeader(matrix){
    const wanted=['transaction date','transaction type','num','name','cliente','broker','amount','balance'];
    for(let i=0;i<Math.min(matrix.length,30);i++){
      const cells=(matrix[i]||[]).map(x=>norm(x));
      const hits=wanted.filter(x=>cells.includes(x)).length;
      if(hits>=5) return {row:i, cells};
    }
    throw new Error('Could not find the expected transaction headers.');
  }

  function parseWorkbook(buffer, source='Uploaded report'){
    const wb=XLSX.read(buffer,{type:'array',cellDates:true});
    const ws=wb.Sheets[wb.SheetNames[0]];
    const matrix=XLSX.utils.sheet_to_json(ws,{header:1,raw:true,defval:null});
    const header=findHeader(matrix);
    const map={}; header.cells.forEach((h,i)=>{ if(h) map[h]=i; });
    const rows=[];
    for(let i=header.row+1;i<matrix.length;i++){
      const r=matrix[i]||[];
      const date=parseDate(r[map['transaction date']]);
      const type=r[map['transaction type']];
      const num=r[map['num']];
      const client=r[map['cliente']];
      const broker=r[map['broker']];
      const amount=Number(r[map['amount']]);
      const balance=Number(r[map['balance']]);
      if(!date || !type || !broker || !Number.isFinite(amount)) continue;
      rows.push({id:`${i}-${num}-${amount}`,date,type:String(type).trim(),num:String(num??'').trim(),client:String(client??'').trim(),broker:String(broker).trim(),brokerKey:norm(broker),amount,balance:Number.isFinite(balance)?balance:null,age:ageDays(date)});
    }
    if(!rows.length) throw new Error('No usable transaction rows were found in the report.');
    return {rows,source,loadedAt:new Date()};
  }

  async function saveLocal(){
    try{ localStorage.setItem('wc_income_rows',JSON.stringify({rows:state.rows.map(r=>({...r,date:r.date.toISOString()})),source:state.source,loadedAt:state.loadedAt.toISOString()})); }catch(e){ console.warn('Local save failed',e); }
  }
  function restoreLocal(){
    try{ const raw=localStorage.getItem('wc_income_rows'); if(!raw)return false; const x=JSON.parse(raw); state.rows=x.rows.map(r=>({...r,date:new Date(r.date)})); state.source=x.source; state.loadedAt=new Date(x.loadedAt); return true; }catch(e){return false;}
  }

  function applyDataset(parsed){
    state.rows=parsed.rows; state.source=parsed.source; state.loadedAt=parsed.loadedAt;
    state.brokers=[...new Map(state.rows.map(r=>[r.brokerKey,title(r.broker)])).entries()].map(([key,name])=>({key,name})).sort((a,b)=>a.name.localeCompare(b.name));
    state.clients=[...new Set(state.rows.map(r=>r.client).filter(Boolean).map(title))].sort();
    $('sidebarStatus').textContent=`${number.format(state.rows.length)} rows ready`;
    $('reportDate').textContent=`Updated ${new Intl.DateTimeFormat('en-US',{month:'short',day:'2-digit',year:'numeric',hour:'numeric',minute:'2-digit'}).format(state.loadedAt)}`;
    const list=$('brokerList'); list.innerHTML=state.brokers.map(b=>`<option value="${esc(b.name)}"></option>`).join('');
    renderDashboard(); renderBrokers(); renderData();
  }

  async function loadFile(file){
    if(!file)return; try{ const parsed=parseWorkbook(await file.arrayBuffer(),file.name); applyDataset(parsed); await saveLocal(); toast(`Loaded ${number.format(parsed.rows.length)} transactions.`); }catch(e){ console.error(e); toast(e.message); } }
  async function tryAutoLoad(){
    if(restoreLocal()){ applyDataset({rows:state.rows,source:state.source,loadedAt:state.loadedAt}); return; }
    try{ const res=await fetch('data/AR_Report.xlsx',{cache:'no-store'}); if(!res.ok)throw new Error(); const parsed=parseWorkbook(await res.arrayBuffer(),'data/AR_Report.xlsx'); applyDataset(parsed); await saveLocal(); }
    catch(e){ $('sidebarStatus').textContent='Waiting for report'; $('dataDescription').textContent='No report loaded. Use Data → Choose Excel report.'; }
  }

  function renderDashboard(){
    const total=state.rows.reduce((s,r)=>s+r.amount,0);
    $('kpiMovements').textContent=number.format(state.rows.length); $('kpiAmount').textContent=fmt(total); $('kpiBrokers').textContent=number.format(state.brokers.length); $('kpiClients').textContent=number.format(state.clients.length);
    const groups=aggregateBrokers().slice(0,10);
    const canvas=$('brokerChart'); if(state.chart)state.chart.destroy();
    state.chart=new Chart(canvas,{type:'bar',data:{labels:groups.map(x=>shorten(x.name,23)),datasets:[{data:groups.map(x=>x.amount),backgroundColor:'#6d4aff',borderRadius:5,barThickness:18}]},options:{indexAxis:'y',plugins:{legend:{display:false},tooltip:{callbacks:{label:c=>fmt(c.raw)}}},scales:{x:{grid:{color:'#eef0f4'},ticks:{font:{size:9},callback:v=>fmt(v)}},y:{grid:{display:false},ticks:{font:{size:9}}}}}});
    $('topBrokersTable').innerHTML=`<table class="mini-table"><thead><tr><th>Broker</th><th>Items</th><th>Amount</th></tr></thead><tbody>${groups.map(x=>`<tr><td>${esc(x.name)}</td><td>${number.format(x.count)}</td><td>${fmt(x.amount)}</td></tr>`).join('')}</tbody></table>`;
  }
  function aggregateBrokers(){ const m=new Map(); for(const r of state.rows){const x=m.get(r.brokerKey)||{key:r.brokerKey,name:title(r.broker),amount:0,count:0,clients:new Set(),oldest:null};x.amount+=r.amount;x.count++;if(r.client)x.clients.add(norm(r.client));if(!x.oldest||r.date<x.oldest)x.oldest=r.date;m.set(r.brokerKey,x);}return [...m.values()].sort((a,b)=>b.amount-a.amount); }
  function shorten(s,n){return s.length>n?s.slice(0,n-1)+'…':s;}

  function renderBrokers(filter=''){ const q=norm(filter); const all=aggregateBrokers(); const list=all.filter(x=>!q||norm(x.name).includes(q)); $('brokerCount').textContent=`${number.format(list.length)} brokers`; $('brokerCards').innerHTML=list.map(x=>`<article class="broker-card" data-broker="${esc(x.key)}"><h3 title="${esc(x.name)}">${esc(x.name)}</h3><p>${number.format(x.clients.size)} clients · ${number.format(x.count)} movements</p><div class="amount">${fmt(x.amount)}</div><div class="broker-meta"><span>Oldest ${dateLabel(x.oldest)}</span><span>View →</span></div></article>`).join('')||'<div class="empty-state"><h2>No brokers found</h2><p>Try another search.</p></div>'; document.querySelectorAll('.broker-card').forEach(c=>c.addEventListener('click',()=>openBroker(c.dataset.broker))); }

  function identify(){
    const broker=norm($('brokerInput').value); const amount=Number(String($('amountInput').value).replace(/[$,\s]/g,'')); const date=$('dateInput').value?new Date($('dateInput').value+'T00:00:00'):null;
    if(!broker||!Number.isFinite(amount)||amount<=0){toast('Enter a broker and a valid payment amount.');return;}
    const candidates=state.rows.filter(r=>r.brokerKey===broker || norm(r.broker).includes(broker)).filter(r=>r.amount>0);
    if(!candidates.length){renderNoMatch();return;}
    const exact=candidates.filter(r=>cents(r.amount)===cents(amount)).sort((a,b)=>(date?Math.abs(a.date-date)-Math.abs(b.date-date):0));
    const combo=findCombination(candidates,amount,date);
    renderIdentification(broker,amount,date,exact,combo,candidates);
  }
  function findCombination(rows,target,date){
    const targetC=cents(target); let candidates=[...rows].filter(r=>cents(r.amount)<=targetC).sort((a,b)=>{const da=date?Math.abs(a.date-date):0,db=date?Math.abs(b.date-date):0;return (a.amount-b.amount)*-1 + (da-db)/86400000*0.0001;}).slice(0,120);
    const dp=new Map([[0,[]]]); let best=null;
    for(const r of candidates){ const rc=cents(r.amount); const entries=[...dp.entries()]; for(const [sum,combo] of entries){const ns=sum+rc;if(ns>targetC)continue;if(!dp.has(ns)){const nc=[...combo,r];dp.set(ns,nc);if(ns===targetC){best=nc;break;}}} if(best)break; if(dp.size>100000)break;}
    return best;
  }
  function renderNoMatch(){ $('identifyEmpty').classList.add('hidden'); const box=$('identifyResults'); box.classList.remove('hidden'); box.innerHTML='<div class="result-head"><div><h2>No candidate found</h2><p>No movement for that broker could be matched to the requested amount.</p></div><span class="badge red">NO MATCH</span></div>'; }
  function renderIdentification(broker,amount,date,exact,combo,candidates){
    $('identifyEmpty').classList.add('hidden'); const box=$('identifyResults'); box.classList.remove('hidden');
    let html=`<div class="result-head"><div><h2>Identification result</h2><p>${number.format(candidates.length)} movements found for <strong>${esc(title(broker))}</strong>.</p></div></div>`;
    if(exact.length){ html+=matchCard('Exact amount match',exact.slice(0,10),amount,date,'green','EXACT MATCH'); }
    if(combo && (!exact.length || combo.length>1)){ html+=matchCard('Possible batch match',combo,amount,date,'green','EXACT COMBINATION'); }
    if(!exact.length && !combo){ const near=[...candidates].sort((a,b)=>Math.abs(a.amount-amount)-Math.abs(b.amount-amount)).slice(0,10); html+=matchCard('Closest open movements',near,amount,date,'yellow','REVIEW'); }
    box.innerHTML=html;
  }
  function matchCard(titleText,rows,payment,date,badgeClass,badgeText){const sum=rows.reduce((s,r)=>s+r.amount,0);const diff=payment-sum;return `<article class="match-card"><div class="match-top"><div class="match-title"><h3>${esc(titleText)}</h3><p>${rows.length} movement${rows.length===1?'':'s'} · ${date?`payment date ${dateLabel(date)}`:'no payment date provided'}</p></div><span class="badge ${badgeClass}">${badgeText}</span></div><div class="match-summary"><div class="summary-item"><small>Selected</small><strong>${fmt(sum)}</strong></div><div class="summary-item"><small>Payment</small><strong>${fmt(payment)}</strong></div><div class="summary-item"><small>Difference</small><strong>${fmt(diff)}</strong></div></div><table class="match-table"><thead><tr><th>Date</th><th>Num</th><th>Client</th><th>Broker</th><th>Amount</th></tr></thead><tbody>${rows.map(r=>`<tr><td>${dateLabel(r.date)}</td><td>${esc(r.num)}</td><td>${esc(title(r.client))}</td><td>${esc(title(r.broker))}</td><td class="amount">${fmt(r.amount)}</td></tr>`).join('')}</tbody></table></article>`;}

  function openBroker(key){ const rows=state.rows.filter(r=>r.brokerKey===key).sort((a,b)=>b.date-a.date); if(!rows.length)return; const name=title(rows[0].broker); const total=rows.reduce((s,r)=>s+r.amount,0); $('modalContent').innerHTML=`<div class="detail-title"><h2>${esc(name)}</h2><p>Movement history available in the current AR report.</p></div><div class="detail-stats"><div class="detail-stat"><small>Total amount</small><strong>${fmt(total)}</strong></div><div class="detail-stat"><small>Movements</small><strong>${number.format(rows.length)}</strong></div><div class="detail-stat"><small>Clients</small><strong>${number.format(new Set(rows.map(r=>norm(r.client))).size)}</strong></div></div><table class="detail-table"><thead><tr><th>Date</th><th>Num</th><th>Client</th><th>Type</th><th>Amount</th><th>Balance</th></tr></thead><tbody>${rows.map(r=>`<tr><td>${dateLabel(r.date)}</td><td>${esc(r.num)}</td><td>${esc(title(r.client))}</td><td>${esc(r.type)}</td><td>${fmt(r.amount)}</td><td>${r.balance==null?'—':fmt(r.balance)}</td></tr>`).join('')}</tbody></table>`; $('modal').classList.remove('hidden'); }

  function renderData(){ const total=state.rows.reduce((s,r)=>s+r.amount,0); $('dataDescription').textContent=`${state.source} · loaded ${new Intl.DateTimeFormat('en-US',{dateStyle:'medium',timeStyle:'short'}).format(state.loadedAt)}`; $('dataStats').innerHTML=[['Rows',number.format(state.rows.length)],['Brokers',number.format(state.brokers.length)],['Clients',number.format(state.clients.length)],['Amount',fmt(total)],['First date',dateLabel(Math.min(...state.rows.map(r=>r.date.getTime()))===Infinity?null:new Date(Math.min(...state.rows.map(r=>r.date.getTime()))))],['Last date',dateLabel(new Date(Math.max(...state.rows.map(r=>r.date.getTime()))))]].map(([a,b])=>`<div class="data-stat"><span>${a}</span><strong>${b}</strong></div>`).join(''); }

  function switchView(view){document.querySelectorAll('.nav-item').forEach(x=>x.classList.toggle('active',x.dataset.view===view));document.querySelectorAll('.view').forEach(x=>x.classList.remove('active-view'));$(`view-${view}`).classList.add('active-view');}
  function globalSearch(q){q=norm(q);if(!q)return;const broker=state.brokers.find(b=>norm(b.name).includes(q));if(broker){switchView('brokers');$('brokerFilter').value=broker.name;renderBrokers(broker.name);return;}const client=state.rows.find(r=>norm(r.client).includes(q));if(client){$('brokerInput').value=client.broker;$('amountInput').value=client.amount;switchView('identify');identify();return;}const num=state.rows.find(r=>norm(r.num)===q);if(num){$('brokerInput').value=num.broker;$('amountInput').value=num.amount;switchView('identify');identify();return;}const amount=Number(q.replace(/[$,\s]/g,''));if(Number.isFinite(amount)){switchView('identify');$('amountInput').value=amount;toast('Enter the broker to narrow the identification.');}}

  function bind(){
    document.querySelectorAll('.nav-item').forEach(b=>b.addEventListener('click',()=>switchView(b.dataset.view)));
    $('identifyBtn').addEventListener('click',identify); $('amountInput').addEventListener('keydown',e=>{if(e.key==='Enter')identify();});
    $('brokerFilter').addEventListener('input',e=>renderBrokers(e.target.value)); $('globalSearch').addEventListener('keydown',e=>{if(e.key==='Enter')globalSearch(e.target.value)});
    $('fileInput').addEventListener('change',e=>loadFile(e.target.files[0])); $('refreshBtn').addEventListener('click',tryAutoLoad);
    $('modalClose').addEventListener('click',()=>$('modal').classList.add('hidden')); document.querySelector('.modal-backdrop').addEventListener('click',()=>$('modal').classList.add('hidden'));
    const dz=$('dropZone'); ['dragenter','dragover'].forEach(ev=>dz.addEventListener(ev,e=>{e.preventDefault();dz.style.borderColor='#6d4aff'})); ['dragleave','drop'].forEach(ev=>dz.addEventListener(ev,e=>{e.preventDefault();dz.style.borderColor=''})); dz.addEventListener('drop',e=>loadFile(e.dataTransfer.files[0]));
  }
  bind(); tryAutoLoad();
  return {loadFile};
})();
