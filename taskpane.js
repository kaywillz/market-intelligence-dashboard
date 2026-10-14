/* Presentation only. Never infer rejection, acceptance, or admission from graphics. */
const el = id => document.getElementById(id);
const field = (s,k) => { const m=String(s||'').match(new RegExp('(?:^|\\|)'+k+'=([^|]*)'));return m?m[1]:''; };
function paint(rows){
 const d=Object.fromEntries(rows.filter(r=>r[0]).map(r=>[String(r[0]),r[1]]));
 el('ticker').textContent=d.Ticker||'Awaiting review';
 el('boundary').textContent=d.ReviewBoundary||'—';
 el('daily').textContent=d.DailyId||'No eligible candle loaded';
 el('state').textContent=d.Status||'Not evaluated';
 const raw=String(d.Understand||'');
 const evidence=String(d.CandleFacts||'');
 const ok=raw.startsWith('ESTABLISHED|');
 el('headline').textContent=ok?'Preliminary market observations':(d.Status||'No assessment available');
 el('summary').textContent=ok?'The independent evaluator returned established observations. Confirmation and trading admission are separate.':'No established UNDERSTAND assessment for this review.';
 el('understand').textContent=raw||'No independent assessment supplied.';
 el('unresolved').textContent=ok?'Rejection, acceptance, confirmation, control and admission are not inferred by this dashboard.':raw||'—';
 const direction=field(evidence,'DIRECTION');
 el('direction').textContent=direction||'Not evaluated';
 el('closeLocation').textContent=(field(evidence,'CLOSE_THIRD')||'—').replaceAll('_',' ').toLowerCase();
 const vol=field(evidence,'VOLUME_RELATION');
 el('volText').textContent=(vol||'Not evaluated').replaceAll('_',' ').toLowerCase();
 const a=Number(field(evidence,'VOLUME')),b=Number(field(evidence,'AVERAGE_VOLUME20'));
 el('volBar').style.width=(a>0&&b>0?Math.min(100,100*a/Math.max(a,b)) :0)+'%';
 const open=Number(d.Open),high=Number(d.High),low=Number(d.Low),close=Number(d.Close);
 const valid=[open,high,low,close].every(Number.isFinite)&&high>low&&high>=Math.max(open,close)&&low<=Math.min(open,close)&&d.Open!==''&&d.High!=='';
 const Y=x=>15+(high-x)/(high-low)*160;
 el('wick').style.display=valid?'':'none';el('body').style.display=valid?'':'none';
 if(valid){el('wick').setAttribute('y1',Y(high));el('wick').setAttribute('y2',Y(low));
 const top=Math.min(Y(open),Y(close)),ht=Math.max(3,Math.abs(Y(open)-Y(close)));
 el('body').setAttribute('y',top);el('body').setAttribute('height',ht);
 const color=close>open?'#368a69':close<open?'#b56a63':'#8b9b92';
 el('body').setAttribute('fill',color);el('wick').setAttribute('stroke',color);}
}
async function refresh(){
 try {await Excel.run(async ctx=>{
  const sh=ctx.workbook.worksheets.getItem('MI_Presentation');
  const rg=sh.getRange('A1:B14');rg.load('values');await ctx.sync();
  const rows=rg.values;
  const d=Object.fromEntries(rows.map(r=>[String(r[0]),r[1]]));
  if(String(d.Commit)!=='READY'){el('state').textContent='Snapshot not ready';return;}
  paint(rows);
 });}catch(e){el('state').textContent='Connection unavailable';el('unresolved').textContent=String(e.message||e);}
}
Office.onReady(()=>{el('refresh').addEventListener('click',refresh);refresh();});
