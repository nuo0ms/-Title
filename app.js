const $=id=>document.getElementById(id);
const fmt=x=>Number(x).toLocaleString('ko-KR',{maximumFractionDigits:3});
const FILES={basic:'./io-2023.xlsx',small:'./io-2023-small.xlsx',mid:'./io-2023-mid.xlsx',large:'./io-2023-large.xlsx'};
const LABEL={basic:'기본부문(380)',small:'소분류(165)',mid:'중분류(83)',large:'대분류(33)'};
let level='basic', loaded={}, tags={}, booted=false, worker, nextId=0, requests=new Map(), sectors=[], rowChoices=[], cells=[], basket=[], result=null;
function text(tag,value){const e=document.createElement(tag);e.textContent=value;return e;}
function call(type,payload){return new Promise((resolve,reject)=>{const id=++nextId;requests.set(id,{resolve,reject});worker.postMessage({id,type,payload});});}
function startWorker(){worker=new Worker('./worker.js');worker.onmessage=({data:d})=>{if(d.type==='status'){$('status').textContent=d.text;return;}const p=requests.get(d.id);if(!p)return;requests.delete(d.id);d.ok?p.resolve(d.result):p.reject(Error(d.error));};worker.onerror=e=>{for(const p of requests.values())p.reject(Error(e.message||'Worker 실행 오류'));requests.clear();};}
const lab=x=>x.code+' '+x.name+(x.tag?' ('+x.tag+')':'');
function options(el,list){el.replaceChildren(...list.map(x=>{const o=document.createElement('option');o.value=x.code;o.textContent=lab(x);return o;}));}
function invalidate(){result=null;$('results').hidden=true;}
const IDS=['link-sort','fd-search','fd-sector','sheet','row','sector','add','run','row-search','col-search','sector-search','amount','level'];
function setEnabled(on){for(const id of IDS)$(id).disabled=!on;}
async function useLevel(lv){
 invalidate();basket=[];renderBasket();setEnabled(false);$('load-basic').disabled=true;
 try{
  if(!booted){$('status').textContent='Python 실행 환경 준비 중… (처음 한 번, 30초 안팎)';startWorker();await call('boot');booted=true;
   try{const r=await fetch('./classification.xlsx');if(r.ok)tags=await call('classify',{buffer:await r.arrayBuffer()});}catch(e){}}
  if(!loaded[lv]){
   $('status').textContent=LABEL[lv]+' 엑셀 내려받는 중…';
   const r=await fetch(FILES[lv]);if(!r.ok)throw Error(FILES[lv]+' 다운로드 실패 (HTTP '+r.status+')');
   $('status').textContent=LABEL[lv]+' 읽기와 계수 검증 중…';
   loaded[lv]=await call('load',{level:lv,buffer:await r.arrayBuffer()});
   if(lv==='basic')for(const s of loaded[lv].sectors)s.tag=tags[s.code]||'';
  }
  level=lv;const d=loaded[lv];sectors=d.sectors;options($('sector'),sectors);
  for(const id of ['sector-search','row-search','col-search'])$(id).value='';
  $('sheet').replaceChildren(...d.sheets.map(s=>{const o=text('option',s);o.value=s;return o;}));$('row').replaceChildren();
  $('validation').textContent=LABEL[lv]+'\n'+JSON.stringify(d.residuals,null,2)+'\n최대 절대 잔차: 허용값 1e-8 미만';
  $('level').value=lv;options($('fd-sector'),sectors);$('fd-search').value='';setEnabled(true);await explore();await loadLinkage();await loadItems();
  $('status').textContent='준비 완료 · '+LABEL[lv]+' · '+sectors.length+'개 부문 · '+d.sheets.length+'개 시트 · 계수 검증 통과';
  $('load-basic').textContent='자료 불러오기 완료';
 }catch(e){$('status').textContent='준비 실패: '+e.message+' — 새로고침 후 재시도하세요.';$('level').disabled=false;$('load-basic').disabled=false;if(loaded[level]){setEnabled(true);}}
}
let exploreSeq=0;
async function explore(){const seq=++exploreSeq;try{const d=await call('explore',{level,sheet:$('sheet').value,row:$('row').value});if(seq!==exploreSeq)return;rowChoices=d.rows;filterRows(d.selected);cells=d.cells;$('unit').textContent=d.unit;$('explore-caption').textContent=$('sheet').value+' / 행: '+d.selected+' '+d.name;renderCells();}catch(e){$('status').textContent=e.message;}}
function filterRows(selected){const q=$('row-search').value.trim();options($('row'),rowChoices.filter(r=>(r.code+' '+r.name).includes(q)));if([...$('row').options].some(o=>o.value===selected))$('row').value=selected;}
function renderCells(){const q=$('col-search').value.trim();$('cells').replaceChildren(...cells.filter(c=>(c.code+' '+c.name).includes(q)).map(c=>{const tr=document.createElement('tr');for(const v of [c.code,c.name,c.value===null?'빈 셀':typeof c.value==='number'?fmt(c.value):c.value])tr.append(text('td',v));return tr;}));}
$('sheet').onchange=()=>{$('row-search').value='';$('row').replaceChildren();explore();};$('row').onchange=explore;
$('row-search').oninput=()=>{filterRows();if($('row').value)explore();else{$('cells').replaceChildren();$('explore-caption').textContent='일치하는 행이 없습니다.';}};
$('col-search').oninput=renderCells;
$('sector-search').oninput=()=>{const q=$('sector-search').value.trim();options($('sector'),sectors.filter(s=>lab(s).includes(q)));};
function renderBasket(){$('basket').replaceChildren(...basket.map((s,i)=>{const li=text('li',s.code+' '+s.name+' / '+fmt(s.amount)+'백만원 ');const b=text('button','삭제');b.type='button';b.onclick=()=>{basket.splice(i,1);invalidate();renderBasket();};li.append(b);return li;}));$('sim-status').textContent=basket.length?basket.length+'개 입력 준비. 계산 버튼을 누르세요.':'시나리오를 추가하세요.';}
$('scenario').onsubmit=e=>{e.preventDefault();const amount=Number($('amount').value),s=sectors.find(s=>s.code===$('sector').value);if(!s||!Number.isFinite(amount)||amount<=0){$('sim-status').textContent='산업과 양의 증가액을 확인하세요.';return;}basket.push({...s,amount});invalidate();renderBasket();};
$('level').onchange=()=>{if(booted)useLevel($('level').value);};
$('clear').onclick=()=>{basket=[];invalidate();renderBasket();};
$('run').onclick=async()=>{if(!basket.length){$('sim-status').textContent='시나리오를 먼저 추가하세요.';return;}const snapshot=basket.map(s=>({...s}));$('run').disabled=true;$('sim-status').textContent='계산 중…';try{const r=await call('simulate',{level,shocks:snapshot});if(JSON.stringify(snapshot)!==JSON.stringify(basket)){$('sim-status').textContent='입력이 변경되었습니다. 다시 계산하세요.';return;}result=r;renderResult();$('sim-status').textContent='계산 완료. 아래 결과는 현재 시나리오에 대응합니다.';}catch(e){invalidate();$('sim-status').textContent='계산 실패: '+e.message;}finally{$('run').disabled=false;}};
function renderResult(){$('results').hidden=false;$('cards').replaceChildren(...[['수요 증가액',result.demand],['생산유발액',result.production],['수입유발액',result.imports],['부가가치유발액',result.valueAdded]].map(([title,v])=>{const d=text('div',title);d.className='card';d.append(text('strong',fmt(v)),text('span','백만원'));return d;}));const max=Math.max(...result.rows.map(r=>r.production),1);$('bars').replaceChildren(...result.rows.slice(0,15).map(r=>{const d=text('div',r.name+' · '+fmt(r.production)+'백만원');d.className='bar';const t=document.createElement('div');t.className='track';const f=document.createElement('div');f.className='fill';f.style.width=Math.max(0,r.production/max*100)+'%';t.append(f);d.append(t);return d;}));$('result-rows').replaceChildren(...result.rows.map(r=>{const tr=document.createElement('tr');[r.code,r.name,fmt(r.production),fmt(r.imports),fmt(r.valueAdded)].forEach(v=>tr.append(text('td',v)));return tr;}));}
function csvCell(v){let s=String(v);if(/^[=+@-]/.test(s)&&typeof v!=='number')s="'"+s;return '"'+s.replaceAll('"','""')+'"';}
$('download').onclick=()=>{if(!result)return;const rows=[['산업코드','산업명','생산유발액_백만원','수입유발액_백만원','부가가치유발액_백만원'],...result.rows.map(r=>[r.code,r.name,r.production,r.imports,r.valueAdded])];const url=URL.createObjectURL(new Blob(['\ufeff'+rows.map(r=>r.map(csvCell).join(',')).join('\r\n')],{type:'text/csv;charset=utf-8'}));const a=document.createElement('a');a.href=url;a.download='io-2023-'+level+'-results.csv';a.click();setTimeout(()=>URL.revokeObjectURL(url),1000);};


let linkRows=[];
async function loadLinkage(){const d=await call('linkage',{level});linkRows=d.rows;renderLinkage();drawLinkChart();}
const NS='http://www.w3.org/2000/svg';
function svg(n,at,t){const e=document.createElementNS(NS,n);for(const k in at)e.setAttribute(k,at[k]);if(t!==undefined)e.textContent=t;return e;}
const KIND=[['전·후방','both','전·후방 연쇄 산업'],['후방','back','후방연쇄 중심'],['전방','fwd','전방연쇄 중심'],['독립','ind','독립적 산업']];
const kindOf=t=>KIND.find(k=>t.startsWith(k[0]))[1];
function drawLinkChart(){
 const W=640,H=470,m={l:56,r:18,t:18,b:48},rows=linkRows,xs=rows.map(r=>r.backward),ys=rows.map(r=>r.forward);
 const x0=Math.min(...xs,.9)*.95,x1=Math.max(...xs,1.1)*1.05,y0=Math.min(...ys,.9)*.95,y1=Math.max(...ys,1.1)*1.05,lg=Math.log;
 const sx=v=>m.l+(lg(v)-lg(x0))/(lg(x1)-lg(x0))*(W-m.l-m.r),sy=v=>H-m.b-(lg(v)-lg(y0))/(lg(y1)-lg(y0))*(H-m.t-m.b);
 const s=svg('svg',{viewBox:`0 0 ${W} ${H}`,role:'img','aria-label':'산업별 후방·전방연쇄효과 산점도'});
 const X=sx(1),Y=sy(1),R=W-m.r,B=H-m.b;
 s.append(svg('rect',{x:X,y:m.t,width:R-X,height:Y-m.t,class:'q-both'}),svg('line',{x1:X,x2:X,y1:m.t,y2:B,class:'ref'}),svg('line',{x1:m.l,x2:R,y1:Y,y2:Y,class:'ref'}),svg('rect',{x:m.l,y:m.t,width:W-m.l-m.r,height:H-m.t-m.b,class:'frame'}));
 for(const v of [.5,.75,1,1.5,2,3,4,6,8,12,16]){if(v>x0&&v<x1){s.append(svg('text',{x:sx(v),y:B+16,'text-anchor':'middle'},v));}if(v>y0&&v<y1){s.append(svg('text',{x:m.l-8,y:sy(v)+4,'text-anchor':'end'},v));}}
 s.append(svg('text',{x:(m.l+R)/2,y:H-8,'text-anchor':'middle',class:'ax'},'후방연쇄효과(영향력계수) → 클수록 다른 산업을 많이 끌어올림'),svg('text',{x:14,y:(m.t+B)/2,'text-anchor':'middle',class:'ax',transform:`rotate(-90 14 ${(m.t+B)/2})`},'전방연쇄효과(감응도계수) →'));
 s.append(svg('text',{x:R-6,y:m.t+16,'text-anchor':'end',class:'qt'},'전·후방 연쇄'),svg('text',{x:R-6,y:B-8,'text-anchor':'end',class:'qt'},'후방연쇄 중심'),svg('text',{x:m.l+6,y:m.t+16,class:'qt'},'전방연쇄 중심'),svg('text',{x:m.l+6,y:B-8,class:'qt'},'독립적'));
 const r0=rows.length>200?3:4.5;
 for(const r of rows){const c=svg('circle',{cx:sx(r.backward),cy:sy(r.forward),r:r0,class:'dot-'+kindOf(r.type)});c.append(svg('title',{},r.code+' '+r.name+'\n후방 '+r.backward.toFixed(3)+' · 전방 '+r.forward.toFixed(3)+'\n'+r.type));s.append(c);}
 const placed=[];
 for(const r of rows.filter(r=>kindOf(r.type)==='both').sort((a,b)=>b.backward+b.forward-a.backward-a.forward).slice(0,5)){let ty=sy(r.forward)-6;while(placed.some(p=>Math.abs(p-ty)<13))ty+=13;placed.push(ty);s.append(svg('text',{x:sx(r.backward)+7,y:ty,class:'pt'},r.name));}
 $('link-chart').replaceChildren(s);
 $('link-legend').replaceChildren(...KIND.map(k=>{const sp=document.createElement('span');const d=document.createElement('i');d.className='sw dot-'+k[1];sp.append(d,k[2]+' '+rows.filter(r=>kindOf(r.type)===k[1]).length+'개');return sp;}));
}

function renderLinkage(){const k=$('link-sort').value;$('link-rows').replaceChildren(...[...linkRows].sort((a,b)=>b[k]-a[k]).map(r=>{const tr=document.createElement('tr');[r.code,r.name,r.backward.toFixed(3),r.forward.toFixed(3),r.type].forEach(v=>tr.append(text('td',v)));return tr;}));}
async function loadItems(){const code=$('fd-sector').value;if(!code){$('fd-rows').replaceChildren();return;}
 const d=await call('items',{level,code}),s=sectors.find(x=>x.code===code),tot=d.production[d.production.length-1]||1;
 $('fd-caption').textContent=(s?s.code+' '+s.name:code)+' — 최종수요 항목별 유발액(백만원)';
 const mx=Math.max(...d.production.slice(0,-1))||1;
 $('fd-chart').replaceChildren(...d.items.slice(0,-1).map((n,i)=>{const b=document.createElement('div');b.className='bar';b.append(text('span',n+' · '+fmt(Math.round(d.production[i]))+'백만원 ('+(d.production[i]/tot*100).toFixed(1)+'%)'));const t=document.createElement('div');t.className='track';const f=document.createElement('div');f.className='fill';f.style.width=(d.production[i]/mx*100)+'%';t.append(f);b.append(t);return b;}));
 $('fd-rows').replaceChildren(...d.items.map((n,i)=>{const tr=document.createElement('tr');[n,fmt(d.production[i]),fmt(d.imports[i]),fmt(d.valueAdded[i]),(d.production[i]/tot*100).toFixed(1)+'%'].forEach(v=>tr.append(text('td',v)));return tr;}));}
$('link-sort').onchange=renderLinkage;$('fd-sector').onchange=loadItems;
$('fd-search').oninput=()=>{const q=$('fd-search').value.trim();options($('fd-sector'),sectors.filter(s=>lab(s).includes(q)));loadItems();};
$('load-basic').onclick=()=>useLevel($('level').value);
