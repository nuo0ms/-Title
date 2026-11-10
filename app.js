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
async function loadLinkage(){const d=await call('linkage',{level});linkRows=d.rows;renderLinkage();}
function renderLinkage(){const k=$('link-sort').value;$('link-rows').replaceChildren(...[...linkRows].sort((a,b)=>b[k]-a[k]).map(r=>{const tr=document.createElement('tr');[r.code,r.name,r.backward.toFixed(3),r.forward.toFixed(3),r.type].forEach(v=>tr.append(text('td',v)));return tr;}));}
async function loadItems(){const code=$('fd-sector').value;if(!code){$('fd-rows').replaceChildren();return;}
 const d=await call('items',{level,code}),s=sectors.find(x=>x.code===code),tot=d.production[d.production.length-1]||1;
 $('fd-caption').textContent=(s?s.code+' '+s.name:code)+' — 최종수요 항목별 유발액(백만원)';
 $('fd-rows').replaceChildren(...d.items.map((n,i)=>{const tr=document.createElement('tr');[n,fmt(d.production[i]),fmt(d.imports[i]),fmt(d.valueAdded[i]),(d.production[i]/tot*100).toFixed(1)+'%'].forEach(v=>tr.append(text('td',v)));return tr;}));}
$('link-sort').onchange=renderLinkage;$('fd-sector').onchange=loadItems;
$('fd-search').oninput=()=>{const q=$('fd-search').value.trim();options($('fd-sector'),sectors.filter(s=>lab(s).includes(q)));loadItems();};
$('load-basic').onclick=()=>useLevel($('level').value);
