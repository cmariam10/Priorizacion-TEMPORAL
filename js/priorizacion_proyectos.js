const all=(window.HUB_PROJECTS||[]).filter(p=>['perfil','prefactibilidad','factibilidad'].includes(String(p.etapa||'').toLowerCase()));
const storeKey='hub_mopt_project_priority_v1';
let saved={};try{saved=JSON.parse(localStorage.getItem(storeKey)||'{}')}catch(e){}
let selected=null;
const $=id=>document.getElementById(id);
const clean=v=>String(v??'').toLowerCase();
function rec(p){return saved[p.id]||{de:'',da:'',ds:'',mi:'',capex:'',excluded:false,reason:''}}
function num(v){if(v===''||v===null||v===undefined)return null;const n=Number(v);return Number.isFinite(n)?Math.max(0,Math.min(100,n)):null}
function completeness(r){const vals=[r.de,r.da,r.ds,r.mi,r.capex];return Math.round(vals.filter(v=>v!==''&&v!==null&&v!==undefined&&Number.isFinite(Number(v))&&Number(v)>=0).length/5*100)}
function hasCore(r){return [r.de,r.da,r.ds,r.mi].every(v=>num(v)!==null)&&Number(r.capex)>0}
function evaluated(r){return hasCore(r)&&!r.excluded}
function score(r){if(!evaluated(r))return null;return +(([r.de,r.da,r.ds,r.mi].map(Number).reduce((a,b)=>a+b,0))/4).toFixed(1)}
function cls(s){if(s==null)return'Pendiente';if(s>=85)return'Muy Alta';if(s>=70)return'Alta';if(s>=55)return'Media';if(s>=40)return'Baja';return'Muy baja'}
function state(p){const r=rec(p);if(r.excluded)return'excluido';return evaluated(r)?'evaluado':'pendiente'}
function filtered(){const q=clean($('q').value),st=$('stage').value,ins=$('inst').value,es=$('evalState').value;return all.filter(p=>(!q||clean(Object.values(p).join(' ')).includes(q))&&(!st||p.etapa===st)&&(!ins||p.inst===ins)&&(!es||state(p)===es))}
function init(){
  [...new Set(all.map(p=>p.etapa).filter(Boolean))].sort().forEach(v=>$('stage').add(new Option(v,v)));
  [...new Set(all.map(p=>p.inst).filter(Boolean))].sort().forEach(v=>$('inst').add(new Option(v,v)));
  ['q','stage','inst','evalState','yCut'].forEach(id=>$(id).addEventListener('input',render));
  $('reset').onclick=()=>{['q','stage','inst','evalState'].forEach(id=>$(id).value='');$('yCut').value=70;render()};
  render();
}
function render(){
  const list=filtered();
  const evals=list.map(p=>({p,r:rec(p),s:score(rec(p))})).filter(x=>x.s!=null);
  const complete=list.filter(p=>hasCore(rec(p))).length;
  $('kTotal').textContent=list.length;
  $('kEval').textContent=evals.length;
  $('kEvalSub').textContent=(list.length?Math.round(evals.length/list.length*100):0)+'% del portafolio filtrado';
  $('kAvg').textContent=evals.length?(evals.reduce((a,x)=>a+x.s,0)/evals.length).toFixed(1):'—';
  $('kHigh').textContent=evals.filter(x=>x.s>=70).length;
  $('kComplete').textContent=complete;
  $('rankingSummary').textContent=`${list.length} proyecto${list.length===1?'':'s'} visible${list.length===1?'':'s'}`;
  renderMatrix(list,evals);
  renderRanking(list);
  if(selected){const p=all.find(x=>x.id===selected);if(p)renderEditor(p)}
}
function renderMatrix(list,evals){
  const host=$('matrix');
  host.innerHTML='<div class="zone z1">Victoria Temprana<small>Alto puntaje · menor inversión</small></div><div class="zone z2">Mina de Oro<small>Alto puntaje · mayor inversión</small></div><div class="zone z3">Baja Prioridad<small>Bajo puntaje · menor inversión</small></div><div class="zone z4">Cuestionable<small>Bajo puntaje · mayor inversión</small></div><div class="axis-x">Inversión requerida (CAPEX)</div><div class="axis-y">Puntaje de Implementación</div>';
  const caps=evals.map(x=>Number(x.r.capex)).filter(v=>v>0).sort((a,b)=>a-b);
  const median=caps.length?(caps[Math.floor((caps.length-1)/2)]+caps[Math.ceil((caps.length-1)/2)])/2:100;
  const max=Math.max(1,...caps,median*1.25);
  const ycut=Math.max(0,Math.min(100,Number($('yCut').value)||70));
  host.style.setProperty('--xcut',(median/max*100)+'%');host.style.setProperty('--ycut',(100-ycut)+'%');
  evals.forEach(({p,r,s})=>{
    const b=document.createElement('button');
    b.className='point '+(s>=70?'high':s>=40?'med':'low')+(selected===p.id?' active':'');
    b.style.left=Math.min(98,Math.max(2,Number(r.capex)/max*100))+'%';
    b.style.top=(100-s)+'%';
    b.title=`${p.name}\nPuntaje: ${s}/100\nCAPEX: US$ ${Number(r.capex).toLocaleString('es-CR')} M`;
    b.setAttribute('aria-label',b.title.replaceAll('\n',', '));
    b.onclick=()=>select(p.id);
    host.appendChild(b);
  });
}
function pillClass(r,s){if(r.excluded)return'excluded';if(s==null)return'pending';if(s>=70)return'high';if(s<40)return'low';return'med'}
function renderRanking(list){
  const rows=list.map(p=>{const r=rec(p),s=score(r);return{p,r,s}}).sort((a,b)=>{
    if(a.r.excluded!==b.r.excluded)return a.r.excluded?1:-1;
    return (b.s??-1)-(a.s??-1);
  });
  $('ranking').innerHTML=rows.map((x,i)=>`<tr class="${selected===x.p.id?'active':''}" onclick="select('${x.p.id}')"><td>${i+1}</td><td class="project-cell"><b>${x.p.id}</b> · ${x.p.name}<small>${x.p.inst||'Institución por definir'} · BPIP ${x.p.bpip||'Por definir'}</small></td><td>${x.p.etapa||'—'}</td><td>${x.r.de||'—'}</td><td>${x.r.da||'—'}</td><td>${x.r.ds||'—'}</td><td>${x.r.mi||'—'}</td><td><b>${x.s??'—'}</b></td><td><span class="pill ${pillClass(x.r,x.s)}">${x.r.excluded?'No elegible / subsanación':cls(x.s)}</span></td><td>${x.r.capex?`US$ ${Number(x.r.capex).toLocaleString('es-CR')} M`:'—'}</td></tr>`).join('')||'<tr><td colspan="10">No hay proyectos con estos filtros.</td></tr>';
}
function select(id){selected=id;const p=all.find(x=>x.id===id);if(p)renderEditor(p);render()}
function openMainFicha(id){try{if(parent&&typeof parent.openMatrixFicha==='function'){parent.openMatrixFicha(null,id)}}catch(e){}}
function renderEditor(p){
  const r=rec(p),s=score(r),comp=completeness(r);
  $('editorBody').innerHTML=`
    <div class="project-title">${p.id} · ${p.name}</div>
    <div class="project-meta">${p.inst||'Institución por definir'} · ${p.etapa||'Etapa por definir'} · BPIP ${p.bpip||'Por definir'}</div>
    <div class="completion"><div class="completion-track"><div class="completion-fill" style="width:${comp}%"></div></div><b>${comp}% completo</b></div>
    <div class="score-grid">
      <label>Dimensión económica (DE)<span>Rentabilidad y eficiencia económica</span><input id="eDE" type="number" min="0" max="100" value="${r.de}"></label>
      <label>Dimensión ambiental (DA)<span>Impacto, descarbonización y resiliencia</span><input id="eDA" type="number" min="0" max="100" value="${r.da}"></label>
      <label>Dimensión social (DS)<span>Cobertura, accesibilidad y aceptación</span><input id="eDS" type="number" min="0" max="100" value="${r.ds}"></label>
      <label>Madurez implementación (MI)<span>Definición, viabilidad y gestión</span><input id="eMI" type="number" min="0" max="100" value="${r.mi}"></label>
    </div>
    <div class="capex-row">
      <label>CAPEX vigente · USD millones<input id="eCapex" type="number" min="0" step="0.1" value="${r.capex}"></label>
      <label>Fuente / observación<input id="eReason" value="${String(r.reason||'').replaceAll('"','&quot;')}" placeholder="Estudio, fecha o referencia"></label>
    </div>
    <label class="exclude"><input id="eExcluded" type="checkbox" ${r.excluded?'checked':''}><span><b>Condición excluyente o pendiente de subsanación</b><br>Marca esta opción cuando exista una condición que impida recomendar la programación del proyecto.</span></label>
    <div class="result-box"><strong>Resultado actual</strong><br>${r.excluded?'No elegible para programación / Pendiente de subsanación':s==null?'Pendiente de información de preinversión':`${s}/100 · ${cls(s)}`}<br><small>Puntaje = 25% DE + 25% DA + 25% DS + 25% MI.</small></div>
    <div class="editor-actions"><button onclick="saveEditor('${p.id}')">Guardar evaluación</button><button class="secondary" onclick="openMainFicha('${p.id}')">Ver ficha del proyecto</button><button class="secondary" onclick="clearEditor('${p.id}')">Limpiar valores</button></div>`;
}
function saveEditor(id){saved[id]={de:$('eDE').value,da:$('eDA').value,ds:$('eDS').value,mi:$('eMI').value,capex:$('eCapex').value,reason:$('eReason').value,excluded:$('eExcluded').checked};localStorage.setItem(storeKey,JSON.stringify(saved));render()}
function clearEditor(id){delete saved[id];localStorage.setItem(storeKey,JSON.stringify(saved));render()}
function report(){requestAnimationFrame(()=>parent.postMessage({type:'hub-priorizacion-height',target:'proyectos',height:Math.max(document.documentElement.scrollHeight,980)},'*'))}
new MutationObserver(report).observe(document.body,{subtree:true,childList:true,attributes:true});window.addEventListener('load',report);window.addEventListener('resize',report);init();
