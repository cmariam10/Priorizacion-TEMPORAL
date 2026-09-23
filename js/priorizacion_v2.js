const API_BASE = "https://hub-proyectos-mopt-api.cmariam10.workers.dev";
const API_SCOPE="api://3a5ead05-e9e6-477b-9a63-24ef1f9fefd5/access_as_user";
const MSAL_CONFIG={auth:{clientId:"3a5ead05-e9e6-477b-9a63-24ef1f9fefd5",authority:"https://login.microsoftonline.com/common",redirectUri:"https://cmariam10.github.io/Priorizacion-TEMPORAL/"},cache:{cacheLocation:"sessionStorage"}};
let msalApp=null,config=null,projects=[],ranking=[],activeId=null;

const INDICATOR_NAMES={
 "ND-01":"TPDA","ND-02":"IRI","ND-03":"Tasa de siniestralidad","ND-04":"Índice de Deficiencia Estructural","ND-05":"Restricción operacional","ND-06":"Interacción peatonal","ND-07":"Accidentes peatonales","ND-08":"Nivel de inestabilidad","ND-09":"Cierres por deslizamientos",
 "IS-01":"Población beneficiaria","IS-02":"Índice de Desarrollo Social","IS-03":"Acceso a actividades socioeconómicas",
 "VI-01":"Magnitud de inversión","VI-02":"Complejidad constructiva","VI-03":"Disponibilidad predial",
 "RR-01":"Exposición multiamenaza","RR-02":"Daño Anual Esperado","RR-03":"Pérdida Anual Esperada","RR-04":"Priorización BSA","RR-05":"Conflictividad social","RR-06":"Criticidad de conectividad",
 "SI-01":"Sensibilidad ambiental territorial","SI-02":"Descarbonización / sostenibilidad","SI-03":"Sostenibilidad operativa",
 "PG-01":"Prioridad institucional","PG-02":"Oportunidad de financiamiento","PG-03":"Continuidad programática","PG-04":"Urgencia gerencial","PG-05":"Madurez mínima"
};
function esc(v){return String(v??"").replace(/[&<>"']/g,m=>({"&":"&amp;","<":"&lt;",">":"&gt;","\"":"&quot;","'":"&#39;"}[m]));}
function fmt(v){return Number.isFinite(Number(v))?Number(v).toFixed(1):"—";}
function toast(msg){const el=document.getElementById("toast");el.textContent=msg;el.hidden=false;clearTimeout(toast.t);toast.t=setTimeout(()=>el.hidden=true,3500);}
async function token(){
 if(!window.msal)throw new Error("No se cargó MSAL.");
 if(!msalApp){msalApp=new msal.PublicClientApplication(MSAL_CONFIG);if(msalApp.initialize)await msalApp.initialize();await msalApp.handleRedirectPromise();}
 let account=msalApp.getActiveAccount()||msalApp.getAllAccounts()[0];
 if(!account){const r=await msalApp.loginPopup({scopes:[API_SCOPE],prompt:"select_account"});account=r.account;msalApp.setActiveAccount(account);if(r.accessToken)return r.accessToken;}
 try{return (await msalApp.acquireTokenSilent({scopes:[API_SCOPE],account})).accessToken;}catch{return (await msalApp.acquireTokenPopup({scopes:[API_SCOPE],account})).accessToken;}
}
async function api(path,options={},auth=false){
 const headers=new Headers(options.headers||{});if(options.body&&!headers.has("Content-Type"))headers.set("Content-Type","application/json");if(auth)headers.set("Authorization",`Bearer ${await token()}`);
 const r=await fetch(API_BASE+path,{...options,headers});const data=await r.json().catch(()=>({}));if(!r.ok)throw new Error(data.error||data.message||`HTTP ${r.status}`);return data;
}
function projectId(p){return String(p?.["ID HUB"]||p?.idHub||p?.id||"").trim();}
function projectName(p){return p?.["Nombre del proyecto"]||p?.["Nombre corto"]||p?.name||"Sin nombre";}
function projectInst(p){return p?.["Institución"]||p?.institution||p?.inst||"";}
async function loadAll(){
 const [c,p,r]=await Promise.all([api("/prioritization/config"),api("/projects"),api("/prioritization/ranking")]);
 config=c;projects=p.projects||[];ranking=r.ranking||[];populateFilters();render();
 const agg=c.finalAggregation||{};
 document.getElementById("methodologyNotice").innerHTML=`<b>Metodología ${esc(c.version)}.</b> Las matrices M-01 a M-05 están parametrizadas. La guía recibida no explicita la proporción numérica Técnico/Gerencial; por eso el Puntaje Final queda vacío hasta que se configuren pesos aprobados.`;
}
function populateFilters(){
 const mf=document.getElementById("matrixFilter");const selected=mf.value;mf.innerHTML='<option value="">Todas</option>'+Object.entries(config.matrices||{}).map(([id,m])=>`<option value="${id}">${id} · ${esc(m.typology)}</option>`).join("");mf.value=selected;
 const inf=document.getElementById("institutionFilter");const si=inf.value;const inst=[...new Set(projects.map(projectInst).filter(Boolean))].sort();inf.innerHTML='<option value="">Todas</option>'+inst.map(x=>`<option>${esc(x)}</option>`).join("");inf.value=si;
}
function mergedRows(){
 const byId=new Map(ranking.map(r=>[String(r.projectId),r]));
 return projects.map(p=>{const id=projectId(p),r=byId.get(id)||{};return {...r,projectId:id,project:r.project||{idHub:id,name:projectName(p),institution:projectInst(p),stage:p["Etapa"]||""}};});
}
function filteredRows(){
 const q=document.getElementById("q").value.trim().toLowerCase(),m=document.getElementById("matrixFilter").value,inst=document.getElementById("institutionFilter").value,cl=document.getElementById("classFilter").value;
 let rows=mergedRows().filter(r=>{const txt=`${r.projectId} ${r.project?.name||""} ${r.project?.institution||""}`.toLowerCase();return(!q||txt.includes(q))&&(!m||r.matrixId===m)&&(!inst||r.project?.institution===inst)&&(!cl||r.classification===cl);});
 rows.sort((a,b)=>{const af=Number.isFinite(+a.finalScore)?+a.finalScore:null,bf=Number.isFinite(+b.finalScore)?+b.finalScore:null;if(af!==null||bf!==null)return(bf??-1)-(af??-1);return(+b.technicalScore||-1)-(+a.technicalScore||-1);});
 return rows;
}
function render(){
 const rows=filteredRows();document.getElementById("kProjects").textContent=projects.length;document.getElementById("kEvaluated").textContent=ranking.length;document.getElementById("kTechnical").textContent=ranking.filter(x=>Number.isFinite(+x.technicalScore)).length;document.getElementById("kFinal").textContent=ranking.filter(x=>Number.isFinite(+x.finalScore)).length;
 document.getElementById("rankingBody").innerHTML=rows.map((r,i)=>`<tr><td>${i+1}</td><td><b>${esc(r.projectId)}</b></td><td>${esc(r.project?.name||"Sin nombre")}<div class="muted">${esc(r.project?.institution||"")}</div></td><td>${r.matrixId?`<span class="pill">${r.matrixId}</span>`:'<span class="pill warn">Sin asignar</span>'}</td><td class="score">${fmt(r.technicalScore)}</td><td>${fmt(r.managerialScore)}</td><td class="score">${fmt(r.finalScore)}</td><td>${esc(r.classification||"—")}</td><td>${r.missingCount??"—"}</td><td><button class="row-action" data-id="${esc(r.projectId)}">Abrir</button></td></tr>`).join("")||'<tr><td colspan="10" class="muted">No hay proyectos que coincidan.</td></tr>';
 document.querySelectorAll(".row-action").forEach(b=>b.onclick=()=>openProject(b.dataset.id));
}
function applicableIndicators(matrixId){
 const m=config.matrices?.[matrixId];if(!m)return[...Object.keys(INDICATOR_NAMES)];
 const ids=[];Object.values(m.criteria||{}).forEach(c=>ids.push(...c.indicators));return[...new Set([...ids,"PG-01","PG-02","PG-03","PG-04","PG-05"])];
}
function renderEditor(data){
 const matrixId=data?.matrix?.matrixId||data?.matrixOverride||"";const ids=applicableIndicators(matrixId);const raw=data?.inputs?.raw||{},scores=data?.inputs?.scores||{},computed=data?.scores?.indicatorScores||{};
 document.getElementById("indicatorEditor").innerHTML=ids.map(id=>`<div class="indicator" data-indicator="${id}"><div class="indicator-head"><b>${id}</b><span>Calculado: ${fmt(computed[id])}</span></div><div>${esc(INDICATOR_NAMES[id]||id)}</div><div class="indicator-inputs"><input class="raw" value="${esc(raw[id]??"")}" placeholder="Dato original / categoría"><input class="manual-score" type="number" min="0" max="100" step="1" value="${esc(scores[id]??"")}" placeholder="0–100"></div><small>Si se ingresa puntaje validado, tiene prioridad sobre la normalización del dato original.</small></div>`).join("");
}
function renderEvidence(spatial){
 const ev=spatial?.evidence||{};const cards=[];
 for(const [k,v] of Object.entries(ev)){if(!v||typeof v!=="object")continue;const m=v.metrics;cards.push(`<div class="evidence-card"><b>${esc(k)}</b><p>Estado: ${esc(v.status||"—")}${v.typeName?`<br>Capa: ${esc(v.typeName)}`:""}${m?`<br>Intersecciones: ${m.intersectCount}<br>En ${m.nearMeters} m: ${m.nearCount}<br>Consultados: ${m.totalFetched}`:""}${v.error?`<br>Error: ${esc(v.error)}`:""}</p></div>`);}
 const sugg=spatial?.suggestedScores||{};for(const [id,s] of Object.entries(sugg))cards.push(`<div class="evidence-card"><b>Sugerencia ${id}: ${s.score}/100</b><p>${esc(s.note||"Derivada de relación espacial; requiere validación metodológica.")}</p></div>`);
 document.getElementById("spatialEvidence").innerHTML=cards.join("")||'<div class="muted">Todavía no hay evidencia espacial calculada.</div>';
}
async function openProject(id){
 activeId=id;const p=projects.find(x=>projectId(x)===id);const r=await api(`/projects/${encodeURIComponent(id)}/prioritization`);const d=r.prioritization||{projectId:id,project:{name:projectName(p),institution:projectInst(p)},inputs:{raw:{},scores:{}}};
 document.getElementById("dialogTitle").textContent=`${id} · ${d.project?.name||projectName(p)}`;document.getElementById("dialogMeta").textContent=`${d.project?.institution||projectInst(p)} · ${d.matrix?.matrixId||"Matriz sin asignar"}`;
 document.getElementById("dTechnical").textContent=fmt(d.scores?.technicalScore);document.getElementById("dManagerial").textContent=fmt(d.scores?.managerialScore);document.getElementById("dFinal").textContent=fmt(d.scores?.finalScore);
 const ms=document.getElementById("matrixOverride");ms.innerHTML='<option value="">Asignación automática</option>'+Object.entries(config.matrices||{}).map(([mid,m])=>`<option value="${mid}">${mid} · ${esc(m.typology)}</option>`).join("");ms.value=d.matrixOverride||"";
 document.getElementById("technicalWeight").value=d.finalAggregation?.technicalWeight??"";document.getElementById("managerialWeight").value=d.finalAggregation?.managerialWeight??"";
 renderEditor(d);renderEvidence(d.spatial);document.getElementById("dialogStatus").innerHTML=d.scores?.missing?.length?`Faltan <b>${d.scores.missing.length}</b> indicadores para completar el cálculo: ${d.scores.missing.join(", ")}.`:`${d.calculatedAt?`Último cálculo: <b>${new Date(d.calculatedAt).toLocaleString()}</b>.`:"Este proyecto todavía no ha sido calculado."}`;
 document.getElementById("projectDialog").showModal();
}
function collectInput(){
 const raw={},scores={};document.querySelectorAll(".indicator").forEach(box=>{const id=box.dataset.indicator,rv=box.querySelector(".raw").value.trim(),sv=box.querySelector(".manual-score").value.trim();if(rv!=="")raw[id]=rv;if(sv!=="")scores[id]=Number(sv);});
 const tw=document.getElementById("technicalWeight").value.trim(),gw=document.getElementById("managerialWeight").value.trim();
 return{matrixOverride:document.getElementById("matrixOverride").value,raw,scores,finalAggregation:{technicalWeight:tw===""?null:Number(tw),managerialWeight:gw===""?null:Number(gw)}};
}
async function saveInput(){if(!activeId)return;await api(`/projects/${encodeURIComponent(activeId)}/prioritization`,{method:"PUT",body:JSON.stringify(collectInput())},true);toast("Datos guardados.");await openProject(activeId);}
async function calculateActive(){if(!activeId)return;await api(`/projects/${encodeURIComponent(activeId)}/prioritization`,{method:"PUT",body:JSON.stringify(collectInput())},true);const r=await api(`/projects/${encodeURIComponent(activeId)}/prioritization/calculate`,{method:"POST"},true);toast("Priorización recalculada.");await loadAll();await openProject(activeId);return r;}
async function recalcBatch(){
 let offset=0,total=null;document.getElementById("btnBatch").disabled=true;
 try{do{const r=await api("/prioritization/recalculate",{method:"POST",body:JSON.stringify({offset,limit:10})},true);total=r.total;offset=r.nextOffset;toast(`Procesados ${Math.min(offset??total,total)} de ${total}`);}while(offset!==null);await loadAll();toast("Portafolio recalculado.");}finally{document.getElementById("btnBatch").disabled=false;}
}
function exportCsv(){const rows=filteredRows(),head=["Ranking","ID HUB","Proyecto","Institución","Matriz","Puntaje Técnico","Puntaje Gerencial","Puntaje Final","Clasificación","Faltantes"];const q=v=>`"${String(v??"").replace(/"/g,'""')}"`;const body=[head,...rows.map((r,i)=>[i+1,r.projectId,r.project?.name,r.project?.institution,r.matrixId,r.technicalScore,r.managerialScore,r.finalScore,r.classification,r.missingCount])].map(a=>a.map(q).join(",")).join("\n");const blob=new Blob(["\ufeff"+body],{type:"text/csv;charset=utf-8"});const a=document.createElement("a");a.href=URL.createObjectURL(blob);a.download="ranking_priorizacion_hub.csv";a.click();URL.revokeObjectURL(a.href);}
["q","matrixFilter","institutionFilter","classFilter"].forEach(id=>document.getElementById(id).addEventListener("input",render));document.getElementById("btnRefresh").onclick=loadAll;document.getElementById("btnBatch").onclick=recalcBatch;document.getElementById("btnCsv").onclick=exportCsv;document.getElementById("btnSave").onclick=saveInput;document.getElementById("btnCalculate").onclick=calculateActive;document.getElementById("matrixOverride").onchange=()=>renderEditor({matrixOverride:document.getElementById("matrixOverride").value,inputs:{raw:{},scores:{}}});
loadAll().catch(e=>{console.error(e);toast(e.message);document.getElementById("methodologyNotice").textContent="No fue posible cargar el módulo: "+e.message;});
