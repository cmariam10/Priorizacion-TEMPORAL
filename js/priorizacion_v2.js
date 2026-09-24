const API_BASE = "https://hub-proyectos-mopt-api.cmariam10.workers.dev";
const API_SCOPE = "api://3a5ead05-e9e6-477b-9a63-24ef1f9fefd5/access_as_user";

// Esta URI ya estaba siendo utilizada por el HUB y debe existir en la
// configuración de autenticación de Microsoft Entra.
const AUTH_REDIRECT_URI = "https://cmariam10.github.io/Priorizacion-TEMPORAL/";

const MSAL_CONFIG = {
  auth: {
    clientId: "3a5ead05-e9e6-477b-9a63-24ef1f9fefd5",
    authority: "https://login.microsoftonline.com/common",
    redirectUri: AUTH_REDIRECT_URI
  },
  cache: {
    cacheLocation: "sessionStorage"
  }
};

const MSAL_CDNS = [
  "https://alcdn.msauth.net/browser/2.38.3/js/msal-browser.min.js",
  "https://cdn.jsdelivr.net/npm/@azure/msal-browser@2.38.3/lib/msal-browser.min.js"
];

let msalApp = null;
let msalInitPromise = null;
let config = null;
let projects = [];
let ranking = [];
let activeId = null;

function loadExternalScript(src) {
  return new Promise((resolve, reject) => {
    const existing = [...document.scripts].find(s => s.src === src);
    if (existing) {
      if (window.msal?.PublicClientApplication) return resolve();
      existing.addEventListener("load", resolve, { once: true });
      existing.addEventListener("error", () => reject(new Error(`No se pudo cargar ${src}`)), { once: true });
      return;
    }

    const script = document.createElement("script");
    script.src = src;
    script.async = true;
    script.onload = resolve;
    script.onerror = () => reject(new Error(`No se pudo cargar ${src}`));
    document.head.appendChild(script);
  });
}

async function ensureMsalLoaded() {
  if (window.msal?.PublicClientApplication) return;

  let lastError = null;
  for (const src of MSAL_CDNS) {
    try {
      await loadExternalScript(src);
      if (window.msal?.PublicClientApplication) return;
    } catch (e) {
      lastError = e;
      console.warn("Falló la carga de MSAL desde", src, e);
    }
  }

  throw new Error(
    "No se pudo cargar MSAL desde los CDN configurados. " +
    "Revise la conexión, el bloqueo del navegador o una política CSP." +
    (lastError ? ` Detalle: ${lastError.message}` : "")
  );
}

async function getMsalApp() {
  if (msalApp) return msalApp;
  if (msalInitPromise) return msalInitPromise;

  msalInitPromise = (async () => {
    await ensureMsalLoaded();
    const app = new window.msal.PublicClientApplication(MSAL_CONFIG);
    if (typeof app.initialize === "function") await app.initialize();

    try {
      const redirectResult = await app.handleRedirectPromise();
      if (redirectResult?.account) app.setActiveAccount(redirectResult.account);
    } catch (e) {
      console.warn("MSAL handleRedirectPromise:", e);
    }

    msalApp = app;
    return app;
  })();

  try {
    return await msalInitPromise;
  } finally {
    msalInitPromise = null;
  }
}

const INDICATOR_NAMES={
 "ND-01":"TPDA","ND-02":"IRI","ND-03":"Tasa de siniestralidad","ND-04":"Índice de Deficiencia Estructural","ND-05":"Restricción operacional","ND-06":"Interacción peatonal","ND-07":"Accidentes peatonales","ND-08":"Nivel de inestabilidad","ND-09":"Cierres por deslizamientos",
 "IS-01":"Población beneficiaria","IS-02":"Índice de Desarrollo Social","IS-03":"Acceso a actividades socioeconómicas",
 "VI-01":"Magnitud de inversión","VI-02":"Complejidad constructiva","VI-03":"Disponibilidad predial",
 "RR-01":"Exposición multiamenaza","RR-02":"Daño Anual Esperado","RR-03":"Pérdida Anual Esperada","RR-04":"Priorización BSA","RR-05":"Conflictividad social","RR-06":"Criticidad de conectividad",
 "SI-01":"Sensibilidad ambiental territorial","SI-02":"Descarbonización / sostenibilidad","SI-03":"Sostenibilidad operativa",
 "PG-01":"Prioridad institucional","PG-02":"Oportunidad de financiamiento","PG-03":"Continuidad programática","PG-04":"Urgencia gerencial","PG-05":"Madurez mínima"
};

const INDICATOR_ACQUISITION = {
  "ND-01":"manual", "ND-02":"manual", "ND-03":"manual",
  "ND-04":"manual", "ND-05":"manual", "ND-06":"geospatial", "ND-07":"geospatial", "ND-08":"geospatial", "ND-09":"manual",
  "IS-01":"manual", "IS-02":"manual", "IS-03":"geospatial",
  "VI-01":"manual", "VI-02":"manual", "VI-03":"geospatial",
  "RR-01":"geospatial", "RR-02":"manual", "RR-03":"manual", "RR-04":"manual", "RR-05":"manual", "RR-06":"geospatial",
  "SI-01":"geospatial", "SI-02":"manual", "SI-03":"manual",
  "PG-01":"managerial", "PG-02":"managerial", "PG-03":"managerial", "PG-04":"managerial", "PG-05":"managerial"
};

const ACQUISITION_LABELS = {
  geospatial: {title:"Automáticos geoespaciales", note:"Se obtienen mediante cruces con APOYO_GEO en OneDrive y fuentes oficiales. No son editables manualmente."},
  official: {title:"Automáticos / fuentes institucionales", note:"Se obtienen del HUB o de fuentes oficiales cuando el conector esté disponible."},
  manual: {title:"Datos ingresados por el usuario", note:"Información institucional que debe registrar o validar la persona funcionaria."},
  managerial: {title:"Priorización gerencial", note:"Criterios gerenciales registrados por la persona funcionaria con su soporte correspondiente."}
};

function esc(v){return String(v??"").replace(/[&<>"']/g,m=>({"&":"&amp;","<":"&lt;",">":"&gt;","\"":"&quot;","'":"&#39;"}[m]));}
function hasScore(v){return v!==null&&v!==undefined&&v!==""&&Number.isFinite(Number(v));}
function fmt(v){return hasScore(v)?Number(v).toFixed(1):"—";}
function scoreLabel(v){return hasScore(v)?Number(v).toFixed(1):"Pendiente";}
function toast(msg){const el=document.getElementById("toast");el.textContent=msg;el.hidden=false;clearTimeout(toast.t);toast.t=setTimeout(()=>el.hidden=true,3500);}
async function token() {
  const app = await getMsalApp();
  let account = app.getActiveAccount() || app.getAllAccounts()[0];

  if (!account) {
    const login = await app.loginPopup({
      scopes: [API_SCOPE],
      prompt: "select_account"
    });
    account = login.account;
    if (account) app.setActiveAccount(account);
    if (login.accessToken) return login.accessToken;
  }

  if (!account) throw new Error("No fue posible obtener la cuenta de Microsoft.");

  try {
    const silent = await app.acquireTokenSilent({ scopes: [API_SCOPE], account });
    return silent.accessToken;
  } catch (silentError) {
    console.warn("Token silencioso no disponible; se abrirá Microsoft.", silentError);
    const popup = await app.acquireTokenPopup({ scopes: [API_SCOPE], account });
    return popup.accessToken;
  }
}

async function api(path, options = {}, auth = false) {
  // Cuando el módulo corre embebido dentro del HUB, reutiliza la sesión
  // y el token que el HUB principal ya administra. Esto evita abrir una
  // segunda instancia MSAL dentro del iframe y bloqueos en loginPopup.
  if (auth && window.parent && window.parent !== window) {
    try {
      if (typeof window.parent.hubApiFetch === "function") {
        return await window.parent.hubApiFetch(path, options);
      }
    } catch (e) {
      // Si el parent no es accesible o falla por alguna razón, usamos
      // el flujo MSAL propio como respaldo para ejecución standalone.
      console.warn("No se pudo reutilizar la autenticación del HUB padre; se usa MSAL local.", e);
    }
  }

  const headers = new Headers(options.headers || {});
  if (options.body && !headers.has("Content-Type")) headers.set("Content-Type", "application/json");
  if (auth) headers.set("Authorization", `Bearer ${await token()}`);

  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), 45000);
  try {
    const response = await fetch(API_BASE + path, { ...options, headers, signal: controller.signal });
    const data = await response.json().catch(() => ({}));
    if (!response.ok) throw new Error(data.error || data.message || `HTTP ${response.status}`);
    return data;
  } catch (e) {
    if (e?.name === "AbortError") {
      throw new Error("La solicitud al HUB superó 45 segundos y fue cancelada.");
    }
    throw e;
  } finally {
    clearTimeout(timer);
  }
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
 rows.sort((a,b)=>{const af=hasScore(a.finalScore)?Number(a.finalScore):null,bf=hasScore(b.finalScore)?Number(b.finalScore):null;if(af!==null||bf!==null)return(bf??-1)-(af??-1);const bt=hasScore(b.technicalScore)?Number(b.technicalScore):-1,at=hasScore(a.technicalScore)?Number(a.technicalScore):-1;return bt-at;});
 return rows;
}
function render(){
 const rows=filteredRows();document.getElementById("kProjects").textContent=projects.length;document.getElementById("kEvaluated").textContent=ranking.length;document.getElementById("kTechnical").textContent=ranking.filter(x=>hasScore(x.technicalScore)).length;document.getElementById("kFinal").textContent=ranking.filter(x=>hasScore(x.finalScore)).length;
 document.getElementById("rankingBody").innerHTML=rows.map((r,i)=>`<tr><td>${i+1}</td><td><b>${esc(r.projectId)}</b></td><td>${esc(r.project?.name||"Sin nombre")}<div class="muted">${esc(r.project?.institution||"")}</div></td><td>${r.matrixId?`<span class="pill">${r.matrixId}</span>`:'<span class="pill warn">Sin asignar</span>'}</td><td class="score">${fmt(r.technicalScore)}</td><td>${fmt(r.managerialScore)}</td><td class="score">${fmt(r.finalScore)}</td><td>${esc(r.classification||"—")}</td><td>${r.missingCount??"—"}</td><td><button class="row-action" data-id="${esc(r.projectId)}">Abrir</button></td></tr>`).join("")||'<tr><td colspan="10" class="muted">No hay proyectos que coincidan.</td></tr>';
 document.querySelectorAll(".row-action").forEach(b=>b.onclick=()=>openProject(b.dataset.id));
}
function applicableIndicators(matrixId){
 const m=config.matrices?.[matrixId];if(!m)return[...Object.keys(INDICATOR_NAMES)];
 const ids=[];Object.values(m.criteria||{}).forEach(c=>ids.push(...c.indicators));return[...new Set([...ids,"PG-01","PG-02","PG-03","PG-04","PG-05"])];
}
function renderEditor(data){
 const matrixId=data?.matrix?.matrixId||data?.matrixOverride||"";
 const ids=applicableIndicators(matrixId);
 const raw=data?.inputs?.raw||{}, scores=data?.inputs?.scores||{}, computed=data?.scores?.indicatorScores||{};
 const groups={geospatial:[],official:[],manual:[],managerial:[]};
 ids.forEach(id=>(groups[INDICATOR_ACQUISITION[id]||"manual"]||groups.manual).push(id));
 const html=[];
 for(const mode of ["geospatial","official","manual","managerial"]){
   const list=groups[mode]; if(!list.length) continue;
   const meta=ACQUISITION_LABELS[mode];
   html.push(`<div class="indicator-group"><div class="indicator-group-head"><div><h4>${esc(meta.title)}</h4><p>${esc(meta.note)}</p></div><span class="mode-pill ${mode}">${mode==="geospatial"?"GIS":mode==="official"?"FUENTE":mode==="managerial"?"GERENCIAL":"MANUAL"}</span></div><div class="indicator-grid">`);
   for(const id of list){
     const value=hasScore(computed[id])?computed[id]:(hasScore(scores[id])?scores[id]:null);
     if(mode==="geospatial"||mode==="official"){
       html.push(`<div class="indicator indicator-auto" data-indicator="${id}" data-mode="${mode}"><div class="indicator-head"><b>${id}</b><span>Calculado: ${fmt(value)}</span></div><div>${esc(INDICATOR_NAMES[id]||id)}</div><div class="auto-value"><strong>${fmt(value)}</strong><span>${mode==="geospatial"?"Resultado automático del análisis espacial":"Dato automático / institucional"}</span></div><small>${value===null?"Pendiente de fuente o de una corrida válida.":"Este valor no se edita manualmente."}</small></div>`);
     }else{
       html.push(`<div class="indicator" data-indicator="${id}" data-mode="${mode}"><div class="indicator-head"><b>${id}</b><span>Calculado: ${fmt(computed[id])}</span></div><div>${esc(INDICATOR_NAMES[id]||id)}</div><div class="indicator-inputs"><input class="raw" value="${esc(raw[id]??"")}" placeholder="Dato original / categoría"><input class="manual-score" type="number" min="0" max="100" step="1" value="${esc(scores[id]??"")}" placeholder="0–100"></div><small>El puntaje validado tiene prioridad sobre la normalización del dato original.</small></div>`);
     }
   }
   html.push(`</div></div>`);
 }
 document.getElementById("indicatorEditor").innerHTML=html.join("");
}


let evidenceMap=null;
let evidenceOverlayControl=null;
function evidenceLayerLabel(k){return ({asp:"Áreas silvestres protegidas",wetlands:"Humedales",corridors:"Corredores biológicos",landslides:"Deslizamientos",crowns:"Coronas de deslizamiento",floods:"Áreas con potencial de inundación"})[k]||k;}
function renderSpatialMap(spatial){
 const el=document.getElementById("spatialMap");
 if(!el||!window.L)return;
 if(!evidenceMap){
   evidenceMap=L.map(el,{preferCanvas:true});
   L.tileLayer("https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png",{maxZoom:19,attribution:"© OpenStreetMap"}).addTo(evidenceMap);
 }
 evidenceMap.eachLayer(layer=>{if(!(layer instanceof L.TileLayer)) evidenceMap.removeLayer(layer);});
 if(evidenceOverlayControl){evidenceMap.removeControl(evidenceOverlayControl);evidenceOverlayControl=null;}
 const overlays={}; const bounds=[];
 const project=spatial?.projectFeature;
 if(project?.geometry){
   const l=L.geoJSON(project,{style:{weight:5},pointToLayer:(f,ll)=>L.circleMarker(ll,{radius:7,weight:3,fillOpacity:.8})}).addTo(evidenceMap);
   overlays["Geometría del proyecto"]=l; if(l.getBounds?.().isValid()) bounds.push(l.getBounds());
 }
 for(const [k,v] of Object.entries(spatial?.evidence||{})){
   const fc=v?.previewGeoJSON; if(!fc?.features?.length) continue;
   const l=L.geoJSON(fc,{style:{weight:2,fillOpacity:.18},pointToLayer:(f,ll)=>L.circleMarker(ll,{radius:5,weight:2,fillOpacity:.5}),onEachFeature:(f,layer)=>{const props=f.properties||{};const rows=Object.entries(props).slice(0,6).map(([a,b])=>`<b>${esc(a)}:</b> ${esc(b)}`).join("<br>"); if(rows)layer.bindPopup(rows);}}).addTo(evidenceMap);
   overlays[evidenceLayerLabel(k)]=l; if(l.getBounds?.().isValid()) bounds.push(l.getBounds());
 }
 evidenceOverlayControl=L.control.layers(null,overlays,{collapsed:false}).addTo(evidenceMap);
 if(bounds.length){let b=bounds[0]; for(let i=1;i<bounds.length;i++)b=b.extend(bounds[i]); evidenceMap.fitBounds(b.pad(.15),{maxZoom:15});}
 else evidenceMap.setView([9.93,-84.08],7);
 setTimeout(()=>evidenceMap.invalidateSize(),50);
}
function renderEvidence(spatial){
 const ev=spatial?.evidence||{};const cards=[];
 for(const [k,v] of Object.entries(ev)){
  if(!v||typeof v!=="object")continue;
  const m=v.metrics;
  const stale=v.refreshStatus==="stale";
  const fresh=v.refreshStatus==="fresh";
  const estado=stale?"guardado (fuente no disponible en esta corrida)":fresh?"actualizado":(v.status||"—");
  cards.push(`<div class="evidence-card"><b>${esc(evidenceLayerLabel(k))}</b><p>Estado: ${esc(estado)}${v.sourceName?`<br>Fuente: ${esc(v.sourceName)}`:""}${v.layerName?`<br>Capa: ${esc(v.layerName)}`:(v.typeName?`<br>Capa: ${esc(v.typeName)}`:"")}${v.file?`<br>Archivo: ${esc(v.file)}`:""}${m?`<br>Intersecciones: ${m.intersectCount}<br>En ${m.nearMeters} m: ${m.nearCount}<br>Consultados en zona: ${m.totalFetched}`:""}${v.lastValidAt?`<br>Último dato válido: ${esc(new Date(v.lastValidAt).toLocaleString())}`:""}${v.refreshError?`<br>Actualización fallida: ${esc(v.refreshError)}`:""}${v.error&&!stale?`<br>Error: ${esc(v.error)}`:""}</p></div>`);
 }
 const sugg=spatial?.suggestedScores||{};
 for(const [id,s] of Object.entries(sugg))cards.push(`<div class="evidence-card"><b>Sugerencia ${id}: ${s.score}/100</b><p>${esc(s.note||"Derivada de relación espacial; requiere validación metodológica.")}</p></div>`);
 const summary=spatial?.summary;
 if(summary)cards.unshift(`<div class="evidence-card"><b>Estado de fuentes</b><p>Actualizadas: ${summary.fresh??summary.ok??0}<br>Desde APOYO_GEO: ${summary.localOk??0}<br>Con dato guardado: ${summary.stale??0}<br>No disponibles: ${summary.unavailable??summary.failed??0}</p></div>`);
 document.getElementById("spatialEvidence").innerHTML=cards.join("")||'<div class="muted">Todavía no hay evidencia espacial calculada.</div>';
 renderSpatialMap(spatial);
}
async function openProject(id){
 activeId=id;const p=projects.find(x=>projectId(x)===id);const r=await api(`/projects/${encodeURIComponent(id)}/prioritization`);const d=r.prioritization||{projectId:id,project:{name:projectName(p),institution:projectInst(p)},inputs:{raw:{},scores:{}}};
 document.getElementById("dialogTitle").textContent=`${id} · ${d.project?.name||projectName(p)}`;document.getElementById("dialogMeta").textContent=`${d.project?.institution||projectInst(p)} · ${d.matrix?.matrixId||"Matriz sin asignar"}`;
 document.getElementById("dTechnical").textContent=scoreLabel(d.scores?.technicalScore);document.getElementById("dManagerial").textContent=scoreLabel(d.scores?.managerialScore);document.getElementById("dFinal").textContent=scoreLabel(d.scores?.finalScore);
 const ms=document.getElementById("matrixOverride");ms.innerHTML='<option value="">Asignación automática</option>'+Object.entries(config.matrices||{}).map(([mid,m])=>`<option value="${mid}">${mid} · ${esc(m.typology)}</option>`).join("");ms.value=d.matrixOverride||"";
 document.getElementById("technicalWeight").value=d.finalAggregation?.technicalWeight??"";document.getElementById("managerialWeight").value=d.finalAggregation?.managerialWeight??"";
 renderEditor(d);renderEvidence(d.spatial);
 const lastValid=d.calculatedAt?new Date(d.calculatedAt).toLocaleString():null;
 const lastAttempt=d.lastAttempt?.calculatedAt?new Date(d.lastAttempt.calculatedAt).toLocaleString():null;
 const attemptValid=d.lastAttempt?.valid;
 let statusHtml; if(!d.matrix?.matrixId){statusHtml="<b>Evaluación pendiente.</b> Primero debe asignarse una matriz metodológica; después el HUB mostrará los indicadores aplicables.";}else if(d.scores?.missing?.length){statusHtml=`Faltan <b>${d.scores.missing.length}</b> indicadores para completar el cálculo: ${d.scores.missing.join(", ")}.`;}else{statusHtml=lastValid?`Último cálculo válido: <b>${lastValid}</b>.`:"Este proyecto todavía no ha sido calculado.";}
 if(lastAttempt && attemptValid===false && lastValid)statusHtml+=`<br>La actualización de <b>${lastAttempt}</b> no reemplazó el resultado guardado.`;
 document.getElementById("dialogStatus").innerHTML=statusHtml;
 document.getElementById("projectDialog").showModal();
}
function collectInput(){
 const raw={},scores={};document.querySelectorAll(".indicator").forEach(box=>{const id=box.dataset.indicator;const rawEl=box.querySelector(".raw"),scoreEl=box.querySelector(".manual-score");if(!rawEl&&!scoreEl)return;const rv=rawEl?.value?.trim()||"",sv=scoreEl?.value?.trim()||"";if(rv!=="")raw[id]=rv;if(sv!=="")scores[id]=Number(sv);});
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
function runUiAction(fn) {
  return async (...args) => {
    try {
      await fn(...args);
    } catch (e) {
      console.error("ERROR PRIORIZACIÓN:", e);
      toast(e?.message || "Ocurrió un error en el módulo de priorización.");

      const status = document.getElementById("dialogStatus");
      if (status) {
        status.innerHTML = `<b>Error:</b> ${esc(e?.message || "Error desconocido")}`;
      }
    }
  };
}

function initPrioritizationEvents() {
  console.log("PRIORIZACIÓN V2: eventos inicializados");

  ["q", "matrixFilter", "institutionFilter", "classFilter"].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.addEventListener("input", render);
  });

  const btnRefresh = document.getElementById("btnRefresh");
  if (btnRefresh) btnRefresh.addEventListener("click", runUiAction(loadAll));

  const btnBatch = document.getElementById("btnBatch");
  if (btnBatch) btnBatch.addEventListener("click", runUiAction(recalcBatch));

  const btnCsv = document.getElementById("btnCsv");
  if (btnCsv) btnCsv.addEventListener("click", exportCsv);

  const btnSave = document.getElementById("btnSave");
  if (btnSave) {
    btnSave.addEventListener("click", async event => {
      event.preventDefault();
      event.stopPropagation();
      console.log("PRIORIZACIÓN V2: clic Guardar", activeId);
      await runUiAction(saveInput)();
    });
  }

  const btnCalculate = document.getElementById("btnCalculate");
  if (btnCalculate) {
    btnCalculate.addEventListener("click", async event => {
      event.preventDefault();
      event.stopPropagation();

      console.log("PRIORIZACIÓN V2: clic CALCULAR", activeId);

      const status = document.getElementById("dialogStatus");

      if (!activeId) {
        const msg = "No hay un proyecto activo para calcular.";
        console.error(msg);
        if (status) status.innerHTML = `<b>Error:</b> ${msg}`;
        return;
      }

      if (status) {
        status.innerHTML = "<b>Calculando…</b> Guardando configuración, consultando geometría y ejecutando análisis geoespacial.";
      }

      btnCalculate.disabled = true;
      btnCalculate.textContent = "Calculando…";

      try {
        console.log("1. Datos enviados:", collectInput());
        console.log("2. Guardando configuración...");

        await api(
          `/projects/${encodeURIComponent(activeId)}/prioritization`,
          { method: "PUT", body: JSON.stringify(collectInput()) },
          true
        );

        console.log("3. Configuración guardada.");
        console.log("4. Iniciando cálculo...");

        const result = await api(
          `/projects/${encodeURIComponent(activeId)}/prioritization/calculate`,
          { method: "POST" },
          true
        );

        console.log("5. RESULTADO CÁLCULO:", result);

        if (status) status.innerHTML = "<b>Cálculo completado.</b> Actualizando resultados…";
        toast("Priorización recalculada.");

        await loadAll();
        await openProject(activeId);
      } catch (e) {
        console.error("ERROR AL CALCULAR:", e);
        if (status) status.innerHTML = `<b>Error al calcular:</b> ${esc(e?.message || "Error desconocido")}`;
        toast(e?.message || "No fue posible calcular.");
      } finally {
        btnCalculate.disabled = false;
        btnCalculate.textContent = "Calcular / recalcular";
      }
    });
  }

  const matrixOverride = document.getElementById("matrixOverride");
  if (matrixOverride) {
    matrixOverride.addEventListener("change", () => {
      console.log("Matriz seleccionada:", matrixOverride.value);
      renderEditor({
        matrixOverride: matrixOverride.value,
        inputs: { raw: {}, scores: {} }
      });
    });
  }

  const btnCloseDialog = document.getElementById("btnCloseDialog");
  if (btnCloseDialog) {
    btnCloseDialog.addEventListener("click", () => {
      document.getElementById("projectDialog")?.close();
    });
  }
}

async function startPrioritization() {
  try {
    initPrioritizationEvents();
    console.log("PRIORIZACIÓN V2: cargando información...");
    await loadAll();
    console.log("PRIORIZACIÓN V2: módulo listo");
  } catch (e) {
    console.error("ERROR INICIALIZANDO PRIORIZACIÓN:", e);
    toast(e?.message || "No fue posible cargar el módulo.");

    const notice = document.getElementById("methodologyNotice");
    if (notice) {
      notice.textContent = "No fue posible cargar el módulo: " + (e?.message || "Error desconocido");
    }
  }
}

if (document.readyState === "loading") {
  document.addEventListener("DOMContentLoaded", startPrioritization, { once: true });
} else {
  startPrioritization();
}
