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
const DECISION_DETAIL_CACHE = new Map();
const decisionFilters = { dateFrom:"", dateTo:"", sort:"ranking" };
const PRIORITIZATION_PARAMS=new URLSearchParams(window.location.search);
const PRIORITIZATION_SCOPE = PRIORITIZATION_PARAMS.get("scope") || "all";
const PRIORITIZATION_VIEW_MODE=PRIORITIZATION_SCOPE==="decision"?"readonly":(PRIORITIZATION_PARAMS.get("mode")||"edit");
const DECISION_IDS=new Set((PRIORITIZATION_PARAMS.get("decisionIds")||"").split(",").map(x=>x.trim()).filter(Boolean));
const IMPLEMENTATION_SELECTED_IDS=new Set((PRIORITIZATION_PARAMS.get("selectedIds")||"").split(",").map(x=>x.trim()).filter(Boolean));
const DECISION_WORKFLOW_CACHE=new Map();
function notifyParentPrioritizationUpdated(projectId, prioritization=null){
  try{parent.postMessage({type:"hub-prioritization-updated",projectId:String(projectId||""),missing:prioritization?.scores?.missing||[],updatedAt:prioritization?.updatedAt||prioritization?.calculatedAt||new Date().toISOString()},"*");}catch(e){}
}

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
  "IS-01":"manual", "IS-02":"manual", "IS-03":"manual",
  "VI-01":"manual", "VI-02":"manual", "VI-03":"manual",
  "RR-01":"geospatial", "RR-02":"manual", "RR-03":"manual", "RR-04":"manual", "RR-05":"manual", "RR-06":"geospatial",
  "SI-01":"geospatial", "SI-02":"manual", "SI-03":"manual",
  "PG-01":"managerial", "PG-02":"managerial", "PG-03":"managerial", "PG-04":"managerial", "PG-05":"managerial"
};

const ACQUISITION_LABELS = {
  geospatial: {title:"Indicadores geoespaciales", note:"El HUB calcula estos indicadores automáticamente. Si requiere continuar sin esperar, puede registrar una valoración manual con justificación; ambas fuentes quedan trazadas."},
  official: {title:"Automáticos / fuentes institucionales", note:"Se obtienen del HUB o de fuentes oficiales cuando el conector esté disponible."},
  manual: {title:"Datos ingresados por el usuario", note:"Información institucional que debe registrar o validar la persona funcionaria."},
  managerial: {title:"Priorización gerencial", note:"Criterios gerenciales registrados por la persona funcionaria con su soporte correspondiente."}
};

function esc(v){return String(v??"").replace(/[&<>"']/g,m=>({"&":"&amp;","<":"&lt;",">":"&gt;","\"":"&quot;","'":"&#39;"}[m]));}
function hasScore(v){return v!==null&&v!==undefined&&v!==""&&Number.isFinite(Number(v));}
function fmt(v){return hasScore(v)?Number(v).toFixed(1):"—";}
function scoreLabel(v){return hasScore(v)?Number(v).toFixed(1):"Pendiente";}
function finalScoreFallback(r){
  if(hasScore(r?.finalScore))return Number(r.finalScore);
  if(hasScore(r?.technicalScore)&&hasScore(r?.managerialScore))return Number(((Number(r.technicalScore)+Number(r.managerialScore))/2).toFixed(1));
  return null;
}
function classificationFromScore(score){
  if(!hasScore(score))return null;
  const n=Number(score);
  return n>=85?"Muy Alta":n>=70?"Alta":n>=55?"Media":n>=40?"Baja":"Muy baja";
}
function normalizedRankingRow(r){
  const finalScore=finalScoreFallback(r);
  return {...r,finalScore,classification:r?.classification||classificationFromScore(finalScore)};
}
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
function parseHubDate(v){
 if(!v)return null;
 if(v instanceof Date&&!isNaN(v))return v;
 if(typeof v==="number"){const d=new Date(v);return isNaN(d)?null:d;}
 const t=String(v).trim();if(!t)return null;
 let d=new Date(t);if(!isNaN(d))return d;
 const m=t.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})(?:\s+.*)?$/);
 if(m){d=new Date(Number(m[3]),Number(m[2])-1,Number(m[1]));return isNaN(d)?null:d;}
 return null;
}
function projectEntryDate(p, detail=null){
 const ex=p?.excelData||{};
 const candidates=[
   p?.registeredAt,p?.createdAt,p?.created_at,p?.registrationDate,p?.entryDate,p?.fechaIngreso,
   p?.workflow?.registeredAt,p?.governance?.registeredAt,p?.governance?.createdAt,
   ex["Fecha de ingreso"],ex["FECHA DE INGRESO"],ex["Fecha ingreso"],ex["Fecha de registro"],ex["FECHA DE REGISTRO"],
   detail?.registeredAt,detail?.createdAt,detail?.project?.registeredAt,detail?.project?.createdAt
 ];
 for(const v of candidates){const d=parseHubDate(v);if(d)return d;}
 return null;
}
function formatEntryDate(v){const d=v instanceof Date?v:parseHubDate(v);return d?d.toLocaleDateString("es-CR",{year:"numeric",month:"2-digit",day:"2-digit"}):"—";}
function detailForDecision(projectId){return DECISION_DETAIL_CACHE.get(String(projectId))||null;}
function workflowForDecision(projectId){return DECISION_WORKFLOW_CACHE.get(String(projectId))||null;}
function decisionProjectFor(projectId){return projects.find(p=>String(projectIdOfDecision(p))===String(projectId))||projects.find(p=>String(projectId(p))===String(projectId))||null;}
function projectIdOfDecision(p){try{return projectId(p)}catch(e){return p?.idHub||p?.id||p?.projectId||""}}
function decisionEvidence(projectId){
 const id=String(projectId||"");
 const wf=workflowForDecision(id)||{};
 const p=decisionProjectFor(id)||{};
 const d=detailForDecision(id)||{};
 const explicitSelected=
   IMPLEMENTATION_SELECTED_IDS.has(id)||
   p?.selectedForImplementation===true||p?.processII?.active===true||p?.implementationPrioritization?.active===true||
   d?.selectedForImplementation===true||d?.processII?.active===true||d?.implementation?.active===true||d?.activeForImplementation===true;
 const text=[wf?.state,wf?.status,wf?.decisionStatus,p?.state,p?.status,p?.decisionStatus,p?.cycleStatus,d?.state,d?.status,d?.decisionStatus,d?.cycleStatus]
   .filter(v=>v!==undefined&&v!==null).join(" | ");
 return {explicitSelected,text};
}
function decisionManaged(projectId){const ev=decisionEvidence(projectId);return ev.explicitSelected||/seleccionado para implementaci|no seleccionado para implementaci|decisión registrada|proceso cerrado|priorizado/i.test(ev.text);}
function decisionStatusLabel(projectId){const ev=decisionEvidence(projectId);if(/no seleccionado para implementaci/i.test(ev.text))return "No seleccionado";if(ev.explicitSelected||/seleccionado para implementaci|priorizado/i.test(ev.text))return "Priorizado";if(/decisión registrada|proceso cerrado/i.test(ev.text))return "Gestionado";return "";}
function ensureDecisionControls(){
 if(PRIORITIZATION_SCOPE!=="decision"||document.getElementById("decisionDateControls"))return;
 const matrix=document.querySelector(".priority-matrix")||[...document.querySelectorAll("section,div")].find(el=>/Victoria Temprana/i.test(el.textContent||"")&&/Mina de Oro/i.test(el.textContent||""));
 if(!matrix)return;
 const panel=document.createElement("div");panel.id="decisionDateControls";
 panel.style.cssText="display:grid;grid-template-columns:repeat(4,minmax(160px,1fr));gap:12px;align-items:end;margin:0 0 18px;padding:14px 16px;border:1px solid #dbe4ef;border-radius:14px;background:#f8fbff";
 panel.innerHTML=`
   <label><b>Fecha de ingreso desde</b><input id="decisionDateFrom" type="date" style="width:100%;margin-top:6px"></label>
   <label><b>Fecha de ingreso hasta</b><input id="decisionDateTo" type="date" style="width:100%;margin-top:6px"></label>
   <label><b>Ordenar por</b><select id="decisionSort" style="width:100%;margin-top:6px"><option value="ranking">Ranking / puntaje final</option><option value="dateDesc">Fecha de ingreso · más reciente</option><option value="dateAsc">Fecha de ingreso · más antigua</option><option value="scoreDesc">Puntaje final · mayor a menor</option><option value="scoreAsc">Puntaje final · menor a mayor</option></select></label>
   <button id="clearDecisionDates" type="button" class="secondary">Limpiar fechas</button>`;
 matrix.parentElement.insertBefore(panel,matrix);
 const from=panel.querySelector("#decisionDateFrom"),to=panel.querySelector("#decisionDateTo"),sort=panel.querySelector("#decisionSort");
 from.value=decisionFilters.dateFrom;to.value=decisionFilters.dateTo;sort.value=decisionFilters.sort;
 from.onchange=()=>{decisionFilters.dateFrom=from.value;render();};
 to.onchange=()=>{decisionFilters.dateTo=to.value;render();};
 sort.onchange=()=>{decisionFilters.sort=sort.value;render();};
 panel.querySelector("#clearDecisionDates").onclick=()=>{decisionFilters.dateFrom="";decisionFilters.dateTo="";from.value="";to.value="";render();};
 if(!document.getElementById("decision-controls-responsive")){const st=document.createElement("style");st.id="decision-controls-responsive";st.textContent='@media(max-width:900px){#decisionDateControls{grid-template-columns:1fr 1fr!important}}@media(max-width:600px){#decisionDateControls{grid-template-columns:1fr!important}}';document.head.appendChild(st);}
}
function ensureDecisionTableHeader(){
 if(PRIORITIZATION_SCOPE!=="decision")return;
 const table=document.getElementById("rankingBody")?.closest("table");const tr=table?.querySelector("thead tr");if(!tr)return;
 if(!tr.querySelector('[data-col="entry-date"]')){
   const th=document.createElement("th");th.dataset.col="entry-date";th.textContent="FECHA DE INGRESO";
   const projectTh=tr.children[2];if(projectTh?.nextSibling)tr.insertBefore(th,projectTh.nextSibling);else tr.appendChild(th);
 }
 const last=tr.lastElementChild;if(last&&!String(last.textContent||"").trim())last.textContent="GESTIÓN";
}
function manageDecisionProject(id){
 try{parent.postMessage({type:"hub-manage-project",projectId:String(id||"")},"*");}catch(e){console.warn(e);}
}
function preliminaryMatrixFromName(p){
 const text=String([p?.["Tipo de Obra"],p?.["Sub tipo de obra"],projectName(p)].filter(Boolean).join(" ")).toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g,"");
 if(/puente.*peaton|peatonal.*puente|pasarela.*peaton|paso.*peaton/.test(text))return "M-04";
 if(/talud|ladera|estabiliz|proteccion.*ladera|deslizamiento|derrumbe/.test(text))return "M-05";
 if(/puente|viaduct/.test(text))return "M-03";
 if(/rehabil|mejor|ampli|conserv|manten|paviment|recarpet|reconstru|reforz|repar|ruta nacional|red vial/.test(text))return "M-01";
 if(/nuev[oa]|construccion|construir|apertura|segment|tramo|carreter|via|vial|radial|circunval|bypass|by pass|intercambio|eje ciclista|ciclovia|ciclista|circuito recreativo/.test(text))return "M-02";
 return "M-01";
}
async function loadAll(){
 const [c,p,r]=await Promise.all([api("/prioritization/config"),api("/projects"),api("/prioritization/ranking")]);
 config=c;projects=p.projects||[];ranking=(r.ranking||[]).map(normalizedRankingRow);
 if(PRIORITIZATION_SCOPE==="decision"){
   projects=projects.filter(x=>DECISION_IDS.has(String(projectId(x))));
   ranking=ranking.filter(x=>DECISION_IDS.has(String(x.projectId)));
   DECISION_DETAIL_CACHE.clear();
   DECISION_WORKFLOW_CACHE.clear();
   const results=await Promise.allSettled(projects.map(async p=>{
     const id=projectId(p);const [x,w]=await Promise.all([api(`/projects/${encodeURIComponent(id)}/prioritization`),api(`/projects/${encodeURIComponent(id)}/workflow`).catch(()=>null)]);return [id,x?.prioritization||x||{},w?.workflow||null];
   }));
   results.forEach(x=>{if(x.status==="fulfilled"){DECISION_DETAIL_CACHE.set(String(x.value[0]),x.value[1]);if(x.value[2])DECISION_WORKFLOW_CACHE.set(String(x.value[0]),x.value[2]);}});
 }
 populateFilters();ensureDecisionControls();ensureDecisionTableHeader();render();
 const agg=c.finalAggregation||{};
 document.getElementById("methodologyNotice").innerHTML=`<b>Metodología ${esc(c.version)}.</b> Las matrices M-01 a M-05 están parametrizadas. Cuando el Puntaje Técnico y el Puntaje Gerencial están completos, el HUB calcula automáticamente el Puntaje Final. Mientras SPS no configure otra ponderación, se aplica una integración 50/50.`;
}
function populateFilters(){
 const mf=document.getElementById("matrixFilter");const selected=mf.value;mf.innerHTML='<option value="">Todas</option>'+Object.entries(config.matrices||{}).map(([id,m])=>`<option value="${id}">${id} · ${esc(m.typology)}</option>`).join("");mf.value=selected;
 const inf=document.getElementById("institutionFilter");const si=inf.value;const inst=[...new Set(projects.map(projectInst).filter(Boolean))].sort();inf.innerHTML='<option value="">Todas</option>'+inst.map(x=>`<option>${esc(x)}</option>`).join("");inf.value=si;
}
function mergedRows(){
 const byId=new Map(ranking.map(r=>[String(r.projectId),r]));
 return projects.map(p=>{const id=projectId(p),r=normalizedRankingRow(byId.get(id)||{}),detail=detailForDecision(id);return {...r,projectId:id,_sourceProject:p,_decisionDetail:detail,_entryDate:projectEntryDate(p,detail),project:r.project||{idHub:id,name:projectName(p),institution:projectInst(p),stage:p["Etapa"]||""}};});
}
function filteredRows(){
 const q=document.getElementById("q").value.trim().toLowerCase(),m=document.getElementById("matrixFilter").value,inst=document.getElementById("institutionFilter").value,cl=document.getElementById("classFilter").value;
 let rows=mergedRows().filter(r=>{
   const txt=`${r.projectId} ${r.project?.name||""} ${r.project?.institution||""}`.toLowerCase();
   if(q&&!txt.includes(q))return false;if(m&&r.matrixId!==m)return false;if(inst&&r.project?.institution!==inst)return false;if(cl&&r.classification!==cl)return false;
   if(PRIORITIZATION_SCOPE==="decision"){
     const d=r._entryDate;
     if(decisionFilters.dateFrom){const f=parseHubDate(decisionFilters.dateFrom+"T00:00:00");if(!d||d<f)return false;}
     if(decisionFilters.dateTo){const t=parseHubDate(decisionFilters.dateTo+"T23:59:59");if(!d||d>t)return false;}
   }
   return true;
 });
 const defaultRanking=(a,b)=>{const af=hasScore(a.finalScore)?Number(a.finalScore):null,bf=hasScore(b.finalScore)?Number(b.finalScore):null;if(af!==null||bf!==null)return(bf??-1)-(af??-1);const bt=hasScore(b.technicalScore)?Number(b.technicalScore):-1,at=hasScore(a.technicalScore)?Number(a.technicalScore):-1;return bt-at;};
 if(PRIORITIZATION_SCOPE==="decision"){
   if(decisionFilters.sort==="dateDesc")rows.sort((a,b)=>(b._entryDate?.getTime()||0)-(a._entryDate?.getTime()||0));
   else if(decisionFilters.sort==="dateAsc")rows.sort((a,b)=>(a._entryDate?.getTime()||8640000000000000)-(b._entryDate?.getTime()||8640000000000000));
   else if(decisionFilters.sort==="scoreAsc")rows.sort((a,b)=>(hasScore(a.finalScore)?Number(a.finalScore):9999)-(hasScore(b.finalScore)?Number(b.finalScore):9999));
   else rows.sort(defaultRanking);
 }else rows.sort(defaultRanking);
 return rows;
}
function indicatorScoreFromRow(r,id){
 const d=r?._decisionDetail||detailForDecision(r?.projectId);const candidates=[r?.inputs?.scores?.[id],r?.scores?.indicatorScores?.[id],r?.indicatorScores?.[id],r?.inputs?.raw?.[id]?.score,d?.inputs?.scores?.[id],d?.scores?.indicatorScores?.[id],d?.indicatorScores?.[id],d?.inputs?.raw?.[id]?.score,d?.spatialOverrides?.[id]?.score,d?.spatial?.suggestedScores?.[id]?.score];
 for(const v of candidates){if(v!==undefined&&v!==null&&v!==""&&Number.isFinite(Number(v)))return Number(v);}return null;
}
function renderDecisionMatrix(rows){
 const host=document.querySelector(".priority-matrix")||[...document.querySelectorAll("section,div")].find(el=>/Victoria Temprana/i.test(el.textContent||"")&&/Mina de Oro/i.test(el.textContent||""));if(!host)return;
 host.querySelectorAll(".decision-point").forEach(n=>n.remove());
 const scored=rows.filter(r=>hasScore(r.finalScore)||hasScore(r.technicalScore));
 scored.forEach((r,index)=>{
   const score=hasScore(r.finalScore)?Number(r.finalScore):Number(r.technicalScore);
   const vi=["VI-01","VI-02","VI-03"].map(id=>indicatorScoreFromRow(r,id)).filter(Number.isFinite);
   const ease=vi.length?vi.reduce((a,b)=>a+b,0)/vi.length:50;
   const effort=100-ease;
   const b=document.createElement("button");b.type="button";b.className="decision-point";
   b.style.left=`${Math.max(4,Math.min(96,effort))}%`;b.style.top=`${Math.max(5,Math.min(95,100-score))}%`;
   b.textContent=r.projectId;
   b.title=`${r.projectId} · ${r.project?.name||"Proyecto"}\nPuntaje final: ${score.toFixed(1)}\nEsfuerzo relativo: ${effort.toFixed(1)}${vi.length<3?"\nEsfuerzo provisional: faltan indicadores VI.":""}`;
   b.onclick=()=>{if(decisionManaged(r.projectId)){alert(`La decisión de ${r.projectId} ya fue registrada (${decisionStatusLabel(r.projectId)||"Gestionado"}).`);return;}manageDecisionProject(r.projectId);};host.appendChild(b);
 });
 if(!document.getElementById("decision-matrix-style")){
   const st=document.createElement("style");st.id="decision-matrix-style";st.textContent=`.priority-matrix{position:relative!important;overflow:visible!important}.decision-point{position:absolute!important;transform:translate(-50%,-50%);min-width:56px;height:32px;padding:0 9px;border-radius:999px;border:3px solid #fff;background:#0d4d8b;color:#fff;font-size:11px;font-weight:800;box-shadow:0 3px 10px rgba(0,0,0,.35);z-index:30;cursor:pointer;white-space:nowrap}.decision-point:hover{transform:translate(-50%,-50%) scale(1.08);z-index:35}`;document.head.appendChild(st);
 }
}
function render(){
 ensureDecisionControls();ensureDecisionTableHeader();
 const rows=filteredRows();document.getElementById("kProjects").textContent=projects.length;document.getElementById("kEvaluated").textContent=ranking.length;document.getElementById("kTechnical").textContent=ranking.filter(x=>hasScore(x.technicalScore)).length;document.getElementById("kFinal").textContent=ranking.filter(x=>hasScore(x.finalScore)).length;
 document.getElementById("rankingBody").innerHTML=rows.map((r,i)=>{const p=projects.find(x=>projectId(x)===r.projectId);const mid=r.matrixId||preliminaryMatrixFromName(p||r.project||{});const prelim=!r.matrixId||r.matrixPreliminary;const dateCell=PRIORITIZATION_SCOPE==="decision"?`<td>${esc(formatEntryDate(r._entryDate))}</td>`:"";const managed=PRIORITIZATION_SCOPE==="decision"&&decisionManaged(r.projectId);const statusLabel=managed?decisionStatusLabel(r.projectId):"";const actions=PRIORITIZATION_SCOPE==="decision"?`<div style="display:flex;gap:6px;flex-wrap:wrap;align-items:center"><button class="row-action" data-id="${esc(r.projectId)}">Ver puntajes</button>${managed?`<button class="row-manage" disabled style="opacity:.6;cursor:not-allowed">Gestionado</button><span class="muted">${esc(statusLabel)}</span>`:`<button class="row-manage" data-id="${esc(r.projectId)}">Gestionar</button>`}</div>`:`<button class="row-action" data-id="${esc(r.projectId)}">Abrir</button>`;return `<tr><td>${i+1}</td><td><b>${esc(r.projectId)}</b></td><td>${esc(r.project?.name||"Sin nombre")}<div class="muted">${esc(r.project?.institution||"")}</div></td>${dateCell}<td><span class="pill ${prelim?"warn":""}" title="${prelim?"Asignación preliminar por nombre; puede corregirse en la ficha.":"Matriz guardada"}">${esc(mid)}${prelim?" · preliminar":""}</span></td><td class="score">${fmt(r.technicalScore)}</td><td>${fmt(r.managerialScore)}</td><td class="score">${fmt(r.finalScore)}</td><td>${esc(r.classification||"—")}</td><td>${r.missingCount??"—"}</td><td>${actions}</td></tr>`;}).join("")||`<tr><td colspan="${PRIORITIZATION_SCOPE==="decision"?11:10}" class="muted">No hay proyectos que coincidan.</td></tr>`;
 document.querySelectorAll(".row-action").forEach(b=>b.onclick=()=>openProject(b.dataset.id));
 document.querySelectorAll(".row-manage").forEach(b=>b.onclick=()=>manageDecisionProject(b.dataset.id));
 if(PRIORITIZATION_SCOPE==="decision"){
   renderDecisionMatrix(rows);
   const notice=document.getElementById("methodologyNotice");if(notice)notice.innerHTML="<b>Vista del Tomador de decisión.</b> Solo se muestran las iniciativas remitidas a su bandeja. Los puntajes y criterios son de consulta y no pueden editarse.";
   const sub=document.getElementById("rankingSubtitle");if(sub)sub.textContent="Ranking de las iniciativas que han llegado al Tomador de decisión, ordenado por Puntaje Final (o Técnico cuando el Final aún no esté disponible).";
   const batch=document.getElementById("btnBatch");if(batch)batch.style.display="none";
 }
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
   if(PRIORITIZATION_SCOPE==="technical"&&mode==="managerial") continue;
   const list=groups[mode]; if(!list.length) continue;
   const meta=ACQUISITION_LABELS[mode];
   html.push(`<div class="indicator-group"><div class="indicator-group-head"><div><h4>${esc(meta.title)}</h4><p>${esc(meta.note)}</p></div><span class="mode-pill ${mode}">${mode==="geospatial"?"GIS":mode==="official"?"FUENTE":mode==="managerial"?"GERENCIAL":"MANUAL"}</span></div><div class="indicator-grid">`);
   for(const id of list){
     const value=hasScore(computed[id])?computed[id]:(hasScore(scores[id])?scores[id]:null);
     if(mode==="geospatial"){
       const o=data?.spatialOverrides?.[id]||null;
       const usingManual=o?.method==="manual";
       const autoScore=data?.spatial?.suggestedScores?.[id]?.score ?? null;
       const options=(config?.inputOptions?.[id]||[]);
       const optionHtml=options.map(x=>`<option value="${esc(x.score)}" data-label="${esc(x.label)}" ${usingManual&&Number(o.score)===Number(x.score)?"selected":""}>${esc(x.label)} · ${esc(x.score)} pts</option>`).join("");
       html.push(`<div class="indicator indicator-auto geo-flex" data-indicator="${id}" data-mode="geospatial">
         <div class="indicator-head"><b>${id}</b><span>Utilizado: ${fmt(value)}</span></div>
         <div>${esc(INDICATOR_NAMES[id]||id)}</div>
         <div class="auto-value"><strong>${fmt(autoScore)}</strong><span>${autoScore===null?"GIS pendiente / no disponible todavía":"Resultado GIS guardado"}</span></div>
         <div style="display:flex;gap:8px;flex-wrap:wrap;margin:10px 0">
           <button type="button" class="geo-use-auto" ${autoScore===null?"disabled":""}>Usar cálculo GIS</button>
           <button type="button" class="geo-use-manual">Ingresar manualmente</button>
         </div>
         <div class="geo-manual-panel" style="display:${usingManual?"block":"none"};padding:10px;border:1px solid #dbe4ef;border-radius:10px;background:#f8fbff">
           <label><b>Condición / categoría</b></label>
           <select class="geo-manual-choice"><option value="">Seleccione…</option>${optionHtml}</select>
           <label style="display:block;margin-top:8px"><b>Justificación obligatoria</b></label>
           <textarea class="geo-manual-justification" rows="2" style="width:100%" placeholder="Indique la fuente o razón de la valoración manual.">${esc(o?.justification||"")}</textarea>
           <small>El HUB asigna el puntaje de la metodología; no se escribe un puntaje libre.</small>
         </div>
         ${GEO_REASON[id]?`<small><b>¿Por qué geoespacial?</b> ${esc(GEO_REASON[id])}</small>`:""}
       </div>`);
     }else if(mode==="official"){
       html.push(`<div class="indicator indicator-auto" data-indicator="${id}" data-mode="${mode}"><div class="indicator-head"><b>${id}</b><span>Calculado: ${fmt(value)}</span></div><div>${esc(INDICATOR_NAMES[id]||id)}</div><div class="auto-value"><strong>${fmt(value)}</strong><span>Dato automático / institucional</span></div><small>${value===null?"Pendiente de fuente o de una corrida válida.":"Este valor no se edita manualmente."}</small></div>`);
     }else{
       const options=(config?.inputOptions?.[id]||[]);
       const savedScore=hasScore(scores[id])?Number(scores[id]):null;
       const optionHtml=options.map(o=>`<option value="${esc(o.score)}" data-label="${esc(o.label)}" ${savedScore===Number(o.score)?"selected":""}>${esc(o.label)} · ${esc(o.score)} pts</option>`).join("");
       html.push(`<div class="indicator" data-indicator="${id}" data-mode="${mode}"><div class="indicator-head"><b>${id}</b><span>Calculado: ${fmt(computed[id])}</span></div><div>${esc(INDICATOR_NAMES[id]||id)}</div><div class="indicator-inputs"><select class="manual-choice"><option value="">Seleccione un rango / categoría</option>${optionHtml}</select></div><small>Seleccione la condición que corresponda. El HUB asigna automáticamente el puntaje de la metodología.</small></div>`);
     }
   }
   html.push(`</div></div>`);
 }
 document.getElementById("indicatorEditor").innerHTML=html.join("");
 document.querySelectorAll(".indicator[data-mode='geospatial']").forEach(box=>{
   const panel=box.querySelector(".geo-manual-panel");
   const autoBtn=box.querySelector(".geo-use-auto");
   const manualBtn=box.querySelector(".geo-use-manual");
   if(manualBtn)manualBtn.onclick=()=>{ if(panel)panel.style.display="block"; box.dataset.geoMethod="manual"; };
   if(autoBtn)autoBtn.onclick=()=>{ if(panel)panel.style.display="none"; box.dataset.geoMethod="automatic"; };
   const existing=data?.spatialOverrides?.[box.dataset.indicator];
   box.dataset.geoMethod=existing?.method==="manual"?"manual":"automatic";
 });
}





const EVIDENCE_REASON = {
  asp: "Se usa para identificar si el proyecto intersecta o se encuentra próximo a áreas silvestres protegidas.",
  wetlands: "Se usa para identificar interacción territorial con humedales registrados.",
  corridors: "Se usa para identificar interacción con corredores biológicos.",
  landslides: "Se usa para identificar exposición del proyecto a zonas de deslizamiento.",
  crowns: "Se usa como evidencia complementaria de inestabilidad y amenaza por deslizamientos.",
  floods: "Se usa para identificar exposición a áreas con potencial de inundación.",
  roadHierarchy: "Se usa para determinar la jerarquía vial sobre la que se localiza el proyecto.",
  connectivity: "La criticidad de conectividad se deriva de la jerarquía de la red vial asociada al proyecto.",
  multiHazard: "La exposición multiamenaza combina las amenazas disponibles que intersectan o se ubican próximas al proyecto."
};

function evidenceLayerLabel(key){
  const labels={
    asp:"Áreas silvestres protegidas",
    wetlands:"Humedales",
    corridors:"Corredores biológicos",
    landslides:"Deslizamientos",
    crowns:"Coronas de deslizamiento",
    floods:"Áreas con potencial de inundación",
    roadHierarchy:"Jerarquía de la Red Vial Nacional",
    connectivity:"Criticidad de conectividad",
    multiHazard:"Exposición multiamenaza"
  };
  return labels[key] || String(key||"").replace(/[_-]+/g," ").replace(/\b\w/g,c=>c.toUpperCase());
}

function renderEvidence(spatial){
 const ev=spatial?.evidence||{};const cards=[];
 for(const [k,v] of Object.entries(ev)){
  if(!v||typeof v!=="object")continue;
  const m=v.metrics;
  const stale=v.refreshStatus==="stale";
  const fresh=v.refreshStatus==="fresh";
  const estado=stale?"guardado (fuente no disponible en esta corrida)":fresh?"actualizado":(v.status||"—");
  cards.push(`<div class="evidence-card"><b>${esc(evidenceLayerLabel(k))}</b><p>${EVIDENCE_REASON[k]?`${esc(EVIDENCE_REASON[k])}<br><br>`:""}Estado: ${esc(estado)}${v.sourceName?`<br>Fuente: ${esc(v.sourceName)}`:""}${v.layerName?`<br>Capa: ${esc(v.layerName)}`:(v.typeName?`<br>Capa: ${esc(v.typeName)}`:"")}${v.file?`<br>Archivo: ${esc(v.file)}`:""}${m?`<br>Intersecciones: ${m.intersectCount}<br>En ${m.nearMeters} m: ${m.nearCount}<br>Consultados en zona: ${m.totalFetched}`:""}${v.lastValidAt?`<br>Último dato válido: ${esc(new Date(v.lastValidAt).toLocaleString())}`:""}${v.refreshError?`<br>Actualización fallida: ${esc(v.refreshError)}`:""}${v.error&&!stale?`<br>Error: ${esc(v.error)}`:""}</p></div>`);
 }
 const sugg=spatial?.suggestedScores||{};
 for(const [id,s] of Object.entries(sugg))cards.push(`<div class="evidence-card"><b>Sugerencia ${id}: ${s.score}/100</b><p>${esc(s.note||"Derivada de relación espacial; requiere validación metodológica.")}</p></div>`);
 const summary=spatial?.summary;
 if(summary)cards.unshift(`<div class="evidence-card"><b>Estado de fuentes</b><p>Actualizadas: ${summary.fresh??summary.ok??0}<br>Desde APOYO_GEO: ${summary.localOk??0}<br>Con dato guardado: ${summary.stale??0}<br>No disponibles: ${summary.unavailable??summary.failed??0}</p></div>`);
 document.getElementById("spatialEvidence").innerHTML=cards.join("")||'<div class="muted">Todavía no hay evidencia espacial calculada.</div>';
}
async function openProject(id){
 activeId=id;const p=projects.find(x=>projectId(x)===id);const r=await api(`/projects/${encodeURIComponent(id)}/prioritization`);const d=r.prioritization||{projectId:id,project:{name:projectName(p),institution:projectInst(p)},inputs:{raw:{},scores:{}}};
 document.getElementById("dialogTitle").textContent=`${id} · ${d.project?.name||projectName(p)}`;document.getElementById("dialogMeta").textContent=`${d.project?.institution||projectInst(p)} · ${d.matrix?.matrixId||"Matriz sin asignar"}`;
 const displayFinal=hasScore(d.scores?.finalScore)?Number(d.scores.finalScore):(hasScore(d.scores?.technicalScore)&&hasScore(d.scores?.managerialScore)?Number(((Number(d.scores.technicalScore)+Number(d.scores.managerialScore))/2).toFixed(1)):null);
 document.getElementById("dTechnical").textContent=scoreLabel(d.scores?.technicalScore);document.getElementById("dManagerial").textContent=scoreLabel(d.scores?.managerialScore);document.getElementById("dFinal").textContent=scoreLabel(displayFinal);
 const ms=document.getElementById("matrixOverride");const effectiveMid=d.matrix?.matrixId||preliminaryMatrixFromName(p||d.project||{});ms.innerHTML=`<option value="">Asignación automática · preliminar ${esc(effectiveMid)}</option>`+Object.entries(config.matrices||{}).map(([mid,m])=>`<option value="${mid}">${mid} · ${esc(m.typology)}</option>`).join("");ms.value=d.matrixOverride||"";
 document.getElementById("technicalWeight").value=d.finalAggregation?.technicalWeight??50;document.getElementById("managerialWeight").value=d.finalAggregation?.managerialWeight??50;
 renderEditor(d);renderEvidence(d.spatial);
 const lastValid=d.calculatedAt?new Date(d.calculatedAt).toLocaleString():null;
 const lastAttempt=d.lastAttempt?.calculatedAt?new Date(d.lastAttempt.calculatedAt).toLocaleString():null;
 const attemptValid=d.lastAttempt?.valid;
 let statusHtml; if(!d.matrix?.matrixId){statusHtml="<b>Evaluación pendiente.</b> Primero debe asignarse una matriz metodológica; después el HUB mostrará los indicadores aplicables.";}else if(d.scores?.missing?.length){statusHtml=`Faltan <b>${d.scores.missing.length}</b> indicadores para completar el cálculo: ${d.scores.missing.join(", ")}.`;}else{statusHtml=lastValid?`Último cálculo válido: <b>${lastValid}</b>.`:"Este proyecto todavía no ha sido calculado.";}
 if(PRIORITIZATION_SCOPE==="technical"&&d.matrix?.matrixId){
   const techMissing=(d.scores?.missing||[]).filter(x=>!String(x).startsWith("PG-"));
   statusHtml=techMissing.length?`Faltan <b>${techMissing.length}</b> variables técnicas para completar la tarea del Técnico evaluador: ${techMissing.join(", ")}.`:`<b>Evaluación técnica completa.</b> La Priorización gerencial (PG-01 a PG-05) será diligenciada posteriormente por el Jerarca Institucional.`;
 }
 if(lastAttempt && attemptValid===false && lastValid)statusHtml+=`<br>La actualización de <b>${lastAttempt}</b> no reemplazó el resultado guardado.`;
 if(d.lastAttempt?.spatialCacheHit)statusHtml+=`<br><span style="color:#15803d">✓ Cruces geoespaciales reutilizados porque la geometría no cambió.</span>`;
 if(d.spatial?.checkedAt)statusHtml+=`<br><small>GIS: ${new Date(d.spatial.checkedAt).toLocaleString()}</small>`;
 document.getElementById("dialogStatus").innerHTML=statusHtml;
 const readonly=PRIORITIZATION_VIEW_MODE==="readonly";
 document.querySelectorAll("#projectDialog input,#projectDialog select").forEach(el=>{el.disabled=readonly;});
 ["btnSave","btnCalculate"].forEach(id=>{const el=document.getElementById(id);if(el)el.style.display=readonly?"none":"";});
 const dlg=document.getElementById("projectDialog");if(!dlg.open)dlg.showModal();
}
function collectInput(){
 // SEGURIDAD: el guardado es incremental. Una pantalla parcial nunca borra valores ya persistidos.
 const raw={},scores={},spatialOverrides={};
 document.querySelectorAll(".indicator").forEach(box=>{
   const id=box.dataset.indicator;
   if(box.dataset.mode==="geospatial"){
     if(box.dataset.geoMethod==="manual"){
       const choice=box.querySelector(".geo-manual-choice");
       const justification=(box.querySelector(".geo-manual-justification")?.value||"").trim();
       if(choice?.value!==""){
         if(!justification)throw new Error(`${id}: debe escribir una justificación para usar el valor manual.`);
         spatialOverrides[id]={method:"manual",score:Number(choice.value),label:choice.selectedOptions?.[0]?.dataset?.label||choice.selectedOptions?.[0]?.textContent||"",justification,updatedAt:new Date().toISOString(),updatedBy:(window.HUB_CURRENT_USER?.email||window.HUB_CURRENT_USER?.displayName||"")};
       }
     }
     return;
   }
   const choice=box.querySelector(".manual-choice");
   if(!choice)return;
   const value=choice.value;
   if(value!==""){
     scores[id]=Number(value);
     raw[id]=choice.selectedOptions?.[0]?.dataset?.label||choice.selectedOptions?.[0]?.textContent||"";
   }
 });
 const tw=document.getElementById("technicalWeight").value.trim(),gw=document.getElementById("managerialWeight").value.trim();
 return{matrixOverride:document.getElementById("matrixOverride").value,raw,scores,spatialOverrides,replaceSpatialOverrides:false,replaceManualInputs:false,finalAggregation:{technicalWeight:tw===""?null:Number(tw),managerialWeight:gw===""?null:Number(gw)}};
}
function ensureBusyOverlay(){
 const dlg=document.getElementById("projectDialog");
 if(!dlg)return null;
 let el=dlg.querySelector("#prioritizationBusyOverlay");
 if(el)return el;
 el=document.createElement("div");
 el.id="prioritizationBusyOverlay";
 Object.assign(el.style,{position:"fixed",inset:"0",zIndex:"2147483647",display:"none",alignItems:"center",justifyContent:"center",background:"rgba(8,38,72,.34)",backdropFilter:"blur(1.5px)"});
 el.innerHTML='<div style="min-width:280px;max-width:460px;margin:20px;padding:22px 26px;border-radius:16px;background:#fff;color:#0b3765;box-shadow:0 20px 60px rgba(0,0,0,.28);font:600 18px/1.35 system-ui,-apple-system,Segoe UI,sans-serif;text-align:center"><div id="prioritizationBusyTitle">Guardando…</div><div id="prioritizationBusyDetail" style="margin-top:7px;font-size:13px;font-weight:500;color:#64748b">Espere un momento.</div></div>';
 dlg.appendChild(el);return el;
}
function showBusyOverlay(title,detail="Espere un momento."){const el=ensureBusyOverlay();if(!el)return;el.querySelector("#prioritizationBusyTitle").textContent=title;el.querySelector("#prioritizationBusyDetail").textContent=detail;el.style.display="flex";}
function hideBusyOverlay(){const el=document.getElementById("prioritizationBusyOverlay");if(el)el.style.display="none";}
async function showBusySuccess(title="Guardado correctamente"){const el=ensureBusyOverlay();if(!el)return;el.querySelector("#prioritizationBusyTitle").textContent=`✓ ${title}`;el.querySelector("#prioritizationBusyDetail").textContent="Los datos quedaron persistidos en el HUB.";await new Promise(r=>setTimeout(r,650));hideBusyOverlay();}

async function saveInput(){
 if(!activeId)return;
 const btn=document.getElementById("btnSave");if(btn)btn.disabled=true;
 showBusyOverlay("Guardando datos…","Registrando la selección y la matriz del proyecto en OneDrive.");
 try{
   const saved=await api(`/projects/${encodeURIComponent(activeId)}/prioritization`,{method:"PUT",body:JSON.stringify(collectInput())},true);
   if(!saved?.ok||saved?.persisted===false)throw new Error("El HUB no confirmó la persistencia de los datos.");
   const refreshed=await api(`/projects/${encodeURIComponent(activeId)}/prioritization`);
   notifyParentPrioritizationUpdated(activeId,refreshed?.prioritization||saved?.prioritization||null);
   await openProject(activeId);
   await showBusySuccess("Datos guardados");
   toast("Datos guardados.");
 }catch(e){hideBusyOverlay();throw e;}finally{if(btn)btn.disabled=false;}
}
async function calculateActive(){if(!activeId)return;await api(`/projects/${encodeURIComponent(activeId)}/prioritization`,{method:"PUT",body:JSON.stringify(collectInput())},true);const r=await api(`/projects/${encodeURIComponent(activeId)}/prioritization/calculate`,{method:"POST",body:JSON.stringify({forceSpatial:false})},true);toast("Priorización recalculada.");await loadAll();await openProject(activeId);return r;}
async function recalcBatch(){
 let offset=0,total=null;document.getElementById("btnBatch").disabled=true;
 try{do{const r=await api("/prioritization/recalculate",{method:"POST",body:JSON.stringify({offset,limit:10})},true);total=r.total;offset=r.nextOffset;toast(`Procesados ${Math.min(offset??total,total)} de ${total}`);}while(offset!==null);await loadAll();toast("Portafolio recalculado.");}finally{document.getElementById("btnBatch").disabled=false;}
}
function exportCsv(){const rows=filteredRows(),head=PRIORITIZATION_SCOPE==="decision"?["Ranking","ID HUB","Proyecto","Institución","Fecha de ingreso","Matriz","Puntaje Técnico","Puntaje Gerencial","Puntaje Final","Clasificación","Faltantes"]:["Ranking","ID HUB","Proyecto","Institución","Matriz","Puntaje Técnico","Puntaje Gerencial","Puntaje Final","Clasificación","Faltantes"];const q=v=>`"${String(v??"").replace(/"/g,'""')}"`;const body=[head,...rows.map((r,i)=>PRIORITIZATION_SCOPE==="decision"?[i+1,r.projectId,r.project?.name,r.project?.institution,formatEntryDate(r._entryDate),r.matrixId,r.technicalScore,r.managerialScore,r.finalScore,r.classification,r.missingCount]:[i+1,r.projectId,r.project?.name,r.project?.institution,r.matrixId,r.technicalScore,r.managerialScore,r.finalScore,r.classification,r.missingCount])].map(a=>a.map(q).join(",")).join("\n");const blob=new Blob(["\ufeff"+body],{type:"text/csv;charset=utf-8"});const a=document.createElement("a");a.href=URL.createObjectURL(blob);a.download="ranking_priorizacion_hub.csv";a.click();URL.revokeObjectURL(a.href);}
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

  const projectDialog = document.getElementById("projectDialog");
  if (projectDialog && !projectDialog.dataset.backdropClose) {
    projectDialog.dataset.backdropClose = "1";
    projectDialog.addEventListener("click", event => {
      if (event.target === projectDialog) {
        projectDialog.close();
        activeId = null;
      }
    });
  }

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

      if (status) status.innerHTML = "<b>Calculando…</b> Guardando configuración, consultando geometría y ejecutando análisis geoespacial.";
      showBusyOverlay("Calculando priorización…","Guardando datos y ejecutando los cruces geoespaciales.");
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
          { method: "POST", body: JSON.stringify({ forceSpatial:false }) },
          true
        );

        console.log("5. RESULTADO CÁLCULO:", result);
        notifyParentPrioritizationUpdated(activeId,result?.prioritization||null);

        if (status) status.innerHTML = "<b>Cálculo completado.</b> Actualizando resultados…";
        await loadAll();
        await openProject(activeId);
        await showBusySuccess("Cálculo guardado");
        toast("Priorización recalculada.");
      } catch (e) {
        hideBusyOverlay();
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


// ============================================================
// PROCESO II · GESTIÓN Y PRIORIZACIÓN DE PROYECTOS
// ============================================================
let implementationConfig=null, implementationRows=[], implementationProjects=[], implementationActiveId=null;
function implEsc(v){return esc(v)}
function implMoney(v){const n=Number(v);return Number.isFinite(n)?`${n.toLocaleString("es-CR",{maximumFractionDigits:1})} M USD`:"—";}
function implementationCss(){return `<style>
body{margin:0;background:#f5f8fc;color:#0f2f58;font-family:Inter,system-ui,-apple-system,Segoe UI,sans-serif}.impl-wrap{padding:24px;max-width:1500px;margin:auto}.impl-hero{background:linear-gradient(120deg,#0c4a7d,#0b7b85);color:white;padding:24px 28px;border-radius:20px;margin-bottom:18px}.impl-card{background:#fff;border:1px solid #dbe5f0;border-radius:18px;padding:20px;margin-bottom:18px}.impl-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(250px,1fr));gap:14px}.impl-kpi{background:#f6f9fc;border:1px solid #dce6ef;border-radius:14px;padding:14px}.impl-table{width:100%;border-collapse:collapse}.impl-table th,.impl-table td{padding:11px;border-bottom:1px solid #e5edf5;text-align:left}.impl-table th{font-size:12px;letter-spacing:.05em;color:#64748b}.impl-btn{border:0;border-radius:10px;background:#084a86;color:white;padding:10px 14px;font-weight:700;cursor:pointer}.impl-btn.secondary{background:#64748b}.impl-btn:disabled{opacity:.45}.impl-filter{display:flex;gap:10px;flex-wrap:wrap}.impl-filter input,.impl-filter select,.impl-field input,.impl-field select{padding:10px 12px;border:1px solid #cddbeb;border-radius:10px;background:#fff;min-width:180px}.impl-dlg{width:min(1080px,94vw);max-height:90vh;border:0;border-radius:18px;padding:0}.impl-dlg::backdrop{background:rgba(15,34,56,.58)}.impl-dlg-head{position:sticky;top:0;background:#f8fbff;padding:18px 22px;border-bottom:1px solid #dae5ef;display:flex;justify-content:space-between;z-index:2}.impl-dlg-body{padding:20px}.impl-dim{margin:18px 0}.impl-dim h3{margin:0 0 10px}.impl-field{display:grid;grid-template-columns:120px 1fr 220px 90px;gap:12px;align-items:center;border:1px solid #e1eaf3;border-radius:12px;padding:12px;margin:8px 0}.impl-score{font-weight:800;text-align:right}.impl-note{font-size:12px;color:#64748b}.impl-status{padding:12px;border-left:4px solid #0c5b9e;background:#eef6ff;border-radius:10px}.impl-bad{border-color:#c53030;background:#fff1f1}.impl-matrix{position:relative;height:390px;border:1px solid #d9e4ee;border-radius:16px;background:linear-gradient(90deg,#eef9ef 0 50%,#f8f0fd 50% 100%);overflow:hidden}.impl-matrix:before{content:"";position:absolute;left:50%;top:0;bottom:0;border-left:1px dashed #7b94ac}.impl-matrix:after{content:"";position:absolute;left:0;right:0;top:50%;border-top:1px dashed #7b94ac}.impl-point{position:absolute;transform:translate(-50%,-50%);background:#0a4f8b;color:#fff;border-radius:999px;padding:7px 9px;font-size:11px;font-weight:800;box-shadow:0 4px 12px #0003}.impl-q{position:absolute;font-weight:800}.q1{top:12px;left:14px;color:#16823b}.q2{top:12px;right:14px;color:#7c3eb0}.q3{bottom:12px;left:14px;color:#b77900}.q4{bottom:12px;right:14px;color:#b62b2b}@media(max-width:760px){.impl-field{grid-template-columns:1fr}.impl-table{font-size:13px}}
</style>`}
async function loadImplementationModule(){
 document.body.innerHTML=implementationCss()+`<div class="impl-wrap"><div class="impl-hero"><h1 style="margin:0 0 8px">Priorización de proyectos</h1><div>Proceso II · Evaluación para programación de inversiones mediante Puntaje de Implementación.</div></div><div class="impl-card"><div class="impl-filter"><input id="implQ" placeholder="Buscar ID, proyecto o institución"><select id="implClass"><option value="">Todas las clasificaciones</option><option>Muy Alta</option><option>Alta</option><option>Media</option><option>Baja</option><option>Muy baja</option></select><select id="implSort"><option value="rank">Ordenar por ranking</option><option value="score">Puntaje mayor a menor</option><option value="capexAsc">CAPEX menor a mayor</option><option value="capexDesc">CAPEX mayor a menor</option></select></div></div><div id="implSummary" class="impl-grid"></div><div class="impl-card"><h2>Matriz Implementación – Inversión</h2><div class="impl-note">Eje Y: Puntaje de Implementación · Eje X: CAPEX vigente.</div><div id="implMatrix" class="impl-matrix"><span class="impl-q q1">Victoria Temprana</span><span class="impl-q q2">Mina de Oro</span><span class="impl-q q3">Baja Prioridad</span><span class="impl-q q4">Cuestionable</span></div></div><div class="impl-card"><h2>Ranking de proyectos</h2><div id="implTable"></div></div><dialog id="implDialog" class="impl-dlg"></dialog></div>`;
 implementationConfig=await api('/project-prioritization/config');
 if(IMPLEMENTATION_SELECTED_IDS.size){
   await Promise.allSettled([...IMPLEMENTATION_SELECTED_IDS].map(id=>api(`/projects/${encodeURIComponent(id)}/implementation-prioritization`,{method:'POST',body:JSON.stringify({selectedBy:'Tomador de decisión',updatedBy:'Autorreparación Proceso II'})},true)));
 }
 const [pr,rr]=await Promise.all([api('/projects'),api('/project-prioritization/ranking')]);implementationProjects=pr.projects||[];implementationRows=rr.ranking||[];
 ['implQ','implClass','implSort'].forEach(id=>document.getElementById(id).addEventListener('input',renderImplementation));renderImplementation();
 const requested=PRIORITIZATION_PARAMS.get('project');if(requested)setTimeout(()=>openImplementationProject(requested),150);
}
function implementationFiltered(){const q=(document.getElementById('implQ')?.value||'').toLowerCase(),cl=document.getElementById('implClass')?.value||'',sort=document.getElementById('implSort')?.value||'rank';let a=implementationRows.filter(r=>!q||`${r.projectId} ${r.project?.name||''} ${r.project?.institution||''}`.toLowerCase().includes(q)).filter(r=>!cl||r.classification===cl);a=[...a];if(sort==='score')a.sort((x,y)=>(Number(y.implementationScore)||-1)-(Number(x.implementationScore)||-1));if(sort==='capexAsc')a.sort((x,y)=>(Number(x.investment)||Infinity)-(Number(y.investment)||Infinity));if(sort==='capexDesc')a.sort((x,y)=>(Number(y.investment)||-1)-(Number(x.investment)||-1));return a;}
function renderImplementation(){const rows=implementationFiltered();const done=rows.filter(r=>Number.isFinite(Number(r.implementationScore))).length;document.getElementById('implSummary').innerHTML=`<div class="impl-kpi"><b>${rows.length}</b><div>Proyectos seleccionados</div></div><div class="impl-kpi"><b>${done}</b><div>Evaluados</div></div><div class="impl-kpi"><b>${rows.filter(r=>r.status==='Información de preinversión incompleta').length}</b><div>Información incompleta</div></div><div class="impl-kpi"><b>${rows.filter(r=>(r.exclusions||[]).length).length}</b><div>Pendientes de subsanación</div></div>`;
 document.getElementById('implTable').innerHTML=`<table class="impl-table"><thead><tr><th>#</th><th>ID HUB</th><th>PROYECTO</th><th>ESTADO</th><th>ECON.</th><th>AMB.</th><th>SOCIAL</th><th>MADUREZ</th><th>IMPLEMENTACIÓN</th><th>CLASIFICACIÓN</th><th>CAPEX</th><th></th></tr></thead><tbody>${rows.map((r,i)=>`<tr><td>${i+1}</td><td><b>${implEsc(r.projectId)}</b></td><td>${implEsc(r.project?.name||'')}<div class="impl-note">${implEsc(r.project?.institution||'')}</div></td><td>${implEsc(r.status||'Información de preinversión incompleta')}</td><td>${r.dimensions?.DE??'—'}</td><td>${r.dimensions?.DA??'—'}</td><td>${r.dimensions?.DS??'—'}</td><td>${r.dimensions?.DM??'—'}</td><td><b>${r.implementationScore??'—'}</b></td><td>${implEsc(r.classification||'—')}</td><td>${implMoney(r.investment)}</td><td><button class="impl-btn" onclick="openImplementationProject('${String(r.projectId).replace(/'/g,"\\'")}')">${PRIORITIZATION_VIEW_MODE==='readonly'?'Ver':'Diligenciar'}</button></td></tr>`).join('')||'<tr><td colspan="12">No hay proyectos seleccionados para implementación.</td></tr>'}</tbody></table>`;
 renderImplementationMatrix(rows);
}
function renderImplementationMatrix(rows){const box=document.getElementById('implMatrix');box.querySelectorAll('.impl-point').forEach(x=>x.remove());const valid=rows.filter(r=>Number.isFinite(Number(r.implementationScore))&&Number.isFinite(Number(r.investment)));if(!valid.length)return;const caps=valid.map(r=>Number(r.investment)),max=Math.max(...caps,1);valid.forEach(r=>{const x=8+84*Math.min(1,Number(r.investment)/max),y=92-84*(Number(r.implementationScore)/100);const b=document.createElement('button');b.className='impl-point';b.style.left=x+'%';b.style.top=y+'%';b.textContent=r.projectId;b.title=`${r.project?.name||''} · Puntaje ${r.implementationScore} · CAPEX ${r.investment}`;b.onclick=()=>openImplementationProject(r.projectId);box.appendChild(b);});}
function implInputHtml(id,cfg,val,score){if(cfg.type==='number')return `<input data-impl-id="${id}" type="number" step="any" value="${val??''}" ${PRIORITIZATION_VIEW_MODE==='readonly'?'disabled':''} placeholder="${implEsc(cfg.unit||'')}">`;return `<select data-impl-id="${id}" ${PRIORITIZATION_VIEW_MODE==='readonly'?'disabled':''}><option value="">Seleccione...</option>${(cfg.options||[]).map(o=>`<option value="${implEsc(o[0])}" ${String(val)===String(o[0])?'selected':''}>${implEsc(o[0])} · ${o[1]} pts</option>`).join('')}</select>`;}
async function openImplementationProject(id){implementationActiveId=id;const r=await api(`/projects/${encodeURIComponent(id)}/implementation-prioritization`);const d=r.implementation||{},dlg=document.getElementById('implDialog'),dims=implementationConfig.dimensions||{};dlg.innerHTML=`<div class="impl-dlg-head"><div><h2 style="margin:0">${implEsc(id)} · Evaluación de implementación</h2><div class="impl-note">${implEsc(d.project?.name||'')} · ${implEsc(d.project?.institution||'')}</div></div><button class="impl-btn secondary" onclick="document.getElementById('implDialog').close()">Cerrar</button></div><div class="impl-dlg-body"><div class="impl-status ${(d.scores?.exclusions||[]).length?'impl-bad':''}"><b>Estado:</b> ${implEsc(d.scores?.status||'Información de preinversión incompleta')} · <b>Faltantes:</b> ${(d.scores?.missing||[]).length}${(d.scores?.exclusions||[]).length?`<br><b>Condiciones excluyentes:</b> ${implEsc((d.scores.exclusions||[]).join(', '))}`:''}</div>${Object.entries(dims).map(([dim,label])=>`<div class="impl-dim"><h3>${implEsc(label)} <span class="impl-note">25%</span></h3>${Object.entries(implementationConfig.indicators).filter(([_,c])=>c.dimension===dim).map(([iid,c])=>`<div class="impl-field"><div><b>${iid}</b></div><div><b>${implEsc(c.name)}</b><div class="impl-note">Fuente: ${implEsc(c.source||'')}</div></div><div>${implInputHtml(iid,c,d.raw?.[iid],d.scores?.indicatorScores?.[iid])}</div><div class="impl-score">${d.scores?.indicatorScores?.[iid]??'—'} pts</div></div>`).join('')}<div class="impl-status"><b>Puntaje ${implEsc(label)}:</b> ${d.scores?.dimensions?.[dim]??'—'}</div></div>`).join('')}<div class="impl-card"><h3>Resultado</h3><div class="impl-grid"><div class="impl-kpi"><b>${d.scores?.implementationScore??'—'}</b><div>Puntaje de Implementación</div></div><div class="impl-kpi"><b>${implEsc(d.scores?.classification||'—')}</b><div>Clasificación</div></div><div class="impl-kpi"><b>${implMoney(d.project?.investment)}</b><div>CAPEX</div></div></div></div>${PRIORITIZATION_VIEW_MODE==='readonly'?'':`<button id="implSave" class="impl-btn">Guardar y recalcular</button>`}</div>`;dlg.showModal();if(PRIORITIZATION_VIEW_MODE!=='readonly')document.getElementById('implSave').onclick=saveImplementationProject;}
async function saveImplementationProject(){if(!implementationActiveId)return;const raw={};document.querySelectorAll('[data-impl-id]').forEach(el=>{if(el.value!=='')raw[el.dataset.implId]=el.value;});const b=document.getElementById('implSave');b.disabled=true;b.textContent='Guardando…';try{const r=await api(`/projects/${encodeURIComponent(implementationActiveId)}/implementation-prioritization`,{method:'PUT',body:JSON.stringify({raw,updatedBy:'Técnico evaluador'})},true);toast('Evaluación guardada.');const rr=await api('/project-prioritization/ranking');implementationRows=rr.ranking||[];renderImplementation();await openImplementationProject(implementationActiveId);try{parent.postMessage({type:'hub-project-prioritization-updated',projectId:implementationActiveId,implementation:r.implementation},'*')}catch(e){}}finally{if(b){b.disabled=false;b.textContent='Guardar y recalcular';}}}
window.openImplementationProject=openImplementationProject;

function removePrioritizationMap(){
  const map=document.getElementById("spatialMap");
  if(map){
    const toolbar=map.previousElementSibling;
    if(toolbar?.classList?.contains("map-toolbar"))toolbar.remove();
    map.remove();
  }
}

async function startPrioritization() {
  if(PRIORITIZATION_SCOPE==="projects"){try{await loadImplementationModule();}catch(e){console.error(e);document.body.innerHTML=`<div style="padding:30px;font-family:system-ui"><h2>No fue posible cargar Priorización de proyectos</h2><pre>${esc(e?.message||e)}</pre></div>`;}return;}
  removePrioritizationMap();
  try {
    initPrioritizationEvents();
    console.log("PRIORIZACIÓN V2: cargando información...");
    await loadAll();
    const requestedProject=new URLSearchParams(window.location.search).get("project");
    if(requestedProject){
      try{await openProject(requestedProject);}catch(e){console.warn("No se pudo abrir automáticamente el proyecto solicitado",e);}
    }
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
const GEO_REASON = {
  "ND-06":"Se calcula espacialmente porque depende de la relación del proyecto con equipamientos y flujos peatonales próximos.",
  "ND-07":"Se calcula con accidentes georreferenciados dentro del área de influencia del proyecto.",
  "ND-08":"Se deriva del cruce del proyecto con deslizamientos y coronas de deslizamiento y su proximidad.",
  "RR-01":"Combina las amenazas disponibles que intersectan o se ubican próximas al proyecto; actualmente usa inundación y, cuando estén disponibles, deslizamientos/coronas.",
  "RR-06":"Se obtiene de la jerarquía de la Red Vial Nacional de CONAVI: una vía principal/primaria recibe mayor criticidad y la puntuación disminuye para jerarquías inferiores.",
  "SI-01":"Se deriva del cruce con áreas silvestres protegidas, humedales, corredores biológicos y otras áreas ambientalmente sensibles."
};

