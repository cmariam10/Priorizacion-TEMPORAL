
const steps=[
{n:1,title:"Registro de idea",phase:1,role:"Técnico / Profesional",desc:"Identifica la necesidad y registra la iniciativa."},
{n:2,title:"Revisión técnica",phase:1,role:"Jefatura Técnica",desc:"Valida completitud, alcance inicial y consistencia."},
{n:3,title:"Alineación institucional",phase:1,role:"Planeamiento",desc:"Revisa alineación con planes y prioridades."},
{n:4,title:"Aprobación institucional",phase:1,role:"Dirección / Gerencia",desc:"Aprueba envío al HUB."},
{n:5,title:"Radicación ante el MOPT",phase:1,role:"Coordinador HUB",desc:"Formaliza la iniciativa y soportes."},
{n:6,title:"Revisión y clasificación",phase:2,role:"Secretaría Técnica",desc:"Verifica requisitos, clasifica y asigna revisión."},
{n:7,title:"Evaluación técnica central",phase:2,role:"Áreas Técnicas",desc:"Evalúa componentes técnicos y sectoriales."},
{n:8,title:"Evaluación automatizada y calificación del proyecto",phase:2,role:"Sistema HUB / Secretaría Técnica HUB",desc:"El HUB califica la iniciativa mediante criterios parametrizados definidos por el MOPT según las características del proyecto."},
{n:9,title:"Priorización y decisión central",phase:2,role:"Comité HUB Central",desc:"Analiza los resultados de la evaluación automatizada y define la priorización final de las iniciativas."},
{n:10,title:"Formulación definitiva",phase:3,role:"Institución ejecutora",desc:"Desarrolla estudios, diseños y presupuesto."},
{n:11,title:"Registro como proyecto",phase:3,role:"Registro HUB",desc:"Registra formalmente el proyecto aprobado."},
{n:12,title:"Asignación de responsables",phase:3,role:"Institución ejecutora",desc:"Define equipo, unidad ejecutora y responsables."},
{n:13,title:"Planificación de ejecución",phase:3,role:"Institución ejecutora",desc:"Define cronograma, hitos y plan de adquisiciones."},
{n:14,title:"Ejecución",phase:4,role:"Institución ejecutora",desc:"Inicia contratación, obra o servicios."},
{n:15,title:"Seguimiento y control",phase:4,role:"Supervisor",desc:"Monitorea avance físico, financiero y riesgos."},
{n:16,title:"Supervisión / interventoría",phase:4,role:"Supervisor",desc:"Verifica calidad, cumplimiento e informes."},
{n:17,title:"Cierre del proyecto",phase:4,role:"Comité / ejecutor",desc:"Cierra, liquida y registra lecciones aprendidas."}
];
const BASE_PROJECTS = [];
let projects = []; // Solo datos reales provenientes del Excel maestro.
let activeId=projects[0]?.id||"";
let currentRole="visitor";
const canRegisterRoles=["tecnico","coordinador","secretaria"];
const roleStepPermissions={tecnico:[],jefatura:[2],planeamiento:[3],direccion:[4],coordinador:[5],secretaria:[6,8],areas:[7],sistema:[8],comite:[9],ejecutor:[10,11,12,13,14,15,16,17]};
let auditLog=[];
const docEditPermissions={
 tecnico:["Ficha de iniciativa","Diagnóstico / justificación"],
 jefatura:["Ficha de iniciativa","Diagnóstico / justificación","Estimación de costos"],
 planeamiento:["Aval institucional"],
 direccion:["Aval institucional"],
 coordinador:["Ficha de iniciativa","Aval institucional","Plan de ejecución"],
 secretaria:["Concepto técnico","Concepto financiero","Plan de ejecución"],
 areas:["Concepto técnico"],
 sistema:["Matriz de criterios"],
 comite:[],
 ejecutor:["Plan de ejecución","Estimación de costos"]
};
let documents=[];
const BASE_DOCUMENTS = [];

const HUB_USERS={
 usuario1:{password:"Hub2026!01",name:"Persona de prueba 1"},
 usuario2:{password:"Hub2026!02",name:"Persona de prueba 2"},
 usuario3:{password:"Hub2026!03",name:"Persona de prueba 3"},
 usuario4:{password:"Hub2026!04",name:"Persona de prueba 4"},
 usuario5:{password:"Hub2026!05",name:"Persona de prueba 5"}
};
let currentUser=null;
function userStorageKey(){return currentUser?`hub_demo_state_v2_${currentUser.username}`:null;}
function resetWorkspace(){
 projects=[];
 documents=JSON.parse(JSON.stringify(BASE_DOCUMENTS));
 auditLog=[]; activeId=projects[0]?.id||"";
}
function saveHubState(){
 const key=userStorageKey(); if(!key)return;
 try{localStorage.setItem(key,JSON.stringify({projects,documents,auditLog}));}catch(e){console.warn("No fue posible guardar el estado HUB",e);}
}
function loadHubState(){
 resetWorkspace();
 const key=userStorageKey(); if(!key)return;
 try{
  const raw=localStorage.getItem(key); if(!raw)return;
  const st=JSON.parse(raw);
  /* Excel es la única fuente maestra: no restaurar proyectos desde localStorage. */
  documents=[]; // Los documentos se cargan únicamente desde OneDrive.
  if(Array.isArray(st.auditLog))auditLog=st.auditLog;
  activeId=projects[0]?.id||"";
 }catch(e){console.warn("No fue posible recuperar el estado HUB",e);}
}
function actorLabel(){
 const roleName=roleSelect?.options?.[roleSelect.selectedIndex]?.text||currentRole;
 return currentUser?`${currentUser.name} · Perfil: ${roleName}`:currentRole;
}

const TOTAL_WORKFLOW_STEPS=steps.length;
const ROLE_DISPLAY_STEPS={tecnico:"1",jefatura:"2",planeamiento:"3",direccion:"4",coordinador:"5",secretaria:"6 y 8",areas:"7",sistema:"8",comite:"9",ejecutor:"10–17"};
let sharedFlowFolder=null, sharedFlowSyncTimer=null, internalNotifications=[];

function notificationStorageKey(){
 return `hub_notif_seen_${currentUser?.username||"anon"}_${currentRole}`;
}
function getSeenNotifications(){
 try{return JSON.parse(localStorage.getItem(notificationStorageKey())||"{}")}catch(e){return {}}
}
function setSeenNotifications(v){localStorage.setItem(notificationStorageKey(),JSON.stringify(v||{}));}
function notificationProjects(){
 if(!currentUser||currentRole==="visitor")return [];
 const allowed=currentRole==="tecnico"?[1]:(roleStepPermissions[currentRole]||[]);
 return projects.filter(p=>(allowed.includes(Number(p.step))||(currentRole==="ejecutor"&&p.step>=10))&&p.sharedUpdatedAt);
}
function unreadNotificationProjects(){
 const seen=getSeenNotifications();
 return notificationProjects().filter(p=>seen[p.id]!==p.sharedUpdatedAt);
}
function ensureTabBadge(tab,prefix,count){
 if(!tab)return;
 let badge=tab.querySelector(`.${prefix}-badge`);
 if(!badge){
  badge=document.createElement("span");badge.className=`${prefix}-badge`;
  badge.style.cssText="margin-left:7px;background:#c62828;color:#fff;border-radius:999px;padding:2px 7px;font-size:11px;font-weight:800;display:none";
  tab.appendChild(badge);
 }
 badge.textContent=String(count);badge.style.display=count>0?"inline-block":"none";
}

async function refreshBadgesAfterSharedSync(){ refreshWorkBadges(); }
window.refreshBadgesAfterSharedSync=refreshBadgesAfterSharedSync;

function refreshWorkBadges(){
 const tabs=[...document.querySelectorAll(".protected-tab")];
 const notifTab=tabs.find(t=>t.textContent.trim().startsWith("Notificaciones"));
 const inboxTab=tabs.find(t=>t.textContent.trim().startsWith("Bandeja"));
 ensureTabBadge(notifTab,"hub-notif",unreadNotificationProjects().length);
 let pending=0;
 if(currentUser&&currentRole!=="visitor"){
  const allowed=currentRole==="tecnico"?[1]:(roleStepPermissions[currentRole]||[]);
  pending=projects.filter(p=>allowed.includes(Number(p.step))||(currentRole==="ejecutor"&&p.step>=10)).length;
 }
 ensureTabBadge(inboxTab,"hub-inbox",pending);
}
async function markNotificationRead(id,stamp){
 const seen=getSeenNotifications();seen[id]=stamp;setSeenNotifications(seen);
 refreshWorkBadges();await renderNotifications();
}
window.markNotificationRead=markNotificationRead;


async function resolveSharedFlowFolder(){
 if(sharedFlowFolder)return sharedFlowFolder;
 const root=await resolveRootFolder();
 sharedFlowFolder=await ensureChildFolder(root.driveId,root.itemId,"_HUB_FLUJO");
 return sharedFlowFolder;
}
function workflowPayload(p){
 return {projectId:p.id,step:Number(p.step)||1,state:p.state||"",owner:p.owner||"",returned:!!p.returned,rejected:!!p.rejected,updatedAt:new Date().toISOString(),updatedBy:actorLabel()};
}
async function saveSharedWorkflow(p,movement=null){
 const f=await resolveSharedFlowFolder();
 const filename=safeOneDriveName(p.id)+".json";
 let previous=null;
 try{
  previous=await graphFetch(`/drives/${f.driveId}/items/${f.itemId}:/${encodeURIComponent(filename)}:/content`);
 }catch(e){
  if(!String(e?.message||e).includes("404"))console.warn("No se pudo leer historial previo de flujo",e);
 }
 const current=workflowPayload(p);
 let history=Array.isArray(previous?.history)?previous.history.slice():[];

 // Migra automáticamente el formato antiguo: conserva la última situación que ya existía.
 if(previous && !Array.isArray(previous.history) && previous.updatedAt){
  history.push({
   eventId:"legacy-"+String(previous.updatedAt),
   step:Number(previous.step)||1,
   state:previous.state||"",
   owner:previous.owner||"",
   returned:!!previous.returned,
   rejected:!!previous.rejected,
   date:previous.updatedAt,
   updatedBy:previous.updatedBy||"",
   action:"Estado histórico previo"
  });
 }

 if(movement){
  const event={
   eventId:movement.eventId||(`${movement.date||new Date().toISOString()}-${Math.random().toString(36).slice(2,8)}`),
   date:movement.date||new Date().toISOString(),
   action:movement.action||"Gestión del proyecto",
   comment:movement.comment||"",
   stepFrom:Number(movement.stepFrom)||null,
   stateFrom:movement.stateFrom||"",
   stepTo:Number(movement.stepTo)||Number(p.step)||1,
   stateTo:movement.stateTo||p.state||"",
   ownerTo:movement.ownerTo||p.owner||"",
   updatedBy:movement.updatedBy||actorLabel()
  };
  if(!history.some(h=>h.eventId===event.eventId))history.push(event);
 }
 history.sort((a,b)=>String(a.date||"").localeCompare(String(b.date||"")));

 const payload={...current,historyCount:history.length,history};
 await graphFetch(`/drives/${f.driveId}/items/${f.itemId}:/${encodeURIComponent(filename)}:/content`,{
  method:"PUT",headers:{"Content-Type":"application/json"},body:JSON.stringify(payload,null,2)
 });
 return payload;
}
function applySharedWorkflow(p,s){
 if(!s)return;
 p.step=Number(s.step)||p.step||1;p.state=s.state||p.state;p.owner=s.owner||p.owner;
 p.returned=!!s.returned;p.rejected=!!s.rejected;p.sharedUpdatedAt=s.updatedAt||null;p.sharedUpdatedBy=s.updatedBy||null;
}
async function loadAllSharedWorkflow(){
 try{
  const f=await resolveSharedFlowFolder();
  const r=await graphFetch(`/drives/${f.driveId}/items/${f.itemId}/children?$select=id,name,file&$top=999`);
  const files=(r.value||[]).filter(x=>x.file&&x.name.endsWith(".json"));
  const byId=new Map();
  for(const item of files){
   try{const s=await graphFetch(`/drives/${f.driveId}/items/${item.id}/content`);if(s?.projectId)byId.set(String(s.projectId),s);}catch(e){}
  }
  projects.forEach(p=>applySharedWorkflow(p,byId.get(String(p.id))));
  return byId;
 }catch(e){console.warn("No se pudo sincronizar el flujo compartido",e);return new Map();}
}
async function recoverLegacyWorkflowFromTrace(){
 // Recupera movimientos hechos antes de crear _HUB_FLUJO, sin tocar proyectos que ya tienen estado compartido.
 try{
  const root=await resolveRootFolder();
  const rr=await graphFetch(`/drives/${root.driveId}/items/${root.itemId}/children?$select=id,name,folder&$top=999`);
  const tf=(rr.value||[]).find(x=>x.folder&&x.name==="_HUB_TRAZABILIDAD");if(!tf)return;
  const tr=await graphFetch(`/drives/${root.driveId}/items/${tf.id}/children?$select=id,name,folder&$top=999`);
  const folders=(tr.value||[]).filter(x=>x.folder);
  const candidates=projects.filter(p=>!p.sharedUpdatedAt);
  for(const p of candidates){
   const pf=folders.find(x=>x.name===safeOneDriveName(p.id));if(!pf)continue;
   const fr=await graphFetch(`/drives/${root.driveId}/items/${pf.id}/children?$select=id,name,file&$top=999`);
   const jsons=(fr.value||[]).filter(x=>x.file&&x.name.endsWith(".json")).slice(-30);
   let latest=null;
   for(const jf of jsons){
    try{const ev=await graphFetch(`/drives/${root.driveId}/items/${jf.id}/content`);
      if(ev?.accion==="GESTION_FLUJO"&&ev?.pasoDestino&&(!latest||String(ev.fecha)>String(latest.fecha)))latest=ev;
    }catch(e){}
   }
   if(latest){
    applySharedWorkflow(p,{step:latest.pasoDestino,state:latest.estadoDestino,owner:steps.find(s=>s.n==latest.pasoDestino)?.role||p.owner,returned:/devuelto/i.test(latest.estadoDestino||""),rejected:/rechazado|archivado/i.test(latest.estadoDestino||""),updatedAt:latest.fecha,updatedBy:latest.nombreHUB||latest.usuarioHUB});
    await saveSharedWorkflow(p);
   }
  }
 }catch(e){console.warn("No se pudo recuperar flujo histórico",e);}
}
function startSharedFlowSync(){
 if(sharedFlowSyncTimer)clearInterval(sharedFlowSyncTimer);
 sharedFlowSyncTimer=setInterval(async()=>{
  if(!currentUser)return;
  const before=projects.map(p=>`${p.id}|${p.step}|${p.state}|${p.owner}`).join(";");
  await loadAllSharedWorkflow();
  const after=projects.map(p=>`${p.id}|${p.step}|${p.state}|${p.owner}`).join(";");
  if(before!==after){renderKpis();renderList();renderMatrix();renderInbox();refreshWorkBadges();if(activeTab==="notifications")await renderNotifications();}
 },30000);
}
// ===== Microsoft OneDrive / Graph =====
const MS_CONFIG={
 clientId:"3a5ead05-e9e6-477b-9a63-24ef1f9fefd5",
 tenantId:"common",
 redirectUri:"https://cmariam10.github.io/Priorizacion-TEMPORAL/",
 rootShareUrl:"https://mariampenar-my.sharepoint.com/:f:/g/personal/mariampena_mariampenar_onmicrosoft_com/IgAqBC0CtqOmS6-OH1fqcWt2AaAj1L_cMwOuhSdlPvdU6Lw?e=oIbXkP",
 excelShareUrl:"https://mariampenar-my.sharepoint.com/:x:/g/personal/mariampena_mariampenar_onmicrosoft_com/IQDKEAZoAhkNTrGSyIHzJhqEAbmFWdlEVNOFICupivCl3rM?e=tGqA5x",
 repositoryDriveId:"b!rN3lIXRiLEqBrnrpDsCpsL9C27sCaXNGgiWM2lm6aiGmKI-8whwHT4Po8BEHpXPF",
 rootItemId:"01GCSG2KRKAQWQFNVDUZF27DQ7K7VHC23W",
 excelItemId:"01GCSG2KWKCADGQAQZBVHLDEWIQHZSMGUE"
};
const GRAPH_SCOPES=["User.Read","Files.ReadWrite"];
const HUB_API={
 baseUrl:"https://hub-proyectos-mopt-api.cmariam10.workers.dev",
 scope:"api://3a5ead05-e9e6-477b-9a63-24ef1f9fefd5/access_as_user"
};
let hubApiIdentity=null;
let msalApp=null, graphAccount=null, graphUser=null, graphRoot=null;
const projectFolderCache=new Map();

function safeOneDriveName(v){
 return String(v||"").replace(/[~"#%&*:<>?/\\{|}]/g,"-").replace(/\s+/g," ").trim().slice(0,180)||"Sin nombre";
}
function projectFolderName(p){return `${safeOneDriveName(p.id)} - ${safeOneDriveName(p.name)}`;}
function encodeShareUrl(url){
 const bytes=new TextEncoder().encode(url); let bin=""; bytes.forEach(b=>bin+=String.fromCharCode(b));
 return "u!"+btoa(bin).replace(/=+$/g,"").replace(/\//g,"_").replace(/\+/g,"-");
}
async function initMicrosoft(){
 if(msalApp)return msalApp;
 if(!window.msal)throw new Error("No se cargó Microsoft Authentication Library (MSAL).");
 msalApp=new msal.PublicClientApplication({auth:{clientId:MS_CONFIG.clientId,authority:`https://login.microsoftonline.com/${MS_CONFIG.tenantId}`,redirectUri:MS_CONFIG.redirectUri},cache:{cacheLocation:"sessionStorage",storeAuthStateInCookie:false}});
 if(typeof msalApp.initialize==="function")await msalApp.initialize();
 const result=await msalApp.handleRedirectPromise();
 graphAccount=result?.account||msalApp.getAllAccounts()[0]||null;
 if(graphAccount)msalApp.setActiveAccount(graphAccount);
 return msalApp;
}
async function hubApiToken(interactive=true){
 await initMicrosoft();
 graphAccount=msalApp.getActiveAccount()||msalApp.getAllAccounts()[0]||graphAccount;
 const inPopup=!!window.opener && window.opener!==window;
 if(!graphAccount){
  if(!interactive||inPopup)throw new Error("Debe iniciar sesión con Microsoft para acceder al HUB.");
  const r=await msalApp.loginPopup({scopes:[HUB_API.scope],prompt:"select_account"});
  graphAccount=r.account;msalApp.setActiveAccount(graphAccount);
  if(r.accessToken)return r.accessToken;
 }
 try{
  return (await msalApp.acquireTokenSilent({scopes:[HUB_API.scope],account:graphAccount})).accessToken;
 }catch(e){
  if(!interactive||inPopup)throw e;
  const r=await msalApp.acquireTokenPopup({scopes:[HUB_API.scope],account:graphAccount});
  if(r.account){graphAccount=r.account;msalApp.setActiveAccount(graphAccount);}
  return r.accessToken;
 }
}
async function hubApiFetch(path,options={}){
 const token=await hubApiToken(true);
 const headers=new Headers(options.headers||{});
 headers.set("Authorization",`Bearer ${token}`);
 headers.set("X-HUB-User",currentUser?.username||"");
 headers.set("X-HUB-Role",currentRole||"");
 const res=await fetch(`${HUB_API.baseUrl}${path}`,{...options,headers});
 const ct=res.headers.get("content-type")||"";
 const data=ct.includes("json")?await res.json():await res.text();
 if(!res.ok)throw new Error(data?.error||data?.message||`HUB API ${res.status}`);
 return data;
}
async function verifyHubMicrosoftIdentity(){
 const r=await hubApiFetch("/auth/me");
 if(!r?.identity?.verified)throw new Error("Microsoft no pudo verificar la identidad.");
 hubApiIdentity=r.identity;
 return hubApiIdentity;
}

async function graphToken(interactive=true){
 await initMicrosoft();
 graphAccount=msalApp.getActiveAccount()||msalApp.getAllAccounts()[0]||graphAccount;

 // Only the top-level HUB page may launch an interactive MSAL popup.
 // A popup opened by MSAL must never try to open another popup.
 const inPopup=!!window.opener && window.opener!==window;
 if(!graphAccount){
  if(!interactive||inPopup)throw new Error("Debe iniciar sesión con Microsoft para acceder a OneDrive.");
  const r=await msalApp.loginPopup({scopes:GRAPH_SCOPES,prompt:"select_account"});
  graphAccount=r.account;
  msalApp.setActiveAccount(graphAccount);
  if(r.accessToken)return r.accessToken;
 }

 try{
  return (await msalApp.acquireTokenSilent({
   scopes:GRAPH_SCOPES,
   account:graphAccount
  })).accessToken;
 }catch(e){
  if(!interactive||inPopup)throw e;
  // Interaction is allowed only from the main HUB window.
  const r=await msalApp.acquireTokenPopup({
   scopes:GRAPH_SCOPES,
   account:graphAccount
  });
  if(r.account){graphAccount=r.account;msalApp.setActiveAccount(graphAccount);}
  return r.accessToken;
 }
}
async function graphFetch(path,options={}){
 const token=await graphToken(true);
 const headers=new Headers(options.headers||{});
 headers.set("Authorization",`Bearer ${token}`);
 const url=path.startsWith("http")?path:`https://graph.microsoft.com/v1.0${path}`;
 let res=await fetch(url,{...options,headers,redirect:"follow"});

 // Defensive handling for Graph/SharePoint redirects that can occasionally
 // surface as 307/308 in cross-tenant shared-item scenarios.
 if((res.status===307||res.status===308)){
  const location=res.headers.get("location");
  if(location){
   res=await fetch(location,{...options,headers,redirect:"follow"});
  }
 }
 if(!res.ok){
  let detail="";
  try{detail=(await res.clone().json())?.error?.message||""}catch(e){}
  const shortPath=String(path).replace(/^https:\/\/graph\.microsoft\.com\/v1\.0/,"");
  throw new Error(`Microsoft Graph ${res.status}: ${detail||res.statusText} [${shortPath}]`);
 }
 if(res.status===204)return null;
 const ct=res.headers.get("content-type")||"";
 return ct.includes("json")?res.json():res.blob();
}

function normalizeSharedDriveItem(item){
 if(!item)return null;
 // If Graph returns a local placeholder for an item from another drive,
 // always switch to the source item identifiers before subsequent requests.
 if(item.remoteItem){
  const r=item.remoteItem;
  return {
   ...r,
   id:r.id,
   name:r.name||item.name,
   webUrl:r.webUrl||item.webUrl,
   folder:r.folder||item.folder,
   file:r.file||item.file,
   parentReference:r.parentReference||item.parentReference
  };
 }
 return item;
}
function sharedItemCacheKey(shareUrl){
 let h=0;for(let i=0;i<shareUrl.length;i++)h=((h<<5)-h)+shareUrl.charCodeAt(i)|0;
 return `hub_shared_item_${h}`;
}

async function resolveSharedDriveItem(shareUrl){
 const cacheKey=sharedItemCacheKey(shareUrl);
 try{
  const cached=JSON.parse(sessionStorage.getItem(cacheKey)||"null");
  if(cached?.driveId&&cached?.itemId){
   const direct=await graphFetch(`/drives/${encodeURIComponent(cached.driveId)}/items/${encodeURIComponent(cached.itemId)}?$select=id,name,webUrl,parentReference,folder,file,remoteItem`);
   return normalizeSharedDriveItem(direct);
  }
 }catch(e){sessionStorage.removeItem(cacheKey);}

 const token=encodeShareUrl(shareUrl);
 const item=normalizeSharedDriveItem(await graphFetch(
  `/shares/${token}/driveItem?$select=id,name,webUrl,parentReference,folder,file,remoteItem`,
  {headers:{"Prefer":"redeemSharingLink"}}
 ));
 if(!item?.id)throw new Error("Microsoft Graph no devolvió el identificador del recurso compartido.");
 const driveId=item.parentReference?.driveId;
 if(!driveId)throw new Error("Microsoft Graph no devolvió el driveId del recurso compartido.");

 // From this point forward the HUB works only with the canonical driveId/itemId,
 // not with the sharing URL.
 sessionStorage.setItem(cacheKey,JSON.stringify({driveId,itemId:item.id}));
 return item;
}

// ===== Base maestra de proyectos: Excel en OneDrive / SharePoint =====
let excelBook=null, excelTable=null, excelHeaders=["Referencia de fila", "ID HUB", "Código del BPIP", "Estado", "Estado (PREVIO, REVISAR)", "Etapa", "Nombre corto", "Nombre del proyecto", "¿Pertenece a algún programa?", "Programa vinculado", "Planes estratégicos vinculados (siglas)", "Institución", "Ubicación político-administrativa (provincia, cantón)", "Ruta Nacional", "Inversión por proyecto (Millones USD)", "Coordenada Este (CRTM 05)", "Coordenada Norte (CRTM 05) ", "Tipo de Obra", "Sub tipo de obra", "Identificación de Fuentes y Modalidades de Financiamiento", "Otras posibles formas de financiamiento", "Comentarios adicionales sobre el financiamiento", "Estado revisión", "Fecha actualización", "Observaciones", "Completo", "Estado HUB"];

let pendingProjectGeometry=null, registrationMap=null, fichaLeafletMap=null, registrationMapLayers=[];
let projectGeometryDirty=false, originalProjectGeometry=null;

let portfolioMap=null, portfolioMapLayer=null, portfolioMapFilterEnabled=false, portfolioMapMoveTimer=null, portfolioMapGeometries=new Map(), portfolioMapInitializing=false, portfolioMapReady=false;

function pointInBounds(lat,lon,b){return b&&b.contains([Number(lat),Number(lon)]);}
function geometryInBounds(g,b){
 if(!g||!b)return false;
 if(g.type==="segment"){
  const pts=(g.coordinates||[]).map(c=>({lat:c[1],lon:c[0]}));
  return pts.some(p=>pointInBounds(p.lat,p.lon,b));
 }
 const p=g.referencePoints?.[0];return p?pointInBounds(p.lat,p.lon,b):false;
}
async function getPortfolioGeometry(p){
 if(portfolioMapGeometries.has(p.id))return portfolioMapGeometries.get(p.id);
 const x=inferGeo(p),g={type:"point",source:"inferred",referencePoints:[{lat:x.lat,lon:x.lon}]};
 portfolioMapGeometries.set(p.id,g);return g;
}

let portfolioSearchFocusLayer=null;

async function focusProjectFromList(id){
 const p=projects.find(x=>x.id===id);
 if(!p||!portfolioMapReady||!portfolioMap)return;

 clearPortfolioSearchFocus();

 // For an explicit user click, always check the persisted geometry first.
 let g=null;
 try{g=await loadProjectGeometry(id);}catch(e){}
 if(g&&validProjectGeometry(g))portfolioMapGeometries.set(id,g);
 else g=await getPortfolioGeometry(p);

 if(g?.type==="segment"&&g.coordinates?.length>1){
  portfolioSearchFocusLayer=L.polyline(g.coordinates.map(c=>[Number(c[1]),Number(c[0])]),{weight:9,opacity:1})
   .bindPopup(`<b>${traceEscape(p.id)}</b><br>${traceEscape(p.name)}`).addTo(portfolioMap);
  portfolioMap.fitBounds(portfolioSearchFocusLayer.getBounds().pad(.25),{maxZoom:16});
 }else{
  const q=g?.referencePoints?.[0],lat=Number(q?.lat),lon=Number(q?.lon);
  if(Number.isFinite(lat)&&Number.isFinite(lon)){
   portfolioSearchFocusLayer=L.circleMarker([lat,lon],{radius:12,weight:4,fillOpacity:.9})
    .bindPopup(`<b>${traceEscape(p.id)}</b><br>${traceEscape(p.name)}`).addTo(portfolioMap);
   portfolioMap.setView([lat,lon],15);
  }
 }
 if(portfolioSearchFocusLayer?.openPopup)portfolioSearchFocusLayer.openPopup();
 document.getElementById("portfolioProjectsMap")?.scrollIntoView({behavior:"smooth",block:"center"});
}
window.focusProjectFromList=focusProjectFromList;

let portfolioSearchFocusTimer=null;

function clearPortfolioSearchFocus(){
 if(portfolioMap&&portfolioSearchFocusLayer){
  try{portfolioMap.removeLayer(portfolioSearchFocusLayer)}catch(e){}
 }
 portfolioSearchFocusLayer=null;
}

async function focusPortfolioMapFromSearch(list){
 if(!portfolioMapReady||!portfolioMap)return;
 const search=document.getElementById("globalSearch")?.value?.trim()||"";
 if(!search){clearPortfolioSearchFocus();return;}

 const matches=(list||[]).filter(p=>!p.rejected);
 if(!matches.length){clearPortfolioSearchFocus();return;}

 clearPortfolioSearchFocus();

 // One exact/unique result: zoom closely and emphasize it.
 if(matches.length===1){
  const p=matches[0],g=await getPortfolioGeometry(p);
  if(g?.type==="segment"&&g.coordinates?.length>1){
   portfolioSearchFocusLayer=L.polyline(g.coordinates.map(c=>[Number(c[1]),Number(c[0])]),{weight:9,opacity:1})
    .bindPopup(`<b>${traceEscape(p.id)}</b><br>${traceEscape(p.name)}`).addTo(portfolioMap);
   portfolioMap.fitBounds(portfolioSearchFocusLayer.getBounds().pad(.25),{maxZoom:16});
  }else{
   const q=g?.referencePoints?.[0],lat=Number(q?.lat),lon=Number(q?.lon);
   if(Number.isFinite(lat)&&Number.isFinite(lon)){
    portfolioSearchFocusLayer=L.circleMarker([lat,lon],{radius:12,weight:4,fillOpacity:.9})
     .bindPopup(`<b>${traceEscape(p.id)}</b><br>${traceEscape(p.name)}`).addTo(portfolioMap);
    portfolioMap.setView([lat,lon],15);
   }
  }
  if(portfolioSearchFocusLayer?.openPopup)portfolioSearchFocusLayer.openPopup();
  return;
 }

 // Several search matches: zoom to all of them.
 const pts=[];
 for(const p of matches){
  const g=await getPortfolioGeometry(p);
  if(g?.type==="segment"&&g.coordinates?.length){
   g.coordinates.forEach(c=>pts.push([Number(c[1]),Number(c[0])]));
  }else{
   const q=g?.referencePoints?.[0],lat=Number(q?.lat),lon=Number(q?.lon);
   if(Number.isFinite(lat)&&Number.isFinite(lon))pts.push([lat,lon]);
  }
 }
 if(pts.length)portfolioMap.fitBounds(pts,{padding:[45,45],maxZoom:13});
}

function currentMapFilteredIds(){
 if(!portfolioMap||!portfolioMapFilterEnabled)return null;
 const b=portfolioMap.getBounds(),ids=new Set();
 portfolioMapGeometries.forEach((g,id)=>{if(geometryInBounds(g,b))ids.add(id);});
 return ids;
}

async function refreshPortfolioMapLayers(){
 if(!portfolioMapReady||!portfolioMap||!portfolioMapLayer)return;
 portfolioMapLayer.clearLayers();

 // Apply only the ordinary UI filters here; temporarily ignore spatial bounds.
 const spatial=portfolioMapFilterEnabled;
 portfolioMapFilterEnabled=false;
 let visible;
 try{visible=getGlobalFilteredProjects();}finally{portfolioMapFilterEnabled=spatial;}

 for(const p of visible){
  const g=await getPortfolioGeometry(p);
  if(g?.type==="segment"&&g.coordinates?.length>1){
   L.polyline(g.coordinates.map(c=>[Number(c[1]),Number(c[0])]),{weight:4,opacity:.8})
    .bindPopup(`<b>${traceEscape(p.id)}</b><br>${traceEscape(p.name)}`).addTo(portfolioMapLayer);
  }else{
   const q=g?.referencePoints?.[0],lat=Number(q?.lat),lon=Number(q?.lon);
   if(Number.isFinite(lat)&&Number.isFinite(lon))
    L.circleMarker([lat,lon],{radius:5,weight:1.5,fillOpacity:.75})
     .bindPopup(`<b>${traceEscape(p.id)}</b><br>${traceEscape(p.name)}`).addTo(portfolioMapLayer);
  }
 }
}
function applyPortfolioMapFilter(){
 if(!portfolioMapFilterEnabled)return;
 renderKpis();renderList();
}
function askMapFilter(){
 if(portfolioMapFilterEnabled)return;
 if(confirm("¿Desea filtrar el listado y los indicadores según el área visible del mapa?")){
  portfolioMapFilterEnabled=true;
  const c=document.getElementById("portfolioMapFilterState");if(c)c.innerHTML='<b>Filtro por mapa activo.</b> Mueva o acerque el mapa para actualizar el listado.';
  applyPortfolioMapFilter();
 }
}
function applyMapFilterFromButton(){
 if(!portfolioMap||!portfolioMapReady){
  alert("El mapa todavía no está listo.");
  return;
 }
 portfolioMapFilterEnabled=true;
 const c=document.getElementById("portfolioMapFilterState");
 if(c)c.innerHTML="<b>Filtro por mapa activo.</b> El listado y los indicadores corresponden al área visible del mapa.";
 applyPortfolioMapFilter();
}
window.applyMapFilterFromButton=applyMapFilterFromButton;

function disablePortfolioMapFilter(){
 portfolioMapFilterEnabled=false;
 const c=document.getElementById("portfolioMapFilterState");if(c)c.innerHTML='Filtro del mapa desactivado. Ajuste el área visible y seleccione “Aplicar filtro con el mapa” cuando lo necesite.';
 renderKpis();renderList();
}

async function preloadPortfolioUserGeometries(){
 if(!currentUser)return;
 try{
  const f=await resolveGeometryFolder();
  const r=await graphFetch(`/drives/${f.driveId}/items/${f.itemId}/children?$select=id,name,file&$top=999`);
  const files=(r.value||[]).filter(x=>x.file&&x.name.toLowerCase().endsWith(".json"));
  await Promise.all(files.map(async x=>{
   try{
    const g=await graphFetch(`/drives/${f.driveId}/items/${x.id}/content`);
    if(g?.projectId)portfolioMapGeometries.set(g.projectId,g);
   }catch(e){}
  }));
 }catch(e){console.warn("No se pudieron precargar geometrías del portafolio",e);}
}
function addProjectToPortfolioMap(p,g){
 if(!portfolioMapLayer||!g)return;
 if(g.type==="segment"&&g.coordinates?.length>1){
  L.polyline(g.coordinates.map(c=>[c[1],c[0]]),{weight:4,opacity:.8})
   .bindPopup(`<b>${traceEscape(p.id)}</b><br>${traceEscape(p.name)}`).addTo(portfolioMapLayer);
 }else{
  const q=g.referencePoints?.[0];
  if(q)L.circleMarker([q.lat,q.lon],{radius:5,weight:2,fillOpacity:.75})
   .bindPopup(`<b>${traceEscape(p.id)}</b><br>${traceEscape(p.name)}`).addTo(portfolioMapLayer);
 }
}

async function initPortfolioMap(){
 const el=document.getElementById("portfolioProjectsMap");
 if(!el||!projects.length||portfolioMapInitializing)return;
 portfolioMapInitializing=true;
 const state=document.getElementById("portfolioMapFilterState");
 try{
  if(state)state.innerHTML="Cargando mapa del portafolio…";
  await loadLeaflet();

  // Wait until the inserted container has a real size.
  await new Promise(r=>requestAnimationFrame(()=>requestAnimationFrame(r)));
  if(!el.clientWidth||!el.clientHeight)throw new Error("El contenedor del mapa todavía no tiene dimensiones.");

  if(portfolioMap){try{portfolioMap.remove()}catch(e){} portfolioMap=null;}
  portfolioMap=L.map(el,{preferCanvas:true,zoomControl:true}).setView([9.93,-84.08],7);

  const tiles=L.tileLayer("https://tile.openstreetmap.org/{z}/{x}/{y}.png",{
   maxZoom:19,
   attribution:"© OpenStreetMap",
   crossOrigin:true
  });
  tiles.addTo(portfolioMap);
  portfolioMapLayer=L.layerGroup().addTo(portfolioMap);

  // Paint inferred locations immediately.
  portfolioMapGeometries.clear();
  const bounds=[];
  projects.filter(p=>!p.rejected).forEach(p=>{
   try{
    const x=inferGeo(p);
    const lat=Number(x.lat),lon=Number(x.lon);
    if(!Number.isFinite(lat)||!Number.isFinite(lon))return;
    const g={type:"point",source:"inferred",referencePoints:[{lat,lon}]};
    portfolioMapGeometries.set(p.id,g);
    L.circleMarker([lat,lon],{radius:5,weight:1.5,fillOpacity:.75})
      .bindPopup(`<b>${traceEscape(p.id)}</b><br>${traceEscape(p.name)}`)
      .addTo(portfolioMapLayer);
    bounds.push([lat,lon]);
   }catch(e){}
  });

  if(bounds.length)portfolioMap.fitBounds(bounds,{padding:[25,25],maxZoom:10});
  portfolioMap.invalidateSize(true);
  portfolioMapReady=true;

  if(state)state.innerHTML=`Mapa cargado · ${bounds.length} proyectos localizados. No está filtrando el listado.`;

  // Enrich asynchronously; the basemap is already visible.
  preloadPortfolioUserGeometries().then(async()=>{
   if(!portfolioMapReady)return;
   await refreshPortfolioMapLayers();
   portfolioMap.invalidateSize(false);
  }).catch(e=>console.warn("Geometrías de usuario no cargadas",e));

  portfolioMap.on("zoomend moveend",()=>{
   clearTimeout(portfolioMapMoveTimer);
   portfolioMapMoveTimer=setTimeout(()=>{
    if(portfolioMapFilterEnabled)applyPortfolioMapFilter();
   },350);
  });
 }catch(e){
  console.error("Error mapa portafolio",e);
  portfolioMapReady=false;
  if(state)state.innerHTML=`<b>No fue posible cargar el mapa.</b> ${traceEscape(e.message||String(e))}`;
  // Visible retry instead of a permanently blank box.
  el.innerHTML='<div class="portfolio-map-retry"><b>El mapa no terminó de inicializarse.</b><br><button type="button" onclick="retryPortfolioMap()">Reintentar cargar mapa</button></div>';
 }finally{
  portfolioMapInitializing=false;
 }
}
function retryPortfolioMap(){
 portfolioMapReady=false;portfolioMapInitializing=false;
 if(portfolioMap){try{portfolioMap.remove()}catch(e){} portfolioMap=null;}
 const el=document.getElementById("portfolioProjectsMap");if(el)el.innerHTML="";
 const state=document.getElementById("portfolioMapFilterState");if(state)state.innerHTML="Reintentando cargar mapa…";
 schedulePortfolioMapInit();
 if(window.bootPortfolioMapIndependent)setTimeout(window.bootPortfolioMapIndependent,80);
}
window.retryPortfolioMap=retryPortfolioMap;
window.disablePortfolioMapFilter=disablePortfolioMapFilter;


function loadLeaflet(){
 return new Promise((resolve,reject)=>{
  if(window.L){resolve(window.L);return;}
  if(!document.getElementById("leaflet-css")){
   const c=document.createElement("link");c.id="leaflet-css";c.rel="stylesheet";
   c.href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css";document.head.appendChild(c);
  }

  // Remove a stale script tag left by a previous HUB version.
  const stale=document.getElementById("leaflet-js");
  if(stale&&!window.L){try{stale.remove()}catch(e){}}

  const sources=[
   "https://unpkg.com/leaflet@1.9.4/dist/leaflet.js",
   "https://cdn.jsdelivr.net/npm/leaflet@1.9.4/dist/leaflet.js"
  ];
  let i=0;
  const tryNext=()=>{
   if(window.L){resolve(window.L);return;}
   if(i>=sources.length){reject(new Error("No fue posible cargar la librería del mapa."));return;}
   const s=document.createElement("script");s.id="leaflet-js";s.src=sources[i++];
   let finished=false;
   const timer=setTimeout(()=>{
    if(finished)return;finished=true;
    try{s.remove()}catch(e){}
    tryNext();
   },8000);
   s.onload=()=>{
    if(finished)return;finished=true;clearTimeout(timer);
    if(window.L)resolve(window.L);else{try{s.remove()}catch(e){};tryNext();}
   };
   s.onerror=()=>{
    if(finished)return;finished=true;clearTimeout(timer);
    try{s.remove()}catch(e){};tryNext();
   };
   document.head.appendChild(s);
  };
  tryNext();
 });
}
function geometryCenter(g){
 const pts=g?.referencePoints||[];if(!pts.length)return {lat:9.93,lon:-84.08};
 return {lat:pts.reduce((a,p)=>a+Number(p.lat),0)/pts.length,lon:pts.reduce((a,p)=>a+Number(p.lon),0)/pts.length};
}
function updateGeometryStatus(){
 const b=document.getElementById("geometryStatus");if(!b)return;const g=pendingProjectGeometry;
 if(!g){b.innerHTML="Seleccione el tipo de ubicación y márquela en el mapa.";return;}
 if(g.type==="point"){const p=g.referencePoints?.[0];b.innerHTML=p?`<b>Ubicación puntual:</b> ${p.lat.toFixed(6)}, ${p.lon.toFixed(6)}`:"Haga clic en el mapa para definir la ubicación.";}
 else{const n=g.referencePoints?.length||0, mode=g.routeMode==="road"?"siguiendo rutas existentes":g.routeMode==="road-choice"?"rutas existentes seleccionadas":g.routeMode==="routing"?"calculando ruta…":g.routeMode==="linear"?"línea recta":"modo de trazado pendiente";b.innerHTML=`<b>Tramo:</b> ${n} punto${n===1?"":"s"} de referencia · ${mode}.`;}
}
function redrawRegistrationGeometry(){
 if(!registrationMap||!window.L)return;
 registrationMapLayers.forEach(x=>{try{registrationMap.removeLayer(x)}catch(e){}});registrationMapLayers=[];
 const g=pendingProjectGeometry;if(!g){updateGeometryStatus();return;}
 (g.referencePoints||[]).forEach((p,i)=>registrationMapLayers.push(L.marker([p.lat,p.lon]).addTo(registrationMap).bindTooltip(g.type==="point"?"Ubicación":`Referencia ${i+1}`)));
 if(g.type==="segment"&&(g.coordinates||[]).length>1)registrationMapLayers.push(L.polyline(g.coordinates.map(c=>[c[1],c[0]]),{weight:5}).addTo(registrationMap));
 if(registrationMapLayers.length){try{registrationMap.fitBounds(L.featureGroup(registrationMapLayers).getBounds().pad(.18),{maxZoom:16})}catch(e){}}
 updateGeometryStatus();
}
async function initRegistrationLocation(existing=null){
 const el=document.getElementById("projectLocationMap");if(!el)return;
 try{
  await loadLeaflet();if(registrationMap){try{registrationMap.remove()}catch(e){}}
  registrationMap=L.map(el).setView([9.93,-84.08],7);
  L.tileLayer("https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png",{maxZoom:19,attribution:"© OpenStreetMap"}).addTo(registrationMap);
  pendingProjectGeometry=existing?JSON.parse(JSON.stringify(existing)):null;
  originalProjectGeometry=existing?JSON.parse(JSON.stringify(existing)):null;
  projectGeometryDirty=false;
  const type=document.getElementById("locationType");if(type&&pendingProjectGeometry)type.value=pendingProjectGeometry.type||"";
  const choice=document.getElementById("segmentRouteChoice"),mode=document.getElementById("segmentRouteMode");
  if(choice)choice.style.display=pendingProjectGeometry?.type==="segment"?"block":"none";
  if(mode&&pendingProjectGeometry?.type==="segment")mode.value=pendingProjectGeometry.routeMode==="road"||pendingProjectGeometry.routeMode==="road-choice"?"road":pendingProjectGeometry.routeMode==="linear"?"linear":"";
  registrationMap.on("click",e=>{
   const t=document.getElementById("locationType")?.value;if(!t){alert("Seleccione primero si la ubicación es puntual o un tramo.");return;}
   if(t==="point")pendingProjectGeometry={type:"point",source:"user",referencePoints:[{lat:e.latlng.lat,lon:e.latlng.lng}]};
   else{
    if(!pendingProjectGeometry||pendingProjectGeometry.type!=="segment")pendingProjectGeometry={type:"segment",source:"user",referencePoints:[],coordinates:[],routeMode:null};
    pendingProjectGeometry.referencePoints.push({lat:e.latlng.lat,lon:e.latlng.lng});
    const selectedMode=document.getElementById("segmentRouteMode")?.value;
    pendingProjectGeometry.coordinates=pendingProjectGeometry.referencePoints.map(p=>[p.lon,p.lat]);
    pendingProjectGeometry.routeMode=selectedMode==="road"?"road-choice":selectedMode==="linear"?"linear":null;
   }
   projectGeometryDirty=true;
   redrawRegistrationGeometry();
   if(t==="segment" && document.getElementById("segmentRouteMode")?.value==="road" && pendingProjectGeometry.referencePoints.length>=2) routeProjectSegment(true);
  });
  redrawRegistrationGeometry();setTimeout(()=>registrationMap.invalidateSize(),100);
 }catch(e){el.innerHTML=`<div class="note">${traceEscape(e.message)}</div>`;}
}
function changeLocationType(){
 const t=document.getElementById("locationType")?.value, choice=document.getElementById("segmentRouteChoice"), mode=document.getElementById("segmentRouteMode");
 if(choice)choice.style.display=t==="segment"?"block":"none";
 if(mode)mode.value="";
 pendingProjectGeometry=t==="point"?{type:"point",source:"user",referencePoints:[]}:t==="segment"?{type:"segment",source:"user",referencePoints:[],coordinates:[],routeMode:null}:null;
 projectGeometryDirty=true;
 redrawRegistrationGeometry();
}
async function changeSegmentRouteMode(){
 const mode=document.getElementById("segmentRouteMode")?.value,g=pendingProjectGeometry;if(!g||g.type!=="segment")return;
 if(mode==="linear"){g.coordinates=g.referencePoints.map(p=>[p.lon,p.lat]);g.routeMode="linear";projectGeometryDirty=true;redrawRegistrationGeometry();}
 else if(mode==="road"){g.routeMode="road-choice";if(g.referencePoints.length>=2)await routeProjectSegment(true);else updateGeometryStatus();}
}
function undoLocationPoint(){const g=pendingProjectGeometry;if(!g)return;if(g.type==="point")g.referencePoints=[];else{g.referencePoints.pop();g.coordinates=g.referencePoints.map(p=>[p.lon,p.lat]);g.routeMode=document.getElementById("segmentRouteMode")?.value==="road"?"road-choice":"linear";}projectGeometryDirty=true;redrawRegistrationGeometry();if(g.type==="segment"&&document.getElementById("segmentRouteMode")?.value==="road"&&g.referencePoints.length>=2)routeProjectSegment(true);}
function clearProjectLocation(){pendingProjectGeometry=null;projectGeometryDirty=true;const t=document.getElementById("locationType");if(t)t.value="";redrawRegistrationGeometry();}
async function routeProjectSegment(silent=false){
 const g=pendingProjectGeometry;
 if(!g||g.type!=="segment"||g.referencePoints.length<2){
  if(!silent)alert("Marque al menos dos puntos de referencia.");
  return;
 }
 const b=document.getElementById("geometryStatus");
 if(b)b.innerHTML="<b>Tramo:</b> buscando la ruta por la red vial existente…";
 try{
  const coords=g.referencePoints.map(p=>`${p.lon},${p.lat}`).join(";");
  const url=`https://router.project-osrm.org/route/v1/driving/${coords}?overview=full&geometries=geojson&steps=false&alternatives=true&continue_straight=false`;
  const r=await fetch(url);
  if(!r.ok)throw new Error("Sin respuesta del servicio de rutas");
  const d=await r.json();
  const routes=(d.routes||[]).filter(x=>x.geometry?.coordinates?.length);
  if(!routes.length)throw new Error("No se encontró ruta vial");
  routes.sort((a,b)=>a.distance-b.distance);
  g.coordinates=routes[0].geometry.coordinates;projectGeometryDirty=true;
  g.routeMode="road";
  g.routeDistanceMeters=routes[0].distance||null;
 }catch(e){
  g.coordinates=g.referencePoints.map(p=>[p.lon,p.lat]);
  g.routeMode="linear";
  if(!silent)alert("No se encontró una conexión por carretera entre esos puntos. Se mantendrá el trazado lineal.");
 }
 redrawRegistrationGeometry();
}
window.changeLocationType=changeLocationType;window.changeSegmentRouteMode=changeSegmentRouteMode;window.undoLocationPoint=undoLocationPoint;window.clearProjectLocation=clearProjectLocation;window.routeProjectSegment=routeProjectSegment;

const INSTITUTION_CODES={MOPT:"M",CONAVI:"V",COSEVI:"S",INCOFER:"F",DGAC:"D",INCOP:"I",JAPDEVA:"J",AYA:"A",AyA:"A",CTP:"T",CNC:"C"};
const FORM_DROPDOWNS={
 "Estado":["Abierta","Activo","Aprobado","Descartada","En atención","Suspendida","Transferida","Otro"],
 "Etapa":["Factibilidad","Idea","Perfil","Prefactibilidad","Diseño"],
 "¿Pertenece a algún programa?":["Sí","No"],
 "Tipo de Obra":["Carretera","Ciclovía","Edificación","Infraestructura Aeroportuaria","Infraestructura Ferroviaria","Infraestructura Logística","Infraestructura obras hidráulicas","Infraestructura Portuaria","Intercambio","Otro","Puente","Seguridad Vial","Transporte Público"],
 "Sub tipo de obra":["Aeródromo","Aeropuerto","Pavimento asfáltico","Edificios","Embarcaderos","Equipamiento","Intermodal","Lastre","Marinas y T. Cruceros","O&M","Otras obras","Puentes","Puentes y Pasos","RVC","Seguridad vial","Servicios","Talleres","Taludes y otros","Terminales","Tren de Carga","Tren de Pasajeros","Tren Mixto","Otro","Pavimento concreto"],
 "Fuente / Modalidad de financiamiento":["APP","Bono","Cooperación Internacional","Fideicomiso","Otro","Préstamo BCIE","Préstamo BID","Préstamo CAF","Presupuesto Nacional","Recursos Propios"]
};
const AUTO_HEADERS=["Referencia de fila","ID HUB","Fecha actualización","Estado HUB"];
function hnorm(v){return String(v||"").trim().toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g,"").replace(/[^a-z0-9]/g,"");}
function headerIndex(name){const n=hnorm(name);return excelHeaders.findIndex(h=>hnorm(h)===n);}
function cell(row,name){const i=headerIndex(name);return i>=0?(row.values?.[0]?.[i]??""):"";}
async function resolveExcelWorkbook(){
 if(excelBook)return excelBook;
 const item=await graphFetch(`/drives/${encodeURIComponent(MS_CONFIG.repositoryDriveId)}/items/${encodeURIComponent(MS_CONFIG.excelItemId)}?$select=id,name,webUrl,parentReference,file`);
 excelBook={driveId:MS_CONFIG.repositoryDriveId,itemId:MS_CONFIG.excelItemId,name:item.name,webUrl:item.webUrl};
 return excelBook;
}
async function resolveExcelTable(){
 if(excelTable&&excelHeaders.length)return excelTable;
 const b=await resolveExcelWorkbook();
 const tables=await graphFetch(`/drives/${b.driveId}/items/${b.itemId}/workbook/tables?$select=id,name`);
 if(!(tables.value||[]).length)throw new Error("El Excel maestro debe tener sus datos convertidos en una Tabla de Excel (Insertar > Tabla).");
 excelTable=tables.value[0];
 const hr=await graphFetch(`/drives/${b.driveId}/items/${b.itemId}/workbook/tables/${encodeURIComponent(excelTable.name)}/headerRowRange`);
 excelHeaders=(hr.values?.[0]||[]).map(String);
 if(headerIndex("ID HUB")<0)throw new Error("La tabla del Excel no contiene la columna ID HUB.");
 return excelTable;
}
async function getExcelRows(){const b=await resolveExcelWorkbook();const t=await resolveExcelTable();const r=await graphFetch(`/drives/${b.driveId}/items/${b.itemId}/workbook/tables/${encodeURIComponent(t.name)}/rows`);return r.value||[];}
function excelRowToProject(r){
 const id=String(cell(r,"ID HUB")||"").trim();
 return {id,name:cell(r,"Nombre del proyecto")||cell(r,"Nombre corto")||"Sin nombre",inst:cell(r,"Institución")||"Sin institución",step:1,state:cell(r,"Estado")||"Borrador",priority:cell(r,"Estado previo")||"",bpip:cell(r,"Código del BPIP")||"Por definir",categoria:cell(r,"Tipo de Obra")||"",subcategoria:cell(r,"Sub tipo de obra")||"",etapa:cell(r,"Etapa")||"Idea",owner:"Técnico / Profesional",returned:false,rejected:false,excelIndex:r.index,excelData:Object.fromEntries(excelHeaders.map((h,i)=>[h,r.values?.[0]?.[i]??""]))};
}
async function loadProjectsFromExcel(){
 const response=await hubApiFetch("/projects");
 const rows=Array.isArray(response)?response:(response?.projects||response?.value||response?.data||[]);
 projects=rows.map((raw,index)=>{
  const data=raw?.excelData||raw?.data||raw;
  const pick=(...names)=>{for(const n of names){if(data&&data[n]!==undefined&&data[n]!==null&&String(data[n]).trim()!=="")return data[n];if(raw&&raw[n]!==undefined&&raw[n]!==null&&String(raw[n]).trim()!=="")return raw[n];}return "";};
  const id=String(pick("ID HUB","id","idHub","projectId")||"").trim();
  return {id,name:pick("Nombre del proyecto","name","nombreProyecto","Nombre corto")||"Sin nombre",inst:pick("Institución","inst","institution","institucion")||"Sin institución",step:Number(pick("step","paso"))||1,state:pick("Estado","state","estado")||"Borrador",priority:pick("Estado (PREVIO, REVISAR)","Estado previo","priority")||"",bpip:pick("Código del BPIP","bpip","codigoBPIP")||"Por definir",categoria:pick("Tipo de Obra","categoria","tipoObra")||"",subcategoria:pick("Sub tipo de obra","subcategoria","subTipoObra")||"",etapa:pick("Etapa","etapa")||"Idea",owner:pick("owner","responsable")||"Técnico / Profesional",returned:!!pick("returned","devuelto"),rejected:!!pick("rejected","rechazado"),excelIndex:index,excelData:data||{}};
 }).filter(p=>p.id);
 activeId=projects[0]?.id||"";
 populateSelects();applyGlobalFilters();renderFlow();renderKpis();renderMatrix();renderDetail();renderList();renderInbox();renderNotifications();renderRegistrationForm();
 return projects;
}
function currentYear2(){return String(new Date().getFullYear()).slice(-2);}
async function generateHubId(institution){
 const code=INSTITUTION_CODES[String(institution||"").trim()]||INSTITUTION_CODES[String(institution||"").trim().toUpperCase()];
 if(!code)throw new Error("No existe código de ID HUB configurado para la institución: "+institution);
 const yy=currentYear2(), rows=await getExcelRows(); let max=0;
 const re=new RegExp(`^${code}-${yy}-(\\d+)$`,`i`);
 rows.forEach(r=>{const m=String(cell(r,"ID HUB")||"").trim().match(re);if(m)max=Math.max(max,Number(m[1])||0);});
 return `${code}-${yy}-${String(max+1).padStart(3,"0")}`;
}
function formValue(header){return document.querySelector(`[data-excel-header="${CSS.escape(header)}"]`)?.value?.trim?.()||"";}
function renderRegistrationForm(data={}){
 const box=document.getElementById("excelProjectForm");if(!box||!excelHeaders.length)return;
 const fields=excelHeaders.filter(h=>!AUTO_HEADERS.some(a=>hnorm(a)===hnorm(h)));
 box.innerHTML=fields.map(h=>{const val=data[h]??"";const opts=FORM_DROPDOWNS[h];
  if(opts)return `<div><label>${h}</label><select data-excel-header="${h}"><option value="">Seleccione…</option>${opts.map(o=>`<option ${String(val)===o?'selected':''}>${o}</option>`).join("")}</select></div>`;
  if(hnorm(h)===hnorm("Institución")){const opts=["MOPT","CONAVI","COSEVI","INCOFER","DGAC","INCOP","JAPDEVA","AYA","CTP","CNC"];return `<div><label>${h} *</label><select data-excel-header="${h}" required><option value="">Seleccione…</option>${opts.map(o=>`<option ${String(val).toUpperCase()===o?'selected':''}>${o}</option>`).join("")}</select></div>`;}
  const long=/nombre del proyecto|coment|observ|planes/i.test(h);return `<div><label>${h}${/nombre del proyecto/i.test(h)?' *':''}</label>${long?`<textarea data-excel-header="${h}" rows="3">${String(val).replace(/</g,"&lt;")}</textarea>`:`<input data-excel-header="${h}" value="${String(val).replace(/"/g,"&quot;")}">`}</div>`;
 }).join("");
 box.insertAdjacentHTML("beforeend",`<div class="location-editor">
  <div><h3>Ubicación del proyecto</h3><p class="muted">Indique si es puntual o un tramo y márquelo directamente sobre el mapa.</p></div>
  <div class="location-controls">
   <div><label>Tipo de ubicación *</label><select id="locationType" onchange="changeLocationType()"><option value="">Seleccione…</option><option value="point">Puntual</option><option value="segment">Tramo</option></select></div>
   <div id="segmentRouteChoice" style="display:none"><label>¿Cómo desea trazar el tramo? *</label><select id="segmentRouteMode" onchange="changeSegmentRouteMode()"><option value="">Seleccione…</option><option value="road">Calcar rutas existentes</option><option value="linear">Línea recta entre puntos</option></select></div>
   <div class="location-help"><b>Puntual:</b> marque un punto.<br><b>Tramo:</b> marque dos o más referencias en orden y elija si el vector debe seguir las carreteras existentes o unir los puntos en línea recta.</div>
  </div>
  <div id="projectLocationMap" class="project-location-map"></div><div id="geometryStatus" class="note"></div>
  <div class="action-row location-actions"><button type="button" onclick="routeProjectSegment()">Recalcular rutas existentes</button><button type="button" class="secondary" onclick="undoLocationPoint()">Deshacer último punto</button><button type="button" class="secondary" onclick="clearProjectLocation()">Limpiar ubicación</button></div>
 </div>`);
 setTimeout(()=>initRegistrationLocation(window.__editingGeometry||null),0);
}
function buildExcelValues(id,existing={}){
 const now=new Date().toISOString();
 return excelHeaders.map(h=>{
  if(hnorm(h)===hnorm("ID HUB"))return id;
  if(hnorm(h)===hnorm("Fecha actualización"))return now;
  if(hnorm(h)===hnorm("Estado HUB"))return "Incluido";
  if(hnorm(h)===hnorm("Referencia de fila"))return existing[h]||"";
  const el=document.querySelector(`[data-excel-header="${CSS.escape(h)}"]`);return el?el.value.trim():(existing[h]??"");
 });
}
function buildProjectApiPayload(){
 const payload={};
 for(const h of excelHeaders){
  if(["ID HUB","Fecha actualización","Estado HUB","Referencia de fila"].some(x=>hnorm(x)===hnorm(h)))continue;
  const el=document.querySelector(`[data-excel-header="${CSS.escape(h)}"]`);
  payload[h]=el?el.value.trim():"";
 }
 return payload;
}
async function appendExcelProject(values){
 const b=await resolveExcelWorkbook(),t=await resolveExcelTable();
 return graphFetch(`/drives/${b.driveId}/items/${b.itemId}/workbook/tables/${encodeURIComponent(t.name)}/rows/add`,{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({index:null,values:[values]})});
}
function excelColumnName(n){let out="";for(let x=n+1;x>0;x=Math.floor((x-1)/26))out=String.fromCharCode(65+((x-1)%26))+out;return out;}
async function patchExcelRow(index,values){
 const b=await resolveExcelWorkbook(),t=await resolveExcelTable();
 // Evita PATCH /rows/{index}, que en algunos libros de Excel Online devuelve ApiNotFound.
 // Calculamos el rango exacto de la fila dentro de la tabla y actualizamos ese rango.
 const tr=await graphFetch(`/drives/${b.driveId}/items/${b.itemId}/workbook/tables/${encodeURIComponent(t.name)}/range`);
 const address=String(tr.address||"");
 const bang=address.lastIndexOf("!");
 if(bang<0)throw new Error("No se pudo determinar la hoja de la tabla de Excel.");
 let sheet=address.slice(0,bang);
 if(sheet.startsWith("'")&&sheet.endsWith("'"))sheet=sheet.slice(1,-1).replace(/''/g,"'");
 const firstCol=Number(tr.columnIndex||0);
 const colCount=Number(tr.columnCount||excelHeaders.length);
 const firstDataRow=Number(tr.rowIndex||0)+2; // Excel es 1-based y la primera fila de la tabla es encabezado.
 const rowNumber=firstDataRow+Number(index);
 const a1=`${excelColumnName(firstCol)}${rowNumber}:${excelColumnName(firstCol+colCount-1)}${rowNumber}`;
 const safeAddress=a1.replace(/'/g,"''");
 return graphFetch(`/drives/${b.driveId}/items/${b.itemId}/workbook/worksheets/${encodeURIComponent(sheet)}/range(address='${safeAddress}')`,{method:"PATCH",headers:{"Content-Type":"application/json"},body:JSON.stringify({values:[values]})});
}
async function editProjectExcel(id){
 const rows=await getExcelRows(),r=rows.find(x=>String(cell(x,"ID HUB")).trim()===id);if(!r)return alert("No se encontró el proyecto en Excel.");
 const data=Object.fromEntries(excelHeaders.map((h,i)=>[h,r.values?.[0]?.[i]??""])),p=projects.find(x=>x.id===id);
 let g=await loadProjectGeometry(id);
 if(!g&&p){const x=inferGeo(p);g={type:"point",source:"inferred",referencePoints:[{lat:x.lat,lon:x.lon}],inferredFrom:x.fuente};}
 window.__editingGeometry=g;renderRegistrationForm(data);showView("registrar");
 const btn=document.getElementById("saveExcelProjectBtn");btn.textContent="Guardar cambios";btn.dataset.editId=id;
}
async function softDeleteProjectExcel(id){
 if(!confirm(`¿Marcar ${id} como Eliminado? El registro se conservará en Excel.`))return;
 const rows=await getExcelRows(),r=rows.find(x=>String(cell(x,"ID HUB")).trim()===id);if(!r)return alert("No se encontró el proyecto en Excel.");
 const vals=[...(r.values?.[0]||[])];while(vals.length<excelHeaders.length)vals.push("");
 const si=headerIndex("Estado HUB"),fi=headerIndex("Fecha actualización");if(si>=0)vals[si]="Eliminado";if(fi>=0)vals[fi]=new Date().toISOString();
 await patchExcelRow(r.index,vals);await loadProjectsFromExcel();alert(`${id} quedó marcado como Eliminado.`);
}
window.editProjectExcel=editProjectExcel;window.softDeleteProjectExcel=softDeleteProjectExcel;window.loadProjectsFromExcel=loadProjectsFromExcel;

async function ensureGraphIdentity(){
 if(!graphUser)graphUser=await graphFetch("/me?$select=displayName,userPrincipalName,mail,id");
 return graphUser;
}
async function resolveRootFolder(){
 if(graphRoot)return graphRoot;
 const item=await graphFetch(`/drives/${encodeURIComponent(MS_CONFIG.repositoryDriveId)}/items/${encodeURIComponent(MS_CONFIG.rootItemId)}?$select=id,name,webUrl,parentReference,folder`);
 if(!item?.folder)throw new Error("HUB_Proyectos_MOPT no corresponde a una carpeta accesible.");
 graphRoot={driveId:MS_CONFIG.repositoryDriveId,itemId:MS_CONFIG.rootItemId,name:item.name,webUrl:item.webUrl};
 return graphRoot;
}

async function resolveGeometryFolder(){
 const root=await resolveRootFolder(),children=await graphFetch(`/drives/${root.driveId}/items/${root.itemId}/children?$select=id,name,folder&$top=999`);
 let f=(children.value||[]).find(x=>x.folder&&x.name==="_HUB_GEOMETRIA");
 if(!f)f=await graphFetch(`/drives/${root.driveId}/items/${root.itemId}/children`,{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({name:"_HUB_GEOMETRIA",folder:{},"@microsoft.graph.conflictBehavior":"fail"})});
 return {driveId:root.driveId,itemId:f.id};
}

let shpWritePromise=null;
function loadShpWrite(){
 if(window.shpwrite)return Promise.resolve(window.shpwrite);
 if(shpWritePromise)return shpWritePromise;
 shpWritePromise=new Promise((resolve,reject)=>{
  const sources=[
   "https://unpkg.com/@mapbox/shp-write@0.4.3/shpwrite.js",
   "https://cdn.jsdelivr.net/npm/@mapbox/shp-write@0.4.3/shpwrite.js"
  ];
  let i=0;
  const next=()=>{
   if(window.shpwrite){resolve(window.shpwrite);return;}
   if(i>=sources.length){reject(new Error("No fue posible cargar el generador de Shapefile."));return;}
   const s=document.createElement("script");s.src=sources[i++];s.async=true;
   const timer=setTimeout(()=>{try{s.remove()}catch(e){};next();},10000);
   s.onload=()=>{clearTimeout(timer);window.shpwrite?resolve(window.shpwrite):next();};
   s.onerror=()=>{clearTimeout(timer);try{s.remove()}catch(e){};next();};
   document.head.appendChild(s);
  };
  next();
 });
 return shpWritePromise;
}
async function ensureChildFolder(parentDriveId,parentItemId,name){
 const r=await graphFetch(`/drives/${parentDriveId}/items/${parentItemId}/children?$select=id,name,folder&$top=999`);
 let f=(r.value||[]).find(x=>x.folder&&x.name.toLowerCase()===name.toLowerCase());
 if(!f)f=await graphFetch(`/drives/${parentDriveId}/items/${parentItemId}/children`,{
  method:"POST",headers:{"Content-Type":"application/json"},
  body:JSON.stringify({name,folder:{},"@microsoft.graph.conflictBehavior":"fail"})
 });
 return {driveId:parentDriveId,itemId:f.id,name:f.name};
}
async function resolveShapeFolder(type){
 const root=await resolveGeometryFolder();
 const name=type==="point"?"Shape Punto":"Shape Lineal";
 return ensureChildFolder(root.driveId,root.itemId,name);
}
function geometryToGeoJSONFeature(id,g){
 const p=projects.find(x=>x.id===id)||{};
 const props={
  ID_HUB:String(id||""),
  BPIP:String(p.bpip||""),
  NOMBRE:String(p.name||"").slice(0,250),
  INSTITUC:String(p.dependency||"").slice(0,100),
  EDITADO:new Date().toISOString(),
  USUARIO:String(actorLabel()||"").slice(0,100)
 };
 if(g.type==="point"){
  const q=g.referencePoints?.[0];
  return {type:"Feature",properties:props,geometry:{type:"Point",coordinates:[Number(q.lon),Number(q.lat)]}};
 }
 return {type:"Feature",properties:props,geometry:{type:"LineString",coordinates:(g.coordinates||[]).map(c=>[Number(c[0]),Number(c[1])])}};
}
function shapeTimestamp(d=new Date()){
 const pad=n=>String(n).padStart(2,"0");
 return `${d.getFullYear()}-${pad(d.getMonth()+1)}-${pad(d.getDate())}_${pad(d.getHours())}-${pad(d.getMinutes())}-${pad(d.getSeconds())}`;
}
async function saveGeometryShapefile(id,g){
 const shp=await loadShpWrite();
 const feature=geometryToGeoJSONFeature(id,g);
 const fc={type:"FeatureCollection",features:[feature]};
 const base=`${safeOneDriveName(id)}_${shapeTimestamp()}`;
 const options={
  folder:base,
  filename:base,
  outputType:"blob",
  compression:"DEFLATE",
  types:{point:base,polyline:base,polygon:base}
 };
 let data=await shp.zip(fc,options);
 if(data instanceof Promise)data=await data;
 const blob=data instanceof Blob?data:new Blob([data],{type:"application/zip"});
 const folder=await resolveShapeFolder(g.type);
 await graphFetch(`/drives/${folder.driveId}/items/${folder.itemId}:/${encodeURIComponent(base+".zip")}:/content`,{
  method:"PUT",
  headers:{"Content-Type":"application/zip"},
  body:blob
 });
 return base+".zip";
}

async function resolveGeometryHistoryFolder(){
 const root=await resolveRootFolder(),children=await graphFetch(`/drives/${root.driveId}/items/${root.itemId}/children?$select=id,name,folder&$top=999`);
 let f=(children.value||[]).find(x=>x.folder&&x.name==="_HUB_GEOMETRIA_HISTORIAL");
 if(!f)f=await graphFetch(`/drives/${root.driveId}/items/${root.itemId}/children`,{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({name:"_HUB_GEOMETRIA_HISTORIAL",folder:{},"@microsoft.graph.conflictBehavior":"fail"})});
 return {driveId:root.driveId,itemId:f.id};
}
function validProjectGeometry(g){
 if(!g||!["point","segment"].includes(g.type)||!Array.isArray(g.referencePoints))return false;
 if(g.type==="point")return g.referencePoints.length>=1&&Number.isFinite(Number(g.referencePoints[0]?.lat))&&Number.isFinite(Number(g.referencePoints[0]?.lon));
 return g.referencePoints.length>=2&&Array.isArray(g.coordinates)&&g.coordinates.length>=2;
}
async function saveProjectGeometry(id,g){
 if(!validProjectGeometry(g))throw new Error("La geometría no es válida. Se conserva la ubicación anterior.");
 const f=await resolveGeometryFolder();
 const filename=safeOneDriveName(id)+".json";
 const payload={projectId:id,updatedAt:new Date().toISOString(),updatedBy:actorLabel(),...g};

 // Read current version first. If it exists, keep an immutable backup before overwrite.
 let previous=null;
 try{previous=await loadProjectGeometry(id);}catch(e){}
 if(previous&&validProjectGeometry(previous)){
  try{
   const h=await resolveGeometryHistoryFolder();
   const stamp=new Date().toISOString().replace(/[:.]/g,"-");
   const backupName=`${safeOneDriveName(id)}_${stamp}.json`;
   await graphFetch(`/drives/${h.driveId}/items/${h.itemId}:/${encodeURIComponent(backupName)}:/content`,{
    method:"PUT",headers:{"Content-Type":"application/json"},body:JSON.stringify(previous,null,2)
   });
  }catch(e){console.warn("No se pudo crear respaldo histórico de geometría",e);}
 }

 await graphFetch(`/drives/${f.driveId}/items/${f.itemId}:/${encodeURIComponent(filename)}:/content`,{
  method:"PUT",headers:{"Content-Type":"application/json"},body:JSON.stringify(payload,null,2)
 });

 // Every user geometry edition also creates a dated Shapefile snapshot.
 // It is stored as ZIP because a Shapefile is a set of files (.shp, .shx, .dbf, .prj).
 try{
  payload.shapeSnapshot=await saveGeometryShapefile(id,payload);
 }catch(e){
  console.error("No se pudo crear el Shapefile de la geometría",e);
  throw new Error(`La geometría JSON fue guardada, pero no se pudo crear su Shapefile: ${e.message}`);
 }

 // Refresh portfolio cache immediately so the new geometry cannot be replaced visually by stale inferred data.
 portfolioMapGeometries.set(id,payload);
 return payload;
}
async function loadProjectGeometry(id){
 try{const f=await resolveGeometryFolder(),r=await graphFetch(`/drives/${f.driveId}/items/${f.itemId}/children?$select=id,name,file&$top=999`),x=(r.value||[]).find(v=>v.file&&v.name===safeOneDriveName(id)+".json");return x?await graphFetch(`/drives/${f.driveId}/items/${x.id}/content`):null;}catch(e){console.warn("Geometría no disponible",e);return null;}
}

async function ensureProjectFolder(project){
 const cached=projectFolderCache.get(project.id); if(cached)return cached;
 const root=await resolveRootFolder(), wanted=projectFolderName(project);
 const children=await graphFetch(`/drives/${root.driveId}/items/${root.itemId}/children?$select=id,name,webUrl,folder&$top=999`);
 let folder=(children.value||[]).find(x=>x.folder&&x.name===wanted);
 if(!folder){
  folder=await graphFetch(`/drives/${root.driveId}/items/${root.itemId}/children`,{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({name:wanted,folder:{},"@microsoft.graph.conflictBehavior":"rename"})});
 }
 const out={driveId:root.driveId,itemId:folder.id,name:folder.name,webUrl:folder.webUrl}; projectFolderCache.set(project.id,out); return out;
}
// HUB OneDrive integration build 2026-09-16b
window.HUB_ONEDRIVE_BUILD = "2026-09-16c-trace-ui";
window.ensureProjectFolder = ensureProjectFolder;
window.syncProjectFolderById = async function(projectId){
  const p=projects.find(x=>x.id===projectId);
  if(!p) throw new Error("No se encontró el proyecto "+projectId+" en el HUB.");
  const folder=await ensureProjectFolder(p);
  await writeTrace(p,"SINCRONIZO_CARPETA_PROYECTO",folder.name,{folderId:folder.itemId});
  return folder;
};

async function uploadOneDriveFile(project,file,itemId=null){
 let folder=await ensureProjectFolder(project);
 let path=itemId?`/drives/${folder.driveId}/items/${itemId}/content`:`/drives/${folder.driveId}/items/${folder.itemId}:/${encodeURIComponent(file.name)}:/content`;
 try{
  return await graphFetch(path,{method:"PUT",headers:{"Content-Type":file.type||"application/octet-stream"},body:file});
 }catch(e){
  // A stale OneDrive item/folder id can return 404. Resolve the root/folder again and retry once.
  if(!String(e?.message||e).includes("404")||itemId)throw e;
  rootFolderCache=null;
  folder=await ensureProjectFolder(project);
  path=`/drives/${folder.driveId}/items/${folder.itemId}:/${encodeURIComponent(file.name)}:/content`;
  return await graphFetch(path,{method:"PUT",headers:{"Content-Type":file.type||"application/octet-stream"},body:file});
 }
}
async function writeTrace(project,action,documentName,extra={}){
 try{
  const root=await resolveRootFolder(); const user=await ensureGraphIdentity();
  let traceFolder;
  const children=await graphFetch(`/drives/${root.driveId}/items/${root.itemId}/children?$select=id,name,folder&$top=999`);
  traceFolder=(children.value||[]).find(x=>x.folder&&x.name==="_HUB_TRAZABILIDAD");
  if(!traceFolder)traceFolder=await graphFetch(`/drives/${root.driveId}/items/${root.itemId}/children`,{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({name:"_HUB_TRAZABILIDAD",folder:{},"@microsoft.graph.conflictBehavior":"fail"})});
  // Organiza la evidencia técnica por proyecto para que OneDrive quede limpio.
  const traceChildren=await graphFetch(`/drives/${root.driveId}/items/${traceFolder.id}/children?$select=id,name,folder&$top=999`);
  const projectTraceName=safeOneDriveName(project.id);
  let projectTrace=(traceChildren.value||[]).find(x=>x.folder&&x.name===projectTraceName);
  if(!projectTrace)projectTrace=await graphFetch(`/drives/${root.driveId}/items/${traceFolder.id}/children`,{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({name:projectTraceName,folder:{},"@microsoft.graph.conflictBehavior":"fail"})});
  const event={fecha:new Date().toISOString(),proyectoId:project.id,proyecto:project.name,accion:action,documento:documentName,usuarioHUB:currentUser?.username||null,nombreHUB:currentUser?.name||null,perfil:currentRole,usuarioMicrosoft:user.userPrincipalName||user.mail||null,nombreMicrosoft:user.displayName||null,...extra};
  const filename=`${new Date().toISOString().replace(/[:.]/g,"-")}_${Math.random().toString(36).slice(2,8)}.json`;
  await graphFetch(`/drives/${root.driveId}/items/${projectTrace.id}:/${encodeURIComponent(filename)}:/content`,{method:"PUT",headers:{"Content-Type":"application/json"},body:JSON.stringify(event,null,2)});
 }catch(e){console.warn("No se pudo escribir la trazabilidad en OneDrive",e);}
}

async function writeManagementHistoryDocument(project){
 try{
  const events=await readProjectTrace(project);
  const flow=events.filter(e=>e.accion==="GESTION_FLUJO").sort((a,b)=>String(a.fecha||"").localeCompare(String(b.fecha||"")));
  const esc=v=>traceEscape(v??"");
  const fmt=v=>{try{return new Date(v).toLocaleString("es-CR",{dateStyle:"medium",timeStyle:"short"})}catch(e){return v||"—"}};
  const timeline=flow.map((e,i)=>{
   const from=steps.find(s=>s.n==Number(e.pasoOrigen));
   const to=steps.find(s=>s.n==Number(e.pasoDestino));
   const actor=e.nombreHUB||e.usuarioHUB||"Usuario HUB";
   const comment=(e.comentario||"").trim();
   return `<section class="event">
    <div class="dot">${i+1}</div>
    <div class="card">
     <div class="event-head"><div><span class="step">Movimiento ${i+1}</span><h2>${esc(e.descripcion||"Gestión del proyecto")}</h2></div><time>${esc(fmt(e.fechaGestion||e.fecha))}</time></div>
     <div class="grid">
      <div><label>Gestionado por</label><strong>${esc(actor)}</strong></div>
      <div><label>Perfil</label><strong>${esc(e.perfil||from?.role||"—")}</strong></div>
      <div><label>Recibido en</label><strong>Paso ${esc(e.pasoOrigen||"—")}/${TOTAL_WORKFLOW_STEPS} · ${esc(from?.title||e.estadoOrigen||"—")}</strong></div>
      <div><label>Enviado a</label><strong>Paso ${esc(e.pasoDestino||"—")}/${TOTAL_WORKFLOW_STEPS} · ${esc(to?.role||"—")}</strong></div>
     </div>
     <div class="comment"><label>Comentarios de la gestión</label><p>${comment?esc(comment):'<span class="muted">Sin comentarios registrados.</span>'}</p></div>
     <div class="result"><b>Estado resultante:</b> ${esc(e.estadoDestino||"—")}</div>
    </div>
   </section>`;
  }).join("");

  const current=steps.find(s=>s.n==Number(project.step));
  const html=`<!doctype html><html lang="es"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
  <title>Línea de tiempo · ${esc(project.id)}</title>
  <style>
   *{box-sizing:border-box}body{margin:0;background:#f3f7fb;color:#17233b;font-family:Arial,Helvetica,sans-serif}
   .hero{background:linear-gradient(120deg,#08477f,#087f88);color:#fff;padding:30px 42px}.hero small{letter-spacing:2px;font-weight:700;opacity:.85}.hero h1{margin:7px 0 4px;font-size:28px}.hero p{margin:0;opacity:.9}
   .wrap{max-width:1100px;margin:24px auto;padding:0 24px}.summary{display:grid;grid-template-columns:repeat(4,1fr);gap:12px;margin-bottom:28px}.summary div{background:#fff;border:1px solid #dbe5ef;border-radius:14px;padding:15px}.summary label,.card label{display:block;font-size:11px;letter-spacing:1px;text-transform:uppercase;color:#64748b;font-weight:700;margin-bottom:5px}.summary strong{font-size:15px}
   .timeline{position:relative;padding-left:62px}.timeline:before{content:"";position:absolute;left:24px;top:12px;bottom:12px;width:3px;background:#c9d9e8}
   .event{position:relative;margin:0 0 22px}.dot{position:absolute;left:-62px;top:4px;width:50px;height:50px;border-radius:50%;background:#0b4d8c;color:#fff;display:flex;align-items:center;justify-content:center;font-weight:800;border:5px solid #e5eef7}
   .card{background:#fff;border:1px solid #d8e3ee;border-radius:16px;padding:20px 22px;box-shadow:0 4px 14px rgba(20,50,80,.06)}.event-head{display:flex;justify-content:space-between;gap:20px;border-bottom:1px solid #e6edf4;padding-bottom:12px;margin-bottom:14px}.event-head h2{font-size:18px;margin:4px 0 0}.event-head time{white-space:nowrap;color:#475569;font-weight:700}.step{font-size:11px;color:#087f88;text-transform:uppercase;font-weight:800;letter-spacing:1px}
   .grid{display:grid;grid-template-columns:1fr 1fr;gap:14px 24px}.comment{margin-top:16px;background:#f7fafc;border-radius:10px;padding:13px}.comment p{margin:0;line-height:1.5}.muted{color:#94a3b8;font-style:italic}.result{margin-top:12px;padding:10px 13px;border-left:4px solid #087f88;background:#edfafa;border-radius:7px}
   .empty{background:#fff;border:1px solid #d8e3ee;border-radius:16px;padding:30px;text-align:center;color:#64748b}
   .foot{text-align:center;color:#64748b;font-size:12px;margin:28px 0}
   @media(max-width:760px){.summary{grid-template-columns:1fr 1fr}.grid{grid-template-columns:1fr}.event-head{display:block}.event-head time{display:block;margin-top:8px}.hero{padding:24px}.wrap{padding:0 14px}}
  </style></head><body>
   <header class="hero"><small>HUB DE PROYECTOS MOPT</small><h1>Línea de tiempo de gestión</h1><p>${esc(project.id)} · ${esc(project.name)}</p></header>
   <main class="wrap">
    <div class="summary">
     <div><label>ID HUB</label><strong>${esc(project.id)}</strong></div>
     <div><label>BPIP</label><strong>${esc(project.bpip||"—")}</strong></div>
     <div><label>Paso actual</label><strong>${esc(project.step||"—")}/${TOTAL_WORKFLOW_STEPS}</strong></div>
     <div><label>Responsable actual</label><strong>${esc(current?.role||project.owner||"—")}</strong></div>
    </div>
    <div class="timeline">${timeline||'<div class="empty">Este proyecto todavía no registra movimientos de gestión.</div>'}</div>
    <div class="foot">Actualizado automáticamente por el HUB · ${esc(fmt(new Date().toISOString()))}</div>
   </main>
  </body></html>`;

  const folder=await ensureProjectFolder(project);
  // Keep the prior readable history document and add the timeline as an additional report.
  await graphFetch(`/drives/${folder.driveId}/items/${folder.itemId}:/${encodeURIComponent("Línea de tiempo - "+safeOneDriveName(project.id)+".html")}:/content`,{
   method:"PUT",headers:{"Content-Type":"text/html;charset=utf-8"},body:html
  });
 }catch(e){console.warn("No se pudo actualizar la línea de tiempo de gestión",e);}
}

function ensureHubProgressOverlay(){
 let el=document.getElementById("hubOperationProgress");
 if(el)return el;
 el=document.createElement("div");el.id="hubOperationProgress";
 el.innerHTML=`<div class="hub-progress-card"><div class="hub-progress-spinner"></div><h3 id="hubProgressTitle">Procesando…</h3><p id="hubProgressText">Espere un momento.</p><div class="hub-progress-track"><div id="hubProgressBar"></div></div><div id="hubProgressPct">0%</div></div>`;
 document.body.appendChild(el);return el;
}
function showHubOperationProgress(title,message,pct=10){
 const el=ensureHubProgressOverlay();el.classList.add("show");
 document.getElementById("hubProgressTitle").textContent=title;
 document.getElementById("hubProgressText").textContent=message;
 setHubOperationProgress(pct);
}
function setHubOperationProgress(pct,message){
 const n=Math.max(0,Math.min(100,Number(pct)||0));
 const b=document.getElementById("hubProgressBar"),l=document.getElementById("hubProgressPct"),t=document.getElementById("hubProgressText");
 if(b)b.style.width=n+"%";if(l)l.textContent=Math.round(n)+"%";if(message&&t)t.textContent=message;
}
function finishHubOperationProgress(message="Proceso completado"){setHubOperationProgress(100,message);setTimeout(()=>document.getElementById("hubOperationProgress")?.classList.remove("show"),700);}
function failHubOperationProgress(message="No fue posible completar el proceso"){setHubOperationProgress(100,message);setTimeout(()=>document.getElementById("hubOperationProgress")?.classList.remove("show"),1400);}

function traceEscape(v){return String(v??"").replace(/[&<>"']/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]));}
function traceActionLabel(a){return ({SUBIO_DOCUMENTO:"Subió documento",ACTUALIZO_DOCUMENTO:"Actualizó documento",ELIMINO_DOCUMENTO:"Eliminó documento",CREO_PROYECTO_Y_CARPETA:"Creó proyecto y carpeta",SINCRONIZO_CARPETA_PROYECTO:"Sincronizó carpeta"})[a]||String(a||"").replaceAll("_"," ");}
function traceActionIcon(a){return a==="ELIMINO_DOCUMENTO"?"🔴":a==="ACTUALIZO_DOCUMENTO"?"🔵":a==="SUBIO_DOCUMENTO"?"🟢":"⚪";}
async function readProjectTrace(project){
 const root=await resolveRootFolder();
 const rootChildren=await graphFetch(`/drives/${root.driveId}/items/${root.itemId}/children?$select=id,name,folder&$top=999`);
 const traceFolder=(rootChildren.value||[]).find(x=>x.folder&&x.name==="_HUB_TRAZABILIDAD"); if(!traceFolder)return [];
 let files=[];
 const traceChildren=await graphFetch(`/drives/${root.driveId}/items/${traceFolder.id}/children?$select=id,name,file,folder&$top=999`);
 // Formato nuevo: _HUB_TRAZABILIDAD/N-451/*.json
 const projectFolder=(traceChildren.value||[]).find(x=>x.folder&&x.name===safeOneDriveName(project.id));
 if(projectFolder){const r=await graphFetch(`/drives/${root.driveId}/items/${projectFolder.id}/children?$select=id,name,file&$top=999`);files.push(...(r.value||[]).filter(x=>x.file&&x.name.endsWith(".json")));}
 // Compatibilidad con registros antiguos guardados directamente en _HUB_TRAZABILIDAD.
 files.push(...(traceChildren.value||[]).filter(x=>x.file&&x.name.endsWith(".json")&&x.name.includes(`_${safeOneDriveName(project.id)}_`)));
 const events=[];
 for(const f of files.slice(-100)){
  try{const ev=await graphFetch(`/drives/${root.driveId}/items/${f.id}/content`); if(ev&&ev.proyectoId===project.id)events.push(ev);}catch(e){console.warn("No se pudo leer un evento de trazabilidad",e);}
 }
 return events.sort((a,b)=>String(b.fecha||"").localeCompare(String(a.fecha||"")));
}
async function loadTraceHistory(projectId){
 const box=document.getElementById("trace-history-box"); if(!box)return;
 const p=projects.find(x=>x.id===projectId); if(!p)return;
 box.innerHTML='<p class="muted">Cargando trazabilidad documental desde OneDrive…</p>';
 try{
  const events=await readProjectTrace(p);
  if(!events.length){box.innerHTML='<div class="note">Todavía no hay movimientos documentales registrados para este proyecto.</div>';return;}
  box.innerHTML=`<div style="overflow:auto"><table><thead><tr><th>Fecha y hora</th><th>Usuario</th><th>Acción</th><th>Documento</th><th>Detalles</th></tr></thead><tbody>${events.map((e,i)=>`<tr><td>${traceEscape(new Date(e.fecha).toLocaleString())}</td><td><b>${traceEscape(e.nombreHUB||e.usuarioHUB||e.nombreMicrosoft||"Usuario")}</b><br><span class="muted">${traceEscape(e.perfil||"")}</span></td><td>${traceActionIcon(e.accion)} ${traceEscape(traceActionLabel(e.accion))}</td><td>${traceEscape(e.documento||"—")}</td><td><details><summary>Ver detalles</summary><div class="muted" style="padding-top:8px"><b>Cuenta Microsoft:</b> ${traceEscape(e.usuarioMicrosoft||"—")}<br><b>Nombre Microsoft:</b> ${traceEscape(e.nombreMicrosoft||"—")}<br><b>Item ID:</b> ${traceEscape(e.itemId||e.folderId||"—")}</div></details></td></tr>`).join("")}</tbody></table></div>`;
 }catch(e){console.error(e);box.innerHTML=`<div class="note">No fue posible cargar la trazabilidad desde OneDrive: ${traceEscape(e.message)}</div>`;}
}
window.loadTraceHistory=loadTraceHistory;

async function syncProjectDocuments(id){
 const p=projects.find(x=>x.id===id); if(!p)return;
 const folder=await ensureProjectFolder(p);
 const result=await graphFetch(`/drives/${folder.driveId}/items/${folder.itemId}/children?$select=id,name,webUrl,size,lastModifiedDateTime,file,folder&$top=999`);
 const remote=(result.value||[]).filter(x=>x.file).map(x=>({project:id,type:"Soporte cargado",name:x.name,state:"En OneDrive",owner:"Repositorio MOPT",itemId:x.id,driveId:folder.driveId,webUrl:x.webUrl,size:x.size,uploadedAt:x.lastModifiedDateTime}));
 documents=documents.filter(d=>d.project!==id); // Elimina registros locales/antiguos de este proyecto.
 remote.forEach(r=>documents.push(r));
 saveHubState(); renderDocuments();
}
async function openOneDriveDocument(d){
 try{if(d.webUrl){window.open(d.webUrl,"_blank","noopener");return;} alert("Este registro de demostración todavía no tiene un archivo físico en OneDrive.");}catch(e){alert(e.message||"No fue posible abrir el documento.");}
}
async function replaceOneDriveDocument(index){
 const d=documents[index],p=projects.find(x=>x.id===d.project); if(!d||!p||!d.itemId){alert("Este documento aún no está vinculado a un archivo físico en OneDrive.");return;}
 const input=document.createElement("input");input.type="file";input.accept=".pdf,.doc,.docx,.xls,.xlsx,.png,.jpg,.jpeg,.txt";
 input.onchange=async()=>{const file=input.files?.[0];if(!file)return;try{const item=await uploadOneDriveFile(p,file,d.itemId);d.name=item.name;d.webUrl=item.webUrl;d.size=item.size;d.uploadedAt=item.lastModifiedDateTime;d.owner=actorLabel();auditLog.unshift({project:p.id,action:"Actualizó / sobrescribió documento",role:actorLabel(),date:new Date().toLocaleString(),comment:`Se actualizó ${item.name} en OneDrive.`});saveHubState();await writeTrace(p,"ACTUALIZO_DOCUMENTO",item.name,{itemId:item.id});renderDocuments();openManageModal(p.id,true);alert("Nueva versión guardada en OneDrive y registrada en trazabilidad.");}catch(e){console.error(e);alert("No fue posible actualizar el documento en OneDrive: "+e.message);}};input.click();
}
async function deleteOneDriveDocument(index){
 const d=documents[index],p=projects.find(x=>x.id===d.project); if(!d||!p||!d.itemId){alert("Este documento no tiene un archivo físico de OneDrive asociado.");return;}
 if(!confirm(`¿Eliminar ${d.name} de OneDrive? La acción quedará registrada en trazabilidad.`))return;
 try{const folder=await ensureProjectFolder(p);await graphFetch(`/drives/${folder.driveId}/items/${d.itemId}`,{method:"DELETE"});const oldName=d.name;documents.splice(index,1);auditLog.unshift({project:p.id,action:"Eliminó documento",role:actorLabel(),date:new Date().toLocaleString(),comment:`Se eliminó ${oldName} de OneDrive.`});saveHubState();await writeTrace(p,"ELIMINO_DOCUMENTO",oldName,{itemId:d.itemId});renderDocuments();openManageModal(p.id,true);alert("Documento eliminado de OneDrive. La trazabilidad se conservó.");}catch(e){console.error(e);if(String(e.message||"").includes("423")){alert("No se puede eliminar el documento porque está abierto o bloqueado en Microsoft 365. Cierre el archivo en Word, Excel o OneDrive y vuelva a intentarlo.");}else{alert("No fue posible eliminar el documento de OneDrive: "+e.message);}}
}
window.syncProjectDocuments=syncProjectDocuments;window.openOneDriveDocument=openOneDriveDocument;window.replaceOneDriveDocument=replaceOneDriveDocument;window.deleteOneDriveDocument=deleteOneDriveDocument;
const roleInfo={
 tecnico:["Registra iniciativas","Adjunta soportes","Atiende devoluciones","Envía a revisión técnica"],
 jefatura:["Revisa la ficha técnica","Solicita ajustes","Valida alcance preliminar"],
 planeamiento:["Evalúa alineación estratégica","Prioriza dentro de la institución","Revisa disponibilidad de recursos"],
 direccion:["Aprueba institucionalmente","Devuelve para ajustes","Autoriza radicación ante MOPT"],
 coordinador:["Radica ante el MOPT","Controla soportes","Recibe devoluciones del HUB"],
 secretaria:["Clasifica iniciativas","Asigna evaluadores técnicos","Controla requisitos"],
 areas:["Realiza evaluación técnica central","Emite observaciones","Recomienda ajustes"],
 sistema:["Aplica evaluación automatizada","Calcula puntaje multicriterio","Genera ranking preliminar"],
 comite:["Prioriza proyectos","Aprueba, devuelve o rechaza","Define paso a preparación"],
 ejecutor:["Formula proyecto definitivo","Ejecuta y reporta avance","Gestiona cierre"]
};
function phaseOf(step){return step<=5?1:step<=9?2:step<=13?3:4}
function phaseName(n){return n==1?"Fase 1 · Gestión institucional":n==2?"Fase 2 · Evaluación central":n==3?"Fase 3 · Preparación":"Fase 4 · Ejecución y cierre"}
function clsStep(n){return n<=5?"phase1":n<=9?"phase2":n<=13?"phase3":"phase4"}
function dotClass(p,n){if(p.rejected&&n==p.step)return"rejected"; if(p.returned&&n==p.step)return"returned"; if(n<p.step)return"done"; if(n==p.step)return"current"; return"future"}
function norm(x){return (x||"").toString().toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g,"")}
function matches(p,q){
 q=norm(q||"").trim();
 if(!q) return true;

 const text=norm([
   p.id,p.name,p.inst,p.state,p.bpip,p.categoria,p.subcategoria,p.etapa,p.priority
 ].join(" "));

 const words=text.split(/\s+/).filter(Boolean);
 const terms=q.split(/\s+/).filter(Boolean);

 // Busca por coincidencias parciales: todas las palabras escritas deben aparecer
 // en cualquier campo del proyecto, aunque no estén juntas ni sean exactas.
 return terms.every(term=>{
   if(text.includes(term)) return true;
   return words.some(word=>word.includes(term));
 });
}
function populateSelects(){
  projectSelect.innerHTML=projects.map(p=>`<option value="${p.id}">${p.id} · ${p.name.slice(0,60)}</option>`).join("");
  if(typeof fichaSelect!=="undefined" && fichaSelect) fichaSelect.innerHTML=projects.map(p=>`<option value="${p.id}">${p.id} · ${p.name.slice(0,85)}</option>`).join("");
  const insts=[...new Set(projects.map(p=>p.inst).filter(Boolean))].sort();
  const states=[...new Set(projects.map(p=>p.state).filter(Boolean))].sort();
  const cats=[...new Set(projects.map(p=>p.categoria).filter(Boolean))].sort();
  const subcats=[...new Set(projects.map(p=>p.subcategoria).filter(Boolean))].sort();
  const pri=[...new Set(projects.map(p=>p.priority).filter(Boolean))].sort();
  instFilter.innerHTML='<option value="">Todas</option>'+insts.map(x=>`<option>${x}</option>`).join("");
  stateFilter.innerHTML='<option value="">Todos</option>'+states.map(x=>`<option>${x}</option>`).join("");
  globalInst.innerHTML='<option value="">Todas</option>'+insts.map(x=>`<option>${x}</option>`).join("");
  globalState.innerHTML='<option value="">Todos</option>'+states.map(x=>`<option>${x}</option>`).join("");
  globalCat.innerHTML='<option value="">Todas</option>'+cats.map(x=>`<option>${x}</option>`).join("");
  globalSubcat.innerHTML='<option value="">Todas</option>'+subcats.map(x=>`<option>${x}</option>`).join("");
  globalPriority.innerHTML='<option value="">Todas</option>'+pri.map(x=>`<option>${x}</option>`).join("");
  globalStep.innerHTML='<option value="">Todos</option>'+steps.map(s=>`<option value="${s.n}">${s.n}. ${s.title}</option>`).join("");
}
function showView(id,btn){
 if(["bandeja","notificaciones","registrar","documentos"].includes(id)&&currentRole==="visitor"){
   alert("Para acceder a este módulo debe iniciar sesión seleccionando un rol.");
   return;
 }
 document.querySelectorAll(".view").forEach(v=>v.classList.add("hidden"));
 document.getElementById(id).classList.remove("hidden");
 document.querySelectorAll(".tab").forEach(t=>t.classList.remove("active"));
 btn?.classList.add("active");
 if(id==="matriz")renderMatrix(); if(id==="historial")renderDetail(); if(id==="listado")renderList(); if(id==="bandeja")renderInbox(); if(id==="notificaciones")renderNotifications(); if(id==="documentos")renderDocuments();
 if(id==="priorizacion"){ const f=document.getElementById("priorizacionFrame"); setTimeout(()=>{try{f.contentWindow.dispatchEvent(new Event("resize"));}catch(e){}},350); }
 if(id==="priorizacionProyectos"){ const f=document.getElementById("priorizacionProyectosFrame"); setTimeout(()=>{try{f.contentWindow.dispatchEvent(new Event("resize"));}catch(e){}},350); }
}
function previewRole(){
 if(!currentUser){roleCard.classList.add("hidden");return;}
 const items=roleInfo[currentRole]||[];
 const baseRoleName=roleSelect?.options?.[roleSelect.selectedIndex]?.dataset.baseLabel||roleSelect?.options?.[roleSelect.selectedIndex]?.text||currentRole;
 const roleName=`${baseRoleName} (${ROLE_DISPLAY_STEPS[currentRole]||"—"}/${TOTAL_WORKFLOW_STEPS})`;
 roleCard.classList.remove("hidden");
 roleCard.innerHTML=`<h3>${currentUser.name} · Perfil activo: ${roleName}</h3><ul>${items.map(i=>`<li>${i}</li>`).join("")}</ul>`;
}
function refreshRoleAccess(){
 const roleEl=document.getElementById("roleSelect");
 if(roleEl)[...roleEl.options].forEach(o=>{if(!o.value)return;if(!o.dataset.baseLabel)o.dataset.baseLabel=o.textContent.replace(/\s*\([^)]*\/17\)\s*$/,"");o.textContent=`${o.dataset.baseLabel} (${ROLE_DISPLAY_STEPS[o.value]||"—"}/${TOTAL_WORKFLOW_STEPS})`;});
 document.querySelectorAll(".protected-tab").forEach(t=>t.classList.remove("hidden"));
 const registerBtn=[...document.querySelectorAll(".protected-tab")].find(t=>t.textContent.trim()==="Registrar");
 if(registerBtn) registerBtn.classList.toggle("hidden",!canRegisterRoles.includes(currentRole));
 previewRole(); renderInbox(); renderNotifications(); renderDocuments();
}
async function changeTestProfile(){
 if(!currentUser)return;
 const roleEl=document.getElementById("roleSelect");
 const sessionEl=document.getElementById("sessionLabel");
 if(!roleEl)return;
 currentRole=roleEl.value||"tecnico";
 if(sessionEl) sessionEl.textContent=` · ${currentUser.name} · Perfil activo: ${roleEl.options[roleEl.selectedIndex].text}`;
 refreshRoleAccess();
 refreshWorkBadges();
}
window.changeTestProfile=changeTestProfile;
async function loginRole(){
 const usernameEl=document.getElementById("usernameInput");
 const passwordEl=document.getElementById("passwordInput");
 const roleEl=document.getElementById("roleSelect");
 const logoutEl=document.getElementById("logoutBtn");
 const sessionEl=document.getElementById("sessionLabel");
 const u=(usernameEl?.value||"").trim().toLowerCase();
 const pw=passwordEl?.value||"";
 const selectedRole=roleEl?.value||"";

 if(!selectedRole){
   if(sessionEl) sessionEl.textContent=" · Seleccione un rol para iniciar sesión";
   roleEl?.focus();
   return;
 }
 if(!u||!pw){
   if(sessionEl) sessionEl.textContent=" · Ingrese usuario y contraseña";
   (!u?usernameEl:passwordEl)?.focus();
   return;
 }
 const account=HUB_USERS[u];
 if(!account||account.password!==pw){
   if(sessionEl) sessionEl.textContent=" · Usuario o contraseña incorrectos";
   alert("Usuario o contraseña incorrectos. Verifique, por ejemplo: usuario5 / Hub2026!05");
   passwordEl?.focus();
   return;
 }
 try{
   currentUser={username:u,...account};
   currentRole=selectedRole;
   if(sessionEl) sessionEl.textContent=" · Verificando identidad Microsoft…";
   const verifiedIdentity=await verifyHubMicrosoftIdentity();
   currentUser.microsoftIdentity=verifiedIdentity;
   loadHubState();
   if(roleEl){roleEl.value=currentRole; roleEl.classList.remove("hidden"); roleEl.removeAttribute("aria-hidden");}
   if(logoutEl) logoutEl.classList.remove("hidden");
   if(usernameEl) usernameEl.disabled=true;
   if(passwordEl) passwordEl.disabled=true;
   populateSelects(); applyGlobalFilters(); renderFlow(); renderKpis(); renderMatrix(); renderDetail(); renderList();
   refreshRoleAccess();
   try{
     await loadProjectsFromExcel();
   }catch(excelErr){console.error(excelErr);alert("Sesión iniciada, pero no fue posible cargar la base maestra mediante el Worker: "+excelErr.message);}
   const roleName=roleEl?.options?.[roleEl.selectedIndex]?.text||currentRole;
   if(sessionEl) sessionEl.textContent=` · ${account.name} · Perfil activo: ${roleName} · Microsoft: ${hubApiIdentity?.email||hubApiIdentity?.displayName||"verificado"}`;
   const firstProtected=document.querySelector('.protected-tab:not(.hidden)');
   if(firstProtected){ firstProtected.scrollIntoView({behavior:'smooth',block:'nearest'}); }
 }catch(err){
   console.error("Error al iniciar sesión",err);
   currentRole="visitor"; currentUser=null;
   if(usernameEl) usernameEl.disabled=false;
   if(passwordEl) passwordEl.disabled=false;
   if(sessionEl) sessionEl.textContent=" · Error al iniciar sesión. Recargue la página e intente de nuevo.";
   alert("Se produjo un error al iniciar sesión: "+(err?.message||err));
 }
 await refreshBadgesAfterSharedSync();
 setTimeout(()=>refreshWorkBadges(),0);
}
window.loginRole=loginRole;
function logoutRole(){
 saveHubState();if(sharedFlowSyncTimer){clearInterval(sharedFlowSyncTimer);sharedFlowSyncTimer=null;}
 currentRole="visitor"; currentUser=null; hubApiIdentity=null; roleSelect.value=""; roleSelect.classList.remove("hidden"); roleSelect.removeAttribute("aria-hidden");
 if(document.getElementById("usernameInput")){document.getElementById("usernameInput").value="";document.getElementById("usernameInput").disabled=false;}
 if(document.getElementById("passwordInput")){document.getElementById("passwordInput").value="";document.getElementById("passwordInput").disabled=false;}
 sessionLabel.textContent=" · Información pública del portafolio"; roleCard.classList.add("hidden"); logoutBtn.classList.add("hidden");
 document.querySelectorAll(".protected-tab").forEach(t=>t.classList.add("hidden"));
 resetWorkspace(); populateSelects(); applyGlobalFilters(); renderMatrix(); renderList();
 showView("inicio",document.querySelector(".public-tab"));
}
function applyRole(){ previewRole(); }
function renderFlow(){
 const groups=[1,2,3,4].map(ph=>steps.filter(s=>s.phase==ph));
 phaseFlowCards.innerHTML=groups.map((group,idx)=>{
   const ph=idx+1;
   const title=ph==1?"FASE 1. GESTIÓN INSTITUCIONAL":ph==2?"FASE 2. EVALUACIÓN Y PRIORIZACIÓN CENTRAL":ph==3?"FASE 3. FORMULACIÓN Y PREPARACIÓN":"FASE 4. EJECUCIÓN, SEGUIMIENTO Y CIERRE";
   return `<div class="phase-block"><div class="phase-header">${title}</div><div class="phase-cards">${group.map(s=>`<article class="phase-card p${ph}"><div class="num">${s.n}</div><h3>${s.title}</h3><p>${s.desc}</p><div class="resp"><b>Responsable:</b><br>${s.role}</div></article>`).join("")}</div></div>`;
 }).join("");
}


let customPortfolioFilters=[];

const EXCEL_FILTER_FIELDS=[
 "ID HUB",
 "Código del BPIP",
 "Etapa",
 "Nombre corto",
 "Nombre del proyecto",
 "¿Pertenece a algún programa?",
 "Programa vinculado",
 "Planes estratégicos vinculados (siglas)",
 "Institución",
 "Ubicación político-administrativa (provincia, cantón)",
 "Ruta Nacional",
 "Inversión por proyecto (Millones USD)",
 "Tipo de Obra",
 "Sub tipo de obra",
 "Identificación de Fuentes y Modalidades de Financiamiento",
 "Otras posibles formas de financiamiento",
 "Comentarios adicionales sobre el financiamiento"
];
const EXCEL_AMOUNT_FIELD="Inversión por proyecto (Millones USD)";

function excelFilterValue(p,field){
 if(p?.excelData && Object.prototype.hasOwnProperty.call(p.excelData,field)) return p.excelData[field];
 const fallback={
  "ID HUB":p?.id,"Código del BPIP":p?.bpip,"Etapa":p?.etapa,
  "Nombre corto":p?.name,"Nombre del proyecto":p?.name,
  "Institución":p?.inst,"Tipo de Obra":p?.categoria,"Sub tipo de obra":p?.subcategoria
 };
 return fallback[field]??"";
}
function parseExcelAmount(v){
 if(v==null||v==="")return null;
 if(typeof v==="number")return Number.isFinite(v)?v:null;
 let s=String(v).trim().replace(/\s/g,"");
 if(!s||/^(NA|NS|N\/A)$/i.test(s))return null;
 if(s.includes(",")&&s.includes(".")){
   if(s.lastIndexOf(",")>s.lastIndexOf("."))s=s.replace(/\./g,"").replace(",",".");
   else s=s.replace(/,/g,"");
 }else if(s.includes(","))s=s.replace(",",".");
 const n=Number(s.replace(/[^\d.-]/g,""));return Number.isFinite(n)?n:null;
}
function availableExcelFilterFields(){
 const headers=new Set(excelHeaders||[]);
 return EXCEL_FILTER_FIELDS.filter(h=>headers.size===0||headers.has(h));
}
function filteredProjectsBefore(index){
 let list=projects.slice();
 for(let i=0;i<index;i++){
   const f=customPortfolioFilters[i]; if(!f?.field)continue;
   if(f.field===EXCEL_AMOUNT_FIELD){
     const min=f.min===""?null:Number(f.min),max=f.max===""?null:Number(f.max);
     list=list.filter(p=>{const n=parseExcelAmount(excelFilterValue(p,f.field));return n!=null&&(min==null||n>=min)&&(max==null||n<=max);});
   }else if(f.values?.length){
     const selected=new Set(f.values.map(String));
     list=list.filter(p=>selected.has(String(excelFilterValue(p,f.field)??"").trim()));
   }
 }
 return list;
}
function excelFieldOptions(field,index){
 return [...new Set(filteredProjectsBefore(index).map(p=>String(excelFilterValue(p,field)??"").trim()).filter(Boolean))]
   .sort((a,b)=>a.localeCompare(b,"es",{numeric:true}));
}
function toggleCustomMulti(i,value,checked){
 const f=customPortfolioFilters[i];if(!f)return;
 const set=new Set(f.values||[]);checked?set.add(value):set.delete(value);f.values=[...set];
 renderCustomFilterBuilder();renderKpis();renderList();if(portfolioMapReady)refreshPortfolioMapLayers();
 setTimeout(()=>document.getElementById("hubMultiMenu"+i)?.classList.add("open"),0);
}
function toggleFilterMenu(i){
 document.querySelectorAll(".hub-multi-menu").forEach((el,j)=>{if(j!==i)el.classList.remove("open")});
 document.getElementById("hubMultiMenu"+i)?.classList.toggle("open");
}
function renderCustomFilterBuilder(){
 const box=document.getElementById("customPortfolioFilterBuilder");if(!box)return;
 const fields=availableExcelFilterFields();
 const active=customPortfolioFilters.filter(f=>f.field&&(f.field===EXCEL_AMOUNT_FIELD?(f.min!==""||f.max!==""):(f.values?.length))).length;
 box.innerHTML=`<div class="custom-filter-head"><div><b>Constructor de filtros</b><div class="muted">Los filtros usan directamente los atributos de HUB_Proyectos.xlsx. Puede seleccionar varios valores en cada lista.</div></div><span>${active} filtro${active===1?"":"s"} activo${active===1?"":"s"}</span></div>
 <div>${customPortfolioFilters.map((f,i)=>{
   const isAmount=f.field===EXCEL_AMOUNT_FIELD;
   const options=f.field&&!isAmount?excelFieldOptions(f.field,i):[];
   const selected=f.values||[];
   return `<div class="custom-filter-row">
    <select onchange="setCustomFilterField(${i},this.value)"><option value="">Seleccione una variable…</option>${fields.filter(k=>k===f.field||!customPortfolioFilters.some((other,j)=>j!==i&&other.field===k)).map(k=>`<option value="${traceEscape(k)}" ${f.field===k?"selected":""}>${traceEscape(k)}</option>`).join("")}</select>
    ${isAmount?`<div class="range-filter"><input type="number" step="any" placeholder="Monto mínimo (USD millones)" value="${traceEscape(f.min??"")}" oninput="setCustomRange(${i},'min',this.value)"><span>—</span><input type="number" step="any" placeholder="Monto máximo (USD millones)" value="${traceEscape(f.max??"")}" oninput="setCustomRange(${i},'max',this.value)"></div>`:
      `<div class="hub-multi"><button type="button" class="hub-multi-trigger" ${f.field?"":"disabled"} onclick="toggleFilterMenu(${i})">${selected.length?`${selected.length} seleccionado${selected.length===1?"":"s"}`:"Todas"} ▾</button>
       <div class="hub-multi-menu" id="hubMultiMenu${i}">${options.length?options.map(v=>`<label><input type="checkbox" ${selected.includes(v)?"checked":""} onchange="toggleCustomMulti(${i},'${String(v).replace(/\\/g,"\\\\").replace(/'/g,"\\'")}',this.checked)"> <span>${traceEscape(v)}</span></label>`).join(""):'<div class="muted" style="padding:10px">Seleccione primero una variable.</div>'}</div></div>`}
    <button type="button" onclick="removeCustomPortfolioFilter(${i})">Quitar</button>
   </div>`}).join("")}</div>
 <div class="custom-filter-actions"><button type="button" onclick="addCustomPortfolioFilter()">+ Agregar variable</button>${customPortfolioFilters.length?'<button type="button" class="secondary" onclick="clearCustomPortfolioFilters()">Limpiar personalizados</button>':""}</div>`;
}
function rerenderCustomFilters(){renderCustomFilterBuilder();renderKpis();renderList();if(portfolioMapReady)refreshPortfolioMapLayers();}
function addCustomPortfolioFilter(){customPortfolioFilters.push({field:"",values:[],min:"",max:""});renderCustomFilterBuilder();}
function removeCustomPortfolioFilter(i){customPortfolioFilters.splice(i,1);rerenderCustomFilters();}
function setCustomFilterField(i,field){customPortfolioFilters[i]={field,values:[],min:"",max:""};rerenderCustomFilters();}
function setCustomRange(i,key,value){if(customPortfolioFilters[i])customPortfolioFilters[i][key]=value;renderKpis();renderList();if(portfolioMapReady)refreshPortfolioMapLayers();}
function clearCustomPortfolioFilters(){customPortfolioFilters=[];rerenderCustomFilters();}
window.addCustomPortfolioFilter=addCustomPortfolioFilter;window.removeCustomPortfolioFilter=removeCustomPortfolioFilter;
window.setCustomFilterField=setCustomFilterField;window.setCustomRange=setCustomRange;window.clearCustomPortfolioFilters=clearCustomPortfolioFilters;
window.toggleCustomMulti=toggleCustomMulti;window.toggleFilterMenu=toggleFilterMenu;

function installCustomPortfolioFilters(){
 const search=document.getElementById("globalSearch");if(!search||document.getElementById("customPortfolioFilterBuilder"))return;
 const panel=search.closest(".filters")||search.parentElement?.parentElement;if(!panel)return;
 [...panel.children].forEach(child=>{if(child.contains(search))return;child.style.display="none";});
 const builder=document.createElement("div");builder.id="customPortfolioFilterBuilder";builder.className="custom-filter-builder";
 panel.insertAdjacentElement("afterend",builder);renderCustomFilterBuilder();
}

function getGlobalFilteredProjects(){
 let q=globalSearch?.value||"", ph=globalPhase?.value||"", inst=globalInst?.value||"", st=globalState?.value||"", cat=globalCat?.value||"", sub=globalSubcat?.value||"", pri=globalPriority?.value||"", step=globalStep?.value||"";
 let result=projects.filter(p=>matches(p,q)&&(!ph||phaseOf(p.step)==+ph)&&(!inst||p.inst==inst)&&(!st||p.state==st)&&(!cat||p.categoria==cat)&&(!sub||p.subcategoria==sub)&&(!pri||p.priority==pri)&&(!step||p.step==+step));
 customPortfolioFilters.forEach(f=>{
  if(!f?.field)return;
  if(f.field===EXCEL_AMOUNT_FIELD){
    const min=f.min===""?null:Number(f.min),max=f.max===""?null:Number(f.max);
    result=result.filter(p=>{const n=parseExcelAmount(excelFilterValue(p,f.field));return n!=null&&(min==null||n>=min)&&(max==null||n<=max);});
  }else if(f.values?.length){
    const selected=new Set(f.values.map(String));
    result=result.filter(p=>selected.has(String(excelFilterValue(p,f.field)??"").trim()));
  }
 });
 if(portfolioMapFilterEnabled){const ids=currentMapFilteredIds();result=result.filter(p=>ids&&ids.has(p.id));}
 return result;
}
function renderKpis(){
  installCustomPortfolioFilters();
  const list=getGlobalFilteredProjects();
  const loaded=projects.length>0;
  kTotal.textContent=loaded?list.length:"—";
  kF1.textContent=loaded?list.filter(p=>phaseOf(p.step)==1).length:"—";
  kF2.textContent=loaded?list.filter(p=>phaseOf(p.step)==2).length:"—";
  kRech.textContent=loaded?list.filter(p=>p.rejected).length:"—";
  if(homeCounter)homeCounter.innerHTML=loaded?`${list.length} de ${projects.length}<br><span style="font-size:12px;color:var(--muted)">proyectos visibles</span>`:`—<br><span style="font-size:12px;color:var(--muted)">base maestra no cargada</span>`;
}
function filteredProjects(){
 let q=matrixSearch.value, ph=phaseFilter.value, inst=instFilter.value, st=stateFilter.value;
 return projects.filter(p=>matches(p,q)&&(!ph||phaseOf(p.step)==+ph)&&(!inst||p.inst==inst)&&(!st||p.state==st));
}
function renderMatrix(){
 const list=filteredProjects();
 matrixCounter.innerHTML=`${list.length} de ${projects.length}<br><span style="font-size:12px;color:var(--muted)">proyectos</span>`;
 const header=`<div class="phase-ribbon"><div class="phase-spacer"></div><div class="phase1">FASE 1 · Gestión institucional</div><div class="phase2">FASE 2 · Evaluación central</div><div class="phase3">FASE 3 · Preparación</div><div class="phase4">FASE 4 · Ejecución y cierre</div></div>
 <div class="header-row"><div class="project-label">Proyecto</div>${steps.map(s=>`<div class="header-cell ${clsStep(s.n)}" title="${s.title}">${s.n}<br>${s.title.split(" ")[0]}</div>`).join("")}</div>`;
 const rows=list.map(p=>`<div class="project-row">
 <div class="project-label matrix-project-name" title="Abrir ficha de ${p.name}"><button type="button" onclick="openMatrixFicha(event,'${p.id}')"><b>${p.id}</b> · ${p.name}</button><small>${p.inst} · ${p.state} · ${phaseName(phaseOf(p.step))}</small></div>
 ${steps.map(s=>`<div class="cell"><span class="dot ${dotClass(p,s.n)}" title="Paso ${s.n} · ${s.title}"></span></div>`).join("")}</div>`).join("");
 const counts=`<div class="counts-row"><div class="project-label">Proyectos que llegaron al paso</div>${steps.map(s=>`<div class="cell"><span class="count">${list.filter(p=>p.step>=s.n).length}</span></div>`).join("")}</div>`;
 matrixBox.innerHTML=header+(rows||'<div class="note">No hay proyectos con esos filtros.</div>')+counts;
}

function openMatrixFicha(ev,id){
 if(ev){ev.preventDefault();ev.stopPropagation();}
 activeId=id;
 openFichasModal();
 renderFicha(id);
}
function selectProject(id){activeId=id; if(projectSelect)projectSelect.value=id; openHistoryModal(id);}
function renderDetail(targetId){
 const p=projects.find(x=>x.id==(targetId||activeId))||projects[0]; if(!p)return "";
 if(projectSelect) projectSelect.value=p.id;
 const step=steps.find(s=>s.n==p.step);
 let timeline=steps.map(s=>`<div class="titem ${dotClass(p,s.n)}"><div class="tnum">${s.n}</div><div class="tbox"><b>${s.title}</b><span class="muted">${s.role} · ${s.desc}</span>${s.n==p.step?`<br><br><span class="badge ${p.rejected?'red':p.returned?'orange':'yellow'}">Paso actual: ${p.state}</span>`:''}</div></div>`).join("");
 const content=`<div class="detail"><div><h3>${p.id} · ${p.name}</h3><p class="muted"><b>Dependencia:</b> ${p.inst}<br><b>BPIP:</b> ${p.bpip||'Por definir'}<br><b>Categoría:</b> ${p.categoria} / ${p.subcategoria}<br><b>Etapa base:</b> ${p.etapa}<br><b>Estado:</b> ${p.state}</p><div class="note"><b>Ubicación actual:</b> ${phaseName(phaseOf(p.step))}, paso ${p.step}: ${step.title}. Responsable sugerido: ${p.owner}.</div></div><div><h3>Lectura del avance</h3><p><span class="badge green">Cumplido</span> <span class="badge yellow">Actual</span> <span class="badge orange">Devuelto</span> <span class="badge red">Rechazado</span></p><p class="muted">Esta ventana resume la trazabilidad del proyecto sin perder la navegación del listado.</p></div></div><hr style="border:0;border-top:1px solid var(--line);margin:22px 0"><h3>Línea de tiempo del proyecto</h3><div class="timeline">${timeline}</div><hr style="border:0;border-top:1px solid var(--line);margin:22px 0"><h3>Trazabilidad documental</h3><p class="muted">Historial de cargas, actualizaciones y eliminaciones registradas en OneDrive.</p><div id="trace-history-box"><p class="muted">Cargando trazabilidad documental desde OneDrive…</p></div>`;
 if(projectDetail)projectDetail.innerHTML=content;
 return content;
}
function openHistoryModal(id){
 activeId=id;
 const p=projects.find(x=>x.id==id);
 modalProjectTitle.textContent=`${p.id} · Historial del proyecto`;
 modalProjectMeta.textContent=`${p.name} · ${p.inst} · ${phaseName(phaseOf(p.step))}`;
 modalProjectBody.innerHTML=renderDetail(id);
 historyModal.classList.add("open");
 loadTraceHistory(id);
}
function closeHistoryModal(e){
 if(e && e.target!==historyModal)return;
 historyModal.classList.remove("open");
}

async function viewProjectDocuments(id){
  const p=projects.find(x=>x.id===id);
  if(!p)return;
  activeId=id;
  manageTitle.textContent=`${p.id} · Documentos del proyecto`;
  manageMeta.textContent=`${p.name} · ${p.inst}`;
  manageBody.innerHTML='<div class="note">Consultando documentos reales en OneDrive…</div>';
  manageModal.classList.add("open");
  try{
    await syncProjectDocuments(id);
    const related=documents.filter(d=>d.project===id);
    if(!related.length){
      manageBody.innerHTML=`<div class="project-doc-empty"><b>Aún no hay documentos cargados.</b><br><span class="muted">Este proyecto no tiene archivos en su carpeta de OneDrive.</span></div>`;
      return;
    }
    manageBody.innerHTML=`<div class="project-doc-view">
      <div class="project-doc-summary"><b>${related.length}</b> ${related.length===1?"documento cargado":"documentos cargados"} en OneDrive</div>
      ${related.map(d=>{
        const idx=documents.indexOf(d);
        return `<div class="project-doc-item">
          <div class="project-doc-main">
            <div class="project-doc-name">${d.name}</div>
            <div class="project-doc-meta">${d.type||"Soporte cargado"} · ${d.state||"En OneDrive"}${d.uploadedAt?` · ${new Date(d.uploadedAt).toLocaleString()}`:""}</div>
          </div>
          <div class="project-doc-actions">
            <button onclick="openOneDriveDocument(documents[${idx}])">Ver documento</button>
          </div>
        </div>`;
      }).join("")}
    </div>`;
  }catch(e){
    console.error(e);
    manageBody.innerHTML=`<div class="note">No fue posible consultar los documentos de OneDrive: ${traceEscape(e.message||e)}</div>`;
  }
}
window.viewProjectDocuments=viewProjectDocuments;

(function injectProjectListStyles(){
  if(document.getElementById("hub-project-list-styles"))return;
  const s=document.createElement("style");
  s.id="hub-project-list-styles";
  s.textContent=`

    .project-click-row{cursor:pointer;transition:background .15s ease,box-shadow .15s ease}
    .project-click-row:hover{background:#f4f8fc}
    .project-click-row:hover td:first-child{box-shadow:inset 4px 0 0 #0b4d8c}
    .project-list-wrap{overflow-x:auto}
    .project-list-table{width:100%;table-layout:fixed;border-collapse:separate;border-spacing:0}
    .project-list-table th{padding:16px 14px;text-align:left;vertical-align:bottom}
    .project-list-table td{padding:20px 14px;vertical-align:top;border-top:1px solid var(--line)}
    .project-list-table .pl-id{width:10%}
    .project-list-table .pl-project{width:36%}
    .project-list-table .pl-info{width:18%}
    .project-list-table .pl-class{width:18%}
    .project-list-table .pl-actions{width:18%}
    .project-id{display:block;font-size:16px;line-height:1.25}
    .project-bpip,.project-sub{display:block;margin-top:6px;color:var(--muted);font-size:13px;line-height:1.35}
    .project-title{font-weight:650;line-height:1.4}
    .project-linkage{margin-top:8px;font-size:12px;color:var(--muted)}
    .project-action-grid{display:flex;flex-wrap:wrap;gap:7px;align-items:center}
    .project-action-grid button{margin:0;white-space:nowrap}
    .project-doc-view{display:grid;gap:12px}
    .project-doc-summary{padding:12px 14px;background:#f5f8fb;border-radius:12px}
    .project-doc-item{display:flex;justify-content:space-between;gap:18px;align-items:center;padding:16px;border:1px solid var(--line);border-radius:14px;background:#fff}
    .project-doc-main{min-width:0}
    .project-doc-name{font-weight:700;overflow-wrap:anywhere}
    .project-doc-meta{margin-top:5px;color:var(--muted);font-size:13px}
    .project-doc-actions{flex:0 0 auto}
    .project-doc-empty{padding:28px;text-align:center;border:1px dashed var(--line);border-radius:14px;background:#fafbfd}



    #fichasModal{z-index:10000!important}
    #fichasModal .modal-card,#fichasModal .ficha-shell{position:relative;z-index:10001}
    #fichasModal .leaflet-container{z-index:1}
    .portfolio-map-section{width:100%;flex:0 0 100%;grid-column:1/-1;margin:22px 0;padding:20px;border:1px solid var(--line);border-radius:18px;background:#fff}

    .map-load-button{margin-left:10px;padding:7px 12px}.portfolio-map-retry{height:100%;display:flex;flex-direction:column;align-items:center;justify-content:center;gap:12px;text-align:center;color:var(--muted)}
    .portfolio-map-actions{display:flex;gap:10px;flex-wrap:wrap;justify-content:flex-end}.portfolio-map-head{display:flex;justify-content:space-between;align-items:center;gap:18px;margin-bottom:12px}
    .portfolio-map-head h2{margin:0 0 4px}.portfolio-map-head p{margin:0;color:var(--muted)}
    .portfolio-projects-map{height:480px;position:relative;border:1px solid var(--line);border-radius:16px;overflow:hidden;background:#eef3f7}
    .location-editor{grid-column:1/-1;margin-top:18px;padding:18px;border:1px solid var(--line);border-radius:16px;background:#fff}
    .location-editor h3{margin:0 0 4px}.location-controls{display:grid;grid-template-columns:minmax(220px,.35fr) 1fr;gap:14px;align-items:end;margin:14px 0}
    .location-help{padding:11px 13px;border-radius:10px;background:#f5f8fb;font-size:13px;line-height:1.5}
    .project-location-map,.ficha-interactive-map{height:390px;border:1px solid var(--line);border-radius:14px;overflow:hidden;background:#eef3f7}
    .location-actions{margin-top:10px;flex-wrap:wrap}
    .ficha-completa{max-width:1500px;margin:auto}
    .ficha-section{margin-top:18px;padding:20px;border:1px solid var(--line);border-radius:16px;background:#fff}
    .ficha-section h4{margin:0 0 6px}
    .ficha-section-head{display:flex;justify-content:space-between;gap:20px;align-items:flex-start}
    .ficha-updated{text-align:right;min-width:210px}.ficha-updated small{display:block;color:var(--muted);margin-bottom:5px}
    .ficha-summary-grid{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:10px;margin-top:14px}
    .ficha-progress{margin-top:14px}
    .ficha-all-data{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:10px;margin-top:14px}
    .ficha-data-item{padding:13px 14px;border:1px solid var(--line);border-radius:12px;background:#fafbfd;min-width:0}
    .ficha-data-item small{display:block;color:var(--muted);font-weight:700;margin-bottom:6px}
    .ficha-data-item div{overflow-wrap:anywhere;line-height:1.35}
    .ficha-map-layout{display:grid;grid-template-columns:1.3fr .7fr;gap:16px;margin-top:14px}
    .ficha-map-data{grid-template-columns:1fr}
    .ficha-history-list{margin-top:12px}
    .ficha-history-row{display:grid;grid-template-columns:190px 1fr;gap:16px;padding:13px 0;border-top:1px solid var(--line)}
    .ficha-history-row:first-child{border-top:0}
    .ficha-history-date{font-weight:700;font-size:13px}
    .ficha-history-empty{padding:18px;border:1px dashed var(--line);border-radius:12px;background:#fafbfd}
    .ficha-bottom-actions{justify-content:flex-end;margin-top:18px}
    @media(max-width:900px){
      .ficha-summary-grid,.ficha-all-data,.ficha-map-layout,.location-controls{grid-template-columns:1fr}
      .ficha-section-head{flex-direction:column}.ficha-updated{text-align:left}
      .ficha-history-row{grid-template-columns:1fr}
    }
    @media(max-width:900px){
      .project-list-table{min-width:900px}
      .project-doc-item{align-items:flex-start;flex-direction:column}
    }
  `;
  document.head.appendChild(s);
})();




function schedulePortfolioMapInit(){
 if(portfolioMapReady||portfolioMapInitializing)return;
 let tries=0;
 const attempt=()=>{
  tries++;
  const el=document.getElementById("portfolioProjectsMap");
  if(!el){if(tries<20)setTimeout(attempt,150);return;}
  const state=document.getElementById("portfolioMapFilterState");
  if(state&&tries===1)state.innerHTML="Preparando mapa del portafolio…";
  if(el.clientWidth>50&&el.clientHeight>50){
   initPortfolioMap();
  }else if(tries<20){
   setTimeout(attempt,150);
  }else if(state){
   state.innerHTML="<b>No fue posible iniciar el mapa:</b> el contenedor no obtuvo un tamaño válido.";
  }
 };
 setTimeout(attempt,50);
}

function ensurePortfolioMapSection(){
 if(document.getElementById("portfolioProjectsMap"))return;
 const heading=[...document.querySelectorAll("h1,h2,h3")].find(x=>x.textContent.trim().toLowerCase().includes("listado general filtrado"));
 if(!heading)return;
 const host=heading.parentElement;
 const section=document.createElement("section");section.className="portfolio-map-section";
 section.innerHTML=`<div class="portfolio-map-head"><div><h2>Mapa del portafolio</h2><p>Visualice la distribución de los proyectos. Los filtros superiores también actualizan el mapa.</p></div><div class="portfolio-map-actions"><button type="button" onclick="applyMapFilterFromButton()">Aplicar filtro con el mapa</button><button type="button" class="secondary" onclick="disablePortfolioMapFilter()">Quitar filtro del mapa</button></div></div><div id="portfolioProjectsMap" class="portfolio-projects-map"></div><div id="portfolioMapFilterState" class="note">Preparando mapa del portafolio… <button type="button" class="secondary map-load-button" onclick="retryPortfolioMap()">Cargar mapa</button></div>`;
 host.insertBefore(section,heading);
 schedulePortfolioMapInit();
}

function renderList(){
  ensurePortfolioMapSection();
  schedulePortfolioMapInit();
  const list=getGlobalFilteredProjects();
  if(portfolioMapReady&&portfolioMap){
   refreshPortfolioMapLayers();
   clearTimeout(portfolioSearchFocusTimer);
   portfolioSearchFocusTimer=setTimeout(()=>focusPortfolioMapFromSearch(list),350);
  }
  if(!projects.length){
    const empty='<div class="note"><b>Base maestra no cargada.</b><br>El HUB no muestra datos de demostración. Los proyectos e indicadores aparecerán al cargar la base real de Excel.</div>';
    if(listBox)listBox.innerHTML=empty;
    if(homeListBox)homeListBox.innerHTML=empty;
    return;
  }
  const table=`<div class="project-list-wrap"><table class="project-list-table">
    <thead><tr>
      <th class="pl-id">ID / BPIP</th>
      <th class="pl-project">Proyecto</th>
      <th class="pl-info">Información</th>
      <th class="pl-class">Clasificación</th>
      <th class="pl-actions">Acciones</th>
    </tr></thead>
    <tbody>${list.slice(0,400).map(p=>{
      const canManage=canManageProject(p);
      return `<tr class="project-click-row" onclick="if(!event.target.closest('button,a,input,select,textarea,label')) focusProjectFromList('${p.id}')">
        <td class="pl-id"><b class="project-id">${p.id}</b><span class="project-bpip">BPIP ${p.bpip||"Por definir"}</span></td>
        <td class="pl-project"><div class="project-title">${p.name}</div><div class="project-sub">${p.inst||"Sin dependencia"}</div></td>
        <td class="pl-info"><span class="badge">${p.state||"Sin estado"}</span><div class="project-sub">${phaseName(phaseOf(p.step))} · Paso ${p.step}</div></td>
        <td class="pl-class"><div>${p.categoria||"Sin categoría"}</div><div class="project-sub">${p.subcategoria||""}</div>${p.priority?`<div class="project-linkage">${p.priority}</div>`:""}</td>
        <td class="pl-actions"><div class="project-action-grid">
          <button onclick="openHistoryModal('${p.id}')">Historial</button>
          <button onclick="openProjectFichaFromList('${p.id}')">Ficha</button><button onclick="viewProjectDocuments('${p.id}')">Documentos</button>
          ${canManage?`<button onclick="openManageModal('${p.id}')">Gestionar</button><button onclick="editProjectExcel('${p.id}')">Editar</button><button class="secondary" onclick="softDeleteProjectExcel('${p.id}')">Eliminar</button>`:""}
        </div></td>
      </tr>`;
    }).join("")}</tbody>
  </table><p class="muted">Mostrando ${Math.min(list.length,400)} de ${list.length} resultados filtrados.</p></div>`;
  if(listBox)listBox.innerHTML=table;
  if(homeListBox)homeListBox.innerHTML=table;
}

function applyGlobalFilters(){
 renderKpis();
 renderList();
 if(!document.getElementById("matriz").classList.contains("hidden"))renderMatrix();
}
function resetGlobalFilters(){
 [globalSearch,globalPhase,globalInst,globalState,globalCat,globalSubcat,globalPriority,globalStep].forEach(el=>{if(el)el.value="";});
 applyGlobalFilters();
}
function renderInbox(){
 if(currentRole==="visitor"){inboxBox.innerHTML='<div class="note">Inicie sesión seleccionando un rol para consultar la bandeja de trabajo.</div>';return;}
 let list=[];
 if(currentRole==="tecnico"){
   list=projects.filter(p=>p.step===1 || p.returned || /devuelto|borrador/i.test(p.state||""));
 }else{
   const allowed=roleStepPermissions[currentRole]||[];
   list=projects.filter(p=>allowed.includes(p.step) || (currentRole==="ejecutor"&&p.step>=10));
 }
 inboxBox.innerHTML=list.length?`<div class="note" style="margin-bottom:12px"><b>${list.length}</b> proyecto${list.length===1?"":"s"} pendiente${list.length===1?"":"s"} para este perfil.</div><table><thead><tr><th>Proyecto</th><th>Acción requerida</th><th>Estado</th><th>Gestión</th></tr></thead><tbody>${list.slice(0,100).map(p=>`<tr class="project-click-row" onclick="if(!event.target.closest('button,a,input,select')) focusProjectFromList('${p.id}')"><td><b>${p.id}</b><br>${p.name}</td><td>${currentRole==="tecnico"?"Registrar, completar ficha o atender devolución":"Revisar paso "+p.step+": "+(steps.find(s=>s.n==p.step)?.title||"")}</td><td>${p.state}</td><td><button onclick="openManageModal('${p.id}')">Gestionar</button></td></tr>`).join("")}</tbody></table>`:'<div class="note">No hay proyectos pendientes para este rol.</div>';
 refreshWorkBadges();
}
async function renderNotifications(){
 if(currentRole==="visitor"){notifBox.innerHTML='<div class="note">Inicie sesión seleccionando un rol para consultar notificaciones internas.</div>';refreshWorkBadges();return;}
 const all=notificationProjects().sort((a,b)=>String(b.sharedUpdatedAt||"").localeCompare(String(a.sharedUpdatedAt||"")));
 const seen=getSeenNotifications();
 internalNotifications=all;
 notifBox.innerHTML=all.length?`<div class="mail-notification-list">${all.map(p=>{
  const s=steps.find(x=>x.n==p.step), unread=seen[p.id]!==p.sharedUpdatedAt;
  return `<div class="note" style="border-left:5px solid ${unread?"#c62828":"#94a3b8"};padding:14px 16px;margin-bottom:10px;${unread?"background:#fff8f8":""}">
   <div style="display:flex;justify-content:space-between;gap:16px"><b>${unread?"📩 Nuevo trabajo":"✓ Leída"} · ${traceEscape(p.id)}</b><span class="muted">${traceEscape(new Date(p.sharedUpdatedAt).toLocaleString())}</span></div>
   <div style="margin-top:6px"><b>${traceEscape(p.name)}</b></div>
   <div class="muted" style="margin-top:5px">Asignado a: ${traceEscape(s?.role||p.owner)} · Paso ${p.step}/${TOTAL_WORKFLOW_STEPS}: ${traceEscape(s?.title||"")}</div>
   <div style="margin-top:9px;display:flex;gap:8px"><button class="primary" onclick="markNotificationRead('${p.id}','${p.sharedUpdatedAt}');openManageModal('${p.id}')">Abrir trabajo</button>${unread?`<button onclick="markNotificationRead('${p.id}','${p.sharedUpdatedAt}')">Marcar como leída</button>`:""}</div>
  </div>`;
 }).join("")}</div>`:'<div class="note">No tiene notificaciones para este perfil.</div>';
 refreshWorkBadges();
}


function canManageProject(p){
 if(currentRole==="visitor")return false;
 if(currentRole==="tecnico"){
   return p.step===1 || p.returned || /devuelto|borrador/i.test(p.state||"");
 }
 return (roleStepPermissions[currentRole]||[]).includes(p.step) || (currentRole==="ejecutor"&&p.step>=10);
}
function openManageModal(id,skipOneDriveSync=false){
 const p=projects.find(x=>x.id==id); if(!p)return;
 activeId=id;
 manageTitle.textContent=`${p.id} · Gestionar proyecto`;
 manageMeta.textContent=`${p.name} · ${p.inst} · Paso ${p.step}: ${steps.find(s=>s.n==p.step)?.title||""}`;
 const allowed=canManageProject(p);
 const relatedDocs=documents.filter(d=>d.project==p.id);
 const audits=auditLog.filter(a=>a.project==p.id);
 const tecnicoActions=currentRole==="tecnico";
 manageBody.innerHTML=`<div class="manage-grid">
  <div class="manage-card">
    <h3>${tecnicoActions?"Registro / ajuste de iniciativa":"Estado actual"}</h3>
    <p class="muted"><b>Fase:</b> ${phaseName(phaseOf(p.step))}<br><b>Paso:</b> ${p.step}. ${steps.find(s=>s.n==p.step)?.title||""}<br><b>Responsable actual:</b> ${p.owner}<br><b>Estado:</b> ${p.state}</p>
    ${allowed?`<label>${tecnicoActions?"Observación para envío o ajuste":"Observación de gestión"}</label><textarea id="manageComment" rows="4" placeholder="${tecnicoActions?"Describa el registro, complemento o ajuste realizado antes de enviarlo a revisión":"Escriba la justificación, observación o instrucción para el siguiente responsable"}"></textarea>
    <div class="action-row">
      ${tecnicoActions?`
        <button class="btn-approve" onclick="manageDecision('${p.id}','enviar_revision')">Enviar a revisión técnica</button>
        <button class="btn-purple" onclick="manageDecision('${p.id}','guardar_borrador')">Guardar borrador / ajuste</button>
      `:`
        <button class="btn-approve" onclick="manageDecision('${p.id}','aprobar')">Aprobar / avanzar</button>
        <button class="btn-return" onclick="manageDecision('${p.id}','devolver')">Devolver para ajustes</button>
        <button class="btn-reject" onclick="manageDecision('${p.id}','rechazar')">Rechazar / archivar</button>
      `}
    </div>`:`<div class="note">Su rol actual no tiene permiso para gestionar este paso. Puede consultar el historial y los documentos disponibles.</div>`}
  </div>
  <div class="manage-card">
    <h3>Documentos del proyecto</h3>
    <div class="doc-drop">Cargar soporte / nueva versión<br><small>Los archivos se almacenan en OneDrive, dentro de la carpeta del proyecto.</small></div>
    <div class="action-row"><button onclick="simulateDocUpload('${p.id}')">Cargar documento</button><button class="secondary" onclick="showView('documentos',[...document.querySelectorAll('.protected-tab')].find(t=>t.textContent.trim()==='Documentos'));closeManageModal()">Ir a documentos</button></div>
    <div style="margin-top:12px">${relatedDocs.length?relatedDocs.map(d=>`<div class="audit-item"><b>${d.type}</b><br>${d.name} · ${d.state}</div>`).join(""):'<p class="muted">Sin documentos asociados todavía.</p>'}</div>
  </div>
  <div class="manage-card">
    <h3>Bitácora de gestión</h3>
    <div class="audit-box">${audits.length?audits.map(a=>`<div class="audit-item"><b>${a.action}</b> · ${a.role}<br><span class="muted">${a.date}</span><br>${a.comment||"Sin observación"}</div>`).join(""):'<div class="muted">Aún no hay movimientos registrados en esta sesión.</div>'}</div>
  </div>
  <div class="manage-card">
    <h3>Lectura rápida del flujo</h3>
    <p class="muted">${tecnicoActions?"El técnico registra la iniciativa, completa información, carga soportes y atiende devoluciones. No aprueba ni rechaza; al enviar, pasa a revisión de la jefatura técnica.":"Al aprobar, el proyecto avanza al siguiente paso y se notifica al responsable correspondiente. Al devolver, se mantiene trazabilidad y vuelve a ajuste. Al rechazar, queda marcado como archivado."}</p>
    <button onclick="openHistoryModal('${p.id}')">Ver historial completo</button>
  </div>
 </div>`;
 manageModal.classList.add("open");
 if(!skipOneDriveSync){syncProjectDocuments(id).then(()=>{if(manageModal.classList.contains("open")&&activeId===id)openManageModal(id,true);}).catch(e=>console.warn("No se pudo sincronizar OneDrive",e));}
}
function closeManageModal(e){
 if(e && e.target!==manageModal)return;
 manageModal.classList.remove("open");
}
async function manageDecision(id,decision){
 const p=projects.find(x=>x.id==id); if(!p)return;
 const pasoOrigen=p.step, estadoOrigen=p.state;
 showHubOperationProgress("Enviando gestión","Preparando el movimiento del proyecto…",12);
 const comment=(document.getElementById("manageComment")||{}).value||"";
 if(!canManageProject(p)){alert("Su rol no tiene permiso para gestionar este paso.");return;}
 let action="";
 if(currentRole==="tecnico"){
   if(decision==="enviar_revision"){
     p.step=2;
     p.state="En revisión técnica";
     p.owner=steps.find(s=>s.n==2)?.role||"Jefatura Técnica / Coordinador";
     p.returned=false;p.rejected=false;
     action="Envió la iniciativa a revisión técnica";
   }else if(decision==="guardar_borrador"){
     p.step=1;
     p.state="Borrador / ajuste en curso";
     p.owner="Técnico / Profesional";
     p.returned=false;p.rejected=false;
     action="Guardó registro o ajuste de la iniciativa";
   }else{
     alert("El rol Técnico / Profesional no aprueba, no rechaza y no devuelve iniciativas. Solo registra, ajusta y envía a revisión.");
     return;
   }
 }else if(decision==="aprobar"){
   if(p.step<17)p.step+=1;
   p.state=p.step>=14?"En ejecución":"En revisión";
   p.owner=steps.find(s=>s.n==p.step)?.role||p.owner;
   p.returned=false;p.rejected=false;
   action="Aprobó y avanzó el proyecto";
 }else if(decision==="devolver"){
   p.state="Devuelto para ajustes";
   p.returned=true;p.rejected=false;
   p.step=1;
   p.owner="Técnico / Profesional";
   action="Devolvió para ajustes al técnico";
 }else{
   p.state="Rechazado / archivado";
   p.rejected=true;p.returned=false;
   action="Rechazó y archivó el proyecto";
 }
 auditLog.unshift({project:id,action,role:actorLabel(),date:new Date().toISOString(),comment});
 saveHubState();
 const gestionFecha=new Date().toISOString();
 try{
  setHubOperationProgress(35,"Guardando trazabilidad…");
  await writeTrace(p,"GESTION_FLUJO",p.id,{descripcion:action,comentario:comment,pasoOrigen,estadoOrigen,pasoDestino:p.step,estadoDestino:p.state,fechaGestion:gestionFecha,recibidoPor:steps.find(s=>s.n==pasoOrigen)?.role||"",entregadoA:steps.find(s=>s.n==p.step)?.role||p.owner});
  setHubOperationProgress(60,"Actualizando el flujo compartido y la bandeja siguiente…");
  await saveSharedWorkflow(p,{
   date:gestionFecha,
   action,
   comment,
   stepFrom:pasoOrigen,
   stateFrom:estadoOrigen,
   stepTo:p.step,
   stateTo:p.state,
   ownerTo:p.owner,
   updatedBy:actorLabel()
  });
  setHubOperationProgress(82,"Actualizando el reporte de gestión…");
  await writeManagementHistoryDocument(p);
 }catch(e){console.warn("No se pudo persistir completamente el movimiento compartido",e);alert("La gestión se realizó en pantalla, pero hubo un problema al sincronizarla con OneDrive: "+e.message);}
 renderKpis();renderList();renderMatrix();renderInbox();await renderNotifications();
 openManageModal(id);
 finishHubOperationProgress("Gestión enviada correctamente");
 setTimeout(()=>alert("Gestión registrada y sincronizada. El siguiente perfil ya puede verla en su bandeja y notificaciones."),500);
}
async function simulateDocUpload(id){
 const p=projects.find(x=>x.id==id);if(!p)return;
 const input=document.createElement("input");input.type="file";input.accept=".pdf,.doc,.docx,.xls,.xlsx,.png,.jpg,.jpeg,.txt";
 input.onchange=async()=>{
  const file=input.files?.[0];if(!file)return;
  try{
   showHubOperationProgress("Cargando documento",`Preparando ${file.name}…`,15);
   setHubOperationProgress(30,"Localizando la carpeta del proyecto en OneDrive…");
   const item=await uploadOneDriveFile(p,file);
   setHubOperationProgress(72,"Archivo cargado. Registrando trazabilidad…");
   documents=documents.filter(d=>!(d.project===id&&d.itemId===item.id));
   documents.unshift({project:id,type:"Soporte cargado",name:item.name,state:"En OneDrive",owner:actorLabel(),itemId:item.id,driveId:item.parentReference?.driveId,webUrl:item.webUrl,size:item.size,uploadedAt:item.lastModifiedDateTime});
   auditLog.unshift({project:id,action:"Cargó soporte documental",role:actorLabel(),date:new Date().toLocaleString(),comment:`Se cargó ${item.name} en OneDrive.`});
   saveHubState();await writeTrace(p,"SUBIO_DOCUMENTO",item.name,{itemId:item.id});
   setHubOperationProgress(92,"Actualizando documentos del proyecto…");
   renderDocuments();openManageModal(id,true);
   finishHubOperationProgress("Documento guardado correctamente");
   setTimeout(()=>alert("Documento guardado correctamente en OneDrive y registrado en trazabilidad."),500);
  }catch(e){
   console.error(e);failHubOperationProgress("No fue posible cargar el documento");
   setTimeout(()=>alert("No fue posible guardar el archivo en OneDrive: "+e.message),500);
  }
 };input.click();
}

function canEditDoc(type){
 if(currentRole==="visitor")return false;
 if(type==="Soporte cargado")return true;
 return (docEditPermissions[currentRole]||[]).includes(type);
}
function renderDocuments(){
 if(currentRole==="visitor"){documentsBox.innerHTML='<div class="note">Inicie sesión para consultar documentos internos.</div>';return;}
 const q=norm(docSearch?.value||""), t=docTypeFilter?.value||"", st=docStateFilter?.value||"";
 const perms=docEditPermissions[currentRole]||[];
 docPermissionNote.innerHTML=perms.length?`Su rol puede editar: <b>${perms.join(", ")}</b>. Los demás documentos quedan en solo lectura.`:"Su rol puede consultar documentos, pero no tiene permiso de edición en esta etapa.";
 const list=documents.filter(d=>(!q||norm([d.project,d.type,d.name,d.state,d.owner].join(" ")).includes(q))&&(!t||d.type==t)&&(!st||d.state==st));
 documentsBox.innerHTML=`<div class="action-row" style="margin-bottom:12px"><button class="secondary" onclick="syncProjectDocuments(activeId).then(()=>renderDocuments()).catch(e=>alert('No fue posible sincronizar OneDrive: '+e.message))">Sincronizar proyecto activo con OneDrive</button></div><table><thead><tr><th>Proyecto</th><th>Documento</th><th>Tipo</th><th>Estado</th><th>Responsable</th><th>Acciones según rol</th></tr></thead><tbody>${list.map(d=>{const idx=documents.indexOf(d);return `<tr><td><b>${d.project}</b></td><td>${d.name}</td><td>${d.type}</td><td><span class="badge">${d.state}</span></td><td>${d.owner}</td><td><div class="doc-actions"><button onclick="openOneDriveDocument(documents[${idx}])">Ver</button>${canEditDoc(d.type)?`<button onclick="openOneDriveDocument(documents[${idx}])">Editar</button><button onclick="replaceOneDriveDocument(${idx})">Cargar versión</button><button class="btn-reject" onclick="deleteOneDriveDocument(${idx})">Eliminar</button>`:`<button class="disabled" onclick="alert('Documento en solo lectura para este rol.')">Solo lectura</button>`}</div></td></tr>`}).join("")}</tbody></table>`;
}

const HUB_GEO_FEATURES = window.HUB_GEO_DATA;

function hashNumber(str,min,max){
 let h=0; for(let i=0;i<String(str).length;i++) h=(h*31+String(str).charCodeAt(i))>>>0;
 return min+(h%(max-min+1));
}
function featureCenter(feature){
 const pts=[];
 function walk(c){
  if(!c)return;
  if(typeof c[0]==='number' && typeof c[1]==='number'){pts.push(c);return;}
  c.forEach(walk);
 }
 walk(feature.coords);
 if(!pts.length)return null;
 const lon=pts.reduce((s,p)=>s+p[0],0)/pts.length;
 const lat=pts.reduce((s,p)=>s+p[1],0)/pts.length;
 return {lat:+lat.toFixed(5),lon:+lon.toFixed(5)};
}
function findGeoFeature(p){
 const id=norm(p.id||'');
 let f=HUB_GEO_FEATURES.find(g=>norm(g.id)===id);
 if(f)return f;
 const n=norm(p.name||'');
 f=HUB_GEO_FEATURES.find(g=>g.name && (norm(g.name).includes(n.slice(0,32)) || n.includes(norm(g.name).slice(0,32))));
 return f||null;
}
function inferGeo(p){
 const feature=findGeoFeature(p);
 if(feature){
  const c=featureCenter(feature)||{lat:9.93,lon:-84.08};
  const loc=[feature.dist,feature.canton,feature.prov].filter(Boolean).join(', ') || 'Ubicación según capa HUB';
  return {lat:c.lat,lon:c.lon,lugar:loc,ambito:p.subcategoria||p.categoria||feature.layer||'Sin clasificación',feature:feature,fuente:'Capa geográfica HUB'};
 }
 const text=norm([p.name,p.inst,p.categoria,p.subcategoria].join(" "));
 const locations=[
  ["limon",9.99,-83.03,"Limón"],["siquirres",10.10,-83.51,"Siquirres, Limón"],["talamanca",9.63,-82.72,"Talamanca, Limón"],["tortuguero",10.54,-83.50,"Tortuguero, Limón"],["matina",10.08,-83.29,"Matina, Limón"],["moin",10.00,-83.08,"Moín, Limón"],["cahuita",9.74,-82.84,"Cahuita, Limón"],["manzanillo",9.63,-82.65,"Manzanillo, Limón"],["colorado",10.78,-83.59,"Colorado, Limón"],
  ["puntarenas",9.98,-84.83,"Puntarenas"],["paquera",9.82,-84.94,"Paquera, Puntarenas"],["playa naranjo",9.85,-84.93,"Playa Naranjo, Puntarenas"],["chomes",10.04,-84.90,"Chomes, Puntarenas"],["uvita",9.16,-83.73,"Uvita, Puntarenas"],["burica",8.05,-82.90,"Punta Burica, Puntarenas"],["coyote",9.78,-85.25,"Puerto Coyote, Puntarenas"],
  ["guanacaste",10.63,-85.44,"Guanacaste"],["liberia",10.63,-85.44,"Liberia, Guanacaste"],["ostional",9.99,-85.70,"Ostional, Guanacaste"],
  ["alajuela",10.02,-84.21,"Alajuela"],["san carlos",10.32,-84.43,"San Carlos, Alajuela"],["florencia",10.36,-84.48,"Florencia, Alajuela"],["penas blancas",10.35,-84.74,"Peñas Blancas, Alajuela"],
  ["cartago",9.86,-83.92,"Cartago"],["ochomogo",9.90,-83.95,"Ochomogo, Cartago"],
  ["heredia",10.00,-84.12,"Heredia"],["barreal",9.98,-84.14,"Barreal, Heredia"],["lagunilla",9.99,-84.12,"Lagunilla, Heredia"],
  ["san jose",9.93,-84.08,"San José"],["zapote",9.92,-84.05,"Zapote, San José"],["uruca",9.95,-84.11,"La Uruca, San José"],["plaza viquez",9.93,-84.07,"Plaza Víquez, San José"],["paso ancho",9.91,-84.08,"Paso Ancho, San José"],["perez zeledon",9.36,-83.70,"Pérez Zeledón, San José"]
 ];
 let found=locations.find(x=>text.includes(x[0]));
 if(!found){
   const defaults={CONAVI:[9.93,-84.08,"Red vial nacional"],DVMP:[9.99,-84.83,"Ámbito marítimo-portuario"],DEN:[9.93,-84.08,"Sedes e infraestructura institucional"],"UA al CAS":[10.02,-84.21,"Proyecto institucional / territorial"],"UE PIV MU":[10.32,-84.43,"Corredor vial estratégico"]};
   const d=defaults[p.inst]||[9.93,-84.08,"Ubicación por definir"];
   found=["",d[0],d[1],d[2]];
 }
 const lat=found[1]+(hashNumber(p.id,0,18)-9)/1000;
 const lon=found[2]+(hashNumber(p.name,0,18)-9)/1000;
 return {lat:+lat.toFixed(5),lon:+lon.toFixed(5),lugar:found[3],ambito:p.subcategoria||p.categoria||"Sin clasificación",feature:null,fuente:'Ubicación inferida'};
}
function projectScore(p){
 const base={Alto:86,Medio:68,Bajo:48,"Sin dato":30}[p.priority]||45;
 const phase=phaseOf(p.step)*5;
 const rejected=p.rejected?-45:0;
 return Math.max(0,Math.min(100,base+phase+hashNumber(p.id,0,8)+rejected));
}
function flattenCoords(coords){
 const pts=[];
 function walk(c){
  if(!c)return;
  if(typeof c[0]==='number' && typeof c[1]==='number'){pts.push(c);return;}
  c.forEach(walk);
 }
 walk(coords);
 return pts;
}
function mapGeometryPath(coords,project){
 function point(pt){return `${project.x(pt[0]).toFixed(1)},${project.y(pt[1]).toFixed(1)}`;}
 function path(c){
  if(!c || !c.length)return '';
  if(typeof c[0][0]==='number') return 'M '+c.map(point).join(' L ');
  return c.map(path).join(' ');
 }
 return path(coords);
}
function mapSvg(p,geo){
 const feature=geo.feature;
 const all=feature?flattenCoords(feature.coords):[[geo.lon,geo.lat]];
 let minLon=Math.min(...all.map(x=>x[0])), maxLon=Math.max(...all.map(x=>x[0])), minLat=Math.min(...all.map(x=>x[1])), maxLat=Math.max(...all.map(x=>x[1]));
 if(maxLon-minLon<0.03){minLon-=0.08;maxLon+=0.08;} else {const pad=(maxLon-minLon)*.18;minLon-=pad;maxLon+=pad;}
 if(maxLat-minLat<0.03){minLat-=0.08;maxLat+=0.08;} else {const pad=(maxLat-minLat)*.18;minLat-=pad;maxLat+=pad;}
 const project={x:lon=>40+((lon-minLon)/(maxLon-minLon))*440,y:lat=>300-((lat-minLat)/(maxLat-minLat))*235};
 const cx=project.x(geo.lon), cy=project.y(geo.lat);
 let geom='';
 if(feature){
  const path=mapGeometryPath(feature.coords,project);
  if(feature.type.includes('Line')) geom=`<path d="${path}" fill="none" stroke="#0f8a8f" stroke-width="8" stroke-linecap="round" stroke-linejoin="round" opacity=".92"/>`;
  else if(feature.type.includes('Point')) geom=`<circle cx="${cx}" cy="${cy}" r="16" fill="#0f8a8f" stroke="#fff" stroke-width="5"/>`;
  else geom=`<path d="${path} Z" fill="#b4d5b5" stroke="#0f8a8f" stroke-width="3" opacity=".75"/>`;
 }
 return `<svg viewBox="0 0 520 360" width="100%" height="100%" role="img" aria-label="Mapa del proyecto con capa geográfica HUB">
  <defs><linearGradient id="gmap" x1="0" x2="1"><stop offset="0" stop-color="#eaf7ff"/><stop offset="1" stop-color="#eef8ef"/></linearGradient></defs>
  <rect width="520" height="360" fill="url(#gmap)"/>
  <path d="M82 82 C140 40 212 62 268 90 C327 119 381 105 443 132 C485 151 482 220 431 245 C379 271 335 249 287 281 C239 313 174 318 133 278 C91 237 37 216 49 158 C54 128 58 101 82 82Z" fill="#ffffff" stroke="#b9cce0" stroke-width="3"/>
  <path d="M95 215 C150 190 190 197 242 210 C291 223 328 218 386 188" fill="none" stroke="#9fc3dd" stroke-width="8" stroke-linecap="round" opacity=".35"/>
  ${geom}
  <g transform="translate(${cx.toFixed(1)} ${cy.toFixed(1)})">
   <path d="M0 -30 C16 -30 29 -18 29 -2 C29 18 0 43 0 43 C0 43 -29 18 -29 -2 C-29 -18 -16 -30 0 -30Z" fill="#bf3434" stroke="#fff" stroke-width="5"/>
   <circle r="9" fill="#fff"/><circle r="4" fill="#073b75"/>
  </g>
  <rect x="22" y="22" width="236" height="64" rx="16" fill="#ffffff" opacity=".95" stroke="#dce5ef"/>
  <text x="40" y="48" font-size="15" font-weight="800" fill="#073b75">Capa geográfica HUB</text>
  <text x="40" y="69" font-size="12" fill="#667085">${geo.fuente}: ${geo.lugar}</text>
  <rect x="305" y="285" width="190" height="42" rx="13" fill="#ffffff" opacity=".95" stroke="#dce5ef"/>
  <text x="321" y="309" font-size="12" font-weight="800" fill="#073b75">Lat ${geo.lat} · Lon ${geo.lon}</text>
 </svg>`;
}
function fichaFilteredProjects(){
 const q=fichaSearch?.value||"";
 return projects.filter(p=>matches(p,q)).slice(0,80);
}



function suspendPortfolioMapForModal(){
 const section=document.querySelector(".portfolio-map-section");
 if(section){
  section.dataset.hiddenByFicha="1";
  section.style.visibility="hidden";
 }
}
function restorePortfolioMapAfterModal(){
 const section=document.querySelector(".portfolio-map-section[data-hidden-by-ficha='1']");
 if(section){
  section.style.visibility="";
  delete section.dataset.hiddenByFicha;
 }
 if(portfolioMap){
  setTimeout(()=>{try{portfolioMap.invalidateSize(false)}catch(e){}},80);
 }
}
window.suspendPortfolioMapForModal=suspendPortfolioMapForModal;
window.restorePortfolioMapAfterModal=restorePortfolioMapAfterModal;

async function openProjectFichaFromList(id){
 suspendPortfolioMapForModal();
 openFichasModal();
 await renderFicha(id);
}
window.openProjectFichaFromList=openProjectFichaFromList;

function openFichasModal(){
 suspendPortfolioMapForModal();
 const modal=document.getElementById('fichas');
 if(!modal) return;
 modal.classList.remove('hidden');
 modal.classList.add('ficha-modal-open');
 document.body.classList.add('fichas-open');
}
function closeFichasModal(){
 restorePortfolioMapAfterModal();
 const modal=document.getElementById('fichas');
 if(!modal) return;
 modal.classList.add('hidden');
 modal.classList.remove('ficha-modal-open');
 document.body.classList.remove('fichas-open');
}
window.addEventListener('keydown',function(e){if(e.key==='Escape')closeFichasModal();});

function renderFichaResults(){
 const list=fichaFilteredProjects();
 if(geoCounter) geoCounter.innerHTML=`${list.length} resultados<br><span style="font-size:12px;color:var(--muted)">de ${projects.length} proyectos</span>`;
 if(fichaResults) fichaResults.innerHTML=list.map((p,i)=>`<button class="geo-result ${p.id===activeId?'active':''}" onclick="renderFicha('${p.id}')">${p.id} · ${p.name}<small>${p.inst} · ${p.state} · ${p.subcategoria||p.categoria||''}</small></button>`).join("") || '<div class="note">No hay proyectos que coincidan con la búsqueda.</div>';
 if(fichaSelect) fichaSelect.innerHTML=list.map(p=>`<option value="${p.id}">${p.id} · ${p.name.slice(0,85)}</option>`).join("");
}

function realMapIframe(p,geo){return `<div id="fichaInteractiveMap" class="ficha-interactive-map"></div><div class="map-chip">${geo.userGeometry?(geo.userGeometry.type==="segment"?"Tramo ingresado por usuario":"Ubicación ingresada por usuario"):`Punto de referencia · ${geo.fuente}`}</div>`;}
function realMapLink(geo){
 const lat=Number(geo.lat)||9.93, lon=Number(geo.lon)||-84.08;
 return `https://www.openstreetmap.org/?mlat=${lat.toFixed(5)}&mlon=${lon.toFixed(5)}#map=14/${lat.toFixed(5)}/${lon.toFixed(5)}`;
}



async function initFichaInteractiveMap(geo){
 const el=document.getElementById("fichaInteractiveMap");if(!el)return;
 try{
  await loadLeaflet();if(fichaLeafletMap){try{fichaLeafletMap.remove()}catch(e){}}
  fichaLeafletMap=L.map(el).setView([geo.lat,geo.lon],13);L.tileLayer("https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png",{maxZoom:19,attribution:"© OpenStreetMap"}).addTo(fichaLeafletMap);
  const g=geo.userGeometry;
  if(g?.type==="segment"&&g.coordinates?.length>1){const line=L.polyline(g.coordinates.map(c=>[c[1],c[0]]),{weight:6}).addTo(fichaLeafletMap);fichaLeafletMap.fitBounds(line.getBounds().pad(.15));}
  else L.marker([geo.lat,geo.lon]).addTo(fichaLeafletMap);
  setTimeout(()=>fichaLeafletMap.invalidateSize(),80);
 }catch(e){el.innerHTML=`<div class="note">${traceEscape(e.message)}</div>`;}
}

function fichaEscape(v){return traceEscape(v===""||v==null?"—":v);}
function fichaLabel(h){return String(h||"").trim();}
function fichaDate(v){
  if(!v)return "—";
  const d=new Date(v);
  return isNaN(d)?String(v):d.toLocaleString();
}
function fichaFlowSummary(p,events=[]){
  const local=auditLog.filter(a=>a.project===p.id).map(a=>({
    fecha:a.date, accion:a.action, detalle:a.comment||"", perfil:a.role||"", step:null
  }));
  const remote=events.filter(e=>e.accion==="GESTION_FLUJO").map(e=>({
    fecha:e.fecha, accion:e.descripcion||traceActionLabel(e.accion), detalle:e.comentario||"", perfil:e.perfil||"", step:e.pasoDestino||e.paso||null
  }));
  const all=[...remote,...local].sort((a,b)=>new Date(b.fecha||0)-new Date(a.fecha||0));
  if(!all.length){
    return `<div class="ficha-history-empty"><b>No hay movimientos históricos de flujo registrados todavía.</b><br><span class="muted">La ficha muestra el estado actual del Excel. Los próximos avances, devoluciones y decisiones quedarán fechados aquí.</span></div>`;
  }
  return `<div class="ficha-history-list">${all.slice(0,12).map(e=>`<div class="ficha-history-row">
    <div class="ficha-history-date">${fichaEscape(fichaDate(e.fecha))}</div>
    <div><b>${fichaEscape(e.accion)}</b>${e.step?` <span class="badge">Paso ${e.step}</span>`:""}<div class="muted">${fichaEscape(e.perfil)}${e.detalle?` · ${fichaEscape(e.detalle)}`:""}</div></div>
  </div>`).join("")}</div>`;
}
async function loadFichaHistory(p){
  const box=document.getElementById("fichaHistoryBox"); if(!box)return;
  try{
    const events=await readProjectTrace(p);
    box.innerHTML=fichaFlowSummary(p,events);
  }catch(e){
    console.warn("No se pudo cargar historial de ficha",e);
    box.innerHTML=fichaFlowSummary(p,[]);
  }
}

async function renderFicha(id){
 const p=projects.find(x=>x.id==(id||activeId))||projects[0]; if(!p)return;
 activeId=p.id;
 const modalTitle=document.getElementById('fichaModalTitle');
 if(modalTitle) modalTitle.textContent=`${p.id} · ${p.name}`;
 let geo=inferGeo(p);const ug=await loadProjectGeometry(p.id);
 if(ug){const c=geometryCenter(ug);geo={...geo,lat:+c.lat.toFixed(6),lon:+c.lon.toFixed(6),lugar:ug.type==="segment"?"Tramo definido por usuario":"Ubicación definida por usuario",fuente:"Ubicación ingresada por usuario",feature:null,userGeometry:ug};}
 const step=steps.find(s=>s.n==p.step)||steps[0], advance=Math.round((p.step/17)*100);
 const data=p.excelData||{};
 const hiddenHeaders=new Set([""]);
 const rows=Object.entries(data).filter(([h])=>!hiddenHeaders.has(hnorm(h)));
 const updated=data["Fecha actualización"]||data["Fecha Actualización"]||"";
 fichaBox.innerHTML=`<article class="ficha-card ficha-completa">
  <div class="ficha-top">
    <div>
      <h3>${fichaEscape(p.name)}</h3>
      <p><b>${fichaEscape(p.inst)}</b> · ${phaseName(phaseOf(p.step))} · Paso ${p.step}: ${fichaEscape(step.title)}</p>
      <div class="ficha-badges"><span class="badge">${fichaEscape(p.state)}</span><span class="badge">${fichaEscape(p.etapa||"Sin etapa")}</span><span class="badge">${fichaEscape(p.categoria||"Sin categoría")}</span></div>
    </div>
    <div class="ficha-code"><small>ID HUB</small><b>${fichaEscape(p.id)}</b><small>BPIP: ${fichaEscape(p.bpip||"Por definir")}</small></div>
  </div>

  <div class="ficha-section ficha-location-first">
    <div class="ficha-section-head">
      <div>
        <h4>Ubicación del proyecto</h4>
        <p class="muted">Localización y referencia espacial registrada para la iniciativa.</p>
      </div>
      <div class="ficha-updated">
        <small>Fecha de actualización</small>
        <b>${fichaEscape(fichaDate(updated))}</b>
      </div>
    </div>
    <div class="ficha-map-layout">
      <div><div class="geo-map">${realMapIframe(p,geo)}</div><a class="map-link" href="${realMapLink(geo)}" target="_blank" rel="noopener">Abrir mapa ampliado ↗</a></div>
      <div class="ficha-summary-grid ficha-map-data">
        <div class="info-pill"><small>Coordenadas de referencia</small><b>${fichaEscape(geo.lat)}, ${fichaEscape(geo.lon)}</b></div>
        <div class="info-pill"><small>Ubicación</small><b>${fichaEscape(geo.lugar)}</b></div>
        <div class="info-pill"><small>Fuente espacial</small><b>${fichaEscape(geo.fuente)}</b></div>
        <div class="info-pill"><small>Capa</small><b>${fichaEscape(geo.feature?geo.feature.layer:"Sin coincidencia exacta")}</b></div>
      </div>
    </div>
  </div>


  <div class="ficha-section">
    <div class="ficha-section-head"><div><h4>Resumen del proyecto</h4><p class="muted">Estado actual y ubicación dentro del flujo HUB.</p></div></div>
    <div class="ficha-summary-grid">
      <div class="info-pill"><small>Dependencia</small><b>${fichaEscape(p.inst)}</b></div>
      <div class="info-pill"><small>Estado</small><b>${fichaEscape(p.state)}</b></div>
      <div class="info-pill"><small>Fase actual</small><b>${fichaEscape(phaseName(phaseOf(p.step)))}</b></div>
      <div class="info-pill"><small>Paso actual</small><b>${p.step}. ${fichaEscape(step.title)}</b></div>
      <div class="info-pill"><small>Responsable</small><b>${fichaEscape(p.owner||step.role)}</b></div>
      <div class="info-pill"><small>Avance metodológico</small><b>${advance}%</b></div>
    </div>
    <div class="progress-line ficha-progress"><span style="width:${advance}%"></span></div>
  </div>

  <div class="ficha-section">
    <h4>Información completa de la base maestra</h4>
    <p class="muted">Se muestran todos los campos disponibles para esta iniciativa en el Excel maestro.</p>
    <div class="ficha-all-data">${rows.map(([h,v])=>`<div class="ficha-data-item"><small>${fichaEscape(fichaLabel(h))}</small><div>${fichaEscape(v)}</div></div>`).join("")}</div>
  </div>


  <div class="ficha-section">
    <h4>Resumen del historial y flujo</h4>
    <p class="muted">Movimientos disponibles del proyecto, con fecha, responsable y decisión registrada.</p>
    <div id="fichaHistoryBox"><p class="muted">Cargando historial…</p></div>
  </div>

  <div class="action-row ficha-bottom-actions">
    <button onclick="viewProjectDocuments('${p.id}')">Ver documentos</button>
    <button class="secondary" onclick="closeFichasModal()">Cerrar ficha</button>
  </div>
 </article>`;
 initFichaInteractiveMap(geo);
 loadFichaHistory(p);
}

async function addProject(){
 if(currentRole==='visitor'||!canRegisterRoles.includes(currentRole)){alert('Su rol no tiene permiso para registrar iniciativas.');return;}
 try{
  await resolveExcelTable();
  const inst=formValue("Institución"); const name=formValue("Nombre del proyecto");
  if(!inst){alert("Seleccione la Institución.");return;} if(!name){alert("Ingrese el Nombre del proyecto.");return;}
  const btn=document.getElementById("saveExcelProjectBtn"),editId=btn?.dataset?.editId||"";
  if(editId){
   const rows=await getExcelRows(),r=rows.find(x=>String(cell(x,"ID HUB")).trim()===editId);if(!r)throw new Error("No se encontró "+editId+" en el Excel.");
   const existing=Object.fromEntries(excelHeaders.map((h,i)=>[h,r.values?.[0]?.[i]??""]));
   await patchExcelRow(r.index,buildExcelValues(editId,existing));
   if(projectGeometryDirty){
    if(!pendingProjectGeometry||!validProjectGeometry(pendingProjectGeometry)){
     throw new Error("La nueva ubicación no es válida. No se modificó la geometría guardada anteriormente.");
    }
    pendingProjectGeometry.source="user";
    await saveProjectGeometry(editId,pendingProjectGeometry);
   }
   window.__editingGeometry=null;pendingProjectGeometry=null;originalProjectGeometry=null;projectGeometryDirty=false;
   if(btn){btn.textContent="Guardar iniciativa";delete btn.dataset.editId;}
   await loadProjectsFromExcel();auditLog.unshift({project:editId,action:"Editó iniciativa",role:actorLabel(),date:new Date().toLocaleString(),comment:"Actualización guardada en Excel maestro."});
   alert(`${editId} actualizado correctamente en Excel.`);return;
  }
  if(!pendingProjectGeometry||!pendingProjectGeometry.referencePoints?.length){alert("Defina la ubicación del proyecto en el mapa.");return;}
  if(pendingProjectGeometry.type==="segment"&&pendingProjectGeometry.referencePoints.length<2){alert("Para un tramo marque al menos dos puntos.");return;}
  if(pendingProjectGeometry.type==="segment"){
   const selectedMode=document.getElementById("segmentRouteMode")?.value;
   if(!selectedMode){alert("Seleccione si el tramo debe calcar rutas existentes o quedar en línea recta.");return;}
   if(selectedMode==="road"&&pendingProjectGeometry.routeMode!=="road")await routeProjectSegment();
   if(selectedMode==="linear"){pendingProjectGeometry.coordinates=pendingProjectGeometry.referencePoints.map(p=>[p.lon,p.lat]);pendingProjectGeometry.routeMode="linear";}
  }
  const registration=await hubApiFetch("/projects",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify(buildProjectApiPayload())});
  const id=String(registration?.idHub||"").trim();
  if(!id)throw new Error("El Worker registró la iniciativa, pero no devolvió el ID HUB.");
  await hubApiFetch(`/projects/${encodeURIComponent(id)}/geometry`,{method:"PUT",headers:{"Content-Type":"application/json"},body:JSON.stringify({...pendingProjectGeometry,source:"user"})});
  pendingProjectGeometry=null;originalProjectGeometry=null;projectGeometryDirty=false;
  await loadProjectsFromExcel();
  auditLog.unshift({project:id,action:"Registró iniciativa",role:actorLabel(),date:new Date().toLocaleString(),comment:"Registro creado mediante HUB API en Excel maestro."});
  renderRegistrationForm();
  const folderName=registration?.carpetaOneDrive?.name||`${id} - ${name}`;
  const folderWarning=registration?.advertenciaCarpeta?`\nAdvertencia de carpeta: ${registration.advertenciaCarpeta}`:"";
  alert(`Iniciativa ${id} registrada correctamente en Excel, ubicación guardada y carpeta gestionada en OneDrive:\n${folderName}${folderWarning}`);
 }catch(e){console.error(e);alert("No fue posible guardar la iniciativa: "+e.message);}
}

resetWorkspace(); populateSelects(); applyRole(); renderFlow(); renderKpis(); renderMatrix(); renderDetail(); renderList(); renderInbox(); renderNotifications(); renderDocuments(); applyGlobalFilters();



window.addEventListener('message',function(e){
  if(e.data && e.data.type==='hub-priorizacion-height'){
    const f=document.getElementById(e.data.target==='proyectos'?'priorizacionProyectosFrame':'priorizacionFrame');
    if(f){f.style.height=Math.max(900,Number(e.data.height)||900)+'px';}
  }
});
function resizePriorizacionFrame(){
 const f=document.getElementById('priorizacionFrame');
 if(!f)return;
 try{const d=f.contentDocument||f.contentWindow.document;const h=Math.max(d.documentElement.scrollHeight,d.body.scrollHeight,900);f.style.height=h+'px';}catch(e){}
}
window.addEventListener('load',()=>{setTimeout(resizePriorizacionFrame,700);setTimeout(resizePriorizacionFrame,2500)});


/* HUB portfolio map independent bootstrap */
(function(){
 let bootTimer=null, bootAttempts=0;
 async function bootPortfolioMapIndependent(){
  const el=document.getElementById("portfolioProjectsMap");
  if(!el || portfolioMapReady || portfolioMapInitializing) return;
  if(el.offsetWidth<100 || el.offsetHeight<100) return;
  bootAttempts++;
  const state=document.getElementById("portfolioMapFilterState");
  if(state) state.innerHTML=`Inicializando mapa… intento ${bootAttempts}`;
  try{
   // Registration/ficha maps already prove Leaflet can work in this HUB.
   // Reuse it if present; otherwise load it.
   if(!window.L) await loadLeaflet();
   if(!window.L) throw new Error("Leaflet no está disponible");

   portfolioMapInitializing=true;
   if(portfolioMap){try{portfolioMap.remove()}catch(e){}}
   el.innerHTML="";

   portfolioMap=L.map(el,{preferCanvas:true}).setView([9.93,-84.08],7);
   L.tileLayer("https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png",{
    maxZoom:19, attribution:"© OpenStreetMap"
   }).addTo(portfolioMap);
   portfolioMapLayer=L.layerGroup().addTo(portfolioMap);

   portfolioMapGeometries.clear();
   const bounds=[];
   for(const p of projects.filter(x=>!x.rejected)){
    const x=inferGeo(p),lat=Number(x.lat),lon=Number(x.lon);
    if(!Number.isFinite(lat)||!Number.isFinite(lon))continue;
    const g={type:"point",source:"inferred",referencePoints:[{lat,lon}]};
    portfolioMapGeometries.set(p.id,g);
    L.circleMarker([lat,lon],{radius:5,weight:1.5,fillOpacity:.75})
      .bindPopup(`<b>${traceEscape(p.id)}</b><br>${traceEscape(p.name)}`)
      .addTo(portfolioMapLayer);
    bounds.push([lat,lon]);
   }
   if(bounds.length)portfolioMap.fitBounds(bounds,{padding:[25,25],maxZoom:10});
   portfolioMap.invalidateSize(true);
   portfolioMapReady=true;
   if(state)state.innerHTML=`Mapa cargado · ${bounds.length} proyectos localizados. Ajuste el área visible y seleccione “Aplicar filtro con el mapa”.`;

   portfolioMap.on("zoomend moveend",()=>{
    clearTimeout(portfolioMapMoveTimer);
    portfolioMapMoveTimer=setTimeout(()=>{
     if(portfolioMapFilterEnabled)applyPortfolioMapFilter();
    },350);
   });

   // Enrich later, never block the visible map.
   setTimeout(()=>preloadPortfolioUserGeometries()
     .then(()=>refreshPortfolioMapLayers())
     .then(()=>portfolioMap&&portfolioMap.invalidateSize(false))
     .catch(e=>console.warn("Geometrías OneDrive:",e)),500);
  }catch(e){
   console.error("Bootstrap mapa:",e);
   portfolioMapReady=false;
   if(state)state.innerHTML=`<b>Error al iniciar mapa:</b> ${traceEscape(e.message||String(e))} <button type="button" class="secondary" onclick="retryPortfolioMap()">Reintentar</button>`;
  }finally{
   portfolioMapInitializing=false;
  }
 }
 window.bootPortfolioMapIndependent=bootPortfolioMapIndependent;

 function startPortfolioWatcher(){
  if(bootTimer)return;
  bootTimer=setInterval(()=>{
   if(portfolioMapReady){clearInterval(bootTimer);bootTimer=null;return;}
   bootPortfolioMapIndependent();
  },700);
 }
 window.startPortfolioWatcher=startPortfolioWatcher;

 if(document.readyState==="loading")document.addEventListener("DOMContentLoaded",startPortfolioWatcher);
 else startPortfolioWatcher();

 // Also watch for the dynamically inserted map container.
 const obs=new MutationObserver(()=>{
  if(document.getElementById("portfolioProjectsMap")&&!portfolioMapReady){
   startPortfolioWatcher();
   setTimeout(bootPortfolioMapIndependent,50);
  }
 });
 obs.observe(document.documentElement,{childList:true,subtree:true});
})();


window.renderFicha=renderFicha;

(()=>{const s=document.createElement('style');s.textContent=`
.custom-filter-builder{margin:14px 0 4px;border:1px solid #d6e2eb;border-radius:18px;background:#fff;padding:18px 20px}
.custom-filter-head{display:flex;justify-content:space-between;gap:18px;align-items:flex-start;margin-bottom:14px}
.custom-filter-head b{display:block;font-size:20px;color:#14243b;margin-bottom:4px}.custom-filter-head span{color:#64748b;white-space:nowrap}
.custom-filter-row{display:grid;grid-template-columns:minmax(260px,1fr) minmax(260px,1fr) auto;gap:10px;margin-bottom:10px}
.custom-filter-row select,.custom-filter-row button,.custom-filter-actions button{min-height:48px;border:1px solid #d5e0e9;border-radius:12px;background:#fff;padding:0 14px;font:inherit}
.custom-filter-row button,.custom-filter-actions button{cursor:pointer;font-weight:700}.custom-filter-actions{display:flex;gap:10px;flex-wrap:wrap;margin-top:12px}.custom-filter-actions .secondary{background:#eef3f7}
@media(max-width:760px){.custom-filter-row{grid-template-columns:1fr}.custom-filter-head{display:block}.custom-filter-head span{display:block;margin-top:7px}}
`;document.head.appendChild(s)})();

(()=>{const s=document.createElement('style');s.textContent=`
#hubOperationProgress{position:fixed;inset:0;background:rgba(10,24,45,.48);z-index:30000;display:none;align-items:center;justify-content:center;padding:20px}
#hubOperationProgress.show{display:flex}.hub-progress-card{width:min(480px,92vw);background:#fff;border-radius:18px;padding:26px 28px;text-align:center;box-shadow:0 20px 60px rgba(0,0,0,.22)}
.hub-progress-card h3{margin:10px 0 6px;color:#123052;font-size:21px}.hub-progress-card p{margin:0 0 18px;color:#64748b}.hub-progress-spinner{width:42px;height:42px;margin:auto;border:4px solid #dce7f0;border-top-color:#0b4d8c;border-radius:50%;animation:hubspin .8s linear infinite}
.hub-progress-track{height:12px;background:#e7eef5;border-radius:999px;overflow:hidden}.hub-progress-track>div{height:100%;width:0;background:linear-gradient(90deg,#0b4d8c,#07858d);transition:width .28s ease}#hubProgressPct{margin-top:8px;font-size:13px;font-weight:800;color:#0b4d8c}@keyframes hubspin{to{transform:rotate(360deg)}}
`;document.head.appendChild(s)})();

(()=>{const s=document.createElement('style');s.textContent=`
/* Corrección de legibilidad - Constructor de filtros */
#customPortfolioFilterBuilder .custom-filter-row button,
#customPortfolioFilterBuilder .custom-filter-actions button{
  color:#123052 !important;
  background:#ffffff !important;
  border:1px solid #cbd9e6 !important;
  font-weight:700 !important;
  opacity:1 !important;
  -webkit-text-fill-color:#123052 !important;
}
#customPortfolioFilterBuilder .custom-filter-actions button:first-child{
  background:#0b4d8c !important;
  color:#ffffff !important;
  -webkit-text-fill-color:#ffffff !important;
  border-color:#0b4d8c !important;
}
#customPortfolioFilterBuilder .custom-filter-actions .secondary{
  background:#eef3f7 !important;
  color:#24364d !important;
  -webkit-text-fill-color:#24364d !important;
}
#customPortfolioFilterBuilder select{
  color:#17233b !important;
  background-color:#ffffff !important;
  opacity:1 !important;
}
#customPortfolioFilterBuilder select:disabled{
  color:#64748b !important;
  background:#f4f7fa !important;
  opacity:1 !important;
}
`;document.head.appendChild(s)})();

(()=>{const s=document.createElement('style');s.textContent=`
#customPortfolioFilterBuilder .range-filter{display:grid;grid-template-columns:1fr auto 1fr;gap:8px;align-items:center}
#customPortfolioFilterBuilder .range-filter input{min-height:48px;border:1px solid #d5e0e9;border-radius:12px;padding:0 14px;font:inherit;color:#17233b;background:#fff;min-width:0}
#customPortfolioFilterBuilder .hub-multi{position:relative}
#customPortfolioFilterBuilder .hub-multi-trigger{width:100%;text-align:left;display:flex;align-items:center;justify-content:space-between}
#customPortfolioFilterBuilder .hub-multi-menu{display:none;position:absolute;z-index:12000;top:calc(100% + 5px);left:0;right:0;max-height:280px;overflow:auto;background:#fff;border:1px solid #cbd9e6;border-radius:12px;box-shadow:0 12px 30px rgba(20,45,70,.16);padding:8px}
#customPortfolioFilterBuilder .hub-multi-menu.open{display:block}
#customPortfolioFilterBuilder .hub-multi-menu label{display:flex;gap:9px;align-items:flex-start;padding:8px;border-radius:8px;cursor:pointer;color:#17233b}
#customPortfolioFilterBuilder .hub-multi-menu label:hover{background:#f1f6fa}
#customPortfolioFilterBuilder .hub-multi-menu input{margin-top:3px}
`;document.head.appendChild(s)})();

(()=>{const s=document.createElement('style');s.textContent=`
/* Selector múltiple compacto y legible */
#customPortfolioFilterBuilder .hub-multi-menu{
  left:0 !important; right:0 !important; width:100% !important;
  max-height:300px !important; padding:6px !important;
}
#customPortfolioFilterBuilder .hub-multi-menu label{
  display:grid !important;
  grid-template-columns:22px minmax(0,1fr) !important;
  align-items:center !important;
  gap:10px !important;
  padding:9px 10px !important;
  min-height:38px !important;
  text-align:left !important;
  letter-spacing:normal !important;
  font-size:14px !important;
  line-height:1.25 !important;
  text-transform:none !important;
}
#customPortfolioFilterBuilder .hub-multi-menu label span{
  display:block !important;
  width:auto !important;
  text-align:left !important;
  letter-spacing:normal !important;
  word-break:normal !important;
  overflow-wrap:anywhere !important;
  color:#17233b !important;
  font-weight:500 !important;
}
#customPortfolioFilterBuilder .hub-multi-menu label input[type="checkbox"]{
  width:17px !important;height:17px !important;margin:0 !important;
  justify-self:center !important;
}
#customPortfolioFilterBuilder .hub-multi-trigger{
  min-height:48px !important;
  color:#17233b !important;
  background:#fff !important;
  font-weight:700 !important;
}
`;document.head.appendChild(s)})();
