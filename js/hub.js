
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
const BASE_PROJECTS = JSON.parse(JSON.stringify(window.HUB_PROJECTS||[]));
let projects = JSON.parse(JSON.stringify(BASE_PROJECTS));
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
 projects=JSON.parse(JSON.stringify(BASE_PROJECTS));
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
  if(Array.isArray(st.projects)&&st.projects.length)projects=st.projects;
  documents=[]; // Los documentos se cargan únicamente desde OneDrive.
  if(Array.isArray(st.auditLog))auditLog=st.auditLog;
  activeId=projects[0]?.id||"";
 }catch(e){console.warn("No fue posible recuperar el estado HUB",e);}
}
function actorLabel(){
 const roleName=roleSelect?.options?.[roleSelect.selectedIndex]?.text||currentRole;
 return currentUser?`${currentUser.name} · Perfil: ${roleName}`:currentRole;
}
// ===== Microsoft OneDrive / Graph =====
const MS_CONFIG={
 clientId:"3a5ead05-e9e6-477b-9a63-24ef1f9fefd5",
 tenantId:"b3eacfd6-2fae-4231-8cc6-ad611b3240f1",
 redirectUri:"https://cmariam10.github.io/Priorizacion-TEMPORAL/",
 rootShareUrl:"https://mariampenar-my.sharepoint.com/:f:/g/personal/mariampena_mariampenar_onmicrosoft_com/IgAqBC0CtqOmS6-OH1fqcWt2AaAj1L_cMwOuhSdlPvdU6Lw?e=oIbXkP",
 excelShareUrl:"https://mariampenar-my.sharepoint.com/:x:/g/personal/mariampena_mariampenar_onmicrosoft_com/IQDKEAZoAhkNTrGSyIHzJhqEAbmFWdlEVNOFICupivCl3rM?e=tGqA5x"
};
const GRAPH_SCOPES=["User.Read","Files.ReadWrite"];
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
async function graphToken(interactive=true){
 await initMicrosoft();
 graphAccount=msalApp.getActiveAccount()||msalApp.getAllAccounts()[0]||graphAccount;
 if(!graphAccount&&interactive){const r=await msalApp.loginPopup({scopes:GRAPH_SCOPES,prompt:"select_account"});graphAccount=r.account;msalApp.setActiveAccount(graphAccount);}
 if(!graphAccount)throw new Error("Debe iniciar sesión con Microsoft para acceder a OneDrive.");
 try{return (await msalApp.acquireTokenSilent({scopes:GRAPH_SCOPES,account:graphAccount})).accessToken;}
 catch(e){if(!interactive)throw e;return (await msalApp.acquireTokenPopup({scopes:GRAPH_SCOPES,account:graphAccount})).accessToken;}
}
async function graphFetch(path,options={}){
 const token=await graphToken(true); const headers=new Headers(options.headers||{}); headers.set("Authorization",`Bearer ${token}`);
 const res=await fetch(path.startsWith("http")?path:`https://graph.microsoft.com/v1.0${path}`,{...options,headers});
 if(!res.ok){let detail="";try{detail=(await res.json())?.error?.message||""}catch(e){}throw new Error(`Microsoft Graph ${res.status}: ${detail||res.statusText}`);}
 if(res.status===204)return null; const ct=res.headers.get("content-type")||""; return ct.includes("json")?res.json():res.blob();
}

// ===== Base maestra de proyectos: Excel en OneDrive / SharePoint =====
let excelBook=null, excelTable=null, excelHeaders=[];
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
 const token=encodeShareUrl(MS_CONFIG.excelShareUrl);
 const item=await graphFetch(`/shares/${token}/driveItem?$select=id,name,webUrl,parentReference,file`);
 const driveId=item.parentReference?.driveId;if(!driveId)throw new Error("No se pudo identificar el drive del Excel maestro.");
 excelBook={driveId,itemId:item.id,name:item.name,webUrl:item.webUrl};return excelBook;
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
 const rows=await getExcelRows();
 projects=rows.filter(r=>String(cell(r,"ID HUB")||"").trim()).filter(r=>hnorm(cell(r,"Estado HUB"))!=="eliminado").map(excelRowToProject);
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
 renderRegistrationForm(Object.fromEntries(excelHeaders.map((h,i)=>[h,r.values?.[0]?.[i]??""])));showView("registrar");
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
 const token=encodeShareUrl(MS_CONFIG.rootShareUrl);
 const item=await graphFetch(`/shares/${token}/driveItem?$select=id,name,webUrl,parentReference,folder`);
 if(!item?.folder)throw new Error("El enlace configurado no corresponde a una carpeta de OneDrive/SharePoint.");
 const driveId=item.parentReference?.driveId; if(!driveId)throw new Error("Microsoft Graph no devolvió el driveId de HUB_Proyectos_MOPT.");
 graphRoot={driveId,itemId:item.id,name:item.name,webUrl:item.webUrl}; return graphRoot;
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
 const folder=await ensureProjectFolder(project); let path;
 if(itemId)path=`/drives/${folder.driveId}/items/${itemId}/content`;
 else path=`/drives/${folder.driveId}/items/${folder.itemId}:/${encodeURIComponent(file.name)}:/content`;
 return graphFetch(path,{method:"PUT",headers:{"Content-Type":file.type||"application/octet-stream"},body:file});
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
 const roleName=roleSelect?.options?.[roleSelect.selectedIndex]?.text||currentRole;
 roleCard.classList.remove("hidden");
 roleCard.innerHTML=`<h3>${currentUser.name} · Perfil activo: ${roleName}</h3><ul>${items.map(i=>`<li>${i}</li>`).join("")}</ul>`;
}
function refreshRoleAccess(){
 document.querySelectorAll(".protected-tab").forEach(t=>t.classList.remove("hidden"));
 const registerBtn=[...document.querySelectorAll(".protected-tab")].find(t=>t.textContent.trim()==="Registrar");
 if(registerBtn) registerBtn.classList.toggle("hidden",!canRegisterRoles.includes(currentRole));
 previewRole(); renderInbox(); renderNotifications(); renderDocuments();
}
function changeTestProfile(){
 if(!currentUser)return;
 const roleEl=document.getElementById("roleSelect");
 const sessionEl=document.getElementById("sessionLabel");
 if(!roleEl)return;
 currentRole=roleEl.value||"tecnico";
 if(sessionEl) sessionEl.textContent=` · ${currentUser.name} · Perfil activo: ${roleEl.options[roleEl.selectedIndex].text}`;
 refreshRoleAccess();
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
   loadHubState();
   currentRole=selectedRole;
   if(roleEl){roleEl.value=currentRole; roleEl.classList.remove("hidden"); roleEl.removeAttribute("aria-hidden");}
   if(logoutEl) logoutEl.classList.remove("hidden");
   if(usernameEl) usernameEl.disabled=true;
   if(passwordEl) passwordEl.disabled=true;
   populateSelects(); applyGlobalFilters(); renderFlow(); renderKpis(); renderMatrix(); renderDetail(); renderList();
   refreshRoleAccess();
   try{await loadProjectsFromExcel();}catch(excelErr){console.error(excelErr);alert("Sesión iniciada, pero no fue posible cargar la base maestra de Excel: "+excelErr.message);}
   const roleName=roleEl?.options?.[roleEl.selectedIndex]?.text||currentRole;
   if(sessionEl) sessionEl.textContent=` · ${account.name} · Perfil activo: ${roleName}`;
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
}
window.loginRole=loginRole;
function logoutRole(){
 saveHubState();
 currentRole="visitor"; currentUser=null; roleSelect.value=""; roleSelect.classList.remove("hidden"); roleSelect.removeAttribute("aria-hidden");
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
function getGlobalFilteredProjects(){
 let q=globalSearch?.value||"", ph=globalPhase?.value||"", inst=globalInst?.value||"", st=globalState?.value||"", cat=globalCat?.value||"", sub=globalSubcat?.value||"", pri=globalPriority?.value||"", step=globalStep?.value||"";
 return projects.filter(p=>matches(p,q)&&(!ph||phaseOf(p.step)==+ph)&&(!inst||p.inst==inst)&&(!st||p.state==st)&&(!cat||p.categoria==cat)&&(!sub||p.subcategoria==sub)&&(!pri||p.priority==pri)&&(!step||p.step==+step));
}
function renderKpis(){
 const list=getGlobalFilteredProjects();
 kTotal.textContent=list.length;
 kF1.textContent=list.filter(p=>phaseOf(p.step)==1).length;
 kF2.textContent=list.filter(p=>phaseOf(p.step)==2).length;
 kRech.textContent=list.filter(p=>p.rejected).length;
 if(homeCounter)homeCounter.innerHTML=`${list.length} de ${projects.length}<br><span style="font-size:12px;color:var(--muted)">proyectos visibles</span>`;
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
    @media(max-width:900px){
      .project-list-table{min-width:900px}
      .project-doc-item{align-items:flex-start;flex-direction:column}
    }
  `;
  document.head.appendChild(s);
})();


function renderList(){
  const list=getGlobalFilteredProjects();
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
      return `<tr>
        <td class="pl-id"><b class="project-id">${p.id}</b><span class="project-bpip">BPIP ${p.bpip||"Por definir"}</span></td>
        <td class="pl-project"><div class="project-title">${p.name}</div><div class="project-sub">${p.inst||"Sin dependencia"}</div></td>
        <td class="pl-info"><span class="badge">${p.state||"Sin estado"}</span><div class="project-sub">${phaseName(phaseOf(p.step))} · Paso ${p.step}</div></td>
        <td class="pl-class"><div>${p.categoria||"Sin categoría"}</div><div class="project-sub">${p.subcategoria||""}</div>${p.priority?`<div class="project-linkage">${p.priority}</div>`:""}</td>
        <td class="pl-actions"><div class="project-action-grid">
          <button onclick="openHistoryModal('${p.id}')">Historial</button>
          <button onclick="viewProjectDocuments('${p.id}')">Documentos</button>
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
 inboxBox.innerHTML=list.length?`<table><thead><tr><th>Proyecto</th><th>Acción requerida</th><th>Estado</th><th>Gestión</th></tr></thead><tbody>${list.slice(0,100).map(p=>`<tr><td><b>${p.id}</b><br>${p.name}</td><td>${currentRole==="tecnico"?"Registrar, completar ficha o atender devolución":"Revisar paso "+p.step+": "+(steps.find(s=>s.n==p.step)?.title||"")}</td><td>${p.state}</td><td><button onclick="openManageModal('${p.id}')">Gestionar</button></td></tr>`).join("")}</tbody></table>`:'<div class="note">No hay proyectos pendientes para este rol.</div>';
}
function renderNotifications(){
 if(currentRole==="visitor"){notifBox.innerHTML='<div class="note">Inicie sesión seleccionando un rol para consultar notificaciones internas.</div>';return;}
 const roleMap={tecnico:1,jefatura:2,planeamiento:3,direccion:4,coordinador:5,secretaria:8,areas:7,comite:9,ejecutor:10};
 const target=roleMap[currentRole];
 const list=projects.filter(p=>p.step==target || (currentRole==="ejecutor"&&p.step>=10)).slice(0,20);
 notifBox.innerHTML=(list.length?list:projects.slice(0,10)).map(p=>`<div class="note"><b>${p.id}</b> · ${p.name}<br>Notificación: el proyecto se encuentra en <b>paso ${p.step}</b>. Responsable actual: ${p.owner}.</div>`).join("");
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
function manageDecision(id,decision){
 const p=projects.find(x=>x.id==id); if(!p)return;
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
 auditLog.unshift({project:id,action,role:actorLabel(),date:new Date().toLocaleString(),comment});
 saveHubState();
 renderKpis();renderList();renderMatrix();renderInbox();renderNotifications();
 openManageModal(id);
 alert("Gestión registrada. Se actualizó la bitácora y se generó la notificación correspondiente.");
}
async function simulateDocUpload(id){
 const p=projects.find(x=>x.id==id); if(!p)return;
 const input=document.createElement("input");input.type="file";input.accept=".pdf,.doc,.docx,.xls,.xlsx,.png,.jpg,.jpeg,.txt";
 input.onchange=async()=>{
  const file=input.files?.[0];if(!file)return;
  try{
   const item=await uploadOneDriveFile(p,file);
   documents=documents.filter(d=>!(d.project===id&&d.itemId===item.id));
   documents.unshift({project:id,type:"Soporte cargado",name:item.name,state:"En OneDrive",owner:actorLabel(),itemId:item.id,driveId:item.parentReference?.driveId,webUrl:item.webUrl,size:item.size,uploadedAt:item.lastModifiedDateTime});
   auditLog.unshift({project:id,action:"Cargó soporte documental",role:actorLabel(),date:new Date().toLocaleString(),comment:`Se cargó ${item.name} en OneDrive.`});
   saveHubState();await writeTrace(p,"SUBIO_DOCUMENTO",item.name,{itemId:item.id});renderDocuments();openManageModal(id,true);alert("Documento guardado correctamente en OneDrive y registrado en trazabilidad.");
  }catch(e){console.error(e);alert("No fue posible guardar el archivo en OneDrive: "+e.message);}
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

function openFichasModal(){
 const modal=document.getElementById('fichas');
 if(!modal) return;
 modal.classList.remove('hidden');
 modal.classList.add('ficha-modal-open');
 document.body.classList.add('fichas-open');
}
function closeFichasModal(){
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

function realMapIframe(p,geo){
 const lat=Number(geo.lat)||9.93, lon=Number(geo.lon)||-84.08;
 const delta=geo.feature?.type?.includes('Line') ? 0.12 : 0.035;
 const bbox=[(lon-delta).toFixed(5),(lat-delta).toFixed(5),(lon+delta).toFixed(5),(lat+delta).toFixed(5)].join('%2C');
 const marker=`${lat.toFixed(5)}%2C${lon.toFixed(5)}`;
 const src=`https://www.openstreetmap.org/export/embed.html?bbox=${bbox}&layer=mapnik&marker=${marker}`;
 const href=`https://www.openstreetmap.org/?mlat=${lat.toFixed(5)}&mlon=${lon.toFixed(5)}#map=14/${lat.toFixed(5)}/${lon.toFixed(5)}`;
 return `<iframe title="Mapa real del proyecto ${p.id}" loading="lazy" referrerpolicy="no-referrer-when-downgrade" src="${src}"></iframe><div class="map-chip">Punto de referencia · ${geo.fuente}</div>`;
}
function realMapLink(geo){
 const lat=Number(geo.lat)||9.93, lon=Number(geo.lon)||-84.08;
 return `https://www.openstreetmap.org/?mlat=${lat.toFixed(5)}&mlon=${lon.toFixed(5)}#map=14/${lat.toFixed(5)}/${lon.toFixed(5)}`;
}

function renderFicha(id){
 const p=projects.find(x=>x.id==(id||activeId))||projects[0]; if(!p)return;
 activeId=p.id;
 const modalTitle=document.getElementById('fichaModalTitle');
 if(modalTitle) modalTitle.textContent=`${p.id} · ${p.name}`;
 const geo=inferGeo(p), score=projectScore(p), step=steps.find(s=>s.n==p.step)||steps[0], advance=Math.round((p.step/17)*100);
 const criteria=[
  ["Impacto / prioridad",p.priority||"Sin dato"],["Madurez del proyecto",p.etapa||"Idea"],["Clasificación",p.categoria||"Sin categoría"],["Ámbito",geo.ambito],["Puntaje HUB estimado",score+" / 100"]
 ];
 fichaBox.innerHTML=`<article class="ficha-card">
  <div class="ficha-top">
    <div>
      <h3>${p.name}</h3>
      <p><b>${p.inst}</b> · ${phaseName(phaseOf(p.step))} · Paso ${p.step}: ${step.title}</p>
      <div class="ficha-badges"><span class="badge">${p.state}</span><span class="badge">${p.priority||'Sin dato'}</span><span class="badge">${geo.lugar}</span></div>
    </div>
    <div class="ficha-code"><small>ID HUB</small><b>${p.id}</b><small>BPIP: ${p.bpip||'Por definir'}</small></div>
  </div>
  <div class="ficha-body">
    <div class="geo-map-card"><h4>Mapa del proyecto</h4><div class="geo-map">${realMapIframe(p,geo)}</div><a class="map-link" href="${realMapLink(geo)}" target="_blank" rel="noopener">Abrir mapa ampliado ↗</a><div class="info-grid"><div class="info-pill"><small>Coordenadas</small><b>${geo.lat}, ${geo.lon}</b></div><div class="info-pill"><small>Ubicación</small><b>${geo.lugar}</b></div><div class="info-pill"><small>Fuente espacial</small><b>${geo.fuente}</b></div><div class="info-pill"><small>Capa</small><b>${geo.feature?geo.feature.layer:'Sin coincidencia exacta'}</b></div></div></div>
    <div class="geo-data-card">
      <h4>Datos básicos e infografía de avance</h4>
      <div class="info-grid">
        <div class="info-pill"><small>Dependencia</small><b>${p.inst}</b></div>
        <div class="info-pill"><small>Estado</small><b>${p.state}</b></div>
        <div class="info-pill"><small>Categoría</small><b>${p.categoria||'Sin dato'}</b></div>
        <div class="info-pill"><small>Subcategoría</small><b>${p.subcategoria||'Sin dato'}</b></div>
        <div class="info-pill"><small>Etapa base</small><b>${p.etapa||'Sin dato'}</b></div>
        <div class="info-pill"><small>Responsable actual</small><b>${p.owner||step.role}</b></div>
      </div>
      <div class="note"><b>Avance dentro del flujo HUB:</b> ${advance}% del recorrido metodológico. <div class="progress-line"><span style="width:${advance}%"></span></div></div>
      <h4 style="margin-top:16px">Criterios visibles de priorización</h4>
      <div class="criteria-list">${criteria.map(c=>`<div class="criteria-item"><span>${c[0]}</span><b>${c[1]}</b></div>`).join("")}</div>
      <div class="action-row"><button onclick="openHistoryModal('${p.id}')">Ver historial</button><button class="secondary" onclick="closeFichasModal()">Cerrar ficha</button></div>
    </div>
  </div>
 </article>`;
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
   if(btn){btn.textContent="Guardar iniciativa";delete btn.dataset.editId;}
   await loadProjectsFromExcel();auditLog.unshift({project:editId,action:"Editó iniciativa",role:actorLabel(),date:new Date().toLocaleString(),comment:"Actualización guardada en Excel maestro."});
   alert(`${editId} actualizado correctamente en Excel.`);return;
  }
  const id=await generateHubId(inst);const values=buildExcelValues(id,{});await appendExcelProject(values);
  await loadProjectsFromExcel();const p=projects.find(x=>x.id===id)||{id,name,inst};
  auditLog.unshift({project:id,action:"Registró iniciativa",role:actorLabel(),date:new Date().toLocaleString(),comment:"Registro creado en Excel maestro."});
  const folder=await ensureProjectFolder(p);await writeTrace(p,"CREO_PROYECTO_Y_CARPETA",folder.name,{folderId:folder.itemId});
  renderRegistrationForm();alert(`Iniciativa ${id} registrada en Excel y carpeta creada en OneDrive:
${folder.name}`);
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
