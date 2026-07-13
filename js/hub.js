
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
let projects = window.HUB_PROJECTS;
let activeId=projects[0]?.id||"";
let currentRole="visitor";
const canRegisterRoles=["tecnico","coordinador","secretaria"];
const roleStepPermissions={tecnico:[],jefatura:[2],planeamiento:[3],direccion:[4],coordinador:[5],secretaria:[6,8],areas:[7],comite:[9],ejecutor:[10,11,12,13,14,15,16,17]};
let auditLog=[];
const docEditPermissions={
 tecnico:["Ficha de iniciativa","Diagnóstico / justificación"],
 jefatura:["Ficha de iniciativa","Diagnóstico / justificación","Estimación de costos"],
 planeamiento:["Aval institucional"],
 direccion:["Aval institucional"],
 coordinador:["Ficha de iniciativa","Aval institucional","Plan de ejecución"],
 secretaria:["Concepto técnico","Concepto financiero","Plan de ejecución"],
 areas:["Concepto técnico"],
 secretaria:["Concepto financiero"],
 comite:[],
 ejecutor:["Plan de ejecución","Estimación de costos"]
};
const documents=[
 {project:"M-20-001",type:"Ficha de iniciativa",name:"Ficha inicial del proyecto",state:"En revisión",owner:"Técnico / Profesional"},
 {project:"M-20-001",type:"Diagnóstico / justificación",name:"Justificación de necesidad",state:"Borrador",owner:"Técnico / Profesional"},
 {project:"M-04-001",type:"Estimación de costos",name:"Estimación preliminar de inversión",state:"Aprobado",owner:"Jefatura Técnica"},
 {project:"V-25-001",type:"Aval institucional",name:"Aval de presentación ante HUB",state:"En revisión",owner:"Dirección / Gerencia"},
 {project:"C-25-003",type:"Concepto técnico",name:"Concepto técnico central",state:"Devuelto para ajustes",owner:"Áreas Técnicas MOPT"},
 {project:"F-25-001",type:"Matriz de criterios",name:"Calificación automatizada y priorización preliminar",state:"Aprobado",owner:"Sistema HUB / Secretaría Técnica HUB"},
 {project:"C-25-002",type:"Plan de ejecución",name:"Cronograma e hitos de ejecución",state:"Borrador",owner:"Institución ejecutora"}
];
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
 if(id==="matriz")renderMatrix(); if(id==="historial")renderDetail(); if(id==="listado")renderList(); if(id==="bandeja")renderInbox(); if(id==="notificaciones")renderNotifications(); if(id==="documentos")renderDocuments(); if(id==="priorizacion"){ const f=document.getElementById("priorizacionFrame"); setTimeout(()=>{try{f.contentWindow.dispatchEvent(new Event("resize"));}catch(e){}},350); }
}
function previewRole(){
 const r=roleSelect.value;
 if(!r){roleCard.classList.add("hidden");return;}
 const items=roleInfo[r]||[];
 roleCard.classList.remove("hidden");
 roleCard.innerHTML=`<h3>${roleSelect.options[roleSelect.selectedIndex].text}</h3><ul>${items.map(i=>`<li>${i}</li>`).join("")}</ul>`;
}
function loginRole(){
 const r=roleSelect.value;
 if(!r){alert("Seleccione un rol para iniciar sesión.");return;}
 currentRole=r;
 sessionLabel.textContent=" · Sesión iniciada como "+roleSelect.options[roleSelect.selectedIndex].text;
 document.querySelectorAll(".protected-tab").forEach(t=>t.classList.remove("hidden"));
 const registerBtn=[...document.querySelectorAll(".protected-tab")].find(t=>t.textContent.trim()=="Registrar");
 if(registerBtn && !canRegisterRoles.includes(currentRole)) registerBtn.classList.add("hidden");
 logoutBtn.classList.remove("hidden");
 previewRole();
 renderInbox(); renderNotifications();
}
function logoutRole(){
 currentRole="visitor";
 roleSelect.value="";
 sessionLabel.textContent=" · Información pública del portafolio";
 roleCard.classList.add("hidden");
 logoutBtn.classList.add("hidden");
 document.querySelectorAll(".protected-tab").forEach(t=>t.classList.add("hidden"));
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
 const header=`<div class="phase-ribbon"><div class="phase1">FASE 1 · Gestión institucional</div><div class="phase2">FASE 2 · Evaluación central</div><div class="phase3">FASE 3 · Preparación</div><div class="phase4">FASE 4 · Ejecución y cierre</div></div>
 <div class="header-row"><div class="project-label">Proyecto</div>${steps.map(s=>`<div class="header-cell ${clsStep(s.n)}" title="${s.title}">${s.n}<br>${s.title.split(" ")[0]}</div>`).join("")}</div>`;
 const rows=list.map(p=>`<div class="project-row" onclick="selectProject('${p.id}')">
 <div class="project-label" title="${p.name}">${p.id} · ${p.name}<small>${p.inst} · ${p.state} · ${phaseName(phaseOf(p.step))}</small></div>
 ${steps.map(s=>`<div class="cell"><span class="dot ${dotClass(p,s.n)}" title="Paso ${s.n} · ${s.title}"></span></div>`).join("")}</div>`).join("");
 const counts=`<div class="counts-row"><div class="project-label">Proyectos que llegaron al paso</div>${steps.map(s=>`<div class="cell"><span class="count">${list.filter(p=>p.step>=s.n).length}</span></div>`).join("")}</div>`;
 matrixBox.innerHTML=header+(rows||'<div class="note">No hay proyectos con esos filtros.</div>')+counts;
}
function selectProject(id){activeId=id; if(projectSelect)projectSelect.value=id; openHistoryModal(id);}
function renderDetail(targetId){
 const p=projects.find(x=>x.id==(targetId||activeId))||projects[0]; if(!p)return "";
 if(projectSelect) projectSelect.value=p.id;
 const step=steps.find(s=>s.n==p.step);
 let timeline=steps.map(s=>`<div class="titem ${dotClass(p,s.n)}"><div class="tnum">${s.n}</div><div class="tbox"><b>${s.title}</b><span class="muted">${s.role} · ${s.desc}</span>${s.n==p.step?`<br><br><span class="badge ${p.rejected?'red':p.returned?'orange':'yellow'}">Paso actual: ${p.state}</span>`:''}</div></div>`).join("");
 const content=`<div class="detail"><div><h3>${p.id} · ${p.name}</h3><p class="muted"><b>Dependencia:</b> ${p.inst}<br><b>BPIP:</b> ${p.bpip||'Por definir'}<br><b>Categoría:</b> ${p.categoria} / ${p.subcategoria}<br><b>Etapa base:</b> ${p.etapa}<br><b>Estado:</b> ${p.state}</p><div class="note"><b>Ubicación actual:</b> ${phaseName(phaseOf(p.step))}, paso ${p.step}: ${step.title}. Responsable sugerido: ${p.owner}.</div></div><div><h3>Lectura del avance</h3><p><span class="badge green">Cumplido</span> <span class="badge yellow">Actual</span> <span class="badge orange">Devuelto</span> <span class="badge red">Rechazado</span></p><p class="muted">Esta ventana resume la trazabilidad del proyecto sin perder la navegación del listado.</p></div></div><hr style="border:0;border-top:1px solid var(--line);margin:22px 0"><h3>Línea de tiempo del proyecto</h3><div class="timeline">${timeline}</div>`;
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
}
function closeHistoryModal(e){
 if(e && e.target!==historyModal)return;
 historyModal.classList.remove("open");
}
function renderList(){
 const list=getGlobalFilteredProjects();
 const table=`<table><thead><tr><th>ID</th><th>Proyecto</th><th>Dependencia</th><th>Fase actual</th><th>Paso</th><th>Estado</th><th>Categoría</th><th>Vinculación</th><th>Acciones</th></tr></thead><tbody>${list.slice(0,400).map(p=>{
   const canManage=canManageProject(p);
   return `<tr><td><b>${p.id}</b><br><span class="muted">${p.bpip||""}</span></td><td>${p.name}</td><td>${p.inst}</td><td>${phaseName(phaseOf(p.step))}</td><td>${p.step}</td><td><span class="badge">${p.state}</span></td><td>${p.categoria||""}<br><span class="muted">${p.subcategoria||""}</span></td><td>${p.priority||""}</td><td><div class="doc-actions"><button onclick="openHistoryModal('${p.id}')">Historial</button>${canManage?`<button onclick="openManageModal('${p.id}')">Gestionar</button>`:""}</div></td></tr>`;
 }).join("")}</tbody></table><p class="muted">Mostrando ${Math.min(list.length,400)} de ${list.length} resultados filtrados.</p>`;
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
function openManageModal(id){
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
    <div class="doc-drop">Cargar soporte / nueva versión<br><small>En este prototipo se simula el cargue documental.</small></div>
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
 auditLog.unshift({project:id,action,role:roleSelect.options[roleSelect.selectedIndex]?.text||currentRole,date:new Date().toLocaleString(),comment});
 renderKpis();renderList();renderMatrix();renderInbox();renderNotifications();
 openManageModal(id);
 alert("Gestión registrada. Se actualizó la bitácora y se generó la notificación correspondiente.");
}
function simulateDocUpload(id){
 const p=projects.find(x=>x.id==id); if(!p)return;
 documents.unshift({project:id,type:"Soporte cargado",name:"Nuevo soporte documental",state:"Borrador",owner:roleSelect.options[roleSelect.selectedIndex]?.text||currentRole});
 auditLog.unshift({project:id,action:"Cargó soporte documental",role:roleSelect.options[roleSelect.selectedIndex]?.text||currentRole,date:new Date().toLocaleString(),comment:"Se cargó una nueva versión documental."});
 renderDocuments();
 openManageModal(id);
}

function canEditDoc(type){
 if(currentRole==="visitor")return false;
 return (docEditPermissions[currentRole]||[]).includes(type);
}
function renderDocuments(){
 if(currentRole==="visitor"){documentsBox.innerHTML='<div class="note">Inicie sesión para consultar documentos internos.</div>';return;}
 const q=norm(docSearch?.value||""), t=docTypeFilter?.value||"", st=docStateFilter?.value||"";
 const perms=docEditPermissions[currentRole]||[];
 docPermissionNote.innerHTML=perms.length?`Su rol puede editar: <b>${perms.join(", ")}</b>. Los demás documentos quedan en solo lectura.`:"Su rol puede consultar documentos, pero no tiene permiso de edición en esta etapa.";
 const list=documents.filter(d=>(!q||norm([d.project,d.type,d.name,d.state,d.owner].join(" ")).includes(q))&&(!t||d.type==t)&&(!st||d.state==st));
 documentsBox.innerHTML=`<table><thead><tr><th>Proyecto</th><th>Documento</th><th>Tipo</th><th>Estado</th><th>Responsable</th><th>Acciones según rol</th></tr></thead><tbody>${list.map(d=>`<tr><td><b>${d.project}</b></td><td>${d.name}</td><td>${d.type}</td><td><span class="badge">${d.state}</span></td><td>${d.owner}</td><td><div class="doc-actions"><button onclick="alert('Vista previa documental del soporte seleccionado.')">Ver</button>${canEditDoc(d.type)?`<button onclick="alert('Edición habilitada para este rol.')">Editar</button><button onclick="alert('Nueva versión cargada y registrada en trazabilidad.')">Cargar versión</button>`:`<button class="disabled" onclick="alert('Documento en solo lectura para este rol.')">Solo lectura</button>`}</div></td></tr>`).join("")}</tbody></table>`;
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
 renderFichaResults();
 renderFicha(activeId||projects[0]?.id);
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
 if(fichaSelect) fichaSelect.value=p.id;
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
      <div class="action-row"><button onclick="openHistoryModal('${p.id}')">Ver historial</button><button class="secondary" onclick="showView('matriz',[...document.querySelectorAll('.public-tab')].find(t=>t.textContent.trim()==='Matriz de proyectos'))">Ver en matriz</button></div>
    </div>
  </div>
 </article>`;
 renderFichaResults();
}

function addProject(){
 if(currentRole==='visitor'||!canRegisterRoles.includes(currentRole)){alert('Su rol no tiene permiso para registrar iniciativas.');return;}
 const id="N-"+String(projects.length+1).padStart(3,"0");
 projects.unshift({id,name:newName.value||"Nueva iniciativa",inst:newInst.value||"Sin dependencia",step:1,state:"Borrador",priority:newPriority.value,bpip:newBpip.value||"Por definir",categoria:newCat.value||"Sin categoría",subcategoria:"",etapa:"Idea",owner:"Técnico / Profesional",returned:false,rejected:false});
 populateSelects(); renderKpis(); renderMatrix(); renderList(); renderFichaResults(); renderFicha(id); selectProject(id); alert("Iniciativa registrada en el paso 1. Se notificó a la jefatura técnica para revisión.");
}
populateSelects(); applyRole(); renderFlow(); renderKpis(); renderMatrix(); renderDetail(); renderList(); renderInbox(); renderNotifications(); renderDocuments(); renderFichaResults(); renderFicha(projects[0]?.id); applyGlobalFilters();



window.addEventListener('message',function(e){
  if(e.data && e.data.type==='hub-priorizacion-height'){
    const f=document.getElementById('priorizacionFrame');
    if(f){f.style.height=Math.max(900,Number(e.data.height)||900)+'px';}
  }
});
function resizePriorizacionFrame(){
 const f=document.getElementById('priorizacionFrame');
 if(!f)return;
 try{const d=f.contentDocument||f.contentWindow.document;const h=Math.max(d.documentElement.scrollHeight,d.body.scrollHeight,900);f.style.height=h+'px';}catch(e){}
}
window.addEventListener('load',()=>{setTimeout(resizePriorizacionFrame,700);setTimeout(resizePriorizacionFrame,2500)});
