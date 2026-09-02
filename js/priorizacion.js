
const projects = window.PRIORIZACION_PROJECTS;
const scenarios={
 'Base metodología':{nd:25,is:25,vi:15,rr:15,si:10,ger:10},
 'Ambiental':{nd:10,is:15,vi:10,rr:25,si:35,ger:5},
 'Social / equidad':{nd:20,is:40,vi:10,rr:15,si:10,ger:5},
 'Económico / logístico':{nd:25,is:30,vi:15,rr:10,si:5,ger:15},
 'Riesgo y resiliencia':{nd:20,is:15,vi:5,rr:40,si:15,ger:5},
 'Madurez rápida':{nd:15,is:15,vi:30,rr:5,si:5,ger:30}
};
const labels={nd:'Necesidad y desempeño',is:'Impacto socioeconómico',vi:'Viabilidad',rr:'Riesgo / resiliencia',si:'Sostenibilidad',ger:'Gerencial'};
const snitLayers=[
 // CAPAS SNIT: usar URL WMS oficial + nombre COMPLETO con namespace. En GitHub Pages deben ir en HTTPS.
 {name:'IGN límite cantonal 5k CO',url:'https://geos.snitcr.go.cr/be/IGN_5_CO/wms?',layers:'IGN_5_CO:limitecantonal_5k',opacity:.45, source:'SNIT-IGN'},
 {name:'IGN límite distrital 5k',url:'https://geos.snitcr.go.cr/be/IGN_5/wms?',layers:'IGN_5:limitedistrital_5k',opacity:.35, source:'SNIT-IGN'},
 {name:'IGN límite nacional 5k',url:'https://geos.snitcr.go.cr/be/IGN_5/wms?',layers:'IGN_5:limitenacional_5k',opacity:.55, source:'SNIT-IGN'},
 {name:'SINAC corredores biológicos',url:'https://geos.snitcr.go.cr/be/SINAC/wms?',layers:'SINAC:corredoresbiologicos',opacity:.55, source:'SNIT-SINAC'},
 {name:'SINAC áreas silvestres protegidas',url:'https://geos.snitcr.go.cr/be/SINAC/wms?',layers:'SINAC:areas_silvestres_protegidas',opacity:.55, source:'SNIT-SINAC'},
 {name:'SINAC humedales',url:'https://geos.snitcr.go.cr/be/SINAC/wms?',layers:'SINAC:humedales',opacity:.55, source:'SNIT-SINAC'},
 {name:'CONAVI zonas conservación vial',url:'https://geos.snitcr.go.cr/be/CONAVI/wms?',layers:'CapasCONAVI:Zonas_de_Conservación_Vial',opacity:.65, source:'SNIT-CONAVI'},
 {name:'SENARA recarga acuífera',url:'https://geos.snitcr.go.cr/be/SENARA/wms?',layers:'SENARA:recarga_acuifera',opacity:.55, source:'SNIT-SENARA'},
 {name:'CENIGA ambiental histórico',url:'https://geodatos.sinia.go.cr/geoserver/CENIGA/wms?',layers:'CENIGA:corredores_biologicos',opacity:.55, source:'SINIA-CENIGA'}
];
let weights={...scenarios['Base metodología']};
let map=null;
const mapStub={setView(){return this},flyTo(){return this},fitBounds(){return this},removeLayer(){},addLayer(){},eachLayer(){},closePopup(){}};
try{
  if(window.L){
    map=L.map('map',{preferCanvas:true}).setView([9.93,-84.08],8);
  }else{
    map=mapStub;
    document.getElementById('map').innerHTML='<div style="height:100%;display:grid;place-items:center;padding:30px;text-align:center;color:#65758b;background:#eef3f8"><div><b style="color:#0b2f5b">El mapa base no pudo conectarse.</b><br>La matriz, los filtros y el listado siguen disponibles. Verifique la conexión a los CDN de Leaflet.</div></div>';
  }
}catch(e){map=mapStub;console.error('Mapa:',e)}
L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png',{attribution:'© OpenStreetMap | SNIT OGC'}).addTo(map);
let markerLayer=L.geoJSON(null,{pointToLayer:(f,latlng)=>L.circleMarker(latlng,{radius:7,fillColor:color(score(f.properties)),color:'#fff',weight:1,fillOpacity:.9}),onEachFeature:(f,l)=>{l.bindPopup(popup(f.properties),{className:'popup'}); l.on('click',()=>selectProject(f.properties.id));}}).addTo(map);
let selectedProjectId=null;
let wmsGroup={};
function color(s){return s>=85?'#c62828':s>=70?'#ef6c00':s>=55?'#f9a825':s>=40?'#78909c':'#b0bec5'}
function cls(s){return s>=85?'Muy alta':s>=70?'Alta':s>=55?'Media':s>=40?'Baja':'Muy baja'}
function score(p){let sum=0,tot=0; for(const k of Object.keys(weights)){sum += (+p[k]||0)*(+weights[k]||0); tot += (+weights[k]||0)} return tot? +(sum/tot).toFixed(1):0}
function clean(v){return (v===undefined||v===null||String(v)==='nan'||Number.isNaN(v))?'Sin dato':v}
function factorLevel(v){v=+v||0; return v>=85?'Muy alto':v>=70?'Alto':v>=55?'Medio':v>=40?'Bajo':'Muy bajo'}
const factorSubvars={
 nd:['TPDA / demanda','IRI / condición del activo','Siniestralidad'],
 is:['Población beneficiaria','IDS / equidad territorial','Acceso a activos económicos'],
 vi:['Costo referencial','Complejidad constructiva','Disponibilidad predial'],
 rr:['Amenazas / BSA','Criticidad de conectividad','Conflictividad social'],
 si:['Sensibilidad ambiental','Descarbonización','Integración con red existente'],
 ger:['Prioridad institucional','Financiamiento','Urgencia / madurez']
};
function factorCards(p){
 let totalWeight=Object.values(weights).reduce((a,b)=>a+(+b||0),0)||1;
 return Object.keys(labels).map(k=>{
   let val=+(p[k]||0), w=+(weights[k]||0), wn=w/totalWeight*100, contrib=val*w/totalWeight;
   return `<div class='factor-card'><div class='factor-head'><span class='factor-name'>${labels[k]}</span><span class='factor-score'>${val.toFixed(1)} / 100</span></div><div class='bar'><div style='width:${Math.max(0,Math.min(100,val))}%'></div></div><div class='contribution'>Peso escenario: <b>${wn.toFixed(1)}%</b> · Aporte al puntaje final: <b>${contrib.toFixed(1)}</b> · Nivel: <b>${factorLevel(val)}</b></div><div class='subvars'><b>Subvariables metodológicas:</b> ${factorSubvars[k].join(' · ')}</div></div>`;
 }).join('');
}
function completeness(p){
 const keys=['nd','is','vi','rr','si','ger'];
 let ok=keys.filter(k=>Number.isFinite(+p[k]) && +p[k]>0).length;
 return Math.round(ok/keys.length*100);
}
function topFactors(p){
 return Object.keys(labels).map(k=>({k,val:+p[k]||0,label:labels[k]})).sort((a,b)=>b.val-a.val);
}
function executiveText(p){
 const s=score(p), c=cls(s), top=topFactors(p), best=top[0], second=top[1], weak=top[top.length-1];
 let base=`Esta iniciativa presenta una prioridad <b>${c}</b> porque combina un desempeño destacado en <b>${best.label}</b> (${best.val.toFixed(1)}/100) y <b>${second.label}</b> (${second.val.toFixed(1)}/100), bajo el escenario de ponderación activo.`;
 if((+p.si||0)<45) base += ` No obstante, la sostenibilidad territorial aparece como alerta, por lo que conviene validar restricciones ambientales antes de avanzar.`;
 if((+p.vi||0)<50) base += ` La viabilidad de implementación requiere gestión adicional, especialmente en madurez, costos, predial o complejidad constructiva.`;
 if((+p.is||0)>=75) base += ` Su impacto socioeconómico sugiere potencial para beneficiar población o mejorar accesibilidad territorial.`;
 if((+p.rr||0)>=75) base += ` Además, aporta a la resiliencia o criticidad funcional de la red.`;
 if(!p.elegible) base += ` Antes de llevarlo a decisión, debe completarse o verificarse el filtro de elegibilidad.`;
 return base;
}
function recommendationText(p){
 const s=score(p), comp=completeness(p);
 if(!p.elegible) return `Recomendación del HUB: mantener la iniciativa en revisión técnica hasta completar elegibilidad, localización, alcance y soportes mínimos.`;
 if(s>=85 && comp>=70) return `Recomendación del HUB: mantener en el grupo de prioridad muy alta y avanzar hacia formulación, prefactibilidad o estructuración según su etapa actual.`;
 if(s>=70) return `Recomendación del HUB: priorizar para revisión del Comité HUB Central, validando primero los factores débiles y los datos que aún requieran confirmación institucional.`;
 if(s>=55) return `Recomendación del HUB: conservar en cartera y mejorar información técnica antes de competir por recursos frente a proyectos de mayor impacto o madurez.`;
 return `Recomendación del HUB: no priorizar en el escenario actual, salvo que exista una instrucción estratégica o una restricción crítica no capturada por los datos disponibles.`;
}
function decisionAlerts(p){
 const arr=[];
 arr.push(p.elegible?`✔ Pasa filtro preliminar de elegibilidad`:`⚠ Elegibilidad incompleta o por validar`);
 arr.push((+p.si||0)<45?`⚠ Validar sensibilidad ambiental y restricciones territoriales`:`✔ Sin alerta ambiental crítica en el puntaje disponible`);
 arr.push((+p.vi||0)<50?`⚠ Requiere fortalecer viabilidad, costos, predial o madurez`:`✔ Viabilidad técnica aceptable para continuar análisis`);
 arr.push((+p.rr||0)>=75?`✔ Alta contribución a resiliencia / criticidad funcional`:`• Revisar criticidad y exposición a amenazas`);
 return `<ul class='alert-list'>${arr.map(x=>`<li>${x}</li>`).join('')}</ul>`;
}

function popup(p){
 let s=score(p);
 let escenario=document.getElementById('scenario')?.value||'Base metodología';
 let conf=completeness(p);
 return `<div class='popup-executive'>
 <h3 class='detail-title'>${clean(p.nombre||p.proyecto)}</h3>
 <span class='badge' style='background:${color(s)}'>${cls(s)} · ${s}/100</span>
 <span class='badge' style='background:#0b2f5b'>${escenario}</span>
 <div class='confidence-box'><div><small>Confianza del análisis</small><div class='bar'><div style='width:${conf}%'></div></div></div><div class='conf-num'>${conf}%</div></div>
 <div class='exec-read'><b>¿Por qué este proyecto es prioritario?</b><br>${executiveText(p)}</div>
 <div class='recommend-box'><b>Recomendación para decisión:</b><br>${recommendationText(p)}</div>
 <table class='meta-table'>
  <tr><td>ID / BPIP</td><td>${clean(p.id)} / ${clean(p.bpip)}</td></tr>
  <tr><td>Institución</td><td>${clean(p.institucion)}</td></tr>
  <tr><td>Etapa / Estado</td><td>${clean(p.etapa)} / ${clean(p.estado)}</td></tr>
  <tr><td>Elegibilidad</td><td>${p.elegible?'Pasa filtro de elegibilidad':'Revisar elegibilidad / información incompleta'}</td></tr>
  <tr><td>Ubicación</td><td>${clean(p.ubicacion)}</td></tr>
  <tr><td>Coordenadas</td><td>${clean(p.y_crtm05)}, ${clean(p.x_crtm05)} CRTM05</td></tr>
 </table>
 <h4 style='margin:10px 0 4px;color:#0b2f5b'>Factores de decisión</h4>
 <div class='scoregrid'>${Object.keys(labels).map(k=>`<div><b>${(+p[k]||0).toFixed(1)}</b>${labels[k]}</div>`).join('')}</div>
 <h4 style='margin:10px 0 4px;color:#0b2f5b'>Alertas y hallazgos para comité</h4>
 ${decisionAlerts(p)}
 ${factorCards(p)}
 <p class='note'><b>Nota metodológica:</b> el puntaje combina indicadores normalizados 0–100 y los pesos del escenario activo. La confianza depende de la completitud de los factores disponibles; las variables no calculables automáticamente deben ser validadas por la dependencia responsable.</p>
 </div>`;
}
function selectProject(id){selectedProjectId=id; let f=projects.features.find(x=>String(x.properties.id)===String(id)); if(f){updateQuadSelected(f.properties); renderQuadrant(filtered()); showProjectPopup(f);}}
function showProjectPopup(f){
 let p=f.properties;
 let c=f.geometry.coordinates;
 map.setView([c[1],c[0]], Math.max(map.getZoom(), 11));
 L.popup({className:'popup',maxWidth:680,autoPan:true,keepInView:true})
  .setLatLng([c[1],c[0]])
  .setContent(popup(p))
  .openOn(map);
}
function showProjectDetail(p){
 let f=projects.features.find(x=>String(x.properties.id)===String(p.id));
 if(f) showProjectPopup(f);
}

function initControls(){
 let sc=document.getElementById('scenario'); Object.keys(scenarios).forEach(n=>sc.add(new Option(n,n))); sc.onchange=()=>{weights={...scenarios[sc.value]}; renderWeights(); update();};
 ['fEtapa','fInst'].forEach(id=>{let key=id==='fEtapa'?'etapa':'institucion'; let vals=[...new Set(projects.features.map(f=>f.properties[key]).filter(Boolean))].sort(); vals.forEach(v=>document.getElementById(id).add(new Option(v,v)));});
 const qEl=document.getElementById('q');
['q','fEtapa','fInst','fClass','fQuad'].forEach(id=>document.getElementById(id).addEventListener('input',update));
 document.getElementById('normalize').onclick=()=>{normalize();renderWeights();update();}; document.getElementById('base').onclick=()=>{document.getElementById('scenario').value='Base metodología'; weights={...scenarios['Base metodología']}; renderWeights(); update();};
 document.getElementById('exportCsv').onclick=exportCsv;
 renderWeights(); renderLayers();
}
function renderWeights(){let box=document.getElementById('weights'); box.innerHTML=''; Object.entries(labels).forEach(([k,lab])=>{let div=document.createElement('div'); div.className='row'; div.innerHTML=`<label>${lab}<input type='range' min='0' max='60' value='${weights[k]||0}' id='w_${k}'></label><input type='text' value='${weights[k]||0}' id='t_${k}'>`; box.appendChild(div); div.querySelector(`#w_${k}`).oninput=e=>{weights[k]=+e.target.value;div.querySelector(`#t_${k}`).value=weights[k];update();}; div.querySelector(`#t_${k}`).oninput=e=>{weights[k]=+e.target.value||0;div.querySelector(`#w_${k}`).value=weights[k];update();};})}
function renderLayers(){
 let box=document.getElementById('layers'); box.innerHTML='';
 snitLayers.forEach((cfg,i)=>{
   let d=document.createElement('div'); d.className='layerbox';
   d.innerHTML=`<label><input type='checkbox' data-i='${i}'> ${cfg.name}</label><div class='note'>${cfg.url}<br><b>Layer:</b> ${cfg.layers}</div>`;
   box.appendChild(d); d.querySelector('input').onchange=e=>toggleWms(i,e.target.checked);
 });
 document.getElementById('addManualLayer').onclick=()=>{
   let url=document.getElementById('manualUrl').value.trim();
   let layer=document.getElementById('manualLayer').value.trim();
   if(!url||!layer){alert('Pega una URL WMS y el nombre exacto de la capa.');return;}
   snitLayers.push({name:'Manual: '+layer,url:url,layers:layer,opacity:.65,source:'Manual'});
   renderLayers();
 };
}
function statusMsg(msg){let el=document.getElementById('layerStatus'); if(el) el.innerHTML=msg;}
// SNIT suele fallar con WMS por teselas/CRS en Leaflet. Este modo usa imagen única GetMap en EPSG:4326.
function normalizeWmsUrl(u){
  u=(u||'').trim();
  if(!u) return u;
  return u.includes('?') ? u : u + '?';
}
function buildWmsImageUrl(c){
  const b=map.getBounds();
  const size=map.getSize();
  const base=normalizeWmsUrl(c.url);
  const params={
    SERVICE:'WMS', VERSION:'1.1.1', REQUEST:'GetMap',
    LAYERS:c.layers, STYLES:'', FORMAT:'image/png', TRANSPARENT:'TRUE',
    SRS:'EPSG:4326',
    BBOX:[b.getWest(), b.getSouth(), b.getEast(), b.getNorth()].join(','),
    WIDTH:Math.max(256, Math.round(size.x)), HEIGHT:Math.max(256, Math.round(size.y))
  };
  const qs=Object.entries(params).map(([k,v])=>k+'='+encodeURIComponent(v)).join('&');
  return base + (base.endsWith('?')?'':'&') + qs;
}
function refreshSingleWms(i){
  const obj=wmsGroup[i]; if(!obj) return;
  const c=obj.cfg;
  const url=buildWmsImageUrl(c);
  const bounds=map.getBounds();
  const img=new Image();
  img.onload=()=>{
    const next=L.imageOverlay(url,bounds,{opacity:c.opacity||.6,interactive:false,attribution:'SNIT Costa Rica'});
    next.addTo(map);
    if(obj.layer) map.removeLayer(obj.layer);
    obj.layer=next;
    obj.testUrl=url;
    statusMsg('✅ Capa cargada como imagen WMS: <b>'+c.name+'</b>');
  };
  img.onerror=()=>{
    obj.testUrl=url;
    statusMsg('⚠️ SNIT no devolvió imagen para <b>'+c.name+'</b>. En la consola del navegador queda el enlace WMS de prueba.');
    console.warn('WMS falló:', url);
  };
  img.src=url;
}
function toggleWms(i,on){
 let c=snitLayers[i];
 if(on){
   wmsGroup[i]={cfg:c,layer:null,testUrl:null};
   refreshSingleWms(i);
 } else if(wmsGroup[i]){
   if(wmsGroup[i].layer) map.removeLayer(wmsGroup[i].layer);
   delete wmsGroup[i]; statusMsg('Capa desactivada.');
 }
}
map.on('moveend zoomend resize',()=>{Object.keys(wmsGroup).forEach(i=>refreshSingleWms(i));});


function stageScore(p){
  let e=String(p.etapa||'').toLowerCase();
  if(e.includes('licit')) return 92;
  if(e.includes('factibilidad')) return 82;
  if(e.includes('prefact')) return 70;
  if(e.includes('perfil')) return 58;
  if(e.includes('idea')) return 35;
  return 50;
}
function completenessScore(p){
  const fields=[p.id,p.nombre||p.proyecto,p.institucion,p.etapa,p.estado,p.ubicacion,p.bpip,p.inversion];
  const valid=fields.filter(v=>v!==null&&v!==undefined&&String(v).trim()!==''&&String(v).toLowerCase()!=='nan'&&String(v).toLowerCase()!=='por definir').length;
  return +(valid/fields.length*100).toFixed(1);
}
function stageReadiness(p){const e=String(p.etapa||'').toLowerCase();if(e.includes('ejec'))return 95;if(e.includes('licit'))return 88;if(e.includes('factib'))return 82;if(e.includes('prefact'))return 70;if(e.includes('perfil'))return 55;if(e.includes('idea'))return 25;return 40}function complexityComponents(p){return{etapa:+(100-stageReadiness(p)).toFixed(1),tecnica:+(100-(+p.vi||0)).toFixed(1),institucional:+(100-(+p.ger||0)).toFixed(1),informacion:+(100-completeness(p)).toFixed(1),restricciones:+(100-(+p.si||0)).toFixed(1)}}function effortScore(p){const c=complexityComponents(p);return +(c.etapa*.25+c.tecnica*.30+c.institucional*.20+c.informacion*.15+c.restricciones*.10).toFixed(1)}function priorityScore(p){return score(p)}function quadrantFromXY(x,y){if(y>=70&&x<50)return'Oportunidades Estratégicas';if(y>=70&&x>=50)return'Proyectos Transformadores';if(y<70&&x<50)return'Intervenciones Complementarias';return'Proyectos de Largo Plazo'}function quadrantName(p){return p._quad||quadrantFromXY(p._effortRaw??effortScore(p),p._priorityRaw??priorityScore(p))}function quadrantAction(q){return q==='Oportunidades Estratégicas'?'Alto valor estratégico y complejidad comparativamente baja.':q==='Proyectos Transformadores'?'Alto valor con exigencias técnicas, financieras o institucionales relevantes.':q==='Intervenciones Complementarias'?'Aporte focalizado y complejidad relativamente manejable.':'Requiere mayor maduración o revisión antes de competir por recursos.'}
function renderQuadrant(feats){
 const chart=document.getElementById('quadChart'); if(!chart) return;
 chart.querySelectorAll('.qpoint').forEach(n=>n.remove());
 const counts={strategic:0,catalyst:0,opportunity:0,develop:0};
 (feats||[]).forEach((f,index)=>{
   const p=f.properties;
   const rawX=Number.isFinite(+p._effortRaw)?+p._effortRaw:effortScore(p);
   const rawY=Number.isFinite(+p._priorityRaw)?+p._priorityRaw:priorityScore(p);
   const jx=hashJitter(String(p.id||index))*1.25;
   const jy=hashJitter(String(p.id||index)+'y')*1.25;
   const x=Math.max(2,Math.min(98,rawX+jx));
   const y=Math.max(2,Math.min(98,rawY+jy));
   const q=quadrantFromXY(rawX,rawY);
   p._effortRaw=rawX; p._priorityRaw=rawY; p._quad=q;
   if(q==='Oportunidades Estratégicas') counts.strategic++;
   else if(q==='Proyectos Transformadores') counts.catalyst++;
   else if(q==='Intervenciones Complementarias') counts.opportunity++;
   else counts.develop++;
   const d=document.createElement('button');
   d.type='button';
   d.className='qpoint'+(String(p.id)===String(selectedProjectId)?' active':'');
   d.style.left=x+'%';
   d.style.top=(100-y)+'%';
   d.style.background=color(rawY);
   d.setAttribute('aria-label',(p.nombre||p.proyecto||'Proyecto')+' '+q);
   d.title=(p.nombre||p.proyecto)+' · '+q+' · Puntaje '+rawY+'/100 · Complejidad '+rawX+'/100';
   d.onclick=(ev)=>{ev.stopPropagation();selectProject(p.id);};
   chart.appendChild(d);
 });
 const ids={qStrategic:counts.strategic,qCatalyst:counts.catalyst,qOpportunity:counts.opportunity,qDevelop:counts.develop};
 Object.entries(ids).forEach(([id,v])=>{const el=document.getElementById(id);if(el)el.textContent=String(v);});
 const hint=chart.querySelector('.quad-hint');
 if(hint) hint.textContent=`Cortes: prioridad 70 · complejidad 50 · ${feats.length} proyectos visibles`;
}
function updateQuadSelected(p){
 let el=document.getElementById('quadSelected'); if(!el) return;
 if(!p){el.innerHTML='<b>Sin proyecto seleccionado.</b><br>Haz clic en un punto de la matriz, del mapa o del ranking para ver su lectura estratégica.'; return;}
 let q=quadrantName(p), x=p._effortRaw??effortScore(p), y=p._priorityRaw??priorityScore(p), c=complexityComponents(p);
 el.innerHTML=`<b>${p.nombre||p.proyecto}</b><br><b>${q}</b><br>Puntaje de priorización: ${y}/100 · Complejidad de implementación: ${x}/100.<br><span style="color:#65758b">Preparación: etapa ${c.etapa}, complejidad técnica ${c.tecnica}, esfuerzo institucional ${c.institucional}, brechas de información ${c.informacion} y restricciones ${c.restricciones}.</span><br>${quadrantAction(q)}`;
}


function update(){
  try{
    const feats=filtered();
    if(markerLayer){markerLayer.clearLayers();markerLayer.addData(feats);}
    const total=feats.length;
    const avg=total?(feats.reduce((a,f)=>a+(+f.properties._score||0),0)/total).toFixed(1):'0';
    const high=feats.filter(f=>(+f.properties._score||0)>=70).length;
    const elig=feats.filter(f=>f.properties.elegible).length;
    const set=(id,v)=>{const el=document.getElementById(id);if(el)el.textContent=String(v);};
    set('kTotal',total);set('kAvg',avg);set('kHigh',high);set('kElig',elig);
    renderRanking(feats);
    renderQuadrant(feats);
    const selected=selectedProjectId?projects.features.find(x=>String(x.properties.id)===String(selectedProjectId)):null;
    updateQuadSelected(selected?selected.properties:null);
    const status=document.getElementById('layerStatus');
    if(status){
      status.dataset.layerMessage='1';
      const parts=[];
      if(fInst?.value)parts.push('institución '+fInst.value);
      if(fEtapa?.value)parts.push('etapa '+fEtapa.value);
      if(fClass?.value)parts.push('prioridad '+fClass.value);
      if(fQuad?.value)parts.push('grupo '+fQuad.value);
      status.textContent=total?`${total} proyectos visibles${parts.length?' · '+parts.join(' · '):''}.`:`No se encontraron proyectos${parts.length?' para '+parts.join(' · '):''}.`;
    }
    setTimeout(()=>{try{map?.invalidateSize();}catch(e){}},100);
    reportHubHeight?.();
  }catch(err){
    console.error('Error actualizando priorización:',err);
    const status=document.getElementById('layerStatus');
    if(status)status.textContent='Error al cargar la información: '+err.message;
  }
}
function normalize(){let tot=Object.values(weights).reduce((a,b)=>a+(+b||0),0)||1; Object.keys(weights).forEach(k=>weights[k]=+(weights[k]*100/tot).toFixed(1));}
function normText(v){return String(v??'').trim().toLowerCase();}
function filtered(){
  const q=normText(qEl?.value), et=normText(fEtapa?.value), inst=normText(fInst?.value), cl=normText(fClass?.value), quad=normText(fQuad?.value);
  const result=projects.features.map(f=>{
    const p=f.properties;
    p._score=score(p);
    p._class=cls(p._score);
    p._effortRaw=effortScore(p);
    p._priorityRaw=priorityScore(p);
    p._quad=quadrantFromXY(p._effortRaw,p._priorityRaw);
    return f;
  }).filter(f=>{
    const p=f.properties;
    const txt=normText(Object.values(p).join(' '));
    return (!q||txt.includes(q)) &&
           (!et||normText(p.etapa)===et) &&
           (!inst||normText(p.institucion)===inst) &&
           (!cl||normText(p._class)===cl) &&
           (!quad||normText(p._quad)===quad);
  }).sort((a,b)=>b.properties._score-a.properties._score);
  window.__lastFilteredFeatures=result;
  return result;
}
function renderRanking(feats){let box=document.getElementById('ranking'); if(!feats.length){box.innerHTML='<div class="detail-empty">No hay proyectos con los filtros seleccionados. Usa Limpiar para ver todo el portafolio.</div>'; return;} box.innerHTML=feats.slice(0,50).map((f,i)=>{let p=f.properties,s=p._score;return `<div class='item' onclick='zoomTo(${f.geometry.coordinates[1]},${f.geometry.coordinates[0]},"${String(p.id).replaceAll('\"','&quot;')}")'><span class='rank'>#${i+1}</span> <b>${p.nombre||p.proyecto}</b> <span class='badge' style='background:${color(s)}'>${p._class} ${s}</span><br><small>${p.institucion||''} · ${p.etapa||''}</small><div class='mini'><div style='width:${s}%'></div></div></div>`}).join('')}
function zoomTo(lat,lon,id){map.setView([lat,lon],12); if(id) selectProject(id)}
function exportCsv(){
  let rows=filtered().map(f=>{let p=f.properties;return [p.id,p.nombre,p.institucion,p.etapa,p.estado,p._score,p._class,p.nd,p.is,p.vi,p.rr,p.si,p.ger,f.geometry.coordinates[1],f.geometry.coordinates[0]]});
  let header='ID,Nombre,Institucion,Etapa,Estado,Puntaje,Clasificacion,ND,IS,VI,RR,SI,Gerencial,Lat,Lon';
  let csv=header+'\n'+rows.map(r=>r.map(x=>'"'+String(x??'').replaceAll('"','""')+'"').join(',')).join('\n');
  let a=document.createElement('a');
  a.href=URL.createObjectURL(new Blob([csv],{type:'text/csv;charset=utf-8'}));
  a.download='ranking_priorizacion_costa_rica.csv';
  a.click();
}
let legend=L.control({position:'bottomleft'}); legend.onAdd=()=>{let d=L.DomUtil.create('div','legend'); d.innerHTML='<b>Prioridad</b><br><span class="dot" style="background:#c62828"></span>Muy alta ≥85<br><span class="dot" style="background:#ef6c00"></span>Alta 70–84<br><span class="dot" style="background:#f9a825"></span>Media 55–69<br><span class="dot" style="background:#78909c"></span>Baja 40–54<br><span class="dot" style="background:#b0bec5"></span>Muy baja <40';return d}; legend.addTo(map);

const qEl=document.getElementById('q');
const fEtapa=document.getElementById('fEtapa');
const fInst=document.getElementById('fInst');
const fClass=document.getElementById('fClass');
const fQuad=document.getElementById('fQuad');
const rankSearch=document.getElementById('rankSearch');
initControls(); document.getElementById('rankSearch').addEventListener('input',()=>{let t=rankSearch.value.toLowerCase();document.querySelectorAll('#ranking .item').forEach(x=>x.style.display=x.textContent.toLowerCase().includes(t)?'':'none')});document.getElementById('resetFilters').addEventListener('click',()=>{['q','fEtapa','fInst','fClass','fQuad'].forEach(id=>document.getElementById(id).value='');document.getElementById('rankSearch').value='';update();map&&map.setView([9.93,-84.08],8)});document.querySelectorAll('.quad-kpi[data-quad]').forEach(el=>el.addEventListener('click',()=>{document.getElementById('fQuad').value=el.dataset.quad;update()}));
update(); map&&map.setView([9.93,-84.08],8);

function reportHubHeight(){
  requestAnimationFrame(()=>{
    const h=Math.max(document.documentElement.scrollHeight,document.body.scrollHeight,900);
    parent.postMessage({type:'hub-priorizacion-height',height:h},'*');
  });
}
window.addEventListener('load',reportHubHeight);
window.addEventListener('resize',reportHubHeight);
new MutationObserver(reportHubHeight).observe(document.body,{subtree:true,childList:true,attributes:true});
setTimeout(reportHubHeight,300);setTimeout(reportHubHeight,1500);setTimeout(reportHubHeight,4000);

