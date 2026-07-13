// Copia de referencia extraída de priorizacion.html
const snitLayers = [
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
