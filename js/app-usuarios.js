function setThemeBtn(){var d=document.documentElement.getAttribute('data-theme')==='dark';var b=document.getElementById('theme-btn');if(b)b.textContent=d?'☀️':'🌙'}
function syncDocsTheme(){
  var ifd=document.getElementById('iframe-documentos');
  if(ifd&&ifd.contentWindow){
    var dk=document.documentElement.getAttribute('data-theme')==='dark';
    try{ifd.contentWindow.postMessage({type:'theme',dark:dk},'*');}catch(e){}
  }
}
function toggleTheme(){
  var html=document.documentElement,dark=html.getAttribute('data-theme')==='dark';
  html.setAttribute('data-theme',dark?'light':'dark');
  localStorage.setItem('kg-theme',dark?'light':'dark');
  setThemeBtn();
  syncDocsTheme();
}
/* ════════════════════════════════════════════════════════════════════
   DATOS PARA LA PRESENTACIÓN (PPTX) — Generador de Documentos
   El módulo de Documentos vive en un iframe aislado y no tiene acceso a
   los datos. Aquí se construye el resumen agregado y se le envía por
   postMessage. REGLA DE SEGURIDAD: solo se entregan datos de las razones
   sociales que la cuenta tiene permitidas; si pide otra, se rechaza.
════════════════════════════════════════════════════════════════════ */
var MESES_PPTX=['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'];
function _mesIdx(m){
  var n=norm(String(m||''));
  for(var i=0;i<MESES_PPTX.length;i++){ if(norm(MESES_PPTX[i])===n) return i; }
  return -1;
}
function _razonDeRegistro(rec){
  return rec.razon || razonDeCentro(centroDeTienda(rec.tienda||'')) || '';
}
function _pctP(ok,tot){ return tot>0?Math.round(ok/tot*100):0; }

/* Razones que la cuenta puede exportar (null = todas) */
function razonesExportables(){
  var r=_razonesAsignadas();
  if(r&&r.length)return r.slice();
  return ['KNO','KSC','KSA'];
}

/* ════════════════════════════════════════════════════════════════════
   CONTEO DE TAREAS DESDE LA FUENTE REAL + DEDUPLICACIÓN
   Los campos pendientes/resueltas de `auditorias` son una foto del momento
   de la carga y quedan desfasados cuando alguien resuelve una tarea. Aquí
   se cuentan las tareas reales, ligadas a su auditoría por
   centro (o tienda) + clase + mes.
════════════════════════════════════════════════════════════════════ */
function _claveTienda(rec){
  var c=String(rec.centro||'').trim();
  if(c && !pareceCifrado(c)) return 'C:'+norm(c);
  return 'T:'+norm(rec.tienda||'');
}
/* Clase: 'ol' (Orden y Limpieza) o 'proc' (Procesos/Colaboración/Cartera) */
function _claseTarea(t){ return tipoNorm(t.tipoTarea||t.tipo_tarea)==='ol' ? 'ol' : 'proc'; }
function _claseAuditoria(a){ return /orden/i.test(String(a.clase||'')) ? 'ol' : 'proc'; }
function _mesDeTarea(t){
  var f=t.fechaCreacion||t.fecha_creacion;
  if(!f)return -1;
  var d=new Date(f);
  return isNaN(d) ? -1 : d.getMonth();
}

/* Elimina duplicados exactos conservando el primero */
function _dedup(arr, claveFn){
  var vistos={}, out=[], dups=0;
  (arr||[]).forEach(function(x){
    var k=claveFn(x);
    if(vistos[k]){dups++;return;}
    vistos[k]=1; out.push(x);
  });
  if(dups)console.warn('Registros duplicados omitidos:',dups);
  out._dups=dups;
  return out;
}

/* Índice: clave|clase|mes -> {tareas, resueltas, pendientes} */
function _indiceTareas(razon){
  var idx={};
  var tareas=_dedup(STORE.tareas||[], function(t){
    return t.id!=null ? 'id:'+t.id
      : [norm(t.tienda), norm(t.nombre||t.actividad||''), t.fechaCreacion||'', norm(t.tipoTarea||'')].join('|');
  });
  tareas.forEach(function(t){
    if(razKey(_razonDeRegistro(t))!==razKey(razon))return;
    var k=_claveTienda(t)+'|'+_claseTarea(t)+'|'+_mesDeTarea(t);
    if(!idx[k])idx[k]={tareas:0,resueltas:0,pendientes:0,pendVenc:0};
    idx[k].tareas++;
    if(esResuelta(t))idx[k].resueltas++;
    else { idx[k].pendientes++; if(tareaVencidaPorFecha(t))idx[k].pendVenc++; }
  });
  return idx;
}

function buildPptxDataset(mes,razon,mesesRango){
  var mIdx=_mesIdx(mes);
  var rzOk=function(rec){ return razKey(_razonDeRegistro(rec))===razKey(razon); };
  var mesOk=function(v){ return _mesIdx(v)===mIdx; };
  var esOrden=function(c){ return /orden/i.test(String(c||'')); };

  /* ── AUDITORÍAS: detalle por tienda y resumen por clase ── */
  /* Universo completo de auditorías = vigentes + FINALIZADAS.
     Las finalizadas salen del módulo de auditorías (para no duplicar en pantalla),
     pero el reporte SÍ debe incluirlas: si no, se perdería su información.
     El dedup por tienda+mes+clase garantiza que ninguna se cuente dos veces. */
  var finsComoAud=(FINALIZADAS||[])
    .filter(function(f){return f && !f._pending;})
    .map(function(f){
      var tot=parseInt(f.total_tareas)||0, res=parseInt(f.resueltas)||0;
      return {
        tienda:f.tienda||'', mes:f.mes||'', clase:f.clase||'',
        /* Las finalizadas no guardan `centro`, pero _claveTienda usa centro
           cuando existe. Sin esto, la copia finalizada quedaba con clave
           'T:tienda' y la auditoría viva con 'C:centro' → el dedup no las
           colapsaba y la tienda aparecía DUPLICADA en el PPTX. Derivamos el
           centro desde la tienda para que ambas claves coincidan. */
        centro:f.centro||centroDeTienda(f.tienda||'')||'', razon:f.razon||'',
        pctCumpl:f.pct_cumpl, tareas:tot, resueltas:res,
        pendientes:Math.max(0, tot-res), _finalizada:true
      };
    });
  var auds=_dedup((STORE.auditorias||[]).concat(finsComoAud).filter(function(a){return rzOk(a);}),
    function(a){ return [_claveTienda(a), norm(a.mes), _claseAuditoria(a)].join('|'); });
  var IDX=_indiceTareas(razon);
  function detalleDe(lista){
    return lista.map(function(a){
      /* En memoria los campos son camelCase (pctCumpl) y vienen como fracción
         (0.44) o como porcentaje (44): se normaliza a 0-100. */
      function aPct(v){
        var n=parseFloat(v)||0;
        return Math.round(n<=1 ? n*100 : n);
      }
      /* Pendientes/resueltas CONTADAS de las tareas reales (siempre al día).
         El snapshot de `auditorias` solo se usa si no hay tareas ligadas. */
      var k=_claveTienda(a)+'|'+_claseAuditoria(a)+'|'+_mesIdx(a.mes);
      var real=IDX[k];
      /* Bug corregido: la etiqueta `a.mes` de la auditoría (ej. "MAYO") a veces
         no coincide con el mes real en que se crearon sus tareas ligadas
         (ej. tareas creadas el 2 de junio para revisar mayo) — el índice de
         tareas se arma por mes de creación (_mesDeTarea), así que la clave
         basada en `a.mes` no encontraba el match y el semáforo caía al
         snapshot, que SIEMPRE reportaba 0 tareas atrasadas (pv=0), pintando
         "vigente" auditorías que en realidad tenían tareas atrasadas.
         Como respaldo, se reintenta la misma clave usando el mes real de
         `a.fecha` antes de rendirse al snapshot. */
      if(!real && a.fecha){
        var _fd=new Date(a.fecha);
        if(!isNaN(_fd)){
          var k2=_claveTienda(a)+'|'+_claseAuditoria(a)+'|'+_fd.getMonth();
          if(k2!==k) real=IDX[k2];
        }
      }
      var r, p, pv;
      if(real){ r=real.resueltas; p=real.pendientes; pv=real.pendVenc||0; }
      else {
        r=+a.resueltas||0; p=+a.pendientes||0;
        /* Sin tareas reales ligadas no se puede saber cuántas están
           atrasadas: se asume el caso conservador (todas las pendientes se
           tratan como atrasadas) en vez de 0, para que el semáforo nunca
           muestre "vigente" sobre datos que no pudo verificar. */
        pv=p;
      }

      return {
        tienda:String(a.tienda||'—'), mes:String(a.mes||''),
        pctAud:aPct(a.pctCumpl!==undefined?a.pctCumpl:a.pct_cumpl),
        pendientes:p, resueltas:r, pendVenc:pv,
        pctResuelto: (r+p)>0 ? Math.round(r/(r+p)*100) : 0,
        fuente: real ? 'tareas' : 'snapshot'
      };
    }).sort(function(x,y){return y.pctAud-x.pctAud;});
  }
  function resumenDe(det){
    var p=0,r=0;
    det.forEach(function(d){p+=d.pendientes;r+=d.resueltas;});
    return {tareas:p+r, pendientes:p, resueltas:r, pctResuelto:_pctP(r,p+r)};
  }
  /* Red de seguridad robusta e independiente del mes.
     Un duplicado REAL es la misma auditoría contada dos veces (la viva + su
     copia finalizada): misma tienda, mismo mes y VALORES IDÉNTICOS. En cambio,
     dos centros físicos distintos que comparten nombre de tienda (p. ej.
     "ENSENADA MAYOREO Y MENUDEO" = KN03 y KN04) tienen valores distintos y
     deben conservarse ambos. Por eso se deduplica por la FIRMA DE VALORES, no
     por centro (que las finalizadas no guardan y centroDeTienda puede resolver
     mal en nombres ambiguos). */
  function dedupDetalle(det){
    var visto={}, out=[];
    det.forEach(function(d){
      var k=[norm(d.tienda||''), _mesIdx(d.mes), d.pctAud, d.pendientes, d.resueltas].join('|');
      if(visto[k])return;
      visto[k]=1; out.push(d);
    });
    return out;
  }
  var audMes=auds.filter(function(a){return mesOk(a.mes);});
  var procesosDetalle=dedupDetalle(detalleDe(audMes.filter(function(a){return !esOrden(a.clase);})));
  var ordenDetalle   =dedupDetalle(detalleDe(audMes.filter(function(a){return  esOrden(a.clase);})));

  /* ── HISTÓRICO por clase (tabla por tienda/mes + % por mes) ── */
  function historico(filtro){
    var filas=[], meses=[], pctMes=[];
    (mesesRango||[]).forEach(function(mn){
      var i=_mesIdx(mn);
      var lote=auds.filter(function(a){return _mesIdx(a.mes)===i && filtro(a.clase);});
      if(!lote.length)return;
      var det=detalleDe(lote);
      dedupDetalle(det).forEach(function(d){ filas.push(d); });
      var res=resumenDe(det);
      meses.push(mn); pctMes.push(res.pctResuelto);
    });
    return {filas:filas, meses:meses, pctMes:pctMes};
  }

  /* ── AJUSTES / MERMAS: resumen y detalle por tienda ── */
  function agregaCond(arr){
    var tot=0,ok=0,tar=0,porTienda={};
    (arr||[]).forEach(function(x){
      if(!rzOk(x)||!mesOk(x.mes))return;
      var t=String(x.tienda||'—');
      if(!porTienda[t])porTienda[t]={tienda:t,cantidad:0,aTiempo:0,destiempo:0};
      porTienda[t].cantidad++;
      tot++;
      if(x.condicion==='A tiempo'){ok++;porTienda[t].aTiempo++;}
      else if(x.condicion==='Destiempo'){tar++;porTienda[t].destiempo++;}
    });
    var det=Object.values(porTienda).sort(function(a,b){return b.cantidad-a.cantidad;});
    return {res:{total:tot, aTiempo:ok, destiempo:tar, pctATiempo:_pctP(ok,ok+tar)}, det:det};
  }
  var AJ=agregaCond(AJUSTES), ME=agregaCond(MERMAS);

  /* ── ACTIVIDADES: programadas / no programadas, con reprogramadas visibles ── */
  var actsMes=(ACTIVIDADES||[]).filter(function(a){return rzOk(a)&&mesOk(a.mes);});
  var prog=[], noProg=[], realizadas=0, pendientes=0;
  actsMes.forEach(function(a){
    var e=String(a.estado||'');
    if(/completad/i.test(e))realizadas++; else pendientes++;   /* reprogramado = pendiente */
    /* En memoria las fechas son camelCase (estInicio/realInicio); se deja el
       fallback snake_case por si el registro viene directo de Supabase. */
    var ini=a.realInicio||a.estInicio||a.real_inicio||a.est_inicio||'';
    var fin=a.realFin||a.estFin||a.real_fin||a.est_fin||'';
    var fila={
      actividad:String(a.actividad||''), estado:e, asignado:String(a.asignado||''),
      inicio:ini, fin:fin, vencida:actEstaVencida(a)
    };
    /* 'programada' puede venir como boolean o como texto "true"/"false" */
    var esProg=(a.programada===true)||(String(a.programada).toLowerCase()==='true')||(a.programada===undefined);
    if(esProg)prog.push(fila); else noProg.push(fila);
  });

  var anio=(new Date()).getFullYear();
  var fa=(STORE.auditorias||[]).find(function(a){return a.fecha;});
  if(fa&&fa.fecha){ var y=new Date(fa.fecha).getFullYear(); if(!isNaN(y))anio=y; }

  return {
    mes:mes, anio:anio, razon:razon,
    razonNombre:({KNO:'Kuroda Norte, S.A. de C.V.',KSC:'KS Comercial, S.A. de C.V.',KSA:'Kuroda SA, S.A. de C.V.'})[String(razon).toUpperCase()]||razon,
    actividades:{ total:actsMes.length, realizadas:realizadas, pendientes:pendientes,
                  pctCompl:_pctP(realizadas, actsMes.length),
                  programadas:prog, noProgramadas:noProg },
    procesos:resumenDe(procesosDetalle), procesosDetalle:procesosDetalle,
    ordenLimpieza:resumenDe(ordenDetalle), ordenDetalle:ordenDetalle,
    ajustes:AJ.res, ajustesDetalle:AJ.det,
    mermas:ME.res, mermasDetalle:ME.det,
    histProcesos:historico(function(c){return !esOrden(c);}),
    histOrden:historico(function(c){return  esOrden(c);})
  };
}

/* Carga bajo demanda de los módulos que alimentan la presentación */
async function asegurarDatosPptx(){
  var tareas=[];
  if(typeof AJUSTES==='undefined'||!AJUSTES.length){ if(typeof loadAjustes==='function')tareas.push(loadAjustes()); }
  if(typeof MERMAS==='undefined'||!MERMAS.length){ if(typeof loadMermas==='function')tareas.push(loadMermas()); }
  if(typeof ACTIVIDADES==='undefined'||!ACTIVIDADES.length){ if(typeof loadActividades==='function')tareas.push(loadActividades()); }
  if(!STORE.auditorias||!STORE.auditorias.length){ if(typeof loadDataFromSupabase==='function')tareas.push(loadDataFromSupabase()); }
  /* Las auditorías finalizadas también alimentan el reporte */
  if(typeof FINALIZADAS==='undefined'||!FINALIZADAS.length){ if(typeof loadFinalizadas==='function')tareas.push(loadFinalizadas()); }
  if(tareas.length){ try{ await Promise.all(tareas); }catch(e){ console.warn('asegurarDatosPptx:',e.message); } }
  /* Defensa extra: si el usuario abre el generador de PPT sin haber pasado
     antes por el dashboard/Auditorías en esta sesión, el texto de estado de
     STORE.tareas puede seguir diciendo "Abierta" con la fecha ya vencida.
     Esto se corrige aquí también, para que ningún reporte (ni ninguna otra
     pantalla que lea el texto de estado) muestre algo vencido como vigente. */
  if(typeof actualizarEstadosVencidos==='function')actualizarEstadosVencidos();
}

/* ════════════════════════════════════════════════════════════════════
   DATASET DE CUMPLIMIENTO POR TIENDA
   Consolida, por sucursal, todo lo que la afecta en el período: auditorías
   (calificación y tareas), ajustes y mermas. Respeta la razón social
   permitida igual que el resto de exportaciones.
════════════════════════════════════════════════════════════════════ */
function buildCumplimientoDataset(mes,razon,mesesRango){
  var mIdx=_mesIdx(mes);
  var rzOk=function(rec){ return razKey(_razonDeRegistro(rec))===razKey(razon); };
  var mesOk=function(v){ return _mesIdx(v)===mIdx; };
  var esOrden=function(c){ return /orden/i.test(String(c||'')); };
  var T={}; /* tienda -> acumulado */
  function get(t){
    var k=norm(t||'—');
    if(!T[k])T[k]={tienda:String(t||'—'), auditorias:0, sumPct:0,
      tareas:0, pendientes:0, resueltas:0,
      ajustes:0, ajTiempo:0, ajDestiempo:0,
      mermas:0, mmTiempo:0, mmDestiempo:0,
      clases:{}};
    return T[k];
  }
  /* Auditorías del mes — vigentes + finalizadas, sin duplicados, con tareas reales */
  var IDX=_indiceTareas(razon);
  var _finsAud=(FINALIZADAS||[]).filter(function(f){return f && !f._pending;}).map(function(f){
    var tot=parseInt(f.total_tareas)||0, res=parseInt(f.resueltas)||0;
    return {tienda:f.tienda||'', mes:f.mes||'', clase:f.clase||'', centro:f.centro||'',
            razon:f.razon||'', pctCumpl:f.pct_cumpl, tareas:tot, resueltas:res,
            pendientes:Math.max(0,tot-res)};
  });
  _dedup((STORE.auditorias||[]).concat(_finsAud).filter(function(a){return rzOk(a)&&mesOk(a.mes);}),
    function(a){ return [_claveTienda(a), norm(a.mes), _claseAuditoria(a)].join('|'); }
  ).forEach(function(a){
    var d=get(a.tienda);
    var pc=parseFloat(a.pctCumpl!==undefined?a.pctCumpl:a.pct_cumpl)||0;
    if(pc>1.5)pc=pc/100;
    d.auditorias++; d.sumPct+=pc*100;
    var k=_claveTienda(a)+'|'+_claseAuditoria(a)+'|'+_mesIdx(a.mes);
    var real=IDX[k];
    if(real){ d.pendientes+=real.pendientes; d.resueltas+=real.resueltas; }
    else { d.pendientes+=(+a.pendientes||0); d.resueltas+=(+a.resueltas||0); }
    d.clases[esOrden(a.clase)?'Orden y Limpieza':'Procesos']=1;
  });
  /* Ajustes y mermas del mes */
  var firmaReg=function(x){
    return [norm(x.tienda), norm(x.mes), x.fechaCorreo||x.fecha_correo||'',
            x.fechaAjuste||x.fecha_ajuste||x.fecha||'', norm(x.condicion||''),
            norm(x.auditor||'')].join('|');
  };
  _dedup(AJUSTES||[], firmaReg).forEach(function(x){
    if(!rzOk(x)||!mesOk(x.mes))return;
    var d=get(x.tienda); d.ajustes++;
    if(x.condicion==='A tiempo')d.ajTiempo++;
    else if(x.condicion==='Destiempo')d.ajDestiempo++;
  });
  _dedup(MERMAS||[], firmaReg).forEach(function(x){
    if(!rzOk(x)||!mesOk(x.mes))return;
    var d=get(x.tienda); d.mermas++;
    if(x.condicion==='A tiempo')d.mmTiempo++;
    else if(x.condicion==='Destiempo')d.mmDestiempo++;
  });

  var tiendas=Object.values(T).map(function(d){
    d.tareas=d.pendientes+d.resueltas;
    d.pctTareas=_pctP(d.resueltas,d.tareas);
    d.pctAud=d.auditorias?Math.round(d.sumPct/d.auditorias):0;
    d.pctAjustes=_pctP(d.ajTiempo, d.ajTiempo+d.ajDestiempo);
    d.pctMermas=_pctP(d.mmTiempo, d.mmTiempo+d.mmDestiempo);
    /* Cumplimiento global de la tienda: promedio de lo que sí aplica */
    var partes=[];
    if(d.tareas)partes.push(d.pctTareas);
    if(d.ajTiempo+d.ajDestiempo)partes.push(d.pctAjustes);
    if(d.mmTiempo+d.mmDestiempo)partes.push(d.pctMermas);
    d.pctGlobal=partes.length?Math.round(partes.reduce(function(a,b){return a+b;},0)/partes.length):0;
    d.clasesTxt=Object.keys(d.clases).join(' · ')||'—';
    return d;
  }).sort(function(a,b){return b.pctGlobal-a.pctGlobal;});

  /* Totales */
  var tot={tareas:0,pendientes:0,resueltas:0,ajustes:0,ajTiempo:0,ajDestiempo:0,mermas:0,mmTiempo:0,mmDestiempo:0};
  tiendas.forEach(function(d){
    tot.tareas+=d.tareas; tot.pendientes+=d.pendientes; tot.resueltas+=d.resueltas;
    tot.ajustes+=d.ajustes; tot.ajTiempo+=d.ajTiempo; tot.ajDestiempo+=d.ajDestiempo;
    tot.mermas+=d.mermas; tot.mmTiempo+=d.mmTiempo; tot.mmDestiempo+=d.mmDestiempo;
  });
  tot.pctTareas=_pctP(tot.resueltas,tot.tareas);
  tot.pctAjustes=_pctP(tot.ajTiempo, tot.ajTiempo+tot.ajDestiempo);
  tot.pctMermas=_pctP(tot.mmTiempo, tot.mmTiempo+tot.mmDestiempo);
  var pg=[]; if(tot.tareas)pg.push(tot.pctTareas);
  if(tot.ajTiempo+tot.ajDestiempo)pg.push(tot.pctAjustes);
  if(tot.mmTiempo+tot.mmDestiempo)pg.push(tot.pctMermas);
  tot.pctGlobal=pg.length?Math.round(pg.reduce(function(a,b){return a+b;},0)/pg.length):0;

  /* Evolución mensual del cumplimiento de tareas (meses elegidos) */
  var evo={meses:[],pct:[]};
  (mesesRango||[]).forEach(function(mn){
    var i=_mesIdx(mn), p=0, r=0;
    _dedup((STORE.auditorias||[]).filter(function(a){return rzOk(a)&&_mesIdx(a.mes)===i;}),
      function(a){ return [_claveTienda(a), norm(a.mes), _claseAuditoria(a)].join('|'); }
    ).forEach(function(a){
      var real=IDX[_claveTienda(a)+'|'+_claseAuditoria(a)+'|'+i];
      if(real){ p+=real.pendientes; r+=real.resueltas; }
      else { p+=(+a.pendientes||0); r+=(+a.resueltas||0); }
    });
    if(p+r===0)return;
    evo.meses.push(mn); evo.pct.push(_pctP(r,p+r));
  });

  var anio=(new Date()).getFullYear();
  var fa=(STORE.auditorias||[]).find(function(a){return a.fecha;});
  if(fa&&fa.fecha){ var y=new Date(fa.fecha).getFullYear(); if(!isNaN(y))anio=y; }

  return {
    mes:mes, anio:anio, razon:razon,
    razonNombre:({KNO:'Kuroda Norte, S.A. de C.V.',KSC:'KS Comercial, S.A. de C.V.',KSA:'Kuroda SA, S.A. de C.V.'})[String(razon).toUpperCase()]||razon,
    tiendas:tiendas, totales:tot, evolucion:evo
  };
}

/* Canal con el iframe de Documentos */
/* ════════ Generador de Dashboard Ejecutivo: razones sociales ════════
   Envía al iframe del generador las razones que la cuenta puede usar. El
   generador ya no es exclusivo de KNO: cada razón (KNO/KSC/KSA) genera su
   propio reporte y guarda su historial por separado. */
function sendRazonesToGenerador(){
  var ifr=document.getElementById('iframe-generador');
  if(!ifr||!ifr.contentWindow)return;
  var razones=(typeof razonesExportables==='function')?razonesExportables():['KNO','KSC','KSA'];
  ifr.contentWindow.postMessage({type:'gen-config',razones:razones,activa:razones[0]},'*');
}
/* Cuando el generador termina de cargar avisa con 'gen-ready'; le respondemos
   con su configuración de razones (por si aún no la tenía).
   Cuando el generador termina un reporte, nos manda el archivo con 'gen-download':
   el dashboard (origen real) hace la descarga, porque el iframe 'srcdoc' del
   generador no puede descargar blobs de forma confiable. */
window.addEventListener('message',function(ev){
  var d=ev.data; if(!d||typeof d!=='object')return;
  var ifr=document.getElementById('iframe-generador');
  if(!ifr||ev.source!==ifr.contentWindow)return; /* solo el iframe del generador */
  if(d.type==='gen-ready'){ sendRazonesToGenerador(); return; }
  if(d.type==='gen-download' && (d.b64||d.blob||d.html)){
    try{
      var mime=d.mime||'application/vnd.openxmlformats-officedocument.presentationml.presentation';
      var blob;
      if(d.html){
        blob=new Blob([d.html],{type:mime});
      } else if(d.b64){
        var bin=atob(d.b64), bytes=new Uint8Array(bin.length);
        for(var i=0;i<bin.length;i++)bytes[i]=bin.charCodeAt(i);
        blob=new Blob([bytes],{type:mime});
      } else { blob=d.blob; }
      var url=URL.createObjectURL(blob);
      var a=document.createElement('a');
      a.href=url; a.download=d.fileName||'Dashboard_Ejecutivo.pptx';
      document.body.appendChild(a); a.click(); a.remove();
      setTimeout(function(){URL.revokeObjectURL(url);},4000);
    }catch(e){ console.error('Descarga del reporte falló',e); toast('⚠ No se pudo descargar el archivo'); }
    return;
  }
});

window.addEventListener('message',function(ev){
  var d=ev.data;
  if(!d||typeof d!=='object')return;
  var ifd=document.getElementById('iframe-documentos');
  if(!ifd||ev.source!==ifd.contentWindow)return; /* solo nuestro iframe */

  if(d.type==='pptx-init'){
    ifd.contentWindow.postMessage({type:'pptx-config',
      razones:razonesExportables(), meses:MESES_PPTX,
      mesActual:MESES_PPTX[(new Date()).getMonth()]},'*');
    return;
  }
  /* La tarjeta 'Cumplimiento por Tienda' de Documentos ahora genera el reporte
     de Cumplimiento del dashboard (generatePpt: KPIs + donut), en vez del
     antiguo reporte de tablas por sucursal. */
  if(d.type==='run-consumos'){
    pedirArchivoConsumos();
    return;
  }
  if(d.type==='consumos-download'){
    generarReporteConsumos();
    return;
  }
  if(d.type==='cumpl-init'){
    (async function(){
      try{
        await asegurarDatosPptx();
        var razones=(typeof razonesExportables==='function')?razonesExportables():[];
        var setT={}; STORE.tareas.concat(STORE.auditorias).forEach(function(x){
          var t=(x.tienda||'').trim(); if(t)setT[t]=1;
        });
        var tiendas=Object.keys(setT).sort();
        ifd.contentWindow.postMessage({type:'cumpl-init-data',razones:razones,tiendas:tiendas},'*');
      }catch(e){
        ifd.contentWindow.postMessage({type:'cumpl-init-data',razones:[],tiendas:[],error:e.message},'*');
      }
    })();
    return;
  }
  if(d.type==='run-cumplimiento'){
    (async function(){
      try{
        await asegurarDatosPptx();
        generatePpt({kpis:d.kpis, secciones:d.secciones, filtro:d.filtro, titulo:d.titulo, empresa:d.empresa});
        ifd.contentWindow.postMessage({type:'cumpl-done'},'*');
      }catch(e){
        ifd.contentWindow.postMessage({type:'cumpl-done',error:e.message},'*');
      }
    })();
    return;
  }
  if(d.type==='cumpl-request'){
    (async function(){
      try{
        var permitidas=razonesExportables();
        var ok=permitidas.some(function(r){return razKey(r)===razKey(d.razon);});
        if(!ok){
          ifd.contentWindow.postMessage({type:'cumpl-data',error:'Tu cuenta no tiene acceso a la razón social '+d.razon},'*');
          return;
        }
        await asegurarDatosPptx();
        var payload=buildCumplimientoDataset(d.mes,d.razon,d.mesesRango);
        ifd.contentWindow.postMessage({type:'cumpl-data',payload:payload},'*');
      }catch(e){
        ifd.contentWindow.postMessage({type:'cumpl-data',error:e.message},'*');
      }
    })();
    return;
  }
  if(d.type==='pptx-request'){
    (async function(){
      try{
        /* SEGURIDAD: nunca entregar datos de una razón no permitida */
        var permitidas=razonesExportables();
        var ok=permitidas.some(function(r){return razKey(r)===razKey(d.razon);});
        if(!ok){
          ifd.contentWindow.postMessage({type:'pptx-data',error:'Tu cuenta no tiene acceso a la razón social '+d.razon},'*');
          return;
        }
        /* Los módulos cargan sus datos al entrar a su vista. Si se abre
           Documentos directamente, AJUSTES/MERMAS/ACTIVIDADES están vacíos y
           la presentación saldría en ceros. Se cargan aquí bajo demanda. */
        await asegurarDatosPptx();
        var payload=buildPptxDataset(d.mes,d.razon,d.mesesRango);
        ifd.contentWindow.postMessage({type:'pptx-data',payload:payload},'*');
      }catch(e){
        ifd.contentWindow.postMessage({type:'pptx-data',error:e.message},'*');
      }
    })();
  }
});

/* ── Detector de versión obsoleta ──────────────────────────────────────
   Si GitHub Pages ya tiene un index.html más nuevo que el que se está
   ejecutando, avisa para recargar. Evita seguir usando una copia cacheada
   (causa habitual de "no veo los cambios" y de archivos generados con la
   versión anterior). */
var BUILD_ID='2026-07-21-fix-gracia-dias-30';
(function(){
  setTimeout(async function(){
    try{
      var r=await fetch(location.href.split('?')[0]+'?v='+Date.now(),{cache:'no-store'});
      var t=await r.text();
      var m=t.match(/var BUILD_ID='([^']+)'/);
      if(m&&m[1]!==BUILD_ID){
        var b=document.createElement('div');
        b.style.cssText='position:fixed;bottom:16px;left:16px;z-index:9999;background:var(--k-orange,#fb6340);color:#fff;padding:11px 16px;border-radius:12px;font-size:13px;font-weight:700;box-shadow:0 8px 24px rgba(0,0,0,.25);cursor:pointer;font-family:inherit';
        b.textContent='⟳ Hay una versión nueva del dashboard — toca para actualizar';
        b.onclick=function(){location.reload(true);};
        document.body.appendChild(b);
      }
    }catch(e){}
  },2500);
})();

setThemeBtn();
/* Refuerzo: el módulo Documentos es visible salvo para cuentas de solo lectura */
(function(){
  var nd=document.getElementById('nav-documentos');
  if(nd&&!(_session&&_session.rol==='viewer'))nd.style.display='';
})();