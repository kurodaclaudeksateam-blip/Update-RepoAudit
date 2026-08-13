/* early declarations for temporal dead zone fix */
var _sb=null, _session=null;

/* ════════════════════════════════════════════════════════════════════
   ESTADO GLOBAL
════════════════════════════════════════════════════════════════════ */
const LS_KEY='cbh_cerezo_store_v1';
let STORE={auditorias:[],tareas:[]};      // datos persistidos
let STAGED={auditorias:[],tareas:[],files:[]}; // buffer de importación
let VIEW='dash';
let charts={};
const MESES=['Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic'];

/* ════════════════════════════════════════════════════════════════════
   PERSISTENCIA (localStorage — acumula entre sesiones)
════════════════════════════════════════════════════════════════════ */
/* Datos NO se persisten en localStorage — la fuente de verdad es Supabase.
   STORE solo vive en memoria durante la sesión para renderizar el dashboard. */
function saveStore(){ /* no-op: los datos viven en Supabase, no en localStorage */ }
function loadStore(){ return false; /* sin caché local de datos */ }
function confirmWipe(){
  openConfirm('🗑️ Borrar datos guardados',
    `Se eliminarán <b>${STORE.tareas.length} tareas</b> y <b>${STORE.auditorias.length} auditorías</b> guardadas en este navegador. Esta acción no se puede deshacer.`,
    'Borrar todo','btn-red',()=>{
      STORE={auditorias:[],tareas:[]};saveStore();closeModal();refreshAll();toast('Datos borrados');
    });
}

/* ════════════════════════════════════════════════════════════════════
   UTILIDADES DE FECHA
════════════════════════════════════════════════════════════════════ */
function parseDate(v){
  if(v===null||v===undefined||v==='')return null;
  /* Fecha real de Excel (SheetJS con cellDates:true la entrega anclada en UTC
     medianoche). Si se deja tal cual, al leerla de vuelta con toISOString()+
     new Date() en una zona horaria negativa (ej. Tijuana UTC-8) se recorre un
     día hacia atrás. Se reconstruye a partir de los componentes UTC pero
     anclada a mediodía LOCAL, para que sobreviva el viaje de ida y vuelta
     sin importar el huso horario del navegador. */
  if(v instanceof Date){
    if(isNaN(v))return null;
    return new Date(v.getUTCFullYear(),v.getUTCMonth(),v.getUTCDate(),12,0,0);
  }
  if(typeof v==='number'&&isFinite(v))return excelSerialToDate(v);
  var s=String(v).trim();
  if(s==='')return null;
  /* Serial de Excel como texto */
  if(/^\d+(\.\d+)?$/.test(s))return excelSerialToDate(parseFloat(s));
  /* ISO yyyy-mm-dd (con o sin hora) → interpretación directa */
  var m=s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})(?:[T ].*)?$/);
  if(m){var dt=new Date(+m[1],+m[2]-1,+m[3],12,0,0);return isNaN(dt)?null:dt;}
  /* dd/mm/yyyy o dd-mm-yyyy (con o sin hora al final) → DÍA / MES / AÑO
     Acepta " 15:42:54" u otro texto tras la fecha (p.ej. exportación del sistema). */
  m=s.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})(?:[ T].*)?$/);
  if(m){
    var d=+m[1], mo=+m[2], y=+m[3];
    if(y<100)y+=2000;
    /* Si el 2º número no puede ser mes (>12) pero el 1º sí, era mm/dd: corregir */
    if(mo>12&&d<=12){var t=d;d=mo;mo=t;}
    var dt2=new Date(y,mo-1,d,12,0,0);
    return isNaN(dt2)?null:dt2;
  }
  var d3=new Date(s);return isNaN(d3)?null:d3;
}
/* Serial de Excel (días desde 1899-12-30 en UTC) -> Date anclada a mediodía LOCAL
   sobre el mismo día calendario, para evitar el corrimiento de zona horaria. */
function excelSerialToDate(serial){
  var u=new Date(Math.round((serial-25569)*86400*1000));
  return new Date(u.getUTCFullYear(),u.getUTCMonth(),u.getUTCDate(),12,0,0);
}
function toISO(d){return d?d.toISOString():null;}
function fromISO(s){
  if(!s)return null;
  // Date-only strings like "2026-03-26" are parsed as UTC by browsers
  // causing off-by-one errors in local timezones (GMT-N)
  // Force local interpretation by appending T00:00:00
  if(typeof s==='string'&&/^\d{4}-\d{2}-\d{2}$/.test(s))return new Date(s+'T00:00:00');
  return new Date(s);
}
function fmtDate(d){return d?d.toLocaleDateString('es-MX',{day:'2-digit',month:'short',year:'numeric'}):'—';}
function fmtInput(d){if(!d)return'';const z=n=>String(n).padStart(2,'0');return `${d.getFullYear()}-${z(d.getMonth()+1)}-${z(d.getDate())}`;}
function pctStr(v){return Math.round((v||0)*100)+'%';}
const daysBetween=(a,b)=>Math.floor((a-b)/86400000);

// etiqueta del bucket de tendencia según granularidad
function bucketKey(d,gran){
  if(!d)return null;
  const y=d.getFullYear(),m=d.getMonth();
  if(gran==='year')return {k:`${y}`,sort:y*100,lbl:`${y}`};
  if(gran==='half'){const h=m<6?1:2;return {k:`${y}-H${h}`,sort:y*100+h,lbl:`${h==1?'1er':'2do'} sem. ${y}`};}
  if(gran==='quarter'){const q=Math.floor(m/3)+1;return {k:`${y}-T${q}`,sort:y*100+q,lbl:`T${q} ${y}`};}
  return {k:`${y}-${m}`,sort:y*100+m,lbl:`${MESES[m]} ${y}`};
}
function bucketKeyEx(d,gran){
  if(!d)return null;
  if(gran==='day'){
    const y=d.getFullYear(),m=d.getMonth(),day=d.getDate();
    const z=n=>String(n).padStart(2,'0');
    return {k:`${y}-${z(m+1)}-${z(day)}`,sort:y*10000+(m+1)*100+day,lbl:`${day} ${MESES[m]}`};
  }
  return bucketKey(d,gran);
}

/* ════════════════════════════════════════════════════════════════════
   PARSEO DE EXCEL (detección por encabezado, tolerante a variaciones)
════════════════════════════════════════════════════════════════════ */
/* norm() se usa como llave de emparejamiento (centro/tienda) entre auditorías,
   tareas, ajustes y mermas. ANTES solo reemplazaba manualmente las vocales y
   la Ñ ya "precompuestas" (un solo carácter Unicode). Pero según el sistema de
   origen (Excel, copiar-pegar, Supabase) el mismo acento puede llegar en forma
   "descompuesta" (una N + tilde combinante como carácter aparte): visualmente
   es idéntico pero como texto NO es el mismo string, así que el reemplazo
   manual no lo detectaba y dos registros de la misma tienda (p.ej. "PEÑASCO")
   dejaban de coincidir según de dónde vinieran. normalize('NFD') descompone
   SIEMPRE ambas formas al mismo patrón (letra + marca combinante), y el regex
   de abajo quita cualquier marca combinante — cubre todos los acentos por
   igual sin depender de una lista manual. */
function norm(s){
  return String(s||'')
    .normalize('NFD').replace(/[\u0300-\u036f]/g,'')
    .replace(/\u00a0/g,' ')
    .toLowerCase().replace(/\s+/g,' ').trim();
}

/* ════════════════════════════════════════════════════════════════════
   DICCIONARIO DE SUCURSALES (nombre canónico por CÓDIGO DE CENTRO)
   El código de centro es la llave confiable; el nombre de tienda varía.
   Al cargar e importar, el nombre se normaliza a partir del centro para
   que auditorías↔tareas emparejen y las cargas actualicen correctamente.
   Para agregar/corregir: edita una línea aquí.
════════════════════════════════════════════════════════════════════ */
var CENTRO_TIENDA={
  'KC00':'CEDIS PALMITO',
  'KC01':'BRAVO',
  'KC02':'CALZADA',
  'KC03':'PROYECTOS',
  'KC04':'PALMITO',
  'KC05':'SANALONA',
  'KC06':'GUAMUCHIL',
  'KC13':'BIENESTAR',
  'KC15':'OUTLET MOCHIS',
  'KC16':'ECONOSALDOS',
  'KC17':'GUASAVE',
  'KC20':'NAVOLATO',
  'KC22':'HUERTAS',
  'KC23':'PAPALOTE',
  'KC24':'COSTARICA',
  'KC26':'PROYECTOS MAZATLAN',
  'KN00':'KN 20 DE NOV MENUDEO',
  'KN01':'KN EXPRESS',
  'KN03':'KN ENSENADA MAYOREO',
  'KN04':'KN ENSENADA MENUDEO',
  'KN05':'KN MEXICALI MENUDEO',
  'KN06':'KN MEXICALI MAYOREO',
  'KN07':'KN PEÑASCO',
  'KN09':'KN MOSAIKOS',
  'KN14':'KN TIJUANA PROYECTOS',
  'KN15':'KN BO VALLE DE GPE.',
  'KN16':'KN CEDIS TECATE',
  'KN17':'KN ROSARITO',
  'KN18':'KN PEDREGAL',
  'KN19':'KN TECATE',
  'KN21':'NOGALES',
  'KS00':'KS MATRIZ',
  'KS01':'KS PROYECTOS',
  'KS02':'KS PLOMELEC PERIFERICO',
  'KS03':'KS TIANGUIS PERISUR',
  'KS05':'KS TIANGUIS HILLO',
  'KS07':'KS OBREGON',
  'KS08':'KS NAVOJOA',
  'KS09':'KS MORELOS',
  'KS10':'CEDIS PALO VERDE',
  'KS11':'KS HUATABAMPO',
  'KS15':'KS SAN JOSE DEL CABO',
  'KS16':'KS CABOS SAN LUCAS',
  'KS17':'KS TIANGUIS OBREGON',
  'KS21':'KS PLOMELEC EXPRESS'
};
/* Alias de centros: algunos centros comparten la misma tienda física bajo más
   de un código (histórico, de sistema origen, etc.). KN06 y KN18 son la misma
   tienda — cualquier dato que llegue etiquetado KN06 (carga de Excel, fila de
   Supabase, captura manual) se centra en KN18 desde este único punto, para no
   fragmentar sus auditorías/tareas en dos y mantener la actualización correcta. */
var CENTRO_ALIAS={'KN06':'KN18'};
function canonCentro(c){
  var cc=String(c||'').toUpperCase().replace(/\s+/g,'').trim();
  return CENTRO_ALIAS[cc]||cc;
}
/* Devuelve el nombre canónico de la sucursal a partir del centro; si el
   centro no está en el diccionario, conserva el nombre recibido. */
function canonTienda(centro,raw){
  var cc=canonCentro(centro);
  if(cc&&CENTRO_TIENDA[cc])return CENTRO_TIENDA[cc];
  return String(raw||'').trim();
}
function centroConocido(centro){return !!CENTRO_TIENDA[canonCentro(centro)];}
/* Razón social por prefijo de centro (hoja "Tiendas válidas"):
   KC→KSA · KN→KNO · KS→KS. Relación única centro↔razón. */
var PREFIJO_RAZON={KC:'KSA',KN:'KNO',KS:'KSC'};
function razonDeCentro(c){return PREFIJO_RAZON[canonCentro(c).slice(0,2)]||'';}

/* ════════ RAZÓN SOCIAL: clave compuesta y visibilidad ════════
   La llave única de una tarea es ID + RAZÓN (una misma tarea puede existir
   en dos razones sociales distintas). razKey normaliza igual que el backfill
   SQL (minúsculas, espacios colapsados). */
function razKey(s){return String(s||'').replace(/\s+/g,' ').trim().toLowerCase();}
function tareaKey(id,razon){return String(id)+'|'+razKey(razon);}
function _razonesAsignadas(){
  /* Los viewers son de solo observación: siempre ven las 3 razones sociales
     (KNO/KSC/KSA) completas, sin importar qué traiga razones_permitidas en su
     cuenta. La restricción por razón es para limitar edición/captura por
     entidad (auditores); un viewer no modifica nada, así que no aplica. */
  if(_session&&_session.rol==='viewer')return null;
  var r=_session&&_session.razones_permitidas;
  if(typeof r==='string'){try{r=JSON.parse(r);}catch(e){r=null;}}
  return (r&&r.length)?r:null;
}
/* Visibilidad por razón social: usuarios con razones asignadas (auditores,
   viewers restringidos) solo ven su(s) razón(es). admin ve todo; usuarios
   sin razones asignadas ("vista a todo") ven todo. */
function razonVisible(razon){
  if(!_session||_session.rol==='admin')return true;
  var rs=_razonesAsignadas();if(!rs)return true;
  return rs.some(function(x){return razKey(x)===razKey(razon);});
}
/* Para módulos sin campo razón (ajustes/mermas/finalizadas): se deriva el
   centro desde la tienda con el diccionario y de ahí la razón social exacta
   (KC→KSA, KN→KNO, KS→KS), evitando el choque entre KSA y KS. */
/* Alias de TIENDA (no de centro): nombres reales usados en Auditorías/Tareas/
   Excel que no calzan ni exacto ni por substring con el nombre del catálogo
   (palabras en otro orden o abreviadas). Sin esto, Ajustes/Mermas de estas
   tiendas quedan sin razón social detectada y se ocultan por seguridad. */
var TIENDA_ALIAS_CENTRO={
  '20 de noviembre':'KN00',        /* catálogo: "20 DE NOV MENUDEO" */
  'valle de guadalupe':'KN15'      /* catálogo: "BO VALLE DE GPE." */
};
var _TIENDA_ALIAS_NORM=null;
function _buildTiendaAliasIndex(){
  if(_TIENDA_ALIAS_NORM)return;
  _TIENDA_ALIAS_NORM={};
  Object.keys(TIENDA_ALIAS_CENTRO).forEach(function(k){_TIENDA_ALIAS_NORM[norm(k)]=TIENDA_ALIAS_CENTRO[k];});
}
/* Compara dos nombres normalizados ignorando el ORDEN de las palabras: cubre
   casos como "PROYECTOS TIJUANA" (dato real) vs "TIJUANA PROYECTOS" (catálogo),
   que la coincidencia por substring simple no detecta porque el orden difiere. */
function _tokMatch(a,b){
  var wa=a.split(' ').filter(Boolean), wb=b.split(' ').filter(Boolean);
  if(!wa.length||!wb.length)return false;
  var corta=wa.length<=wb.length?wa:wb, larga=wa.length<=wb.length?wb:wa;
  return corta.every(function(w){return larga.indexOf(w)>=0;});
}
var _TIENDA_CENTRO=null;
var _TIENDA_CENTRO_KEYS=null;
var _TIENDA_CENTRO_AMBIGUAS=null; /* claves normalizadas que existen en más de una razón social */
function _buildTiendaCentroIndex(){
  if(_TIENDA_CENTRO)return;
  _TIENDA_CENTRO={};_TIENDA_CENTRO_KEYS=[];_TIENDA_CENTRO_AMBIGUAS={};
  var _razonPorClave={}; /* clave -> {razon: centro} visto hasta ahora, para detectar choque entre razones */
  Object.keys(CENTRO_TIENDA).forEach(function(c){
    var full=CENTRO_TIENDA[c];
    var corto=full.replace(/^(KC|KN|KS)\s+/i,'');
    var nf=norm(full), nc=norm(corto);
    var raz=razonDeCentro(c);
    [nf,nc].forEach(function(k){
      if(!k)return;
      _razonPorClave[k]=_razonPorClave[k]||{};
      /* Si la misma clave normalizada ya pertenece a OTRA razón social, es
         ambigua: no se puede resolver ciegamente al primer centro insertado
         (bug detectado con "PROYECTOS" = KC03/KSA y KS01/KS). Se marca para
         que centroDeTienda() la trate como no resuelta y caiga al fallback
         de coincidencia parcial, que sí valida ambigüedad entre razones. */
      if(Object.keys(_razonPorClave[k]).length&&!_razonPorClave[k][raz])_TIENDA_CENTRO_AMBIGUAS[k]=true;
      _razonPorClave[k][raz]=c;
    });
    if(!_TIENDA_CENTRO[nf])_TIENDA_CENTRO[nf]=c;
    if(!_TIENDA_CENTRO[nc])_TIENDA_CENTRO[nc]=c;
    _TIENDA_CENTRO_KEYS.push({nf:nf,nc:nc,c:c});
  });
}
function centroDeTienda(t){
  _buildTiendaCentroIndex();_buildTiendaAliasIndex();
  var nt=norm(t);
  if(!nt)return'';
  if(_TIENDA_ALIAS_NORM[nt])return _TIENDA_ALIAS_NORM[nt];
  if(_TIENDA_CENTRO[nt]&&!_TIENDA_CENTRO_AMBIGUAS[nt])return _TIENDA_CENTRO[nt];
  /* Fallback por coincidencia parcial: se recopilan TODOS los candidatos y
     solo se resuelve si TODOS pertenecen a la misma razón social. Nombres
     genéricos como "PROYECTOS" existen a la vez en KC03 (KSA) y KS01 (KSC):
     tomar el primer match a ciegas atribuía el registro a la razón
     equivocada (mermas/ajustes de KSC apareciendo en KSA, o viceversa). Si el
     nombre es ambiguo entre razones distintas, se retorna '' para que el
     registro se oculte por seguridad en vez de adivinar mal. También se
     compara ignorando el orden de las palabras (_tokMatch) para cubrir
     nombres reales con las mismas palabras en otro orden ("PROYECTOS
     TIJUANA" vs "TIJUANA PROYECTOS"). */
  var candidatos=[];
  for(var i=0;i<_TIENDA_CENTRO_KEYS.length;i++){
    var k=_TIENDA_CENTRO_KEYS[i];
    if(k.nf.indexOf(nt)>=0||nt.indexOf(k.nf)>=0||k.nc.indexOf(nt)>=0||nt.indexOf(k.nc)>=0||
       _tokMatch(nt,k.nf)||_tokMatch(nt,k.nc))candidatos.push(k.c);
  }
  if(!candidatos.length)return '';
  var razones=uniq(candidatos.map(function(c){return razonDeCentro(c);}));
  if(razones.length===1)return candidatos[0];
  return '';
}
function tiendaVisible(tienda){
  if(!_session||_session.rol==='admin')return true;
  var rs=_razonesAsignadas();if(!rs)return true;
  var raz=razonDeCentro(centroDeTienda(tienda));
  if(!raz)return false; /* tienda no identificable: se oculta por seguridad */
  return rs.some(function(x){return razKey(x)===razKey(raz);});
}
function esc(s){return String(s==null?'':s).replace(/&/g,'&amp;').replace(/"/g,'&quot;').replace(/</g,'&lt;').replace(/>/g,'&gt;');}

/* Normaliza el Estado de una tarea al importar desde Excel. Acepta tanto los
   textos en español ("Abierta", "Abierta atrasada", "Resuelta", "Resuelta
   Atrasada") como los códigos en inglés/snake_case usados por otros sistemas
   de exportación: open, open_and_late, resolved, resolved_and_late (también
   con espacios o guiones en vez de guión bajo, y sin importar mayúsculas). */
var ESTADO_CODIGOS={
  'open':'Abierta',
  'open_and_late':'Abierta atrasada',
  'resolved':'Resuelta',
  'resolved_and_late':'Resuelta Atrasada'
};
function normalizarEstadoTarea(raw){
  var s=String(raw||'').replace(/\u00a0/g,' ').trim();
  if(!s)return s;
  var key=s.toLowerCase().replace(/[\s\-]+/g,'_');
  if(ESTADO_CODIGOS[key])return ESTADO_CODIGOS[key];
  return s; /* ya viene en español u otro texto: se deja tal cual */
}
function findCol(headers,keys){
  for(let i=0;i<headers.length;i++){const h=norm(headers[i]);
    if(keys.some(k=>h.includes(k)))return i;}
  return -1;
}
/* Devuelve el valor de la columna principal; si viene vacío (p.ej. una fórmula
   sin recalcular en la plantilla), usa la columna de respaldo "PEGA AQUÍ ↴". */
function findColRe(headers,re){
  for(let i=0;i<headers.length;i++){if(re.test(norm(headers[i])))return i;}
  return -1;
}
function valCol(row,mainIdx,fallbackIdx){
  var v=mainIdx>=0?row[mainIdx]:undefined;
  if(v!==undefined&&v!==null&&String(v).trim()!=='')return v;
  if(fallbackIdx!=null&&fallbackIdx>=0){
    var fb=row[fallbackIdx];
    if(fb!==undefined&&fb!==null&&String(fb).trim()!=='')return fb;
  }
  return v;
}
function tipoNorm(t){
  const n=norm(t);
  if(n.includes('orden')||n.includes('limpieza'))return 'ol';
  if(n.includes('cartera'))return 'cartera';
  if(n.includes('col')||n.includes('colab')||n.includes('auditoria col'))return 'col';
  return 'col';
}
function parseWorkbook(buf){
  const wb=XLSX.read(buf,{type:'array',cellDates:true});
  const out={auditorias:[],tareas:[]};
  wb.SheetNames.forEach(sn=>{
    const raw=XLSX.utils.sheet_to_json(wb.Sheets[sn],{header:1,defval:'',raw:false,cellDates:true});
    /* Copia "cruda" de la misma hoja: si una celda es una fecha real de Excel,
       raw:true la entrega como objeto Date verdadero (basado en el valor interno,
       sin pasar por texto). Esto evita que raw:false la reformatee a un string
       cuyo orden día/mes puede ser ambiguo o distinto al de la plantilla. */
    const rawTyped=XLSX.utils.sheet_to_json(wb.Sheets[sn],{header:1,defval:'',raw:true,cellDates:true});
    function fechaCelda(i,idx,fallbackIdx){
      const rt=rawTyped[i]||[];
      if(idx!=null&&idx>=0&&rt[idx]instanceof Date)return rt[idx];
      const r=raw[i]||[];
      const v=valCol(r,idx,fallbackIdx);
      if(v!==undefined&&v!==null&&String(v).trim()!=='')return v;
      if(fallbackIdx!=null&&fallbackIdx>=0&&rt[fallbackIdx]instanceof Date)return rt[fallbackIdx];
      return v;
    }
    if(raw.length<2)return;
    // localizar fila de encabezado (la que tenga "razon" o "id" o "tienda")
    let hi=0;
    for(let i=0;i<Math.min(6,raw.length);i++){
      const joined=raw[i].map(norm).join('|');
      if(joined.includes('razon')||joined.includes('tienda')||joined.includes('centro')||joined.includes('nombre de tarea')
         ||joined.includes('id ticket')||joined.includes('codigo del local')||joined.includes('nombre del local')){hi=i;break;}
    }
    const H=raw[hi];

    /* ═══ Formato NATIVO del sistema (exportación "StoreAction") ═══
       Se detecta por sus encabezados propios. Permite cargar el reporte
       tal cual sale del sistema, sin pasar por la plantilla. */
    const idxIdTicket=findCol(H,['id ticket']);
    const idxCodLocal=findCol(H,['codigo del local','codigo local']);
    const idxNomLocal=findCol(H,['nombre del local','nombre local']);
    const esStoreAction=(idxIdTicket>=0 && idxCodLocal>=0);
    if(esStoreAction){
      const c={
        id:idxIdTicket,
        centro:idxCodLocal,
        tienda:idxNomLocal,
        razon:findCol(H,['cadena']),
        titulo:findCol(H,['titulo del ticket','titulo ticket']),
        descripcion:findCol(H,['descripcion ticket','descripcion del ticket']),
        areaResp:findCol(H,['area responsable']),
        areaRev:findCol(H,['area que revisa','area revisa']),
        estado:findCol(H,['estado']),
        fechaCre:findCol(H,['fecha de creacion','fecha creacion']),
        fechaCierre:findCol(H,['fecha de cierre','fecha cierre']),
        actividad:findCol(H,['nombre actividad','nombre de actividad'])
      };
      for(let i=hi+1;i<raw.length;i++){
        const r=raw[i];const idv=r[c.id];
        if(idv===''||idv===undefined||idv===null)continue;
        if(norm(idv)==='id ticket')continue;
        const num=String(idv).replace(/\D/g,'');
        if(!num)continue;
        const nombreAct=String(r[c.actividad]||'').trim();
        /* Regla: actividad en blanco = Auditoría de Colaboración;
           si trae texto, ese texto define el tipo (Orden y Limpieza, etc.).
           tipoTarea se normaliza SIEMPRE a uno de los 3 valores canónicos que
           ya usa el resto del sistema (igual criterio que tipoNormLocal) —
           antes se guardaba el texto crudo de "Nombre Actividad" tal cual
           (p.ej. "AUDITORÍA DE ORDEN Y LIMPIEZA v02"), que nunca coincide con
           el canónico ya guardado en Supabase y marcaba la tarea como
           "cambiada" en cada re-importación aunque nada real hubiera cambiado. */
        const nombreActN=norm(nombreAct);
        const tipoTarea = !nombreAct ? 'TAREAS AUDITORIA COL'
          : (nombreActN.includes('orden')||nombreActN.includes('limpieza')) ? 'TAREAS ORDEN Y LIMPIEZA'
          : nombreActN.includes('cartera') ? 'TAREAS CARTERA'
          : 'TAREAS AUDITORIA COL';
        /* El centro trae el nombre canónico de tienda; se respeta el del diccionario.
           canonCentro también resuelve alias de centro (p.ej. KN06 se centra en KN18). */
        const centro=canonCentro(String(r[c.centro]||'').trim());
        out.tareas.push({
          razon:String(r[c.razon]||'').trim(),
          centro,
          id:isNaN(Number(idv))?String(idv).trim():Number(num),
          fechaCreacion:toISO(parseDate(fechaCelda(i,c.fechaCre,-1))),
          tienda:canonTienda(centro,r[c.tienda]),
          areaResp:String(r[c.areaResp]||'').trim(),
          areaRev:String(r[c.areaRev]||'').trim(),
          /* "Nombre Actividad" vacío = auditoría de Colaboración (dato manual
             que el usuario captura en el archivo); NO significa que a la
             tarea le falte nombre. El "Título del Ticket" siempre trae el
             nombre real de la tarea, venga o no "Nombre Actividad". */
          actividad:nombreAct||String(r[c.titulo]||'').trim(),
          nombre:String(r[c.titulo]||'').trim(),
          estado:normalizarEstadoTarea(r[c.estado]),
          fechaTerm:null,   /* pendiente: aún no aparece en el reporte del sistema */
          fechaCumpl:toISO(parseDate(fechaCelda(i,c.fechaCierre,-1))),
          tipoTarea
        });
      }
      return; /* hoja procesada como StoreAction */
    }
    const idxId=findCol(H,['id ','id']);
    const idxNombreTarea=findCol(H,['nombre de tarea','nombre tarea']);
    const idxTipoTarea=findCol(H,['tipo tarea','tipo de tarea']);
    const idxEstado=findCol(H,['estado']);
    const idxPctCumpl=findCol(H,['porcentaje cumplimiento','porcentaje de cumplimiento','% cumplimiento']);
    const idxClase=findCol(H,['clase']);

    // ¿Es hoja de TAREAS? (tiene ID + nombre de tarea/estado)
    const esTareas=(idxId>=0 && (idxNombreTarea>=0||idxTipoTarea>=0||idxEstado>=0) && idxPctCumpl<0);
    // ¿Es hoja de CONTROL/auditorías? (tiene % cumplimiento + clase, sin ID de tarea)
    const esControl=(idxPctCumpl>=0 || idxClase>=0) && idxNombreTarea<0;

    if(esTareas){
      const c={razon:findCol(H,['razon']),centro:findCol(H,['centro']),id:idxId,
        fechaCre:findCol(H,['fecha de creacion','fecha creacion']),
        tienda:findCol(H,['tienda']),areaResp:findCol(H,['area de responsabilidad','responsabilidad']),
        areaRev:findCol(H,['area revisora','revisora']),act:findCol(H,['actividad']),
        nombre:idxNombreTarea,estado:idxEstado,
        term:findCol(H,['fecha de termino','fecha termino']),
        cumpl:findCol(H,['fecha de cumplimiento','fecha cumplimiento']),tipo:idxTipoTarea,
        /* Columnas "PEGA AQUÍ ↴" de la plantilla: respaldo cuando la columna
           principal viene vacía (p.ej. si trae una fórmula sin recalcular). */
        pegaCre:findColRe(H,/pega aqui.*creacion/),
        pegaTerm:findColRe(H,/pega aqui.*termino/),
        pegaCumpl:findColRe(H,/pega aqui.*cumplimiento/)};
      for(let i=hi+1;i<raw.length;i++){
        const r=raw[i];const idv=r[c.id];
        if(idv===''||idv===undefined||idv===null)continue;
        if(norm(idv)==='id')continue;
        const num=String(idv).replace(/\D/g,'');
        if(!num)continue;
        out.tareas.push({
          razon:String(r[c.razon]||'').trim(),centro:canonCentro(String(r[c.centro]||'').trim()),
          id:isNaN(Number(idv))?String(idv).trim():Number(num),
          fechaCreacion:toISO(parseDate(fechaCelda(i,c.fechaCre,c.pegaCre))),
          tienda:canonTienda(String(r[c.centro]||'').trim(),r[c.tienda]),
          areaResp:String(r[c.areaResp]||'').trim(),areaRev:String(r[c.areaRev]||'').trim(),
          actividad:String(r[c.act]||'').trim(),nombre:String(r[c.nombre]||'').trim(),
          estado:normalizarEstadoTarea(r[c.estado]),
          fechaTerm:toISO(parseDate(fechaCelda(i,c.term,c.pegaTerm))),fechaCumpl:toISO(parseDate(fechaCelda(i,c.cumpl,c.pegaCumpl))),
          tipoTarea:String(r[c.tipo]||'').trim()
        });
      }
    }else if(esControl){
      const c={razon:findCol(H,['razon']),centro:findCol(H,['centro']),tienda:findCol(H,['tienda']),
        fecha:findCol(H,['fecha de creacion','fecha creacion']),mes:findCol(H,['mes auditoria','mes']),
        pct:idxPctCumpl,tareas:findCol(H,['tareas']),
        pend:findCol(H,['pendientes']),res:findCol(H,['resueltos','resueltas']),
        pctRes:findCol(H,['porcentaje resuelto','% resuelto']),clase:idxClase,
        pegaFecha:findColRe(H,/pega aqui.*creacion/)};
      for(let i=hi+1;i<raw.length;i++){
        const r=raw[i];const centroA=canonCentro(String(r[c.centro]||'').trim());
        const tienda=canonTienda(centroA,r[c.tienda]);
        const razon=String(r[c.razon]||'').trim();
        if(!tienda&&!razon)continue;
        const f=parseDate(fechaCelda(i,c.fecha,c.pegaFecha));
        out.auditorias.push({
          razon,centro:centroA,tienda,
          fecha:toISO(f),mes:String(r[c.mes]||'').trim(),
          pctCumpl:parseFloat(String(r[c.pct]).replace('%',''))||0,
          tareas:parseInt(r[c.tareas])||0,pendientes:parseInt(r[c.pend])||0,
          resueltas:parseInt(r[c.res])||0,pctResuelto:parseFloat(String(r[c.pctRes]).replace('%',''))||0,
          clase:String(r[c.clase]||'').trim()
        });
      }
    }
  });
  // normalizar pctCumpl/pctResuelto a fracción (si vienen como 74 -> 0.74)
  out.auditorias.forEach(a=>{if(a.pctCumpl>1.5)a.pctCumpl/=100;if(a.pctResuelto>1.5)a.pctResuelto/=100;});
  return out;
}

/* ════════════════════════════════════════════════════════════════════
   IMPORTAR (carga -> staging -> generar/combinar con dedupe por ID)
════════════════════════════════════════════════════════════════════ */
function stageFiles(fileList){
  const files=[...fileList];if(!files.length)return;
  const prog=document.getElementById('prog'),fill=prog.firstElementChild;
  prog.style.display='block';fill.style.width='20%';
  let done=0;
  files.forEach(file=>{
    const reader=new FileReader();
    reader.onload=e=>{
      try{
        const parsed=parseWorkbook(e.target.result);
        STAGED.auditorias.push(...parsed.auditorias);
        STAGED.tareas.push(...parsed.tareas);
        STAGED.files.push(file.name);
      }catch(err){toast('⚠ Error en '+file.name+': '+err.message);}
      done++;fill.style.width=(20+done/files.length*70)+'%';
      if(done===files.length){
        fill.style.width='100%';setTimeout(()=>{prog.style.display='none';fill.style.width='0'},400);
        renderStaged();
      }
    };
    reader.readAsArrayBuffer(file);
  });
  document.getElementById('file-input').value='';
}
function renderStaged(){
  const box=document.getElementById('staged-box');
  const n=STAGED.tareas.length+STAGED.auditorias.length;
  if(!n){box.innerHTML='';document.getElementById('btn-generar').disabled=true;return;}
  /* Validar centros contra el diccionario de sucursales */
  var centros=[...new Set([...STAGED.tareas,...STAGED.auditorias]
    .map(function(r){return canonCentro(r.centro);}).filter(Boolean))];
  var desconocidos=centros.filter(function(c){return !CENTRO_TIENDA[c];});
  var okCount=centros.length-desconocidos.length;
  var valida='<span class="staged-chip" style="background:#dcfce7;color:#166534;margin-left:8px">✓ '+okCount+' centro(s) reconocido(s)</span>';
  var alerta=desconocidos.length?
    '<span class="staged-chip" style="background:#fef3c7;color:#92400e;margin-left:8px" title="Estos centros no están en el diccionario; se usará el nombre tal cual venga y podrían no empatar">⚠ Sin diccionario: '+desconocidos.join(', ')+'</span>':'';
  box.innerHTML=`<span class="staged-chip">📥 En cola: ${STAGED.tareas.length} tareas · ${STAGED.auditorias.length} auditorías
    <span class="x" onclick="clearStaged()" title="Descartar cola">✕</span></span>`+valida+alerta;
  document.getElementById('btn-generar').disabled=false;
}
function clearStaged(){STAGED={auditorias:[],tareas:[],files:[]};renderStaged();}

function audKey(a){return [norm(a.razon),norm(a.centro),norm(a.tienda),a.fecha||'',norm(a.clase)].join('|');}

async function generarImport(){
  if(!STAGED.tareas.length&&!STAGED.auditorias.length){toast('No hay datos en cola para generar');return;}
  var client=getSbClient();
  if(!client){toast('⚠ Sin conexión a Supabase. No se puede validar.');return;}

  /* Deduplicar por ID+razón: un mismo archivo (p.ej. exportación StoreAction)
     puede traer la misma fila repetida más de una vez. Sin esto, cada copia
     se procesa por separado y se cuenta/actualiza dos veces. Se conserva la
     última copia (son idénticas cuando el archivo trae duplicados exactos). */
  var _porClave={};
  STAGED.tareas.forEach(function(t){_porClave[tareaKey(t.id,t.razon)]=t;});
  STAGED.tareas=Object.values(_porClave);

  toast('⏳ Validando contra Supabase…');

  /* ── Traer registros existentes de Supabase por ID ── */
  var stagedIds=STAGED.tareas.map(function(t){return String(t.id);}).filter(Boolean);
  var existentes={};
  try{
    /* Buscar en chunks de 200 IDs */
    for(var i=0;i<stagedIds.length;i+=200){
      var chunk=stagedIds.slice(i,i+200);
      var r=await client.from('tareas').select('*').in('tarea_id',chunk);
      if(r.error){toast('⚠ Error consultando Supabase: '+r.error.message);return;}
      /* Descifrar para comparar y para armar la clave ID+RAZÓN */
      var rDec=await decArr(r.data||[],FIELDS.tareas);
      rDec.forEach(function(row){existentes[tareaKey(row.tarea_id,row.razon)]=row;});
    }
  }catch(e){toast('⚠ Error: '+e.message);return;}

  /* ── Clasificar cada tarea staged ── */
  var nuevas=[], idénticas=[], diferentes=[];
  STAGED.tareas.forEach(function(t){
    var ex=existentes[tareaKey(t.id,t.razon)];
    if(!ex){nuevas.push(t);return;}
    /* Comparar campo por campo */
    if(tareaIgualSupabase(t,ex)){idénticas.push(t);}
    else{diferentes.push({nueva:t,actual:ex});}
  });

  /* ── Auditorías: validar por clave compuesta ── */
  var audStaged=STAGED.auditorias.slice();

  /* ── Si hay diferencias, pedir permiso ── */
  if(diferentes.length>0){
    askUpdatePermission(nuevas, idénticas, diferentes, audStaged);
  } else {
    /* Solo insertar nuevas, omitir idénticas */
    await commitToSupabase(nuevas, [], audStaged, idénticas.length, []);
  }
}

/* Compara una tarea staged con la fila de Supabase — true si son idénticas */
function tareaIgualSupabase(t, ex){
  function n(v){return (v===null||v===undefined||v==='')?'':String(v).trim();}
  function fdate(v){return v?String(v).split('T')[0]:'';}
  return n(t.razon)===n(ex.razon) &&
         n(t.centro)===n(ex.centro) &&
         n(t.tienda)===n(ex.tienda) &&
         n(t.areaResp)===n(ex.area_resp) &&
         n(t.areaRev)===n(ex.area_rev) &&
         n(t.actividad)===n(ex.actividad) &&
         n(t.nombre)===n(ex.nombre) &&
         n(t.tipoTarea)===n(ex.tipo_tarea) &&
         n(t.estado)===n(ex.estado) &&
         fdate(t.fechaCreacion)===fdate(ex.fecha_creacion) &&
         /* fecha_term NO se compara: es de captura manual y no debe verse
            afectada por cargas/actualizaciones de Excel (ver toRowActualizar). */
         fdate(t.fechaCumpl)===fdate(ex.fecha_cumpl);
}

/* Lista qué campos cambiaron entre staged y Supabase, con el valor anterior
   y el nuevo, para que se pueda decidir con información real si conviene
   actualizar en vez de solo ver "Otros campos". Cubre los mismos campos que
   compara tareaIgualSupabase — así nada queda oculto en un cajón genérico. */
function camposDiferentes(t, ex){
  function n(v){return (v===null||v===undefined||v==='')?'':String(v).trim();}
  function fdate(v){return v?String(v).split('T')[0]:'';}
  function corto(v,max){v=n(v)||'—';return v.length>max?v.slice(0,max)+'…':v;}
  var campos=[];
  if(n(t.estado)!==n(ex.estado))campos.push('Estado: "'+corto(ex.estado,30)+'" → "'+corto(t.estado,30)+'"');
  if(fdate(t.fechaCumpl)!==fdate(ex.fecha_cumpl))campos.push('F.Cumplimiento: "'+corto(fdate(ex.fecha_cumpl),20)+'" → "'+corto(fdate(t.fechaCumpl),20)+'"');
  if(fdate(t.fechaCreacion)!==fdate(ex.fecha_creacion))campos.push('F.Creación: "'+corto(fdate(ex.fecha_creacion),20)+'" → "'+corto(fdate(t.fechaCreacion),20)+'"');
  if(n(t.nombre)!==n(ex.nombre))campos.push('Nombre: "'+corto(ex.nombre,40)+'" → "'+corto(t.nombre,40)+'"');
  if(n(t.actividad)!==n(ex.actividad))campos.push('Actividad: "'+corto(ex.actividad,40)+'" → "'+corto(t.actividad,40)+'"');
  if(n(t.tipoTarea)!==n(ex.tipo_tarea))campos.push('Tipo de tarea: "'+corto(ex.tipo_tarea,30)+'" → "'+corto(t.tipoTarea,30)+'"');
  if(n(t.areaResp)!==n(ex.area_resp))campos.push('Área Resp.: "'+corto(ex.area_resp,25)+'" → "'+corto(t.areaResp,25)+'"');
  if(n(t.areaRev)!==n(ex.area_rev))campos.push('Área Revisa: "'+corto(ex.area_rev,25)+'" → "'+corto(t.areaRev,25)+'"');
  if(n(t.tienda)!==n(ex.tienda))campos.push('Tienda: "'+corto(ex.tienda,30)+'" → "'+corto(t.tienda,30)+'"');
  if(n(t.centro)!==n(ex.centro))campos.push('Centro: "'+corto(ex.centro,15)+'" → "'+corto(t.centro,15)+'"');
  if(n(t.razon)!==n(ex.razon))campos.push('Razón: "'+corto(ex.razon,15)+'" → "'+corto(t.razon,15)+'"');
  if(!campos.length)campos.push('⚠ Diferencia no identificada — revisar manualmente');
  return campos;
}

function askUpdatePermission(nuevas, idénticas, diferentes, audStaged){
  var rows=diferentes.slice(0,10).map(function(d){
    var difs=camposDiferentes(d.nueva, d.actual);
    return '<tr><td><b>'+d.nueva.id+'</b></td><td>'+(d.nueva.tienda||'')+'</td>'+
           '<td class="tsub" style="font-size:11px">'+difs.join('<br>')+'</td></tr>';
  }).join('');
  var more=diferentes.length>10?'<div class="tsub" style="padding:6px 12px">…y '+(diferentes.length-10)+' más</div>':'';

  var html='<p style="font-size:13.5px;color:var(--txt);margin-bottom:14px">'+
    'Validación contra <b>Supabase</b> completada:</p>'+
    '<div style="display:flex;gap:18px;flex-wrap:wrap;margin-bottom:16px">'+
      '<div class="stat" style="border:none"><b style="color:var(--green);font-size:22px">'+nuevas.length+'</b><span style="font-size:11px;color:var(--muted)">nuevas (se agregarán)</span></div>'+
      '<div class="stat" style="border:none"><b style="color:var(--muted);font-size:22px">'+idénticas.length+'</b><span style="font-size:11px;color:var(--muted)">idénticas (se omiten)</span></div>'+
      '<div class="stat" style="border:none"><b style="color:var(--orange);font-size:22px">'+diferentes.length+'</b><span style="font-size:11px;color:var(--muted)">con cambios</span></div>'+
    '</div>'+
    '<p style="font-size:12.5px;color:var(--txt);margin-bottom:8px"><b>Tareas que YA existen en Supabase con cambios:</b></p>'+
    '<div class="tbl-scroll"><table class="dt" style="min-width:0"><thead><tr><th>ID</th><th>Sucursal</th><th>Cambios detectados</th></tr></thead><tbody>'+rows+'</tbody></table></div>'+more+
    '<p style="font-size:12px;color:var(--muted);margin-top:14px">'+
      '<b style="color:var(--orange)">Actualizar</b>: aplica los cambios a las tareas existentes · '+
      '<b style="color:var(--muted)">Solo nuevas</b>: ignora los cambios y conserva lo que ya está en Supabase.</p>';

  openModal('🔄 Tareas existentes con cambios', html, [
    {label:'Solo agregar nuevas',cls:'btn-ghost',fn:function(){commitToSupabase(nuevas,[],audStaged,idénticas.length,[]);}},
    {label:'Actualizar las que cambiaron',cls:'btn-orange',fn:function(){commitToSupabase(nuevas,diferentes.map(function(d){return d.nueva;}),audStaged,idénticas.length,diferentes.map(function(d){return d.actual;}));}}
  ]);
}

async function commitToSupabase(nuevas, actualizar, audStaged, omitidas, prevTareas){
  var client=getSbClient();
  if(!client){toast('⚠ Sin conexión a Supabase');return;}
  closeModal();
  toast('⏳ Guardando en Supabase…');

  var insOk=0, updOk=0, audOk=0, err=null;

  function toRow(t){return{
    tarea_id:String(t.id),tarea_key:tareaKey(t.id,t.razon),
    razon:t.razon||null,centro:t.centro||null,tienda:t.tienda||null,
    area_resp:t.areaResp||null,area_rev:t.areaRev||null,actividad:t.actividad||null,
    nombre:t.nombre||null,tipo_tarea:t.tipoTarea||null,estado:t.estado||null,
    fecha_creacion:t.fechaCreacion?String(t.fechaCreacion).split('T')[0]:null,
    fecha_term:t.fechaTerm?String(t.fechaTerm).split('T')[0]:null,
    fecha_cumpl:t.fechaCumpl?String(t.fechaCumpl).split('T')[0]:null
  };}
  /* Para ACTUALIZAR tareas ya existentes: igual que toRow pero sin fecha_term.
     La fecha de término es de captura manual (modal Editar tarea); las cargas/
     recargas de Excel nunca deben pisarla, solo se sincronizan el resto de
     campos (estado, fecha_cumpl, etc.) con lo que trae el archivo. */
  function toRowActualizar(t){var row=toRow(t);delete row.fecha_term;return row;}

  try{
    /* INSERT nuevas — con upsert por tarea_id: si alguna ya existiera en
       Supabase (o viniera repetida en el archivo), se actualiza en vez de
       chocar con la restricción única tareas_tarea_id_unique. */
    if(nuevas.length){
      var rows=nuevas.map(toRow);
      /* dedup por tarea_id dentro del mismo lote (conserva la última) */
      var _seenIds={},_rows2=[];
      for(var z=rows.length-1;z>=0;z--){
        if(_seenIds[rows[z].tarea_key])continue;
        _seenIds[rows[z].tarea_key]=1;_rows2.unshift(rows[z]);
      }
      rows=_rows2;
      for(var i=0;i<rows.length;i+=500){
        var r=await client.from('tareas').upsert(rows.slice(i,i+500),{onConflict:'tarea_key',ignoreDuplicates:false});
        if(r.error){err='Insert: '+r.error.message;break;}
        insOk+=Math.min(500,rows.length-i);
      }
    }
    /* UPDATE las que cambiaron (con permiso) */
    if(!err&&actualizar.length){
      for(var j=0;j<actualizar.length;j++){
        var row=toRowActualizar(actualizar[j]);
        var ru=await client.from('tareas').update(row).eq('tarea_key',row.tarea_key);
        if(ru.error){err='Update: '+ru.error.message;break;}
        updOk++;
      }
    }
    /* Auditorías — dedup por CENTRO (razón+centro+fecha+clase), independiente del
       nombre de tienda: actualiza en vez de duplicar cuando el nombre varía. */
    if(!err&&audStaged.length){
      var arows=audStaged.filter(Boolean).map(function(a){return{
        razon:a.razon||null,centro:a.centro||null,tienda:a.tienda||null,
        fecha:a.fecha?String(a.fecha).split('T')[0]:null,mes:a.mes||null,
        pct_cumpl:parseFloat(a.pctCumpl)||0,tareas:parseInt(a.tareas)||0,
        pendientes:parseInt(a.pendientes)||0,resueltas:parseInt(a.resueltas)||0,
        pct_resuelto:parseFloat(a.pctResuelto)||0,clase:a.clase||null
      };});
      var _kAud=function(razon,centro,fecha,clase){
        return [norm(razon),canonCentro(centro),String(fecha||'').split('T')[0],norm(clase)].join('|');
      };
      /* dedup dentro del mismo archivo (conserva la última ocurrencia) */
      var _seen={},arows2=[];
      for(var z=arows.length-1;z>=0;z--){
        var kk=_kAud(arows[z].razon,arows[z].centro,arows[z].fecha,arows[z].clase);
        if(_seen[kk])continue;_seen[kk]=1;arows2.unshift(arows[z]);
      }
      arows=arows2;
      /* traer existentes para emparejar por centro (no por nombre) */
      var exAud=[];
      try{var re=await client.from('auditorias').select('id,razon,centro,fecha,clase').limit(20000);if(!re.error)exAud=re.data||[];}catch(_e){}
      var exMap={};
      exAud.forEach(function(x){exMap[_kAud(x.razon,x.centro,x.fecha,x.clase)]=x.id;});
      for(var ai=0;ai<arows.length&&!err;ai++){
        var ar=arows[ai];var k=_kAud(ar.razon,ar.centro,ar.fecha,ar.clase);
        if(exMap[k]){
          var ru2=await client.from('auditorias').update(ar).eq('id',exMap[k]);
          if(ru2.error)err='Aud upd: '+ru2.error.message;else audOk++;
        }else{
          var ri2=await client.from('auditorias').insert([ar]).select('id');
          if(ri2.error)err='Aud ins: '+ri2.error.message;else{audOk++;if(ri2.data&&ri2.data[0])exMap[k]=ri2.data[0].id;}
        }
      }
    }
  }catch(e){err=e.message;}

  clearStaged();
  if(err){toast('⚠ '+err);return;}

  /* ── Snapshot para poder revertir esta carga ──
     STORE.* aún tiene el estado PREVIO (loadDataFromSupabase corre después). */
  var snap=null;
  try{
    var tNew=(nuevas||[]).map(function(t){return tareaKey(t.id,t.razon);});
    var tPrev=(prevTareas||[]).map(function(rw){var cc=Object.assign({},rw);delete cc.updated_at;return cc;});
    var aNew=[],aPrev=[];
    if((audStaged||[]).length && typeof arows!=='undefined' && arows){
      arows.forEach(function(ar){
        var prev=(STORE.auditorias||[]).find(function(x){
          return norm(x.razon)===norm(ar.razon)&&canonCentro(x.centro)===canonCentro(ar.centro)&&
                 String(x.fecha||'').split('T')[0]===ar.fecha&&norm(x.clase)===norm(ar.clase);
        });
        if(prev){
          aPrev.push({razon:prev.razon||null,centro:prev.centro||null,tienda:prev.tienda||null,
            fecha:prev.fecha?String(prev.fecha).split('T')[0]:null,mes:prev.mes||null,
            pct_cumpl:parseFloat(prev.pctCumpl)||0,tareas:parseInt(prev.tareas)||0,
            pendientes:parseInt(prev.pendientes)||0,resueltas:parseInt(prev.resueltas)||0,
            pct_resuelto:parseFloat(prev.pctResuelto)||0,clase:prev.clase||null});
        }else{
          aNew.push({razon:ar.razon,centro:ar.centro,tienda:ar.tienda,fecha:ar.fecha});
        }
      });
    }
    if(tNew.length||tPrev.length||aNew.length||aPrev.length)
      snap={t_new:tNew,t_prev:tPrev,a_new:aNew,a_prev:aPrev};
  }catch(e){console.warn('snapshot:',e);}

  /* Registrar esta carga (con respaldo) para el seguimiento de desempeño */
  await registrarCargaExcel(insOk+updOk, audOk, snap);

  /* Recargar datos frescos desde Supabase */
  await loadDataFromSupabase();
  await loadCargas();
  if(VIEW==='desempeno')renderDesempeno();
  toast('✓ '+insOk+' nuevas · '+updOk+' actualizadas · '+omitidas+' omitidas (idénticas)');
}

/* Helper: obtener cliente Supabase */
function getSbClient(){
  if(_sb)return _sb;
  try{return supabase.createClient(SB_URL,SB_KEY,{
    auth:{persistSession:false,autoRefreshToken:false,detectSessionInUrl:false},
    realtime:{enabled:false},
    global:{headers:{'x-client-info':'monitor-cumplimiento'}}
  });}catch(e){return null;}
}
/* ════════════════════════════════════════════════════════════════════
   FILTROS
════════════════════════════════════════════════════════════════════ */
function uniq(arr){return [...new Set(arr.filter(Boolean))].sort();}
function fillFilters(){patchedFillFilters();}
/* Detectar si un valor parece cifrado (base64:base64 largo) */
function pareceCifrado(v){
  if(!v||typeof v!=='string')return false;
  return /^[A-Za-z0-9+/=]+:[A-Za-z0-9+/=]+$/.test(v)&&v.length>30;
}
function limpiarOpciones(arr){
  return arr.filter(function(v){return v&&!pareceCifrado(v);});
}
function fillCentroTienda(){
  const razon=document.getElementById('f-razon').value;
  const filtA=STORE.tareas.filter(t=>razon==='ALL'||t.razon===razon);
  const filtAud=STORE.auditorias.filter(a=>razon==='ALL'||a.razon===razon);
  const centros=limpiarOpciones(uniq([...filtA.map(t=>t.centro),...filtAud.map(a=>a.centro)]));
  const tiendas=limpiarOpciones(uniq([...filtA.map(t=>t.tienda),...filtAud.map(a=>a.tienda)]));
  const fc=document.getElementById('f-centro'),ft=document.getElementById('f-tienda');
  const cv=fc.value, tv=ft.value;
  fc.innerHTML='<option value="ALL">Todos</option>'+centros.map(c=>`<option>${c}</option>`).join('');
  ft.innerHTML='<option value="ALL">Todas</option>'+tiendas.map(t=>`<option>${t}</option>`).join('');
  if([...fc.options].some(o=>o.value===cv))fc.value=cv;
  if([...ft.options].some(o=>o.value===tv))ft.value=tv;
}
function setQuick(q,el){
  document.querySelectorAll('#quick-row .qbtn').forEach(b=>b.classList.remove('active'));
  el.classList.add('active');
  const today=new Date(),y=today.getFullYear(),m=today.getMonth();
  /* Los rangos miran hacia atrás: terminan HOY (nunca al futuro) y arrancan
     N meses antes. "Mes" = mes en curso; "Trimestre" = últimos 3 meses; etc. */
  let desde=null,hasta=new Date(today);
  if(q==='all'){desde=null;hasta=null;}
  else if(q==='month')desde=new Date(y,m,1);
  else if(q==='quarter')desde=new Date(y,m-2,1);   /* últimos 3 meses */
  else if(q==='half')desde=new Date(y,m-5,1);      /* últimos 6 meses */
  else if(q==='year')desde=new Date(y,m-11,1);     /* últimos 12 meses */
  document.getElementById('f-desde').value=fmtInput(desde);
  document.getElementById('f-hasta').value=fmtInput(hasta);
  /* Al usar rango rápido, resetear filtro de mes a "Todos" */
  const g=document.getElementById('f-gran');
  g.value='ALL';
  render();
}
function onFilter(src){
  if(src==='manual')document.querySelectorAll('#quick-row .qbtn').forEach(b=>b.classList.remove('active'));
  fillCentroTienda();render();
}
function getFilterState(){
  return{
    razon:document.getElementById('f-razon').value,
    centro:document.getElementById('f-centro').value,
    tienda:document.getElementById('f-tienda').value,
    tipo:document.getElementById('f-tipo').value,
    desde:document.getElementById('f-desde').value?new Date(document.getElementById('f-desde').value+'T00:00:00'):null,
    hasta:document.getElementById('f-hasta').value?new Date(document.getElementById('f-hasta').value+'T23:59:59'):null,
    gran:'month',
    mesGran:document.getElementById('f-gran').value /* 'ALL' o '01'-'12' */
  };
}
function matchTipo(t,tipo){
  if(tipo==='ALL')return true;
  return tipoNorm(t.tipoTarea)===tipo;
}
function inRange(d,f){
  if(!d)return (!f.desde&&!f.hasta);
  if(f.desde&&d<f.desde)return false;
  if(f.hasta&&d>f.hasta)return false;
  return true;
}
function filteredTareas(){
  const f=getFilterState();
  const anio=new Date().getFullYear();
  return STORE.tareas.filter(t=>{
    if(f.razon!=='ALL'&&t.razon!==f.razon)return false;
    if(f.centro!=='ALL'&&t.centro!==f.centro)return false;
    if(f.tienda!=='ALL'&&t.tienda!==f.tienda)return false;
    if(!matchTipo(t,f.tipo))return false;
    if(f.desde||f.hasta){if(!inRange(fromISO(t.fechaCreacion),f))return false;}
    /* Filtro por mes seleccionado */
    if(f.mesGran&&f.mesGran!=='ALL'){
      const cr=fromISO(t.fechaCreacion);
      if(!cr)return false;
      if(cr.getMonth()+1!==parseInt(f.mesGran)||cr.getFullYear()!==anio)return false;
    }
    return true;
  });
}
function filteredAuditorias(){
  const f=getFilterState();
  const anio=new Date().getFullYear();
  /* Fecha de la auditoría con fallback: si `a.fecha` no se parsea (dato común
     en la base), se deriva del nombre del mes `a.mes`. Sin esto, al activar el
     filtro de mes o un rango de fechas se perdían TODAS las auditorías y tanto
     el dashboard como el PPTX de Cumplimiento salían vacíos. */
  const audFecha=a=>{
    const d=fromISO(a.fecha); if(d)return d;
    const mi=(typeof mesIndexFromNombre==='function')?mesIndexFromNombre(a.mes):-1;
    return mi>=0?new Date(anio,mi,15):null;
  };
  return STORE.auditorias.filter(a=>{
    if(f.razon!=='ALL'&&a.razon!==f.razon)return false;
    if(f.centro!=='ALL'&&a.centro!==f.centro)return false;
    if(f.tienda!=='ALL'&&a.tienda!==f.tienda)return false;
    const fa=audFecha(a);
    if(f.desde||f.hasta){if(!inRange(fa,f))return false;}
    /* Filtro por mes seleccionado (igual que tareas) */
    if(f.mesGran&&f.mesGran!=='ALL'){
      if(!fa)return false;
      if(fa.getMonth()+1!==parseInt(f.mesGran)||fa.getFullYear()!==anio)return false;
    }
    return true;
  });
}

/* ════════════════════════════════════════════════════════════════════
   CLASIFICACIÓN DE TAREAS
   fechaCumpl es la señal autoritativa: si una tarea ya tiene fecha de
   cumplimiento capturada, se considera resuelta (no pendiente) sin
   importar si el texto de "estado" quedó desactualizado. Esto mantiene
   sincronizados Auditorías, Cartera, Dashboard y Tareas entre sí.
════════════════════════════════════════════════════════════════════ */
function esPendiente(t){if(t.fechaCumpl)return false;return norm(t.estado).includes('abierta');}
function esResuelta(t){if(t.fechaCumpl)return true;return norm(t.estado).includes('resuelta');}
/* Vencida = pendiente y su fechaTerm ya pasó — calculado por FECHA, no por
   el texto "atrasada" guardado en estado. El texto solo se corrige cuando
   corre actualizarEstadosVencidos() (se dispara al abrir el dashboard o
   Auditorías); si el usuario abre el generador de PPT directamente en esa
   sesión sin haber pasado antes por esas pantallas, el texto podía seguir
   diciendo "Abierta" aunque la fecha ya hubiera vencido, y el reporte
   pintaba esa tarea como vigente. Esta función no depende de ese texto. */
function tareaVencidaPorFecha(t){
  if(esResuelta(t))return false;
  var ft=fromISO(t.fechaTerm); if(!ft)return false;
  var hoy=new Date(); hoy.setHours(0,0,0,0);
  var ftD=new Date(ft); ftD.setHours(0,0,0,0);
  return ftD<hoy;
}
/* Deriva el estado correcto SOLO a partir de las fechas, en vez de depender de
   que alguien elija bien "Resuelta"/"Resuelta Atrasada"/"Abierta"/"Abierta
   atrasada" a mano en un campo aparte:
   - Si ya tiene fecha de cumplimiento → Resuelta (a tiempo) o Resuelta Atrasada
     (si esa fecha quedó después de la fecha de término).
   - Si no tiene fecha de cumplimiento → Abierta, o Abierta atrasada si la
     fecha de término ya pasó.
   Esta es la misma comparación que usa esPendiente/esResuelta; aquí se usa
   para fijar el TEXTO del campo estado, no solo para el cálculo funcional. */
function estadoAutomatico(fechaTerm, fechaCumpl){
  var ft=fromISO(fechaTerm);
  if(fechaCumpl){
    var fc=fromISO(fechaCumpl);
    if(!ft||!fc)return 'Resuelta';
    var ftD=new Date(ft);ftD.setHours(0,0,0,0);
    var fcD=new Date(fc);fcD.setHours(0,0,0,0);
    return fcD>ftD?'Resuelta Atrasada':'Resuelta';
  }
  if(!ft)return 'Abierta';
  var hoy=new Date();hoy.setHours(0,0,0,0);
  var ftD2=new Date(ft);ftD2.setHours(0,0,0,0);
  return ftD2<hoy?'Abierta atrasada':'Abierta';
}
/* Resuelta atrasada = ya está resuelta pero su fecha de cumplimiento quedó
   después de su fecha de término — calculado por FECHA, igual que
   tareaVencidaPorFecha hace con las abiertas. Antes esto se leía solo del
   texto guardado en "estado" ("Resuelta Atrasada"), que se queda
   desactualizado cuando una tarea se vuelve a importar/corrige y su fecha
   de cumplimiento cambia pero el texto del estado no se reescribe: la
   auditoría se pintaba en amarillo (🟡) para tareas que, por fecha, ya
   estaban resueltas a tiempo. Si no hay fechas suficientes para decidir,
   se usa el texto como respaldo (mismo comportamiento de antes). */
function tareaResueltaAtrasadaReal(t){
  if(!esResuelta(t))return false;
  var fc=fromISO(t.fechaCumpl), ft=fromISO(t.fechaTerm);
  if(fc&&ft){
    var fcD=new Date(fc);fcD.setHours(0,0,0,0);
    var ftD=new Date(ft);ftD.setHours(0,0,0,0);
    return fcD>ftD;
  }
  return norm(t.estado).includes('resuelta')&&norm(t.estado).includes('atrasad');
}
function diasVenc(t){ // días hasta fechaTerm (negativo = vencida)
  const ft=fromISO(t.fechaTerm);if(!ft)return null;
  return daysBetween(ft,new Date());
}
function estadoBadge(estado){
  const n=norm(estado);
  if(n.includes('resuelta')&&n.includes('atrasad'))return `<span class="badge b-orange">Resuelta atrasada</span>`;
  if(n.includes('resuelta'))return `<span class="badge b-green">✓ Resuelta</span>`;
  if(n.includes('abierta')&&n.includes('atrasad'))return `<span class="badge b-red">🔴 Abierta atrasada</span>`;
  if(n.includes('abierta'))return `<span class="badge b-blue">Abierta</span>`;
  return `<span class="badge b-gray">${estado||'—'}</span>`;
}

/* ════════════════════════════════════════════════════════════════════
   RENDER PRINCIPAL
════════════════════════════════════════════════════════════════════ */

/* Actualizar automáticamente tareas "Abierta" cuya fechaTerm ya venció → "Abierta atrasada" */
var _estadoSyncEnVuelo={};
function actualizarEstadosVencidos(){
  var actualizadas=0;
  var client=getSbClient();
  STORE.tareas.forEach(function(t){
    var correcto=estadoAutomatico(t.fechaTerm,t.fechaCumpl);
    if(norm(t.estado)===norm(correcto))return; /* ya coincide, nada que hacer */
    t.estado=correcto;
    actualizadas++;
    /* Persistir la corrección en Supabase — de lo contrario la próxima carga
       trae de vuelta el estado viejo desde la base de datos. Guarda
       anti-duplicados: una sola escritura en curso por tarea a la vez. */
    if(client&&t.id){
      var tk=tareaKey(t.id,t.razon);
      if(!_estadoSyncEnVuelo[tk]){
        _estadoSyncEnVuelo[tk]=true;
        client.from('tareas').update({estado:correcto}).eq('tarea_key',tk)
          .then(function(r){delete _estadoSyncEnVuelo[tk];if(r.error)console.warn('sync estado:',r.error.message);})
          .catch(function(e){delete _estadoSyncEnVuelo[tk];console.warn('sync estado:',e.message);});
      }
    }
  });
  if(actualizadas>0)console.log('[AUDITORIA] '+actualizadas+' tarea(s) con estado corregido automáticamente según sus fechas');
}

/* Plantillas oficiales (archivos reales, incrustados). El usuario elige cuál
   descargar por su nombre. */
var PLANTILLA_MASTER_B64='UEsDBBQAAAAIAMl4+FxGx01IlQAAAM0AAAAQAAAAZG9jUHJvcHMvYXBwLnhtbE3PTQvCMAwG4L9SdreZih6kDkQ9ip68zy51hbYpbYT67+0EP255ecgboi6JIia2mEXxLuRtMzLHDUDWI/o+y8qhiqHke64x3YGMsRoPpB8eA8OibdeAhTEMOMzit7Dp1C5GZ3XPlkJ3sjpRJsPiWDQ6sScfq9wcChDneiU+ixNLOZcrBf+LU8sVU57mym/8ZAW/B7oXUEsDBBQAAAAIAMl4+FzFWh7WPwEAANcCAAARAAAAZG9jUHJvcHMvY29yZS54bWzNkktOwzAQhq9SZZ/YSWkRVmqJh8qGSkgFgdhZ9rS1iB+yHdIeilNwMRzTplSwYYfkzcz8882v8dTcEm4c3DtjwQUJfrRVjfaE21m2CcEShDzfgGK+iAodiyvjFAsxdGtkGX9la0AVxlOkIDDBAkM9MLcDMaO14IQ7YMG4PV7wAW9b1ySY4AgaUKCDR2VRooz2E+1u29ToCEiwhul1Gwf/iQY6f1wm1KG9ZwVwyn/BQQy8lP0Vmioo2yu3Xg6qruuKbpx0cR8lel7cLdPqcql9YJpD7PKShJ2FWXaY/DS+vnmYZ7TC1TTH5/E9lBNSYTKZvPRmT/wdDSsj5Er+C8fVWe8YT0l58c3xwSCt44k1zIfFPnG1o3NwmmlhRrctOAfOFKPLVsj4ux/vrEY/G3qGgzfppdEUJ8UQpuj0huknUEsDBBQAAAAIAMl4+Fy2UZiG2QIAACwMAAATAAAAeGwvdGhlbWUvdGhlbWUxLnhtbM2WW2/bIBiG7yftPyDuW3xMnahO1bixdjFp0tr9AILxocXYMqyHfz8MiQ91miVaKi0Xjvn8+gFe+D5zffNaMvBMG1FUPIT2pQUB5aRKCp6F8NdDfBFAICTmCWYVpyF8owLeLL9+ucYLmdOSAvU+FwscwlzKeoGQICqMxWVVU66epVVTYqmaTYaSBr8obsmQY1kzVOKCQ8BxqbA/0rQgFDy0SLjcwddMXbgUbYCw5p7oHodvaG3yZLd/osk2EWvAM2YhtPQPouU16gRMTnWx/m11W0Hy5Ex0duzNr+46nmN4U916vY7WdsfTAkyImsW0by8O7NWOORCZ2yk7snzLG+sHfHein69WK38+0ru93pvoA2vm3Tojvdfr/en4V7dRNBvp/V4/m3p9NZ95Y70W5azgT3tXsFuZTpJW7NteeaDkwW7BexUa7BzzPpcf7aMSP1ZNrAR6cbEsOJBvNU0xUboIl5umwG0HeEHx4IkJEfEuhN4By4IforNC4U+j90A0nJieZjlsFIzdyzdGvwvduahYkcQqqBta1tlY5+p228FIlzW4vxdbUiZAXQm1VvBDlE7mgksT8wdZ2cl1KxNDoNsKj4W6V8dBbVMTjqTa/iEqGrigdhrAbb20Z47pAgiCGU1UxCyfLBj9SYk0tJGV/2CryHFCt77ax1kQ/N2BAXXuns/YIdY7g7PWYWfRdNszPm6BFzUU3/EhILgOYapyUt2WtYILnkGAWaY+iUSa8deNkHdY5GYKOjV2VZ73PMf32kGeD+gG9nmA6L0BNE2Vbx9E+qZ6ZiB7n55fjPaNbJPF/08N846sYd4p1cbbVZtxpsw/JQGdgzMYJmCNZQ7ai9pmRUOY+VC2mfVQtWkHupoOZAgvTHkBTRfcqLEFg95a1OcVxN7O4Mg1OtE495OM8/f45p9gG5qmCBodAtCeE3W1eVSoO3Wo+M2kMCePV9ngaHde6hJUv7r8A1BLAwQUAAAACADJePhc3pf6HJMIAABVGgAAGAAAAHhsL3dvcmtzaGVldHMvc2hlZXQxLnhtbK1Za2/bOBb9K4SLLVJMJrYeTt04DeA4acY7aVPYyexiv9ESY7MjiS4p5fXr91xKFt00ofphCjSWpSvy8D7OPaSP75X+26yFKNlDnhXmY29dlpujft8ka5Fzc6A2osCTW6VzXuKrXvXNRgue2pfyrB8OBof9nMuid3Js733VJ8eqKjNZiK+amSrPuX48FZm6/9gLetsbc7lal3Sjf3K84SuxEOXN5qvGt347SipzURipCqbF7cfeaXh0Go/oBWvxlxT3ZueambW6v9AyvcTMWMigx56UyhcJz8QXgp9hugHu0pKXSv1NL81SMsQoIhNJSVNxfNyJqchgfhrB3HyvZ8d1i45e3b3e4vhk3YRlL7kRU5X9R6blmqbtsVTc8ior3c3RweFhPDgMh+2zubr/QzR+iQ9imi5RmbF/2X39VtxjSWVKlW+H7rFcFvUnf2gcumNvV/ziG2HzRmiXVU9kF3HGS35yrNU909aK8ITBdpQWIbyWkMUpTODuAB+4LQsK/KLUeCwxYnny9k0QjsJBNGZfM16UMss4FswSrlecvX0zCoNwzCZVKkul374Jo/djbvbZhMIgU54K+vINUwvDHtlnAQeb434JyDR8P8F/QG3xRi2sqIYVvQLrz0qrlDNkRinY3p9frt4BTDACzObJQhVKczYVRakVLBbTVywWlabHk3ceWMMW1rCGFb/qrQ8ROeQP9Y2zyc3Z7PpqDq8MhuPJwjP+YTv+YT3+8JXxFwIVuBSIgEEQRMZyjB5H47TK1PMoUHSGH+IxW6qytipYDwBHo2DMLkQhNNesz6YqX8qC6x7bU/Qc4Q7orSlFWLPzh0RkvXcHHvTvW/TvO9BLfSfYhsPrk4v5+cVkzvgzzEUl7jilirWaXs1hNpv3J9Prm8nl7H94I8Pj75Vgj5yJB4nEKo7gCA+8UQtv5IdnaLQcw+aYXCCACUea3Epk/EbpbSr9hiw2+DsFEIHPTyJZu0ow0gEDVOKjimfyie97EH5oEX7oQsgKtTN6Air3RSYYtCMTifmGblaHpV2dXs4uJkjc2RXikAqk250opKayx5xrSu3etRRFikjcIXLhcJyh1k1vTCZ8XQeTGUpTLe84q419OB0dBYEfJ6/AYvWkpUw4WgwRALyy5olgtzwr4X6TaLmUGmhrj9EFX2Jt4gEB4d5sBhO2WEI/lmukCbHdVyxPEhBcz4WpRFYqyuF/td+YQW+qCnbHM2TS7MtsOptckrfQnIQWRSI5srgASYNjlQ+dY8gg8qOzacILVlqUNWNTyGxlCcNZacOCFM5tQieU0PtEKttC0ALdN6kyMqcVFfD7cDDOhVY+Egc7thjjjoRW4K1HlmHwnLoLWMmVTKpMW1c/BjhRWosVAgwLTm8qYsRNRZkXjY03vI7Lg2FHSdgKtwVghG5yp064quB1ORgW4SKDd8umEtRKHf1ExbCFdOBLNJxENmTsY4TAdYSgoyX8OBW70iny6JFdynwjBXiHPW8MxOwlyN/rJcfpQUPqdnqSiT8bO4YNRr/QH6Ntf5xez/6anU3Ozn29MXDsGHTQo787OknySmckoRN+eNb5fF4KHb2GHfR6SUmbVTnyZs6ftpP+wLYT22Wo3RwxCJp9Bs2C4EGaHLCFhDGk137TU+qV+KDtCLwORkUzBQso1DkqqyaHjG2EBmVxJLYV1YkwCnRgSi2LlUwVqxlEu6XsEfhUGkE3gjEVJKpmVWkC6iXc0BFu2EG4C8nKCiSJGuS2z5qm6lgmc1lSu+a2NsEru+j22QaLhLcVrNDJspeiceTD6Gg37KJdx58UKyNXmOXFrkWu/QEm+axdn9dljmHDDoadWACI2EdwaL7UwtJVI7vYnkFnx6di3/S7MeP5EoyKbm+Zje9UDEWz2vWoD52j2LCDYg3yGkQFFY+C+17JmsDr6df8kSTYpqw04d92AnjsThq0d9IcKYBhi5uKRpzMd0Jum0JmR1Da601HduH7X+CveMtf/75ZXHu5K3TMGHaITz93NTuof4q3HKWGHZR6DgkBBbGEb79csbpBb0uH5IurnmepbDtnWiV1vrWp12hBH7rIsWrUwao3xpJVM/r5f9FRrrplKtvbrqB+/K5hvTswiPYBc5wadXMqASORYkhWE8dCvFRouw2+HVdZNQNpgSqUhc3tTHgd5Agz6iDMs1aBoKYRPZTaEWuI6EHmitTLjkmpLbZbu6WxmgXIrEDbuVknoxfgzjY+6lIPkeOyKP6F6hs21ff5fP7Zu6mOHA1FHTTkLb760OKfqr3IKbuoQ9ldc4QrUf6qkyuIZbshbkhin1WvVEXZtQ2LHBFGHVv5ziLz+sDRYtRBi5cKW5D+JyteMXYhaJ9kt3eC9ezDHsqmt2vQI91R71Sa2jIb1cE5seOcePALSXjYJOFLS/dN4xgk7mCQ6e6G4vlOMbGb9f1tiNWtxO3MxRgVu0vFyj61egSJ4sPnmCXuYBYIxBJpmYgDNq0n2mASJCtpRkpJKyLarK2PF/q1vyyL/HRk+OyU0BsuxxlxB2fMxYrExMqed2XCGx1HGHEHYZAYCZEEXxT4gDKMFRIRquNh7H4WoeJLUe9i62qxJ0lBfXqa1sXjXaRjiriDKbZwLrFSy9OmFrwQVxohAM+naT/P+xz/UC708Xue/56m3vkdG8QdbLDjjlR8gwNoqXbZ+JZvkBa5eEoye/xAqjfldJygBQXkiC3r/NGID6lOOkep9eZSenVb7Ggk7qCRLcCLiuuU1y45eMjMg5vOno9oQBXFzhnIi/P3dw7a6dePz3hVFgi7uMXsg4P3gKHrc/b6S6k2tgeid2AfYC/XArpak8EwCEZBMAijwzAckKdvlSpffuR+bak2TGk6c+L0s8fHHklozWXZg5ZBaS7kk6BzRSSZlk8KZtnZRpK2ArvdoXSxCXF36BeE9pekk/8DUEsDBBQAAAAIAMl4+Fzi3u04lhMAAEOgAAAYAAAAeGwvd29ya3NoZWV0cy9zaGVldDIueG1svZ1dcxrZuYX/CqVUzkXKZdHdCJAjq0oj2RnIzNglT5Jzzh2WWhY1iFagFWXm16eRgL32yl6LXbnIxfiDR/3QaN5Fl7zovc+em9Uv6/u6bnv/fFgs1++P7tv28d3x8frmvn6Yrd82j/WyI3fN6mHWdn9dfTteP67q2e3LQQ+L47LfHx4/zObLo/Ozl8c+r87Pmqd2MV/Wn1e99dPDw2z163f1onl+f1Qc7R64nn+7bzcPHJ+fPc6+1V/q9i+Pn1fd3473ltv5Q71cz5tlb1XfvT+6KN5NB+PB5oiXL/nrvH5ew5976/vm+U+r+e0P3VN3r6R/1Nu8uq9N88sGT243D22ebVn3fv3yuJi/PH+vbR5/qO/ay3qx6J6jPOrNbtr5P+rP3Ze9P/ratG3zsOHdmbeztnvobtX8Vi9fzqJe1N3Xduf3+G9f/CoJ0vXfX19D+e5/P16V1dH+hW5ODf+8e0UfX77j3Xfw62xdXzaLv81v2/vuhLtXdVvfzZ4WbXhw/HY8Ho6K8ehkD6+b5+/r7fd48Pblm3bTLNYvv/aeXw8bva2qsl8VZXfY13rdftx+R26e1t3r2D3hUe9hvnz9ffbP7f8xkBSlOKDcHlDSAZU6oNoeUPEzVG+Hw0F/mHWag61kQJJy/Lbon1ajHMfJ1nHCJzIUBwy3Bwz5gL44YLQ9YJT7zRzvvvv9l8l5/V/5MidXs3Z2frZqnnurly/f/B+v9s+7n4Fu8G82X3GxieDrS+keni83Mf3Srjo874zt+fXst//5XTmo/rj8w9lx2z3V5uHjm+3R3x04+rJetqsmdeTl9siROPLneb28nSUOvDrwlD/W697F0+28bVbdiVejP85ST//h0IkvupSlDvx44Lx/37t8eujeSx6682+bhOBPh174rHs3XScO/P7AgZ+7b9jmSevUwZMDB1/X66d60TapY6cHX/Lu6Pjg424I95NY7gbuY/lqG7/YNleJ8K3ZotN/R99rNNFoukVFhKLzqvbnVenzqvR5aTTRaFodPK/B/rwG+rwG+rw0mmg0HRw8r5P9eZ3o8zrR56XRRKPpycHzGu7Pa6jPa6jPS6OJRtPhwfMa7c9rpM9rpM9Lo4lG09HB8xrvz2usz2usz0ujiUbT8cHzOt2f16k+r1N9XhpNNJqeHjyvor8/sc1VXJ3ZjqVOzbCJYdMdc2dXhLMrzNkV5uw0mxg23TF3duHNvzDv/oV5+zdsYti0OHwFKMIloDDXgMJcBAybGDYtDl8HinAhKMyVoDCXAsMmhk2Lw1eDIlwOCnM9KMwFwbCJYdPi8DWhCBeFwlwVCnNZMGxi2LQ4fGUowqWhMNeGwlwcDJsYNi0OXx+KcIEozBWiMJcIwyaGTYvDV4kiXCYKc50o9Fv+94ZNDJvumDm7MlwrSnOtKM21wrCJYdPy8LWiDNeK0lwrSnOtMGxi2LQ8fK0o4QcF95OC+1HB/azgflg4fK0ow7WiTL53b02H39fL8L5eJt9nt6bD78FleA8uk++JW9Ph98syvF+Wyfevrenwe1sZ3tvK5HvN1nT4fagM70Nl8n1hazr8nlGG94wymeGt6XC+q5DvKpm37Q95h7NYhSxWyWxsTYdzU4XcVMk53poyfiKGH4nNjKdZbAozXpkZT7PYFGa8MjOeZrEpzHhlZjzNYlOY8crMeJrFpjDjlZnxNItNYcYrM+NpFv/7Q5jxgZnxNItNYcYHZsbTLDaFGR+YGU+z2BRmfGBmPM1iE/xDjZnxNItNYcYHZsbTLDaFGR+YGU+z2BRmfGBmPM1iU5jxgZnxNItNYcYHZsbTLP63rDDjJ2bG0yw2hRk/MTOeZrEpzPiJmfE0i01hxk/MjKdZbAozfmJmPM1iE/zzoZnxNItNYcZPzIynWWwKM35iZjzNYlOY8RMz42kWm8KMn5gZT7P430XDjA/NjKdZbAozPjQznmaxKcz40Mx4msWmMONDM+NpFpvCjA/NjKdZbAozPjQznmaxCf5R28x4msWmMONDM+NpFpvCjA/NjKdZbAozPjQznmbxv7GHGR+ZGU+z2BRmfGRmPM1iU5jxkZnxNItNYcZHZsbTLDaFGR+ZGU+z2BRmfGRmPM1iU5jxkZnxNItNULWYGU+z2BRmfGRmPM1iU5jxkZnxNIv7mjDjYzPjaRabwoyPzYynWWwKMz42M55msSnM+NjMeJrFpjDjYzPjaRabwoyPzYynWWwKMz42M55msSnM+NjMeJrFJigAzYynWWwKMz42M55mcfcXZvzUzHiaxaYw46dmxtMsNoUZPzUznmaxKcz4qZnxNItNYcZPzYynWWwKM35qZjzNYlOY8VMz42kWm8KMn5oZT7PYFGb81Mx4msUmqKXNjKcZ9chYJJspF5Bk0Pv2zaALSDKoaftm1gUkGbSqfTPuApIMStC+mXgBSQadZd8MvYAkg4qxb+ZeQJJBI9g3oy8gyaDA65vpF5Bk0Lf1TQAEpM8qQALSn23YyrI++ICffHAJSEOSQQLSnxzYyTISgJ8rSBf9O1lGAvBjAOlefifLSAC29ukafSfLSACW7OnWeyfLSAB24umSeifLSABW2OlOeSfLSAA2zukKeCfLSAAUxEW6sd3KMurcAvrcIl2w7mQZCSjxozouARllaQFtaeHqUgFJBglwjamAJIMEuNJUQJJBAlxvKiDJIAGuOhWQZJAA154KSDJIgCtQBaTPXEECXIcqIMkgAa5GFZBkkADXpApIMvxsmUtARptaQJ1auD5VQJJBAlylKiDJIAGuVRWQZJAAV6wKSDJIgOtWBSQZJMDVqwLS5/ogAa5hFZBkkABXsgpIMkiA61kFJBkkwFWtApIMPwzpEpBRtxbQtxaucBWQZJAA17kKSDJIgKtdBSQZJMA1rwKSDBLgylcBYxnUr4XrXwUkGSTAVbACkgwS4FpYAUkGCXBFrIAkgwS4LlZAkuGnd10CMvrYAgrZwjWyApIMEuBKWQFJBglwvayAJIMEuGpWwFgG5Wzh2lkBSQYJcAWtgCSDBLiOVkCSQQJcTSsgySABrqkVkGSQAFfWCkgy/Li5S0BGYVtAY1u4ylZAkkECXGsrIMkgAa64FTCWQXVbuO5WQJJBAlx9KyDJIAGuwRWQZJAAV+IKSDJIgOtxBSQZJMBVuQKSDBLg2lwBSYb3R7gEZDS6BVS6het0BSQZJMDVugLGMih2C9fsCkgySIArdwUkGSTA9bsCkgwS4CpeAUkGCXAtr4AkgwS4oldAkkECXNcrIMkgAa7uFZBkeEOPS0BG5VtA51u40lfAWAa1b+F6XwFJBglw1a+AJIMEuPZXQJJBAlwBLCDJIAGuAxaQZJAAVwMLSDJIgGuCBSQZJMCVwQKSDBLg+mABSYZ3oLkEZHTCZR9vGDMJEJBkcH+X64QFJBncjuU6YQFJBndPuU5YQJLBDVSuExaQZHAPleuEBSQZ3EblOmEBSQZ3UrlOWECSwc1UrhMWkGRwP5XrhAWMZdAJl64TFpBkeIejS0AakgwS4DphAUkGCXCdsIAkgwS4TlhAkkECXCcsIMkgAa4TFpBkkADXCQtIMkiA64QFJBkkwHXCAsYyvGnYdcICkgwS4DphAUmGt+S6BOTcQIt30NpbaHPuocWbaO1dtDm30eJ9tPZG2pw7afFWWnsvbc7NtHg3rb2dNud+Wryh1t5Rm3NLLd5Ta2+qzeiES+iES9cJC0gySIDrhAUkGSTAdcICkgzvIbc3kWckADrh0nXCApIMEuA6YQFJBglwnbCAJIMEuE5YQJJBAlwnLCDJIAGuExYwlkEnXLpOWECSQQJcJywgySABrhMWkGSQANcJC0gyXEjBrqSQkQDohEvXCQtIMkiA64QFJBkkwHXCApIMEuA6YQFJBglwnbCAsQw64dJ1wgKSDBLgOmEBSQYJcJ2wgCSDBLhOWECSQQJcJywgyXA1EbucSEYCoBMuXScsIMkgAa4TFpBkkADXCQtIMkiA64QFjGXQCZeuExaQZJAA1wkLSDJIgOuEBSQZJMB1wgKSDBLgOmEBSQYJcJ2wgCTDJXXsmjoZCYBOuHSdsIAkgwS4TlhAkkECXCcsYCyDTrh0nbCAJIMEuE5YQJJBAlwnLCDJIAGuExaQZJAA1wkLSDJIgOuEBSQZJMB1wgKSDNeVsgtLZSQAOuHSdcICkgwS4DphAWMZdMKl64QFJBkkwHXCApIMEuA6YQFJBglwnbCAJIMEuE5YQJJBAlwnLCDJIAGuExaQZJAA1wkLSDJcXM2urpaRAOiES9cJCxjLoBMuXScsIMkgAa4TFpBkkADXCQtIMkiA64QFJBkkwHXCApIMEuA6YQFJBglwnbCAJIMEuE5YQJJBAlwnLCDJcIVBu8RgzhqDuMigXWUwY5lB6IQr1wkLSDJYatB1wgKSDFYbdJ2wgCSDBQddJywgyWDNQdcJC0gyWHbQdcICkgxWHnSdsIAkg8UHXScsIMlg/UHXCQsYy6ATrlwnLCDJcKVNl4A0JBkkwHXCApIMEuA6YQFJBglwnbCAJIMEuE5YQJJBAlwnLCDJIAGuExaQZJAA1wkLSDJIgOuEBYxl0AlXrhMWkGSQANcJC0gyXG7WrjebkQDohCvXCQtIMkiA64QFJBkkwHXCApIMEuA6YQFJBglwnbCAJIMEuE5YQJJBAlwnLGAsw6WW7VrLOYst42rLdrnlnPWWccFlu+Jy1pLLuOayXXQ5IwG47LJddzln4WVcedkuvZyz9jIuvmxXX85ZfhnXX7YLMOeswIxLMNs1mHMWYcZVmO0yzBmdcAWdcOU6YQFJBglwnbCAJIMEuE5YQJJBAlwnLCDJcOFxu/J4RgKgE65cJywgySABrhMWkGSQANcJC0gySIDrhAUkGSTAdcICxjLohCvXCQtIMkiA64QFJBkkwHXCApIMEuA6YQFJBglwnbCAJMPV9+3y+xkJgE64cp2wgCSDBLhOWECSQQJcJywgySABrhMWMJZBJ1y5TlhAkkECXCcsIMkgAa4TFpBkkADXCQtIMkiA64QFJBkkwHXCApIMt6Cwe1BkJAA64cp1wgKSDBLgOmEBSQYJcJ2wgLEMOuHKdcICkgwS4DphAUkGCXCdsIAkgwS4TlhAkkECXCcsIMkgAa4TFpBkkADXCQtIMtyHxW7EkpEA6IQr1wkLSDJIgOuEBYxl0AlXrhMWkGSQANcJC0gySIDrhAUkGSTAdcICkgwS4DphAUkGCXCdsIAkgwS4TlhAkkECXCcsIMlwMyK7G1FGAqATrlwnLGAsg064cp2wgCSDBLhOWECSQQJcJywgySABrhMWkGSQANcJC0gySIDrhAUkGSTAdcICkgwS4DphAUkGCXCdsIAkwx257JZcOXty4aZcdleujG25oBMeuE5YQJLB1lyuExaQZLA7l+uEBSQZbNDlOmEBSQZ7dLlOWECSwTZdrhMWkGSwU5frhAUkGWzW5TphAUkG+3W5TljAWAad8MB1wgKSDHemcwlIQ5JBAlwnLCDJIAGuExaQZJAA1wkLSDJIgOuEBSQZJMB1wgKSDBLgOmEBSQYJcJ2wgCSDBLhOWMBYBp3wwHXCApIMEuA6YQFJhtsz2v0ZMxIAnfDAdcICkgwS4DphAUkGCXCdsIAkgwS4TlhAkkECXCcsIMkgAa4TFpBkkADXCQsYy6ATHrhOWECSQQJcJywgySABrhMWkGS4R6ndpDQjAdAJD1wnLCDJIAGuExaQZJAA1wkLSDJIgOuEBSQZJMB1wgKSDBLgOmEBYxluzmt3583Znhf357Ub9Obs0Itb9No9enM26cVdeu02vVn79OJGvXan3owE4F69drPenN16cbteu19vzoa9uGOv3bI3Z89e3LTX7tqbs20v7ttrN+7N6IQH0AkPXCcsIMkgAa4TFpBkkADXCQtIMkiA64QFJBkkwHXCApIMd6u221VnJAA64YHrhAUkGSTAdcICkgwS4DphAUkGCXCdsICxDDrhgeuEBSQZJMB1wgKSDBLgOmEBSQYJcJ2wgCSDBLhOWECSQQJcJywgyXDLdrtne0YCoBMeuE5YQJJBAlwnLCDJIAGuExYwlkEnPHCdsIAkgwS4TlhAkkECXCcsIMkgAa4TFpBkkADXCQtIMkiA64QFJBkkwHXCApIMEuA6YQFJBglwnbCAJIMEuE5YwFgGnfDAdcICkgwS4DphAUkGCXCdsIAkgwS4TlhAkkECXCcs4KvseH1f1+3VrJ2dn912v/51tph3v8+b5bp30zwtuwM25xyj3vrvq/ru/dFV+e5q8FIS3zfPV6vm8ap5Xr4/6r8+MFk+PrU/1uv17Fu9f/DDatWs8MHZYtE8f7eYLX/pzrA7wV8fu8cX83XbPetds3p4WsyK86MPy3rVvPlYf11tfv9xtvqteXPxdTVfdH/+tXkzfVrON78uul8vvjXrtnnzpX5s5/VDd8CbTzft0+b3n5p/bB+5mt+8/uno7Hj/JGfH8atUr/qifHfx33nVf/7p05s/f7ns/rv4T070Q/nuw3/nRC/+cjX5+dP15OJL7+pD7/LTDxfffbq+uJx8+unNHvU+XV99+Kn3f70fJj9+nnz4/wtAlxfXP3+4PvAi6YH1+dljd5bdLHybd8O6qO+6We2/3bxDr+bf7vd/aZvHl1P/2rRt8/Dyx/t6dluvNl/Q8bumaXd/Oe6e5blZ/fISivN/AVBLAwQUAAAACADJePhcQZXySVcEAAB9DwAAGAAAAHhsL3dvcmtzaGVldHMvc2hlZXQzLnhtbL1XUW/iOBD+K1FOOt2dogYMpV2WIlEClJbtVkW79+wmBqw6npzjwMKv33ESUsrFbLWq+kBI5ptvPGOPx+PeBtRzumJMOz9iIdMrd6V10vX9NFyxmKZnkDCJyAJUTDV+qqWfJorRKCfFwieNRsePKZduv5fLHlS/B5kWXLIH5aRZHFO1vWYCNldu090LHvlypY3A7/cSumRzpr8lDwq//MpKxGMmUw7SUWxx5Q6a3VnL6OcK3znbpAfvTrqCzUTxaIYDYxwN19kBxPOQCnZvvBc4WgOlJuIngGdDmkZG0XggmbOdJ4LnPjkakhlb6CETyBoQ16Gh5mv2gGpX7hNoDbHBMRpNNYoWCnZM5r4xwVAXfU7+p1wYKYxeN9Fq+l8emHmvAjduHb7vYxznK4Az+kRTNgTxL4/0yoTkOhFb0EzoF+HlWafTbnTIeYU9wuaGlVPePmub4UIQaf50Ni+mwixFd/e2XSfmsvinP8rFOiC02hYCKQnkiGBiriW0SkLrmNCxENoloX1EILYYzkvC+fEIthg6JaHz1hguSsLFW2O4LAmf3kr4ZCFYgzYrWqxc43jprJRqsY8Xj9hmqkleUfwisfKkDaim/Z6CjaNyfZN/L0NXGYk7MDQaA1Mf8vRBKZemhMy1QpSjQd1/pLs//yDt1mf5T8/XOJIR+2FJvj5NHpjNxyMaOb7zoGCLuxTqzAwLMy2LmS8srSEFp0lDLBJLUOh86+IzreGPTvMHKV9KGkENc3yaOcICVcubnObhDC0VjWlU5+3Nae6YhSvq4MhnzlTykNcNP32ziTGXNfzb0/wgUzTkZbIYO9yE4vwVlUuQ/l1j8+4Xiwh4GmmqauOZ/WIBE9jCUeL4uC2qvUGqLUByQ+3ckDlTX/LbigytSGBFRlZkbEUmVuTGikytyK0VubMiszrk1US2qolsWSfSigytSGBFRlZkbEUmVuTGikytyK0VubMiszqkmEj/oHJH+PxOBZZN09KkTgiZ1OYQPYb2rcyAdAfnpssy3VigIAlgI/NmzAimMsk0VtAUG75KOFIK1KGQCuwVrwWVz0Uvtk1QLniqcVTTh2aCNvvu3f1X724+xN/A7fmVfK9B+o1KSjCo1+7a3A9IN/gY90drKjKsUSC9QRZxDYpTbyS1YkvqTeW6LDVeXjmc2wwbWx7uP6nzePN+QY9Jd/wxQWMJTQQzJxKGik2ASsF7ZMn+rAFvzlFs6uyOqvcLcEK6k48JcF6cL949vJ/3Q9IdflBOSqbAG7MnZf6/ULXDhHtSXOD7FrzbTHLzFCYtl4AdnDdnieYsRoL3NdSZ+b+HdSkJ8OzP335nLo4EaXFLRJ+WHOuQwPsUBnZ2gW29KprI4gMvbnmExbWr6DnxssqUUThvNi+bzQZpdQhptLFfXwDoeujlVpolDu5Nsx2NH1duAkoryvE6l9CEqTnfsbwzX6HWDlBNBAk3XS5O+5opzcMDiWmPqxt3/ydQSwMEFAAAAAgAyXj4XBv83/jlDQAAHXIAABgAAAB4bC93b3Jrc2hlZXRzL3NoZWV0NC54bWyNnV1zGmcWhP8Kpcu9sJgZvpSyXeVYld2tym654v24JtLIYoM0LIzjxL8+A9i8fTqn23OTxHrgQcjdSUwzLy8/dftfDo9t209+e9o+H15dPfb97rvr68PdY/u0Przodu3zQB66/dO6H365/3B92O3b9f3pTk/b63o6XVw/rTfPV69fnr72bv/6Zfex326e23f7yeHj09N6//v37bb79Oqquvr6hZ82Hx774xeuX7/crT+079v+37t3++FX1xfL/eapfT5suufJvn14dfWm+u62WSyP9zjd5D+b9tMB/nlyeOw+/XW/uf9xeOjhmUyvJp+77un93Xrb/vP4/W+Hx5sOXz0+55+77pfjnf5+f7zh8Xt4bie/v99tN6fvatJ3ux/bh/5tux3u9aapF1eT9V2/+bV9N9zw1dXPXd93T8dbDM+oX/fDlx723ef2+fTdtdt2uO3wfe/+dOOz5Kw9PZvJ4f+nZ3fb1MvL87u+PCn8569P9ofTb8bww/15fWjfdtv/bu77x+Nzu5rctw/rj9u+fHH1YrGYTRf1/MJ+6j79rf3y05+9mB0f7q7bHk5/nXw636sZVHcfD8N3/dV9NXnaPJ//vv7ty+8b3KFaiDvUX+7Q0B3qWtxh9uUOp2/s+vydnZ717bpfv3657z5N9qdbH59A+T4vT2n4vbw73uLNMWunxx++unk+xvF9vx/oZhD2r/+1aZ/v1395ed0Pj3L80vXdlzt+7+/4Q3v3uJ687fb7tsvu/nbM3d/8b/im2+zut+e7N+Lubz7eb/puH+94PfxQLj+Z+vIDqE+m2cl0bGd5hmdST/+M3mp0m/nCQzeXh27kQzf6oTW6zXzhoWeXh56dbjpPHvpM6ip5aI1uz2ihH3r+9aG/n58tdfIAKQqaxUWz0JoUBc3yollqTYqCZnXRrLQmRUFzc9HcaE2KgqaaXjzHf8kpUc6iqSqmyphSFk11MdXGlLJoaoqpMaaURdOsmGbGlLJoKqmuTKxzFk0l2JVJds6iqWS7MuHOWTSVeFcm3zmLppLwykQ8Z/Hf2yXjtcl4zqKpZLw2Gc9ZNJWM1ybjOYumkvHaZDxn0VQyXpuM5yyaSsZrk/GcRVPJeG0ynrNoKhmvTcZzFk0l47XJeM6iqWS8NhnPWfwfhJLxxmQ8Z9FUMt6YjOcsmkrGG5PxnEVTyXhjMp6zaCoZb0zGcxZNJeONyXjOoqlkvDEZz1k0lYw3JuM5i6aS8cZkPGfRVDLemIznLP6faMn4zGQ8Z9FUMj4zGc9ZNJWMz0zGcxZNJeMzk/GcRVPJ+MxkPGfRVDI+MxnPWTSVjM9MxnMWTSXjM5PxnEVTyfjMZDxn0VQyPjMZz1n8I0/J+NxkPGfRVDI+NxnPWTSVjM9NxnMWTSXjc5PxnEVTyfjcZDxn0QR/ynR/zPx2xucl43OT8ZxFU8n43GQ8Z9FUMj43Gc9ZNJWMz03Gcxb/PF4yvjAZz1k0lYwvTMZzFk0l4wuT8ZxFU8n4wmQ8Z9FUMr4wGc9ZNJWML0zGcxZN8GKKezXl2xlflIwvTMZzFk0l4wuT8ZxFU8n4wmQ8Z/HFopLxpcl4zqKpZHxpMp6zaCoZX5qM5yyaSsaXJuM5i6aS8aXJeM6iqWR8aTKes2gqGV+ajOcsmuA1Q/ei4bczviwZX5qM5yyaSsaXJuM5i69kloyvTMZzFk0l4yuT8ZxFU8n4ymQ8Z9FUMr4yGc9ZNJWMr0zGcxZNJeMrk/GcRVPJ+MpkPGfRVDK+MhnPWTTBS+PutfFvZ3xVMr4yGc9ZfJm9ZPzGZDxn0VQyfmMynrNoKhm/MRnPWTSVjN+YjOcsmkrGb0zGcxZNJeM3JuM5i6aS8RuT8ZxFU8n4jcl4zqKpZPzGZDxn0QQLkJuAxmxAOALZFWjEDDSFHWjqhqAckgymoKnJuoAkgzVo6uagHJIMBqGpW4RySDLYhKZuFMohyWAWmrpdKIckg2Vo6qahHJIMxqGpW4dySDLYh6ZuIMohbZfQALuDjhpCcQm1U+iIBuAYatfQMXMo7qF2EB2ziOIkajfRMaMorqJ2Fh2zi+IwapfRMdMobqN2HB2zjuI8avfRMQMpLqR2Ih2xkVYwklZuJRWQZNAAN5QKSDJ8O4BrQA5JBg1wc6mAJIMGuMVUQJJBA9xoKiDJoAFuNxWQZNAAN50KSDJogFtPBSQZNMANqALSe0WgAW5DFZBk0AA3owpIMmiAW1IFJBm+J8a+KWZEA2BOrdyeKiDJoAFuUhWQZNAAt6oKSDJogBtWBSQZNMBtqwKSDBrg5lUB6T1O0AC3sApIMmiAG1kFJBk0wO2sApIMGuCmVgFJhm8Ms+8MG9EA2FsrN7gKSDJogNtcBSQZNMDNrgKSDBrgllcBSQYNcOOrgFEG82vl9lcBSQYNcBOsgCSDBrgVVkCSQQPcECsgyaABbosVkGT47kj79sgRDYBBtnKLrIAkgwa4UVZAkkED3C4rIMmgAW6aFTDKYJyt3DorIMmgAW6gFZBk0AC30QpIMmiAm2kFJBk0wC21ApIMGuDGWgFJhm8Rtu8RHtEAWGwrN9kKSDJogFttBSQZNMANtwJGGUy3ldtuBSQZNMDNtwKSDBrgFlwBSQYNcCOugCSDBrgdV0CSQQPclCsgyaABbs0VkGT4Pnn7RvkRDYBJt3KbroAkgwa4WVfAKINht3LLroAkgwa4cVdAkkED3L4rIMmgAW7iFZBk0AC38gpIMmiAG3oFJBk0wG29ApIMGuDmXgFJhheL2KtFRjQANt/Kjb4CRhnMvpXbfQUkGTTATb8Ckgwa4NZfAUkGDXADsIAkgwa4DVhAkkED3AwsIMmgAW4JFpBk0AA3BgtIMmiA24MFJBleMWUvmRpzzRReNGWvmhpx2RRswrXbhAUkGVw65TZhAUkGV0+5TVhAksEFVG4TFpBkcA2V24QFJBlcRuU2YQFJBldSuU1YQJLBxVRuExaQZHA9lduEBYwy2IRrtwkLSDK8ctBeOjiiAbAJ124TFpBk0AC3CQtIMmiA24QFJBk0wG3CApIMGuA2YQFJBg1wm7CAJIMGuE1YQJJBA9wmLGCU4YWz9srZMZfO4rWz9uLZUVfP4uWzrgE5JBk0wF5CO+YaWryI1l5FO+YyWryO1l5IO+ZKWryU1l5LO+ZiWrya1l5OO+Z6Wryg1l5RO+aSWrym1l5UO2ITrmETrt0mLCDJoAFuExaQZNAAtwkLSDK8htxeRD6iAbAJ124TFpBk0AC3CQtIMmiA24QFJBk0wG3CApIMGuA2YQFJBg1wm7CAUQabcO02YQFJBg1wm7CAJIMGuE1YQJJBA9wmLCDJ8CAFe5LCiAbAJly7TVhAkkED3CYsIMmgAW4TFpBk0AC3CQtIMmiA24QFjDLYhGu3CQtIMmiA24QFJBk0wG3CApIMGuA2YQFJBg1wm7CAJMPTROxxIiMaAJtw7TZhAUkGDXCbsIAkgwa4TVhAkkED3CYsYJTBJly7TVhAkkED3CYsIMmgAW4TFpBk0AC3CQtIMmiA24QFJBk0wG3CApIMj9SxZ+qMaABswrXbhAUkGTTAbcICkgwa4DZhAaMMNuHabcICkgwa4DZhAUkGDXCbsIAkgwa4TVhAkkED3CYsIMmgAW4TFpBk0AC3CQtIMjxXyh4sNaIBsAnXbhMWkGTQALcJCxhlsAnXbhMWkGTQALcJC0gyaIDbhAUkGTTAbcICkgwa4DZhAUkGDXCbsIAkgwa4TVhAkkED3CYsIMnwcDV7utqIBsAmXLtNWMAog024dpuwgCSDBrhNWECSQQPcJiwgyaABbhMWkGTQALcJC0gyaIDbhAUkGTTAbcICkgwa4DZhAUkGDXCbsIAkwxMG7RGDY84YxEMG7SmDI44ZhE24cZuwgCSDowbdJiwgyeC0QbcJC0gyOHDQbcICkgzOHHSbsIAkg2MH3SYsIMng5EG3CQtIMjh80G3CApIMzh90m7CAUQabcOM2YQFJhidt2qM2RzQANuHGbcICkgwa4DZhAUkGDXCbsIAkgwa4TVhAkkED3CYsIMmgAW4TFpBk0AC3CQtIMmiA24QFjDLYhBu3CQtIMmiA24QFJBkeN+sakEOSQQPcJiwgyaABbhMWkGTQALcJC0gyaIDbhAUkGTTAbcICkgwa4DZhAUkGDXCbsIBRhkct27OWxxy2jKct2+OWx5y3jAcu2xOXRx25jGcu20OXRzQAj1225y6POXgZT162Ry+POXsZD1+2py+POX4Zz1+2BzCPOYEZj2C2ZzCPOYQZT2G2xzCP2IQb2IQbtwkLSDJogNuEBSQZNMBtwgKSDBrgNmEBSYYHj9uTx0c0ADbhxm3CApIMGuA2YQFJBg1wm7CAJIMGuE1YQJJBA9wmLGCUwSbcuE1YQJJBA9wmLCDJoAFuExaQZNAAtwkLSDJogNuEBSQZnr5vj98f0QDYhBu3CQtIMmiA24QFJBk0wG3CApIMGuA2YQGjDDbhxm3CApIMGuA2YQFJBg1wm7CAJIMGuE1YQJJBA9wmLCDJoAFuExaQZPgRFPYzKEY0ADbhxm3CAp5l1/DJdMdPMfzHev9h83yYbNuH4bbTF8edb3/+YLrzL/pud/rwuvNHAp4/x65d37f74w3mVbWqqmk9/J7X0+N/eh66rs9R+dTEj7tJt9+0z/36+ImDr6523b7frzf91WS33rX795vP7fFzuCaPw60+d8PNtre7zfll0smv7b7f3MFXjh+5d/lEyNd/AFBLAwQUAAAACADJePhckx8OOdUHAAAOMQAAGAAAAHhsL3dvcmtzaGVldHMvc2hlZXQ1LnhtbL2bXVOjSBSG/wqVrZqLLWtCd8iXo6nCBGbd1Wgljju1d5i0hhpCs0DM6K/fBiLNxzktKrM3GunzHk4/nHR4Q3uy5+GPaMNYrP3cen502tnEcXDc7UarDds60WceMF+M3PNw68Tiz/ChGwUhc9apaOt1qa4PulvH9TuTk/TYdTg54bvYc312HWrRbrt1wqcz5vH9aYd0Xg4s3IdNnBzoTk4C54EtWfwtuA7FX908y9rdMj9yua+F7P60Y5Jj2xgmgjTi1mX7qPBaizZ8/zV01xfizGIiekd75ny7XDkemyfle+J0ujiaTPmO8x+J6HydBCYl+Ex7WgaemxalxTy4YPfxlHlCZdJxR3NWsfvIrkXcaeeOxzHfJgFiPrETi0P3IX9mfloc85iIFVUHteAsSSFr9G82NTo+/m7Pstl18ykVX79M1U6vhCB750Rsyr2/3XW8SWbW0dbs3tl5sTw4+jwYGPqA9vOxBd//wQ7ojc9GcroV96L0p7bPVD2RarWLRNEvuTva1vWz387Pw0UrCIiBCOhBQKsCigh6B0GvKhghAuMg6FcEFDvD4CAYpKSzuadcZ07sTE5CvtfCNDpBJEnk0ESvrJIIM2nldIbiqOsn3b6MQzHqioTx5MZl/tr5/aQbi7Mkh7qrg/AsE/YQ4QUX/doVF/nTb5T2v/jMAXJM1TkuWQSIZuqKbbbaOJq5i3noPjsrV5zf6H3xoSlYTTLdOp67Vuex1dMwd2tXVFMWdsUVyi8Tza8GVstfczLWbqypeWNB1wKdAg8dBD0msXwWcgg7zWBlumSVfJwYA52OTrqPRabNwmzk9MnqfRwFzkosOGJ5jlj4yDoT21rMzfnsSvv6zVosrMWVpoDZy2FiF0TAHGnX1mxhfTUvIJzKpoZgYgIUZg+mNKzAhMOMXgUmcnoY5tVyai4089K8Pf9naqpIGjlJAyc50KbW7HypaE5MrGhOTILyNGBQ/QrPZmE2cvoWmrOfI+2jSHVdo7o2s7T51e25dXm2AKlierRHMQHKtA/DGleYNguzkdO3wHSQMx3gTIlmfb9eWMslxBLToSwxAcpyAELq6xWWSBipsERO3wLLYc5y+MH+xPQoU0xgs7sQoToEcQ2q7/pmYTZSQAtURznVEU61p1nzpTU3ZyaEExMqFlFUggMdwaSqb/lmYTZSQAtAxznQ8ftumDCZAicqwXGOm+GEw4bVT3mkgBZwEl3aAf29HYoqFUxxDQ410UC4BhWqWFz1ThSroQ2uBZtF3s0VU6q4ohoFV9KIl4XFVd/9WA1tcJWGieCOSe9rl9b386l5cS5ezL/NrCuQL5YB/aRCFZdO+AyzhV3RyKiybRZnYxV89I6fSPNEcPf0WsNiSlXDot8BoEhhb1RH2izOxipoo12lkSK4k3qN6jtsFKrBqcIOqU61WZyNVdAGVemliMJMvUIVU6qoYhqcKuyR6lSbxdlYBW1QlW6KKOzUULsWyW6utGvr029UH38xl1N4dcWSqABjGhwwbJxGVSeAxQ2rgN/ksN6wvkp/RXCDRYba4mppLs5vYKSoU1IgxTQ4Utg1jXtVpEhcrWd/nb0i0l8RhcFqZFvRBPjdAKbA0YL+iei15aBZnI1V0AZa6bQIbrXesBxgSXC8mALHC/opotcWg2ZxNlbBRxcDKl0XxV3XK19Vo0oUKKpAgVLYSdWWAjiO6BVnZmMVfBiotFsUt1uvAcWUOFBMgQMFLRTRq1//N4yzsQo+DLTwYErxZOr1RwCoWvWA6s0ui4LuCcDaLM7GKmhhXaXSaFHFYyr1jQCqVFF9s9GioIEihFSpNouzsQraoCqNFn230UKVKqpvNloUNFAA1WZxNlbBh5cA6bKowmU1+6oFzYCvrZjCvAtdSDGjoIcipPZh1SzOxipoo12l16IKr9XovhVNgKPFFDha0D0BaJvF2VgFbaCVRosqjJZyT8U7bBaqwZmC9onQ6pNBLK62Dvw6m0WlzaLvfo6FKlVUMQ1OFXZOlFapNouzsQraoJo7rBmtGZJCpeBYeeNL7idmvdqduMwEj5UzEZmpdgtayASOlTNRmal211XIBI6VM/VkptqdRiETOFbOZMhMtU/XQiZwrJypLzPVPk4KmcCxcqaBzFRbPQuZwLFypqHMVFszCpnAsXKmkcxUe58UMoFj5Uyyx3uKHofHyluSZI8bih6Hx8qZZI8bih6Hx8qZZI8bih6Hx8qZZI8bih6Hx8qZZI8bih6Hx8qZZI8bih6Hx8qZZI8bih6Hx8qZZI8bih6Hx7JM3cJm2bX4mW3zTLY7R9qK7/w425VZHnrZ5nxGj8+MccJuw/ezkAczvvfTndrJgXM/2MWXLIqcB5YftMKQh8WDjufx/Znn+D+yjdpPgTjuuVEszprsUt95Dpl00ru3o+I22s5JNx9+CaQTPT9KxdzKVWOzmNLj6f8zi3Rb1NHh2fNR6omO0s9u8fqJH/25893kpyd+mg88ivnRkgWxy7ZCcHS1infJ7zl/PByZuavs1XtYVA5E2bZ9UdODKy69x+7Fldc/D0U7h9lW6eyPmAfpDLNd8NnOauasWZgE9AkZEaLT3oBSPem6e85jeEj+m8Au0HjoMj9O6xD3DzyMQ8eNO1rgBCxcus8s2eiibZKNzFyEebPATT7lBfZHFsbuqnAk2QSe/wvE5D9QSwMEFAAAAAgAyXj4XM9QP6OyBgAAbicAABgAAAB4bC93b3Jrc2hlZXRzL3NoZWV0Ni54bWytmlFzmzgUhf+Kxp3p7L7ESBDbaZPMyJhN3GLwgJ1t+0ZsJWGKgQXStPn1K2GC3N0b7kzwQ1Mbc67PAcF3JXP+lBXfywchKvJzl6TlxeChqvIPw2G5eRC7qDzJcpHKT+6yYhdV8m1xPyzzQkTbWrRLhswwRsNdFKeDy/N627K4PM8eqyROxbIg5eNuFxW/piLJni4GdPCyIYjvHyq1YXh5nkf3IhTVOl8W8t2wrbKNdyIt4ywlhbi7GHD6wbZGSlDvcROLp/LgNSkfsqerIt668ptlEGNAqug2FInYVGJbf/dzlu3CTZQIT8VJ5DZD7qUOwW2WfVdF5lslVJZSQX6FeRLXJkmV5a64q2yRJMqINSDRpop/iKXc72Jwm1VVtlM7yHxVVMlNd0X2LNLabO1ApdhL9jWm1JT7/lMHU69/2zH/X9XfpGykpaw+IMP2KBy+fjk6f9UnT56M26gUdpb8HW+rBxV+QLbiLnpMKr1xcjIaWcaInbafBdnTtWjOlnViqa/bZElZ/yVPTSk2IJvHUtp9qT0guzjd/x/9bM7zgcB8TcAaAfuPQB1yUGA2gvoQDvfO6tSzqIouz4vsiRT13iqAabxUaSPJk71Re3A1NutycmucquEbVoX8NJYFq0tbpFWRnQ8r+SVqy3DT6KbdulUs0m1E/ngso4KIshIkzXa3hSDipzyn2Z9ARbu7YhA9v3/HLPNjSspsE0fJ7yWGMnEbm7XpWF1z/ErNz7ZhQNm6VbYzm4dkyd3FfOVDQfZyNc7gbw15h3ez9W5i3inkvVs1DfgN6Nns49lqPVuYZwZ57lbZ3P3GZxxybfVxfdq6PsVcm5DrbtUy8L869soPId+nfXyPWt8jzLcF+e5WdYzrUR/X49b1GHN9CrnuVoXc467vgYNk3Mf2pLU9wWyPINvdqqs1X6zt67kL+Z708X3W+j5DfFNwcHerpnPHc8IVDyDfZ318U0OTycCcgwMFkfnrleusyMKXhx28Nhv9W/0fkJVi/sERg8gc2/f8kLsz+M7SqN/qXpOTYuikY9B9t0wO+JDfOKDzXtikmpsUAycDoY/IPAlOl8M3RtqLnlTjk2L8ZCA/Edn12glWHB4uvQhKNUIpxlAG3mYQ2ZJLHPkreLz0gijVFKUYRhmIUURm++r+OLdBJNFeKKWapRSDKYPvMd2ytnchC/6Nr1zugSF6gZVqslIErR7cpaMywgwyc4jn35CF461nDnzx7stMXi3jd6XQnKUIaD24X0dlxPmyDJwQvnzPenhnGrUMQa0Hd8CojDhe6HiyeZcD6asfwCegqfLGEJq3DOGtB7fDqOwgxOujqKnyxhAHE1YEux7cHaMy6f2LvBm5884QrE8ITWCGENiDe2VUdhCiYziZfUJoFjOExZ4BNkCojCyd9++YcfaRhzYcwOoTQCOZIUj2jDMwACaTvXPI55/h/rNRv9G8hjJDoOxR+GLGZGQ1/7SWk0XSOT9vyrwxhcYzQ/DswVMYVEamPrnhrusovl0tnRMwxLhPCI1nhnEWnsegMrJfSFs5Noc7PNaHzEyTmWGIhacyqIwE8kIIXlsE7INmU6PZxBhLJ+AyIIrmpTMLnCsOLj2YfZBsaiSbGFspeAtCZR2DxuwDYlOD2MSIyuDl126Z58sj7oD3HLMPfM2DZWOEoiHcTaMyydxVMP8GWt9rz2Ct+ontQ5lHG3ExyAtRiuKHGMiCNukKpEFsIkQNX1kIx2TdCGjkR82k2WwikA3hhXJURpauv3Bcx5aXdzD/S/6De4ym0FHTaXibCIVDeDqByiS8uXe1Vr++yGThGlx+bKocNZomuomgOYT7c1Smo13PXRc+Z+PjB9OUNxFch3C7i8qIP5WY8cEVhEZ81EQa+ybC79CAyYnJiFoK/OSDCzuN+JiJLN0LWAjUQ7ijR2Wyow8cF74PNuKjJtINgoWQPqQgrBBZ+wut7JKdYAb2CU2Jo+bS3YOFdA8hBZmFysj1mq/4lC+W4C2i0R81lO4sLKxFgGczqIyE3COf/FDNZlxi8ymc7fhNhnXwwzXWLcCTHFRWx9kndNc2vAhvHb/VsHSrYWE9Azz7QWWaWB13eOv4fYal+wwLaxjgHh2V6S6qY+3VOmKfMTx4dEc9FraIivs4LUki7mRx42QsD2Oxf3Jn/6bK8vrpnv3DUvsHfUS0FYXa4ZTSCaUGM0eMGZaMeZdlFfyRfgztMSdZEYu0itSzWNJ0VlRFFFcDkke5KML4WdRhH+Rez5ncLZnlsZqpSj78EEUVbw62qGeS2kfsLv8FUEsDBBQAAAAIAMl4+FyOP945FgQAAIsbAAANAAAAeGwvc3R5bGVzLnhtbNVZW4+jNhT+K4j3LhcTEqokUobZkSq11Uo7D/vqBJNYMpeCM03219cXEhOwZ8gMtFOiEba/c75z8cEYz7KmZ4K+HxCi1ikjeb2yD5SWvzpOvTugDNZfihLlDEmLKoOUdau9U5cVgknNlTLi+K4bOhnEub1e5sfsKaO1tSuOOV3Zru2sl2mRqxEvsOUIk4UZsl4gWdkxJHhbYS68O8CqZr6IcU+owwyTsxzwhUhBisqizDvUiNQ/G3nRcyR918imwpDoTUjGar9d2U9PYB54s1ta9320W5MJeY1hAptMhA9zf+GOYWL8dBvz4oUbEGwmNfEYzedueL+JXhX2LTQpArZFMS/1X9wvfhAxex6/5osI3G929BIy5/4r2MxuCcNBuX/j2flYmBqvhpXxB57o+1ajC+WixShuNRPGhFxXvtCWA+tlCSlFVf7EOkJHDPYgq2k/n0tWVfsKnj1/Zg9WqAuCE25yH7dz4c9C8PVB5PwWYIt4GMeCv8X5QWthMA8WOmtqeRrRmoptkhC0pOLGpnpbVAmqrpMN7MvQeklQSpl6hfcHfqdFyfNRUFpkrJFguC9yKCrhotHWtMQLemXTg3jB3i5m3uPsUfrGRRsbAzWErHBnoAKTvPg9UEMKvzNGnCfohBL24ASDIuzKvxFfV/yt6Lryr8XWNFhZ7BAh3znbj1RtgRjnKbXkVum3hO+SLL5gXJqsoJqmpJEdzt9mk9wtWh+8i9cq8UtBH44smlz0/zoWFH2rUIpPon9Krw6Y2L1J2X3F7rfZmVVYluS8IXifZ0imdrDB9RJe9KxDUeGfzBpfx3dsAFW29YIqinftkb8rWD6jE23eBM4pHeQz+L/4DAyzeIeDr8yiiX2cGgn+E/ZxMjOblD2clH0+Pnv0b61a0bRr4sRL7rRrrudPSr9Q7EGb3f/Ey2PL59mn9nnYAzSdz28l8vM45QWTvv3a9BO8oNr04yw2TrOlbO1bb3at11GLfy2v7D/5mSBRFNb2iAnFuWbHyjgp3BJ0S8pUEpTCI6HPV3Blq/YfKMHHzL9KfeOxNFKq/Tvf6cvTCrFXZ7aa3XrcdNmnSedjl19coYuoc5U+YtJxXf6nRzhmsmPywKTDx/XIwhiP6y6MCMf0bCadhVGHj+uR2OU/HaIOAbpIxC59pFEEQBiaMioPMHoexKa8hSE/9tCzmeJRByVdRH35Ds+1ebbNFfJ6HZjm9LUKMUVqrkSe0/tyzRF93vgVRfrZNtmRmN6OqXYkpkPUyVEXASCO9Xa4fdMTbEaiSI+oc+lejV4PoHoehPynQ9SxaQ+5HrH35xQAvW8AmBD+NJoR/ZPF2fSzoM5pbxFeh3o2Ph5FesQUDx83IaZ4JGLyQB+PRIA49Hc67yPn8p5y1L/f1v8AUEsDBBQAAAAIAMl4+FyXirscwAAAABMCAAALAAAAX3JlbHMvLnJlbHOdkrluwzAMQH/F0J4wB9AhiDNl8RYE+QFWog/YEgWKRZ2/r9qlcZALGXk9PBLcHmlA7TiktoupGP0QUmla1bgBSLYlj2nOkUKu1CweNYfSQETbY0OwWiw+QC4ZZre9ZBanc6RXiFzXnaU92y9PQW+ArzpMcUJpSEszDvDN0n8y9/MMNUXlSiOVWxp40+X+duBJ0aEiWBaaRcnToh2lfx3H9pDT6a9jIrR6W+j5cWhUCo7cYyWMcWK0/jWCyQ/sfgBQSwMEFAAAAAgAyXj4XOPhrXivAQAAKAUAAA8AAAB4bC93b3JrYm9vay54bWy1lF9u2zAMxq9i6H2z6zRBF8QBihVbA+xPsRbtsyLRMVdZMiQ6bnubHWCn6MVK2zDmYIDRFz9J/AiRP3wCuWmcf9w79xg9lcaGTBRE1TqOgyqglOGjq8ByJne+lMShP8Sh8iB1KACoNHGaJKu4lGjFdjPUuvHxdtNe7hGa8E9vw+iIAfdokJ4z0d0NiKhEiyW+gM5EIqJQuObaeXxxlqS5Vd4Zk4mzPnEPnlD9J9+2PHdyHzrl6QGtdk0mPpwlFyJ6Pg2bLnpATUUm0kW6XA3aNeChIC6xuPjEIsn9L0noMrFMmCtHH6hr1GFKRXgE7slpjmpyX9AQ+CtJ8NW7ukJ7aGnYjHjkRufccPa2r/17jHd5jgqunKpLsNQ778G0gDYUWAURWVlCJnY2kK+VYh1Caw832uneKmK4kfF+jZzwO91hzod0WWsk51//yjFQOgGUzgzUfh5qqU8cWkwALWYG+l0HOoE5n4A5nxfmO/CjMctygmU5L8sdgtUyRMfXP4Y/bEy1mqBadXM3DJuGHC3oH1wxsM77Q934qD12/c7Ja2M+c/jTfnNSD2M7bK7tG1BLAwQUAAAACADJePhcZX9aaswAAADNBAAAGgAAAHhsL19yZWxzL3dvcmtib29rLnhtbC5yZWxzxdS9DoIwEAfwVyF9AE8Q0RhgcnE1vECDx0egtOmdEd5e1AGbOLiQTs1d0//9lmt6xV5yqwdqWkPBqPqBMtEwmxMAlQ0qSRttcJhvKm2V5Lm0NRhZdrJGiLbbBOx3hsjT78ygmAz+k6irqi3xrMu7woF/BMND244aRBZBIW2NnAkY+6VN8D7CzZwsgsstE/ZyCwX4BkUOKPIP2jmgnX9Q7IBi/6C9A9r7ByUOKFkRRDz1SIvmUzvjDyuO5/ktLtPf5afp7vXxZQDn98qfUEsDBBQAAAAIAMl4+Fxw+bXcLgEAAHcGAAATAAAAW0NvbnRlbnRfVHlwZXNdLnhtbM2Vz07DMAzGX6XqdWozBuyA1l2AK+zAC4TWXaPmn2JvdG+P226TQKNiKhK9NGpsf78vtpSs3g4eMGqMtpjFFZF/EALzCozE1HmwHCldMJL4N2yFl3kttyAW8/lS5M4SWEqo1YjXqyco5U5T9NzwNipnsziAxjh67BNbVhZL77XKJXFc7G3xjZIcCSlXdjlYKY8zTojFRUIb+RlwrHvdQwiqgGgjA71Iw1mi0QLpoAHTYYkLHl1ZqhwKl+8Ml6ToA8gCKwAyOu1FZ8Nk4g5D/70Zze9khoCcuQnOI08swPW400ja6sSzEARSw0c8E1l69PmgnXYBxS/Z3N4PF+puHii6ZXyPv874rH+lj8VEfNxOxMfdRHzcT8TH8h99vDtX//VV2K6pkcqe+KJ7b9afUEsBAhQDFAAAAAgAyXj4XEbHTUiVAAAAzQAAABAAAAAAAAAAAAAAAIABAAAAAGRvY1Byb3BzL2FwcC54bWxQSwECFAMUAAAACADJePhcxVoe1j8BAADXAgAAEQAAAAAAAAAAAAAAgAHDAAAAZG9jUHJvcHMvY29yZS54bWxQSwECFAMUAAAACADJePhctlGYhtkCAAAsDAAAEwAAAAAAAAAAAAAAgAExAgAAeGwvdGhlbWUvdGhlbWUxLnhtbFBLAQIUAxQAAAAIAMl4+Fzel/ockwgAAFUaAAAYAAAAAAAAAAAAAACAgTsFAAB4bC93b3Jrc2hlZXRzL3NoZWV0MS54bWxQSwECFAMUAAAACADJePhc4t7tOJYTAABDoAAAGAAAAAAAAAAAAAAAgIEEDgAAeGwvd29ya3NoZWV0cy9zaGVldDIueG1sUEsBAhQDFAAAAAgAyXj4XEGV8klXBAAAfQ8AABgAAAAAAAAAAAAAAICB0CEAAHhsL3dvcmtzaGVldHMvc2hlZXQzLnhtbFBLAQIUAxQAAAAIAMl4+Fwb/N/45Q0AAB1yAAAYAAAAAAAAAAAAAACAgV0mAAB4bC93b3Jrc2hlZXRzL3NoZWV0NC54bWxQSwECFAMUAAAACADJePhckx8OOdUHAAAOMQAAGAAAAAAAAAAAAAAAgIF4NAAAeGwvd29ya3NoZWV0cy9zaGVldDUueG1sUEsBAhQDFAAAAAgAyXj4XM9QP6OyBgAAbicAABgAAAAAAAAAAAAAAICBgzwAAHhsL3dvcmtzaGVldHMvc2hlZXQ2LnhtbFBLAQIUAxQAAAAIAMl4+FyOP945FgQAAIsbAAANAAAAAAAAAAAAAACAAWtDAAB4bC9zdHlsZXMueG1sUEsBAhQDFAAAAAgAyXj4XJeKuxzAAAAAEwIAAAsAAAAAAAAAAAAAAIABrEcAAF9yZWxzLy5yZWxzUEsBAhQDFAAAAAgAyXj4XOPhrXivAQAAKAUAAA8AAAAAAAAAAAAAAIABlUgAAHhsL3dvcmtib29rLnhtbFBLAQIUAxQAAAAIAMl4+Fxlf1pqzAAAAM0EAAAaAAAAAAAAAAAAAACAAXFKAAB4bC9fcmVscy93b3JrYm9vay54bWwucmVsc1BLAQIUAxQAAAAIAMl4+Fxw+bXcLgEAAHcGAAATAAAAAAAAAAAAAACAAXVLAABbQ29udGVudF9UeXBlc10ueG1sUEsFBgAAAAAOAA4AnAMAANRMAAAAAA==';
var PLANTILLA_GENERADOR_B64='UEsDBBQAAAAIAC0j9FxGx01IlQAAAM0AAAAQAAAAZG9jUHJvcHMvYXBwLnhtbE3PTQvCMAwG4L9SdreZih6kDkQ9ip68zy51hbYpbYT67+0EP255ecgboi6JIia2mEXxLuRtMzLHDUDWI/o+y8qhiqHke64x3YGMsRoPpB8eA8OibdeAhTEMOMzit7Dp1C5GZ3XPlkJ3sjpRJsPiWDQ6sScfq9wcChDneiU+ixNLOZcrBf+LU8sVU57mym/8ZAW/B7oXUEsDBBQAAAAIAC0j9Fz2ykoBFAEAAFwCAAARAAAAZG9jUHJvcHMvY29yZS54bWzFkl1LwzAUhv/K6H17knSrELqCH3jlULCieHdIz7Zg04Yk2s1fb1u3zqH3XuY973nyBJIrK1Xr6MG1llzQ5Gc7UzdeKruMtiFYCeDVlgz6pG80/XDdOoOhP7oNWFRvuCEQjGVgKGCFAWEAxnYiRkVeKakcYWjdAV+pCW/fXT3CKgVUk6EmeOAJh6h48uRyOC0PoEDO+O+Aqok2pn8ixwlEh+bO66nVdV3SpWOv9+fwsrp7HJ8a68YHbBT1W17LsLe0jI43P6fXN+VtVAgmWMx5zLKSL+RcyJS/Dq5nfidh01Z6rf/bOIvZRSxYyeZSLKTIfhgfBYu8/xI1+rA6BFf74t4rdLNLgx/6U2EOvxtjdv6Rii9QSwMEFAAAAAgALSP0XFjKdHfyBQAAjhoAABMAAAB4bC90aGVtZS90aGVtZTEueG1s7VnNixs3FL8X+j+IuTvzPeNZ4g322E7a7CYh66TkKM/IHmU1IzOSd9eEQElOvRQKaeml0FsPpTTQQEMv/WMWEtr0j6hmxh8aW85H45SUxgZbevq9p5/ek56kmYuXzlICTlDOMM1amnnB0ADKIhrjbNzSbg36jaYGGIdZDAnNUEubIaZd2v/4o4twjycoRUDoZ2wPtrSE88merrNIiCG7QCcoE20jmqeQi2o+1uMcngq7KdEtw/D0FOJMAxlMhdmB0AExAtdHIxwhbX9hvkfET8ZZIYhIfhSVfVY6EjY+Nos/NmMhycEJJC1N9BTT0wE64xogkHHR0NKM8qPp+xf1pRLhW3QlvX75mevNFeJjq9TLx8OlouO4jtde2rcq+5u4nt/zet7SXgmAUSRGam5g3U7Q6bpzrASqigrbXb9rmzW8ZN/ewLfd4lvD2yu8s4Hv98OVDyVQVXQVPvGt0Knh3RXe28D7Rrvr+DV8CUoIzo430Ibr2eFitEvIiJIrSnjgOn3fmsNXKF2aXZV+xrfNtRTepXlfAMrgQo4zwGcTNIKRwIWQ4GGOwQEeJ2LiTWBGmRAbltE3bPFbfJ2yVHoE7iEoaVeiiG2ICj6ARTme8Jb2qbCqSZBnT5+eP3hy/uDX84cPzx/8PO97U+8KzMay3osfvvrru8/Bn798/+LR12o8k/HPf/ri+W+/v8w8r9H65vHzJ4+fffvlHz8+UsDbORzK8AFOEQPX0Cm4SVMxQEUHaJi/mcYggbimAROBVAB7PKkBr80gUeE6qO7C27nIFCrg5endGtejJJ9yrABeTdIa8JBS0qG5cjhXi77k4UyzsbrzfCrjbkJ4ouo7XAtwbzoRUx6rTIYJqtG8QUS04RhliIOijR4jpFC7g3HNr4c4yimjIw7uYNCBWOmSAR5ytdIVnIq4zFQERahrvjm8DTqUqMx30UkdKZYFJCqTiNTceBlOOUyVjGFKZOQB5ImK5NEsj2oOZ1xEeowIBb0YMabSuZ7PanSvigyjDvshmaV1ZM7xsQp5ACmVkV16HCYwnSg54yyRsZ+wYzFFIbhBuZIEra+Qoi7iALOt4b6NEX+zZX1LZCD1BClaprlqSSBaX48zMoIom28EtZSe4uyV+X0ts7v/TmZ/Zzl999m8nWPlmlrP4dtw/8HM3YXT7AYSi+VD4v6QuP+PiXvbWt59ul5laF0+q5dm0q0H9xEm5IjPCDpgZW5nYnhxXwjLSqm0vCdMElGcd1fDjXNYlkFO+WeYJ0cJnIhuzLKHMZubHjMwoUzsDtpW2+XuMk0PaVxJTXNxNRUKkK/kYndZyMVexCup56/uYEvzZW3MZAJuafT1SUid1UnYChK+/XokTGNXLAIFi6b5Mha6FBWx/gAsnmu4TsVIzDdIUFzEqdJfRHfnkd7mzPqwLcXwAmdnka6RkKZbnYQ0DRMYo3XxjmMdBOpQW0oafvNdxFrfzA0kq9fAqVhztivMRHDS0kbiXCiK6UTYY0XehGSctbSIzx39TzLLJGe8C1lSwcqmavwp5igHBKdirsthINmKm2n5xvtLLjDeP8/p60FGoxGK+BbJqiraKiPK1rcEFxU6FaSPkvgUDMk0vwmFo1zfLBwYY8aX3oxxLk3ulRfX0tV8KdYema2WKCSTBM53FDmZV/CyvKQjjaNkuj4qXeXC4bi/i1331UprSXPLBuJvzWLvbpOXWNlqVq4y1wVN4+W7xNtvCBK1ppqaraa2be/Y4YFA6s7b4jdrazTfcjdYn7W6dK4saxvvJujwrpj5XXFcnRLOqvv/mbgjhIunylUmKKWL7HLGwTTHLe2e4bad0HLDhtF0ew3HdoxG023bjbbr2mbPNY1ux7ovnMKT1HSrvvviPkNm85cvpXzjBUy6OGZfiGiq0/IcrJfK5QsY06q9gKnOyWBQtGsAC8/c86x+YAcdrxHY7X7D6XaajSD0Oo2uF/rdfjd0m0H/vgZOSrDTtkPH6zUbnhmGDcczCvrNoOE7ltV2/Haz57Tvz30tRr74X7i35LX/N1BLAwQUAAAACAAtI/Rc1GiR2msHAAAAEwAAGAAAAHhsL3dvcmtzaGVldHMvc2hlZXQxLnhtbI1Ya1PbOBT9Kxp32gkz2TwMgUKAmZBmu50tj0nK7mfFURKlsuVKMq9fv+dKthMoePkAia3Hvbr3nHOvcnqvzU+7FsKxh1Rl9ixaO5efdLs2WYuU247ORYaRpTYpd3g0q67NjeALvyhV3bjXO+ymXGbR+al/d2POT3XhlMzEjWG2SFNuHi+E0vdnUT+qXkzlau3oRff8NOcrMRPuNr8xeOrWuyxkKjIrdcaMWJ5FF/HJRXxEC/yMf6S4tzvfmV3r+69GLr7DMg7Sixgdbq71Txr+tqBXmC+USBxtyvFxJ8ZCqbNoRJ798mZG3qduve/u98renz4cON6cWzHW6l+5cOuz6HPEFmLJC+Wm+v4vUR5xQPslWln/n92HufsRSwrrdFquhQOpzMInfygjszO/3+u9sSIuV8Te72DIe/mFO35+avQ9MzSK7ejLBeYjPPvHEYN3MqNEzZzBsMRCd36jeOakUpx9+vA57sdD9lVkwvCFNmyyEUmBqOnTroMtWtBN8AcbtaGD2tCBN3TQe8PQP58+xPFgqOSCs5wbjtNx9auQwjDDnzB4sD/MmNWJ5Iq1/r66brO/Z/Q3ajPhks5egxOHtROHwYn+G06Mg51Us2WRJYAFb9j1KGKU0cM6E3WOa3NHwVz8ljlOhxXWcbLbH3IARjFukjWCCpiDWlZkjrPbK5YKi1FYSnMlU4nXusMmillpHbiJMQckc4YNaGqO/OBrptO5EbRQvbDU8jM2AttpFk2uJtNrFvfiw6jNoj8nF9P6udPp7DHdhiHsxhT5lehMJ4LxNTbbPxq2K2suPLsCs34VggmbGAlSMJGR/UQonHcUkzvCcrbWG/gRjXeOREO5weZWW/ZJuSHO8mnlhv576bumF9FepyE1x3UGjpsTfsOtZv0a218KeJJwctY7R46GEDUY6/f+Hwg0pwkJf/EnlsA0LCK8a11G7FnGdpNEHg+OD4YsutR3oIhmic4lN9sByCoOEo2hzoYVGQ8TIvaI/HlU8N8MZIW4o3lZBSLK4R1/RALxmpsVdmqJTYft4mOvw640m2sDrNKOy0AhkxaKB9zZkHgLCGkAI+HpXMI4ErzgTtumPPa3OtWP35HJuM7kJEv4XDxBpzxriGi2SApjuWqyt/+OVO6/g9RzpSl4DqAWIfoe+9FXpDdz4oSBVRGyFo0WUG1Q2Ctq+bplxQpRHPQgeIkPbK6zBd9jj20cZs43YGP0A1sv+Am7ur5EMtiXCfs+YrPb8e10NvoeddhUEG75E4yWU75OppOrHxP2gUAQVVPpkehb5wP44Aq5rCL5UkWqMJIrFcWlsYLJFRS7BK7CkTh2Qn3HixbKUJEiCtDqklceYDhROV84iXhhSXWuyM8JYWyHguAVZSt4yruqOMm0CHGC/qUU3kZIDbaZHLwDUvt1IMY8d4XxCvuxBhXYkojcNerD4TtAddgMqomPkuNzRZwubbIu+xjVnlSxSrZuIq8J1Y+NYK0eONxvs9C6afZxr15Y6i0D+VUBGJ6wG5n8lNmqzX4YbnNEAdydZM6IFad8XUM86MVdkHvkYTa6+eMCPYbhyc82mxYOpL9BHknTha2HOuymEAtgq7CQEu0MllYBtD7hCXdlH6BXmkoO94KM9oy5gqH9xDayagZalwKHKX3DSULVG+s0J//e9KLsF3YB/hxTZXnzePbqBo8zgRhxI0Nhc5qskbRVbMAUJ7IVPKU9UmlTfAy2pxuyfOfoaTikhQIAsJi5lIoO4ukFP8AmvimIQpBM3gjnz1sMfW6G82zrapnqMmK8WEinTcgmyCk9k5qMHr8D0cfNiJ5Jr4u1mlhCo5JLKr9UgmjwhWMJMqtIJFpoRMjxCrkSOb6TC6RrjxRy41EKqmSJ3tKG4L6DNSRVVDL9WJOGwu0l5AV9ljKrtS1I+YJQ4KgXVUGagta/rICz20tGmdSlcpci6l1C6RytPKmoEnPFqShHo9sv335cT7+NiNo2L+j0+8NAj9eF+BXJbVOwfBeBVVWX16o1v48j1w+Hja1UvE1p3H+HXh7UvAJxHIU3EKUsAZZdwPa4yWD8/+CK42ZwfeeVPRZ99OUuNJkhM1EZwNeLVOnhXkho9iydoSXaIFU+sLRUud0eowQUL+B24LgDoENRCkpSdsQgoKaOKeVZ4SlZNdONudjepuKDd+RiUOdiVsw9pMu7RZONwTvCP2gO/y3ifieemFIQt3ZoHXF0sMoLY33DaXUelH2gSmQX3r25dpW4R+PQcE4ewLfQDbxy83z1CvRa9Cnfv98j/aU/6E3JYnhZEQ7FAozl2SLwCHKR6tRfUurKQfjJ6xtyRjjIUXA8Q0nAfrcIRoskZJqkrr295vLnt1p/7fI1w5eM5ioQH21zc9QMjBE1U06a1+6Ez24YuFZUVWpHQx7raGGhKRKvmlXzRcJMAmXZSmQhjEYmkNvHGu+IZK/tmYfXvq0zVaNqXteh7s6vF/Tb0CVwITOIiljigL3OEcBoAkLDg9O5/zkEYAIO/Ne14LhY0QSML7V21QP9RlL/6HX+H1BLAwQUAAAACAAtI/RctFES1+wIAABvPQAAGAAAAHhsL3dvcmtzaGVldHMvc2hlZXQyLnhtbK1ba3PiOBb9Ky6mZmqmdjYg+U0nVIU0NFo73SlIZj47oCSuBszaTjL771e2hYVAkmXDl25AvudI9+hxrh1ffybpz+wN49z4Z7PeZje9tzzfDfv9bPmGN1F2lezwlrS8JOkmysnX9LWf7VIcrcqgzboPBwOnv4nibW90Xf72kI6uk/d8HW/xQ2pk75tNlP5vjNfJ500P9PY/zOPXt7z4oT+63kWveIHzp91DSr71a5RVvMHbLE62Ropfbnq3YBjaZhFQXvFXjD+zg89GHj0v8Bovc7wqmYqhPSfJz6IRkZ8GRQ/LCwrIiPz3ge/wen3TG8MB6dh/S5bic92LIvTw855vWqaDDO85yvBdsv47XuVvhJTArPBL9L7O58nnDNMxWldWAblM1ln5r/FZXW76V57nuMBz7Z7xjLN8GpcpMZbvWZ5s9qA9YxNvq/+jf2jKDlDcK8sC1sCBOiCQgsAjEE8j1qSx5lGsfdABYaBFA62jQOheOY4q0KaB9smQwcA3XVmYQ8OcozDnIN3CQJcGuoqONqbJoyDecacl1/v0ev8cUlDPjuPp0ThosJ8ToJwU/WqaltP8a5RHo+s0+TTS4nqCW3y4LVYx+U4GRHpUrPv+kjZNQPlbvC1W/yJPSVtMwPLRN5zibY6Hxvcf9+P5xPg2mU++P04McN3PSXxxUQ0yawsCBSBBWxCTB+mTUddDh/XQYTl0KAO/e9/s1vEmJgwJ2QmMXZoscZZkBt7iNDHgADqCvo4rVNM5TegdbRLlumpyJF15JJ1YRUNj8XT3NF/chuJUt8QQZrolhirR5j7RE7ME9WR5TrZLvMsTQXemVaQlifxVlIXOZKg9WdCZLGxFxuXVqiewVWL4SrFEU7SKA7KJLxrpHY0ZKBfLb79Ay/wiCJ9U4aYk+iFe/oy3r6IZUAVCwZqZdcVEFNOsMT9Gg+v+x6GwXbFDJTano13raFcxslXHVr1h3D59RY8/5uhWpCvFKYBfRv+a+tf9l4KfH9pddRXwyqvQ9Pex/dtr/sX1f/2zt4h7f/a+J70/hJETW5mUxzTKdhHZIEUy2nIZu6Miu1nI7uihEp2T0qmldMoYu0lJ0d47rmKBU8mHJPJVV0FYy+foyucoUzHZ5il+jcjpZvx4TkX7xtSRi3guNnKapTyXI1RycIK6taCunqCmSFCXW4+hRFD3WFBXV1C3ISEfZD823S/EsixuH/49JiYwjZY/Rcq6cmUvRoLcZokvRhYqyTitvVprT09rS6R1FQvp1gsHYq29473X09XaU2Zm/p5HmfFADvzCruJMKbYnF/tyLMhrVvtybKGSjZPbr+X29eS2RXL73NJGMr3947Xt6+pdRQo0mlJMT57YmTQY+QfTdPF0/zuyhsgTdyGQooTHKKE1DE9RuLSDASs5B5p+x2nwO3ug/QYrU4Fed7DswKBZB777BxUz0Js3rrDLgDvkpyaQdBmcdBloTp3/0Fg44JTjh8OqYAD1huMJh0NraJ+uA+lw4PFCAFB3JRSxrap/0f48a40iWvZBaxRHUcgBk4lg6ongC0UwORFCKDEa++sORDC1RTDb3RgQa9ASRCxBSxClAqyWBpaeAmAglMDil7UFJRJYJxJY2hJYnW+j0NBW91G606EOdEF3urAdHT8BWBEObM0JILr3NgZ87Y2kE8A+mQDa5TdQ17KKmyiguVCedUZHGuhBZ/QQCG8d8Cqy+htoFuBAWIEDvgQPLdlxdlKEA+0qHKjLWOVdFNBcJc/OwEca+MEZ+CEQ3kHgtWSlN9CsvYGw+KbR9ZZsmxIt3ROnpV2Ag6Z6tfGWCmguimdnsyANluBslhAI7yLw6rJiG9Ba2OSuptrRNkv0VEpdNba56wGaa9TZ5eiQBl1wOboQCOt8Xg5WDANfIYevkMO/3I0JiqXW43J8SIMvuCBfuOdTKAJZmQwHckX2bSJFaJvo1gFtYoU7cQZDknbx9jaTI6ETJOIChkiGFMiRwj2S+H4Gn5y6CB9DWh/bp4jfZE90i7+KGZLDaYlversUZzj9wL2RoXpWzJ5hQtkMkFV+ovJ/1hpFVHUHrVF81RhZDQStg9LqaFIdNmlUXeLRtwQRD74liHLsNtPX7lxd0dBW1VV3OtSBLuhOF7aj49PLfDmkzhqKNjN52wSq3aai5oEaXrkzOtJADxrQlU5Zjc9nmTlm6CqODFdxZKi9n7ImgRoO9gx8pIEfNODf43QTidOs/5gIMusKFdYVKqwrbDJ3jeUC1HCsZ7MgDZZAgyVaEWu0S1LjLtnsxCWDmonPP/OqUOFVocKrQrWXa1M6QA2rejk6pEEXQPnzGnjywAY6Q5J89TMPk3lRU+FFTYUXpW0XqQ4oljLlF+RDaj4+U+zpkFkZU9MTZQooMgWkrp02Hbh2oh5Jh8S1y5HQCRI5x4ZIgMSPrn5YNDah1HZPFW1I0RaK2/ge1B51YsqMiPSvUkU37GftYUS3fYP2MFDhl0xmR027ncuVjLItiniQbVGUY3TYGJ3OlpuGtvvD0M50qANd0J0ubEfHp9dl6VWbIoV1NjUsXWd0pIEedEYP1eh8rjyWK7WbURpgU8OTnYGPNPCDM/BDNT6fMZ9lrMncNHpZU8NCnc2CNFiCs1lCNQv/59u1q5pYarfSxo9aGubocnRIgy5ooJMXZWpwPpmAJRNczmpSLHU2L8eHNPgCS2rrQuvY1pFdcEjSqLZ1FmS5g1LzSZsOzCfBJgmSmE85EjpBIifBEAmQ+F4yT2TJngTerjbxNs6KijNJa8t1+/UefUeLx/nt1x/zk4eEPAkzJZbT0uwocdlpbLmdzQ4N7fSuCDviLPURoXoFQ/+AsNgBYak3VeWRaulvpzbbTu2m/a3xSLL1dx6b7Ty2eidos43byo2A7wBbvja83NZHsbR6wJ6O2KZ0A6FNBxuI5Q0JjWTZ9w9eVtzg9LV8qzczlsn7Nq/eW6x/pW8Sw+Fd9bIju7x6Dfk+Sl/jbWas8QsJHVyR3qXVW7zl5zzZlZ+KVzKTPE82+29vOFrhtPhGuv6SJPn+S0FSv2M9+j9QSwMEFAAAAAgALSP0XBaFnwXSBAAA7SwAAA0AAAB4bC9zdHlsZXMueG1s3Vptb+JGEP4rliPdh6qJ1+92D7jeISFV6lWRLh/64aTIwAKu/FazpHC/vjteA2vYSQmJwQ0I2d7ZZ+aZ2dlX3FuyTUK/LShl2jpNsmVfXzBW/GIYy8mCptHyLi9oxiWzvEwjxh/LubEsShpNlwBKE8MixDPSKM70QS9bpaOULbVJvspYXzd3RZq4/DblhZ6ja0LdMJ/Svv54+5N28/PNDbkj5PH24/fmI0g//L3K2cdbcfn0qSr99fFWNwY9ozY56M3yTLJs6aKEE4hSqj1FSV8fRkk8LmOAzaI0Tjai2IKCSZ7kpca4zxR485LlDyE2xROEo9aTxlleVsaFhS7aGddKj2ySK9i8kJ+X8qqcj/v6aETIF+JeNJy1YTBNrmF4SM41jKjk2kamb31uqPRegCY+GTYJWaejbdsxXffYnQb6JYjqAkNRnCS7oSjURcGgV0SM0TIb8YcKUxUeibT6/mFT8K4zL6ONabn6yYBlnsRTMDkfNpN1G6lxLYizKV1TPiDz8Ri0SxrPtVV390DXWAyu35I7JwzDwPEd4juu5VmXIWBJBAJOICSBwz++6fptMzjqKe17a17KTB3UKzWqu7dvc/u+6wauGVoO/72UQHXh/XScl1NaNnqqKBr0EjpjHF7G8wVcWV6AjZyxPOU30zia51lUdeMtQkZq1bKqr7NFtSyaKJlBxdrCSfWrmhWVk6rzelvGJ9UXVS/km7wY6YpjxntpjEZw/48t8dIUM65P2XgvVE4Y3S7UURAm9Q0fvic0Sb6Bzj9nuzHc5prXM2m7SWCzmYmdJ7/lA39dKtSIB+MAFJ4DcuwTUYZMXLgheeA85wJB9WpF/JSzLyseuax6hs0yvS/pLF5Xz+vZf/qKaefeREWRbD4n8TxLqdjan2gQzgCiLU5b5GX8g1uDpfOEF1CxVVjPcFJ2Z0iRk0id2w6Sdmuv3ZG1269wmbzWZadVlzHtr0hsdbfxW20uS9ZuXbO5JFKuTMrtCCm7M4ktkfJkUk5HSB32h+uRwka9dmef17jc2uzTSVKvmRLPSw7Tb3VtIuWe2x2fxR8Yb9QlrCOnrUuFtDtpdNoy48Kkur4w9roYqe40nzqnrDda/qmT41B7J6aI7pDyO54cwdsnh49r76TL3VtcXnXDgpHyukjqqrsobIr03noDf+jyudrbXsZJ7INWl3Fhq9rNdtegJn5S+lbJ/ERLFk8glVle6No/ZVQ80DWr/zitktqoz2ClM+XGifKuVIM3KPr6H/D6UrI3p41XccLibOfuIeA+L6EfRX/R/WQqg1wV6CuPS1TS5T7HZIjdPErm3Fk0TmiTPKc2pbNolbCHnbCv7++/0mm8Sq1drXuIb11rf/87nOGL9zKqU3tuqz63H9aP5Xws/f9N6g8ADiWj6qOWYBghU0tAhtnBGGAYgcLsvCd/AtQfIcO4BUpJgGICFCNQKsmw+mJ21Bh4yUPtaRjatudhER0OlQyGWNw8D35qbRg3QGB2wNLLYo23Np4hz+cB1qbPZQjmKZ6JmKd4rEGijhsgwlDd2pgdQGCtgOUO2FfbgZxSY2wbWhXjhvVgXBKGmARyUZ2jnodEx4Ovun2wXmLbYaiWgEzNwLYxCfRGXIIxAA6YxBbz4MF8ZGznKWP/gvPgX1BLAwQUAAAACAAtI/Rcl4q7HMAAAAATAgAACwAAAF9yZWxzLy5yZWxznZK5bsMwDEB/xdCeMAfQIYgzZfEWBPkBVqIP2BIFikWdv6/apXGQCxl5PTwS3B5pQO04pLaLqRj9EFJpWtW4AUi2JY9pzpFCrtQsHjWH0kBE22NDsFosPkAuGWa3vWQWp3OkV4hc152lPdsvT0FvgK86THFCaUhLMw7wzdJ/MvfzDDVF5UojlVsaeNPl/nbgSdGhIlgWmkXJ06IdpX8dx/aQ0+mvYyK0elvo+XFoVAqO3GMljHFitP41gskP7H4AUEsDBBQAAAAIAC0j9Fw1sp8bggEAAAIDAAAPAAAAeGwvd29ya2Jvb2sueG1stZJPT+MwEMW/SuQ7xA3ailZ1L1t2qYQAAYKz60yaEf4T2RMCfHomjqLtaqXVXvZkzxtl5vn3shlCfD2E8Fq8O+uTEi1Rty7LZFpwOp2HDjx3mhCdJi7jsUxdBF2nFoCcLSspl6XT6MV2M8+6j+V2M16eEYb0Sx/L4g0THtAifSiR7xZE4dCjw0+olZCiSG0YrkPEz+BJ20cTg7VKLKbGM0RC84f8OPp50oeUlfcX9HUYlDhbyEtRfPxeDrl6wZpaJaqL6tty1q4Bjy3xiIvLFYukDw+aMCixXPF3DcZEeZESXGpD+Aa8M5vWPYUfaAniThP8jKHv0B9HNwyjPKGRyc3nhH0d/wV8aBo0sAumd+BpIh/BjgZ9arFLovDagRJ7nyj2xrAOacTDi/b1hIrY3An4uEZuxH2dbf4/S1e3Vw93RSWr5Ymf6i9+qoxtZlVDgx7qW56VWOf4zX0sxiO/a7WQ1Yrz6a39ztqdvwl6ei9yHrxhB5Y0x3QupZwCmf/J7RdQSwMEFAAAAAgALSP0XI33LFq0AAAAiQIAABoAAAB4bC9fcmVscy93b3JrYm9vay54bWwucmVsc8WSTQqDMBBGrxJygI7a0kVRV924LV4g6PiD0YTMlOrta3WhgS66ka7CNyHvezCJH6gVt2agprUkxl4PlMiG2d4AqGiwV3QyFof5pjKuVzxHV4NVRadqhCgIruD2DJnGe6bIJ4u/EE1VtQXeTfHsceAvYHgZ11GDyFLkytXIiYRRb2OC5QhPM1mKrEyky8pQwr+FIk8oOlCIeNJIm82avfrzgfU8v8WtfYnr0N/J5eMA3s9L31BLAwQUAAAACAAtI/RcbqckvB4BAABXBAAAEwAAAFtDb250ZW50X1R5cGVzXS54bWzFlM9OwzAMxl+lynVqMnbggNZdgCvswAuE1l2j5p9ib3Rvj9tuk0CjYioSl0aN7e/n+IuyfjtGwKxz1mMhGqL4oBSWDTiNMkTwHKlDcpr4N+1U1GWrd6BWy+W9KoMn8JRTryE26yeo9d5S9tzxNprgC5HAosgex8SeVQgdozWlJo6rg6++UfITQXLlkIONibjgBKGuEvrIz4BT3esBUjIVZFud6EU7zlKdVUhHCyinJa70GOralFCFcu+4RGJMoCtsAMhZOYoupsnEE4bxezebP8hMATlzm0JEdizB7bizJX11HlkIEpnpI16ILD37fNC7XUH1SzaP9yOkdvAD1bDMn/FXjy/6N/ax+sc+3kNo//qq96t02vgzXw3vyeYTUEsBAhQDFAAAAAgALSP0XEbHTUiVAAAAzQAAABAAAAAAAAAAAAAAAIABAAAAAGRvY1Byb3BzL2FwcC54bWxQSwECFAMUAAAACAAtI/Rc9spKARQBAABcAgAAEQAAAAAAAAAAAAAAgAHDAAAAZG9jUHJvcHMvY29yZS54bWxQSwECFAMUAAAACAAtI/RcWMp0d/IFAACOGgAAEwAAAAAAAAAAAAAAgAEGAgAAeGwvdGhlbWUvdGhlbWUxLnhtbFBLAQIUAxQAAAAIAC0j9FzUaJHaawcAAAATAAAYAAAAAAAAAAAAAACAgSkIAAB4bC93b3Jrc2hlZXRzL3NoZWV0MS54bWxQSwECFAMUAAAACAAtI/RctFES1+wIAABvPQAAGAAAAAAAAAAAAAAAgIHKDwAAeGwvd29ya3NoZWV0cy9zaGVldDIueG1sUEsBAhQDFAAAAAgALSP0XBaFnwXSBAAA7SwAAA0AAAAAAAAAAAAAAIAB7BgAAHhsL3N0eWxlcy54bWxQSwECFAMUAAAACAAtI/Rcl4q7HMAAAAATAgAACwAAAAAAAAAAAAAAgAHpHQAAX3JlbHMvLnJlbHNQSwECFAMUAAAACAAtI/RcNbKfG4IBAAACAwAADwAAAAAAAAAAAAAAgAHSHgAAeGwvd29ya2Jvb2sueG1sUEsBAhQDFAAAAAgALSP0XI33LFq0AAAAiQIAABoAAAAAAAAAAAAAAIABgSAAAHhsL19yZWxzL3dvcmtib29rLnhtbC5yZWxzUEsBAhQDFAAAAAgALSP0XG6nJLweAQAAVwQAABMAAAAAAAAAAAAAAIABbSEAAFtDb250ZW50X1R5cGVzXS54bWxQSwUGAAAAAAoACgCEAgAAvCIAAAAA';
var PLANTILLA_FILES=[
  {nombre:'Auditorías, Actividades, Ajustes y Mermas', archivo:'Plantilla_Auditorias_Actividades_Ajustes_Mermas.xlsx', b64:PLANTILLA_MASTER_B64},
  {nombre:'Generador Ejecutivo', archivo:'Plantilla_Generador_Ejecutivo.xlsx', b64:PLANTILLA_GENERADOR_B64}
];
function _b64ToBlob(b64){
  var bin=atob(b64), len=bin.length, bytes=new Uint8Array(len);
  for(var i=0;i<len;i++)bytes[i]=bin.charCodeAt(i);
  return new Blob([bytes],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
}
function _descargarBlob(blob,nombre){
  var url=URL.createObjectURL(blob), a=document.createElement('a');
  a.href=url; a.download=nombre; document.body.appendChild(a); a.click();
  setTimeout(function(){URL.revokeObjectURL(url);a.remove();},100);
}
function _bajarPlantilla(i){
  var f=PLANTILLA_FILES[i]; if(!f)return;
  _descargarBlob(_b64ToBlob(f.b64), f.archivo);
  var ov=document.getElementById('plantilla-overlay'); if(ov)ov.remove();
  toast('\u2713 '+f.archivo+' descargada');
}
function descargarPlantillas(){
  var ex=document.getElementById('plantilla-overlay'); if(ex){ex.remove();return;}
  var ov=document.createElement('div');
  ov.id='plantilla-overlay';
  ov.style.cssText='position:fixed;inset:0;background:rgba(15,23,42,.5);display:flex;align-items:center;justify-content:center;z-index:9999';
  ov.onclick=function(e){if(e.target===ov)ov.remove();};
  var box='<div style="background:#fff;border-radius:14px;max-width:440px;width:92%;padding:22px;box-shadow:0 20px 50px rgba(0,0,0,.25)">'+
    '<div style="font-size:17px;font-weight:800;color:#0F172A;margin-bottom:4px">Descargar plantilla</div>'+
    '<div style="font-size:13px;color:#64748B;margin-bottom:16px">Elige la plantilla que quieres descargar.</div>';
  PLANTILLA_FILES.forEach(function(f,i){
    box+='<button onclick="_bajarPlantilla('+i+')" style="display:block;width:100%;text-align:left;padding:12px 14px;margin-bottom:8px;border:1px solid #E2E8F0;border-radius:10px;background:#F8FAFC;cursor:pointer;font-size:14px;font-weight:600;color:#0F172A">\ud83d\udcc4 '+f.nombre+'<div style="font-size:11px;color:#94A3B8;font-weight:400;margin-top:2px">'+f.archivo+'</div></button>';
  });
  box+='<button onclick="document.getElementById(\'plantilla-overlay\').remove()" style="width:100%;padding:9px;margin-top:6px;border:none;border-radius:8px;background:#fff;color:#64748B;cursor:pointer;font-size:13px">Cancelar</button></div>';
  ov.innerHTML=box;
  document.body.appendChild(ov);
}


/* ══════════════════════════════════════════════════════════════════
   ANÁLISIS DE CONSUMOS DE COMBUSTIBLE
   Detecta posibles fugas / robos de combustible y errores de carga a
   partir del Excel de consumos (formato flexible). Flujo:
   1) analizarConsumos() lee y analiza el archivo
   2) envía un resumen al iframe para PREVISUALIZAR (dashboard)
   3) al confirmar, generarReporteConsumos() crea el Excel con colores
      por severidad (motor ExcelJS, cargado bajo demanda).
   Parámetros ajustables en CONSUMOS_CFG.
   ══════════════════════════════════════════════════════════════════ */
/* Umbrales por tipo de combustible: un camión diésel (F450/F550) rinde
   legítimamente 2-4 km/L y carga 300-400 L, mientras un auto de gasolina
   rinde 8-14 km/L y carga <150 L. Un umbral único marcaba como "fuga" a los
   diésel y dejaba pasar autos de gasolina ineficientes. PERFIL afina por tipo;
   REND_MIN/REND_MAX/LIT_MAX son el respaldo cuando no hay columna Producto. */
var CONSUMOS_CFG={
  PMIN:16, PMAX:34, DUP_MIN:60, DIST_MIN:30,
  PERFIL:{
    diesel:  { REND_MIN:1.8, REND_MAX:10, LIT_MAX:420 },
    gasolina:{ REND_MIN:5,   REND_MAX:22, LIT_MAX:150 }
  },
  REND_MIN:3, REND_MAX:25, LIT_MAX:200
};
var _CONS=null; /* caché del último análisis: {filas,col,nombre} */

function _cIdx(H, aliases){
  for(var i=0;i<H.length;i++){ var h=norm(H[i]);
    for(var j=0;j<aliases.length;j++){ if(h && h.indexOf(aliases[j])>=0) return i; }
  }
  return -1;
}
function _numC(v){ if(v===null||v===undefined)return null;
  var s=String(v).replace(/[^0-9.,\-]/g,'').replace(/,/g,''); if(s===''||s==='-')return null;
  var n=parseFloat(s); return (isNaN(n)||!isFinite(n))?null:n; }

function pedirArchivoConsumos(){
  var inp=document.getElementById('consumos-file-input');
  if(!inp){
    inp=document.createElement('input'); inp.type='file'; inp.id='consumos-file-input';
    inp.accept='.xlsx,.xls,.csv,.htm,.html'; inp.style.display='none';
    inp.onchange=function(){ if(inp.files&&inp.files[0]) analizarConsumos(inp.files[0]); inp.value=''; };
    document.body.appendChild(inp);
  }
  inp.click();
}

function analizarConsumos(file){
  if(typeof XLSX==='undefined'){toast('⚠ XLSX no cargado');return;}
  toast('⏳ Analizando consumos…');
  var rd=new FileReader();
  rd.onload=function(e){
    try{
      var raw=_leerFilasConsumos(e.target.result);
      if(!raw||raw.length<2){ toast('⚠ El archivo no tiene datos legibles (¿es el envoltorio del export?). Exporta como .xlsx.'); return; }
      var hi=0,best=-1;
      for(var i=0;i<Math.min(12,raw.length);i++){
        var nn=raw[i].filter(function(c){return String(c).trim()!=='';}).length;
        if(nn>best){best=nn;hi=i;}
      }
      var H=raw[hi].map(function(x){return String(x);});
      var col={
        fecha:_cIdx(H,['fecha','dia']),
        placa:_cIdx(H,['placa','unidad','economico','económico','vehiculo','vehículo','no. eco','num eco','no eco']),
        conductor:_cIdx(H,['conductor','operador','chofer','empleado','responsable']),
        litros:_cIdx(H,['litros','cantidad','lts','volumen','ltrs']),
        monto:_cIdx(H,['monto','importe','total','costo','pago']),
        precio:_cIdx(H,['precio','p.u','p/l','unitario']),
        odom:_cIdx(H,['odometro','odómetro','kilometraje','kilometros','kilómetros','km','kms']),
        estacion:_cIdx(H,['estacion','estación','gasolinera','establecimiento']),
        producto:_cIdx(H,['producto','combustible','magna','premium','diesel','gasolina','tipo']),
        tarjeta:_cIdx(H,['tarjeta','no. tarjeta','no tarjeta'])
      };
      if(col.litros<0 && col.monto<0){ toast('⚠ No se encontraron columnas de litros ni monto'); return; }
      var filas=[];
      for(var r=hi+1;r<raw.length;r++){
        var row=raw[r]; if(!row||row.every(function(c){return String(c).trim()==='';}))continue;
        var g=function(k){return col[k]>=0?row[col[k]]:'';};
        var lit=_numC(g('litros')), mon=_numC(g('monto')), odo=_numC(g('odom')), pre=_numC(g('precio'));
        if(lit===null && mon===null && !String(g('placa')).trim())continue;
        filas.push({ fila:r+1, fecha:String(g('fecha')).trim(), placa:String(g('placa')).trim().toUpperCase(),
          conductor:String(g('conductor')).trim(), litros:lit, monto:mon, precio:pre, odom:odo,
          estacion:String(g('estacion')).trim(), producto:String(g('producto')).trim(),
          tarjeta:String(g('tarjeta')).trim(), obs:[], sev:0 });
      }
      if(!filas.length){toast('⚠ El archivo no contiene filas de datos');return;}
      analizarFilasConsumos(filas, col);
      _CONS={filas:filas, col:col, nombre:file.name};
      var resumen=_consumosResumen(filas, col, file.name);
      var ifd=document.getElementById('iframe-documentos');
      if(ifd&&ifd.contentWindow) ifd.contentWindow.postMessage({type:'consumos-preview',resumen:resumen},'*');
      toast('✓ Análisis listo: '+resumen.conObs+' observación(es)');
    }catch(err){ console.error(err); toast('⚠ Error al analizar: '+err.message); }
  };
  rd.readAsArrayBuffer(file);
}

/* Lectura robusta: intenta XLSX binario; si no hay filas, intenta como HTML. */
function _leerFilasConsumos(arrbuf){
  try{
    var wb=XLSX.read(arrbuf,{type:'array',cellDates:false});
    if(wb&&wb.SheetNames&&wb.SheetNames.length){
      var raw=XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]],{header:1,defval:'',raw:false});
      if(raw&&raw.length>=2)return raw;
    }
  }catch(e){}
  /* Fallback HTML (algunos export .xls son HTML con la tabla embebida) */
  try{
    var txt=new TextDecoder('utf-8').decode(new Uint8Array(arrbuf));
    var doc=new DOMParser().parseFromString(txt,'text/html');
    var tabla=doc.querySelector('table');
    if(tabla){ var wb2=XLSX.utils.table_to_book(tabla,{raw:false});
      return XLSX.utils.sheet_to_json(wb2.Sheets[wb2.SheetNames[0]],{header:1,defval:'',raw:false}); }
  }catch(e){}
  return null;
}

/* Día calendario normalizado (yyyy-mm-dd) usando el parser robusto del
   dashboard: entiende dd/mm/yyyy, ISO y seriales de Excel. Sin esto, el
   orden de cargas se rompía (Date.parse no lee dd/mm/yyyy) y el
   emparejamiento carga-anterior↔carga-reciente producía falsos rendimientos. */
function _diaConsumo(fecha){
  var d=parseDate(fecha); if(!d)return '';
  var mm=('0'+(d.getMonth()+1)).slice(-2), dd=('0'+d.getDate()).slice(-2);
  return d.getFullYear()+'-'+mm+'-'+dd;
}
function _tsConsumo(fecha){ var d=parseDate(fecha); return d?d.getTime():null; }

/* Clasifica el producto: 'diesel' | 'gasolina' | 'nofuel' (servicio/portal) |
   null (desconocido → usa umbrales de respaldo). */
function _perfilCombustible(prod){
  var p=norm(prod); if(!p)return null;
  if(p.indexOf('uso y mant')>=0||p.indexOf('portal')>=0||p.indexOf('servicio')>=0||p.indexOf('cuota')>=0||p.indexOf('comision')>=0||p.indexOf('mantenimiento')>=0) return 'nofuel';
  if(p.indexOf('diesel')>=0) return 'diesel';
  if(p.indexOf('gasolina')>=0||p.indexOf('magna')>=0||p.indexOf('premium')>=0||p.indexOf('regular')>=0||p.indexOf('octan')>=0) return 'gasolina';
  return null;
}
function _limitesConsumo(perfil){
  var C=CONSUMOS_CFG;
  if(perfil==='diesel'||perfil==='gasolina') return C.PERFIL[perfil];
  return {REND_MIN:C.REND_MIN, REND_MAX:C.REND_MAX, LIT_MAX:C.LIT_MAX};
}

function analizarFilasConsumos(filas, col){
  var C=CONSUMOS_CFG;
  var add=function(f,txt,sev,cat){ f.obs.push(txt); if(cat){f.cats=f.cats||[];if(f.cats.indexOf(cat)<0)f.cats.push(cat);} f.sev=Math.max(f.sev,sev); };
  filas.forEach(function(f){
    f.cats=f.cats||[];
    f._ts=_tsConsumo(f.fecha);     /* orden cronológico correcto */
    f._dia=_diaConsumo(f.fecha);   /* clave de día para duplicados/frecuencia */
    f._perfil=_perfilCombustible(f.producto);
    if(f._perfil==='nofuel'){ f._skip=true; return; } /* cargos de servicio: no se analizan como combustible */
    var L=_limitesConsumo(f._perfil);
    if(f.litros!==null && f.litros<=0) add(f,'Litros en cero o negativos',3,'Litros en cero/negativo');
    if(f.monto!==null && f.monto<=0) add(f,'Monto en cero o negativo',3,'Monto en cero/negativo');
    if(f.litros!==null && f.litros>L.LIT_MAX) add(f,'Carga de '+f.litros+' L excede lo razonable para '+(f._perfil||'la unidad')+' ('+L.LIT_MAX+' L): posible error de captura o fuga',3,'Litros excesivos por carga');
    var pl=null;
    if(f.precio!==null && f.precio>0) pl=f.precio;
    else if(f.litros>0 && f.monto>0) pl=f.monto/f.litros;
    f._pl=pl;
    if(pl!==null && (pl<C.PMIN||pl>C.PMAX)) add(f,'Precio por litro fuera de rango ($'+pl.toFixed(2)+'/L): posible error de carga',3,'Precio por litro fuera de rango');
  });
  var porPlaca={};
  filas.forEach(function(f){ if(!f.placa||f._skip)return; (porPlaca[f.placa]=porPlaca[f.placa]||[]).push(f); });
  Object.keys(porPlaca).forEach(function(p){
    /* Orden: por fecha real; las filas sin fecha van al final; desempate por odómetro. */
    var arr=porPlaca[p].slice().sort(function(a,b){
      if(a._ts!=null && b._ts!=null){ if(a._ts!==b._ts) return a._ts-b._ts; }
      else if(a._ts!=null) return -1;
      else if(b._ts!=null) return 1;
      return (a.odom||0)-(b.odom||0);
    });

    /* ── Duplicados (comparación con la carga inmediata anterior) ── */
    var prev=null;
    arr.forEach(function(f){
      if(prev){
        var sameDay=f._dia && prev._dia && f._dia===prev._dia;
        if(sameDay && f.monto!==null && prev.monto!==null && Math.abs(f.monto-prev.monto)<=1 && f.litros!==null && prev.litros!==null && Math.abs(f.litros-prev.litros)<=1)
          add(f,'Posible carga duplicada (misma unidad, día y monto que fila '+prev.fila+')',3,'Carga duplicada');
      }
      prev=f;
    });

    /* ── Rendimiento carga-a-carga (método tanque-a-tanque) ──
       Distancia entre dos lecturas de odómetro válidas ÷ TODOS los litros del
       tramo (incluye cargas parciales del mismo día). Solo se mide con avance
       ≥ DIST_MIN km, para no marcar "fuga" en un tramo corto con tanque lleno.
       El umbral depende del tipo de combustible de la carga de cierre. La
       última lectura válida se arrastra: una fila sin odómetro no rompe la cadena. */
    var baseOdo=null, baseFila=null, accLit=0, lastOdo=null;
    arr.forEach(function(f){
      var odoOk = f.odom!==null && f.odom>0;
      if(odoOk){
        if(lastOdo!==null && f.odom<lastOdo)
          add(f,'Odómetro ('+f.odom+') menor a una lectura anterior ('+lastOdo+'): dato inconsistente',2,'Odómetro inconsistente');
        if(baseOdo!==null && f.odom>baseOdo){
          var km=f.odom-baseOdo;
          var litTramo=accLit + (f.litros>0?f.litros:0);
          if(km>=C.DIST_MIN && litTramo>0){
            var kmL=km/litTramo; f._kmL=kmL;
            var L=_limitesConsumo(f._perfil);
            if(kmL<L.REND_MIN) add(f,'Rendimiento muy bajo ('+kmL.toFixed(1)+' km/L, '+km+' km / '+litTramo.toFixed(0)+' L desde fila '+baseFila+'): posible fuga o robo',3,'Rendimiento bajo (posible fuga)');
            else if(kmL>L.REND_MAX) add(f,'Rendimiento muy alto ('+kmL.toFixed(1)+' km/L, '+km+' km / '+litTramo.toFixed(0)+' L desde fila '+baseFila+'): posible error de odómetro o carga no registrada',2,'Rendimiento alto (odómetro/carga)');
            baseOdo=f.odom; baseFila=f.fila; accLit=0;
          } else {
            accLit=litTramo; /* tramo corto: acumula litros hacia el próximo cierre */
          }
        } else {
          baseOdo=f.odom; baseFila=f.fila; accLit=0;
        }
        lastOdo=Math.max(lastOdo||0, f.odom);
      } else if(f.litros>0){
        accLit+=f.litros;
      }
    });

    /* ── Frecuencia atípica de cargas por día ── */
    var porDia={};
    arr.forEach(function(f){ var d=f._dia; if(d)(porDia[d]=porDia[d]||[]).push(f); });
    Object.keys(porDia).forEach(function(d){
      if(porDia[d].length>2) porDia[d].forEach(function(f){ add(f,porDia[d].length+' cargas de la misma unidad el '+d+': frecuencia atípica',1,'Frecuencia atípica de cargas'); });
    });
  });
}

/* Normaliza una placa/unidad para agrupar sin duplicados por espacios extra
   o diferencias de mayúsculas/minúsculas (la placa ya llega en mayúsculas
   desde analizarConsumos, esto solo colapsa espacios repetidos y recorta). */
function _normPlacaKey(p){ return String(p||'').toUpperCase().replace(/\s+/g,' ').trim(); }

function _consumosResumen(filas, col, nombre){
  var flag=filas.filter(function(f){return f.obs.length;});
  var alta=flag.filter(function(f){return f.sev>=3;}).length;
  var media=flag.filter(function(f){return f.sev===2;}).length;
  var baja=flag.filter(function(f){return f.sev===1;}).length;
  var tipos={}; flag.forEach(function(f){ (f.cats||[]).forEach(function(c){ tipos[c]=(tipos[c]||0)+1; }); });
  /* Vehículos: se agrupan por placa NORMALIZADA (para que la misma unidad no
     salga repetida por variaciones de espacios/mayúsculas) y se guarda el
     detalle de cada carga individual, para poder revisarlas una por una. */
  var vh={};
  filas.forEach(function(f){
    if(!f.placa)return;
    var key=_normPlacaKey(f.placa);
    var v=vh[key]=vh[key]||{labelCounts:{},cargas:0,lit:0,mon:0,obs:0,sevMax:0,detalle:[]};
    v.labelCounts[f.placa]=(v.labelCounts[f.placa]||0)+1;
    v.cargas++;
    if(f.litros)v.lit+=f.litros;
    if(f.monto)v.mon+=f.monto;
    v.obs+=f.obs.length;
    v.sevMax=Math.max(v.sevMax, f.sev||0);
    v.detalle.push({
      fila:f.fila, fecha:f.fecha,
      litros:f.litros!=null?Math.round(f.litros*10)/10:null,
      monto:f.monto!=null?Math.round(f.monto):null,
      precio:f._pl!=null?Math.round(f._pl*100)/100:null,
      estacion:f.estacion, producto:f.producto,
      sev:(f.sev>=3?'Alta':f.sev===2?'Media':f.sev===1?'Baja':'OK'),
      obs:f.obs.join(' · ')||'Sin observaciones: carga dentro de rango normal'
    });
  });
  var sevOrden={Alta:3,Media:2,Baja:1,OK:0};
  var totLit=0,totMon=0; filas.forEach(function(f){ if(f.litros)totLit+=f.litros; if(f.monto)totMon+=f.monto; });
  return {
    nombre:nombre, total:filas.length, conObs:flag.length, alta:alta, media:media, baja:baja,
    totLit:Math.round(totLit), totMon:Math.round(totMon),
    tipos:Object.keys(tipos).map(function(k){return {k:k,n:tipos[k]};}).sort(function(a,b){return b.n-a.n;}),
    /* Todos los vehículos (no solo los que tienen observación), para poder
       verificar también los que están bien; ordenados por severidad y luego
       por nombre para facilitar la revisión. */
    vehiculos:Object.keys(vh).map(function(key){
      var v=vh[key];
      var label=Object.keys(v.labelCounts).sort(function(a,b){return v.labelCounts[b]-v.labelCounts[a];})[0];
      v.detalle.sort(function(a,b){return a.fila-b.fila;});
      return {
        placa:label, cargas:v.cargas, lit:Math.round(v.lit), mon:Math.round(v.mon), obs:v.obs,
        sev:(v.sevMax>=3?'Alta':v.sevMax===2?'Media':v.sevMax===1?'Baja':'OK'),
        detalle:v.detalle
      };
    }).sort(function(a,b){
      if(sevOrden[b.sev]!==sevOrden[a.sev]) return sevOrden[b.sev]-sevOrden[a.sev];
      if(b.obs!==a.obs) return b.obs-a.obs;
      return a.placa.localeCompare(b.placa);
    }),
    columnas:Object.keys(col).filter(function(k){return col[k]>=0;}),
    muestra:flag.slice(0,15).map(function(f){return {fila:f.fila,fecha:f.fecha,placa:f.placa,litros:f.litros,monto:f.monto,sev:(f.sev>=3?'Alta':f.sev===2?'Media':'Baja'),obs:f.obs.join(' · ')};})
  };
}

function _loadExcelJS(cb){
  if(window.ExcelJS){cb();return;}
  toast('⏳ Preparando motor de Excel…');
  var s=document.createElement('script');
  s.src='https://cdn.jsdelivr.net/npm/exceljs@4.4.0/dist/exceljs.min.js';
  s.onload=function(){cb();};
  s.onerror=function(){toast('⚠ No se pudo cargar el motor de Excel (revisa tu conexión)');};
  document.head.appendChild(s);
}

function generarReporteConsumos(){
  if(!_CONS){toast('⚠ Primero analiza un archivo');return;}
  _loadExcelJS(function(){ _buildConsumosXLSX(_CONS); });
}

function _buildConsumosXLSX(data){
  var filas=data.filas, col=data.col, nombre=data.nombre, C=CONSUMOS_CFG;
  var flag=filas.filter(function(f){return f.obs.length;});
  var HEADER='FF0F172A', WHITE='FFFFFFFF';
  var fillAlta='FFF8CBCB', fillMedia='FFFDE9B8', fillBaja='FFFBF3C0';
  var txtAlta='FF991B1B', txtMedia='FF92400E', txtBaja='FF854D0E';
  var s={style:'thin',color:{argb:'FFE2E8F0'}}; var BORDER={top:s,left:s,bottom:s,right:s};
  var wb=new ExcelJS.Workbook(); wb.creator='Grupo Kuroda'; wb.created=new Date();

  function headerRow(ws,ncols){
    var row=ws.getRow(1); row.height=22;
    for(var i=1;i<=ncols;i++){ var c=row.getCell(i);
      c.fill={type:'pattern',pattern:'solid',fgColor:{argb:HEADER}};
      c.font={bold:true,color:{argb:WHITE},size:11}; c.alignment={vertical:'middle'}; c.border=BORDER; }
  }

  /* ── Observaciones ── */
  var ws=wb.addWorksheet('Observaciones',{views:[{state:'frozen',ySplit:1}]});
  ws.columns=[
    {header:'Fila',key:'fila',width:7},{header:'Fecha',key:'fecha',width:14},
    {header:'Vehículo',key:'placa',width:16},{header:'Combustible',key:'producto',width:22},
    {header:'Conductor',key:'conductor',width:20},
    {header:'Litros',key:'litros',width:9},{header:'Monto',key:'monto',width:12},
    {header:'$/L',key:'pl',width:9},{header:'Odómetro',key:'odom',width:11},
    {header:'Rend. km/L',key:'kmL',width:11},{header:'Severidad',key:'sev',width:11},
    {header:'Observaciones',key:'obs',width:72}
  ];
  flag.forEach(function(f){
    var sev=f.sev>=3?'Alta':f.sev===2?'Media':'Baja';
    var row=ws.addRow({fila:f.fila,fecha:f.fecha,placa:f.placa,producto:f.producto,conductor:f.conductor,
      litros:f.litros,monto:f.monto,pl:(f._pl!=null?+f._pl.toFixed(2):null),odom:f.odom,
      kmL:(f._kmL!=null?+f._kmL.toFixed(1):null),sev:sev,obs:f.obs.join(' · ')});
    var fillC=sev==='Alta'?fillAlta:sev==='Media'?fillMedia:fillBaja;
    var txtC=sev==='Alta'?txtAlta:sev==='Media'?txtMedia:txtBaja;
    row.eachCell(function(c){ c.border=BORDER; c.alignment={vertical:'middle'}; });
    row.getCell('obs').alignment={vertical:'middle',wrapText:true};
    row.getCell('sev').fill={type:'pattern',pattern:'solid',fgColor:{argb:fillC}};
    row.getCell('sev').font={bold:true,color:{argb:txtC}};
    row.getCell('sev').alignment={vertical:'middle',horizontal:'center'};
    row.getCell('obs').fill={type:'pattern',pattern:'solid',fgColor:{argb:fillC}};
    row.getCell('monto').numFmt='"$"#,##0.00'; row.getCell('pl').numFmt='"$"#,##0.00';
  });
  headerRow(ws,12);
  ws.autoFilter='A1:L1';
  if(!flag.length) ws.addRow({fila:'',fecha:'',placa:'',producto:'',conductor:'',litros:'',monto:'',pl:'',odom:'',kmL:'',sev:'',obs:'Sin observaciones: no se detectaron inconsistencias.'});

  /* ── Resumen ── */
  var rs=wb.addWorksheet('Resumen');
  rs.getColumn(1).width=42; rs.getColumn(2).width=22;
  var t=rs.addRow(['REPORTE DE CONSUMOS DE COMBUSTIBLE']); t.font={bold:true,size:15,color:{argb:HEADER}}; rs.mergeCells('A1:B1');
  var t2=rs.addRow(['Grupo Kuroda — análisis de fugas y errores de carga']); t2.font={italic:true,color:{argb:'FF64748B'}}; rs.mergeCells('A2:B2');
  rs.addRow([]);
  var kv=[
    ['Archivo origen', nombre||''],
    ['Generado', new Date().toLocaleString('es-MX')],
    ['Registros analizados', filas.length],
    ['Registros con observación', flag.length],
    ['Severidad Alta (fuga / duplicada / precio)', flag.filter(function(f){return f.sev>=3;}).length],
    ['Severidad Media (odómetro / rendimiento)', flag.filter(function(f){return f.sev===2;}).length],
    ['Severidad Baja (frecuencia)', flag.filter(function(f){return f.sev===1;}).length]
  ];
  kv.forEach(function(p){ var r=rs.addRow(p); r.getCell(1).font={bold:true,color:{argb:'FF334155'}}; });
  var sevColorRows=[{i:5,c:fillAlta},{i:6,c:fillMedia},{i:7,c:fillBaja}];
  sevColorRows.forEach(function(o){ var r=rs.getRow(3+o.i); if(r){ r.getCell(2).fill={type:'pattern',pattern:'solid',fgColor:{argb:o.c}}; r.getCell(2).font={bold:true}; } });
  rs.addRow([]);
  var ch=rs.addRow(['Criterios usados','']); ch.getCell(1).font={bold:true,color:{argb:HEADER}};
  rs.addRow(['Precio $/L esperado', C.PMIN+' a '+C.PMAX]);
  rs.addRow(['Diésel — rend. km/L (mín/máx) · litros máx', C.PERFIL.diesel.REND_MIN+' / '+C.PERFIL.diesel.REND_MAX+'  ·  '+C.PERFIL.diesel.LIT_MAX+' L']);
  rs.addRow(['Gasolina — rend. km/L (mín/máx) · litros máx', C.PERFIL.gasolina.REND_MIN+' / '+C.PERFIL.gasolina.REND_MAX+'  ·  '+C.PERFIL.gasolina.LIT_MAX+' L']);
  rs.addRow(['Sin tipo — rend. km/L (mín/máx) · litros máx', C.REND_MIN+' / '+C.REND_MAX+'  ·  '+C.LIT_MAX+' L']);
  rs.addRow(['Distancia mín. para medir rendimiento (km)', C.DIST_MIN]);

  /* ── Por tipo ── */
  var tipos={}; flag.forEach(function(f){ (f.cats||[]).forEach(function(c){ tipos[c]=(tipos[c]||0)+1; }); });
  var wt=wb.addWorksheet('Por tipo'); wt.columns=[{header:'Tipo de observación',key:'k',width:60},{header:'Ocurrencias',key:'n',width:14}];
  Object.keys(tipos).sort(function(a,b){return tipos[b]-tipos[a];}).forEach(function(k){ wt.addRow({k:k,n:tipos[k]}); });
  headerRow(wt,2);

  /* ── Por vehículo ── */
  var vh={}; filas.forEach(function(f){ if(!f.placa)return; var v=vh[f.placa]=vh[f.placa]||{cargas:0,lit:0,mon:0,obs:0};
    v.cargas++; if(f.litros)v.lit+=f.litros; if(f.monto)v.mon+=f.monto; v.obs+=f.obs.length; });
  var wv=wb.addWorksheet('Por vehículo'); wv.columns=[
    {header:'Vehículo',key:'p',width:18},{header:'Cargas',key:'c',width:10},
    {header:'Litros',key:'l',width:12},{header:'Monto',key:'m',width:14},{header:'Observaciones',key:'o',width:14}];
  Object.keys(vh).sort(function(a,b){return vh[b].obs-vh[a].obs;}).forEach(function(p){
    var r=wv.addRow({p:p,c:vh[p].cargas,l:+vh[p].lit.toFixed(1),m:+vh[p].mon.toFixed(2),o:vh[p].obs});
    r.getCell('m').numFmt='"$"#,##0.00';
    if(vh[p].obs>0){ r.getCell('o').fill={type:'pattern',pattern:'solid',fgColor:{argb:fillAlta}}; r.getCell('o').font={bold:true,color:{argb:txtAlta}}; }
  });
  headerRow(wv,5);

  wb.xlsx.writeBuffer().then(function(buf){
    var blob=new Blob([buf],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
    _descargarBlob(blob,'Reporte_Consumos_Combustible_'+new Date().toISOString().slice(0,10)+'.xlsx');
    toast('✓ Reporte Excel generado');
  }).catch(function(err){ toast('⚠ Error al generar Excel: '+err.message); });
}

/* Barra de datos: Vigentes = auditorías activas que NO están finalizadas
   (mismo criterio que la vista de Auditorías); Finalizadas = archivadas.
   Se llama en refreshAll y también al terminar loadFinalizadas (cargan async). */
function actualizarStrip(){
  var elV=document.getElementById('ds-aud');
  var finList=(typeof FINALIZADAS!=='undefined')?FINALIZADAS.filter(function(f){return !f._pending;}):[];
  var vigentes=(typeof estaFinalizada==='function')
    ? STORE.auditorias.filter(function(a){return !estaFinalizada(a);}).length
    : STORE.auditorias.length;
  if(elV)elV.textContent=vigentes;
  var elF=document.getElementById('ds-fin'); if(elF)elF.textContent=finList.length;
  var elT=document.getElementById('ds-tar'); if(elT)elT.textContent=STORE.tareas.length;
  var elP=document.getElementById('ds-pend'); if(elP)elP.textContent=STORE.tareas.filter(esPendiente).length;
}

function refreshAll(){
  actualizarEstadosVencidos(); /* sincronizar estados antes de KPIs */
  actualizarStrip();
  fillFilters();render();
  /* Si la vista de auditorías está activa, re-renderizarla para reflejar cambios en tareas */
  if(VIEW==='auditorias') renderAuditoriasView();
}
function render(){
  const tareas=filteredTareas(), aud=filteredAuditorias();
  document.getElementById('rec-count').textContent=`${tareas.length} tarea(s) · ${aud.length} auditoría(s)`;
  if(VIEW==='dash')renderDashboard(tareas,aud);
  else if(VIEW==='evaluacion'){if(typeof renderEvaluacion==='function')renderEvaluacion();}
  else if(VIEW==='desempeno'){if(typeof renderDesempeno==='function')renderDesempeno();}
  else if(VIEW==='auditorias'){ renderAuditoriasView(); }
  else renderTareasView(tareas);
}

function renderDashboard(tareas,aud){
  renderKPIs(tareas,aud);
  renderTareasInsight(tareas);
  renderTrend(aud);
  renderResPend(tareas);
  renderCentro(aud);
  renderDonut(tareas);
  renderRanking(aud);
  renderPendRank(tareas);
  renderVencTable(tareas);
}

function renderTareasInsight(tareas){
  const nameOf=t=>(t.nombre&&t.nombre.trim())||(t.actividad&&t.actividad.trim())||'Sin nombre';

  const freq={};
  tareas.forEach(t=>{const k=nameOf(t);freq[k]=(freq[k]||0)+1;});
  const topFreq=Object.entries(freq).sort((a,b)=>b[1]-a[1]).slice(0,6);

  const byName={};
  tareas.forEach(t=>{
    const k=nameOf(t);
    if(!byName[k])byName[k]={total:0,bad:0,good:0};
    byName[k].total++;
    const n=norm(t.estado);
    if(n.includes('abierta')&&n.includes('atrasad'))byName[k].bad++;
    else if(n.includes('resuelta')&&!n.includes('atrasad'))byName[k].good++;
  });

  const peor=Object.entries(byName).sort((a,b)=>b[1].total-a[1].total).slice(0,6);
  const mejor=Object.entries(byName).sort((a,b)=>a[1].total-b[1].total).slice(0,6);
  const rankColors=['#2563eb','#0d9488','#ea580c','#7c3aed','#dc2626','#d97706'];

  function freqRows(arr){
    if(!arr.length)return`<div class="insight-empty">Sin datos</div>`;
    return arr.map(([name,cnt],i)=>`
      <div class="insight-row">
        <span class="insight-rank" style="background:${rankColors[i]}">${i+1}</span>
        <span class="insight-name" title="${name}">${name}</span>
        <span class="insight-chip" style="background:${rankColors[i]}18;color:${rankColors[i]}">${cnt}x</span>
      </div>`).join('');
  }

  function simpleRows(arr,col,colPale){
    if(!arr.length)return`<div class="insight-empty">Sin datos</div>`;
    return arr.map(([k,d],i)=>`
      <div class="insight-row">
        <span class="insight-rank" style="background:${col}">${i+1}</span>
        <span class="insight-name" title="${k}">${k}</span>
        <span class="insight-chip" style="background:${colPale};color:${col}">${d.total}x</span>
      </div>`).join('');
  }

  const grid=document.getElementById('insight-row');
  if(grid){
    grid.style.gridTemplateColumns='repeat(3,1fr)';
    grid.innerHTML=`
      <div class="insight-card">
        <div class="insight-card-hdr"><span class="ico">🔁</span><span class="ttl">Tareas más frecuentes</span></div>
        ${freqRows(topFreq)}
      </div>
      <div class="insight-card">
        <div class="insight-card-hdr"><span class="ico">🔴</span><span class="ttl">Peor calificadas</span><span style="font-size:9.5px;color:var(--muted);margin-left:auto;font-weight:600">mayor frecuencia</span></div>
        ${simpleRows(peor,'#dc2626','#fee2e2')}
      </div>
      <div class="insight-card">
        <div class="insight-card-hdr"><span class="ico">✅</span><span class="ttl">Mejor calificadas</span><span style="font-size:9.5px;color:var(--muted);margin-left:auto;font-weight:600">menor frecuencia</span></div>
        ${simpleRows(mejor,'#16a34a','#dcfce7')}
      </div>`;
  }
}

function renderKPIs(tareas,aud){
  const avgCumpl=aud.length?aud.reduce((a,r)=>a+r.pctCumpl,0)/aud.length:0;
  const total=tareas.length;
  const res=tareas.filter(esResuelta).length;
  const pend=tareas.filter(esPendiente).length;
  const vencidas=tareas.filter(t=>esPendiente(t)&&diasVenc(t)!==null&&diasVenc(t)<0).length;
  const pctRes=total?res/total:0;

  // delta de tendencia: comparar último bucket vs anterior (% cumplimiento)
  const f=getFilterState();
  /* Fallback: si `a.fecha` no se parsea (dato común), se usa `a.mes` con el año
     actual, igual que en la gráfica de tendencia. Sin esto salía "sin tendencia". */
  const _anioK=new Date().getFullYear();
  const _fechaK=a=>{const d=fromISO(a.fecha);if(d)return d;const mi=(typeof mesIndexFromNombre==='function')?mesIndexFromNombre(a.mes):-1;return mi>=0?new Date(_anioK,mi,15):null;};
  const buckets={};
  aud.forEach(a=>{const d=_fechaK(a);const b=bucketKey(d,f.gran);if(b){(buckets[b.k]=buckets[b.k]||{sort:b.sort,s:0,n:0});buckets[b.k].s+=a.pctCumpl;buckets[b.k].n++;}});
  const ord=Object.values(buckets).sort((a,b)=>a.sort-b.sort);
  let delta=null;
  if(ord.length>=2){const cur=ord[ord.length-1].s/ord[ord.length-1].n,prev=ord[ord.length-2].s/ord[ord.length-2].n;delta=(cur-prev)*100;}

  const deltaHtml=delta===null?`<span class="delta flat">— sin tendencia</span>`:
    delta>=0?`<span class="delta up">▲ +${delta.toFixed(1)} pts</span> vs período previo`:
    `<span class="delta down">▼ ${delta.toFixed(1)} pts</span> vs período previo`;

  // distribución de estados (mismos conteos que el donut)
  const resOk=tareas.filter(t=>esResuelta(t)&&!norm(t.estado).includes('atrasad')).length;
  const resAtr=tareas.filter(t=>esResuelta(t)&&norm(t.estado).includes('atrasad')).length;
  const abOk=tareas.filter(t=>esPendiente(t)&&!norm(t.estado).includes('atrasad')).length;
  const abAtr=tareas.filter(t=>esPendiente(t)&&norm(t.estado).includes('atrasad')).length;

  const kpis=[
    {c:'k-blue',ico:'📊',lbl:'Cumplimiento prom.',val:Math.round(avgCumpl*100)+'%',sub:deltaHtml},
    {c:'k-teal',ico:'📋',lbl:'Tareas en período',val:total,sub:`${aud.length} auditorías`},
    {c:'k-green',ico:'✅',lbl:'Resueltas',val:res,sub:pctStr(pctRes)+' del total'},
    {c:'k-orange',ico:'⏳',lbl:'Pendientes',val:pend,sub:total?pctStr(pend/total)+' sin cerrar':'—'},
    {c:'k-red',ico:'🚨',lbl:'Pend. vencidas',val:vencidas,sub:'fuera de fecha de término'},
    {c:'k-teal',ico:'🎯',lbl:'% Resolución',val:Math.round(pctRes*100)+'%',sub:'tareas cerradas'},
    {c:'k-blue',ico:'🏬',lbl:'Sucursales',val:uniq(tareas.map(t=>t.tienda)).length,sub:'con tareas en período'},
    {c:'k-orange',ico:'📈',lbl:'Cumpl. ponderado',val:Math.round((aud.reduce((a,r)=>a+r.pctCumpl*(r.tareas||1),0)/(aud.reduce((a,r)=>a+(r.tareas||1),0)||1))*100)+'%',sub:'por nº de tareas'},
    {c:'k-green',ico:'🟢',lbl:'Resueltas a tiempo',val:resOk,sub:total?pctStr(resOk/total)+' del total':'—'},
    {c:'k-orange',ico:'🟠',lbl:'Resueltas atrasadas',val:resAtr,sub:total?pctStr(resAtr/total)+' del total':'—'},
    {c:'k-blue',ico:'🔵',lbl:'Abiertas en plazo',val:abOk,sub:total?pctStr(abOk/total)+' del total':'—'},
    {c:'k-red',ico:'🔴',lbl:'Abiertas atrasadas',val:abAtr,sub:total?pctStr(abAtr/total)+' del total':'—'},
  ];
  // render only selected KPIs
  const sel=loadKpiSelection()||[];
  const visible=kpis.filter(k=>sel.includes(k.lbl));
  document.getElementById('kpi-row').innerHTML=(visible.length?visible:kpis).map(k=>`
    <div class="kpi-card ${k.c}">
      <div class="kpi-top"><div class="kpi-lbl">${k.lbl}</div><div class="kpi-ico">${k.ico}</div></div>
      <div class="kpi-val">${k.val}</div>
      <div class="kpi-sub">${k.sub}</div>
    </div>`).join('');
  // store catalog for the config modal
  window._kpiCatalog=kpis;
}

/* ══════════════════ KPI CONFIG ══════════════════ */
const KPI_CFG_KEY='cerezo_kpi_selection';
function loadKpiSelection(){
  try{const v=localStorage.getItem(KPI_CFG_KEY);return v?JSON.parse(v):null;}catch{return null;}
}
function saveKpiCfgToStorage(labels){
  try{localStorage.setItem(KPI_CFG_KEY,JSON.stringify(labels));}catch{}
}
function openKpiCfg(){
  const catalog=window._kpiCatalog||[];
  const sel=loadKpiSelection()||catalog.map(k=>k.lbl);
  const body=document.getElementById('kpi-cfg-body');
  // group by category
  const groups=[
    {lbl:'Cumplimiento',keys:['Cumplimiento prom.','Cumpl. ponderado','% Resolución']},
    {lbl:'Volumen de tareas',keys:['Tareas en período','Resueltas','Pendientes','Pend. vencidas','Sucursales']},
    {lbl:'Distribución de estado',keys:['Resueltas a tiempo','Resueltas atrasadas','Abiertas en plazo','Abiertas atrasadas']},
  ];
  let html='';
  groups.forEach(g=>{
    html+=`<div class="kpi-cfg-sep">${g.lbl}</div>`;
    g.keys.forEach(key=>{
      const kpi=catalog.find(k=>k.lbl===key);
      if(!kpi)return;
      const chk=sel.includes(key);
      html+=`<label class="kpi-cfg-item${chk?' checked':''}" onclick="toggleKpiItem(this)">
        <input type="checkbox" ${chk?'checked':''} data-key="${key}" onclick="event.stopPropagation()">
        <span class="kpi-cfg-ico">${kpi.ico}</span>
        <div><div class="kpi-cfg-txt">${key}</div></div>
      </label>`;
    });
  });
  body.innerHTML=html;
  document.getElementById('kpi-cfg-overlay').classList.add('show');
}
function closeKpiCfg(){document.getElementById('kpi-cfg-overlay').classList.remove('show');}
function toggleKpiItem(el){
  const cb=el.querySelector('input');
  cb.checked=!cb.checked;
  el.classList.toggle('checked',cb.checked);
}
function kpiCfgSelectAll(v){
  document.querySelectorAll('#kpi-cfg-body input[type=checkbox]').forEach(cb=>{
    cb.checked=v;cb.closest('.kpi-cfg-item').classList.toggle('checked',v);
  });
}
function saveKpiCfg(){
  const selected=[...document.querySelectorAll('#kpi-cfg-body input[type=checkbox]')]
    .filter(cb=>cb.checked).map(cb=>cb.dataset.key);
  saveKpiCfgToStorage(selected.length?selected:(window._kpiCatalog||[]).map(k=>k.lbl));
  closeKpiCfg();
  render();
}

function destroyChart(id){if(charts[id]){charts[id].destroy();charts[id]=null;}}
/* Garantiza que el canvas exista; lo recrea si fue destruido por un "empty" previo */
function ensureCanvas(canvasId){
  var cv=document.getElementById(canvasId);
  if(cv)return cv;
  /* Buscar el chart-wrap que originalmente contenía este canvas */
  var wraps=document.querySelectorAll('.chart-wrap');
  for(var i=0;i<wraps.length;i++){
    if(wraps[i].getAttribute('data-canvas')===canvasId){
      wraps[i].innerHTML='<canvas id="'+canvasId+'"></canvas>';
      return document.getElementById(canvasId);
    }
  }
  return null;
}
/* Muestra mensaje vacío SIN destruir el canvas permanentemente */
function showEmpty(canvasId,msg){
  var cv=document.getElementById(canvasId);
  var wrap=cv?cv.parentElement:null;
  if(!wrap){
    var wraps=document.querySelectorAll('.chart-wrap');
    for(var i=0;i<wraps.length;i++){if(wraps[i].getAttribute('data-canvas')===canvasId){wrap=wraps[i];break;}}
  }
  if(wrap)wrap.innerHTML='<div class="empty">'+msg+'</div>';
}

function renderTrend(aud){
  const f=getFilterState();
  /* La tendencia agrupa por FECHA de la auditoría. Pero en algunos datasets la
     auditoría no trae `fecha` parseable (viene solo el nombre del mes), y antes
     eso dejaba la gráfica vacía aunque el resto del dashboard sí tenía datos.
     Fallback: si no hay fecha usable, se arma la fecha desde `a.mes` con un año
     de referencia (el de la primera auditoría con fecha válida; si ninguna la
     tiene, el año actual). */
  let baseYear=null;
  aud.forEach(a=>{if(baseYear==null){const d=fromISO(a.fecha);if(d)baseYear=d.getFullYear();}});
  if(baseYear==null)baseYear=new Date().getFullYear();
  const fechaAud=a=>{
    const d=fromISO(a.fecha); if(d)return d;
    const mi=mesIndexFromNombre(a.mes); if(mi>=0)return new Date(baseYear,mi,1);
    return null;
  };
  let gran=f.gran;
  const uniqueMonths=new Set(aud.map(a=>{const d=fechaAud(a);return d?d.getFullYear()+'-'+d.getMonth():null}).filter(Boolean));
  if(uniqueMonths.size<=1&&gran==='month')gran='day';
  const buckets={};
  aud.forEach(a=>{const d=fechaAud(a);const b=bucketKeyEx(d,gran);if(!b)return;
    if(!buckets[b.k])buckets[b.k]={lbl:b.lbl,sort:b.sort,s:0,n:0,res:0,pen:0};
    buckets[b.k].s+=a.pctCumpl;buckets[b.k].n++;});
  const ord=Object.values(buckets).sort((a,b)=>a.sort-b.sort);
  const labels=ord.map(b=>b.lbl);
  const data=ord.map(b=>Math.round(b.s/b.n*100));
  const note=document.getElementById('trend-note');
  if(note)note.textContent=ord.length?`${ord.length} período(s)`:'';
  destroyChart('trend');
  if(!labels.length){showEmpty('chart-trend','Sin auditorías en el rango seleccionado.');return;}
  const ctx=ensureCanvas('chart-trend');
  if(!ctx){return;}
  charts.trend=new Chart(ctx,{type:'line',
    data:{labels,datasets:[{data,borderColor:'#2563eb',backgroundColor:'rgba(37,99,235,.12)',
      fill:true,tension:.35,pointRadius:5,pointBackgroundColor:'#2563eb',borderWidth:2.5}]},
    options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false},
      tooltip:{callbacks:{label:c=>c.raw+'% cumplimiento'}}},
      scales:{y:{min:0,max:100,ticks:{callback:v=>v+'%',font:{size:10}},grid:{color:'#f0f3f7'}},
        x:{ticks:{font:{size:10}},grid:{display:false}}}}});
}

function renderResPend(tareas){
  const f=getFilterState();
  const buckets={};
  tareas.forEach(t=>{const d=fromISO(t.fechaCreacion);const b=bucketKey(d,f.gran);if(!b)return;
    if(!buckets[b.k])buckets[b.k]={lbl:b.lbl,sort:b.sort,res:0,pen:0};
    if(esResuelta(t))buckets[b.k].res++;else if(esPendiente(t))buckets[b.k].pen++;});
  const ord=Object.values(buckets).sort((a,b)=>a.sort-b.sort);
  const labels=ord.map(b=>b.lbl);
  destroyChart('respend');
  if(!labels.length){showEmpty('chart-respend','Sin tareas en el rango seleccionado.');return;}
  const ctx=ensureCanvas('chart-respend');
  if(!ctx){return;}
  charts.respend=new Chart(ctx,{type:'bar',
    data:{labels,datasets:[
      {label:'Resueltas',data:ord.map(b=>b.res),backgroundColor:'#16a34a',borderRadius:5,stack:'a'},
      {label:'Pendientes',data:ord.map(b=>b.pen),backgroundColor:'#dc2626',borderRadius:5,stack:'a'}]},
    options:{responsive:true,maintainAspectRatio:false,
      plugins:{legend:{display:true,position:'top',labels:{font:{size:10},boxWidth:12,padding:8}}},
      scales:{x:{stacked:true,ticks:{font:{size:10}},grid:{display:false}},
        y:{stacked:true,beginAtZero:true,ticks:{font:{size:10},stepSize:1},grid:{color:'#f0f3f7'}}}}});
}

function renderCentro(aud){
  const byC={};
  aud.forEach(a=>{const k=a.centro||a.tienda||'—';if(!byC[k])byC[k]={s:0,n:0};byC[k].s+=a.pctCumpl;byC[k].n++;});
  const ent=Object.entries(byC).map(([k,d])=>({k,v:Math.round(d.s/d.n*100)})).sort((a,b)=>b.v-a.v);
  destroyChart('centro');
  if(!ent.length){showEmpty('chart-centro','Sin datos.');return;}
  const ctx=ensureCanvas('chart-centro');
  if(!ctx){return;}
  charts.centro=new Chart(ctx,{type:'bar',
    data:{labels:ent.map(e=>e.k),datasets:[{data:ent.map(e=>e.v),
      backgroundColor:ent.map(e=>e.v>=70?'#16a34a':e.v>=50?'#d97706':'#dc2626'),borderRadius:6}]},
    options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false},
      tooltip:{callbacks:{label:c=>c.raw+'% cumplimiento'}}},
      scales:{y:{min:0,max:100,ticks:{callback:v=>v+'%',font:{size:10}},grid:{color:'#f0f3f7'}},
        x:{ticks:{font:{size:9}},grid:{display:false}}}}});
}

function renderDonut(tareas){
  const resOk=tareas.filter(t=>esResuelta(t)&&!norm(t.estado).includes('atrasad')).length;
  const resAtr=tareas.filter(t=>esResuelta(t)&&norm(t.estado).includes('atrasad')).length;
  const abAtr=tareas.filter(t=>esPendiente(t)&&norm(t.estado).includes('atrasad')).length;
  const ab=tareas.filter(t=>esPendiente(t)&&!norm(t.estado).includes('atrasad')).length;
  const segs=[['Resueltas a tiempo','#16a34a',resOk],['Resueltas atrasadas','#ea580c',resAtr],
    ['Abiertas en plazo','#2563eb',ab],['Abiertas atrasadas','#dc2626',abAtr]];
  destroyChart('donut');
  const ctx=document.getElementById('chart-donut');
  if(!ctx){return;}
  charts.donut=new Chart(ctx,{type:'doughnut',
    data:{labels:segs.map(s=>s[0]),datasets:[{data:segs.map(s=>s[2]),
      backgroundColor:segs.map(s=>s[1]),borderWidth:3,borderColor:'#fff'}]},
    options:{responsive:true,maintainAspectRatio:false,cutout:'68%',
      plugins:{legend:{display:false},tooltip:{callbacks:{label:c=>`${c.label}: ${c.raw}`}}}}});
  document.getElementById('donut-legend').innerHTML=segs.map(s=>
    `<div class="legend-item"><span class="legend-dot" style="background:${s[1]}"></span>
      <span class="legend-lbl">${s[0]}</span><span class="legend-val">${s[2]}</span></div>`).join('');
}

function renderRanking(aud){
  const byS={};
  aud.forEach(a=>{const k=a.tienda||a.centro;if(!byS[k])byS[k]={s:0,n:0,centro:a.centro};byS[k].s+=a.pctCumpl;byS[k].n++;});
  const arr=Object.entries(byS).map(([t,d])=>({t,centro:d.centro,v:d.s/d.n})).sort((a,b)=>a.v-b.v).slice(0,7);
  const el=document.getElementById('rank-list');
  if(!arr.length){el.innerHTML='<div class="empty">Sin auditorías.</div>';return;}
  el.innerHTML=arr.map((s,i)=>{
    const pct=Math.round(s.v*100);const c=pct>=70?'#16a34a':pct>=50?'#d97706':'#dc2626';
    return `<div class="rank-item">
      <div class="rank-num" style="background:${c}">${i+1}</div>
      <div class="rank-name">${s.t}<small>${s.centro||''}</small></div>
      <div class="rank-bar-wrap"><div class="rank-bg"><div class="rank-fill" style="width:${pct}%;background:${c}"></div></div>
        <span class="rank-pct" style="color:${c}">${pct}%</span></div></div>`;
  }).join('');
}

function renderPendRank(tareas){
  const byS={};
  tareas.filter(esPendiente).forEach(t=>{const k=t.tienda;byS[k]=(byS[k]||0)+1;});
  const max=Math.max(1,...Object.values(byS));
  const arr=Object.entries(byS).map(([t,n])=>({t,n})).sort((a,b)=>b.n-a.n).slice(0,7);
  const el=document.getElementById('pend-rank');
  if(!arr.length){el.innerHTML='<div class="empty">✅ Sin tareas pendientes en el período.</div>';return;}
  el.innerHTML=arr.map((s,i)=>`<div class="rank-item">
    <div class="rank-num" style="background:#ea580c">${i+1}</div>
    <div class="rank-name">${s.t}</div>
    <div class="rank-bar-wrap"><div class="rank-bg"><div class="rank-fill" style="width:${s.n/max*100}%;background:#ea580c"></div></div>
      <span class="rank-pct" style="color:#ea580c">${s.n}</span></div></div>`).join('');
}

function vencInfo(t){
  const dv=diasVenc(t);
  if(dv===null)return{txt:'Sin fecha',cls:'b-gray',color:'#94a3b8',ord:9999};
  if(dv<0)return{txt:`Vencida ${Math.abs(dv)}d`,cls:'b-red',color:'#dc2626',ord:dv};
  if(dv<=7)return{txt:`${dv}d restantes`,cls:'b-orange',color:'#ea580c',ord:dv};
  return{txt:`${dv}d restantes`,cls:'b-blue',color:'#2563eb',ord:dv};
}
function renderVencTable(tareas){
  const pend=tareas.filter(esPendiente).map(t=>({t,v:vencInfo(t)})).sort((a,b)=>a.v.ord-b.v.ord).slice(0,25);
  const el=document.getElementById('venc-table');
  if(!pend.length){el.innerHTML='<div class="empty">✅ No hay tareas pendientes para el filtro actual.</div>';return;}
  el.innerHTML=`<table class="dt">
    <thead><tr><th>#</th><th>Tarea</th><th>Sucursal</th><th>Área</th><th>Tipo</th><th>F. Término</th><th class="c">Vencimiento</th></tr></thead>
    <tbody>${pend.map(({t,v},i)=>`<tr>
      <td><span class="rank-pill" style="background:${v.color}">${i+1}</span></td>
      <td><div class="tname" style="max-width:320px;overflow:hidden;text-overflow:ellipsis">${t.nombre}</div></td>
      <td class="tname">${t.tienda}<div class="tsub">${t.centro}</div></td>
      <td class="tsub">${t.areaResp}</td>
      <td><span class="badge b-gray">${tipoNorm(t.tipoTarea)==='ol'?'O&L':tipoNorm(t.tipoTarea)==='cartera'?'Cartera':'Colab.'}</span></td>
      <td class="tsub">${fmtDate(fromISO(t.fechaTerm))}</td>
      <td class="cell-c"><span class="badge ${v.cls}">${v.txt}</span></td>
    </tr>`).join('')}</tbody></table>`;
}

/* ════════════════════════════════════════════════════════════════════
   VISTA TAREAS (editable)
════════════════════════════════════════════════════════════════════ */
function renderTareasView(tareas){
  document.getElementById('tareas-count').textContent=`${tareas.length} tarea(s)`;
  const el=document.getElementById('tareas-table');
  if(!tareas.length){el.innerHTML='<div class="empty">Sin tareas para el filtro actual.</div>';return;}
  const sorted=[...tareas].sort((a,b)=>(fromISO(b.fechaCreacion)||0)-(fromISO(a.fechaCreacion)||0));
  el.innerHTML=`<table style="width:100%;border-collapse:collapse;font-size:12px">
    <thead><tr style="background:#1a1f3c;color:#fff;font-size:10.5px;text-transform:uppercase;letter-spacing:.04em">
      <th style="padding:7px 8px;font-weight:700;white-space:nowrap">ID</th>
      <th style="padding:7px 8px;font-weight:700;text-align:left">Tarea / Actividad</th>
      <th style="padding:7px 8px;font-weight:700;text-align:left">Sucursal</th>
      <th style="padding:7px 8px;font-weight:700;text-align:left">Área</th>
      <th style="padding:7px 8px;font-weight:700;text-align:center">Tipo</th>
      <th style="padding:7px 8px;font-weight:700;text-align:center">Estado</th>
      <th style="padding:7px 8px;font-weight:700;text-align:center;white-space:nowrap">F. Término</th>
      <th style="padding:7px 8px;font-weight:700;text-align:center;white-space:nowrap">F. Cumpl.</th>
      <th style="padding:7px 8px;font-weight:700;text-align:center">✎</th>
    </tr></thead>
    <tbody>${sorted.map((t,i)=>{
      const bg=i%2===0?'var(--bg)':'var(--soft)';
      const tipo=tipoNorm(t.tipoTarea)==='ol'?'O&L':tipoNorm(t.tipoTarea)==='cartera'?'Cart.':'Colab.';
      const tipoCls=tipoNorm(t.tipoTarea)==='ol'?'#16a34a':tipoNorm(t.tipoTarea)==='cartera'?'#7c3aed':'#2563eb';
      const nEst=norm(t.estado||'');
      const estCol=nEst.includes('atrasad')?'#dc2626':nEst.includes('resuelta')&&nEst.includes('atrasad')?'#d97706':nEst.includes('resuelta')?'#16a34a':'#2563eb';
      const estTxt=nEst.includes('abierta')&&nEst.includes('atrasad')?'Ab. Atr.':nEst.includes('abierta')?'Abierta':nEst.includes('resuelta')&&nEst.includes('atrasad')?'Res. Atr.':'Resuelta';
      const ftColor=fromISO(t.fechaTerm)&&fromISO(t.fechaTerm)<new Date()&&esPendiente(t)?'#dc2626':'var(--muted)';
      return `<tr style="background:${bg};border-bottom:1px solid var(--rowline)">
        <td style="padding:5px 8px;color:var(--muted);font-size:11px;white-space:nowrap;font-family:monospace">${t.id}</td>
        <td style="padding:5px 8px;max-width:300px">
          <div style="font-weight:700;font-size:12px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="${(t.nombre||t.actividad||'').replace(/"/g,'&quot;')}">${t.nombre||t.actividad||'—'}</div>
          <div style="font-size:10px;color:var(--muted);overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${(t.nombre&&t.actividad&&norm(t.actividad)!==norm(t.nombre))?t.actividad:''}</div>
        </td>
        <td style="padding:5px 8px;max-width:160px">
          <div style="font-weight:700;font-size:11.5px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${t.tienda||'—'}</div>
          <div style="font-size:10px;color:var(--muted)">${t.centro||''}</div>
        </td>
        <td style="padding:5px 8px;font-size:11px;color:var(--muted);max-width:120px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${t.areaResp||'—'}</td>
        <td style="padding:5px 8px;text-align:center"><span style="font-size:10px;font-weight:700;color:${tipoCls};background:${tipoCls}18;padding:2px 6px;border-radius:8px">${tipo}</span></td>
        <td style="padding:5px 8px;text-align:center"><span style="font-size:10px;font-weight:700;color:${estCol};background:${estCol}18;padding:2px 6px;border-radius:8px;white-space:nowrap">${estTxt}</span></td>
        <td style="padding:5px 8px;text-align:center;font-size:11px;color:${ftColor};white-space:nowrap">${fmtDate(fromISO(t.fechaTerm))||'—'}</td>
        <td style="padding:5px 8px;text-align:center;font-size:11px;color:var(--muted);white-space:nowrap">${fmtDate(fromISO(t.fechaCumpl))||'—'}</td>
        <td style="padding:5px 8px;text-align:center">${editBtn(t.id)}</td>
      </tr>`;
    }).join('')}</tbody></table>`;
}

const ESTADOS=['Abierta','Abierta atrasada','Resuelta','Resuelta Atrasada'];
function editBtn(id){
  if(_session&&['admin','admin_auditor','auditor'].includes(_session.rol)){
    return '<button class="icon-btn" onclick="openEditIfAllowed(\''+String(id).replace(/\'/g,'')+'\')" >✎</button>';
  }
  return '<span style="color:var(--border)">—</span>';
}
function openEditIfAllowed(id){
  if(!_session||!['admin','admin_auditor','auditor'].includes(_session.rol)){  
    toast('⚠ Sin permisos para editar tareas');return;
  }
  openEdit(id);
}
function openEdit(id){
  const t=id?STORE.tareas.find(x=>String(x.id)===String(id)):null;
  const isNew=!t;
  const v=t||{id:'',razon:'KNO',centro:'',tienda:'',areaResp:'',areaRev:'',actividad:'',nombre:'',
    estado:'Abierta',fechaCreacion:toISO(new Date()),fechaTerm:null,fechaCumpl:null,tipoTarea:'TAREAS ORDEN Y LIMPIEZA'};
  const html=`
    <div class="form-grid">
      <div class="form-field"><label>ID</label><input id="e-id" value="${v.id}" ${isNew?'':'readonly'} placeholder="ID único"></div>
      <div class="form-field"><label>Tipo de tarea</label>
        <select id="e-tipo">
          <option value="TAREAS ORDEN Y LIMPIEZA" ${tipoNorm(v.tipoTarea)==='ol'?'selected':''}>Orden y Limpieza</option>
          <option value="TAREAS AUDITORIA COL" ${tipoNorm(v.tipoTarea)==='col'?'selected':''}>Auditoría Colaboración</option>
          <option value="TAREAS CARTERA" ${tipoNorm(v.tipoTarea)==='cartera'?'selected':''}>Cartera</option>
        </select></div>
      <div class="form-field full"><label>Nombre de tarea${v.nombre?'':' <span style="font-weight:400;color:var(--muted)">(vacío — se está mostrando "'+esc(v.actividad||'')+'" como nombre en las listas)</span>'}</label><textarea id="e-nombre" rows="2" placeholder="${esc(v.actividad||'')}">${v.nombre||''}</textarea></div>
      <div class="form-field"><label>Razón</label><input id="e-razon" value="${v.razon||''}"></div>
      <div class="form-field"><label>Centro</label><input id="e-centro" value="${v.centro||''}"></div>
      <div class="form-field"><label>Sucursal / Tienda</label><input id="e-tienda" value="${v.tienda||''}"></div>
      <div class="form-field"><label>Área de responsabilidad</label><input id="e-areaResp" value="${v.areaResp||''}"></div>
      <div class="form-field"><label>Área revisora</label><input id="e-areaRev" value="${v.areaRev||''}"></div>
      <div class="form-field"><label>Actividad <span style="font-weight:400;color:var(--muted)">(tipo de auditoría)</span></label>
        <select id="e-act">
          <option value="AUDITORIA ORDEN Y LIMPIEZA" ${(tipoNormLocal(v.actividad)||tipoNorm(v.tipoTarea))==='ol'?'selected':''}>Auditoría Orden y Limpieza</option>
          <option value="AUDITORIAS DE COLABORACION" ${(tipoNormLocal(v.actividad)||tipoNorm(v.tipoTarea))==='col'?'selected':''}>Auditoría de Colaboración</option>
          <option value="AUDITORIA CARTERA" ${(tipoNormLocal(v.actividad)||tipoNorm(v.tipoTarea))==='cartera'?'selected':''}>Cartera</option>
        </select></div>
      <div class="form-field"><label>Estado <span style="font-weight:400;color:var(--muted)">(automático, según fechas)</span></label>
        <div id="e-estado-preview" style="padding:8px 10px;border:1px solid var(--border);border-radius:8px;font-size:13px;font-weight:600"></div></div>
      <div class="form-field"><label>Fecha de creación</label><input type="date" id="e-fcre" value="${fmtInput(fromISO(v.fechaCreacion))}"></div>
      <div class="form-field"><label>Fecha de término</label><input type="date" id="e-fterm" value="${fmtInput(fromISO(v.fechaTerm))}" oninput="actualizarEstadoPreview()"></div>
      <div class="form-field"><label>Fecha de cumplimiento</label><input type="date" id="e-fcumpl" value="${fmtInput(fromISO(v.fechaCumpl))}" oninput="actualizarEstadoPreview()"></div>
    </div>`;
  const foot=[{label:'Cancelar',cls:'btn-ghost',fn:closeModal}];
  if(!isNew&&_session&&['admin','admin_auditor','auditor'].includes(_session.rol))foot.push({label:'Eliminar',cls:'btn-red',fn:()=>deleteTask(v.id,v.razon)});
  foot.push({label:isNew?'Crear tarea':'Guardar cambios',cls:'btn-blue',fn:()=>saveTask(isNew)});
  openModal(isNew?'➕ Nueva tarea':`✎ Editar tarea <b>#${v.id}</b>`,html,foot);
  actualizarEstadoPreview();
}
/* Recalcula y pinta el badge de Estado del modal de edición de tarea según
   las fechas de término/cumplimiento capturadas — se llama al abrir el modal
   y cada vez que el usuario cambia cualquiera de las dos fechas. */
function actualizarEstadoPreview(){
  var el=document.getElementById('e-estado-preview');
  if(!el)return;
  var ft=dval('e-fterm')?dval('e-fterm')+'T12:00:00':null;
  var fc=dval('e-fcumpl')?dval('e-fcumpl')+'T12:00:00':null;
  var est=estadoAutomatico(ft?toISO(new Date(ft)):null, fc?toISO(new Date(fc)):null);
  el.textContent=est;
  var colores={'Resuelta':'#16a34a','Resuelta Atrasada':'#d97706','Abierta':'#2563eb','Abierta atrasada':'#dc2626'};
  el.style.color=colores[est]||'inherit';
  el.style.borderColor=colores[est]||'var(--border)';
}
function dval(id){const e=document.getElementById(id);return e?e.value.trim():'';}
function ddate(id){const v=dval(id);return v?toISO(new Date(v+'T12:00:00')):null;}
function saveTask(isNew){
  const id=dval('e-id');
  if(!id){toast('⚠ El ID es obligatorio');return;}
  if(!dval('e-nombre')){toast('⚠ El nombre de tarea es obligatorio');return;}
  if(isNew&&STORE.tareas.some(t=>String(t.id)===id)){toast('⚠ Ya existe una tarea con ese ID');return;}
  const fCre=ddate('e-fcre'), fTerm=ddate('e-fterm'), fCumpl=ddate('e-fcumpl');
  const rec={
    id:isNaN(Number(id))?id:Number(id),
    razon:dval('e-razon'),centro:canonCentro(dval('e-centro')),tienda:dval('e-tienda'),
    areaResp:dval('e-areaResp'),areaRev:dval('e-areaRev'),actividad:dval('e-act'),
    nombre:dval('e-nombre'),estado:estadoAutomatico(fTerm,fCumpl),tipoTarea:dval('e-tipo'),
    fechaCreacion:fCre,fechaTerm:fTerm,fechaCumpl:fCumpl
  };
  const idx=STORE.tareas.findIndex(t=>String(t.id)===id);
  if(idx>=0){
    /* Preserve sb_uuid from original record */
    rec.sb_uuid=STORE.tareas[idx].sb_uuid||null;
    STORE.tareas[idx]={...STORE.tareas[idx],...rec};
  } else {
    STORE.tareas.push(rec);
  }
  saveStore();closeModal();refreshAll();toast(isNew?'✓ Tarea creada':'✓ Tarea actualizada');
  /* Sync to Supabase */
  syncTaskToSupabase(rec, isNew);
}
function deleteTask(id,razon){
  STORE.tareas=STORE.tareas.filter(function(t){
    return !(String(t.id)===String(id)&&razKey(t.razon)===razKey(razon));
  });
  saveStore();closeModal();refreshAll();toast('Tarea eliminada');
  /* Delete from Supabase — por clave compuesta ID+RAZÓN */
  deleteTaskFromSupabase(String(id),razon);
}
function openFileIfAllowed(){
  if(!_session||!['admin','admin_auditor','auditor'].includes(_session.rol)){  
    toast('⚠ Sin permisos para cargar archivos');return;
  }
  document.getElementById('file-input').click();
}
async function syncTaskToSupabase(rec, isNew){
  var client=_sb;
  if(!client){try{client=supabase.createClient(SB_URL,SB_KEY,{
    auth:{persistSession:false,autoRefreshToken:false,detectSessionInUrl:false},
    realtime:{enabled:false},
    global:{headers:{'x-client-info':'monitor-cumplimiento'}}
  });}catch(e){return;}}
  var tid=String(rec.id);
  var _rowRaw={
    tarea_id:tid,tarea_key:tareaKey(tid,rec.razon),razon:rec.razon||null,centro:rec.centro||null,
    tienda:rec.tienda||null,area_resp:rec.areaResp||null,area_rev:rec.areaRev||null,
    actividad:rec.actividad||null,nombre:rec.nombre||null,
    tipo_tarea:rec.tipoTarea||null,estado:rec.estado||null,
    fecha_creacion:rec.fechaCreacion?String(rec.fechaCreacion).split('T')[0]:null,
    fecha_term:rec.fechaTerm?String(rec.fechaTerm).split('T')[0]:null,
    fecha_cumpl:rec.fechaCumpl?String(rec.fechaCumpl).split('T')[0]:null
  };
  var row=await encObj(_rowRaw,FIELDS.tareas);
  try{
    var ru, ri;
    if(!isNew){
      /* ── CONTROL DE CONCURRENCIA ──
         Verificar si la tarea fue modificada por otro usuario desde que la cargamos */
      if(rec.sb_uuid&&rec.sb_updated){
        var cc=await client.from('tareas').select('updated_at').eq('id',rec.sb_uuid).single();
        if(cc.data&&cc.data.updated_at&&cc.data.updated_at!==rec.sb_updated){
          var ok=confirm('⚠ Esta tarea fue modificada por otro usuario mientras la editabas.\n\n¿Deseas sobrescribir esos cambios con los tuyos?\n\n(Cancelar = conservar los cambios del otro usuario)');
          if(!ok){
            toast('↩ Cambios descartados. Recarga para ver la versión actual.');
            return;
          }
        }
      }
      /* Try UPDATE by sb_uuid first (most reliable), fallback to tarea_id */
      if(rec.sb_uuid){
        ru=await client.from('tareas').update(row).eq('id',rec.sb_uuid);
      } else {
        ru=await client.from('tareas').update(row).eq('tarea_key',_rowRaw.tarea_key);
      }
      if(ru.error){
        /* Error on update — try insert */
        ri=await client.from('tareas').upsert([row],{onConflict:'tarea_key'});
        if(ri.error)toast('⚠ Supabase: '+ri.error.message);
        else toast('☁️ Tarea sincronizada en Supabase');
      } else {
        /* Check if update actually affected rows via select */
        var chk=await client.from('tareas').select('tarea_id',{count:'exact',head:true}).eq('tarea_key',_rowRaw.tarea_key);
        if(chk.count===0){
          /* Row didn't exist — insert */
          ri=await client.from('tareas').upsert([row],{onConflict:'tarea_key'});
          if(ri.error)toast('⚠ Supabase insert: '+ri.error.message);
          else{toast('☁️ Tarea creada en Supabase');}
        } else {
          toast('☁️ Tarea actualizada en Supabase');
        }
      }
    } else {
      /* Nueva tarea — INSERT */
      ri=await client.from('tareas').upsert([row],{onConflict:'tarea_key'});
      if(ri.error)toast('⚠ Supabase: '+ri.error.message);
      else toast('☁️ Tarea guardada en Supabase');
    }
  }catch(e){toast('⚠ Error Supabase: '+e.message);console.error('syncTaskToSupabase:',e);}
}
async function deleteTaskFromSupabase(taskId,razon){
  var client=_sb;
  if(!client){try{client=supabase.createClient(SB_URL,SB_KEY,{
    auth:{persistSession:false,autoRefreshToken:false,detectSessionInUrl:false},
    realtime:{enabled:false},
    global:{headers:{'x-client-info':'monitor-cumplimiento'}}
  });}catch(e){return;}}
  try{
    /* Delete por clave compuesta ID+RAZÓN (no toca la misma tarea de otra razón) */
    var r=await client.from('tareas').delete().eq('tarea_key',tareaKey(taskId,razon));
    if(r.error)toast('⚠ Supabase delete: '+r.error.message);
    else toast('☁️ Tarea eliminada de Supabase');
  }catch(e){console.warn('deleteTaskFromSupabase:',e);}
}

/* ════════════════════════════════════════════════════════════════════
   PENDIENTES POR SUCURSAL (modal + PNG)
════════════════════════════════════════════════════════════════════ */
function pendientesPorSucursalHTML(){
  const tareas=filteredTareas().filter(esPendiente);
  if(!tareas.length)return '<div class="empty">✅ No hay tareas pendientes con el filtro actual.</div>';
  const byS={};
  tareas.forEach(t=>{(byS[t.tienda]=byS[t.tienda]||[]).push(t);});
  const stores=Object.entries(byS).sort((a,b)=>b[1].length-a[1].length);
  return stores.map(([tienda,list])=>{
    list.sort((a,b)=>vencInfo(a).ord-vencInfo(b).ord);
    const centro=list[0].centro||'';
    return `<div class="store-block">
      <div class="store-block-hdr"><span class="sname">🏬 ${tienda}</span>
        <span class="tsub">${centro}</span><span class="scount">${list.length} pendiente(s)</span></div>
      <div class="tbl-scroll"><table class="dt" style="min-width:0">
        <thead><tr><th>ID</th><th>Tarea</th><th>Área</th><th>Tipo</th><th>Estado</th><th>F. Término</th><th class="c">Vencimiento</th></tr></thead>
        <tbody>${list.map(t=>{const v=vencInfo(t);return `<tr>
          <td><b>${t.id}</b></td>
          <td><div class="tname" style="max-width:300px;overflow:hidden;text-overflow:ellipsis">${t.nombre}</div></td>
          <td class="tsub">${t.areaResp}</td>
          <td><span class="badge b-gray">${tipoNorm(t.tipoTarea)==='ol'?'O&L':tipoNorm(t.tipoTarea)==='cartera'?'Cartera':'Colab.'}</span></td>
          <td>${estadoBadge(t.estado)}</td>
          <td class="tsub">${fmtDate(fromISO(t.fechaTerm))}</td>
          <td class="cell-c"><span class="badge ${v.cls}">${v.txt}</span></td>
        </tr>`;}).join('')}</tbody></table></div></div>`;
  }).join('');
}
function openPendientes(){
  actualizarEstadosVencidos(); /* actualizar Abierta→Abierta atrasada antes de mostrar */
  const html=`<div id="pend-content">${pendientesPorSucursalHTML()}</div>`;
  openModal('🏬 Tareas pendientes por sucursal',html,[
    {label:'🖼️ Descargar PNG',cls:'btn-teal',fn:()=>downloadPendientesPNG()},
    {label:'Cerrar',cls:'btn-ghost',fn:closeModal}
  ]);
}

/* ════════════════════════════════════════════════════════════════════
   EXPORTACIÓN PNG (siempre activa)  ·  render off-screen ancho completo
════════════════════════════════════════════════════════════════════ */
const PNG_CSS=`
  *{box-sizing:border-box;margin:0;padding:0;font-family:'Inter','Segoe UI',sans-serif}
  body{background:#fff;color:#1f2530}
  .wrap{background:#fff;padding:30px 34px;display:inline-block;min-width:920px}
  .hd{border-bottom:2px solid #e7eaef;padding-bottom:14px;margin-bottom:20px}
  .hd h2{font-size:19px;font-weight:800;color:#1f2530}.hd h2 b{color:#2563eb}
  .hd p{font-size:12px;color:#7c8696;margin-top:3px}
  .sec{margin-bottom:26px}
  .sec-t{font-size:12px;font-weight:800;text-transform:uppercase;letter-spacing:.05em;color:#7c8696;
    border-bottom:1.5px solid #e7eaef;padding-bottom:7px;margin-bottom:12px;display:flex;align-items:center;gap:8px}
  .sdot{width:9px;height:9px;border-radius:3px;display:inline-block}
  table{border-collapse:collapse;width:max-content;min-width:100%}
  thead th{font-size:10.5px;font-weight:800;text-transform:uppercase;letter-spacing:.04em;padding:9px 13px;
    text-align:left;background:#f8fafc;border-bottom:2px solid #e7eaef;white-space:nowrap;color:#7c8696}
  thead th.c{text-align:center}
  tbody td{padding:8px 13px;border-bottom:1px solid #f0f3f7;font-size:12px;white-space:nowrap;color:#1f2530}
  .tname{font-weight:700}.tsub{font-size:10px;color:#7c8696}.cell-c{text-align:center}
  .badge{border-radius:6px;padding:3px 9px;font-size:10px;font-weight:800;white-space:nowrap;display:inline-block}
  .b-green{background:#dcfce7;color:#166534}.b-red{background:#fee2e2;color:#991b1b}
  .b-blue{background:#dbeafe;color:#1e40af}.b-orange{background:#ffedd5;color:#9a3412;border:1px solid #fed7aa}.b-gray{background:#f1f3f7;color:#475569}
  .rank-pill{display:inline-flex;align-items:center;justify-content:center;width:23px;height:23px;border-radius:7px;font-size:10px;font-weight:800;color:#fff}
  .store-block{border:1px solid #e7eaef;border-radius:12px;margin-bottom:14px;overflow:hidden}
  .store-block-hdr{display:flex;align-items:center;gap:10px;padding:11px 15px;background:#f8fafc;border-bottom:1px solid #e7eaef}
  .store-block-hdr .sname{font-size:13.5px;font-weight:800}.store-block-hdr .scount{margin-left:auto;font-size:12px;font-weight:800;color:#dc2626}
  .kpi-row{display:grid;grid-template-columns:repeat(4,1fr);gap:12px;margin-bottom:8px}
  .kpi-b{border:1px solid #e7eaef;border-radius:12px;padding:13px 15px;border-left:4px solid #2563eb}
  .kpi-b .l{font-size:10px;font-weight:700;text-transform:uppercase;color:#7c8696;letter-spacing:.04em}
  .kpi-b .v{font-size:24px;font-weight:800;margin-top:5px}
  .kpi-b .s{font-size:10.5px;color:#7c8696;margin-top:2px}
  .ft{margin-top:16px;padding-top:11px;border-top:1px solid #e7eaef;font-size:10px;color:#94a3b8;text-align:right}
  .empty{padding:30px;text-align:center;color:#7c8696;font-size:13px}
`;
async function renderPNG(innerHTML,filename,btnEl){
  if(btnEl){btnEl.disabled=true;}
  toast('Generando imagen…');
  const iframe=document.createElement('iframe');
  iframe.style.cssText='position:fixed;left:-99999px;top:0;border:none;visibility:hidden';
  document.body.appendChild(iframe);
  const doc=iframe.contentDocument;
  doc.open();
  doc.write(`<!DOCTYPE html><html><head><meta charset="UTF-8"><style>${PNG_CSS}</style></head>
    <body><div class="wrap" id="cap">${innerHTML}</div></body></html>`);
  doc.close();
  await new Promise(r=>setTimeout(r,260));
  try{
    const cap=doc.getElementById('cap');
    const w=cap.scrollWidth,h=cap.scrollHeight;
    iframe.style.width=w+'px';iframe.style.height=h+'px';
    await new Promise(r=>setTimeout(r,80));
    const canvas=await html2canvas(cap,{scale:2.4,backgroundColor:'#ffffff',logging:false,
      width:w,height:h,windowWidth:w,windowHeight:h});
    const a=document.createElement('a');
    a.download=filename;a.href=canvas.toDataURL('image/png',1.0);a.click();
    toast('✓ PNG descargado');
  }catch(err){toast('⚠ Error al generar PNG: '+err.message);}
  finally{document.body.removeChild(iframe);if(btnEl)btnEl.disabled=false;}
}
function pngHeader(title){
  const f=getFilterState();
  const today=new Date().toLocaleDateString('es-MX',{day:'2-digit',month:'long',year:'numeric'});
  const parts=[];
  if(f.razon!=='ALL')parts.push('Razón: '+f.razon);
  if(f.centro!=='ALL')parts.push('Centro: '+f.centro);
  if(f.tienda!=='ALL')parts.push('Sucursal: '+f.tienda);
  if(f.desde||f.hasta)parts.push(`Rango: ${f.desde?fmtDate(f.desde):'inicio'} → ${f.hasta?fmtDate(f.hasta):'hoy'}`);
  return `<div class="hd"><h2>${title}</h2><p>${parts.join(' · ')||'Todos los datos'} · Generado el ${today}</p></div>`;
}
function downloadPendientesPNG(){
  renderPNG(pngHeader('🏬 Tareas pendientes por sucursal')+
    `<div class="sec"><div class="sec-t"><span class="sdot" style="background:#ea580c"></span>Detalle por sucursal</div>${pendientesPorSucursalHTML()}</div>`+
    `<div class="ft">📊 Monitor de Cumplimiento — Grupo Cerezo</div>`,
    `pendientes_por_sucursal_${new Date().toISOString().slice(0,10)}.png`);
}
function downloadTareasPNG(){
  const tareas=filteredTareas();
  const inner=document.getElementById('tareas-table').innerHTML;
  renderPNG(pngHeader('📝 Tareas filtradas')+
    `<div class="sec"><div class="sec-t"><span class="sdot" style="background:#2563eb"></span>${tareas.length} tarea(s)</div>${inner}</div>`+
    `<div class="ft">📊 Monitor de Cumplimiento — Grupo Cerezo</div>`,
    `tareas_${new Date().toISOString().slice(0,10)}.png`);
}
function downloadDashboardPNG(){
  const tareas=filteredTareas(),aud=filteredAuditorias();
  const avgCumpl=aud.length?Math.round(aud.reduce((a,r)=>a+r.pctCumpl,0)/aud.length*100):0;
  const res=tareas.filter(esResuelta).length,pend=tareas.filter(esPendiente).length;
  const venc=tareas.filter(t=>esPendiente(t)&&diasVenc(t)!==null&&diasVenc(t)<0).length;
  const kpiHTML=`<div class="kpi-row">
    <div class="kpi-b" style="border-left-color:#2563eb"><div class="l">Cumplimiento prom.</div><div class="v">${avgCumpl}%</div><div class="s">${aud.length} auditorías</div></div>
    <div class="kpi-b" style="border-left-color:#0d9488"><div class="l">Tareas</div><div class="v">${tareas.length}</div><div class="s">en período</div></div>
    <div class="kpi-b" style="border-left-color:var(--k-greenok)"><div class="l">Resueltas</div><div class="v">${res}</div><div class="s">${tareas.length?Math.round(res/tareas.length*100):0}%</div></div>
    <div class="kpi-b" style="border-left-color:#dc2626"><div class="l">Pendientes</div><div class="v">${pend}</div><div class="s">${venc} vencidas</div></div>
  </div>`;
  // ranking de sucursales menor cumplimiento
  const byS={};aud.forEach(a=>{const k=a.tienda;if(!byS[k])byS[k]={s:0,n:0,c:a.centro};byS[k].s+=a.pctCumpl;byS[k].n++;});
  const rk=Object.entries(byS).map(([t,d])=>({t,c:d.c,v:Math.round(d.s/d.n*100)})).sort((a,b)=>a.v-b.v).slice(0,10);
  const rankHTML=rk.length?`<table><thead><tr><th>#</th><th>Sucursal</th><th>Centro</th><th class="c">% Cumplimiento</th></tr></thead>
    <tbody>${rk.map((s,i)=>{const c=s.v>=70?'#16a34a':s.v>=50?'#d97706':'#dc2626';
      return `<tr><td><span class="rank-pill" style="background:${c}">${i+1}</span></td>
      <td class="tname">${s.t}</td><td class="tsub">${s.c||''}</td>
      <td class="cell-c"><span class="badge" style="background:${c}22;color:${c}">${s.v}%</span></td></tr>`;}).join('')}</tbody></table>`:'<div class="empty">Sin auditorías.</div>';
  // pendientes
  const pendHTML=pendientesPorSucursalHTML();
  renderPNG(pngHeader('📊 Resumen de cumplimiento')+
    `<div class="sec"><div class="sec-t"><span class="sdot" style="background:#2563eb"></span>Indicadores</div>${kpiHTML}</div>`+
    `<div class="sec"><div class="sec-t"><span class="sdot" style="background:#dc2626"></span>Sucursales por nivel de cumplimiento</div>${rankHTML}</div>`+
    `<div class="sec"><div class="sec-t"><span class="sdot" style="background:#ea580c"></span>Tareas pendientes por sucursal</div>${pendHTML}</div>`+
    `<div class="ft">📊 Monitor de Cumplimiento — Grupo Cerezo</div>`,
    `resumen_cumplimiento_${new Date().toISOString().slice(0,10)}.png`,
    event&&event.target?event.target.closest('button'):null);
}

/* ════════════════════════════════════════════════════════════════════
   CSV
════════════════════════════════════════════════════════════════════ */
function exportCSV(){
  const rows=filteredTareas();
  if(!rows.length){toast('Sin tareas para exportar');return;}
  const cols=['id','razon','centro','tienda','areaResp','areaRev','actividad','nombre','estado','fechaCreacion','fechaTerm','fechaCumpl','tipoTarea'];
  const head=cols.join(',')+'\n';
  const body=rows.map(r=>cols.map(c=>{
    let v=r[c]; if((c==='fechaCreacion'||c==='fechaTerm'||c==='fechaCumpl'))v=v?fmtDate(fromISO(v)):'';
    return `"${String(v??'').replace(/"/g,'""')}"`;}).join(',')).join('\n');
  const blob=new Blob(['\uFEFF'+head+body],{type:'text/csv;charset=utf-8'});
  const a=document.createElement('a');a.href=URL.createObjectURL(blob);
  a.download='tareas_cerezo.csv';a.click();toast('✓ CSV descargado');
}

/* ════════════════════════════════════════════════════════════════════
   MODAL / TOAST / VISTAS
════════════════════════════════════════════════════════════════════ */
function openModal(title,html,footBtns){
  document.getElementById('modal-title').innerHTML=title;
  document.getElementById('modal-body').innerHTML=html;
  const foot=document.getElementById('modal-foot');
  if(footBtns&&footBtns.length){
    foot.style.display='flex';foot.innerHTML='';
    footBtns.forEach(b=>{const el=document.createElement('button');el.className='btn '+b.cls;el.textContent=b.label;el.onclick=b.fn;foot.appendChild(el);});
  }else foot.style.display='none';
  document.getElementById('modal-overlay').classList.add('show');
}
function openConfirm(title,html,okLabel,okCls,okFn){
  openModal(title,`<p style="font-size:13.5px;color:var(--txt);line-height:1.55">${html}</p>`,
    [{label:'Cancelar',cls:'btn-ghost',fn:closeModal},{label:okLabel,cls:okCls,fn:okFn}]);
}
function closeModal(e){
  if(e&&e.type==='click'&&e.target!==e.currentTarget)return;
  document.getElementById('modal-overlay').classList.remove('show');
}
document.addEventListener('keydown',e=>{if(e.key==='Escape')closeModal();});
let toastT;
function toast(msg){
  const t=document.getElementById('toast');t.textContent=msg;t.classList.add('show');
  clearTimeout(toastT);toastT=setTimeout(()=>t.classList.remove('show'),2600);
}
/* Contenido completo (HTML/CSS/JS) del Generador de Dashboard Ejecutivo,
   codificado en base64 para cargarlo como recurso autocontenido dentro de
   un iframe aislado (view-generador) — así no hay ningún riesgo de choque
   de nombres de función o de estilos con el resto del dashboard. */
/* Decodifica un string base64 a texto UTF-8 correcto (atob() por sí solo
   rompe acentos y emojis, ya que trata cada byte como Latin-1). */
function b64ToUtf8(b64){
  var binary=atob(b64);
  var bytes=new Uint8Array(binary.length);
  for(var i=0;i<binary.length;i++)bytes[i]=binary.charCodeAt(i);
  return new TextDecoder('utf-8').decode(bytes);
}
function setView(v){
  /* Guardia de solo lectura: el rol viewer no accede a herramientas de
     creación (Documentos / Generador), aunque intente forzar la vista. */
  if(_session&&_session.rol==='viewer'&&['documentos','generador'].includes(v)&&!tieneVista(v)){
    toast('⚠ Tu cuenta es de solo lectura');
    return;
  }
  VIEW=v;
  document.getElementById('view-dash').style.display=v==='dash'?'flex':'none';
  document.getElementById('view-tareas').style.display=v==='tareas'?'flex':'none';
  var va=document.getElementById('view-actividades');
  if(va)va.style.display=v==='actividades'?'flex':'none';
  var vau=document.getElementById('view-auditorias');
  if(vau)vau.style.display=v==='auditorias'?'flex':'none';
  var vaj=document.getElementById('view-ajustes');
  if(vaj)vaj.style.display=v==='ajustes'?'flex':'none';
  var vmr=document.getElementById('view-mermas');
  if(vmr)vmr.style.display=v==='mermas'?'flex':'none';
  var vfin=document.getElementById('view-finalizadas');
  if(vfin)vfin.style.display=v==='finalizadas'?'flex':'none';
  var vdesp=document.getElementById('view-desempeno');
  if(vdesp)vdesp.style.display=v==='desempeno'?'block':'none';
  var vev=document.getElementById('view-evaluacion');
  if(vev)vev.style.display=v==='evaluacion'?'flex':'none';
  var vgen=document.getElementById('view-generador');
  if(vgen)vgen.style.display=v==='generador'?'flex':'none';
  var vdoc=document.getElementById('view-documentos');
  if(vdoc)vdoc.style.display=v==='documentos'?'flex':'none';
  document.getElementById('nav-dash').classList.toggle('active',v==='dash');
  document.getElementById('nav-tareas').classList.toggle('active',v==='tareas');
  var na=document.getElementById('nav-actividades');
  if(na)na.classList.toggle('active',v==='actividades');
  var nau=document.getElementById('nav-auditorias');
  if(nau)nau.classList.toggle('active',v==='auditorias');
  var naj=document.getElementById('nav-ajustes');
  if(naj)naj.classList.toggle('active',v==='ajustes');
  var nmr=document.getElementById('nav-mermas');
  if(nmr)nmr.classList.toggle('active',v==='mermas');
  var nfin=document.getElementById('nav-finalizadas');
  if(nfin)nfin.classList.toggle('active',v==='finalizadas');
  var ndesp=document.getElementById('nav-desempeno');
  if(ndesp)ndesp.classList.toggle('active',v==='desempeno');
  var nev=document.getElementById('nav-evaluacion');
  if(nev)nev.classList.toggle('active',v==='evaluacion');
  var ngen=document.getElementById('nav-generador');
  if(ngen)ngen.classList.toggle('active',v==='generador');
  var ndoc=document.getElementById('nav-documentos');
  if(ndoc)ndoc.classList.toggle('active',v==='documentos');
  if(v==='actividades'){ if(!ACTIVIDADES.length) loadActividades(); else{ fillActFilters(); renderActividades(); } }
  else if(v==='auditorias'){
    if(!STORE.auditorias.length) loadDataFromSupabase().then(renderAuditoriasView);
    else{ renderAuditoriasView(); }
  }
  else if(v==='ajustes'){ if(!AJUSTES.length) loadAjustes(); else{ fillAjFilters(); renderAjustes(); } }
  else if(v==='mermas'){ if(!MERMAS.length) loadMermas(); else{ fillMrFilters(); renderMermas(); } }
  else if(v==='finalizadas'){ if(!FINALIZADAS.length) loadFinalizadas(); else renderFinalizadas(); }
  else if(v==='desempeno'){
    /* Cargar todos los módulos que el desempeño necesita antes de renderizar */
    var loaders=[];
    if(!AJUSTES.length) loaders.push(loadAjustes());
    if(!MERMAS.length)  loaders.push(loadMermas());
    if(!ACTIVIDADES.length) loaders.push(loadActividades());
    if(!FINALIZADAS.length) loaders.push(loadFinalizadas());
    if(!CARGAS.length) loaders.push(loadCargas());
    if(loaders.length) Promise.all(loaders).then(renderDesempeno);
    else renderDesempeno();
  }
  else if(v==='evaluacion'){
    var loaders2=[];
    if(!AJUSTES.length) loaders2.push(loadAjustes());
    if(!MERMAS.length) loaders2.push(loadMermas());
    if(!STORE.auditorias.length) loaders2.push(loadDataFromSupabase());
    if(loaders2.length) Promise.all(loaders2).then(renderEvaluacion);
    else renderEvaluacion();
  }
  else if(v==='generador'){
    var ifr=document.getElementById('iframe-generador');
    /* Enviar al generador las razones que la cuenta puede usar. El generador
       ajusta su marca y su historial (localStorage) por razón social. */
    sendRazonesToGenerador();
    if(ifr&&!ifr.getAttribute('src')){
      /* assets/generador.html vive en el mismo origen que este dashboard
         (mismo dominio en GitHub Pages), así que localStorage (usado por el
         Historial de Cargas y la Tendencia del generador) persiste igual
         que antes — ya no hace falta empaquetarlo como base64 embebido. */
      ifr.src='assets/generador.html';
    }
  }
  else if(v==='documentos'){
    var ifd=document.getElementById('iframe-documentos');
    if(ifd&&!ifd.getAttribute('src')){ ifd.src='assets/documentos.html'; }
    syncDocsTheme();
  }
  else render();
}

/* ════════════════════════════════════════════════════════════════════
   MÓDULO AUDITORÍAS (vista por clase, datos de tabla auditorias)
════════════════════════════════════════════════════════════════════ */
function fillAudFilters(){
  /* Alcance por razón social: la razón activa en el filtro global (#f-razon)
     también acota qué meses/tiendas/centros aparecen aquí — antes se listaban
     mezclados los de las 3 razones sin importar cuál estuviera seleccionada. */
  var razFA=(document.getElementById('f-razon')||{}).value||'ALL';
  var audFA=STORE.auditorias.filter(function(a){return razFA==='ALL'||razKey(a.razon)===razKey(razFA);});
  /* Deduplicar meses sin distinguir mayúsculas/minúsculas (en Supabase el
     campo mes está guardado con mezcla de mayúsculas y Formato Título para
     algunos registros) — se muestra una sola vez, en Formato Título. */
  var mesesVistos={};
  audFA.forEach(function(a){
    if(!a.mes)return;
    var k=norm(a.mes);
    if(!mesesVistos[k])mesesVistos[k]=a.mes.charAt(0).toUpperCase()+a.mes.slice(1).toLowerCase();
  });
  var meses=Object.values(mesesVistos).sort(function(a,b){return mesIndexFromNombre(a)-mesIndexFromNombre(b);});
  var tiendas=[...new Set(audFA.map(a=>a.tienda).filter(Boolean))].sort();
  var centros=[...new Set(audFA.map(a=>a.centro).filter(Boolean))].sort();
  var sM=document.getElementById('aud-f-mes'), sT=document.getElementById('aud-f-tienda'), sC=document.getElementById('aud-f-centro');
  if(!sM)return;
  var vM=sM.value,vT=sT.value,vC=sC.value;
  sM.innerHTML='<option value="ALL">Todos</option>'+meses.map(m=>'<option>'+m+'</option>').join('');
  sT.innerHTML='<option value="ALL">Todas</option>'+limpiarOpciones(tiendas).map(t=>'<option>'+t+'</option>').join('');
  sC.innerHTML='<option value="ALL">Todos</option>'+limpiarOpciones(centros).map(c=>'<option>'+c+'</option>').join('');
  sM.value=vM;sT.value=vT;sC.value=vC;
}

function filteredAudByView(){
  var mes=document.getElementById('aud-f-mes').value;
  var tienda=document.getElementById('aud-f-tienda').value;
  var centro=document.getElementById('aud-f-centro').value;
  var razon=(document.getElementById('f-razon')||{}).value||'ALL';
  return STORE.auditorias.filter(function(a){
    if(razon!=='ALL'&&razKey(a.razon)!==razKey(razon))return false;
    if(mes!=='ALL'&&norm(a.mes)!==norm(mes))return false;
    if(tienda!=='ALL'&&norm(a.tienda)!==norm(tienda))return false;
    if(centro!=='ALL'&&norm(a.centro)!==norm(centro))return false;
    return true;
  });
}

function pctFmt(v){if(v==null)return'—';var n=v>1.5?v:v*100;return Math.round(n)+'%';}

/* Mapa mes-nombre → índice 0-based */
var MESES_NOMBRE=['enero','febrero','marzo','abril','mayo','junio','julio','agosto','septiembre','octubre','noviembre','diciembre'];
function mesIndexFromNombre(nombre){
  if(!nombre)return -1;
  var n=nombre.toLowerCase().trim();
  return MESES_NOMBRE.indexOf(n);
}

/* Mapea clase de auditoría → tipoTarea de tareas */
function tipoTareaDeClase(clase){
  var c=norm(clase||'');
  if(c.includes('colaboracion')||c.includes('colab'))return 'col';
  if(c.includes('orden')||c.includes('limpieza'))return 'ol';
  if(c.includes('cartera'))return 'cartera';
  return null;
}
function tipoNormLocal(tipo){
  var t=norm(tipo||'');
  if(t.includes('orden')||t.includes('limpieza'))return 'ol';
  if(t.includes('cartera'))return 'cartera';
  if(t.includes('col')||t.includes('auditoria'))return 'col';
  return '';
}

/* Calcula stats dinámicamente cruzando STORE.tareas con la auditoría.
   Filtra por tienda+centro+tipoTarea(clase); si hay 2+ auditorías de la misma
   tienda/centro/clase también filtra por mes de fechaCreacion. */
/* Compara si dos registros (auditoría/tarea) son de la MISMA tienda.
   El "centro" (KN00, KN07, ...) es un código controlado y estable; el nombre
   de tienda es texto libre y puede capturarse distinto entre registros
   (p. ej. "20 DE NOVIEMBRE" vs "20 DE NOV MENUDEO" para el mismo KN00, o
   "PEÑASCO" capturado sin el acento en algún renglón). Exigir que AMBOS
   campos coincidieran dejaba auditorías "desconectadas" de sus tareas reales
   (sin cálculo en vivo, sin ⟳, con los pendientes mostrados en azul genérico
   en vez de su color real) cada vez que el nombre de tienda no coincidía
   letra por letra. Con centro disponible en los dos lados, es el único
   criterio: es el identificador confiable.
   Cuando falta el centro en alguno de los dos registros, ANTES se comparaba
   el texto de tienda literal — pero eso seguía dejando auditorías
   desconectadas cuando el nombre venía con distinto prefijo/redacción de un
   lado que del otro (p.ej. "PEÑASCO" vs "KN PEÑASCO"). Ahora, en ese caso, se
   deriva el centro desde el nombre de tienda con el mismo diccionario/alias
   que ya usan ajustes y mermas (centroDeTienda), y solo si eso tampoco
   resuelve se cae al texto literal como último recurso. */
function mismaTiendaAud(x,a){
  var cx=norm(x.centro), ca=norm(a.centro);
  if(cx&&ca)return cx===ca;
  var dx=cx||(typeof centroDeTienda==='function'?norm(centroDeTienda(x.tienda||'')):'');
  var da=ca||(typeof centroDeTienda==='function'?norm(centroDeTienda(a.tienda||'')):'');
  if(dx&&da)return dx===da;
  return norm(x.tienda)===norm(a.tienda);
}

function calcAudStats(a, rowsMismaClase){
  var tipoClase = tipoTareaDeClase(a.clase);

  /* Hermanas = TODAS las auditorías del mismo centro+clase que existen
     en STORE, no solo las que se están pintando. Antes se usaba rowsMismaClase
     (ya filtrado por mes de la vista y por estaFinalizada), así que una
     auditoría podía quedarse "sola" y llevarse las tareas de todos los meses:
     KN EXPRESS julio contaba 19 en vez de 10. */
  var audsMismaTienda = (STORE.auditorias||[]).filter(function(x){
    return mismaTiendaAud(x,a) && norm(x.clase)===norm(a.clase);
  });
  if(!audsMismaTienda.length) audsMismaTienda=[a];
  var mesMes = mesIndexFromNombre(a.mes);

  /* Filtrar tareas: mismo centro (ver mismaTiendaAud) + tipo correspondiente a la clase */
  var tt = STORE.tareas.filter(function(t){
    if(!mismaTiendaAud(t,a))return false;
    if(tipoClase && tipoNormLocal(t.tipoTarea)!==tipoClase)return false;
    return true;
  });

  /* Una auditoría solo cuenta las tareas de SU mes (ventana del mes ± GRACIA_DIAS).
     GRACIA_DIAS era 5 y dejaba fuera tareas reales: en la práctica los tickets
     de una auditoría de "ABRIL"/"JUNIO" se crean en Supabase varios días o
     incluso semanas después del cierre del mes (se vieron casos de 1, 2 y
     hasta 15 días de rezago), así que 5 días era insuficiente y auditorías
     con centro perfectamente coincidente quedaban "desconectadas" (sin ⟳, con
     los pendientes en azul genérico) solo por ese margen tan angosto.
     Esto se aplica siempre que la auditoría tenga mes reconocible, haya o no
     hermanas: si no, la única auditoría visible de esa tienda/clase se llevaba
     las tareas de todos los meses.

     Cuando hay 2+ hermanas, cada tarea se asigna a la del mes MÁS CERCANO, así
     que nunca puede contarse dos veces ni depende de ningún total declarado.
     Por eso ampliar el margen es seguro: no revive el bug de "una auditoría
     se lleva tareas de cualquier mes", porque la desambiguación entre
     auditorías vecinas sigue intacta — solo deja de descartar de entrada
     tareas legítimas con unos días/semanas de rezago normal. */
  if(mesMes>=0){
    var GRACIA_DIAS=30;
    var _fa=a.fecha?fromISO(a.fecha):null;
    var anioAud;
    if(_fa){
      anioAud=_fa.getFullYear();
    }else{
      /* Sin fecha propia en la auditoría: asumir el año actual del reloj
         dejaba fuera tareas reales del mismo centro/tipo cuando la auditoría
         era de un año anterior (la ventana quedaba centrada en el año
         equivocado y no matcheaba nada, aunque el centro sí coincidiera).
         Se infiere el año a partir de las tareas reales de ESTE centro+tipo
         cuyo mes de creación coincide por nombre con el mes de la auditoría;
         si hay varias, se usa la más reciente. Solo si no hay ninguna
         coincidencia así, se cae al año actual (comportamiento anterior). */
      var _aniosCand=STORE.tareas.filter(function(t){
        if(!mismaTiendaAud(t,a))return false;
        if(tipoClase&&tipoNormLocal(t.tipoTarea)!==tipoClase)return false;
        var d=fromISO(t.fechaCreacion);
        return d&&d.getMonth()===mesMes;
      }).map(function(t){return fromISO(t.fechaCreacion).getFullYear();});
      anioAud=_aniosCand.length?Math.max.apply(null,_aniosCand):new Date().getFullYear();
    }
    var desdeMes=new Date(anioAud,mesMes,1-GRACIA_DIAS);
    var hastaMes=new Date(anioAud,mesMes+1,GRACIA_DIAS);
    /* Precalcular mes+año de cada hermana con mes reconocible */
    var hermanas=audsMismaTienda.map(function(sib){
      var sm=mesIndexFromNombre(sib.mes);
      if(sm<0)return null;
      var sf=sib.fecha?fromISO(sib.fecha):null;
      var sy=sf?sf.getFullYear():anioAud;
      return {aud:sib,mesIdx:sy*12+sm};
    }).filter(Boolean);
    tt=tt.filter(function(t){
      var d=fromISO(t.fechaCreacion);
      if(!d)return false;
      if(d<desdeMes||d>hastaMes)return false; /* fuera del margen de ESTA auditoría: no es candidata */
      if(hermanas.length<2)return true; /* no hay realmente con quién competir */
      var tMes=d.getFullYear()*12+d.getMonth();
      var mejorDist=Infinity,mejorEsEsta=false;
      hermanas.forEach(function(h){
        var dist=Math.abs(tMes-h.mesIdx);
        if(dist<mejorDist||(dist===mejorDist&&h.aud===a)){mejorDist=dist;mejorEsEsta=(h.aud===a);}
      });
      return mejorEsEsta;
    });
  }

  if(!tt.length){
    if(typeof console!=='undefined'&&console.warn){
      window._diagAudSyncVistos=window._diagAudSyncVistos||{};
      var _dk=norm(a.centro)+'|'+norm(a.tienda)+'|'+norm(a.mes)+'|'+norm(a.clase);
      if(!window._diagAudSyncVistos[_dk]){
        window._diagAudSyncVistos[_dk]=1;
        var _candTipo=(STORE.tareas||[]).filter(function(t){return !tipoClase||tipoNormLocal(t.tipoTarea)===tipoClase;});
        var _candCentro=_candTipo.filter(function(t){return mismaTiendaAud(t,a);});
        var _candTiendaLit=_candTipo.filter(function(t){return norm(t.tienda)===norm(a.tienda);});
        var _centroDerivAud=(typeof centroDeTienda==='function')?centroDeTienda(a.tienda||''):'';
        var _fechasCand=_candCentro.map(function(t){return t.fechaCreacion;}).slice(0,8);
        console.warn('[Sync auditoría↔tareas] "'+a.tienda+'" ('+a.mes+', '+a.clase+') sin tareas emparejadas.'+
          ' fecha auditoría='+JSON.stringify(a.fecha||'(vacía)')+
          ' | centro auditoría='+JSON.stringify(a.centro||'')+
          ' | centro derivado del nombre='+JSON.stringify(_centroDerivAud||'(no resuelto — revisar CENTRO_TIENDA)')+
          ' | tareas del mismo tipo en STORE='+_candTipo.length+
          ' | con mismo centro='+_candCentro.length+' → fechaCreacion de esas: '+JSON.stringify(_fechasCand)+
          ' | con mismo nombre de tienda literal='+_candTiendaLit.length+
          '. Si "centro derivado" da vacío, falta esa tienda en CENTRO_TIENDA. Si hay tareas con mismo centro pero'+
          ' sus fechas quedan lejos del mes de la auditoría, es la ventana de fecha (±5 días del mes), no el centro.');
      }
    }
    return {tareas:a.tareas||0,resueltas:a.resueltas||0,pendientes:a.pendientes||0,
            abiertas:a.pendientes||0,abiertasAtrasadas:0,resueltasAtrasadas:0,
            pctResuelto:a.pctResuelto||0,tieneAtrasada:false,dinamico:false};
  }

  var total = tt.length;
  /* Pendiente = cualquier estado "abierta*"; resuelto = todo lo demás (resuelta,
     resuelta atrasada, cerrada, etc.). Así resueltas + pendientes = total SIEMPRE
     y coincide con la tabla Cartera (que usa !esPendiente). */
  var pendientes = tt.filter(esPendiente).length;
  var resueltas = total - pendientes;
  var resueltasAtrasadas = tt.filter(tareaResueltaAtrasadaReal).length;
  var pendientesTareas = tt.filter(esPendiente);
  var abiertasAtrasadas = pendientesTareas.filter(tareaVencidaPorFecha).length;
  var abiertas = pendientesTareas.length - abiertasAtrasadas;
  var pctResuelto = total>0 ? resueltas/total : 0;

  return {
    tareas:total, resueltas:resueltas, pendientes:pendientes,
    abiertas:abiertas, abiertasAtrasadas:abiertasAtrasadas,
    resueltasAtrasadas:resueltasAtrasadas,
    pctResuelto:pctResuelto, dinamico:true
  };
}

/* Sincroniza TODAS las auditorías contra sus tareas reales, sin depender de
   que el usuario llegue a abrir/filtrar esa fila en el módulo de Auditorías.
   Antes syncAuditoriaDinamico solo corría para las filas que efectivamente se
   pintaban en pantalla (audTablaPorClase/audTablaPorClaseCarteraRender) — una
   auditoría que nadie llegara a mirar se quedaba con números viejos aunque la
   tarea ya se hubiera actualizado. Esto corre en segundo plano en cuanto se
   cargan los datos, cubriendo todas las auditorías del usuario de una vez. */
async function sincronizarTodasAuditorias(){
  var porClase={};
  STORE.auditorias.forEach(function(a){
    var k=norm(a.clase||'');
    (porClase[k]=porClase[k]||[]).push(a);
  });
  var huboFinalizacionNueva=false;
  for(var k in porClase){
    for(var i=0;i<porClase[k].length;i++){
      var a=porClase[k][i];
      var stats=calcAudStats(a,porClase[k]);
      syncAuditoriaDinamico(a,stats);
      if(stats.dinamico&&stats.pendientes===0&&stats.tareas>0){
        await registrarFinalizada(a);
        huboFinalizacionNueva=true;
      }
    }
  }
  /* Si alguna auditoría se acaba de finalizar en este ciclo, hay que volver a
     pintar la vista para que la exclusión de "ya finalizadas" la tome en
     cuenta de inmediato — si no, se sigue viendo en la tabla activa hasta la
     siguiente recarga, aunque ya esté registrada en Finalizadas. */
  if(huboFinalizacionNueva&&VIEW==='auditorias')renderAuditoriasView();
}


/* Cuando calcAudStats calcula en vivo (dinamico=true) y el resultado difiere
   de lo que quedó guardado en el renglón de la auditoría (tareas/pendientes/
   resueltas/pct_resuelto), esos campos se quedaban congelados en Supabase para
   siempre — la pantalla mostraba el número correcto, pero cualquier otra
   lectura directa de la tabla 'auditorias' (reportes, otra vista) seguía
   viendo el valor viejo. Esto persiste el número correcto de vuelta, una sola
   vez por cambio real (con guarda anti-duplicados mientras la escritura está
   en curso). Se llama justo cuando se pinta cada fila de la tabla. */
var _audSyncEnVuelo={};
function syncAuditoriaDinamico(a, stats){
  if(!stats.dinamico)return;
  var pctR=stats.pctResuelto||0;
  var cambio = (a.tareas||0)!==stats.tareas || (a.pendientes||0)!==stats.pendientes ||
               (a.resueltas||0)!==stats.resueltas || Math.abs((a.pctResuelto||0)-pctR)>0.001;
  if(!cambio)return;
  var client=getSbClient();
  if(!client)return;
  var akey=[norm(a.razon||''),norm(a.centro||a.tienda||''),String(a.fecha||'').split('T')[0],norm(a.clase||'')].join('|');
  if(_audSyncEnVuelo[akey])return;
  _audSyncEnVuelo[akey]=true;
  /* Reflejar de inmediato en memoria para no reintentar en el próximo render */
  a.tareas=stats.tareas;a.pendientes=stats.pendientes;a.resueltas=stats.resueltas;a.pctResuelto=pctR;
  var row={tareas:stats.tareas,pendientes:stats.pendientes,resueltas:stats.resueltas,pct_resuelto:pctR};
  var q=client.from('auditorias').update(row);
  q=a.centro?q.ilike('centro',a.centro):q.ilike('tienda',a.tienda);
  q=q.ilike('mes',(a.mes||'')).ilike('clase',a.clase);
  if(a.fecha)q=q.eq('fecha',String(a.fecha).split('T')[0]);
  q.then(function(r){delete _audSyncEnVuelo[akey];if(r.error)console.warn('sync auditoria:',r.error.message);})
   .catch(function(e){delete _audSyncEnVuelo[akey];console.warn('sync auditoria:',e.message);});
}

/* Color del número de pendientes:
   🔴 rojo   → hay abierta atrasada
   🔵 azul   → hay abierta vigente (sin atraso)
   🟡 amarillo → 0 pendientes pero hay resueltas atrasadas
   🟢 verde  → 0 pendientes y sin resueltas atrasadas */
function pendColor(stats){
  if(stats.abiertasAtrasadas>0) return '#dc2626';
  if(stats.abiertas>0)          return '#2563eb';
  if((stats.resueltasAtrasadas||0)>0) return '#d97706';
  return '#16a34a';
}

function audTablaPorClase(titulo,color,rows){
  /* Una auditoría archivada en Finalizadas ya no es "vigente": se excluye para
     que no aparezca (ni se sume) en los dos módulos a la vez. */
  rows=(rows||[]).filter(function(a){return !estaFinalizada(a);});
  var canEditAud=_session&&['admin','admin_auditor','auditor'].includes(_session.rol);
  /* Cache global indexado por clave única (no por título que puede colisionar) */
  window._audItemCache=window._audItemCache||{};
  var filas=rows.map(function(a){
    var stats=calcAudStats(a,rows);
    var pend=stats.pendientes;
    var pc=pendColor(stats);
    /* Persistir en Supabase el número correcto si difiere de lo guardado */
    syncAuditoriaDinamico(a,stats);
    /* Auto-enviar a finalizadas si resuelto al 100% */
    if(stats.dinamico&&stats.pendientes===0&&stats.tareas>0){
      registrarFinalizada(a);
    }
    var mark=stats.dinamico?'<span style="font-size:9px;color:var(--teal);margin-left:3px" title="En tiempo real">⟳</span>':'';
    /* Color de fila según estado leyenda — adaptado a modo oscuro */
    var _isDark=document.documentElement.getAttribute('data-theme')==='dark';
    var rowBg=stats.abiertasAtrasadas>0?(_isDark?'rgba(252,129,129,.10)':'#fff5f5'):
              stats.abiertas>0?(_isDark?'rgba(67,24,255,.12)':'#eff6ff'):
              (stats.resueltasAtrasadas||0)>0?(_isDark?'rgba(251,177,64,.10)':'#fffbeb'):
              (_isDark?'rgba(1,181,116,.10)':'#f0fdf4');
    var rowBd=stats.abiertasAtrasadas>0?'3px solid #dc2626':
              stats.abiertas>0?'3px solid #2563eb':
              (stats.resueltasAtrasadas||0)>0?'3px solid #d97706':'3px solid #16a34a';
    /* Clave única por auditoría */
    var akey=[norm(a.tienda||''),norm(a.mes||''),norm(a.clase||''),a.fecha||''].join('|');
    window._audItemCache[akey]=a;
    var editBtn=canEditAud?'<td style="text-align:center"><button class="icon-btn" onclick="openEditAuditoria(\''+akey+'\')" title="Editar auditoría">✎</button></td>':'<td></td>';
    return '<tr style="background:'+rowBg+';border-left:'+rowBd+'">'+ 
      '<td><b>'+(a.tienda||'—')+'</b></td>'+
      '<td>'+(a.mes||'—')+'</td>'+
      '<td style="text-align:center">'+pctFmt(a.pctCumpl)+'</td>'+
      '<td style="text-align:center;font-weight:800">'+stats.tareas+'</td>'+
      '<td style="text-align:center">'+pend+'</td>'+
      '<td style="text-align:center;font-size:10px">'+
      ((stats.abiertas||0)>0?'<span style="color:var(--k-blue);font-weight:700">'+(stats.abiertas||0)+'↑</span> ':'')+
      ((stats.abiertasAtrasadas||0)>0?'<span style="color:var(--k-red);font-weight:700">'+(stats.abiertasAtrasadas||0)+'🔴</span>':
       (stats.abiertas||0)===0&&(stats.abiertasAtrasadas||0)===0?'<span style="color:var(--k-greenok)">✓</span>':'')+
      '</td>'+
      '<td style="text-align:center">'+stats.resueltas+mark+'</td>'+
      '<td style="text-align:center">'+pctFmt(stats.pctResuelto)+mark+'</td>'+
      editBtn+
    '</tr>';
  }).join('');
  var editTh=canEditAud?'<th class="c"></th>':'<th></th>';
  var leyenda='<span style="font-size:10px;font-weight:600;color:var(--muted);margin-left:8px;display:inline-flex;align-items:center;gap:10px">'+
    '<span><span style="color:var(--k-red);font-weight:800">●</span> Atrasada</span>'+
    '<span><span style="color:var(--k-blue);font-weight:800">●</span> Vigente</span>'+
    '<span><span style="color:var(--k-orange);font-weight:800">●</span> Resuelta c/atraso</span>'+
    '<span><span style="color:var(--k-greenok);font-weight:800">●</span> Al corriente</span></span>';
  var tablaHTML='<div class="slbl" style="margin:4px 0 4px;color:'+color+'">'+
    '<span class="dot" style="background:'+color+'"></span>'+titulo+
    ' <span style="font-weight:600;font-size:10px;color:var(--muted)">('+rows.length+')</span>'+
    leyenda+'</div>'+
    '<div class="card" style="padding:8px 10px;margin-bottom:10px"><div class="tbl-scroll"><table class="dt">'+
    '<thead><tr><th>Tienda</th><th>Mes Auditoría</th><th class="c">% Cumplimiento</th><th class="c">Total</th><th class="c">Pendientes</th><th class="c">Vigentes/Atraso</th><th class="c">Resueltos</th><th class="c">% Resuelto</th>'+editTh+'</tr></thead>'+
    '<tbody>'+(filas||'<tr><td colspan="9" style="text-align:center;color:var(--muted);padding:18px">Sin registros</td></tr>')+'</tbody></table></div></div>';

  return tablaHTML+'<div style="margin-bottom:22px"></div>';
}

/* Wrapper: decide si usar tabla normal o tabla cartera */
function audTablaPorClaseCartera(titulo,color,rows,esCartera){
  if(esCartera)return audTablaPorClaseCarteraRender(titulo,color,rows);
  return audTablaPorClase(titulo,color,rows);
}

/* Tabla especial CARTERA — sin % cumplimiento, solo seguimiento de tareas */
function audTablaPorClaseCarteraRender(titulo,color,rows){
  rows=(rows||[]).filter(function(a){return !estaFinalizada(a);});
  var canEditAud=_session&&['admin','admin_auditor','auditor'].includes(_session.rol);
  window._audItemCache=window._audItemCache||{};
  var filas=rows.map(function(a){
    /* Usar calcAudStats (igual que las demás clases) para heredar el acotado
       por mes/año y la salvaguarda anti-sobre-emparejamiento entre auditorías
       del mismo centro. Antes esta tabla filtraba aparte y nunca se enviaba
       a finalizadas: una auditoría de Cartera con 0 pendientes se quedaba
       "viva" para siempre en este módulo. */
    var stats=calcAudStats(a,rows);
    var total=stats.tareas;
    var pend=stats.pendientes;
    var resu=stats.resueltas;
    var pc=pend>0?'#dc2626':'#16a34a';
    /* Persistir en Supabase el número correcto si difiere de lo guardado */
    syncAuditoriaDinamico(a,stats);
    /* Auto-enviar a finalizadas si resuelto al 100% (igual que audTablaPorClase) */
    if(stats.dinamico&&stats.pendientes===0&&stats.tareas>0){
      registrarFinalizada(a);
    }
    var akey=[norm(a.tienda||''),norm(a.mes||''),norm(a.clase||''),a.fecha||''].join('|');
    window._audItemCache[akey]=a;
    var editBtn=canEditAud?'<td style="text-align:center"><button class="icon-btn" data-akey="'+akey+'" onclick="openEditAuditoriaKey(this)" title="Editar">✎</button></td>':'<td></td>';
    return '<tr>'+
      '<td><b>'+(a.tienda||'—')+'</b></td>'+
      '<td>'+(a.mes||'—')+'</td>'+
      '<td style="text-align:center">'+total+'</td>'+
      '<td style="text-align:center;color:'+pc+';font-weight:800">'+pend+'</td>'+
      '<td style="text-align:center;color:var(--k-greenok);font-weight:700">'+resu+'</td>'+
      editBtn+'</tr>';
  }).join('');
  var editTh=canEditAud?'<th class="c"></th>':'<th></th>';
  var tablaHTML='<div class="slbl" style="margin:4px 0 4px;color:'+color+'">'+
    '<span class="dot" style="background:'+color+'"></span>'+titulo+
    ' <span style="font-weight:600;font-size:10px;color:var(--muted)">('+rows.length+')</span>'+
    '<span style="font-size:10px;font-weight:600;color:var(--k-purple);margin-left:10px">Sin % cumplimiento · solo seguimiento</span></div>'+
    '<div class="card" style="padding:8px 10px;margin-bottom:10px"><div class="tbl-scroll"><table class="dt">'+
    '<thead><tr><th>Tienda</th><th>Mes</th><th class="c">Total Tareas</th>'+
    '<th class="c">Pendientes</th><th class="c">Resueltas</th>'+editTh+'</tr></thead>'+
    '<tbody>'+(filas||'<tr><td colspan="6" style="text-align:center;color:var(--muted);padding:18px">Sin registros</td></tr>')+'</tbody></table></div></div>';
  return tablaHTML+'<div style="margin-bottom:22px"></div>';
}

function openEditAuditoriaKey(el){
  var akey=el.getAttribute('data-akey');
  var a=window._audItemCache&&window._audItemCache[akey];
  if(!a)return;
  openEditAuditoria(akey);
}

/* ── Editar auditoría (admin, señior y jr; eliminar solo admin/señior) ── */
function openEditAuditoria(akey){
  if(!_session||!['admin','admin_auditor','auditor'].includes(_session.rol)){toast('⚠ Sin permisos para editar auditorías');return;}
  var a=(window._audItemCache&&window._audItemCache[akey]);if(!a)return;
  var meses=['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'];
  var html='<div class="form-grid">'+
    '<div class="form-field"><label>Tienda</label><input id="ea-tienda" value="'+esc(a.tienda||'')+'"></div>'+
    '<div class="form-field"><label>Centro</label><input id="ea-centro" value="'+esc(a.centro||'')+'"></div>'+
    '<div class="form-field"><label>Razón</label><select id="ea-razon">'+
      ['KNO','KSC','KSA'].map(function(rz){return'<option value="'+rz+'"'+(norm(a.razon||'')===norm(rz)?' selected':'')+'>'+rz+'</option>';}).join('')+
    '</select></div>'+
    '<div class="form-field full"><label>Clase</label><select id="ea-clase">'+
      '<option value="AUDITORIAS DE COLABORACION"'+(norm(a.clase||'').includes('colab')?' selected':'')+'>Colaboración</option>'+
      '<option value="AUDITORIA ORDEN Y LIMPIEZA"'+(norm(a.clase||'').includes('orden')?' selected':'')+'>Orden y Limpieza</option>'+
      '<option value="AUDITORIA CARTERA"'+(norm(a.clase||'').includes('cartera')?' selected':'')+'>Cartera</option>'+
    '</select></div>'+
    '<div class="form-field"><label>Mes</label><select id="ea-mes">'+meses.map(function(m){return'<option'+(norm(a.mes||'')===norm(m)?' selected':'')+'>'+m+'</option>';}).join('')+'</select></div>'+
    '<div class="form-field"><label>Fecha auditoría</label><input type="date" id="ea-fecha" value="'+(a.fecha||'')+'"></div>'+
    '<div class="form-field"><label>% Cumplimiento (0–100)</label><input type="number" id="ea-pct" min="0" max="100" step="0.1" value="'+Math.round((a.pctCumpl||0)*100)+'"></div>'+
    '<div class="form-field"><label>Total tareas</label><input type="number" id="ea-tareas" min="0" value="'+(a.tareas||0)+'"></div>'+
    '<div class="form-field"><label>Pendientes</label><input type="number" id="ea-pend" min="0" value="'+(a.pendientes||0)+'"></div>'+
    '<div class="form-field"><label>Resueltos</label><input type="number" id="ea-res" min="0" value="'+(a.resueltas||0)+'"></div>'+
  '</div>';
  var _btnsEA=[{label:'Cancelar',cls:'btn-ghost',fn:closeModal}];
  if(['admin','admin_auditor','auditor'].includes(_session.rol))_btnsEA.push({label:'Eliminar',cls:'btn-red',fn:function(){deleteAuditoria(a);}});
  _btnsEA.push({label:'Guardar cambios',cls:'btn-blue',fn:function(){saveEditAuditoria(a);}});
  openModal('✎ Editar auditoría — '+esc(a.tienda||''),html,_btnsEA);
}

async function saveEditAuditoria(original){
  if(!_session||!['admin','admin_auditor','auditor'].includes(_session.rol)){toast('⚠ Sin permisos para editar auditorías');return;}
  function dv(id){var e=document.getElementById(id);return e?e.value.trim():'';}
  var pct=parseFloat(dv('ea-pct'))||0;
  var updated={
    tienda:dv('ea-tienda').toUpperCase(),centro:canonCentro(dv('ea-centro')),
    razon:(dv('ea-razon')||original.razon||'').toUpperCase(),mes:dv('ea-mes').toUpperCase(),
    fecha:dv('ea-fecha'),pctCumpl:pct>1?pct/100:pct,
    tareas:parseInt(dv('ea-tareas'))||0,
    pendientes:parseInt(dv('ea-pend'))||0,
    resueltas:parseInt(dv('ea-res'))||0,
    pctResuelto:parseInt(dv('ea-tareas'))>0?(parseInt(dv('ea-res'))/parseInt(dv('ea-tareas'))):0,
    clase:dv('ea-clase')||original.clase
  };
  var idx=STORE.auditorias.findIndex(function(x){
    return (original.centro?mismaTiendaAud(x,original):norm(x.tienda)===norm(original.tienda))&&
      norm(x.mes)===norm(original.mes)&&norm(x.clase)===norm(original.clase)&&(!original.fecha||x.fecha===original.fecha);
  });
  if(idx>=0)STORE.auditorias[idx]=Object.assign({},STORE.auditorias[idx],updated);
  closeModal();refreshAll();
  var client=getSbClient();
  if(!client){toast('✓ Guardado local (sin Supabase)');return;}
  try{
    var row={razon:updated.razon||null,centro:updated.centro||null,tienda:updated.tienda,
      fecha:updated.fecha,mes:updated.mes,pct_cumpl:updated.pctCumpl,
      tareas:updated.tareas,pendientes:updated.pendientes,resueltas:updated.resueltas,
      pct_resuelto:updated.pctResuelto,clase:updated.clase};
    /* Buscar la fila por 'centro' (código estable) en vez de 'tienda' (texto libre
       que puede diferir de lo guardado originalmente en Supabase); si no hay
       centro, se recurre a tienda como respaldo. */
    var q=client.from('auditorias').update(row);
    q=original.centro?q.ilike('centro',original.centro):q.ilike('tienda',original.tienda);
    q=q.ilike('mes',original.mes||'').ilike('clase',original.clase);
    if(original.fecha)q=q.eq('fecha',original.fecha);
    var r=await q;
    if(r.error)toast('⚠ '+r.error.message);
    else{toast('☁️ Auditoría actualizada');if(VIEW==='auditorias')renderAuditoriasView();}
  }catch(e){toast('⚠ '+e.message);}
}

async function deleteAuditoria(a){
  if(!_session||!['admin','admin_auditor','auditor'].includes(_session.rol)){toast('⚠ Sin permisos para eliminar auditorías');return;}
  if(!confirm('¿Eliminar auditoría de '+a.tienda+' ('+a.mes+')?'))return;
  closeModal();
  STORE.auditorias=STORE.auditorias.filter(function(x){
    return !((a.centro?mismaTiendaAud(x,a):norm(x.tienda)===norm(a.tienda))&&norm(x.mes)===norm(a.mes)&&norm(x.clase)===norm(a.clase)&&x.fecha===a.fecha);
  });
  refreshAll();
  var client=getSbClient();
  if(!client){toast('✓ Eliminado local');return;}
  try{
    /* Igual que en saveEditAuditoria: preferir 'centro' (código estable) sobre
       'tienda' (texto libre) para localizar la fila real en Supabase. */
    var q=client.from('auditorias').delete();
    q=a.centro?q.ilike('centro',a.centro):q.ilike('tienda',a.tienda);
    q=q.ilike('mes',a.mes||'').ilike('clase',a.clase);
    if(a.fecha)q=q.eq('fecha',a.fecha);
    var r=await q;
    if(r.error)toast('⚠ '+r.error.message);
    else{toast('☁️ Auditoría eliminada');if(VIEW==='auditorias')renderAuditoriasView();}
  }catch(e){toast('⚠ '+e.message);}
}

function renderAuditoriasView(){
  fillAudFilters();
  var arr=filteredAudByView();
  /* Excluir auditorías ya resueltas (0 pendientes) de la tabla activa.
     Antes esto además exigía que existiera un registro en Finalizadas con
     la misma clave — pero la retención de 7 días borra ese registro pasado
     ese plazo, y una auditoría que sigue genuinamente con 0 pendientes se
     quedaba sin "memoria" de que ya se había finalizado y reaparecía en la
     tabla activa. El estado real (0 pendientes) ya se calcula aquí mismo en
     vivo, así que basta con eso — no hace falta que el registro de
     Finalizadas siga existiendo para saber que ya terminó. */
  arr=arr.filter(function(a){
    var tipoClase=tipoTareaDeClase(a.clase);
    var tareasDeAud=STORE.tareas.filter(function(t){
      if(norm(t.tienda)!==norm(a.tienda)||norm(t.centro)!==norm(a.centro))return false;
      if(tipoClase&&tipoNormLocal(t.tipoTarea)!==tipoClase)return false;
      return true;
    });
    var sinPend=tareasDeAud.length>0&&tareasDeAud.filter(esPendiente).length===0;
    return !sinPend;
  });
  /* Contar auditorías con al menos 1 tarea pendiente */
  var audConPend=arr.filter(function(a){
    var stats=calcAudStats(a,arr);
    return stats.pendientes>0;
  }).length;
  document.getElementById('aud-count').textContent='';
  var cont=document.getElementById('auditorias-tables');
  if(!arr.length){cont.innerHTML='<div class="empty" style="padding:30px">Sin auditorías. Verifica los filtros o carga datos desde el módulo principal.</div>';return;}

  /* Clasificar en los dos tipos conocidos + agrupar el resto por clase real */
  function claseCanonica(a){
    var c=norm(a.clase||'');
    if(c.includes('colaboracion')||c.includes('colab'))return '__colab';
    if(c.includes('orden')||c.includes('limpieza'))return '__orden';
    if(c.includes('cartera'))return '__cartera';
    return a.clase||'SIN CLASE';
  }

  var grupos={};
  arr.forEach(function(a){
    var k=claseCanonica(a);
    if(!grupos[k])grupos[k]=[];
    grupos[k].push(a);
  });

  /* Omitir duplicados DENTRO de cada clase: misma tienda/centro + mismo mes +
     mismo % de cumplimiento casi seguro es la misma auditoría capturada o
     importada dos veces. Se compara solo dentro de la misma clase — una
     Colaboración y una Orden y Limpieza del mismo centro/mes son auditorías
     distintas y nunca se consideran duplicado entre sí. */
  Object.keys(grupos).forEach(function(k){
    var vistos={};
    grupos[k]=grupos[k].filter(function(a){
      var pct=parseFloat(a.pctCumpl)||0;if(pct>1)pct=pct/100;
      var key=[norm(a.centro||a.tienda||''),norm(a.mes||''),Math.round(pct*100)].join('|');
      if(vistos[key])return false;
      vistos[key]=true;
      return true;
    });
  });

  var html='';
  /* Primero los dos grupos conocidos en orden fijo */
  if(grupos['__colab']&&grupos['__colab'].length)
    html+=audTablaPorClaseCartera('AUDITORIAS DE COLABORACION','#2563eb',grupos['__colab'],false);
  if(grupos['__orden']&&grupos['__orden'].length)
    html+=audTablaPorClaseCartera('ORDEN Y LIMPIEZA','#16a34a',grupos['__orden'],false);
  if(grupos['__cartera']&&grupos['__cartera'].length)
    html+=audTablaPorClaseCartera('AUDITORIA CARTERA','#7c3aed',grupos['__cartera'],true);
  /* Luego cualquier otra clase */
  Object.keys(grupos).forEach(function(k){
    if(k==='__colab'||k==='__orden'||k==='__cartera')return;
    html+=audTablaPorClase(k.toUpperCase(),'#7c8696',grupos[k]);
  });

  /* Si hay 0 auditorías en colab o en orden, mostrar sección vacía informativa */
  if(!grupos['__colab']||!grupos['__colab'].length)
    html+='<div class="slbl" style="color:var(--k-blue);margin-bottom:8px"><span class="dot" style="background:var(--k-blue)"></span>AUDITORIAS DE COLABORACION <span style="font-size:10px;color:var(--muted);font-weight:500">(0 con los filtros actuales)</span></div>';
  if(!grupos['__orden']||!grupos['__orden'].length)
    html+='<div class="slbl" style="color:var(--k-greenok);margin-bottom:8px"><span class="dot" style="background:var(--k-greenok)"></span>ORDEN Y LIMPIEZA <span style="font-size:10px;color:var(--muted);font-weight:500">(0 con los filtros actuales)</span></div>';

  cont.innerHTML=html;
}

function downloadAuditoriasPNG(){
  toast('⏳ Generando PNG...');
  var el=document.getElementById('auditorias-tables');
  if(typeof html2canvas==='undefined'){toast('⚠ Librería de captura no disponible');return;}
  html2canvas(el,{scale:2,backgroundColor:'#fff'}).then(function(canvas){
    var link=document.createElement('a');
    link.download='auditorias_por_clase.png';
    link.href=canvas.toDataURL('image/png');
    link.click();
    toast('✓ PNG descargado');
  }).catch(function(e){toast('⚠ Error: '+e.message);});
}

/* ════════════════════════════════════════════════════════════════════
   NUEVA AUDITORÍA (modal)
════════════════════════════════════════════════════════════════════ */
function openNuevaAuditoriaIfAllowed(){
  if(!_session||!['admin','admin_auditor','auditor'].includes(_session.rol)){toast('⚠ Sin permisos para crear auditorías');return;}
  openNuevaAuditoria();
}
function openNuevaAuditoria(){
  /* Rellenar datalists con valores existentes */
  var tiendas=uniq(STORE.auditorias.map(function(a){return a.tienda;}).concat(STORE.tareas.map(function(t){return t.tienda;})));
  var centros=uniq(STORE.auditorias.map(function(a){return a.centro;}).concat(STORE.tareas.map(function(t){return t.centro;})));
  var razones=uniq(STORE.auditorias.map(function(a){return a.razon;}).concat(STORE.tareas.map(function(t){return t.razon;})));
  document.getElementById('na-tienda-list').innerHTML=limpiarOpciones(tiendas).map(function(v){return'<option value="'+v+'">';}).join('');
  document.getElementById('na-centro-list').innerHTML=limpiarOpciones(centros).map(function(v){return'<option value="'+v+'">';}).join('');
  document.getElementById('na-razon-list').innerHTML=razones.map(function(v){return'<option value="'+v+'">';}).join('');
  /* Reset campos */
  ['na-tienda','na-centro','na-razon','na-pct','na-tareas','na-pend','na-res'].forEach(function(id){
    document.getElementById(id).value='';
  });
  document.getElementById('na-mes').value='';
  document.getElementById('na-fecha').value='';
  document.getElementById('na-pct-res-preview').textContent='—';
  document.getElementById('na-pct-res-preview').style.color='var(--blue)';
  document.getElementById('na-err').textContent='';
  document.getElementById('aud-modal-overlay').classList.add('show');
}

function closeNuevaAuditoria(){
  document.getElementById('aud-modal-overlay').classList.remove('show');
}

function naAutoCalc(){
  var total=parseInt(document.getElementById('na-tareas').value)||0;
  var pend=parseInt(document.getElementById('na-pend').value)||0;
  var res=parseInt(document.getElementById('na-res').value)||0;
  /* Si llenan tareas y pendientes, auto-calcular resueltos */
  if(total>0&&pend>=0&&document.getElementById('na-res').value===''){
    var autoRes=total-pend;
    if(autoRes>=0) document.getElementById('na-res').value=autoRes;
    res=autoRes>=0?autoRes:0;
  }
  var pctRes=total>0?Math.round(res/total*100):0;
  var prev=document.getElementById('na-pct-res-preview');
  prev.textContent=total>0?(pctRes+'%'):'—';
  prev.style.color=pctRes>=80?'#16a34a':pctRes>=50?'#d97706':'#dc2626';
}

async function guardarNuevaAuditoria(){
  var errEl=document.getElementById('na-err');
  errEl.textContent='';
  var tienda=document.getElementById('na-tienda').value.trim().toUpperCase();
  var mes=document.getElementById('na-mes').value;
  if(!tienda){errEl.textContent='La tienda es obligatoria';return;}
  if(!mes){errEl.textContent='El mes es obligatorio';return;}

  var clase=document.getElementById('na-clase').value;
  var centro=canonCentro(document.getElementById('na-centro').value.trim().toUpperCase());
  var razon=document.getElementById('na-razon').value.trim().toUpperCase();
  var fechaVal=document.getElementById('na-fecha').value;
  var pct=parseFloat(document.getElementById('na-pct').value)||0;
  var total=parseInt(document.getElementById('na-tareas').value)||0;
  var pend=parseInt(document.getElementById('na-pend').value)||0;
  var res=parseInt(document.getElementById('na-res').value)||0;
  var pctRes=total>0?res/total:0;

  /* Fecha: si no se puso usar hoy */
  var fecha=fechaVal?fechaVal:new Date().toISOString().split('T')[0];

  var newAud={
    razon:razon||'KNO', centro:centro, tienda:tienda,
    fecha:fecha, mes:mes.toUpperCase(),
    pctCumpl:pct>1?pct/100:pct,
    tareas:total, pendientes:pend, resueltas:res,
    pctResuelto:pctRes, clase:clase
  };

  /* Agregar a STORE local — sincronizar dashboard y vista de auditorías */
  STORE.auditorias.unshift(newAud);
  closeNuevaAuditoria();
  refreshAll();
  fillAudFilters();
  if(VIEW==='auditorias') renderAuditoriasView();
  toast('✓ Auditoría agregada');

  /* Intentar sync a Supabase */
  var client=getSbClient();
  if(!client){toast('⚠ Sin Supabase — guardado solo en local');return;}
  try{
    var row={
      razon:newAud.razon||null, centro:newAud.centro||null, tienda:newAud.tienda,
      fecha:newAud.fecha, mes:newAud.mes,
      pct_cumpl:newAud.pctCumpl, tareas:newAud.tareas,
      pendientes:newAud.pendientes, resueltas:newAud.resueltas,
      pct_resuelto:newAud.pctResuelto, clase:newAud.clase
    };
    var r=await client.from('auditorias').upsert([row],{onConflict:'razon,centro,tienda,fecha',ignoreDuplicates:false});
    if(r.error)toast('⚠ Supabase: '+r.error.message);
    else toast('☁️ Auditoría guardada en Supabase');
  }catch(e){toast('⚠ Error Supabase: '+e.message);}
}

/* ════════════════════════════════════════════════════════════════════
   MÓDULO ACTIVIDADES
════════════════════════════════════════════════════════════════════ */
let ACTIVIDADES=[];
var _actMesManual=false; /* true cuando el usuario elige el mes a mano: evita que una carga lo reescriba */

/* Una actividad "Completado" cuenta como cumplida (a tiempo) mientras se haya
   terminado DENTRO DEL MISMO MES que la fecha estimada de fin — no se exige
   que coincida el día exacto. Solo se marca "atrasada" si Real Fin cae en un
   mes posterior al de Est. Fin. Se usa tanto en la tabla de Actividades
   (columna Cumplimiento) como en Desempeño, para que ambos coincidan siempre. */
/* ¿Una actividad NO completada ya está vencida? Mismo criterio mensual que
   actividadEnTiempoDesempeno: se considera vencida si su mes estimado de fin
   ya quedó atrás respecto al mes en curso. Sin fecha estimada no se juzga. */
/* ════════════════════════════════════════════════════════════════════
   ACTIVIDADES COMPARTIDAS ENTRE AUDITORES
   El campo "asignado" admite uno o varios auditores separados por " / ",
   coma, "y" o "&". También acepta "Ambos" / "Todos", que se expande a los
   auditores de la razón social correspondiente. La actividad se evalúa y
   califica PARA CADA UNO de los auditores involucrados.
════════════════════════════════════════════════════════════════════ */
function auditoresDeActividad(a){
  var txt=String((a&&a.asignado)||'').trim();
  if(!txt)return [];
  /* "Ambos" / "Todos" → todos los auditores de esa razón social */
  if(/^(ambos|todos|ambas|equipo)$/i.test(norm(txt))){
    var rz=a.razon||'';
    var lista=(_AUDITORES_CONOCIDOS||[]).filter(function(n){
      var meta=_AUDITORES_META[norm(n)];
      if(!rz||!meta||!meta.razones)return true;
      return meta.razones.some(function(r){return razKey(r)===razKey(rz);});
    });
    return lista.length?lista:[];
  }
  return txt.split(/\s*(?:\/|,|&|\+|\by\b)\s*/i)
            .map(function(s){return s.trim();})
            .filter(function(s){return s.length>1;});
}

function actividadVencidaDesempeno(a){
  if(!a.estFin)return false;
  var ef=fromISO(a.estFin); if(!ef)return false;
  var hoy=new Date();
  return (ef.getFullYear()*12+ef.getMonth()) < (hoy.getFullYear()*12+hoy.getMonth());
}
/* El desempeño ahora se mide con la columna Estado: una actividad marcada
   "Completado" cuenta como cumplida. Ya no se compara Est.Fin vs Real Fin
   (la columna Cumplimiento se eliminó de la tabla de Actividades). */
function actividadEnTiempoDesempeno(a){
  return norm(a.estado||'').includes('completad');
}

async function loadActividades(){
  var client=getSbClient();
  if(!client){toast('⚠ Sin conexión a Supabase');return;}
  try{
    var r=await client.from('actividades').select('*').order('est_inicio',{ascending:true}).limit(5000);
    if(r.error){toast('⚠ '+r.error.message);return;}
    var _actRaw=r.data||[];
    var _actDec=await decArr(_actRaw,FIELDS.actividades);
    ACTIVIDADES=_actDec.map(function(a){return{
      id:a.id, mes:a.mes||'', actividad:a.actividad||'', estado:a.estado||'',
      categoria:a.categoria||'', asignado:a.asignado||'',
      estInicio:a.est_inicio, estFin:a.est_fin, durEstimada:a.dur_estimada,
      realInicio:a.real_inicio, realFin:a.real_fin, cumplimiento:a.cumplimiento,
      durReal:a.dur_real, pctAdicional:a.pct_adicional,
      cumplFechaInicio:a.cumpl_fecha_inicio, variacionDias:a.variacion_dias,
      comentario:a.comentario||'', apoyos:a.apoyos||'',
      programada:a.programada!==false, sb_updated:a.updated_at||null,
      creadoPor:a.creado_por||'', razon:(a.razon&&!pareceCifrado(a.razon))?a.razon:''
    };});
    /* Visibilidad por razón social: usuarios restringidos solo ven las
       actividades de su(s) razón(es). Las antiguas sin razón quedan visibles
       solo para cuentas con acceso total (para poder asignarles razón). */
    ACTIVIDADES=ACTIVIDADES.filter(function(a){return razonVisible(a.razon);});
    fillActFilters();
    renderActividades();
    toast('✓ '+ACTIVIDADES.length+' actividades cargadas');
  }catch(e){toast('⚠ Error: '+e.message);}
}

function fillActFilters(){
  var asigs=[...new Set(ACTIVIDADES.reduce(function(acc,a){
    return acc.concat(auditoresDeActividad(a));},[]).filter(Boolean))].sort();
  var selA=document.getElementById('act-f-asignado');
  var vA=selA.value;
  selA.innerHTML='<option value="ALL">Todos</option>'+asigs.map(a=>'<option>'+a+'</option>').join('');
  selA.value=vA;
  /* Filtro RAZÓN SOCIAL: separa las actividades entre las 3 razones y permite
     detectar rápidamente las que quedaron sin razón asignada (invisibles para
     personal restringido). Si la cuenta está fija a una sola razón, se bloquea. */
  var _rzAsigFF=_razonesAsignadas();
  var selR=document.getElementById('act-f-razon');
  if(selR){
    var vR=selR.value;
    var hayBlancas=ACTIVIDADES.some(function(a){return!a.razon;});
    var razonesDetectadas=[...new Set(['KNO','KSC','KSA'].concat(ACTIVIDADES.map(a=>a.razon).filter(Boolean)))];
    if(_rzAsigFF&&_rzAsigFF.length===1){
      selR.innerHTML='<option value="'+_rzAsigFF[0]+'">'+_rzAsigFF[0]+'</option>';
      selR.disabled=true;
    }else{
      selR.disabled=false;
      selR.innerHTML='<option value="ALL">Todas</option>'+
        razonesDetectadas.map(function(r){return'<option value="'+r+'">'+r+'</option>';}).join('')+
        (hayBlancas?'<option value="SIN_RAZON">⚠ Sin razón</option>':'');
      if(vR&&[...selR.options].some(function(o){return o.value===vR;}))selR.value=vR;
    }
  }
  /* Botón de asignación masiva: solo admin/admin_auditor, y solo si hay
     actividades sin razón que reclasificar */
  var btnAR=document.getElementById('btn-asignar-razon');
  if(btnAR){
    var esAdminFF=_session&&['admin','admin_auditor'].includes(_session.rol);
    var hayBlancas2=ACTIVIDADES.some(function(a){return!a.razon;});
    btnAR.style.display=(esAdminFF&&hayBlancas2)?'':'none';
  }
  /* Mes por defecto = mes en curso (si tiene datos); si no, el más reciente con
     datos. Se recalcula en cada carga salvo que el usuario haya elegido el mes a
     mano (_actMesManual), para que al importar el mes actual aparezca solo. */
  var MORD_ACT=['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'];
  var sm=document.getElementById('act-f-mes');
  if(sm&&!_actMesManual){
    var conDatos=function(m){return ACTIVIDADES.some(function(a){return norm(a.mes)===norm(m);});};
    var mesHoy=MORD_ACT[new Date().getMonth()];
    var tgt=null;
    if(conDatos(mesHoy))tgt=mesHoy;
    else{
      var bi=-1;
      ACTIVIDADES.forEach(function(a){var i=MORD_ACT.map(norm).indexOf(norm(a.mes));if(i>bi)bi=i;});
      if(bi>=0)tgt=MORD_ACT[bi];
    }
    var opt=tgt&&[...sm.options].find(function(o){return norm(o.value)===norm(tgt);});
    sm.value=opt?opt.value:'ALL';
  }
}
function actMesChanged(){_actMesManual=true;renderActividades();}

function filteredActividades(){
  var mes=document.getElementById('act-f-mes').value;
  var est=document.getElementById('act-f-estado').value;
  var cat=document.getElementById('act-f-categoria').value;
  var asig=document.getElementById('act-f-asignado').value;
  var selR=document.getElementById('act-f-razon');
  var raz=selR?selR.value:'ALL';
  return ACTIVIDADES.filter(function(a){
    if(mes!=='ALL'&&norm(a.mes)!==norm(mes))return false;
    if(est!=='ALL'&&norm(a.estado)!==norm(est))return false;
    if(cat!=='ALL'&&a.categoria!==cat)return false;
    if(asig!=='ALL'&&!auditoresDeActividad(a).some(function(x){return norm(x)===norm(asig);}))return false;
    if(raz==='SIN_RAZON'){if(a.razon)return false;}
    else if(raz!=='ALL'&&razKey(a.razon)!==razKey(raz))return false;
    return true;
  });
}

/* ¿Una actividad NO completada ya venció su Est.Fin? Comparación EXACTA por
   día (no por mes, a diferencia de actividadVencidaDesempeno que es mensual
   y solo aplica a Desempeño). Sin esto, una actividad "Sin comenzar"/"En
   curso"/"Reprogramado" cuyo Est.Fin ya pasó se seguía viendo con su color
   normal (vigente) de forma indefinida, tanto en la tabla de Actividades
   como en el PPT descargable. Se usa para recolorear en rojo y así avisar
   de inmediato que quedó atrasada. */
function actEstaVencida(a){
  if(!a)return false;
  if(norm(a.estado||'').includes('completad'))return false;
  /* Si ya tiene Real Finalización capturada, el trabajo ya se hizo — aunque
     el campo Estado se haya quedado en "En curso"/"Sin comenzar" por olvido.
     Sin esto, cualquier actividad con Real Fin ya registrada seguía
     marcándose "Vencida" solo porque nadie actualizó el texto de Estado. */
  if(a.realFin)return false;
  var ef=fromISO(a.estFin); if(!ef)return false;
  var hoy=new Date(); hoy.setHours(0,0,0,0);
  var efD=new Date(ef); efD.setHours(0,0,0,0);
  return efD<hoy;
}

function actEstadoBadge(a){
  var e=norm(a&&a.estado);
  var color='#1d4ed8', bg='#dbeafe';  // default sin comenzar: azul claro / letra azul
  if(e.includes('completad')){color='#15803d';bg='#dcfce7';}           // verde
  else if(e.includes('curso')){color='#1a1a1a';bg='#fb923c';}          // naranja, letra negra
  else if(e.includes('reprogramad')){color='#ffffff';bg='#2563eb';}    // azul, letra blanca
  else if(e.includes('comenzar')||e.includes('sin')){color='#1d4ed8';bg='#dbeafe';} // azul claro, letra azul
  var texto=(a&&a.estado)||'—';
  if(actEstaVencida(a))texto+=' · Vencida';   // se avisa en el texto, pero ya no se pinta de rojo: siempre el color fijo del estado
  return '<span style="display:inline-block;padding:2px 9px;border-radius:20px;font-size:11px;font-weight:700;color:'+color+';background:'+bg+'">'+texto+'</span>';
}

/* Mismo criterio de color que actEstadoBadge, pero como estilo inline para
   el <select> editable de la columna Estado (edición directa en la tabla). */
function estadoSelStyle(a){
  var e=norm(a&&a.estado);
  var color='#1d4ed8', bg='#dbeafe';
  if(e.includes('completad')){color='#15803d';bg='#dcfce7';}
  else if(e.includes('curso')){color='#1a1a1a';bg='#fb923c';}
  else if(e.includes('reprogramad')){color='#ffffff';bg='#2563eb';}
  else if(e.includes('comenzar')||e.includes('sin')){color='#1d4ed8';bg='#dbeafe';}
  // Ya no se pinta de rojo por vencida: el color del select siempre es el fijo del estado (el aviso queda en el title/tooltip)
  return 'padding:3px 8px;border-radius:20px;font-size:11px;font-weight:700;color:'+color+';background:'+bg+';border:none;cursor:pointer;appearance:auto';
}

/* Edición rápida de un campo de Actividades directamente desde la tabla, sin
   abrir el modal. Actualiza en memoria, persiste en Supabase (cifrando el
   campo si corresponde) y vuelve a pintar la tabla para reflejar cambios
   dependientes (p. ej. Dur.real cuando cambian Est.Inicio/Est.Fin). */
var _CAMPO_ACT_COL={estado:'estado',categoria:'categoria',asignado:'asignado',estInicio:'est_inicio',estFin:'est_fin'};
async function updateCampoActividadInline(id,campo,valor){
  var col=_CAMPO_ACT_COL[campo];
  if(!col)return;
  var a=ACTIVIDADES.find(function(x){return String(x.id)===String(id);});
  if(!a)return;
  var anterior=a[campo];
  a[campo]=valor||null;
  try{
    var client=getSbClient();
    if(!client){toast('⚠ Sin conexión a Supabase');a[campo]=anterior;renderActividades();return;}
    var payload={};payload[col]=valor||null;
    var row=(FIELDS.actividades.indexOf(col)>=0)?await encObj(payload,FIELDS.actividades):payload;
    var ru=await client.from('actividades').update(row).eq('id',id);
    if(ru.error){toast('⚠ '+ru.error.message);a[campo]=anterior;renderActividades();return;}
    toast('☁️ Actualizado');
    renderActividades();
  }catch(e){toast('⚠ Error: '+e.message);a[campo]=anterior;renderActividades();}
}

/* Estilo neutro (sin color semántico) para los <select>/<input> editables de
   Categoría, Asignado, Est.Inicio y Est.Fin. */
function inlineFieldStyle(){
  return 'padding:3px 8px;border-radius:8px;font-size:11.5px;font-weight:600;color:var(--txt);'+
    'background:var(--soft);border:1px solid var(--border);cursor:pointer;font-family:inherit';
}

function fmtActFecha(d){if(!d)return'—';var x=fromISO(d);if(!x)return'—';return fmtInput(x).split('-').reverse().join('/');}

function renderActividades(){
  var arr=filteredActividades();
  document.getElementById('act-count').textContent=arr.length+' actividad(es)';
  var canEdit=_session&&['admin','admin_auditor','auditor'].includes(_session.rol);
  var cont=document.getElementById('actividades-table');

  if(!arr.length){cont.innerHTML='<div class="empty" style="padding:30px">Sin actividades. '+(canEdit?'Crea una con "Nueva actividad".':'')+'</div>';return;}

  function estadoCelda(a){
    if(!canEdit)return actEstadoBadge(a);
    var estados=['Completado','En curso','Reprogramado','Sin comenzar'];
    var opts=estados.map(function(e){return '<option value="'+e+'"'+(norm(a.estado||'')===norm(e)?' selected':'')+'>'+e+'</option>';}).join('');
    var titVenc=actEstaVencida(a)?' title="⚠ Vencida: Est.Fin ya pasó"':'';
    return '<select class="act-estado-inline"'+titVenc+' style="'+estadoSelStyle(a)+'" onchange="updateCampoActividadInline(\''+a.id+'\',\'estado\',this.value)">'+opts+'</select>';
  }
  function categoriaCelda(a){
    if(!canEdit)return (a.categoria||'—');
    var cats=['Evaluacion','Auditoria','Entrega','Inventario','Revision','Apoyo Juridico','Apoyo a RH'];
    var opts='<option value=""'+(!a.categoria?' selected':'')+'>—</option>'+
      cats.map(function(c){return '<option value="'+c+'"'+(norm(a.categoria)===norm(c)?' selected':'')+'>'+c+'</option>';}).join('');
    return '<select class="act-cat-inline" style="'+inlineFieldStyle()+'" onchange="updateCampoActividadInline(\''+a.id+'\',\'categoria\',this.value)">'+opts+'</select>';
  }
  /* Lista de auditores disponibles para la razón social de la actividad —
     mismo criterio que pintarChipsAuditores, para que "Todos" sea consistente
     entre el modal completo y la edición rápida en la tabla, en CUALQUIER
     razón social. */
  function auditoresDeRazon(rzSel){
    function enRazon(nombre){
      if(!rzSel)return true;
      var meta=_AUDITORES_META[norm(nombre)];
      if(!meta||!meta.razones)return true;
      return meta.razones.some(function(r){return razKey(r)===razKey(rzSel);});
    }
    var mapa={};
    (_AUDITORES_CONOCIDOS||[]).forEach(function(n){if(n&&!pareceCifrado(n)&&enRazon(n))mapa[norm(n)]=n;});
    (ACTIVIDADES||[]).forEach(function(x){auditoresDeActividad(x).forEach(function(n){if(n&&!pareceCifrado(n)&&!mapa[norm(n)]&&enRazon(n))mapa[norm(n)]=n;});});
    return Object.values(mapa).sort(function(x,y){return x.localeCompare(y);});
  }
  function asignadoCelda(a){
    if(!canEdit)return (a.asignado||'—');
    var lista=auditoresDeRazon(a.razon);
    var actual=(a.asignado||'').trim();
    var todosVal=lista.join(' / ');
    var opts='<option value=""'+(!actual?' selected':'')+'>—</option>'+
      lista.map(function(n){return '<option value="'+esc(n)+'"'+(norm(actual)===norm(n)?' selected':'')+'>'+esc(n)+'</option>';}).join('');
    if(lista.length>1){
      opts+='<option value="'+esc(todosVal)+'"'+(norm(actual)===norm(todosVal)?' selected':'')+'>👥 Todos</option>';
    }
    var yaExiste=!actual||lista.some(function(n){return norm(n)===norm(actual);})||norm(actual)===norm(todosVal);
    if(!yaExiste)opts+='<option value="'+esc(actual)+'" selected>'+esc(actual)+'</option>';
    return '<select class="act-asig-inline" style="'+inlineFieldStyle()+'" onchange="updateCampoActividadInline(\''+a.id+'\',\'asignado\',this.value)">'+opts+'</select>';
  }
  function fechaCelda(a,campo){
    if(!canEdit)return fmtActFecha(a[campo]);
    var x=fromISO(a[campo]);
    var val=x?fmtInput(x):'';
    return '<input type="date" class="act-fecha-inline" style="'+inlineFieldStyle()+'" value="'+val+'" onchange="updateCampoActividadInline(\''+a.id+'\',\''+campo+'\',this.value)">';
  }
  function filaActividad(a){
    /* La columna se llama "Dur.real" pero calculaba Est.Inicio→Est.Fin (la
       duración ESTIMADA), no la real. Esto hacía parecer que una actividad
       ya tenía avance real registrado cuando en realidad solo tenía fechas
       estimadas — y por eso una actividad sin Real Fin capturado (todavía
       pendiente de verdad) podía leerse como si ya tuviera duración real. */
    var dur=calcDias(a.realInicio,a.realFin);
    return '<tr>'+
      '<td>'+(a.mes||'—')+'</td>'+
      '<td style="min-width:200px"><b>'+(a.actividad||'—')+'</b></td>'+
      '<td>'+estadoCelda(a)+'</td>'+
      '<td>'+categoriaCelda(a)+'</td>'+
      '<td>'+asignadoCelda(a)+'</td>'+
      '<td>'+fechaCelda(a,'estInicio')+'</td>'+
      '<td>'+fechaCelda(a,'estFin')+'</td>'+
      '<td style="text-align:center">'+(dur!==''?dur:'—')+'</td>'+
      '<td class="cell-c">'+(canEdit?'<button class="icon-btn" onclick="openActividad(\''+a.id+'\')">✎</button>':'<span style="color:var(--border)">—</span>')+'</td>'+
    '</tr>';
  }

  var thead='<thead><tr><th>Mes</th><th>Actividad/Proyecto</th><th>Estado</th><th>Categoría</th><th>Asignado</th>'+
    '<th>Est.Inicio</th><th>Est.Fin</th><th>Dur.real</th><th></th></tr></thead>';

  var prog=arr.filter(function(a){return a.programada;});
  var noProg=arr.filter(function(a){return !a.programada;});

  var html='';
  // Tabla PROGRAMADAS
  html+='<div class="slbl" style="margin:4px 0 8px"><span class="dot" style="background:var(--green)"></span>Programadas <span style="font-weight:600;font-size:10px;color:var(--muted)">('+prog.length+')</span></div>';
  if(prog.length){
    html+='<table class="dt">'+thead+'<tbody>'+prog.map(filaActividad).join('')+'</tbody></table>';
  } else {
    html+='<div class="empty" style="padding:18px;font-size:12px">Sin actividades programadas.</div>';
  }
  // Tabla NO PROGRAMADAS
  html+='<div class="slbl" style="margin:22px 0 8px"><span class="dot" style="background:#dc2626"></span>No programadas <span style="font-weight:600;font-size:10px;color:var(--muted)">('+noProg.length+')</span></div>';
  if(noProg.length){
    html+='<table class="dt">'+thead+'<tbody>'+noProg.map(filaActividad).join('')+'</tbody></table>';
  } else {
    html+='<div class="empty" style="padding:18px;font-size:12px">Sin actividades no programadas.</div>';
  }

  cont.innerHTML=html;
}

function openActividadIfAllowed(id){
  if(!_session||!['admin','admin_auditor','auditor'].includes(_session.rol)){  
    toast('⚠ Sin permisos para editar actividades');return;
  }
  openActividad(id);
}

function openActividad(id){
  var a=id?ACTIVIDADES.find(x=>String(x.id)===String(id)):null;
  var isNew=!a;
  if(!a)a={mes:'',actividad:'',estado:'Sin comenzar',categoria:'',asignado:'',estInicio:'',estFin:'',durEstimada:'',realInicio:'',realFin:'',durReal:'',comentario:'',apoyos:'',programada:true,razon:''};
  /* Razón social: si el usuario tiene UNA sola asignada, se fija; si tiene
     varias, elige entre ellas; acceso total puede escribir cualquiera. */
  var _rzAsig=_razonesAsignadas();
  var _rzVal=(a.razon||((_rzAsig&&_rzAsig.length===1)?_rzAsig[0]:'')||'').toUpperCase();
  var _RAZONES_FIJAS=['KNO','KSC','KSA'];
  var _rzOpts=[...new Set(_RAZONES_FIJAS.concat(_rzAsig||[]).concat(STORE.tareas.map(function(t){return t.razon;})).concat(STORE.auditorias.map(function(x){return x.razon;})).filter(Boolean).map(function(r){return r.toUpperCase();}))];
  function di(d){if(!d)return'';var x=fromISO(d);return x?fmtInput(x):'';}
  var meses=['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'];
  var cats=['Evaluacion','Auditoria','Entrega','Inventario','Revision','Apoyo Juridico','Apoyo a RH'];
  var estados=['Completado','En curso','Reprogramado','Sin comenzar'];
  var _rzBloqueada=_rzAsig&&_rzAsig.length===1;
  var body='<div style="display:grid;grid-template-columns:1fr 1fr;gap:12px">'+
    '<div class="fg" style="grid-column:1/3"><label>Actividad / Proyecto *</label><input id="a-actividad" value="'+esc(a.actividad)+'" placeholder="Nombre de la actividad"></div>'+
    '<div class="fg"><label>Razón social '+(_rzBloqueada?'':'*')+'</label><select id="a-razon" onchange="cambioRazonActividad()"'+(_rzBloqueada?' disabled title="Tu cuenta está asignada a '+esc(_rzAsig[0])+'"':'')+'>'+
      (_rzBloqueada?'':'<option value="">— Selecciona —</option>')+
      _rzOpts.map(function(r){return'<option value="'+esc(r)+'"'+(r===_rzVal?' selected':'')+'>'+esc(r)+'</option>';}).join('')+
      '</select>'+
      (_rzBloqueada?'':'<div style="font-size:10px;color:var(--muted);margin-top:4px">Obligatorio: sin razón, la actividad queda oculta para el personal con acceso restringido. Elige la razón del equipo (senior + jr) que la captura.</div>')+
      '</div>'+
    '<div class="fg"><label>Mes</label><select id="a-mes">'+['<option value="">—</option>'].concat(meses.map(m=>'<option'+(norm(a.mes)===norm(m)?' selected':'')+'>'+m+'</option>')).join('')+'</select></div>'+
    '<div class="fg"><label>Estado</label><select id="a-estado">'+estados.map(e=>'<option'+(norm(a.estado)===norm(e)?' selected':'')+'>'+e+'</option>').join('')+'</select></div>'+
    '<div class="fg"><label>Categoría</label><select id="a-categoria"><option value="">—</option>'+cats.map(c=>'<option'+(norm(a.categoria)===norm(c)?' selected':'')+'>'+c+'</option>').join('')+'</select></div>'+
    '<div class="fg" style="grid-column:1/3"><label>Asignado a <span style="font-weight:400;color:var(--muted);text-transform:none">(elige uno o varios: la actividad se califica para todos)</span></label>'+
      '<input id="a-asignado" value="'+esc(a.asignado)+'" placeholder="Nombre, o varios separados por /" oninput="pintarChipsAuditores()">'+
      '<div id="a-asig-chips" style="display:flex;flex-wrap:wrap;gap:6px;margin-top:7px"></div></div>'+
    '<div class="fg"><label>Estimado Inicio</label><input type="date" id="a-estinicio" value="'+di(a.estInicio)+'" oninput="calcDurEst()"></div>'+
    '<div class="fg"><label>Estimado Finalización</label><input type="date" id="a-estfin" value="'+di(a.estFin)+'" oninput="calcDurEst()"></div>'+
    '<div class="fg"><label>Duración estimada (días)</label><input type="number" id="a-durest" value="'+calcDias(a.estInicio,a.estFin)+'" readonly style="background:#f1f5f9"></div>'+
    '<div class="fg"><label>Real Inicio</label><input type="date" id="a-realinicio" value="'+di(a.realInicio)+'" oninput="calcDurReal()"></div>'+
    '<div class="fg"><label>Real Finalización</label><input type="date" id="a-realfin" value="'+di(a.realFin)+'" oninput="calcDurReal()"></div>'+
    '<div class="fg"><label>Duración real (días)</label><input type="number" id="a-durreal" value="'+calcDias(a.realInicio,a.realFin)+'" readonly style="background:#f1f5f9"></div>'+
    '<div class="fg" style="grid-column:1/3"><label>Comentario</label><input id="a-comentario" value="'+esc(a.comentario)+'"></div>'+
    '<div class="fg" style="grid-column:1/3"><label>Apoyos requeridos</label><input id="a-apoyos" value="'+esc(a.apoyos)+'"></div>'+
    '<label style="grid-column:1/3;display:flex;align-items:center;gap:8px;font-size:13px;color:var(--txt)"><input type="checkbox" id="a-programada"'+(a.programada?' checked':'')+'> Actividad programada (desmarcar si es NO PROGRAMADA)</label>'+
  '</div>';
  var foot=[{label:isNew?'Crear':'Guardar',cls:'btn-blue',fn:function(){saveActividad(a.id,isNew);}}];
  if(!isNew&&_session&&['admin','admin_auditor'].includes(_session.rol))foot.push({label:'Eliminar',cls:'btn-red',fn:function(){deleteActividad(a.id);}});
  openModal(isNew?'➕ Nueva actividad':'✎ Editar actividad',body,foot);
  pintarChipsAuditores();   /* chips de asignación (uno o varios auditores) */
}

/* Calcula días entre dos fechas (inclusive) */
function calcDias(d1,d2){
  if(!d1||!d2)return '';
  var a=fromISO(d1), b=fromISO(d2);
  if(!a||!b)return '';
  var diff=Math.round((b-a)/(1000*60*60*24));
  return diff>=0?diff:'';
}
function calcDurEst(){
  var v=calcDias(document.getElementById('a-estinicio').value,document.getElementById('a-estfin').value);
  document.getElementById('a-durest').value=v;
}
function calcDurReal(){
  var v=calcDias(document.getElementById('a-realinicio').value,document.getElementById('a-realfin').value);
  document.getElementById('a-durreal').value=v;
}

/* Chips para asignar uno o VARIOS auditores (o "Todos") a una actividad.
   Solo se muestran los auditores de la RAZÓN SOCIAL elegida en la actividad:
   no se puede asignar una tarea de KNO a un auditor de KSC. La lista se
   deduplica sin distinguir mayúsculas ni acentos (FERNANDO GUERRERO y
   Fernando Guerrero son la misma persona) conservando el nombre de la BD. */
function pintarChipsAuditores(){
  var cont=document.getElementById('a-asig-chips'), inp=document.getElementById('a-asignado');
  if(!cont||!inp)return;
  var rzSel=(document.getElementById('a-razon')||{}).value||'';

  /* ¿Este auditor pertenece a la razón social de la actividad? */
  function enRazon(nombre){
    if(!rzSel)return true;
    var meta=_AUDITORES_META[norm(nombre)];
    if(!meta||!meta.razones)return true;   /* auditor sin restricción: aplica a todas */
    return meta.razones.some(function(r){return razKey(r)===razKey(rzSel);});
  }

  /* Deduplicar por nombre normalizado; el alta en Usuarios tiene prioridad */
  var mapa={};
  (_AUDITORES_CONOCIDOS||[]).forEach(function(n){
    if(n&&!pareceCifrado(n)&&enRazon(n))mapa[norm(n)]=n;
  });
  (ACTIVIDADES||[]).forEach(function(a){
    auditoresDeActividad(a).forEach(function(n){
      if(n&&!pareceCifrado(n)&&!mapa[norm(n)]&&enRazon(n))mapa[norm(n)]=n;
    });
  });
  var lista=Object.values(mapa).sort(function(a,b){return a.localeCompare(b);});
  var sel=auditoresDeActividad({asignado:inp.value}).map(norm);
  cont.innerHTML=lista.map(function(n){
    var on=sel.indexOf(norm(n))>=0;
    return '<button type="button" data-aud="'+esc(n)+'" style="border:1px solid '+(on?'transparent':'var(--border)')+
      ';background:'+(on?'var(--k-blue,#5e72e4)':'transparent')+';color:'+(on?'#fff':'var(--txt)')+
      ';border-radius:20px;padding:5px 12px;font-size:12px;font-weight:600;cursor:pointer;font-family:inherit">'+
      (on?'✓ ':'')+esc(n)+'</button>';
  }).join('')+
  (lista.length>1?'<button type="button" data-aud="__AMBOS__" style="border:1px dashed var(--k-orange,#fb6340);background:transparent;color:var(--k-orange,#fb6340);border-radius:20px;padding:5px 12px;font-size:12px;font-weight:700;cursor:pointer;font-family:inherit">👥 Todos</button>':'');
  if(!lista.length){
    cont.innerHTML='<span style="font-size:12px;color:var(--muted)">No hay auditores registrados en '+esc(rzSel||'esta razón social')+'.</span>';
    return;
  }
  cont.querySelectorAll('button').forEach(function(b){
    b.onclick=function(){
      var n=b.getAttribute('data-aud');
      if(n==='__AMBOS__'){ inp.value=lista.join(' / '); pintarChipsAuditores(); return; }
      var act=auditoresDeActividad({asignado:inp.value});
      var i=act.map(norm).indexOf(norm(n));
      if(i>=0)act.splice(i,1); else act.push(n);
      inp.value=act.join(' / ');
      pintarChipsAuditores();
    };
  });
}

/* Al cambiar la razón social: quitar auditores que ya no pertenecen a ella */
function cambioRazonActividad(){
  var inp=document.getElementById('a-asignado');
  var rzSel=(document.getElementById('a-razon')||{}).value||'';
  if(inp&&rzSel){
    var quedan=auditoresDeActividad({asignado:inp.value}).filter(function(n){
      var meta=_AUDITORES_META[norm(n)];
      if(!meta||!meta.razones)return true;
      return meta.razones.some(function(r){return razKey(r)===razKey(rzSel);});
    });
    inp.value=quedan.join(' / ');
  }
  pintarChipsAuditores();
}

async function saveActividad(id,isNew){
  var actividad=document.getElementById('a-actividad').value.trim();
  if(!actividad){toast('⚠ La actividad es obligatoria');return;}
  /* Razón social obligatoria: sin ella la actividad queda invisible para
     cualquier usuario con acceso restringido (ver razonVisible/_razonesAsignadas) */
  var _razonFinal=(function(){var v=(document.getElementById('a-razon')||{value:''}).value.trim().toUpperCase();
    if(v)return v;var rs=_razonesAsignadas();return(rs&&rs.length===1)?rs[0]:'';})();
  if(!_razonFinal){toast('⚠ La razón social es obligatoria');return;}
  function dv(elId){var v=document.getElementById(elId).value;return v||null;}
  function nv(elId){var v=document.getElementById(elId).value;return v!==''?parseFloat(v):null;}
  var _rowActRaw={
    actividad:actividad, mes:document.getElementById('a-mes').value||null,
    estado:document.getElementById('a-estado').value||null,
    categoria:document.getElementById('a-categoria').value||null,
    asignado:document.getElementById('a-asignado').value.trim()||null,
    razon:_razonFinal,
    creado_por:(_session&&(_session.nombre||_session.username))||null,
    est_inicio:dv('a-estinicio'), est_fin:dv('a-estfin'), dur_estimada:nv('a-durest'),
    real_inicio:dv('a-realinicio'), real_fin:dv('a-realfin'), dur_real:nv('a-durreal'),
    comentario:document.getElementById('a-comentario').value.trim()||null,
    apoyos:document.getElementById('a-apoyos').value.trim()||null,
    programada:document.getElementById('a-programada').checked
  };
  var row=await encObj(_rowActRaw,FIELDS.actividades);
  var client=getSbClient();
  if(!client){toast('⚠ Sin conexión a Supabase');return;}
  closeModal();
  try{
    if(isNew){
      var ri=await client.from('actividades').insert([row]);
      if(ri.error){toast('⚠ '+ri.error.message);return;}
      toast('☁️ Actividad creada en Supabase');
    } else {
      var ru=await client.from('actividades').update(row).eq('id',id);
      if(ru.error){toast('⚠ '+ru.error.message);return;}
      toast('☁️ Actividad actualizada en Supabase');
    }
    await loadActividades();
  }catch(e){toast('⚠ Error: '+e.message);}
}

/* ════════════════════════════════════════════════════════════════════
   IMPORTAR ACTIVIDADES DESDE EXCEL
   Columnas esperadas (ver plantilla): Razón, Actividad, Mes, Categoría,
   Asignado, Estado, Programada, Fecha Est Inicio, Fecha Est Fin,
   Duración Estimada, Comentario, Apoyos.
   (Fecha Real Inicio, Fecha Real Fin y Duración Real ya no forman parte de
   la plantilla; si el archivo las trae igual, el importador las sigue
   leyendo por compatibilidad, pero no son obligatorias.)
   La razón es OBLIGATORIA por fila salvo que la cuenta tenga una sola
   razón asignada (entonces se usa esa automáticamente para todas).
════════════════════════════════════════════════════════════════════ */
/* Elige la hoja correcta del libro por nombre (ignora "Instrucciones" y
   "Tiendas válidas"). Si ningún nombre coincide, busca la hoja cuya cabecera
   contenga alguna columna requerida. Devuelve {ws,raw}. */
function pickSheet(wb,names,headerKeys){
  function nm(s){return String(s||'').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/\s+/g,' ').trim();}
  var want=names.map(nm),skip=['instrucciones','tiendas validas'],order=wb.SheetNames.slice();
  function read(sn){return XLSX.utils.sheet_to_json(wb.Sheets[sn],{header:1,defval:''});}
  var byName=order.find(function(sn){return want.indexOf(nm(sn))>=0;});
  if(byName)return {ws:wb.Sheets[byName],raw:read(byName)};
  var keys=headerKeys.map(nm);
  for(var i=0;i<order.length;i++){
    if(skip.indexOf(nm(order[i]))>=0)continue;
    var raw=read(order[i]),h=(raw[0]||[]).map(nm);
    if(keys.some(function(k){return h.some(function(x){return x.indexOf(k)>=0;});}))return {ws:wb.Sheets[order[i]],raw:raw};
  }
  var f=order[0];return {ws:wb.Sheets[f],raw:read(f)};
}
function importActividadesExcel(files){
  if(!files||!files.length)return;
  if(!_session||!['admin','admin_auditor','auditor'].includes(_session.rol)){toast('⚠ Sin permisos');return;}
  var file=files[0];
  document.getElementById('act-file-input').value='';
  var reader=new FileReader();
  reader.onload=function(e){
    try{
      var wb=XLSX.read(e.target.result,{type:'array',cellDates:false});
      var _pk=pickSheet(wb,['actividades'],['actividad','proyecto']);var ws=_pk.ws,raw=_pk.raw;
      if(raw.length<2){toast('⚠ Excel vacío o sin datos');return;}
      function normH(s){return String(s||'').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/[^a-z0-9]+/g,' ').trim();}
      var H=raw[0].map(normH);
      function fc(keys){
        var nk=keys.map(normH);
        for(var ki=0;ki<nk.length;ki++)for(var hi=0;hi<H.length;hi++)if(H[hi]===nk[ki])return hi;
        for(var ki=0;ki<nk.length;ki++)for(var hi=0;hi<H.length;hi++)if(H[hi].includes(nk[ki]))return hi;
        return -1;
      }
      var iRz=fc(['razon','razon social']);
      var iAc=fc(['actividad','actividad/proyecto','actividad proyecto','proyecto']);
      var iMe=fc(['mes']);
      var iCa=fc(['categoria']);
      var iAs=fc(['asignado','asignado a']);
      var iEs=fc(['estado']);
      var iPr=fc(['programada']);
      var iEI=fc(['fecha est inicio','est inicio','fecha estimada inicio']);
      var iEF=fc(['fecha est fin','est fin','fecha estimada fin']);
      var iDE=fc(['duracion estimada','dur estimada']);
      var iRI=fc(['fecha real inicio','real inicio']);
      var iRF=fc(['fecha real fin','real fin']);
      var iDR=fc(['duracion real','dur real']);
      var iCo=fc(['comentario']);
      var iAp=fc(['apoyos']);
      if(iAc<0){toast('⚠ No se encontró la columna "Actividad". Usa la plantilla oficial.');return;}
      var _rzAsigImp=_razonesAsignadas();
      var _rzUnica=(_rzAsigImp&&_rzAsigImp.length===1)?_rzAsigImp[0]:null;
      if(iRz<0&&!_rzUnica){toast('⚠ Falta la columna "Razón" (KNO/KSC/KSA). Usa la plantilla oficial.');return;}
      function toISOfecha(v){
        if(!v&&v!==0)return null;
        var n=Number(v);
        if(!isNaN(n)&&n>30000&&n<80000){
          var ms=(n-25569)*86400*1000;var d=new Date(ms);
          return d.getUTCFullYear()+'-'+String(d.getUTCMonth()+1).padStart(2,'0')+'-'+String(d.getUTCDate()).padStart(2,'0');
        }
        var s=String(v).trim();
        var m=s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
        if(m)return m[3]+'-'+m[2].padStart(2,'0')+'-'+m[1].padStart(2,'0');
        if(/^\d{4}-\d{2}-\d{2}$/.test(s))return s;
        var dt=new Date(s);
        if(!isNaN(dt))return dt.getFullYear()+'-'+String(dt.getMonth()+1).padStart(2,'0')+'-'+String(dt.getDate()).padStart(2,'0');
        return null;
      }
      var yoSesion=(_session&&(_session.nombre||_session.username))||'';
      var rows=[], omitidas=0;
      for(var i=1;i<raw.length;i++){
        var r=raw[i];
        var actividad=String(r[iAc]||'').trim();
        if(!actividad)continue;
        var razon=(iRz>=0?String(r[iRz]||'').trim().toUpperCase():'')||_rzUnica||'';
        if(!razon){omitidas++;continue;}
        var progRaw=iPr>=0?norm(String(r[iPr]||'')):'';
        var programada=iPr<0?true:!(progRaw==='no'||progRaw==='false'||progRaw==='0');
        var _estI=iEI>=0?toISOfecha(r[iEI]):null;
        var _estF=iEF>=0?toISOfecha(r[iEF]):null;
        var _realI=iRI>=0?toISOfecha(r[iRI]):null;
        var _realF=iRF>=0?toISOfecha(r[iRF]):null;
        function _durOrNull(v){return v===''?null:v;}
        rows.push({
          razon:razon,
          actividad:actividad,
          mes:iMe>=0?String(r[iMe]||'').trim():'',
          categoria:iCa>=0?String(r[iCa]||'').trim():'',
          asignado:iAs>=0?String(r[iAs]||'').trim():'',
          estado:iEs>=0?String(r[iEs]||'').trim()||'Sin comenzar':'Sin comenzar',
          programada:programada,
          est_inicio:_estI,
          est_fin:_estF,
          dur_estimada:_durOrNull(calcDias(_estI,_estF)),
          real_inicio:_realI,
          real_fin:_realF,
          dur_real:_durOrNull(calcDias(_realI,_realF)),
          comentario:iCo>=0?String(r[iCo]||'').trim():'',
          apoyos:iAp>=0?String(r[iAp]||'').trim():'',
          creado_por:yoSesion||null
        });
      }
      if(!rows.length){toast('⚠ Sin filas válidas'+(omitidas?' ('+omitidas+' omitidas por falta de razón)':''));return;}
      toast('⏳ Importando '+rows.length+' actividad(es)'+(omitidas?' — '+omitidas+' omitidas sin razón':'')+'…');
      commitActividadesExcel(rows);
    }catch(err){toast('⚠ Error: '+err.message);console.error(err);}
  };
  reader.readAsArrayBuffer(file);
}

async function commitActividadesExcel(rows){
  var client=getSbClient();
  if(!client){toast('⚠ Sin Supabase');return;}
  try{
    var okN=0,failN=0,dupN=0;
    for(var i=0;i<rows.length;i++){
      try{
        var row=await encObj(rows[i],FIELDS.actividades);
        var r=await client.from('actividades').insert([row]);
        if(r.error){
          /* 409 / 23505 = la actividad ya existe (restricción única). No es un
             fallo real: se omite en vez de reventar la importación. */
          var msg=String(r.error.message||'')+String(r.error.code||'');
          if(/duplicate|already exists|23505|conflict/i.test(msg))dupN++;
          else { failN++; console.warn('actividad:',r.error.message); }
        } else okN++;
      }catch(e){failN++;}
    }
    toast('☁️ '+okN+' actividad(es) importadas'+
      (dupN?(' — '+dupN+' ya existían (omitidas)'):'')+
      (failN?(' — '+failN+' con error'):''));
    _actMesManual=false; /* tras importar, volver al mes en curso para ver lo recién cargado */
    await loadActividades();
  }catch(e){toast('⚠ Error: '+e.message);}
}

async function deleteActividad(id){
  if(!confirm('¿Eliminar esta actividad permanentemente?'))return;
  var client=getSbClient();
  if(!client){toast('⚠ Sin conexión');return;}
  closeModal();
  try{
    var r=await client.from('actividades').delete().eq('id',id);
    if(r.error){toast('⚠ '+r.error.message);return;}
    toast('☁️ Actividad eliminada');
    await loadActividades();
  }catch(e){toast('⚠ Error: '+e.message);}
}

/* ════════════════════════════════════════════════════════════════════
   ASIGNACIÓN MASIVA DE RAZÓN SOCIAL (Actividades sin razón)
   Permite a un admin reclasificar de un tirón las actividades antiguas
   que se cargaron sin razón (por eso quedaban ocultas para el personal
   restringido a KNO/KSC/KSA), opcionalmente acotando por "asignado".
════════════════════════════════════════════════════════════════════ */
function _actividadesSinRazon(){return ACTIVIDADES.filter(function(a){return!a.razon;});}
function openAsignarRazonMasiva(){
  var pend=_actividadesSinRazon();
  var asigs=[...new Set(pend.map(function(a){return a.asignado;}).filter(Boolean))].sort();
  var sel=document.getElementById('ar-asignado');
  sel.innerHTML='<option value="ALL">Todos los pendientes ('+pend.length+')</option>'+
    asigs.map(function(a){var n=pend.filter(function(x){return x.asignado===a;}).length;return'<option value="'+esc(a)+'">'+esc(a)+' ('+n+')</option>';}).join('');
  document.getElementById('ar-razon').value='';
  document.getElementById('ar-err').textContent='';
  refreshAsignarRazonCount();
  document.getElementById('ar-overlay').classList.add('show');
}
function closeAsignarRazonMasiva(){document.getElementById('ar-overlay').classList.remove('show');}
function refreshAsignarRazonCount(){
  var asig=document.getElementById('ar-asignado').value;
  var pend=_actividadesSinRazon().filter(function(a){return asig==='ALL'||a.asignado===asig;});
  document.getElementById('ar-count').textContent=pend.length+' actividad(es) se actualizarán con la razón elegida.';
}
async function ejecutarAsignarRazonMasiva(){
  var razon=(document.getElementById('ar-razon').value||'').trim().toUpperCase();
  var asig=document.getElementById('ar-asignado').value;
  var err=document.getElementById('ar-err');
  if(!razon){err.textContent='⚠ Elige una razón social destino';return;}
  var pend=_actividadesSinRazon().filter(function(a){return asig==='ALL'||a.asignado===asig;});
  if(!pend.length){err.textContent='⚠ No hay actividades pendientes con ese filtro';return;}
  var client=getSbClient();
  if(!client){err.textContent='⚠ Sin conexión a Supabase';return;}
  var btn=document.getElementById('ar-btn-aplicar');
  btn.disabled=true;btn.textContent='⏳ Asignando…';
  try{
    var okN=0, failN=0;
    for(var i=0;i<pend.length;i++){
      try{
        var row=await encObj({razon:razon},FIELDS.actividades);
        var r=await client.from('actividades').update(row).eq('id',pend[i].id);
        if(r.error)failN++; else okN++;
      }catch(e){failN++;}
    }
    closeAsignarRazonMasiva();
    toast('☁️ '+okN+' actividad(es) asignadas a '+razon+(failN?(' — '+failN+' con error'):''));
    await loadActividades();
  }finally{
    btn.disabled=false;btn.textContent='✓ Asignar razón';
  }
}

/* ════════════════════════════════════════════════════════════════════
   ARRANQUE  (carga datos guardados o semilla del Excel)
════════════════════════════════════════════════════════════════════ */
(function init(){
  /* No cargar seed ni localStorage al arrancar.
     Los datos se cargan desde Supabase tras el login. */
  STORE={auditorias:[],tareas:[]};
  refreshAll();
})();
/* ════════════════════════════════════════════════════════════════════
   GENERADOR POWERPOINT
════════════════════════════════════════════════════════════════════ */
/* Modal 'Generar PowerPoint' eliminado. El reporte de Cumplimiento ahora se
   dispara desde la tarjeta 'Cumplimiento por Tienda' del módulo Documentos,
   que envía el mensaje 'run-cumplimiento' al padre (ver listener más abajo). */

function generatePpt(opts){
  if(typeof PptxGenJS==='undefined'){toast('⚠ PptxGenJS no cargado');return;}
  opts=opts||{};
  var _elT=document.getElementById('ppt-titulo'), _elE=document.getElementById('ppt-empresa');
  var titulo=opts.titulo||(_elT&&_elT.value)||'Reporte de Cumplimiento';
  var empresa=opts.empresa||(_elE&&_elE.value)||'Grupo Kuroda';
  var today=new Date().toLocaleDateString('es-MX',{day:'2-digit',month:'long',year:'numeric'});

  /* ── datos: filtro propio del reporte (razón, tienda, periodo, rango de
     fechas) o, si no se envía, el filtro actual del dashboard.
     El reporte refleja tanto VIGENTES como FINALIZADAS: las vigentes son las
     auditorías activas que NO están finalizadas (para no duplicar), y se les
     suman las archivadas. ── */
  function pptFiltrar(){
    var anio=new Date().getFullYear();
    var fl = opts.filtro || (function(){
      var f=getFilterState();
      var gd=function(id){var e=document.getElementById(id);return e?e.value:'';};
      return {razon:f.razon, tienda:f.tienda, mes:(f.mesGran||'ALL'), desde:gd('f-desde'), hasta:gd('f-hasta')};
    })();
    var mi=(fl.mes&&fl.mes!=='ALL')?parseInt(fl.mes):null;
    var desde=fl.desde?new Date(fl.desde+'T00:00:00'):null;
    var hasta=fl.hasta?new Date(fl.hasta+'T23:59:59'):null;
    /* Restricción por razón social del usuario: solo su(s) razón(es). Si tiene
       las tres asignadas (admin), no se restringe. La razón se deriva del centro
       o de la tienda cuando el registro no la trae (p.ej. finalizadas). */
    var permit=(typeof razonesExportables==='function')?razonesExportables().map(razKey):null;
    var restrict=!!(permit&&permit.length>0&&permit.length<3);
    var razonDe=function(x){
      if(x.razon)return x.razon;
      var c=x.centro||((typeof centroDeTienda==='function')?centroDeTienda(x.tienda||''):'');
      return (typeof razonDeCentro==='function')?razonDeCentro(c):'';
    };
    var okR=function(x){
      var rz=razKey(razonDe(x));
      if(restrict && permit.indexOf(rz)<0)return false;
      return !fl.razon||fl.razon==='ALL'||rz===razKey(fl.razon);
    };
    var okT=function(x){return !fl.tienda||fl.tienda==='ALL'||norm(x.tienda)===norm(fl.tienda);};
    var audFecha=function(a){var d=fromISO(a.fecha);if(d)return d;var k=mesIndexFromNombre(a.mes);return k>=0?new Date(anio,k,15):null;};
    var okFA=function(a){var fa=audFecha(a);
      if(desde&&(!fa||fa<desde))return false; if(hasta&&(!fa||fa>hasta))return false;
      if(mi){if(!fa)return false;if(fa.getMonth()+1!==mi||fa.getFullYear()!==anio)return false;} return true;};
    var T=STORE.tareas.filter(function(t){
      if(!okR(t)||!okT(t))return false;
      var cr=fromISO(t.fechaCreacion);
      if(desde&&(!cr||cr<desde))return false;
      if(hasta&&(!cr||cr>hasta))return false;
      if(mi){ if(!cr)return false; if(cr.getMonth()+1!==mi||cr.getFullYear()!==anio)return false; }
      return true;
    });
    var estaFin=(typeof estaFinalizada==='function')?estaFinalizada:function(){return false;};
    var audVig=STORE.auditorias.filter(function(a){return okR(a)&&okT(a)&&okFA(a)&&!estaFin(a);});
    var audFin=(typeof FINALIZADAS!=='undefined'?FINALIZADAS:[])
      .filter(function(f){return !f._pending;})
      .map(function(f){
        var centro=f.centro||((typeof centroDeTienda==='function')?centroDeTienda(f.tienda||''):'')||'';
        return {tienda:f.tienda||'', centro:centro, razon:f.razon||'', mes:f.mes||'', clase:f.clase||'',
          fecha:f.fecha||f.fecha_inicio||'',
          pctCumpl:(f.pctCumpl!=null?f.pctCumpl:(parseFloat(f.pct_cumpl)||0)),
          tareas:(f.tareas!=null?f.tareas:(parseInt(f.total_tareas)||0)),
          resueltas:(f.resueltas!=null?f.resueltas:(parseInt(f.resueltas)||0)), _fin:true};
      })
      .filter(function(a){return okR(a)&&okT(a)&&okFA(a);});
    return {tareas:T, aud:audVig.concat(audFin), nVig:audVig.length, nFin:audFin.length};
  }
  var _pf=pptFiltrar();
  var tareas=_pf.tareas;
  var aud=_pf.aud;
  var nVig=_pf.nVig, nFin=_pf.nFin;
  var total=tareas.length||1;
  var avgCumpl=aud.length?aud.reduce(function(a,r){return a+r.pctCumpl;},0)/aud.length:0;
  var res=tareas.filter(esResuelta).length;
  var pend=tareas.filter(esPendiente).length;
  var vencidas=tareas.filter(function(t){return esPendiente(t)&&diasVenc(t)!==null&&diasVenc(t)<0;}).length;
  var pctRes=res/total;
  var resOk=tareas.filter(function(t){return esResuelta(t)&&!norm(t.estado).includes('atrasad');}).length;
  var resAtr=tareas.filter(function(t){return esResuelta(t)&&norm(t.estado).includes('atrasad');}).length;
  var abOk=tareas.filter(function(t){return esPendiente(t)&&!norm(t.estado).includes('atrasad');}).length;
  var abAtr=tareas.filter(function(t){return esPendiente(t)&&norm(t.estado).includes('atrasad');}).length;
  var sucursales=uniq(tareas.map(function(t){return t.tienda;})).length;
  var cumplPond=Math.round((aud.reduce(function(a,r){return a+r.pctCumpl*(r.tareas||1);},0)/(aud.reduce(function(a,r){return a+(r.tareas||1);},0)||1))*100);
  function pp(n){return Math.round((n||0)/total*100)+'%';}

  /* ── paleta corporativa (mismos colores que el generador) ── */
  var BLU='2563EB', TEA='0D9488', GRN='16A34A', ORG='EA580C',
      RED='DC2626', AMB='D97706', NAV='0F172A', SLT='334155',
      MUT='94A3B8', WHT='FFFFFF', OFF='F8FAFC', LIN='E2E8F0',
      PBLU='DBEAFE', PTEA='CCFBF1', PGRN='DCFCE7',
      PORG='FFEDD5', PRED='FEE2E2', PAMB='FEF3C7';
  var W=10, H=5.625;
  /* Tonos empresariales adicionales para jerarquía y encabezados */
  var NAV2='1E293B', BLU2='1D4ED8', GOLD='C9A84C', HDR='F1F5F9';

  var pptx=new PptxGenJS();
  pptx.layout='LAYOUT_16x9';
  pptx.title=titulo; pptx.company=empresa;

  /* Qué secciones se incluyen (configurable desde la tarjeta 'Cumplimiento').
     'portada' siempre se recomienda, pero también es toggle. */
  var SECCIONES=[
    {id:'kpis',        t:'KPIs de Tendencia'},
    {id:'analisis',    t:'Análisis de Tareas'},
    {id:'mapa_calor',  t:'Mapa de Calor — Tienda y Actividad'},
    {id:'distribucion',t:'Distribución de Estado'},
    {id:'menor_cumpl', t:'Sucursales con Menor Calificacion'},
    {id:'pendientes',  t:'Sucursales con más Pendientes'}
  ];
  var secSel=(opts.secciones&&opts.secciones.length)?opts.secciones:['portada'].concat(SECCIONES.map(function(x){return x.id;}));
  function secOn(id){return secSel.indexOf(id)>=0;}

  /* ── helper: slide con encabezado corporativo (barra de acento + banda sutil) ── */
  function mkSlide(ac, tag){
    var s=pptx.addSlide();
    s.background={color:WHT};
    /* barra de acento superior + hairline navy para dar peso */
    s.addShape(pptx.ShapeType.rect,{x:0,y:0,w:W,h:.07,fill:{color:ac},line:{color:ac}});
    s.addShape(pptx.ShapeType.rect,{x:0,y:.07,w:W,h:.012,fill:{color:NAV},line:{color:NAV}});
    /* pie */
    s.addShape(pptx.ShapeType.rect,{x:0,y:H-.28,w:W,h:.28,fill:{color:OFF},line:{color:OFF}});
    s.addShape(pptx.ShapeType.rect,{x:0,y:H-.28,w:W,h:.012,fill:{color:LIN},line:{color:LIN}});
    s.addText(empresa.toUpperCase(),{x:.35,y:H-.25,w:5,h:.2,fontSize:7,bold:true,color:SLT,charSpacing:2,valign:'middle'});
    s.addText(today,{x:5,y:H-.25,w:4.65,h:.2,fontSize:7,color:MUT,align:'right',valign:'middle'});
    if(tag) s.addText(tag.toUpperCase(),{x:.35,y:.14,w:9.3,h:.2,fontSize:7.5,bold:true,color:ac,charSpacing:2.5,valign:'middle'});
    return s;
  }
  /* ── helper: título de sección con acento (consistente en todas) ── */
  function slideTitle(s, text, ac){
    s.addShape(pptx.ShapeType.rect,{x:.35,y:.42,w:.09,h:.34,fill:{color:ac},line:{color:ac}});
    s.addText(text,{x:.53,y:.34,w:9.1,h:.46,fontSize:17,bold:true,color:NAV,valign:'middle'});
  }

  /* ── helper: KPI card (rediseño: acento lateral, jerarquía limpia,
     tipografía que escala con la altura de la tarjeta) ── */
  function kCard(s,x,y,w,h,val,lbl,sub,ac,bg){
    s.addShape(pptx.ShapeType.roundRect,{x:x,y:y,w:w,h:h,rectRadius:.05,fill:{color:WHT},line:{color:LIN,pt:.75},
      shadow:{type:'outer',blur:4,offset:1,angle:90,color:'CBD5E1',opacity:.28}});
    s.addShape(pptx.ShapeType.rect,{x:x,y:y+.06,w:.07,h:h-.12,fill:{color:ac},line:{color:ac}});
    var vfs=h>=1.25?32:(h>=1.0?27:22);
    s.addText(String(val),{x:x+.2,y:y+.08,w:w-.32,h:h*.5,fontSize:vfs,bold:true,color:ac,align:'left',valign:'middle'});
    s.addText(lbl,{x:x+.21,y:y+h*.57,w:w-.34,h:.24,fontSize:9,bold:true,color:NAV,valign:'middle'});
    s.addText(sub,{x:x+.21,y:y+h*.77,w:w-.34,h:.2,fontSize:7.5,color:MUT,valign:'middle'});
  }

  /* ══════════ PORTADA ══════════ */
  if(secOn('portada'))(function(){
    var s=pptx.addSlide();
    s.background={color:NAV};
    /* capas tonales para profundidad en el lado oscuro */
    s.addShape(pptx.ShapeType.rect,{x:0,y:3.15,w:4.0,h:2.475,fill:{color:NAV2},line:{color:NAV2}});
    s.addShape(pptx.ShapeType.rect,{x:0,y:4.35,w:4.0,h:1.275,fill:{color:'16233B'},line:{color:'16233B'}});
    /* filo azul + hilo dorado a la izquierda */
    s.addShape(pptx.ShapeType.rect,{x:0,y:0,w:.16,h:H,fill:{color:BLU},line:{color:BLU}});
    s.addShape(pptx.ShapeType.rect,{x:.16,y:0,w:.03,h:H,fill:{color:GOLD},line:{color:GOLD}});
    /* panel claro derecho (cubre cualquier decoración a la derecha) */
    s.addShape(pptx.ShapeType.rect,{x:4.0,y:0,w:6.0,h:H,fill:{color:WHT},line:{color:WHT}});
    s.addShape(pptx.ShapeType.rect,{x:4.0,y:0,w:6.0,h:.06,fill:{color:GOLD},line:{color:GOLD}});
    /* logo */
    s.addShape(pptx.ShapeType.roundRect,{x:.4,y:.32,w:1.04,h:1.04,rectRadius:.06,fill:{color:WHT},line:{color:'2A3A55',pt:.8}});
    var kLogo=document.querySelector('.lp-panel-logo img');
    if(kLogo&&kLogo.src&&kLogo.src.startsWith('data:')){
      s.addImage({data:kLogo.src,x:.45,y:.37,w:.94,h:.94});
    } else {
      s.addShape(pptx.ShapeType.roundRect,{x:.47,y:.39,w:.9,h:.9,rectRadius:.06,fill:{color:BLU},line:{color:'3D6BDC',pt:.8}});
      s.addText('K',{x:.47,y:.39,w:.9,h:.9,fontSize:28,bold:true,color:WHT,align:'center',valign:'middle'});
    }
    /* eyebrow + empresa */
    s.addText('REPORTE EJECUTIVO',{x:.44,y:1.5,w:3.3,h:.22,fontSize:8,bold:true,color:GOLD,charSpacing:3});
    s.addText(empresa,{x:.44,y:1.74,w:3.3,h:.34,fontSize:13,bold:true,color:'9FC1FF'});
    /* línea dorada */
    s.addShape(pptx.ShapeType.rect,{x:.44,y:2.16,w:2.6,h:.032,fill:{color:GOLD},line:{color:GOLD}});
    /* título */
    s.addText(titulo,{x:.44,y:2.32,w:3.35,h:1.3,fontSize:23,bold:true,color:WHT,fontFace:'Calibri',lineSpacingMultiple:1.14,valign:'top'});
    /* fecha */
    s.addText(today,{x:.44,y:3.72,w:3.3,h:.28,fontSize:10,color:'7FB3FF'});
    s.addText('INFORME CONFIDENCIAL',{x:.44,y:H-.4,w:3.2,h:.26,fontSize:7,color:'6B86B8',bold:true,charSpacing:1.5});

    /* métricas resumen (tarjetas limpias con acento y sombra sutil) */
    var mets=[{l:'En curso',v:nVig,c:BLU},{l:'Finalizadas',v:nFin,c:GRN},{l:'Tareas',v:tareas.length,c:TEA},{l:'Cumplimiento',v:Math.round(avgCumpl*100)+'%',c:ORG}];
    mets.forEach(function(m,i){
      var mx=4.35+i*1.4;
      s.addShape(pptx.ShapeType.roundRect,{x:mx,y:1.05,w:1.28,h:1.16,rectRadius:.05,fill:{color:WHT},line:{color:LIN,pt:.8},
        shadow:{type:'outer',blur:4,offset:1,angle:90,color:'CBD5E1',opacity:.3}});
      s.addShape(pptx.ShapeType.rect,{x:mx+.06,y:1.11,w:.05,h:1.04,fill:{color:m.c},line:{color:m.c}});
      s.addText(String(m.v),{x:mx+.15,y:1.12,w:1.08,h:.66,fontSize:24,bold:true,color:m.c,align:'left',valign:'middle'});
      s.addText(m.l.toUpperCase(),{x:mx+.16,y:1.8,w:1.08,h:.28,fontSize:7,color:SLT,bold:true,charSpacing:.3});
    });
    /* índice — solo las secciones incluidas */
    var secs=SECCIONES.filter(function(x){return secOn(x.id);}).map(function(x){return x.t;});
    s.addText('CONTENIDO',{x:4.35,y:2.62,w:5.3,h:.22,fontSize:7.5,bold:true,color:MUT,charSpacing:2.5});
    s.addShape(pptx.ShapeType.rect,{x:4.35,y:2.86,w:5.3,h:.012,fill:{color:LIN},line:{color:LIN}});
    secs.forEach(function(sec,i){
      var sy=2.96+i*.42;
      s.addText(String(i+1).padStart(2,'0'),{x:4.35,y:sy,w:.4,h:.3,fontSize:10,bold:true,color:BLU,valign:'middle'});
      s.addText(sec,{x:4.8,y:sy,w:4.85,h:.3,fontSize:10,color:NAV,valign:'middle'});
    });
  })();

  /* ══════════ 1. KPIs DE TENDENCIA (configurable) ══════════ */
  if(secOn('kpis'))(function(){
    var s=mkSlide(BLU,'KPIs de Tendencia');
    /* Catálogo completo con ID estable; el usuario elige cuáles aparecen
       desde la tarjeta 'Cumplimiento' (módulo Documentos). */
    var CAT=[
      {id:'cumpl_prom', v:Math.round(avgCumpl*100)+'%', l:'Cumplimiento prom.',   su:'promedio de auditorías',    ac:BLU, bg:PBLU},
      {id:'tareas',     v:tareas.length,                l:'Tareas en período',     su:aud.length+' auditorías',    ac:TEA, bg:PTEA},
      {id:'resueltas',  v:res,                          l:'Resueltas',             su:pp(res)+' del total',        ac:GRN, bg:PGRN},
      {id:'pendientes', v:pend,                         l:'Pendientes',            su:pp(pend)+' sin cerrar',      ac:ORG, bg:PORG},
      {id:'venc',       v:vencidas,                     l:'Pend. vencidas',        su:'fuera de fecha de término', ac:RED, bg:PRED},
      {id:'pct_res',    v:Math.round(pctRes*100)+'%',   l:'% Resolución',          su:'tareas cerradas',           ac:TEA, bg:PTEA},
      {id:'sucursales', v:sucursales,                   l:'Sucursales',            su:'con tareas en período',     ac:BLU, bg:PBLU},
      {id:'cumpl_pond', v:cumplPond+'%',                l:'Cumpl. ponderado',      su:'por nº de tareas',          ac:ORG, bg:PORG},
      {id:'res_ok',     v:resOk,                        l:'Resueltas a tiempo',    su:pp(resOk)+' del total',      ac:GRN, bg:PGRN},
      {id:'res_atr',    v:resAtr,                       l:'Resueltas atrasadas',   su:pp(resAtr)+' del total',     ac:AMB, bg:PAMB},
      {id:'ab_ok',      v:abOk,                         l:'Abiertas en plazo',     su:pp(abOk)+' del total',       ac:BLU, bg:PBLU},
      {id:'ab_atr',     v:abAtr,                        l:'Abiertas atrasadas',    su:pp(abAtr)+' del total',      ac:RED, bg:PRED}
    ];
    var sel=(opts.kpis&&opts.kpis.length)?opts.kpis:CAT.map(function(k){return k.id;});
    var kpis=CAT.filter(function(k){return sel.indexOf(k.id)>=0;});
    if(!kpis.length)kpis=CAT.slice();

    slideTitle(s,'KPIs de Tendencia',BLU);
    s.addText(kpis.length+' de '+CAT.length+' indicadores',{x:7.0,y:.4,w:2.65,h:.3,fontSize:9,color:MUT,align:'right',valign:'middle'});

    /* Grid responsivo: columnas y tamaño de tarjeta se ajustan al nº de KPIs. */
    var n=kpis.length;
    var cols = n<=2?n : n<=4?n : n<=6?3 : 4;
    var rows=Math.ceil(n/cols);
    var gx=.14, gy=.14, sx=.35, top=.92;
    var availW=9.30, availH=(H-.30)-top;
    var cw=(availW-(cols-1)*gx)/cols;
    var ch=Math.min(1.5,(availH-(rows-1)*gy)/rows);
    var blockH=rows*ch+(rows-1)*gy;
    var sy=top+Math.max(0,(availH-blockH)/2);
    kpis.forEach(function(k,i){
      var r=Math.floor(i/cols), c=i%cols;
      var inRow=(r===rows-1)?(n-cols*r):cols;
      var rowW=inRow*cw+(inRow-1)*gx;
      var rx=sx+Math.max(0,(availW-rowW)/2);
      kCard(s, rx+c*(cw+gx), sy+r*(ch+gy), cw, ch, k.v, k.l, k.su, k.ac, k.bg);
    });
  })();

  /* ══════════ 2. ANÁLISIS DE TAREAS ══════════ */
  if(secOn('analisis'))(function(){
    var s=mkSlide(AMB,'Análisis de Tareas');
    slideTitle(s,'Análisis de Tareas',AMB);
    var nameOf=function(t){return (t.nombre&&t.nombre.trim())||(t.actividad&&t.actividad.trim())||'Sin nombre';};
    var freq={};
    tareas.forEach(function(t){var k=nameOf(t);freq[k]=(freq[k]||0)+1;});
    var topFreq=Object.entries(freq).sort(function(a,b){return b[1]-a[1];}).slice(0,6);
    var byName={};
    tareas.forEach(function(t){
      var k=nameOf(t); if(!byName[k])byName[k]={total:0,bad:0,good:0};
      byName[k].total++;
      var n=norm(t.estado);
      if(n.includes('abierta')&&n.includes('atrasad'))byName[k].bad++;
      else if(n.includes('resuelta')&&!n.includes('atrasad'))byName[k].good++;
    });
    var peor=Object.entries(byName).sort(function(a,b){return b[1].total-a[1].total;}).slice(0,6);
    var mejor=Object.entries(byName).sort(function(a,b){return a[1].total-b[1].total;}).slice(0,6);
    var rk=[BLU,TEA,ORG,'7C3AED',RED,AMB];
    var cols3=[
      {lbl:'🔁 Tareas más frecuentes',ac:BLU,items:topFreq,vf:function(e){return e[1]+'x';},tc:BLU},
      {lbl:'🔴 Peor calificadas',     ac:RED,items:peor,   vf:function(e){return e[1].total+'x';},tc:RED},
      {lbl:'✅ Mejor calificadas',    ac:GRN,items:mejor,  vf:function(e){return e[1].total+'x';},tc:GRN}
    ];
    var colW=3.0, rh=.55, sy=.86;
    cols3.forEach(function(col,ci){
      var cx=.32+ci*(colW+.09);
      s.addShape(pptx.ShapeType.rect,{x:cx,y:sy,w:colW,h:.3,fill:{color:col.ac},line:{color:col.ac}});
      s.addText(col.lbl,{x:cx+.1,y:sy,w:colW-.14,h:.3,fontSize:8.5,bold:true,color:WHT,valign:'middle'});
      col.items.forEach(function(e,i){
        var ry=sy+.3+i*rh;
        s.addShape(pptx.ShapeType.rect,{x:cx,y:ry,w:colW,h:rh-.04,fill:{color:i%2===0?WHT:OFF},line:{color:LIN,pt:.4}});
        /* badge blanco con borde */
        s.addShape(pptx.ShapeType.rect,{x:cx+.07,y:ry+.1,w:.24,h:.24,fill:{color:WHT},line:{color:rk[i],pt:.8}});
        s.addText(String(i+1),{x:cx+.07,y:ry+.1,w:.24,h:.24,fontSize:8,bold:true,color:rk[i],align:'center',valign:'middle'});
        s.addText(String(e[0]).slice(0,34),{x:cx+.37,y:ry+.08,w:colW-.5,h:.22,fontSize:8,color:NAV,bold:i===0});
        s.addText(col.vf(e),{x:cx+.37,y:ry+.31,w:colW-.5,h:.18,fontSize:7.5,color:col.tc,bold:true});
      });
    });
  })();

  /* ══════════ 3. MAPA DE CALOR — TIENDA × ACTIVIDAD ══════════ */
  if(secOn('mapa_calor'))(function(){
    var s=mkSlide(TEA,'Mapa de Calor — Tienda y Actividad');
    slideTitle(s,'Cumplimiento por Tienda y Actividad',TEA);
    var nameOf=function(t){return (t.nombre&&t.nombre.trim())||(t.actividad&&t.actividad.trim())||'Sin nombre';};

    var freqAct={};
    tareas.forEach(function(t){var k=nameOf(t);freqAct[k]=(freqAct[k]||0)+1;});
    var topAct=Object.entries(freqAct).sort(function(a,b){return b[1]-a[1];}).slice(0,5).map(function(e){return e[0];});

    var freqStore={};
    tareas.forEach(function(t){var k=t.tienda||'Sin tienda';freqStore[k]=(freqStore[k]||0)+1;});
    var topStores=Object.entries(freqStore).sort(function(a,b){return b[1]-a[1];}).slice(0,7).map(function(e){return e[0];});

    if(!topAct.length||!topStores.length){
      s.addText('Sin tareas registradas en el período.',{x:.35,y:2.5,w:9.3,h:.4,fontSize:12,color:MUT,align:'center'});
      return;
    }

    function pctCell(store,act){
      var subset=tareas.filter(function(t){return (t.tienda||'Sin tienda')===store&&nameOf(t)===act;});
      if(!subset.length)return null;
      return Math.round(subset.filter(esResuelta).length/subset.length*100);
    }
    function heatBg(v){if(v==null)return LIN;if(v>=80)return GRN;if(v>=60)return AMB;if(v>=40)return ORG;return RED;}
    function heatFg(v){if(v==null)return MUT;return WHT;}

    var firstColW=1.95, sy=.86, headerH=.42;
    var cellW=(9.32-firstColW)/topAct.length;
    var cellH=Math.min(.46,(H-.34-sy-headerH)/topStores.length);

    s.addShape(pptx.ShapeType.rect,{x:.34,y:sy,w:firstColW,h:headerH,fill:{color:NAV},line:{color:WHT,pt:.5}});
    s.addText('Sucursal',{x:.34,y:sy,w:firstColW,h:headerH,fontSize:8,bold:true,color:WHT,align:'left',valign:'middle'});
    topAct.forEach(function(act,j){
      var x=.34+firstColW+j*cellW;
      s.addShape(pptx.ShapeType.rect,{x:x,y:sy,w:cellW,h:headerH,fill:{color:NAV},line:{color:WHT,pt:.5}});
      s.addText(act.slice(0,22),{x:x+.03,y:sy,w:cellW-.06,h:headerH,fontSize:6.5,bold:true,color:WHT,align:'center',valign:'middle'});
    });

    topStores.forEach(function(store,i){
      var y=sy+headerH+i*cellH;
      s.addShape(pptx.ShapeType.rect,{x:.34,y:y,w:firstColW,h:cellH-.02,fill:{color:i%2===0?OFF:WHT},line:{color:LIN,pt:.4}});
      s.addText(store.slice(0,20),{x:.4,y:y,w:firstColW-.08,h:cellH-.02,fontSize:8,bold:true,color:NAV,valign:'middle'});
      topAct.forEach(function(act,j){
        var v=pctCell(store,act);
        var x=.34+firstColW+j*cellW;
        s.addShape(pptx.ShapeType.rect,{x:x,y:y,w:cellW-.02,h:cellH-.02,fill:{color:heatBg(v)},line:{color:WHT,pt:.75}});
        s.addText(v!=null?v+'%':'—',{x:x,y:y,w:cellW-.02,h:cellH-.02,fontSize:9,bold:true,color:heatFg(v),align:'center',valign:'middle'});
      });
    });

    var legend=[{c:GRN,l:'≥80% Óptimo'},{c:AMB,l:'60–79% Aceptable'},{c:ORG,l:'40–59% En riesgo'},{c:RED,l:'<40% Crítico'}];
    var ly=sy+headerH+topStores.length*cellH+.14;
    legend.forEach(function(g,i){
      var lx=.34+i*2.35;
      s.addShape(pptx.ShapeType.rect,{x:lx,y:ly,w:.2,h:.15,fill:{color:g.c},line:{color:g.c}});
      s.addText(g.l,{x:lx+.28,y:ly-.03,w:2.05,h:.2,fontSize:7,color:SLT,valign:'middle'});
    });
    s.addText('% de tareas resueltas por tienda y actividad (top '+topAct.length+' actividades por frecuencia)',{x:.34,y:ly+.24,w:9.3,h:.2,fontSize:7,italic:true,color:MUT});
  })();

  /* ══════════ 4. DISTRIBUCIÓN DONUT ══════════ */
  if(secOn('distribucion'))(function(){
    var s=mkSlide(TEA,'Distribución de Estado de Tareas');
    slideTitle(s,'Distribución de Estado de Tareas',TEA);
    var segs=[
      {lbl:'Resueltas a tiempo',  cnt:resOk,  ac:GRN},
      {lbl:'Resueltas atrasadas', cnt:resAtr, ac:ORG},
      {lbl:'Abiertas en plazo',   cnt:abOk,   ac:BLU},
      {lbl:'Abiertas atrasadas',  cnt:abAtr,  ac:RED}
    ];
    var tot3=tareas.length||1;
    /* Donut chart nativo de PowerPoint */
    var chartData=[{name:'Estado',labels:segs.map(function(g){return g.lbl;}),values:segs.map(function(g){return g.cnt;})}];
    s.addChart(pptx.ChartType.doughnut, chartData, {
      x:.3, y:.78, w:4.5, h:4.2,
      holeSize:60,
      showLegend:false, showLabel:false, showValue:false, showPercent:false,
      chartColors:[GRN,ORG,BLU,RED]
    });
    /* leyenda derecha — igual que el generador */
    segs.forEach(function(g,i){
      var ly=1.2+i*1.0;
      var pct3=Math.round(g.cnt/tot3*100);
      s.addShape(pptx.ShapeType.ellipse,{x:4.95,y:ly+.1,w:.22,h:.22,fill:{color:g.ac},line:{color:g.ac}});
      s.addText(g.lbl,{x:5.26,y:ly,w:3.4,h:.28,fontSize:11,color:NAV,valign:'middle'});
      s.addText(String(g.cnt),{x:8.7,y:ly,w:.95,h:.28,fontSize:13,bold:true,color:g.ac,align:'right',valign:'middle'});
      s.addText('('+pct3+'%)',{x:5.26,y:ly+.28,w:3.4,h:.2,fontSize:9,color:MUT});
    });
    s.addText('Total: '+tareas.length+' tareas',{x:4.95,y:4.7,w:4.7,h:.2,fontSize:8,color:MUT,align:'center'});
  })();

  /* ══════════ 5. SUCURSALES MENOR CUMPLIMIENTO ══════════ */
  if(secOn('menor_cumpl'))(function(){
    var s=mkSlide(RED,'Sucursales con Menor Calificacion');
    slideTitle(s,'Sucursales con Menor Calificacion',RED);
    var byS={};
    aud.forEach(function(a){var k=a.tienda||a.centro;if(!byS[k])byS[k]={s:0,n:0,c:a.centro};byS[k].s+=a.pctCumpl;byS[k].n++;});
    var arr=Object.entries(byS).map(function(e){return{t:e[0],c:e[1].c,v:Math.round(e[1].s/e[1].n*100)};}).sort(function(a,b){return a.v-b.v;}).slice(0,9);
    if(!arr.length){
      s.addText('Sin auditorías registradas.',{x:.35,y:2.5,w:9.3,h:.4,fontSize:12,color:MUT,align:'center'});
      return;
    }
    var sy=.88, rh=.43;
    s.addShape(pptx.ShapeType.rect,{x:.34,y:sy,w:9.32,h:.3,fill:{color:NAV},line:{color:NAV}});
    [{t:'#',x:.42,w:.28,a:'center'},{t:'Sucursal',x:.77,w:3.9,a:'left'},{t:'Centro',x:4.72,w:2.3,a:'left'},{t:'% Cumpl.',x:7.08,w:.9,a:'right'},{t:'Progreso',x:8.05,w:1.55,a:'left'}].forEach(function(h){
      s.addText(h.t,{x:h.x,y:sy+.05,w:h.w,h:.2,fontSize:7.5,bold:true,color:WHT,align:h.a,valign:'middle',charSpacing:.8});
    });
    arr.forEach(function(r,i){
      var y=sy+.3+i*rh;
      var rc=r.v>=70?GRN:r.v>=50?AMB:RED;
      s.addShape(pptx.ShapeType.rect,{x:.34,y:y,w:9.32,h:rh-.04,fill:{color:i%2===0?WHT:OFF},line:{color:LIN,pt:.4}});
      /* badge blanco con borde */
      s.addShape(pptx.ShapeType.rect,{x:.42,y:y+.09,w:.26,h:.22,fill:{color:WHT},line:{color:rc,pt:.8}});
      s.addText(String(i+1),{x:.42,y:y+.09,w:.26,h:.22,fontSize:8.5,bold:true,color:rc,align:'center',valign:'middle'});
      s.addText(r.t,{x:.75,y:y+.09,w:3.9,h:.22,fontSize:9.5,bold:true,color:NAV});
      s.addText(r.c||'—',{x:4.7,y:y+.09,w:2.3,h:.22,fontSize:9,color:SLT});
      s.addText(r.v+'%',{x:7.06,y:y+.09,w:.94,h:.22,fontSize:11,bold:true,color:rc,align:'right'});
      s.addShape(pptx.ShapeType.rect,{x:8.03,y:y+.15,w:1.55,h:.1,fill:{color:LIN},line:{color:LIN}});
      s.addShape(pptx.ShapeType.rect,{x:8.03,y:y+.15,w:Math.max(.02,r.v/100*1.55),h:.1,fill:{color:rc},line:{color:rc}});
    });
  })();

  /* ══════════ 6. SUCURSALES MÁS PENDIENTES ══════════ */
  if(secOn('pendientes'))(function(){
    var s=mkSlide(ORG,'Sucursales con más Pendientes');
    slideTitle(s,'Sucursales con más Pendientes',ORG);
    var byStore={};
    tareas.filter(esPendiente).forEach(function(t){
      var k=t.tienda||'Sin tienda';
      if(!byStore[k])byStore[k]={cnt:0,venc:0};
      byStore[k].cnt++;
      if(diasVenc(t)!==null&&diasVenc(t)<0)byStore[k].venc++;
    });
    var topS=Object.entries(byStore).map(function(e){return{t:e[0],cnt:e[1].cnt,venc:e[1].venc};}).sort(function(a,b){return b.cnt-a.cnt;}).slice(0,9);
    if(!topS.length){
      s.addText('No hay tareas pendientes.',{x:.35,y:2.5,w:9.3,h:.4,fontSize:12,color:MUT,align:'center'});
      return;
    }
    var maxC=topS[0].cnt||1, sy=.88, rh=.43;
    s.addShape(pptx.ShapeType.rect,{x:.34,y:sy,w:9.32,h:.3,fill:{color:NAV},line:{color:NAV}});
    [{t:'#',x:.42,w:.28,a:'center'},{t:'Sucursal',x:.77,w:4.2,a:'left'},{t:'Pendientes',x:5.03,w:2.6,a:'left'},{t:'Vencidas',x:7.7,w:1.92,a:'right'}].forEach(function(h){
      s.addText(h.t,{x:h.x,y:sy+.05,w:h.w,h:.2,fontSize:7.5,bold:true,color:WHT,align:h.a,valign:'middle',charSpacing:.8});
    });
    topS.forEach(function(r,i){
      var y=sy+.3+i*rh;
      s.addShape(pptx.ShapeType.rect,{x:.34,y:y,w:9.32,h:rh-.04,fill:{color:i%2===0?WHT:OFF},line:{color:LIN,pt:.4}});
      s.addShape(pptx.ShapeType.rect,{x:.42,y:y+.09,w:.26,h:.22,fill:{color:WHT},line:{color:ORG,pt:.8}});
      s.addText(String(i+1),{x:.42,y:y+.09,w:.26,h:.22,fontSize:8.5,bold:true,color:ORG,align:'center',valign:'middle'});
      s.addText(r.t,{x:.75,y:y+.09,w:4.22,h:.22,fontSize:9.5,bold:true,color:NAV});
      s.addShape(pptx.ShapeType.rect,{x:5.01,y:y+.15,w:2.55,h:.1,fill:{color:LIN},line:{color:LIN}});
      s.addShape(pptx.ShapeType.rect,{x:5.01,y:y+.15,w:Math.max(.03,r.cnt/maxC*2.55),h:.1,fill:{color:ORG},line:{color:ORG}});
      s.addText(String(r.cnt),{x:5.01,y:y+.09,w:.7,h:.22,fontSize:10,bold:true,color:ORG});
      s.addText(String(r.venc),{x:7.68,y:y+.09,w:1.94,h:.22,fontSize:10,bold:true,color:r.venc>0?RED:MUT,align:'right'});
    });
  })();

  toast('⏳ Generando PowerPoint…');
  var fname='Cumplimiento_'+new Date().toISOString().slice(0,10)+'.pptx';
  pptx.writeFile({fileName:fname}).then(function(){toast('✓ PowerPoint descargado');});
}

/* ════════════════════════════════════════════════════════════════════
   GENERADOR PPT ACTIVIDADES AUDITORÍA
════════════════════════════════════════════════════════════════════ */


/* generateActPpt eliminado: reporte de Actividades Auditoría retirado por solicitud. */




/* ════════════════════════════════════════════════════════════════════
   CIFRADO AES-GCM — Datos en reposo cifrados antes de llegar a Supabase
   La clave se deriva de la contraseña del usuario con PBKDF2.
   Campos de ID, fechas, números y condición NO se cifran (needed por Supabase).
   Todo el texto sensible (tiendas, nombres, actividades, etc.) SÍ se cifra.
════════════════════════════════════════════════════════════════════ */
/* ── CLAVE COMPARTIDA DE ORGANIZACIÓN ──────────────────────────────────
   ANTES: la clave se derivaba de la contraseña de CADA usuario y de un salt
   aleatorio guardado en localStorage. Consecuencia: cada usuario derivaba una
   clave distinta y NO podía leer lo que cifró otro (el viewer veía el texto
   cifrado en pantalla), y si se limpiaba el localStorage se perdía el acceso
   a los datos para siempre.

   AHORA: la clave se deriva de un secreto FIJO de la organización y un salt
   FIJO. Todos los usuarios —incluidos los nuevos y los de solo lectura—
   derivan la MISMA clave y leen los mismos datos, desde cualquier navegador.

   Nota de seguridad: al ser una clave común de la app, el cifrado protege los
   datos EN REPOSO en Supabase (quien vea la tabla no lee texto plano), pero no
   es un secreto frente a alguien que inspeccione el código del sitio. Es el
   compromiso necesario para que varios usuarios compartan los mismos datos.

   Se conserva una clave "legacy" derivada de la contraseña para poder LEER los
   datos antiguos cifrados con el esquema anterior (retrocompatibilidad). */
var _cryptoKey=null;  /* Clave compartida — cifra y descifra todo lo nuevo */
var _legacyKey=null;  /* Clave antigua (contraseña del usuario) — solo lectura */
var _cryptoSalt=null;

const CRYPTO_SALT_KEY='ksa_crypto_salt';
const PBKDF2_ITER=200000;
/* Secreto y salt compartidos de la organización (fijos, iguales para todos) */
const ORG_SECRET='KurodaGrupo::clave-compartida::v1';
const ORG_SALT='KurodaGrupo::salt::v1';

function getCryptoSalt(){
  var s=localStorage.getItem(CRYPTO_SALT_KEY);
  if(s)return Uint8Array.from(atob(s),c=>c.charCodeAt(0));
  return null; /* sin salt guardado no hay clave legacy que reconstruir */
}

async function _derivar(password,salt){
  var enc=new TextEncoder();
  var base=await crypto.subtle.importKey('raw',enc.encode(password),{name:'PBKDF2'},false,['deriveKey']);
  return crypto.subtle.deriveKey(
    {name:'PBKDF2',salt:salt,iterations:PBKDF2_ITER,hash:'SHA-256'},
    base,{name:'AES-GCM',length:256},false,['encrypt','decrypt']
  );
}

/* Clave compartida: NO depende de la contraseña ni del navegador. */
async function initSharedKey(){
  if(_cryptoKey)return _cryptoKey;
  _cryptoKey=await _derivar(ORG_SECRET,new TextEncoder().encode(ORG_SALT));
  return _cryptoKey;
}

/* En el login se derivan ambas: la compartida (siempre) y, si existe el salt
   antiguo en este navegador, la legacy para poder leer datos sin migrar. */
async function initCryptoKey(password){
  await initSharedKey();
  try{
    var salt=getCryptoSalt();
    if(salt&&password)_legacyKey=await _derivar(password,salt);
  }catch(e){ _legacyKey=null; }
}

async function enc(text){
  if(!_cryptoKey||text===null||text===undefined||text==='')return text;
  var iv=crypto.getRandomValues(new Uint8Array(12));
  var enc2=new TextEncoder();
  var ct=await crypto.subtle.encrypt({name:'AES-GCM',iv},_cryptoKey,enc2.encode(String(text)));
  /* Guardar como "iv_b64:ct_b64" */
  return btoa(String.fromCharCode(...iv))+':'+btoa(String.fromCharCode(...new Uint8Array(ct)));
}

async function dec(cipher){
  if(!cipher||typeof cipher!=='string'||!cipher.includes(':'))return cipher;
  if(!_cryptoKey&&!_legacyKey)return cipher;
  var parts=cipher.split(':');
  if(parts.length!==2)return cipher;
  var iv,ct;
  try{
    iv=Uint8Array.from(atob(parts[0]),c=>c.charCodeAt(0));
    ct=Uint8Array.from(atob(parts[1]),c=>c.charCodeAt(0));
  }catch(e){ return cipher; }
  /* 1) Clave compartida (datos nuevos / ya migrados) */
  if(_cryptoKey){
    try{
      var pt=await crypto.subtle.decrypt({name:'AES-GCM',iv},_cryptoKey,ct);
      return new TextDecoder().decode(pt);
    }catch(e){}
  }
  /* 2) Clave legacy (datos cifrados con el esquema anterior, aún sin migrar) */
  if(_legacyKey){
    try{
      var pt2=await crypto.subtle.decrypt({name:'AES-GCM',iv},_legacyKey,ct);
      return new TextDecoder().decode(pt2);
    }catch(e){}
  }
  return cipher; /* no descifrable con ninguna clave */
}

/* ════════════════════════════════════════════════════════════════════
   MIGRACIÓN DE CIFRADO — re-cifra los datos guardados con el esquema
   ANTIGUO (clave por contraseña) usando la CLAVE COMPARTIDA, para que
   todos los usuarios (viewers y cuentas nuevas) puedan leerlos.
   La ejecuta un admin desde el navegador donde SÍ se ven los datos
   (es decir, donde la clave legacy funciona). Es idempotente: lo que ya
   está en el esquema nuevo se deja intacto.
════════════════════════════════════════════════════════════════════ */
var TABLAS_CIFRADAS={
  tareas:'tareas', auditorias:'auditorias', actividades:'actividades',
  ajustes:'ajustes', mermas:'mermas', tareas_finalizadas:'tareas_finalizadas'
};

async function _abreCon(key,val){
  if(!key||!val||typeof val!=='string'||!val.includes(':'))return false;
  var p=val.split(':'); if(p.length!==2)return false;
  try{
    var iv=Uint8Array.from(atob(p[0]),c=>c.charCodeAt(0));
    var ct=Uint8Array.from(atob(p[1]),c=>c.charCodeAt(0));
    await crypto.subtle.decrypt({name:'AES-GCM',iv},key,ct);
    return true;
  }catch(e){ return false; }
}
/* ¿Este texto se descifra SOLO con la clave legacy? (=> hay que migrarlo) */
async function _soloLegacy(val){
  if(!val||typeof val!=='string'||!val.includes(':'))return false;
  var parts=val.split(':'); if(parts.length!==2)return false;
  var iv,ct;
  try{
    iv=Uint8Array.from(atob(parts[0]),c=>c.charCodeAt(0));
    ct=Uint8Array.from(atob(parts[1]),c=>c.charCodeAt(0));
  }catch(e){ return false; }
  if(_cryptoKey){
    try{ await crypto.subtle.decrypt({name:'AES-GCM',iv},_cryptoKey,ct); return false; }catch(e){}
  }
  if(_legacyKey){
    try{ await crypto.subtle.decrypt({name:'AES-GCM',iv},_legacyKey,ct); return true; }catch(e){}
  }
  return false;
}

/* Diagnóstico: recorre las tablas y reporta, por cada una, cuántos registros
   están en texto plano, cuántos abre la clave compartida, cuántos abre la clave
   de ESTE usuario (legacy) y cuántos no abre ninguna (cifrados por otro usuario). */
async function diagnosticarCifrado(){
  var client=getSbClient();
  if(!client){toast('⚠ Sin conexión a Supabase');return;}
  await initSharedKey();
  var st=document.getElementById('mig-status');
  if(st)st.textContent='Analizando…';
  function fb(s){return Uint8Array.from(atob(s),c=>c.charCodeAt(0));}
  async function abre(key,val){
    if(!key)return false;
    try{var p=val.split(':');await crypto.subtle.decrypt({name:'AES-GCM',iv:fb(p[0])},key,fb(p[1]));return true;}catch(e){return false;}
  }
  var rep=[];
  for(var t in TABLAS_CIFRADAS){
    var campos=FIELDS[t]; if(!campos||!campos.length)continue;
    try{
      var r=await client.from(TABLAS_CIFRADAS[t]).select('*').limit(300);
      if(r.error){rep.push(t+': ⚠ '+r.error.message);continue;}
      var filas=r.data||[];
      var plano=0,sh=0,lg=0,no=0;
      for(var i=0;i<filas.length;i++){
        var v=String(filas[i][campos[0]]||'');
        if(!v||!v.includes(':')){plano++;continue;}
        if(await abre(_cryptoKey,v))sh++;
        else if(await abre(_legacyKey,v))lg++;
        else no++;
      }
      rep.push(t+' ('+filas.length+'): plano '+plano+' · compartida '+sh+' · tu clave '+lg+' · ilegibles '+no);
    }catch(e){rep.push(t+': ⚠ '+e.message);}
  }
  if(st)st.textContent='';
  alert('DIAGNÓSTICO DE CIFRADO\n(usuario: '+((_session&&_session.username)||'?')+')\n\n'+rep.join('\n')+
    '\n\n• "plano" = visible para todos.\n• "compartida" = ya migrado, visible para todos.\n• "tu clave" = lo puedes migrar ahora con el botón.\n• "ilegibles" = cifrado con la contraseña de OTRO usuario: debe migrarlo esa persona desde su navegador, o volver a importar el Excel.');
}

/* ════════════════════════════════════════════════════════════════════
   LIMPIEZA DE REGISTROS ILEGIBLES
   Registros cifrados con la clave de un usuario que ya no se puede
   reconstruir: nadie —ni el admin— puede leerlos. Como la importación de
   Excel deduplica por 'tienda|fechas' y la tienda está cifrada, reimportar
   sin limpiar generaría DUPLICADOS. Esta función descarga primero un
   respaldo y luego borra solo esos registros, dejando la tabla lista para
   reimportar el Excel (los nuevos ya se cifran con la clave compartida).
════════════════════════════════════════════════════════════════════ */
async function limpiarIlegibles(){
  if(!_session||!['admin','admin_auditor'].includes(_session.rol)){toast('⚠ Solo un admin');return;}
  var client=getSbClient();
  if(!client){toast('⚠ Sin conexión a Supabase');return;}
  await initSharedKey();
  var st=document.getElementById('mig-status');
  if(st)st.textContent='Analizando…';

  var porTabla={}, respaldo={}, totalIleg=0;
  for(var t in TABLAS_CIFRADAS){
    var campos=FIELDS[t]; if(!campos||!campos.length)continue;
    var r=await client.from(TABLAS_CIFRADAS[t]).select('*').limit(20000);
    if(r.error){console.warn(t,r.error.message);continue;}
    var ileg=[];
    var filas=r.data||[];
    for(var i=0;i<filas.length;i++){
      var v=String(filas[i][campos[0]]||'');
      if(!v||!v.includes(':'))continue;                       /* texto plano: se respeta */
      if(await _abreCon(_cryptoKey,v))continue;               /* ya migrado */
      if(await _abreCon(_legacyKey,v))continue;               /* migrable: no se toca */
      ileg.push(filas[i]);                                    /* nadie puede leerlo */
    }
    if(ileg.length){ porTabla[t]=ileg.length; respaldo[t]=ileg; totalIleg+=ileg.length; }
  }
  if(st)st.textContent='';
  if(!totalIleg){ toast('✓ No hay registros ilegibles'); return; }

  var detalle=Object.keys(porTabla).map(function(k){return '• '+k+': '+porTabla[k];}).join('\n');
  if(!confirm('Se encontraron '+totalIleg+' registro(s) que NADIE puede descifrar:\n\n'+detalle+
    '\n\nSe descargará un respaldo (.json) y luego se BORRARÁN de la base para que puedas reimportar el Excel sin duplicados.\n\n¿Continuar?'))return;

  /* 1) Respaldo antes de borrar */
  try{
    var blob=new Blob([JSON.stringify(respaldo,null,1)],{type:'application/json'});
    var a=document.createElement('a');
    a.href=URL.createObjectURL(blob);
    a.download='respaldo_registros_ilegibles_'+(new Date().toISOString().slice(0,10))+'.json';
    document.body.appendChild(a); a.click(); a.remove();
  }catch(e){ if(!confirm('No se pudo descargar el respaldo ('+e.message+'). ¿Borrar de todos modos?'))return; }

  if(!confirm('Último paso: se borrarán '+totalIleg+' registro(s) ilegibles.\n\nEsta acción no se puede deshacer (ya tienes el respaldo .json).\n\n¿Confirmas?'))return;

  /* 2) Borrado */
  if(st)st.textContent='Borrando…';
  var borrados=0, errores=0;
  for(var tt in respaldo){
    var ids=respaldo[tt].map(function(x){return x.id;}).filter(Boolean);
    for(var j=0;j<ids.length;j+=50){
      var lote=ids.slice(j,j+50);
      var del=await client.from(TABLAS_CIFRADAS[tt]).delete().in('id',lote);
      if(del.error){errores++;console.warn(tt,del.error.message);}else borrados+=lote.length;
    }
  }
  var msg='✓ '+borrados+' registro(s) borrados'+(errores?' · '+errores+' error(es)':'')+'. Ahora reimporta el Excel.';
  toast(msg); if(st)st.textContent=msg;
  setTimeout(function(){location.reload();},2500);
}

async function migrarCifrado(){
  if(!_session||_session.rol==='viewer'){toast('⚠ Tu cuenta es de solo lectura');return;}
  var st=document.getElementById('mig-status');
  function setSt(t){if(st)st.textContent=t;}
  await initSharedKey();
  var okLegacy=await asegurarLegacyKey();
  if(!okLegacy){
    alert('No se pudo obtener la clave antigua en este navegador.\n\nEjecuta la migración desde el navegador del admin donde los datos SÍ se ven correctamente (ahí está guardado el salt original) e ingresa su contraseña.');
    return;
  }
  if(!confirm('Se volverán a cifrar los datos antiguos con la clave compartida para que TODOS los usuarios (incluidos los de solo lectura) puedan verlos.\n\nSe recomienda hacerlo una sola vez. ¿Continuar?'))return;
  var client=getSbClient();
  if(!client){toast('⚠ Sin conexión a Supabase');return;}
  setSt('Migrando…');
  var totalMig=0, errores=0, ilegibles=0;
  for(var t in TABLAS_CIFRADAS){
    var campos=FIELDS[t]; if(!campos)continue;
    try{
      var r=await client.from(TABLAS_CIFRADAS[t]).select('*').limit(5000);
      if(r.error){console.warn('Migración '+t+':',r.error.message);errores++;continue;}
      var filas=r.data||[];
      for(var i=0;i<filas.length;i++){
        var fila=filas[i], upd={}, hay=false;
        for(var c=0;c<campos.length;c++){
          var campo=campos[c], val=fila[campo];
          if(await _soloLegacy(val)){
            var plano=await dec(val);        /* abre con legacy */
            upd[campo]=await enc(plano);     /* re-cifra con la compartida */
            hay=true;
          }else if(val&&typeof val==='string'&&val.includes(':')&&!(await _abreCon(_cryptoKey,val))&&!(await _abreCon(_legacyKey,val))){
            ilegibles++;
          }
        }
        if(hay){
          var up=await client.from(TABLAS_CIFRADAS[t]).update(upd).eq('id',fila.id);
          if(up.error){errores++;console.warn('Update '+t+' id '+fila.id+':',up.error.message);}
          else totalMig++;
        }
      }
    }catch(e){errores++;console.warn('Migración '+t+':',e.message);}
  }
  var msg;
  if(totalMig>0) msg='✓ Migrados '+totalMig+' campo(s)'+(errores?' · '+errores+' error(es)':'');
  else if(ilegibles>0) msg='⚠ '+ilegibles+' campo(s) cifrados con la clave de OTRO usuario: no los puedes migrar tú';
  else if(errores) msg='⚠ No se migró nada · '+errores+' error(es)';
  else msg='✓ No había datos pendientes de migrar';
  toast(msg); setSt(msg);
  if(totalMig===0&&ilegibles>0){
    alert('Se encontraron '+ilegibles+' campo(s) que NO se pueden abrir con tu clave.\n\nFueron cifrados con la contraseña de otro usuario (el que cargó esos datos).\n\nSoluciones:\n1) Que ESA persona inicie sesión en SU navegador (el mismo donde cargó los datos) y pulse "Migrar datos ahora".\n2) O vuelve a importar el Excel de esos módulos: los datos nuevos ya se guardan con la clave compartida y todos los verán.');
  }
  if(totalMig>0)setTimeout(function(){location.reload();},1500);
}

/* Cifrar un objeto — solo los campos especificados */
async function encObj(obj,fields){
  var r=Object.assign({},obj);
  for(var f of fields){if(r[f]!=null&&r[f]!=='')r[f]=await enc(r[f]);}
  return r;
}

/* Descifrar un objeto — solo los campos especificados */
async function decObj(obj,fields){
  var r=Object.assign({},obj);
  for(var f of fields){if(r[f]!=null&&r[f]!=='')r[f]=await dec(r[f]);}
  return r;
}

/* Descifrar array */
async function decArr(arr,fields){
  var res=[];
  for(var obj of arr)res.push(await decObj(obj,fields));
  return res;
}

/* Campos a cifrar por tabla */
var FIELDS={
  tareas:['tienda','area_resp','area_rev','actividad','nombre','tipo_tarea','estado','razon','centro'],
  auditorias:['tienda','razon','centro','clase','mes'],
  actividades:['actividad','estado','categoria','asignado','comentario','apoyos','mes','creado_por','razon'],
  ajustes:['tienda','mes','condicion','auditor'],
  mermas:['tienda','local_foranea','mes','condicion','auditor'],
  tareas_finalizadas:['tienda','clase','mes','completado_por'],
  seguimiento_semanas:['tienda','clase','actividad']
};

/* ════════════════════════════════════════════════════════════════════
   SEGUIMIENTO DE CARGAS EXCEL — atribuye cada importación al auditor
   que la realizó (sesión activa), para reflejarlo en Desempeño.
   Requiere la tabla public.cargas_excel (ver SQL entregado).
   Degradación elegante: si la tabla no existe, CARGAS queda vacío.
════════════════════════════════════════════════════════════════════ */
var CARGAS=[];
async function loadCargas(){
  var client=getSbClient();
  if(!client)return;
  try{
    var r=await client.from('cargas_excel').select('*').order('fecha',{ascending:false}).limit(5000);
    if(r.error){console.warn('cargas_excel no disponible:',r.error.message);return;}
    CARGAS=r.data||[];
  }catch(e){console.warn('loadCargas:',e);}
}
async function registrarCargaExcel(nTareas,nAuditorias,snapshot){
  try{
    var aud=(_session&&(_session.nombre||_session.username))||null;
    var nT=parseInt(nTareas)||0, nA=parseInt(nAuditorias)||0;
    if(!aud||(nT+nA)===0)return;
    var client=getSbClient();
    if(!client)return;
    var row={auditor:aud,n_tareas:nT,n_auditorias:nA,total:nT+nA};
    if(snapshot){row.snapshot=snapshot;row.revertida=false;}
    var r=await client.from('cargas_excel').insert([row]);
    if(r.error){
      /* Posible que falten las columnas snapshot/revertida → reintentar sin ellas
         para no perder el registro de la carga */
      var r2=await client.from('cargas_excel').insert([{auditor:aud,n_tareas:nT,n_auditorias:nA,total:nT+nA}]);
      if(r2.error)console.warn('registrarCargaExcel:',r2.error.message);
    }
  }catch(e){console.warn('registrarCargaExcel:',e);}
}

/* ── Revertir una carga: deshace inserciones y restaura el estado previo ──
   Solo la carga MÁS RECIENTE (para no pisar cambios posteriores), admin, con confirmación. */
async function revertirCarga(id){
  var rol=_session&&_session.rol;
  if(!['admin','admin_auditor'].includes(rol)){toast('⚠ Solo administradores pueden revertir');return;}
  var client=getSbClient();
  if(!client){toast('⚠ Sin conexión a Supabase');return;}
  toast('⏳ Verificando carga…');
  var r=await client.from('cargas_excel').select('*').eq('id',id).limit(1);
  if(r.error||!r.data||!r.data.length){toast('⚠ No se encontró la carga');return;}
  var c=r.data[0];
  if(c.revertida){toast('Esta carga ya fue revertida');return;}
  if(!c.snapshot){toast('⚠ Esta carga no tiene respaldo (se hizo antes de activar esta función)');return;}
  /* Verificar que sea la carga más reciente NO revertida */
  var rl=await client.from('cargas_excel').select('id,fecha,revertida').eq('revertida',false).order('fecha',{ascending:false}).limit(1);
  if(!rl.error&&rl.data&&rl.data.length&&String(rl.data[0].id)!==String(id)){
    toast('⚠ Solo puede revertirse la última carga');return;
  }
  var s=c.snapshot;
  var resumen='Se revertirá la carga del <b>'+(c.fecha?String(c.fecha).split('T')[0]:'—')+'</b> (por '+esc(c.auditor||'—')+'):<br><br>'+
    '• <b>'+((s.t_new||[]).length)+'</b> tarea(s) nueva(s) se eliminarán<br>'+
    '• <b>'+((s.t_prev||[]).length)+'</b> tarea(s) volverán a su valor anterior<br>'+
    '• <b>'+((s.a_new||[]).length)+'</b> auditoría(s) nueva(s) se eliminarán<br>'+
    '• <b>'+((s.a_prev||[]).length)+'</b> auditoría(s) volverán a su valor anterior<br><br>'+
    '<b style="color:var(--k-red)">Esta acción no se puede deshacer.</b>';
  openModal('↩ Revertir última carga','<p style="font-size:13px;color:var(--txt);line-height:1.7">'+resumen+'</p>',[
    {label:'Cancelar',cls:'btn-ghost',fn:closeModal},
    {label:'Sí, revertir',cls:'btn-red',fn:function(){ejecutarRevertCarga(id,s);}}
  ]);
}
async function ejecutarRevertCarga(id,s){
  closeModal();
  var client=getSbClient();
  if(!client){toast('⚠ Sin conexión a Supabase');return;}
  toast('⏳ Revirtiendo…');
  var err=null;
  try{
    var tn=(s.t_new||[]).map(String);
    var tnKeys=tn.filter(function(x){return x.indexOf('|')>=0;});
    var tnIds=tn.filter(function(x){return x.indexOf('|')<0;}); /* snapshots antiguos */
    for(var i=0;i<tnKeys.length&&!err;i+=200){
      var d=await client.from('tareas').delete().in('tarea_key',tnKeys.slice(i,i+200));
      if(d.error)err='Eliminar tareas: '+d.error.message;
    }
    for(var i2=0;i2<tnIds.length&&!err;i2+=200){
      var d2=await client.from('tareas').delete().in('tarea_id',tnIds.slice(i2,i2+200));
      if(d2.error)err='Eliminar tareas: '+d2.error.message;
    }
    if(!err&&(s.t_prev||[]).length){
      var up=await client.from('tareas').upsert(s.t_prev,{onConflict:'tarea_key'});
      if(up.error)err='Restaurar tareas: '+up.error.message;
    }
    for(var k=0;k<(s.a_new||[]).length&&!err;k++){
      var key=s.a_new[k];
      var da=await client.from('auditorias').delete()
        .ilike('razon',key.razon||'').ilike('centro',key.centro||'').ilike('tienda',key.tienda||'').eq('fecha',key.fecha);
      if(da.error)err='Eliminar auditorías: '+da.error.message;
    }
    if(!err&&(s.a_prev||[]).length){
      var ua=await client.from('auditorias').upsert(s.a_prev,{onConflict:'razon,centro,tienda,fecha'});
      if(ua.error)err='Restaurar auditorías: '+ua.error.message;
    }
    if(!err){
      var mu=await client.from('cargas_excel').update({revertida:true}).eq('id',id);
      if(mu.error)err='Marcar revertida: '+mu.error.message;
    }
  }catch(e){err=e.message;}
  if(err){toast('⚠ '+err);return;}
  await loadDataFromSupabase();
  await loadCargas();
  if(VIEW==='desempeno')renderDesempeno();
  if(VIEW==='auditorias')renderAuditoriasView();
  toast('✓ Carga revertida correctamente');
}


/* ════════════════════════════════════════════════════════════════════
   MÓDULO DESEMPEÑO DE AUDITORES
════════════════════════════════════════════════════════════════════ */
/* Lista de auditores conocidos (cargada al listar usuarios) */
var _AUDITORES_CONOCIDOS=[];

/* Carga la lista de auditores DIRECTAMENTE de la tabla de usuarios, para que
   el módulo Desempeño la tenga siempre disponible (antes solo se llenaba al
   abrir el modal de Usuarios, que un viewer nunca abre). Se refresca sola:
   cualquier usuario nuevo con rol de auditor aparece sin tocar el código. */
var _audLoading=false, _audCargados=false;
var _AUDITORES_META={}; /* norm(nombre) -> {nombre, razones:[] | null} */
async function cargarAuditoresDesempeno(){
  if(_audLoading||_audCargados)return;
  _audLoading=true;
  try{
    if(!_sb)_sb=initSupabase();
    if(!_sb){_audLoading=false;return;}
    var r=await _sb.rpc('listar_usuarios');
    if(r.error)throw r.error;
    var data=r.data||[];
    /* Es auditor si está marcado como tal o si su rol lo es */
    var lista=data.filter(function(u){
      return u.activo&&(u.es_auditor||['auditor','admin_auditor'].includes(u.rol));
    });
    var nombres=[]; _AUDITORES_META={};
    for(var i=0;i<lista.length;i++){
      var u=lista[i];
      var n=u.nombre||u.username||'';
      if(pareceCifrado(n)){ try{ n=await dec(n); }catch(e){} }
      if(!n||pareceCifrado(n))continue;
      /* Razones sociales asignadas a ese auditor (vacío/null = todas) */
      var rz=u.razones_permitidas;
      if(typeof rz==='string'){try{rz=JSON.parse(rz);}catch(e){rz=null;}}
      _AUDITORES_META[norm(n)]={nombre:n,razones:(rz&&rz.length)?rz:null};
      nombres.push(n);
    }
    _AUDITORES_CONOCIDOS=[...new Set(nombres)].sort();
    _audCargados=true;
    _audLoading=false;
    if(VIEW==='desempeno')renderDesempeno();
  }catch(e){ _audLoading=false; console.warn('cargarAuditoresDesempeno:',e.message); }
}

function getAuditores(){
  /* Prioridad: lista de usuarios marcados como es_auditor */
  if(_AUDITORES_CONOCIDOS.length)return _AUDITORES_CONOCIDOS;
  /* Fallback: inferir desde datos existentes */
  var fromAct=ACTIVIDADES.map(function(a){return a.creadoPor||'';}).filter(Boolean);
  var fromAj=AJUSTES.map(function(a){return a.auditor||'';}).filter(Boolean);
  var fromMr=MERMAS.map(function(m){return m.auditor||'';}).filter(Boolean);
  return [...new Set([...fromAct,...fromAj,...fromMr])].filter(Boolean).sort();
}

/* Alias de auditor: en Ajustes/Mermas/Actividades el nombre a veces se
   captura corto o parcial ("Fernando", "OSCAR") en vez del nombre completo
   registrado en la tabla de usuarios ("FERNANDO GUERRERO", "OSCAR AMAVIZCA").
   Sin esto, el módulo de Desempeño los trata como DOS auditores distintos
   (se ve duplicado en el selector y las estadísticas quedan partidas).
   _AUDITOR_ALIAS_NORM mapea cada prefijo de nombre (por token: "oscar",
   "oscar amavizca", ...) al nombre completo registrado, PERO solo cuando ese
   prefijo pertenece a un único auditor conocido. Si dos auditores distintos
   comparten el mismo primer nombre, el prefijo queda ambiguo y NO se fusiona
   — se prefiere mostrarlos separados a adivinar mal y mezclar personas. */
var _AUDITOR_ALIAS_NORM={};
function _buildAuditorAliasIndex(){
  _AUDITOR_ALIAS_NORM={};
  var porPrefijo={};
  (_AUDITORES_CONOCIDOS||[]).forEach(function(nombreCompleto){
    var tokens=norm(nombreCompleto).split(' ').filter(Boolean);
    for(var i=1;i<=tokens.length;i++){
      var prefijo=tokens.slice(0,i).join(' ');
      porPrefijo[prefijo]=porPrefijo[prefijo]||[];
      if(porPrefijo[prefijo].indexOf(nombreCompleto)<0)porPrefijo[prefijo].push(nombreCompleto);
    }
  });
  Object.keys(porPrefijo).forEach(function(pref){
    if(porPrefijo[pref].length===1)_AUDITOR_ALIAS_NORM[pref]=porPrefijo[pref][0];
  });
}

function renderDesempeno(){
  var el=document.getElementById('view-desempeno');
  if(!el)return;
  cargarAuditoresDesempeno(); /* se auto-refresca al terminar */
  _buildAuditorAliasIndex(); /* se reconstruye en cada render: barato (lista corta) y siempre al día */
  var MORD=['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'];
  var mesAct=MORD[new Date().getMonth()];
  var mesFilter=(document.getElementById('desp-mes')||{}).value||mesAct;
  var buscar=norm(((document.getElementById('desp-buscar')||{}).value||'').trim());
  /* preservar foco/caret del buscador entre re-renders */
  var _bEl=document.getElementById('desp-buscar');
  var _bFoco=_bEl&&document.activeElement===_bEl;
  var _bCaret=_bEl?_bEl.selectionStart:null;

  function mMatch(m){ return mesFilter==='ALL'||norm(String(m||''))===norm(mesFilter); }

  /* RAZÓN SOCIAL EN DESEMPEÑO — se respeta SIEMPRE el aislamiento de datos:
     un auditor solo ve registros de SU(S) razón(es) social(es). Únicamente los
     roles admin y viewer ven las tres razones completas. Por tanto, el módulo
     usa el filtro global de razón tal como quede aplicado a cada cuenta (a los
     auditores se les fija y bloquea en su razón). Consecuencia esperada: los
     totales de un mismo auditor pueden diferir según el alcance de quien
     consulta — es el precio correcto de no mezclar razones. */
  var _razonSelDesp=(document.getElementById('f-razon')||{}).value||'ALL';
  function rzOk(rec){
    if(_razonSelDesp==='ALL')return true;
    var r=rec.razon||razonDeCentro(centroDeTienda(rec.tienda||''));
    return razKey(r)===razKey(_razonSelDesp);
  }

  /* ── Agrupar SOLO por auditor (sin separar por razón) — clave normalizada ── */
  var byAuditor={};
  var nombreReal={}; /* guardar el primer nombre encontrado con su capitalización original */
  function resolveAuditorKey(k){
    if(pareceCifrado(k)){
      var conocido=_AUDITORES_CONOCIDOS.find(function(n){return !pareceCifrado(n);});
      if(conocido)k=conocido;
      else k='Auditor ('+k.substring(0,8)+'...)';
    }
    var kn=norm(k);
    /* Fusionar variante corta ("Fernando") con el nombre completo registrado
       ("FERNANDO GUERRERO") si es inequívoco — ver _buildAuditorAliasIndex. */
    if(_AUDITOR_ALIAS_NORM[kn]){ k=_AUDITOR_ALIAS_NORM[kn]; kn=norm(k); }
    return {k:k,kn:kn};
  }
  /* actividadEnTiempoDesempeno() es global (compartida con la tabla de
     Actividades) para que ambas vistas coincidan siempre en el mismo criterio. */
  function ensure(k){
    var _rk=resolveAuditorKey(k);k=_rk.k;
    var kn=_rk.kn;
    if(!byAuditor[kn]){
      byAuditor[kn]={
        nombre:k,
        actProgTotal:0,actProgOk:0,actProgAtrasada:0,actProgRepr:0,actProgPend:0,
        actNoPTotal:0,actNoPOk:0,actNoPAtrasada:0,actNoPRepr:0,actNoPPend:0,
        actProgEval:0,actNoPEval:0,actEnCurso:0,actVencidas:0,
        ajTotal:0,ajOk:0,ajTard:0,ajSinValidar:0,
        mrTotal:0,mrOk:0,mrTard:0,mrSinValidar:0,
        cgTotal:0,cgTareas:0,cgAuditorias:0,cgRegistros:0,cgUltima:null
      };
    }
    return byAuditor[kn];
  }

  /* Actividades — el campo asignado puede traer VARIOS auditores: la actividad
     se contabiliza para todos ellos (trabajo conjunto = mérito conjunto). */
  ACTIVIDADES.filter(function(a){return mMatch(a.mes)&&rzOk(a);}).forEach(function(a){
    var auds=auditoresDeActividad(a);
    if(!auds.length&&a.creadoPor)auds=[a.creadoPor];
    if(!auds.length)return;
    auds.forEach(function(aud){
    var d=ensure(aud);
    var n=norm(a.estado||'');
    var prog=(a.programada!==false);
    var completada=n.includes('completad');
    var repro=n.includes('reprogramad');
    /* Evaluable = completada (a tiempo o tarde) o bien pendiente/reprogramada
       cuyo plazo YA venció. Lo que sigue abierto y aún no vence no penaliza:
       es trabajo en curso, no un incumplimiento. */
    var vencida=(!completada&&actividadVencidaDesempeno(a));
    if(prog){
      d.actProgTotal++;
      if(completada){
        d.actProgEval++;
        if(actividadEnTiempoDesempeno(a))d.actProgOk++;else d.actProgAtrasada++;
      }else{
        if(repro)d.actProgRepr++;else d.actProgPend++;
        if(vencida){d.actProgEval++;d.actVencidas++;}else d.actEnCurso++;
      }
    }else{
      d.actNoPTotal++;
      if(completada){
        d.actNoPEval++;
        if(actividadEnTiempoDesempeno(a))d.actNoPOk++;else d.actNoPAtrasada++;
      }else{
        if(repro)d.actNoPRepr++;else d.actNoPPend++;
        if(vencida){d.actNoPEval++;d.actVencidas++;}else d.actEnCurso++;
      }
    }
    }); /* cierra auds.forEach */
  });

  /* Ajustes — condicion '—' (sin fecha de validación aún) no es "Destiempo":
     es un registro todavía abierto, no una falta del auditor. Se cuenta aparte
     y no entra al denominador de puntualidad, para no penalizar lo que aún
     no se ha podido resolver. */
  AJUSTES.filter(function(a){return mMatch(a.mes)&&rzOk(a);}).forEach(function(a){
    var aud=a.auditor||'';
    if(!aud)return;
    var d=ensure(aud);
    if(a.condicion==='A tiempo'){d.ajTotal++;d.ajOk++;}
    else if(a.condicion==='Destiempo'){d.ajTotal++;d.ajTard++;}
    else{d.ajSinValidar++;}
  });

  /* Mermas — mismo criterio que Ajustes */
  MERMAS.filter(function(m){return mMatch(m.mes)&&rzOk(m);}).forEach(function(m){
    var aud=m.auditor||'';
    if(!aud)return;
    var d=ensure(aud);
    if(m.condicion==='A tiempo'){d.mrTotal++;d.mrOk++;}
    else if(m.condicion==='Destiempo'){d.mrTotal++;d.mrTard++;}
    else{d.mrSinValidar++;}
  });
  
  /* Cargas de Excel: cargas_excel no tiene columna de razón/tienda. Para no
     mostrar auditores de otras razones (usuario restringido o razón seleccionada
     en el filtro), solo se suman cargas a auditores que YA tengan registros
     visibles de otras fuentes (sí filtradas por razón); nunca se crea tarjeta
     nueva por esto. */
  var _restringidoDesp=(!!_razonesAsignadas()&&!(_session&&_session.rol==='admin'))||_razonSelDesp!=='ALL';
  (CARGAS||[]).filter(function(c){
    if(mesFilter==='ALL')return true;
    var f=c.fecha?new Date(c.fecha):null;
    return f&&!isNaN(f)&&norm(MORD[f.getMonth()])===norm(mesFilter);
  }).forEach(function(c){
    var aud=c.auditor||'';
    if(!aud)return;
    var _rk=resolveAuditorKey(aud);
    if(_restringidoDesp&&!byAuditor[_rk.kn])return;
    var d=ensure(aud);
    var nT=parseInt(c.n_tareas)||0, nA=parseInt(c.n_auditorias)||0;
    d.cgTotal++;
    d.cgTareas+=nT; d.cgAuditorias+=nA;
    d.cgRegistros+=(parseInt(c.total)|| (nT+nA));
    var f=c.fecha?new Date(c.fecha):null;
    if(f&&!isNaN(f)&&(!d.cgUltima||f>d.cgUltima))d.cgUltima=f;
  });

  /* ¿Este auditor pertenece al alcance de razones que el usuario puede ver?
     - Si el usuario ve todas las razones (admin/viewer, filtro en ALL) → sí.
     - Si el auditor no tiene razones asignadas (ve todo) → sí.
     - Si hay filtro/restricción de razón → solo si el auditor tiene esa razón.
     Así, un auditor de otra razón social NO aparece (ni siquiera en ceros). */
  function auditorEnAlcance(nombre){
    var meta=_AUDITORES_META[norm(nombre)];
    var alcance=(_razonSelDesp!=='ALL')?[_razonSelDesp]:_razonesAsignadas();
    if(!alcance)return true;              /* el usuario ve todas las razones */
    if(!meta||!meta.razones)return true;  /* auditor sin restricción de razón */
    return meta.razones.some(function(r){
      return alcance.some(function(a){return razKey(r)===razKey(a);});
    });
  }
  /* Los auditores dados de alta aparecen aunque no tengan registros en el mes,
     PERO solo los que están dentro del alcance de razón social del usuario. */
  (_AUDITORES_CONOCIDOS||[]).forEach(function(n){ if(n&&auditorEnAlcance(n))ensure(n); });

  var auditores=Object.values(byAuditor).sort(function(a,b){return a.nombre.localeCompare(b.nombre);});
  var totalAuditores=auditores.length;
  /* Opciones del selector: auditores dados de alta + los que tengan registros.
     Se deduplica SIN distinguir mayúsculas/acentos ("Fernando Guerrero" y
     "FERNANDO GUERRERO" son la misma persona) y se conserva el nombre tal como
     está registrado en la tabla de usuarios. */
  var _optMap={};
  (_AUDITORES_CONOCIDOS||[]).forEach(function(n){
    if(n&&!pareceCifrado(n)&&auditorEnAlcance(n))_optMap[norm(n)]=n; /* el de la BD tiene prioridad */
  });
  auditores.forEach(function(d){
    var n=d.nombre;
    if(n&&!pareceCifrado(n)&&!_optMap[norm(n)])_optMap[norm(n)]=n;
  });
  var _listaAuditoresOpts=Object.values(_optMap).sort(function(a,b){return a.localeCompare(b);});
  /* Filtro por auditor (selector) */
  var audSel=(document.getElementById('desp-auditor')||{}).value||'ALL';
  if(audSel!=='ALL')auditores=auditores.filter(function(d){return norm(d.nombre)===norm(audSel);});
  if(buscar)auditores=auditores.filter(function(d){return norm(d.nombre).includes(buscar);});
  var colores=['#2563eb','#0d9488','#16a34a','#ea580c','#7c3aed','#dc2626','#d97706'];

  function pct(ok,tot){return tot>0?Math.round(ok/tot*100):100;}
  function bar(p,col){
    return '<div style="background:var(--rowline);border-radius:4px;height:7px;overflow:hidden;margin-top:5px">'+
      '<div style="width:'+Math.min(p,100)+'%;background:'+col+';height:100%;border-radius:4px"></div></div>';
  }
  function colPct(p){return p>=80?'#16a34a':p>=50?'#d97706':'#dc2626';}
  var _MCG=['ene','feb','mar','abr','may','jun','jul','ago','sep','oct','nov','dic'];
  function fmtCargaFecha(f){try{return f.getDate()+' '+_MCG[f.getMonth()]+' '+f.getFullYear();}catch(e){return '—';}}
  /* Última carga revertible (más reciente, no revertida, con respaldo) y rol admin */
  var isAdminDesp=_session&&['admin','admin_auditor'].includes(_session.rol);
  var _ultCarga=null;
  (CARGAS||[]).forEach(function(c){
    if(c.revertida||!c.snapshot)return;
    var f=c.fecha?new Date(c.fecha):null;
    if(!f||isNaN(f))return;
    if(!_ultCarga||f>_ultCarga.f)_ultCarga={id:c.id,f:f,audKn:norm(c.auditor||'')};
  });

  /* ═══ BASE COMÚN DE CALIFICACIÓN (misma fórmula para todos) ═══
     Puntualidad ponderada: Ajustes y Mermas pesan 2 (puntualidad dura),
     Actividades pesa 1. Se pondera sobre REGISTROS, no sobre promedios de
     porcentajes, así el resultado es comparable entre auditores aunque cada
     uno tenga distinta mezcla de áreas. Solo entra lo EVALUABLE (lo que sigue
     abierto y aún no vence queda fuera). */
  var PESO={aj:2, mr:2, act:1};
  function baseEval(d){
    var actEval=d.actProgEval+d.actNoPEval;
    var actOk=d.actProgOk+d.actNoPOk;
    return {
      wOk: d.ajOk*PESO.aj + d.mrOk*PESO.mr + actOk*PESO.act,
      wTot: d.ajTotal*PESO.aj + d.mrTotal*PESO.mr + actEval*PESO.act,
      n: d.ajTotal + d.mrTotal + actEval
    };
  }
  /* Promedio del equipo (para suavizar a quien tiene pocos registros) */
  var _teamOk=0,_teamTot=0;
  Object.values(byAuditor).forEach(function(d){var b=baseEval(d);_teamOk+=b.wOk;_teamTot+=b.wTot;});
  var promEquipo=_teamTot>0?(_teamOk/_teamTot*100):100;
  /* Carga máxima del equipo, para el bono por volumen */
  var _maxN=0;
  Object.values(byAuditor).forEach(function(d){var b=baseEval(d);if(b.n>_maxN)_maxN=b.n;});

  var cards=auditores.map(function(d,i){
    var col=colores[i%colores.length];
    var pAj=pct(d.ajOk,d.ajTotal);
    var pMr=pct(d.mrOk,d.mrTotal);
    /* Act. Programadas/No Programadas: el % se calcula sobre el TOTAL asignado
       (no solo sobre lo "evaluable"), para que una actividad aún pendiente sí
       reduzca el porcentaje en vez de mostrar 100% mientras falte una por
       completar. El "Global ponderado" del auditor (baseEval) sigue usando
       solo lo evaluable, ya que ese cálculo es de puntualidad, no de avance. */
    var pActProg=pct(d.actProgOk,d.actProgTotal);
    var pActNoP=pct(d.actNoPOk,d.actNoPTotal);
    var actTotal=d.actProgTotal+d.actNoPTotal;
    var actEval=d.actProgEval+d.actNoPEval;
    var actOk=d.actProgOk+d.actNoPOk;
    var totAll=d.ajTotal+d.mrTotal+actTotal;
    var totOk=d.ajOk+d.mrOk+actOk;
    var pAct=pct(actOk,actEval);

    /* ── Global de desempeño del auditor ──
       1) PUNTUALIDAD: ratio ponderado sobre registros evaluables (misma
          fórmula para todos). Ajustes/Mermas pesan 2, Actividades 1.
          Lo que aún no vence NO entra: no se penaliza trabajo en curso.
       2) SUAVIZADO: con pocos registros el % es poco confiable, así que se
          acerca al promedio del equipo (media bayesiana, K=5 registros).
       3) VOLUMEN: bono de hasta +5 pts según la carga de trabajo real
          respecto al auditor con más registros. Evita que quien hace mucho
          quede debajo de quien hizo 1 solo registro perfecto.
       Auditorías y Tareas son seguimiento y NO cuentan. */
    var _b=baseEval(d);
    var medibles=_b.n;
    /* Global = PUNTUALIDAD PONDERADA REAL, calculada EXACTAMENTE IGUAL para
       todos los auditores: (registros a tiempo ÷ registros evaluables), con
       Ajustes y Mermas pesando 2 y Actividades 1. Sin bonos ni suavizados que
       distorsionen: si un KPI está en 80%, el global lo refleja. Lo que aún no
       vence no entra al denominador (no se penaliza el trabajo en curso). */
    var pPunt=_b.wTot>0?(_b.wOk/_b.wTot*100):0;
    var pGlobal=medibles>0?Math.round(pPunt):0;
    var muestraChica=medibles>0&&medibles<5;
    var pPuntR=Math.round(pPunt);

    function kpiBlock(icon,label,total,ok,tard,repr,pend,pbar,col2,subfooter){
      return '<div style="background:var(--soft);border-radius:10px;padding:12px;border-top:3px solid '+col2+'">'+
        '<div style="font-size:9.5px;font-weight:700;color:var(--muted);text-transform:uppercase;letter-spacing:.04em;margin-bottom:6px">'+icon+' '+label+'</div>'+
        '<div style="font-size:26px;font-weight:900;color:'+col2+';line-height:1">'+total+'</div>'+
        '<div style="display:flex;gap:4px;margin-top:6px;flex-wrap:wrap">'+
          (ok!==null?'<span style="font-size:10px;background:#dcfce7;color:#15803d;padding:2px 7px;border-radius:10px;font-weight:700">✓ '+ok+'</span>':'')+
          (tard?'<span style="font-size:10px;background:#fee2e2;color:#b91c1c;padding:2px 7px;border-radius:10px;font-weight:700">✗ '+tard+'</span>':'')+
          (repr?'<span style="font-size:10px;background:#dbeafe;color:#1d4ed8;padding:2px 7px;border-radius:10px;font-weight:700">↺ '+repr+'</span>':'')+
          (pend?'<span style="font-size:10px;background:#fee2e2;color:#b91c1c;padding:2px 7px;border-radius:10px;font-weight:700">⏳ '+pend+'</span>':'')+
        '</div>'+
        bar(pbar,col2)+
        '<div style="font-size:10px;color:var(--muted);margin-top:4px;font-weight:700">'+subfooter+'</div>'+
      '</div>';
    }

    return '<div class="card" style="padding:0;overflow:hidden">'+
      '<div style="background:'+col+';padding:14px 18px;display:flex;align-items:center;gap:14px">'+
        '<div style="width:42px;height:42px;border-radius:50%;background:rgba(255,255,255,.22);display:flex;align-items:center;justify-content:center;font-size:18px;font-weight:900;color:#fff;flex-shrink:0">'+
          d.nombre.charAt(0).toUpperCase()+
        '</div>'+
        '<div style="flex:1;min-width:0">'+
          '<div style="font-size:15px;font-weight:800;color:#fff;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">'+d.nombre+'</div>'+
          '<div style="font-size:11px;color:rgba(255,255,255,.75);font-weight:600">'+(mesFilter==='ALL'?'Todos los meses':mesFilter)+' · '+totAll+' registros</div>'+
        '</div>'+
        /* Mini-badges % por KPI */
        '<div style="display:flex;flex-direction:column;gap:6px;flex-shrink:0;min-width:220px">'+
          '<div style="display:flex;gap:6px;flex-wrap:wrap;justify-content:flex-end">'+
            (d.ajTotal?'<div style="background:rgba(255,255,255,.18);border-radius:8px;padding:4px 10px;text-align:center"><div style="font-size:13px;font-weight:900;color:#fff">'+pAj+'%</div><div style="font-size:8px;color:rgba(255,255,255,.7);font-weight:700">AJU</div></div>':'')+
            (d.mrTotal?'<div style="background:rgba(255,255,255,.18);border-radius:8px;padding:4px 10px;text-align:center"><div style="font-size:13px;font-weight:900;color:#fff">'+pMr+'%</div><div style="font-size:8px;color:rgba(255,255,255,.7);font-weight:700">MER</div></div>':'')+
            (d.actProgTotal?'<div style="background:rgba(255,255,255,.18);border-radius:8px;padding:4px 10px;text-align:center"><div style="font-size:13px;font-weight:900;color:#fff">'+pActProg+'%</div><div style="font-size:8px;color:rgba(255,255,255,.7);font-weight:700">ACT-P</div></div>':'')+
            (d.actNoPTotal?'<div style="background:rgba(255,255,255,.18);border-radius:8px;padding:4px 10px;text-align:center"><div style="font-size:13px;font-weight:900;color:#fff">'+pActNoP+'%</div><div style="font-size:8px;color:rgba(255,255,255,.7);font-weight:700">ACT-NP</div></div>':'')+
          '</div>'+
          '<div style="background:rgba(255,255,255,.25);border-radius:10px;padding:8px 14px;text-align:center" title="Global = puntualidad ponderada sobre '+medibles+' registro(s) evaluable(s). Misma fórmula para todos los auditores: registros a tiempo ÷ registros evaluables, con Ajustes y Mermas pesando doble que Actividades. Lo que aún no vence no penaliza'+(d.actEnCurso?' ('+d.actEnCurso+' actividad(es) en curso)':'')+'. Promedio del equipo: '+Math.round(promEquipo)+'%. Auditorías y Tareas son seguimiento, no cuentan.">'+
            '<div style="font-size:24px;font-weight:900;color:#fff;line-height:1">'+pGlobal+'%'+
              (muestraChica?'<span style="font-size:11px;font-weight:700;vertical-align:super" title="Muestra chica: menos de 5 registros, el % es poco representativo">*</span>':'')+
            '</div>'+
            '<div style="font-size:9px;color:rgba(255,255,255,.8);font-weight:700;text-transform:uppercase;margin-top:1px">Global ponderado'+(muestraChica?' · '+medibles+' reg.':'')+'</div>'+
          '</div>'+
        '</div>'+
      '</div>'+
      '<div style="padding:14px 16px;display:grid;grid-template-columns:1fr 1fr;gap:10px">'+
        /* Fila 1: Ajustes, Mermas */
        kpiBlock('⚖️','Ajustes',d.ajTotal,d.ajOk,d.ajTard,0,d.ajSinValidar,pAj,'#0d9488',pAj+'% a tiempo'+(d.ajSinValidar?' · '+d.ajSinValidar+' sin validar':''))+
        kpiBlock('🗂️','Mermas',d.mrTotal,d.mrOk,d.mrTard,0,d.mrSinValidar,pMr,'#7c3aed',pMr+'% a tiempo'+(d.mrSinValidar?' · '+d.mrSinValidar+' sin validar':''))+
      '</div>'+
      /* Fila 2: Actividades Programadas y No Programadas */
      '<div style="padding:0 16px 14px;display:grid;grid-template-columns:1fr 1fr;gap:10px">'+
        kpiBlock('📋','Act. Programadas',d.actProgTotal,d.actProgOk,d.actProgAtrasada,d.actProgRepr,d.actProgPend,pActProg,'#2563eb',(d.actProgTotal?pActProg+'% completado sobre '+d.actProgTotal+' total':'sin actividades aún')+(d.actProgAtrasada?' · '+d.actProgAtrasada+' tarde':'')+((d.actProgTotal-d.actProgEval)>0?' · '+(d.actProgTotal-d.actProgEval)+' en curso':''))+
        kpiBlock('📌','Act. No Programadas',d.actNoPTotal,d.actNoPOk,d.actNoPAtrasada,d.actNoPRepr,d.actNoPPend,pActNoP,'#ea580c',(d.actNoPTotal?pActNoP+'% completado sobre '+d.actNoPTotal+' total':'sin actividades aún')+(d.actNoPAtrasada?' · '+d.actNoPAtrasada+' tarde':'')+((d.actNoPTotal-d.actNoPEval)>0?' · '+(d.actNoPTotal-d.actNoPEval)+' en curso':''))+
      '</div>'+
      /* Fila 3: Seguimiento de cargas de Excel — solo se muestra si hay cargas
         en el período; si no, se omite el bloque por completo (antes se
         mostraba una caja vacía "Sin cargas en este período" en cada tarjeta). */
      (d.cgTotal>0?(
      '<div style="padding:0 16px 14px">'+
        '<div style="background:var(--soft);border-radius:10px;padding:12px;border-top:3px solid #0891b2">'+
          '<div style="font-size:9.5px;font-weight:700;color:var(--muted);text-transform:uppercase;letter-spacing:.04em;margin-bottom:8px">📥 Seguimiento de cargas (Excel)</div>'+
            '<div style="display:flex;align-items:flex-end;gap:20px;flex-wrap:wrap">'+
              '<div><div style="font-size:26px;font-weight:900;color:#0891b2;line-height:1">'+d.cgTotal+'</div><div style="font-size:9px;color:var(--muted);font-weight:700">CARGAS</div></div>'+
              '<div><div style="font-size:20px;font-weight:800;color:var(--txt);line-height:1">'+d.cgRegistros+'</div><div style="font-size:9px;color:var(--muted);font-weight:700">REGISTROS</div></div>'+
              '<div style="display:flex;gap:5px;flex-wrap:wrap">'+
                (d.cgTareas?'<span style="font-size:10px;background:#e0f2fe;color:#075985;padding:2px 8px;border-radius:10px;font-weight:700">'+d.cgTareas+' tareas</span>':'')+
                (d.cgAuditorias?'<span style="font-size:10px;background:#fef3c7;color:#92400e;padding:2px 8px;border-radius:10px;font-weight:700">'+d.cgAuditorias+' auditorías</span>':'')+
              '</div>'+
              (d.cgUltima?'<div style="margin-left:auto;font-size:10px;color:var(--muted);font-weight:700">Última: '+fmtCargaFecha(d.cgUltima)+'</div>':'')+
            '</div>'+
            ((isAdminDesp&&_ultCarga&&_ultCarga.audKn===norm(d.nombre))?
              '<button class="btn-ghost" style="font-size:11px;padding:5px 11px;margin-top:10px;border-radius:8px" onclick="revertirCarga('+_ultCarga.id+')" title="Deshace la última importación y restaura el estado anterior">↩ Revertir última carga</button>'
              :'')+
        '</div>'+
      '</div>'):'')+
    '</div>';
  }).join('');

  el.innerHTML='<div style="display:flex;flex-direction:column;gap:14px">'+
    '<div class="card" style="padding:14px 18px">'+
      '<div style="display:flex;align-items:center;gap:10px;flex-wrap:wrap">'+
        '<span style="font-size:13px;font-weight:700;color:var(--nav)">📊 Desempeño de Auditores</span>'+
        '<select id="desp-mes" onchange="renderDesempeno()" style="padding:7px 10px;border-radius:var(--radius-sm);border:1px solid var(--border);font-size:13px;font-family:inherit;background:var(--soft)">'+
          '<option value="ALL">Todos los meses</option>'+
          MORD.map(function(m){return'<option'+(m===mesFilter?' selected':'')+'>'+m+'</option>';}).join('')+
        '</select>'+
        /* Filtro de auditor: se llena solo con los usuarios que tienen rol de
           auditor; los nuevos aparecen automáticamente al crearse. */
        '<select id="desp-auditor" onchange="renderDesempeno()" title="Filtra por auditor. La lista se actualiza sola con los usuarios que tengan rol de auditor." style="padding:7px 10px;border-radius:var(--radius-sm);border:1px solid var(--border);font-size:13px;font-family:inherit;background:var(--soft);max-width:200px">'+
          '<option value="ALL">Todos los auditores</option>'+
          _listaAuditoresOpts.map(function(n){return'<option value="'+esc(n)+'"'+(norm(n)===norm(audSel)?' selected':'')+'>'+esc(n)+'</option>';}).join('')+
        '</select>'+
        '<div style="position:relative;flex:1;min-width:160px;max-width:280px">'+
          '<span style="position:absolute;left:10px;top:50%;transform:translateY(-50%);font-size:12px;color:var(--muted);pointer-events:none">🔍</span>'+
          '<input id="desp-buscar" type="search" placeholder="Buscar auditor…" value="'+esc(buscar)+'" oninput="renderDesempeno()" '+
            'style="width:100%;padding:7px 10px 7px 30px;border-radius:var(--radius-sm);border:1px solid var(--border);font-size:13px;font-family:inherit;background:var(--soft)">'+
        '</div>'+
        '<span style="font-size:11px;color:var(--muted);font-weight:600">'+
          (buscar?auditores.length+' de '+totalAuditores+' auditor(es)':totalAuditores+' auditor(es)')+
        '</span>'+
        '<span style="font-size:10.5px;color:var(--muted);font-weight:600;margin-left:auto" title="Los auditores solo ven datos de su razón social; admin y viewer ven las tres.">'+
          (_razonSelDesp==='ALL'?'Alcance: todas las razones':'Alcance: '+esc(_razonSelDesp))+
        '</span>'+
      '</div>'+
    '</div>'+
    (cards||'<div class="card" style="padding:40px;text-align:center;color:var(--muted)">'+
      (buscar
        ? 'Ningún auditor coincide con “'+esc(buscar)+'”.'
        : 'Sin datos para este período.<br><small>Los ajustes, mermas, actividades y cargas de Excel aparecen aquí cuando se guardan con sesión de auditor activa.</small>')+
    '</div>')+
  '</div>';

  if(_bFoco){var nb=document.getElementById('desp-buscar');if(nb){nb.focus();try{nb.setSelectionRange(_bCaret,_bCaret);}catch(e){}}}
}


/* ════════════════════════════════════════════════════════════════════
   SUPABASE + LOGIN + SYNC
════════════════════════════════════════════════════════════════════ */
/* ── SUPABASE CONFIG (hardcoded) ── */
/* SB_URL / SB_KEY / SB_SESSION_KEY ahora viven en config/supabase-config.js (cargado antes que este archivo) */
/* _sb and _session declared at top */

function initSupabase(){
  try{_sb=supabase.createClient(SB_URL,SB_KEY,{
    auth:{persistSession:false,autoRefreshToken:false,detectSessionInUrl:false},
    realtime:{enabled:false},
    global:{headers:{'x-client-info':'monitor-cumplimiento'}}
  });return _sb;}catch(e){console.warn('Supabase init error:',e);return null;}
}
/* saveSupabaseCfg/toggleSupabaseCfg removed — config hardcoded */
function showLoginErr(msg){document.getElementById('lp-err').textContent=msg;}
function lpTogglePass(){
  var inp=document.getElementById('lp-pass');
  var chk=document.getElementById('lp-show-chk');
  var btn=document.getElementById('lp-eye');
  var show=chk?chk.checked:(inp.type==='text');
  if(!chk){show=inp.type==='password';}
  inp.type=show?'text':'password';
  if(chk)chk.checked=show;
  if(btn)btn.textContent=show?'🙈':'👁️';
}

async function doLogin(){
  const user=document.getElementById('lp-user').value.trim();
  const pass=document.getElementById('lp-pass').value;
  if(!user||!pass){showLoginErr('Completa usuario y contraseña');return;}
  const btn=document.getElementById('lp-btn');
  btn.disabled=true; btn.textContent='Verificando…'; showLoginErr('');

  /* intentar con Supabase */
  if(!_sb) _sb=initSupabase();
  if(_sb){
    try{
      const {data,error}=await _sb.rpc('login_usuario',{p_username:user,p_password:pass});
      if(error) throw error;
      if(data&&data.length>0){
        _session=data[0];
        localStorage.setItem(SB_SESSION_KEY,JSON.stringify(_session));
        /* Derivar clave AES-GCM de la contraseña ANTES de cargar datos */
        await initCryptoKey(pass);
        /* Descifrar campos de sesión que pueden venir cifrados */
        if(_session.nombre)_session.nombre=await dec(_session.nombre);
        if(_session.username)_session.username=await dec(_session.username);
        /* Re-guardar sesión con datos descifrados */
        localStorage.setItem(SB_SESSION_KEY,JSON.stringify(_session));
        onLoginSuccess();return;
      } else {
        showLoginErr('Usuario o contraseña incorrectos');
        btn.disabled=false;btn.textContent='Ingresar';return;
      }
    }catch(e){
      console.warn('Supabase login error:',e.message);
      showLoginErr('Error de conexión: '+e.message);
      btn.disabled=false;btn.textContent='Ingresar';
      return;
    }
  }
  /* Sin Supabase configurado — solo autenticación en la nube */
  showLoginErr('No se pudo conectar con Supabase. Verifica tu conexión a internet.');
  btn.disabled=false;btn.textContent='Ingresar';
}

function setPill(state){
  const pill=document.getElementById('conn-pill');
  const dot=document.getElementById('conn-dot');
  const lbl=document.getElementById('conn-label');
  if(!pill)return;
  pill.classList.remove('offline','sb-ok');
  if(state==='supabase'){
    pill.classList.add('sb-ok');
    lbl.textContent='Conectado a Supabase';
  } else {
    pill.classList.add('offline');
    lbl.textContent='Sin conexión a Supabase';
  }
}
/* Etiqueta legible de cada rol (según definición de Grupo Kuroda) */
function rolLabel(rol){
  return {admin:'Administrador',admin_auditor:'Auditor Señior',
          auditor:'Auditor Jr',viewer:'Viewer · solo lectura'}[rol]||rol;
}
/* ¿El usuario tiene otorgada explícitamente una vista? (vistas_permitidas) */
function tieneVista(v){
  var l=_session&&_session.vistas_permitidas;
  if(typeof l==='string'){try{l=JSON.parse(l);}catch(e){l=null;}}
  return !!(l&&l.includes(v));
}
/* ¿Puede modificar un módulo restringido? admin/señior siempre; Auditor Jr
   solo si se le dio el acceso a ese módulo al crear/editar su usuario. */
function puedeModificarModulo(v){
  if(!_session)return false;
  if(['admin','admin_auditor'].includes(_session.rol))return true;
  if(v==='usuarios')return false; /* exclusivo de admin/admin_auditor, sin excepción */
  return _session.rol==='auditor'&&tieneVista(v);
}
/* ¿Puede el usuario ver el módulo Generador de Dashboard Ejecutivo?
   Recibe un objeto tipo sesión/usuario ({rol, razones_permitidas,
   vistas_permitidas}) para poder reutilizarse tanto con _session como con
   cada fila de la tabla de administración de usuarios.
   Reglas:
   · Admin y Auditor Señior: acceso por defecto.
   · Cuentas con la razón social KNO asignada: acceso por defecto (el
     Generador nació exclusivo para KNO; sus usuarios deben conservar el
     acceso aunque no tengan la vista marcada explícitamente — este era el
     permiso que faltaba y por el que dejaban de ver el módulo).
   · Cualquier otra cuenta (Auditor Jr / Viewer sin KNO asignada): solo con
     acceso explícito en vistas_permitidas ('generador'). */
function puedeAccederGenerador(u){
  if(!u)return false;
  if(['admin','admin_auditor'].includes(u.rol))return true;
  var raz=u.razones_permitidas;
  if(typeof raz==='string'){try{raz=JSON.parse(raz);}catch(e){raz=null;}}
  if(raz&&raz.length&&raz.some(function(x){return razKey(x)==='kno';}))return true;
  var vistas=u.vistas_permitidas;
  if(typeof vistas==='string'){try{vistas=JSON.parse(vistas);}catch(e){vistas=null;}}
  return !!(vistas&&vistas.length&&vistas.includes('generador'));
}

function onLoginSuccess(){
  document.getElementById('login-page').classList.add('hidden');
  const el=document.getElementById('topbar-user');
  if(el&&_session){
    el.innerHTML='<span style="display:block;line-height:1.15">'+esc(_session.nombre||_session.username)+'</span>'+
      '<span style="display:block;font-size:9.5px;font-weight:700;letter-spacing:.06em;text-transform:uppercase;color:var(--blue);line-height:1.1">'+esc(rolLabel(_session.rol))+'</span>';
  }
  checkAdminUI();
  applyRazonRestriction();
  applyVistasRestriction();
  if(!_sb)_sb=initSupabase();
  if(_sb) loadDataFromSupabase();
  else { STORE={auditorias:[],tareas:[]}; refreshAll(); setPill('offline'); toast('⚠ Sin conexión a Supabase'); }
}

function applyRazonRestriction(){
  var fr=document.getElementById('f-razon');
  if(!fr||!_session)return;
  var razones=_razonesAsignadas();
  if(!razones||!razones.length)return;
  if(razones.length===1){
    fr.value=razones[0];
    fr.disabled=true;
    fr.title='Tu cuenta solo tiene acceso a '+razones[0];
  } else {
    fr.disabled=false;
  }
}

function applyVistasRestriction(){
  if(!_session)return;
  var vistas=_session.vistas_permitidas;
  if(typeof vistas==='string'){try{vistas=JSON.parse(vistas);}catch(e){vistas=null;}}
  var lista=(vistas&&vistas.length)?vistas:null;
  var navMap={
    'nav-dash':'dash','nav-tareas':'tareas','nav-pend':'sucursales',
    'nav-actividades':'actividades','nav-auditorias':'auditorias',
    'nav-ajustes':'ajustes','nav-mermas':'mermas','nav-finalizadas':'finalizadas',
    'nav-desempeno':'desempeno','nav-evaluacion':'evaluacion',
    'nav-generador':'generador','nav-documentos':'documentos','nav-usuarios':'usuarios'
  };
  /* Módulos restringidos para Auditor Jr: solo con acceso explícito.
     'usuarios' se quitó de aquí a propósito: la gestión de usuarios es
     exclusiva de admin/admin_auditor y ya no puede habilitarse para
     'auditor' aunque se le marque en sus vistas_permitidas. */
  var RESTRINGIDAS_JR=['finalizadas','desempeno','evaluacion'];
  var visibles=[];
  Object.entries(navMap).forEach(function(e){
    var key=e[1],permitido;
    if(key==='usuarios'){
      /* Usuarios: exclusivo de admin y admin_auditor. Ningún otro rol —
         incluido 'auditor' con acceso explícito por vistas_permitidas —
         puede ver ni abrir este módulo. */
      permitido=['admin','admin_auditor'].includes(_session.rol);
    }else if(key==='generador'){
      /* Generador de reportes ejecutivos (PPTX). Ya NO es exclusivo de KNO
         en cuanto a datos (funciona con cualquier razón social: KNO/KSC/KSA),
         pero las cuentas con razón KNO asignada conservan el acceso
         automático al módulo por ser su origen histórico. Ver
         puedeAccederGenerador() para la regla completa. */
      permitido=puedeAccederGenerador(_session);
    }else if(key==='documentos'){
      /* Herramienta de CREACIÓN de documentos: nunca para viewer (solo lectura).
         Para el resto es de utilidad general y se muestra por defecto (las
         listas de vistas antiguas no incluían 'documentos'). */
      permitido=(_session.rol!=='viewer');
    }else if(_session.rol==='auditor'&&RESTRINGIDAS_JR.includes(key)){
      permitido=!!(lista&&lista.includes(key));
    }else{
      permitido=!lista||lista.includes(key);
    }
    if(permitido&&key!=='usuarios')visibles.push(key);
    var navEl=document.getElementById(e[0]);
    if(navEl)navEl.style.display=permitido?'':'none';
  });
  /* Si la vista activa quedó oculta, ir a la primera visible */
  if(visibles.length&&!visibles.includes(VIEW||'dash')){
    setView(visibles[0]||'dash');
  }
}

function patchedFillFilters(){
  var allT=STORE.tareas, allA=STORE.auditorias;
  /* Si el usuario tiene razones restringidas, filtrar los datos antes de mostrar */
  var razones=_razonesAsignadas();
  if(razones&&razones.length){
    allT=allT.filter(function(t){return razones.some(function(r){return norm(t.razon||'').includes(norm(r));});});
    allA=allA.filter(function(a){return razones.some(function(r){return norm(a.razon||'').includes(norm(r));});});
  }
  var razonesList=uniq([...allT.map(function(t){return t.razon;}),...allA.map(function(a){return a.razon;})]);
  var fr=document.getElementById('f-razon');
  if(!fr)return;
  /* Si solo una razón permitida, fijarla y deshabilitar */
  if(razones&&razones.length===1){
    fr.innerHTML='<option value="'+razones[0]+'">'+razones[0]+'</option>';
    fr.disabled=true;
  } else {
    fr.disabled=false;
    fr.innerHTML='<option value="ALL">Todas</option>'+razonesList.map(function(r){return'<option>'+r+'</option>';}).join('');
  }
  fillCentroTienda();
}

function doLogout(){
  localStorage.removeItem(SB_SESSION_KEY);
  _session=null;
  _cryptoKey=null; /* Destruir clave de cifrado de memoria */
  document.getElementById('login-page').classList.remove('hidden');
  document.getElementById('lp-pass').value='';
  showLoginErr('');
}

/* ── carga datos desde Supabase ── */
async function recargarDatos(){
  var btn=document.getElementById('btn-recargar');
  if(!_sb)_sb=initSupabase();
  if(!_sb){toast('⚠ Sin conexión a Supabase');return;}
  if(btn){btn.disabled=true;btn.textContent='⏳ Recargando…';}
  await loadDataFromSupabase();
  if(btn){btn.disabled=false;btn.textContent='🔄 Recargar';}
}
async function loadDataFromSupabase(){
  if(!_sb){STORE={auditorias:[],tareas:[]};refreshAll();return;}
  try{
    toast('⏳ Cargando datos…');
    const [{data:aud},{data:tar}]=await Promise.all([
      _sb.from('auditorias').select('*').order('fecha',{ascending:false}).limit(5000),
      _sb.from('tareas').select('*').order('fecha_term',{ascending:true}).limit(10000)
    ]);
    /* Descifrar auditorias */
    var audDec=aud?await decArr(aud,FIELDS.auditorias):[];
    STORE.auditorias=audDec.map(function(a){
      function safe(v){return(v&&!pareceCifrado(v))?v:'';}
      return{
        razon:safe(a.razon)||'',centro:canonCentro(safe(a.centro))||'',tienda:canonTienda(safe(a.centro),safe(a.tienda)),
        fecha:a.fecha,mes:safe(a.mes)||'',
        pctCumpl:parseFloat(a.pct_cumpl)||0,
        tareas:a.tareas||0,pendientes:a.pendientes||0,
        resueltas:a.resueltas||0,pctResuelto:parseFloat(a.pct_resuelto)||0,
        clase:safe(a.clase)||''
      };
    });
    /* Descifrar tareas */
    var tarDec=tar?await decArr(tar,FIELDS.tareas):[];
    STORE.tareas=tarDec.map(function(t){
      function safe(v){return(v&&!pareceCifrado(v))?v:'';}
      return{
        id:t.tarea_id||t.id,sb_uuid:t.id,sb_updated:t.updated_at||null,
        razon:safe(t.razon)||'',centro:canonCentro(safe(t.centro))||'',
        tienda:canonTienda(safe(t.centro),safe(t.tienda)),areaResp:safe(t.area_resp)||'',areaRev:safe(t.area_rev)||'',
        actividad:safe(t.actividad)||'',nombre:safe(t.nombre)||'',tipoTarea:safe(t.tipo_tarea)||'',
        estado:safe(t.estado)||'Abierta',
        fechaCreacion:t.fecha_creacion,fechaTerm:t.fecha_term,fechaCumpl:t.fecha_cumpl
      };
    });
    /* Visibilidad por razón social: usuarios restringidos (auditor/viewer con
       razón asignada) solo ven auditorías y tareas de su(s) razón(es). Admin y
       cuentas sin restricción ven todo. Consistente con Actividades/Ajustes. */
    STORE.auditorias=STORE.auditorias.filter(function(a){return razonVisible(a.razon);});
    STORE.tareas=STORE.tareas.filter(function(t){return razonVisible(t.razon);});
    /* Deduplicar auditorías: puede haber más de un renglón para la misma
       auditoría (misma tienda+mes+clase) — por ejemplo si el mismo Excel se
       cargó dos veces con una fecha de captura ligeramente distinta y el
       emparejamiento por Centro+Mes+Clase+Fecha no coincidió con el registro
       ya existente. Sin este paso, "Vigentes" cuenta cada duplicado como una
       auditoría aparte. Se conserva un solo renglón por auditoría (el de
       fecha más reciente). */
    (function dedupStoreAuditorias(){
      var vistos={},out=[];
      STORE.auditorias.slice().sort(function(x,y){return String(y.fecha||'').localeCompare(String(x.fecha||''));})
        .forEach(function(a){
          var k=audKeyMes(a);
          if(vistos[k])return;
          vistos[k]=true;
          out.push(a);
        });
      STORE.auditorias=out;
    })();
    refreshAll();
    await loadFinalizadas();
    if(VIEW==='auditorias')renderAuditoriasView();
    await sincronizarTodasAuditorias();
    setPill('supabase');toast('✓ Conectado a Supabase — '+STORE.tareas.length+' tareas, '+STORE.auditorias.length+' auditorías');
  }catch(e){
    console.warn('Error cargando Supabase:',e);
    STORE={auditorias:[],tareas:[]};refreshAll();
    setPill('offline');toast('⚠ Error cargando datos de Supabase: '+e.message);
  }
}

/* ── sincronizar Excel subido → Supabase ── */
async function syncToSupabase(parsed){
  var client=_sb;
  if(!client){
    try{client=supabase.createClient(SB_URL,SB_KEY,{
    auth:{persistSession:false,autoRefreshToken:false,detectSessionInUrl:false},
    realtime:{enabled:false},
    global:{headers:{'x-client-info':'monitor-cumplimiento'}}
  });}catch(e){return;}
  }
  var upload_id='upload_'+Date.now(), audOk=0, tarOk=0, errMsg=null;
  try{
    if(parsed.auditorias&&parsed.auditorias.length>0){
      var arowsRaw=parsed.auditorias.filter(Boolean).map(function(a){return{
        razon:a.razon||null,centro:a.centro||null,tienda:a.tienda||null,
        fecha:a.fecha?String(a.fecha).split('T')[0]:null,
        mes:a.mes||null,pct_cumpl:parseFloat(a.pctCumpl)||0,
        tareas:parseInt(a.tareas)||0,pendientes:parseInt(a.pendientes)||0,
        resueltas:parseInt(a.resueltas)||0,
        pct_resuelto:parseFloat(a.pctResuelto)||0,
        clase:a.clase||null,upload_id:upload_id
      };});
      var arows=[];for(var _ai=0;_ai<arowsRaw.length;_ai++)arows.push(await encObj(arowsRaw[_ai],FIELDS.auditorias));
      if(arows.length){
        var ar=await client.from('auditorias').upsert(arows,{onConflict:'razon,centro,tienda,fecha',ignoreDuplicates:false});
        if(ar.error)errMsg='Aud: '+ar.error.message;
        else audOk=arows.length;
      }
    }
    if(parsed.tareas&&parsed.tareas.length>0){
      var trowsRaw=parsed.tareas.filter(Boolean).map(function(t){return{
        tarea_id:String(t.id||t.tarea_id||''),
        tarea_key:tareaKey(String(t.id||t.tarea_id||''),t.razon),
        razon:t.razon||null,centro:t.centro||null,tienda:t.tienda||null,
        area_resp:t.areaResp||null,area_rev:t.areaRev||null,
        actividad:t.actividad||null,nombre:t.nombre||null,
        tipo_tarea:t.tipoTarea||null,estado:t.estado||null,
        fecha_creacion:t.fechaCreacion?String(t.fechaCreacion).split('T')[0]:null,
        fecha_term:t.fechaTerm?String(t.fechaTerm).split('T')[0]:null,
        fecha_cumpl:t.fechaCumpl?String(t.fechaCumpl).split('T')[0]:null,
        upload_id:upload_id
      };});
      var trows=[];for(var _ti=0;_ti<trowsRaw.length;_ti++)trows.push(await encObj(trowsRaw[_ti],FIELDS.tareas));
      for(var ci=0;ci<trows.length;ci+=500){
        var tr=await client.from('tareas').upsert(trows.slice(ci,ci+500),{onConflict:'tarea_key',ignoreDuplicates:false});
        if(tr.error){errMsg=(errMsg?errMsg+' | ':'')+'Tar: '+tr.error.message;break;}
        tarOk+=Math.min(500,trows.length-ci);
      }
    }
    if(errMsg)toast('⚠ '+errMsg);
    else if(audOk>0||tarOk>0){
      toast('☁️ Guardado en Supabase: '+audOk+' auditorías · '+tarOk+' tareas');
      if(typeof setPill==='function')setPill('supabase');
    }
  }catch(e){toast('⚠ Error Supabase: '+e.message);console.error('syncToSupabase:',e);}
}

/* ── inicialización ── */
(function initApp(){
  initSupabase();
  try{
    const raw=localStorage.getItem(SB_SESSION_KEY);
    if(raw&&_sb){
      /* Se pide la contraseña para derivar TAMBIÉN la clave antigua y poder
         leer los datos que aún no se han migrado al esquema compartido.
         (La clave compartida no depende de la contraseña; la legacy sí.) */
      var sess=JSON.parse(raw);
      var uInput=document.getElementById('lp-user');
      if(uInput&&sess.username)uInput.value=sess.username;
      var hint=document.getElementById('login-hint');
      if(hint)hint.textContent='Sesión guardada — ingresa tu contraseña para continuar';
    }
  }catch(e){}
})();



/* ════════════════════════════════════════════════════════════════════
   GESTIÓN DE USUARIOS (solo admin)
════════════════════════════════════════════════════════════════════ */
function openUsuarios(){
  if(!puedeModificarModulo('usuarios')){toast('⚠ Sin permisos');return;}
  document.getElementById('usr-overlay').classList.add('show');
  loadUsuarios();
  mostrarBotonMigracion();
  restringirCheckboxRazonesUI('nu');
}
/* Quien crea o edita un usuario solo puede otorgarle sus MISMAS razones
   sociales (nunca "Todas" ni una razón ajena a la suya). Antes esto solo se
   validaba al guardar (con un mensaje de error); ahora además se oculta/
   deshabilita en la propia UI para que ni siquiera se pueda intentar marcar
   una razón que no le pertenece. Se usa tanto en "Nuevo usuario" (prefix
   'nu') como en "Editar usuario" (prefix 'eu'). */
function restringirCheckboxRazonesUI(prefix){
  var propias=_razonesAsignadas();
  var cbs=document.querySelectorAll('.'+prefix+'-razon-cb');
  var allCb=document.getElementById(prefix+'-razon-all');
  var allWrap=allCb?allCb.closest('label'):null;
  if(!propias){
    /* admin sin restricción de razón: todas las opciones disponibles */
    cbs.forEach(function(cb){cb.disabled=false;var l=cb.closest('label');if(l)l.style.display='';});
    if(allCb){allCb.disabled=false;if(allWrap)allWrap.style.display='';}
    return;
  }
  var propiasNorm=propias.map(razKey);
  if(allCb){allCb.checked=false;allCb.disabled=true;if(allWrap)allWrap.style.display='none';}
  cbs.forEach(function(cb){
    var permitido=propiasNorm.includes(razKey(cb.value));
    var l=cb.closest('label');
    if(l)l.style.display=permitido?'':'none';
    if(!permitido){cb.checked=false;cb.disabled=true;}
    else if(propiasNorm.length===1){cb.checked=true;cb.disabled=true;} /* única razón propia: queda fija */
    else{cb.disabled=false;}
  });
}
/* Panel de migración de cifrado — visible para cualquier admin. Si en este
   navegador no está la clave antigua (p. ej. la sesión se restauró sin pedir
   contraseña), se solicita la contraseña al momento de migrar. */
function mostrarBotonMigracion(){
  /* Panel de migración de cifrado oculto en la UI de Gestión de Usuarios.
     Se deja la función (no-op de visibilidad) para no romper las llamadas
     existentes ni las funciones migrarCifrado/diagnosticarCifrado/limpiarIlegibles. */
  try{
    var w=document.getElementById('mig-wrap');
    if(w)w.style.display='none';
  }catch(e){}
}
/* Reconstruye la clave antigua a partir de la contraseña del admin y el salt
   guardado en ESTE navegador. Necesaria para leer los datos aún sin migrar. */
async function asegurarLegacyKey(){
  if(_legacyKey)return true;
  var salt=getCryptoSalt();
  if(!salt)return false; /* este navegador nunca cifró: no hay salt antiguo */
  var pass=prompt('Para migrar los datos antiguos, escribe la contraseña del administrador (la misma con la que se cargaron los datos):');
  if(!pass)return false;
  try{
    _legacyKey=await _derivar(pass,salt);
    return true;
  }catch(e){ _legacyKey=null; return false; }
}
function closeUsuarios(){
  document.getElementById('usr-overlay').classList.remove('show');
  document.getElementById('usr-err').textContent='';
}

/* Alcance por razón social del módulo Usuarios: una cuenta admin_auditor
   con razones_permitidas asignadas (p.ej. KNO) solo debe ver y poder editar
   usuarios de esa(s) misma(s) razón(es) — no la lista completa del sistema.
   El admin sin restricción de razón (razones_permitidas vacío/null) sigue
   viendo a todos, igual que en el resto del dashboard (_razonesAsignadas()/
   razonVisible()). Un usuario objetivo sin razón asignada (razones_permitidas
   vacío = "todas") se considera fuera del alcance de una cuenta restringida,
   para que un admin_auditor de KNO no pueda ver ni tocar cuentas de alcance
   global. */
function filtrarUsuariosPorRazonSesion(data){
  var propias=_razonesAsignadas();
  if(!propias)return data; /* sesión sin restricción de razón: ve todo */
  var propiasNorm=propias.map(razKey);
  return data.filter(function(u){
    var raz=u.razones_permitidas;
    if(typeof raz==='string'){try{raz=JSON.parse(raz);}catch(e){raz=null;}}
    if(!raz||!raz.length)return false;
    return raz.some(function(r){return propiasNorm.includes(razKey(r));});
  });
}

async function loadUsuarios(){
  const tbody=document.getElementById('usr-tbody');
  tbody.innerHTML='<tr><td colspan="8" style="text-align:center;color:var(--muted);padding:20px">Cargando...</td></tr>';
  if(!_sb){tbody.innerHTML='<tr><td colspan="8" style="text-align:center;color:var(--muted)">Sin conexión a Supabase</td></tr>';return;}
  try{
    const {data,error}=await _sb.rpc('listar_usuarios');
    if(error)throw error;
    var visibles=filtrarUsuariosPorRazonSesion(data||[]);
    if(!visibles.length){tbody.innerHTML='<tr><td colspan="8" style="text-align:center;color:var(--muted)">Sin usuarios</td></tr>';return;}
    /* Actualizar lista de auditores medibles (con el universo completo de
       usuarios, no solo los visibles para esta sesión: un admin_auditor de
       KNO igual debe medir auditores de otras razones en Desempeño). */
    window._usrEditMap={};
    _AUDITORES_CONOCIDOS=(data||[]).filter(function(u){return u.es_auditor&&u.activo;}).map(function(u){return u.nombre||u.username;});
    tbody.innerHTML=visibles.map(function(u){
      var rolClass='rol-'+u.rol;
      var razones=u.razones_permitidas;
      if(typeof razones==='string'){try{razones=JSON.parse(razones);}catch(e){razones=null;}}
      var canEdit=_session&&['admin','admin_auditor'].includes(_session.rol);
      var razonesLabel=razones&&razones.length?razones.join(', '):'Todas';
      var esAuditorLabel=u.es_auditor?'<span style="color:var(--green);font-weight:700">✓</span>':'<span style="color:var(--muted)">—</span>';
      /* Verificación de permisos: acceso calculado al Generador según el rol
         y las razones de esta cuenta (misma regla que applyVistasRestriction,
         vía puedeAccederGenerador), para poder auditar de un vistazo que cada
         usuario tenga el acceso que le corresponde por su rol. */
      var tieneGen=puedeAccederGenerador({rol:u.rol,razones_permitidas:razones,vistas_permitidas:u.vistas_permitidas});
      var genLabel=tieneGen?'<span style="color:var(--green);font-weight:700">✓</span>':'<span style="color:var(--muted)">—</span>';
      /* Guardar datos en mapa global indexado por id */
      window._usrEditMap[u.id]={id:u.id,username:u.username,nombre:u.nombre||'',rol:u.rol,es_auditor:!!u.es_auditor,razones_permitidas:razones||[],vistas_permitidas:u.vistas_permitidas?JSON.parse(u.vistas_permitidas):[]  };
      return '<tr>'+
        '<td style="font-weight:700;font-family:monospace">'+u.username+'</td>'+
        '<td>'+(u.nombre||'—')+'</td>'+
        '<td><span class="rol-badge '+rolClass+'">'+u.rol+'</span></td>'+
        '<td style="font-size:11px;color:var(--muted);font-weight:600">'+razonesLabel+'</td>'+
        '<td style="text-align:center">'+esAuditorLabel+'</td>'+
        '<td style="text-align:center">'+genLabel+'</td>'+
        '<td><span style="color:'+(u.activo?'var(--green)':'var(--red)')+'">'+
          (u.activo?'✓ Activo':'✕ Inactivo')+'</span></td>'+
        '<td style="display:flex;gap:6px;flex-wrap:wrap">'+
          (canEdit?'<button class="btn btn-ghost" style="padding:4px 8px;font-size:11px;color:var(--blue)" onclick="openEditUsuario(\''+u.id+'\')">Editar</button>':'')+
          '<button class="btn btn-ghost" style="padding:4px 8px;font-size:11px" onclick="toggleActivo(\''+u.id+'\','+(!u.activo)+')">'+
            (u.activo?'Desactivar':'Activar')+'</button>'+
          (u.username!=='ADMIN'&&canEdit?'<button class="btn btn-ghost" style="padding:4px 8px;font-size:11px;color:var(--red)" onclick="eliminarUsuario(\''+u.id+'\',\''+u.username+'\')">Eliminar</button>':'')+
        '</td></tr>';
    }).join('');
  }catch(e){
    tbody.innerHTML='<tr><td colspan="8" style="color:var(--red);text-align:center">Error: '+e.message+'</td></tr>';
  }
}

function onRolChange(){
  /* Mostrar/ocultar opciones de auditor según el rol */
  var rol=document.getElementById('nu-rol').value;
  var opts=document.getElementById('nu-auditor-opts');
  if(opts)opts.style.display=['auditor','admin_auditor','viewer','editor'].includes(rol)?'flex':'none';
}
function onRazonAllChange(cb){
  /* Si marca "Todas", desmarcar el resto */
  if(cb.checked){
    document.querySelectorAll('.nu-razon-cb').forEach(function(el){el.checked=false;});
  }
}

async function crearUsuario(){
  var user=document.getElementById('nu-user').value.trim().toUpperCase();
  var pass=document.getElementById('nu-pass').value;
  var nombre=document.getElementById('nu-nombre').value.trim();
  var rol=document.getElementById('nu-rol').value;
  var esAuditor=document.getElementById('nu-es-auditor').checked;
  var razonAll=document.getElementById('nu-razon-all').checked;
  var razonesSel=[...document.querySelectorAll('.nu-razon-cb:checked')].map(function(el){return el.value;});
  var razonesPermitidas=razonAll||!razonesSel.length?null:razonesSel;
  var errEl=document.getElementById('usr-err');
  errEl.textContent='';
  if(!user||!pass){errEl.textContent='Usuario y contraseña requeridos';return;}
  if(pass.length<3){errEl.textContent='Contraseña mínimo 3 caracteres';return;}
  if(!_sb){errEl.textContent='Sin conexión a Supabase';return;}
  /* Un admin_auditor con razón(es) asignada(s) solo puede crear usuarios
     dentro de su propio alcance: ni 'Todas' ni una razón ajena. */
  var propias=_razonesAsignadas();
  if(propias){
    var propiasNorm=propias.map(razKey);
    if(!razonesPermitidas||!razonesPermitidas.length||razonesPermitidas.some(function(r){return !propiasNorm.includes(razKey(r));})){
      errEl.textContent='Solo puedes crear usuarios dentro de tu(s) razón(es): '+propias.join(', ');
      return;
    }
  }
  try{
    const {data,error}=await _sb.rpc('crear_usuario',{
      p_username:user, p_password:pass, p_rol:rol,
      p_nombre:nombre||null,
      p_es_auditor:esAuditor,
      p_razones_permitidas:razonesPermitidas?JSON.stringify(razonesPermitidas):null
    });
    if(error)throw error;
    document.getElementById('nu-user').value='';
    document.getElementById('nu-pass').value='';
    document.getElementById('nu-nombre').value='';
    document.getElementById('nu-es-auditor').checked=false;
    document.getElementById('nu-razon-all').checked=false;
    document.querySelectorAll('.nu-razon-cb').forEach(function(el){el.checked=false;});
    toast('✓ Usuario '+user+' creado');
    loadUsuarios();
  }catch(e){
    errEl.textContent='Error: '+e.message;
  }
}

async function toggleActivo(id, nuevoEstado){
  if(!_sb)return;
  try{
    await _sb.rpc('actualizar_usuario',{p_id:id, p_activo:nuevoEstado});
    loadUsuarios();
    toast('✓ Usuario '+(nuevoEstado?'activado':'desactivado'));
  }catch(e){toast('Error: '+e.message);}
}

function openEditUsuario(uid){
  if(!_session||!['admin','admin_auditor'].includes(_session.rol)){toast('⚠ Sin permisos');return;}
  var u=(window._usrEditMap||{})[uid];
  if(!u){toast('⚠ No se encontraron datos del usuario');return;}
  var ROLES=['auditor','admin_auditor','viewer','admin'];
  document.getElementById('eu-overlay').classList.add('show');
  document.getElementById('eu-titulo').textContent='Editar — '+u.username;
  document.getElementById('eu-id').value=u.id;
  document.getElementById('eu-nombre').value=u.nombre||'';
  var rolSel=document.getElementById('eu-rol');
  rolSel.innerHTML=ROLES.map(function(r){return'<option value="'+r+'"'+(r===u.rol?' selected':'')+'>'+r+'</option>';}).join('');
  document.getElementById('eu-es-auditor').checked=!!u.es_auditor;
  var raz=u.razones_permitidas||[];
  document.querySelectorAll('.eu-razon-cb').forEach(function(cb){cb.checked=raz.includes(cb.value);});
  document.getElementById('eu-razon-all').checked=!raz.length;
  restringirCheckboxRazonesUI('eu');
  /* Cargar vistas permitidas */
  var vistas=u.vistas_permitidas||[];
  document.querySelectorAll('.eu-vista-cb').forEach(function(cb){
    cb.checked=!vistas.length||vistas.includes(cb.value);
  });
  document.getElementById('eu-pass').value='';
  document.getElementById('eu-err').textContent='';
}
function closeEditUsuario(){
  document.getElementById('eu-overlay').classList.remove('show');
}
function onEuRazonAllChange(cb){
  if(cb.checked)document.querySelectorAll('.eu-razon-cb').forEach(function(el){el.checked=false;});
}

async function guardarEditUsuario(){
  var id=document.getElementById('eu-id').value;
  var nombre=document.getElementById('eu-nombre').value.trim();
  var rol=document.getElementById('eu-rol').value;
  var esAuditor=document.getElementById('eu-es-auditor').checked;
  var razonAll=document.getElementById('eu-razon-all').checked;
  var razonesSel=[...document.querySelectorAll('.eu-razon-cb:checked')].map(function(el){return el.value;});
  var razonesPermitidas=razonAll||!razonesSel.length?null:razonesSel;
  /* Vistas: si están todas marcadas = null (sin restricción), si hay desmarcadas = guardar solo las marcadas */
  var todasVistas=['dash','tareas','sucursales','actividades','auditorias','ajustes','mermas','finalizadas','documentos'];
  var vistasSel=[...document.querySelectorAll('.eu-vista-cb:checked')].map(function(el){return el.value;});
  var vistasPermitidas=vistasSel.length===todasVistas.length?null:vistasSel;
  var newPass=document.getElementById('eu-pass').value;
  var errEl=document.getElementById('eu-err');
  errEl.textContent='';
  if(!_sb){errEl.textContent='Sin conexión a Supabase';return;}
  /* Mismo alcance por razón que en crearUsuario: un admin_auditor
     restringido no puede reasignar un usuario a 'Todas' ni a una razón
     fuera de la(s) suya(s) — y openEditUsuario() ya evita que abra la
     edición de un usuario que estuviera fuera de su alcance. */
  var propiasEU=_razonesAsignadas();
  if(propiasEU){
    var propiasEUNorm=propiasEU.map(razKey);
    if(!razonesPermitidas||!razonesPermitidas.length||razonesPermitidas.some(function(r){return !propiasEUNorm.includes(razKey(r));})){
      errEl.textContent='Solo puedes asignar tu(s) propia(s) razón(es): '+propiasEU.join(', ');
      return;
    }
  }
  try{
    var {error}=await _sb.rpc('actualizar_usuario',{
      p_id:id,
      p_nombre:nombre||null,
      p_rol:rol,
      p_es_auditor:esAuditor,
      p_razones_permitidas:razonesPermitidas?JSON.stringify(razonesPermitidas):null,
      p_vistas_permitidas:vistasPermitidas?JSON.stringify(vistasPermitidas):null,
      p_password:newPass||null
    });
    if(error)throw error;
    closeEditUsuario();
    toast('✓ Usuario actualizado');
    loadUsuarios();
  }catch(e){errEl.textContent='Error: '+e.message;}
}

async function eliminarUsuario(id, username){
  if(!confirm('¿Eliminar usuario '+username+'?'))return;
  if(!_sb)return;
  try{
    await _sb.rpc('eliminar_usuario',{p_id:id});
    loadUsuarios();
    toast('✓ Usuario '+username+' eliminado');
  }catch(e){toast('Error: '+e.message);}
}

/* Mostrar botón admin solo si el usuario tiene rol admin */
function checkAdminUI(){
  /* nav-usuarios y nav-desempeno se gobiernan en applyVistasRestriction
     (Auditor Jr puede tenerlos SI se le dio el acceso). */
  applyVistasRestriction();
  var excelBtn=document.getElementById('btn-cargar-excel');
  if(excelBtn){
    var canUpload=_session&&['admin','admin_auditor','auditor'].includes(_session.rol);
    excelBtn.style.opacity=canUpload?'1':'0.4';
    excelBtn.title=canUpload?'Cargar Excel':'Solo admin, admin_auditor y auditor pueden cargar archivos';
  }
  var addBtn=document.getElementById('btn-agregar-tarea');
  if(addBtn){
    var canEdit=_session&&['admin','admin_auditor','auditor'].includes(_session.rol);
    addBtn.style.display=canEdit?'':'none';
  }
  var addAct=document.getElementById('btn-add-actividad');
  if(addAct){
    var canEditA=_session&&['admin','admin_auditor','auditor'].includes(_session.rol);
    addAct.style.display=canEditA?'':'none';
  }
  var actExcel=document.getElementById('btn-excel-actividad');
  if(actExcel) actExcel.style.display=(_session&&['admin','admin_auditor','auditor'].includes(_session.rol))?'':'none';
  var audBtn=document.getElementById('btn-nueva-auditoria');
  if(audBtn) audBtn.style.display=(_session&&['admin','admin_auditor','auditor'].includes(_session.rol))?'':'none';
  var ajBtn=document.getElementById('btn-nuevo-ajuste');
  if(ajBtn) ajBtn.style.display=(_session&&['admin','admin_auditor','auditor'].includes(_session.rol))?'':'none';
  var ajExcel=document.getElementById('btn-excel-ajuste');
  if(ajExcel) ajExcel.style.display=(_session&&['admin','admin_auditor','auditor'].includes(_session.rol))?'':'none';
  var ajTh=document.getElementById('aj-edit-th');
  if(ajTh) ajTh.style.display=(_session&&['admin','admin_auditor','auditor'].includes(_session.rol))?'':'none';
  var mrBtn=document.getElementById('btn-nueva-merma');
  if(mrBtn) mrBtn.style.display=(_session&&['admin','admin_auditor','auditor'].includes(_session.rol))?'':'none';
  var mrExcel=document.getElementById('btn-excel-merma');
  if(mrExcel) mrExcel.style.display=(_session&&['admin','admin_auditor','auditor'].includes(_session.rol))?'':'none';
  var mrTh=document.getElementById('mr-edit-th');
  if(mrTh) mrTh.style.display=(_session&&['admin','admin_auditor','auditor'].includes(_session.rol))?'':'none';
  var finTh=document.getElementById('fin-edit-th');
  if(finTh) finTh.style.display=(_session&&['admin','admin_auditor'].includes(_session.rol))?'':'none';
}

/* ── Inicializar semanas si no existen (crea 5 por defecto) ── *//* ════════════════════════════════════════════════════════════════════
   MÓDULO AJUSTES — Supabase
   Tabla: ajustes
   Columnas: id (uuid PK), tienda (text), mes (text), año (int),
             fecha_correo (date), fecha_ajuste (date),
             dias (int generado), condicion (text generado),
             created_at (timestamptz)
════════════════════════════════════════════════════════════════════ */
/* Año mínimo que se muestra en Ajustes y Mermas (registros previos se descartan) */
var AÑO_MIN_MODULOS=2026;
var AJUSTES=[];
var _ajPerManual=false; /* true cuando el usuario elige mes/año a mano: la carga no lo reescribe */
var _chartAjDonut=null;

function diasEntre(d1,d2){
  if(!d1||!d2)return null;
  var a=new Date(d1+'T00:00:00'),b=new Date(d2+'T00:00:00');
  if(b<a)return 0;
  /* Contar solo días hábiles lunes-viernes */
  var habil=0,cur=new Date(a);
  while(cur<b){
    cur.setDate(cur.getDate()+1);
    var dw=cur.getDay();
    if(dw!==0&&dw!==6)habil++;
  }
  return habil;
}
function condicionAjuste(dias){
  if(dias===null)return'—';
  return dias<=3?'A tiempo':'Destiempo';
}
function mesNombre(m){
  var n=['Enero','Febrero','Marzo','Abril','Mayo','Junio',
         'Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'];
  return n[parseInt(m)-1]||m;
}

async function loadAjustes(){
  var client=getSbClient();
  if(!client){toast('⚠ Sin Supabase');return;}
  try{
    var r=await client.from('ajustes').select('*').order('fecha_correo',{ascending:true}).order('fecha_ajuste',{ascending:true}).limit(5000);
    if(r.error){toast('⚠ '+r.error.message);return;}
    var MN=['ENERO','FEBRERO','MARZO','ABRIL','MAYO','JUNIO',
      'JULIO','AGOSTO','SEPTIEMBRE','OCTUBRE','NOVIEMBRE','DICIEMBRE'];
    var _ajDataRaw=r.data||[];
    var _ajDataDec=await decArr(_ajDataRaw,FIELDS.ajustes);
    AJUSTES=_ajDataDec.map(function(a){
      var dias=diasEntre(a.fecha_correo,a.fecha_ajuste);
      /* SIEMPRE derivar mes y año desde la FECHA DE AJUSTE (predominante). Si aún
         no hay fecha de ajuste (pendiente de validar), se usa fecha_correo como
         respaldo para que el registro no desaparezca de los filtros. Se evita así
         además depender de formatos sucios ("julio2023.", "agosto", etc.) que
         pueden estar en el campo mes de Supabase. */
      var mesLimpio='';
      var añoLimpio=null;
      var _fechaBaseAj=a.fecha_ajuste||a.fecha_correo;
      if(_fechaBaseAj){
        var parts=_fechaBaseAj.split('-');
        if(parts.length>=2){
          añoLimpio=parseInt(parts[0]);
          mesLimpio=MN[parseInt(parts[1])-1]||'';
        }
      }
      /* Fallback: intentar parsear el campo mes almacenado si no hay ninguna fecha */
      if(!mesLimpio&&a.mes){
        var raw=String(a.mes).toUpperCase().replace(/[0-9.]/g,'').trim();
        /* Extraer solo la parte del nombre del mes */
        var found=MN.find(function(m){return raw.startsWith(m)||raw===m;});
        mesLimpio=found||raw;
      }
      if(!añoLimpio&&a.año)añoLimpio=parseInt(a.año);
      return{id:a.id,tienda:a.tienda||'',mes:mesLimpio,año:añoLimpio,
        fechaCorreo:a.fecha_correo,fechaAjuste:a.fecha_ajuste,
        dias:dias,condicion:condicionAjuste(dias),auditor:a.auditor||'',
        razon:razonDeCentro(centroDeTienda(a.tienda||''))};
    });
    AJUSTES=AJUSTES.filter(function(a){return tiendaVisible(a.tienda);});
    /* Solo se conservan registros de 2026 en adelante (los años anteriores no se necesitan) */
    AJUSTES=AJUSTES.filter(function(a){return (parseInt(a.año)||0)>=AÑO_MIN_MODULOS;});
    fillAjFilters();renderAjustes();
    toast('✓ '+AJUSTES.length+' ajuste(s) cargados');
  }catch(e){toast('⚠ '+e.message);}
}

function fillAjFilters(){
  var MORD=['ENERO','FEBRERO','MARZO','ABRIL','MAYO','JUNIO',
    'JULIO','AGOSTO','SEPTIEMBRE','OCTUBRE','NOVIEMBRE','DICIEMBRE'];
  /* Meses en orden cronológico, no alfabético */
  var mesSet=new Set(AJUSTES.map(function(a){return(a.mes||'').toUpperCase();}).filter(Boolean));
  var meses=MORD.filter(function(m){return mesSet.has(m);});
  /* Completar con cualquier mes no estándar que pueda venir */
  mesSet.forEach(function(m){if(MORD.indexOf(m)<0)meses.push(m);});
  var años=[...new Set(AJUSTES.map(function(a){return a.año?String(a.año):'';}).filter(Boolean))].sort().reverse();
  var tiendas=[...new Set(AJUSTES.map(function(a){return a.tienda;}).filter(Boolean))].sort();
  var sm=document.getElementById('aj-f-mes');
  var sa=document.getElementById('aj-f-año');
  var st=document.getElementById('aj-f-tienda');
  if(!sm)return;
  var vm=sm.value,va=sa.value,vt=st.value;
  sm.innerHTML='<option value="ALL">Todos</option>'+meses.map(function(m){return'<option value="'+m+'">'+m+'</option>';}).join('');
  sa.innerHTML='<option value="ALL">Todos</option>'+años.map(function(y){return'<option value="'+y+'">'+y+'</option>';}).join('');
  st.innerHTML='<option value="ALL">Todas</option>'+limpiarOpciones(tiendas).map(function(t){return'<option>'+t+'</option>';}).join('');
  /* Período objetivo: mes/año en curso si hay datos; si no, el más reciente con datos.
     Nunca "Todos" (para no volcar todo el histórico). */
  var mesHoy=MORD[new Date().getMonth()];
  var añoHoy=String(new Date().getFullYear());
  function _tienePer(m,y){return AJUSTES.some(function(a){return (a.mes||'').toUpperCase()===m&&String(a.año||'')===String(y);});}
  var tgtMes=mesHoy, tgtAño=añoHoy;
  if(!_tienePer(mesHoy,añoHoy)){
    var best=null;
    AJUSTES.forEach(function(a){
      var mi=MORD.indexOf((a.mes||'').toUpperCase()); var y=parseInt(a.año)||0;
      if(mi<0||!y)return;
      var key=y*100+mi;
      if(!best||key>best.key)best={key:key,mes:(a.mes||'').toUpperCase(),año:String(a.año)};
    });
    if(best){tgtMes=best.mes;tgtAño=best.año;}
  }
  /* Mes/año por defecto = período en curso si tiene datos; si no, el más reciente
     con datos. Se recalcula en cada carga salvo que el usuario haya elegido a mano
     (_ajPerManual), para que el mes en curso aparezca solo (incl. tras importar). */
  if(_ajPerManual){
    if(vm!=='ALL'&&[...sm.options].some(function(o){return o.value===vm;}))sm.value=vm;
    if(va!=='ALL'&&[...sa.options].some(function(o){return o.value===va;}))sa.value=va;
  }else{
    sm.value=[...sm.options].some(function(o){return o.value===tgtMes;})?tgtMes:'ALL';
    sa.value=[...sa.options].some(function(o){return o.value===tgtAño;})?tgtAño:'ALL';
  }
  if(vt!=='ALL'&&[...st.options].some(function(o){return o.value===vt;}))st.value=vt;
}
function ajPerChanged(){_ajPerManual=true;renderAjustes();}
function filteredAjustes(){
  var mes=(document.getElementById('aj-f-mes').value||'ALL');
  var año=(document.getElementById('aj-f-año').value||'ALL');
  var tienda=(document.getElementById('aj-f-tienda').value||'ALL');
  return AJUSTES.filter(function(a){
    if(mes!=='ALL'&&(a.mes||'').toUpperCase()!==mes.toUpperCase())return false;
    if(año!=='ALL'&&String(a.año||'')!==String(año))return false;
    if(tienda!=='ALL'&&norm(a.tienda)!==norm(tienda))return false;
    return true;
  });
}

function fmtFecha(iso){
  if(!iso)return'—';var p=iso.split('-');return p[2]+'/'+p[1]+'/'+p[0];
}

function renderAjustes(){
  var arr=filteredAjustes();
  /* Ordenar por fecha de correo: día 1 primero, último día al final */
  arr=arr.slice().sort(function(a,b){
    var da=a.fechaCorreo||'';var db=b.fechaCorreo||'';
    return da<db?-1:da>db?1:0;
  });
  var isAdmin=_session&&['admin','admin_auditor','auditor'].includes(_session.rol);
  document.getElementById('aj-count').textContent=arr.length+' registro(s)';
  var th=document.getElementById('aj-edit-th');
  if(th)th.style.display=isAdmin?'':'none';

  var tbody=document.getElementById('aj-tbody');
  if(!arr.length){
    tbody.innerHTML='<tr><td colspan="8" style="text-align:center;color:var(--muted);padding:24px">Sin registros para los filtros seleccionados.</td></tr>';
  }else{
    tbody.innerHTML=arr.map(function(a,i){
      var cls=a.condicion==='A tiempo'?'b-atiempo':'b-destiempo';
      var editTd=isAdmin?'<td style="text-align:center"><button class="icon-btn" onclick="openEditAjuste(\''+a.id+'\')" title="Editar">✎</button></td>':'';
      return '<tr>'+
        '<td style="color:var(--muted);font-size:11px">'+(i+1)+'</td>'+
        '<td><b>'+esc(a.tienda)+'</b></td>'+
        '<td>'+esc(a.mes)+'</td>'+
        '<td style="text-align:center">'+fmtFecha(a.fechaCorreo)+'</td>'+
        '<td style="text-align:center">'+fmtFecha(a.fechaAjuste)+'</td>'+
        '<td style="text-align:center;font-weight:800">'+(a.dias!==null?a.dias:'—')+'</td>'+
        '<td style="text-align:center"><span class="'+cls+'">'+esc(a.condicion)+'</span></td>'+
        editTd+'</tr>';
    }).join('');
  }
  renderAjDonut(arr);
}

function renderAjDonut(arr){
  var ok=arr.filter(function(a){return a.condicion==='A tiempo';}).length;
  var mal=arr.filter(function(a){return a.condicion==='Destiempo';}).length;
  var total=arr.length;
  document.getElementById('aj-donut-note').textContent=total?total+' registros':'';
  if(_chartAjDonut){_chartAjDonut.destroy();_chartAjDonut=null;}
  /* Siempre restaurar el canvas antes de decidir qué mostrar */
  var wrap=document.getElementById('chart-aj-donut');
  if(!wrap||wrap.tagName!=='CANVAS'){
    var parent=wrap?wrap.parentElement:document.querySelector('#view-ajustes .ajustes-donut-wrap div');
    if(parent)parent.innerHTML='<canvas id="chart-aj-donut"></canvas>';
  }
  var legend=document.getElementById('aj-donut-legend');
  if(!total){
    var cv=document.getElementById('chart-aj-donut');
    if(cv)cv.parentElement.innerHTML='<div class="empty" style="display:flex;align-items:center;justify-content:center;width:165px;height:165px;color:var(--muted);font-size:12px">Sin datos</div>';
    if(legend)legend.innerHTML='';
    return;
  }
  var canvas=document.getElementById('chart-aj-donut');
  if(!canvas)return;
  _chartAjDonut=new Chart(canvas,{
    type:'doughnut',
    data:{labels:['A tiempo','Destiempo'],
      datasets:[{data:[ok,mal],backgroundColor:['#16a34a','#dc2626'],borderWidth:2,borderColor:'#fff'}]},
    options:{responsive:true,maintainAspectRatio:false,cutout:'65%',plugins:{legend:{display:false},
      tooltip:{callbacks:{label:function(ctx2){return ctx2.label+': '+ctx2.raw+' ('+Math.round(ctx2.raw/total*100)+'%)';}}}}}
  });
  if(legend)legend.innerHTML=
    '<div class="legend-item"><span class="legend-dot" style="background:#16a34a"></span>'+
    '<span class="legend-lbl">A tiempo</span><span class="legend-val" style="margin-left:12px">'+ok+'</span></div>'+
    '<div class="legend-item"><span class="legend-dot" style="background:#dc2626"></span>'+
    '<span class="legend-lbl">Destiempo</span><span class="legend-val" style="margin-left:12px">'+mal+'</span></div>';
}

/* ── Modal nuevo ajuste ── */
function openNuevoAjuste(){
  if(!_session||!['admin','admin_auditor','auditor'].includes(_session.rol)){toast('⚠ Sin permisos');return;}
  var meses=['Enero','Febrero','Marzo','Abril','Mayo','Junio',
    'Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'];
  var tiendas=uniq(STORE.tareas.map(function(t){return t.tienda;})
    .concat(STORE.auditorias.map(function(a){return a.tienda;})));
  var yoSesion=(_session&&(_session.nombre||_session.username))||'';
  var auditoresLista=uniq((getAuditores()||[]).concat(yoSesion?[yoSesion]:[]));
  var html='<div class="form-grid">'+
    '<div class="form-field full"><label>Realizado por (auditor) *</label>'+
      '<input id="naj-auditor" list="naj-auditor-list" value="'+esc(yoSesion)+'" placeholder="Nombre del auditor" autocomplete="off">'+
      '<datalist id="naj-auditor-list">'+auditoresLista.map(function(a){return'<option value="'+esc(a)+'">';}).join('')+'</datalist>'+
      '<div style="font-size:10px;color:var(--muted);margin-top:3px">Se rellena con tu usuario; puedes cambiarlo si el ajuste lo hizo otra persona.</div></div>'+
    '<div class="form-field"><label>Tienda *</label>'+
      '<input id="naj-tienda" list="naj-tienda-list" placeholder="Nombre de tienda">'+
      '<datalist id="naj-tienda-list">'+tiendas.map(function(t){return'<option value="'+esc(t)+'">';}).join('')+'</datalist></div>'+
    '<div class="form-field"><label>Mes *</label>'+
      '<select id="naj-mes">'+meses.map(function(m){return'<option>'+m+'</option>';}).join('')+'</select></div>'+
    '<div class="form-field"><label>Fecha correo *</label><input type="date" id="naj-correo" oninput="najautoCalc()"></div>'+
    '<div class="form-field"><label>Fecha ajuste *</label><input type="date" id="naj-ajuste" oninput="najautoCalc()"></div>'+
    '<div class="form-field full" style="background:var(--soft);border-radius:var(--radius-sm);padding:10px 12px;border:1px solid var(--border)">'+
      '<div style="font-size:11px;font-weight:700;color:var(--muted);text-transform:uppercase;letter-spacing:.04em;margin-bottom:4px">Días / Condición (calculado)</div>'+
      '<div id="naj-preview" style="font-size:20px;font-weight:800;color:var(--blue)">—</div></div>'+
  '</div><div id="naj-err" style="font-size:12px;color:var(--red);min-height:16px;margin-top:8px"></div>';
  openModal('➕ Nuevo ajuste',html,[
    {label:'Cancelar',cls:'btn-ghost',fn:closeModal},
    {label:'Guardar',cls:'btn-blue',fn:guardarNuevoAjuste}
  ]);
}

function najautoCalc(){
  var c=document.getElementById('naj-correo').value;
  var a=document.getElementById('naj-ajuste').value;
  var dias=diasEntre(c,a);
  var prev=document.getElementById('naj-preview');
  if(dias===null||!c||!a){prev.textContent='—';prev.style.color='var(--blue)';return;}
  var cond=condicionAjuste(dias);
  prev.textContent=dias+' día(s) — '+cond;
  prev.style.color=cond==='A tiempo'?'#16a34a':'#dc2626';
}

async function guardarNuevoAjuste(){
  var errEl=document.getElementById('naj-err');errEl.textContent='';
  var auditor=(document.getElementById('naj-auditor').value||'').trim()
              ||(_session&&(_session.nombre||_session.username))||'Desconocido';
  var tienda=document.getElementById('naj-tienda').value.trim().toUpperCase();
  var mes=document.getElementById('naj-mes').value;
  var correo=document.getElementById('naj-correo').value;
  var ajuste=document.getElementById('naj-ajuste').value;
  if(!tienda){errEl.textContent='La tienda es obligatoria';return;}
  if(!correo||!ajuste){errEl.textContent='Ambas fechas son obligatorias';return;}
  var dias=diasEntre(correo,ajuste);
  /* Año predominante = fecha de ajuste; si aún no existe, se usa fecha de correo */
  var año=parseInt(((ajuste||correo)).split('-')[0]);
  var _ajRaw={tienda,mes:mes.toUpperCase(),año,fecha_correo:correo,fecha_ajuste:ajuste,
    dias,condicion:condicionAjuste(dias),auditor:auditor||null};
  var row=await encObj(_ajRaw,FIELDS.ajustes);
  var client=getSbClient();
  if(!client){toast('⚠ Sin Supabase');return;}
  closeModal();
  try{
    var r=await client.from('ajustes').insert([row]).select();
    if(r.error){toast('⚠ '+r.error.message);return;}
    var inserted=(r.data&&r.data[0])||{};
    /* Guardar en memoria los valores EN CLARO (no los que devuelve Supabase,
       que vienen cifrados). Solo se toma el id del registro insertado. */
    AJUSTES.push({id:inserted.id,tienda:tienda,mes:mes.toUpperCase(),año:año,
      fechaCorreo:correo,fechaAjuste:ajuste,
      dias:dias,condicion:condicionAjuste(dias),auditor:auditor||''});
    fillAjFilters();renderAjustes();
    toast('☁️ Ajuste guardado');
  }catch(e){toast('⚠ '+e.message);}
}

/* ── Modal editar ajuste (solo admin) ── */
function openEditAjuste(id){
  if(!_session||!['admin','admin_auditor','auditor'].includes(_session.rol)){toast('⚠ Sin permisos');return;}
  var a=AJUSTES.find(function(x){return x.id===id;});if(!a)return;
  var meses=['Enero','Febrero','Marzo','Abril','Mayo','Junio',
    'Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'];
  var yoSesion=(_session&&(_session.nombre||_session.username))||'';
  var auditorVal=(a.auditor!==undefined&&a.auditor!==null&&a.auditor!=='')?a.auditor:yoSesion;
  var auditoresLista=uniq((getAuditores()||[]).concat(yoSesion?[yoSesion]:[]));
  var html='<div class="form-grid">'+
    '<div class="form-field full"><label>Realizado por (auditor)</label>'+
      '<input id="eaj-auditor" list="eaj-auditor-list" value="'+esc(auditorVal)+'" placeholder="Nombre del auditor" autocomplete="off">'+
      '<datalist id="eaj-auditor-list">'+auditoresLista.map(function(x){return'<option value="'+esc(x)+'">';}).join('')+'</datalist></div>'+
    '<div class="form-field"><label>Tienda</label><input id="eaj-tienda" value="'+esc(a.tienda)+'"></div>'+
    '<div class="form-field"><label>Mes</label><select id="eaj-mes">'+
      meses.map(function(m){return'<option'+(norm(a.mes)===norm(m)?' selected':'')+'>'+m+'</option>';}).join('')+'</select></div>'+
    '<div class="form-field"><label>Fecha correo</label><input type="date" id="eaj-correo" value="'+(a.fechaCorreo||'')+'" oninput="eajautoCalc()"></div>'+
    '<div class="form-field"><label>Fecha ajuste</label><input type="date" id="eaj-ajuste" value="'+(a.fechaAjuste||'')+'" oninput="eajautoCalc()"></div>'+
    '<div class="form-field full" style="background:var(--soft);border-radius:var(--radius-sm);padding:10px 12px;border:1px solid var(--border)">'+
      '<div style="font-size:11px;font-weight:700;color:var(--muted);text-transform:uppercase;letter-spacing:.04em;margin-bottom:4px">Días / Condición</div>'+
      '<div id="eaj-preview" style="font-size:20px;font-weight:800;color:'+(a.condicion==='A tiempo'?'#16a34a':'#dc2626')+'">'+
        (a.dias!==null?a.dias:0)+' día(s) — '+esc(a.condicion)+'</div></div>'+
  '</div>';
  openModal('✎ Editar ajuste — '+esc(a.tienda),html,[
    {label:'Cancelar',cls:'btn-ghost',fn:closeModal},
    {label:'Eliminar',cls:'btn-red',fn:function(){deleteAjuste(id);}},
    {label:'Guardar',cls:'btn-blue',fn:function(){saveEditAjuste(id);}}
  ]);
}

function eajautoCalc(){
  var c=document.getElementById('eaj-correo').value;
  var a=document.getElementById('eaj-ajuste').value;
  var dias=diasEntre(c,a);
  var prev=document.getElementById('eaj-preview');
  if(dias===null){prev.textContent='—';return;}
  var cond=condicionAjuste(dias);
  prev.textContent=dias+' día(s) — '+cond;
  prev.style.color=cond==='A tiempo'?'#16a34a':'#dc2626';
}

async function saveEditAjuste(id){
  function dv(elId){var e=document.getElementById(elId);return e?e.value.trim():'';}
  var tienda=dv('eaj-tienda').toUpperCase();
  var mes=dv('eaj-mes').toUpperCase();
  var correo=dv('eaj-correo');
  var ajuste=dv('eaj-ajuste');
  var _ajActual=AJUSTES.find(function(x){return x.id===id;})||{};
  var auditor=dv('eaj-auditor')||_ajActual.auditor||(_session&&(_session.nombre||_session.username))||'Desconocido';
  var dias=diasEntre(correo,ajuste);
  /* Año predominante = fecha de ajuste; si aún no existe, se usa fecha de correo */
  var año=parseInt(((ajuste||correo)||'').split('-')[0])||null;
  var _ajUpdRaw={tienda,mes,año,fecha_correo:correo,fecha_ajuste:ajuste,dias,condicion:condicionAjuste(dias),auditor:auditor};
  var row=await encObj(_ajUpdRaw,FIELDS.ajustes);
  var client=getSbClient();if(!client)return;
  closeModal();
  try{
    var r=await client.from('ajustes').update(row).eq('id',id);
    if(r.error){toast('⚠ '+r.error.message);return;}
    var idx=AJUSTES.findIndex(function(x){return x.id===id;});
    if(idx>=0)AJUSTES[idx]=Object.assign({},AJUSTES[idx],{tienda,mes,año,fechaCorreo:correo,fechaAjuste:ajuste,dias,condicion:condicionAjuste(dias),auditor:auditor});
    fillAjFilters();renderAjustes();toast('☁️ Ajuste actualizado');
  }catch(e){toast('⚠ '+e.message);}
}

async function deleteAjuste(id){
  if(!confirm('¿Eliminar este ajuste?'))return;
  var client=getSbClient();if(!client)return;
  closeModal();
  try{
    var r=await client.from('ajustes').delete().eq('id',id);
    if(r.error){toast('⚠ '+r.error.message);return;}
    AJUSTES=AJUSTES.filter(function(x){return x.id!==id;});
    fillAjFilters();renderAjustes();toast('☁️ Ajuste eliminado');
  }catch(e){toast('⚠ '+e.message);}
}
/* ── Importar Excel de Ajustes ── */
function importAjustesExcel(files){
  if(!files||!files.length)return;
  if(!_session||!['admin','admin_auditor','auditor'].includes(_session.rol)){toast('⚠ Sin permisos');return;}
  var file=files[0];
  document.getElementById('aj-file-input').value='';
  var reader=new FileReader();
  reader.onload=function(e){
    try{
      /* cellDates:false → recibir seriales numéricos y convertir manualmente */
      var wb=XLSX.read(e.target.result,{type:'array',cellDates:false});
      var _pk=pickSheet(wb,['ajustes'],['tienda','ajuste','correo']);var ws=_pk.ws,raw=_pk.raw;
      if(raw.length<2){toast('⚠ Excel vacío o sin datos');return;}
      /* Normalizar cabeceras: minúsculas + sin tildes + sin asteriscos/símbolos + sin espacios extra */
      function normH(s){return String(s||'').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/[^a-z0-9]+/g,' ').trim();}
      var H=raw[0].map(normH);
      /* Buscar columna por lista de variantes, con coincidencia parcial como respaldo */
      function fc(keys){
        var nk=keys.map(normH);
        for(var ki=0;ki<nk.length;ki++)for(var hi=0;hi<H.length;hi++)if(H[hi]===nk[ki])return hi;
        for(var ki=0;ki<nk.length;ki++)for(var hi=0;hi<H.length;hi++)if(H[hi].includes(nk[ki])||nk[ki].includes(H[hi]))return hi;
        return -1;
      }
      var iT=fc(['tienda','sucursal']);
      var iC=fc(['correo','fecha correo','fecha_correo','f correo','fcorreo']);
      var iA=fc(['ajuste','fecha ajuste','fecha_ajuste','f ajuste','fajuste']);
      var iAud=fc(['auditor','realizado por','realizadopor','responsable']);
      if(iT<0||iC<0||iA<0){
        toast('⚠ Columnas no encontradas. Necesita: Tienda, Correo, Ajuste');return;
      }
      var yoSesion=(_session&&(_session.nombre||_session.username))||'';
      var MN=['ENERO','FEBRERO','MARZO','ABRIL','MAYO','JUNIO',
              'JULIO','AGOSTO','SEPTIEMBRE','OCTUBRE','NOVIEMBRE','DICIEMBRE'];
      /* Convierte serial Excel numérico O string de fecha a 'yyyy-mm-dd' */
      function toISO(v){
        if(!v&&v!==0)return null;
        var n=Number(v);
        if(!isNaN(n)&&n>30000&&n<80000){
          /* Epoch Unix en ms: (serial - 25569) * 86400 * 1000 */
          var ms=(n-25569)*86400*1000;
          var d=new Date(ms);
          return d.getUTCFullYear()+'-'+
            String(d.getUTCMonth()+1).padStart(2,'0')+'-'+
            String(d.getUTCDate()).padStart(2,'0');
        }
        var s=String(v).trim();
        var m=s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
        if(m)return m[3]+'-'+m[2].padStart(2,'0')+'-'+m[1].padStart(2,'0');
        if(/^\d{4}-\d{2}-\d{2}$/.test(s))return s;
        var dt=new Date(s);
        if(!isNaN(dt))return dt.getFullYear()+'-'+
          String(dt.getMonth()+1).padStart(2,'0')+'-'+
          String(dt.getDate()).padStart(2,'0');
        return null;
      }
      var rows=[];
      for(var i=1;i<raw.length;i++){
        var r=raw[i];
        var tienda=String(r[iT]||'').trim().toUpperCase();
        if(!tienda)continue;
        var fechaC=toISO(r[iC]);
        var fechaA=toISO(r[iA]);
        /* Basta la fecha de correo. Un ajuste del mes en curso puede no tener
           todavía fecha de ajuste: antes se descartaba la fila entera y el mes
           no aparecía. Ahora entra como pendiente (condición '—'). */
        if(!fechaC)continue;
        /* Derivar mes y año SIEMPRE desde la FECHA DE AJUSTE (predominante); si la
           fila aún no tiene fecha de ajuste (pendiente de validar), se usa fecha
           de correo como respaldo. Se ignora el campo Mes del Excel porque viene
           con formatos inconsistentes: "julio2023.", "julio", etc. */
        var _fechaBaseAjImp=fechaA||fechaC;
        var parts=_fechaBaseAjImp.split('-');
        var año=parseInt(parts[0]);
        var mes=MN[parseInt(parts[1])-1]||'';
        var dias=fechaA?diasEntre(fechaC,fechaA):null;
        var auditor=(iAud>=0?String(r[iAud]||'').trim():'')||yoSesion||'Desconocido';
        rows.push({tienda,mes,año,fecha_correo:fechaC,fecha_ajuste:fechaA,
          dias,condicion:condicionAjuste(dias),auditor:auditor});
      }
      if(!rows.length){toast('⚠ Sin filas válidas en el Excel');return;}
      toast('⏳ Importando '+rows.length+' registros…');
      commitAjustesExcel(rows);
    }catch(err){toast('⚠ Error: '+err.message);console.error(err);}
  };
  reader.readAsArrayBuffer(file);
}

async function commitAjustesExcel(rows){
  var client=getSbClient();
  if(!client){toast('⚠ Sin Supabase');return;}
  try{
    /* Traer existentes para actualizar en vez de duplicar. Clave: tienda +
       fecha_correo + fecha_ajuste (las fechas no están cifradas). */
    var ex=[];
    try{var re=await client.from('ajustes').select('id,fecha_correo,fecha_ajuste,tienda').limit(20000);if(!re.error)ex=re.data||[];}catch(_e){}
    var exDec=await decArr(ex,FIELDS.ajustes);
    var _k=function(t,fc,fa){return norm(t)+'|'+String(fc||'')+'|'+String(fa||'');};
    var exMap={};
    exDec.forEach(function(x){exMap[_k(x.tienda,x.fecha_correo,x.fecha_ajuste)]=x.id;});
    /* dedup dentro del mismo archivo (conserva la última) */
    var seen={},dedup=[];
    for(var z=rows.length-1;z>=0;z--){var kk=_k(rows[z].tienda,rows[z].fecha_correo,rows[z].fecha_ajuste);if(seen[kk])continue;seen[kk]=1;dedup.unshift(rows[z]);}
    var ins=0,upd=0,err=null;
    for(var i=0;i<dedup.length&&!err;i++){
      var raw=dedup[i];
      var row=await encObj(raw,FIELDS.ajustes);   /* ← CIFRAR antes de guardar */
      var k=_k(raw.tienda,raw.fecha_correo,raw.fecha_ajuste);
      if(exMap[k]){
        var ru=await client.from('ajustes').update(row).eq('id',exMap[k]);
        if(ru.error)err=ru.error.message;else upd++;
      }else{
        var ri=await client.from('ajustes').insert([row]).select('id');
        if(ri.error)err=ri.error.message;else{ins++;if(ri.data&&ri.data[0])exMap[k]=ri.data[0].id;}
      }
    }
    if(err){toast('⚠ '+err);return;}
    toast('☁️ '+ins+' nuevos · '+upd+' actualizados');
    _ajPerManual=false; /* tras importar, volver al mes en curso */
    await loadAjustes();
  }catch(e){toast('⚠ '+e.message);}
}
/* ════════════════════════════════════════════════════════════════════
   MÓDULO MERMAS — Supabase
   Tabla: mermas
   Columnas: id (uuid PK), tienda (text), local_foranea (text),
             mes (text), año (int), fecha_autorizacion (date),
             fecha_validacion (date), dias (int), condicion (text),
             created_at, updated_at
════════════════════════════════════════════════════════════════════ */
var MERMAS=[];
var _mrPerManual=false; /* true cuando el usuario elige mes/año a mano: la carga no lo reescribe */
var _chartMrDonut=null;

async function loadMermas(){
  var client=getSbClient();
  if(!client){toast('⚠ Sin Supabase');return;}
  try{
    var r=await client.from('mermas').select('*').order('fecha_autorizacion',{ascending:true}).order('fecha_validacion',{ascending:true}).limit(5000);
    if(r.error){toast('⚠ '+r.error.message);return;}
    var _mrDataRaw=r.data||[];
    var _mrDataDec=await decArr(_mrDataRaw,FIELDS.mermas);
    var MNmr=['ENERO','FEBRERO','MARZO','ABRIL','MAYO','JUNIO',
      'JULIO','AGOSTO','SEPTIEMBRE','OCTUBRE','NOVIEMBRE','DICIEMBRE'];
    MERMAS=_mrDataDec.map(function(m){
      var dias=diasEntre(m.fecha_autorizacion,m.fecha_validacion);
      /* SIEMPRE derivar mes y año desde la FECHA DE VALIDACIÓN (predominante). Si
         aún no hay fecha de validación (pendiente), se usa fecha_autorizacion
         como respaldo para que el registro no desaparezca de los filtros. */
      var mesLimpioMr='';
      var añoLimpioMr=null;
      var _fechaBaseMr=m.fecha_validacion||m.fecha_autorizacion;
      if(_fechaBaseMr){
        var partsMr=_fechaBaseMr.split('-');
        if(partsMr.length>=2){
          añoLimpioMr=parseInt(partsMr[0]);
          mesLimpioMr=MNmr[parseInt(partsMr[1])-1]||'';
        }
      }
      /* Fallback: intentar parsear el campo mes almacenado si no hay ninguna fecha */
      if(!mesLimpioMr&&m.mes){
        var rawMr=String(m.mes).toUpperCase().replace(/[0-9.]/g,'').trim();
        var foundMr=MNmr.find(function(mn){return rawMr.startsWith(mn)||rawMr===mn;});
        mesLimpioMr=foundMr||rawMr;
      }
      if(!añoLimpioMr&&m.año)añoLimpioMr=parseInt(m.año);
      return{id:m.id,tienda:m.tienda||'',localForanea:m.local_foranea||'',
        mes:mesLimpioMr,año:añoLimpioMr,
        fechaAut:m.fecha_autorizacion,fechaVal:m.fecha_validacion,
        dias:dias,condicion:condicionAjuste(dias),auditor:m.auditor||'',
        razon:razonDeCentro(centroDeTienda(m.tienda||''))};
    });
    MERMAS=MERMAS.filter(function(m){return tiendaVisible(m.tienda);});
    /* Solo se conservan registros de 2026 en adelante (los años anteriores no se necesitan) */
    MERMAS=MERMAS.filter(function(m){return (parseInt(m.año)||0)>=AÑO_MIN_MODULOS;});
    fillMrFilters();renderMermas();
    toast('✓ '+MERMAS.length+' merma(s) cargadas');
  }catch(e){toast('⚠ '+e.message);}
}

function fillMrFilters(){
  var MORD=['ENERO','FEBRERO','MARZO','ABRIL','MAYO','JUNIO',
    'JULIO','AGOSTO','SEPTIEMBRE','OCTUBRE','NOVIEMBRE','DICIEMBRE'];
  var mesSet=new Set(MERMAS.map(function(m){return(m.mes||'').toUpperCase();}).filter(Boolean));
  var meses=MORD.filter(function(m){return mesSet.has(m);});
  mesSet.forEach(function(m){if(MORD.indexOf(m)<0)meses.push(m);});
  var años=[...new Set(MERMAS.map(function(m){return m.año?String(m.año):'';}).filter(Boolean))].sort().reverse();
  var tiendas=[...new Set(MERMAS.map(function(m){return m.tienda;}).filter(Boolean))].sort();
  var sm=document.getElementById('mr-f-mes');
  var sa=document.getElementById('mr-f-año');
  var st=document.getElementById('mr-f-tienda');
  if(!sm)return;
  var vm=sm.value,va=sa.value,vt=st.value;
  sm.innerHTML='<option value="ALL">Todos</option>'+meses.map(function(m){return'<option value="'+m+'">'+m+'</option>';}).join('');
  sa.innerHTML='<option value="ALL">Todos</option>'+años.map(function(y){return'<option value="'+y+'">'+y+'</option>';}).join('');
  st.innerHTML='<option value="ALL">Todas</option>'+limpiarOpciones(tiendas).map(function(t){return'<option>'+t+'</option>';}).join('');
  var mesHoy=MORD[new Date().getMonth()];
  var añoHoy=String(new Date().getFullYear());
  function _tienePerM(m,y){return MERMAS.some(function(x){return (x.mes||'').toUpperCase()===m&&String(x.año||'')===String(y);});}
  var tgtMes=mesHoy, tgtAño=añoHoy;
  if(!_tienePerM(mesHoy,añoHoy)){
    var best=null;
    MERMAS.forEach(function(x){
      var mi=MORD.indexOf((x.mes||'').toUpperCase()); var y=parseInt(x.año)||0;
      if(mi<0||!y)return;
      var key=y*100+mi;
      if(!best||key>best.key)best={key:key,mes:(x.mes||'').toUpperCase(),año:String(x.año)};
    });
    if(best){tgtMes=best.mes;tgtAño=best.año;}
  }
  /* Mes/año por defecto = período en curso si tiene datos; si no, el más reciente.
     Se recalcula en cada carga salvo elección manual (_mrPerManual). */
  if(_mrPerManual){
    if(vm!=='ALL'&&[...sm.options].some(function(o){return o.value===vm;}))sm.value=vm;
    if(va!=='ALL'&&[...sa.options].some(function(o){return o.value===va;}))sa.value=va;
  }else{
    sm.value=[...sm.options].some(function(o){return o.value===tgtMes;})?tgtMes:'ALL';
    sa.value=[...sa.options].some(function(o){return o.value===tgtAño;})?tgtAño:'ALL';
  }
  if(vt!=='ALL'&&[...st.options].some(function(o){return o.value===vt;}))st.value=vt;
}
function mrPerChanged(){_mrPerManual=true;renderMermas();}
function filteredMermas(){
  var mes=(document.getElementById('mr-f-mes').value||'ALL');
  var año=(document.getElementById('mr-f-año').value||'ALL');
  var tienda=(document.getElementById('mr-f-tienda').value||'ALL');
  return MERMAS.filter(function(m){
    if(mes!=='ALL'&&(m.mes||'').toUpperCase()!==mes.toUpperCase())return false;
    if(año!=='ALL'&&String(m.año||'')!==String(año))return false;
    if(tienda!=='ALL'&&norm(m.tienda)!==norm(tienda))return false;
    return true;
  });
}

function renderMermas(){
  var arr=filteredMermas();
  /* Ordenar por fecha de autorización: día 1 primero, último día al final */
  arr=arr.slice().sort(function(a,b){
    var da=a.fechaAut||'';var db=b.fechaAut||'';
    return da<db?-1:da>db?1:0;
  });
  var isAdmin=_session&&['admin','admin_auditor','auditor'].includes(_session.rol);
  document.getElementById('mr-count').textContent=arr.length+' registro(s)';
  var th=document.getElementById('mr-edit-th');
  if(th)th.style.display=isAdmin?'':'none';

  var tbody=document.getElementById('mr-tbody');
  if(!arr.length){
    tbody.innerHTML='<tr><td colspan="9" style="text-align:center;color:var(--muted);padding:24px">Sin registros para los filtros seleccionados.</td></tr>';
  }else{
    tbody.innerHTML=arr.map(function(m,i){
      var cls=m.condicion==='A tiempo'?'b-atiempo':'b-destiempo';
      var lf=m.localForanea||'—';
      var lfColor=norm(lf).includes('foranea')||norm(lf).includes('foránea')?'color:#ea580c;font-weight:700':'color:var(--k-blue);font-weight:700';
      var editTd=isAdmin?'<td style="text-align:center"><button class="icon-btn" onclick="openEditMerma(\''+m.id+'\')" title="Editar">✎</button></td>':'';
      return '<tr>'+
        '<td style="color:var(--muted);font-size:11px">'+(i+1)+'</td>'+
        '<td><b>'+esc(m.tienda)+'</b></td>'+
        '<td style="'+lfColor+'">'+esc(lf)+'</td>'+
        '<td>'+esc(m.mes)+'</td>'+
        '<td style="text-align:center">'+fmtFecha(m.fechaAut)+'</td>'+
        '<td style="text-align:center">'+fmtFecha(m.fechaVal)+'</td>'+
        '<td style="text-align:center;font-weight:800">'+(m.dias!==null?m.dias:'—')+'</td>'+
        '<td style="text-align:center"><span class="'+cls+'">'+esc(m.condicion)+'</span></td>'+
        editTd+'</tr>';
    }).join('');
  }
  renderMrDonut(arr);
}

function renderMrDonut(arr){
  var ok=arr.filter(function(m){return m.condicion==='A tiempo';}).length;
  var mal=arr.filter(function(m){return m.condicion==='Destiempo';}).length;
  var total=arr.length;
  document.getElementById('mr-donut-note').textContent=total?total+' registros':'';
  if(_chartMrDonut){_chartMrDonut.destroy();_chartMrDonut=null;}
  var wrap=document.getElementById('chart-mr-donut');
  if(!wrap||wrap.tagName!=='CANVAS'){
    var parent=wrap?wrap.parentElement:document.querySelector('#view-mermas .ajustes-donut-wrap div');
    if(parent)parent.innerHTML='<canvas id="chart-mr-donut"></canvas>';
  }
  var legend=document.getElementById('mr-donut-legend');
  if(!total){
    var cv=document.getElementById('chart-mr-donut');
    if(cv)cv.parentElement.innerHTML='<div class="empty" style="display:flex;align-items:center;justify-content:center;width:165px;height:165px;color:var(--muted);font-size:12px">Sin datos</div>';
    if(legend)legend.innerHTML='';
    return;
  }
  var canvas=document.getElementById('chart-mr-donut');
  if(!canvas)return;
  _chartMrDonut=new Chart(canvas,{
    type:'doughnut',
    data:{labels:['A tiempo','Destiempo'],
      datasets:[{data:[ok,mal],backgroundColor:['#16a34a','#dc2626'],borderWidth:2,borderColor:'#fff'}]},
    options:{responsive:true,maintainAspectRatio:false,cutout:'65%',plugins:{legend:{display:false},
      tooltip:{callbacks:{label:function(ctx2){return ctx2.label+': '+ctx2.raw+' ('+Math.round(ctx2.raw/(total||1)*100)+'%)';}}}}}
  });
  if(legend)legend.innerHTML=
    '<div class="legend-item"><span class="legend-dot" style="background:#16a34a"></span>'+
    '<span class="legend-lbl">A tiempo</span><span class="legend-val" style="margin-left:12px">'+ok+'</span></div>'+
    '<div class="legend-item"><span class="legend-dot" style="background:#dc2626"></span>'+
    '<span class="legend-lbl">Destiempo</span><span class="legend-val" style="margin-left:12px">'+mal+'</span></div>';
}

/* ════════════════════════════════════════════════════════════════════
   MÓDULO EVALUACIÓN KPIs — Mermas · Ajustes · Auditorías (dinámico)
════════════════════════════════════════════════════════════════════ */
var _evCharts={};
function evDestroy(k){if(_evCharts[k]){_evCharts[k].destroy();_evCharts[k]=null;}}

function evKpiCard(label,pct,n,color,icon,note){
  var safe=isFinite(pct)?Math.round(pct):0;
  return '<div class="ev-kpi-card">'+
    '<div class="ev-kpi-top"><span class="ev-kpi-icon" style="background:linear-gradient(135deg,'+color+',#1a1f3c)">'+icon+'</span>'+
    '<div><div class="ev-kpi-lbl">'+label+'</div><div class="ev-kpi-n">'+n+' registros</div></div></div>'+
    '<div class="ev-kpi-bar"><div class="ev-kpi-bar-fill" style="width:'+safe+'%;background:linear-gradient(90deg,'+color+',#0ce7fe)"></div></div>'+
    '<div class="ev-kpi-pct" style="color:'+color+'">'+safe+'% cumplimiento</div>'+
    (note?'<div style="font-size:10px;color:var(--muted);font-weight:600;margin-top:6px">'+note+'</div>':'')+
  '</div>';
}

function evMesesUltimos(n){
  var MN=['ENE','FEB','MAR','ABR','MAY','JUN','JUL','AGO','SEP','OCT','NOV','DIC'];
  var out=[],d=new Date();
  d.setDate(1);
  for(var i=n-1;i>=0;i--){
    var dd=new Date(d.getFullYear(),d.getMonth()-i,1);
    out.push({key:MN[dd.getMonth()],año:dd.getFullYear(),mesIdx:dd.getMonth()});
  }
  return out;
}

var _evMesManual=false; /* true cuando el usuario elige mes a mano en Evaluación */
function evMesChanged(){_evMesManual=true;renderEvaluacion();}
function renderEvaluacion(){
  var host=document.getElementById('view-evaluacion');
  if(!host)return;
  if(!host.dataset.built){
    host.dataset.built='1';
    host.innerHTML=
    '<div class="ev-hdr">'+
      '<div><h2>🎯 Evaluación de KPIs</h2><p>Cumplimiento de Mermas, Ajustes, Auditorías y Actividades — actualizado en vivo</p></div>'+
      '<div style="display:flex;align-items:center;gap:8px;flex-wrap:wrap">'+
        '<select id="ev-mes" onchange="evMesChanged()" title="Mes (por defecto el mes en curso)" style="padding:7px 10px;border-radius:var(--radius-sm);border:1px solid var(--border);font-size:13px;font-family:inherit;background:var(--soft)"></select>'+
        '<div style="position:relative">'+
          '<span style="position:absolute;left:10px;top:50%;transform:translateY(-50%);font-size:12px;color:var(--muted);pointer-events:none">🔍</span>'+
          '<select id="ev-auditor" onchange="renderEvaluacion()" style="padding:7px 10px 7px 30px;border-radius:var(--radius-sm);border:1px solid var(--border);font-size:13px;font-family:inherit;background:var(--soft)"></select>'+
        '</div>'+
        '<button class="btn btn-blue" onclick="renderEvaluacion()">⟳ Actualizar</button>'+
      '</div>'+
    '</div>'+
    '<div class="ev-kpi-grid" id="ev-kpi-grid"></div>'+
    '<div class="ev-chart-grid">'+
      '<div class="card ev-chart-card"><h3>Tendencia de cumplimiento por mes</h3><div class="ev-chart-box" style="height:240px"><canvas id="ev-trend"></canvas></div></div>'+
      '<div class="card ev-chart-card"><h3>Ajustes: a tiempo vs destiempo</h3><div class="ev-chart-box" style="height:240px"><canvas id="ev-aj-donut"></canvas></div></div>'+
      '<div class="card ev-chart-card"><h3>Mermas: a tiempo vs destiempo</h3><div class="ev-chart-box" style="height:240px"><canvas id="ev-mr-donut"></canvas></div></div>'+
      '<div class="card ev-chart-card" style="grid-column:span 2"><h3>Auditorías: % resuelto promedio por clase</h3><div class="ev-chart-box" style="height:260px"><canvas id="ev-aud-bar"></canvas></div></div>'+
      '<div class="card ev-chart-card"><h3>Actividades: a tiempo vs fuera de tiempo</h3><div class="ev-chart-box" style="height:240px"><canvas id="ev-act-donut"></canvas></div></div>'+
    '</div>';
  }

  /* ── Filtros: barra global aporta razón/tienda/fechas; el mes lo controla
     este módulo con su propio selector (por defecto el mes en curso). ── */
  var gf=(typeof getFilterState==='function')?getFilterState():null;
  function refDate(s){return s?new Date(String(s).split('T')[0]+'T12:00:00'):null;}
  /* gBase: razón + tienda + rango de fechas de la barra global (SIN mes: el mes
     se aplica aparte para poder mostrar la tendencia de 6 meses completa). */
  function gBase(rec,fechaRef){
    if(!gf)return true;
    if(gf.razon&&gf.razon!=='ALL'){
      var _rz=rec.razon||razonDeCentro(centroDeTienda(rec.tienda||''));
      if(razKey(_rz)!==razKey(gf.razon))return false;
    }
    if(gf.tienda&&gf.tienda!=='ALL'&&norm(rec.tienda||'')!==norm(gf.tienda))return false;
    if(gf.desde||gf.hasta){
      if(!fechaRef)return false;
      if(gf.desde&&fechaRef<gf.desde)return false;
      if(gf.hasta&&fechaRef>gf.hasta)return false;
    }
    return true;
  }

  var MNFULL=['ENERO','FEBRERO','MARZO','ABRIL','MAYO','JUNIO','JULIO','AGOSTO','SEPTIEMBRE','OCTUBRE','NOVIEMBRE','DICIEMBRE'];

  /* Filtro por auditor (solo Ajustes y Mermas) */
  var selEl=document.getElementById('ev-auditor');
  var audSel=selEl?selEl.value:'ALL';
  var auds=getAuditores();
  if(selEl){
    selEl.innerHTML='<option value="ALL">Todos los auditores</option>'+
      auds.map(function(n){return'<option value="'+esc(n)+'"'+(norm(n)===norm(audSel)?' selected':'')+'>'+esc(n)+'</option>';}).join('');
    if(audSel!=='ALL'&&!auds.some(function(n){return norm(n)===norm(audSel);}))audSel='ALL';
    selEl.value=audSel;
  }
  function audMatch(name){return audSel==='ALL'||norm(name||'')===norm(audSel);}

  /* Auditorías base (razón/centro/tienda/fechas de la barra global, SIN mes) */
  var audAll=(STORE.auditorias||[]).filter(function(a){
    if(gf){
      if(gf.razon&&gf.razon!=='ALL'&&razKey(a.razon)!==razKey(gf.razon))return false;
      if(gf.centro&&gf.centro!=='ALL'&&norm(a.centro)!==norm(gf.centro))return false;
      if(gf.tienda&&gf.tienda!=='ALL'&&norm(a.tienda)!==norm(gf.tienda))return false;
      if(gf.desde||gf.hasta){var d=refDate(a.fecha);if(!d)return false;if(gf.desde&&d<gf.desde)return false;if(gf.hasta&&d>gf.hasta)return false;}
    }
    return true;
  });

  /* Datasets base (todos los meses) para la tendencia */
  var AJall=AJUSTES.filter(function(a){return audMatch(a.auditor)&&gBase(a,refDate(a.fechaAjuste||a.fechaCorreo));});
  var MRall=MERMAS.filter(function(m){return audMatch(m.auditor)&&gBase(m,refDate(m.fechaVal||m.fechaAut));});
  /* Actividades: mismo criterio de asignado que Desempeño (asignado, o si no,
     quien la creó), y misma comparación de fechas (Real Fin vs Est. Fin) que
     ya se usa ahí para no divergir entre módulos. */
  var ACTall=ACTIVIDADES.filter(function(a){return audMatch(a.asignado||a.creadoPor)&&gBase(a,refDate(a.realFin||a.estFin));});

  /* Selector de mes propio: por defecto el mes en curso (o el más reciente con
     datos). Se recalcula salvo elección manual (_evMesManual). */
  var mesSel=document.getElementById('ev-mes');
  if(mesSel){
    var vEv=mesSel.value;
    var mesHoyEv=MNFULL[new Date().getMonth()];
    function _hayMesEv(mn){return AJall.concat(MRall).some(function(x){return (x.mes||'').toUpperCase()===mn;})||audAll.some(function(a){return (a.mes||'').toUpperCase()===mn;})||ACTall.some(function(a){return (a.mes||'').toUpperCase()===mn;});}
    var tgtEv=mesHoyEv;
    if(!_hayMesEv(mesHoyEv)){
      var bi=-1;audAll.concat(AJall,MRall,ACTall).forEach(function(x){var i=MNFULL.indexOf((x.mes||'').toUpperCase());if(i>bi)bi=i;});
      if(bi>=0)tgtEv=MNFULL[bi];
    }
    mesSel.innerHTML='<option value="ALL">Todos los meses</option>'+MNFULL.map(function(m){return'<option value="'+m+'">'+m.charAt(0)+m.slice(1).toLowerCase()+'</option>';}).join('');
    if(_evMesManual&&vEv&&[...mesSel.options].some(function(o){return o.value===vEv;}))mesSel.value=vEv;
    else mesSel.value=[...mesSel.options].some(function(o){return o.value===tgtEv;})?tgtEv:'ALL';
  }
  var mesEv=mesSel?mesSel.value:'ALL';
  function mMatchEv(mn){return mesEv==='ALL'||(String(mn||'').toUpperCase()===mesEv);}

  /* Datasets del mes seleccionado (para KPIs y donas) */
  var AJ=AJall.filter(function(a){return mMatchEv(a.mes);});
  var MR=MRall.filter(function(m){return mMatchEv(m.mes);});
  var audArr=audAll.filter(function(a){return mMatchEv(a.mes);});
  var ACT=ACTall.filter(function(a){return mMatchEv(a.mes);});

  var ajOk=AJ.filter(function(a){return a.condicion==='A tiempo';}).length;
  var ajTot=AJ.length;
  var mrOk=MR.filter(function(m){return m.condicion==='A tiempo';}).length;
  var mrTot=MR.length;
  /* Actividades: cumplida = Completado + a tiempo (mismo criterio que
     Desempeño vía actividadEnTiempoDesempeno). El total considera todas las
     actividades del periodo, no solo las completadas — igual filosofía que
     Ajustes/Mermas, donde el total incluye lo que aún no se resuelve. */
  var actOk=ACT.filter(function(a){return norm(a.estado||'').includes('completad')&&actividadEnTiempoDesempeno(a);}).length;
  var actTot=ACT.length;

  /* % resuelto EN VIVO (coherente con el módulo de Auditorías); si no hay tareas
     cruzables, cae al valor almacenado. */
  function prPct(a){
    if(typeof calcAudStats==='function'){
      var s=calcAudStats(a,audArr);
      if(s&&isFinite(s.pctResuelto)&&s.tareas>0)return s.pctResuelto*100;
    }
    var v=parseFloat(a.pctResuelto);if(!isFinite(v))v=0;return v>1.5?v:v*100;
  }
  var audProm=audArr.length?audArr.reduce(function(s,a){return s+prPct(a);},0)/audArr.length:0;

  document.getElementById('ev-kpi-grid').innerHTML=
    evKpiCard('Ajustes',ajTot?ajOk/ajTot*100:0,ajTot,'#4318ff','⚖️')+
    evKpiCard('Mermas',mrTot?mrOk/mrTot*100:0,mrTot,'#0ce7fe','🗂️')+
    evKpiCard('Auditorías',audProm,audArr.length,'#01b574','🔍',audSel!=='ALL'?'Global · sin desglose por auditor':'')+
    evKpiCard('Actividades',actTot?actOk/actTot*100:0,actTot,'#f59e0b','📋');

  /* Tendencia últimos 6 meses — usa los datasets base (TODOS los meses) para no
     colapsar al filtrar por un mes concreto. */
  var meses=evMesesUltimos(6);
  var serieAj=[],serieMr=[],serieAud=[],serieAct=[],lbls=[];
  meses.forEach(function(m){
    lbls.push(m.key);
    var ajM=AJall.filter(function(a){return (a.mes||'').toUpperCase()===MNFULL[m.mesIdx]&&Number(a.año)===m.año;});
    var mrM=MRall.filter(function(x){return (x.mes||'').toUpperCase()===MNFULL[m.mesIdx]&&Number(x.año)===m.año;});
    var audM=audAll.filter(function(a){return (a.mes||'').toUpperCase().startsWith(MNFULL[m.mesIdx].slice(0,3));});
    var actM=ACTall.filter(function(a){return (a.mes||'').toUpperCase().startsWith(MNFULL[m.mesIdx].slice(0,3));});
    serieAj.push(ajM.length?Math.round(ajM.filter(function(a){return a.condicion==='A tiempo';}).length/ajM.length*100):null);
    serieMr.push(mrM.length?Math.round(mrM.filter(function(x){return x.condicion==='A tiempo';}).length/mrM.length*100):null);
    serieAud.push(audM.length?Math.round(audM.reduce(function(s,a){return s+prPct(a);},0)/audM.length):null);
    serieAct.push(actM.length?Math.round(actM.filter(function(a){return norm(a.estado||'').includes('completad')&&actividadEnTiempoDesempeno(a);}).length/actM.length*100):null);
  });

  evDestroy('trend');
  _evCharts.trend=new Chart(document.getElementById('ev-trend'),{
    type:'line',
    data:{labels:lbls,datasets:[
      {label:'Ajustes',data:serieAj,borderColor:'#4318ff',backgroundColor:'rgba(67,24,255,.15)',tension:.4,fill:true,spanGaps:true},
      {label:'Mermas',data:serieMr,borderColor:'#0ce7fe',backgroundColor:'rgba(12,231,254,.15)',tension:.4,fill:true,spanGaps:true},
      {label:'Auditorías',data:serieAud,borderColor:'#01b574',backgroundColor:'rgba(1,181,116,.15)',tension:.4,fill:true,spanGaps:true},
      {label:'Actividades',data:serieAct,borderColor:'#f59e0b',backgroundColor:'rgba(245,158,11,.15)',tension:.4,fill:true,spanGaps:true}
    ]},
    options:{responsive:true,maintainAspectRatio:false,animation:{duration:600},
      scales:{y:{min:0,max:100,ticks:{callback:function(v){return v+'%';}}}},
      plugins:{legend:{position:'bottom',labels:{boxWidth:10}}}}
  });

  evDestroy('ajdonut');
  _evCharts.ajdonut=new Chart(document.getElementById('ev-aj-donut'),{
    type:'doughnut',
    data:{labels:['A tiempo','Destiempo'],datasets:[{data:[ajOk,ajTot-ajOk],backgroundColor:['#4318ff','#fc8181'],borderWidth:0}]},
    options:{responsive:true,maintainAspectRatio:false,cutout:'70%',animation:{duration:500},plugins:{legend:{position:'bottom',labels:{boxWidth:10}}}}
  });

  evDestroy('mrdonut');
  _evCharts.mrdonut=new Chart(document.getElementById('ev-mr-donut'),{
    type:'doughnut',
    data:{labels:['A tiempo','Destiempo'],datasets:[{data:[mrOk,mrTot-mrOk],backgroundColor:['#0ce7fe','#fbb140'],borderWidth:0}]},
    options:{responsive:true,maintainAspectRatio:false,cutout:'70%',animation:{duration:500},plugins:{legend:{position:'bottom',labels:{boxWidth:10}}}}
  });

  /* Auditorías por clase */
  var porClase={};
  audArr.forEach(function(a){
    var c=a.clase||'SIN CLASE';
    if(!porClase[c])porClase[c]={s:0,n:0};
    porClase[c].s+=prPct(a);porClase[c].n++;
  });
  var clases=Object.keys(porClase);
  var datosClase=clases.map(function(c){return Math.round(porClase[c].s/porClase[c].n);});
  evDestroy('audbar');
  _evCharts.audbar=new Chart(document.getElementById('ev-aud-bar'),{
    type:'bar',
    data:{labels:clases.map(function(c){return c.length>22?c.slice(0,22)+'…':c;}),
      datasets:[{label:'% resuelto promedio',data:datosClase,backgroundColor:'#9f7aea',borderRadius:8,maxBarThickness:38}]},
    options:{responsive:true,maintainAspectRatio:false,animation:{duration:600},
      scales:{y:{min:0,max:100,ticks:{callback:function(v){return v+'%';}}}},
      plugins:{legend:{display:false}}}
  });

  evDestroy('actdonut');
  _evCharts.actdonut=new Chart(document.getElementById('ev-act-donut'),{
    type:'doughnut',
    data:{labels:['A tiempo','Fuera de tiempo / pendiente'],datasets:[{data:[actOk,actTot-actOk],backgroundColor:['#f59e0b','#94a3b8'],borderWidth:0}]},
    options:{responsive:true,maintainAspectRatio:false,cutout:'70%',animation:{duration:500},plugins:{legend:{position:'bottom',labels:{boxWidth:10}}}}
  });
}

function mrFormHtml(m){
  m=m||{};
  var meses=['Enero','Febrero','Marzo','Abril','Mayo','Junio',
    'Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'];
  var tiendas=uniq(STORE.tareas.map(function(t){return t.tienda;})
    .concat(STORE.auditorias.map(function(a){return a.tienda;})));
  var yoSesion=(_session&&(_session.nombre||_session.username))||'';
  /* En nuevo: pre-llena con la sesión. En editar: conserva el auditor guardado. */
  var auditorVal=(m.auditor!==undefined&&m.auditor!==null&&m.auditor!=='')?m.auditor:yoSesion;
  var auditoresLista=uniq((getAuditores()||[]).concat(yoSesion?[yoSesion]:[]));
  return '<div class="form-grid">'+
    '<div class="form-field full"><label>Realizado por (auditor) *</label>'+
      '<input id="mrf-auditor" list="mrf-auditor-list" value="'+esc(auditorVal)+'" placeholder="Nombre del auditor" autocomplete="off">'+
      '<datalist id="mrf-auditor-list">'+auditoresLista.map(function(a){return'<option value="'+esc(a)+'">';}).join('')+'</datalist>'+
      '<div style="font-size:10px;color:var(--muted);margin-top:3px">Se rellena con tu usuario; puedes cambiarlo si la merma la revisó otra persona.</div></div>'+
    '<div class="form-field"><label>Tienda *</label>'+
      '<input id="mrf-tienda" list="mrf-tienda-list" value="'+esc(m.tienda||'')+'" placeholder="Nombre de tienda">'+
      '<datalist id="mrf-tienda-list">'+tiendas.map(function(t){return'<option value="'+esc(t)+'">';}).join('')+'</datalist></div>'+
    '<div class="form-field"><label>Local / Foránea *</label>'+
      '<select id="mrf-lf">'+
        '<option value="Local"'+(m.localForanea==='Local'?' selected':'')+'>Local</option>'+
        '<option value="Foranea"'+(m.localForanea&&m.localForanea!=='Local'?' selected':'')+'>Foránea</option>'+
      '</select></div>'+
    '<div class="form-field"><label>Mes</label>'+
      '<select id="mrf-mes">'+meses.map(function(ms){return'<option'+(norm(m.mes||'')===norm(ms)?' selected':'')+'>'+ms+'</option>';}).join('')+'</select></div>'+
    '<div class="form-field"><label>Fecha autorización *</label><input type="date" id="mrf-aut" value="'+(m.fechaAut||'')+'" oninput="mrfAutoCalc()"></div>'+
    '<div class="form-field"><label>Fecha validación *</label><input type="date" id="mrf-val" value="'+(m.fechaVal||'')+'" oninput="mrfAutoCalc()"></div>'+
    '<div class="form-field" style="background:var(--soft);border-radius:var(--radius-sm);padding:10px 12px;border:1px solid var(--border)">'+
      '<div style="font-size:11px;font-weight:700;color:var(--muted);text-transform:uppercase;letter-spacing:.04em;margin-bottom:4px">Días / Condición</div>'+
      '<div id="mrf-preview" style="font-size:20px;font-weight:800;color:var(--blue)">'+
        (m.dias!=null?m.dias+' día(s) — '+m.condicion:'—')+'</div></div>'+
  '</div><div id="mrf-err" style="font-size:12px;color:var(--red);min-height:16px;margin-top:8px"></div>';
}
function mrfAutoCalc(){
  var a=document.getElementById('mrf-aut').value;
  var v=document.getElementById('mrf-val').value;
  var dias=diasEntre(a,v);
  var prev=document.getElementById('mrf-preview');
  if(dias===null||!a||!v){prev.textContent='—';prev.style.color='var(--blue)';return;}
  var cond=condicionAjuste(dias);
  prev.textContent=dias+' día(s) — '+cond;
  prev.style.color=cond==='A tiempo'?'#16a34a':'#dc2626';
}

function openNuevaMerma(){
  if(!_session||!['admin','admin_auditor','auditor'].includes(_session.rol)){toast('⚠ Sin permisos');return;}
  openModal('➕ Nueva merma',mrFormHtml(),[
    {label:'Cancelar',cls:'btn-ghost',fn:closeModal},
    {label:'Guardar',cls:'btn-orange',fn:guardarNuevaMerma}
  ]);
}

async function guardarNuevaMerma(){
  var errEl=document.getElementById('mrf-err');errEl.textContent='';
  var auditorMr=(document.getElementById('mrf-auditor').value||'').trim()
                ||(_session&&(_session.nombre||_session.username))||'Desconocido';
  var tienda=document.getElementById('mrf-tienda').value.trim().toUpperCase();
  var lf=document.getElementById('mrf-lf').value;
  var mes=document.getElementById('mrf-mes').value;
  var aut=document.getElementById('mrf-aut').value;
  var val=document.getElementById('mrf-val').value;
  if(!tienda){errEl.textContent='La tienda es obligatoria';return;}
  if(!aut||!val){errEl.textContent='Ambas fechas son obligatorias';return;}
  var dias=diasEntre(aut,val);
  /* Año predominante = fecha de validación; si aún no existe, se usa fecha de autorización */
  var año=parseInt(((val||aut)||'').split('-')[0])||null;
  var _mrRaw={tienda,local_foranea:lf,mes:mes.toUpperCase(),año,
    fecha_autorizacion:aut,fecha_validacion:val,dias,condicion:condicionAjuste(dias),auditor:auditorMr||null};
  var row=await encObj(_mrRaw,FIELDS.mermas);
  var client=getSbClient();if(!client){toast('⚠ Sin Supabase');return;}
  closeModal();
  try{
    var r=await client.from('mermas').insert([row]).select();
    if(r.error){toast('⚠ '+r.error.message);return;}
    var ins=(r.data&&r.data[0])||{};
    /* Guardar en memoria los valores EN CLARO (no los cifrados que devuelve
       Supabase). Solo se toma el id del registro insertado. */
    MERMAS.push({id:ins.id,tienda:tienda,localForanea:lf,
      mes:mes.toUpperCase(),año:año,fechaAut:aut,
      fechaVal:val,dias:dias,condicion:condicionAjuste(dias),auditor:auditorMr||''});
    fillMrFilters();renderMermas();toast('☁️ Merma guardada');
  }catch(e){toast('⚠ '+e.message);}
}

function openEditMerma(id){
  if(!_session||!['admin','admin_auditor','auditor'].includes(_session.rol)){toast('⚠ Sin permisos');return;}
  var m=MERMAS.find(function(x){return x.id===id;});if(!m)return;
  openModal('✎ Editar merma — '+esc(m.tienda),mrFormHtml(m),[
    {label:'Cancelar',cls:'btn-ghost',fn:closeModal},
    {label:'Eliminar',cls:'btn-red',fn:function(){deleteMerma(id);}},
    {label:'Guardar',cls:'btn-orange',fn:function(){saveEditMerma(id);}}
  ]);
}

async function saveEditMerma(id){
  function dv(el){var e=document.getElementById(el);return e?e.value.trim():'';}
  var tienda=dv('mrf-tienda').toUpperCase();
  var lf=dv('mrf-lf');
  var mes=dv('mrf-mes').toUpperCase();
  var aut=dv('mrf-aut');
  var val=dv('mrf-val');
  var auditor=dv('mrf-auditor')||(_session&&(_session.nombre||_session.username))||'Desconocido';
  var dias=diasEntre(aut,val);
  /* Año predominante = fecha de validación; si aún no existe, se usa fecha de autorización */
  var año=parseInt(((val||aut)||'').split('-')[0])||null;
  var _mrUpdRaw={tienda,local_foranea:lf,mes,año,fecha_autorizacion:aut,
    fecha_validacion:val,dias,condicion:condicionAjuste(dias),
    auditor:auditor};
  var row=await encObj(_mrUpdRaw,FIELDS.mermas);
  var client=getSbClient();if(!client)return;
  closeModal();
  try{
    var r=await client.from('mermas').update(row).eq('id',id);
    if(r.error){toast('⚠ '+r.error.message);return;}
    var idx=MERMAS.findIndex(function(x){return x.id===id;});
    if(idx>=0)MERMAS[idx]=Object.assign({},MERMAS[idx],{tienda,localForanea:lf,mes,año,
      fechaAut:aut,fechaVal:val,dias,condicion:condicionAjuste(dias),auditor:auditor});
    fillMrFilters();renderMermas();toast('☁️ Merma actualizada');
  }catch(e){toast('⚠ '+e.message);}
}

async function deleteMerma(id){
  if(!confirm('¿Eliminar esta merma?'))return;
  var client=getSbClient();if(!client)return;
  closeModal();
  try{
    var r=await client.from('mermas').delete().eq('id',id);
    if(r.error){toast('⚠ '+r.error.message);return;}
    MERMAS=MERMAS.filter(function(x){return x.id!==id;});
    fillMrFilters();renderMermas();toast('☁️ Merma eliminada');
  }catch(e){toast('⚠ '+e.message);}
}

/* ── Importar Excel Mermas ── */
function importMermasExcel(files){
  if(!files||!files.length)return;
  if(!_session||!['admin','admin_auditor','auditor'].includes(_session.rol)){toast('⚠ Sin permisos');return;}
  var file=files[0];
  document.getElementById('mr-file-input').value='';
  var reader=new FileReader();
  reader.onload=function(e){
    try{
      /* cellDates:false para recibir seriales numéricos y convertirlos manualmente */
      var wb=XLSX.read(e.target.result,{type:'array',cellDates:false});
      var _pk=pickSheet(wb,['mermas'],['tienda','merma','autoriz','validacion']);var ws=_pk.ws,raw=_pk.raw;
      if(raw.length<2){toast('⚠ Excel vacío');return;}
      /* Normalizar cabeceras: minúsculas + sin tildes + sin espacios extra */
      function normH(s){return String(s||'').toLowerCase().trim().normalize('NFD').replace(/[\u0300-\u036f]/g,'');}
      var H=raw[0].map(normH);
      /* Buscar columna por lista de variantes, con coincidencia parcial como fallback */
      function fc(keys){
        var nk=keys.map(normH);
        for(var ki=0;ki<nk.length;ki++)for(var hi=0;hi<H.length;hi++)if(H[hi]===nk[ki])return hi;
        for(var ki=0;ki<nk.length;ki++)for(var hi=0;hi<H.length;hi++)if(H[hi].includes(nk[ki])||nk[ki].includes(H[hi]))return hi;
        return -1;
      }
      var iT=fc(['tienda','sucursal']);
      var iLF=fc(['local_foranea','local/foranea','local foranea','localforanea','tipo']);
      var iM=fc(['mes','month']);
      /* "Autoriacion" es error tipográfico en el Excel — aceptar sin z */
      var iA=fc(['autorizacion','autoriacion','autoriacion','fecha autorizacion','fecha_autorizacion']);
      var iV=fc(['validacion','fecha validacion','fecha_validacion']);
      var iAud=fc(['auditor','realizado por','realizadopor','responsable']);
      if(iT<0||iA<0||iV<0){
        toast('⚠ Columnas detectadas: '+H.join(' | '));
        return;
      }
      var yoSesion=(_session&&(_session.nombre||_session.username))||'';
      /* Convertir serial de Excel (número) o string a yyyy-mm-dd */
      function toISOfecha(v){
        if(!v&&v!==0)return null;
        var n=Number(v);
        if(!isNaN(n)&&n>30000&&n<80000){
          /* Excel base: 25569 días desde epoch Unix = 1900-01-01 con bug de bisiesto */
          var ms=(n-25569)*86400*1000;
          var d=new Date(ms);
          return d.getUTCFullYear()+'-'+String(d.getUTCMonth()+1).padStart(2,'0')+'-'+String(d.getUTCDate()).padStart(2,'0');
        }
        if(v instanceof Date){
          return v.getFullYear()+'-'+String(v.getMonth()+1).padStart(2,'0')+'-'+String(v.getDate()).padStart(2,'0');
        }
        var s=String(v).trim();
        var m=s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
        if(m)return m[3]+'-'+m[2].padStart(2,'0')+'-'+m[1].padStart(2,'0');
        if(/^\d{4}-\d{2}-\d{2}$/.test(s))return s;
        var dt=new Date(s);
        if(!isNaN(dt))return dt.getFullYear()+'-'+String(dt.getMonth()+1).padStart(2,'0')+'-'+String(dt.getDate()).padStart(2,'0');
        return null;
      }
      var mNames=['ENERO','FEBRERO','MARZO','ABRIL','MAYO','JUNIO',
        'JULIO','AGOSTO','SEPTIEMBRE','OCTUBRE','NOVIEMBRE','DICIEMBRE'];
      var rows=[];
      for(var i=1;i<raw.length;i++){
        var r=raw[i];
        var tienda=String(r[iT]||'').trim().toUpperCase();
        if(!tienda)continue;
        var fechaA=toISOfecha(r[iA]);
        var fechaV=toISOfecha(r[iV]);
        /* Basta la fecha de autorización: una merma del mes en curso puede seguir
           sin validar. Antes se descartaba la fila y el mes no aparecía. */
        if(!fechaA)continue;
        var lf=iLF>=0?String(r[iLF]||'').trim():'Local';
        if(!lf)lf='Local';
        /* Mes/año SIEMPRE derivados desde la FECHA DE VALIDACIÓN (predominante);
           si la fila aún no tiene fecha de validación (pendiente), se usa fecha
           de autorización como respaldo. Se ignora la columna Mes del Excel por
           la misma razón que en Ajustes: formatos inconsistentes. */
        var _fechaBaseMrImp=fechaV||fechaA;
        var mes=_fechaBaseMrImp?(mNames[parseInt(_fechaBaseMrImp.split('-')[1])-1]||''):'';
        if(!mes&&iM>=0){
          mes=String(r[iM]||'').trim().toUpperCase();
          if(mes&&mes.includes(','))mes=mes.split(',')[0].trim().toUpperCase();
        }
        var año=parseInt((_fechaBaseMrImp||'').split('-')[0])||null;
        var dias=fechaV?diasEntre(fechaA,fechaV):null;
        var auditor=(iAud>=0?String(r[iAud]||'').trim():'')||yoSesion||'Desconocido';
        rows.push({tienda,local_foranea:lf,mes,año,fecha_autorizacion:fechaA,
          fecha_validacion:fechaV,dias,condicion:condicionAjuste(dias),auditor:auditor});
      }
      if(!rows.length){toast('⚠ Sin filas válidas. Revisa que las columnas de fecha tengan datos');return;}
      toast('⏳ Importando '+rows.length+' registros…');
      commitMermasExcel(rows);
    }catch(err){toast('⚠ Error: '+err.message);console.error(err);}
  };
  reader.readAsArrayBuffer(file);
}

async function commitMermasExcel(rows){
  var client=getSbClient();if(!client){toast('⚠ Sin Supabase');return;}
  try{
    /* Traer existentes para actualizar en vez de duplicar. Clave: tienda +
       fecha_autorizacion + fecha_validacion (las fechas no están cifradas). */
    var ex=[];
    try{var re=await client.from('mermas').select('id,fecha_autorizacion,fecha_validacion,tienda').limit(20000);if(!re.error)ex=re.data||[];}catch(_e){}
    var exDec=await decArr(ex,FIELDS.mermas);
    var _k=function(t,fa,fv){return norm(t)+'|'+String(fa||'')+'|'+String(fv||'');};
    var exMap={};
    exDec.forEach(function(x){exMap[_k(x.tienda,x.fecha_autorizacion,x.fecha_validacion)]=x.id;});
    var seen={},dedup=[];
    for(var z=rows.length-1;z>=0;z--){var kk=_k(rows[z].tienda,rows[z].fecha_autorizacion,rows[z].fecha_validacion);if(seen[kk])continue;seen[kk]=1;dedup.unshift(rows[z]);}
    var ins=0,upd=0,err=null;
    for(var i=0;i<dedup.length&&!err;i++){
      var raw=dedup[i];
      var row=await encObj(raw,FIELDS.mermas);   /* ← CIFRAR antes de guardar */
      var k=_k(raw.tienda,raw.fecha_autorizacion,raw.fecha_validacion);
      if(exMap[k]){
        var ru=await client.from('mermas').update(row).eq('id',exMap[k]);
        if(ru.error)err=ru.error.message;else upd++;
      }else{
        var ri=await client.from('mermas').insert([row]).select('id');
        if(ri.error)err=ri.error.message;else{ins++;if(ri.data&&ri.data[0])exMap[k]=ri.data[0].id;}
      }
    }
    if(err){toast('⚠ '+err);return;}
    toast('☁️ '+ins+' nuevas · '+upd+' actualizadas');
    _mrPerManual=false; /* tras importar, volver al mes en curso */
    await loadMermas();
  }catch(e){toast('⚠ '+e.message);}
}
/* ════════════════════════════════════════════════════════════════════
   MÓDULO AUDITORÍAS FINALIZADAS — Supabase
   Tabla: tareas_finalizadas (retención de 7 días desde fecha_finalizacion;
          ver purgarFinalizadasVencidas)
   Columnas: id, aud_key, tienda, mes, clase, pct_cumpl,
             total_tareas, resueltas, fecha_finalizacion, created_at
   Flujo: una auditoría sin tareas pendientes se registra aquí. Se filtra por
          el MES DE LA AUDITORÍA (campo mes) y por defecto muestra el mes en
          curso. Pasados 7 días desde su finalización, se elimina sola.
════════════════════════════════════════════════════════════════════ */
var FINALIZADAS=[];
var _finMesManual=false; /* true cuando el usuario elige mes/año a mano en Finalizadas */
var _finLoaded=false;

async function loadFinalizadas(){
  var client=getSbClient();
  if(!client){toast('⚠ Sin Supabase');return;}
  try{
    var r=await client.from('tareas_finalizadas').select('*')
      .order('fecha_finalizacion',{ascending:false}).limit(5000);
    if(r.error){toast('⚠ '+r.error.message);return;}
    var _finRaw=r.data||[];
    var _finDec=await decArr(_finRaw,FIELDS.tareas_finalizadas);

    /* Helper: recuperar nombre legible de tienda cruzando con auditorías y tareas */
    function resolverTienda(f){
      var t=f.tienda||'';
      /* Si parece cifrado (contiene : y caracteres base64) búscala en STORE */
      var parece_cifrado=/^[A-Za-z0-9+/=]+:[A-Za-z0-9+/=]+$/.test(t)&&t.length>40;
      if(!parece_cifrado)return t;
      /* Buscar en auditorías por aud_key */
      if(f.aud_key){
        var parts=f.aud_key.split('|');
        var tiendaNorm=parts[0]||'';
        /* Cruzar con STORE.auditorias */
        var aud=STORE.auditorias.find(function(a){return norm(a.tienda||'')===tiendaNorm;});
        if(aud&&aud.tienda)return aud.tienda;
        /* Cruzar con STORE.tareas */
        var tarea=STORE.tareas.find(function(ta){return norm(ta.tienda||'')===tiendaNorm;});
        if(tarea&&tarea.tienda)return tarea.tienda;
        /* Último recurso: capitalizar la clave normalizada */
        if(tiendaNorm)return tiendaNorm.toUpperCase();
      }
      return t;
    }

    FINALIZADAS=_finDec.map(function(f){
      /* duracion_dias: calcular desde fechas reales de tareas cuando sea posible */
      var fechaIni=f.fecha_inicio;
      var fechaFinRec=f.fecha_finalizacion;
      /* Si fecha_inicio es null, buscarla en STORE.tareas */
      if((!fechaIni||!fechaFinRec)&&STORE.tareas.length>0){
        var tipoC=tipoTareaDeClase(f.clase||'');
        var td2=STORE.tareas.filter(function(t){
          if(norm(t.tienda)!==norm(f.tienda||''))return false;
          if(tipoC&&tipoNormLocal(t.tipoTarea)!==tipoC)return false;
          return true;
        });
        if(!fechaIni&&td2.length>0){
          td2.forEach(function(t){
            var fc=t.fechaCreacion?String(t.fechaCreacion).split('T')[0]:null;
            if(fc&&(!fechaIni||fc<fechaIni))fechaIni=fc;
          });
        }
        if(!fechaFinRec&&td2.length>0){
          td2.forEach(function(t){
            var fd=t.fechaCumpl?String(t.fechaCumpl).split('T')[0]:
                   (t.fechaTerm?String(t.fechaTerm).split('T')[0]:null);
            if(fd&&(!fechaFinRec||fd>fechaFinRec))fechaFinRec=fd;
          });
        }
      }
      var dur=(f.duracion_dias!=null)?f.duracion_dias:
        (fechaIni&&fechaFinRec?diasEntre(fechaIni,fechaFinRec):null);
      /* pct_cumpl: normalizar — puede llegar como 0.74 o como 74 */
      var pct=parseFloat(f.pct_cumpl)||0;
      if(pct>1)pct=pct/100;
      /* total_tareas y resueltas desde STORE.
         IMPORTANTE: se debe acotar por CENTRO además de tienda+clase — si no,
         una tienda con dos auditorías de la misma clase en meses distintos
         (una ya finalizada, otra todavía vigente con pendientes) mezcla las
         tareas de ambas al recalcular. Eso hacía que la finalizada antigua
         se "desarchivara" sola por pendientes que en realidad eran de la
         auditoría vigente, descuadrando Vigentes contra Finalizadas.
         Se busca primero la auditoría original (misma tienda+mes+clase) en
         STORE.auditorias para tomar su centro exacto, igual que hace
         registrarFinalizada() al archivar. Si no se encuentra (por ejemplo,
         la auditoría original ya no está en memoria), se cae al criterio
         anterior (tienda+clase) como respaldo. */
      var total=f.total_tareas||0;
      var res=f.resueltas||0;
      if(STORE.tareas.length>0&&f.aud_key){
        var tipoClase=tipoTareaDeClase(f.clase||'');
        var audOriginal=(STORE.auditorias||[]).find(function(a){return audKeyMes(a)===audKeyMes(f);});
        var centroOriginal=audOriginal?norm(audOriginal.centro||''):'';
        var td=STORE.tareas.filter(function(t){
          if(norm(t.tienda)!==norm(f.tienda||''))return false;
          if(tipoClase&&tipoNormLocal(t.tipoTarea)!==tipoClase)return false;
          if(centroOriginal&&norm(t.centro||'')!==centroOriginal)return false;
          return true;
        });
        if(td.length>0){total=td.length;res=td.filter(function(t){return!esPendiente(t);}).length;}
      }
      return Object.assign({},f,{
        tienda: resolverTienda(f),
        mes: (f.mes&&f.mes.length<20)?f.mes:(f.aud_key?(f.aud_key.split('|')[1]||'').toUpperCase():f.mes),
        clase: (f.clase&&f.clase.length<60)?f.clase:(f.aud_key?f.aud_key.split('|')[2]||f.clase:f.clase),
        duracion_dias:dur,pct_cumpl:pct,total_tareas:total,resueltas:res,
        completadoPor:f.completado_por||''});
    });
    FINALIZADAS=FINALIZADAS.filter(function(f){return f._pending||tiendaVisible(f.tienda);});
    /* Las auditorías finalizadas SE CONSERVAN: son el histórico del que se
       alimentan los reportes. Antes se borraban a los 7 días.
       Solo se devuelven a Auditorías las que quedaron archivadas incompletas. */
    deduplicarFinalizadas();
    reconciliarFinalizadas();
    _finLoaded=true;
    fillFinFilters();renderFinalizadas();
    if(typeof actualizarStrip==='function')actualizarStrip(); /* re-sincroniza Vigentes/Finalizadas */
    toast('\u2713 '+FINALIZADAS.length+' auditoría(s) finalizada(s)');
  }catch(e){toast('\u26a0 '+e.message);}
}

/* Las auditorías finalizadas se conservan en este módulo 7 días desde su
   fecha de finalización; pasado ese plazo se eliminan solas, tanto de
   Supabase como de la lista en pantalla, para no acumular historial
   indefinidamente. Se corre cada vez que se recargan los datos. */
/* Días que una auditoría debe permanecer al 100% resuelta antes de pasar a
   Finalizadas. Evita moverla por una resolución momentánea. */
var DIAS_PARA_FINALIZAR=3;

/* Clave única de una auditoría (misma fórmula en todo el dashboard) */
function audKeyDe(a){
  return [norm(a.tienda||''),norm(a.mes||''),norm(a.clase||''),a.fecha||''].join('|');
}
/* Clave SIN fecha: tienda|mes|clase. El mes ya conserva el criterio mensual,
   y así una auditoría viva casa con su registro archivado aunque la fecha
   del renglón difiera (esa era la causa del duplicado KN EXPRESS). */
function audKeyMes(a){
  if(a&&a.aud_key)return String(a.aud_key).split('|').slice(0,3).join('|');
  return [norm(a.tienda||''),norm(a.mes||''),norm(a.clase||'')].join('|');
}
/* Una finalizada solo "tapa" a su auditoría si de verdad está al 100%. */
function finCompleta(f){
  if(f._pending)return true;
  var tot=+f.total_tareas||0,res=+f.resueltas||0;
  return tot>0&&res>=tot;
}
/* ¿Esta auditoría ya está archivada en Finalizadas?
   Se usa para NO contarla otra vez entre las vigentes (evita duplicar). */
function estaFinalizada(a){
  var ak=audKeyMes(a);
  return (FINALIZADAS||[]).some(function(f){return finCompleta(f)&&audKeyMes(f)===ak;});
}

/* Deduplicar: la tabla puede llegar con más de un renglón para la misma
   auditoría (misma tienda+mes+clase) — típicamente por doble inserción
   cuando dos sesiones abiertas al mismo tiempo disparan registrarFinalizada()
   casi al mismo instante y ambas pasan el chequeo "existe" antes de que
   cualquiera termine de guardar en Supabase. Sin esto, Finalizadas mostraba
   la misma auditoría dos (o más) veces — y esos duplicados se sumaban en
   los reportes que alimentan de este módulo (tareas/resueltas contadas de
   más). Se conserva un solo renglón por auditoría (el de fecha_finalizacion
   más reciente) y el resto se elimina también de Supabase. */
function deduplicarFinalizadas(){
  var vistos={},dupIds=[];
  var ordenadas=FINALIZADAS.slice().sort(function(x,y){
    return String(y.fecha_finalizacion||'').localeCompare(String(x.fecha_finalizacion||''));
  });
  ordenadas.forEach(function(f){
    var k=audKeyMes(f);
    if(vistos[k]){ if(f.id&&!f._pending)dupIds.push(f.id); return; }
    vistos[k]=true;
  });
  if(!dupIds.length)return;
  FINALIZADAS=FINALIZADAS.filter(function(f){return dupIds.indexOf(f.id)===-1;});
  var client=getSbClient();
  if(!client)return;
  client.from('tareas_finalizadas').delete().in('id',dupIds)
    .then(function(r){if(r.error)console.warn('deduplicar finalizadas:',r.error.message);})
    .catch(function(e){console.warn('deduplicar finalizadas:',e.message);});
}

/* Reconciliación: toda finalizada con resueltas < total_tareas nunca debió
   archivarse. Se saca de Finalizadas (y de Supabase) para que vuelva a
   Auditorías y se re-archive sola cuando llegue al 100% + DIAS_PARA_FINALIZAR. */
function reconciliarFinalizadas(){
  var malas=FINALIZADAS.filter(function(f){return !f._pending&&!finCompleta(f);});
  if(!malas.length)return;
  var ids=malas.map(function(f){return f.id;}).filter(Boolean);
  FINALIZADAS=FINALIZADAS.filter(function(f){return finCompleta(f);});
  var client=getSbClient();
  if(!client||!ids.length)return;
  client.from('tareas_finalizadas').delete().in('id',ids)
    .then(function(r){if(r.error)console.warn('reconciliar finalizadas:',r.error.message);})
    .catch(function(e){console.warn('reconciliar finalizadas:',e.message);});
}

var RETENCION_FINALIZADAS_DIAS=7;
function purgarFinalizadasVencidas(){
  var hoy=new Date();hoy.setHours(0,0,0,0);
  var vencidas=[],vigentes=[];
  FINALIZADAS.forEach(function(f){
    if(f._pending){vigentes.push(f);return;} /* aún no confirmada en Supabase: no tocar */
    var fecha=f.fecha_finalizacion?String(f.fecha_finalizacion).split('T')[0]:null;
    var fh=fecha?fromISO(fecha):null;
    if(!fh){vigentes.push(f);return;} /* sin fecha: no se puede juzgar, se conserva */
    var dias=Math.floor((hoy-fh)/(1000*60*60*24));
    if(dias>RETENCION_FINALIZADAS_DIAS)vencidas.push(f);else vigentes.push(f);
  });
  if(!vencidas.length)return;
  FINALIZADAS=vigentes;
  var client=getSbClient();
  if(!client)return;
  var ids=vencidas.map(function(f){return f.id;}).filter(Boolean);
  if(!ids.length)return;
  client.from('tareas_finalizadas').delete().in('id',ids)
    .then(function(r){if(r.error)console.warn('purga finalizadas:',r.error.message);})
    .catch(function(e){console.warn('purga finalizadas:',e.message);});
}

function fillFinFilters(){
  var MORD=['ENERO','FEBRERO','MARZO','ABRIL','MAYO','JUNIO',
    'JULIO','AGOSTO','SEPTIEMBRE','OCTUBRE','NOVIEMBRE','DICIEMBRE'];
  var mesSet=new Set(FINALIZADAS.map(function(f){return(f.mes||'').toUpperCase();}).filter(Boolean));
  var meses=MORD.filter(function(m){return mesSet.has(m);});
  var añoSet=[...new Set(FINALIZADAS.map(function(f){
    return f.fecha_finalizacion?f.fecha_finalizacion.split('-')[0]:'';
  }).filter(Boolean))].sort().reverse();
  var tiendas=limpiarOpciones([...new Set(FINALIZADAS.map(function(f){return f.tienda;}).filter(Boolean))].sort());
  var sm=document.getElementById('fin-f-mes');
  var sa=document.getElementById('fin-f-año');
  var st=document.getElementById('fin-f-tienda');
  if(!sm)return;
  var vm=sm.value,va=sa.value,vt=st.value;
  sm.innerHTML='<option value="ALL">Todos</option>'+meses.map(function(m){return'<option value="'+m+'">'+m+'</option>';}).join('');
  sa.innerHTML='<option value="ALL">Todos</option>'+añoSet.map(function(y){return'<option value="'+y+'">'+y+'</option>';}).join('');
  st.innerHTML='<option value="ALL">Todas</option>'+limpiarOpciones(tiendas).map(function(t){return'<option>'+t+'</option>';}).join('');
  /* Mes de auditoría y año por defecto = mes/año en curso si tienen datos; si no,
     el más reciente con datos. Se respeta la elección manual (_finMesManual). */
  var mesHoyF=MORD[new Date().getMonth()];
  var añoHoyF=String(new Date().getFullYear());
  var tgtF=mesSet.has(mesHoyF)?mesHoyF:(meses.length?meses[meses.length-1]:'ALL');
  var tgtAñoF=añoSet.indexOf(añoHoyF)>=0?añoHoyF:(añoSet.length?añoSet[0]:'ALL');
  if(_finMesManual){
    if([...sm.options].some(function(o){return o.value===vm;}))sm.value=vm;
    if([...sa.options].some(function(o){return o.value===va;}))sa.value=va;
  }else{
    sm.value=[...sm.options].some(function(o){return o.value===tgtF;})?tgtF:'ALL';
    sa.value=[...sa.options].some(function(o){return o.value===tgtAñoF;})?tgtAñoF:'ALL';
  }
  if([...st.options].some(function(o){return o.value===vt;}))st.value=vt;
}
function finMesChanged(){_finMesManual=true;renderFinalizadas();}

function filteredFinalizadas(){
  var mes=document.getElementById('fin-f-mes').value||'ALL';
  var año=document.getElementById('fin-f-año').value||'ALL';
  var tienda=document.getElementById('fin-f-tienda').value||'ALL';
  return FINALIZADAS.filter(function(f){
    if(mes!=='ALL'&&(f.mes||'').toUpperCase()!==mes)return false;
    if(año!=='ALL'){var y=f.fecha_finalizacion?f.fecha_finalizacion.split('-')[0]:'';if(y!==año)return false;}
    if(tienda!=='ALL'&&norm(f.tienda)!==norm(tienda))return false;
    return true;
  });
}

function renderFinalizadas(){
  var arr=filteredFinalizadas();
  var isAdmin=puedeModificarModulo('finalizadas');
  document.getElementById('fin-count').textContent=arr.length+' finalizada(s)';
  var th=document.getElementById('fin-edit-th');if(th)th.style.display=isAdmin?'':'none';
  var tbody=document.getElementById('fin-tbody');
  var nCols=isAdmin?9:8;

  if(!arr.length){
    tbody.innerHTML='<tr><td colspan="'+nCols+'" style="text-align:center;color:var(--muted);padding:24px">Sin registros finalizados.</td></tr>';
    return;
  }

  function getDur(f){
    if(f.duracion_dias!=null)return f.duracion_dias;
    var fi=f.fecha_inicio;var ff=f.fecha_finalizacion;
    if(!fi||!ff){
      var tipoC=tipoTareaDeClase(f.clase||'');
      var td=STORE.tareas.filter(function(t){
        if(norm(t.tienda)!==norm(f.tienda||''))return false;
        if(tipoC&&tipoNormLocal(t.tipoTarea)!==tipoC)return false;
        return true;
      });
      if(!fi&&td.length>0)td.forEach(function(t){
        var fc=t.fechaCreacion?String(t.fechaCreacion).split('T')[0]:null;
        if(fc&&(!fi||fc<fi))fi=fc;
      });
      if(!ff&&td.length>0)td.forEach(function(t){
        var fd=t.fechaCumpl?String(t.fechaCumpl).split('T')[0]:(t.fechaTerm?String(t.fechaTerm).split('T')[0]:null);
        if(fd&&(!ff||fd>ff))ff=fd;
      });
    }
    return (fi&&ff)?diasEntre(fi,ff):null;
  }

  function claseGrupo(f){
    var c=norm(f.clase||'');
    var _isDark=document.documentElement.getAttribute('data-theme')==='dark';
    if(c.includes('colaboracion')||c.includes('colab'))return{key:'colab',label:'COLABORACIÓN',color:'#2563eb',bg:_isDark?'rgba(67,24,255,.12)':'#eff6ff'};
    if(c.includes('orden')||c.includes('limpieza'))return{key:'ol',label:'ORDEN Y LIMPIEZA',color:'#16a34a',bg:_isDark?'rgba(1,181,116,.10)':'#f0fdf4'};
    if(c.includes('cartera'))return{key:'cartera',label:'CARTERA',color:'#7c3aed',bg:_isDark?'rgba(159,122,234,.12)':'#f5f3ff'};
    return{key:'otro',label:(f.clase||'OTRO').toUpperCase(),color:'#7c8696',bg:_isDark?'rgba(255,255,255,.04)':'#f8fafc'};
  }

  /* Agrupar por tipo */
  var grupos={};
  var orden=['colab','ol','cartera','otro'];
  arr.forEach(function(f){
    var g=claseGrupo(f);
    if(!grupos[g.key])grupos[g.key]={meta:g,filas:[]};
    grupos[g.key].filas.push(f);
  });

  function renderGrupo(g){
    var filas=g.filas;
    var meta=g.meta;
    var arrConPct=filas.filter(function(f){return f.pct_cumpl>0;});
    var promPct=arrConPct.length?Math.round(arrConPct.reduce(function(s,f){return s+(f.pct_cumpl||0);},0)/arrConPct.length*100):null;
    var arrConDur=filas.filter(function(f){var d=getDur(f);return d!=null&&d>=0;});
    var promDias=arrConDur.length?Math.round(arrConDur.reduce(function(s,f){return s+getDur(f);},0)/arrConDur.length):null;
    var promSems=promDias!=null?Math.round(promDias/7*10)/10:null;
    var totalTareasSum=filas.reduce(function(s,f){return s+(parseInt(f.total_tareas)||0);},0);
    var totalResSum=filas.reduce(function(s,f){return s+(parseInt(f.resueltas)||0);},0);

    var rows=filas.map(function(f,i){
      var pct=(f.pct_cumpl>0)?Math.round(f.pct_cumpl*100)+'%':'—';
      var dur=getDur(f);
      var diasStr=dur!=null?dur+'d':'—';
      var semsStr=dur!=null?(Math.round(dur/7*10)/10)+' sem':'—';
      var editTd=isAdmin?'<td style="text-align:center"><button class="icon-btn" data-fid="'+esc(f.id||'')+'" onclick="finDeleteClick(this)" title="Eliminar">✕</button></td>':'';
      var fechaCell=isAdmin?
        '<td style="text-align:center"><input class="seg-date-input" type="date" value="'+(f.fecha_finalizacion||'')+'" '+
          'data-fid="'+esc(f.id||'')+'" onchange="updateFinFecha(this)" style="width:110px"></td>':
        '<td style="text-align:center">'+fmtFecha(f.fecha_finalizacion)+'</td>';
      /* Tienda y % Cumpl. editables inline, igual que en Actividades:
         inputs con onchange que actualizan memoria y persisten en Supabase
         (updateCampoFinInline), sin abrir ningún modal aparte. */
      var tiendaCell=isAdmin?
        '<td><input value="'+esc(f.tienda||'')+'" data-fid="'+esc(f.id||'')+'" data-campo="tienda" '+
          'onchange="updateCampoFinInline(this)" style="'+inlineFieldStyle()+';width:150px;font-weight:700;color:'+meta.color+'"></td>':
        '<td><b style="color:'+meta.color+'">'+esc(f.tienda||'—')+'</b></td>';
      var pctCell=isAdmin?
        '<td style="text-align:center"><input type="number" min="0" max="100" step="1" value="'+((f.pct_cumpl>0)?Math.round(f.pct_cumpl*100):'')+'" '+
          'data-fid="'+esc(f.id||'')+'" data-campo="pct_cumpl" onchange="updateCampoFinInline(this)" '+
          'style="'+inlineFieldStyle()+';width:58px;text-align:center;font-weight:700;color:var(--k-greenok)"></td>':
        '<td style="text-align:center;font-weight:700;color:var(--k-greenok)">'+pct+'</td>';
      return '<tr style="background:'+meta.bg+'">'+
        '<td style="color:var(--muted);font-size:11px">'+(i+1)+'</td>'+
        tiendaCell+
        pctCell+
        '<td style="text-align:center;font-weight:700">'+(parseInt(f.total_tareas)||0)+'</td>'+
        '<td style="text-align:center;color:var(--k-greenok);font-weight:700">'+(parseInt(f.resueltas)||0)+'</td>'+
        fechaCell+
        '<td style="text-align:center;color:var(--muted)">'+diasStr+'</td>'+
        '<td style="text-align:center;color:var(--muted)">'+semsStr+'</td>'+
        editTd+'</tr>';
    }).join('');

    var resumen='<tr style="background:'+meta.color+'22;font-weight:800;border-top:2px solid '+meta.color+'">'+
      '<td colspan="2" style="color:'+meta.color+';padding:8px 12px">Subtotal '+meta.label+'</td>'+
      '<td style="text-align:center;color:'+meta.color+'">'+(promPct!=null?promPct+'%':'N/A')+'</td>'+
      '<td style="text-align:center;color:'+meta.color+'">'+totalTareasSum+'</td>'+
      '<td style="text-align:center;color:'+meta.color+'">'+totalResSum+'</td>'+
      '<td></td>'+
      '<td style="text-align:center;color:'+meta.color+'">'+(promDias!=null?promDias+' d':'—')+'</td>'+
      '<td style="text-align:center;color:'+meta.color+'">'+(promSems!=null?promSems+' sem':'—')+'</td>'+
      (isAdmin?'<td></td>':'')+'</tr>';

    return '<tr><td colspan="'+nCols+'" style="padding:10px 12px 4px;font-weight:800;font-size:12px;color:'+meta.color+';text-transform:uppercase;letter-spacing:.04em;border-top:2px solid '+meta.color+';background:var(--soft)">'+
      '● '+meta.label+' ('+filas.length+')</td></tr>'+rows+resumen;
  }

  /* Header común (theme-aware) */
  var headerRow='<tr style="background:var(--soft);color:var(--text);font-size:11px;font-weight:700;border-bottom:2px solid var(--border)">'+
    '<th style="padding:8px 10px">#</th><th>Tienda</th>'+
    '<th class="c">% Cumpl.</th><th class="c">Tareas</th>'+
    '<th class="c">Resueltas</th><th class="c">Fecha Fin</th>'+
    '<th class="c">Días</th><th class="c">Semanas</th>'+
    (isAdmin?'<th class="c">✕</th>':'')+'</tr>';

  var html=orden.map(function(k){
    return grupos[k]?renderGrupo(grupos[k]):'';
  }).join('');

  tbody.innerHTML=headerRow+html;
}

function finDeleteClick(el){
  var id=el.getAttribute('data-fid');
  if(id)deleteFinalizada(id);
}

async function updateFinFecha(el){
  if(!_session||!['admin','admin_auditor'].includes(_session.rol))return;
  var id=el.getAttribute('data-fid');
  var newDate=el.value;
  if(!id||!newDate)return;
  /* Actualizar en cache local */
  var rec=FINALIZADAS.find(function(f){return f.id===id;});
  if(rec)rec.fecha_finalizacion=newDate;
  var client=getSbClient();if(!client)return;
  try{
    var r=await client.from('tareas_finalizadas').update({fecha_finalizacion:newDate}).eq('id',id);
    if(r.error)toast('⚠ '+r.error.message);
    else toast('✓ Fecha actualizada');
  }catch(e){toast('⚠ '+e.message);}
}

/* Edición rápida de un campo de Finalizadas directamente desde la tabla,
   igual que updateCampoActividadInline en el módulo de Actividades: actualiza
   en memoria (FINALIZADAS, de donde también leen los reportes descargables),
   cifra el campo si corresponde y persiste en Supabase. */
var _CAMPO_FIN_COL={tienda:'tienda',pct_cumpl:'pct_cumpl'};
async function updateCampoFinInline(el){
  if(!_session||!['admin','admin_auditor'].includes(_session.rol)){toast('⚠ Solo administradores');renderFinalizadas();return;}
  var id=el.getAttribute('data-fid');
  var campo=el.getAttribute('data-campo');
  var col=_CAMPO_FIN_COL[campo];
  if(!id||!col)return;
  var f=FINALIZADAS.find(function(x){return String(x.id)===String(id);});
  if(!f)return;
  var anterior=f[campo];
  var valorSb=el.value;
  if(campo==='pct_cumpl'){
    var n=parseFloat(el.value);
    if(isNaN(n)){f[campo]=anterior;renderFinalizadas();return;}
    valorSb=Math.max(0,Math.min(100,n))/100;
    f.pct_cumpl=valorSb;
  }else{
    f[campo]=el.value||null;
    valorSb=el.value||null;
  }
  try{
    var client=getSbClient();
    if(!client){toast('⚠ Sin conexión a Supabase');f[campo]=anterior;renderFinalizadas();return;}
    var payload={};payload[col]=valorSb;
    var row=(FIELDS.tareas_finalizadas.indexOf(col)>=0)?await encObj(payload,FIELDS.tareas_finalizadas):payload;
    var ru=await client.from('tareas_finalizadas').update(row).eq('id',id);
    if(ru.error){toast('⚠ '+ru.error.message);f[campo]=anterior;renderFinalizadas();return;}
    toast('☁️ Actualizado');
    renderFinalizadas();
  }catch(e){toast('⚠ Error: '+e.message);f[campo]=anterior;renderFinalizadas();}
}

/* ── Registrar auditoría como finalizada ──
   fecha_inicio  = fechaCreacion más antigua de las tareas
   fecha_finalizacion = fechaCumpl o fechaTerm más reciente
   duracion_dias = fecha_finalizacion - fecha_inicio */
var _finRegistrandoEnVuelo={};
async function registrarFinalizada(a){
  var client=getSbClient();if(!client)return;
  var ak=[norm(a.tienda||""),norm(a.mes||""),norm(a.clase||""),a.fecha||""].join("|");
  var akm=audKeyMes(a);
  var existe=FINALIZADAS.find(function(f){return audKeyMes(f)===akm;});
  if(existe)return;
  /* Evita el duplicado en origen: si ya hay un registro en curso para esta
     misma auditoría (otra llamada disparada en el mismo ciclo, antes de que
     la primera termine su INSERT y aparezca en FINALIZADAS), se aborta aquí
     en vez de dejar que ambas lleguen a insertar la misma fila dos veces. */
  if(_finRegistrandoEnVuelo[akm])return;
  _finRegistrandoEnVuelo[akm]=true;
  FINALIZADAS.unshift({aud_key:ak,_pending:true});

  var tipoClase=tipoTareaDeClase(a.clase);
  var tareasDeAud=STORE.tareas.filter(function(t){
    if(norm(t.tienda)!==norm(a.tienda||''))return false;
    if(norm(t.centro)!==norm(a.centro||''))return false;
    if(tipoClase&&tipoNormLocal(t.tipoTarea)!==tipoClase)return false;
    return true;
  });
  var totalTareas=tareasDeAud.length;
  var resueltas=tareasDeAud.filter(function(t){return!esPendiente(t);}).length;

  /* Blindaje: solo se archiva al 100%. Sin esto se colaban auditorías a medias
     (KN EXPRESS entró con 9/19) porque el conteo de aquí no siempre coincide
     con el de calcAudStats. */
  if(totalTareas<=0||resueltas<totalTareas){
    FINALIZADAS=FINALIZADAS.filter(function(f){return !(f.aud_key===ak&&f._pending);});
    delete _finRegistrandoEnVuelo[akm];
    return;
  }

  /* fecha_inicio = fechaCreacion mas antigua */
  var fechaInicio=null;
  tareasDeAud.forEach(function(t){
    var fc=t.fechaCreacion?String(t.fechaCreacion).split('T')[0]:null;
    if(fc&&(!fechaInicio||fc<fechaInicio))fechaInicio=fc;
  });
  if(!fechaInicio&&a.fecha)fechaInicio=String(a.fecha).split('T')[0];

  /* fecha_finalizacion = fechaCumpl o fechaTerm mas reciente */
  var fechaFin=null;
  tareasDeAud.forEach(function(t){
    var fd=t.fechaCumpl?String(t.fechaCumpl).split('T')[0]:
           (t.fechaTerm?String(t.fechaTerm).split('T')[0]:null);
    if(fd&&(!fechaFin||fd>fechaFin))fechaFin=fd;
  });
  if(!fechaFin){
    fechaFin=new Date().toISOString().split('T')[0];
  }

  /* Solo se archiva cuando lleva al menos DIAS_PARA_FINALIZAR días al 100%.
     Si se resolvió hoy, se queda en Auditorías unos días más. */
  var _hoy=new Date(); _hoy.setHours(0,0,0,0);
  var _ff=fromISO(fechaFin);
  if(_ff){
    var _diasAl100=Math.floor((_hoy-_ff)/(1000*60*60*24));
    if(_diasAl100<DIAS_PARA_FINALIZAR){
      /* aún no: quitar la marca provisional y salir */
      FINALIZADAS=FINALIZADAS.filter(function(f){return !(f.aud_key===ak&&f._pending);});
      delete _finRegistrandoEnVuelo[akm];
      return;
    }
  }

  var duracion=(fechaInicio&&fechaFin)?diasEntre(fechaInicio,fechaFin):null;
  var pct=parseFloat(a.pctCumpl)||0;
  if(pct>1)pct=pct/100;

  var _finRowRaw={
    aud_key:ak,tienda:a.tienda||'',mes:(a.mes||'').toUpperCase(),
    clase:a.clase||'',pct_cumpl:pct,
    total_tareas:totalTareas,resueltas:resueltas,
    fecha_inicio:fechaInicio,fecha_finalizacion:fechaFin,
    duracion_dias:duracion,
    completado_por:(_session&&(_session.nombre||_session.username))||null
  };
  var row=await encObj(_finRowRaw,FIELDS.tareas_finalizadas);
  try{
    var r=await client.from('tareas_finalizadas').insert([row]).select();
    if(r.error){
      FINALIZADAS=FINALIZADAS.filter(function(f){return f.aud_key!==ak||!f._pending;});
      if(!r.error.message.includes('duplicate')&&!r.error.message.includes('unique'))
        console.warn('registrarFinalizada:',r.error.message);
      delete _finRegistrandoEnVuelo[akm];
      return;
    }
    var ix=FINALIZADAS.findIndex(function(f){return f.aud_key===ak&&f._pending;});
    /* Siempre usar los datos originales sin cifrar para mostrar en UI */
    var saved=Object.assign({},_finRowRaw);
    saved.id=(r.data&&r.data[0])?r.data[0].id:('tmp_'+Date.now());
    if((saved.pct_cumpl||0)>1)saved.pct_cumpl=saved.pct_cumpl/100;
    if(ix>=0)FINALIZADAS[ix]=saved;
    delete _finRegistrandoEnVuelo[akm];
    if(VIEW==='finalizadas'){fillFinFilters();renderFinalizadas();}
  }catch(e){
    FINALIZADAS=FINALIZADAS.filter(function(f){return f.aud_key!==ak||!f._pending;});
    delete _finRegistrandoEnVuelo[akm];
    console.warn('registrarFinalizada:',e.message);
  }
}

async function deleteFinalizada(id){
  if(!_session||!['admin','admin_auditor'].includes(_session.rol)){toast('⚠ Solo administradores');return;}
  if(!confirm('¿Eliminar este registro de finalización?'))return;
  var client=getSbClient();if(!client)return;
  try{
    var r=await client.from('tareas_finalizadas').delete().eq('id',id);
    if(r.error){toast('⚠ '+r.error.message);return;}
    FINALIZADAS=FINALIZADAS.filter(function(f){return f.id!==id;});
    fillFinFilters();renderFinalizadas();toast('☁️ Registro eliminado');
  }catch(e){toast('⚠ '+e.message);}
}
