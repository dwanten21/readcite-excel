// @ts-check
// Modelo de datos: mapeo columnas de Excel ↔ objetos del núcleo, CRUD, búsqueda, filtros y colecciones.
// Los objetos usan los nombres de campo de la web (text, author…) para que apa/nombres/reto funcionen sin traducir.
// Todas las funciones son puras: devuelven arrays nuevos y nunca mutan la entrada.

/** @typedef {{id:string,text:string,author:string,book:string,topic:string,concept:string,chapter:string,page:string,
 *   year:string,originalYear:string,publisher:string,translator:string,collection:string,understood:string,why:string,
 *   likes:number,liked:boolean,savedAt:number|null,createdAt:number|null,user:string,ratio:string,imagePos:string,
 *   image:string,mask:any}} Cita */
/** @typedef {{titulo:string,autor:string,paginas:number,estado:string,leidas:number,anio:string}} Libro */
/** @typedef {{nombre:string,descripcion:string,creada:number|null}} Coleccion */
/** @typedef {{citaId:string,usuario:string,cuando:number|null,texto:string}} Comentario */

// ── Esquema de las tablas (orden = orden de columnas en Excel) ──
/** campo del núcleo → encabezado exacto en tblCitas */
export const COLUMNAS_CITAS=/** @type {const} */([
  ['id','id'],['text','texto'],['author','autor'],['book','libro'],['topic','tema'],['concept','concepto'],
  ['chapter','capitulo'],['page','pagina'],['year','anio_edicion'],['originalYear','anio_original'],['publisher','editorial'],
  ['translator','traductor'],['collection','coleccion'],['understood','entendido'],['why','por_que'],['likes','likes'],
  ['liked','me_gusta'],['savedAt','guardada_fecha'],['createdAt','creada_fecha'],['user','usuario'],['ratio','ratio_tarjeta'],
  ['imagePos','posicion_imagen'],['image','imagen_path'],['mask','mask_json']]);
export const COLUMNAS_LIBROS=['titulo','autor','paginas','estado','leidas','anio'];
export const COLUMNAS_COLECCIONES=['nombre','descripcion','creada_fecha'];
export const COLUMNAS_COMENTARIOS=['cita_id','usuario','cuando','texto'];
export const COLUMNAS_CONFIG=['clave','valor'];

export const ESTADOS=/** @type {const} */(['leyendo','leidos','quiero','abandonados']);
export const ETIQUETAS_ESTADO={leyendo:'Leyendo',leidos:'Leídos',quiero:'Quiero leer',abandonados:'Abandonados'};
export const RATIOS=['poster','square'];
export const POSICIONES=['top','center','bottom'];
export const CONFIG_DEF={idioma:'es',traduccion_auto:true,usuario_local:'Tú',paginas_dia:20,
  cuaderno_citas_por_pagina:2,cuaderno_pestanas_espejo:true,cuaderno_anio:''};

// ── Fechas: serial de Excel (días desde 1899-12-30, hora local) ↔ ms ──
/** @param {unknown} v @returns {number|null} */
export function serialAMs(v){
  if(v===''||v==null)return null;const n=Number(v);if(!isFinite(n))return null;
  const d=Math.floor(n),f=n-d;const base=new Date(1899,11,30);base.setDate(base.getDate()+d);
  return base.getTime()+Math.round(f*86400000);
}
/** @param {number|null|undefined} ms @returns {number|string} */
export function msASerial(ms){
  if(ms==null||!isFinite(ms))return '';const t=new Date(ms);const medianoche=new Date(t.getFullYear(),t.getMonth(),t.getDate());
  const dias=Math.round((medianoche.getTime()-new Date(1899,11,30).getTime())/86400000);
  return dias+(t.getTime()-medianoche.getTime())/86400000;
}

/**
 * Protección contra inyección de fórmulas («CSV injection»): un texto que empiece por = + - @ (o tabulador/retorno)
 * se escribe con un apóstrofo delante, que Excel usa como prefijo de texto y no muestra ni devuelve al leer.
 * Así una «cita» como =WEBSERVICE("http://…"&A1) se guarda como texto y nunca se ejecuta.
 * @param {unknown} v
 */
export const textoSeguro = v => (typeof v === 'string' && /^[=+\-@\t\r]/.test(v) ? "'" + v : v);

const txt=/** @param {unknown} v */v=>v==null?'':String(v).trim();
const ent=/** @param {unknown} v */v=>{const n=parseInt(String(v));return isFinite(n)?n:0};
const bool=/** @param {unknown} v */v=>v===true||v===1||/^(true|verdadero|sí|si|1)$/i.test(String(v));

// ── Filas ↔ objetos ──
/** Fila de tblCitas (array, mismo orden que los encabezados dados) → Cita.
 * @param {unknown[]} fila @param {string[]} encabezados */
export function filaACita(fila,encabezados){
  /** @type {Record<string,unknown>} */const r={};encabezados.forEach((h,i)=>r[h]=fila[i]);
  /** @type {any} */const c={};
  for(const [campo,col] of COLUMNAS_CITAS){
    const v=r[col];
    if(campo==='likes')c[campo]=ent(v);
    else if(campo==='liked')c[campo]=bool(v);
    else if(campo==='savedAt'||campo==='createdAt')c[campo]=serialAMs(v);
    else if(campo==='mask'){try{c[campo]=txt(v)?JSON.parse(txt(v)):null}catch(e){c[campo]=null}}
    else c[campo]=txt(v);
  }
  return /** @type {Cita} */(c);
}
/** Cita → fila en el orden de `encabezados`. Columnas desconocidas → undefined (no se tocan).
 * @param {Partial<Cita>} c @param {string[]} encabezados */
export function citaAFila(c,encabezados){
  const porCol=Object.fromEntries(COLUMNAS_CITAS.map(([campo,col])=>[col,campo]));
  return encabezados.map(h=>{const campo=porCol[h];if(!campo)return undefined;const v=/** @type {any} */(c)[campo];
    if(campo==='savedAt'||campo==='createdAt')return msASerial(v);
    if(campo==='mask')return v?JSON.stringify(v):'';
    if(campo==='liked')return !!v;
    if(campo==='likes')return v||0;
    return v==null?'':v;});
}
/** @param {unknown[]} f @param {string[]} h @returns {Libro} */
export function filaALibro(f,h){const r=/** @type {any} */(Object.fromEntries(h.map((k,i)=>[k,f[i]])));
  return {titulo:txt(r.titulo),autor:txt(r.autor),paginas:ent(r.paginas),estado:ESTADOS.includes(txt(r.estado))?txt(r.estado):'quiero',leidas:Math.max(0,ent(r.leidas)),anio:txt(r.anio)}}
/** @param {unknown[]} f @param {string[]} h @returns {Coleccion} */
export function filaAColeccion(f,h){const r=/** @type {any} */(Object.fromEntries(h.map((k,i)=>[k,f[i]])));return {nombre:txt(r.nombre),descripcion:txt(r.descripcion),creada:serialAMs(r.creada_fecha)}}
/** @param {unknown[]} f @param {string[]} h @returns {Comentario} */
export function filaAComentario(f,h){const r=/** @type {any} */(Object.fromEntries(h.map((k,i)=>[k,f[i]])));return {citaId:txt(r.cita_id),usuario:txt(r.usuario),cuando:serialAMs(r.cuando),texto:txt(r.texto)}}
/** Filas clave/valor → objeto de configuración con valores por defecto. @param {unknown[][]} filas */
export function leerConfig(filas){
  /** @type {any} */const c={...CONFIG_DEF};
  filas.forEach(([k,v])=>{const key=txt(k);if(!key||v===''||v==null)return;const def=/** @type {any} */(CONFIG_DEF)[key];
    c[key]=typeof def==='boolean'?bool(v):typeof def==='number'?(parseInt(String(v))||def):v});
  return c;
}

// ── CRUD de citas ──
/** Campos obligatorios que faltan (mismos que la web: cita, autor, libro, lo que entendí). @param {Partial<Cita>} f */
export function camposFaltantes(f){
  return /** @type {const} */([['text','cita'],['author','autor'],['book','libro'],['understood','lo que entendiste']]).filter(([k])=>!txt(/** @type {any} */(f)[k])).map(([,l])=>l);
}
/** @param {string} id @param {string} msg */
export const errorCita=(id,msg)=>`Cita ${id}: ${msg}`;

/** Crea una cita desde el formulario (reglas de submitCreate). No muta `citas`.
 * @param {Cita[]} citas @param {Partial<Cita>} form @param {{ahora:number, usuario:string}} ctx
 * @returns {{ok:true, cita:Cita, citas:Cita[]}|{ok:false, error:string}} */
export function crearCita(citas,form,{ahora,usuario}){
  const falta=camposFaltantes(form);
  if(falta.length)return {ok:false,error:'Completa '+falta.join(', ')};
  let id='u'+ahora;while(citas.some(c=>c.id===id))id='u'+(++ahora);
  const base=filaACita([],[]);
  /** @type {Cita} */const cita={...base,...form,ratio:form.ratio||'poster',imagePos:form.imagePos||'center',id,user:usuario,likes:0,liked:false,
    topic:txt(form.topic)||'Sin tema',savedAt:null,createdAt:ahora};
  return {ok:true,cita,citas:[cita,...citas]};
}

/** Edita campos de una cita (id, createdAt y user no se cambian).
 * @param {Cita[]} citas @param {string} id @param {Partial<Cita>} cambios */
export function editarCita(citas,id,cambios){
  const i=citas.findIndex(c=>c.id===id);if(i<0)return {ok:false,error:errorCita(id,'no existe')};
  const {id:_a,createdAt:_b,user:_c,...resto}=cambios;const nueva={...citas[i],...resto};
  const falta=camposFaltantes(nueva);if(falta.length)return {ok:false,error:errorCita(id,'falta '+falta.join(', '))};
  if('topic' in resto&&!txt(nueva.topic))nueva.topic='Sin tema';
  const out=citas.slice();out[i]=nueva;return {ok:true,cita:nueva,citas:out};
}

/** Borra una cita y sus comentarios. @param {Cita[]} citas @param {Comentario[]} comentarios @param {string} id */
export function borrarCita(citas,comentarios,id){
  if(!citas.some(c=>c.id===id))return {ok:false,error:errorCita(id,'no existe')};
  return {ok:true,citas:citas.filter(c=>c.id!==id),comentarios:comentarios.filter(c=>c.citaId!==id)};
}

/** @param {Cita[]} citas @param {string} id @param {(c:Cita)=>Cita} fn */
function actualizar(citas,id,fn){const i=citas.findIndex(c=>c.id===id);if(i<0)return {ok:false,error:errorCita(id,'no existe')};const out=citas.slice();out[i]=fn(citas[i]);return {ok:true,cita:out[i],citas:out}}
/** @param {Cita[]} citas @param {string} id @param {number} ahora */
export const guardarCita=(citas,id,ahora)=>actualizar(citas,id,c=>({...c,savedAt:ahora}));
/** @param {Cita[]} citas @param {string} id */
export const quitarGuardada=(citas,id)=>actualizar(citas,id,c=>({...c,savedAt:null}));
/** @param {Cita[]} citas @param {string} id */
export const alternarMeGusta=(citas,id)=>actualizar(citas,id,c=>({...c,liked:!c.liked}));

/** Efectos colaterales de crear una cita (como la web): libro nuevo → 'leyendo'; colección nueva → se registra.
 * @param {Cita} cita @param {Libro[]} libros @param {Coleccion[]} colecciones @param {number} ahora */
export function efectosAlCrear(cita,libros,colecciones,ahora){
  const libroNuevo=libros.some(l=>l.titulo===cita.book)?null:{titulo:cita.book,autor:cita.author,paginas:0,estado:'leyendo',leidas:0,anio:cita.originalYear||cita.year||''};
  const coleccionNueva=cita.collection&&!colecciones.some(c=>c.nombre===cita.collection)?{nombre:cita.collection,descripcion:'',creada:ahora}:null;
  return {libroNuevo,coleccionNueva};
}

// ── Búsqueda, filtros y orden (lógica de renderVals) ──
/** @typedef {{query?:string,author?:string,book?:string,topic?:string,collection?:string,mine?:boolean,soloGuardadas?:boolean,sort?:'likes'|'recent'}} Filtros */

/** Likes visibles: base + 1 si me gusta. @param {Cita} c */
export const likesTotales=c=>(c.likes||0)+(c.liked?1:0);

/** Filtra y ordena. Texto: contiene, sin distinguir mayúsculas, en cita/autor/libro/tema/concepto.
 * Resto: igualdad exacta, combinados con Y. 'likes' = más likes primero (estable); 'recent' = más nuevas primero.
 * @param {Cita[]} citas @param {Filtros} F @param {string} usuarioLocal */
export function filtrarCitas(citas,F,usuarioLocal){
  const q=(F.query||'').trim().toLowerCase();/** @param {string} s */const norm=s=>(s||'').toLowerCase();
  const r=citas.filter(x=>{
    if(q&&![x.text,x.author,x.book,x.topic,x.concept].some(v=>norm(v).includes(q)))return false;
    if(F.author&&x.author!==F.author)return false;if(F.book&&x.book!==F.book)return false;if(F.topic&&x.topic!==F.topic)return false;if(F.collection&&x.collection!==F.collection)return false;
    if(F.mine&&x.user!==usuarioLocal)return false;if(F.soloGuardadas&&!x.savedAt)return false;return true;});
  if((F.sort||'likes')==='likes')return r.slice().sort((a,b)=>likesTotales(b)-likesTotales(a));
  return ordenarRecientes(r);
}
/** Más nuevas primero; sin fecha al final, conservando su orden. @param {Cita[]} citas */
export function ordenarRecientes(citas){return citas.slice().sort((a,b)=>(b.createdAt??-Infinity)-(a.createdAt??-Infinity))}

/** Número de filtros activos (texto aparte), como `activeFilters`. @param {Filtros} F */
export const filtrosActivos=F=>['author','book','topic','collection'].filter(k=>/** @type {any} */(F)[k]).length+(F.mine?1:0);

/** Valores distintos de un campo con su recuento, de más a menos (`count`). @param {Cita[]} citas @param {keyof Cita} clave
 * @returns {[string,number][]} */
export function contar(citas,clave){/** @type {Map<string,number>} */const m=new Map();citas.forEach(x=>{const k=/** @type {string} */(x[clave]);if(!k)return;m.set(k,(m.get(k)||0)+1)});return [...m.entries()].sort((a,b)=>b[1]-a[1])}

/** Opciones de los desplegables de filtros. @param {Cita[]} citas @param {Coleccion[]} colecciones */
export function opcionesFiltros(citas,colecciones){
  return {autores:contar(citas,'author').map(x=>x[0]),libros:contar(citas,'book').map(x=>x[0]),temas:contar(citas,'topic').map(x=>x[0]),colecciones:nombresColecciones(citas,colecciones)};
}

// ── Colecciones ──
/** Nombres: los usados en citas (orden de aparición) + los registrados sin citas. @param {Cita[]} citas @param {Coleccion[]} colecciones */
export function nombresColecciones(citas,colecciones){return [...new Set([...citas.map(x=>x.collection).filter(Boolean),...colecciones.map(c=>c.nombre).filter(Boolean)])]}
/** @param {Cita[]} citas @param {Coleccion[]} colecciones */
export function resumenColecciones(citas,colecciones){
  return nombresColecciones(citas,colecciones).map(nombre=>{const qs=citas.filter(x=>x.collection===nombre);const reg=colecciones.find(c=>c.nombre===nombre);
    return {nombre,descripcion:reg?reg.descripcion:'',count:qs.length,citas:qs,muestras:qs.slice(0,3).map((q,i)=>extracto(q.text,i===0?110:60))}});
}
/** “…” recortado como las teselas de colección. @param {string} t @param {number} n */
export const extracto=(t,n)=>`“${t.slice(0,n)}${t.length>n?'…':''}”`;
/** @param {Coleccion[]} colecciones @param {{nombre:string,descripcion?:string}} datos @param {number} ahora */
export function crearColeccion(colecciones,datos,ahora){
  const nombre=txt(datos.nombre);if(!nombre)return {ok:false,error:'La colección necesita un nombre'};
  if(colecciones.some(c=>c.nombre===nombre))return {ok:false,error:`Ya existe la colección «${nombre}»`};
  const c={nombre,descripcion:txt(datos.descripcion),creada:ahora};return {ok:true,coleccion:c,colecciones:[...colecciones,c]};
}

// ── Comentarios ──
/** @param {Comentario[]} comentarios @param {string} citaId */
export const comentariosDe=(comentarios,citaId)=>comentarios.filter(c=>c.citaId===citaId);
/** @param {Comentario[]} comentarios @param {Cita[]} citas @param {{citaId:string,texto:string,usuario:string,ahora:number}} d */
export function anadirComentario(comentarios,citas,{citaId,texto,usuario,ahora}){
  const t=txt(texto);if(!t)return {ok:false,error:errorCita(citaId,'el comentario está vacío')};
  if(!citas.some(c=>c.id===citaId))return {ok:false,error:errorCita(citaId,'no existe')};
  const c={citaId,usuario,cuando:ahora,texto:t};return {ok:true,comentario:c,comentarios:[...comentarios,c]};
}

// ── Biblioteca ──
/** Libros de la biblioteca: los de tblLibros + los que solo aparecen en citas (estado 'quiero', sin páginas), como la web.
 * @param {Cita[]} citas @param {Libro[]} libros */
export function biblioteca(citas,libros){
  const out=libros.map(l=>({...l,citas:citas.filter(c=>c.book===l.titulo).length}));
  const vistos=new Set(libros.map(l=>l.titulo));
  citas.forEach(c=>{if(!c.book||vistos.has(c.book))return;vistos.add(c.book);
    out.push({titulo:c.book,autor:c.author,paginas:0,estado:'quiero',leidas:0,anio:c.originalYear||c.year||'',citas:citas.filter(y=>y.book===c.book).length})});
  return out;
}
/** @param {ReturnType<typeof biblioteca>} libros @param {string} estado */
export const librosPorEstado=(libros,estado)=>libros.filter(l=>l.estado===estado);
