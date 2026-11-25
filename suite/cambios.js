/* MÓDULO · CAMBIOS DE INGENIERÍA (bitácora de NCD). Fecha de emisión = fecha en que el cambio se aprueba. */
(function(g){
const Suite=g.Suite,N=Suite.norm,esc=Suite.esc,C=Suite.cambios={ev:[],idx:{},info:''};
/* [id, nombre, patrón, módulos que deben enterarse] */
const TIPOS=[
 ['sop','Soportes y placas de carga',/SOPORT|PLACA.? DE CARGA|UBOLT|GUIA|CLIP/,['soportes','alineacion']],
 ['pres','Presión o temperatura de prueba',/PRESION|VALORES DE P|PRUEBA/,['pruebas']],
 ['fig8','Fig. 8, bridas y anillos',/FIG|ANILLO|BRIDA/,['pruebas','materiales']],
 ['jun','Juntas, dimensiones y válvulas',/JUNTA|DIMENSION|ARREGLO|VALV/,['soldadura','materiales','alineacion']],
 ['tray','Trayectoria, elevación y conexiones',/ELEVAC|TRAYECT|CONEX|CONTINUAC|BOQUILLA/,['montaje','alineacion']],
 ['diam','Diámetro, línea nueva y drenaje',/DIAMETRO|NUEVA LINEA|DRENAJE/,['materiales','soldadura','alineacion']]];
const tipos=t=>{const u=N(t).toUpperCase();return TIPOS.filter(x=>x[2].test(u)).map(x=>x[0])};
/* Sin tipo reconocido: se avisa a Alineación por prudencia */
const mods=tp=>tp.length?[...new Set(TIPOS.filter(x=>tp.includes(x[0])).flatMap(x=>x[3]))]:['alineacion'];

C.parse=wb=>{
 const A=XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]],{header:1,raw:true,defval:null,blankrows:false});
 const hi=A.findIndex(r=>r.some(c=>c&&N(c).toLowerCase().includes('scds/ncd')));if(hi<0)return null;
 const H=A[hi].map(c=>N(c).toLowerCase()),ix=s=>H.findIndex(h=>h.includes(s));
 const c={n:ix('scds/ncd'),t:ix('descripcion'),d:ix('documentos afectados'),e:ix('fecha de emision'),k:ix('fecha de carga a cwx')};
 if(c.n<0||c.d<0||c.e<0)return null;
 return A.slice(hi+1).filter(r=>r[c.n]).map(r=>{const tp=tipos(r[c.t]);
  return{n:String(r[c.n]).trim(),t:String(r[c.t]||'').trim(),d:String(r[c.d]||'').trim(),k:Suite.cruce(r[c.d]),e:Suite.fecha(r[c.e]),c:c.k>=0?Suite.fecha(r[c.k]):null,tp,m:mods(tp)}})
};
C.set=(ev,info)=>{C.ev=ev;C.info=info||'';C.idx={};for(const x of ev)if(x.k)(C.idx[x.k]=C.idx[x.k]||[]).push(x)};
C.cargar=async()=>{try{const v=await Suite.db('g','ncd');if(v&&v.ev)C.set(v.ev,v.info)}catch(e){}return C.ev.length};
C.importar=async f=>{const ev=C.parse(XLSX.read(await f.arrayBuffer(),{type:'array',cellDates:true}));
 if(!ev)throw new Error('No encontré las columnas "No. SCDS/NCD", "Documentos afectados" y "Fecha de Emisión". Elige la bitácora de NCD.');
 const info='Bitácora: '+f.name+' · cargada el '+Suite.hoy();C.set(ev,info);await Suite.db('s','ncd',{ev,info});return ev.length};
C.de=iso=>C.idx[Suite.cruce(iso)]||[];
/* Isométricos de la app de alineación evaluados ANTES de una NCD ya aprobada para ese isométrico */
C.alertas=S=>Object.keys(S).map(k=>{const fe=S[k].fecha,
 L=C.de(k).filter(x=>x.m.includes('alineacion')&&x.e&&fe&&x.e>fe).sort((a,b)=>b.e.localeCompare(a.e));
 return L.length?{iso:k,fe,u:L[0],n:L.length}:null}).filter(Boolean).sort((a,b)=>b.u.e.localeCompare(a.u.e));
/* Aviso para el panel de un isométrico en la app de alineación */
C.aviso=(iso,fe)=>{const L=C.de(iso).filter(x=>x.m.includes('alineacion'));if(!L.length)return'';
 const nw=fe?L.filter(x=>x.e&&x.e>fe).sort((a,b)=>b.e.localeCompare(a.e)):[],sf=L.filter(x=>!x.e).length;
 if(nw.length)return'<p class="wa" style="padding:8px">⚠ Reevaluar: '+esc(nw[0].n)+' emitida el '+nw[0].e+' · '+esc(nw[0].t)+(nw.length>1?' (+'+(nw.length-1)+' NCD más)':'')+'</p>';
 return'<p class="mu" style="margin:6px 0">Cambios de ingeniería: '+L.length+' NCD'+(fe?', todas anteriores a la evaluación':'')+(sf?' · '+sf+' sin fecha de emisión':'')+'.</p>'};
C.vista=el=>{
 let S={};try{S=JSON.parse(localStorage.getItem('evalAlin1'))||{}}catch(e){}
 const ev=C.ev,A=C.alertas(S),k=(l,v,c)=>'<div class="c '+(c||'')+'"><span class="mu">'+l+'</span><b>'+v+'</b></div>';
 const T=TIPOS.map(x=>{const L=ev.filter(y=>y.tp.includes(x[0]));return[x[1],L.length,new Set(L.map(y=>y.n)).size,x[3].join(', ')]});
 const O=ev.filter(y=>!y.tp.length);T.push(['Sin tipo reconocido',O.length,new Set(O.map(y=>y.n)).size,'alineacion']);
 el.innerHTML='<div class="c"><b>Cambios de ingeniería (NCD)</b><div class="mu">'+esc(C.info||'Aún no hay bitácora cargada.')+'</div><p><button class="p" id="nb">Cargar bitácora NCD</button><input id="nf" type="file" accept=".xlsx,.xlsm,.xls" hidden></p><div class="mu">La fecha de emisión es la fecha en que el cambio se aprueba. Si es posterior a la evaluación de alineación, el isométrico queda en "Reevaluar".</div></div>'+
 (ev.length?'<div class="kp">'+k('Renglones',ev.length)+k('NCD distintas',new Set(ev.map(x=>x.n)).size)+k('Isométricos con NCD',Object.keys(C.idx).length)+k('Sin fecha de emisión',ev.filter(x=>!x.e).length,'wa')+k('Evaluaciones a reevaluar',A.length,A.length?'er':'ok')+'</div>'+
 '<div class="c"><b>Reevaluar en Alineación</b>'+(A.length?'<div class="tw"><table><tr><th>Isométrico</th><th>Evaluado el</th><th>Última NCD</th><th>Emisión</th><th>Cambio</th><th>NCD nuevas</th></tr>'+A.slice(0,300).map(r=>'<tr><td>'+esc(r.iso)+'</td><td>'+esc(r.fe)+'</td><td>'+esc(r.u.n)+'</td><td>'+r.u.e+'</td><td style="white-space:normal">'+esc(r.u.t)+'</td><td>'+r.n+'</td></tr>').join('')+'</table></div>':'<p class="mu">'+(Object.keys(S).length?'Ninguna evaluación quedó anterior a una NCD.':'No hay evaluaciones guardadas en este dispositivo.')+'</p>')+'</div>'+
 '<div class="c"><b>Tipo de cambio</b> <span class="mu">· clasificación por palabras del título</span><div class="tw"><table><tr><th>Tipo</th><th>Renglones</th><th>NCD</th><th>Módulos a avisar</th></tr>'+T.map(r=>'<tr><td style="white-space:normal">'+esc(r[0])+'</td><td>'+r[1]+'</td><td>'+r[2]+'</td><td>'+esc(r[3])+'</td></tr>').join('')+'</table></div></div>':'');
 el.querySelector('#nb').onclick=()=>el.querySelector('#nf').click();
 el.querySelector('#nf').onchange=async e=>{const f=e.target.files[0];e.target.value='';if(!f)return;
  try{await C.importar(f)}catch(x){alert(x.message||'No se pudo leer el archivo.')}C.vista(el)};
};
/* NCD aprobadas después de la fecha de un dato, para un módulo; y tabla de alertas {llave:fecha} */
C.nuevas=(k,mod,d)=>C.de(k).filter(x=>x.m.includes(mod)&&x.e&&d&&x.e>d).sort((a,b)=>b.e.localeCompare(a.e));
C.alertasMod=(mod,F)=>Object.keys(F).map(k=>{const L=C.nuevas(k,mod,F[k]);return L.length?[k,F[k],L[0].n,L[0].e,L[0].t,L.length]:null}).filter(Boolean);
Suite.mods.push({cargar:C.cargar,id:'cambios',nom:'Cambios de ingeniería',desc:'NCD y alertas de reevaluación',vista:C.vista,activo:true});
})(typeof window!=='undefined'?window:globalThis);
