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
 let S0={};try{S0=JSON.parse(localStorage.getItem('evalAlin1'))||{}}catch(e){}
 const ev=C.ev,A=C.alertas(S0),NC=new Set(ev.map(x=>x.n)).size,T=TIPOS.map(x=>{const L=ev.filter(y=>y.tp.includes(x[0]));return[x[1],L.length,new Set(L.map(y=>y.n)).size,x[3].join(', ')]}),O=ev.filter(y=>!y.tp.length);T.push(['Sin tipo reconocido',O.length,new Set(O.map(y=>y.n)).size,'alineacion']);
 const mxT=Math.max(1,...T.map(r=>r[1])),M={};ev.forEach(x=>{if(x.e){const m=x.e.slice(0,7);M[m]=(M[m]||new Set()).add(x.n)}});const mk=Object.keys(M).sort().slice(-14),mxM=Math.max(1,...mk.map(m=>M[m].size));
 el.innerHTML='<div class="c"><b>Cambios de ingeniería (NCD)</b><div class="mu">'+esc(C.info||'Aún no hay bitácora cargada.')+'</div><p><button class="p" id="nb">Cargar bitácora NCD</button><input id="nf" type="file" accept=".xlsx,.xlsm,.xls" hidden></p><div class="mu">La fecha de emisión es la fecha en que el cambio se aprueba. Si es posterior a la evaluación de alineación, el isométrico queda en "Reevaluar".</div></div>'+
 (ev.length?'<div class="ui"><div class="kp">'+S.kp('Renglones',ev.length)+S.kp('NCD distintas',NC)+S.kp('Isométricos con NCD',Object.keys(C.idx).length)+S.kp('Sin fecha de emisión',ev.filter(x=>!x.e).length,'wa')+S.kp('Evaluaciones a reevaluar',A.length,A.length?'er':'ok')+'</div>'+
 '<div class="c"><b>Tipo de cambio</b><div class="mu" style="margin:2px 0 8px">Clasificación por palabras del título; largo de la barra = renglones.</div>'+T.map(r=>S.hb(esc(r[0])+'<br><span class="mu" style="font-size:11px">avisa a: '+esc(r[3])+'</span>',100*r[1]/mxT,r[1]+' renglones · '+r[2]+' NCD','var(--ac)')).join('')+'</div>'+
 '<div class="c"><b>NCD emitidas por mes</b><div class="mu" style="margin:2px 0 8px">Cuántas NCD distintas se aprobaron cada mes (últimos 14 meses con datos).</div>'+mk.map(m=>S.hb(m,100*M[m].size/mxM,M[m].size+' NCD','var(--ok)')).join('')+'</div>'+
 S.tabla('Reevaluar en Alineación',['Isométrico','Evaluado el','Última NCD','Emisión','Cambio','#NCD nuevas'],A.slice(0,300).map(r=>[r.iso,r.fe,r.u.n,r.u.e,r.u.t,r.n]),300,A.length?'Evaluaciones de alineación anteriores a una NCD aprobada.':'Ninguna evaluación quedó anterior a una NCD.')+'</div>':'');
 el.querySelector('#nb').onclick=()=>el.querySelector('#nf').click();
 el.querySelector('#nf').onchange=async e=>{const f=e.target.files[0];e.target.value='';if(!f)return;
  try{await C.importar(f)}catch(x){alert(x.message||'No se pudo leer el archivo.')}C.vista(el)};
};
/* NCD aprobadas después de la fecha de un dato, para un módulo; y tabla de alertas {llave:fecha} */
C.nuevas=(k,mod,d)=>C.de(k).filter(x=>x.m.includes(mod)&&x.e&&d&&x.e>d).sort((a,b)=>b.e.localeCompare(a.e));
C.alertasMod=(mod,F)=>Object.keys(F).map(k=>{const L=C.nuevas(k,mod,F[k]);return L.length?[k,F[k],L[0].n,L[0].e,L[0].t,L.length]:null}).filter(Boolean);
Suite.mods.push({cargar:C.cargar,id:'cambios',nom:'Cambios de ingeniería',desc:'NCD y alertas de reevaluación',vista:C.vista,activo:true});
})(typeof window!=='undefined'?window:globalThis);
