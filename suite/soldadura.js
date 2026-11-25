/* MÓDULO · SOLDADURA (reporte FR de Weld Console: solo juntas reportadas, sin universo total).
   Lee Excel grandes en modo denso y CSV por flujo (sin cargar todo en memoria). */
(function(g){const S=g.Suite,W=S.soldadura={d:{},w:{},p:{},info:''},AC=/ACCEPT|ACEPT/;
const fe=v=>{const m=typeof v==='string'&&v.trim().match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);return m?S.fecha(m[3]+'-'+m[2].padStart(2,'0')+'-'+m[1].padStart(2,'0')):S.fecha(v)};
W.nuevo=()=>({d:{},w:{},p:{},I:null,R:null,n:0,pri:null});
W.fila=(st,r)=>{
 if(!st.I){if(!st.pri)st.pri=r;
  if(st.n++<40&&r.some(c=>/^drawing\s*(no|number)/.test(S.norm(c).toLowerCase()))){
   const H=r.map(c=>S.norm(c).toLowerCase()),c=n=>{const i=H.indexOf(n);return i>=0?i:H.findIndex(h=>h.startsWith(n))};
   st.I={d:c('drawing'),l:c('location'),f:c('fab date'),w:c('balance welders'),rp:c('repair'),pk:c('package number'),cm:c('commodity'),ln:c('line'),sh:c('sheet')};
   st.R=['vt result','rt results','pt result','ht result','pwht result'].map(c)}
  return}
 const I=st.I,{d,w,p}=st;if(!r[I.d])return;
 const k=S.cruce(r[I.d])||(I.cm>=0&&I.ln>=0&&I.sh>=0?[r[I.cm],r[I.ln],r[I.sh]].join('-'):null);if(!k)return;
 const f=fe(r[I.f]),t=d[k]=d[k]||{iso:String(r[I.d]),fw:0,sw:0,vt:0,rej:0,rep:0,last:''},rs=st.R.map(i=>i>=0?String(r[i]||'').toUpperCase():''),
 rej=rs.some(x=>/REJ|RECH/.test(x)),fw=String(r[I.l]||'').trim().toUpperCase()==='FW',vt=AC.test(rs[0]);
 fw?t.fw++:t.sw++;if(vt)t.vt++;if(rej)t.rej++;if(+r[I.rp]>0)t.rep++;if(f&&f>t.last)t.last=f;
 const pk=I.pk>=0&&r[I.pk]?String(r[I.pk]).trim():'';if(pk){const q=p[pk]=p[pk]||{n:0,fw:0,vt:0,rej:0};q.n++;if(fw)q.fw++;if(vt)q.vt++;if(rej)q.rej++}
 const wl=I.w>=0?String(r[I.w]||'').trim().toUpperCase():'';if(wl&&wl!=='N/A'){const x=w[wl]=w[wl]||{n:0,rej:0};x.n++;if(rej)x.rej++}};
/* Hoja de Excel (modo denso o normal) */
W.hoja=(ws,st)=>{st.n=0;const D=Array.isArray(ws)?ws:ws['!data'];
 if(D){for(let i=0;i<D.length;i++){if(!st.I&&st.n>=40)break;const x=D[i];W.fila(st,x?x.map(c=>c?c.v:null):[])}}
 else XLSX.utils.sheet_to_json(ws,{header:1,raw:true,defval:null,blankrows:false}).forEach(r=>W.fila(st,r))};
W.parse=wb=>{const st=W.nuevo();for(const n of wb.SheetNames){W.hoja(wb.Sheets[n],st);if(st.I)break}
 if(!st.I)throw new Error('No encontré la columna "Drawing No". Hojas: '+wb.SheetNames.join(', '));return{d:st.d,w:st.w,p:st.p}};
W.leerXlsx=async(f,st)=>{const buf=await f.arrayBuffer(),nom=XLSX.read(buf,{type:'array',bookSheets:true}).SheetNames;
 for(const n of nom){const ws=XLSX.read(buf,{type:'array',sheets:n,dense:true,cellDates:true,cellStyles:false,cellNF:false,cellHTML:false,cellFormula:false}).Sheets[n];
  if(ws)W.hoja(ws,st);if(st.I)break}
 return nom};
/* CSV por flujo: detecta coma o punto y coma, comillas y saltos de línea */
W.leerCsv=async(f,st)=>{const rd=f.stream().getReader(),dec=new TextDecoder('utf-8');let q=false,pq=false,fld='',row=[],dl=null;
 const fin=()=>{row.push(fld);fld='';if(row.length>1||row[0]!=='')W.fila(st,row);row=[]};
 for(;;){const {done,value}=await rd.read(),t=done?'':dec.decode(value,{stream:true});
  if(dl===null&&t){const l=t.split('\n')[0];dl=l.split(';').length>l.split(',').length?';':','}
  for(let i=0;i<t.length;i++){const c=t[i];
   if(pq){pq=false;if(c==='"'){fld+='"';continue}q=false}
   if(q){if(c==='"')pq=true;else fld+=c;continue}
   if(c==='"')q=true;else if(c===dl){row.push(fld);fld=''}else if(c==='\n')fin();else if(c!=='\r')fld+=c}
  if(done){if(fld!==''||row.length)fin();break}}};
W.cargar=async()=>{try{const v=await S.db('g','soldadura');if(v&&v.d){Object.assign(W,v)}}catch(e){}return Object.keys(W.d).length};
W.importar=async f=>{const b=typeof document!=='undefined'&&document.getElementById('nb');
 if(b){b.textContent='Leyendo archivo… puede tardar';b.disabled=true}await new Promise(r=>setTimeout(r,60));
 const st=W.nuevo();let nom=[];
 try{if(/\.csv$/i.test(f.name))await W.leerCsv(f,st);else nom=await W.leerXlsx(f,st)}
 catch(e){throw new Error('El navegador no pudo abrir el archivo por su tamaño. En Excel usa Guardar como → CSV UTF-8 y cárgalo de nuevo. ('+(e.message||e)+')')}
 if(!st.I)throw new Error('No encontré la columna "Drawing No". '+(nom.length?'Hojas: '+nom.join(', ')+'. ':'')+'Primera fila leída: '+(st.pri||[]).slice(0,8).map(c=>S.norm(c)).join(' | '));
 Object.assign(W,{d:st.d,w:st.w,p:st.p,info:'FR: '+f.name+' · cargado el '+S.hoy()});await S.db('s','soldadura',{d:W.d,w:W.w,p:W.p,info:W.info})};
W.resumen=k=>{const x=W.d[k];return x?(x.fw+x.sw)+' juntas reportadas ('+x.fw+' de sitio) · VT aceptadas '+x.vt+' · rechazos '+x.rej+' · reparadas '+x.rep+(x.last?' · última '+x.last:''):''};
W.vista=el=>{const L=Object.values(W.d),T=f=>L.reduce((a,x)=>a+x[f],0),F={};Object.keys(W.d).forEach(k=>{if(W.d[k].last)F[k]=W.d[k].last});
 const A=S.cambios.alertasMod('soldadura',F),pc=(a,b)=>Math.round(1000*a/(b||1))/10+' %';
 el.innerHTML=S.cab('Soldadura',W.info,'Cargar reporte FR (Excel o CSV)')+(L.length?'<div class="kp">'+S.kp('Isométricos con juntas',L.length)+S.kp('Juntas reportadas',T('fw')+T('sw'))+S.kp('De sitio (FW)',T('fw'))+S.kp('VT aceptadas',T('vt'),'ok')+S.kp('Con rechazo',T('rej'),T('rej')?'er':'ok')+S.kp('Reparadas',T('rep'),'wa')+S.kp('Revisar por NCD',A.length,A.length?'er':'ok')+'</div>'+
 '<div class="c mu">Solo incluye juntas ya reportadas: no hay universo total, así que no se calcula % soldado.</div>'+
 S.tabla('Por soldador (top 30)',['Soldador','Juntas','Con rechazo','% rechazo'],Object.keys(W.w).sort((a,b)=>W.w[b].n-W.w[a].n).slice(0,30).map(k=>[k,W.w[k].n,W.w[k].rej,pc(W.w[k].rej,W.w[k].n)]),30)+
 S.tabla('Por paquete de prueba (campo Package Number del FR)',['Paquete','Juntas','De sitio','VT aceptadas','Con rechazo'],Object.keys(W.p).sort().map(k=>[k,W.p[k].n,W.p[k].fw,W.p[k].vt,W.p[k].rej]),500)+
 S.tabla('Juntas anteriores a una NCD aprobada · revisar',['Isométrico','Última junta','Última NCD','Emisión','Cambio','NCD'],A.map(x=>[W.d[x[0]].iso,...x.slice(1)])):'');
 const nf=el.querySelector('#nf');if(nf)nf.accept='.xlsx,.xlsm,.xls,.csv';
 S.ligar(el,W.importar,W.vista)};
S.mods.push({id:'soldadura',nom:'Soldadura',desc:'Juntas, soldadores y rechazos (FR)',vista:W.vista,cargar:W.cargar,resumen:W.resumen});
})(typeof window!=='undefined'?window:globalThis);
