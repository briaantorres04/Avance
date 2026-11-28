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
 const f=fe(r[I.f]),t=d[k]=d[k]||{iso:String(r[I.d]),tot:0,sold:0,fw:0,fws:0,sv:0,vt:0,rej:0,rep:0,last:''},rs=st.R.map(i=>i>=0?String(r[i]||'').toUpperCase():''),
 rej=rs.some(x=>/REJ|RECH/.test(x)),fw=String(r[I.l]||'').trim().toUpperCase()==='FW',vt=AC.test(rs[0]),sold=!!f,sv=sold&&rs[0].trim()==='';
 /* Universo = todas las juntas del FR; soldada = tiene fecha de fabricación; sin VT = soldada sin resultado de inspección visual */
 t.tot++;if(fw)t.fw++;if(sold){t.sold++;if(fw)t.fws++;if(sv)t.sv++;if(vt)t.vt++;if(f>t.last)t.last=f}if(rej)t.rej++;if(+r[I.rp]>0)t.rep++;
 const pk=I.pk>=0&&r[I.pk]?String(r[I.pk]).trim():'';if(pk){const q=p[pk]=p[pk]||{n:0,sold:0,fw:0,fws:0,sv:0,vt:0,rej:0};q.n++;if(fw)q.fw++;if(sold){q.sold++;if(fw)q.fws++;if(sv)q.sv++;if(vt)q.vt++}if(rej)q.rej++}
 const wl=sold&&I.w>=0?String(r[I.w]||'').trim().toUpperCase():'';if(wl&&wl!=='N/A'){const x=w[wl]=w[wl]||{n:0,rej:0};x.n++;if(rej)x.rej++}};
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
W.cargar=async()=>{try{const v=await S.db('g','soldadura');if(v&&v.d&&v.v===2){Object.assign(W,v)}}catch(e){}return Object.keys(W.d).length};
W.importar=async f=>{const b=typeof document!=='undefined'&&document.getElementById('nb');
 if(b){b.textContent='Leyendo archivo… puede tardar';b.disabled=true}await new Promise(r=>setTimeout(r,60));
 const st=W.nuevo();let nom=[];
 try{if(/\.csv$/i.test(f.name))await W.leerCsv(f,st);else nom=await W.leerXlsx(f,st)}
 catch(e){throw new Error('El navegador no pudo abrir el archivo por su tamaño. En Excel usa Guardar como → CSV UTF-8 y cárgalo de nuevo. ('+(e.message||e)+')')}
 if(!st.I)throw new Error('No encontré la columna "Drawing No". '+(nom.length?'Hojas: '+nom.join(', ')+'. ':'')+'Primera fila leída: '+(st.pri||[]).slice(0,8).map(c=>S.norm(c)).join(' | '));
 Object.assign(W,{d:st.d,w:st.w,p:st.p,info:'FR: '+f.name+' · cargado el '+S.hoy()});await S.db('s','soldadura',{v:2,d:W.d,w:W.w,p:W.p,info:W.info})};
const pc=(a,b)=>Math.round(1000*a/(b||1))/10+' %';
W.resumen=k=>{const x=W.d[k];return x?x.sold+' de '+x.tot+' juntas soldadas ('+pc(x.sold,x.tot)+') · por soldar '+(x.tot-x.sold)+' ('+(x.fw-x.fws)+' de sitio) · sin VT '+x.sv+' · VT aceptadas '+x.vt+' · rechazos '+x.rej+' · reparadas '+x.rep+(x.last?' · última '+x.last:''):''};
W.vista=el=>{const L=Object.values(W.d),T=f=>L.reduce((a,x)=>a+x[f],0),F={};Object.keys(W.d).forEach(k=>{if(W.d[k].last)F[k]=W.d[k].last});
 const A=S.cambios.alertasMod('soldadura',F),cp=L.filter(x=>x.fw>x.fws).sort((a,b)=>(b.fw-b.fws)-(a.fw-a.fws)),mo=x=>((S.montaje&&S.montaje.d)||{})[S.cruce(x.iso)],op=cp.filter(x=>{const m=mo(x);return m&&m.e==='Montado'}),
 tot=T('tot'),sold=T('sold'),vt=T('vt'),sv=T('sv'),ot=sold-vt-sv,por=tot-sold,fw=T('fw'),fws=T('fws'),pn=(a,b)=>Math.round(100*a/(b||1)),
 ST=[['VT aceptada',vt,'var(--ok)'],['Soldada sin VT',sv,'var(--wa)'],['Soldada, VT no aceptada',ot,'var(--er)'],['Por soldar',por,'var(--mu)']].filter(x=>x[1]>0);
 el.innerHTML=S.cab('Soldadura',W.info,'Cargar reporte FR (Excel o CSV)')+(L.length?'<div class="ui"><div class="c"><b>Avance de soldadura</b>'+S.hb('Juntas soldadas',pn(sold,tot),sold+' de '+tot+' · '+pn(sold,tot)+' %','var(--ok)')+S.hb('Juntas de sitio soldadas',pn(fws,fw),fws+' de '+fw+' · '+pn(fws,fw)+' %','var(--ac)')+S.hb('VT aceptada sobre soldadas',pn(vt,sold),vt+' de '+sold+' · '+pn(vt,sold)+' %','var(--ok)')+'<div class="mu" style="margin-top:6px">Universo = todas las juntas del FR. Soldada = tiene fecha de fabricación. Sin VT = soldada sin resultado de inspección visual.</div></div>'+
 '<div class="kp">'+S.kp('Isométricos',L.length)+S.kp('Por soldar',por,'wa')+S.kp('De sitio por soldar',fw-fws,'wa')+S.kp('Oportunidades de soldadura',op.length,'wa')+S.kp('Soldadas sin VT',sv,sv?'wa':'ok')+S.kp('Con rechazo',T('rej'),T('rej')?'er':'ok')+S.kp('Revisar por NCD',A.length,A.length?'er':'ok')+'</div>'+
 '<div class="c"><b>¿En qué estado están las '+tot+' juntas?</b><div class="hb" style="margin-top:8px"><div class="t" style="height:30px">'+S.stk(ST.map(x=>[100*x[1]/(tot||1),x[2],x[0]+': '+x[1]]))+'</div></div><div class="leg">'+ST.map(x=>'<span><i style="background:'+x[2]+'"></i>'+x[0]+' · <b>'+x[1]+'</b></span>').join('')+'</div></div>'+
 S.tabla('Por paquete de prueba (Package Number del FR)',['Paquete','#Juntas','#Soldadas','% soldado','#Sitio por soldar','#Sin VT','#Con rechazo'],Object.keys(W.p).sort().map(k=>{const q=W.p[k];return[k,q.n,q.sold,S.mb(pn(q.sold,q.n),'var(--ok)'),q.fw-q.fws,q.sv,q.rej]}),500)+
 S.tabla('Oportunidad de soldadura: isométrico 100 % montado con juntas de sitio por soldar',['Isométrico','Unidad','#De sitio por soldar','#Por soldar','% soldado'],op.map(x=>[x.iso,mo(x).un,x.fw-x.fws,x.tot-x.sold,S.mb(pn(x.sold,x.tot),'var(--ok)')]))+
 S.tabla('Isométricos con juntas de sitio por soldar (montados o no)',['Isométrico','#Juntas','#Soldadas','#De sitio por soldar','#Por soldar'],cp.map(x=>[x.iso,x.tot,x.sold,x.fw-x.fws,x.tot-x.sold]))+
 S.tabla('Por soldador (top 30, juntas soldadas)',['Soldador','#Juntas','#Con rechazo','% rechazo'],Object.keys(W.w).sort((a,b)=>W.w[b].n-W.w[a].n).slice(0,30).map(k=>[k,W.w[k].n,W.w[k].rej,S.mb(pn(W.w[k].rej,W.w[k].n),'var(--er)')]),30)+
 S.tabla('Juntas anteriores a una NCD aprobada · revisar',['Isométrico','Última junta','Última NCD','Emisión','Cambio','#NCD'],A.map(x=>[W.d[x[0]].iso,...x.slice(1)]))+'</div>':(W.info?'<div class="c mu">Datos de una versión anterior: vuelve a cargar el reporte FR.</div>':''));
 const nf=el.querySelector('#nf');if(nf)nf.accept='.xlsx,.xlsm,.xls,.csv';
 S.ligar(el,W.importar,W.vista)};
S.mods.push({id:'soldadura',nom:'Soldadura',desc:'Juntas, soldadores y rechazos (FR)',vista:W.vista,cargar:W.cargar,resumen:W.resumen});
})(typeof window!=='undefined'?window:globalThis);
