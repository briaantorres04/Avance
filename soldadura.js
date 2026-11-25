/* MÓDULO · SOLDADURA (reporte FR de Weld Console: solo juntas reportadas, sin universo total) */
(function(g){const S=g.Suite,W=S.soldadura={d:{},w:{},p:{},info:''},AC=/ACCEPT|ACEPT/;
W.parse=wb=>{const A=S.filas(wb.Sheets[wb.SheetNames[0]]),hi=A.findIndex(r=>r.some(c=>S.norm(c).toLowerCase()==='drawing no'));
 if(hi<0)throw new Error('No encontré la columna "Drawing No". Elige el reporte FR de Weld Console.');
 const H=A[hi].map(c=>S.norm(c).toLowerCase()),c=n=>H.indexOf(n),I={d:c('drawing no'),l:c('location'),f:c('fab date'),w:c('balance welders'),rp:c('repair'),pk:c('package number'),cm:c('commodity'),ln:c('line'),sh:c('sheet')},
 R=['vt result','rt results','pt result','ht result','pwht result'].map(c),d={},w={},p={};
 for(const r of A.slice(hi+1)){if(!r[I.d])continue;const k=S.cruce(r[I.d])||(I.cm>=0?[r[I.cm],r[I.ln],r[I.sh]].join('-'):null);if(!k)continue;
  const f=S.fecha(r[I.f]),t=d[k]=d[k]||{iso:String(r[I.d]),fw:0,sw:0,vt:0,rej:0,rep:0,last:''},rs=R.map(i=>i>=0?String(r[i]||'').toUpperCase():''),
  rej=rs.some(x=>/REJ|RECH/.test(x)),fw=String(r[I.l]).trim().toUpperCase()==='FW',vt=AC.test(rs[0]);
  fw?t.fw++:t.sw++;if(vt)t.vt++;if(rej)t.rej++;if(+r[I.rp]>0)t.rep++;if(f&&f>t.last)t.last=f;
  const pk=I.pk>=0&&r[I.pk]?String(r[I.pk]).trim():'';if(pk){const q=p[pk]=p[pk]||{n:0,fw:0,vt:0,rej:0};q.n++;if(fw)q.fw++;if(vt)q.vt++;if(rej)q.rej++}
  const wl=I.w>=0?String(r[I.w]||'').trim().toUpperCase():'';if(wl&&wl!=='N/A'){const x=w[wl]=w[wl]||{n:0,rej:0};x.n++;if(rej)x.rej++}}
 return{d,w,p}};
W.cargar=async()=>{try{const v=await S.db('g','soldadura');if(v&&v.d){Object.assign(W,v)}}catch(e){}return Object.keys(W.d).length};
W.importar=async f=>{const o=W.parse(await S.xl(f));Object.assign(W,o,{info:'FR: '+f.name+' · cargado el '+S.hoy()});await S.db('s','soldadura',{d:W.d,w:W.w,p:W.p,info:W.info})};
W.resumen=k=>{const x=W.d[k];return x?(x.fw+x.sw)+' juntas reportadas ('+x.fw+' de sitio) · VT aceptadas '+x.vt+' · rechazos '+x.rej+' · reparadas '+x.rep+(x.last?' · última '+x.last:''):''};
W.vista=el=>{const L=Object.values(W.d),T=f=>L.reduce((a,x)=>a+x[f],0),F={};Object.keys(W.d).forEach(k=>{if(W.d[k].last)F[k]=W.d[k].last});
 const A=S.cambios.alertasMod('soldadura',F),pc=(a,b)=>Math.round(1000*a/(b||1))/10+' %';
 el.innerHTML=S.cab('Soldadura',W.info,'Cargar reporte FR')+(L.length?'<div class="kp">'+S.kp('Isométricos con juntas',L.length)+S.kp('Juntas reportadas',T('fw')+T('sw'))+S.kp('De sitio (FW)',T('fw'))+S.kp('VT aceptadas',T('vt'),'ok')+S.kp('Con rechazo',T('rej'),T('rej')?'er':'ok')+S.kp('Reparadas',T('rep'),'wa')+S.kp('Revisar por NCD',A.length,A.length?'er':'ok')+'</div>'+
 '<div class="c mu">Solo incluye juntas ya reportadas: no hay universo total, así que no se calcula % soldado.</div>'+
 S.tabla('Por soldador (top 30)',['Soldador','Juntas','Con rechazo','% rechazo'],Object.keys(W.w).sort((a,b)=>W.w[b].n-W.w[a].n).slice(0,30).map(k=>[k,W.w[k].n,W.w[k].rej,pc(W.w[k].rej,W.w[k].n)]),30)+
 S.tabla('Por paquete de prueba (campo Package Number del FR)',['Paquete','Juntas','De sitio','VT aceptadas','Con rechazo'],Object.keys(W.p).sort().map(k=>[k,W.p[k].n,W.p[k].fw,W.p[k].vt,W.p[k].rej]))+
 S.tabla('Juntas anteriores a una NCD aprobada · revisar',['Isométrico','Última junta','Última NCD','Emisión','Cambio','NCD'],A.map(x=>[W.d[x[0]].iso,...x.slice(1)])):'');
 S.ligar(el,W.importar,W.vista)};
S.mods.push({id:'soldadura',nom:'Soldadura',desc:'Juntas, soldadores y rechazos (FR)',vista:W.vista,cargar:W.cargar,resumen:W.resumen});
})(typeof window!=='undefined'?window:globalThis);
