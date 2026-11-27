/* MÓDULO · PAQUETES DE PRUEBA: qué paquetes están listos para entregarse a prueba y qué les falta.
   Fuente: Excel de paquetes liberados (hojas PAQUETE GENERAL e ISOMETRICOS X PAQUETE). Cruza montaje, soportes, soldadura y NCD. */
(function(g){const S=g.Suite,esc=S.esc,r0=v=>Math.round(v).toLocaleString('es-MX'),P=S.paquetes={pk:[],li:[],inv:[],info:'',t:'res',es:'',sis:'',sel:'',R:null,mult:{}},
nk=s=>String(s==null?'':s).toUpperCase().replace(/[\s\-]+/g,''),ORD=['Listo para entregar','Entregado, sin probar','Casi lista','En proceso','Sin datos de avance','Sin isométricos'];
if(typeof document!=='undefined'&&!document.getElementById('st-sop')){const s=document.createElement('style');s.id='st-sop';s.textContent='input,select{background:var(--s2);color:var(--tx);border:1px solid var(--bd);padding:8px;font:inherit;max-width:100%}.f2{display:flex;gap:8px;flex-wrap:wrap;margin:12px 0}.f2 button{flex:1 1 120px}.stk{display:flex;height:14px;border:1px solid var(--bd);min-width:160px}.stk i{display:block}.fl{display:flex;gap:10px;flex-wrap:wrap;align-items:flex-end;margin-top:8px}';document.head.appendChild(s)}
/* Un hito cuenta si la celda trae dato; N/A, guion o vacío no cuentan */
const hx=v=>v!=null&&!/^(n\/?a|-+|0)?$/i.test(String(v).trim());
P.parse=wb=>{let G=null,L=null;const h=(A,t)=>A.findIndex(r=>r.some(c=>S.norm(c).toLowerCase()===t));
 for(const n of wb.SheetNames){const A=S.filas(wb.Sheets[n]),a=h(A,'circuito'),b=h(A,'paquete de prueba');if(a>=0&&!G)G={A,h:a};if(b>=0&&!L)L={A,h:b}}
 if(!G||!L)throw new Error('No encontré las hojas con las columnas CIRCUITO y Paquete de Prueba. Elige el Excel de paquetes liberados mecánicamente.');
 const H=G.A[G.h].map(c=>S.norm(c).toLowerCase()),c=n=>H.indexOf(n),I={s:c('sistema'),c:c('circuito'),t:c('tipo de prueba'),m:c('ml'),ph:c('ph'),lv:c('lavado'),sp:c('soplado'),en:c('fecha de entrega a ph'),pr:c('probado'),st:c('status'),ob:c('observaciones')},
 pk=G.A.slice(G.h+1).filter(r=>r[I.c]).map(r=>({id:String(r[I.c]).trim(),nk:nk(r[I.c]),sis:String(r[I.s]||'').trim(),tipo:String(r[I.t]||'').trim(),ml:+r[I.m]||0,ph:S.fecha(r[I.ph]),lav:S.fecha(r[I.lv]),sop:S.fecha(r[I.sp]),ent:S.fecha(r[I.en]),hx:{ph:hx(r[I.ph]),lv:hx(r[I.lv]),sp:hx(r[I.sp]),en:hx(r[I.en])},prob:String(r[I.pr]||'').trim().toLowerCase()==='probado',st:String(r[I.st]||'').trim().toUpperCase(),ob:String(r[I.ob]||'').trim()})),
 K=L.A[L.h].map(x=>S.norm(x).toLowerCase()),q=n=>K.indexOf(n),J={pl:q('planta'),d:q('iso description'),s:q('sistema'),p:q('paquete de prueba')},li=[],inv=[];
 L.A.slice(L.h+1).forEach(r=>{const k=S.cruce(r[J.d])||S.cruce(r[1]);if(!k)return;const p=String(r[J.p]==null?'':r[J.p]).trim();
  if(/^CP-/i.test(p))li.push({k,paq:p,nk:nk(p),pl:String(r[J.pl]||''),sis:String(r[J.s]||''),mt:+r[J.p+1]||0});else inv.push([k,p||'(vacío)'])});
 return{pk,li,inv}};
P.cargar=async()=>{try{const v=await S.db('g','paquetes');if(v&&v.v===1)Object.assign(P,{pk:v.pk,li:v.li,inv:v.inv,info:v.info})}catch(e){}return P.pk.length};
P.importar=async f=>{const o=P.parse(await S.xl(f));Object.assign(P,o,{info:'Paquetes: '+f.name+' · cargado el '+S.hoy()});await S.db('s','paquetes',{v:1,pk:P.pk,li:P.li,inv:P.inv,info:P.info})};
/* Preparación de cada paquete con lo que sepan los otros módulos (los que no tengan datos no bloquean) */
const ne=o=>o&&Object.keys(o).length?o:null;
P.src=()=>{const a=ne(S.montaje&&S.montaje.d),b=ne(S.soportes&&S.soportes.d),c=ne(S.soldadura&&S.soldadura.d);return'Datos usados: Montaje '+(a?'✓':'✗ sin cargar')+' · Soportes '+(b?'✓':'✗ sin cargar')+' · Soldadura '+(c?'✓':'✗ sin cargar')+'. Lo que no está cargado no se considera pendiente.'};
P.calc=()=>{const MD=ne(S.montaje&&S.montaje.d),SD=ne(S.soportes&&S.soportes.d),WD=ne(S.soldadura&&S.soldadura.d),Wn={},R={},mult={};
 if(S.soldadura&&S.soldadura.p)Object.keys(S.soldadura.p).forEach(k=>Wn[nk(k)]=S.soldadura.p[k]);
 P.pk.forEach(p=>R[p.nk]={...p,reg:true,isos:{}});
 P.li.forEach(l=>{const q=R[l.nk]=R[l.nk]||{id:l.paq,nk:l.nk,sis:l.sis,tipo:'',ml:0,reg:false,isos:{},prob:false};q.isos[l.k]=(q.isos[l.k]||0)+l.mt;(mult[l.k]=mult[l.k]||new Set()).add(l.nk)});
 Object.values(R).forEach(q=>{const ks=Object.keys(q.isos),w=Wn[q.nk],f=[];q.n=ks.length;q.mi=Object.values(q.isos).reduce((a,v)=>a+v,0);q.sm=q.sp=q.nc=q.ncp=q.jp=q.vt=q.rj=0;
  ks.forEach(k=>{const m=MD&&MD[k];if(m&&m.e!=='Montado')q.sm++;const s=SD&&SD[k];if(s)q.sp+=s.por;
   S.cambios.de(k).forEach(x=>{if(x.m.includes('pruebas')){q.nc++;if(q.ph&&x.e&&x.e>q.ph)q.ncp++}});
   if(!w&&WD&&WD[k]){q.jp+=WD[k].tot-WD[k].sold;q.vt+=WD[k].sv;q.rj+=WD[k].rej}});
  if(w){q.jp=w.n-w.sold;q.vt=w.sv;q.rj=w.rej}
  if(q.sm)f.push(q.sm+' isos sin montar');if(q.sp)f.push(q.sp+' soportes');if(q.jp)f.push(q.jp+' juntas por soldar');if(q.vt)f.push(q.vt+' sin VT');if(q.rj)f.push(q.rj+' con rechazo');
  q.f=f;q.cat=(q.sm?1:0)+(q.sp?1:0)+(q.jp?1:0)+(q.vt||q.rj?1:0);q.dat=!!(MD||SD||WD||w);
  q.e=q.prob?'Probado':!q.n?'Sin isométricos':q.hx&&q.hx.en?'Entregado, sin probar':!q.dat?'Sin datos de avance':!q.cat?'Listo para entregar':q.cat===1?'Casi lista':'En proceso'});
 P.R=R;P.mult=mult;return R};
P.resumen=k=>{if(!P.pk.length)return'';if(!P.R)P.calc();return P.li.filter(l=>l.k===k).map(l=>{const q=P.R[l.nk];return l.paq+' ('+(q?q.e:'—')+(q&&q.f.length?' · falta: '+q.f.join(', '):'')+')'}).join(' | ')};
const td=v=>'<td>'+esc(v)+'</td>',tb=(t,H,R)=>'<div class="c"><b>'+t+'</b>'+(R.length?'<div class="tw"><table><tr>'+H.map(h=>'<th>'+h+'</th>').join('')+'</tr>'+R.join('')+'</table></div>':'<p class="mu">Sin casos.</p>')+'</div>',
 btn=(a,l,x,p)=>'<button '+(p?'class="p" ':'')+'data-a="'+a+'" '+(x||'')+'>'+l+'</button>',ml=q=>q.reg?q.ml:q.mi;
P.res=()=>{const L=Object.values(P.R),pe=L.filter(q=>!q.prob),n=e=>pe.filter(q=>q.e===e).length,sum=(a,f)=>a.reduce((s,q)=>s+f(q),0),pc=(a,b)=>Math.round(100*a/(b||1))+' %',S2={};
 L.forEach(q=>{const u=S2[q.sis||'—']=S2[q.sis||'—']||{n:0,pr:0,ml:0,mp:0,mo:0};u.n++;u.ml+=ml(q);if(q.prob){u.pr++;u.mp+=ml(q)}else if(q.e==='Listo para entregar'||q.e==='Entregado, sin probar')u.mo+=ml(q)});
 const hit=[['Entregados a PH',q=>q.hx&&q.hx.en],['Con fecha de prueba (PH)',q=>q.hx&&q.hx.ph],['Lavados',q=>q.hx&&q.hx.lv],['Soplados',q=>q.hx&&q.hx.sp],['Probados',q=>q.prob]],
 sh=P.li.length?new Set(Object.keys(P.mult).filter(k=>P.mult[k].size>1)).size:0;
 return'<div class="c mu">'+P.src()+'</div><div class="kp">'+S.kp('Paquetes',L.length)+S.kp('Probados',L.length-pe.length,'ok')+S.kp('Pendientes',pe.length,'wa')+S.kp('ML pendientes',r0(sum(pe,ml)),'wa')+S.kp('Listos para entregar',n('Listo para entregar'),'ok')+S.kp('Entregados, sin probar',n('Entregado, sin probar'),'wa')+S.kp('Casi listas',n('Casi lista'))+S.kp('NCD posterior a la prueba',L.filter(q=>q.ncp>0).length,L.some(q=>q.ncp>0)?'er':'ok')+'</div>'+
 tb('Embudo de hitos',['Hito','Paquetes','% del total'],hit.map(h=>{const c=L.filter(h[1]).length;return'<tr>'+td(h[0])+td(c)+td(pc(c,L.length))+'</tr>'}))+
 '<div class="c"><b>Metros por sistema</b><div class="mu">Verde: probado · Amarillo: listo o entregado sin probar · Rojo: resto</div><div class="tw"><table><tr><th>Sistema</th><th>Paquetes</th><th>Probados</th><th>ML</th><th></th><th>% probado</th></tr>'+Object.keys(S2).sort().map(k=>{const u=S2[k],p=v=>Math.round(100*v/(u.ml||1));return'<tr>'+td(k)+td(u.n)+td(u.pr)+td(r0(u.ml))+'<td><div class="stk"><i style="width:'+p(u.mp)+'%;background:var(--ok)"></i><i style="width:'+p(u.mo)+'%;background:var(--wa)"></i><i style="flex:1;background:var(--er);opacity:.55"></i></div></td>'+td(p(u.mp)+' %')+'</tr>'}).join('')+'</table></div></div>'+
 tb('Calidad de los datos',['Revisión','Casos'],[['Paquetes que están en una hoja y no en la otra',L.filter(q=>!q.reg||(q.reg&&!q.n)).length],['Isométricos con paquete inválido (ENTERRADO, CANCELADA, vacío…)',P.inv.length],['Isométricos en más de un paquete (fronteras)',sh],['Paquetes con ML distinto de la suma de sus isométricos (más de 1 m)',L.filter(q=>q.reg&&q.n&&Math.abs(q.mi-q.ml)>1).length]].map(r=>'<tr>'+td(r[0])+td(r[1])+'</tr>'))};
P.opo=()=>{const L=Object.values(P.R).filter(q=>!q.prob),ss=[...new Set(L.map(q=>q.sis))].sort(),O=L.filter(q=>(!P.es||q.e===P.es)&&(!P.sis||q.sis===P.sis)).sort((a,b)=>ORD.indexOf(a.e)-ORD.indexOf(b.e)||a.f.length-b.f.length||ml(b)-ml(a));
 return'<div class="c"><div class="fl"><label class="mu">Estado<br><select data-c="es"><option value="">Todos</option>'+ORD.map(o=>'<option'+(o===P.es?' selected':'')+'>'+o+'</option>').join('')+'</select></label><label class="mu">Sistema<br><select data-c="sis"><option value="">Todos</option>'+ss.map(s=>'<option'+(s===P.sis?' selected':'')+'>'+esc(s)+'</option>').join('')+'</select></label></div><div class="mu" style="margin-top:6px">'+P.src()+' Listo = isométricos montados, soportes instalados, juntas soldadas y con VT, sin rechazos. Casi lista = falta un solo tipo de pendiente. Los módulos sin datos no bloquean.</div></div>'+
 tb('Paquetes pendientes ('+O.length+')',['Paquete','Sistema','Tipo','ML','Isos','Estado','Qué falta','NCD pruebas',''],O.slice(0,400).map(q=>'<tr>'+td(q.id)+td(q.sis)+td(q.tipo||'—')+td(r1(ml(q)))+td(q.n)+td(q.e)+'<td style="white-space:normal">'+esc(q.f.join(' · ')||'—')+'</td>'+td(q.nc?q.nc+(q.ncp?' ⚠':''):'—')+'<td>'+btn('ver','Ver','data-k="'+q.nk+'"')+'</td></tr>'))};
const r1=v=>Math.round(v*10)/10;
P.det=()=>{const q=P.R[P.sel];if(!q)return'<div class="c mu">Elige un paquete con “Ver” en Oportunidades.</div>';
 const MD=S.montaje&&S.montaje.d||{},SD=S.soportes&&S.soportes.d||{},WD=S.soldadura&&S.soldadura.d||{};
 return'<div class="c"><b>'+esc(q.id)+'</b> <span class="b">'+q.e+'</span><div class="mu">'+esc(q.sis)+' · '+(q.tipo||'sin tipo')+' · ML '+r1(ml(q))+' · Entrega a PH '+(q.ent||'—')+' · PH '+(q.ph||'—')+' · Lavado '+(q.lav||'—')+' · Soplado '+(q.sop||'—')+(q.ob?' · '+esc(q.ob):'')+'</div><div class="mu">'+(q.f.length?'Falta: '+esc(q.f.join(' · ')):'Sin pendientes detectados')+'</div></div>'+
 tb('Isométricos del paquete',['Isométrico','m','Montaje','Soportes por instalar','Juntas soldadas','NCD pruebas','También en'],Object.keys(q.isos).map(k=>{const m=MD[k],s=SD[k],w=WD[k],o=[...(P.mult[k]||[])].filter(x=>x!==q.nk).map(x=>(P.R[x]||{}).id).filter(Boolean);
  return'<tr>'+td(k)+td(r1(q.isos[k]))+td(m?m.nm+'/'+m.n+' · '+m.e:'—')+td(s?s.por:'—')+td(w?w.sold+'/'+w.tot+(w.sv?' · '+w.sv+' sin VT':'')+(w.rej?' · '+w.rej+' rech.':''):'—')+td(S.cambios.de(k).filter(x=>x.m.includes('pruebas')).length||'—')+td(o.join(', ')||'—')+'</tr>'}))};
P.acc=(a,b,el)=>{if(a==='tab')P.t=b.t;else if(a==='ver'){P.sel=b.k;P.t='det'}P.vista(el)};
P.vista=el=>{const tabs=[['res','Resumen'],['opo','Oportunidades'],['det','Detalle']];if(P.pk.length)P.calc();
 el.innerHTML=S.cab('Paquetes de prueba',P.info,'Cargar paquetes liberados')+(P.pk.length?'<div class="f2">'+tabs.map(t=>btn('tab',t[1],'data-t="'+t[0]+'"',P.t===t[0])).join('')+'</div>'+({res:P.res,opo:P.opo,det:P.det}[P.t]||P.res)():'');
 el.onclick=e=>{const b=e.target.closest('[data-a]');if(b)P.acc(b.dataset.a,b.dataset,el)};
 el.onchange=e=>{const t=e.target,c=t.dataset.c;if(!c)return;if(c==='es')P.es=t.value;else if(c==='sis')P.sis=t.value;P.vista(el)};
 S.ligar(el,P.importar,P.vista)};
S.mods.push({id:'paquetes',nom:'Paquetes de prueba',desc:'Listos para entregar y qué les falta',vista:P.vista,cargar:P.cargar,resumen:P.resumen});
})(typeof window!=='undefined'?window:globalThis);
