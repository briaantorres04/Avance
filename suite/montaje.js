/* MÓDULO · MONTAJE: metros por unidad, oportunidades (isométrico completo disponible), complejidad por peso y asignación de equipos */
(function(g){const S=g.Suite,esc=S.esc,has=v=>v!=null&&String(v).trim()!=='',num=v=>+String(v).replace(/,/g,'')||0,r0=v=>Math.round(v).toLocaleString('es-MX'),
M=S.montaje={H:[],R:[],cfg:null,d:null,info:'',eq:[],as:{},th:[250,1000,3000],t:'res',un:'',fp:'',sel:''};
if(typeof document!=='undefined'){const s=document.createElement('style');s.textContent='input,select{background:var(--s2);color:var(--tx);border:1px solid var(--bd);padding:8px;font:inherit;max-width:100%}.f2{display:flex;gap:8px;flex-wrap:wrap;margin:12px 0}.f2 button{flex:1 1 120px}.stk{display:flex;height:14px;border:1px solid var(--bd);min-width:160px}.stk i{display:block}.fl{display:flex;gap:10px;flex-wrap:wrap;align-items:flex-end;margin-top:8px}';document.head.appendChild(s)}
/* Disponible = la columna de almacén trae dato (fecha o texto) distinto de No / Pendiente / Falta */
const dis=v=>has(v)&&!/^(no|n\/?a|pend|falta|-|0)/i.test(String(v).trim());
M.parse=wb=>{const A=S.filas(wb.Sheets[wb.SheetNames.find(n=>/alineaci/i.test(n))||wb.SheetNames[0]]),hi=A.findIndex(r=>r.includes('Isometrico')&&r.includes('Fecha de Montaje'));
 if(hi<0)throw new Error('No encontré las columnas Isometrico y Fecha de Montaje. Elige el Excel del avance (hoja Alineacion).');
 const H=A[hi].map(c=>c==null?'':String(c)),f=['Isometrico','Planta','Unidad','Mts','Prioridad Base Iso','Fecha de Montaje','Fecha Alineacion'].filter(n=>H.indexOf(n)<0);
 if(f.length)throw new Error('Al Excel le faltan estas columnas: '+f.join(', ')+'.');
 return{H,R:A.slice(hi+1).map(r=>H.map((_,i)=>{const c=r[i];return c instanceof Date?S.fecha(c):c==null?'':c}))}};
M.auto=()=>{const H=M.H.map(h=>S.norm(h).toLowerCase()),f=re=>H.findIndex(h=>re.test(h)),i=n=>M.H.indexOf(n);
 M.cfg={iso:i('Isometrico'),pl:i('Planta'),un:i('Unidad'),mt:i('Mts'),pr:i('Prioridad Base Iso'),fm:i('Fecha de Montaje'),fa:i('Fecha Alineacion'),peso:f(/peso|weight|\bkg\b/),alm:f(/almacen/),sp:f(/spool|pieza|mark/)}};
M.calc=()=>{if(!M.H.length||!M.cfg){M.d=null;return}const c=M.cfg,d={};
 M.R.forEach(r=>{const k=has(r[c.iso])&&S.cruce(r[c.iso]);if(!k)return;
  const t=d[k]=d[k]||{k,iso:String(r[c.iso]),pl:r[c.pl],un:String(r[c.un]),pr:String(r[c.pr]),sp:[],mt:0,mm:0,fm:'',a:0},
  x={id:c.sp>=0&&has(r[c.sp])?String(r[c.sp]):'#'+(t.sp.length+1),mt:num(r[c.mt]),m:has(r[c.fm]),w:c.peso>=0?num(r[c.peso]):0,av:c.alm>=0&&dis(r[c.alm])};
  t.sp.push(x);t.mt+=x.mt;if(x.m){t.mm+=x.mt;const f=S.fecha(r[c.fm]);if(f&&f>t.fm)t.fm=f}if(has(r[c.fa]))t.a++});
 Object.values(d).forEach(t=>{const p=t.sp.filter(x=>!x.m);t.n=t.sp.length;t.nm=t.n-p.length;t.nd=p.filter(x=>x.av).length;t.np=p.length-t.nd;
  t.w=Math.max(0,...p.map(x=>x.w));t.mp=t.mt-t.mm;
  /* Criterio: oportunidad = todos los spools pendientes ya están disponibles (cualquier mezcla de montados y en almacén) */
  t.e=!p.length?'Montado':c.alm<0?'Sin dato de almacén':!t.np?'Oportunidad':t.nd?'Parcial':'Sin material';
  t.cx=c.peso<0||!p.length?'—':t.w<M.th[0]?'Baja':t.w<M.th[1]?'Media':t.w<M.th[2]?'Alta':'Crítica'});
 M.d=d};
M.sug=t=>{if(M.cfg.peso<0||!t.w)return'—';if(t.w<M.th[0])return'Manual';const e=M.eq.filter(x=>x.tp==='Izaje'&&x.est!=='Mantenimiento'&&x.c>=t.w).sort((a,b)=>a.c-b.c)[0];return e?e.nom+' ('+e.c+' kg)':'Sin equipo con capacidad'};
M.guardar=()=>S.db('s','montaje_cfg',{cfg:M.cfg,eq:M.eq,as:M.as,th:M.th}).catch(()=>{});
M.cargar=async()=>{try{const v=await S.db('g','montaje');if(v&&v.v===2){M.H=v.H;M.R=v.R;M.info=v.info}const c=await S.db('g','montaje_cfg');if(c)Object.assign(M,c)}catch(e){}M.calc();return M.d?Object.keys(M.d).length:0};
M.importar=async f=>{const o=M.parse(await S.xl(f));M.H=o.H;M.R=o.R;M.info='Avance: '+f.name+' · cargado el '+S.hoy();M.auto();M.calc();await S.db('s','montaje',{v:2,H:M.H,R:M.R,info:M.info});M.guardar()};
M.resumen=k=>{const x=M.d&&M.d[k];return x?x.nm+' de '+x.n+' spools montados'+(x.fm?' · último '+x.fm:'')+' · disponibles '+x.nd+' · faltan '+x.np+' · '+x.e+(x.cx!=='—'?' · complejidad '+x.cx:'')+' · '+Math.round(x.mt*10)/10+' m':''};
const td=v=>'<td>'+esc(v)+'</td>',tb=(t,H,R)=>'<div class="c"><b>'+t+'</b>'+(R.length?'<div class="tw"><table><tr>'+H.map(h=>'<th>'+h+'</th>').join('')+'</tr>'+R.join('')+'</table></div>':'<p class="mu">Sin casos.</p>')+'</div>',
 btn=(a,l,x,p)=>'<button '+(p?'class="p" ':'')+'data-a="'+a+'" '+(x||'')+'>'+l+'</button>';
M.res=()=>{const L=Object.values(M.d),c=M.cfg,U={},F={};
 L.forEach(t=>{const u=U[t.un]=U[t.un]||{n:0,sp:0,mt:0,mm:0,op:0};u.n++;u.sp+=t.n;u.mt+=t.mt;u.mm+=t.mm;if(t.e==='Oportunidad')u.op+=t.mp;if(t.fm)F[t.k]=t.fm});
 const A=S.cambios.alertasMod('montaje',F),n=e=>L.filter(t=>t.e===e).length,op=L.filter(t=>t.e==='Oportunidad').reduce((a,t)=>a+t.mp,0),
 sel=(f,l)=>'<label class="mu">'+l+'<br><select data-c="cfg" data-f="'+f+'"><option value="-1">— ninguna —</option>'+M.H.map((h,i)=>'<option value="'+i+'"'+(c[f]===i?' selected':'')+'>'+esc(h)+'</option>').join('')+'</select></label>';
 return(c.alm<0?'<div class="c er">Falta la columna de almacén: elige abajo la que indica si el spool está disponible. Sin ella no se pueden detectar oportunidades.</div>':'')+
 '<div class="kp">'+S.kp('Isométricos',L.length)+S.kp('Montados al 100 %',n('Montado'),'ok')+S.kp('En oportunidad',n('Oportunidad'),'wa')+S.kp('Metros en oportunidad',r0(op),'wa')+S.kp('Parciales',n('Parcial'))+S.kp('Sin material',n('Sin material'),'er')+S.kp('Revisar por NCD',A.length,A.length?'er':'ok')+'</div>'+
 '<div class="c"><b>Metros por unidad</b><div class="mu">Verde: montado · Amarillo: oportunidad (isométrico completo disponible) · Rojo: pendiente sin material completo</div><div class="tw"><table><tr><th>Unidad</th><th>Isométricos</th><th>Spools</th><th>Metros</th><th></th><th>% montado</th><th>m en oportunidad</th></tr>'+
 Object.keys(U).sort().map(k=>{const u=U[k],p=v=>Math.round(100*v/(u.mt||1));return'<tr>'+td(k)+td(u.n)+td(u.sp)+td(r0(u.mt))+'<td><div class="stk"><i style="width:'+p(u.mm)+'%;background:var(--ok)"></i><i style="width:'+p(u.op)+'%;background:var(--wa)"></i><i style="flex:1;background:var(--er);opacity:.55"></i></div></td>'+td(p(u.mm)+' %')+td(r0(u.op))+'</tr>'}).join('')+'</table></div></div>'+
 '<div class="c"><b>Columnas del avance que alimentan el análisis</b><div class="mu">Disponible = la columna de almacén trae dato (fecha o texto) distinto de No, Pendiente o Falta.</div><div class="fl">'+sel('peso','Peso del spool (kg)')+sel('alm','Disponibilidad en almacén')+sel('sp','Número de spool')+'</div>'+
 '<div class="fl"><span class="mu">Complejidad por peso del spool pendiente más pesado (kg): Baja menor a</span>'+M.th.map((v,i)=>'<input type="number" style="width:90px" data-c="th" data-i="'+i+'" value="'+v+'">').join('<span class="mu">Media/Alta/Crítica desde</span>')+'</div></div>'+
 S.tabla('Montaje anterior a una NCD aprobada · revisar',['Isométrico','Último montaje','Última NCD','Emisión','Cambio','NCD'],A.map(x=>[M.d[x[0]].iso,...x.slice(1)]))};
M.opo=()=>{const L=Object.values(M.d),us=[...new Set(L.map(t=>t.un))].sort(),f=t=>!M.un||t.un===M.un,
 O=L.filter(t=>t.e==='Oportunidad'&&f(t)).sort((a,b)=>a.pr.localeCompare(b.pr)||b.mp-a.mp),P=L.filter(t=>t.e==='Parcial'&&f(t)).sort((a,b)=>a.np-b.np||b.mp-a.mp).slice(0,100),
 row=(t,asg)=>{const a=M.as[t.k];return'<tr>'+td(t.iso)+td(t.pr)+td(t.nm+' / '+t.nd+' / '+t.np)+td(Math.round(t.mp*10)/10)+td(t.w?r0(t.w):'—')+td(t.cx)+td(M.sug(t))+(asg?td(a?a.f:'—')+'<td>'+btn('asig','Asignar','data-k="'+esc(t.k)+'"')+'</td>':'')+'</tr>'},
 H=['Isométrico','Prior.','Spools montados / disponibles / faltan','m pendientes','Peso máx. (kg)','Complejidad','Equipo sugerido'],
 e=M.sel&&M.d[M.sel],iz=M.eq.filter(x=>x.tp==='Izaje'&&x.est!=='Mantenimiento'),pl=M.eq.filter(x=>x.tp==='Plataforma'&&x.est!=='Mantenimiento'),a0=M.as[M.sel]||{};
 return'<div class="c"><label class="mu">Unidad<br><select data-c="un"><option value="">Todas</option>'+us.map(u=>'<option'+(u===M.un?' selected':'')+'>'+esc(u)+'</option>').join('')+'</select></label><div class="mu" style="margin-top:6px">Oportunidad = todos los spools pendientes ya están en almacén (cualquier mezcla de montados y disponibles).</div></div>'+
 (e?'<div class="c" style="border-color:var(--ac)"><b>Asignar '+esc(e.iso)+'</b> <span class="mu">'+r0(e.w)+' kg el spool más pesado</span><div class="fl"><label class="mu">Fecha<br><input id="as_f" type="date" value="'+(a0.f||M.fp||S.hoy())+'"></label><label class="mu">Equipo de izaje<br><select id="as_iz"><option value="">Sin equipo</option>'+iz.map(x=>'<option value="'+x.id+'"'+(a0.iz===x.id?' selected':'')+'>'+esc(x.nom)+' · '+x.c+' kg'+(x.c<e.w?' ⚠':'')+'</option>').join('')+'</select></label><label class="mu">Plataforma<br><select id="as_pl"><option value="">Sin plataforma</option>'+pl.map(x=>'<option value="'+x.id+'"'+(a0.pl===x.id?' selected':'')+'>'+esc(x.nom)+(x.c?' · '+x.c+' m':'')+'</option>').join('')+'</select></label>'+btn('guardar','Asignar','',1)+btn('cancel','Cancelar')+(M.as[M.sel]?btn('quitar','Quitar','data-k="'+esc(M.sel)+'"'):'')+'</div></div>':'')+
 tb('En oportunidad ('+O.length+')',[...H,'Asignado',''],O.slice(0,300).map(t=>row(t,1)))+tb('Casi listos: faltan spools por llegar al almacén',H,P.map(t=>row(t,0)))};
M.eqv=()=>'<div class="c"><b>Equipos de izaje y plataformas</b>'+M.eq.map(x=>'<div style="padding:8px 0;border-bottom:1px solid var(--bd)"><b>'+esc(x.nom)+'</b> <span class="mu">'+x.tp+' · '+(x.tp==='Izaje'?x.c+' kg':x.c?x.c+' m':'')+' · '+x.est+'</span> '+btn('eqmt',x.est==='Mantenimiento'?'Marcar disponible':'Mantenimiento','data-id="'+x.id+'"')+' '+btn('eqdel','Quitar','data-id="'+x.id+'"')+'</div>').join('')+
 '<div class="fl"><input id="eq_n" placeholder="Nombre o clave"><select id="eq_t"><option>Izaje</option><option>Plataforma</option></select><input id="eq_c" type="number" placeholder="Capacidad kg (izaje) o altura m"></div><div class="fl">'+btn('eqadd','Agregar equipo','',1)+btn('eqimp','Importar plataformas de Alineación')+'</div></div>';
M.prg=()=>{const f=M.fp||S.hoy(),L=Object.keys(M.as).filter(k=>M.as[k].f===f&&M.d[k]),use={},nm=id=>(M.eq.find(x=>x.id===id)||{}).nom||'—';
 L.forEach(k=>['iz','pl'].forEach(x=>{const id=M.as[k][x];if(id)(use[id]=use[id]||[]).push(k)}));
 const av=(k,x)=>{const id=M.as[k][x],e=M.eq.find(y=>y.id===id),t=M.d[k];return id?[nm(id),e&&x==='iz'&&e.c<t.w?'⚠ excede capacidad':'',use[id].length>1?'⚠ compartido con '+(use[id].length-1)+' iso(s)':''].filter(Boolean).join(' · '):'—'};
 return'<div class="c"><div class="fl"><label class="mu">Día<br><input type="date" data-c="fp" value="'+f+'"></label>'+btn('share','Compartir programa','',1)+'</div><div class="mu" style="margin-top:6px">'+L.length+' isométricos · '+r0(L.reduce((a,k)=>a+M.d[k].mp,0))+' m pendientes</div></div>'+
 tb('Actividades del día',['Isométrico','Unidad','m','Peso máx. (kg)','Izaje','Plataforma',''],L.map(k=>{const t=M.d[k];return'<tr>'+td(t.iso)+td(t.un)+td(Math.round(t.mp*10)/10)+td(r0(t.w))+td(av(k,'iz'))+td(av(k,'pl'))+'<td>'+btn('quitar','Quitar','data-k="'+esc(k)+'"')+'</td></tr>'}))};
M.share=()=>{const f=M.fp||S.hoy(),nm=id=>(M.eq.find(x=>x.id===id)||{}).nom||'sin equipo';let t='Programa de montaje '+f+'\n';
 Object.keys(M.as).filter(k=>M.as[k].f===f&&M.d[k]).forEach(k=>{const a=M.as[k],d=M.d[k];t+='\n• '+d.iso+' · '+Math.round(d.mp*10)/10+' m · '+nm(a.iz)+' · '+nm(a.pl)});
 if(navigator.share)navigator.share({text:t}).catch(()=>{});else if(navigator.clipboard)navigator.clipboard.writeText(t).then(()=>alert('Programa copiado.'));else prompt('Copia el programa:',t)};
M.acc=(a,b,el)=>{const q=i=>el.querySelector(i);
 if(a==='tab')M.t=b.t;else if(a==='asig')M.sel=b.k;else if(a==='cancel')M.sel='';
 else if(a==='guardar'){const f=q('#as_f').value||S.hoy();M.as[M.sel]={f,iz:q('#as_iz').value,pl:q('#as_pl').value};M.sel='';M.fp=f;M.guardar()}
 else if(a==='quitar'){delete M.as[b.k||M.sel];M.sel='';M.guardar()}
 else if(a==='eqadd'){const n=q('#eq_n').value.trim();if(!n)return;M.eq.push({id:'e'+Date.now(),nom:n,tp:q('#eq_t').value,c:+q('#eq_c').value||0,est:'Disponible'});M.guardar()}
 else if(a==='eqmt'){const e=M.eq.find(x=>x.id===b.id);e.est=e.est==='Mantenimiento'?'Disponible':'Mantenimiento';M.guardar()}
 else if(a==='eqdel'){M.eq=M.eq.filter(x=>x.id!==b.id);M.guardar()}
 else if(a==='eqimp'){let E=[];try{E=JSON.parse(localStorage.getItem('evalAlin1e'))||[]}catch(e){}E.forEach(e=>{if(e.nom&&!M.eq.some(x=>x.nom===e.nom))M.eq.push({id:'a'+e.id,nom:e.nom,tp:'Plataforma',c:+e.alt||0,est:e.est||'Disponible'})});M.guardar()}
 else if(a==='share'){M.share();return}
 M.vista(el)};
M.vista=el=>{const tabs=[['res','Resumen'],['opo','Oportunidades'],['eq','Equipos'],['prg','Programa']];
 el.innerHTML=S.cab('Montaje',M.info,'Cargar avance de montaje')+(M.d?'<div class="f2">'+tabs.map(t=>btn('tab',t[1],'data-t="'+t[0]+'"',M.t===t[0])).join('')+'</div>'+({res:M.res,opo:M.opo,eq:M.eqv,prg:M.prg}[M.t]||M.res)():'');
 el.onclick=e=>{const b=e.target.closest('[data-a]');if(b)M.acc(b.dataset.a,b.dataset,el)};
 el.onchange=e=>{const t=e.target,c=t.dataset.c;if(!c)return;
  if(c==='cfg'){M.cfg[t.dataset.f]=+t.value;M.calc();M.guardar()}else if(c==='th'){const i=+t.dataset.i;M.th[i]=+t.value||M.th[i];M.calc();M.guardar()}else if(c==='un')M.un=t.value;else if(c==='fp')M.fp=t.value||S.hoy();
  M.vista(el)};
 S.ligar(el,M.importar,M.vista)};
S.mods.push({id:'montaje',nom:'Montaje',desc:'Metros, oportunidades y equipos de izaje',vista:M.vista,cargar:M.cargar,resumen:M.resumen});
})(typeof window!=='undefined'?window:globalThis);
