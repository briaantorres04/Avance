/* MÓDULO · MONTAJE (avance de montaje por isométrico: mismo Excel que "Cargar avance" de la app de alineación) */
(function(g){const S=g.Suite,M=S.montaje={d:{},info:''};
M.parse=wb=>{const A=S.filas(wb.Sheets[wb.SheetNames.find(n=>/alineaci/i.test(n))||wb.SheetNames[0]]),hi=A.findIndex(r=>r.includes('Isometrico')&&r.includes('Fecha de Montaje'));
 if(hi<0)throw new Error('No encontré las columnas Isometrico y Fecha de Montaje. Elige el Excel del avance (hoja Alineacion).');
 const H=A[hi],N=['Isometrico','Planta','Unidad','Mts','Prioridad Base Iso','Fecha de Montaje','Fecha Alineacion'],f=N.filter(n=>H.indexOf(n)<0);
 if(f.length)throw new Error('Al Excel le faltan estas columnas: '+f.join(', ')+'.');
 const [i,p,u,m,r,fm,fa]=N.map(n=>H.indexOf(n)),has=v=>v!=null&&String(v).trim()!=='',d={};
 for(const x of A.slice(hi+1)){const k=has(x[i])&&S.cruce(x[i]);if(!k)continue;
  const t=d[k]||(d[k]={iso:String(x[i]),pl:x[p],un:String(x[u]),pr:x[r],n:0,m:0,a:0,mt:0,mm:0,fm:''}),mt=typeof x[m]==='number'?x[m]:0;
  t.n++;t.mt+=mt;if(has(x[fm])){t.m++;t.mm+=mt;const f=S.fecha(x[fm]);if(f&&f>t.fm)t.fm=f}if(has(x[fa]))t.a++}
 return d};
M.cargar=async()=>{try{const v=await S.db('g','montaje');if(v&&v.d){M.d=v.d;M.info=v.info}}catch(e){}return Object.keys(M.d).length};
M.importar=async f=>{const d=M.parse(await S.xl(f)),info='Avance: '+f.name+' · cargado el '+S.hoy();M.d=d;M.info=info;await S.db('s','montaje',{d,info})};
M.resumen=k=>{const x=M.d[k];return x?x.m+' de '+x.n+' spools montados'+(x.fm?' · último '+x.fm:'')+' · '+x.a+' alineados · '+Math.round(x.mt*10)/10+' m':''};
M.vista=el=>{const L=Object.values(M.d),lis=L.filter(x=>x.m===x.n&&x.a===0&&x.mt>6),U={},F={};
 L.forEach(x=>{const u=U[x.un]=U[x.un]||[0,0,0];u[0]++;u[1]+=x.mt;u[2]+=x.mm});Object.keys(M.d).forEach(k=>{if(M.d[k].fm)F[k]=M.d[k].fm});
 const A=S.cambios.alertasMod('montaje',F),r=v=>Math.round(v).toLocaleString('es-MX');
 el.innerHTML=S.cab('Montaje',M.info,'Cargar avance de montaje')+(L.length?'<div class="kp">'+S.kp('Isométricos',L.length)+S.kp('Montados al 100 %',L.filter(x=>x.m===x.n).length,'ok')+S.kp('Listos para alinear',lis.length,'ok')+S.kp('Metros montados',r(L.reduce((a,x)=>a+x.mm,0)))+S.kp('Revisar por NCD',A.length,A.length?'er':'ok')+'</div>'+
 S.tabla('Avance por unidad',['Unidad','Isométricos','Metros','Metros montados','% montado'],Object.keys(U).sort().map(u=>[u,U[u][0],r(U[u][1]),r(U[u][2]),Math.round(100*U[u][2]/(U[u][1]||1))+' %']))+
 S.tabla('Montaje anterior a una NCD aprobada · revisar',['Isométrico','Último montaje','Última NCD','Emisión','Cambio','NCD'],A.map(x=>[M.d[x[0]].iso,...x.slice(1)])):'');
 S.ligar(el,M.importar,M.vista)};
S.mods.push({id:'montaje',nom:'Montaje',desc:'Avance de spools y metros',vista:M.vista,cargar:M.cargar,resumen:M.resumen});
})(typeof window!=='undefined'?window:globalThis);
