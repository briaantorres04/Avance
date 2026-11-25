/* MÓDULO · SOPORTES (lee el checklist que ya cargas en la app de alineación con "Cargar soportes") */
(function(g){const S=g.Suite,P=S.soportes={d:{},info:''};
const EXC=['5S1','5S2','5LP','5DS2','5DS1','5ST','5DA1C','5DA2C','5DA3C'],isEx=c=>EXC.some(p=>String(c).toUpperCase().startsWith(p));
P.set=(sp,info)=>{P.info=info||'';P.d={};for(const key of Object.keys(sp)){const k=S.cruce(key)||key,it=sp[key].filter(x=>!(x[10]&&isEx(x[0])));if(!it.length)continue;
 const ok=x=>x[2]==='Instalado'||x[2]==='Liberado',pe=it.filter(x=>x[2]==='Pendiente');
 P.d[k]={iso:key,n:it.length,ok:it.filter(ok).length,pend:pe.length,mat:pe.filter(x=>x[7]).length,rech:it.filter(x=>x[2]==='Rechazado').length,ult:it.map(x=>S.fecha(x[11])).filter(Boolean).sort().pop()||''}}};
P.cargar=async()=>{try{const v=await S.db('g','sp',null,'alinSop');if(v&&v.d)P.set(v.d,v.i)}catch(e){}return Object.keys(P.d).length};
P.resumen=k=>{const x=P.d[k];return x?x.ok+' de '+x.n+' instalados · '+x.pend+' pendientes ('+x.mat+' con material recibido) · '+x.rech+' rechazados'+(x.ult?' · checklist al '+x.ult:''):''};
P.vista=el=>{const L=Object.values(P.d),T=f=>L.reduce((a,x)=>a+x[f],0),F={};L.forEach(x=>{if(x.ult)F[S.cruce(x.iso)||x.iso]=x.ult});
 const A=S.cambios.alertasMod('soportes',F);
 el.innerHTML=S.cab('Soportes',P.info||'',null)+(L.length?'<div class="kp">'+S.kp('Isométricos',L.length)+S.kp('Soportes requeridos',T('n'))+S.kp('Instalados o liberados',T('ok'),'ok')+S.kp('Pendientes',T('pend'),'wa')+S.kp('Pendientes con material',T('mat'),'wa')+S.kp('Rechazados',T('rech'),T('rech')?'er':'ok')+S.kp('Revisar por NCD',A.length,A.length?'er':'ok')+'</div>'+
 S.tabla('Por instalar con material ya recibido',['Isométrico','Pendientes','Con material','Rechazados','Checklist al'],L.filter(x=>x.mat>0).sort((a,b)=>b.mat-a.mat).map(x=>[x.iso,x.pend,x.mat,x.rech,x.ult||'—']))+
 S.tabla('Checklist anterior a una NCD aprobada · revisar',['Isométrico','Checklist al','Última NCD','Emisión','Cambio','NCD'],A):'<div class="c mu">Sin datos de soportes en este dispositivo. En la app de alineación usa “Cargar soportes” con el CSV del checklist y vuelve aquí.</div>')};
S.mods.push({id:'soportes',nom:'Soportes',desc:'Checklist, pendientes y NCD',vista:P.vista,cargar:P.cargar,resumen:P.resumen});
})(typeof window!=='undefined'?window:globalThis);
