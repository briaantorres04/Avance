/* MÓDULO · CONSULTA POR ISOMÉTRICO: junta el estado de todos los módulos con la llave de cruce */
(function(g){const S=g.Suite;
S.mods.push({id:'alin',resumen:k=>{let T={};try{T=JSON.parse(localStorage.getItem('evalAlin1'))||{}}catch(e){}
 const i=Object.keys(T).find(x=>S.cruce(x)===k);return i?(T[i].res||'')+' · evaluado el '+(T[i].fecha||'—')+(T[i].est?' · '+T[i].est:''):''}});
S.mods.push({id:'consulta',nom:'Consulta por isométrico',desc:'Estado en todos los módulos',vista:el=>{
 el.innerHTML='<div class="ui"><div class="c"><b>Consulta por isométrico</b><div class="mu" style="margin:2px 0 8px">Escribe la llave corta (por ejemplo AP-334065-01) o el código completo.</div><input id="q" placeholder="Ej. AP-334065-01" style="width:100%;padding:12px;font-size:16px"></div><div id="r"></div></div>';
 const q=el.querySelector('#q'),r=el.querySelector('#r'),nom={alin:'Alineación'};
 q.oninput=()=>{const k=S.cruce(q.value)||q.value.trim().toUpperCase();if(!k){r.innerHTML='';return}
  const L=S.mods.filter(m=>m.resumen).map(m=>[nom[m.id]||m.nom,m.resumen(k)]).filter(x=>x[1]),N=S.cambios.de(k).slice().sort((a,b)=>(b.e||'').localeCompare(a.e||''));
  r.innerHTML='<div class="ui">'+(L.length||N.length?S.tabla('Estado de '+S.esc(k),['Módulo','Estado'],L.map(x=>[{h:S.bd(x[0],'var(--ac)')},{h:'<span style="white-space:normal">'+S.esc(x[1])+'</span>'}]))+S.tabla('Cambios de ingeniería',['NCD','Emisión','Cambio'],N.map(x=>[x.n,x.e||'sin fecha',x.t])):'<div class="c mu">Sin datos para '+S.esc(k)+'.</div>')+'</div>'}}});
})(typeof window!=='undefined'?window:globalThis);
