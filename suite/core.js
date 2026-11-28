/* SUITE DE SUPERVISIÓN · NÚCLEO COMPARTIDO (lo usan todos los módulos y la app de alineación) */
(function(g){
const Suite=g.Suite=g.Suite||{};
Suite.mods=[];
/* Configuración por planta: el prefijo del isométrico cambia, la llave de cruce no */
Suite.cfg={prefijos:{'406':'ARRMHP2'}};
Suite.esc=s=>String(s==null?'':s).replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
Suite.norm=s=>String(s==null?'':s).normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/\s+/g,' ').trim();
Suite.hoy=()=>new Date().toLocaleDateString('en-CA');
/* Llave de cruce SISTEMA-SECUENCIA-HOJA (ej. AP-334065-01) desde cualquier formato de isométrico.
   Devuelve null si el documento no es un isométrico (p. ej. un plano DTI). */
Suite.cruce=s=>{const m=String(s==null?'':s).trim().toUpperCase().match(/([A-Z]+-\d{4,7}-\d{2})$/);return m?m[1]:null};
/* Cualquier fecha (Date, serial de Excel, texto) a AAAA-MM-DD; null si no es válida (años fuera de 2000-2100) */
const iso=(y,m,d)=>y>=2000&&y<=2100?y+'-'+String(m).padStart(2,'0')+'-'+String(d).padStart(2,'0'):null;
Suite.fecha=v=>{
 if(v==null||v==='')return null;
 if(v instanceof Date)return isNaN(v)?null:iso(v.getFullYear(),v.getMonth()+1,v.getDate());
 if(typeof v==='number'){if(v<30000)return null;const d=new Date(Math.round((v-25569)*864e5));return iso(d.getUTCFullYear(),d.getUTCMonth()+1,d.getUTCDate())}
 const t=String(v).trim();if(!t)return null;
 if(/^\d{4}-\d{2}-\d{2}/.test(t))return iso(+t.slice(0,4),+t.slice(5,7),+t.slice(8,10));
 const d=new Date(t);return isNaN(d)?null:iso(d.getFullYear(),d.getMonth()+1,d.getDate());
};
/* Almacenamiento compartido en IndexedDB (BD "suiteDB"): db('g',clave) lee, db('s',clave,valor) guarda */
Suite.db=(m,k,v,n)=>new Promise((ok,no)=>{const q=indexedDB.open(n||'suiteDB',1);q.onupgradeneeded=()=>q.result.createObjectStore('kv');
 q.onsuccess=()=>{const tx=q.result.transaction('kv',m==='g'?'readonly':'readwrite'),st=tx.objectStore('kv'),r=m==='g'?st.get(k):st.put(v,k);tx.oncomplete=()=>ok(r.result);tx.onerror=()=>no(tx.error)};q.onerror=()=>no(q.error)});
/* Utilidades de módulos: lectura de Excel, tablas, KPIs y botón de carga */
Suite.filas=ws=>XLSX.utils.sheet_to_json(ws,{header:1,raw:true,defval:null,blankrows:false});
Suite.xl=async f=>XLSX.read(await f.arrayBuffer(),{type:'array',cellDates:true});
Suite.kp=(l,v,c)=>'<div class="c '+(c||'')+'"><span class="mu">'+l+'</span><b>'+v+'</b></div>';
/* Tabla: encabezado con # = columna numérica; celdas {h:'<html>'} sin escapar; números se alinean a la derecha */
Suite.tabla=(t,H,R,max,nota)=>'<div class="c"><b>'+t+'</b>'+(nota?'<div class="mu" style="margin:2px 0 8px">'+nota+'</div>':'')+(R.length?'<div class="tw"><table><tr>'+H.map(h=>h[0]==='#'?'<th class="n">'+h.slice(1)+'</th>':'<th>'+h+'</th>').join('')+'</tr>'+R.slice(0,max||200).map(r=>'<tr>'+r.map(v=>v&&typeof v==='object'?'<td>'+v.h+'</td>':typeof v==='number'?'<td class="n">'+v.toLocaleString('es-MX')+'</td>':'<td>'+Suite.esc(v)+'</td>').join('')+'</tr>').join('')+'</table></div>':'<p class="mu">Sin casos.</p>')+'</div>';
Suite.cab=(t,info,btn)=>'<div class="c"><b>'+t+'</b><div class="mu">'+Suite.esc(info||'Aún no hay datos cargados.')+'</div>'+(btn?'<p><button class="p" id="nb">'+btn+'</button><input id="nf" type="file" accept=".xlsx,.xlsm,.xls" hidden></p>':'')+'</div>';
Suite.ligar=(el,fn,vista)=>{const b=el.querySelector('#nb');if(!b)return;b.onclick=()=>el.querySelector('#nf').click();
 el.querySelector('#nf').onchange=async e=>{const f=e.target.files[0];e.target.value='';if(!f)return;try{await fn(f)}catch(x){alert(x.message||'No se pudo leer el archivo.')}vista(el)}};

/* Diseño compartido (clase .ui): barras, etiquetas, leyendas y minibarras */
if(typeof document!=='undefined'&&!document.getElementById('st-ui')){const s=document.createElement('style');s.id='st-ui';s.textContent='.ui .c{margin-bottom:14px}.ui table{border-collapse:separate;border-spacing:0}.ui th{position:sticky;top:0;z-index:1;background:var(--s2);font-size:11px;letter-spacing:.04em;text-transform:uppercase}.ui td,.ui th{padding:10px 12px;border-bottom:1px solid var(--bd)}.ui tr:nth-child(even) td{background:rgba(127,127,127,.07)}.ui tr:hover td{background:rgba(127,127,127,.16)}.ui .n{text-align:right;font-variant-numeric:tabular-nums}.ui .tw{max-height:560px;overflow:auto}.ui .bg{display:inline-block;padding:2px 10px;border-radius:999px;font-size:11px;font-weight:700;border:1px solid currentColor;white-space:nowrap}.ui .ch{display:inline-block;margin:2px 4px 2px 0;padding:2px 8px;border-radius:6px;font-size:11px;background:rgba(127,127,127,.18)}.ui .pip{display:inline-block;width:12px;height:12px;border-radius:50%;margin-right:5px;vertical-align:middle}.ui .hb{display:flex;align-items:center;gap:12px;margin:9px 0}.ui .hb .l{flex:0 0 210px}.ui td button{padding:5px 14px;min-height:0;font-size:12px}.ui .hb .t{flex:1;height:24px;background:var(--s2);border:1px solid var(--bd);position:relative}.ui .hb .t>i{position:absolute;left:0;top:0;bottom:0}.ui .sg{display:flex;height:100%}.ui .sg i{display:block;height:100%}.ui .hb b{flex:0 0 190px;font-family:Consolas,monospace;font-size:12px}.ui .leg{display:flex;gap:16px;flex-wrap:wrap;margin-top:10px}.ui .leg i{display:inline-block;width:12px;height:12px;margin-right:6px;vertical-align:-2px;border-radius:2px}.ui .tl{display:flex;gap:8px;flex-wrap:wrap;margin-top:10px}.ui .tl span{padding:6px 12px;border:1px solid var(--bd);border-radius:6px;font-size:12px}@media(max-width:700px){.ui .hb{flex-wrap:wrap}.ui .hb .l,.ui .hb b{flex:1 1 100%}}.ui .mb{display:inline-block;width:70px;height:8px;background:var(--s2);border:1px solid var(--bd);vertical-align:middle;margin-right:6px}.ui .mb i{display:block;height:100%}.ui input,.ui select{background:var(--s2);color:var(--tx);border:1px solid var(--bd);padding:8px;font:inherit;max-width:100%}.ui .f2{display:flex;gap:8px;flex-wrap:wrap;margin:12px 0}.ui .f2 button{flex:1 1 120px}.ui .fl{display:flex;gap:10px;flex-wrap:wrap;align-items:flex-end;margin-top:8px}';document.head.appendChild(s)}
Suite.pc=(a,b)=>Math.round(100*a/(b||1));
Suite.hb=(l,p,txt,col,seg)=>'<div class="hb"><span class="l">'+l+'</span><div class="t">'+(seg||'<i style="width:'+Math.max(0,Math.min(100,p))+'%;background:'+col+'"></i>')+'</div><b>'+txt+'</b></div>';
Suite.stk=(a,w)=>'<div class="sg" style="width:'+(w==null?100:w)+'%">'+a.map(x=>'<i title="'+Suite.esc(x[2]||'')+'" style="width:'+Math.max(0,x[0])+'%;background:'+x[1]+(x[3]?';opacity:'+x[3]:'')+'"></i>').join('')+'</div>';
Suite.bd=(t,c)=>'<span class="bg" style="color:'+(c||'var(--mu)')+'">'+Suite.esc(t)+'</span>';
Suite.ch=t=>'<span class="ch">'+Suite.esc(t)+'</span>';
Suite.mb=(p,c)=>({h:'<span class="mb"><i style="width:'+Math.max(0,Math.min(100,p))+'%;background:'+c+'"></i></span>'+p+' %'});
Suite.leg=a=>'<div class="leg">'+a.map(x=>'<span><i style="background:'+x[1]+(x[2]?';opacity:'+x[2]:'')+'"></i>'+x[0]+'</span>').join('')+'</div>';
})(typeof window!=='undefined'?window:globalThis);
