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
Suite.tabla=(t,H,R,max)=>'<div class="c"><b>'+t+'</b>'+(R.length?'<div class="tw"><table><tr>'+H.map(h=>'<th>'+h+'</th>').join('')+'</tr>'+R.slice(0,max||200).map(r=>'<tr>'+r.map(v=>'<td>'+Suite.esc(v)+'</td>').join('')+'</tr>').join('')+'</table></div>':'<p class="mu">Sin casos.</p>')+'</div>';
Suite.cab=(t,info,btn)=>'<div class="c"><b>'+t+'</b><div class="mu">'+Suite.esc(info||'Aún no hay datos cargados.')+'</div>'+(btn?'<p><button class="p" id="nb">'+btn+'</button><input id="nf" type="file" accept=".xlsx,.xlsm,.xls" hidden></p>':'')+'</div>';
Suite.ligar=(el,fn,vista)=>{const b=el.querySelector('#nb');if(!b)return;b.onclick=()=>el.querySelector('#nf').click();
 el.querySelector('#nf').onchange=async e=>{const f=e.target.files[0];e.target.value='';if(!f)return;try{await fn(f)}catch(x){alert(x.message||'No se pudo leer el archivo.')}vista(el)}};
})(typeof window!=='undefined'?window:globalThis);
