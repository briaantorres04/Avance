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
Suite.db=(m,k,v)=>new Promise((ok,no)=>{const q=indexedDB.open('suiteDB',1);q.onupgradeneeded=()=>q.result.createObjectStore('kv');
 q.onsuccess=()=>{const tx=q.result.transaction('kv',m==='g'?'readonly':'readwrite'),st=tx.objectStore('kv'),r=m==='g'?st.get(k):st.put(v,k);tx.oncomplete=()=>ok(r.result);tx.onerror=()=>no(tx.error)};q.onerror=()=>no(q.error)});
})(typeof window!=='undefined'?window:globalThis);
