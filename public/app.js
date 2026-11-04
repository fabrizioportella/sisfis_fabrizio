const ids=['expediente','empresa','placa','supervisor','antecedentes',

'analisis','descargos','conclusiones','recomendaciones'];

let editId=null;

const $=id=>document.getElementById(id);

$('tieneDescargo').onchange=()=> $('zonaDescargo').hidden=!$('tieneDescargo').checked;

function datos(){const x={};ids.forEach(k=>x[k]=$(k).value);

x.tieneDescargo=$('tieneDescargo').checked;return x;}

function cargar(x){ids.forEach(k=>$(k).value=x[k]||'');

$('tieneDescargo').checked=!!x.tieneDescargo;

$('zonaDescargo').hidden=!x.tieneDescargo;editId=x.id;vista(x);}

function numerar(s){return (s||'Sin información').split('\n').filter(x=>x.trim())

.map((x,i)=>`${i+1}. ${x}`).join('\n');}

function sinNumeracion(s){return (s||'Sin información').split('\n').filter(x=>x.trim())

.map(x=>x.replace(/^\s*\d+(?:\.\d+)*\.?\s*/, '')).join('\n')||'Sin información';}

function vista(x){let n=1;const seccion=(t,s,numerada=true)=>
`${n++}. ${t}\n${numerada?numerar(s):sinNumeracion(s)}\n\n`;

$('vista').textContent=`INFORME DE FISCALIZACIÓN\nExpediente: ${x.expediente}\n`

+`Empresa: ${x.empresa} | Placa: ${x.placa||'-'}\nSupervisor: ${x.supervisor||'-'}\n\n`

+seccion('ANTECEDENTES',x.antecedentes,false)+seccion('ANÁLISIS',x.analisis)

+(x.tieneDescargo?seccion('ANÁLISIS DE DESCARGOS',x.descargos):'')

+seccion('CONCLUSIONES',x.conclusiones)+seccion('RECOMENDACIONES',x.recomendaciones,false);}

async function listar(){const r=await fetch('/api/informes');const a=await r.json();

$('lista').replaceChildren();a.forEach(x=>{const p=document.createElement('p');

p.textContent=`${x.expediente} – ${x.empresa} `;

const e=document.createElement('button');e.textContent='Editar';e.onclick=()=>cargar(x);

const d=document.createElement('a');d.className='button-link';d.textContent='Descargar Word';
d.href='/api/informes/'+encodeURIComponent(x.id)+'/docx';

const b=document.createElement('button');b.textContent='Eliminar';

b.onclick=async()=>{if(confirm('¿Eliminar?')){await fetch('/api/informes/'+x.id,{method:'DELETE'});listar();}};

p.append(e,d,b);$('lista').append(p);});}

$('form').onsubmit=async ev=>{ev.preventDefault();const x=datos();

const r=await fetch('/api/informes'+(editId?'/'+editId:''),{

method:editId?'PUT':'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(x)});

if(!r.ok){alert('Revise expediente y empresa');return;}

const saved=await r.json();vista(saved);editId=null;$('form').reset();

$('zonaDescargo').hidden=true;listar();};listar();