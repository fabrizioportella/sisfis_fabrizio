const fieldIds=[
 'expediente','fechaInforme','ciudadInforme','destinatario','supervisor','cargoSupervisor',
 'empresaSupervisora','asunto','referencia','objetivo','empresa','ruc',
 'registroHidrocarburos','actividad','placa','placaTracto','fechaRegistro','capacidad',
 'compartimientos','direccionLegal','distrito','provincia','departamento','antecedentes',
 'numeroResolucion','fechaResolucion','fechaCartaEmpresa','lugarFiscalizacion',
 'fechaAsignacion','especialistaRegional','numeroAsignacion','cartaLinea',
 'fechaConsultaRegistro','fechaVisita','horaInicio','analisis','fuenteActa',
 'nombreRecibeActa','dniRecibeActa','cargoRecibeActa','horaFin','descargos',
 'conclusiones','recomendaciones','representanteNombre','representanteDni',
 'representanteCargo','supervisorCip'
];

let editId=null;
const imageData={imagenRegistro:'',firmaRepresentante:'',firmaFiscalizador:''};
const $=id=>document.getElementById(id);
const form=$('form');
const dateFields=new Set([
 'fechaInforme','fechaRegistro','fechaResolucion','fechaCartaEmpresa','fechaAsignacion',
 'fechaConsultaRegistro','fechaVisita'
]);
const imageFields=[
 ['imagenRegistro','vistaImagenRegistro']
];
const signatureFields=['firmaRepresentante','firmaFiscalizador'];

function mostrarImagen(field,src) {
 imageData[field]=src||'';
 const previewId=imageFields.find(([key])=>key===field)[1];
 const preview=$(previewId);
 preview.src=src||'';
 preview.hidden=!src;
}

imageFields.forEach(([field])=>{
 $(field).onchange=()=>{
  const file=$(field).files[0];
  if(!file) return;
  if(!['image/jpeg','image/png','image/webp'].includes(file.type)) {
   alert('Seleccione una imagen JPG, PNG o WebP.');
   $(field).value='';
   return;
  }
  if(file.size>5*1024*1024) {
   alert('Cada imagen debe pesar como máximo 5 MB.');
   $(field).value='';
   return;
  }
  const reader=new FileReader();
  reader.onload=()=>mostrarImagen(field,reader.result);
  reader.onerror=()=>alert('No se pudo leer la imagen seleccionada.');
  reader.readAsDataURL(file);
 };
});

imageFields.forEach(([field])=>{
 document.querySelector(`[data-clear-image="${field}"]`).onclick=()=>{
  $(field).value='';
  mostrarImagen(field,'');
 };
});

function dibujarFirma(field,src) {
 const canvas=$(field);
 const context=canvas.getContext('2d');
 context.clearRect(0,0,canvas.width,canvas.height);
 imageData[field]=src||'';
 if(!src) return Promise.resolve();
 return new Promise((resolve,reject)=>{
  const image=new Image();
  image.onload=()=>{
   context.drawImage(image,0,0,canvas.width,canvas.height);
   resolve();
  };
  image.onerror=()=>reject(new Error(`No se pudo cargar la firma de ${field==='firmaRepresentante'?'representante legal':'fiscalizador'}.`));
  image.src=src;
 });
}

signatureFields.forEach(field=>{
 const canvas=$(field);
 const context=canvas.getContext('2d');
 let drawing=false;
 let hasStroke=false;
 const point=event=>{
  const bounds=canvas.getBoundingClientRect();
  return {
   x:Math.max(0,Math.min(canvas.width,(event.clientX-bounds.left)*canvas.width/bounds.width)),
   y:Math.max(0,Math.min(canvas.height,(event.clientY-bounds.top)*canvas.height/bounds.height))
  };
 };
 canvas.addEventListener('pointerdown',event=>{
  event.preventDefault();
  canvas.setPointerCapture(event.pointerId);
  const {x,y}=point(event);
  drawing=true;
  hasStroke=true;
  context.beginPath();
  context.arc(x,y,2,0,Math.PI*2);
  context.fill();
  context.beginPath();
  context.moveTo(x,y);
 });
 canvas.addEventListener('pointermove',event=>{
  if(!drawing) return;
  event.preventDefault();
  const {x,y}=point(event);
  context.lineTo(x,y);
  context.stroke();
  context.beginPath();
  context.moveTo(x,y);
 });
 const finishStroke=event=>{
  if(!drawing) return;
  event.preventDefault();
  drawing=false;
  context.beginPath();
  if(hasStroke) imageData[field]=canvas.toDataURL('image/png');
 };
 canvas.addEventListener('pointerup',finishStroke);
 canvas.addEventListener('pointercancel',finishStroke);
 context.lineCap='round';
 context.lineJoin='round';
 context.lineWidth=4;
 context.strokeStyle='#19334a';
 context.fillStyle='#19334a';
 document.querySelector(`[data-clear-signature="${field}"]`).onclick=()=>{
  context.clearRect(0,0,canvas.width,canvas.height);
  imageData[field]='';
  hasStroke=false;
 };
});

function limpiarImagenes() {
 imageFields.forEach(([field])=>{
  $(field).value='';
  mostrarImagen(field,'');
 });
 signatureFields.forEach(field=>{
  $(field).getContext('2d').clearRect(0,0,$(field).width,$(field).height);
  imageData[field]='';
 });
}

function fechaParaControl(value) {
 if(!value) return '';
 if(/^\d{4}-\d{2}-\d{2}$/.test(value)) return value;
 const months={
  enero:'01',febrero:'02',marzo:'03',abril:'04',mayo:'05',junio:'06',
  julio:'07',agosto:'08',septiembre:'09',octubre:'10',noviembre:'11',diciembre:'12'
 };
 const match=/^(\d{1,2}) de ([a-záéíóú]+) de (\d{4})$/i.exec(value.trim());
 const month=match&&months[match[2].toLocaleLowerCase('es')];
 return match&&month?`${match[3]}-${month}-${match[1].padStart(2,'0')}`:'';
}

function agregarFila(tipo,datos={}) {
 const contenedor=$(tipo);
 const fila=document.createElement('fieldset');
 fila.className='repeat-item';
 const campos=tipo==='hallazgos'
  ? [['incumplimiento','Incumplimiento normativo y base legal','textarea'],
     ['descripcion','Descripción del incumplimiento verificado','textarea'],
     ['medioProbatorio','Medio probatorio','textarea']]
  : [['descripcion','Descripción del anexo','input']];
 const leyenda=document.createElement('legend');
 leyenda.textContent=tipo==='hallazgos'
  ? `Hallazgo ${contenedor.children.length+1}` : `Anexo ${contenedor.children.length+1}`;
 fila.append(leyenda);
 campos.forEach(([key,label,elemento])=>{
  const control=document.createElement(elemento==='textarea'?'textarea':'input');
  control.id=`${tipo}-${contenedor.children.length}-${key}`;
  control.dataset.key=key;
  control.value=datos[key]||'';
  const etiqueta=document.createElement('label');
  etiqueta.htmlFor=control.id;
  etiqueta.textContent=label;
  fila.append(etiqueta,control);
 });
 const quitar=document.createElement('button');
 quitar.type='button';
 quitar.className='secondary-button';
 quitar.textContent='Quitar';
 quitar.onclick=()=>{fila.remove();renumerarFilas(tipo);};
 fila.append(quitar);
 contenedor.append(fila);
}

function renumerarFilas(tipo) {
 Array.from($(tipo).children).forEach((fila,index)=>{
  fila.querySelector('legend').textContent=tipo==='hallazgos'
   ? `Hallazgo ${index+1}` : `Anexo ${index+1}`;
 });
}

function valoresFilas(tipo) {
 return Array.from($(tipo).children).map(fila=>{
  const datos={};
  fila.querySelectorAll('[data-key]').forEach(control=>datos[control.dataset.key]=control.value);
  return datos;
 });
}

$('agregarHallazgo').onclick=()=>agregarFila('hallazgos');
$('agregarAnexo').onclick=()=>agregarFila('anexos');
$('tieneDescargo').onchange=()=> $('zonaDescargo').hidden=!$('tieneDescargo').checked;

function datos() {
 const x={};
 fieldIds.forEach(id=>x[id]=$(id).value);
 x.tieneDescargo=$('tieneDescargo').checked;
 x.hallazgos=valoresFilas('hallazgos');
 x.anexos=valoresFilas('anexos').map(item=>item.descripcion);
 imageFields.forEach(([field])=>x[field]=imageData[field]);
 signatureFields.forEach(field=>x[field]=imageData[field]);
 return x;
}

async function cargar(x) {
 fieldIds.forEach(id=>{
  if($(id)) $(id).value=dateFields.has(id)?fechaParaControl(x[id]):x[id]||'';
 });
 $('tieneDescargo').checked=!!x.tieneDescargo;
 $('zonaDescargo').hidden=!x.tieneDescargo;
 $('hallazgos').replaceChildren();
 $('anexos').replaceChildren();
 (x.hallazgos||[]).forEach(item=>agregarFila('hallazgos',item));
 (x.anexos||[]).forEach(item=>agregarFila('anexos',{descripcion:item}));
 imageFields.forEach(([field])=>mostrarImagen(field,x[field]));
 await Promise.all(signatureFields.map(field=>dibujarFirma(field,x[field])));
 editId=x.id;
 vista(x);
}

function limpiarNumeracion(s) {
 return (s||'Sin información').split(/\r\n|\r|\n/).filter(x=>x.trim())
  .map(x=>x.replace(/^\s*\d+(?:\.\d+)*\.?\s*/,'')).join('\n')||'Sin información';
}

function seccion(t,s) {return `${t}\n${s||'Sin información'}\n\n`;}

function fechaLegible(value) {
 if(!value) return '';
 const match=/^(\d{4})-(\d{2})-(\d{2})$/.exec(value);
 if(!match) return value;
 return new Intl.DateTimeFormat('es-ES',{
  day:'numeric',month:'long',year:'numeric',timeZone:'UTC'
 }).format(new Date(Date.UTC(Number(match[1]),Number(match[2])-1,Number(match[3]))));
}

function vista(x) {
 const hallazgos=(x.hallazgos||[]).map((item,i)=>
  `${i+1}. ${item.incumplimiento||'-'}\n${item.descripcion||'-'}\nMedio probatorio: ${item.medioProbatorio||'-'}`
 ).join('\n\n')||'Sin hallazgos registrados';
 $('vista').textContent=
  `INFORME DE SUPERVISIÓN\n${x.ciudadInforme||''} ${fechaLegible(x.fechaInforme)}\nExpediente: ${x.expediente||''}\n`
  +`A: ${x.destinatario||''}\nDe: ${x.supervisor||''}\nAsunto: ${x.asunto||''}\nReferencia: ${x.referencia||''}\n\n`
  +seccion('OBJETIVO:',x.objetivo)
  +seccion('AGENTE FISCALIZADO Y UNIDAD:',
   `Empresa: ${x.empresa||''} | RUC: ${x.ruc||''}\nRegistro de Hidrocarburos: ${x.registroHidrocarburos||''}\n`
   +`Actividad: ${x.actividad||''} | Placa: ${x.placa||''} | Tracto: ${x.placaTracto||''}\n`
   +`Capacidad: ${x.capacidad||''} | Compartimientos: ${x.compartimientos||''}\n`
   +`Dirección legal: ${x.direccionLegal||''}, ${x.distrito||''}, ${x.provincia||''}, ${x.departamento||''}`)
  +seccion('ANTECEDENTES:',limpiarNumeracion(x.antecedentes))
  +seccion('MANDATO, COORDINACIÓN Y ASIGNACIÓN:',
   `Resolución ${x.numeroResolucion||''} (${x.fechaResolucion||''})\nComunicación de la empresa: ${x.fechaCartaEmpresa||''}\n`
   +`Lugar: ${x.lugarFiscalizacion||''}\nAsignación ${x.numeroAsignacion||''}, fecha ${x.fechaAsignacion||''}, `
   +`especialista: ${x.especialistaRegional||''}, carta línea: ${x.cartaLinea||''}`)
  +seccion('REGISTRO Y VISITA:',
   `Consulta del registro: ${x.fechaConsultaRegistro||''}\nVisita: ${x.fechaVisita||''} a las ${x.horaInicio||''}\n`
   +`Lugar: ${x.lugarFiscalizacion||''}`)
  +seccion('ANÁLISIS DE HECHOS:',x.analisis)
  +seccion('HALLAZGOS:',hallazgos)
  +(x.tieneDescargo?seccion('ANÁLISIS DE DESCARGOS:',x.descargos):'')
  +seccion('RECEPCIÓN DEL ACTA:',
   `${x.nombreRecibeActa||''} | DNI: ${x.dniRecibeActa||''} | ${x.cargoRecibeActa||''} | Fin: ${x.horaFin||''}`)
  +seccion('ANEXOS:',(x.anexos||[]).map((a,i)=>`ANEXO ${i+1}. ${a}`).join('\n')||'Sin anexos')
  +seccion('CONCLUSIONES:',x.conclusiones)
  +seccion('RECOMENDACIONES:',limpiarNumeracion(x.recomendaciones))
  +`FIRMAS\nRepresentante: ${x.representanteNombre||''} | DNI: ${x.representanteDni||''} | ${x.representanteCargo||''}\n`
  +`Fiscalizador: ${x.supervisor||''} | CIP: ${x.supervisorCip||''}`;
}

async function listar() {
 const response=await fetch('/api/informes');
 if(!response.ok) throw new Error('No se pudieron cargar los informes');
 const informes=await response.json();
 $('lista').replaceChildren();
 informes.forEach(x=>{
  const item=document.createElement('p');
  item.textContent=`${x.expediente} – ${x.empresa} `;
  const editar=document.createElement('button');
  editar.textContent='Editar';
  editar.onclick=async()=>{
   const detail=await fetch('/api/informes/'+encodeURIComponent(x.id));
   if(!detail.ok) {alert('No se pudo cargar el informe');return;}
   try {
    await cargar(await detail.json());
   } catch(error) {
    console.error(error);
    alert(error.message);
   }
  };
  const descargar=document.createElement('a');
  descargar.className='button-link';
  descargar.textContent='Descargar Word';
  descargar.href='/api/informes/'+encodeURIComponent(x.id)+'/docx';
  const eliminar=document.createElement('button');
  eliminar.textContent='Eliminar';
  eliminar.onclick=async()=>{
   if(confirm('¿Eliminar?')) {
    const response=await fetch('/api/informes/'+encodeURIComponent(x.id),{method:'DELETE'});
    if(!response.ok) {alert('No se pudo eliminar el informe');return;}
    await listar();
   }
  };
  item.append(editar,descargar,eliminar);
  $('lista').append(item);
 });
}

form.onsubmit=async event=>{
 event.preventDefault();
 const x=datos();
 const response=await fetch('/api/informes'+(editId?'/'+encodeURIComponent(editId):''),{
  method:editId?'PUT':'POST',
  headers:{'Content-Type':'application/json'},
  body:JSON.stringify(x)
 });
  if(!response.ok) {
   const contentType=response.headers.get('content-type')||'';
   const error=contentType.includes('application/json')
    ?(await response.json()).error
    :'No se pudo guardar el informe. Revise los campos obligatorios y el tamaño de las imágenes.';
   alert(error||'No se pudo guardar el informe.');
   return;
  }
 const guardado=await response.json();
 vista(guardado);
 editId=null;
 form.reset();
 $('zonaDescargo').hidden=true;
 $('hallazgos').replaceChildren();
 $('anexos').replaceChildren();
 limpiarImagenes();
 await listar();
};

listar().catch(error=>{console.error(error);alert(error.message);});
