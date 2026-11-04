const express = require('express');

const fs = require('fs');

const path = require('path');

const JSZip = require('jszip');

const sharp = require('sharp');

const {DOMParser, XMLSerializer} = require('@xmldom/xmldom');

const app = express();

const file = path.join(__dirname, 'informes.json');

const template = path.join(__dirname, 'Informe_Supervision_Modelo_Identico_Anonimizado.docx');

app.use(express.json({limit: '24mb'}));

app.use(express.static(path.join(__dirname, 'public')));

const read = () => fs.existsSync(file) ? JSON.parse(fs.readFileSync(file,'utf8')) : [];

const save = data => fs.writeFileSync(file, JSON.stringify(data,null,2));

const valid = x => x && typeof x.expediente === 'string' && x.expediente.trim()

 && typeof x.empresa === 'string' && x.empresa.trim();

const imageFields=['imagenRegistro','firmaRepresentante','firmaFiscalizador'];
const maxImageBytes=5*1024*1024;
const allowedImageMimeTypes=new Set(['image/jpeg','image/png','image/webp']);
const summarizeInforme=informe=>{
 const {imagenRegistro,firmaRepresentante,firmaFiscalizador,...summary}=informe;
 return summary;
};

function decodeImageData(data) {
 if(data===undefined||data==='') return null;
 if(typeof data!=='string') throw new Error('El archivo de imagen no es válido');
 const match=/^data:(image\/(?:jpeg|png|webp));base64,([A-Za-z0-9+/]+={0,2})$/.exec(data);
 if(!match||!allowedImageMimeTypes.has(match[1])) {
  throw new Error('Use una imagen JPG, PNG o WebP');
 }
 const buffer=Buffer.from(match[2],'base64');
 if(buffer.length>maxImageBytes) throw new Error('Cada imagen debe pesar como máximo 5 MB');
 if(!buffer.length||buffer.toString('base64')!==match[2]) throw new Error('El archivo de imagen está dañado');
 return buffer;
}

async function validateReportImages(informe) {
 for(const field of imageFields) {
  const buffer=decodeImageData(informe[field]);
  if(!buffer) continue;
  const metadata=await sharp(buffer,{limitInputPixels:25_000_000}).metadata();
  if(!['jpeg','png','webp'].includes(metadata.format)||!metadata.width||!metadata.height) {
   throw new Error('Use una imagen JPG, PNG o WebP válida');
  }
 }
}

async function prepareImage(data, format, width, height) {
 const buffer=decodeImageData(data);
 const background=format==='jpeg'
  ?{r:255,g:255,b:255,alpha:1}:{r:255,g:255,b:255,alpha:0};
 const source=buffer||await sharp({
  create:{width,height,channels:format==='jpeg'?3:4,background}
 }).png().toBuffer();
 const image=sharp(source,{limitInputPixels:25_000_000}).rotate()
  .resize(width,height,{fit:'contain',background});
 return format==='jpeg'
  ?image.flatten({background:'#ffffff'}).jpeg({quality:88}).toBuffer()
  :image.png().toBuffer();
}

app.get('/api/informes', (req,res) => res.json(read().map(summarizeInforme)));

app.get('/api/informes/:id', (req,res) => {

 const x=read().find(a=>a.id===req.params.id);

 return x ? res.json(x) : res.status(404).json({error:'No encontrado'});

});

const wordNamespace = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const cleanText = value => String(value || '').replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '');
const dateFields = new Set([
 'fechaInforme','fechaRegistro','fechaResolucion','fechaCartaEmpresa','fechaAsignacion',
 'fechaConsultaRegistro','fechaVisita'
]);
function formatDate(value) {
 const match=/^(\d{4})-(\d{2})-(\d{2})$/.exec(value||'');
 if(!match) return cleanText(value||'Sin información');
 const [,year,month,day]=match;
 const date=new Date(Date.UTC(Number(year),Number(month)-1,Number(day)));
 if(date.getUTCFullYear()!==Number(year)||date.getUTCMonth()!==Number(month)-1
  ||date.getUTCDate()!==Number(day)) return cleanText(value);
 return new Intl.DateTimeFormat('es-ES',{
  day:'numeric',month:'long',year:'numeric',timeZone:'UTC'
 }).format(date);
}
const withoutLeadingNumbers = value => cleanText(value).split(/\r\n|\r|\n/)
 .map(line=>line.replace(/^\s*\d+(?:\.\d+)*\.?\s*/, '')).join('\n');
const wordElements = (node, name) => Array.from(node.getElementsByTagNameNS(wordNamespace,name));
const paragraphText = paragraph => wordElements(paragraph,'t').map(node=>node.textContent).join('');

function replaceTextAcrossNodes(paragraph, search, replacement) {
 const nodes=wordElements(paragraph,'t');
 const original=nodes.map(node=>node.textContent);
 const text=original.join('');
 const offsets=[];
 let length=0;
 for(const value of original) {
  offsets.push(length);
  length+=value.length;
 }
 const matches=[];
 for(let position=text.indexOf(search);position!==-1;
  position=text.indexOf(search,position+search.length)) matches.push(position);
 for(const position of matches.reverse()) {
  const end=position+search.length;
  const first=offsets.findIndex((offset,index)=>
   offset+original[index].length>position);
  const last=offsets.findIndex((offset,index)=>offset+original[index].length>=end);
  if(first===-1||last===-1) continue;
  const prefix=nodes[first].textContent.slice(0,position-offsets[first]);
  const suffix=nodes[last].textContent.slice(end-offsets[last]);
  if(first===last) nodes[first].textContent=`${prefix}${replacement}${suffix}`;
  else {
   nodes[first].textContent=prefix+replacement;
   for(const node of nodes.slice(first+1,last)) node.textContent='';
   nodes[last].textContent=suffix;
  }
 }
}

function setParagraphText(paragraph, value) {
 let nodes=wordElements(paragraph,'t');
 if(!nodes.length) {
  const document=paragraph.ownerDocument;
  const run=document.createElementNS(wordNamespace,'w:r');
  const text=document.createElementNS(wordNamespace,'w:t');
  run.appendChild(text);
  paragraph.appendChild(run);
  nodes=[text];
 }
 const lines=cleanText(value).split(/\r\n|\r|\n/);
 if(lines.length>1) setParagraphAlignment(paragraph,'left');
 nodes[0].textContent=lines.shift();
 for(const node of nodes.slice(1)) node.textContent='';
 let cursor=nodes[0];
 for(const line of lines) {
  const run=cursor.parentNode;
  const lineBreak=run.ownerDocument.createElementNS(wordNamespace,'w:br');
  const text=run.ownerDocument.createElementNS(wordNamespace,'w:t');
  text.setAttributeNS('http://www.w3.org/XML/1998/namespace','xml:space','preserve');
  text.appendChild(run.ownerDocument.createTextNode(line));
  run.insertBefore(lineBreak,cursor.nextSibling);
  run.insertBefore(text,lineBreak.nextSibling);
  cursor=text;
 }
}

function setParagraphAlignment(paragraph, alignment) {
 const document=paragraph.ownerDocument;
 let properties=Array.from(paragraph.childNodes).find(node=>node.nodeType===1
  && node.namespaceURI===wordNamespace && node.localName==='pPr');
 if(!properties) {
  properties=document.createElementNS(wordNamespace,'w:pPr');
  paragraph.insertBefore(properties,paragraph.firstChild);
 }
 let justification=Array.from(properties.childNodes).find(node=>node.nodeType===1
  && node.namespaceURI===wordNamespace && node.localName==='jc');
 if(!justification) {
  justification=document.createElementNS(wordNamespace,'w:jc');
  properties.appendChild(justification);
 }
 justification.setAttributeNS(wordNamespace,'w:val',alignment);
}

function setParagraphStyle(paragraph, style) {
 const document=paragraph.ownerDocument;
 let properties=Array.from(paragraph.childNodes).find(node=>node.nodeType===1
  && node.namespaceURI===wordNamespace && node.localName==='pPr');
 if(!properties) {
  properties=document.createElementNS(wordNamespace,'w:pPr');
  paragraph.insertBefore(properties,paragraph.firstChild);
 }
 let paragraphStyle=Array.from(properties.childNodes).find(node=>node.nodeType===1
  && node.namespaceURI===wordNamespace && node.localName==='pStyle');
 if(!paragraphStyle) {
  paragraphStyle=document.createElementNS(wordNamespace,'w:pStyle');
  properties.insertBefore(paragraphStyle,properties.firstChild);
 }
 paragraphStyle.setAttributeNS(wordNamespace,'w:val',style);
}

function removeParagraphNumbering(paragraph) {
 const properties=Array.from(paragraph.childNodes).find(node=>node.nodeType===1
  && node.namespaceURI===wordNamespace && node.localName==='pPr');
 if(!properties) return;
 const numbering=Array.from(properties.childNodes).find(node=>node.nodeType===1
  && node.namespaceURI===wordNamespace && node.localName==='numPr');
 if(numbering) properties.removeChild(numbering);
}

function setCellText(cell, value) {
 const paragraphs=wordElements(cell,'p').filter(p=>p.parentNode===cell);
 if(!paragraphs.length) return;
 setParagraphText(paragraphs[0],value);
 clearParagraphs(paragraphs.slice(1));
}

function directChildren(parent, name) {
 return Array.from(parent.childNodes).filter(node=>node.nodeType===1
  && node.namespaceURI===wordNamespace && node.localName===name);
}

function findParagraph(paragraphs, text) {
 return paragraphs.find(paragraph=>paragraphText(paragraph).includes(text));
}

function clearParagraphs(paragraphs) {
 for(const paragraph of paragraphs) setParagraphText(paragraph,'');
}

function parseXml(xml) {
 const errors=[];
 const document=new DOMParser({onError:(level,message)=>{
  if(level!=='warning') errors.push(message);
 }}).parseFromString(xml,'application/xml');
 if(errors.length) throw new Error(`XML inválido: ${errors[0]}`);
 return document;
}

async function generarDocumento(informe) {
 const archive = await JSZip.loadAsync(fs.readFileSync(template));
 const documentFile = archive.file('word/document.xml');
 if (!documentFile) throw new Error('La plantilla no contiene word/document.xml');

 const document = parseXml(await documentFile.async('string'));
 const body = document.getElementsByTagNameNS(wordNamespace,'body').item(0);
 if(!body) throw new Error('No se pudo leer el cuerpo de la plantilla Word');
 const paragraphs=Array.from(body.childNodes).filter(node=>node.nodeType===1
  && node.namespaceURI===wordNamespace && node.localName==='p');
 const tables=Array.from(body.childNodes).filter(node=>node.nodeType===1
  && node.namespaceURI===wordNamespace && node.localName==='tbl');
 const value=field=>dateFields.has(field)
  ?formatDate(informe[field]):cleanText(informe[field]||'Sin información');
 if(paragraphs[0]) replaceTextAcrossNodes(paragraphs[0],'«osinumero»',value('expediente'));
 const reportDate=findParagraph(paragraphs,'Huánuco, 25 de octubre de 2026Expediente:');
 if(reportDate) setParagraphText(reportDate,
  `${value('ciudadInforme')}, ${value('fechaInforme')}\nExpediente: ${value('expediente')}`);

 const recipient=findParagraph(paragraphs,'Oficina Regional Huánuco - OSINERGMIN');
 if(recipient) setParagraphText(recipient,`A: ${value('destinatario')}`);
 const supervisor=findParagraph(paragraphs,'De:Carlos Ejemplo Rivera');
 if(supervisor) setParagraphText(supervisor,`De: ${value('supervisor')}`);
 const supervisorCompany=findParagraph(paragraphs,'Supervisor ES Supervisión Técnica');
 if(supervisorCompany) setParagraphText(supervisorCompany,
  [value('cargoSupervisor'),value('empresaSupervisora')].filter(Boolean).join(' — '));
 const subject=findParagraph(paragraphs,'Asunto   :');
 if(subject) setParagraphText(subject,`Asunto: ${value('asunto')}`);
 const reference=findParagraph(paragraphs,'Referencia:   Registro de Hidrocarburos');
 if(reference) setParagraphText(reference,`Referencia: ${value('referencia')}`);

 const objective=findParagraph(paragraphs,'El objetivo de la presente visita');
 if(objective) setParagraphText(objective,value('objetivo'));

 const antecedentsHeading=findParagraph(paragraphs,'ANTECEDENTES:');
 const legalHeading=findParagraph(paragraphs,'BASE LEGAL:');
 if(antecedentsHeading && legalHeading) {
  const start=paragraphs.indexOf(antecedentsHeading)+1;
  const end=paragraphs.indexOf(legalHeading);
  const entries=paragraphs.slice(start,end).filter(paragraph=>paragraphText(paragraph).trim());
  entries.forEach(removeParagraphNumbering);
 }

 const description=findParagraph(paragraphs,'El medio de transporte terrestre de combustibles líquidos');
 if(description) setParagraphText(description,
  `El medio de transporte terrestre de ${value('actividad')}, operado por ${value('empresa')}, cuenta con la unidad de placa principal ${value('placa')}, tracto ${value('placaTracto')}, Registro de Hidrocarburos N.° ${value('registroHidrocarburos')}, RUC N.° ${value('ruc')}, capacidad autorizada de ${value('capacidad')} y ${value('compartimientos')} compartimientos.`);

 const address=findParagraph(paragraphs,'La fiscalización realizada al medio de transporte terrestre');
 if(address) setParagraphText(address,
  `La fiscalización del medio de transporte de placa principal ${value('placa')} se llevó a cabo en ${value('lugarFiscalizacion')}.`);
 const visit=findParagraph(paragraphs,'Con fecha 24 de septiembre de 2026, siendo las 10:16 horas');
 if(visit) setParagraphText(visit,
  `Con fecha ${value('fechaVisita')}, siendo las ${value('horaInicio')} horas, se inició la fiscalización de la unidad de placa principal ${value('placa')}, operada por ${value('empresa')}, con Registro de Hidrocarburos N.° ${value('registroHidrocarburos')}.\nANÁLISIS DE HECHOS:\n${value('analisis')}\nSe constataron los hallazgos descritos en el siguiente cuadro:`);

 const registrationTable=tables[0];
 const findingsTable=tables[1];
 const signaturesTable=tables[2];
 if(tables.length!==3) throw new Error('La plantilla no contiene las tres tablas originales');
 const registrationCells=directChildren(
  directChildren(registrationTable,'tr')[1],'tc');
 setCellText(registrationCells[0],
  `Consulta del Registro de Hidrocarburos de fecha ${value('fechaConsultaRegistro')}: Registro N.° ${value('registroHidrocarburos')}, titular ${value('empresa')}, RUC N.° ${value('ruc')}, placa principal ${value('placa')}, tracto ${value('placaTracto')}, capacidad de ${value('capacidad')} y ${value('compartimientos')} compartimientos.`);

 const source=findParagraph(paragraphs,'Fuente: Acta de Fiscalización de Seguridad');
 if(source) setParagraphText(source,`Fuente: ${value('fuenteActa')}`);
 const sampleVisit=findParagraph(paragraphs,'Culminada la fiscalización, se suscribió el ACTA');
 if(sampleVisit) setParagraphText(sampleVisit,
  `Culminada la fiscalización, se suscribió el Acta correspondiente al Expediente N.° ${value('expediente')}, recibida por ${value('nombreRecibeActa')}, identificado(a) con DNI N.° ${value('dniRecibeActa')}, en calidad de ${value('cargoRecibeActa')}. La diligencia concluyó a las ${value('horaFin')} horas.`);
 const annexIntro=findParagraph(paragraphs,'De la visita realizada a la unidad vehicular se recabó');
 if(annexIntro) setParagraphText(annexIntro,'De la visita de fiscalización se recabó la siguiente documentación:');
 const annexParagraphs=paragraphs.filter(paragraph=>/^ANEXO\s*\d*\.?/i.test(paragraphText(paragraph).trim())
  || paragraphText(paragraph).includes('VIDEO DE OPERATIVIDAD'));
 const annexes=Array.isArray(informe.anexos)?informe.anexos:[];
 if(!annexParagraphs.length) throw new Error('La plantilla no contiene espacios para los anexos');
 const annexInsertionPoint=annexParagraphs[annexParagraphs.length-1].nextSibling;
 while(annexParagraphs.length<annexes.length) {
  const copy=annexParagraphs[0].cloneNode(true);
  body.insertBefore(copy,annexInsertionPoint);
  annexParagraphs.push(copy);
 }
 annexParagraphs.forEach((paragraph,index)=>{
  removeParagraphNumbering(paragraph);
  const annex=annexes[index];
  setParagraphText(paragraph,annex?`ANEXO ${index+1}. ${cleanText(annex)}`:'');
 });

 const conclusion=findParagraph(paragraphs,'TRANSPORTES MODELO DEMOSTRATIVO S.A.C., la Ficha');
 if(conclusion) setParagraphText(conclusion,value('conclusiones'));
 const recommendation=findParagraph(paragraphs,'Comunicar a TRANSPORTES MODELO DEMOSTRATIVO');
 if(recommendation) {
  removeParagraphNumbering(recommendation);
  setParagraphText(recommendation,withoutLeadingNumbers(informe.recomendaciones) || 'Sin información');
 }
 const extraRecommendation=findParagraph(paragraphs,'Asimismmo, se otorgue el plazo establecido');
 if(extraRecommendation) setParagraphText(extraRecommendation,'');

 const reportDetail=findParagraph(paragraphs,'Con fecha 25 de octubre de 2026, se elabora el presente informe');
 if(reportDetail) setParagraphText(reportDetail,
  `Antecedentes adicionales:\n${withoutLeadingNumbers(informe.antecedentes)||'Sin información'}\n\nCon fecha ${value('fechaInforme')}, se elabora el presente informe de fiscalización, correspondiente a la Carta Línea N.° ${value('cartaLinea')} y al Expediente N.° ${value('expediente')}, para su evaluación.`);
 const chronology=[
  ['Con fecha 29 de agosto de 2024,',`Con fecha ${value('fechaRegistro')}, ${value('empresa')} obtuvo el Registro de Hidrocarburos N.° ${value('registroHidrocarburos')} para operar la unidad vehicular de placa principal N.° ${value('placa')} y tracto ${value('placaTracto')}, con capacidad autorizada de ${value('capacidad')} y ${value('compartimientos')} compartimientos, con dirección legal en ${value('direccionLegal')}, distrito de ${value('distrito')}, provincia de ${value('provincia')}, departamento de ${value('departamento')}.`],
  ['Mediante la Resolución de Autoridad de Fiscalización',`Mediante la Resolución de Autoridad de Fiscalización N.° ${value('numeroResolucion')}, de fecha ${value('fechaResolucion')}, se dispuso que ${value('empresa')} ponga a disposición la unidad de placa principal N.° ${value('placa')}, con Registro de Hidrocarburos N.° ${value('registroHidrocarburos')}.`],
  ['Con fecha 11 de septiembre de 2026,',`Con fecha ${value('fechaCartaEmpresa')}, ${value('empresa')} comunicó que la fiscalización de la unidad ${value('placa')} se realizaría en ${value('lugarFiscalizacion')}.`],
  ['Con fecha 18 de septiembre de 2026,',`Con fecha ${value('fechaAsignacion')}, ${value('especialistaRegional')} mediante Asignación N.° ${value('numeroAsignacion')} asignó la Carta Línea N.° ${value('cartaLinea')} a ${value('empresaSupervisora')}, para efectuar la fiscalización a la unidad de placa ${value('placa')}.`],
  ['Con fecha 24 de septiembre de 2026, el Ing.',`Con fecha ${value('fechaVisita')}, ${value('supervisor')} realizó la visita de fiscalización al medio de transporte de placa principal N.° ${value('placa')}, operado por ${value('empresa')}, correspondiente al Expediente N.° ${value('expediente')}.`]
 ];
 chronology.forEach(([marker,text])=>{
  const paragraph=findParagraph(paragraphs,marker);
  if(paragraph) setParagraphText(paragraph,text);
 });

 const conclusionHeading=findParagraph(paragraphs,'CONCLUSIONES:');
 if(informe.tieneDescargo && conclusionHeading) {
  const position=paragraphs.indexOf(conclusionHeading);
  const blanks=paragraphs.slice(0,position).filter(paragraph=>!paragraphText(paragraph).trim());
  const heading=blanks[blanks.length-2];
  const detail=blanks[blanks.length-1];
  if(heading && detail) {
   setParagraphStyle(heading,'Ttulo2');
   setParagraphText(heading,'ANÁLISIS DE DESCARGOS:');
   setParagraphStyle(detail,'Sangradetextonormal');
   setParagraphText(detail,value('descargos'));
  }
 }

 const findingRows=directChildren(findingsTable,'tr');
 const findingTemplate=findingRows[1];
 if(!findingTemplate) throw new Error('La plantilla no contiene filas para los hallazgos');
 findingRows.slice(1).forEach(row=>findingsTable.removeChild(row));
 const findings=Array.isArray(informe.hallazgos)?informe.hallazgos:[];
 (findings.length?findings:[{}]).forEach((finding,index)=>{
  const row=findingTemplate.cloneNode(true);
  const cells=directChildren(row,'tc');
  setCellText(cells[0],findings.length?String(index+1):'');
  setCellText(cells[1],finding.incumplimiento||'');
  setCellText(cells[2],finding.descripcion||'');
  setCellText(cells[3],finding.medioProbatorio||'');
  findingsTable.appendChild(row);
 });

 const signatureRow=directChildren(signaturesTable,'tr')[1];
 const signatures=directChildren(signatureRow,'tc');
 setCellText(signatures[0],
  `${value('representanteNombre')}\nDNI N.° ${value('representanteDni')}\n${value('representanteCargo')}`);
 setCellText(signatures[1],
  `${value('supervisor')}\nCIP N.° ${value('supervisorCip')}\n${value('cargoSupervisor')}\n${value('empresaSupervisora')}`);

 archive.file('word/media/image1.jpg',await prepareImage(informe.imagenRegistro,'jpeg',808,344));
 archive.file('word/media/image2.png',await prepareImage(informe.firmaRepresentante,'png',172,128));
 archive.file('word/media/image3.png',await prepareImage(informe.firmaFiscalizador,'png',249,119));
 archive.file('word/document.xml',new XMLSerializer().serializeToString(document));
 const headerFile=archive.file('word/header1.xml');
 if(headerFile) {
  const header=parseXml(await headerFile.async('string'));
  const sampleBrand=wordElements(header,'p')
   .find(paragraph=>paragraphText(paragraph).includes('SUPERVISIÓN TÉCNICA DEMOSTRATIVA'));
  if(sampleBrand) setParagraphText(sampleBrand,'');
  archive.file('word/header1.xml',new XMLSerializer().serializeToString(header));
 }
 return archive.generateAsync({type: 'nodebuffer', compression: 'DEFLATE'});
}

app.get('/api/informes/:id/docx', async (req,res) => {
 const informe=read().find(a=>a.id===req.params.id);
 if(!informe) return res.status(404).json({error:'No encontrado'});
 try {
  const document=await generarDocumento(informe);
  const expediente=String(informe.expediente).replace(/[^a-zA-Z0-9_-]/g, '_');
  res.type('application/vnd.openxmlformats-officedocument.wordprocessingml.document');
  res.attachment(`Informe-${expediente}.docx`);
  res.send(document);
 } catch(error) {
  console.error('No se pudo generar el documento Word:', error);
  res.status(500).json({error:'No se pudo generar el documento Word'});
 }
});

app.post('/api/informes', async (req,res) => {
 if(!valid(req.body)) return res.status(400).json({error:'Expediente y empresa obligatorios'});
 try {
  await validateReportImages(req.body);
 } catch(error) {
  return res.status(400).json({error:error.message});
 }
 const data=read(); const x={...req.body,id:String(Date.now())};
 data.push(x); save(data); res.status(201).json(summarizeInforme(x));
});

app.put('/api/informes/:id', async (req,res) => {
 if(!valid(req.body)) return res.status(400).json({error:'Datos incompletos'});
 const data=read(); const i=data.findIndex(a=>a.id===req.params.id);
 if(i<0) return res.status(404).json({error:'No encontrado'});
 try {
  await validateReportImages(req.body);
 } catch(error) {
  return res.status(400).json({error:error.message});
 }
 data[i]={...req.body,id:req.params.id};save(data);res.json(summarizeInforme(data[i]));
});

app.delete('/api/informes/:id', (req,res) => {

 const data=read(); const next=data.filter(a=>a.id!==req.params.id);

 if(next.length===data.length) return res.status(404).json({error:'No encontrado'});

 save(next);res.status(204).end();

});

const port = process.env.PORT || 3000;
app.listen(port,()=>console.log(`Servidor: http://localhost:${port}`));