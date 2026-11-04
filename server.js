const express = require('express');

const fs = require('fs');

const path = require('path');

const JSZip = require('jszip');

const {DOMParser, XMLSerializer} = require('@xmldom/xmldom');

const app = express();

const file = path.join(__dirname, 'informes.json');

const template = path.join(__dirname, 'Informe_Supervision_Modelo_Identico_Anonimizado.docx');

app.use(express.json());

app.use(express.static(path.join(__dirname, 'public')));

const read = () => fs.existsSync(file) ? JSON.parse(fs.readFileSync(file,'utf8')) : [];

const save = data => fs.writeFileSync(file, JSON.stringify(data,null,2));

const valid = x => x && typeof x.expediente === 'string' && x.expediente.trim()

 && typeof x.empresa === 'string' && x.empresa.trim();

app.get('/api/informes', (req,res) => res.json(read()));

app.get('/api/informes/:id', (req,res) => {

 const x=read().find(a=>a.id===req.params.id);

 return x ? res.json(x) : res.status(404).json({error:'No encontrado'});

});

const wordNamespace = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const cleanText = value => String(value || '').replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '');
const withoutLeadingNumbers = value => cleanText(value).split(/\r\n|\r|\n/)
 .map(line=>line.replace(/^\s*\d+(?:\.\d+)*\.?\s*/, '')).join('\n');
const wordElements = (node, name) => Array.from(node.getElementsByTagNameNS(wordNamespace,name));
const paragraphText = paragraph => wordElements(paragraph,'t').map(node=>node.textContent).join('');

function replaceTextAcrossNodes(paragraph, search, replacement) {
 const nodes=wordElements(paragraph,'t');
 let text=nodes.map(node=>node.textContent).join('');
 let position=text.indexOf(search);
 while(position!==-1) {
  let offset=0;
  const first=nodes.findIndex(node=>{
   const next=offset+node.textContent.length;
   const found=position<next;
   if(!found) offset=next;
   return found;
  });
  let endOffset=0;
  const last=nodes.findIndex((node,index)=>{
   endOffset+=node.textContent.length;
   return index>=first && position+search.length<=endOffset;
  });
  const startOffset=nodes.slice(0,first).reduce((sum,node)=>sum+node.textContent.length,0);
  const prefix=nodes[first].textContent.slice(0,position-startOffset);
  const suffix=nodes[last].textContent.slice(position+search.length
   -nodes.slice(0,last).reduce((sum,node)=>sum+node.textContent.length,0));
  if(first===last) nodes[first].textContent=`${prefix}${replacement}${suffix}`;
  else {
   nodes[first].textContent=`${prefix}${replacement}`;
   for(const node of nodes.slice(first+1,last)) node.textContent='';
   nodes[last].textContent=suffix;
  }
  text=nodes.map(node=>node.textContent).join('');
  position=text.indexOf(search,position+replacement.length);
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
 if(paragraphs.length) setParagraphText(paragraphs[0],value);
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
 const value=(field)=>cleanText(informe[field] || 'Sin información');

 if(paragraphs[0]) replaceTextAcrossNodes(paragraphs[0],'«osinumero»',value('expediente'));

 const expediente=findParagraph(paragraphs,'Huánuco, 25 de octubre de 2026Expediente:');
 if(expediente) setParagraphText(expediente,`Expediente: ${value('expediente')}`);

 const supervisor=findParagraph(paragraphs,'De:Carlos Ejemplo Rivera');
 if(supervisor) setParagraphText(supervisor,`De: ${value('supervisor')}`);
 const oldSupervisorCompany=findParagraph(paragraphs,'Supervisor ES Supervisión Técnica');
 if(oldSupervisorCompany) setParagraphText(oldSupervisorCompany,'');

 const subject=findParagraph(paragraphs,'Asunto   :');
 if(subject) setParagraphText(subject,`Asunto: Fiscalización de ${value('empresa')} — placa ${value('placa')}`);
 const reference=findParagraph(paragraphs,'Referencia:   Registro de Hidrocarburos');
 if(reference) setParagraphText(reference,`Referencia: Expediente N.° ${value('expediente')}`);

 const objective=findParagraph(paragraphs,'El objetivo de la presente visita');
 if(objective) setParagraphText(objective,
  `El objetivo de la presente visita de fiscalización es verificar las condiciones de seguridad de la unidad vehicular de placa ${value('placa')}, operada por ${value('empresa')}.`);

 const antecedentsHeading=findParagraph(paragraphs,'ANTECEDENTES:');
 const legalHeading=findParagraph(paragraphs,'BASE LEGAL:');
 if(antecedentsHeading && legalHeading) {
  const start=paragraphs.indexOf(antecedentsHeading)+1;
  const end=paragraphs.indexOf(legalHeading);
  const entries=paragraphs.slice(start,end).filter(paragraph=>paragraphText(paragraph).trim());
  if(entries.length) {
   entries.forEach(removeParagraphNumbering);
   setParagraphText(entries[0],withoutLeadingNumbers(informe.antecedentes) || 'Sin información');
   clearParagraphs(entries.slice(1));
  }
 }

 const description=findParagraph(paragraphs,'El medio de transporte terrestre de combustibles líquidos');
 if(description) setParagraphText(description,
  `La unidad vehicular de placa ${value('placa')}, operada por ${value('empresa')}, corresponde al informe de fiscalización del expediente ${value('expediente')}.`);

 const address=findParagraph(paragraphs,'La fiscalización realizada al medio de transporte terrestre');
 if(address) setParagraphText(address,'');
 const visit=findParagraph(paragraphs,'Con fecha 24 de septiembre de 2026, siendo las 10:16 horas');
 if(visit) setParagraphText(visit,`ANÁLISIS DE HECHOS:\n${value('analisis')}`);

 const source=findParagraph(paragraphs,'Fuente: Acta de Fiscalización de Seguridad');
 if(source) setParagraphText(source,'Fuente: Información ingresada en el sistema.');
 const sampleVisit=findParagraph(paragraphs,'Culminada la fiscalización, se suscribió el ACTA');
 if(sampleVisit) setParagraphText(sampleVisit,'');
 const annexIntro=findParagraph(paragraphs,'De la visita realizada a la unidad vehicular se recabó');
 if(annexIntro) setParagraphText(annexIntro,'');
 const annexes=paragraphs.filter(paragraph=>/^ANEXO\s*\d*\.?/i.test(paragraphText(paragraph).trim())
  || paragraphText(paragraph).includes('VIDEO DE OPERATIVIDAD'));
 clearParagraphs(annexes);
 const conclusion=findParagraph(paragraphs,'TRANSPORTES MODELO DEMOSTRATIVO S.A.C., la Ficha');
 if(conclusion) setParagraphText(conclusion,value('conclusiones'));
 const recommendation=findParagraph(paragraphs,'Comunicar a TRANSPORTES MODELO DEMOSTRATIVO');
 if(recommendation) {
  removeParagraphNumbering(recommendation);
  setParagraphText(recommendation,withoutLeadingNumbers(informe.recomendaciones) || 'Sin información');
 }
 const extraRecommendation=findParagraph(paragraphs,'Asimismmo, se otorgue el plazo establecido');
 if(extraRecommendation) setParagraphText(extraRecommendation,'');

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

 if(tables.length!==3) throw new Error('La plantilla no contiene las tres tablas originales');
 setCellText(tables[0].getElementsByTagNameNS(wordNamespace,'tr').item(1)
  .getElementsByTagNameNS(wordNamespace,'tc').item(0),
  `Expediente: ${value('expediente')}\nEmpresa: ${value('empresa')}\nPlaca: ${value('placa')}`);
 const findingsRows=tables[1].getElementsByTagNameNS(wordNamespace,'tr');
 for(let row=1;row<findingsRows.length;row++) {
  const cells=findingsRows.item(row).getElementsByTagNameNS(wordNamespace,'tc');
  for(let col=1;col<cells.length;col++) setCellText(cells.item(col),'');
 }
 if(findingsRows.length>1) {
  const cells=findingsRows.item(1).getElementsByTagNameNS(wordNamespace,'tc');
  setCellText(cells.item(2),value('analisis'));
 }
 const signatures=tables[2].getElementsByTagNameNS(wordNamespace,'tr').item(1)
  .getElementsByTagNameNS(wordNamespace,'tc');
 setCellText(signatures.item(0),'');
 setCellText(signatures.item(1),`Supervisor:\n${value('supervisor')}`);

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

app.post('/api/informes', (req,res) => {

 if(!valid(req.body)) return res.status(400).json({error:'Expediente y empresa obligatorios'});

 const data=read(); const x={...req.body,id:String(Date.now())};

 data.push(x); save(data); res.status(201).json(x);

});

app.put('/api/informes/:id', (req,res) => {

 if(!valid(req.body)) return res.status(400).json({error:'Datos incompletos'});

 const data=read(); const i=data.findIndex(a=>a.id===req.params.id);

 if(i<0) return res.status(404).json({error:'No encontrado'});

 data[i]={...req.body,id:req.params.id};save(data);res.json(data[i]);

});

app.delete('/api/informes/:id', (req,res) => {

 const data=read(); const next=data.filter(a=>a.id!==req.params.id);

 if(next.length===data.length) return res.status(404).json({error:'No encontrado'});

 save(next);res.status(204).end();

});

const port = process.env.PORT || 3000;
app.listen(port,()=>console.log(`Servidor: http://localhost:${port}`));