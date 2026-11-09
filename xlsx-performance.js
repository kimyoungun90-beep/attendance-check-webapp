/* Native OOXML optimization. Preserve cell values, cached results and appearance. */
(() => {
'use strict';
const decode=x=>x.replace(/&quot;/g,'"').replace(/&apos;/g,"'").replace(/&lt;/g,'<').replace(/&gt;/g,'>').replace(/&amp;/g,'&');
const read=b=>fflate.strFromU8(b),write=s=>fflate.strToU8(s);
const wholeColumn=/(?:'([^']+)'|([A-Za-z_\uAC00-\uD7A3][A-Za-z0-9_\uAC00-\uD7A3]*))!\$([A-Z]+):\$([A-Z]+)|\$([A-Z]+):\$([A-Z]+)/g;
function normalizeWorksheet(source){let duplicateCells=0,restoredStyles=0;const cellPattern=/<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g;
 const xml=source.replace(/<row\b([^>]*?)(?:\/>|>([\s\S]*?)<\/row>)/g,(full,attrs,body='')=>{const cells=new Map();let duplicates=0;for(const m of body.matchAll(cellPattern)){const address=m[1].match(/\br="([A-Z]+[0-9]+)"/)?.[1];if(!address)continue;const old=cells.get(address);let value=m[0];if(old){duplicates++;duplicateCells++;const style=old.match(/\bs="([^"]+)"/)?.[1];if(style&&!/\bs="[^"]+"/.test(m[1])){value=value.replace('<c ',`<c s="${style}" `);restoredStyles++;}}cells.set(address,value);}if(!duplicates)return full;const column=a=>{let n=0;for(const x of a.match(/^[A-Z]+/)[0])n=n*26+x.charCodeAt(0)-64;return n;};return `<row${attrs}>${[...cells].sort((a,b)=>column(a[0])-column(b[0])).map(x=>x[1]).join('')}${body.replace(cellPattern,'')}</row>`;});return {xml,duplicateCells,restoredStyles};}
function optimize(input){
 const zip=fflate.unzipSync(new Uint8Array(input)),path='xl/workbook.xml';if(!zip[path])throw Error('지원하는 Excel 통합문서가 아닙니다.');
 let workbook=read(zip[path]);const originalWorkbook=workbook;
 const rels=read(zip['xl/_rels/workbook.xml.rels']),sheetPaths={},lastRows={};
 for(const m of workbook.matchAll(/<sheet\b([^>]+)\/?\s*>/g)){const name=decode(m[1].match(/\bname="([^"]+)"/)?.[1]||''),rid=m[1].match(/\br:id="([^"]+)"/)?.[1];const relation=[...rels.matchAll(/<Relationship\b([^>]+)\/?\s*>/g)].find(r=>r[1].match(/\bId="([^"]+)"/)?.[1]===rid);if(!relation)continue;const t=relation[1].match(/\bTarget="([^"]+)"/)?.[1];const p=t.startsWith('/')?t.slice(1):'xl/'+t;sheetPaths[name]=p;if(zip[p]){const xml=read(zip[p]);let max=1;for(const cell of xml.matchAll(/<c\b[^>]*\br="[A-Z]+([0-9]+)"[^>]*>([\s\S]*?)<\/c>/g))if(/<(?:v|is|f)\b/.test(cell[2]))max=Math.max(max,Number(cell[1]));lastRows[name]=max;}}
 // Inspect formulas/rules/validations and charts rather than treating cached link data as active use.
 const active=[];for(const [p,b]of Object.entries(zip)){if(!p.endsWith('.xml')||p.startsWith('xl/externalLinks/')||p===path||p.endsWith('sharedStrings.xml')||p.endsWith('calcChain.xml'))continue;const xml=read(b);for(const m of xml.matchAll(/<(?:f|formula|formula1|formula2|c:f)\b[^>]*>([\s\S]*?)<\/(?:f|formula|formula1|formula2|c:f)>/g))active.push(decode(m[1]));}
 const names=[...workbook.matchAll(/<definedName\b([^>]*)>([\s\S]*?)<\/definedName>/g)].map(m=>({full:m[0],name:decode(m[1].match(/\bname="([^"]+)"/)?.[1]||''),body:decode(m[2])}));
 const isBad=n=>/#REF!|\[[0-9]+\]/i.test(n.body);const candidates=names.filter(isBad),keep=new Set();const mentions=(formula,name)=>{const escaped=name.replace(/[.*+?^${}()|[\]\\]/g,'\\$&');return new RegExp('(^|[^\\w.\\uAC00-\\uD7A3])'+escaped+'($|[^\\w.\\uAC00-\\uD7A3])','i').test(formula);};
 const roots=[...active,...names.filter(n=>!isBad(n)).map(n=>n.body)];let more=true;while(more){more=false;for(const n of candidates){if(!keep.has(n.name)&&roots.some(f=>mentions(f,n.name))){keep.add(n.name);roots.push(n.body);more=true;}}}
 let removedNames=0;for(const n of candidates)if(!keep.has(n.name)){workbook=workbook.replace(n.full,'');removedNames++;}
 const retainedNames=names.filter(n=>!isBad(n)||keep.has(n.name));let removedLinks=0;
 if(![...active,...retainedNames.map(n=>n.body)].some(f=>/\[[0-9]+\]/.test(f))){workbook=workbook.replace(/<externalReferences\b[^>]*>[\s\S]*?<\/externalReferences>/,'');let r=rels.replace(/<Relationship\b[^>]*\bType="[^"]*\/externalLink"[^>]*\/>/g,()=>{removedLinks++;return '';});zip['xl/_rels/workbook.xml.rels']=write(r);for(const p of Object.keys(zip))if(p.startsWith('xl/externalLinks/'))delete zip[p];let types=read(zip['[Content_Types].xml']).replace(/<Override\b[^>]*\bPartName="\/xl\/externalLinks\/[^"]*"[^>]*\/>/g,'');zip['[Content_Types].xml']=write(types);}
 let boundedFormulas=0,duplicateCells=0,restoredStyles=0;
 function boundedFormula(formula,local){return formula.replace(wholeColumn,(full,quoted,plain,first,last,localFirst,localLast)=>{const sheet=quoted||plain||local;const end=Math.max(2000,Math.ceil(((lastRows[sheet]||0)+512)/1000)*1000);if(!lastRows[sheet])return full;return (quoted?`'${quoted}'!`:plain?plain+'!':'')+`$${first||localFirst}$1:$${last||localLast}$${end}`;});}
 for(const [sheet,p]of Object.entries(sheetPaths)){const normalized=normalizeWorksheet(read(zip[p]));duplicateCells+=normalized.duplicateCells;restoredStyles+=normalized.restoredStyles;let xml=normalized.xml;xml=xml.replace(/(<f\b[^>]*>)([\s\S]*?)(<\/f>)/g,(full,open,f,close)=>{const next=boundedFormula(f,sheet);if(next!==f)boundedFormulas++;return open+next+close;});zip[p]=write(xml);}
 // Old generated files requested a complete recalculation on every open. Keep automatic
 // calculation for edits without forcing an unnecessary full calculation at each launch.
 let resetForcedCalculation=false;
 workbook=workbook.replace(/<calcPr\b([^>]*?)\/>/,(full,attrs)=>{if(!/\b(?:fullCalcOnLoad|forceFullCalc)="1"|\bcalcId="0"/.test(attrs))return full;resetForcedCalculation=true;attrs=attrs.replace(/\s(?:calcMode|fullCalcOnLoad|forceFullCalc|calcId)="[^"]*"/g,'');return `<calcPr${attrs} calcId="191029" calcMode="auto" fullCalcOnLoad="0" forceFullCalc="0"/>`;});
 // Imported calculation chains can refer to formulas that the monthly raw-data rewrite
 // replaced with values. Drop only a chain containing such stale references.
 let staleCalculationEntries=0;
 if(zip['xl/calcChain.xml']){const formulasById={},sheets=[...workbook.matchAll(/<sheet\b([^>]+)\/?\s*>/g)];for(const m of sheets){const name=decode(m[1].match(/\bname="([^"]+)"/)?.[1]||''),id=m[1].match(/\bsheetId="([^"]+)"/)?.[1],p=sheetPaths[name];if(!p)continue;const set=new Set();for(const cell of read(zip[p]).matchAll(/<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g))if(/<f\b/.test(cell[2]||''))set.add(cell[1].match(/\br="([^"]+)"/)?.[1]);formulasById[id]=set;}let sheetId='1';for(const item of read(zip['xl/calcChain.xml']).matchAll(/<c\b([^>]*?)\/>/g)){sheetId=item[1].match(/\bi="([^"]+)"/)?.[1]||sheetId;const address=item[1].match(/\br="([^"]+)"/)?.[1];if(!formulasById[sheetId]?.has(address))staleCalculationEntries++;}
  if(staleCalculationEntries){delete zip['xl/calcChain.xml'];zip['xl/_rels/workbook.xml.rels']=write(read(zip['xl/_rels/workbook.xml.rels']).replace(/<Relationship\b[^>]*\bType="[^"]*\/calcChain"[^>]*\/>/g,''));zip['[Content_Types].xml']=write(read(zip['[Content_Types].xml']).replace(/<Override\b[^>]*\bPartName="\/xl\/calcChain.xml"[^>]*\/>/g,''));}}
 zip[path]=write(workbook);const output=fflate.zipSync(zip,{level:6});return {bytes:output,stats:{removedNames,removedLinks,boundedFormulas,duplicateCells,restoredStyles,resetForcedCalculation,staleCalculationEntries,beforeBytes:new Uint8Array(input).byteLength,afterBytes:output.byteLength}};
}
window.ZenielXlsxPerformance={optimize,normalizeWorksheet};
})();
