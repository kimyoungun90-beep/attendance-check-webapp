/* Calendar-aware week boundaries, Monday through Sunday. */
(() => {
'use strict';
function applyWeeks(input,month){
 if(!/^\d{4}-(0[1-9]|1[0-2])$/.test(month))throw Error('주차 서식의 대상월을 확인해주세요.');
 const zip=fflate.unzipSync(new Uint8Array(input)),read=b=>fflate.strFromU8(b),write=s=>fflate.strToU8(s),workbook=read(zip['xl/workbook.xml']),rels=read(zip['xl/_rels/workbook.xml.rels']);
 const sh=[...workbook.matchAll(/<sheet\b([^>]*?)\/?\s*>/g)].find(m=>/\bname="MX"/.test(m[1]));if(!sh)return new Uint8Array(input);
 const rid=sh[1].match(/\br:id="([^"]+)"/)?.[1],rel=[...rels.matchAll(/<Relationship\b([^>]*?)\/>/g)].find(m=>m[1].match(/\bId="([^"]+)"/)?.[1]===rid),target=rel?.[1].match(/\bTarget="([^"]+)"/)?.[1];if(!target)throw Error('마감 시트를 찾을 수 없습니다.');const path=target.startsWith('/')?target.slice(1):'xl/'+target;
 let styles=read(zip['xl/styles.xml']),sheet=read(zip[path]);
 const section=name=>styles.match(new RegExp(`<${name}\\b[^>]*>([\\s\\S]*?)<\\/${name}>`))?.[1]||'';
 const borders=[...section('borders').matchAll(/<border\b[^>]*?(?:\/>|>[\s\S]*?<\/border>)/g)].map(m=>m[0]),fills=[...section('fills').matchAll(/<fill\b[^>]*?(?:\/>|>[\s\S]*?<\/fill>)/g)].map(m=>m[0]),xfs=[...section('cellXfs').matchAll(/<xf\b[^>]*?(?:\/>|>[\s\S]*?<\/xf>)/g)].map(m=>m[0]);
 const first=new Date(+month.slice(0,4),+month.slice(5)-1,1).getDay(),days=new Date(+month.slice(0,4),+month.slice(5),0).getDate(),offset=(first+6)%7;
 const fillIDs=['DFE8F3','F1F5FA'].map(color=>{const value=`<fill><patternFill patternType="solid"><fgColor rgb="FF${color}"/><bgColor indexed="64"/></patternFill></fill>`;let i=fills.indexOf(value);if(i<0){i=fills.length;fills.push(value);}return i;});
 const cache=new Map();function styleFor(base,day,header){const weekday=(first+day-1)%7,inside=day<=days,left=inside&&(day===1||weekday===1),right=inside&&(day===days||weekday===0),group=Math.floor((offset+day-1)/7)%2,key=[base,inside,left,right,header,group].join('|');if(cache.has(key))return cache.get(key);
  let xf=xfs[base]||xfs[0],bid=Number(xf.match(/\bborderId="([0-9]+)"/)?.[1]||0),border=borders[bid]||'<border/>';
  if(/\/>$/.test(border))border=border.replace(/\/>$/,'></border>');border=border.replace(/<(?:left|right)\b[^>]*?(?:\/>|>[\s\S]*?<\/(?:left|right)>)/g,'');const side=(name,strong)=>!inside?`<${name}/>`:`<${name} style="${strong?'medium':'thin'}"><color rgb="FF${strong?'203960':'B5C0CD'}"/></${name}>`;border=border.replace(/(<border\b[^>]*>)/,'$1'+side('left',left)+side('right',right));let newBorder=borders.indexOf(border);if(newBorder<0){newBorder=borders.length;borders.push(border);}xf=xf.replace(/\bborderId="[0-9]+"/,`borderId="${newBorder}"`);if(!/\bapplyBorder=/.test(xf))xf=xf.replace('<xf ','<xf applyBorder="1" ');else xf=xf.replace(/\bapplyBorder="[^"]*"/,'applyBorder="1"');
  if(header&&inside){xf=xf.replace(/\bfillId="[0-9]+"/,`fillId="${fillIDs[group]}"`);if(!/\bapplyFill=/.test(xf))xf=xf.replace('<xf ','<xf applyFill="1" ');else xf=xf.replace(/\bapplyFill="[^"]*"/,'applyFill="1"');}
  let id=xfs.indexOf(xf);if(id<0){id=xfs.length;xfs.push(xf);}cache.set(key,id);return id;
 }
 sheet=sheet.replace(/<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g,(full,attrs)=>{const address=attrs.match(/\br="([A-Z]+)([0-9]+)"/);if(!address)return full;const row=+address[2],column=XLSX.utils.decode_col(address[1]),day=column-11;if(row<7||row>46||day<1||day>31)return full;const base=Number(attrs.match(/\bs="([0-9]+)"/)?.[1]||0),id=styleFor(base,day,row===7);return /\bs="[0-9]+"/.test(attrs)?full.replace(/\bs="[0-9]+"/,`s="${id}"`):full.replace('<c ',`<c s="${id}" `);});
 for(const [name,values]of [['borders',borders],['fills',fills],['cellXfs',xfs]])styles=styles.replace(new RegExp(`<${name}\\b([^>]*)>[\\s\\S]*?<\\/${name}>`),(full,attrs)=>`<${name}${attrs.replace(/\bcount="[0-9]+"/,`count="${values.length}"`)}>${values.join('')}</${name}>`);
 zip[path]=write(sheet);zip['xl/styles.xml']=write(styles);return fflate.zipSync(zip,{level:6});
}
window.ZenielWeeklyLayout={applyWeeks};
})();
