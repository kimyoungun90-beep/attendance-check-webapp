/* ZENIEL COSTCO Attendance - new implementation 1.0.0
   All analysis runs in the browser. Import the previous month's exported XLSX to continue history.
*/
(() => {
'use strict';
const VERSION='1.2.1';
const STATE_SHEET='_STATE';
const FIELDS=['filePlan','fileAttendance','fileLeave','fileHours'];
const $=id=>document.getElementById(id);
const emptyState=()=>({version:1,months:{},holidays:[],hires:{},overrides:{}});
let history=emptyState(), holidays=[], hires={}, result=null, lastReport=null, activeView='late';
const uid=(n,d,t)=>[String(n||'').trim(),d||'',t||''].join('|');
const clone=o=>JSON.parse(JSON.stringify(o));
const iso=(y,m,d)=>`${String(y).padStart(4,'0')}-${String(m).padStart(2,'0')}-${String(d).padStart(2,'0')}`;
const validISO=x=>/^\d{4}-\d{2}-\d{2}$/.test(x)&&new Date(x+'T12:00:00').toISOString().startsWith(x);
const isoDate=x=>{if(x==null||x==='')return '';if(x instanceof Date)return iso(x.getFullYear(),x.getMonth()+1,x.getDate());let s=String(x).trim();if(/^\d{8}$/.test(s))return `${s.slice(0,4)}-${s.slice(4,6)}-${s.slice(6)}`;let m=s.match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})/);if(m)return iso(+m[1],+m[2],+m[3]);if(typeof x==='number'&&x>10000000)return isoDate(String(x));if(typeof x==='number'){let d=new Date(Date.UTC(1899,11,30)+Math.round(x)*86400000);return iso(d.getUTCFullYear(),d.getUTCMonth()+1,d.getUTCDate());}return ''};
const dayStep=(d,n)=>{let x=new Date(d+'T12:00:00');x.setDate(x.getDate()+n);return iso(x.getFullYear(),x.getMonth()+1,x.getDate())};
const daysBetween=(a,b)=>Math.round((new Date(b+'T12:00:00')-new Date(a+'T12:00:00'))/86400000);
const monday=d=>{let x=new Date(d+'T12:00:00');return dayStep(d,-((x.getDay()+6)%7))};
const nice=d=>d?`${Number(d.slice(5,7))}월 ${Number(d.slice(8))}일`:'';
const rangeDates=(a,b,max=366)=>{if(!validISO(a)||!validISO(b)||a>b)return[];let n=Math.min(daysBetween(a,b)+1,max);return Array.from({length:n},(_,i)=>dayStep(a,i));};
const trim=x=>String(x??'').trim();
const canonStore=x=>trim(x).replace(/코스트코/g,'').replace(/[\s()]/g,'').replace(/점$/,'').replace(/매장$/,'').replace(/^대구혁신$/,'혁신').replace(/^부산$/,'부산');
const shiftOf=x=>{let m=trim(x).match(/근무\s*([ABCＡＢＣ])/i);return m?({'Ａ':'A','Ｂ':'B','Ｃ':'C'}[m[1].toUpperCase()]||m[1].toUpperCase()):''};
const vacationType=x=>{let s=trim(x).replace(/\s/g,'');if(!s)return {kind:'',units:0};if(/대체휴무|대체휴일|대휴/.test(s))return{kind:'대체휴일',units:/0[.]5|반일|반차|오전|오후/.test(s)?.5:1};if(/보상휴가/.test(s))return{kind:'보상휴가',units:/0[.]5|반차|오전|오후/.test(s)?.5:1};if(/반차/.test(s))return{kind:'반차',units:.5};if(/연차/.test(s))return{kind:'연차',units:1};if(/휴무|휴일/.test(s))return{kind:'휴무',units:1};if(/공가|병가|경조|휴가/.test(s))return{kind:'기타휴가',units:1};if(/교육/.test(s))return{kind:'교육',units:0};if(/근무/.test(s))return{kind:'근무',units:0};return{kind:'기타',units:0};};
function asMinutes(x){if(x===null||x===undefined||x==='')return null;if(typeof x==='number'&&Number.isFinite(x)){let v=((x%1)+1)%1;return Math.round(v*1440)%1440;}let s=String(x).trim(),m=[...s.matchAll(/(\d{1,2}):(\d{2})(?::\d{2})?/g)].at(-1);if(m)return (+m[1]<=23&&+m[2]<=59)? +m[1]*60+ +m[2]:null; if(/^\d{3,4}$/.test(s)){s=s.padStart(4,'0');let h=+s.slice(0,2),mi=+s.slice(2);return h<24&&mi<60?h*60+mi:null;}return null;}
const timeFmt=m=>m==null?'':`${String(Math.floor(m/60)).padStart(2,'0')}:${String(m%60).padStart(2,'0')}`;
const checkinScore=m=>m<=10?-1:m<60?-2:-3;
const dateInMonth=(d,mo)=>d&&d.slice(0,7)===mo;
const dayArrToObj=(arr)=>Object.fromEntries(arr.map(v=>[v.date,v]));
const escapeHTML=v=>String(v??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const keyEvent=e=>`${e.date}|${e.start}|${e.end}|${e.title}|${e.amount}`;

function repairRange(sh){let maxR=0,maxC=0;for(const key of Object.keys(sh)){if(/^[A-Z]+[1-9][0-9]*$/.test(key)){let c=XLSX.utils.decode_cell(key);maxR=Math.max(maxR,c.r);maxC=Math.max(maxC,c.c);}}sh['!ref']=XLSX.utils.encode_range({s:{r:0,c:0},e:{r:maxR,c:maxC}});}
function readPlan(workbook,month){const sh=workbook.Sheets[workbook.SheetNames[0]];if(sh)repairRange(sh);if(!sh)throw Error('매장근무계획 시트가 없습니다.');const rows=XLSX.utils.sheet_to_json(sh,{header:1,defval:'',raw:true});const head=rows[0]||[];const cols=[];head.forEach((h,j)=>{let m=String(h).match(/^\s*(\d{1,2})일\s*$/);if(m){let d=+m[1],max=new Date(Number(month.slice(0,4)),Number(month.slice(5)),0).getDate();if(d<=max)cols.push([j,`${month}-${String(d).padStart(2,'0')}`]);}});if(!cols.length)throw Error('매장근무계획에서 01일~31일 헤더를 찾지 못했습니다.');let people=[],plan=[],issues=[];for(let ri=1;ri<rows.length;ri++){let r=rows[ri];let name=trim(r[3]),id=trim(r[1]);if(!name)continue;if(!id)throw Error(`근무계획 ${ri+1}행 ${name}: 사번이 없어 직원을 연결할 수 없습니다.`);let store=trim(r[0]);people.push({id,name,store,recordRow:ri+1});for(let [col,date] of cols){let value=trim(r[col]);if(!value)continue;plan.push({id,name,store,date,value,row:ri+1});}}return{people,plan,issues};}
function readAttendance(workbook){const sh=workbook.Sheets[workbook.SheetNames[0]];if(sh)repairRange(sh);if(!sh)throw Error('근태관리 시트가 없습니다.');const rows=XLSX.utils.sheet_to_json(sh,{header:1,defval:'',raw:true});if(!String(rows[0]?.[0]).includes('이름'))throw Error('근태관리의 A열 이름 헤더를 확인해주세요.');let data=[];for(let i=1;i<rows.length;i++){let r=rows[i],date=isoDate(r[1]),name=trim(r[0]);if(!name||!date)continue;data.push({id:trim(r[3]),name,date,inTime:asMinutes(r[2]),outTime:asMinutes(r[4]),editedIn:asMinutes(r[5]),editedOut:asMinutes(r[6]),inStore:trim(r[7]),outStore:trim(r[8]),row:i+1});}return data;}
function readLeave(workbook){const sh=workbook.Sheets[workbook.SheetNames[0]];if(sh)repairRange(sh);if(!sh)throw Error('연차신청현황 시트가 없습니다.');const rows=XLSX.utils.sheet_to_json(sh,{header:1,defval:'',raw:true});if(!String(rows[0]?.[2]).includes('이름'))throw Error('연차신청현황의 C열 이름 헤더를 확인해주세요.');let data=[];for(let i=1;i<rows.length;i++){let r=rows[i],name=trim(r[2]),start=isoDate(r[6]),end=isoDate(r[7]);if(!name||!start)continue;data.push({id:trim(r[3]),name,start,end:end||start,count:Number(r[8])||0,type:trim(r[9]),status:trim(r[4]),requested:isoDate(r[5]),row:i+1});}return data;}
function readHours(workbook){let sh=workbook.Sheets['영업시간DB']||workbook.Sheets[workbook.SheetNames[0]];if(sh)repairRange(sh);if(!sh)throw Error('영업시간DB 시트가 없습니다.');const rows=XLSX.utils.sheet_to_json(sh,{header:1,defval:'',raw:true});if(!String(rows[0]?.[0]).includes('점포'))throw Error('영업시간DB의 A열 점포명 헤더를 확인해주세요.');let times={};for(let i=1;i<rows.length;i++){let r=rows[i],store=canonStore(r[0]);if(!store)continue;times[store]={A:asMinutes(r[3]),B:asMinutes(r[4]),C:asMinutes(r[5])};}return times;}
function identityMaps(plan){const byId=new Map(),byName=new Map(),persons=new Map(),issues=[];for(const p of plan){if(!persons.has(p.id))persons.set(p.id,{id:p.id,name:p.name,store:p.store});else if(persons.get(p.id).name!==p.name)issues.push({kind:'사번/이름 충돌',name:p.name,date:'',info:`${p.id} · ${persons.get(p.id).name} / ${p.name}`});byId.set(p.id,p.id);if(!byName.has(p.name))byName.set(p.name,new Set());byName.get(p.name).add(p.id);}function resolve(id,name,date='',source=''){let canonical=byId.get(id);if(canonical){if(name&&persons.get(canonical)?.name!==name){issues.push({kind:'사번·이름 불일치',name,date,info:`${source} 사번 ${id} → ${persons.get(canonical)?.name}, 행 이름 ${name}`});return null;}return canonical;}let choices=byName.get(name);if(choices?.size===1){let candidate=[...choices][0];issues.push({kind:'다른 사번 연결',name,date,info:`${source} ${id||'(없음)'} → 계획 사번 ${candidate} (이름 동일)`});return candidate;}issues.push({kind:'직원 연결 불가',name,date,info:`${source} · ${id||'(사번 없음)'}`});return null;}return{persons,resolve,issues};}
function combineMonths(state){let plan=[],attendance=[],leave=[];for(let [month,data] of Object.entries(state.months||{})){plan.push(...(data.plan||[]));attendance.push(...(data.attendance||[]));leave.push(...(data.leave||[]));}return{plan,attendance,leave};}
function classifyHour(schedule,hours,store){let shift=shiftOf(schedule);if(!shift)return{shift:'',basis:null};let lookup=hours[canonStore(store)];return{shift,basis:lookup?.[shift]??null};}
function resolveScope(state,month){const merged=combineMonths(state);const thisMonth=state.months[month];if(!thisMonth)throw Error('선택한 분석월 자료가 없습니다.');let identities=identityMaps(merged.plan);let plans=new Map(),people=new Map(),issues=identities.issues;
 for(let p of merged.plan){let id=identities.resolve(p.id,p.name,p.date,'근무계획');if(!id)continue;let k=uid(id,p.date,'plan');if(plans.has(k)&& (plans.get(k).value!==p.value||plans.get(k).store!==p.store))issues.push({kind:'동일 날짜 근무계획 중복',name:p.name,date:p.date,info:`${plans.get(k).value} / ${p.value} · ${plans.get(k).store} / ${p.store}`});plans.set(k,{...p,id});if(!people.has(id))people.set(id,{...p});if(dateInMonth(p.date,month))people.set(id,{...p});}
 let attendMap=new Map();for(let a of merged.attendance){let id=identities.resolve(a.id,a.name,a.date,'근태관리');if(!id)continue;let k=uid(id,a.date,'att');if(!attendMap.has(k))attendMap.set(k,[]);attendMap.get(k).push({...a,id});}
 let leaves=[];for(let a of merged.leave){let id=identities.resolve(a.id,a.name,a.start,'연차신청현황');if(id)leaves.push({...a,id});}
 return{plans,people,attendMap,leaves,issues,month};}
const overrideKey=(type,id,date)=>`${type}|${id}|${date}`;
function makeFindings(scope,month,hours,cutoff,useModified,overrides){let late=[],missingIn=[],missingOut=[],issues=[...scope.issues],extra=[];
const currPlan=[...scope.plans.values()].filter(p=>dateInMonth(p.date,month)&&p.date<=cutoff);
for(let p of currPlan){let key=uid(p.id,p.date,'att'),records=scope.attendMap.get(key)||[],type=vacationType(p.value),plannedWork=type.kind==='근무',isEducation=type.kind==='교육',allPunched=records.filter(r=>r.inTime!=null||r.editedIn!=null||r.outTime!=null||r.editedOut!=null);let rIn=allPunched.map(r=>({r,at:useModified&&r.editedIn!=null?r.editedIn:r.inTime})).filter(x=>x.at!=null).sort((a,b)=>a.at-b.at)[0];let rOut=allPunched.map(r=>({r,at:useModified&&r.editedOut!=null?r.editedOut:r.outTime})).filter(x=>x.at!=null).sort((a,b)=>b.at-a.at)[0];let cin=rIn?.at??null,cout=rOut?.at??null,shift=classifyHour(p.value,hours,p.store),base={id:p.id,name:p.name,store:p.store,date:p.date,plan:p.value,shift:shift.shift,scheduled:shift.basis,checkin:cin,checkout:cout,actualStore:rIn?.r.inStore||'',sourceCount:records.length};
 if(plannedWork&&!isEducation&&cin==null){let key=overrideKey('출근 미등록',p.id,p.date);missingIn.push({...base,key,score:-3,reason:'근무계획상 근무 · 출근 기록 없음',override:overrides[key]||''});}
 if(plannedWork&&cin!=null){if(shift.shift&&shift.basis==null)issues.push({kind:'출근시간 DB 미설정',name:p.name,date:p.date,info:`${p.store} · ${shift.shift}조`});else if(shift.shift&&shift.basis!=null&&cin>=shift.basis){let delta=cin-shift.basis,score=checkinScore(delta),key=overrideKey('지각',p.id,p.date);late.push({...base,delta,score,key,reason:`${delta}분 지각 (기준시각 포함)`,override:overrides[key]||''});}
 else if(!shift.shift)issues.push({kind:'근무조 미지정',name:p.name,date:p.date,info:`매장계획 ${p.value} · 출근 여부만 확인`});}
 if(records.length>1)issues.push({kind:'근태 중복 기록',name:p.name,date:p.date,info:`${records.length}행 / 최초 출근·최종 퇴근으로 판정`});
 if(cin!=null&&cout==null){let key=overrideKey('퇴근 미등록',p.id,p.date);missingOut.push({...base,key,score:0,reason:'출근 기록 있으나 퇴근 기록 없음',override:overrides[key]||''});}
 if(!plannedWork&&cin!=null&&!isEducation)extra.push({...base,reason:`${type.kind||'기타'} 계획일 출근 기록`});
}
// An attendance entry not tied to current plan still needs an out-punch check.
for(let [k,records] of scope.attendMap){let rep=records[0],date=rep.date;if(!dateInMonth(date,month)||date>cutoff)continue;if(scope.plans.has(uid(rep.id,date,'plan')))continue;let inAny=records.some(r=>(useModified?r.editedIn??r.inTime:r.inTime)!=null),outAny=records.some(r=>(useModified?r.editedOut??r.outTime:r.outTime)!=null);if(inAny&&!outAny){let i=records.find(r=>(useModified?r.editedIn??r.inTime:r.inTime)!=null),key=overrideKey('퇴근 미등록',rep.id,date);missingOut.push({id:rep.id,name:rep.name,store:i.inStore||'',date,plan:'계획없음',checkin:(useModified?i.editedIn??i.inTime:i.inTime),checkout:null,shift:'',scheduled:null,key,score:0,reason:'근무계획에 없는 출근 기록 · 퇴근 누락',override:overrides[key]||''});}issues.push({kind:'근무계획 외 근태 기록',name:rep.name,date,info:`사번 ${rep.id}, 계획 없음`});}
return{late,missingIn,missingOut,issues,extra};}
function eventInfoFor(p,event,scope,hires){let hire=hires[p.id]||'',at=scope.plans.get(uid(p.id,event.date,'plan'));
if(hire&&hire>event.date)return{eligibility:'발생 대상 아님',holidayPlan:at?.value||'',offOnEvent:false,workedOnEvent:false,check:'입사일이 발생일 이후'};
if(!at)return{eligibility:'발생일 근무계획 미확인',holidayPlan:'',offOnEvent:false,workedOnEvent:false,check:'발생일 근무계획 자료 필요'};
let type=vacationType(at.value);let offOnEvent=type.kind==='휴무',workedOnEvent=type.kind==='근무'||type.kind==='교육';return{eligibility:hire?'입사일 확인':'근무계획 기반 (입사일 미등록)',holidayPlan:at.value,offOnEvent,workedOnEvent,check:!offOnEvent&&!workedOnEvent?'발생일 휴무·근무 여부 확인 필요':''};}
function buildHolidays(scope,month,events,hires,cutoff){let rows=[],issues=[],entitlements=new Map();let eventRelevance=events.filter(e=>e.date.slice(0,7)<=month&&e.end.slice(0,7)>=Math.min(month,e.date.slice(0,7)) || (e.start.slice(0,7)<=month&&e.end.slice(0,7)>=month));const currentPeople=[...scope.people.values()].filter(p=>[...scope.plans.values()].some(d=>d.id===p.id&&dateInMonth(d.date,month)));for(let event of events){if(event.date.slice(0,7)>month&&event.start.slice(0,7)>month)continue;for(let p of currentPeople){let e=eventInfoFor(p,event,scope,hires),all=[...scope.plans.values()].filter(d=>d.id===p.id&&d.date>=event.start&&d.date<=event.end&&vacationType(d.value).kind==='대체휴일');let useDates=all.map(d=>d.date),used=all.reduce((s,d)=>s+vacationType(d.value).units,0);let relatedDate=event.date<=cutoff;let status='';
if(e.eligibility==='발생 대상 아님'){status=used>0?'발생 대상 아님·사용 확인':'발생 대상 아님';}
else if(!relatedDate||!e.holidayPlan){status=used>0?'발생일 자료대기·사전사용':'발생일 자료대기';}
else if(e.offOnEvent){status=used>0?'공휴일 휴무 + 대체휴일 중복 의심':'공휴일 당일 휴무 · 정상';}
else if(e.workedOnEvent){status=used>event.amount?'대체휴일 초과 사용':used===event.amount?'대체휴일 사용 완료':used===0?(event.end>cutoff?'사용기간 진행중 · 미사용':'대체휴일 미소진'):`일부 사용 (${used}/${event.amount})`;}
else status='당일 근무 구분 확인';
if(!hires[p.id]&&e.eligibility!=='발생 대상 아님')issues.push({kind:'입사일 확인 필요',name:p.name,date:event.date,info:`${event.title||'대체휴일'} 발생대상 검증을 위해 입사일을 등록해주세요.`});if(e.check&&status!=='발생 대상 아님')issues.push({kind:'대체휴일 발생자 확인',name:p.name,date:event.date,info:`${status} · ${e.check}`});
if(status.includes('중복')||status.includes('초과')||status.includes('대상 아님·사용'))issues.push({kind:'대체휴일 사용 검토',name:p.name,date:event.date,info:`${status}: ${useDates.join(', ')}`});
rows.push({event:event.title||'대체휴일',eventDate:event.date,start:event.start,end:event.end,id:p.id,name:p.name,store:p.store,hire:hires[p.id]||'',plan:e.holidayPlan,used,useDates:useDates.join(', '),eligibility:e.eligibility,status,amount:event.amount});
if(e.eligibility!=='발생 대상 아님'&&e.offOnEvent){let wk=monday(event.date),key=uid(p.id,wk,'weekly'),previous=entitlements.get(key)||0;entitlements.set(key,previous+event.amount);}else if(e.eligibility!=='발생 대상 아님'&&e.workedOnEvent){let chosen=[...new Set(useDates)].slice(0,Math.ceil(event.amount));for(let date of chosen){let key=uid(p.id,monday(date),'weekly'),previous=entitlements.get(key)||0;entitlements.set(key,previous+Math.min(1,event.amount));}} }
}
// If 2+ events have overlapping usage windows, the cause of a make-up day cannot be inferred from the plan alone.
for(let e of events){let overlap=events.filter(x=>x!==e&&x.start<=e.end&&x.end>=e.start);if(overlap.length)issues.push({kind:'대체휴일 규칙 기간 중첩',name:'전체',date:e.date,info:`${e.title||'공휴일'} · 겹치는 발생 규칙 ${overlap.map(x=>x.title||x.date).join(', ')} (사용처 귀속 확인 필요)`});}
return{rows,issues,entitlements};}
function buildWeeks(scope,month,cutoff,holidayData){const seen=new Set(),rows=[],issues=[];const allPlans=[...scope.plans.values()],planIdx=scope.plans;let curr=allPlans.filter(p=>dateInMonth(p.date,month));for(let p of curr){let week=monday(p.date),key=uid(p.id,week,'weekly');if(seen.has(key))continue;seen.add(key);let end=dayStep(week,6),days=rangeDates(week,end,7),known=days.map(d=>planIdx.get(uid(p.id,d,'plan'))).filter(Boolean);let types=known.map(k=>vacationType(k.value));let rest=0,sub=0,annual=0,half=0,other=0;types.forEach(t=>{if(t.kind==='휴무')rest+=t.units;else if(t.kind==='대체휴일')sub+=t.units;else if(t.kind==='연차')annual+=t.units;else if(t.kind==='반차')half+=t.units;else if(['보상휴가','기타휴가'].includes(t.kind))other+=t.units;});let total=rest+sub+annual+half+other;let completed=known.length===7&&end<=cutoff;let holidayExtra=holidayData.entitlements.get(key)||0;let limit=2+holidayExtra;let status=!completed?'월경계/일자 자료대기':rest>limit?'정기휴무 초과 확인':rest<2?'정기휴무 부족 확인':rest>2?'정상 (공휴일 휴무 예외)':sub>0?'정기휴무 정상 · 대체휴일 별도확인':'정상';if(!completed&&days.some(d=>d>cutoff))status='주차 미확정 (다음 분석 필요)';let item={id:p.id,name:p.name,store:p.store,start:week,end,known:known.length,rest,sub,annual,half,other,total,holidayExtra,limit,status};rows.push(item);if(status.includes('초과')||status.includes('부족'))issues.push({kind:'주차 휴무 검토',name:p.name,date:week,info:`${status} · ${total}일 / 공휴일 예외 포함 기준 ${limit}일`});}
return{rows,issues};}
function leaveAudit(scope,month){let rows=[],issues=[];for(let l of scope.leaves){if(l.end.slice(0,7)<month||l.start.slice(0,7)>month)continue;let days=rangeDates(l.start,l.end,100),schedule=days.map(d=>scope.plans.get(uid(l.id,d,'plan'))).filter(Boolean);let leavePlan=schedule.filter(p=>{let k=vacationType(p.value).kind;return ['연차','반차','대체휴일','기타휴가','보상휴가'].includes(k);});let status=l.status!=='승인'?'승인 대기/미확정':!leavePlan.length?'계획 대조 필요':'계획에 휴가 표기됨';let item={...l,schedule:leavePlan.map(p=>p.date+':'+p.value).join(' / '),compare:status};rows.push(item);if(status!=='계획에 휴가 표기됨')issues.push({kind:'연차신청현황 대조',name:l.name,date:l.start,info:`${l.type} ${l.status} · ${status}`});}return{rows,issues};}
function buildLeaveLedger(scope,month,cutoff){let items=[],seen=new Set();for(let p of scope.plans.values()){
 if(p.date.slice(0,4)!==month.slice(0,4)||p.date.slice(0,7)>month)continue;
 let type=vacationType(p.value);if(!['연차','반차'].includes(type.kind))continue;
 let k=uid(p.id,p.date,'leave-use');if(seen.has(k))continue;seen.add(k);
 let matching=scope.leaves.filter(a=>a.id===p.id&&a.start<=p.date&&a.end>=p.date&&/연차|반차/.test(a.type));
 let approval=matching.some(a=>a.status==='승인')?'신청 승인':matching.length?'승인 대기':'신청 내역 미확인';
 items.push({id:p.id,name:p.name,store:p.store,date:p.date,type:type.kind,units:type.units,plan:p.value,approval,status:p.date>cutoff?'사용 예정':'계획상 사용'});
}items.sort((a,b)=>a.id.localeCompare(b.id)||a.date.localeCompare(b.date));let totals=new Map();for(let r of items){let key=r.id,now=totals.get(key)||0;r.yearCumulative=now+r.units;totals.set(key,r.yearCumulative);}
return items.sort((a,b)=>a.date.localeCompare(b.date)||a.name.localeCompare(b.name,'ko'));}
function statusFromOverride(raw,allowed){let v=trim(raw).toUpperCase();if(['인정','소명인정','감점제외','제외','정시인정','출근인정','퇴근인정'].includes(v))return'제외';if(allowed&&['무단결근확정','무단결근 확정','무단결근'].includes(v))return'무단결근 확정';return'';}
function applyScores(findings,overrides){
 let late=findings.late.map(v=>({...v,excluded:statusFromOverride(overrides[v.key],false)==='제외'}));
 let missingIn=findings.missingIn.map(v=>{let flag=statusFromOverride(overrides[v.key],true);return {...v,excluded:flag==='제외',score:flag==='무단결근 확정'?-20:-3,flag};});
 let missingOut=findings.missingOut.map(v=>({...v,excluded:statusFromOverride(overrides[v.key],false)==='제외'}));
 const perId=new Map();
 function record(id){if(!perId.has(id))perId.set(id,{late:0,missingIn:0,missingOut:0,missOutCount:0,penalty:0});return perId.get(id);}
 for(let f of [...late,...missingIn,...missingOut])record(f.id);
 const lateById=new Map();
 for(let f of late){if(f.excluded){f.score=0;continue;}let items=lateById.get(f.id)||[];items.push(f);lateById.set(f.id,items);}
 for(let [id,items] of lateById){let p=record(id),count=items.length;p.late=count;for(let f of items){f.score=f.delta<=10&&count<3?0:checkinScore(f.delta);p.penalty+=f.score;}}
 for(let f of missingIn){if(f.excluded){f.score=0;continue;}let p=record(f.id);p.missingIn++;p.penalty+=f.score;}
 const outById=new Map();
 for(let f of missingOut){if(f.excluded){f.score=0;f.sequence=0;continue;}let items=outById.get(f.id)||[];items.push(f);outById.set(f.id,items);}
 for(let [id,items] of outById){let p=record(id);p.missingOut=items.length;p.missOutCount=items.length;let count=0;for(let f of items.sort((a,b)=>a.date.localeCompare(b.date))){count++;f.sequence=count;f.cumulativeScore=count>=3?-count:0;f.score=items.length>=3?-1:0;}if(count>=3)p.penalty-=count;}
 return{late,missingIn,missingOut,perId};
}
function analyze(state,month,hours,cutoff,useModified){let scope=resolveScope(state,month),findings=makeFindings(scope,month,hours,cutoff,useModified,state.overrides||{}),holiday=buildHolidays(scope,month,state.holidays||[],state.hires||{},cutoff),weekly=buildWeeks(scope,month,cutoff,holiday),leave=leaveAudit(scope,month),scored=applyScores(findings,state.overrides||{}),summaries=[];let ids=[...new Set(state.months[month].plan.map(p=>p.id))];for(let id of ids){let person=scope.people.get(id),points=scored.perId.get(id)||{late:0,missingIn:0,missingOut:0,penalty:0};if(!person)continue;let weekIssues=weekly.rows.filter(r=>r.id===id&&(r.status.includes('초과')||r.status.includes('부족')));let hIssues=holiday.rows.filter(r=>r.id===id&&(r.status.includes('중복')||r.status.includes('초과')||r.status.includes('대상 아님·사용')));summaries.push({...person,...points,weekCheck:weekIssues.length,holidayCheck:hIssues.length,warning:points.penalty<=-5?'경고장 검토':'-',month});}
let leaveLedger=buildLeaveLedger(scope,month,cutoff);let issues=[...findings.issues,...holiday.issues,...weekly.issues,...leave.issues];let uniqIssues=[],seen=new Set();for(let i of issues){let k=JSON.stringify(i);if(!seen.has(k)){seen.add(k);uniqIssues.push(i);}}return{month,cutoff,summaries,scored,weekly:weekly.rows,holiday:holiday.rows,leave:leave.rows,leaveLedger,issues:uniqIssues,events:state.holidays,hires:state.hires};}

// Spreadsheet report helpers. All numeric scoring is generated from the same scored model as the dashboard.
const C={dark:'203960',green:'35557A',mint:'E9EEF6',pale:'F4F6F9',white:'FFFFFF',text:'303E53',muted:'7C8799',red:'B23E45',rose:'FFF0F0',amber:'DFA34C',sky:'EAF3FA',line:'DEE9E1'};
const sFill=rgb=>({fgColor:{rgb}});
const styleHeader={font:{name:'맑은 고딕',bold:true,color:{rgb:C.white},sz:10},fill:sFill(C.dark),alignment:{vertical:'center',horizontal:'center',wrapText:true},border:{bottom:{style:'thin',color:{rgb:C.green}}}};
const styleBody=(i,bad)=>({font:{name:'맑은 고딕',color:{rgb:bad?C.red:C.text},sz:10},fill:sFill(bad?C.rose:(i%2?C.white:C.pale)),alignment:{vertical:'center',wrapText:true},border:{bottom:{style:'hair',color:{rgb:C.line}}}});
const maxWidth=(s)=>{let t=String(s??'');return Math.min(40,Math.max(10,Math.ceil([...t].length*1.65)+2));};
function addSheet(wb,title,headers,items,widths=[],opts={}){let top=[`${title}  |  ${opts.month||''}  ·  ZENIEL`],intro=[opts.note||`${opts.month||''} · ${title} · ${VERSION}`];let rows=[[...top], [...intro],[],headers,...items];let ws=XLSX.utils.aoa_to_sheet(rows);let n=headers.length,end=Math.max(4,rows.length),last=XLSX.utils.encode_col(n-1);ws['!merges']=[{s:{r:0,c:0},e:{r:0,c:n-1}},{s:{r:1,c:0},e:{r:1,c:n-1}}];ws['!cols']=headers.map((h,i)=>({wch:widths[i]||Math.max(11,Math.min(24,maxWidth(h)))}));ws['!rows']=[{hpt:33},{hpt:25},{hpt:9},{hpt:28},...items.map(()=>({hpt:25}))];ws['!autofilter']={ref:`A4:${last}${end}`};for(let r=0;r<=1;r++){let cell=ws[XLSX.utils.encode_cell({r,c:0})];if(cell)cell.s={fill:sFill(r===0?C.dark:C.mint),font:{name:'맑은 고딕',sz:r===0?16:10,bold:r===0,color:{rgb:r===0?C.white:C.green}},alignment:{vertical:'center',horizontal:'left'}};}
for(let c=0;c<n;c++){let ad=XLSX.utils.encode_cell({r:3,c});if(ws[ad])ws[ad].s=clone(styleHeader);}for(let r=4;r<rows.length;r++){let vals=rows[r];let bad=opts.badRow?opts.badRow(vals):false;for(let c=0;c<n;c++){let ad=XLSX.utils.encode_cell({r,c});if(ws[ad])ws[ad].s=styleBody(r,bad);}}
XLSX.utils.book_append_sheet(wb,ws,title.slice(0,31));return ws;}
const stamp=d=>d||'';
function buildWorkbookLegacy(result,state){let wb=XLSX.utils.book_new(),month=result.month;let summary=result.summaries.sort((a,b)=>a.store.localeCompare(b.store,'ko')||a.name.localeCompare(b.name,'ko'));addSheet(wb,'최종 요약',['점포','사번','이름','지각(회)','출근 미등록','퇴근 미등록','지각/미등록 감점','경고장','휴무 검토','대체휴일 검토'],summary.map(r=>[r.store,r.id,r.name,r.late,r.missingIn,r.missingOut,r.penalty,r.warning,r.weekCheck,r.holidayCheck]),[21,14,13,14,15,15,19,14,14,16],{month,badRow:r=>r[6]<0});
let records=(arr,extra)=>arr.slice().sort((a,b)=>a.date.localeCompare(b.date)||a.name.localeCompare(b.name,'ko')).map(x=>[x.date,x.store,x.id,x.name,x.plan,x.shift||'',timeFmt(x.scheduled),timeFmt(x.checkin),timeFmt(x.checkout),...(extra(x)),x.excluded?'소명 인정':'자동판정']);
addSheet(wb,'지각',['날짜','계획 점포','사번','이름','매장계획','근무조','기준시간','출근','퇴근','지각(분)','감점','상태'],records(result.scored.late,r=>[r.delta,r.excluded?0:r.score]),[14,20,14,13,16,10,12,11,11,12,10,14],{month,badRow:r=>r[10]<0});
addSheet(wb,'출근 미등록',['날짜','점포','사번','이름','매장계획','근무조','기준시간','출근','퇴근','사유','감점','상태'],records(result.scored.missingIn,r=>[r.flag==='무단결근 확정'?'무단결근 확정':r.reason,r.excluded?0:r.score]),[14,20,14,13,16,10,12,11,11,39,10,14],{month,badRow:r=>r[10]<0});
addSheet(wb,'퇴근 미등록',['날짜','점포','사번','이름','매장계획','근무조','기준시간','출근','퇴근','사유','해당 회차 누적감점','상태'],records(result.scored.missingOut,r=>[r.reason,r.excluded?0:r.score]),[14,20,14,13,16,10,12,11,11,32,22,14],{month,note:'퇴근누락 월 1~2회 0점 / 3회 -3점 / 4회 -4점. 각 행은 해당 회차까지 누적감점(합산 금지). 개인별 합계는 최종 요약 확인',badRow:r=>r[10]<0});
addSheet(wb,'주간 휴무',['주차 시작','주차 종료','점포','사번','이름','확인일수','정기휴무','대체휴일','연차','반차','기타휴가','쉬는날 합계','공휴일 예외','판정기준','판정'],result.weekly.map(r=>[r.start,r.end,r.store,r.id,r.name,r.known,r.rest,r.sub,r.annual,r.half,r.other,r.total,r.holidayExtra,r.limit,r.status]),[14,14,20,14,13,12,12,13,10,10,13,15,12,12,30],{month,note:'주차는 월~일, 전월/다음 월 계획 누적 연결. 미확정 주차는 단정하지 않음.',badRow:r=>String(r[14]).includes('초과')});
addSheet(wb,'대체휴일 판정',['공휴일','발생일','사용 시작','사용 종료','점포','사번','이름','입사일','당일 매장계획','발생 개수','대체휴일 사용 개수','사용일','발생 자격','판정'],result.holiday.map(r=>[r.event,r.eventDate,r.start,r.end,r.store,r.id,r.name,r.hire,r.plan,r.amount,r.used,r.useDates,r.eligibility,r.status]),[18,14,14,14,20,14,12,13,20,14,16,35,26,36],{month,note:'대체휴일 발생 기준 수동 등록. 이벤트가 중첩되면 어떤 행사로 사용했는지 확인 필요.',badRow:r=>/중복|초과|대상 아님·사용/.test(r[13])});
addSheet(wb,'연차 신청 대조',['시작일','종료일','사번','이름','신청유형','신청 사용개수','승인상태','계획 휴가 내용','대조 결과'],result.leave.map(r=>[r.start,r.end,r.id,r.name,r.type,r.count,r.status,r.schedule,r.compare]),[14,14,14,12,18,18,14,42,23],{month,note:'연차 신청 승인 상태와 매장근무계획 대조. 연차 발생·수당 지급 수량은 별도 인력 기초DB 필요.',badRow:r=>r[8]!=='계획에 휴가 표기됨'});
addSheet(wb,'연차 사용 누적',['사용일','점포','사번','이름','종류','차감일수','매장계획','연차신청','예정/기록','연도 누적'],result.leaveLedger.map(r=>[r.date,r.store,r.id,r.name,r.type,r.units,r.plan,r.approval,r.status,r.yearCumulative]),[15,20,14,13,13,12,18,19,17,14],{month,note:'해당 연도에 업로드한 매장계획 자료만 누적합니다. 예정 기록은 사용확정과 다르며 연차 잔여·수당 계산은 별도 기준DB가 필요합니다.'});
const overrideRows=[...result.scored.late.map(r=>['지각',r]),...result.scored.missingIn.map(r=>['출근 미등록',r]),...result.scored.missingOut.map(r=>['퇴근 미등록',r])].map(([type,r])=>[type,r.date,r.id,r.name,r.reason||r.plan,r.excluded?'인정':r.flag==='무단결근 확정'?'무단결근 확정':'자동판정',r.key,trim(state.overrides[r.key]),'처리결과는 웹앱 검토 영역에서 지정 후 다시 다운로드']);
addSheet(wb,'소명 및 판정',['구분','날짜','사번','이름','원인','상태','일치키','웹앱 입력값','안내'],overrideRows,[16,14,14,13,40,20,45,20,58],{month,note:'자동판정 유지. 웹앱 검토 영역에서 소명인정/무단결근 확정을 선택하면 요약 재계산 후 다운로드됩니다.'});
addSheet(wb,'검증 필요',['구분','이름','날짜','내용'],result.issues.map(r=>[r.kind,r.name,r.date,r.info]),[27,15,15,76],{month});
addSheet(wb,'대체휴일 설정',['발생 사유','발생일','개수','사용 시작일','사용 종료일'],state.holidays.map(r=>[r.title,r.date,r.amount,r.start,r.end]),[25,16,13,18,18],{month,note:'사용기간은 발생일 전부터 시작 가능. 입사일이 발생일 이후인 직원은 미발생.'});
addSheet(wb,'입사일 설정',['사번','이름','입사일'],Object.entries(state.hires).map(([id,date])=>[id,summary.find(x=>x.id===id)?.name||'',date]),[18,17,18],{month,note:'입사일 없는 직원은 발생 대상 여부를 단정하지 않고 검토 대상으로 남김.'});
addSheet(wb,'이용 안내',['항목','설명'],[
['월별 사용','매장근무계획/근태/연차신청/영업시간DB를 올리고 분석월을 선택합니다.'],
['다음달 이월','이 결과 엑셀을 보관하고 다음 달에는 이전 결과 파일로 함께 업로드합니다.'],
['주차 계산','월요일~일요일 기준. 월 경계 주차는 다음 월 업로드 후 확정 판정합니다.'],
['지각 기준','기준시각 도착=지각. 0~10분 -1점, 11~60분 -2점, 61분 이상 -3점.'],
['퇴근누락','1~2회 0점, 3회 -3점, 4회 -4점, 이후 횟수만큼 감점.'],
['무단결근','출근 미등록은 -3점, 무단결근 확정은 웹앱에서 수동 선택 시 -20점.'],
['교육','매장계획에 교육이 포함된 날은 지각 제외. 타매장 출근 여부는 무시.'],
['대체휴일 중복','발생일에 휴무인데 다른 날 동일 기간에 대체휴일이 있으면 검토. 겹치는 공휴일은 원인 검증 필요.'],
['연차','승인 내역과 매장계획 교차검증만 수행. 연차수당 산정은 별도 인력DB 필요.'],
['저장방식','브라우저 로컬에서 계산하고 이전 기록은 보고서의 숨김 상태 시트에 담습니다. 서버 저장 없음.'],
], [26,94],{month});
const data=JSON.stringify(state);let chunks=data.match(/[\s\S]{1,29000}/g)||[''];let st=XLSX.utils.aoa_to_sheet([['ZENIEL_STATE','1.0.0'],...chunks.map((chunk,i)=>[i+1,chunk])]);XLSX.utils.book_append_sheet(wb,st,STATE_SHEET);wb.Workbook=wb.Workbook||{};wb.Workbook.Sheets=wb.SheetNames.map(n=>({Hidden:n===STATE_SHEET?1:0}));wb.Props={Title:`코스트코 근태분석 ${month}`,Subject:'누적 근태·휴무 분석',Author:'ZENIEL Attendance System',Comments:'자동판정·검증 결과'};return wb;}

function buildWorkbook(result,state){
 const wb=buildWorkbookLegacy(result,state),month=result.month;
 const summary=result.summaries.slice().sort((a,b)=>a.store.localeCompare(b.store,'ko')||a.name.localeCompare(b.name,'ko'));
 const late=result.scored.late.slice().sort((a,b)=>a.date.localeCompare(b.date)||a.name.localeCompare(b.name,'ko'));
 const missingIn=result.scored.missingIn.slice().sort((a,b)=>a.date.localeCompare(b.date)||a.name.localeCompare(b.name,'ko'));
 const missingOut=result.scored.missingOut.slice().sort((a,b)=>a.date.localeCompare(b.date)||a.name.localeCompare(b.name,'ko'));
 const allScored=[...late.map(r=>['지각',r]),...missingIn.map(r=>['출근 미등록',r]),...missingOut.map(r=>['퇴근 미등록',r])].sort((a,b)=>a[1].date.localeCompare(b[1].date)||a[1].name.localeCompare(b[1].name,'ko'));
 const end=Math.max(5,allScored.length+4);
 const ref=col=>`'소명 및 판정'!$${col}$5:$${col}$${end}`;
 const formula=(ws,a,s)=>{let p=ws[a]||{};ws[a]={...p,t:'n',f:s.replace(/^=/,''),v:0};};
 const formulaText=(ws,a,s)=>{let p=ws[a]||{};ws[a]={...p,t:'s',f:s.replace(/^=/,''),v:''};};
 function replaceSheet(title,headers,items,widths,opts){let temp=XLSX.utils.book_new();addSheet(temp,title,headers,items,widths,opts);wb.Sheets[title]=temp.Sheets[title];return wb.Sheets[title];}
 const reviewRows=allScored.map(([type,r])=>{
  const basic=type==='퇴근 미등록'?0:type==='출근 미등록'?-3:checkinScore(r.delta);
  const orig=trim(state.overrides?.[r.key]||'');
  const confirmation=statusFromOverride(orig,true)==='제외'?'제외':statusFromOverride(orig,true)==='무단결근 확정'?'무단결근 확정':'';
  return[type,r.date,r.id,r.name,r.reason||r.plan,'자동판정',r.key,confirmation,'자동판정',basic,r.excluded?0:r.score,'H열에 제외 입력 → 자동 반영',state.notes?.[r.key]||''];
 });
 const review=replaceSheet('소명 및 판정',
  ['구분','날짜','사번','이름','원인','자동판정','일치키','확인자','최종처리','기본점수','최종반영감점','작성 안내','소명사유'],
  reviewRows,[16,14,14,13,39,19,44,19,20,14,19,38,31],
  {month,note:'확인자(H열)에 "제외" 입력 시 같은 직원의 지각 총횟수도 재계산됩니다. 10분 이내 1~2회는 0점, 3회부터 건당 -1점.'});
 review['A3']={t:'s',v:'H열 [확인자]에 제외 입력 → 월 지각 횟수/전체 점수 자동 재계산  |  제외 삭제 → 복원  |  소명사유(M열) 기록',s:{font:{name:'맑은 고딕',bold:true,color:{rgb:C.green},sz:10},fill:sFill(C.mint),alignment:{vertical:'center'}}};
 review['!merges'].push({s:{r:2,c:0},e:{r:2,c:12}});review['!rows'][2]={hpt:26};
 for(let i=0;i<reviewRows.length;i++){
  const r=i+5;
  formulaText(review,`I${r}`,`IF(H${r}="제외","감점 제외",IF(H${r}="무단결근 확정","무단결근 확정","자동판정"))`);
  const lateness=`COUNTIFS(${ref('A')},"지각",${ref('C')},C${r},${ref('H')},"<>제외")`;
  const checkout=`COUNTIFS(${ref('A')},"퇴근 미등록",${ref('C')},C${r},${ref('H')},"<>제외")`;
  formula(review,`K${r}`,`IF(H${r}="제외",0,IF(A${r}="퇴근 미등록",IF(${checkout}>=3,-1,0),IF(A${r}="지각",IF(AND(J${r}=-1,${lateness}<3),0,J${r}),IF(AND(A${r}="출근 미등록",H${r}="무단결근 확정"),-20,J${r}))))`);
  const old=review[`H${r}`]||{t:'s',v:''};
  review[`M${r}`].s={...styleBody(i,false),fill:sFill('FFF4D7')};
  review[`H${r}`]={...old,s:{...styleBody(i,false),fill:sFill('FFF4D7'),font:{name:'맑은 고딕',bold:true,color:{rgb:C.green},sz:10}}};
 }
 const sumRows=summary.map(r=>[r.store,r.id,r.name,r.late,r.missingIn,r.missingOut,r.penalty,r.warning,r.weekCheck,r.holidayCheck]);
 const summ=replaceSheet('최종 요약',['점포','사번','이름','지각(회)','출근 미등록','퇴근 미등록','지각/미등록 감점','경고장','휴무 검토','대체휴일 검토'],sumRows,[21,14,13,14,15,15,19,14,14,16],{month,note:'소명 및 판정 H열 "제외" 적용 즉시 10분 이내 지각 감점과 최종 점수가 재계산됩니다.',badRow:r=>r[6]<0});
 const counts=(type,id)=>`COUNTIFS(${ref('A')},"${type}",${ref('C')},${id},${ref('H')},"<>제외")`;
 for(let i=0;i<sumRows.length;i++){const r=i+5;formula(summ,`D${r}`,counts('지각',`$B${r}`));formula(summ,`E${r}`,counts('출근 미등록',`$B${r}`));formula(summ,`F${r}`,counts('퇴근 미등록',`$B${r}`));formula(summ,`G${r}`,`SUMIF(${ref('C')},$B${r},${ref('K')})`);formulaText(summ,`H${r}`,`IF(G${r}<=-5,"경고장 검토","-")`);}
 function details(type,entries){
  return entries.map(r=>[r.date,r.store,r.id,r.name,r.plan,r.shift||'',timeFmt(r.scheduled),timeFmt(r.checkin),timeFmt(r.checkout),type==='지각'?r.delta:type==='퇴근 미등록'?r.reason:r.flag==='무단결근 확정'?'무단결근 확정':r.reason,r.excluded?0:r.score,r.excluded?'소명 인정':'자동판정',r.key]);
 }
 function replaceDetail(title,type,items,widths,note){
  const headers=['날짜',type==='지각'?'계획 점포':'점포','사번','이름','매장계획','근무조','기준시간','출근','퇴근',type==='지각'?'지각(분)':'사유',type==='퇴근 미등록'?'해당 회차 누적감점':'감점','상태','내부연결키'];
  const ws=replaceSheet(title,headers,details(type,items),[...widths,46],{month,note,badRow:r=>r[10]<0});ws['!cols'][12].hidden=true;
  for(let i=0;i<items.length;i++){
   const row=i+5,confirm=`IFERROR(INDEX(${ref('H')},MATCH($M${row},${ref('G')},0)),"")`;
   if(type==='퇴근 미등록'){const remain=`COUNTIFS(${ref('A')},"퇴근 미등록",${ref('C')},$C${row},${ref('H')},"<>제외",${ref('B')},"<="&$A${row})`;formula(ws,`K${row}`,`IF(${confirm}="제외",0,IF(${remain}>=3,-(${remain}),0))`);}
   else formula(ws,`K${row}`,`IFERROR(INDEX(${ref('K')},MATCH($M${row},${ref('G')},0)),0)`);
   formulaText(ws,`L${row}`,`IF(${confirm}="제외","소명 인정","자동판정")`);
  }
 }
 replaceDetail('지각','지각',late,[14,20,14,13,16,10,12,11,11,12,10,14],'월 지각 전체 3회부터 0~10분 지각도 건당 -1점. 11~59분 -2점, 60분 이상 -3점. H열 제외 시 재계산.');
 replaceDetail('출근 미등록','출근 미등록',missingIn,[14,20,14,13,16,10,12,11,11,39,10,14],'소명 시트 H열 제외 시 출근 미등록 감점 제외. 무단결근은 확인 후 확정.');
 replaceDetail('퇴근 미등록','퇴근 미등록',missingOut,[14,20,14,13,16,10,12,11,11,32,22,14],'퇴근누락 1~2회 0점, 3회 -3점, 4회 -4점. 상세 점수는 해당 회차 누적감점으로 합산하지 않음.');
 replaceSheet('주간 휴무',['주차 시작','주차 종료','점포','사번','이름','확인일수','정기휴무','대체휴일','연차','반차','기타휴가','쉬는날 합계','공휴일 예외','판정기준','판정'],result.weekly.map(r=>[r.start,r.end,r.store,r.id,r.name,r.known,r.rest,r.sub,r.annual,r.half,r.other,r.total,r.holidayExtra,r.limit,r.status]),[14,14,20,14,13,12,12,13,10,10,13,15,12,12,34],{month,note:'정기휴무 주 2회를 별도로 검사. 정당한 연차 및 대체휴일은 단순 휴무 초과에 합산하지 않음.',badRow:r=>String(r[14]).includes('초과')});
 replaceSheet('이용 안내',['항목','설명'],[
 ['월별 사용','원본 4개를 업로드해 분석월을 선택하고 실행합니다.'],
 ['다음달 이월','이 결과 엑셀과 다음달 원본 4개를 함께 업로드합니다.'],
 ['소명 제외','소명 및 판정 H열 확인자에 제외 입력 → 지각 횟수 및 감점 자동 재계산.'],
 ['소명 되돌리기','제외를 지우면 자동 점수가 복원됩니다.'],
 ['주차 계산','월~일 주차, 월 경계 주차는 다음 분석월과 결합해 확정.'],
 ['지각 기준','기준시간과 같은 시각도 지각. 월 전체 지각 2회 이내는 10분 이내 감점 0점. 3회부터 10분 이내 건당 -1점. 11~59분 -2점, 60분 이상 -3점. 예: 10분 이내 2회+15분 1회=-4점.'],
 ['퇴근누락','1~2회 0점, 3회 -3점, 4회 -4점, 이후 건수만큼 감점.'],
 ['교육','근무계획 교육일은 지각 제외. 다른 매장에서 출근해도 감점하지 않음.'],
 ['정기휴무·연차','정기휴무 2회와 정당한 연차·대체휴일은 따로 관리.'],
 ['대체휴일','당일 휴무와 기간 내 사용을 비교하고 중복 가능성은 검증 필요로 표기.'],
 ['다음달 소명','제외를 수정해 저장한 결과 파일을 다음달 이전 결과로 재업로드하면 소명 이력 복원.'],
 ],[26,98],{month});
 wb.Workbook=wb.Workbook||{};wb.Workbook.CalcPr={calcMode:'auto',fullCalcOnLoad:true,forceFullCalc:true};
 const stateSheet=wb.Sheets[STATE_SHEET];if(stateSheet&&stateSheet['B1'])stateSheet['B1'].v='1.0.2';

 // Save correct cached values as well as editable formulas for non-calculating viewers.
 for(let i=0;i<summary.length;i++){const r=i+5,p=summary[i];for(const [c,v] of Object.entries({D:p.late,E:p.missingIn,F:p.missingOut,G:p.penalty,H:p.warning})){if(summ[c+r])summ[c+r].v=v;}}
 for(let i=0;i<allScored.length;i++){const r=i+5,entry=allScored[i][1];review['K'+r].v=entry.excluded?0:entry.score;review['I'+r].v=entry.excluded?'감점 제외':entry.flag==='무단결근 확정'?'무단결근 확정':'자동판정';}
 for(const [title,items] of [['지각',late],['출근 미등록',missingIn],['퇴근 미등록',missingOut]]){const ws=wb.Sheets[title];for(let i=0;i<items.length;i++){const r=i+5,x=items[i];ws['K'+r].v=x.excluded?0:title==='퇴근 미등록'?(x.cumulativeScore||0):x.score;ws['L'+r].v=x.excluded?'소명 인정':'자동판정';}}
 // Calendar view highlights actual check-in times on late days, including exemptions.
 const calendarRows=summary.map(p=>[p.store,p.id,p.name,...Array.from({length:new Date(+month.slice(0,4),+month.slice(5),0).getDate()},(_,i)=>{let d=month+'-'+String(i+1).padStart(2,'0'),x=late.find(x=>x.id===p.id&&x.date===d);return x?timeFmt(x.checkin):'';}),late.filter(x=>x.id===p.id).length,late.filter(x=>x.id===p.id&&x.excluded).length]);
 const calendar=addSheet(wb,'지각 월간표',['점포','사번','이름',...Array.from({length:new Date(+month.slice(0,4),+month.slice(5),0).getDate()},(_,i)=>(i+1)+'일'),'발생','면제'],calendarRows,[20,14,13,...Array(31).fill(9)],{month,note:'날짜별 실제 출근시간 표시 · 면제된 지각도 포함 · 자세한 사유는 소명 및 판정 확인'});
 const cw=wb.Sheets['지각 월간표'];for(let i=0;i<summary.length;i++)for(const x of late.filter(x=>x.id===summary[i].id)){let address=XLSX.utils.encode_cell({r:i+4,c:Number(x.date.slice(8))+2});if(cw[address])cw[address].s={...styleBody(i,!x.excluded),fill:sFill(x.excluded?'FFF2D5':'FBE7E7'),font:{name:'맑은 고딕',bold:true,sz:11,color:{rgb:x.excluded?'936C21':'B23E45'}}};}
 // Keep operational sheets first; state is retained for the next upload.
 wb.SheetNames=['지각 월간표',...wb.SheetNames.filter(n=>n!=='지각 월간표')];wb.Workbook.Sheets=wb.SheetNames.map(n=>({Hidden:n===STATE_SHEET?1:0}));
 wb.Props={...(wb.Props||{}),Comments:'소명 수정 시 전체 지각횟수 및 감점 재계산 (1.0.2)'};
 return wb;
}
function fromPrevious(workbook){let sheet=workbook.Sheets[STATE_SHEET];if(!sheet)throw Error('이전 결과에 누적 저장 시트가 없습니다. 이 프로그램에서 생성한 보고서를 선택해주세요.');let raw=XLSX.utils.sheet_to_json(sheet,{header:1,defval:''});if(raw[0]?.[0]!=='ZENIEL_STATE')throw Error('지원하지 않는 이전 결과 형식입니다.');let json=raw.slice(1).map(r=>String(r[1]||'')).join('');let st=JSON.parse(json);if(!st.months||!st.holidays||!st.hires)throw Error('이전 결과의 누적 데이터가 손상되었습니다.');let review=workbook.Sheets['소명 및 판정'];if(review){let rows=XLSX.utils.sheet_to_json(review,{header:1,defval:'',raw:true});if(trim(rows[3]?.[7])==='확인자'){st.overrides=st.overrides||{};st.notes=st.notes||{};for(let row of rows.slice(4)){let key=trim(row[6]);if(!key||!/^((지각)|(출근 미등록)|(퇴근 미등록))\|/.test(key))continue;st.notes[key]=trim(row[12]);let decision=trim(row[7]),cls=statusFromOverride(decision,true);if(cls==='제외')st.overrides[key]='제외';else if(cls==='무단결근 확정')st.overrides[key]='무단결근 확정';else delete st.overrides[key];}}}return st;}
function mergeHistory(prev,uiEvents,uiHires){let x=clone(prev),map=new Map((x.holidays||[]).map(e=>[keyEvent(e),e]));for(let e of uiEvents)map.set(keyEvent(e),e);x.holidays=[...map.values()];x.hires={...(x.hires||{}),...uiHires};x.overrides={...(x.overrides||{})};return x;}
function reportBytes(wb){
 const archive=fflate.unzipSync(new Uint8Array(XLSX.write(wb,{type:'array',bookType:'xlsx',compression:true})));
 wb.SheetNames.forEach((name,index)=>{if(name===STATE_SHEET)return;const path=`xl/worksheets/sheet${index+1}.xml`;if(!archive[path])return;let xml=fflate.strFromU8(archive[path]);const views='<sheetViews><sheetView showGridLines="0" workbookViewId="0"><pane xSplit="3" ySplit="4" topLeftCell="D5" activePane="bottomRight" state="frozen"/></sheetView></sheetViews>';xml=/<sheetViews[\s\S]*?<\/sheetViews>/.test(xml)?xml.replace(/<sheetViews[\s\S]*?<\/sheetViews>/,views):xml.replace(/<sheetFormatPr/,views+'<sheetFormatPr');archive[path]=fflate.strToU8(xml);});
 return fflate.zipSync(archive,{level:6});
}
function exportWorkbook(){if(!result||!lastReport)throw Error('먼저 분석을 실행해주세요.');let wb=buildWorkbook(result,lastReport);let name=`ZENIEL_코스트코_근태휴무_${result.month.replace('-','')}_V${VERSION}.xlsx`;const blob=new Blob([reportBytes(wb)],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});window.ZenielWorkspace.saveBlob(blob,name);}

function renderEvents(){let tbody=$('eventRows');if(!tbody)return;tbody.innerHTML=holidays.length?holidays.map((e,i)=>`<tr><td>${escapeHTML(e.title)}</td><td>${e.date}</td><td>${e.start} ~ ${e.end}</td><td>${e.amount}일</td><td><button class="mini-delete" type="button" data-delete="${i}">삭제</button></td></tr>`).join(''):'<tr><td colspan="5" class="empty-row">등록된 대체휴일 발생 기준이 없습니다.</td></tr>';tbody.querySelectorAll('[data-delete]').forEach(b=>b.onclick=()=>{holidays.splice(Number(b.dataset.delete),1);renderEvents();});}
function renderHires(){let el=$('hireTags');el.innerHTML=Object.keys(hires).length?Object.entries(hires).map(([id,d])=>`<span>${escapeHTML(id)} · ${escapeHTML(d)} <button data-hire="${escapeHTML(id)}" title="입사일 제거">×</button></span>`).join(''):'<span>등록된 입사일 없음</span>';el.querySelectorAll('[data-hire]').forEach(b=>b.onclick=()=>{delete hires[b.dataset.hire];renderHires();});}
function fillPeople(plan){const s=$('personSelect');let first='<option value="">직원 선택</option>';let unique=new Map(plan.people.map(p=>[p.id,p]));s.innerHTML=first+[...unique.values()].sort((a,b)=>a.name.localeCompare(b.name,'ko')).map(p=>`<option value="${escapeHTML(p.id)}">${escapeHTML(p.name)} · ${escapeHTML(p.store)} (${escapeHTML(p.id)})</option>`).join('');}
async function readFile(file){if(!file)throw Error('업로드할 엑셀 파일을 선택해주세요.');if(typeof XLSX==='undefined')throw Error('엑셀 분석 라이브러리를 불러오지 못했습니다. 네트워크 연결을 확인한 뒤 새로고침 해주세요.');let bytes=await file.arrayBuffer();let wb=XLSX.read(bytes,{type:'array',cellDates:false,cellStyles:true});Object.values(wb.Sheets).forEach(repairRange);return wb;}
async function guessMonthFromAttendance(f){try{let wb=await readFile(f),entries=readAttendance(wb),months={};entries.forEach(r=>months[r.date.slice(0,7)]=(months[r.date.slice(0,7)]||0)+1);let candidate=Object.keys(months).sort((a,b)=>months[b]-months[a])[0];if(candidate&&!$('analysisMonth').dataset.manual){$('analysisMonth').value=candidate;setMessage(`근태관리 파일에서 ${candidate} 분석월을 자동 인식했습니다.`);await loadPlanPeople();}}catch(e){console.warn('분석월 자동 감지 실패',e);}}
async function loadPlanPeople(){let f=$('filePlan').files[0];if(!f)return;try{let book=await readFile(f),month=$('analysisMonth').value||`${new Date().getFullYear()}-${String(new Date().getMonth()+1).padStart(2,'0')}`;fillPeople(readPlan(book,month));setMessage('매장근무계획에서 직원 목록을 읽었습니다.','');}catch(e){setMessage('근무계획 확인 실패: '+e.message,'error');}}
async function loadPrevious(){let f=$('filePrevious').files[0];if(!f)return;try{let wb=await readFile(f),s=fromPrevious(wb);const incoming=new Set(Object.keys(s.months));const kept=Object.fromEntries(Object.entries(history.overrides||{}).filter(([key])=>!incoming.has(key.split('|').at(-1).slice(0,7))));s={...s,months:{...history.months,...s.months},hires:{...history.hires,...s.hires},overrides:{...kept,...s.overrides},notes:{...history.notes,...s.notes},holidays:[...new Map([...(history.holidays||[]),...(s.holidays||[])].map(e=>[keyEvent(e),e])).values()]};history=s;holidays=clone(s.holidays||[]);hires=clone(s.hires||{});renderEvents();renderHires();lastReport=null;result=null;$('downloadBtn').disabled=true;window.dispatchEvent(new CustomEvent('attendance-updated',{detail:{state:history,result:null}}));setMessage(`이전 결과 불러오기 완료 · ${Object.keys(s.months).length}개월 누적 · 대체휴일 규칙 ${holidays.length}개`,'');}catch(e){setMessage('이전 결과 읽기 오류: '+e.message,'error');}}
function setMessage(text,type=''){let el=$('status');if(!el)return;el.textContent=text;el.className='status'+(type?' '+type:'');}
function makeCells(keys,rows,limit=50){let h='<thead><tr>'+keys.map(x=>`<th>${escapeHTML(x)}</th>`).join('')+'</tr></thead>';let b=rows.slice(0,limit).map(r=>'<tr>'+r.map(c=>`<td>${escapeHTML(c)}</td>`).join('')+'</tr>').join('');if(!b)b=`<tr><td colspan="${keys.length}" class="empty-row">현재 해당 내역이 없습니다.</td></tr>`;return h+'<tbody>'+b+'</tbody>';}
function renderPreview(){if(!result)return;let data,headers;switch(activeView){case 'week':headers=['주차','이름','점포','정기휴무','대체휴일','합계','판정'];data=result.weekly.filter(r=>r.status!=='정상').map(r=>[`${r.start}~${r.end}`,r.name,r.store,r.rest,r.sub,r.total,r.status]);break;case 'holiday':headers=['발생일','이름','당일 계획','다른 날 대체휴일','판정'];data=result.holiday.map(r=>[r.eventDate,r.name,r.plan,r.useDates||'-',r.status]);break;case 'issues':headers=['구분','이름','날짜','확인 내용'];data=result.issues.map(r=>[r.kind,r.name,r.date,r.info]);break;default:headers=['날짜','점포','이름','조','기준시간','출근','지각','감점'];data=result.scored.late.map(r=>[r.date,r.store,r.name,r.shift,timeFmt(r.scheduled),timeFmt(r.checkin),`${r.delta}분`,r.excluded?0:r.score]);} $('previewTable').innerHTML=makeCells(headers,data);}
function updateStats(){if(!result)return;$('statPeople').textContent=result.summaries.length+'명';$('statLate').textContent=result.scored.late.length+'건';$('statMissing').textContent=(result.scored.missingIn.length+result.scored.missingOut.length)+'건';$('statHoliday').textContent=result.holiday.filter(r=>/중복|초과|대상 아님·사용/.test(r.status)).length+'건';}
function renderReview(){if(!result){$('reviewList').innerHTML='<p>분석 후 검토할 근태 내역이 나타납니다.</p>';return;}let host=$('reviewList');if(!host)return;let records=[...result.scored.late.map(r=>({type:'지각',...r})),...result.scored.missingIn.map(r=>({type:'출근 미등록',...r})),...result.scored.missingOut.map(r=>({type:'퇴근 미등록',...r}))].sort((a,b)=>a.date.localeCompare(b.date)||a.name.localeCompare(b.name,'ko'));host.innerHTML=records.length?records.map((r,i)=>`<div class="review-row"><div><strong>${escapeHTML(r.name)}</strong><small>${escapeHTML(r.date)} · ${escapeHTML(r.type)} · ${escapeHTML(r.reason||r.plan)}</small></div><select data-key="${escapeHTML(r.key)}"><option value="">자동판정 유지</option><option value="인정" ${resultState(r.key)==='인정'?'selected':''}>증빙 인정 · 감점 제외</option>${r.type==='출근 미등록'?`<option value="무단결근 확정" ${resultState(r.key)==='무단결근 확정'?'selected':''}>무단결근 확정 · -20점</option>`:''}</select><input class="review-note" aria-label="소명 사유" placeholder="사유 입력 (병원 방문, 차량 고장 등)" data-note="${escapeHTML(r.key)}" value="${escapeHTML(lastReport.notes?.[r.key]||'')}"></div>`).join(''):'<p>소명 처리할 근태 이상 내역이 없습니다.</p>';host.querySelectorAll('select').forEach(s=>s.addEventListener('change',()=>{lastReport.overrides[s.dataset.key]=s.value;refreshScored();}));host.querySelectorAll('[data-note]').forEach(input=>input.addEventListener('change',()=>{lastReport.notes=lastReport.notes||{};lastReport.notes[input.dataset.note]=input.value;notifyUpdate();}));}
function resultState(key){let s=lastReport?.overrides?.[key]||'';return statusFromOverride(s,true)==='제외'?'인정':s;}
function refreshScored(){if(!result||!lastReport)return;let raw=result._raw;if(!raw)return;result.scored=applyScores(raw,lastReport.overrides||{});const perId=result.scored.perId;result.summaries=result.summaries.map(person=>{let score=perId.get(person.id)||{late:0,missingIn:0,missingOut:0,penalty:0};return{...person,...score,warning:score.penalty<=-5?'경고장 검토':'-'};});updateStats();renderPreview();renderReview();notifyUpdate();}
async function run(){let btn=$('runBtn');try{if(btn.disabled)return;btn.disabled=true;setMessage('원본 엑셀을 읽고 날짜별·직원별로 대조하고 있습니다.','busy');let month=$('analysisMonth').value;if(!/^\d{4}-\d{2}$/.test(month))throw Error('분석 대상월을 선택해주세요.');for(let id of FIELDS)if(!$(id).files?.length)throw Error('월별 원본 파일 4개를 모두 선택해주세요.');let [p,a,l,h]=await Promise.all(FIELDS.map(id=>readFile($(id).files[0])));let plan=readPlan(p,month),attendance=readAttendance(a),leave=readLeave(l),hours=readHours(h);if(!plan.people.length)throw Error('선택한 근무계획에 직원이 없습니다.');let srcAtt=attendance.filter(r=>dateInMonth(r.date,month)),srcLeave=leave.filter(r=>!(r.end.slice(0,7)<month||r.start.slice(0,7)>month));if(!srcAtt.length)throw Error('선택한 월의 근태 기록이 없습니다. 분석월과 원본을 확인해주세요.');let lastAtt=srcAtt.map(r=>r.date).sort().at(-1)||`${month}-01`;let cutoff=$('cutoffDate')?.value||lastAtt;if(cutoff.slice(0,7)!==month)throw Error('판정 기준일은 분석 대상월 안의 날짜여야 합니다.');let state=mergeHistory(history,holidays,hires);state.months[month]={plan:plan.plan,attendance:srcAtt,leave:srcLeave,hours,cutoff,useModified:$('useModified').checked}; // Replacing a month is idempotent, no duplicate accumulation.
let x=analyze(state,month,hours,cutoff,$('useModified').checked);history=state;lastReport=state;result=x;result._raw=makeFindings(resolveScope(state,month),month,hours,cutoff,$('useModified').checked,state.overrides); // transient, not saved in the history workbook
updateStats();renderPreview();renderReview();$('downloadBtn').disabled=false;setMessage(`분석 완료 · ${month} · 계획 ${plan.people.length}명 · 근태 ${srcAtt.length}건 · 연차신청 ${srcLeave.length}건 · 주차 ${result.weekly.length}건 · 원본을 모두 대조했습니다.`);notifyUpdate();location.hash='report';}
catch(e){console.error(e);setMessage('분석 오류: '+(e.message||String(e)),'error');}finally{btn.disabled=false;}}
function init(){$('analysisMonth').value='';$('analysisMonth').addEventListener('change',()=>{$('analysisMonth').dataset.manual='1';if($('filePlan').files.length)loadPlanPeople();});const fileLabels={filePlan:'namePlan',fileAttendance:'nameAttendance',fileLeave:'nameLeave',fileHours:'nameHours'};
Object.entries(fileLabels).forEach(([k,target])=>$(k).addEventListener('change',()=>{let f=$(k).files[0];$(target).textContent=f?f.name:'파일 선택';$(k).closest('.upload-tile').classList.toggle('selected',!!f);if(k==='filePlan'&&f)loadPlanPeople();if(k==='fileAttendance'&&f)guessMonthFromAttendance(f);}));$('filePrevious').addEventListener('change',loadPrevious);
$('addEvent').onclick=()=>{let e={title:trim($('eventTitle').value)||'대체휴일',date:$('eventDate').value,start:$('eventStart').value,end:$('eventEnd').value,amount:Number($('eventAmount').value)};if(!validISO(e.date)||!validISO(e.start)||!validISO(e.end)||e.start>e.end){setMessage('대체휴일 발생일과 사용 시작·종료일을 올바르게 지정해주세요.','error');return;}if(holidays.some(x=>keyEvent(x)===keyEvent(e))){setMessage('이미 등록한 대체휴일 규칙입니다.','error');return;}holidays.push(e);holidays.sort((a,b)=>a.date.localeCompare(b.date));renderEvents();setMessage('대체휴일 발생 규칙이 추가되었습니다. 분석을 실행하면 결과에 반영됩니다.');};
$('saveHire').onclick=()=>{let id=$('personSelect').value,date=$('personHire').value;if(!id||!validISO(date)){setMessage('직원을 선택하고 유효한 입사일을 지정해주세요.','error');return;}hires[id]=date;renderHires();setMessage('입사일이 등록되었습니다. 다음 달에 재사용하려면 분석 결과 엑셀을 보관하세요.');};$('personSelect').onchange=()=>{$('personHire').value=hires[$('personSelect').value]||'';};
$('runBtn').onclick=run;$('downloadBtn').onclick=()=>{try{exportWorkbook();}catch(e){setMessage(e.message,'error');}};
$('resetBtn').onclick=()=>{if(!confirm('현재 파일 선택과 분석 화면을 초기화할까요? 저장된 누계와 설정은 유지됩니다.'))return;for(let id of [...FIELDS,'filePrevious'])$(id).value='';for(let [id,target]of Object.entries(fileLabels)){$(target).textContent='파일 선택';$(id).closest('.upload-tile').classList.remove('selected');}result=null;lastReport=null;$('personSelect').innerHTML='<option value="">근무계획을 업로드하면 직원 목록이 나타납니다.</option>';renderEvents();renderHires();$('downloadBtn').disabled=true;['statPeople','statLate','statMissing','statHoliday'].forEach(id=>$(id).textContent='—');$('previewTable').innerHTML='<thead><tr><th>분석 전</th></tr></thead><tbody><tr><td class="empty-row">파일을 올리고 분석을 실행하면 미리보기가 표시됩니다.</td></tr></tbody>';renderReview();setMessage('초기화 완료.');};
document.querySelectorAll('.tab').forEach(b=>b.addEventListener('click',()=>{activeView=b.dataset.view;document.querySelectorAll('.tab').forEach(x=>x.classList.toggle('active',x===b));renderPreview();}));renderEvents();renderHires();}
function notifyUpdate(){if(lastReport){history=lastReport;window.dispatchEvent(new CustomEvent('attendance-updated',{detail:{state:lastReport,result}}));}}
if(typeof window!=='undefined')window.CostcoAnalyzer={openMonth:month=>{let d=history.months[month];if(!d?.hours)throw Error('이전 형식 자료는 원본 4개로 다시 분석해주세요.');lastReport=history;result=analyze(history,month,d.hours,d.cutoff,!!d.useModified);result._raw=makeFindings(resolveScope(history,month),month,d.hours,d.cutoff,!!d.useModified,history.overrides||{});updateStats();renderPreview();renderReview();$('downloadBtn').disabled=false;setMessage(month+' 저장된 분석 결과를 열었습니다.');notifyUpdate();location.hash='report';},getState:()=>lastReport||history,getResult:()=>result,restoreState:s=>{history=s;holidays=clone(s.holidays||[]);hires=clone(s.hires||{});lastReport=null;result=null;renderEvents();renderHires();},readFile,resolveScope,reportBytes,buildWorkbook,VERSION,readPlan,readAttendance,readLeave,readHours,analyze,emptyState,isoDate,vacationType,asMinutes,makeFindings,applyScores,buildWorkbook,fromPrevious};
if(typeof document!=='undefined')document.addEventListener('DOMContentLoaded',init);
})();
