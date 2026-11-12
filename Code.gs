/** ABM Sync — วางในสคริปต์ที่เปิดจาก Google Sheet (Extensions > Apps Script) */
const TOKEN='CHANGE_ME'; // ตั้งรหัสผ่านของคุณ แล้วใส่รหัสเดียวกันในแอป
const CH=40000;
/* ไม่ต้องกด Run ในตัวแก้ไข — ใช้ผ่านการ Deploy เป็น Web app เท่านั้น */
function doGet(){return out({ok:true,msg:'ABM Sync พร้อมใช้งาน'})}
let T0=Date.now();
function out(o){o=o||{};o.v=2;o.ms=Date.now()-T0;return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON)}
function doPost(e){T0=Date.now();const lock=LockService.getScriptLock();
  try{lock.waitLock(30000);const q=JSON.parse(e.postData.contents);
    if(TOKEN==='CHANGE_ME'||q.token!==TOKEN)return out({ok:false,error:'รหัสผ่านไม่ถูกต้อง'});
    const ss=SpreadsheetApp.getActiveSpreadsheet(),m=meta(ss);
    if(q.action==='status')return out({ok:true,rev:m.rev,updatedAt:m.updatedAt});
    if(q.action==='pull')return out({ok:true,rev:m.rev,updatedAt:m.updatedAt,data:readAll(ss)});
    if(q.action==='push'){if(!q.force&&Number(q.baseRev)!==m.rev)return out({ok:false,conflict:true,rev:m.rev});
      writeAll(ss,q.data);const rev=m.rev+1;saveMeta(ss,rev,q.device);return out({ok:true,rev})}
    return out({ok:false,error:'unknown action'})
  }catch(err){return out({ok:false,error:String(err)})}finally{try{lock.releaseLock()}catch(x){}}}
function sheet(ss,n){return ss.getSheetByName(n)||ss.insertSheet(n)}
function meta(ss){const s=ss.getSheetByName('_meta');if(!s)return{rev:0,updatedAt:''};const o={};s.getDataRange().getValues().forEach(r=>o[r[0]]=r[1]);return{rev:Number(o.rev)||0,updatedAt:String(o.updatedAt||'')}}
function saveMeta(ss,rev,dev){const s=sheet(ss,'_meta');s.clear();s.getRange(1,1,3,2).setNumberFormat('@').setValues([['rev',rev],['updatedAt',new Date().toISOString()],['device',dev||'']])}
function chunks(s){const a=[];for(let i=0;i<s.length;i+=CH)a.push(s.slice(i,i+CH));return a.length?a:['']}
function flat(x){const o={};if(x&&typeof x==='object')Object.keys(x).forEach(k=>{const v=x[k];if(v!==null&&typeof v!=='object'&&String(v).length<200)o[k]=v});return o}
function writeTable(ss,name,items){const keys=[];items.forEach(it=>Object.keys(it.cols).forEach(k=>{if(keys.indexOf(k)<0)keys.push(k)}));
  const n=Math.max(1,...items.map(it=>it.js.length)),js=Array.from({length:n},(_,i)=>i?'_json'+(i+1):'_json');
  const all=[keys.concat(js)].concat(items.map(it=>keys.map(k=>it.cols[k]===undefined?'':it.cols[k]).concat(js.map((_,i)=>it.js[i]||''))));
  const sh=sheet(ss,name);sh.clear();sh.getRange(1,1,all.length,all[0].length).setNumberFormat('@').setValues(all);sh.setFrozenRows(1)}
function writeAll(ss,d){const kv={},old=ss.getSheetByName('_kv');
  if(old){const v=old.getDataRange().getValues(),h=v[0]||[],jc=[];h.forEach((x,i)=>{if(String(x).indexOf('_json')===0)jc.push(i)});const kc=h.indexOf('key');
    v.slice(1).forEach(r=>{const s=jc.map(c=>r[c]).join('');if(s)kv[r[kc]]=s})}
  Object.keys(d).forEach(k=>{const v=d[k];
    if(Array.isArray(v))writeTable(ss,k,v.map(x=>({cols:flat(x),js:chunks(JSON.stringify(x))})));
    else kv[k]=JSON.stringify(v)});
  writeTable(ss,'_kv',Object.keys(kv).map(k=>({cols:{key:k},js:chunks(kv[k])})))
  if(d.settings&&typeof d.settings==='object')view(ss,'settings_view',d.settings)}
/* แผ่นอ่านง่าย (อ่านอย่างเดียว ไม่ถูกดึงกลับเข้าแอป) */
function view(ss,name,o){const rows=[['field','value']].concat(Object.keys(o).map(f=>[f,String(o[f]).slice(0,200)])),sh=sheet(ss,name);sh.clear();sh.getRange(1,1,rows.length,2).setNumberFormat('@').setValues(rows)} // อัปเดตเฉพาะ key ที่ส่งมา ที่เหลือคงเดิม
function readAll(ss){const d={};
  ss.getSheets().forEach(sh=>{const n=sh.getName();if(n==='_meta')return;const v=sh.getDataRange().getValues();if(!v.length)return;
    const jc=[];v[0].forEach((h,i)=>{if(String(h).indexOf('_json')===0)jc.push(i)});if(!jc.length)return;
    const kc=v[0].indexOf('key'),arr=[];
    v.slice(1).forEach(r=>{const s=jc.map(c=>r[c]).join('');if(!s)return;if(n==='_kv')d[r[kc]]=JSON.parse(s);else arr.push(JSON.parse(s))});
    if(n!=='_kv')d[n]=arr});
  return d}
