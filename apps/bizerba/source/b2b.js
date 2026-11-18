/* ---------------------------------------------------------------- BRAIN2 state */
const B2_CATS={
 "…":["Audit Trail","About BRAIN2"],
 "Product data":["Code Structure","Label Designer","Layout Management","Master Data","Master Data Exceptions","Text"],
 "Process data":["Devices","Weight Classes","Automation Parameters","Label Parameters","Data Transfer"],
 "Result data":["Individual Values","Totals","Give-away","Rejects"],
 "Reporting":["Reports"],
 "Jobs":["Job List"],
 "System data":["Users & Groups","Device Access","Database Backup","Licences"]};
const B2DB=()=>{const d=freshDB();const keep=["plu","text","dtext","logo","code","label","auto","wc","layouts","stattxt"];const o={};keep.forEach(k=>o[k]=d[k]);return o};
const B2={view:"term",cat:"Product data",mod:"Label Designer",navOpen:true,db:null,devices:[],audit:[],transfers:[],results:[],exc:[],status:{},ev:store.get("b2ev",{}),
 ld:{proj:"2SFGScunthorpeFreshPack2 V2",open:["77"],cur:"77",sel:null,zoom:4,plu:7777,menu:null,dock:"proj",dirty:{},undo:[],redo:[],unit:"Dot"},
 md:{q:"",cust:"",sel:6419,tab:"General",dirty:false,work:null},tx:{kind:"text",sel:22,work:null},cs:{sel:1,work:null},wc:{sel:7292,work:null,cell:null},
 dev:{sel:null},rd:{line:"all",rep:"Give-away report"},au:{q:""},lm:{sel:"205"},ex:{dev:"",plu:"",field:"lb1",val:""},modal:null};
B2.db=B2DB();
Object.keys(B2.db.plu).forEach(n=>B2.status[n]="Released");
Object.entries(LINES).forEach(([k,L],li)=>L.b2.forEach((name,i)=>B2.devices.push({id:name,line:k,i,bus:(L.dev[i]||L.dev[0]).bus,model:L.model,ip:"10.43.12."+(20+li*4+i),online:L.online!==false,legacy:L.online===false,last:"24/10/2026 02:00",lastLayouts:"02/10/2026 14:52"})));
/* the slave labeller's own values, held as master data exceptions (FP2-L04 photo: LB1 40 mm, label distance 0 mm, max height 132 mm, conveying 60 m/min) */
["FP2-L03-Glmi-S","FP2-L04-Glmi-S"].forEach(d=>[["lb1",40,"LB1 susp. path"],["dist",0,"Label distance"],["pkh",132,"Max pckg. height"],["conv",60,"Conveying speed"]].forEach(([f,v,l])=>B2.exc.push({dev:d,plu:6419,field:f,label:l,val:v,note:"Slave labeller (photo FP2-L04)",by:"brain2admin",at:"02/10/2026"})));
const B2USER="Admin",B2WS="STBIZERBA.2SFG.NET:2015";
function b2Audit(mod,obj,field,before,after){B2.audit.unshift({at:new Date(),user:B2USER,ws:B2WS,mod,obj,field,before:String(before??""),after:String(after??"")});if(B2.audit.length>400)B2.audit.length=400}
/* seed some history so Audit Trail and Result data open with content (marked as trainer examples) */
(function seed(){const ago=(d,h)=>new Date(Date.now()-d*864e5-h*36e5);
 [[3,4,"t.services","Master Data","PLU 6419","Shelf-life 1","8","9"],[3,4,"t.services","Master Data","PLU 6419","Text field 10","","=$N$ IF [[GL0A < 7 OR GL0A > 10]]…"],[2,6,"brain2admin","Label Designer","Layout 205","Logo field 01","GB-1606.png","GB-2037.png"],[1,2,"v.compliance","Text","General text 22","Text","…1 hr 20 mins.","…1 hr 25 mins."],[1,1,"t.services","Process data","FP2-L04-Glmi-M","Download","","PLU, texts, weight classes"],[0,9,"t.services","Weight Classes","Table 7292","Upper limit","2900","2874"]].forEach(a=>B2.audit.push({at:ago(a[0],a[1]),user:a[2],ws:"2SFG-TECH-PC02",mod:a[3],obj:a[4],field:a[5],before:a[6],after:a[7],seed:true}));
 B2.audit.sort((a,b)=>b.at-a.at);
 const gen=(line,plu,n,mean,sd,h)=>{for(let i=0;i<n;i++){const w=Math.round((mean+gauss()*sd)*10000)/10000;const p=B2.db.plu[plu];if(!p)return;const lo=p.minw/1000,hi=p.maxw/1000;B2.results.push({line,plu,w,cls:w<lo?"Underweight":w>hi?"Overweight":"OK",ok:w>=lo&&w<=hi,ej:w<lo||w>hi,flags:(w<lo||w>hi)?8:0,printed:1,t:Date.now()-h*36e5+i*4000,nom:p.mode==="Fixed weight"?p.fixw:p.minw,mode:p.mode,seed:true})}};
 gen("FP2-L04",6419,220,2.62,.09,7);gen("FP2-L03",6706,180,2.28,.07,6);if(B2.db.plu[7037])gen("FP2-L03",7037,240,.6115,.006,3);
 B2.transfers.push({at:ago(2,22),user:"BRAIN2 scheduler",dev:"all online devices",what:"Night synchronization: PLU, texts, weight classes",res:"OK (6 devices)"},{at:ago(1,1),user:"t.services",dev:"FP2-L04-Glmi-M",what:"Download PLU, texts, weight classes",res:"OK"},{at:ago(24,1),user:"brain2admin",dev:"FP1-L03-Glmi-M",what:"Export layouts 173.173 CO-OP",res:"OK"},{at:ago(24,1),user:"brain2admin",dev:"FP1-L03-Glmi-M",what:"Download PLU data",res:"Refused: device platform is view only"})})();
function b2Result(r){B2.results.push(r);if(B2.results.length>4000)B2.results.splice(0,500)}
function b2Ev(k){B2.ev[k]=true;store.set("b2ev",B2.ev);if(S)checkDrills()}
const devLine=id=>B2.devices.find(d=>d.id===id);
/* compare BRAIN2's database with what is on a line */
function lineDiff(k){const L=DBS[k]||null;if(!L)return{n:0,list:[],unknown:true};const out=[];
 Object.entries(B2.db.text).forEach(([n,t])=>{if(L.text[n]!==t)out.push("Text "+n)});
 Object.keys(B2.db.plu).forEach(n=>{const a=B2.db.plu[n],b=L.plu[n];if(!b){out.push("PLU "+n+" missing");return}for(const f of["text","tare","minw","maxw","life1","txt1","lab","wc","tf10","cs1"])if(String(a[f])!==String(b[f])){out.push(`PLU ${n} · ${f}`);break}});
 Object.keys(B2.db.wc).forEach(n=>{if(JSON.stringify(B2.db.wc[n])!==JSON.stringify(L.wc[n]))out.push("Weight class "+n)});
 Object.keys(L.plu).forEach(n=>{if(!B2.db.plu[n])out.push("PLU "+n+" only on device")});
 return{n:out.length,list:out}}
function withLine(k,fn){const s0=S,d0=DB;S=ST[k];DB=DBS[k];try{return fn()}finally{S=s0;DB=d0}}
/* download master data to one line (both its devices); exceptions for the -S device go into each article's slave values */
function b2Download(k,what){const L=LINES[k];if(L.online===false)return{ok:false,msg:"The device platform does not accept PLU, text and weight class data from BRAIN2 (view only). Change these at the terminal."};
 const t=DBS[k]||(DBS[k]=freshDB());const src=clone(B2.db);
 if(what.plu){const prev=t.plu;t.plu=src.plu;Object.values(t.plu).forEach(p=>{p.slave=clone(p.slave||{});const o=prev[p.plu];if(!o)return;
   /* 0 = keep previously loaded data (as in Data Maintenance) */
   ["aut","lab","wc","logo1","logo2","logo3","cf1","cf2","cf3"].forEach(f=>{if(Number(p[f])===0)p[f]=o[f]});[1,2,3,4,5].forEach(i=>{if(String(p["cs"+i])==="0")p["cs"+i]=o["cs"+i]});if(Number(p.unit)===0&&!(p.sprice>0)){p.mode=o.mode;p.unit=o.unit;p.fixp=o.fixp}});
  B2.exc.filter(x=>devLine(x.dev)&&devLine(x.dev).line===k).forEach(x=>{const tg=x.plu?[t.plu[x.plu]]:Object.values(t.plu);const isS=devLine(x.dev).i===1;tg.forEach(p=>{if(!p)return;if(isS){p.slave=p.slave||{};p.slave[x.field]=x.val}else p[x.field]=x.val})});
  t.label=src.label;t.auto=src.auto;t.code=src.code;t.logo=src.logo}
 if(what.text){t.text=src.text;t.dtext=src.dtext}
 if(what.wc)t.wc=src.wc;
 if(ST[k])withLine(k,()=>{S.wcJob=null;S.jobText={};S.jobOv=null;loadJob();log("BRAIN2 download: "+Object.keys(what).filter(x=>what[x]).join(", ")+" (local changes overwritten)");S.ev.brain=true});
 return{ok:true}}
function b2Upload(k){const t=DBS[k];if(!t)return{ok:false,msg:"No data read yet from this device"};const src=clone(t);
 ["plu","text","wc"].forEach(x=>{B2.db[x]=src[x]});Object.keys(B2.db.plu).forEach(n=>{if(!B2.status[n])B2.status[n]="Released"});return{ok:true}}
function b2ExportLayouts(ids,devs){const res=[];devs.forEach(id=>{const d=devLine(id);const t=DBS[d.line]||(DBS[d.line]=freshDB());t.layouts=t.layouts||{};ids.forEach(n=>{t.layouts[n]=clone(B2.db.layouts[n]);const lb=Object.values(t.label).find(l=>l.layout===String(n));if(!lb)t.label[n]={name:B2.db.layouts[n].name,size:B2.db.layouts[n].w+" x "+B2.db.layouts[n].h,type:"Standard",layout:String(n),speed:150,vpos:300,hpos:40,feed:16,rev:"W/o",sales:"with",lbf:2,lbb:2}});t.layoutsFromB2=true;d.lastLayouts=fmtD(new Date())+" "+new Date().toTimeString().slice(0,5);res.push(id)});return res}

/* ---------------------------------------------------------------- shell */
function b2Shell(){if(B2.app==="dm")return dmShell();const mods=B2_CATS[B2.cat];
 return`<div class="b2win"><div class="b2title"><span class="b2tl">${brainIcon()}</span><span>BRAIN2</span><span class="b2wc"><i>?</i><i>─</i><i>▢</i><i>✕</i></span></div>
 <div class="b2main ${B2.navOpen?"":"navc"}"><aside class="b2nav"><div class="b2navh"><span class="b2burger" data-b="burger" title="Applications">▭ ▾</span><span class="b2mt">${esc(B2.mod)}</span><span class="b2col" data-b="navtog">${B2.navOpen?"‹":"›"}</span></div>
 ${B2.navOpen?`<nav class="b2mods">${mods.map(m=>`<button class="b2m ${m===B2.mod?"on":""}" data-b="mod" data-v="${esc(m)}">${esc(m)}</button>`).join("")}</nav>`:""}</aside>
 <section class="b2c">${b2Module()}</section></div>
 <footer class="b2cats">${Object.keys(B2_CATS).map(c=>`<button class="b2cat ${c===B2.cat?"on":""}" data-b="cat" data-v="${esc(c)}">${esc(c)}</button>`).join("")}<span class="b2user"><span>${esc(B2USER)} (Two Sisters Administrator)</span><span>${esc(B2WS)}</span></span><span class="b2conn" title="Connected to the BRAIN2 server">${connIcon()}</span></footer>
 ${B2.modal?b2Modal():""}</div>`}
function connIcon(){return`<svg width="30" height="26" viewBox="0 0 30 26"><rect x="2" y="9" width="16" height="8" rx="1" fill="#555"/><rect x="17" y="11" width="6" height="4" fill="#555"/><path d="M6 9V5M14 9V5" stroke="#555" stroke-width="2"/><circle cx="22" cy="18" r="7" fill="#3a9b3a"/><path d="M18.5 18l2.5 2.5 4.5-5" stroke="#fff" stroke-width="2" fill="none"/></svg>`}
function brainIcon(){return`<svg width="26" height="22" viewBox="0 0 26 22"><path d="M8 3c-3 0-5 2-5 4-2 1-2 4 0 5-1 3 1 5 4 5 1 2 4 2 5 0V4c-1-1-2-1-4-1zM18 3c3 0 5 2 5 4 2 1 2 4 0 5 1 3-1 5-4 5-1 2-4 2-5 0V4c1-1 2-1 4-1z" fill="#fff" stroke="#333" stroke-width="1.3"/><path d="M6 8h4M6 13h4M16 8h4M16 13h4" stroke="#333"/></svg>`}
function b2Module(){const m=B2.mod;const F={"Label Designer":ldHTML,"Layout Management":lmHTML,"Code Structure":csHTML,"Master Data":mdHTML,"Master Data Exceptions":exHTML,"Text":txHTML,
 "Devices":devHTML,"Weight Classes":wcHTML,"Automation Parameters":()=>parHTML("auto"),"Label Parameters":()=>parHTML("label"),"Data Transfer":dtHTML,"Device Access":daHTML,
 "Individual Values":ivHTML,"Job List":jobHTML,"Totals":totHTML,"Give-away":gaHTML,"Rejects":rjHTML,"Reports":repHTML,"Audit Trail":auHTML,"Users & Groups":daHTML,"Database Backup":bkHTML,"Licences":licHTML,"About BRAIN2":aboutHTML};
 return(F[m]||(()=>`<div class="b2pad">Not modelled.</div>`))()}
const b2h=(t,sub,btns)=>`<div class="b2hd"><div><h3>${esc(t)}</h3>${sub?`<p>${sub}</p>`:""}</div><div class="b2tb">${btns||""}</div></div>`;
const bb=(act,label,o)=>`<button class="b2btn ${o&&o.pri?"pri":""}" data-b="${act}" ${o&&o.v!=null?`data-v="${esc(o.v)}"`:""} ${o&&o.dis?"disabled":""}>${label}</button>`;
