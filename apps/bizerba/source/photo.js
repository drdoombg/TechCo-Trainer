/* ============================================================ photographed main menu
   The FP2 terminals show a grouped main menu (photos 24/10/2026): Article data, Production data,
   Display, Configuration, Device, Service. Branches that were photographed are built key for key;
   the others are filled with the manual's menus (marked as assumed in the manual panel). */
let MENU_MODE=store.get("menuMode","photo");
const FREE={free:true};
const MK=(label,o)=>Object.assign({label},o||{});
const PMENU=MK("Main menu",{kids:[
 MK("Article data",{ico:"db",photo:true,kids:[
  MK("PLU",{sk:"PLU"}),
  MK("Edit current article data",{fn:"ace",sec:"12.6"}),
  MK("Article data maintenance",{fn:"acm",lv:3,sec:"12.6"}),
  MK("Templates",{photo:true,kids:[MK("Template 1",{fn:"tpl",big:"P1"}),MK("Template 2",{fn:"tpl",big:"P2"}),MK("Template 3",{fn:"tpl",big:"P3"}),MK("Template 4",{fn:"tpl"}),MK("Template 5",{fn:"tpl"}),MK("Template 6",{fn:"tpl"}),MK("Template from PLU",{fn:"tplplu",greyAlways:true}),MK("More templates",{fn:"tplmore",menuLook:true}),MK("Create/edit template",{sk:"Create/edit template",lv:3}),MK("Delete unused cells",{sk:"Delete unused cells",lv:3}),MK("Delete template",{sk:"Delete template",lv:3})]}),
  MK("Article dataparts management",{photo:true,kids:[
   MK("PLU",{ref:["PLU"]}),
   MK("Texts",{photo:true,kids:[MK("General texts",{photo:true,kids:[FREE,FREE,MK("Text editor",{fn:"ted",lv:3,sec:"13.3"}),MK("Delete text",{sk:"Delete text",lv:3}),MK("Delete text range",{sk:"Delete text range",lv:3}),MK("Text No. occupied",{sk:"Text No. occupied"}),FREE,FREE]}),MK("Static texts",{ref:["Static texts"]}),MK("Date text",{ref:["Date text"]}),FREE,FREE,FREE,FREE,FREE]}),
   MK("Labels",{refPick:[["Labels","Label data local"],["Labels","Lab. data global only"],["Labels","Lab. data glob.+local"],["Labels","Delete lab. data local."],["Labels","Delete label data"],["Labels","Delete labelLayout"],["Labels","Delete logo"],["Labels","Delete font"]]}),
   MK("Weight class",{photo:true,kids:[MK("Load weight class",{fn:"wcload"}),MK("Edit current weight class data",{sk:"Weight class editor"}),MK("Weight class data",{sk:"Weight class data",lv:3}),MK("Delete weight class",{fn:"wcdel",lv:3}),FREE,FREE,FREE,FREE]}),
   MK("Code",{ref:["Code"]}),MK("Automat",{ref:["Automat"]}),MK("Scanning rule",{ref:["Scanning rules"]}),MK("Tendency control",{ref:["Tendency regulation"]})]}),
  FREE,FREE,
  MK("Article data configuration",{refs:[["Article data"],["Prices"],["Labeling mode"],["Date"]]})]}),
 MK("Production data",{ico:"prod",refs:[["Processing"],["Texts"],["Configure total"],["Release total"],["Totalization"],["Numerators"],["Weight class &statistics"],["Templates"]]}),
 MK("Display",{ico:"disp",ref:["Display"]}),
 MK("Configuration",{ico:"conf",refs:[["Labels"],["Printer + interface"],["Interfaces"],["Languages & characters"],["Date"],["Setup"],["Storage medium"],["Device operation"]]}),
 FREE,FREE,
 MK("Device",{ico:"dev",split:true,ref:["Device menu"]}),
 MK("Service",{ico:"svc",refs:[["Service menu"],["Password editor"],["Printer"],["Scale menu"],["Database"]]})]});
function findByLabel(label,n){n=n||TREE;for(const c of n.ch){if(c.label===label)return c;const r=findByLabel(label,c);if(r)return r}return null}
function pmKids(v){
 if(v.kids)return v.kids.map(k=>k.free?{free:true}:k.ref?{label:k.label,m:findNode(k.ref),vk:k}:{label:k.label,v:k});
 if(v.refs)return v.refs.map(r=>{const n=findNode(r);return n?{label:n.label,m:n}:null}).filter(Boolean);
 if(v.refPick)return v.refPick.map(r=>{const n=findNode(r);return n?{label:n.label,m:n}:null}).filter(Boolean);
 if(v.ref){const n=findNode(v.ref);return n?n.ch.map(c=>({label:c.label,m:c})):[]}
 return[]}
function pmResolve(path){let cur={v:PMENU};
 for(const t of path){if(cur.v){if(cur.v.ref){const n=findNode(cur.v.ref);cur={m:n};}
   if(cur.v){const e=pmKids(cur.v).find(x=>!x.free&&x.label===t);if(!e)break;cur=e.m?{m:e.m,vk:e.vk}:{v:e.v};continue}}
  const c=cur.m.ch.find(x=>x.label===t);if(!c)break;cur={m:c}}
 if(cur.v&&cur.v.ref)cur={m:findNode(cur.v.ref),vk:cur.v};
 return cur}
const PFN={
 ace:()=>openACE("current"),
 acm:()=>loadRecord(n=>openACE("maint",n)),
 tpl:()=>msgDlg("Template with this number does not exist!"),
 tplplu:()=>{},
 tplmore:()=>num("Template number","",7,()=>msgDlg("Template with this number does not exist!")),
 wcload:()=>loadDlg("Weight class No.",P().wc>0?P().wc:"",v=>{v=Number(v);if(!DB.wc[v]){msgDlg("Weight class with this number does not exist!");return}loading(()=>{S.wcJob=clone(DB.wc[v]);log(`Weight class table ${v} loaded for the current labeling job`);toast(`Weight class ${v} loaded for this job`)})},[]),
 wcdel:()=>num("Delete weight class","",P().wc,v=>{v=Number(v);if(!DB.wc[v]){msgDlg("Weight class with this number does not exist!");return}if(v===P().wc){toast("The active PLU uses this table. Change the PLU first.");return}conf("Security check",`Delete weight class ${v}?`,()=>{delete DB.wc[v];log(`Weight class ${v} deleted`)})}),
 ted:()=>openTED()};
function vKey(e){
 if(e.free)return{label:"Free",free:true,grey:true};
 if(e.m){const vis=nodeVisible(e.m,S.level)||e.m.ch.some(c=>nodeVisible(c,S.level));if(!vis)return{label:"Free",free:true,grey:true};
  const k=keyFor(e.m);k.label=e.label;if(e.m.ch.length)k.menu=true;if(e.vk&&e.vk.ico)k.icon=e.vk.ico;return k}
 const v=e.v;const k={label:v.label,vnode:v};
 if(v.kids||v.refs||v.ref||v.refPick){k.menu=true;k.icon=v.ico;k.desc=v.photo?"Photographed on the FP2 terminal.":"Group on the photographed main menu. The trainer fills it with the manual's menus: "+(v.refs||v.refPick||[v.ref]).map(r=>r.join(" / ")).join(", ")+" (assumed grouping).";
  if(v.split){k.split=["Stop","Start"];k.cur=()=>S.running?"Start":"Stop"}return k}
 if(v.sk){const sp=SPECIAL[v.sk]||{};const node=findByLabel(v.sk);Object.assign(k,node?keyFor(node):{},sp);k.label=v.label;k.menu=false;if(node)k.node=node}
 if(v.fn){k.run=PFN[v.fn];k.fnKey=true;if(v.sec)k.sec=v.sec}
 if(v.big)k.big=v.big;if(v.menuLook)k.menu2=true;
 if(v.greyAlways)k.grey=true;
 if(v.lv&&S.level<v.lv){k.grey=true;k.greyWhy=`Needs authorization level ${v.lv}`}
 return k}
function photoList(){const cur=pmResolve(S.path);
 if(cur.m){const L=cur.m.ch.filter(c=>nodeVisible(c,S.level)).map(keyFor);return L}
 return pmKids(cur.v).map(vKey)}
function crumb(){return S.path.join(" \\ ")}

/* ---------------------------------------------------------------- photo dialogs */
function msgDlg(text,ok){S.dlg={k:"msg",text,ok};render()}
function loading(fn){S.dlg={k:"loading"};render();setTimeout(()=>{S.dlg=null;fn();render()},700)}
/* "Load" / "Load record" dialog: PLU (0..999999), arrows step through occupied numbers */
function loadDlg(field,cur,ok,extra,title){S.dlg={k:"load",title:title||"Load",field,val:String(cur??""),fresh:true,ok,extra:extra||[]};render()}
function pluStep(cur,dir){const ks=Object.keys(DB.plu).map(Number).sort((a,b)=>a-b);if(!ks.length)return cur;const c=Number(cur)||0;
 if(dir>0)return ks.find(x=>x>c)??ks[0];const r=ks.filter(x=>x<c);return r.length?r[r.length-1]:ks[ks.length-1]}
function loadRecord(ok){loadDlg("PLU",S.plu,v=>{v=Number(v);if(!DB.plu[v]){msgDlg(`PLU ${v} does not exist. Use Next availb. no. to create a new one.`);return}ok(v)},[
  {label:"PLU",ico:"⌕",fn:()=>pluList(()=>true,v=>ok(v))},
  {label:"Next availb. no.",fn:()=>{let k=1;while(DB.plu[k])k++;DB.plu[k]=Object.assign(clone(DB.plu[S.plu]),{text:"",plu:k});ok(k)}},
  {label:"Article text",ico:"⌕",fn:()=>pluList(()=>true,v=>ok(v))}],"Load record")}

/* ---------------------------------------------------------------- Edit current article data / Article data maintenance
   One table, one column per device on the line (2: Master, 3: Slave), tabs on the right as photographed. */
const ACE_TABS={
 General:[["text","Article text","text"],["country","Country","sel",["England","Scotland","Wales","Northern Ireland","United Kingdom"]],["upi","Unit price intervals","sel",["per unit","per kg","per 100 g"]],["unit","Price","money"],["sprice","Special price","money"],["recalc","Recalculate unit price(s)","sel",["Without recalculation","With recalculation"]],["wtypes","Weight types","ro",()=>"kg 001"],["mode","Labeling mode","sel",["Weight","Fixed price","Fixed weight","Fixed value"]],["fixw","Fixed weight","kg"],["tare","Tare","kg"],["minw","Minimum weight","kg"],["maxw","Maximum weight","kg"],["ingsel","Ingredient selection","sel",["Ingredients in %","Without"]],["ingp","Ingredient proportion","pct"],["pcs","Pieces/ package","int"],["life1","Shelf-life 1","int"],["life2","Shelf-life 2","int"],["dt1","Date text field 1","dtext"],["dt2","Date text field 2","dtext"],["dt3","Date text field 3","dtext"],["grp","Pr. group number","int"],["logo1","Logo number logo field1","logo"],["logo2","Logo number logo field2","logo"],["logo3","Logo number logo field3","logo"],["lseq","Label sequence number","roint"]],
 Texts:[["text","Article text","text"],["txt1","Text field 1","gtext"],["tf2","Text field 2","text"],["tf3","Text field 3","text"],["tf4","Text field 4","text"],["tf5","Text field 5","text"],["tf6","Text field 6","text"],["tf7","Text field 7","text"],["tf8","Text field 8","text"],["tf9","Text field 9","text"],["tf10","Text field 10","text"]],
 Barcode:[["cf1","Code field 1","code"],["cf2","Code field 2","code"],["cf3","Code field 3","code"],["cs1","Code sub-string 1","digits"],["cs2","Code sub-string 2","digits"],["cs3","Code sub-string 3","digits"],["cs4","Code sub-string 4","digits"],["cs5","Code sub-string 5","digits"]],
 "Label parameters":[["lab","Label param. No.","box"],["ltype","Label type","labsel"]],
 "Aut.lab. parameters":[["aut","Autom. labeler parameter no.","box"],["ref","Ref. point lab. distan.","asel",["Pckg. edge front side","Pckg. edge back side","Pckg. edge middle"]],["lpos","Label position","aro",()=>"Label position 1"],["opmode","Operating mode","asel",["Transport","Weighing Continuous","Weighing Start/Stop"]],["lcmode","Labeler conv. mode","aro"],["lb1","LB1 susp. path","amm"],["lb23","LB2/3 susp. path","aromm"],["dist","Label distance","amm"],["pkh","Max pckg. height","amm"],["speed","Speed","aromin"],["conv","Conveying speed","aromin"]],
 "Package selection":[["wc","Weight class No.","wcbox"],["selcrit","Selection criteria","sel",["Unknown","Weight determination","Checking contents","Sliding contents control","Weight classification","FPV Europe","FPV Canada","IPEM","USA - Non USDA standardized","USA - USDA standardized","USA - USDA only liquid mater","Free selection"]]]};
const ACE_ORDER=["General","Texts","Barcode","Label parameters","Aut.lab. parameters","Package selection"];
/* job settings that the Aut.lab. tab shows live (they belong to the labeling job, not the PLU record) */
const ACE_JOB={ref:"ref",opmode:"opmode",lcmode:"lcmode",lb1:"lb1",lb23:"lb23",dist:"dist",speed:"speed",conv:"conv"};
function openACE(mode,plu){if(S.running&&mode==="maint"&&plu===S.plu){toast("Stop the automat before changing the running article");return}
 const src=mode==="current"?P():DB.plu[plu];if(!src){toast("No article");return}
 const A=AT()||{};const job={};Object.keys(ACE_JOB).forEach(k=>job[k]=S.set[ACE_JOB[k]]);job.pkh=A.pkh||0;
 S.ace={mode,plu:mode==="current"?S.plu:plu,tab:mode==="pkg"?"Package selection":"General",data:clone(src),job,slave:clone(src.slave||{}),top:0,sel:null,dirty:false};S.editor=null;S.matrix=null;S.quick=false;render()}
function aceVal(row,col){const A=S.ace,d=A.data,k=row[0],t=row[2];
 const dv=(kk)=>{if(col===1&&A.slave[kk]!=null)return A.slave[kk];return A.job[kk]};
 switch(t){
  case"text":return d[k]||"";case"sel":return d[k]||row[3][0];case"money":return"£ "+f2(d[k]||0);
  case"kg":return f3((d[k]||0)/1000)+"kg";case"pct":return(d[k]||0)+"%";case"int":case"roint":return String(d[k]||0);
  case"ro":case"aro":return row[3]?row[3]():String(dv(k)??"");
  case"dtext":return d[k]>0?d[k]+": "+(DB.dtext[d[k]]||""):"";
  case"logo":return d[k]>0?`${d[k]} (${DB.logo[d[k]]||"?"})`:"0";
  case"gtext":return d[k]>0?(DB.text[d[k]]??""):(S.ace.mode==="current"&&S.jobText&&S.jobText[1]!=null?S.jobText[1]:"");
  case"code":return d[k]>0&&DB.code[d[k]]?`${DB.code[d[k]].name}\nCode substring 1: ${d.cs1}`:"";
  case"digits":return d[k]??"0";
  case"box":return String(d[k]??"");
  case"labsel":{const L=DB.label[d.lab];return L?`${d.lab}: ${L.name}`:"—"}
  case"asel":return dv(k);case"amm":case"aromm":return(dv(k)??0)+"mm";case"aromin":return(dv(k)??0)+"m/min";
  case"wcbox":return d.wc>0?String(d.wc):"0"}
 return""}
function aceEdit(row,col){const A=S.ace,d=A.data,k=row[0],t=row[2];const done=()=>{A.dirty=true;render()};
 const setDev=(v)=>{if(col===1)A.slave[k]=v;else A.job[k]=v;done()};
 if(/^(ro|aro|aromm|aromin|roint)$/.test(t)){toast("Display only in this table");return}
 if(t==="text")return textIn(row[1],d[k]||"",v=>{d[k]=v;done()});
 if(t==="sel")return sel(row[1],row[3],d[k]||row[3][0],v=>{d[k]=v;done()});
 if(t==="money")return num(row[1],"£",d[k]||0,v=>{d[k]=Number(v)||0;done()});
 if(t==="kg")return num(row[1],"kg",f3((d[k]||0)/1000),v=>{d[k]=Math.round(Number(v)*1000)||0;done()});
 if(t==="pct"||t==="int")return num(row[1],t==="pct"?"%":"",d[k]||0,v=>{d[k]=Number(v)||0;done()});
 if(t==="dtext")return sel(row[1],["0  none"].concat(Object.entries(DB.dtext).map(([n,x])=>n+"  "+x)),"",v=>{d[k]=Number(v.split(" ")[0]);done()});
 if(t==="logo")return sel(row[1],["0  none"].concat(Object.entries(DB.logo).map(([n,x])=>n+"  ("+x+")")),"",v=>{d[k]=Number(v.split(" ")[0]);done()});
 if(t==="gtext")return sel("Text field 1 · general text number",["0  none"].concat(Object.entries(DB.text).map(([n,x])=>n+"  "+x)),"",v=>{d[k]=Number(v.split(" ")[0]);done()});
 if(t==="code")return sel(row[1],["0  none"].concat(Object.entries(DB.code).map(([n,x])=>n+"  "+x.name)),"",v=>{d[k]=Number(v.split(" ")[0]);done()});
 if(t==="digits")return num(row[1],"",d[k]||"0",v=>{d[k]=String(v||"0");done()});
 if(t==="box"){const kind=k==="lab"?"label":"auto";return num(row[1],"",d[k],v=>{v=Number(v);if(v!==0&&!DB[kind][v]){msgDlg("Table with this number does not exist!");return}d[k]=v;done()})}
 if(t==="labsel")return sel(row[1],Object.entries(DB.label).map(([n,L])=>n+": "+L.name),"",v=>{d.lab=Number(v.split(":")[0]);done()});
 if(t==="asel")return sel(row[1],row[3],A.job[k],setDev);
 if(t==="amm")return num(row[1]+(col===1?" · slave":""),"mm",col===1&&A.slave[k]!=null?A.slave[k]:A.job[k],v=>setDev(Number(v)||0));
 if(t==="wcbox")return toast("Use the Weight class button to edit the table")}
function aceSave(close){const A=S.ace;
 const applyJob=()=>{Object.keys(ACE_JOB).forEach(k=>{if(A.job[k]!=null)S.set[ACE_JOB[k]]=A.job[k]})};
 if(A.mode==="current"){S.jobOv={plu:S.plu,data:Object.assign(clone(A.data),{slave:clone(A.slave)})};applyJob();A.dirty=false;log("Edit current article data: accepted for the current labeling job");toast("Changes active for the current labeling job. A PLU call-up brings back the database values.");S.ev.ace=true;if(close)S.ace=null;checkDrills();render();return}
 const go=()=>{DB.plu[A.plu]=Object.assign(clone(A.data),{slave:clone(A.slave)});if(A.plu===S.plu){S.jobOv=null;applyJob()}A.dirty=false;log(`Article data maintenance: PLU ${A.plu} saved in the database`);
  toast(`PLU ${A.plu} saved`+(LINE().online===false?"":". This line is online: the next BRAIN2 download overwrites it."));S.ev.acm=true;if(close)S.ace=null;checkDrills();render()};
 conf("Security check",`Save PLU ${A.plu} in the database?`,go)}
function aceHTML(){const A=S.ace,rows=ACE_TABS[A.tab],devs=LINE().dev;const cols=devs.length>1?[0,1]:[0];const vis=11;const top=Math.min(A.top,Math.max(0,rows.length-vis));A.top=top;
 const title=A.mode==="maint"?`Article data maintenance - PLU: ${A.plu}`:A.mode==="pkg"?"Package selection":"Edit current article data";
 let h=`<div class="ace"><div class="ace-t">${esc(title)}${A.dirty?' <span class="sm">· changed</span>':""}</div>
 <div class="ace-ico"><span class="aico" data-a="acefind" title="Search">⌕</span><span class="aico" data-a="aceprint" title="Print">⎙</span><span class="aico" data-a="acehelp" title="Help">?</span></div>
 <table class="acet"><tr><th></th>${cols.map(c=>`<th class="devh">${devs[c].bus}: ${esc(devs[c].name)}</th>`).join("")}</tr>`;
 rows.slice(top,top+vis).forEach((r,i)=>{const ri=top+i;const ro=/^(ro|aro|aromm|aromin|roint|box|wcbox)$/.test(r[2]);
  h+=`<tr><td><div class="rowb ${ro?"ro":""}" data-acer="${ri}">${esc(r[1])}</div></td>${cols.map(c=>{const v=aceVal(r,c);const on=A.sel&&A.sel[0]===ri&&A.sel[1]===c;
   return`<td class="acv ${ro?"ro":""} ${on?"on":""} ${r[2]==="box"||r[2]==="wcbox"?"box":""}" data-acec="${ri},${c}"><span>${esc(v).replace(/\n/g,"<br>")}</span>${ro?"":"<i class=\"tri\"></i>"}</td>`}).join("")}${r[2]==="wcbox"?`<td class="acebtns"><div class="cbtn" data-a="acewcdel" style="width:76px;height:32px">Delete weight class</div><div class="cbtn" data-a="acewc" style="width:76px;height:32px">Weight class</div></td>`:""}</tr>`});
 h+=`</table>`;
 if(rows.length>vis)h+=`<div class="ace-up ${top>0?"":"dim"}" data-a="aceup">⇧</div><div class="ace-dn ${top+vis<rows.length?"":"dim"}" data-a="acedn">⇩</div>`;
 if(A.mode!=="pkg")h+=`<div class="ace-tabs">${ACE_ORDER.map(t=>`<div class="atab ${t===A.tab?"on":""}" data-acetab="${t}">${esc(t)}</div>`).join("")}</div>`;
 h+=`<div class="ace-bot"><div class="cbtn" data-a="aceclose" title="Close">✕</div>${A.mode==="maint"?`<div class="cbtn" data-a="acesave" title="Save">💾</div>`:""}<div class="cbtn" data-a="aceok" title="Accept"><span class="ck">✓</span></div></div></div>`;
 return h}
function aceClick(desc){const A=S.ace;
 if(desc.acetab){A.tab=desc.acetab;A.top=0;A.sel=null;render();return true}
 if(desc.acer!=null){const r=ACE_TABS[A.tab][desc.acer];showRowHelp(r);A.sel=[desc.acer,0];render();return true}
 if(desc.acec){const [ri,c]=desc.acec.split(",").map(Number);const r=ACE_TABS[A.tab][ri];if(A.sel&&A.sel[0]===ri&&A.sel[1]===c)aceEdit(r,c);else{A.sel=[ri,c];S.aceDev=c;render()}return true}
 const a=desc.a;if(!a||!a.startsWith("ace"))return false;
 if(a==="aceup"){A.top=Math.max(0,A.top-6);render();return true}
 if(a==="acedn"){A.top=A.top+6;render();return true}
 if(a==="acefind"){sel("Find attribute",ACE_ORDER.flatMap(t=>ACE_TABS[t].map(r=>t+" › "+r[1])),"",v=>{const [t,l]=v.split(" › ");A.tab=t;const i=ACE_TABS[t].findIndex(r=>r[1]===l);A.top=Math.max(0,i-3);A.sel=[i,0]});return true}
 if(a==="aceprint"){toast("Article record sent to the list printer (simulated)");return true}
 if(a==="acehelp"){info(A.mode==="maint"?"Article data maintenance":"Edit current article data",A.mode==="maint"?"Changes the PLU record in this terminal's database. Save with the disk key or ✓. On an online line BRAIN2 overwrites local records at its next download, so permanent changes belong in BRAIN2.":"Changes the article for the labeling job that is running now. Nothing is written to the database: the next PLU call-up loads the stored record again. Columns are the devices on the system bus: 2 is the master, 3 the slave labeller.");return true}
 if(a==="acewc"){SPECIAL["Weight class editor"].run();S.ace=null;render();return true}
 if(a==="acewcdel"){PFN.wcdel();return true}
 if(a==="aceclose"){if(A.dirty)conf("Close","Discard your changes?",()=>{S.ace=null;S.aceDev=null;render()});else{S.ace=null;S.aceDev=null}render();return true}
 if(a==="acesave"){aceSave(false);return true}
 if(a==="aceok"){if(!A.dirty){S.ace=null;S.aceDev=null;render();return true}if(A.mode==="current")aceSave(true);else aceSave(true);return true}
 return false}
function showRowHelp(r){const lbl=r[1];const node=findByLabel(lbl)||findByLabel(lbl.replace(/\s*\(s\)$/,""));if(node&&node.secs.length)showManual({label:lbl,node,sec:null});else $("#manBody").innerHTML=`<div class="t">${esc(lbl)}</div><div class="meta">Field of the article record (photographed on FP2-L04). No single manual section covers it.</div>`}

/* ---------------------------------------------------------------- full-screen text editor (photographed) */
function openTED(n){const t=n>0?(DB.text[n]??""):"";S.ted={n:n||0,val:t,shift:false,over:false,align:"Left",size:"Normal"};S.dlg=null;render()}
function tedHTML(){const T=S.ted;const lines=T.val.split("\n");
 const row1=[["tload","Load"],["tsize","Character size"],["talign","Alignment"],["tctrl","Add contr. characters"],["tscope","Scope"],["tover","Overwrite",true],["tsave","Save"],["texit","Exit"]];
 let h=`<div class="ted"><div class="ted-h"><span>Text${T.n?" "+T.n:""}</span><span>Character ${T.val.length} &nbsp; Paragraph ${lines.length} &nbsp; Text length ${T.val.length} &nbsp; max text length 2000</span></div>
 <div class="ted-a" style="text-align:${T.align.toLowerCase()}">${esc(T.val)}<span class="caret">|</span></div>
 <div class="ted-r1">${row1.map(([a,l,sp])=>`<div class="sk2" data-a="${a}">${sp?`<span class="spl"><span class="${T.over?"":"on"}">Off</span><span class="${T.over?"on":""}">On</span></span>`:`<span class="v">▬</span>`}<span>${esc(l)}</span></div>`).join("")}</div>`;
 h+=qwerty({shift:T.shift},true)+`</div>`;return h}
function tedKey(k){const T=S.ted;if(k==="ESC"){S.ted=null;render();return}if(k==="SHIFT"){T.shift=!T.shift;render();return}if(k==="2ND"||k==="ALT")return;
 if(k==="BS")T.val=T.val.slice(0,-1);else if(k==="OK")T.val+="\n";else T.val+=(/[A-ZÜÖÄ]/.test(k)&&!T.shift)?k.toLowerCase():k;if(T.val.length>2000)T.val=T.val.slice(0,2000);render()}
function tedAct(a){const T=S.ted;
 if(a==="tload"){num("Load text · text number","",T.n||"",v=>{v=Number(v);if(DB.text[v]==null){msgDlg("Text with this number does not exist!");return}T.n=v;T.val=DB.text[v];S.dlg=null;render()});return true}
 if(a==="tsize"){sel("Character size",["Normal","Double height","Double width","Small"],T.size,v=>{T.size=v});return true}
 if(a==="talign"){sel("Alignment",["Left","Center","Right"],T.align,v=>{T.align=v});return true}
 if(a==="tctrl"){sel("Add contr. characters",["Line feed","Tab","Bold on/off","Underline on/off"],"",v=>{T.val+=v==="Line feed"?"\n":v==="Tab"?"\t":"";render()});return true}
 if(a==="tscope"){toast("Scope: whole text");return true}
 if(a==="tover"){T.over=!T.over;render();return true}
 if(a==="tsave"){num("Save text · text number","",T.n||"",v=>{v=Number(v);if(!(v>0)){toast("Enter a text number");return}conf("Security check",`Save text ${v} in the database?`,()=>{DB.text[v]=T.val;T.n=v;log(`Text ${v} saved in the database (text editor)`);S.ev.tedSaved=true;toast(`Text ${v} saved`+(P().txt1===v?". The active PLU uses it in Text field 1.":""));checkDrills()})});return true}
 if(a==="texit"){S.ted=null;render();return true}
 return false}
