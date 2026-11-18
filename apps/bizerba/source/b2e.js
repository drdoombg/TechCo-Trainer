/* ---------------------------------------------------------------- modals */
function b2Modal(){const M=B2.modal;let b="";
 if(M.k==="export")b=exportHTML(M);else if(M.k==="dl")b=dlHTML(M);
 else if(M.k==="confirm")b=`<div class="wiz sm2"><div class="wizh"><b>${esc(M.title)}</b></div><div class="wizb"><p>${esc(M.text)}</p></div><div class="wizf">${bb("mdyes","Yes",{pri:true})}${M.no?bb("mdno","No"):""}${bb("wizcancel",M.no?"Cancel":"No")}</div></div>`;
 else if(M.k==="info")b=`<div class="wiz sm2"><div class="wizh"><b>${esc(M.title)}</b></div><div class="wizb">${M.html}</div><div class="wizf">${bb("wizcancel","OK",{pri:true})}</div></div>`;
 else if(M.k==="pick")b=`<div class="wiz sm2"><div class="wizh"><b>${esc(M.title)}</b></div><div class="wizb"><div class="chkl">${M.items.map(([v,l])=>`<div class="pick" data-b="${M.act}" data-v="${esc(v)}">${esc(l)}</div>`).join("")}</div></div><div class="wizf">${bb("wizcancel","Cancel")}</div></div>`;
 else if(M.k==="new")b=`<div class="wiz sm2"><div class="wizh"><b>${esc(M.title)}</b></div><div class="wizb">${fi(M.label,`<input id="b2New" value="${esc(M.val||"")}">`)}</div><div class="wizf">${bb("newok","OK",{pri:true})}${bb("wizcancel","Cancel")}</div></div>`;
 return`<div class="b2modal">${b}</div>`}

/* ---------------------------------------------------------------- render + fit */
let b2Scale=1;
function b2Render(){const st=$("#b2stage");if(!st||B2.view!=="b2")return;const sx=st.parentElement.scrollTop;
 const fs=document.activeElement&&document.activeElement.id&&st.contains(document.activeElement)?{id:document.activeElement.id,s:document.activeElement.selectionStart}:null;
 const g=st.querySelector(".b2grid.tall");const gs=g?g.scrollTop:0;const tr=st.querySelector(".ldtree");const ts=tr?tr.scrollTop:0;const cv=st.querySelector(".ldcanvas");const cs=cv?[cv.scrollTop,cv.scrollLeft]:null;const dl=st.querySelector("#dmList");const dls=dl?dl.scrollTop:0;const dt=st.querySelector("#dmTxList");const dts=dt?dt.scrollTop:0;const dtr=st.querySelector(".dmtree");const dtrs=dtr?dtr.scrollTop:0;
 st.innerHTML=b2Shell();
 const g2=st.querySelector(".b2grid.tall");if(g2){g2.scrollTop=gs;const on=g2.querySelector("tr.on");if(on&&gs===0)g2.scrollTop=Math.max(0,on.offsetTop-80)}const tr2=st.querySelector(".ldtree");if(tr2)tr2.scrollTop=ts;const cv2=st.querySelector(".ldcanvas");if(cv2&&cs){cv2.scrollTop=cs[0];cv2.scrollLeft=cs[1]}const dl2=st.querySelector("#dmList");if(dl2){dl2.scrollTop=dls;const o=dl2.querySelector("option:checked");if(o&&dls===0)dl2.scrollTop=Math.max(0,o.offsetTop-40)}const dt2=st.querySelector("#dmTxList");if(dt2){dt2.scrollTop=dts;const o=dt2.querySelector("option:checked");if(o&&dts===0)dt2.scrollTop=Math.max(0,o.offsetTop-60)}const dtr2=st.querySelector(".dmtree");if(dtr2)dtr2.scrollTop=dtrs;
 if(fs){const el=document.getElementById(fs.id);if(el){el.focus();try{if(fs.s!=null)el.setSelectionRange(fs.s,fs.s)}catch(e){}}}}
function b2Fit(){const wrap=$("#b2wrap");if(!wrap||B2.view!=="b2")return;const W=1440,H=860;const availW=wrap.parentElement.clientWidth;let s;
 if(viewMode==="laptop"){const top=wrap.getBoundingClientRect().top+window.scrollY;const availH=Math.max(320,window.innerHeight-top-12);s=Math.min(availW/W,availH/H,1.2)}else s=Math.min(1,availW/W);
 b2Scale=s;$("#b2stage").style.transform=`scale(${s})`;wrap.style.height=(H*s)+"px";wrap.style.width=(W*s)+"px"}
function setView(v,mod){if(v==="ld")v="b2";B2.view=v==="term"?"term":"b2";B2.app=v==="dm"?"dm":"brain";store.set("appView",v);
 if(mod){const cat=Object.keys(B2_CATS).find(c=>B2_CATS[c].includes(mod));if(cat){B2.cat=cat;B2.mod=mod}}
 $$("#appTabs button").forEach(b=>b.setAttribute("aria-pressed",b.dataset.view===v));
 const term=B2.view==="term";$("#term").hidden=!term;$("#lineCard").hidden=!term;$("#lines").hidden=!term;$("#explain").hidden=!term;$("#b2wrap").hidden=term;
 if(!term){if(B2.app==="dm")dmInit();b2Render();b2Fit();setTimeout(b2Fit,30)}else{render();fit();setTimeout(fit,30)}}
/* ---------------------------------------------------------------- events */
function b2Click(e){if(B2.app==="dm"&&!B2.modal){dmClick(e);return}const t=e.target.closest("[data-b]");const st=$("#b2stage");
 if((B2.ld.menu&&!e.target.closest(".ldmenu"))||(B2.ld.ctx&&!e.target.closest(".ldctx .sub"))){const had=B2.ld.ctx&&e.target.closest(".ldctx");B2.ld.menu=null;if(!had)B2.ld.ctx=null;if(!t){b2Render();return}}
 const el=e.target.closest("[data-lel]");if(el&&!t){B2.ld.sel=el.dataset.lel;b2Render();return}
 if(!t){if(e.target.closest(".ldpaper")===null&&e.target.closest(".ldcanvas")&&!e.target.closest(".lel")){B2.ld.sel=null;b2Render()}return}
 if(t.disabled)return;const a=t.dataset.b,v=t.dataset.v;const D=B2.ld,M=B2.md;
 const R=()=>b2Render();
 switch(a){
  case"cat":B2.cat=v;B2.mod=B2_CATS[v].includes(B2.mod)?B2.mod:(v==="Product data"?"Master Data":B2_CATS[v][0]);break;
  case"mod":B2.mod=v;if(v==="Audit Trail")b2Ev("auditViewed");break;
  case"navtog":B2.navOpen=!B2.navOpen;break;
  case"burger":B2.modal={k:"info",title:"BRAIN2",html:"Applications on this workstation: Product data, Process data, Result data, Reporting and the global functions under …. Pick them on the bar at the bottom."};break;
  /* Label Designer */
  case"ldmenu":D.menu=D.menu===v?null:v;break;
  case"ldmi":ldMenu(v);break;
  case"ldopen":if(!D.open.includes(v))D.open.push(v);D.cur=v;D.sel=null;B2.modal=null;break;
  case"ldclose":e.stopPropagation();ldClose(v);break;
  case"ldnew":newLayout();break;case"ldsave":ldSave();break;case"ldtest":ldMenu("File:Print test label");break;case"ldexport":openExport();break;
  case"ldundo":ldUndo();break;case"lddel":ldDel();break;case"ldcopy":ldMenu("Edit:Copy");ldMenu("Edit:Paste");break;
  case"ldadd":case"ldaddm":ldAdd(v);break;
  case"ldsaveall":Object.keys(D.dirty).filter(n=>D.dirty[n]).forEach(n=>{const c=D.cur;D.cur=n;ldSave();D.cur=c});break;
  case"ldredo":ldRedo();break;case"ldlprop":D.sel=null;break;case"ldgrid":D.grid=!D.grid;break;case"ldsnap":D.snap=!D.snap;break;case"ldlegend":D.legend=!D.legend;break;case"ldruler":D.rulers=!D.rulers;break;
  case"ldzfit":ldMenu("View:Fit to window");break;case"ldz100":D.zoom=4;break;case"ldprev":D.data=!D.data;break;
  case"ldset":{const E=ldEl();if(E){const [k,val]=v.split(":");ldPush();E[k]=val}break}
  case"ldfont":{const E=ldEl();if(E){const ks=Object.keys(FONTS);ldPush();E.font=ks[(ks.indexOf(E.font)+1)%ks.length];E.size=FONTS[E.font]}break}
  case"ldfit":{const E=ldEl();if(E){ldPush();if(v==="h")E.h=Math.max(1,Math.round((E.size||1.5)*1.25*10)/10);else E.w=Math.max(2,Math.round(String(fldName(E)).length*(E.size||1.5)*.5))}break}
  case"ldorder":ldOrder(v);D.ctx=null;break;case"ldalg":ldAlign(v);break;
  case"ldalpha":D.alpha=v==="1";break;case"ldunit":D.unit=v;break;case"ldcoll":{D.coll=D.coll||{};D.coll[v]=!D.coll[v];break}
  case"lddock":D.dock=v;break;case"ldselel":D.sel=v;break;
  case"ldpvsel":D.pvsel=v;if(e.detail>=2){if(!D.open.includes(v))D.open.push(v);D.cur=v;D.sel=null}break;
  case"ldpvopen":{const n=String(D.pvsel??D.cur);if(n&&B2.db.layouts[n]){if(!D.open.includes(n))D.open.push(n);D.cur=n;D.sel=null}break}
  case"ldpvdel":{B2.lm.sel=String(D.pvsel??D.cur);const n=B2.lm.sel;const used=Object.values(B2.db.label).some(l=>l.layout===String(n));if(used){B2.modal={k:"info",title:"Not deleted",html:`Layout ${n} is the Label type of a label parameter set. Change the set in Data Maintenance first.`};break}B2.modal={k:"confirm",title:"Delete layout",text:`Delete layout ${n}?`,yes:()=>{delete B2.db.layouts[n];D.open=D.open.filter(x=>x!==n);if(D.cur===n)D.cur=D.open[0]||null;b2Audit("Label Designer","Layout "+n,"Deleted","","")}};break}
  case"ldpvref":toast("Project view refreshed");break;
  case"ldchg":ldChangeType(v);D.ctx=null;break;
  case"ldalign":{const E=ldEl();if(E){ldPush();E.halign=v}break}
  case"ldzin":D.zoom=Math.min(14,D.zoom+1);break;case"ldzout":D.zoom=Math.max(2,D.zoom-1);break;
  /* wizards */
  case"wiznext":{const X=B2.modal;if(X.k==="export"){if(X.step===1&&!X.res.length){toast("Select at least one layout");return}if(X.step===2){if(!X.devs.length){toast("Select at least one device");return}runExport(X);return}X.step++}break}
  case"wizback":B2.modal.step--;break;
  case"wizcancel":case"wizclose":B2.modal=null;break;
  case"dlgo":runDl(B2.modal);return;
  case"mdyes":{const f=B2.modal.yes;B2.modal=null;f();break}
  case"mdno":{const f=B2.modal.no;B2.modal=null;f&&f();break}
  case"newok":{const f=B2.modal.ok;const val=$("#b2New").value;B2.modal=null;f(val);break}
  /* Layout Management */
  case"lmsel":B2.lm.sel=v;break;
  case"lmopen":setView("ld");if(!D.open.includes(B2.lm.sel))D.open.push(B2.lm.sel);D.cur=B2.lm.sel;B2.cat="Product data";B2.mod="Label Designer";break;
  case"lmcopy":{const s=B2.db.layouts[B2.lm.sel];let n=s.no+100;while(B2.db.layouts[n])n++;B2.db.layouts[n]=Object.assign(clone(s),{no:n,name:n+" "+s.name.replace(/^\d+ /,"")+" copy",ver:1,status:"Draft",changed:fmtD(new Date()),by:B2USER});B2.lm.sel=String(n);b2Audit("Layout Management","Layout "+n,"Copied from",s.no,n);break}
  case"lmdel":{const n=B2.lm.sel;const used=Object.values(B2.db.label).some(l=>l.layout===String(n));if(used){B2.modal={k:"info",title:"Not deleted",html:`Layout ${n} is used by a label parameter table. Point the table at another layout first (Process data › Label Parameters).`};break}B2.modal={k:"confirm",title:"Delete layout",text:`Delete layout ${n}?`,yes:()=>{delete B2.db.layouts[n];D.open=D.open.filter(x=>x!==n);b2Audit("Layout Management","Layout "+n,"Deleted","","");B2.lm.sel=Object.keys(B2.db.layouts)[0]}};break}
  case"lmrel":{const L=B2.db.layouts[B2.lm.sel];const bad=lmChecks(L).filter(c=>!c[1]);if(bad.length){B2.modal={k:"info",title:"Release refused",html:"<ul>"+bad.map(c=>`<li>${esc(c[0])}</li>`).join("")+"</ul>"};break}const old=L.status;L.status="Released";b2Audit("Layout Management","Layout "+L.no,"Status",old,"Released");b2Ev("layoutReleased");toast(`Layout ${L.no} version ${L.ver} released`);break}
  case"lmexp":openExport([B2.lm.sel]);break;
  /* Code Structure */
  case"cssel":B2.cs.sel=Number(v);break;
  case"csnew":{let n=1;while(B2.db.code[n])n++;B2.db.code[n]={name:"New code structure",type:"EAN13",segs:[["Constant","2",1],["PLU number","",6],["Price check digit","",1],["Price","",4],["Check digit","",1]]};B2.cs.sel=n;b2Audit("Code Structure","Code "+n,"Created","","");break}
  case"csdel":{const n=B2.cs.sel;if(Object.values(B2.db.plu).some(p=>p.cf1===n||p.cf2===n||p.cf3===n)){B2.modal={k:"info",title:"Not deleted",html:"Articles still point to this code structure in a Code field."};break}delete B2.db.code[n];B2.cs.sel=Number(Object.keys(B2.db.code)[0]);break}
  case"cssave":b2Audit("Code Structure","Code "+B2.cs.sel,"Saved","","");toast("Code structure saved");break;
  case"csadd":B2.db.code[B2.cs.sel].segs.push(["Constant","0",1]);break;
  case"csrm":B2.db.code[B2.cs.sel].segs.splice(Number(v),1);break;
  case"csup":{const s=B2.db.code[B2.cs.sel].segs,i=Number(v);if(i>0)[s[i-1],s[i]]=[s[i],s[i-1]];break}
  /* Master Data */
  case"mdsel":if(M.dirty){B2.modal={k:"confirm",title:"Master Data",text:"Discard unsaved changes to PLU "+M.sel+"?",yes:()=>{M.dirty=false;M.work=null;M.sel=Number(v)}};break}M.sel=Number(v);M.work=null;break;
  case"mdtab":M.tab=v;break;
  case"mdsave":{const before=B2.db.plu[M.sel];Object.keys(M.work).forEach(k=>{if(JSON.stringify(M.work[k])!==JSON.stringify(before[k]))b2Audit("Master Data","PLU "+M.sel,k,typeof before[k]==="object"?"…":before[k],typeof M.work[k]==="object"?"…":M.work[k])});B2.db.plu[M.sel]=M.work;M.work=null;M.dirty=false;if(B2.status[M.sel]==="Released")B2.status[M.sel]="Changed";b2Ev("mdSaved");toast(`PLU ${M.sel} saved in BRAIN2. Release it, then download to the lines.`);break}
  case"mddiscard":M.work=null;M.dirty=false;break;
  case"mdnew":case"mdcopy":B2.modal={k:"new",title:a==="mdnew"?"New article":"Copy article "+M.sel,label:"New PLU number",val:String(Math.max(...Object.keys(B2.db.plu).map(Number))+1),ok:val=>{const n=Number(val);if(!(n>0)||B2.db.plu[n]){toast("Enter a free PLU number");return}B2.db.plu[n]=a==="mdnew"?Object.assign(clone(B2.db.plu[M.sel]),{plu:n,text:"",cust:"",tf2:"",tf10:""}):Object.assign(clone(B2.db.plu[M.sel]),{plu:n,text:B2.db.plu[M.sel].text+" (copy)"});B2.status[n]="Draft";M.sel=n;M.q="";b2Audit("Master Data","PLU "+n,"Created","","Draft")}};break;
  case"mddel":B2.modal={k:"confirm",title:"Delete article",text:`Delete PLU ${M.sel} from BRAIN2? It disappears from the lines at the next download.`,yes:()=>{delete B2.db.plu[M.sel];b2Audit("Master Data","PLU "+M.sel,"Deleted","","");M.sel=Number(Object.keys(B2.db.plu)[0]);M.work=null;M.dirty=false}};break;
  case"mdrel":mdRelease(M.sel);break;
  case"mdsend":B2.cat="Process data";B2.mod="Devices";B2.dev.sel=B2.dev.sel||"FP2-L04-Glmi-M";break;
  /* Exceptions */
  case"exrm":{const x=B2.exc.splice(Number(v),1)[0];b2Audit("Master Data Exceptions",x.dev,x.label||x.field,x.val,"");break}
  case"exadd":{const dev=$("#exDev").value,plu=Number($("#exPlu").value)||0,field=$("#exField").value,val=Number($("#exVal").value);if($("#exVal").value===""||isNaN(val)){toast("Enter a number");return}if(plu&&!B2.db.plu[plu]){toast("PLU not found");return}const f=EX_FIELDS.find(x=>x[0]===field);B2.exc.push({dev,plu,field,label:f[1],val,note:"",by:B2USER,at:fmtD(new Date())});B2.ex={dev,plu:plu?String(plu):"",field,val:""};b2Audit("Master Data Exceptions",dev,f[1],"",val);b2Ev("excSaved");toast("Exception saved. It reaches the device with the next download.");break}
  /* Text */
  case"txkind":B2.tx.kind=v;B2.tx.sel=Number(Object.keys(v==="text"?B2.db.text:B2.db.dtext)[0]);B2.tx.work=null;break;
  case"txsel":B2.tx.sel=Number(v);B2.tx.work=null;break;
  case"txnew":{const src=B2.tx.kind==="text"?B2.db.text:B2.db.dtext;let n=B2.tx.kind==="text"?21:1;while(src[n]!=null)n++;src[n]="";B2.tx.sel=n;B2.tx.work="";b2Audit("Text",(B2.tx.kind==="text"?"General text ":"Date text ")+n,"Created","","");break}
  case"txdel":{const n=B2.tx.sel;const used=Object.values(B2.db.plu).some(p=>p.txt1===n);if(B2.tx.kind==="text"&&used){B2.modal={k:"info",title:"Not deleted",html:"Articles use this text in Text field 1."};break}const src=B2.tx.kind==="text"?B2.db.text:B2.db.dtext;delete src[n];B2.tx.sel=Number(Object.keys(src)[0]);break}
  case"txsave":{const src=B2.tx.kind==="text"?B2.db.text:B2.db.dtext;const n=B2.tx.sel;b2Audit("Text",(B2.tx.kind==="text"?"General text ":"Date text ")+n,"Text",src[n],B2.tx.work);src[n]=B2.tx.work;B2.tx.work=null;b2Ev("textSaved");toast(`Text ${n} saved in BRAIN2. Download it to the lines to print it.`);break}
  /* Devices */
  case"dvsel":B2.dev.sel=v;break;
  case"dvdl":dlModal(B2.dev.sel);break;
  case"dvul":{const d=devLine(B2.dev.sel);const r=b2Upload(d.line);if(!r.ok){B2.modal={k:"info",title:"Upload",html:esc(r.msg)};break}B2.transfers.unshift({at:new Date(),user:B2USER,dev:d.id,what:"Upload PLU, texts, weight classes",res:"OK"});b2Audit("Process data",d.id,"Upload","","PLU, texts, weight classes");b2Ev("upload");toast(`Data read from ${d.id} into BRAIN2`);break}
  case"dvcmp":{const d=devLine(B2.dev.sel);const df=lineDiff(d.line);B2.modal={k:"info",title:"Compare · "+d.id,html:df.unknown?"No data read from this device yet.":df.n?"<ul>"+df.list.slice(0,40).map(x=>`<li>${esc(x)}</li>`).join("")+"</ul>":"No differences."};break}
  case"dvtime":toast("Device time synchronized with the BRAIN2 server");break;
  case"dvping":{const d=devLine(B2.dev.sel);toast(`Reply from ${d.ip}: time=${2+Math.floor(Math.random()*4)} ms`);break}
  /* Weight classes */
  case"wcsel":B2.wc.sel=Number(v);B2.wc.work=null;B2.wc.cell=null;break;
  case"wccol":B2.wc.cell=Number(v);break;
  case"wcadd":{const W=B2.wc.work=B2.wc.work||clone(B2.db.wc[B2.wc.sel]);let i=B2.wc.cell!=null?B2.wc.cell:1;if(W.cols[i].up==null)i=Math.max(0,i-1);const lo=i?W.cols[i-1].up:W.lower;const up=Math.round((lo+(W.cols[i].up??W.upper))/2);W.cols.splice(i,0,Object.assign(clone(W.cols[i]),{up}));B2.wc.cell=i;break}
  case"wcrm":{const W=B2.wc.work=B2.wc.work||clone(B2.db.wc[B2.wc.sel]);if(W.cols.length<=1)break;W.cols.splice(B2.wc.cell,1);W.cols[W.cols.length-1].up=null;B2.wc.cell=null;break}
  case"wcsave":{const n=B2.wc.sel;b2Audit("Weight Classes","Table "+n,"Saved",B2.db.wc[n].cols.length+" columns",B2.wc.work.cols.length+" columns");B2.db.wc[n]=B2.wc.work;B2.wc.work=null;b2Ev("wcSaved");toast(`Weight class table ${n} saved. Download it to the lines.`);break}
  case"wcdiscard":B2.wc.work=null;break;
  /* other */
  case"dtrun":{const ok=[];B2.devices.forEach(d=>{if(!d.legacy&&d.i===0){if(b2Download(d.line,{plu:true,text:true,wc:true}).ok)ok.push(d.line)}});B2.transfers.unshift({at:new Date(),user:B2USER,dev:"all online devices",what:"Synchronization: PLU, texts, weight classes",res:`OK (${ok.length} lines) · FP1-L03 skipped (view only)`});b2Ev("dl");if(S)render();toast("Synchronization finished");break}
  case"bknow":B2.lastBk=fmtD(new Date())+" "+new Date().toTimeString().slice(0,5)+" · manual · 413 MB";toast("Backup finished");break;
  case"repsave":case"repcsv":case"repxls":b2Ev("report");toast(`Report saved to C:\\BRAIN2\\Reports\\${B2.rd.rep.replace(/\s+/g,"_")}.${a==="repsave"?"pdf":a==="repcsv"?"csv":"xlsx"} (simulated: files can't be saved from this page)`);break;
 }
 R()}
function b2Change(e){if(B2.app==="dm"&&!B2.modal){dmChange(e);return}const t=e.target;const D=B2.ld;
 if(t.id==="ldElSel"){D.sel=t.value||null;b2Render();return}
 if(t.id==="ldPlu"){D.plu=Number(t.value);b2Render();return}
 if(t.id==="ldTok"&&t.value){const E=ldEl();if(E){ldPush();E.text=(E.text||"")+"{"+t.value+"}"}b2Render();return}
 if(t.dataset.lp){const k=t.dataset.lp;const toMM=x=>{const n=parseFloat(String(x).replace(",","."));return isNaN(n)?0:(t.dataset.unit&&D.unit==="Dot"?n/8:n)};
  if(k.startsWith("L.")){const L=ldLayout();ldPush();const f=k.slice(2);L[f]=/^(w|h|dist)$/.test(f)?toMM(t.value):t.value;b2Render();return}
  const E=ldEl();if(!E)return;ldPush();let v=t.value;
  if(k==="fld"){const f=(FLD_TYPES.find(x=>x[1]===v)||[])[0];if(f){D.undo.pop();ldChangeType(f)}b2Render();return}
  if(["x","y","w","h"].includes(k))v=toMM(v);else if(k==="size"){v=parseFloat(v);if(isNaN(v)){b2Render();return}}else if(k==="rot"||k==="n")v=Number(v);else if(k==="bold"||k==="cur")v=v==="true";else if(k==="text")v=v.replace(/⏎/g,"\n");else if(k==="font"){E.size=FONTS[v]||E.size}
  E[k]=v;if(k==="n"||k==="data")E.name=fldName(E);b2Render();return}
 if(t.dataset.md){const M=B2.md;M.work=M.work||clone(B2.db.plu[M.sel]);const k=t.dataset.md;let v=t.value;if(t.type==="number"||t.dataset.n)v=Number(v);if(k==="cs1"||/^cs\d$/.test(k))v=String(t.value);M.work[k]=v;M.dirty=true;b2Render();return}
 if(t.dataset.cs){B2.db.code[B2.cs.sel][t.dataset.cs]=t.value;b2Render();return}
 if(t.dataset.cseg){const [i,j]=t.dataset.cseg.split(",").map(Number);B2.db.code[B2.cs.sel].segs[i][j]=j===2?Number(t.value):t.value;b2Render();return}
 if(t.dataset.wcf){const W=B2.wc.work=B2.wc.work||clone(B2.db.wc[B2.wc.sel]);const k=t.dataset.wcf;W[k]=["lower","upper","stops"].includes(k)?Number(t.value):t.value;b2Render();return}
 if(t.dataset.wcc){const W=B2.wc.work=B2.wc.work||clone(B2.db.wc[B2.wc.sel]);const [i,k]=t.dataset.wcc.split(",");const c=W.cols[Number(i)];if(k==="up"){const v=Number(t.value);const lo=i>0?W.cols[i-1].up:0;const hi=W.cols[Number(i)+1]&&W.cols[Number(i)+1].up!=null?W.cols[Number(i)+1].up:99999;if(!(v>lo&&v<hi)){toast(`Enter a value between ${lo} and ${hi} g`);b2Render();return}c.up=v}else c[k]=k==="txt1"||k==="life1"?Number(t.value):t.value;b2Render();return}
 if(t.dataset.par){const [kind,n,f]=t.dataset.par.split(",");const o=B2.db[kind][n];const before=o[f];o[f]=t.type==="number"?Number(t.value):t.value;b2Audit(kind==="auto"?"Automation Parameters":"Label Parameters","Table "+n,f,before,o[f]);return}
 if(t.id==="rdLine"){B2.rd.line=t.value;b2Render();return}
 if(t.id==="repSel"){B2.rd.rep=t.value;b2Render();return}
 if(t.id==="mdCust"){B2.md.cust=t.value;b2Render();return}
 if(t.id==="dtNight"){B2.night=t.checked;return}
 if(t.dataset.wx){const M=B2.modal;const arr=t.dataset.wx==="res"?M.res:t.dataset.wx==="logo"?M.logos:M.devs;const i=arr.indexOf(t.value);if(t.checked&&i<0)arr.push(t.value);if(!t.checked&&i>=0)arr.splice(i,1);return}
 if(t.dataset.dl){B2.modal.what[t.dataset.dl]=t.checked;return}}
function b2Input(e){if(B2.app==="dm"&&!B2.modal){dmInput(e);return}const t=e.target;
 if(t.id==="ldTxt"){const E=ldEl();if(E){if(!B2.ld._typing){ldPush();B2.ld._typing=true;clearTimeout(B2.ld._tt);}B2.ld._tt=setTimeout(()=>B2.ld._typing=false,800);E.text=t.value;const pap=$("#b2stage .ldpaper");const L=ldLayout();const p=B2.db.plu[B2.ld.plu]||Object.values(B2.db.plu)[0];if(pap)pap.innerHTML=renderLayout(L,p,layoutCtx(p,null,"FP2-L03",B2.db),B2.ld.zoom,{design:true,data:B2.ld.data,sel:B2.ld.sel,db:B2.db})}return}
 if(t.id==="txArea"){B2.tx.work=t.value;const b=$('#b2stage [data-b="txsave"]');if(b)b.disabled=false;return}
 if(t.id==="mdQ"){B2.md.q=t.value;clearTimeout(b2Input.t);b2Input.t=setTimeout(b2Render,250);return}
 if(t.id==="auQ"){B2.au.q=t.value;clearTimeout(b2Input.t);b2Input.t=setTimeout(b2Render,250);return}
 if(t.id==="wcQ"){B2.wc.q=t.value;clearTimeout(b2Input.t);b2Input.t=setTimeout(b2Render,250);return}}
/* drag and resize label elements */
let ldDrag=null;
function b2Down(e){if(B2.app==="dm"||e.button===2)return;const h=e.target.closest("[data-lres]"),el=e.target.closest("[data-lel]");if(!h&&!el)return;const L=ldLayout();const E=L.els.find(x=>x.id===(h?h.dataset.lres:el.dataset.lel));if(!E)return;
 const wasDirty=!!B2.ld.dirty[B2.ld.cur];B2.ld.sel=E.id;ldPush();ldDrag={wasDirty,E,res:!!h,x0:e.clientX,y0:e.clientY,ex:E.x,ey:E.y,ew:E.w,eh:E.h};e.preventDefault();b2Render()}
function b2Move(e){if(!ldDrag)return;const k=B2.ld.zoom*b2Scale;const dx=(e.clientX-ldDrag.x0)/k,dy=(e.clientY-ldDrag.y0)/k;const E=ldDrag.E;const sn=v=>Math.round(v*2)/2;
 if(ldDrag.res){E.w=Math.max(.5,sn(ldDrag.ew+dx));E.h=Math.max(.4,sn(ldDrag.eh+dy))}else{E.x=sn(ldDrag.ex+dx);E.y=sn(ldDrag.ey+dy)}
 const pap=$("#b2stage .ldpaper");const L=ldLayout();const p=B2.db.plu[B2.ld.plu]||Object.values(B2.db.plu)[0];if(pap)pap.innerHTML=renderLayout(L,p,layoutCtx(p,null,"FP2-L03",B2.db),B2.ld.zoom,{design:true,data:B2.ld.data,sel:B2.ld.sel,db:B2.db})}
function b2Up(){if(!ldDrag)return;const E=ldDrag.E;if(E.x!==ldDrag.ex||E.y!==ldDrag.ey||E.w!==ldDrag.ew||E.h!==ldDrag.eh){B2.ld.dirty[B2.ld.cur]=true;b2Audit("Label Designer","Layout "+B2.ld.cur,E.name,`${ldDrag.ex},${ldDrag.ey}`,`${E.x},${E.y}`);b2Ev("layoutEdited")}else{B2.ld.undo.pop();B2.ld.dirty[B2.ld.cur]=ldDrag.wasDirty}ldDrag=null;b2Render()}
function b2Key(e){if(B2.view!=="b2")return;if(B2.app==="dm"){if((e.ctrlKey||e.metaKey)&&e.key==="s"){e.preventDefault();dmSave();b2Render()}if(e.key==="Escape"&&B2.dm.modal){B2.dm.modal=null;b2Render()}return}if(e.key==="Escape"&&(B2.ld.ctx||B2.ld.menu)){B2.ld.ctx=null;B2.ld.menu=null;b2Render();return}if(e.target.closest("input,textarea,select"))return;const E=ldEl();
 if(e.key==="F1"){e.preventDefault();B2.modal={k:"info",title:"Help · "+B2.mod,html:b2HelpText()};b2Render();return}
 if(B2.mod!=="Label Designer"||!E)return;const st={ArrowLeft:[-.5,0],ArrowRight:[.5,0],ArrowUp:[0,-.5],ArrowDown:[0,.5]}[e.key];
 if(st){e.preventDefault();ldPush();E.x+=st[0];E.y+=st[1];b2Render()}else if(e.key==="Delete"){ldDel();b2Render()}else if((e.ctrlKey||e.metaKey)&&e.key==="s"){e.preventDefault();ldSave();b2Render()}else if((e.ctrlKey||e.metaKey)&&e.key==="z"){e.preventDefault();ldUndo();b2Render()}}
function b2HelpText(){const H={"Label Designer":"Drag an element to move it, drag its bottom-right handle to resize it, or type exact values in the Property window. Arrow keys nudge 0.5 mm. Text in braces is a data field filled from the article: {Date 2} prints the use-by date, {Text field 1} the cooking time. Save the layout, release it in Layout Management, then File › Export to… sends it to the lines through _connect.BRAIN.",
 "Master Data":"Select an article, change it in the tabs, Save. Released articles that you change go back to Changed until released again. Send to devices opens Process data › Devices.",
 "Devices":"Select a device and Download. Both devices on a line receive the data. FP1 Line 3 is legacy: BRAIN2 can read it and send layouts, but it refuses master data.",
 "Weight Classes":"Click a column heading to select it, Add column splits it in two. Give each column its own Text field 1 to print a different cooking time per weight band. Save, then download."};return esc(H[B2.mod]||"Pick a module on the left. The bar at the bottom switches between Product data, Process data, Result data and Reporting.")}
function b2Ctx(e){if(B2.app==="dm"||B2.mod!=="Label Designer")return;const el=e.target.closest("[data-lel]");const pap=e.target.closest(".ldpaper");if(!el&&!pap)return;e.preventDefault();const st=$("#b2stage").getBoundingClientRect();B2.ld.sel=el?el.dataset.lel:null;B2.ld.ctx={x:(e.clientX-st.left)/b2Scale,y:(e.clientY-st.top)/b2Scale};b2Render()}
function b2Pos(e){if(B2.app==="dm"||B2.mod!=="Label Designer")return;const pap=e.target.closest&&e.target.closest(".ldpaper");const o=$("#ldPos");if(!pap||!o)return;const r=pap.getBoundingClientRect();const k=B2.ld.zoom*b2Scale;const x=Math.max(0,Math.round((e.clientX-r.left)/k*8)),y=Math.max(0,Math.round((e.clientY-r.top)/k*8));B2.ld.pos=x+", "+y;o.textContent="[Dots] "+B2.ld.pos}
function b2Wire(){const st=$("#b2stage");st.addEventListener("click",b2Click);st.addEventListener("contextmenu",b2Ctx);st.addEventListener("mousemove",b2Pos);st.addEventListener("change",b2Change);st.addEventListener("input",b2Input);st.addEventListener("pointerdown",b2Down);
 window.addEventListener("pointermove",b2Move);window.addEventListener("pointerup",b2Up);document.addEventListener("keydown",b2Key);
 $$("#appTabs button").forEach(b=>b.addEventListener("click",()=>setView(b.dataset.view,b.dataset.mod)));dmInit();window.addEventListener("resize",b2Fit)}

/* ---------------------------------------------------------------- BRAIN2 drills */
DRILLS.push(
 {g:"BRAIN2",id:"b2_text",t:"Change a cooking time permanently through BRAIN2",d:"In BRAIN2 › Product data › Text change general text 22 (a whole-bird cooking time). Save, download to your FP2 line (Process data › Devices › Download), then on the terminal call up a whole bird PLU whose Text field 1 is 22 and run 3 good packs.",h:"Bizerba Brain 2 tab · Text › select 22 › edit › Save. Process data › Devices › select FP2-L03-Glmi-M or FP2-L04-Glmi-M › Download… › Start download. Terminal: plu change to e.g. 006706 › start. The Text 1 column in the Line view shows the new text.",ck:()=>B2.ev.textSaved&&B2.ev.dl&&(S.ev.ctDb||0)>=3},
 {g:"BRAIN2",id:"b2_wc",t:"Cooking time per weight band, set in BRAIN2",d:"In Process data › Weight Classes split the accepted column of a whole-bird table and give each column its own Text field 1. Save, download to the line, run until two different cooking times print.",h:"Weight Classes › find the PLU's table › click the middle column heading › Add column › set Upper limit › Text field 1 per column › Save › Devices › Download. On the terminal the table needs enough columns: Service level (4/5) › Production data › Processing › Weight class &statistics › Wgt class columns.",ck:()=>B2.ev.wcSaved&&B2.ev.dl&&Object.keys(S.ev.ctWc||{}).length>=2},
 {g:"BRAIN2",id:"b2_fp1",t:"Try to send master data to FP1 Line 3",d:"Download PLU data to FP1-L03-Glmi-M and read why BRAIN2 refuses. Then think through how the same change is made on that line.",h:"Process data › Devices › FP1-L03-Glmi-M › Download. FP1 Line 3 is view only for master data; weight classes and cooking times are changed at its terminal (see the FP1 drill under Cooking time).",ck:()=>B2.ev.fp1Refused},
 {g:"BRAIN2",id:"b2_exc",t:"Give the slave labeller its own value",d:"Add a master data exception for a -S device (e.g. Label distance 0 mm for PLU 006419), download, and check the Slave column in Edit current article data on the terminal.",h:"Product data › Master Data Exceptions › New exception › Add. Devices › Download. Terminal level 2+: Article data › Edit current article data › Aut.lab. parameters.",ck:()=>B2.ev.excSaved&&B2.ev.dl},
 {g:"BRAIN2",id:"b2_layout",t:"Edit a label layout and print it on the line",d:"In Label Designer move or change an element of the layout your PLU prints (205 for M&S whole birds), save, export it to your line's devices, then run a pack on the terminal and look at the label preview.",h:"Label Designer › double-click 205.205 M&S NFO 2 in the Project view › drag a field or edit it in the Property window › Save (disk icon) › Export to… (page with green arrow) › tick 205 › Next › tick the FP2 devices › Next. Terminal: start the line.",ck:()=>B2.ev.layoutExported&&(S.ev.layoutPrinted||0)>=1},
 {g:"BRAIN2",id:"b2_release",t:"Quality gate: release a new article",d:"Copy an article to a new PLU number, fix what the release check complains about, release it.",h:"Master Data › Copy › new number › Release. The check refuses an empty article text, a missing label or weight class table, and a shelf-life outside the Text field 10 window.",ck:()=>B2.ev.released},
 {g:"BRAIN2",id:"b2_audit",t:"Find who changed PLU 6419's shelf-life",d:"Use the Audit Trail to find the user, workstation and old and new value.",h:"Bottom bar … › Audit Trail › filter 6419.",ck:()=>B2.ev.auditViewed},
 {g:"BRAIN2",id:"b2_report",t:"Give-away report for your line",d:"Run some packs on the terminal, then produce the give-away report for that line.",h:"Reporting › Reports › Give-away report › line filter on the Result data screens › Save as PDF.",ck:()=>B2.ev.report});
