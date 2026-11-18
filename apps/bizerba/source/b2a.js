/* ============================================================ BRAIN2 (PC software)
   One central database (B2.db) on the "server"; each line's terminal has its own copy (DBS[line]).
   Download = BRAIN2 -> device through _connect.BRAIN; Upload = device -> BRAIN2.
   UI follows the photo of BRAIN2 at the site (24/10/2026): Product data / Process data / Result data / Reporting,
   Product data modules Code Structure, Label Designer, Layout Management, Master Data, Master Data Exceptions, Text. */

/* ---------------------------------------------------------------- label layouts (shared by Label Designer and the terminal preview) */
function defaultLayouts(){
 const T=(id,name,x,y,w,h,text,o)=>Object.assign({id,name,type:"text",x,y,w,h,rot:0,text,font:"Arial",size:2.4,bold:false,halign:"Left",valign:"Top",frame:"No border",fill:"No filling",excess:"Yes",desc:"Yes"},o||{});
 const retail=(no,name,cust,hdr)=>({no,name,cust,w:80,h:80,ver:3,status:"Released",changed:"14/10/2026 09:12",by:"brain2admin",els:[
  T("e1","Text field 01",3,3,74,5,hdr,{size:3,bold:true,halign:"Center"}),
  T("e2","Text field 02",3,9,74,9,"{Article text}",{size:3.2,bold:true}),
  T("e3","Text field 03",3,20,30,3,"Use by",{size:2.2}),
  T("e4","Date field 01",33,19,44,5,"{Date 2}",{type:"date",size:4,bold:true,halign:"Right"}),
  {id:"e5",name:"Line 01",type:"line",x:3,y:26,w:74,h:0.4,rot:0},
  T("e6","Text field 04",3,28,30,3,"Net weight",{size:2.2}),
  T("e7","Weight field 01",40,27,37,5,"{Net weight}",{type:"weight",size:4,bold:true,halign:"Right"}),
  T("e8","Text field 05",3,34,40,3,"{Unit price}",{size:2.2}),
  T("e9","Price field 01",40,33,37,5,"{Price}",{type:"price",size:4,bold:true,halign:"Right"}),
  T("e10","Text field 06",3,41,74,9,"{Text field 1}",{size:2,excess:"Yes"}),
  {id:"e11",name:"Barcode field 01",type:"barcode",x:12,y:52,w:56,h:18,rot:0,code:"Code field 1"},
  T("e12","Text field 07",3,72,74,3,"{PLU} {Batch} {Cons. no.}",{size:1.8}),
  T("e13","Text field 08",3,75.5,74,3,"Keep refrigerated. {Text field 5}",{size:1.8})]});
 const L={
  "205":{no:205,name:"205 M&S NFO 2",cust:"M&S",w:70,h:50,ver:7,status:"Released",changed:"02/10/2026 14:40",by:"brain2admin",els:[
   T("e1","Text field 01",22,2,40,3,"COOKS IN:",{size:1.9,halign:"Center"}),
   T("e2","Text field 02",14,5.5,52,4,"{Text field 1 short}",{size:2.2,halign:"Center"}),
   T("e3","Text field 03",4,13,14,3,"USE BY",{size:1.8}),
   T("e4","Date field 01",3,16,42,13,"{Date 2}",{type:"date",font:"Arial Narrow",size:11}),
   {id:"e5",name:"Logo field 01",type:"logo",x:50,y:15,w:13,h:10,rot:0,logo:"Logo field 1"},
   T("e6","Text field 04",48,26,17,2.5,"Not for EU",{size:1.6,halign:"Center"}),
   T("e7","Text field 05",48,29,17,3.5,"SC1606",{size:2.2,bold:true,halign:"Center"}),
   T("e8","Text field 06",4,38,60,3,"{Shelf-life 1}   ⧈{Time}   {Text field 6} {Day of year}",{size:1.9})]},
  "3":{no:3,name:"3 Date label 58x40",cust:"",w:58,h:40,ver:2,status:"Released",changed:"11/09/2026 10:02",by:"brain2admin",els:[
   T("e1","Text field 01",3,3,52,3,"USE BY",{size:1.8}),
   T("e2","Date field 01",3,7,52,10,"{Date 2}",{type:"date",font:"Arial Narrow",size:8.5}),
   T("e3","Text field 02",3,20,52,3,"{PLU} L{Line} {Cons. no.} {Time}",{size:1.7}),
   T("e4","Weight field 01",30,24,25,4,"{Net weight}",{type:"weight",size:2.2,halign:"Right"}),
   T("e5","Text field 03",3,29,52,9,"{Text field 1}",{size:1.6})]},
  "4":retail(4,"4 Retail 80x80","", "{Customer}"),
  "173":retail(173,"173 CO-OP","Co-op","Co-op"),
  "174":retail(174,"174 ASDA","Asda","ASDA"),
  "175":retail(175,"175 ASDA","Asda","ASDA Extra Special"),
  "176":retail(176,"176 ICELAND","Iceland","ICELAND"),
  "177":retail(177,"177 ICELAND","Iceland","ICELAND Luxury"),
  "178":retail(178,"178 Asda","Asda","Asda Just Essentials")};
 L["175"].ver=2;L["178"].status="Draft";L["178"].changed="23/10/2026 16:05";L["178"].by="t.services";
 Object.assign(L,photoLayouts());
 return L}

/* ---------------------------------------------------------------- layouts made the BRAIN2 way (photos 27/10/2026)
   Every element is a FIELD with a field type, chosen in Label Designer (right click › Change field type):
   - fixed text, price, weight, unit price, date, clock, device number, logo, code: set up in BRAIN2
   - Text field 01–10: only a slot; what prints comes from the article's General text no. 1–10 in Data Maintenance
   Units are mm here, shown in Dots (8 dots/mm) in the Property window like the photo. */
const FONTS={"Font 001":1.2,"Font 002":1.5,"Font 003":2,"Font 004":1.5,"Font 005":2.5,"Font 006":3,"Font 007":4,"Font 008":5,"Font 009":6.5,"Font 010":8};
const FLD_TYPES=[["tf","Text field"],["multi","Multi"],["simple","Simple text"],["static","Static text"],["fixed","Fixed text"],["date","Date field"],["clock","Clock field"],["weight","Weight field"],["price","Price field"],["unit","Unit price field"],["devno","Device number field"],["code","Code field"],["logo","Logo field"]];
const SIMPLE_DATA=["GDA: Energy in kcal","GDA: Energy in kJ","Article text","PLU number","Producer","Best before temperature","Portion size","Line number"];
function fldName(e){const NN=String(e.n||1).padStart(2,"0");return{tf:"Text field "+NN,simple:"Simple text "+NN+" ("+(e.data||"GDA: Energy in kcal")+")",static:"Static text "+NN,fixed:"Text "+NN,multi:"Multi field "+NN,date:"Date field "+NN,clock:"Clock "+NN,weight:"Weight field "+NN,price:"Price field "+NN,unit:"Unit price "+NN,devno:"Device number "+NN,code:"Code "+NN,logo:"Logo "+NN}[e.fld]||e.name}
function photoLayouts(){
 const F=(fld,n,x,y,w,h,o)=>{const e=Object.assign({id:"f"+Math.random().toString(36).slice(2,8),fld,n,type:"field",x,y,w,h,rot:0,font:"Font 004",size:1.5,bold:false,halign:"Center",valign:"Top",frame:"No border",fill:"No Fill",excess:"Yes",desc:"Yes",autoline:"Yes",lspace:"Auto"},o||{});e.name=fldName(e);return e};
 let id=0;const fix=L=>{L.els.forEach(e=>e.id="e"+(++id));return L};
 const L={};
 /* 077 77-test: the test label made on 27/10/2026 (photos of Label Designer and of the printed labels) */
 L["77"]=fix({no:77,name:"77-test",cust:"",w:65,h:265,dpm:8,dist:10/8,ext:"No",ver:4,status:"Released",changed:"27/10/2026 12:03",by:"Admin",photo:true,els:[
  F("unit",1,12,10,10,4,{size:2.4,font:"Font 005"}),
  F("static",2,12,16,8,4.6,{size:1.4}),
  F("simple",2,12,22,8,3.4,{size:1.3,data:"GDA: Energy in kcal",font:"Font 002"}),
  F("tf",1,12,27,8,3.6,{size:1.1,font:"Font 001"}),
  F("price",1,12,32,21,7,{size:5,bold:true,font:"Font 008",halign:"Left"}),
  F("weight",1,12,40,17,5.5,{size:3.4,font:"Font 007",halign:"Left"}),
  F("devno",1,12,47,8,5,{size:3.8,bold:true,font:"Font 007",halign:"Left"}),
  F("logo",1,11,54,11,9),
  F("date",2,12,65,9,3.6,{size:1.7,font:"Font 003",fmt:"DDMMM"}),
  F("fixed",1,12,71,7,5,{size:1.7,bold:true,font:"Font 003",text:"TEST\n1"}),
  F("date",1,12,79,9,3.6,{size:1.7,font:"Font 003",fmt:"Julian"}),
  F("clock",1,12,86,13,4.6,{size:2.8,font:"Font 006",halign:"Left"}),
  F("code",1,12,94,28,13),
  ...Array.from({length:10},(_,i)=>F("tf",i+1,12,110+i*8,9,3.3,{size:1.45}))]});
 /* 205 M&S NFO 2 (FP2 whole birds). Code line rebuilt like Victor's Co-op label: Device number, Clock 1, Text field 6, Date 1 (Julian).
    The cooking time is a Multi field filled from the weight class / Text field 1 ("COOKS IN:" above it). Positions from the M&S label photo (approximate). */
 L["205"]=fix({no:205,name:"205 M&S NFO 2",cust:"M&S",w:70,h:50,dpm:8,dist:10/8,ext:"No",ver:7,status:"Released",changed:"02/10/2026 14:40",by:"brain2admin",els:[
  F("fixed",1,22,2,40,3,{size:1.9,text:"COOKS IN:"}),
  F("multi",2,14,5.5,52,4,{size:2.2,text:"{Text field 1 short}"}),
  F("fixed",3,4,13,14,3,{size:1.8,text:"USE BY",halign:"Left"}),
  F("date",2,3,16,44,13,{size:10.5,font:"Font 010",fmt:"DDMMM",halign:"Left"}),
  F("logo",1,50,15,13,10),
  F("fixed",4,48,26,17,2.5,{size:1.6,text:"Not for EU"}),
  F("fixed",5,48,29,17,3.5,{size:2.2,bold:true,text:"SC1606"}),
  F("devno",1,4,38,6,3,{size:1.9,halign:"Left"}),F("clock",1,10,38,11,3,{size:1.9,halign:"Left"}),F("tf",6,22,38,4,3,{size:1.9,halign:"Left"}),F("date",1,27,38,8,3,{size:1.9,fmt:"Julian",halign:"Left"})]});
 /* 173 Co-op: Victor's annotated label. The artwork is pre-printed on the label stock; the Bizerba prints only these fields.
    Code line = Device number, Clock 1, Text field 6, Date 1 (Julian, day of year). Positions measured off the photo (approximate). */
 L["173"]=fix({no:173,name:"173 CO-OP",cust:"Co-op",w:60,h:80,dpm:8,dist:10/8,ext:"No",ver:5,status:"Released",changed:"14/10/2026 09:12",by:"brain2admin",preprint:"Co-op",els:[
  F("price",1,2,30,18,5,{size:3.4,bold:true,halign:"Left"}),
  F("unit",1,21,30,18,5,{size:3.4,bold:true,halign:"Left",cur:true}),
  F("date",2,41,29,17,5,{size:4,bold:true,font:"Font 007",fmt:"DDMMM"}),
  F("weight",1,41,37,13,6,{size:4.4,bold:true,font:"Font 008",halign:"Left"}),
  F("fixed",1,54,38.5,4,4,{size:3,text:"℮",bold:true}),
  F("devno",1,3,41,6,2.6,{size:1.7,halign:"Left"}),F("clock",1,9,41,9,2.6,{size:1.7,halign:"Left"}),F("tf",6,19,41,3,2.6,{size:1.7,halign:"Left"}),F("date",1,23,41,6,2.6,{size:1.7,fmt:"Julian",halign:"Left"}),
  F("code",1,3,44,34,14),
  F("logo",1,42,45,14,8)]});
 /* the other layouts in the photographed Project view; contents not photographed (assumed: retail fields) */
 const R=(no,name,cust)=>fix({no,name,cust,w:80,h:80,dpm:8,dist:10/8,ext:"No",ver:2,status:"Released",changed:"02/10/2026 14:40",by:"brain2admin",assumed:true,els:[
  F("fixed",1,3,3,74,5,{size:3,bold:true,text:cust.toUpperCase()}),F("simple",1,3,9,74,8,{size:3,bold:true,data:"Article text",halign:"Left"}),
  F("fixed",2,3,20,30,3,{size:2,text:"Use by",halign:"Left"}),F("date",2,33,19,44,5,{size:4,bold:true,fmt:"DDMMM",halign:"Right"}),
  F("weight",1,40,27,37,5,{size:4,bold:true,halign:"Right"}),F("price",1,40,33,37,5,{size:4,bold:true,halign:"Right"}),F("unit",1,3,34,30,4,{size:2.2,halign:"Left",cur:true}),
  F("tf",1,3,41,74,9,{size:2,halign:"Left"}),F("code",1,12,52,56,16),
  F("devno",1,3,72,8,3,{size:1.8,halign:"Left"}),F("clock",1,11,72,12,3,{size:1.8,halign:"Left"}),F("date",1,24,72,10,3,{size:1.8,fmt:"Julian",halign:"Left"}),F("logo",1,62,70,14,9)]});
 L["83"]=R(83,"83 customer label","");L["82"]=R(82,"82 booths ac","Booths");L["85"]=R(85,"85 OCADO new","Ocado");L["87"]=R(87,"87- M&S BDF PZ C WR","M&S");L["91"]=R(91,"91 M&S EVENT","M&S");L["92"]=R(92,"92 M&S Truffle","M&S");
 return L}
const LOGO_ART={"GB-2037.png":{t:"GB|2037",shape:"circle"},"GB-1606.png":{t:"GB|1606",shape:"oval"},"UK2037EC.bmp":{t:"UK|2037|EC",shape:"oval"},"TRACTOR.bmp":{pic:"tractor"}};
const TRACTOR_SVG=`<svg viewBox="0 0 40 30" width="100%" height="100%"><rect x="13" y="4" width="12" height="11" fill="#111"/><rect x="15" y="6" width="8" height="6" fill="#fff"/><rect x="6" y="13" width="28" height="7" fill="#111"/><circle cx="11" cy="23" r="6" fill="#111"/><circle cx="11" cy="23" r="2.2" fill="#fff"/><circle cx="31" cy="24" r="4.5" fill="#111"/><circle cx="31" cy="24" r="1.6" fill="#fff"/><rect x="30" y="7" width="2" height="7" fill="#111"/></svg>`;
const MON=["JAN","FEB","MAR","APR","MAY","JUN","JUL","AUG","SEP","OCT","NOV","DEC"];
const julian=d=>Math.floor((d-new Date(d.getFullYear(),0,0))/864e5);
function fmtDate(d,f){if(!d)return"";if(f==="DDMMM")return String(d.getDate()).padStart(2,"0")+MON[d.getMonth()];if(f==="Julian")return String(julian(d)).padStart(3,"0");if(f==="DD.MM.YY")return String(d.getDate()).padStart(2,"0")+"."+String(d.getMonth()+1).padStart(2,"0")+"."+String(d.getFullYear()).slice(2);return fmtD(d)}
/* device number printed by the Device number field (FP2 numbers assumed: the 27/10 test label printed 11) */
const DEVNO={"FP2-L03":11,"FP2-L04":12,"FP2-L05":15,"FP2-L06":16,"FP1-L03":334,"FP1-CW":21};
/* tokens a layout can use, resolved from an article and (on the terminal) a pack record */
function layoutCtx(p,rec,lineKey,db){db=db||DB;const w=rec?rec.w:(p.mode==="Fixed weight"?p.fixw/1000:(p.mean||1000)/1000);
 const d1=rec?rec.d1:new Date();const d2=rec?rec.d2:addDays(d1,p.life1||0);const price=rec?rec.price:(p.mode==="Weight"?Math.round(w*p.unit*100)/100:p.mode==="Fixed weight"?Math.round(p.fixw/1000*p.unit*100)/100:p.fixp);
 const t1=rec?rec.t1:(p.txt1>0?(db.text[p.txt1]||""):"");const doy=julian(d1);
 const ln=(lineKey||"").slice(-2);const unitV=price>0&&w>0&&p.mode!=="Weight"?price/w:p.unit;
 return{"Article text":p.text,"Customer":(p.cust||"").toUpperCase(),"Text field 1":t1,"Text field 1 short":shortT(t1)==="–"?"":shortT(t1),"Date 1":fmtD(d1),"Date 2":fmtD(d2),"Date 3":fmtD(addDays(d1,p.life2||0)),"Shelf-life 1":String(rec&&rec.life!=null?rec.life:p.life1),
  "Net weight":p.mode==="Fixed weight"?p.fixw+"g ℮":f3(w)+" kg","Unit price":p.unit>0?"£"+f2(p.unit)+"/kg":"","Price":price>0?"£"+f2(price):"","Tare":p.tare+" g","PLU":pad6(p.plu),
  "Batch":rec?rec.batch:"L"+ln+fmtD(d1).replace(/\//g,""),"Cons. no.":rec?String(rec.c).slice(-6):"000000","Time":(rec&&rec.tm)||new Date().toTimeString().slice(0,5),"Day of year":String(doy),"Line":ln,
  _d1:p.d1==="Without"?null:d1,_d2:d2,_w:w,_price:price,_unit:unitV,_devno:String(DEVNO[lineKey]||DEVNO["FP2-L03"]),_fixedW:p.mode==="Fixed weight",
  ...Object.fromEntries([2,3,4,5,6,7,8,9,10].map(i=>["Text field "+i,p["tf"+i]||""]))}}
function fillTokens(s,ctx){return String(s||"").replace(/\{([^}]+)\}/g,(m,k)=>ctx[k]!=null?ctx[k]:m)}
/* code structures (GGCAB): segments [type, value, digits] */
function codeDigits(p,ctx,db,n){n=n||1;const cs=db.code[p["cf"+n]];if(!cs||!p["cf"+n])return null;let s="";
 for(const [t,v,len] of cs.segs){if(t==="Constant")s+=String(v).padStart(len,"0").slice(-len);else if(/^Code substring (\d)/.test(t)){const i=t.match(/(\d)/)[1];s+=String(p["cs"+i]||"0").padStart(len,"0").slice(-len)}
  else if(t==="PLU number")s+=String(p.plu).padStart(len,"0").slice(-len);else if(t==="Price"){const pr=String(Math.round((ctx._price!=null?ctx._price:(parseFloat((ctx.Price||"0").replace(/[^\d.]/g,""))||0))*100));s+=pr.padStart(len,"0").slice(-len)}
  else if(t==="Net weight"){s+=String(Math.round((ctx._w||0)*1000)).padStart(len,"0").slice(-len)}
  else if(t==="Price check digit"){s+="#"}else if(t==="Check digit"){}else if(/^AI/.test(t)){return null}}
 s=s.replace("#",()=>priceCheck4(s.slice(-4).padStart(4,"0")));
 if(cs.type==="EAN13"){s=s.slice(0,12).padEnd(12,"0");return s+ean13check(s)}return null}
function qrSvg(txt,px){let h=0;for(const c of String(txt))h=(h*31+c.charCodeAt(0))>>>0;const n=21;let r="";const rnd=()=>{h^=h<<13;h^=h>>>17;h^=h<<5;return(h>>>0)/4294967296};
 const fin=(x,y)=>`<rect x="${x}" y="${y}" width="7" height="7" fill="#000"/><rect x="${x+1}" y="${y+1}" width="5" height="5" fill="#fff"/><rect x="${x+2}" y="${y+2}" width="3" height="3" fill="#000"/>`;
 for(let y=0;y<n;y++)for(let x=0;x<n;x++){if((x<8&&y<8)||(x>12&&y<8)||(x<8&&y>12))continue;if(rnd()<.5)r+=`<rect x="${x}" y="${y}" width="1" height="1" fill="#000"/>`}
 return`<svg viewBox="0 0 21 21" width="${px}" height="${px}" shape-rendering="crispEdges"><rect width="21" height="21" fill="#fff"/>${r}${fin(0,0)}${fin(14,0)}${fin(0,14)}</svg>`}
const CLOCK_GLYPH=`<span class="clk">⊡</span>`;
/* what a field prints (design = the placeholder Label Designer shows, like the photo) */
function fieldText(e,p,ctx,design,db){const NN=String(e.n||1).padStart(2,"0");
 if(design){return{tf:"Text field\n"+NN,static:"Static text\n"+NN,simple:NN+" ("+(e.data||"")+")",fixed:e.text||"",multi:e.text||"",date:e.n===2?"23.09.76":"14.12.01",clock:"CLock "+NN,weight:"99,999 kg",price:"99,99 €",unit:"99.99",devno:"777"}[e.fld]??""}
 switch(e.fld){
  case"tf":{const v=e.n===1?ctx["Text field 1"]||"":ctx["Text field "+e.n]||"";return/^=\$/.test(v)?"":v} /* "=$…$" texts are commands for the terminal (e.g. the GL0A use-by check), not printed */
  case"static":return(db.stattxt&&db.stattxt[e.n])||"";
  case"simple":return e.data==="Article text"?p.text:e.data==="PLU number"?String(p.plu):e.data==="Line number"?ctx.Line:/^GDA/.test(e.data||"")?(p.nut&&p.nut.use&&p.nut.kcal?String(e.data.includes("kJ")?p.nut.kj:p.nut.kcal):""):"";
  case"fixed":return e.text||"";
  case"multi":return fillTokens(e.text,ctx);
  case"date":{const d=e.n===1?ctx._d1:e.n===2?ctx._d2:null;return d?fmtDate(d,e.fmt||(e.n===1?"Julian":"DDMMM")):""}
  case"clock":return"\u0001"+ctx.Time;
  case"weight":return ctx._fixedW?ctx["Net weight"]:f3(ctx._w)+"kg";
  case"price":return ctx._price>0?"£"+f2(ctx._price):"£";
  case"unit":return ctx._unit>0?(e.cur?"£":"")+f2(ctx._unit):(e.cur?"£":"");
  case"devno":return ctx._devno}
 return""}
/* render a layout at k px per mm; sel = selected element id (designer only) */
function renderLayout(L,p,ctx,k,opt){opt=opt||{};const db=opt.db||DB;let h=`<div class="lyt${L.preprint?" pre":""}" style="width:${L.w*k}px;height:${L.h*k}px${opt.rot180?";transform:rotate(180deg)":""}">`;
 if(L.preprint&&!opt.design)h+=`<div class="prenote" style="font-size:${1.6*k}px">pre-printed ${esc(L.preprint)} stock</div>`;
 L.els.forEach(e=>{const st=`left:${e.x*k}px;top:${e.y*k}px;width:${Math.max(1,e.w*k)}px;height:${Math.max(1,e.h*k)}px;${e.rot?`transform:rotate(${e.rot}deg);transform-origin:0 0;`:""}`;
  const cls=`lel ${e.fld?"fld":""} ${opt.sel===e.id?"sel":""} ${e.frame&&e.frame!=="No border"?"fr"+(e.frame==="Thick"?"2":""):""} ${e.fill==="Black"?"fb":e.fill==="Grey"?"fg":""}`;
  let inner="";
  if(e.fld){
   if(e.fld==="logo"){const n=p["logo"+(e.n||1)];const nm=db.logo&&db.logo[n];const art=nm&&LOGO_ART[nm];
    if(opt.design&&!opt.data)inner=`<div class="lgph"><span style="font-size:${Math.min(e.w,e.h)*.55*k}px">0</span></div>`;
    else if(art&&art.pic)inner=TRACTOR_SVG;
    else if(art)inner=`<div class="oval ${art.shape}" style="border-width:${Math.max(1,.35*k)}px;font-size:${(art.t.split("|").length>2?2:2.6)*k}px">${art.t.split("|").join("<br>")}</div>`;
    else if(nm)inner=`<div class="lph" style="font-size:${1.6*k}px">${esc(nm)}</div>`}
   else if(e.fld==="code"){if(opt.design&&!opt.data)inner=`<div class="bcw">${eanSvg("1111111111116",e.w*k)}</div>`;else{const d=codeDigits(p,ctx,db,e.n||1);inner=d?`<div class="bcw">${eanSvg(d,e.w*k)}</div>`:""}}
   else{let t=fieldText(e,p,ctx,opt.design&&!opt.data,db);let fs=(e.size||1.5)*k;if(e.fld==="tf"&&!(opt.design&&!opt.data)&&(p.sendfmt==="Yes"||(db.textfmt&&p.gtn&&db.textfmt[p.gtn[e.n]]==="Yes")))fs*=.72;const isClock=t.startsWith("\u0001");if(isClock)t=t.slice(1);
    const ja={Left:"flex-start",Center:"center",Right:"flex-end"}[e.halign||"Left"],va={Top:"flex-start",Middle:"center",Bottom:"flex-end"}[e.valign||"Top"];
    inner=`<div class="ftx" style="font-size:${fs}px;font-weight:${e.bold||/^(price|weight|devno|clock)$/.test(e.fld)?700:600};justify-content:${ja};align-items:${va};text-align:${(e.halign||"Left").toLowerCase()};${opt.design&&!opt.data?"white-space:pre;":e.autoline==="No"||/^(price|weight|unit|devno|clock|date)$/.test(e.fld)?"white-space:pre;":""}${e.fill==="Black"?"color:#fff;":""}"><span>${isClock?CLOCK_GLYPH:""}${esc(t).replace(/\n/g,"<br>")}</span></div>`}}
  else if(e.type==="line")inner=`<div style="position:absolute;inset:0;background:#111"></div>`;
  else if(e.type==="rect")inner="";
  else if(e.type==="ellipse")inner=`<div style="position:absolute;inset:0;border:${Math.max(1,.3*k)}px solid #111;border-radius:50%"></div>`;
  else if(e.type==="logo"){const n=p[e.logo==="Logo field 2"?"logo2":e.logo==="Logo field 3"?"logo3":"logo1"];const nm=db.logo&&db.logo[n];const art=nm&&LOGO_ART[nm];
   inner=art&&art.pic?TRACTOR_SVG:art?`<div class="oval" style="border-width:${Math.max(1,.35*k)}px;font-size:${2.6*k}px">${art.t.split("|").join("<br>")}</div>`:nm?`<div class="lph" style="font-size:${1.6*k}px">${esc(nm)}</div>`:(opt.design?`<div class="lph" style="font-size:${1.6*k}px">${esc(e.logo||"Logo")}</div>`:"")}
  else if(e.type==="barcode"){const d=codeDigits(p,ctx,db,1);inner=d?`<div style="position:absolute;inset:0;display:flex;align-items:center;justify-content:center">${eanSvg(d,e.w*k)}</div>`:(p.code&&p.code.v?`<div class="lph" style="font-size:${1.8*k}px">${esc(p.code.v)}</div>`:opt.design?`<div class="lph" style="font-size:${1.6*k}px">${esc(e.code||"Code field 1")} · no code structure</div>`:"")}
  else if(e.type==="qr")inner=`<div style="position:absolute;inset:0;display:flex;align-items:center;justify-content:center">${qrSvg(fillTokens(e.text||"{PLU}",ctx),Math.min(e.w,e.h)*k)}</div>`;
  else{const t=fillTokens(e.text,ctx);const cond=e.font==="Arial Narrow";
   inner=`<div class="ltx" style="font-size:${e.size*k}px;font-weight:${e.bold?700:400};text-align:${(e.halign||"Left").toLowerCase()};${cond?"font-family:'Arial Narrow','Roboto Condensed',Arial,sans-serif;letter-spacing:-.04em;transform:scaleX(.8);transform-origin:left top;width:125%;":""}${e.excess==="No"?"white-space:nowrap;overflow:hidden;":""}${e.fill==="Black"?"color:#fff;":""}">${esc(t).replace(/\n/g,"<br>")}</div>`}
  h+=`<div class="${cls}" ${opt.design?`data-lel="${e.id}"`:""} style="${st}">${inner}${opt.design&&opt.sel===e.id?'<i class="hd tl"></i><i class="hd tr"></i><i class="hd bl"></i><i class="hd br" data-lres="'+e.id+'"></i>':""}</div>`});
 return h+`</div>`}
