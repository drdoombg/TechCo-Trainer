/* ---------------------------------------------------------------- Label Designer (photos 27/10/2026: BRAIN2 › Product data › Label Designer)
   Menu File Edit View Elements Properties Window, three toolbar rows, Project view / Label window dock, Property window
   (Categorized, Alphabetical, mm, Dot), document tab, grey work area with the print legend, Text field editor, status bar.
   Toolbar icons are drawn from the photos; where a photo is blurred the tooltip says what the icon is taken to be. */
const LD_MENUS={File:["New layout","Open layout…","Save","Save as…","Close","—","Print test label","Export to…","—","Project properties"],
 Edit:["Undo","Redo","—","Cut","Copy","Paste","Delete","—","Select all"],
 View:["Zoom in","Zoom out","Fit to window","Zoom 100 %","—","Project view","Label window","Property window","Text field editor","—","Grid","Show article data (preview)"],
 Elements:["Text field","Simple text","Static text","Fixed text","—","Date field","Clock field","Weight field","Price field","Unit price field","Device number field","—","Code field","Logo field","—","Line","Rectangle","Ellipse"],
 Properties:["Label properties","Field properties","—","Change field type…"],Window:["Close all documents","—","Next document"]};
const FLD_OF={"Text field":"tf","Simple text":"simple","Static text":"static","Fixed text":"fixed","Date field":"date","Clock field":"clock","Weight field":"weight","Price field":"price","Unit price field":"unit","Device number field":"devno","Code field":"code","Logo field":"logo"};
/* convert the older layouts (token texts) into fields, so every layout behaves the BRAIN2 way */
function toFields(L){if(!L||L._f)return L;L._f=true;L.dpm=L.dpm||8;L.dist=L.dist??1.25;L.ext=L.ext||"No";let fx=0;
 L.els=L.els.map(e=>{if(e.fld||/^(line|rect|ellipse)$/.test(e.type))return e;const b={x:e.x,y:e.y,w:e.w,h:e.h,rot:e.rot||0,id:e.id,font:"Font 004",size:e.size||2,bold:!!e.bold,halign:e.halign||"Left",valign:"Top",frame:e.frame==="Thin"?"Thin":e.frame==="Thick"?"Thick":"No border",fill:e.fill==="Black"?"Black":e.fill==="Grey"?"Grey":"No Fill",excess:e.excess||"Yes",desc:e.desc||"Yes",autoline:"Yes",lspace:"Auto"};
  let f;const t=(e.text||"").trim();let m;
  if(e.type==="barcode")f={fld:"code",n:Number((e.code||"1").replace(/\D/g,""))||1};
  else if(e.type==="logo")f={fld:"logo",n:Number((e.logo||"1").replace(/\D/g,""))||1};
  else if(t==="{Date 2}")f={fld:"date",n:2,fmt:"DD.MM.YY"};else if(t==="{Date 1}")f={fld:"date",n:1,fmt:"DD.MM.YY"};
  else if(t==="{Net weight}")f={fld:"weight",n:1};else if(t==="{Price}")f={fld:"price",n:1};else if(t==="{Unit price}")f={fld:"unit",n:1,cur:true};
  else if(t==="{Article text}")f={fld:"simple",n:1,data:"Article text"};
  else if((m=t.match(/^\{Text field (\d+)\}$/)))f={fld:"tf",n:Number(m[1])};
  else if(/\{/.test(t))f={fld:"multi",n:++fx,text:e.text||""};
  else f={fld:"fixed",n:++fx,text:e.text||""};
  const o=Object.assign(b,f);o.name=fldName(o);if(o.fld==="date")o.fmt=o.n===1?"Julian":"DDMMM";return o});return L}
function ldLayout(){const L=B2.db.layouts[B2.ld.cur];return toFields(L)}
function ldEl(){const L=ldLayout();return L&&L.els.find(e=>e.id===B2.ld.sel)}
const pad3=n=>String(n).padStart(3,"0");
const ldTitle=L=>pad3(L.no)+"."+L.name;
const ldDoc=L=>pad3(L.no)+"-"+L.name;
/* ---- toolbar icons (16 px) */
const IC={
 newd:'<path d="M3 1h7l3 3v11H3z" fill="#fff" stroke="#556"/><path d="M10 1v3h3" fill="none" stroke="#556"/><rect x="1" y="9" width="7" height="6" fill="#f0c040" stroke="#8a6d10"/>',
 open:'<path d="M1 4h5l1 2h7v8H1z" fill="#e8c35a" stroke="#8a6d10"/><rect x="5" y="1" width="6" height="7" fill="#fff" stroke="#667"/>',
 save:'<rect x="2" y="2" width="12" height="12" fill="#6b7d99" stroke="#3a4656"/><rect x="4" y="2" width="8" height="5" fill="#fff"/><rect x="5" y="10" width="6" height="4" fill="#cfd7e3"/>',
 saveall:'<rect x="1" y="1" width="10" height="10" fill="#8796ad" stroke="#3a4656"/><rect x="5" y="5" width="10" height="10" fill="#6b7d99" stroke="#3a4656"/><rect x="7" y="5" width="6" height="4" fill="#fff"/>',
 page:'<path d="M3 1h7l3 3v11H3z" fill="#fff" stroke="#556"/><path d="M10 1v3h3" fill="none" stroke="#556"/>',
 print:'<rect x="4" y="1" width="8" height="5" fill="#fff" stroke="#556"/><rect x="1" y="6" width="14" height="6" fill="#8a8f96" stroke="#444"/><rect x="4" y="10" width="8" height="5" fill="#fff" stroke="#556"/>',
 export:'<path d="M3 1h6l3 3v6H3z" fill="#fff" stroke="#556"/><circle cx="11" cy="11" r="4.5" fill="#3a9b3a"/><path d="M8.5 11h4M11 8.5l2.5 2.5L11 13.5" stroke="#fff" stroke-width="1.5" fill="none"/>',
 cut:'<circle cx="4.5" cy="12" r="2.5" fill="none" stroke="#2a52a0" stroke-width="1.6"/><circle cx="11.5" cy="12" r="2.5" fill="none" stroke="#2a52a0" stroke-width="1.6"/><path d="M6 10L11 1M10 10L5 1" stroke="#555" stroke-width="1.3"/>',
 copy:'<rect x="1" y="1" width="9" height="10" fill="#fff" stroke="#556"/><rect x="6" y="5" width="9" height="10" fill="#fff" stroke="#556"/><path d="M8 8h5M8 10h5M8 12h4" stroke="#99a"/>',
 paste:'<rect x="2" y="2" width="10" height="13" fill="#c99a4a" stroke="#7a5a20"/><rect x="5" y="1" width="4" height="3" fill="#ddd" stroke="#666"/><rect x="7" y="7" width="8" height="8" fill="#fff" stroke="#556"/>',
 delx:'<rect x="1.5" y="1.5" width="13" height="13" fill="#fff" stroke="#223"/><path d="M3 3l10 10M13 3L3 13" stroke="#223" stroke-width="2"/>',
 undo:'<path d="M5 3L1 7l4 4" fill="none" stroke="#2a52a0" stroke-width="2"/><path d="M1.5 7H10a4 4 0 0 1 0 8H7" fill="none" stroke="#2a52a0" stroke-width="2"/>',
 redo:'<path d="M11 3l4 4-4 4" fill="none" stroke="#2a52a0" stroke-width="2"/><path d="M14.5 7H6a4 4 0 0 0 0 8h3" fill="none" stroke="#2a52a0" stroke-width="2"/>',
 lprop:'<rect x="2" y="1" width="12" height="14" fill="#fff" stroke="#556"/><path d="M4 4h8M4 7h8M4 10h5" stroke="#556"/>',
 grid:'<path d="M1 5h14M1 11h14M5 1v14M11 1v14" stroke="#556"/>',
 snap:'<rect x="1" y="1" width="14" height="14" fill="#fff" stroke="#556"/><path d="M4 8l3 3 5-6" fill="none" stroke="#2a52a0" stroke-width="2"/>',
 legend:'<rect x="1" y="1" width="14" height="14" fill="#fff" stroke="#556"/><rect x="3" y="3" width="4" height="10" fill="#e0c060"/><path d="M9 4h4M9 7h4M9 10h4" stroke="#556"/>',
 ruler:'<rect x="1" y="5" width="14" height="6" fill="#fff" stroke="#556"/><path d="M4 5v3M7 5v2M10 5v3M13 5v2" stroke="#556"/>',
 zin:'<circle cx="6.5" cy="6.5" r="5" fill="#fff" stroke="#2a52a0" stroke-width="1.5"/><path d="M4 6.5h5M6.5 4v5" stroke="#2a52a0" stroke-width="1.5"/><path d="M10 10l5 5" stroke="#555" stroke-width="2"/>',
 zout:'<circle cx="6.5" cy="6.5" r="5" fill="#fff" stroke="#2a52a0" stroke-width="1.5"/><path d="M4 6.5h5" stroke="#2a52a0" stroke-width="1.5"/><path d="M10 10l5 5" stroke="#555" stroke-width="2"/>',
 zfit:'<rect x="1" y="2" width="14" height="12" fill="#fff" stroke="#556"/><rect x="4" y="5" width="8" height="6" fill="none" stroke="#2a52a0" stroke-dasharray="2 1"/>',
 z100:'<rect x="1" y="2" width="14" height="12" fill="#fff" stroke="#556"/><text x="8" y="11" font-size="6" text-anchor="middle" fill="#2a52a0" font-family="Arial" font-weight="700">100%</text>',
 screen:'<rect x="1" y="2" width="14" height="10" fill="#cfd7e3" stroke="#556"/><path d="M5 14h6" stroke="#556" stroke-width="2"/>',
 fr0:'<rect x="1.5" y="1.5" width="13" height="13" fill="#eaf2fb" stroke="#9ab" stroke-dasharray="2 2"/>',
 fr1:'<rect x="1.5" y="1.5" width="13" height="13" fill="#fff" stroke="#223"/>',
 frr:'<rect x="1.5" y="1.5" width="13" height="13" rx="4" fill="#fff" stroke="#223"/>',
 frd:'<rect x="1.5" y="1.5" width="13" height="13" fill="#fff" stroke="#223"/><rect x="4" y="4" width="8" height="8" fill="none" stroke="#223"/>',
 frt:'<rect x="2" y="2" width="12" height="12" fill="#fff" stroke="#223" stroke-width="2.5"/>',
 frs:'<rect x="4" y="4" width="11" height="11" fill="#223"/><rect x="1.5" y="1.5" width="11" height="11" fill="#fff" stroke="#223"/>',
 tN:'<text x="8" y="11" font-size="7" text-anchor="middle" font-family="Arial" font-weight="700" fill="#223">ABC</text>',
 tI:'<rect x="0" y="2" width="16" height="12" fill="#223"/><text x="8" y="11" font-size="7" text-anchor="middle" font-family="Arial" font-weight="700" fill="#fff">ABC</text>',
 tP:'<rect x="0" y="2" width="16" height="12" fill="#888" opacity=".6"/><text x="8" y="11" font-size="7" text-anchor="middle" font-family="Arial" font-weight="700" fill="#223">ABC</text>',
 tG:'<text x="8" y="11" font-size="7" text-anchor="middle" font-family="Arial" font-weight="700" fill="#999">ABC</text>',
 tX:'<text x="8" y="11" font-size="7" text-anchor="middle" font-family="Arial" font-weight="700" fill="#223">ABC</text><path d="M2 3l12 10M14 3L2 13" stroke="#223" stroke-width="1.6"/>',
 fB:'<rect x="1" y="1" width="14" height="14" fill="#222"/>',
 fD:'<rect x="1" y="1" width="14" height="14" fill="#555"/><path d="M2 2h1M5 2h1M8 2h1M11 2h1M3 5h1M6 5h1M9 5h1M12 5h1M2 8h1M5 8h1M8 8h1M11 8h1M3 11h1M6 11h1M9 11h1M12 11h1" stroke="#fff"/>',
 fM:'<rect x="1" y="1" width="14" height="14" fill="#999"/>',
 fL:'<rect x="1" y="1" width="14" height="14" fill="#ddd" stroke="#999"/>',
 fN:'<rect x="1.5" y="1.5" width="13" height="13" fill="#fff" stroke="#999"/>',
 pV:'<rect x="1.5" y="1.5" width="13" height="13" fill="#fff" stroke="#999"/><path d="M5 2v12M11 2v12" stroke="#556" stroke-dasharray="2 1"/>',
 pD:'<rect x="1.5" y="1.5" width="13" height="13" fill="#fff" stroke="#999"/><path d="M3 14L8 2M8 14L13 2" stroke="#556"/>',
 pH:'<rect x="1.5" y="1.5" width="13" height="13" fill="#fff" stroke="#999"/><path d="M3 2v12M13 2v12" stroke="#556"/>',
 pX:'<rect x="1.5" y="1.5" width="13" height="13" fill="#fff" stroke="#999"/><path d="M2 2l12 12M14 2L2 14" stroke="#556"/>',
 font:'<text x="2" y="13" font-size="13" font-family="Times New Roman" font-weight="700" fill="#2a52a0">T</text><text x="9" y="14" font-size="8" font-family="Times New Roman" fill="#555">T</text>',
 hL:'<rect x="1.5" y="1.5" width="13" height="13" fill="#fff" stroke="#556"/><path d="M3 4h9M3 7h6M3 10h9M3 13h5" stroke="#223"/>',
 hC:'<rect x="1.5" y="1.5" width="13" height="13" fill="#fff" stroke="#556"/><path d="M3.5 4h9M5 7h6M3.5 10h9M5.5 13h5" stroke="#223"/>',
 hR:'<rect x="1.5" y="1.5" width="13" height="13" fill="#fff" stroke="#556"/><path d="M4 4h9M7 7h6M4 10h9M8 13h5" stroke="#223"/>',
 vT:'<rect x="1.5" y="1.5" width="13" height="13" fill="#fff" stroke="#556"/><rect x="3" y="3" width="10" height="2.5" fill="#223"/>',
 vM:'<rect x="1.5" y="1.5" width="13" height="13" fill="#fff" stroke="#556"/><rect x="3" y="6.8" width="10" height="2.5" fill="#223"/>',
 vB:'<rect x="1.5" y="1.5" width="13" height="13" fill="#fff" stroke="#556"/><rect x="3" y="10.5" width="10" height="2.5" fill="#223"/>',
 fitT:'<rect x="1.5" y="1.5" width="13" height="13" fill="#fff" stroke="#556"/><path d="M4 8h8M4 8l2-2M4 8l2 2M12 8l-2-2M12 8l-2 2" stroke="#223"/>',
 fitH:'<rect x="1.5" y="1.5" width="13" height="13" fill="#fff" stroke="#556"/><rect x="4" y="4" width="8" height="8" fill="#cfd7e3" stroke="#223"/>',
 gear:'<circle cx="8" cy="8" r="4.5" fill="none" stroke="#556" stroke-width="2.5" stroke-dasharray="2.2 1.3"/><circle cx="8" cy="8" r="1.8" fill="#556"/>',
 lres:'<rect x="1" y="2" width="14" height="12" fill="#f3f0d8" stroke="#8a7a30"/><rect x="3" y="4" width="6" height="5" fill="#e3b33a"/>',
 link:'<path d="M2 9a3 3 0 0 1 3-3h3M14 7a3 3 0 0 1-3 3H8" fill="none" stroke="#555" stroke-width="2"/><path d="M5 8h6" stroke="#555" stroke-width="1.5"/>',
 bottle:'<rect x="5" y="1" width="6" height="3" fill="#8a5a20"/><rect x="3" y="4" width="10" height="11" rx="2" fill="#c98a30" stroke="#7a5a20"/>',
 table:'<rect x="1" y="2" width="14" height="12" fill="#fff" stroke="#556"/><path d="M1 6h14M1 10h14M6 2v12" stroke="#556"/><rect x="1" y="2" width="14" height="4" fill="#5a7ab0" opacity=".6"/>',
 bars:'<path d="M2 2v12M4 2v12M5 2v12M8 2v12M10 2v12M13 2v12M14 2v12" stroke="#222" stroke-width="1"/>',
 tfAb:'<rect x="0.5" y="2.5" width="15" height="11" fill="#e8f0fb" stroke="#2a52a0"/><text x="8" y="11.5" font-size="8.5" text-anchor="middle" font-family="Arial" font-weight="700" fill="#223">Ab</text>',
 stAb:'<rect x="0.5" y="2.5" width="15" height="11" fill="#fff" stroke="#999"/><text x="8" y="11.5" font-size="8.5" text-anchor="middle" font-family="Arial" fill="#777">Ab</text>',
 n123:'<rect x="0.5" y="2.5" width="15" height="11" fill="#e8f0fb" stroke="#2a52a0"/><text x="8" y="11.5" font-size="7.5" text-anchor="middle" font-family="Arial" font-weight="700" fill="#223">123</text>',
 p123:'<rect x="0.5" y="2.5" width="15" height="11" fill="#fbecec" stroke="#a03030"/><text x="8" y="11.5" font-size="7.5" text-anchor="middle" font-family="Arial" font-weight="700" fill="#a03030">123</text>',
 pic:'<rect x="1" y="2" width="14" height="12" fill="#cfe3f7" stroke="#556"/><path d="M1 13l5-5 3 3 2-2 4 4" fill="#4a8a3a"/><circle cx="11" cy="5.5" r="1.6" fill="#e0b020"/>',
 fix:'<rect x="0.5" y="3.5" width="15" height="9" fill="#fff" stroke="#556"/><text x="8" y="11" font-size="6.5" text-anchor="middle" font-family="Arial" font-weight="700" fill="#223">FIX</text>',
 tT:'<rect x="1.5" y="1.5" width="13" height="13" fill="#fff" stroke="#556"/><text x="8" y="12.5" font-size="11" text-anchor="middle" font-family="Times New Roman" font-weight="700" fill="#223">T</text>',
 graph:'<path d="M1 14h14M2 12l4-5 3 3 5-7" fill="none" stroke="#999" stroke-width="1.5"/>',
 front:'<rect x="1" y="1" width="9" height="9" fill="#d9a63a" stroke="#7a5a20"/><rect x="6" y="6" width="9" height="9" fill="#e9c66a" stroke="#7a5a20"/>',
 back:'<rect x="6" y="6" width="9" height="9" fill="#e9c66a" stroke="#7a5a20"/><rect x="1" y="1" width="9" height="9" fill="#d9a63a" stroke="#7a5a20"/>',
 fwd:'<rect x="1" y="1" width="7" height="7" fill="#d9a63a" stroke="#7a5a20"/><rect x="4" y="4" width="7" height="7" fill="#e9c66a" stroke="#7a5a20"/><rect x="8" y="8" width="7" height="7" fill="#bbb" stroke="#666"/>',
 bwd:'<rect x="8" y="8" width="7" height="7" fill="#bbb" stroke="#666"/><rect x="4" y="4" width="7" height="7" fill="#e9c66a" stroke="#7a5a20"/><rect x="1" y="1" width="7" height="7" fill="#d9a63a" stroke="#7a5a20"/>',
 aL:'<path d="M2 1v14" stroke="#223" stroke-width="1.5"/><rect x="3" y="3" width="10" height="3" fill="#556"/><rect x="3" y="9" width="6" height="3" fill="#556"/>',
 aR:'<path d="M14 1v14" stroke="#223" stroke-width="1.5"/><rect x="3" y="3" width="10" height="3" fill="#556"/><rect x="7" y="9" width="6" height="3" fill="#556"/>',
 aT:'<path d="M1 2h14" stroke="#223" stroke-width="1.5"/><rect x="3" y="3" width="3" height="10" fill="#556"/><rect x="9" y="3" width="3" height="6" fill="#556"/>',
 aB:'<path d="M1 14h14" stroke="#223" stroke-width="1.5"/><rect x="3" y="3" width="3" height="10" fill="#556"/><rect x="9" y="7" width="3" height="6" fill="#556"/>',
 dH:'<path d="M1 3v10M15 3v10" stroke="#223"/><rect x="4" y="5" width="3" height="6" fill="#556"/><rect x="9" y="5" width="3" height="6" fill="#556"/><path d="M2 8h1M13 8h1" stroke="#223"/>',
 cV:'<path d="M8 1v14" stroke="#223"/><rect x="3" y="3" width="10" height="3" fill="#556"/><rect x="5" y="10" width="6" height="3" fill="#556"/>',
 sW:'<path d="M1 3v10M15 3v10M2 8h12" stroke="#223"/><path d="M2 8l2-2M2 8l2 2M14 8l-2-2M14 8l-2 2" stroke="#223"/>',
 sH:'<path d="M3 1h10M3 15h10M8 2v12" stroke="#223"/><path d="M8 2l-2 2M8 2l2 2M8 14l-2-2M8 14l2-2" stroke="#223"/>',
 cH:'<path d="M1 8h14" stroke="#223"/><rect x="3" y="2" width="3" height="12" fill="#556"/><rect x="10" y="4" width="3" height="8" fill="#556"/>',
 green:'<path d="M3 10a5 5 0 1 1 3 4" fill="none" stroke="#3a9b3a" stroke-width="2"/><path d="M1 8l2 3 3-2" fill="none" stroke="#3a9b3a" stroke-width="1.6"/><circle cx="12" cy="4" r="2" fill="#8a5a20"/>',
 x:'<path d="M3 3l10 10M13 3L3 13" stroke="#223" stroke-width="2"/>',
 prop:'<rect x="1" y="3" width="11" height="10" fill="#fff" stroke="#556"/><path d="M3 6h7M3 9h5" stroke="#556"/><rect x="8" y="1" width="7" height="7" fill="#c99a4a" stroke="#7a5a20"/>',
 refresh:'<path d="M3 1h7l3 3v11H3z" fill="#fff" stroke="#556"/><path d="M5.5 10a2.5 2.5 0 1 0 1-3.5" fill="none" stroke="#2a52a0" stroke-width="1.4"/><path d="M6 5l.5 2-2 .3" fill="none" stroke="#2a52a0" stroke-width="1.2"/>',
 cat:'<rect x="1" y="1" width="14" height="14" fill="#fff" stroke="#556"/><path d="M3 4h2M4 3v2M3 10h2M4 9v2M7 4h6M7 6h5M7 10h6M7 12h5" stroke="#556"/>',
 az:'<text x="1" y="7" font-size="7" font-family="Arial" font-weight="700" fill="#223">A</text><text x="1" y="15" font-size="7" font-family="Arial" font-weight="700" fill="#223">Z</text><path d="M12 2v12M9.5 11.5L12 14l2.5-2.5" fill="none" stroke="#223" stroke-width="1.4"/>',
 mm:'<rect x="1" y="2" width="14" height="4" fill="#fff" stroke="#556"/><path d="M3 2v2M5 2v2M7 2v2M9 2v2M11 2v2M13 2v2" stroke="#556"/><text x="8" y="14" font-size="6.5" text-anchor="middle" font-family="Arial" font-weight="700" fill="#223">mm</text>',
 dot:'<rect x="1" y="2" width="14" height="4" fill="#fff" stroke="#556"/><path d="M3 2v2M5 2v2M7 2v2M9 2v2M11 2v2M13 2v2" stroke="#556"/><text x="8" y="14" font-size="6.5" text-anchor="middle" font-family="Arial" font-weight="700" fill="#223">Dot</text>'};
const ico16=n=>`<svg viewBox="0 0 16 16" width="16" height="16" aria-hidden="true">${IC[n]||""}</svg>`;
/* toolbar buttons: [icon, action, value, tooltip, pressed(fn), disabled(fn)] */
function ldToolbars(E,L){const D=B2.ld;const on=(k,v)=>E&&String(E[k])===String(v);const dis=!E;
 const b=(i,a,v,t,pr,ds)=>`<button class="ldi ${pr?"pr":""}" data-b="${a}" ${v!=null?`data-v="${esc(v)}"`:""} title="${esc(t)}" ${ds?"disabled":""}>${ico16(i)}</button>`;
 const dd=`<span class="ldd">▾</span>`,sep=`<i class="sep"></i>`,grip=`<i class="grip"></i>`;
 const r1=grip+b("newd","ldnew",null,"New layout")+dd+b("open","ldmi","File:Open layout…","Open layout")+b("save","ldsave",null,"Save (Ctrl+S)",false,!L)+b("saveall","ldsaveall",null,"Save all",false,!Object.values(D.dirty).some(Boolean))+sep+b("page","ldmi","File:Project properties","Project properties")+sep+b("print","ldtest",null,"Print test label",false,!L)+b("export","ldexport",null,"Export to… (send layouts to the devices)")+sep+
  b("cut","ldmi","Edit:Cut","Cut",false,dis)+b("copy","ldmi","Edit:Copy","Copy",false,dis)+b("paste","ldmi","Edit:Paste","Paste",false,!D.clip)+b("delx","lddel",null,"Delete",false,dis)+sep+b("undo","ldundo",null,"Undo",false,!D.undo.length)+b("redo","ldredo",null,"Redo",false,!(D.redo||[]).length)+dd+
  `${grip}<select id="ldElSel" class="ldsel" title="Field selection">${`<option value="">&lt;No selection&gt;</option>`+(L?L.els.map(e=>`<option value="${e.id}" ${e.id===D.sel?"selected":""}>${esc(e.name)}</option>`).join(""):"")}</select>`+
  grip+b("lprop","ldlprop",null,"Label properties",!E)+b("grid","ldgrid",null,"Grid",D.grid)+b("snap","ldsnap",null,"Snap to grid",D.snap)+b("legend","ldlegend",null,"Print legend (thermal bars, carrier tape, tolerance area)",D.legend)+sep+b("ruler","ldruler",null,"Rulers",D.rulers)+sep+b("zin","ldzin",null,"Zoom in")+b("zout","ldzout",null,"Zoom out")+b("zfit","ldzfit",null,"Fit to window")+b("z100","ldz100",null,"Zoom 100 %",D.zoom===4)+sep+b("screen","ldprev",null,"Show article data (preview)",D.data)+dd;
 const r2=grip+b("fr0","ldset","frame:No border","No border",on("frame","No border"),dis)+b("fr1","ldset","frame:Thin","Frame",on("frame","Thin"),dis)+b("frr","ldset","frame:Rounded","Rounded frame",on("frame","Rounded"),dis)+b("frd","ldset","frame:Double","Double frame",on("frame","Double"),dis)+b("frt","ldset","frame:Thick","Thick frame",on("frame","Thick"),dis)+b("frs","ldset","frame:Shadow","Shadow frame",on("frame","Shadow"),dis)+sep+
  b("tN","ldset","tstyle:Normal","Normal text",!E||!E.tstyle||E.tstyle==="Normal",dis)+b("tI","ldset","tstyle:Inverse","Inverse text",on("tstyle","Inverse"),dis)+b("tP","ldset","tstyle:Pattern","Text on pattern",on("tstyle","Pattern"),dis)+b("tG","ldset","tstyle:Grey","Grey text",on("tstyle","Grey"),dis)+b("tX","ldset","tstyle:Off","Do not print text (icon read from photo)",on("tstyle","Off"),dis)+dd+sep+
  b("fB","ldset","fill:Black","Filling black",on("fill","Black"),dis)+b("fD","ldset","fill:Dense","Filling dense",on("fill","Dense"),dis)+b("fM","ldset","fill:Grey","Filling grey",on("fill","Grey"),dis)+b("fL","ldset","fill:Light","Filling light",on("fill","Light"),dis)+b("fN","ldset","fill:No Fill","No Fill",on("fill","No Fill"),dis)+sep+
  b("pV","ldset","pattern:Vertical","Thermal bars vertical (icon read from photo)",false,dis)+b("pD","ldset","pattern:Diagonal","Diagonal pattern (icon read from photo)",false,dis)+b("pH","ldset","pattern:Columns","Columns (icon read from photo)",false,dis)+b("pX","ldset","pattern:Cross","Cross pattern (icon read from photo)",false,dis)+dd+sep+
  b("font","ldfont",null,"Font…",false,dis)+sep+b("hL","ldset","halign:Left","Align left",on("halign","Left"),dis)+b("hC","ldset","halign:Center","Center",on("halign","Center"),dis)+b("hR","ldset","halign:Right","Align right",on("halign","Right"),dis)+sep+
  b("vT","ldset","valign:Top","Top",on("valign","Top"),dis)+b("vM","ldset","valign:Middle","Middle",on("valign","Middle"),dis)+b("vB","ldset","valign:Bottom","Bottom",on("valign","Bottom"),dis)+dd+sep+
  b("fitT","ldfit","w","Fit width to content",false,dis)+b("fitH","ldfit","h","Fit height to content",false,dis)+b("gear","ldmi","Properties:Field properties","Field properties",false,dis)+sep+b("lprop","ldlprop",null,"Label properties")+b("table","ldmi","View:Label window","Label window (field list)");
 const r3=grip+b("lres","ldlprop",null,"Label resource")+sep+b("link","ldaddm","code","Code field (icon read from photo)")+b("bottle","ldaddm","logo","Logo field (icon read from photo)")+b("table","ldaddm","weight","Weight field (icon read from photo)")+b("bars","ldaddm","code","Code field")+b("tfAb","ldaddm","tf","Text field (content from Data Maintenance)")+b("stAb","ldaddm","simple","Simple text")+b("n123","ldaddm","devno","Number field: device number")+b("p123","ldaddm","price","Price field")+b("pic","ldaddm","logo","Logo field")+sep+b("fix","ldaddm","fixed","Fixed text")+b("tT","ldaddm","static","Static text")+sep+b("graph","ldaddm","line","Graphic (line)",false,false)+dd+sep+
  b("front","ldorder","front","Bring to front",false,dis)+b("back","ldorder","back","Send to back",false,dis)+b("fwd","ldorder","fwd","One forward",false,dis)+b("bwd","ldorder","bwd","One backward",false,dis)+dd+sep+
  b("aL","ldalg","L","Align left edges",false,dis)+b("aR","ldalg","R","Align right edges",false,dis)+b("aT","ldalg","T","Align top edges",false,dis)+b("aB","ldalg","B","Align bottom edges",false,dis)+sep+b("dH","ldalg","CH","Center on label horizontally",false,dis)+b("cV","ldalg","CV","Center on label vertically",false,dis)+sep+b("sW","ldalg","W","Full label width",false,dis)+b("sH","ldalg","H","Same height as text",false,dis)+b("cH","ldalg","M","Middle",false,dis)+dd;
 return`<div class="ldtb">${r1}</div><div class="ldtb">${r2}</div><div class="ldtb">${r3}</div>`}
/* ---- Property window */
const DOT=8;
function ldVal(mm){return B2.ld.unit==="Dot"?Math.round(mm*DOT):Number(mm).toFixed(1)}
function ldProps(E,L){const D=B2.ld;const coll=D.coll||(D.coll={Position:true,"Extended layout settings":true});
 const grp=(t,body)=>`<div class="pgc" data-b="ldcoll" data-v="${esc(t)}">${coll[t]?"▷":"◢"} ${esc(t)}</div>${coll[t]?"":body}`;
 const row=(k,v)=>`<div class="pgr"><span title="${esc(k)}">${esc(k)}</span>${v}</div>`;
 const sel=(k,opts,cur)=>`<select data-lp="${k}">${opts.map(o=>`<option ${String(cur??(E&&E[k]))===String(o)?"selected":""}>${esc(o)}</option>`).join("")}</select>`;
 const num=(k,mm)=>`<input data-lp="${k}" data-unit="1" value="${ldVal(mm)}">`;
 const tb=`<div class="pwtb"><button class="ldi ${D.alpha?"":"pr"}" data-b="ldalpha" data-v="0" title="Categorized">${ico16("cat")}</button><button class="ldi ${D.alpha?"pr":""}" data-b="ldalpha" data-v="1" title="Alphabetical">${ico16("az")}</button><i class="sep"></i><button class="ldi ${D.unit==="mm"?"pr":""}" data-b="ldunit" data-v="mm" title="Values in mm">${ico16("mm")}</button><button class="ldi ${D.unit==="Dot"?"pr":""}" data-b="ldunit" data-v="Dot" title="Values in Dots (8 dots/mm)">${ico16("dot")}</button></div>`;
 const U=D.unit==="Dot"?" Dot":" mm";let h;
 if(!E){if(!L)return tb+`<div class="ldpg"></div>`;
  const rows={Resource:[row("Type",`<span class="ro">Label</span>`),row("Name",`<input data-lp="L.name" value="${esc(L.name)}">`)],
   Properties:[row("Label width",`<input data-lp="L.w" data-unit="1" value="${ldVal(L.w)}"><em>${U}</em>`),row("Label length",`<input data-lp="L.h" data-unit="1" value="${ldVal(L.h)}"><em>${U}</em>`),row("Distance to next label",`<input data-lp="L.dist" data-unit="1" value="${ldVal(L.dist)}"><em>${U}</em>`),row("Extended layout",`<select data-lp="L.ext">${["No","Yes"].map(o=>`<option ${L.ext===o?"selected":""}>${o}</option>`).join("")}</select>`)],
   "Extended layout settings":[row("Label sequence",`<span class="ro">1 of 1</span>`),row("Customer",`<input data-lp="L.cust" value="${esc(L.cust||"")}">`)],
   _:[row("Resolution",`<span class="ro">8 Dots / mm</span>`)]};
  h=D.alpha?Object.values(rows).flat().sort().join(""):grp("Resource",rows.Resource.join(""))+grp("Properties",rows.Properties.join(""))+grp("Extended layout settings",rows["Extended layout settings"].join(""))+rows._.join("");
  return tb+`<div class="ldpg">${h}</div><div class="pwdesc"><b>Label width</b><br>Width of the label across the print head in ${D.unit==="Dot"?"dots (8 dots = 1 mm)":"mm"}.</div>`}
 const G=[E.fld?row("Field",sel("fld",FLD_TYPES.map(x=>x[1]),(FLD_TYPES.find(x=>x[0]===E.fld)||[0,""])[1])):row("Field",`<span class="ro">${esc({line:"Line",rect:"Rectangle",ellipse:"Ellipse"}[E.type]||E.type)}</span>`)];
 if(/^(tf|static|date|clock|code|logo)$/.test(E.fld))G.push(row(E.fld==="tf"?"Text field no.":E.fld==="date"?"Date no.":E.fld==="code"?"Code no.":E.fld==="logo"?"Logo no.":"Number",sel("n",E.fld==="tf"?[1,2,3,4,5,6,7,8,9,10]:E.fld==="date"?[1,2,3]:[1,2,3])));
 if(E.fld==="simple")G.push(row("Data",sel("data",SIMPLE_DATA)));
 if(E.fld==="date")G.push(row("Date format",sel("fmt",["Julian","DDMMM","DD.MM.YY","DD/MM/YYYY"])));
 if(E.fld==="unit")G.push(row("Currency sign",sel("cur",[false,true])));
 const P=[row("X",num("x",E.x)),row("Y",num("y",E.y)),row("Width",num("w",E.w)),row("Height",num("h",E.h)),row("Rotation",sel("rot",[0,90,180,270]))];
 const A=[row("Frame",sel("frame",["No border","Thin","Rounded","Double","Thick","Shadow"])),row("Filling",sel("fill",["No Fill","Black","Dense","Grey","Light"]))];
 const isTxt=E.fld&&!/^(code|logo)$/.test(E.fld);
 const Sx=isTxt?[row("Horizontal",sel("halign",["Left","Center","Right"])),row("Vertical",sel("valign",["Top","Middle","Bottom"])),row("Excess length",sel("excess",["Yes","No"])),row("Descender",sel("desc",["Yes","No"])),row("Autom. Line",sel("autoline",["Yes","No"])),row("Line spacing",sel("lspace",["Auto","1.0","1.2","1.5"])),
  row("Font",sel("font",Object.keys(FONTS).map(f=>f),E.font)),row("Font (sorted)",`<input data-lp="size" value="${String(Number(E.size).toFixed(2)).padStart(5,"0")} mm (${E.bold?"bold":"normal"})">`),row("Bold",sel("bold",[false,true])),row("Content",E.fld==="fixed"||E.fld==="multi"?`<input data-lp="text" value="${esc((E.text||"").replace(/\n/g,"⏎"))}">`:`<span class="ro">${esc(E.name)}</span>`)]:[];
 if(D.alpha)h=[...G,...P,...A,...Sx].sort().join("");
 else h=grp("General",G.join(""))+grp("Position",P.join(""))+grp("Appearance",A.join(""))+(Sx.length?grp("Settings",Sx.join("")):"");
 const desc=E.fld==="multi"?"<b>Multi</b><br>Mix of fixed text and data in one field (photo: Property window Field = Multi on the Co-op layout).":E.fld==="tf"?`<b>Text field ${String(E.n).padStart(2,"0")}</b><br>A slot only. What prints comes from the article's General text no. ${E.n} in Data Maintenance (Articles › Texts).`:E.fld==="fixed"?"<b>Fixed text</b><br>Typed here in BRAIN2 and prints the same on every article.":E.fld==="date"?`<b>Date field ${String(E.n).padStart(2,"0")}</b><br>${E.n===1?"Date 1: printed as Julian code (day of year) on the site's labels.":"Date 2: the use-by date (shelf life 1 from the article)."} The 23.09.76 / 14.12.01 shown here is only the designer's preview.`:E.fld==="static"?"<b>Static text</b><br>Number into the static text table (STATTXT), sent with Download Data.":`<b>${esc(E.name)}</b><br>Value filled by the terminal when the label prints.`;
 return tb+`<div class="ldpg">${h}</div><div class="pwdesc">${desc}</div>`}
/* ---- main view */
function ldHTML(){const D=B2.ld;D.unit=D.unit||"Dot";if(D.legend==null)D.legend=true;if(D.rulers==null)D.rulers=true;const L=ldLayout();const p=B2.db.plu[D.plu]||Object.values(B2.db.plu)[0];const E=ldEl();
 const z=D.zoom;let h=`<div class="ld2"><div class="ldmenu"><i class="grip"></i>${Object.keys(LD_MENUS).map(m=>`<span class="ldm ${D.menu===m?"on":""}" data-b="ldmenu" data-v="${m}">${m}</span>`).join("")}
  ${D.menu?`<div class="lddrop" style="left:${Object.keys(LD_MENUS).indexOf(D.menu)*62+10}px">${LD_MENUS[D.menu].map(i=>i==="—"?"<hr>":`<div data-b="ldmi" data-v="${esc(D.menu+":"+i)}">${esc(i)}</div>`).join("")}</div>`:""}</div>
 ${ldToolbars(E,L)}
 <div class="ldbody"><div class="lddock"><div class="ldpanel"><div class="ldph"><span>${D.dock==="lw"?"Label window":"Project view (2SFGS..."}</span><span class="pin">▾ ⊓ ✕</span></div>
  ${D.dock==="lw"?`<div class="ldtree">${L?L.els.map(e=>`<div class="ldtn ${e.id===D.sel?"on":""}" data-b="ldselel" data-v="${e.id}">${esc(e.name)}</div>`).join(""):""}</div>`:
  `<div class="pvtb"><button class="ldi" data-b="ldpvopen" title="Open">${ico16("green")}</button><i class="sep"></i><button class="ldi" data-b="ldpvdel" title="Delete layout">${ico16("x")}</button><button class="ldi" data-b="ldlprop" title="Properties">${ico16("prop")}</button><button class="ldi" data-b="ldpvref" title="Refresh">${ico16("refresh")}</button></div>
  <div class="ldtree pv">${Object.values(B2.db.layouts).sort((a,b)=>a.no-b.no).map(x=>`<div class="ldtn ${String(x.no)===String(D.pvsel??D.cur)?"on":""}" data-b="ldpvsel" data-v="${x.no}" data-dbl="ldopen">${esc(ldTitle(x))}${D.dirty[x.no]?" *":""}</div>`).join("")}</div>`}
  <div class="dtabs"><span class="${D.dock!=="lw"?"on":""}" data-b="lddock" data-v="proj">Project view...</span><span class="${D.dock==="lw"?"on":""}" data-b="lddock" data-v="lw">Label wind...</span></div></div>
  <div class="ldpanel grow"><div class="ldph"><span>Property window</span><span class="pin">▾ ⊓ ✕</span></div>${ldProps(E,L)}</div></div>
 <div class="ldcenter"><div class="lddocs">${D.open.filter(n=>B2.db.layouts[n]).map(n=>`<span class="lddoc ${n===D.cur?"on":""}" data-b="ldopen" data-v="${n}">${ico16("lres")}${esc(ldDoc(B2.db.layouts[n]))}${D.dirty[n]?" *":""} <b data-b="ldclose" data-v="${n}">✕</b></span>`).join("")}</div>
  <div class="ldwork"><div class="ldcanvas" id="ldCanvas">${L?`${D.rulers?`<div class="ldruler h" style="width:${L.w*z+40}px">${Array.from({length:Math.ceil(L.w/10)+1},(_,i)=>`<span style="left:${i*10*z}px">${i}</span>`).join("")}</div><div class="ldruler v" style="height:${L.h*z+40}px">${Array.from({length:Math.ceil(L.h/10)+1},(_,i)=>`<span style="top:${i*10*z}px">${i}</span>`).join("")}</div>`:""}
   <div class="ldpaper ${D.legend?"lg":""} ${D.grid?"grd":""}" style="width:${L.w*z}px;height:${L.h*z}px;--z:${z}px">${renderLayout(L,p,layoutCtx(p,null,"FP2-L03",B2.db),z,{design:true,data:D.data,sel:D.sel,db:B2.db})}</div>`:`<div class="b2pad">Open a layout from the Project view (double-click).</div>`}</div>
   <div class="ldgrey">${D.legend?`<div class="ldleg">${[["Thermal bars with different lengths","#3a6fd8"],["Carrier tape","#e3c75a"],["Tolerance area (Print on paper not guaranteed)","#e8a0a0"],["Calibration area with conflict","#d04040"],["Calibration area without conflict","#40a040"]].map(([t,c])=>`<label><i style="background:${c}"></i>${esc(t)}</label>`).join("")}</div>`:""}</div></div>
  <div class="ldted"><div class="ldph"><span>Text field editor</span><span class="pin">▾ ⊓ ✕</span></div>
   ${E&&(E.fld==="fixed"||E.fld==="multi")?`<div class="ldtedb"><textarea id="ldTxt" rows="2" spellcheck="false">${esc(E.text||"")}</textarea><span class="sm">${E.fld==="fixed"?"Fixed text: typed here, the same on every article.":"Multi field: fixed text mixed with data, e.g. {Shelf-life 1}, {Time}, {PLU}."}</span></div>`:
     `<div class="ldtedb ro"><div>${E?esc(E.name):""}</div><div class="sm">${E?(E.fld==="tf"?`Content is set per article in Data Maintenance › Articles › Texts › General text no ${E.n}.`:E.fld==="static"?"Content from the static text table.":"Content filled by the terminal."):""}</div></div>`}</div></div></div>
 <div class="ldstatus"><span>Press F1, to display the Help.</span><span class="pn">Project name: 2SFGScunthorpeFreshPack2 V2</span><span>1009</span><span>Gx labeling mode</span><span id="ldPos">[Dots] ${D.pos||"0, 0"}</span><span class="ck">CAP</span><span class="ck">NUM</span><span class="ck">SCRL</span></div>
 ${D.ctx?ldCtxMenu():""}</div>`;
 return h}
function ldCtxMenu(){const D=B2.ld,c=D.ctx;const E=ldEl();
 const items=E?[["Edit:Cut","Cut"],["Edit:Copy","Copy"],["Edit:Paste","Paste"],["Edit:Delete","Delete"],["—"],["sub","Change field type ▸"],["—"],["Order:front","Bring to front"],["Order:back","Send to back"],["—"],["Properties:Field properties","Properties"]]:[["Edit:Paste","Paste"],["—"],["Properties:Label properties","Label properties"]];
 return`<div class="ldctx" style="left:${c.x}px;top:${c.y}px">${items.map(([v,t])=>v==="—"?"<hr>":v==="sub"?`<div class="sub">${t}<div class="ldsub">${FLD_TYPES.map(([f,n])=>`<div data-b="ldchg" data-v="${f}" class="${E.fld===f?"cur":""}">${esc(n)}</div>`).join("")}</div></div>`:`<div data-b="${v.startsWith("Order:")?"ldorder":"ldmi"}" data-v="${esc(v.startsWith("Order:")?v.slice(6):v)}">${esc(t)}</div>`).join("")}</div>`}
const pr=(k,v)=>`<div class="pgr"><span>${k}</span>${v}</div>`;
const f1=n=>Number(n).toFixed(1);
function ldPush(){const D=B2.ld;D.undo.push(JSON.stringify(ldLayout()));if(D.undo.length>40)D.undo.shift();D.redo=[];D.dirty[D.cur]=true}
function ldAdd(t){const L=ldLayout();if(!L)return;ldPush();
 if(/^(line|rect|ellipse)$/.test(t)){const e={id:"e"+Date.now().toString(36),name:{line:"Line",rect:"Rectangle",ellipse:"Ellipse"}[t]+" "+String(L.els.filter(e=>e.type===t).length+1).padStart(2,"0"),type:t,x:5,y:5,w:t==="line"?30:20,h:t==="line"?.4:10,rot:0};L.els.push(e);B2.ld.sel=e.id;return}
 const n=t==="tf"?Math.min(10,L.els.filter(e=>e.fld==="tf").reduce((a,e)=>Math.max(a,e.n),0)+1):t==="date"?2:L.els.filter(e=>e.fld===t).length+1;
 const e={id:"e"+Date.now().toString(36),fld:t,n,type:"field",x:5,y:5,w:t==="code"?28:t==="logo"?10:t==="price"?16:10,h:t==="code"?13:t==="logo"?9:t==="price"?7:4.5,rot:0,font:"Font 004",size:t==="price"?5:1.8,bold:false,halign:"Center",valign:"Top",frame:"No border",fill:"No Fill",excess:"Yes",desc:"Yes",autoline:"Yes",lspace:"Auto",text:t==="fixed"?"Text":"",data:t==="simple"?"Article text":undefined,fmt:t==="date"?"DDMMM":undefined};
 e.name=fldName(e);L.els.push(e);B2.ld.sel=e.id;b2Audit("Label Designer","Layout "+L.no,"Field added","",e.name)}
function ldChangeType(f){const E=ldEl();if(!E)return;ldPush();const old=E.name;E.fld=f;E.type="field";if(f==="simple"&&!E.data)E.data="Article text";if(f==="date"){E.n=E.n>3?2:E.n||2;E.fmt=E.n===1?"Julian":"DDMMM"}if(/^(code|logo|clock|devno|price|weight|unit)$/.test(f)&&E.n>3)E.n=1;if(f==="fixed"&&!E.text)E.text="Text";E.name=fldName(E);b2Audit("Label Designer","Layout "+B2.ld.cur,"Change field type",old,E.name);b2Ev("fieldType")}
function ldSave(){const D=B2.ld,L=ldLayout();if(!L)return;L.ver++;L.changed=fmtD(new Date())+" "+new Date().toTimeString().slice(0,5);L.by=B2USER;if(L.status==="Released")L.status="Changed";D.dirty[D.cur]=false;b2Audit("Label Designer","Layout "+L.no,"Saved","version "+(L.ver-1),"version "+L.ver);b2Ev("layoutSaved");toast(`Layout ${pad3(L.no)} saved (version ${L.ver}). Export it to the devices, and in Data Maintenance point a label parameter set at it.`)}
function ldMenu(cmd){const [m,i]=cmd.split(":");B2.ld.menu=null;B2.ld.ctx=null;const D=B2.ld;
 if(i==="Save")return ldSave();if(i==="Export to…")return openExport();if(i==="Print test label")return toast("Test label sent to the PC's printer at original size (simulated)");
 if(i==="New layout")return newLayout();if(i==="Save as…"){const L=ldLayout();if(!L)return;B2.modal={k:"new",title:"Save as",label:"New layout number",val:String(Number(L.no)+1),ok:val=>{const n=Number(val);if(!(n>0)||B2.db.layouts[n]){toast("Enter a free layout number");return}B2.db.layouts[n]=Object.assign(clone(L),{no:n,name:n+" "+L.name.replace(/^\d+\s*/,"")+" copy",ver:1,status:"Draft",changed:fmtD(new Date()),by:B2USER});D.open.push(String(n));D.cur=String(n);b2Audit("Label Designer","Layout "+n,"Saved as copy of",L.no,n);b2Ev("layoutCopied")}};return}
 if(i==="Open layout…"){B2.modal={k:"pick",title:"Open layout",items:Object.values(B2.db.layouts).sort((a,b)=>a.no-b.no).map(x=>[String(x.no),ldTitle(x)]),act:"ldopen"};return}
 if(i==="Close")return ldClose(D.cur);if(i==="Close all documents"){D.open=[];D.cur=null;return}
 if(i==="Undo")return ldUndo();if(i==="Redo")return ldRedo();if(i==="Delete")return ldDel();
 if(i==="Copy"||i==="Cut"){if(ldEl())D.clip=clone(ldEl());if(i==="Cut")ldDel();return}
 if(i==="Paste"&&D.clip&&ldLayout()){ldPush();const e=Object.assign(clone(D.clip),{id:"e"+Date.now().toString(36),x:D.clip.x+2,y:D.clip.y+2});ldLayout().els.push(e);D.sel=e.id;return}
 if(i==="Select all"){toast("Select all: pick fields one at a time in this trainer");return}
 if(i==="Zoom in")D.zoom=Math.min(14,D.zoom+1);if(i==="Zoom out")D.zoom=Math.max(1,D.zoom-1);if(i==="Zoom 100 %")D.zoom=4;if(i==="Fit to window"){const L=ldLayout();if(L)D.zoom=Math.max(1,Math.min(14,Math.floor(Math.min(560/L.h,760/L.w))))}
 if(i==="Project view")D.dock="proj";if(i==="Label window")D.dock="lw";if(i==="Grid")D.grid=!D.grid;if(i==="Show article data (preview)")D.data=!D.data;
 if(i==="Next document"&&D.open.length){D.cur=D.open[(D.open.indexOf(D.cur)+1)%D.open.length]}
 if(m==="Elements"){const f=FLD_OF[i]||{Line:"line",Rectangle:"rect",Ellipse:"ellipse"}[i];if(f)ldAdd(f)}
 if(i==="Label properties"||i==="Project properties")B2.ld.sel=null;
 if(i==="Field properties"&&!ldEl())toast("Select a field first");
 if(i==="Change field type…"){if(!ldEl()){toast("Select a field first");return}D.ctx={x:300,y:200}}}
function ldUndo(){const D=B2.ld;const s=D.undo.pop();if(!s)return;(D.redo=D.redo||[]).push(JSON.stringify(B2.db.layouts[D.cur]));B2.db.layouts[D.cur]=JSON.parse(s)}
function ldRedo(){const D=B2.ld;const s=(D.redo||[]).pop();if(!s)return;D.undo.push(JSON.stringify(B2.db.layouts[D.cur]));B2.db.layouts[D.cur]=JSON.parse(s)}
function ldDel(){const L=ldLayout(),e=ldEl();if(!e)return;ldPush();L.els=L.els.filter(x=>x!==e);b2Audit("Label Designer","Layout "+L.no,"Field deleted",e.name,"");B2.ld.sel=null}
function ldClose(n){const D=B2.ld;if(D.dirty[n]){B2.modal={k:"confirm",title:"Label Designer",text:`Save changes to ${ldDoc(B2.db.layouts[n])}?`,yes:()=>{const c=D.cur;D.cur=n;ldSave();D.cur=c;ldClose(n)},no:()=>{D.dirty[n]=false;ldClose(n)}};return}D.open=D.open.filter(x=>x!==n);if(D.cur===n)D.cur=D.open[0]||null}
function newLayout(){let n=100;while(B2.db.layouts[n])n++;B2.db.layouts[n]={no:n,name:n+" New layout",cust:"",w:60,h:40,dpm:8,dist:1.25,ext:"No",ver:1,status:"Draft",changed:fmtD(new Date()),by:B2USER,els:[],_f:true};B2.ld.open.push(String(n));B2.ld.cur=String(n);B2.ld.sel=null;b2Audit("Label Designer","Layout "+n,"Created","","Draft")}
function ldOrder(v){const L=ldLayout(),e=ldEl();if(!e)return;ldPush();const i=L.els.indexOf(e);L.els.splice(i,1);if(v==="front")L.els.push(e);else if(v==="back")L.els.unshift(e);else if(v==="fwd")L.els.splice(Math.min(L.els.length,i+1),0,e);else L.els.splice(Math.max(0,i-1),0,e)}
function ldAlign(v){const L=ldLayout(),E=ldEl();if(!E)return;ldPush();if(v==="L")E.x=0;if(v==="R")E.x=L.w-E.w;if(v==="T")E.y=0;if(v==="B")E.y=L.h-E.h;if(v==="CH")E.x=(L.w-E.w)/2;if(v==="CV")E.y=(L.h-E.h)/2;if(v==="W"){E.x=0;E.w=L.w}if(v==="H")E.h=Math.max(1,(E.size||2)*1.2);if(v==="M")E.valign="Middle"}

/* ---------------------------------------------------------------- export wizard (File › Export to…) */
function openExport(pre){B2.modal={k:"export",step:1,res:pre||[B2.ld.cur].filter(Boolean),logos:[],devs:[],log:[],busy:false}}
function exportHTML(M){const st=["Select resources","Select devices","Transfer","Result"];
 let b=`<div class="wiz"><div class="wizh"><b>Export to…</b><span>Step ${M.step} of 4 · ${st[M.step-1]}</span></div><div class="wizb">`;
 if(M.step===1)b+=`<p>Choose the layouts and logos to transfer to the devices.</p><div class="chkl">${Object.values(B2.db.layouts).sort((a,b)=>a.no-b.no).map(x=>`<label><input type="checkbox" data-wx="res" value="${x.no}" ${M.res.includes(String(x.no))?"checked":""}> ${esc(ldTitle(x))} <span class="sm">v${x.ver} · ${esc(x.status)}</span></label>`).join("")}${Object.entries(B2.db.logo).map(([n,x])=>`<label><input type="checkbox" data-wx="logo" value="${n}" ${M.logos.includes(n)?"checked":""}> ${String(n).padStart(4,"0")}.${esc(x)} <span class="sm">logo</span></label>`).join("")}</div>`;
 if(M.step===2)b+=`<p>Transfer selected resources to the devices.</p><div class="chkl">${Object.entries(LINES).map(([k,L])=>`<div class="sm" style="margin-top:6px"><b>${esc(k)}</b> · ${esc(L.model)}</div>`+B2.devices.filter(d=>d.line===k).map(d=>`<label><input type="checkbox" data-wx="dev" value="${d.id}" ${M.devs.includes(d.id)?"checked":""}> ${esc(d.id)} <span class="sm">${d.ip}${d.legacy?" · legacy, layouts only":""}</span></label>`).join("")).join("")}</div>`;
 if(M.step>=3)b+=`<div class="xtree">${M.devs.map(id=>`<div>⊟ ${esc(id)}</div>`+M.res.map(n=>B2.db.layouts[n]?`<div class="i1">🗋 ${esc(ldTitle(B2.db.layouts[n]))}</div>`:"").join("")+M.logos.map(n=>`<div class="i1">🖼 ${String(n).padStart(4,"0")}.${esc(B2.db.logo[n])}</div>`).join("")).join("")}</div><div class="xlog">${M.log.map(l=>`<div class="${l.cls||""}"><span>${esc(l.t)}</span>${l.p!=null?`<i class="pbar"><b style="width:${l.p}%"></b></i>`:""}</div>`).join("")}</div>${M.busy&&M.cur?`<div class="cbox"><b>Communication with _connect.BRAIN</b><div>${esc(M.cur)}</div><i class="pbar"><b style="width:${M.curp||0}%"></b></i></div>`:""}`;
 b+=`</div><div class="wizf">${bb("wizback","< Back",{dis:M.step===1||M.step>=3})}${M.step<3?bb("wiznext","Next >",{pri:true}):bb("wizclose",M.busy?"Running…":"Finish",{pri:true,dis:M.busy})}${bb("wizcancel","Cancel",{dis:M.busy})}</div></div>`;return b}
function runExport(M){M.busy=true;M.step=3;const devs=M.devs.slice();const items=[];devs.forEach(id=>{items.push([id,`Open connection to device '${id}'`]);items.push([id,"Get platform information"]);items.push([id,`Transfer ${M.res.length} layout(s)${M.logos.length?" and "+M.logos.length+" logo(s)":""}`]);items.push([id,"Close connection"])});
 let i=0;const step=()=>{if(i>=items.length){const done=b2ExportLayouts(M.res,devs);M.busy=false;M.cur=null;M.step=4;M.log.push({t:`Transfer finished: ${done.length} device(s), layouts ${M.res.map(pad3).join(", ")}`,cls:"okl"});
   const noSlot=M.res.filter(n=>!Object.values(B2.db.label).some(l=>l.layout===String(n)));if(noSlot.length)M.log.push({t:`No label parameter set points at layout ${noSlot.map(pad3).join(", ")} yet. In Data Maintenance › Label Parameters create the set (the "slot"), then match it to the PLUs.`,cls:"sm"});
   B2.transfers.unshift({at:new Date(),user:B2USER,dev:devs.join(", "),what:"Export layouts "+M.res.join(", "),res:"OK"});b2Audit("Label Designer","Export","Layouts "+M.res.join(", "),"",devs.join(", "));b2Ev("layoutExported");if(S)render();b2Render();return}
  const [id,t]=items[i];M.cur=t.startsWith("Get platform")?`Get platform information from device '${id}'…`:t;M.curp=0;const row={t,p:0};M.log.push(row);b2Render();
  let p=0;const tk=setInterval(()=>{p+=25;row.p=Math.min(100,p);M.curp=row.p;if(p>=100){clearInterval(tk);if(t.startsWith("Get platform")){const d=devLine(id);M.log.push({t:`  ${d.model} · GT-12C · ${d.legacy?"legacy platform: label layouts accepted, master data read only":"BX-NET ok"}`,cls:"sm"})}i++;step()}else b2Render()},reduced()?20:110)};step()}
