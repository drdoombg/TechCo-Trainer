/* the subtitle names the site only if you filled it in under Settings */
function renderHead(){const n=(SITE.name||"").trim();
 $("#subtitle").textContent=`Replica of the Bizerba GLM-I terminals${n?" at "+n:""}, Bizerba Brain 2 and the separate Data Management program. Built from the FP2 photos, the 27/10/2026 BRAIN2 / Data Maintenance photos and test labels, the GLM-I manual and the Scunthorpe price list.`}

/* boot */
const startLine=store.get("line","FP2-L04");useLine(LINES[startLine]?startLine:"FP2-L04");
setMode(viewMode);renderHead();renderSet();renderRef();setTimeout(fit,60);
b2Wire();{let v=store.get("appView","term");if(v==="ld")v="b2";setView(v,null)}
window.__trainer={callPLU,start,stop,useLine,get ST(){return ST},get DBS(){return DBS},B2,get S(){return S},get DB(){return DB},get PRODUCTS(){return PRODUCTS},TREE,MAN,importPLU,parseCSV};
})();
