pdfjsLib.GlobalWorkerOptions.workerSrc="https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";
const $=id=>document.getElementById(id);
let excelRows=[], results=[];

function cleanCol(x){return String(x??"").trim().replace(/\s+/g," ")}
function sval(x){return x==null?"":String(x).trim()}
function poVal(x){return sval(x).replace(/\.0$/,"")}
function shipVal(x){return poVal(x).replace(/^0+/,"")}
function num(x){let n=Number(String(x??"").replace(/,/g,""));return Number.isFinite(n)?n:0}
function norm(s){return sval(s).toUpperCase().replace(/\b(M|W)-/g,"").replace(/[^A-Z0-9]+/g," ").replace(/\s+/g," ").trim()}
function filenamePO(name){let m=sval(name).match(/(?:^|\D)(\d{10})(?!\d)/);return m?m[1]:""}
function esc(s){return sval(s).replace(/[&<>"']/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]))}
function fmt(n){return Number.isInteger(n)?String(n):String(Math.round(n*100)/100)}

function levenshtein(a,b){
  const m=a.length,n=b.length,dp=Array(n+1).fill(0); for(let j=0;j<=n;j++)dp[j]=j;
  for(let i=1;i<=m;i++){let prev=dp[0];dp[0]=i;for(let j=1;j<=n;j++){let t=dp[j];dp[j]=Math.min(dp[j]+1,dp[j-1]+1,prev+(a[i-1]===b[j-1]?0:1));prev=t}} return dp[n]
}
function compatible(expected,desc){
  const a=norm(expected),b=norm(desc); if(!a)return true;if(a.includes(b)||b.includes(a))return true;
  const toks=a.split(" ").filter(x=>x.length>=3),hits=toks.filter(x=>b.includes(x)).length,ts=100*hits/Math.max(1,toks.length);
  const fs=100*(1-levenshtein(a,b)/Math.max(1,a.length,b.length));return Math.max(ts,fs)>=58
}

async function readExcel(file){
  const data=await file.arrayBuffer(), wb=XLSX.read(data,{type:"array"}), ws=wb.Sheets[wb.SheetNames[0]];
  const raw=XLSX.utils.sheet_to_json(ws,{defval:""}); if(!raw.length)throw Error("Excel sheet is empty.");
  return raw.map(r=>{const o={};Object.entries(r).forEach(([k,v])=>o[cleanCol(k)]=v);return o});
}
function excelSummary(po){
  const rows=excelRows.filter(r=>poVal(r["Purchasing Doc"])===po);if(!rows.length)return null;
  const req=["Vendor Name","Purchasing Doc","Material","Style name","Color","Open Qty."];
  const miss=req.filter(c=>!(c in rows[0]));if(miss.length)throw Error("Missing required Excel columns: "+miss.join(", "));
  const map=new Map();
  for(const r of rows){
    const mat=sval(r.Material).toUpperCase(),side=sval(r.Size).toUpperCase()||"UNSPECIFIED",q=num(r["Open Qty."]);
    if(!map.has(mat))map.set(mat,{material:mat,style:sval(r["Style name"]),color:sval(r.Color),qty:0,sides:{}});
    const x=map.get(mat);x.qty+=q;x.sides[side]=(x.sides[side]||0)+q;
  }
  return {po_no:po,vendor:sval(rows[0]["Vendor Name"]),ship_to:shipVal(rows[0]["Ship to ID"]),qty:rows.reduce((a,r)=>a+num(r["Open Qty."]),0),items:[...map.values()]};
}
function getFobColumn(){
  if(!excelRows.length)return null;
  const keys=Object.keys(excelRows[0]);
  return keys.find(k=>norm(k)==="NET FOB PRICE") || keys.find(k=>norm(k).includes("NET FOB")&&norm(k).includes("PRICE")) || null;
}
function getFobValue(row){
  const c=getFobColumn(); if(!c)return null;
  const raw=sval(row[c]).replace(/[$,\s]/g,""); if(raw==="")return null;
  const n=Number(raw); return Number.isFinite(n)?Math.round(n*1000000)/1000000:null;
}
function globalFobCheck(ex){
  const col=getFobColumn();
  if(!col)return {ok:false,issues:["Excel column 'Net FOB price' was not found."],details:[]};
  if(!ex)return {ok:false,issues:["PO not found in Excel; Net FOB price cannot be checked."],details:[]};
  const gm=new Map();
  for(const r of excelRows){
    const mat=sval(r.Material).toUpperCase(), price=getFobValue(r); if(!mat||price===null)continue;
    if(!gm.has(mat))gm.set(mat,{prices:new Set(),pos:new Set()});
    gm.get(mat).prices.add(price); const po=poVal(r["Purchasing Doc"]); if(po)gm.get(mat).pos.add(po);
  }
  const issues=[],details=[];
  for(const item of ex.items){
    const g=gm.get(item.material); if(!g){details.push({material:item.material,prices:[],pos:[],ok:true});continue}
    const prices=[...g.prices].sort((a,b)=>a-b),pos=[...g.pos].sort(),ok=prices.length<=1;
    details.push({material:item.material,prices,pos,ok});
    if(!ok)issues.push(`${item.material}: Net FOB price mismatch — ${prices.map(fmt).join(" / ")} — PO ${pos.join(" / ")}`);
  }
  return {ok:issues.length===0,issues,details};
}

async function pdfText(file){
  const doc=await pdfjsLib.getDocument({data:await file.arrayBuffer()}).promise;let pages=[];
  for(let p=1;p<=doc.numPages;p++){const pg=await doc.getPage(p),tc=await pg.getTextContent();pages.push(tc.items.map(i=>i.str).join(" "))}
  return pages.join("\n");
}
function one(re,text){const m=text.match(re);return m?cleanCol(m[1]):""}
function parsePDF(text,filename){
  const po=one(/Purchase\s*Order\s*No\.?\s*:?\s*(\d{7,12})/i,text);
  const vendorNo=one(/Vendor\s*No\.?\s*:?\s*(\d+)/i,text),ship=one(/Ship-To\s*:?\s*(\d+)/i,text);
  const vendorBlock=one(/Vendor\s*Address\s*(.*?)\s*Information/is,text);
  const vendorParts=vendorBlock.split(/\s{2,}|\n/).map(cleanCol).filter(Boolean);
  let vendor=vendorParts[0]||"";
  const vraw=(text.match(/Vendor\s*Address\s*(.*?)\s*Information/is)||[])[1]||"";
  const known=vraw.match(/(VIETNAM JIN CHANG SHOES CO|LI FENG YUEN FOOTWEAR \(CAMBODIA\)|SUPERIOR QUALITY WORLDWIDE INC\.|ROFU \(THAILAND\) LTD\.|GOLDEN PROSPER FOOTWEAR CO\.,? LTD\.?)/i);
  if(known) vendor=known[1];
  else {
    const boundary=vraw.search(/\b(?:LOT|NO\.?|NATIONAL\s+ROAD|ROAD|STREET|INDUSTRY|INDUSTRIAL|DISTRICT|PROVINCE|PHNOM\s+PENH|BANGKOK|VIETNAM|CAMBODIA|THAILAND)\b/i);
    if(boundary>0) vendor=cleanCol(vraw.slice(0,boundary));
  }
  let vendorAddress=cleanCol(vraw.slice(vraw.toUpperCase().indexOf(vendor.toUpperCase())+vendor.length));
  const sraw=(text.match(/Shipping\s*Address\s*:?\s*(.*?)\s*Ship-To\s*:/is)||[])[1]||"";
  const shippingAddress=cleanCol(sraw);

  // Normalize flattened PDF text then split item blocks at item-number + S-material anchors.
  const flat=cleanCol(text), anchor=/\b(\d{1,4})\s+(S\d{6,})\b/g, anchors=[]; let m;
  while((m=anchor.exec(flat))) anchors.push({idx:m.index,item:m[1],material:m[2].toUpperCase(),end:anchor.lastIndex});
  const items=[];
  for(let i=0;i<anchors.length;i++){
    const a=anchors[i],end=i+1<anchors.length?anchors[i+1].idx:flat.length,block=flat.slice(a.end,end).split(/Total Value|Terms and Conditions|Purchase Order No\./i)[0];
    let sides={}; for(const sm of block.matchAll(/\b(LF|RT)\s+EA\s+(\d+(?:\.\d+)?)/gi))sides[sm[1].toUpperCase()]=(sides[sm[1].toUpperCase()]||0)+num(sm[2]);
    let qty=0;
    // Main line typically: price 1 EA QTY amount. Avoid LF/RT rows.
    let qm=block.match(/\b\d+(?:\.\d+)?\s+1\s+EA\s+(\d+(?:\.\d+)?)\s+\d[\d,]*\.\d{2}\b/i);
    if(qm)qty=num(qm[1]); else if(Object.keys(sides).length)qty=Object.values(sides).reduce((x,y)=>x+y,0);
    let desc=block.replace(/\b(LF|RT)\s+EA\s+\d+(?:\.\d+)?/gi," ").replace(/\b\d+(?:\.\d+)?\s+1\s+EA\s+\d+(?:\.\d+)?\s+\d[\d,]*\.\d{2}\b/g," ");
    items.push({item:a.item,material:a.material,description:cleanCol(desc),qty,sides});
  }
  return {filename,po_no:po,vendor_no:vendorNo,vendor,vendor_address:vendorAddress,ship_to:ship,shipping_address:shippingAddress,qty:items.reduce((a,x)=>a+x.qty,0),items};
}
function compare(ex,pdf){
  if(!ex)return {status:"MISMATCH",checks:[],issues:["PO not found in Excel"],warnings:[],fob:globalFobCheck(ex)};
  const checks=[
    ["PO Number",ex.po_no===pdf.po_no,`Excel ${ex.po_no} / PDF ${pdf.po_no}`],
    ["Filename",filenamePO(pdf.filename)===pdf.po_no,`Filename PO ${filenamePO(pdf.filename)||"not found"} / PDF ${pdf.po_no}`],
    ["Vendor",norm(ex.vendor)===norm(pdf.vendor),`Excel ${ex.vendor} / PDF ${pdf.vendor}`],
    ["Ship-To",String(ex.ship_to)===String(pdf.ship_to),`Excel ${ex.ship_to} / PDF ${pdf.ship_to}`],
    ["Total Qty",Math.abs(ex.qty-pdf.qty)<1e-9,`Excel ${fmt(ex.qty)} / PDF ${fmt(pdf.qty)}`]
  ].map(([field,ok,detail])=>({field,ok,detail}));
  const pm=new Map(pdf.items.map(x=>[x.material,x])),issues=[],warnings=[];
  for(const x of ex.items){
    const p=pm.get(x.material);if(!p){issues.push(`${x.material}: missing in PDF`);continue}
    if(Math.abs(x.qty-p.qty)>1e-9)issues.push(`${x.material}: qty Excel ${fmt(x.qty)} / PDF ${fmt(p.qty)}`);
    for(const [side,q] of Object.entries(x.sides)){if(p.sides[side]!=null&&Math.abs(q-p.sides[side])>1e-9)issues.push(`${x.material} ${side}: Excel ${fmt(q)} / PDF ${fmt(p.sides[side])}`)}
    if(!compatible(x.style,p.description))warnings.push(`${x.material}: style text may be abbreviated — Excel "${x.style}" / PDF "${p.description}"`);
    if(!compatible(x.color,p.description))warnings.push(`${x.material}: color text may be abbreviated — Excel "${x.color}" / PDF "${p.description}"`);
  }
  const em=new Set(ex.items.map(x=>x.material));for(const p of pdf.items)if(!em.has(p.material))issues.push(`${p.material}: present in PDF but not Excel`);
  const fob=globalFobCheck(ex); return {status:checks.every(x=>x.ok)&&!issues.length?"PASS":"MISMATCH",checks,issues,warnings,fob};
}

function render(){
  $("results").classList.remove("hidden");$("mFiles").textContent=results.length;$("mUnique").textContent=new Set(results.map(r=>r.pdf.po_no)).size;
  $("mPass").textContent=results.filter(r=>r.comp.status==="PASS").length;$("mMismatch").textContent=results.filter(r=>r.comp.status!=="PASS").length;$("mDup").textContent=results.filter(r=>r.duplicate).length;
  const tb=$("summaryTable").querySelector("tbody");tb.innerHTML="";
  for(const r of results){
    const fn=r.comp.checks.find(x=>x.field==="Filename"), itemOK=!r.comp.issues.length;
    const tr=document.createElement("tr");tr.innerHTML=`<td>${esc(r.pdf.po_no)}</td><td>${esc(r.pdf.vendor)}</td><td>${esc(r.pdf.ship_to)}</td><td>${r.ex?fmt(r.ex.qty):""}</td><td>${fmt(r.pdf.qty)}</td>
    <td class="${itemOK?"pass":"bad"}">${itemOK?"PASS":"MISMATCH"}</td><td class="${fn?.ok?"pass":"bad"}">${fn?.ok?"PASS":"MISMATCH"}</td><td class="${r.duplicate?"warn":""}">${r.duplicate?"YES":"NO"}</td>
    <td class="${r.comp.status==="PASS"?"pass":"bad"}">${r.comp.status}</td><td>${esc([...r.comp.issues,...r.comp.checks.filter(x=>!x.ok).map(x=>x.detail)].join(" | "))}</td>`;tb.appendChild(tr)
  }
  const d=$("details");d.innerHTML="";
  for(const r of results){
    const el=document.createElement("details"), label=`${r.pdf.po_no||"PO not detected"} — ${r.pdf.vendor} — ${r.comp.status}${r.duplicate?" — DUPLICATE":""}`;
    let checks=r.comp.checks.map(c=>`<tr><td>${esc(c.field)}</td><td class="${c.ok?"pass":"bad"}">${c.ok?"PASS":"MISMATCH"}</td><td>${esc(c.detail)}</td></tr>`).join("");
    let issues=r.comp.issues.length?`<p class="bad">${r.comp.issues.map(esc).join("<br>")}</p>`:`<p class="pass">All item-level checks passed.</p>`;
    let warnings=r.comp.warnings.length?`<p class="warn">${r.comp.warnings.map(esc).join("<br>")}</p>`:"";
    el.innerHTML=`<summary>${esc(label)}</summary><p><b>Filename:</b> ${esc(r.pdf.filename)}</p><p class="addr"><b>Vendor address:</b> ${esc(r.pdf.vendor_address)}</p><p class="addr"><b>Shipping address:</b> ${esc(r.pdf.shipping_address)}</p><table class="detail-table"><tr><th>Check</th><th>Result</th><th>Detail</th></tr>${checks}</table>${issues}${warnings}`;d.appendChild(el)
  }
}
function exportXlsx(){
  const summary=results.map(r=>{const fn=r.comp.checks.find(x=>x.field==="Filename");return {"PO":r.pdf.po_no,"Vendor Name":r.pdf.vendor,"Vendor Address (PDF)":r.pdf.vendor_address,"Ship-To":r.pdf.ship_to,"Shipping Address (PDF)":r.pdf.shipping_address,"Excel Qty":r.ex?.qty??"","PDF Qty":r.pdf.qty,"Net FOB Price Check":r.comp.fob?.ok?"PASS":"MISMATCH","PO/Filename":fn?.ok?"PASS":"MISMATCH","Duplicate":r.duplicate?"YES":"NO","Result":r.comp.status,"Issues":[...r.comp.issues,...r.comp.checks.filter(x=>!x.ok).map(x=>x.detail)].join(" | "),"FOB Issues":(r.comp.fob?.issues||[]).join(" | "),"Warnings":r.comp.warnings.join(" | "),"Filename":r.pdf.filename}});
  const detail=[];for(const r of results){const em=new Map((r.ex?.items||[]).map(x=>[x.material,x])),pm=new Map(r.pdf.items.map(x=>[x.material,x]));for(const mat of new Set([...em.keys(),...pm.keys()])){const e=em.get(mat),p=pm.get(mat);detail.push({"PO":r.pdf.po_no,"Material":mat,"Style (Excel)":e?.style||"","Color (Excel)":e?.color||"","Excel Qty":e?.qty??"","PDF Qty":p?.qty??"","Excel LF":e?.sides.LF??"","PDF LF":p?.sides.LF??"","Excel RT":e?.sides.RT??"","PDF RT":p?.sides.RT??"","PDF Description":p?.description||""})}}
  const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(summary),"Summary");XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(detail),"Item Detail");XLSX.writeFile(wb,"PO_Check_Result.xlsx")
}
function ready(){ $("checkBtn").disabled=!($("excelFile").files.length&&$("pdfFiles").files.length) }
$("excelFile").onchange=()=>{$("excelName").textContent=$("excelFile").files[0]?.name||"Choose Excel file";ready()}
$("pdfFiles").onchange=()=>{$("pdfName").textContent=$("pdfFiles").files.length?`${$("pdfFiles").files.length} PDF selected`:"Choose PDF files";ready()}
$("downloadBtn").onclick=exportXlsx;
$("checkBtn").onclick=async()=>{
  $("message").innerHTML='<div class="info">Checking files…</div>';results=[];$("results").classList.add("hidden");
  try{
    excelRows=await readExcel($("excelFile").files[0]);const seen={};
    for(const f of $("pdfFiles").files){const text=await pdfText(f),pdf=parsePDF(text,f.name),ex=pdf.po_no?excelSummary(pdf.po_no):null,duplicate=!!seen[pdf.po_no];seen[pdf.po_no]=(seen[pdf.po_no]||0)+1;results.push({pdf,ex,duplicate,comp:compare(ex,pdf)})}
    $("message").innerHTML="";render()
  }catch(e){$("message").innerHTML=`<div class="error">${esc(e.message||e)}</div>`}
};


function badge(ok,labelOK="PASS",labelBad="MISMATCH"){return `<span class="badge ${ok?"ok":"bad"}">${ok?labelOK:labelBad}</span>`}
function issueSummary(r){
 const failed=r.comp.checks.filter(x=>!x.ok).map(x=>x.field);
 const xs=[...failed,...r.comp.issues.map(x=>x.split(":")[0])]; if(r.comp.fob&&!r.comp.fob.ok)xs.push("Net FOB Price");
 return [...new Set(xs)].join(", ")||"—";
}
function getCheck(r,name){return r.comp.checks.find(x=>x.field===name)}
function render(){
 $("results").classList.remove("hidden");
 $("mFiles").textContent=results.length;$("mUnique").textContent=new Set(results.map(r=>r.pdf.po_no)).size;
 $("mPass").textContent=results.filter(r=>r.comp.status==="PASS").length;
 $("mMismatch").textContent=results.filter(r=>r.comp.status!=="PASS").length;
 $("mDup").textContent=results.filter(r=>r.duplicate).length;
 renderRows();
}
function renderRows(){
 const q=($("searchBox")?.value||"").toUpperCase(), f=$("filterResult")?.value||"All Results";
 const tb=$("summaryTable").querySelector("tbody");tb.innerHTML="";
 results.forEach((r,i)=>{
   const hay=[r.pdf.po_no,r.pdf.vendor,r.pdf.filename].join(" ").toUpperCase();
   if(q&&!hay.includes(q))return;
   if(f==="PASS"&&r.comp.status!=="PASS")return;if(f==="MISMATCH"&&r.comp.status==="PASS")return;if(f==="DUPLICATE"&&!r.duplicate)return;
   const vendor=getCheck(r,"Vendor"),ship=getCheck(r,"Ship-To"),fn=getCheck(r,"Filename"),itemOK=!r.comp.issues.length;
   const tr=document.createElement("tr");
   tr.innerHTML=`<td>${i+1}</td><td><b>${esc(r.pdf.po_no)}</b></td><td class="wraptext">${esc(r.pdf.vendor)}</td><td>${esc(r.pdf.ship_to)}</td><td>${r.ex?fmt(r.ex.qty):"—"}</td><td>${fmt(r.pdf.qty)}</td>
   <td>${badge(!!vendor?.ok)}</td><td>${badge(!!ship?.ok)}</td><td>${badge(itemOK)}</td><td>${badge(!!r.comp.fob?.ok)}</td><td>${badge(!!fn?.ok)}</td>
   <td>${r.duplicate?'<span class="badge warn">YES</span>':'<span class="badge neutral">NO</span>'}</td>
   <td>${badge(r.comp.status==="PASS",r.comp.status,r.comp.status)}</td><td class="wraptext">${esc(issueSummary(r))}</td><td><button class="viewbtn" onclick="showCompare(${i})">View</button></td>`;
   tb.appendChild(tr);
 })
}
function showCompare(i){
 const r=results[i],p=$("comparePanel"), vendor=getCheck(r,"Vendor"),ship=getCheck(r,"Ship-To"),fn=getCheck(r,"Filename"),qty=getCheck(r,"Total Qty");
 const failures=[...r.comp.checks.filter(x=>!x.ok).map(x=>`<li><b>${esc(x.field)}:</b> ${esc(x.detail)}</li>`),...r.comp.issues.map(x=>`<li>${esc(x)}</li>`),...(r.comp.fob?.issues||[]).map(x=>`<li><b>Net FOB:</b> ${esc(x)}</li>`)].join("")||"<li>No blocking issues.</li>";
 p.classList.remove("hidden");
 p.innerHTML=`<div class="compare-title"><span>Detailed Comparison — PO ${esc(r.pdf.po_no)} — ${r.comp.status}</span><button class="viewbtn" onclick="document.getElementById('comparePanel').classList.add('hidden')">Close</button></div>
 <div class="compare-grid">
  <div class="compare-card"><h3>From Excel (Source)</h3>
   <div class="kv"><label>PO Number</label><b>${esc(r.ex?.po_no||"—")}</b></div><div class="kv"><label>Vendor Name</label><span>${esc(r.ex?.vendor||"—")}</span></div>
   <div class="kv"><label>Ship-To ID</label><span>${esc(r.ex?.ship_to||"—")}</span></div><div class="kv"><label>Open Qty</label><span>${r.ex?fmt(r.ex.qty):"—"}</span></div>
  </div>
  <div class="compare-card"><h3>From PDF</h3>
   <div class="kv"><label>PO Number</label><b>${esc(r.pdf.po_no)}</b></div><div class="kv"><label>Vendor Name</label><span>${esc(r.pdf.vendor)}</span></div>
   <div class="kv"><label>Vendor Address</label><span>${esc(r.pdf.vendor_address)}</span></div><div class="kv"><label>Ship-To ID</label><span>${esc(r.pdf.ship_to)}</span></div>
   <div class="kv"><label>Shipping Address</label><span>${esc(r.pdf.shipping_address)}</span></div><div class="kv"><label>PDF Qty</label><span>${fmt(r.pdf.qty)}</span></div>
   <div class="kv"><label>Filename</label><span>${esc(r.pdf.filename)}</span></div>
  </div>
  <div class="compare-card"><h3>Result: ${r.comp.status}</h3><div class="issuebox"><b>Issues found</b><ul>${failures}</ul></div>
   <div class="kv"><label>Vendor</label><span>${badge(!!vendor?.ok)}</span></div><div class="kv"><label>Ship-To</label><span>${badge(!!ship?.ok)}</span></div>
   <div class="kv"><label>Total Qty</label><span>${badge(!!qty?.ok)}</span></div><div class="kv"><label>Net FOB Price</label><span>${badge(!!r.comp.fob?.ok)}</span></div><div class="kv"><label>Filename</label><span>${badge(!!fn?.ok)}</span></div>
  </div>
 </div>`;
 p.scrollIntoView({behavior:"smooth",block:"start"});
}
$("searchBox").oninput=renderRows;$("filterResult").onchange=renderRows;
$("clearBtn").onclick=()=>{location.reload()};
$("downloadMismatchBtn").onclick=()=>{
 const keep=results, subset=results.filter(r=>r.comp.status!=="PASS");results=subset;exportXlsx();results=keep;
};


function buildFobGroups(){
 const col=getFobColumn(), gm=new Map(); if(!col)return {col:null,groups:[]};
 for(const r of excelRows){
   const mat=sval(r.Material).toUpperCase(), price=getFobValue(r), po=poVal(r["Purchasing Doc"]);
   if(!mat||price===null)continue;
   if(!gm.has(mat))gm.set(mat,new Map());
   const pm=gm.get(mat); if(!pm.has(price))pm.set(price,new Set()); if(po)pm.get(price).add(po);
 }
 return {col,groups:[...gm.entries()].map(([material,prices])=>({material,prices,ok:prices.size<=1}))};
}
function renderFob(){
 const x=buildFobGroups(), sec=$("fobSection"), tb=$("fobTable").querySelector("tbody"); sec.classList.remove("hidden");tb.innerHTML="";
 if(!x.col){tb.innerHTML='<tr><td colspan="5"><span class="badge bad">Column "Net FOB price" not found</span></td></tr>';return}
 $("fobMaterials").textContent=x.groups.length;$("fobPass").textContent=x.groups.filter(g=>g.ok).length;$("fobBad").textContent=x.groups.filter(g=>!g.ok).length;
 $("fobGoodMsg").classList.toggle("hidden",x.groups.some(g=>!g.ok));
 // mismatches first, then passes
 x.groups.sort((a,b)=>Number(a.ok)-Number(b.ok)||a.material.localeCompare(b.material));
 for(const g of x.groups){
   const entries=[...g.prices.entries()].sort((a,b)=>a[0]-b[0]);
   const prices=entries.map(([p])=>fmt(p)).join(" / ");
   const allpos=[...new Set(entries.flatMap(([,ps])=>[...ps]))].join(", ");
   const compare=entries.map(([p,ps])=>`<div class="pricegroup"><b>${fmt(p)}</b><span>PO: ${[...ps].join(", ")||"—"}</span></div>`).join("");
   const tr=document.createElement("tr");tr.innerHTML=`<td><b>${esc(g.material)}</b></td><td>${badge(g.ok)}</td><td>${esc(prices)}</td><td class="wraptext">${esc(allpos)}</td><td class="wraptext">${compare}</td>`;tb.appendChild(tr);
 }
}
function issueSummaryV4(r){
 const failed=r.comp.checks.filter(x=>!x.ok).map(x=>x.field), xs=[...failed,...r.comp.issues.map(x=>x.split(":")[0])];
 return [...new Set(xs)].join(", ")||"—";
}
function renderRowsV4(){
 const q=($("searchBox")?.value||"").toUpperCase(),f=$("filterResult")?.value||"All Results",tb=$("summaryTable").querySelector("tbody");tb.innerHTML="";
 results.forEach((r,i)=>{
  const mats=(r.ex?.items||[]).map(x=>x.material).join(", "),hay=[r.pdf.po_no,mats,r.pdf.vendor,r.pdf.filename].join(" ").toUpperCase();
  if(q&&!hay.includes(q))return;if(f==="PASS"&&r.comp.status!=="PASS")return;if(f==="MISMATCH"&&r.comp.status==="PASS")return;if(f==="DUPLICATE"&&!r.duplicate)return;
  const vendor=getCheck(r,"Vendor"),ship=getCheck(r,"Ship-To"),fn=getCheck(r,"Filename"),itemOK=!r.comp.issues.length;
  const tr=document.createElement("tr");tr.innerHTML=`<td>${i+1}</td><td><b>${esc(r.pdf.po_no)}</b></td><td class="wraptext">${esc(mats)}</td><td class="wraptext">${esc(r.pdf.vendor)}</td><td>${esc(r.pdf.ship_to)}</td><td>${r.ex?fmt(r.ex.qty):"—"}</td><td>${fmt(r.pdf.qty)}</td><td>${badge(!!vendor?.ok)}</td><td>${badge(!!ship?.ok)}</td><td>${badge(itemOK)}</td><td>${badge(!!fn?.ok)}</td><td>${r.duplicate?'<span class="badge warn">YES</span>':'<span class="badge neutral">NO</span>'}</td><td>${badge(r.comp.status==="PASS",r.comp.status,r.comp.status)}</td><td class="wraptext">${esc(issueSummaryV4(r))}</td><td><button class="viewbtn" onclick="showCompare(${i})">View</button></td>`;tb.appendChild(tr);
 });
}
const oldRender=render; render=function(){oldRender();renderRowsV4();}
$("searchBox").oninput=renderRowsV4;$("filterResult").onchange=renderRowsV4;
$("excelFile").onchange=()=>{
 $("excelName").textContent=$("excelFile").files[0]?.name||"Choose Excel file";
 $("excelOnlyBtn").disabled=!$("excelFile").files.length;
 $("fobSection").classList.add("hidden");
 ready();
};
$("excelOnlyBtn").onclick=async()=>{
 $("message").innerHTML='<div class="info">Checking Excel…</div>';
 try{excelRows=await readExcel($("excelFile").files[0]);renderFob();$("message").innerHTML=""}
 catch(e){$("message").innerHTML=`<div class="error">${esc(e.message||e)}</div>`}
};
$("downloadFobBtn").onclick=()=>{
 const x=buildFobGroups(), rows=[];
 for(const g of x.groups){for(const [price,pos] of [...g.prices.entries()].sort((a,b)=>a[0]-b[0])){rows.push({"Material":g.material,"Status":g.ok?"PASS":"MISMATCH","Net FOB Price":price,"Purchasing Doc(s)":[...pos].join(", ")})}}
 const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(rows),"Net FOB Check");XLSX.writeFile(wb,"Net_FOB_Price_Check.xlsx");
};

// V4.2: final isolated Excel-only binding
(function(){
  const fileInput=document.getElementById("excelFile");
  const fileName=document.getElementById("excelName");
  const excelBtn=document.getElementById("excelOnlyBtn");
  const msg=document.getElementById("message");
  const fobSec=document.getElementById("fobSection");
  if(!fileInput||!excelBtn) return;

  fileInput.addEventListener("change", function(){
    const hasFile=this.files && this.files.length>0;
    if(fileName) fileName.textContent=hasFile ? this.files[0].name : "Choose Excel file";
    excelBtn.disabled=!hasFile;
    excelBtn.classList.toggle("readybtn",hasFile);
    if(fobSec) fobSec.classList.add("hidden");
    if(msg) msg.innerHTML="";
    ready();
  });

  excelBtn.addEventListener("click", async function(ev){
    ev.preventDefault();
    const file=fileInput.files && fileInput.files[0];
    if(!file){
      if(msg) msg.innerHTML='<div class="error">Please choose Source Excel first.</div>';
      return;
    }
    excelBtn.disabled=true;
    const oldText=excelBtn.textContent;
    excelBtn.textContent="Checking Excel...";
    if(msg) msg.innerHTML='<div class="info">Reading Excel and checking Net FOB price by Material...</div>';
    try{
      excelRows=await readExcel(file);
      renderFob();
      if(msg) msg.innerHTML='<div class="successmsg">✓ Excel check completed.</div>';
      if(fobSec) fobSec.scrollIntoView({behavior:"smooth",block:"start"});
    }catch(e){
      console.error(e);
      if(msg) msg.innerHTML='<div class="error">Excel Check Error: '+esc(e && e.message ? e.message : String(e))+'</div>';
    }finally{
      excelBtn.disabled=false;
      excelBtn.textContent=oldText;
      excelBtn.classList.add("readybtn");
    }
  });
})();
