
const AVSENDER = { navn:"Toddum Gårdsysteri AS", adresse:"Gullvikvegen 193", postnr:"7810", poststed:"NAMSOS", orgnr:"929953088", epost:"silje@toddum.com", tlf:"94307328", ref:"Silje Kvaale By", hentested:"Spillum - Namsos" };
const TRANSPORT = ["Kristensen Transport","Nær Engros Namsos","Namdalingen","Henter selv","Egen levering","Post/Bring","Annet"];
const TRANSPORT_ADR = { "Kristensen Transport":"Kristensen Transport AS\nUtvordvegen 850, 7777 Nord-Statland" };
const STATUS = [["ny","Ny"],["produsert","Produsert"],["sendt","Sendt"],["fakturert","Fakturert"]];
const FREKVENS = [["ingen","Ingen fast ordre"],["hver","Hver uke"],["partall","Partallsuker"],["oddetall","Oddetallsuker"]];

const S = { db:null, fase:"start", bruker:null, feil:"", mangler:[], importStatus:"", pdfVindu:null, pollN:0, produkter:[], kunder:[], ordre:[], tab:"uke", ukeDato:new Date(), open:new Set(), override:new Map(), loaded:{p:false,k:false,o:false}, kundeSel:null, kundeSok:"", leggTil:"", sugIdx:0, pending:0, unsubOrdre:null, confirmDel:null, nyttProdukt:false };

/* ---------- hjelpere ---------- */
const $ = s => document.querySelector(s);
const esc = s => String(s ?? "").replace(/[&<>"']/g, c => ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]));
const pad = n => String(n).padStart(2,"0");
function isoUke(d){ const t=new Date(Date.UTC(d.getFullYear(),d.getMonth(),d.getDate())); const day=t.getUTCDay()||7; t.setUTCDate(t.getUTCDate()+4-day); const y=t.getUTCFullYear(); const w=Math.ceil(((t-Date.UTC(y,0,1))/864e5+1)/7); return {y,w,key:`${y}-${pad(w)}`}; }
function mandag(d){ const x=new Date(d.getFullYear(),d.getMonth(),d.getDate()); const day=x.getDay()||7; x.setDate(x.getDate()-day+1); return x; }
const fmtD = d => `${pad(d.getDate())}.${pad(d.getMonth()+1)}`;
function dmy(iso){ if(!iso) return ""; const [y,m,d]=iso.split("-"); return `${d}.${m}.${y.slice(2)}`; }
function dmyLang(iso){ if(!iso) return ""; const [y,m,d]=iso.split("-"); return `${d}.${m}.${y}`; }
const iDag = () => { const d=new Date(); return `${d.getFullYear()}-${pad(d.getMonth()+1)}-${pad(d.getDate())}`; };
function batchFra(iso){ if(!iso) return ""; const [y,m,d]=iso.split("-"); return String(Number(d+m+y.slice(2))*2); }
const uke = () => isoUke(S.ukeDato);
const prodSort = () => S.produkter.filter(p=>p.aktiv!==false).sort((a,b)=>(a.sort??99)-(b.sort??99)||a.varenr.localeCompare(b.varenr));
const prodAlle = () => [...S.produkter].sort((a,b)=>(a.sort??99)-(b.sort??99)||a.varenr.localeCompare(b.varenr));
const prod = v => S.produkter.find(p=>p.varenr===v);
const kunde = id => S.kunder.find(k=>k.id===id);
const nf = (n,dec=0) => n.toLocaleString("nb-NO",{minimumFractionDigits:dec,maximumFractionDigits:Math.max(dec,1)});
function linjer(o){ return Object.entries(o.linjer||{}).filter(([v,l])=>l&&l.antall>0).map(([v,l])=>({varenr:v, antall:l.antall, batch:l.batch||"", p:prod(v)})).sort((a,b)=>((a.p?.sort??99)-(b.p?.sort??99))); }
function dpTekst(antall,p){ if(!p||!p.perDP) return ""; const dp=Math.floor(antall/p.perDP), r=antall%p.perDP; return r? (dp? `${dp} DP + ${r}`:"") : `${dp} DP`; }
function kgNetto(antall,p){ return p? antall*(p.stkNettoG||0)/1000 : 0; }
function kgBrutto(antall,p){ if(!p) return 0; if(!p.perDP) return antall*(p.stkBruttoG||p.stkNettoG||0)/1000; const dp=Math.floor(antall/p.perDP), r=antall%p.perDP; return (dp*(p.dpBruttoG||0) + r*(p.stkBruttoG||0))/1000; }
function adresseLinjer(k,o){ if(!k) return [o?.kundeNavn||""]; if(k.levAdresse) return [k.levNavn||k.navn, ...k.levAdresse.split(/\n/)]; return [k.navn, k.adresse, k.adresse2, [k.postnr,k.poststed].filter(Boolean).join(" ")].filter(Boolean); }
function trengerCMR(o){ const k=kunde(o.kundeId); return o.transport==="Kristensen Transport" || (k&&k.cmr); }
function toast(msg){ const t=$("#toast"); t.textContent=msg; t.hidden=false; clearTimeout(toast.t); toast.t=setTimeout(()=>t.hidden=true,3200); }
function errMsg(e){ const s=e&&e.status; if(s===0) return "Ingen nettforbindelse. Endringen ble ikke lagret."; if(s===401) return "Innloggingen har gått ut. Last inn siden på nytt."; if(s===403) return "Du har ikke skrivetilgang til Toddum-området i SharePoint. Be Silje legge deg til i teamet."; return "Klarte ikke å lagre i SharePoint. Prøv igjen om litt."; }
function ukeParitet(w){ return w%2===0?"partall":"oddetall"; }
function slug(s){ return s.toLowerCase().replace(/[æ]/g,"ae").replace(/[ø]/g,"o").replace(/[å]/g,"a").replace(/[^a-z0-9]+/g,"-").replace(/^-|-$/g,"").slice(0,40); }

/* ---------- lagring ---------- */
function skriv(path, data){
  if(!S.db){ toast("Ordreboka er ikke koblet til SharePoint."); return Promise.resolve(); }
  const [k,id]=path.split("/");
  S.pending++; visSync();
  const p = data===null ? Store.slett(k,id) : Store.lagre(k, {...data, [k==="produkter"?"varenr":"id"]:id});
  return p.catch(e=>{ toast(errMsg(e)); throw e; }).finally(()=>{ S.pending--; visSync(); });
}
function lagreOrdre(o){
  o={...o, endret:new Date().toISOString()};
  S.override.set(o.id,o);
  const i=S.ordre.findIndex(x=>x.id===o.id); if(i>=0) S.ordre[i]=o; else S.ordre.push(o);
  planRender();
  const {id,...data}=o;
  skriv("ordre/"+id, data).then(()=>{ if(S.override.get(id)===o) S.override.delete(id); }).catch(()=>{ S.override.delete(id); });
}
function visSync(){ const el=$("#sync"); if(!el) return; const st = !S.db ? "" : S.pending? "Lagrer …" : "Lagret i SharePoint"; el.innerHTML = S.bruker ? `<span>${esc(st)}</span><span class="who">${esc(S.bruker.navn)} · <button class="linkbtn" id="loggUt">Logg ut</button></span>` : esc(st); const lo=$("#loggUt"); if(lo) lo.onclick=()=>Store.loggUt(); }

/* ---------- oppstart ---------- */
async function start(){
  bindGlobalt();
  render();
  try{
    const st=await Store.init();
    S.bruker=st.bruker||null;
    if(st.tilstand==="ikke-konfigurert"){ S.fase="ikke-konfigurert"; render(); return; }
    if(st.tilstand==="logget-ut"){ S.fase="login"; render(); return; }
    if(st.tilstand==="oppsett" || location.hash==="#oppsett"){ S.fase="oppsett"; S.mangler=st.mangler||[]; render(); return; }
    await lastAlt();
  }catch(e){ console.error(e); S.fase="feil"; S.feil = e.status===403 ? "Kontoen din har ikke tilgang til Toddum-området i SharePoint. Be Silje legge deg til som medlem i teamet Toddum." : (e.message||String(e)); render(); }
}
async function lastAlt(){
  S.fase="klar"; S.db=true; render();
  const [p,k]=await Promise.all([Store.hent("produkter"),Store.hent("kunder")]);
  S.produkter=p; S.kunder=k.sort((a,b)=>a.navn.localeCompare(b.navn,"nb")); S.loaded.p=S.loaded.k=true;
  await hentUke();
  if(!S.pollT){ S.pollT=setInterval(poll,20000); document.addEventListener("visibilitychange",()=>{ if(!document.hidden) poll(); }); }
}
async function hentUke(){
  const key=uke().key;
  try{
    const liste=await Store.hent("ordre",key);
    if(key!==uke().key) return;
    for(const [id,o] of S.override){ if(o.uke!==key) continue; const i=liste.findIndex(x=>x.id===id); if(i>=0) liste[i]=o; else liste.push(o); }
    const slettet=S.slettet||new Set();
    S.ordre=liste.filter(o=>!slettet.has(o.id)); S.loaded.o=true; planRender();
  }catch(e){ S.loaded.o=true; planRender(); toast(errMsg(e)); }
}
async function poll(){
  if(document.hidden || S.fase!=="klar" || S.pending) return;
  S.pollN++;
  await hentUke();
  if(S.pollN%6===0 && S.tab!=="kunder" && S.tab!=="produkter"){
    try{ const [p,k]=await Promise.all([Store.hent("produkter"),Store.hent("kunder")]); S.produkter=p; S.kunder=k.sort((a,b)=>a.navn.localeCompare(b.navn,"nb")); planRender(); }catch(_){}
  }
}
function abonnerUke(){ if(!S.db) return; S.loaded.o=false; S.ordre=[]; hentUke(); }

/* ---------- rendering ---------- */
let renderT=null;
function planRender(){ if(renderT) return; renderT=setTimeout(()=>{ renderT=null; render(); },0); }
function render(){
  const ae=document.activeElement, fid=ae&&ae.id, fval=ae&&("value" in ae)?ae.value:null, fsel=ae&&ae.selectionStart!=null?[ae.selectionStart,ae.selectionEnd]:null;
  const u=uke(), m=mandag(S.ukeDato), s=new Date(m); s.setDate(s.getDate()+6);
  $("#weekLbl").innerHTML=`<b>Uke ${u.w}</b><span>${fmtD(m)}–${fmtD(s)}.${s.getFullYear()}</span>`;
  document.querySelectorAll(".tab").forEach(t=>t.setAttribute("aria-selected", String(t.dataset.tab===S.tab)));
  const main=$("#main");
  if(S.fase!=="klar" && S.fase!=="start") main.innerHTML=visFase();
  else if(!(S.loaded.p&&S.loaded.k)) { /* behold lastevisning */ }
  else if(S.tab==="uke") main.innerHTML=visUke();
  else if(S.tab==="ysteri") main.innerHTML=visYsteri();
  else if(S.tab==="kunder") main.innerHTML=visKunder();
  else main.innerHTML=visProdukter();
  if(fid){ const el=document.getElementById(fid); if(el && el!==document.activeElement){ if(fval!=null && el.value!==fval && el.dataset.keep!=="0") el.value=fval; el.focus(); try{ if(fsel) el.setSelectionRange(fsel[0],fsel[1]); }catch(_){} } }
  visSync();
}

function ikkeTilkoblet(){ return ""; }
function visFase(){
  const logo=`<img src="assets/logo.jpg" alt="Toddum Gård Gårdsysteri" class="biglogo">`;
  if(S.fase==="login") return `<div class="center">${logo}<h2>Ordrebok</h2><p class="sub">Logg inn med Toddum-kontoen din (navn@toddum.com).</p><button class="btn primary" data-act="loggInn">Logg inn med Microsoft</button></div>`;
  if(S.fase==="ikke-konfigurert") return `<div class="center">${logo}<h2>Ordreboka er ikke koblet til Microsoft 365 ennå</h2><p class="sub">Fyll inn klient-ID fra appregistreringen i Entra i <code>config.js</code>.</p></div>`;
  if(S.fase==="feil") return `<div class="center">${logo}<h2>Ordreboka fikk ikke startet</h2><p class="sub">${esc(S.feil)}</p><button class="btn" data-act="lastPaNytt">Prøv igjen</button></div>`;
  if(S.fase==="oppsett"){
    const navn={produkter:"Ordrebok Produkter",kunder:"Ordrebok Kunder",ordre:"Ordrebok Ordre"};
    return `<div class="card form" style="max-width:640px;margin:0 auto"><h3>Oppsett av ordreboka</h3>
      <p class="hint">Ordreboka lagrer i tre lister i SharePoint-området Toddum. Dette gjøres bare én gang.</p>
      <h4>1. Lister i SharePoint</h4>
      ${S.mangler.length?`<p>Disse listene mangler: ${S.mangler.map(k=>`<b>${navn[k]}</b>`).join(", ")}.</p><div class="actions"><button class="btn primary" data-act="opprettLister">Opprett listene</button></div>`:`<p>Alle listene finnes.</p>`}
      <h4>2. Importer kunder, produkter og ordre</h4>
      <p class="hint">Velg eksportfila fra den gamle ordreboka (JSON). Poster som allerede finnes, hoppes over, så du kan kjøre importen flere ganger.</p>
      <input type="file" id="importFil" accept=".json,application/json" ${S.mangler.length?"disabled":""}>
      ${S.importStatus?`<p class="notice">${esc(S.importStatus)}</p>`:""}
      <h4>3. Ferdig</h4>
      <div class="actions"><button class="btn primary" data-act="ferdigOppsett" ${S.mangler.length?"disabled":""}>Åpne ordreboka</button><button class="btn" data-act="eksporter" ${S.mangler.length?"disabled":""}>Last ned sikkerhetskopi (JSON)</button></div></div>`;
  }
  return "";
}

function forslagFaste(){
  const u=uke(), har=new Set(S.ordre.map(o=>o.kundeId));
  return S.kunder.filter(k=>!k.inaktiv && k.fast && Object.keys(k.fast).length && (k.fastFrekvens==="hver" || k.fastFrekvens===ukeParitet(u.w)) && !har.has(k.id));
}

function visUke(){
  const u=uke();
  const ordre=[...S.ordre].sort((a,b)=>(a.kundeNavn||"").localeCompare(b.kundeNavn||"","nb"));
  const faste=forslagFaste();
  let totKg=0, totStk=0; ordre.forEach(o=>linjer(o).forEach(l=>{ totKg+=kgNetto(l.antall,l.p); totStk+=l.antall; }));
  const sug = S.leggTil.trim() ? kundeTreff(S.leggTil).slice(0,8) : [];
  return `${ikkeTilkoblet()}
  <div class="toolbar">
    <div class="grow">
      <input class="txt" id="leggTil" placeholder="Legg til ordre: skriv kundenavn …" autocomplete="off" value="${esc(S.leggTil)}" data-act="leggTilSok" aria-label="Legg til ordre for kunde">
      ${sug.length?`<div class="suggest">${sug.map((k,i)=>`<button data-act="nyOrdre" data-id="${esc(k.id)}" class="${i===S.sugIdx?"act":""}"><span>${esc(k.navn)}</span><small>${esc(k.poststed||"")}</small></button>`).join("")}</div>`:""}
    </div>
    <button class="btn" data-act="hentFaste" ${faste.length?"":"disabled"} title="${faste.map(k=>esc(k.navn)).join(", ")}">Hent faste ordre${faste.length?` (${faste.length})`:""}</button>
    <button class="btn" data-act="csv" ${ordre.length?"":"disabled"}>Fakturagrunnlag (CSV)</button>
  </div>
  <div class="stats"><span><b>${ordre.length}</b> ordre</span><span><b>${nf(totStk)}</b> enheter</span><span><b>${nf(totKg,1)}</b> kg ost (netto)</span><span>${STATUS.map(([k,l])=>`${l}: <b>${ordre.filter(o=>(o.status||"ny")===k).length}</b>`).join(" · ")}</span></div>
  ${!S.loaded.o?`<div class="empty"><h3>Henter ordre for uke ${u.w} …</h3></div>`:
    ordre.length? `<div class="orders">${ordre.map(visOrdre).join("")}</div>` :
    `<div class="empty"><h3>Ingen ordre i uke ${u.w} ennå</h3>Skriv et kundenavn i feltet over for å legge inn første ordre${faste.length?`, eller trykk «Hent faste ordre» for å legge inn ${faste.length} faste ordre på én gang`:""}.</div>`}`;
}

function visOrdre(o){
  const k=kunde(o.kundeId), åpen=S.open.has(o.id), ls=linjer(o), st=o.status||"ny";
  const kg=ls.reduce((s,l)=>s+kgNetto(l.antall,l.p),0);
  const under=[o.transport, o.sendes?`sendes ${dmy(o.sendes)}`:"", o.bestnr?`best.nr ${o.bestnr}`:""].filter(Boolean).join(" · ");
  let h=`<article class="card" id="o-${esc(o.id)}">
  <div class="ohead" data-act="toggle" data-id="${esc(o.id)}">
    <div class="name">${esc(o.kundeNavn||k?.navn||"Ukjent kunde")}<small>${esc(under||"Ingen transport valgt")}${o.kilde==="fast"?" · fast ordre":""}</small></div>
    <span class="hint">${nf(kg,1)} kg</span>
    <span class="pill st-${st}">${esc((STATUS.find(s=>s[0]===st)||[,st])[1])}</span>
    <span aria-hidden="true">${åpen?"▴":"▾"}</span>
  </div>`;
  if(!åpen){
    h+= ls.length? `<div class="summary">${ls.map(l=>`<span class="chip g-${esc(l.p?.gruppe||"")}">${nf(l.antall)} × ${esc(l.p?.kort||l.varenr)}</span>`).join("")}</div>` : `<div class="summary"><span class="hint">Ingen varer lagt inn</span></div>`;
    return h+`</article>`;
  }
  const batchStd=batchFra(o.prodDato);
  h+=`<div class="obody">
    <div class="lines">${prodSort().map(p=>{ const l=(o.linjer||{})[p.varenr]; const a=l?.antall||0;
      return `<div class="line bl-${esc(p.gruppe)} ${a?"has":""}">
        <div class="pn">${esc(p.kort)}<small>#${esc(p.varenr)} · ${esc(p.pakning)}${a?` · ${esc(dpTekst(a,p))}`:""}</small></div>
        <input id="q-${esc(o.id)}-${esc(p.varenr)}" inputmode="decimal" data-act="qty" data-id="${esc(o.id)}" data-v="${esc(p.varenr)}" value="${a||""}" placeholder="–" aria-label="Antall ${esc(p.kort)}" title="Antall ${esc(p.enhet||"stk")}. Skriv f.eks. 4dp for 4 detaljistpakninger">
        ${a?`<div class="bt">Batch <input id="b-${esc(o.id)}-${esc(p.varenr)}" data-act="batch" data-id="${esc(o.id)}" data-v="${esc(p.varenr)}" value="${esc(l.batch||"")}" placeholder="${esc(batchStd||"–")}" aria-label="Batch ${esc(p.kort)}"></div>`:""}
      </div>`; }).join("")}</div>
    <p class="hint" style="margin:-8px 0 0">Antall er i ${"beger/spann"}. Skriv <b>4dp</b> for fire detaljistpakninger. Batch fylles automatisk fra produksjonsdato (DDMMÅÅ × 2) hvis feltet står tomt.</p>
    <div class="meta">
      <label class="f">Transport<select class="txt" id="t-${esc(o.id)}" data-act="felt" data-f="transport" data-id="${esc(o.id)}"><option value="">Velg …</option>${TRANSPORT.map(t=>`<option ${o.transport===t?"selected":""}>${esc(t)}</option>`).join("")}</select></label>
      <label class="f">Sendes<input type="date" class="txt" id="s-${esc(o.id)}" data-act="felt" data-f="sendes" data-id="${esc(o.id)}" value="${esc(o.sendes||"")}"></label>
      <label class="f">Produksjonsdato<input type="date" class="txt" id="p-${esc(o.id)}" data-act="felt" data-f="prodDato" data-id="${esc(o.id)}" value="${esc(o.prodDato||"")}"></label>
      <label class="f">Bestillingsnr.<input class="txt" id="n-${esc(o.id)}" data-act="felt" data-f="bestnr" data-id="${esc(o.id)}" value="${esc(o.bestnr||"")}"></label>
      <label class="f">Status<select class="txt" id="st-${esc(o.id)}" data-act="felt" data-f="status" data-id="${esc(o.id)}">${STATUS.map(([v,l])=>`<option value="${v}" ${st===v?"selected":""}>${l}</option>`).join("")}</select></label>
      <label class="f wide">Merknad til ysteriet<input class="txt" id="m-${esc(o.id)}" data-act="felt" data-f="merknad" data-id="${esc(o.id)}" value="${esc(o.merknad||"")}" placeholder="${esc(k?.merknad||"")}"></label>
    </div>
    <div class="actions">
      <button class="btn primary" data-act="pakkseddel" data-id="${esc(o.id)}" ${ls.length?"":"disabled"}>Pakkseddel (PDF)</button>
      ${trengerCMR(o)?`<button class="btn primary" data-act="cmr" data-id="${esc(o.id)}" ${ls.length?"":"disabled"}>CMR × 3 (PDF)</button>`:`<button class="btn" data-act="cmr" data-id="${esc(o.id)}" ${ls.length?"":"disabled"}>CMR (PDF)</button>`}
      <select class="txt" style="width:auto" id="ff-${esc(o.id)}" aria-label="Frekvens for fast ordre">${FREKVENS.slice(1).map(([v,l])=>`<option value="${v}" ${(k?.fastFrekvens===v)?"selected":""}>${l}</option>`).join("")}</select>
      <button class="btn" data-act="somFast" data-id="${esc(o.id)}" ${ls.length&&k?"":"disabled"}>Lagre som fast ordre</button>
      <span class="sp"></span>
      <button class="btn danger ghost" data-act="slett" data-id="${esc(o.id)}">${S.confirmDel===o.id?"Trykk igjen for å slette":"Slett ordre"}</button>
    </div>
  </div></article>`;
  return h;
}

function kundeTreff(q){
  const t=q.toLowerCase().trim(); if(!t) return [];
  return S.kunder.filter(k=>!k.inaktiv && (k.navn.toLowerCase().includes(t) || (k.kundenr||"").includes(t) || (k.poststed||"").toLowerCase().includes(t)))
    .sort((a,b)=>(a.navn.toLowerCase().startsWith(t)?0:1)-(b.navn.toLowerCase().startsWith(t)?0:1)||a.navn.localeCompare(b.navn,"nb"));
}

function ysteriData(){
  const ordre=[...S.ordre].filter(o=>linjer(o).length).sort((a,b)=>(a.sendes||"9").localeCompare(b.sendes||"9")||(a.kundeNavn||"").localeCompare(b.kundeNavn||"","nb"));
  const brukt=new Set(); ordre.forEach(o=>linjer(o).forEach(l=>brukt.add(l.varenr)));
  const kol=prodAlle().filter(p=>brukt.has(p.varenr));
  const sum={}; ordre.forEach(o=>linjer(o).forEach(l=>sum[l.varenr]=(sum[l.varenr]||0)+l.antall));
  return {ordre,kol,sum};
}
function kommentar(o){ const k=kunde(o.kundeId); return [o.sendes?`Sendes ${dmy(o.sendes)}`:"", o.transport, o.merknad||k?.merknad, o.bestnr?`Best.nr ${o.bestnr}`:""].filter(Boolean).join(" · "); }

function visYsteri(){
  const u=uke(), {ordre,kol,sum}=ysteriData();
  if(!S.loaded.o) return `<div class="empty"><h3>Henter ordre …</h3></div>`;
  const totKg=kol.reduce((s,p)=>s+kgNetto(sum[p.varenr],p),0);
  return `${ikkeTilkoblet()}<h2>Ysteriliste uke ${u.w}</h2>
  <p class="sub">Det som skal produseres og pakkes denne uka, per kunde og totalt. Last ned som PDF til utskrift, eller kopier som tekst til SMS eller e-post.</p>
  <div class="toolbar"><button class="btn primary" data-act="ysteriPdf" ${ordre.length?"":"disabled"}>Last ned PDF</button><button class="btn" data-act="ysteriTekst" ${ordre.length?"":"disabled"}>Kopier som tekst</button></div>
  ${!ordre.length?`<div class="empty"><h3>Ingen varer bestilt i uke ${u.w}</h3>Legg inn ordre under «Ukeordre», så kommer de med her.</div>`:
  `<div class="tablewrap"><table class="grid"><thead><tr><th>Kunde</th>${kol.map(p=>`<th class="n"><span class="chip g-${esc(p.gruppe)}">${esc(p.kort)}</span></th>`).join("")}<th>Kommentar</th></tr></thead>
  <tbody>${ordre.map(o=>`<tr><td><b>${esc(o.kundeNavn)}</b></td>${kol.map(p=>{const l=(o.linjer||{})[p.varenr]; return `<td class="n">${l&&l.antall?nf(l.antall):""}</td>`;}).join("")}<td class="hint">${esc(kommentar(o))}</td></tr>`).join("")}</tbody>
  <tfoot><tr><td>Sum enheter</td>${kol.map(p=>`<td class="n">${nf(sum[p.varenr])}</td>`).join("")}<td></td></tr>
  <tr><td>Detaljistpakninger</td>${kol.map(p=>`<td class="n">${esc(dpTekst(sum[p.varenr],p))}</td>`).join("")}<td></td></tr>
  <tr><td>Kg ost (netto)</td>${kol.map(p=>`<td class="n">${nf(kgNetto(sum[p.varenr],p),1)}</td>`).join("")}<td>Totalt ${nf(totKg,1)} kg</td></tr></tfoot></table></div>`}`;
}

function visKunder(){
  const q=S.kundeSok.toLowerCase();
  const liste=S.kunder.filter(k=>!q || k.navn.toLowerCase().includes(q) || (k.kundenr||"").includes(q) || (k.poststed||"").toLowerCase().includes(q));
  const k = S.kundeSel==="__ny" ? {id:"",navn:"",kundenr:"",orgnr:"",adresse:"",adresse2:"",postnr:"",poststed:"",epost:"",telefon:"",levNavn:"",levAdresse:"",transport:"",cmr:false,leveringsdag:"",merknad:"",instruks:"",fast:{},fastFrekvens:"ingen",inaktiv:false} : kunde(S.kundeSel);
  return `${ikkeTilkoblet()}<h2>Kunder</h2><p class="sub">${S.kunder.length} kunder. Leveringsadresse, transport og CMR-instruks brukes på pakkseddel og CMR.</p>
  <div class="two">
    <div class="card"><div style="padding:10px;display:flex;gap:8px"><input class="txt" id="kundeSok" data-act="kundeSok" placeholder="Søk kunde …" value="${esc(S.kundeSok)}" aria-label="Søk kunde"><button class="btn" data-act="nyKunde">Ny kunde</button></div>
      <div class="klist">${liste.map(x=>`<button data-act="velgKunde" data-id="${esc(x.id)}" class="${x.id===S.kundeSel?"sel":""}"><span>${esc(x.navn)}${x.inaktiv?" <small>(inaktiv)</small>":""}</span><small>${esc([x.transport, x.fastFrekvens&&x.fastFrekvens!=="ingen"?"fast":""].filter(Boolean).join(" · ")||x.poststed||"")}</small></button>`).join("")||`<p class="hint" style="padding:12px">Ingen treff.</p>`}</div></div>
    <div class="card">${k?kundeSkjema(k):`<div class="form"><p class="hint">Velg en kunde i lista for å se og endre leveringsadresse, transport, CMR-instruks og fast ordre.</p></div>`}</div>
  </div>`;
}
function kundeSkjema(k){
  const fi=(f,l,extra="")=>`<label class="f">${l}<input class="txt" id="kf-${f}" data-f="${f}" value="${esc(k[f]||"")}" ${extra}></label>`;
  const fast=Object.entries(k.fast||{}).filter(([,a])=>a>0);
  return `<form class="form" id="kundeForm" data-id="${esc(k.id)}">
    <h3>${esc(k.navn||"Ny kunde")}</h3>
    <h4>Faktura</h4>
    <div class="row">${fi("navn","Navn","required")}${fi("kundenr","Kundenr.")}${fi("orgnr","Org.nr.")}${fi("epost","E-post")}${fi("telefon","Telefon")}</div>
    <div class="row">${fi("adresse","Adresse")}${fi("adresse2","Adresse 2")}${fi("postnr","Postnr.")}${fi("poststed","Poststed")}</div>
    <h4>Levering</h4>
    <div class="row">${fi("levNavn","Leveres til (navn)","placeholder='Samme som kunde'")}<label class="f">Leveringsadresse<input class="txt" id="kf-levAdresse" data-f="levAdresse" value="${esc(k.levAdresse||"")}" placeholder="Gate, postnr. sted"></label>
      <label class="f">Transport<select class="txt" id="kf-transport" data-f="transport"><option value="">Ikke valgt</option>${TRANSPORT.map(t=>`<option ${k.transport===t?"selected":""}>${esc(t)}</option>`).join("")}</select></label>${fi("leveringsdag","Fast leveringsdag","placeholder='f.eks. tirsdag'")}</div>
    <label class="chk"><input type="checkbox" id="kf-cmr" data-f="cmr" ${k.cmr?"checked":""}> Trenger CMR</label>
    <label class="f">Instruks på CMR<input class="txt" id="kf-instruks" data-f="instruks" value="${esc(k.instruks||"")}" placeholder="f.eks. Kjølevare - behandles forsiktig"></label>
    <label class="f">Fast merknad til ysteriet<input class="txt" id="kf-merknad" data-f="merknad" value="${esc(k.merknad||"")}"></label>
    <h4>Fast ordre</h4>
    <div class="row"><label class="f">Hvor ofte<select class="txt" id="kf-fastFrekvens" data-f="fastFrekvens">${FREKVENS.map(([v,l])=>`<option value="${v}" ${(k.fastFrekvens||"ingen")===v?"selected":""}>${l}</option>`).join("")}</select></label></div>
    <div class="summary" style="padding:0">${fast.length?fast.map(([v,a])=>{const p=prod(v);return `<span class="chip g-${esc(p?.gruppe||"")}">${nf(a)} × ${esc(p?.kort||v)}</span>`;}).join(""):`<span class="hint">Ingen faste varer. Legg inn en ordre for kunden og trykk «Lagre som fast ordre».</span>`}</div>
    <label class="chk"><input type="checkbox" id="kf-inaktiv" data-f="inaktiv" ${k.inaktiv?"checked":""}> Inaktiv kunde</label>
    <div class="actions"><button class="btn primary" type="submit">Lagre kunde</button>${fast.length?`<button class="btn" type="button" data-act="fjernFast">Fjern faste varer</button>`:""}</div>
  </form>`;
}

function visProdukter(){
  const ps=prodAlle();
  return `${ikkeTilkoblet()}<h2>Produkter</h2><p class="sub">Hentet fra EPD-eksporten. Vekt per detaljistpakning (DP) brukes til bruttovekt på CMR, nettovekt til kg-summen i ysterilista. Endringer lagres når du forlater feltet.</p>
  <div class="toolbar"><button class="btn" data-act="nyttProdukt">Nytt produkt</button></div>
  ${S.nyttProdukt?`<form class="card form" id="nyttProduktForm" style="margin-bottom:14px"><div class="row">
    <label class="f">Varenr.<input class="txt" id="np-varenr" required placeholder="f.eks. 0040"></label>
    <label class="f">Kort navn<input class="txt" id="np-kort" required placeholder="Stracciatella 1 kg"></label>
    <label class="f">Gruppe<select class="txt" id="np-gruppe">${["125g","10g","Burrata","Stracciatella","Revet"].map(g=>`<option>${g}</option>`).join("")}</select></label>
    <label class="f">Pakning (DP)<input class="txt" id="np-pakning" placeholder="2 x 1 kg"></label>
    <label class="f">Enheter per DP<input class="txt" id="np-perDP" inputmode="numeric" value="2"></label>
    <label class="f">Brutto per DP (g)<input class="txt" id="np-dpBruttoG" inputmode="numeric"></label>
    <label class="f">Netto per enhet (g)<input class="txt" id="np-stkNettoG" inputmode="numeric"></label>
  </div><div class="actions"><button class="btn primary" type="submit">Legg til produkt</button><button class="btn ghost" type="button" data-act="avbrytProdukt">Avbryt</button></div></form>`:""}
  <div class="tablewrap"><table class="grid"><thead><tr><th>Varenr.</th><th>Kort navn</th><th>Gruppe</th><th>Pakning</th><th class="n">Per DP</th><th class="n">Brutto DP (g)</th><th class="n">Netto enhet (g)</th><th>GTIN</th><th>Aktiv</th></tr></thead>
  <tbody>${ps.map(p=>`<tr>
    <td>#${esc(p.varenr)}</td>
    <td><input class="txt" id="pr-${esc(p.varenr)}-kort" data-act="prod" data-v="${esc(p.varenr)}" data-f="kort" value="${esc(p.kort)}" style="min-width:180px"></td>
    <td><span class="chip g-${esc(p.gruppe)}">${esc(p.gruppe)}</span></td>
    <td><input class="txt" id="pr-${esc(p.varenr)}-pakning" data-act="prod" data-v="${esc(p.varenr)}" data-f="pakning" value="${esc(p.pakning)}" style="width:110px"></td>
    <td class="n"><input class="txt" id="pr-${esc(p.varenr)}-perDP" data-act="prod" data-v="${esc(p.varenr)}" data-f="perDP" data-num="1" value="${esc(p.perDP)}" style="width:60px;text-align:right"></td>
    <td class="n"><input class="txt" id="pr-${esc(p.varenr)}-dpBruttoG" data-act="prod" data-v="${esc(p.varenr)}" data-f="dpBruttoG" data-num="1" value="${esc(p.dpBruttoG)}" style="width:90px;text-align:right"></td>
    <td class="n"><input class="txt" id="pr-${esc(p.varenr)}-stkNettoG" data-act="prod" data-v="${esc(p.varenr)}" data-f="stkNettoG" data-num="1" value="${esc(p.stkNettoG)}" style="width:80px;text-align:right"></td>
    <td class="hint">${esc(p.gtin||"")}</td>
    <td><input type="checkbox" id="pr-${esc(p.varenr)}-aktiv" data-act="prodAktiv" data-v="${esc(p.varenr)}" ${p.aktiv!==false?"checked":""} aria-label="Aktiv"></td></tr>`).join("")}</tbody></table></div>`;
}

/* ---------- handlinger ---------- */
function bindGlobalt(){
  $("#prevWeek").onclick=()=>byttUke(-7);
  $("#nextWeek").onclick=()=>byttUke(7);
  $("#thisWeek").onclick=()=>{ S.ukeDato=new Date(); S.open.clear(); abonnerUke(); render(); };
  document.querySelectorAll(".tab").forEach(t=>t.onclick=()=>{ S.tab=t.dataset.tab; render(); });
  const main=$("#main");
  main.addEventListener("click", klikk);
  main.addEventListener("change", endring);
  main.addEventListener("input", e=>{
    const a=e.target.dataset.act;
    if(a==="leggTilSok"){ S.leggTil=e.target.value; S.sugIdx=0; render(); }
    if(a==="kundeSok"){ S.kundeSok=e.target.value; render(); }
  });
  main.addEventListener("keydown", e=>{
    if(e.target.id==="leggTil"){
      const sug=kundeTreff(S.leggTil).slice(0,8);
      if(e.key==="ArrowDown"){ S.sugIdx=Math.min(S.sugIdx+1,sug.length-1); render(); e.preventDefault(); }
      else if(e.key==="ArrowUp"){ S.sugIdx=Math.max(S.sugIdx-1,0); render(); e.preventDefault(); }
      else if(e.key==="Enter" && sug[S.sugIdx]){ nyOrdre(sug[S.sugIdx].id); e.preventDefault(); }
      else if(e.key==="Escape"){ S.leggTil=""; render(); }
    }
    if(e.key==="Enter" && (e.target.dataset.act==="qty"||e.target.dataset.act==="batch")){ e.target.blur(); }
  });
  main.addEventListener("submit", e=>{
    e.preventDefault();
    if(e.target.id==="kundeForm") lagreKunde(e.target);
    if(e.target.id==="nyttProduktForm") leggTilProdukt();
  });
}
function byttUke(dager){ const d=new Date(S.ukeDato); d.setDate(d.getDate()+dager); S.ukeDato=d; S.open.clear(); abonnerUke(); render(); }

function klikk(e){
  const el=e.target.closest("[data-act]"); if(!el) return;
  const a=el.dataset.act, id=el.dataset.id;
  if(a==="toggle"){ if(e.target.closest("select,input,button:not([data-act=toggle])")) return; S.open.has(id)?S.open.delete(id):S.open.add(id); S.confirmDel=null; render(); }
  else if(a==="nyOrdre") nyOrdre(id);
  else if(a==="hentFaste") hentFaste();
  else if(a==="csv") lastNedCSV();
  else if(a==="pakkseddel"){ forberedVindu(); pakkseddel(finnOrdre(id)); }
  else if(a==="cmr"){ forberedVindu(); cmr(finnOrdre(id)); }
  else if(a==="somFast") somFast(finnOrdre(id));
  else if(a==="slett"){ if(S.confirmDel===id){ S.confirmDel=null; (S.slettet=S.slettet||new Set()).add(id); S.ordre=S.ordre.filter(o=>o.id!==id); S.override.delete(id); render(); skriv("ordre/"+id,null).then(()=>toast("Ordren er slettet.")).catch(()=>{}); } else { S.confirmDel=id; render(); setTimeout(()=>{ if(S.confirmDel===id){ S.confirmDel=null; render(); } },4000); } }
  else if(a==="ysteriPdf"){ forberedVindu(); ysteriPdf(); }
  else if(a==="ysteriTekst") ysteriTekst();
  else if(a==="velgKunde"){ S.kundeSel=id; render(); }
  else if(a==="nyKunde"){ S.kundeSel="__ny"; render(); }
  else if(a==="fjernFast"){ const k=kunde(S.kundeSel); if(k){ const n={...k,fast:{},fastFrekvens:"ingen"}; oppdaterKunde(n).then(()=>toast(`Fast ordre for ${k.navn} er fjernet.`)); } }
  else if(a==="nyttProdukt"){ S.nyttProdukt=true; render(); }
  else if(a==="loggInn") Store.loggInn();
  else if(a==="lastPaNytt") location.reload();
  else if(a==="opprettLister"){ el.disabled=true; el.textContent="Oppretter …"; Store.opprettLister().then(()=>{ S.mangler=[]; render(); toast("Listene er opprettet i SharePoint."); }).catch(e=>{ toast("Klarte ikke å opprette listene: "+(e.message||e)); render(); }); }
  else if(a==="ferdigOppsett"){ history.replaceState(null,"",location.pathname); lastAlt(); }
  else if(a==="eksporter"){ Store.eksporter().then(d=>lagre(`Ordrebok sikkerhetskopi ${iDag()}.json`, JSON.stringify(d,null,1))).catch(e=>toast(errMsg(e))); }
  else if(a==="avbrytProdukt"){ S.nyttProdukt=false; render(); }
}
const finnOrdre = id => S.ordre.find(o=>o.id===id);

function parseAntall(s,p){
  s=String(s).trim().toLowerCase().replace(",",".");
  if(!s) return 0;
  const m=s.match(/^(\d+(?:\.\d+)?)\s*(dp|d)?$/);
  if(!m) return null;
  let n=parseFloat(m[1]); if(m[2]) n*= (p?.perDP||1);
  return Math.round(n*10)/10;
}
function endring(e){
  const t=e.target, a=t.dataset.act, id=t.dataset.id;
  if(t.id==="importFil" && t.files && t.files[0]){
    const fil=t.files[0];
    fil.text().then(txt=>{ const d=JSON.parse(txt); S.importStatus="Importerer …"; render();
      return Store.importer(d,(n,tot)=>{ S.importStatus=`Importerer ${n} av ${tot} …`; planRender(); }); })
     .then(r=>{ S.importStatus=`Ferdig. ${r.ferdig} poster lagt inn, ${r.hoppetOver} fantes fra før.`; render(); })
     .catch(err=>{ S.importStatus="Importen stoppet: "+(err.message||err); render(); });
    return;
  }
  if(a==="qty"){
    const o=finnOrdre(id); if(!o) return; const p=prod(t.dataset.v); const n=parseAntall(t.value,p);
    if(n===null){ toast("Skriv et tall, eller f.eks. 4dp for fire detaljistpakninger."); t.value=(o.linjer?.[t.dataset.v]?.antall)||""; return; }
    const lin={...(o.linjer||{})};
    if(n>0) lin[t.dataset.v]={...(lin[t.dataset.v]||{}), antall:n}; else delete lin[t.dataset.v];
    t.dataset.keep="0"; t.value=n||"";
    lagreOrdre({...o, linjer:lin});
  } else if(a==="batch"){
    const o=finnOrdre(id); if(!o) return; const lin={...(o.linjer||{})}; const v=t.dataset.v;
    if(lin[v]) lin[v]={...lin[v], batch:t.value.trim()}; lagreOrdre({...o, linjer:lin});
  } else if(a==="felt"){
    const o=finnOrdre(id); if(!o) return; lagreOrdre({...o, [t.dataset.f]:t.value});
  } else if(a==="prod"){
    const p=prod(t.dataset.v); if(!p) return; let v=t.value.trim(); if(t.dataset.num){ v=Number(v.replace(",",".")); if(!isFinite(v)){ toast("Skriv et tall."); return; } }
    const n={...p,[t.dataset.f]:v}; S.produkter=S.produkter.map(x=>x.varenr===p.varenr?n:x); const {varenr,...d}=n; skriv("produkter/"+varenr,d);
  } else if(a==="prodAktiv"){
    const p=prod(t.dataset.v); if(!p) return; const n={...p,aktiv:t.checked}; S.produkter=S.produkter.map(x=>x.varenr===p.varenr?n:x); const {varenr,...d}=n; skriv("produkter/"+varenr,d);
  }
}

function nyOrdre(kundeId){
  const k=kunde(kundeId); if(!k) return;
  const finnes=S.ordre.find(o=>o.kundeId===kundeId);
  S.leggTil="";
  if(finnes){ S.open.add(finnes.id); render(); document.getElementById("o-"+finnes.id)?.scrollIntoView({block:"start",behavior:"smooth"}); toast(`${k.navn} har allerede en ordre denne uka.`); return; }
  const u=uke(); const id=`${u.key}-${slug(k.navn)||"kunde"}-${Math.random().toString(36).slice(2,6)}`;
  const o={id, uke:u.key, kundeId:k.id, kundeNavn:k.navn, linjer:{}, transport:k.transport||"", sendes:"", prodDato:"", bestnr:"", merknad:"", status:"ny", kilde:"manuell", opprettet:new Date().toISOString()};
  S.open.add(id); lagreOrdre(o);
  setTimeout(()=>{ const first=document.querySelector(`#o-${CSS.escape(id)} .line input`); first&&first.focus(); },50);
}
function hentFaste(){
  const faste=forslagFaste(); const u=uke();
  faste.forEach(k=>{
    const lin={}; Object.entries(k.fast).forEach(([v,a])=>{ if(a>0) lin[v]={antall:a}; });
    const id=`${u.key}-${slug(k.navn)||"kunde"}-${Math.random().toString(36).slice(2,6)}`;
    lagreOrdre({id, uke:u.key, kundeId:k.id, kundeNavn:k.navn, linjer:lin, transport:k.transport||"", sendes:"", prodDato:"", bestnr:"", merknad:"", status:"ny", kilde:"fast", opprettet:new Date().toISOString()});
  });
  toast(`${faste.length} faste ordre er lagt inn for uke ${u.w}.`);
}
function somFast(o){
  const k=kunde(o.kundeId); if(!k) return; const fr=document.getElementById("ff-"+o.id)?.value||"hver";
  const fast={}; linjer(o).forEach(l=>fast[l.varenr]=l.antall);
  oppdaterKunde({...k, fast, fastFrekvens:fr}).then(()=>toast(`Lagret som fast ordre for ${k.navn} (${FREKVENS.find(f=>f[0]===fr)[1].toLowerCase()}).`));
}
function oppdaterKunde(k){ S.kunder=S.kunder.map(x=>x.id===k.id?k:x); planRender(); const {id,...d}=k; return skriv("kunder/"+id,d); }
function lagreKunde(form){
  const gammel = form.dataset.id ? kunde(form.dataset.id) : null;
  const k = gammel ? {...gammel} : {fast:{},fastFrekvens:"ingen"};
  form.querySelectorAll("[data-f]").forEach(el=>{ k[el.dataset.f] = el.type==="checkbox" ? el.checked : el.value.trim(); });
  if(!k.navn){ toast("Kunden må ha et navn."); return; }
  if(!gammel){ k.id = k.kundenr || ("k-"+slug(k.navn)+"-"+Math.random().toString(36).slice(2,5)); if(kunde(k.id)){ toast("Det finnes allerede en kunde med dette kundenummeret."); return; } S.kunder=[...S.kunder,k].sort((a,b)=>a.navn.localeCompare(b.navn,"nb")); S.kundeSel=k.id; }
  oppdaterKunde(k).then(()=>toast(`${k.navn} er lagret.`));
}
function leggTilProdukt(){
  const g=id=>document.getElementById(id).value.trim();
  const varenr=g("np-varenr").replace(/^#/,"");
  if(!/^[A-Za-z0-9_-]+$/.test(varenr)){ toast("Varenummeret kan bare inneholde tall og bokstaver."); return; }
  if(prod(varenr)){ toast("Det finnes allerede et produkt med dette varenummeret."); return; }
  const p={varenr, kort:g("np-kort"), navn:g("np-kort"), gruppe:g("np-gruppe"), pakning:g("np-pakning"), perDP:Number(g("np-perDP"))||1, dpBruttoG:Number(g("np-dpBruttoG"))||0, stkNettoG:Number(g("np-stkNettoG"))||0, stkBruttoG:0, enhet:"stk", varetype:"Fersk ost", sort:50, aktiv:true};
  S.produkter=[...S.produkter,p]; S.nyttProdukt=false; render(); const {varenr:v,...d}=p; skriv("produkter/"+v,d).then(()=>toast(`${p.kort} er lagt til.`));
}

/* ---------- PDF og filer ---------- */
function forberedVindu(){ try{ S.pdfVindu=window.open("","_blank"); if(S.pdfVindu) S.pdfVindu.document.write("<p style=\"font-family:sans-serif;padding:20px\">Lager PDF …</p>"); }catch(_){ S.pdfVindu=null; } }
async function lagre(filename, data){
  const type = filename.endsWith(".pdf") ? "application/pdf" : filename.endsWith(".csv") ? "text/csv;charset=utf-8" : "application/json";
  const blob = data instanceof Blob ? data : new Blob([data],{type});
  const url = URL.createObjectURL(blob);
  const w = S.pdfVindu; S.pdfVindu=null;
  if(w && !w.closed && filename.endsWith(".pdf")){ w.location.href=url; setTimeout(()=>URL.revokeObjectURL(url),120000); return; }
  if(w && !w.closed) w.close();
  const a=document.createElement("a"); a.href=url; a.download=filename; document.body.appendChild(a); a.click(); a.remove();
  setTimeout(()=>URL.revokeObjectURL(url),120000);
}
async function asset(sti){ const r=await fetch(sti); if(!r.ok) throw new Error("Fant ikke "+sti); return new Uint8Array(await r.arrayBuffer()); }
function safe(font, s){ s=String(s??"").replace(/[–—]/g,"-").replace(/[‘’]/g,"'").replace(/[“”]/g,'"').replace(/[​\r]/g,""); let out=""; for(const ch of s){ if(ch==="\n"){out+=ch;continue;} try{ font.encodeText(ch); out+=ch; }catch(_){ out+="?"; } } return out; }
function wrap(font,size,text,maxW){ const words=String(text).split(/\s+/); const lines=[]; let cur=""; for(const w of words){ const t=cur?cur+" "+w:w; if(font.widthOfTextAtSize(t,size)<=maxW) cur=t; else { if(cur) lines.push(cur); cur=w; while(font.widthOfTextAtSize(cur,size)>maxW && cur.length>1){ let i=cur.length-1; while(i>1 && font.widthOfTextAtSize(cur.slice(0,i),size)>maxW) i--; lines.push(cur.slice(0,i)); cur=cur.slice(i); } } } if(cur) lines.push(cur); return lines.length?lines:[""]; }
function filnavn(s){ return s.replace(/[\\/:*?"<>|]/g,"").replace(/\s+/g," ").trim(); }

async function pakkseddel(o){
  if(!o) return;
  const {PDFDocument,StandardFonts,rgb}=PDFLib;
  const doc=await PDFDocument.create(); const page=doc.addPage([595.28,841.89]);
  const f=await doc.embedFont(StandardFonts.Helvetica), fb=await doc.embedFont(StandardFonts.HelveticaBold);
  const logo=await doc.embedJpg(await asset("assets/logo.jpg"));
  const k=kunde(o.kundeId); const ink=rgb(0.17,0.16,0.15), grey=rgb(0.45,0.43,0.41);
  const T=(s,x,y,size=10,font=f,color=ink)=>page.drawText(safe(font,s),{x,y,size,font,color});
  T("Retur:",50,800,8); T(AVSENDER.navn,50,790,8); T(AVSENDER.adresse,50,780,8); T(`${AVSENDER.postnr} ${AVSENDER.poststed}`,50,770,8);
  page.drawImage(logo,{x:262,y:745,width:72,height:72});
  adresseLinjer(k,o).forEach((l,i)=>T(l,50,700-i*14,11));
  page.drawRectangle({x:330,y:690,width:215,height:24,color:rgb(0.95,0.95,0.94)});
  T("Dato",338,698,10,fb); T(dmyLang(iDag()),440,698,10,fb);
  const rows=[["Bestillingsnr.",o.bestnr||""],["Uke",String(uke().w)],["Vår ref.",AVSENDER.ref]].filter(r=>r[1]);
  rows.forEach((r,i)=>{ T(r[0],338,672-i*14,10); T(r[1],440,672-i*14,10); });
  T("PAKKSEDDEL",50,600,13,fb);
  const ls=linjer(o); const batchStd=batchFra(o.prodDato);
  const batcher=[...new Set(ls.map(l=>l.batch||batchStd).filter(Boolean))];
  let y=578;
  if(o.prodDato){ T(`Prod: ${dmy(o.prodDato)}`,50,y,10); y-=13; }
  if(o.sendes){ T(`Sendt: ${dmy(o.sendes)}`,50,y,10); y-=13; }
  if(batcher.length===1){ T(`Batch: ${batcher[0]}`,50,y,10); y-=13; }
  y-=18;
  const flere=batcher.length>1;
  T("Beskrivelse",50,y,9,f,grey); if(flere) T("Batch",400,y,9,f,grey); T("DP",455,y,9,f,grey); const aw=f.widthOfTextAtSize("Antall",9); T("Antall",545-aw,y,9,f,grey);
  y-=7; page.drawLine({start:{x:50,y},end:{x:545,y},thickness:0.6,color:rgb(0.8,0.79,0.77)}); y-=18;
  for(const l of ls){
    const txt=`#${l.varenr} - ${l.p?.kort||l.varenr}${l.p?.pakning?` (${l.p.pakning})`:""}`;
    T(txt,50,y,10); if(flere) T(l.batch||batchStd,400,y,10); T(dpTekst(l.antall,l.p),455,y,10);
    const s=nf(l.antall); T(s,545-f.widthOfTextAtSize(s,10),y,10); y-=20;
  }
  const kg=ls.reduce((s,l)=>s+kgBrutto(l.antall,l.p),0);
  y-=4; page.drawLine({start:{x:50,y:y+12},end:{x:545,y:y+12},thickness:0.4,color:rgb(0.85,0.84,0.82)});
  T(`Bruttovekt ca. ${nf(kg,1)} kg`,50,y,9,f,grey);
  const foot=`${AVSENDER.navn} | Foretaksregisteret NO ${AVSENDER.orgnr} MVA | ${AVSENDER.epost} | Tlf. ${AVSENDER.tlf}`;
  const fw=f.widthOfTextAtSize(foot,8); page.drawText(safe(f,foot),{x:(595.28-fw)/2,y:40,size:8,font:f,color:ink});
  const bytes=await doc.save();
  await lagre(filnavn(`Pakkseddel ${o.kundeNavn} ${dmy(o.sendes||iDag())}.pdf`), new Blob([bytes],{type:"application/pdf"}));
}

async function cmr(o){
  if(!o) return;
  const {PDFDocument,StandardFonts}=PDFLib;
  const k=kunde(o.kundeId); const ls=linjer(o); const batchStd=batchFra(o.prodDato);
  const src=await PDFDocument.load(await asset("assets/cmr.pdf"));
  const f=await src.embedFont(StandardFonts.Helvetica);
  const form=src.getForm();
  form.getFields().forEach(fl=>{ try{ fl.setText&&fl.setText(""); }catch(_){} });
  const set=(name,val,size)=>{ try{ const tf=form.getTextField(name); tf.setText(safe(f,val)); if(size) tf.setFontSize(size); }catch(err){ console.warn("CMR-felt",name,err); } };
  const adr=adresseLinjer(k,o); const adrTekst=adr[0]+"\n"+adr.slice(1).join(", ");
  set("Sender",`${AVSENDER.navn}\n${AVSENDER.adresse}, ${AVSENDER.postnr} Namsos - NORWAY`,9);
  set("Consignee",adrTekst,9);
  set("Place",adrTekst,9);
  set("Carrier",TRANSPORT_ADR[o.transport]||o.transport||"",9);
  const dato=o.sendes||iDag();
  set("Place and date",`${AVSENDER.hentested}, ${dmy(dato)}`,9);
  const col=(fn)=>ls.map(fn).join("\n");
  const totKg=ls.reduce((s,l)=>s+kgBrutto(l.antall,l.p),0);
  set("Marks and Nos", col(l=>l.p?.kort||l.varenr)+"\n\nSum",7.5);
  set("Number", col(l=>{ const d=dpTekst(l.antall,l.p); return d||`${nf(l.antall)} ${({kg:"kg",liter:"l"})[l.p?.enhet]||"stk"}`; }),7.5);
  set("Method", col(l=>l.p?.pakning||""),7.5);
  set("Nature", col(l=>l.p?.varetype||"Fersk ost"),7.5);
  set("Statistical nr", col(l=>l.batch||batchStd||""),7.5);
  set("Weight", col(l=>nf(kgBrutto(l.antall,l.p),1))+"\n\n"+nf(totKg,1),7.5);
  const instr=[k?.instruks, o.bestnr?`Ordre: ${o.bestnr}`:"", o.merknad].filter(Boolean).join("\n");
  set("Instructions",instr,9);
  set("Established in","Namsos",9); set("on",dmy(dato),9);
  form.updateFieldAppearances(f);
  form.flatten();
  const ut=await PDFDocument.create();
  const sider=await ut.copyPages(src,[0,0,0]); sider.forEach(p=>ut.addPage(p));
  const bytes=await ut.save();
  await lagre(filnavn(`CMR ${o.kundeNavn} ${dmy(dato)}.pdf`), new Blob([bytes],{type:"application/pdf"}));
}

async function ysteriPdf(){
  const {PDFDocument,StandardFonts,rgb}=PDFLib; const u=uke(); const {ordre,kol,sum}=ysteriData();
  const doc=await PDFDocument.create(); const f=await doc.embedFont(StandardFonts.Helvetica), fb=await doc.embedFont(StandardFonts.HelveticaBold);
  const W=841.89,H=595.28, M=32; const ink=rgb(0.17,0.16,0.15), grey=rgb(0.45,0.43,0.41), band=rgb(0.95,0.93,0.88);
  const kundeW=150, komW=200; const cw=Math.max(40,(W-2*M-kundeW-komW)/Math.max(1,kol.length));
  let page, y;
  const ny=()=>{ page=doc.addPage([W,H]); const m=mandag(S.ukeDato), s=new Date(m); s.setDate(s.getDate()+6);
    page.drawText(safe(fb,`Ysteriliste uke ${u.w}`),{x:M,y:H-M-6,size:16,font:fb,color:ink});
    page.drawText(safe(f,`${fmtD(m)}–${fmtD(s)}.${s.getFullYear()} · skrevet ut ${dmyLang(iDag())}`),{x:M,y:H-M-22,size:9,font:f,color:grey});
    y=H-M-46;
    page.drawRectangle({x:M,y:y-22,width:W-2*M,height:30,color:band});
    page.drawText("Kunde",{x:M+4,y:y-4,size:8,font:fb,color:ink});
    kol.forEach((p,i)=>{ const lines=wrap(fb,7,p.kort,cw-6).slice(0,3); lines.forEach((ln,j)=>{ const w=fb.widthOfTextAtSize(safe(fb,ln),7); page.drawText(safe(fb,ln),{x:M+kundeW+i*cw+cw-4-w,y:y-1-j*8,size:7,font:fb,color:ink}); }); });
    page.drawText("Kommentar",{x:M+kundeW+kol.length*cw+6,y:y-4,size:8,font:fb,color:ink});
    y-=34; };
  ny();
  const rad=(navn,vals,kom,bold=false)=>{
    const fn=bold?fb:f; const komL=wrap(f,7.5,safe(f,kom||""),komW-10).slice(0,3); const nL=wrap(fn,9,safe(fn,navn),kundeW-8).slice(0,2);
    const h=Math.max(16, 10*Math.max(komL.length,nL.length)+6);
    if(y-h<M+10) ny();
    nL.forEach((l,j)=>page.drawText(l,{x:M+4,y:y-j*10,size:9,font:fn,color:ink}));
    vals.forEach((v,i)=>{ if(!v) return; const s=safe(fn,v); const w=fn.widthOfTextAtSize(s,10); page.drawText(s,{x:M+kundeW+i*cw+cw-4-w,y,size:10,font:fn,color:ink}); });
    komL.forEach((l,j)=>page.drawText(l,{x:M+kundeW+kol.length*cw+6,y:y-j*9,size:7.5,font:f,color:grey}));
    y-=h; page.drawLine({start:{x:M,y:y+10},end:{x:W-M,y:y+10},thickness:0.4,color:rgb(0.85,0.84,0.82)});
  };
  ordre.forEach(o=>rad(o.kundeNavn, kol.map(p=>{const l=(o.linjer||{})[p.varenr]; return l&&l.antall?nf(l.antall):"";}), kommentar(o)));
  y-=4;
  rad("Sum enheter", kol.map(p=>nf(sum[p.varenr])), "", true);
  rad("Detaljistpakninger", kol.map(p=>dpTekst(sum[p.varenr],p)), "", true);
  const tot=kol.reduce((s,p)=>s+kgNetto(sum[p.varenr],p),0);
  rad("Kg ost (netto)", kol.map(p=>nf(kgNetto(sum[p.varenr],p),1)), `Totalt ${nf(tot,1)} kg`, true);
  const bytes=await doc.save();
  await lagre(`Ysteriliste uke ${u.w}.pdf`, new Blob([bytes],{type:"application/pdf"}));
}
function ysteriTekst(){
  const u=uke(), {ordre,kol,sum}=ysteriData();
  let t=`YSTERILISTE UKE ${u.w}\n\nSUM\n`+kol.map(p=>`${p.kort}: ${sum[p.varenr]}${dpTekst(sum[p.varenr],p)?` (${dpTekst(sum[p.varenr],p)})`:""}`).join("\n");
  t+="\n\nPER KUNDE\n"+ordre.map(o=>`${o.kundeNavn}: `+linjer(o).map(l=>`${l.antall} ${l.p?.kort||l.varenr}`).join(", ")+(kommentar(o)?`\n  ${kommentar(o)}`:"")).join("\n");
  const ok=()=>toast("Ysterilista er kopiert. Lim den inn i SMS eller e-post.");
  try{ navigator.clipboard.writeText(t).then(ok,()=>fallbackKopi(t)); }catch(_){ fallbackKopi(t); }
}
function fallbackKopi(t){ const ta=document.createElement("textarea"); ta.value=t; ta.className="txt"; ta.style.height="240px"; ta.readOnly=true; $("#main").prepend(ta); ta.select(); toast("Teksten er markert. Kopier den med Ctrl+C eller ⌘C."); }
function lastNedCSV(){
  const u=uke(); const rader=[["Uke","Kundenr","Kunde","Orgnr","Sendes","Bestillingsnr","Varenr","Vare","Antall","Status"]];
  [...S.ordre].sort((a,b)=>(a.kundeNavn||"").localeCompare(b.kundeNavn||"","nb")).forEach(o=>{ const k=kunde(o.kundeId); linjer(o).forEach(l=>rader.push([u.w,k?.kundenr||"",o.kundeNavn,k?.orgnr||"",dmyLang(o.sendes),o.bestnr||"",l.varenr,l.p?.kort||"",l.antall,o.status||"ny"])); });
  const csv="﻿"+rader.map(r=>r.map(c=>{ c=String(c??""); return /[;"\n]/.test(c)?`"${c.replace(/"/g,'""')}"`:c; }).join(";")).join("\r\n");
  lagre(`Fakturagrunnlag uke ${u.w}.csv`, csv);
}

start();
window.addEventListener("unhandledrejection",e=>{ if(S.pdfVindu){ try{S.pdfVindu.close();}catch(_){} S.pdfVindu=null; } });
if("serviceWorker" in navigator) window.addEventListener("load",()=>navigator.serviceWorker.register("sw.js").catch(()=>{}));
