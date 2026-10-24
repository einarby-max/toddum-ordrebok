
const AVSENDER = { navn:"Toddum Gårdsysteri AS", adresse:"Gullvikvegen 193", postnr:"7810", poststed:"NAMSOS", orgnr:"929953088", epost:"silje@toddum.com", tlf:"94307328", ref:"Silje Kvaale By", hentested:"Spillum - Namsos" };
/* Startliste for transportører. Legges i SharePoint første gang; deretter redigeres de under fanen Transport. */
const TRANSPORT_START = [
  {navn:"Kristensen Transport", adresse:"Kristensen Transport AS\nUtvordvegen 850, 7777 Nord-Statland", cmr:true},
  {navn:"Nær Engros Namsos"}, {navn:"Namdalingen"}, {navn:"Henter selv"}, {navn:"Egen levering"}, {navn:"Post/Bring"}, {navn:"Annet"} ];
const STATUS = [["ny","Ny"],["produsert","Produsert"],["sendt","Sendt"],["fakturert","Fakturert"]];
const FREKVENS = [["ingen","Ingen fast ordre"],["hver","Hver uke"],["partall","Partallsuker"],["oddetall","Oddetallsuker"]];

const S = { db:null, fase:"start", bruker:null, feil:"", mangler:[], importStatus:"", pdfVindu:null, pollN:0, produkter:[], kunder:[], ordre:[], transport:[], visDP:new Map(), wordImport:null, tEdit:null, utkast:new Map(), lagret:new Map(), siste:new Map(), kunLes:false, tab:"uke", ukeDato:new Date(), open:new Set(), override:new Map(), loaded:{p:false,k:false,o:false}, kundeSel:null, kundeSok:"", leggTil:"", sugIdx:0, pending:0, unsubOrdre:null, confirmDel:null, nyttProdukt:false };

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
const transportør = navn => S.transport.find(t=>t.navn===navn);
const transportNavn = () => S.transport.filter(t=>t.aktiv!==false).sort((a,b)=>(a.sort||0)-(b.sort||0)||a.navn.localeCompare(b.navn,"nb")).map(t=>t.navn);
function transportValg(valgt){ const n=transportNavn(); if(valgt && !n.includes(valgt)) n.push(valgt); return n.map(t=>`<option ${valgt===t?"selected":""}>${esc(t)}</option>`).join(""); }
function transportAdresse(navn){ const t=transportør(navn); return (t&&t.adresse) ? t.adresse : (navn||""); }
function trengerCMR(o){ const k=kunde(o.kundeId); return !!(transportør(o.transport)?.cmr || (k&&k.cmr)); }
/* Antall lagres alltid i salgsenheten (beger, spann, kg, liter). Kunder som bestiller i D-pak, vises og justeres i D-pak. */
function iDP(o,p){ if(!p || !(p.perDP>2)) return false;   /* spann (2 per DP) telles alltid i spann */ const v=S.visDP.get(o.id); if(v!=null) return v; return kunde(o.kundeId)?.antallI==="dp"; }
const enhetNavn = (p,n) => ({spann:"spann",kg:"kg",liter:"l",pose:n===1?"pose":"poser",beger:"beger"})[p?.enhet] || (/spann/i.test(p?.kort||"")?"spann":"stk");
function antallTekst(a,p){ if(!a) return ""; const e=enhetNavn(p,a); if(!p || !(p.perDP>2)) return `${nf(a)} ${e}`; const d=dpTekst(a,p); return d&&!/^\d+ DP$/.test(d) ? `${nf(a)} ${e} = ${d.replace(/\+ (\d+)$/,"+ $1 løse")}` : `${nf(a)} ${e}${d?` = ${d}`:""}`; }
function toast(msg){ const t=$("#toast"); t.textContent=msg; t.hidden=false; clearTimeout(toast.t); toast.t=setTimeout(()=>t.hidden=true,3200); }
function errMsg(e){ const s=e&&e.status; if(s===0) return "Ingen nettforbindelse. Endringen ble ikke lagret."; if(s===401) return "Innloggingen har gått ut. Last inn siden på nytt."; if(s===403) return "Du har bare lesetilgang til ordreboka. Endringen ble ikke lagret."; return "Klarte ikke å lagre i SharePoint. Prøv igjen om litt."; }
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
  if(S.kunLes){ toast("Du har bare lesetilgang til ordreboka."); return Promise.resolve(); }
  o={...o, endret:new Date().toISOString()};
  S.override.set(o.id,o);
  const i=S.ordre.findIndex(x=>x.id===o.id); if(i>=0) S.ordre[i]=o; else S.ordre.push(o);
  S.lagret.set(o.id,{st:"lagrer"});
  planRender();
  const {id,...data}=o;
  const p=skriv("ordre/"+id, data).then(()=>{ if(S.override.get(id)===o) S.override.delete(id); if(S.siste.get(id)===p) S.lagret.set(id,{st:"ok",tid:new Date()}); planRender(); })
    .catch(e=>{ S.override.delete(id); S.lagret.set(id,{st:"feil"}); if(e&&e.status===403) settKunLes(); planRender(); throw e; });
  S.siste.set(id,p); p.catch(()=>{});
  return p;
}
function settKunLes(){ if(S.kunLes) return; S.kunLes=true; render(); }
async function lagreKnapp(id){
  const ae=document.activeElement; if(ae && ae.blur && ae.closest && ae.closest(`#o-${CSS.escape(id)}`)) ae.blur();   // utløser lagring av feltet man står i
  await new Promise(r=>setTimeout(r,60));
  const p=S.siste.get(id); const o=finnOrdre(id);
  try{ if(p) await p; S.lagret.set(id,{st:"ok",tid:new Date()}); S.open.delete(id); render(); toast(`✓ Ordren til ${o?.kundeNavn||"kunden"} er lagret.`); }
  catch(e){ toast(errMsg(e)+" Prøv å trykke Lagre igjen."); render(); }
}
function lagretTekst(id){ const s=S.lagret.get(id); if(!s) return `<span class="lst ok">✓ Lagret</span>`; if(s.st==="lagrer") return `<span class="lst">Lagrer …</span>`; if(s.st==="feil") return `<span class="lst feil">Ikke lagret – trykk Lagre</span>`; return `<span class="lst ok">✓ Lagret kl. ${pad(s.tid.getHours())}.${pad(s.tid.getMinutes())}</span>`; }
function visSync(){ const el=$("#sync"); if(!el) return; const st = !S.db ? "" : S.pending? "Lagrer …" : "Lagret i SharePoint"; el.innerHTML = S.bruker ? `<span>${esc(st)}</span><span class="who">${esc(S.bruker.navn)} · <button class="linkbtn" id="loggUt">Logg ut</button></span>` : esc(st); const lo=$("#loggUt"); if(lo) lo.onclick=()=>Store.loggUt(); }

/* ---------- oppstart ---------- */
async function start(){
  bindGlobalt();
  render();
  try{
    const st=await Store.init();
    S.bruker=st.bruker||null;
    /* Gjester (private e-postadresser) har bare lesetilgang i SharePoint. Vis det fra start. */
    if(S.bruker && S.bruker.epost && !/@toddum\.(com|onmicrosoft\.com)$/i.test(S.bruker.epost)) S.kunLes=true;
    if(st.tilstand==="ikke-konfigurert"){ S.fase="ikke-konfigurert"; render(); return; }
    if(st.tilstand==="logget-ut"){ S.fase="login"; render(); return; }
    if(st.tilstand==="oppsett" || location.hash==="#oppsett"){ S.fase="oppsett"; S.mangler=st.mangler||[]; render(); return; }
    await lastAlt();
  }catch(e){ console.error(e); S.fase="feil"; S.feil = e.status===403 ? "Kontoen din har ikke tilgang til Toddum-området i SharePoint. Be Silje legge deg til som medlem i teamet Toddum." : (e.message||String(e)); render(); }
}
async function lastAlt(){
  S.fase="klar"; S.db=true; render();
  const [p,k,t]=await Promise.all([Store.hent("produkter"),Store.hent("kunder"),Store.hent("transport").catch(()=>[])]);
  S.produkter=p; S.kunder=k.sort((a,b)=>a.navn.localeCompare(b.navn,"nb")); S.loaded.p=S.loaded.k=true;
  settTransport(t);
  await hentUke();
  if(!S.pollT){ S.pollT=setInterval(poll,20000); document.addEventListener("visibilitychange",()=>{ if(!document.hidden) poll(); }); }
}
function settTransport(liste){
  if(liste.length || !Store.harListe("transport")){ S.transport = liste.length ? liste : TRANSPORT_START.map((t,i)=>({id:slug(t.navn),adresse:"",kontakt:"",telefon:"",epost:"",cmr:false,merknad:"",aktiv:true,sort:i,...t})); return; }
  /* Tom liste i SharePoint: legg inn startlista én gang. */
  S.transport=TRANSPORT_START.map((t,i)=>({id:slug(t.navn),adresse:"",kontakt:"",telefon:"",epost:"",cmr:false,merknad:"",aktiv:true,sort:i,...t}));
  S.transport.forEach(lagreTransport);
}
function lagreTransport(t){ const i=S.transport.findIndex(x=>x.id===t.id); if(i>=0) S.transport[i]=t; else S.transport.push(t); planRender(); const {id,...d}=t; return skriv("transport/"+id,d); }
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
    try{ const [p,k]=await Promise.all([Store.hent("produkter"),Store.hent("kunder")]); S.produkter=p; S.kunder=k.sort((a,b)=>a.navn.localeCompare(b.navn,"nb")); if(S.tab!=="transport"){ const t=await Store.hent("transport"); if(t.length) S.transport=t; } planRender(); }catch(_){}
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
  taVarePåUtkast();
  main.classList.toggle("ro", !!S.kunLes);
  if(S.fase!=="klar" && S.fase!=="start") main.innerHTML=visFase();
  else if(!(S.loaded.p&&S.loaded.k)) { /* behold lastevisning */ }
  else if(S.tab==="uke") main.innerHTML=visUke();
  else if(S.tab==="ysteri") main.innerHTML=visYsteri();
  else if(S.tab==="kunder") main.innerHTML=visKunder();
  else if(S.tab==="transport") main.innerHTML=visTransport();
  else main.innerHTML=visProdukter();
  gjenopprettUtkast();
  if(fid){ const el=document.getElementById(fid); if(el && el!==document.activeElement){ if(fval!=null && el.value!==fval && el.dataset.keep!=="0") el.value=fval; el.focus(); try{ if(fsel) el.setSelectionRange(fsel[0],fsel[1]); }catch(_){} } }
  visSync();
}

/* Skjemaer merket data-bevar beholder det som er skrevet, men ikke lagret, når siden tegnes på nytt
   (appen henter nye ordre hvert 20. sekund og tegner da hele siden). */
function taVarePåUtkast(){
  document.querySelectorAll("form[data-bevar]").forEach(fm=>{
    const key=fm.id+"|"+(fm.dataset.id||""), m={};
    fm.querySelectorAll("input,textarea,select").forEach(el=>{
      if(!el.id) return;
      if(el.type==="checkbox"){ if(el.checked!==el.defaultChecked) m[el.id]={c:el.checked}; }
      else if(el.tagName==="SELECT"){ const def=([...el.options].find(o=>o.defaultSelected)||el.options[0]||{}).value; if(el.value!==def) m[el.id]={v:el.value}; }
      else if(el.value!==el.defaultValue) m[el.id]={v:el.value};
    });
    if(Object.keys(m).length) S.utkast.set(key,m); else S.utkast.delete(key);
  });
}
function gjenopprettUtkast(){
  document.querySelectorAll("form[data-bevar]").forEach(fm=>{
    const m=S.utkast.get(fm.id+"|"+(fm.dataset.id||"")); if(!m) return;
    for(const [id,v] of Object.entries(m)){ const el=document.getElementById(id); if(!el||!fm.contains(el)) continue; if("c" in v) el.checked=v.c; else el.value=v.v; }
  });
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
      <div class="actions"><button class="btn primary" data-act="ferdigOppsett" ${S.mangler.length?"disabled":""}>Åpne ordreboka</button><button class="btn" data-act="eksporter" ${S.mangler.length?"disabled":""}>Last ned sikkerhetskopi (JSON)</button></div>
      <h4>Slett alle ordre</h4>
      <p class="hint">Sletter alle ordre i alle uker. Kunder og produkter beholdes. Last ned en sikkerhetskopi først, og skriv SLETT i feltet for å bekrefte.</p>
      <div class="actions"><input class="txt" id="slettBekreft" placeholder="Skriv SLETT" style="max-width:160px" autocomplete="off"><button class="btn danger" data-act="slettAlle" ${S.backupTatt?"":"disabled"}>Slett alle ordre</button></div></div>`;
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
    <button class="btn" data-act="kopierForrige" title="Lager ordre for kundene fra forrige uke med samme varer og antall. Kunder som allerede har ordre denne uka, hoppes over.">Kopier forrige uke</button>
    <label class="btn" title="Last opp ukelista fra Word (.docx). Du ser forslaget før noe lagres.">Importer Word-liste<input type="file" id="wordFil" accept=".docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document" hidden></label>
    <button class="btn" data-act="unimicro" ${ordre.length?"":"disabled"} title="Lager en Excel-fil i Unimicros importmal for ordre. Ordre med status Fakturert tas ikke med.">Ordre til Unimicro (Excel)</button>
    <button class="btn danger ghost" data-act="tomUke" ${ordre.length?"":"disabled"}>${S.confirmTom?`Trykk igjen for å slette ${ordre.length} ordre`:"Tøm uka"}</button>
  </div>
  ${S.kunLes?`<p class="notice"><b>Lesetilgang.</b> Du kan se ukas ordre og ysterilista, men ikke endre noe.</p>`:""}
  ${S.wordImport?visWordImport():""}
  <div class="stats"><span><b>${ordre.length}</b> ordre</span><span><b>${nf(totStk)}</b> enheter</span><span><b>${nf(totKg,1)}</b> kg ost (netto)</span><span>${STATUS.map(([k,l])=>`${l}: <b>${ordre.filter(o=>(o.status||"ny")===k).length}</b>`).join(" · ")}</span></div>
  ${!S.loaded.o?`<div class="empty"><h3>Henter ordre for uke ${u.w} …</h3></div>`:
    ordre.length? `<div class="orders">${ordre.map(visOrdre).join("")}</div>` :
    `<div class="empty"><h3>Ingen ordre i uke ${u.w} ennå</h3>Skriv et kundenavn i feltet over for å legge inn første ordre${faste.length?`, eller trykk «Hent faste ordre» for å legge inn ${faste.length} faste ordre på én gang`:""}. Med «Kopier forrige uke» henter du alle ordrene fra uka før og justerer bare antallet.</div>`}`;
}

function visWordImport(){
  const w=S.wordImport, u=uke(), har=new Set(S.ordre.map(o=>o.kundeId));
  const aktive=S.kunder.filter(k=>!k.inaktiv);
  const klar=w.rader.filter(r=>r.med && r.kundeId && Object.keys(r.linjer).length && !har.has(r.kundeId));
  const chips=r=>Object.entries(r.linjer).map(([v,n])=>{ const p=prod(v); return `<span class="chip g-${esc(p?.gruppe||"")}">${esc(antallTekst(n,p))} ${esc(p?.kort||v)}</span>`; }).join("");
  return `<div class="card wimp" id="wordImport">
    <h3>Import fra Word: ${esc(w.filnavn)}</h3>
    <p class="hint">${w.uke&&w.uke!==u.w?`<b style="color:var(--warn)">Lista er merket uke ${w.uke}, men du står i uke ${u.w}.</b> Bytt uke øverst hvis ordrene skal inn i uke ${w.uke}. `:""}Sjekk kunde og varer for hver rad. Velg riktig kunde der den mangler. Det du velger, huskes til neste import. Kommentaren blir merknad til ysteriet. Kunder som allerede har ordre denne uka, hoppes over.</p>
    <div class="tablewrap"><table class="grid"><thead><tr><th>Med</th><th>I Word</th><th>Kunde i ordreboka</th><th>Varer (tolket)</th><th>Kommentar</th></tr></thead><tbody>
    ${w.rader.map((r,i)=>{ const finnes=r.kundeId&&har.has(r.kundeId); const hopp=!r.med||finnes||!Object.keys(r.linjer).length;
      return `<tr class="${hopp?"hopp":""}">
      <td><input type="checkbox" id="wi-med-${i}" data-act="wiMed" data-i="${i}" ${r.med?"checked":""} aria-label="Ta med ${esc(r.tekst)}"></td>
      <td><b>${esc(r.tekst)}</b><div class="hint">${Object.entries(r.celler).map(([g,t])=>`${esc(g)}: ${esc(t)}`).join("<br>")}</div></td>
      <td><select class="txt" id="wi-k-${i}" data-act="wiKunde" data-i="${i}" style="min-width:200px"><option value="">Velg kunde …</option>${aktive.map(k=>`<option value="${esc(k.id)}" ${k.id===r.kundeId?"selected":""}>${esc(k.navn)}</option>`).join("")}</select>${finnes?`<div class="hint">Har allerede ordre denne uka</div>`:""}</td>
      <td>${chips(r)||`<span class="hint">Ingen varer</span>`}${r.ukjent.length?`<div class="ukjent">${esc(r.ukjent.join(" · "))}</div>`:""}</td>
      <td class="hint">${esc(r.kommentar)}</td></tr>`; }).join("")}
    </tbody></table></div>
    <div class="actions"><button class="btn primary" data-act="wiLegg" ${klar.length?"":"disabled"}>Legg inn ${klar.length} ordre i uke ${u.w}</button><button class="btn ghost" data-act="wiAvbryt">Avbryt</button>
      <span class="hint">Tekst merket «Ikke tolket» må legges inn for hånd etterpå.</span></div>
  </div>`;
}
function leggInnWord(){
  const w=S.wordImport, u=uke(), har=new Set(S.ordre.map(o=>o.kundeId)); let n=0, delvis=0;
  for(const r of w.rader){
    if(!r.med || !r.kundeId || har.has(r.kundeId) || !Object.keys(r.linjer).length) continue;
    const k=kunde(r.kundeId); if(!k) continue;
    const lin={}; Object.entries(r.linjer).forEach(([v,a])=>lin[v]={antall:a});
    const id=`${u.key}-${slug(k.navn)||"kunde"}-${Math.random().toString(36).slice(2,6)}`;
    lagreOrdre({id, uke:u.key, kundeId:k.id, kundeNavn:k.navn, linjer:lin, transport:k.transport||"", sendes:"", prodDato:"", bestnr:"", merknad:[r.kommentar, r.ukjent.length?`Fra Word, ikke tolket: ${r.ukjent.join("; ")}`:""].filter(Boolean).join(" · "), status:"ny", kilde:"word", opprettet:new Date().toISOString()});
    har.add(k.id); n++; if(r.ukjent.length) delvis++;
  }
  S.wordImport=null; render();
  toast(`${n} ordre er lagt inn fra Word${delvis?`. ${delvis} har tekst som ikke ble tolket – se merknaden på ordren`:""}.`);
}

function visOrdre(o){
  const k=kunde(o.kundeId), åpen=S.open.has(o.id), ls=linjer(o), st=o.status||"ny";
  const kg=ls.reduce((s,l)=>s+kgNetto(l.antall,l.p),0);
  const under=[o.transport, o.sendes?`sendes ${dmy(o.sendes)}`:"", o.bestnr?`best.nr ${o.bestnr}`:""].filter(Boolean).join(" · ");
  let h=`<article class="card" id="o-${esc(o.id)}">
  <div class="ohead" data-act="toggle" data-id="${esc(o.id)}">
    <div class="name">${esc(o.kundeNavn||k?.navn||"Ukjent kunde")}<small>${esc(under||"Ingen transport valgt")}${o.kilde==="fast"?" · fast ordre":""}</small></div>
    ${!åpen && S.lagret.get(o.id)?.st==="ok" ? lagretTekst(o.id) : !åpen && S.lagret.get(o.id)?.st==="feil" ? lagretTekst(o.id) : ""}
    <span class="hint">${nf(kg,1)} kg</span>
    <span class="pill st-${st}">${esc((STATUS.find(s=>s[0]===st)||[,st])[1])}</span>
    <span aria-hidden="true">${åpen?"▴":"▾"}</span>
  </div>`;
  if(!åpen){
    h+= ls.length? `<div class="summary">${ls.map(l=>`<span class="chip g-${esc(l.p?.gruppe||"")}">${nf(l.antall)} × ${esc(l.p?.kort||l.varenr)}</span>`).join("")}</div>` : `<div class="summary"><span class="hint">Ingen varer lagt inn</span></div>`;
    return h+`</article>`;
  }
  const batchStd=batchFra(o.prodDato);
  const dpModus=iDP(o,{perDP:6}), harDP=prodSort().some(p=>p.perDP>2);
  h+=`<div class="obody">
    ${harDP?`<div class="modus"><span class="hint">Antall i</span><label class="seg"><input type="radio" name="dm-${esc(o.id)}" id="dm-${esc(o.id)}-stk" data-act="dpModus" data-id="${esc(o.id)}" value="stk" ${dpModus?"":"checked"}><span>Beger / stk</span></label><label class="seg"><input type="radio" name="dm-${esc(o.id)}" id="dm-${esc(o.id)}-dp" data-act="dpModus" data-id="${esc(o.id)}" value="dp" ${dpModus?"checked":""}><span>D-pak</span></label></div>`:""}
    <div class="lines">${prodSort().map(p=>{ const l=(o.linjer||{})[p.varenr]; const a=l?.antall||0; const idp=iDP(o,p);
      return `<div class="line bl-${esc(p.gruppe)} ${a?"has":""}">
        <div class="pn">${esc(p.kort)}<small>#${esc(p.varenr)} · ${esc(p.pakning)}</small>${a?`<small class="ant">${esc(antallTekst(a,p))}</small>`:""}</div>
        <div class="qty"><button type="button" class="stp" data-act="minus" data-id="${esc(o.id)}" data-v="${esc(p.varenr)}" aria-label="Én ${idp?"D-pak":"enhet"} mindre ${esc(p.kort)}" ${a?"":"disabled"}>−</button><input id="q-${esc(o.id)}-${esc(p.varenr)}" inputmode="decimal" data-act="qty" data-id="${esc(o.id)}" data-v="${esc(p.varenr)}" value="${esc(visAntall(a,o,p))}" placeholder="–" aria-label="Antall ${esc(p.kort)} i ${idp?"D-pak":enhetNavn(p,2)}" title="${idp?"Antall D-pak. 1,5 = 1 D-pak og 3 løse.":"Antall "+enhetNavn(p,2)+". Skriv 2dp for to D-pak."}"><button type="button" class="stp" data-act="pluss" data-id="${esc(o.id)}" data-v="${esc(p.varenr)}" aria-label="Én ${idp?"D-pak":"enhet"} mer ${esc(p.kort)}">+</button><span class="u">${idp?"D-pak":esc(enhetNavn(p,2))}</span></div>
        ${a?`<div class="bt">Batch <input id="b-${esc(o.id)}-${esc(p.varenr)}" data-act="batch" data-id="${esc(o.id)}" data-v="${esc(p.varenr)}" value="${esc(l.batch||"")}" placeholder="${esc(batchStd||"–")}" aria-label="Batch ${esc(p.kort)}"></div>`:""}
      </div>`; }).join("")}</div>
    <p class="hint" style="margin:-8px 0 0">Bruk − og + eller skriv antall. Med «Beger / stk» kan du også skrive <b>1,5dp</b> eller <b>1dp+3</b>. Med «D-pak» betyr 1,5 én D-pak og tre løse, og <b>3b</b> gir tre løse. Batch fylles automatisk fra produksjonsdato (DDMMÅÅ × 2) hvis feltet står tomt.</p>
    <div class="meta">
      <label class="f">Transport<select class="txt" id="t-${esc(o.id)}" data-act="felt" data-f="transport" data-id="${esc(o.id)}"><option value="">Velg …</option>${transportValg(o.transport)}</select></label>
      <label class="f">Sendes<input type="date" class="txt" id="s-${esc(o.id)}" data-act="felt" data-f="sendes" data-id="${esc(o.id)}" value="${esc(o.sendes||"")}"></label>
      <label class="f">Produksjonsdato<input type="date" class="txt" id="p-${esc(o.id)}" data-act="felt" data-f="prodDato" data-id="${esc(o.id)}" value="${esc(o.prodDato||"")}"></label>
      <label class="f">Bestillingsnr.<input class="txt" id="n-${esc(o.id)}" data-act="felt" data-f="bestnr" data-id="${esc(o.id)}" value="${esc(o.bestnr||"")}"></label>
      <label class="f">Status<select class="txt" id="st-${esc(o.id)}" data-act="felt" data-f="status" data-id="${esc(o.id)}">${STATUS.map(([v,l])=>`<option value="${v}" ${st===v?"selected":""}>${l}</option>`).join("")}</select></label>
      <label class="f wide">Merknad til ysteriet<input class="txt" id="m-${esc(o.id)}" data-act="felt" data-f="merknad" data-id="${esc(o.id)}" value="${esc(o.merknad||"")}" placeholder="${esc(k?.merknad||"")}"></label>
    </div>
    ${S.kunLes?"":`<div class="actions lagrerad"><button class="btn primary lagre" data-act="lagreOrdre" data-id="${esc(o.id)}">Lagre</button>${lagretTekst(o.id)}<span class="hint">Endringer lagres også av seg selv. «Lagre» bekrefter og lukker ordren.</span></div>`}
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
  <tbody>${ordre.map(o=>`<tr><td><b>${esc(o.kundeNavn)}</b></td>${kol.map(p=>{const l=(o.linjer||{})[p.varenr]; return `<td class="n">${l&&l.antall?nf(l.antall):""}</td>`;}).join("")}<td class="hint">${esc(kommentarFull(o))}</td></tr>`).join("")}</tbody>
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
  return `<form class="form" id="kundeForm" data-id="${esc(k.id)}" data-bevar="1">
    <h3>${esc(k.navn||"Ny kunde")}</h3>
    <h4>Faktura</h4>
    <div class="row">${fi("navn","Navn","required")}${fi("kundenr","Kundenr.")}${fi("orgnr","Org.nr.")}${fi("epost","E-post")}${fi("telefon","Telefon")}</div>
    <div class="row">${fi("adresse","Adresse")}${fi("adresse2","Adresse 2")}${fi("postnr","Postnr.")}${fi("poststed","Poststed")}</div>
    <label class="f">Kundepriser eks. mva (avvik fra listepris)<input class="txt" id="kf-priser" data-f="priser" value="${esc(priserTekst(k.priser))}" placeholder="f.eks. 0032=140; 0001=38"></label>
    <h4>Levering</h4>
    <div class="row">${fi("levNavn","Leveres til (navn)","placeholder='Samme som kunde'")}<label class="f">Leveringsadresse<input class="txt" id="kf-levAdresse" data-f="levAdresse" value="${esc(k.levAdresse||"")}" placeholder="Gate, postnr. sted"></label>
      <label class="f">Transport<select class="txt" id="kf-transport" data-f="transport"><option value="">Ikke valgt</option>${transportValg(k.transport)}</select></label>${fi("leveringsdag","Fast leveringsdag","placeholder='f.eks. tirsdag'")}
      <label class="f">Bestiller antall i<select class="txt" id="kf-antallI" data-f="antallI"><option value="" ${k.antallI==="dp"?"":"selected"}>Beger / stk (løse)</option><option value="dp" ${k.antallI==="dp"?"selected":""}>D-pak</option></select></label></div>
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

function nettLenke(u){ if(!u) return ""; const href=/^https?:\/\//i.test(u)?u:"https://"+u; return `<a href="${esc(href)}" target="_blank" rel="noopener">${esc(u.replace(/^https?:\/\//i,"").replace(/\/$/,""))}</a>`; }
function transportSkjema(t){
  const ny=!t.id; const fi=(f,l,extra="")=>`<label class="f">${l}<input class="txt" id="tf-${f}" name="${f}" value="${esc(t[f]||"")}" ${extra}></label>`;
  return `<form class="card form tform" id="transportForm" data-id="${esc(t.id||"__ny")}" data-bevar="1">
    <h3>${ny?"Ny transportør":esc(t.navn)}</h3>
    <div class="row">${fi("navn","Navn","required placeholder='f.eks. Bring Cargo'")}${fi("kontakt","Kontaktperson")}${fi("telefon","Telefon","inputmode='tel'")}${fi("epost","E-post","inputmode='email'")}${fi("nettside","Hjemmeside","inputmode='url' placeholder='f.eks. kristensentransport.no'")}</div>
    <label class="f">Adresse på CMR<textarea class="txt" id="tf-adresse" name="adresse" rows="2" placeholder="Firmanavn&#10;Gate, postnr. sted (står den tom, brukes bare navnet)">${esc(t.adresse||"")}</textarea></label>
    <label class="f">Merknad<textarea class="txt" id="tf-merknad" name="merknad" rows="2" placeholder="f.eks. hentetider, ruter, priser">${esc(t.merknad||"")}</textarea></label>
    <div class="actions"><label class="chk"><input type="checkbox" id="tf-cmr" name="cmr" ${t.cmr?"checked":""}> Krever CMR</label><label class="chk"><input type="checkbox" id="tf-aktiv" name="aktiv" ${t.aktiv!==false?"checked":""}> I bruk</label></div>
    <div class="actions"><button class="btn primary lagre" type="submit">Lagre</button><button class="btn ghost" type="button" data-act="avbrytTransport">Avbryt</button></div>
  </form>`;
}
function visTransport(){
  const liste=[...S.transport].sort((a,b)=>(b.aktiv!==false)-(a.aktiv!==false)||(a.sort||0)-(b.sort||0)||a.navn.localeCompare(b.navn,"nb"));
  const bruk=n=>S.kunder.filter(k=>k.transport===n&&!k.inaktiv).length;
  const nyT={navn:"",kontakt:"",telefon:"",epost:"",nettside:"",adresse:"",merknad:"",cmr:false,aktiv:true};
  return `<h2>Transport</h2><p class="sub">Transportørene her kan velges på ordre og kunder. Adressen brukes i feltet for transportør på CMR. Er «Krever CMR» krysset av, foreslås CMR × 3 for alle ordre med denne transportøren.</p>
  ${S.kunLes?"":`<div class="toolbar"><button class="btn" data-act="nyTransport" ${S.tEdit?"disabled":""}>Ny transportør</button></div>`}
  ${S.tEdit==="__ny"?transportSkjema(nyT):""}
  <div class="tlist">${liste.map(t=> S.tEdit===t.id ? transportSkjema(t) : `<div class="card tcard ${t.aktiv===false?"inaktiv":""}">
    <div class="thead"><div><b class="tnavn">${esc(t.navn)}</b>${t.cmr?` <span class="pill st-ny">CMR</span>`:""}${t.aktiv===false?` <span class="pill st-fakturert">Ikke i bruk</span>`:""}</div>
      ${S.kunLes?"":`<button class="btn sm" data-act="redigerTransport" data-id="${esc(t.id)}" ${S.tEdit?"disabled":""}>Rediger</button>`}</div>
    <div class="tinfo">
      ${[t.kontakt&&`<span>${esc(t.kontakt)}</span>`, t.telefon&&`<a href="tel:${esc(t.telefon.replace(/\s/g,""))}">${esc(t.telefon)}</a>`, t.epost&&`<a href="mailto:${esc(t.epost)}">${esc(t.epost)}</a>`, t.nettside&&nettLenke(t.nettside)].filter(Boolean).join(" · ")||`<span class="hint">Ingen kontaktinformasjon</span>`}
    </div>
    ${t.adresse?`<div class="hint">CMR: ${esc(t.adresse).replace(/\n/g,", ")}</div>`:""}
    ${t.merknad?`<div class="hint">${esc(t.merknad)}</div>`:""}
    <div class="hint">${bruk(t.navn)} kunder har denne som fast transport</div>
  </div>`).join("")}</div>`;
}
function lagreTransportSkjema(form){
  const v=n=>(form.elements[n]?.value||"").trim(), c=n=>!!form.elements[n]?.checked;
  const navn=v("navn"); if(!navn){ toast("Skriv navnet på transportøren."); return; }
  const gammel=form.dataset.id!=="__ny" ? S.transport.find(x=>x.id===form.dataset.id) : null;
  if(S.transport.some(x=>x!==gammel && x.navn.toLowerCase()===navn.toLowerCase())){ toast("Det finnes allerede en transportør med dette navnet."); return; }
  let id=gammel?.id; if(!id){ id=slug(navn)||"transport"; if(S.transport.some(x=>x.id===id)) id+="-"+Math.random().toString(36).slice(2,5); }
  const t={...(gammel||{sort:S.transport.length}), id, navn, kontakt:v("kontakt"), telefon:v("telefon"), epost:v("epost"), nettside:v("nettside"), adresse:v("adresse"), merknad:v("merknad"), cmr:c("cmr"), aktiv:c("aktiv")};
  S.tEdit=null; S.utkast.delete("transportForm|"+form.dataset.id);
  lagreTransport(t).then(()=>toast(`✓ ${navn} er lagret.${gammel&&gammel.navn!==navn?` Kunder og ordre med «${gammel.navn}» beholder det gamle navnet til de endres.`:""}`)).catch(()=>{});
}

function visProdukter(){
  const ps=prodAlle();
  return `${ikkeTilkoblet()}<h2>Produkter</h2><p class="sub">Hentet fra EPD-eksporten. Vekt per detaljistpakning (DP) brukes til bruttovekt på CMR, nettovekt til kg-summen i ysterilista. Prisen er listepris eks. mva per salgsenhet og brukes i ordrefila til Unimicro. Kundepriser legges på kunden. Endringer lagres når du forlater feltet.</p>
  <div class="toolbar"><button class="btn" data-act="nyttProdukt">Nytt produkt</button><span class="sp" style="flex:1"></span><button class="btn ghost" data-act="aapneOppsett">Oppsett, sikkerhetskopi og sletting</button></div>
  ${S.nyttProdukt?`<form class="card form" id="nyttProduktForm" style="margin-bottom:14px"><div class="row">
    <label class="f">Varenr.<input class="txt" id="np-varenr" required placeholder="f.eks. 0040"></label>
    <label class="f">Kort navn<input class="txt" id="np-kort" required placeholder="Stracciatella 1 kg"></label>
    <label class="f">Pris eks. mva<input class="txt" id="np-pris" inputmode="decimal"></label>
    <label class="f">Gruppe<select class="txt" id="np-gruppe">${["125g","10g","Burrata","Stracciatella","Revet"].map(g=>`<option>${g}</option>`).join("")}</select></label>
    <label class="f">Pakning (DP)<input class="txt" id="np-pakning" placeholder="2 x 1 kg"></label>
    <label class="f">Enheter per DP<input class="txt" id="np-perDP" inputmode="numeric" value="2"></label>
    <label class="f">Brutto per DP (g)<input class="txt" id="np-dpBruttoG" inputmode="numeric"></label>
    <label class="f">Netto per enhet (g)<input class="txt" id="np-stkNettoG" inputmode="numeric"></label>
  </div><div class="actions"><button class="btn primary" type="submit">Legg til produkt</button><button class="btn ghost" type="button" data-act="avbrytProdukt">Avbryt</button></div></form>`:""}
  <div class="tablewrap"><table class="grid"><thead><tr><th>Varenr.</th><th>Kort navn</th><th class="n">Pris eks. mva</th><th>Gruppe</th><th>Pakning</th><th class="n">Per DP</th><th class="n">Brutto DP (g)</th><th class="n">Netto enhet (g)</th><th>GTIN</th><th>Aktiv</th></tr></thead>
  <tbody>${ps.map(p=>`<tr>
    <td>#${esc(p.varenr)}</td>
    <td><input class="txt" id="pr-${esc(p.varenr)}-kort" data-act="prod" data-v="${esc(p.varenr)}" data-f="kort" value="${esc(p.kort)}" style="min-width:180px"></td>
    <td class="n"><input class="txt" id="pr-${esc(p.varenr)}-pris" data-act="prod" data-v="${esc(p.varenr)}" data-f="pris" data-num="1" value="${p.pris?esc(String(p.pris).replace(".",",")):""}" placeholder="–" style="width:80px;text-align:right"></td>
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
  /* Lagre reageres på allerede ved mousedown: ellers kan siden tegnes på nytt (fordi feltet man sto i lagres) før klikket kommer fram. */
  main.addEventListener("mousedown", e=>{ const b=e.target.closest("[data-act=lagreOrdre]"); if(!b) return; e.preventDefault(); S.lagreMd=Date.now(); lagreKnapp(b.dataset.id); });
  main.addEventListener("click", e=>{ if(e.target.closest("[data-act=lagreOrdre]") && Date.now()-(S.lagreMd||0)<1500) return; klikk(e); });
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
    if(e.target.id==="transportForm") lagreTransportSkjema(e.target);
  });
}
function byttUke(dager){ const d=new Date(S.ukeDato); d.setDate(d.getDate()+dager); S.ukeDato=d; S.open.clear(); abonnerUke(); render(); }

function klikk(e){
  const el=e.target.closest("[data-act]"); if(!el) return;
  const a=el.dataset.act, id=el.dataset.id;
  if(a==="toggle"){ if(e.target.closest("select,input,button:not([data-act=toggle])")) return; S.open.has(id)?S.open.delete(id):S.open.add(id); S.confirmDel=null; render(); }
  else if(a==="nyOrdre") nyOrdre(id);
  else if(a==="lagreOrdre") lagreKnapp(id);
  else if(a==="wiLegg") leggInnWord();
  else if(a==="wiAvbryt"){ S.wordImport=null; render(); }
  else if(a==="pluss"||a==="minus"){ const o=finnOrdre(id); if(o) stegAntall(o,el.dataset.v,a==="pluss"?1:-1); }
  else if(a==="hentFaste") hentFaste();
  else if(a==="kopierForrige"){ el.disabled=true; kopierForrige().finally(()=>{ el.disabled=false; }); }
  else if(a==="tomUke"){ if(S.confirmTom){ S.confirmTom=false; tomUke(); } else { S.confirmTom=true; render(); setTimeout(()=>{ if(S.confirmTom){ S.confirmTom=false; render(); } },5000); } }
  else if(a==="slettAlle") slettAlle();
  else if(a==="csv") lastNedCSV();
  else if(a==="unimicro") lastNedUnimicro();
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
  else if(a==="nyTransport"){ S.tEdit="__ny"; render(); setTimeout(()=>document.getElementById("tf-navn")?.focus(),30); }
  else if(a==="redigerTransport"){ S.tEdit=id; render(); setTimeout(()=>document.getElementById("tf-kontakt")?.focus(),30); }
  else if(a==="avbrytTransport"){ S.utkast.delete("transportForm|"+(S.tEdit||"")); S.tEdit=null; render(); }
  else if(a==="loggInn") Store.loggInn();
  else if(a==="aapneOppsett"){ history.replaceState(null,"",location.pathname+"#oppsett"); S.fase="oppsett"; S.mangler=[]; S.importStatus=""; render(); }
  else if(a==="lastPaNytt") location.reload();
  else if(a==="opprettLister"){ el.disabled=true; el.textContent="Oppretter …"; Store.opprettLister().then(()=>{ S.mangler=[]; render(); toast("Listene er opprettet i SharePoint."); }).catch(e=>{ toast("Klarte ikke å opprette listene: "+(e.message||e)); render(); }); }
  else if(a==="ferdigOppsett"){ history.replaceState(null,"",location.pathname); lastAlt(); }
  else if(a==="eksporter"){ Store.eksporter().then(d=>{ S.backupTatt=true; render(); return lagre(`Ordrebok sikkerhetskopi ${iDag()}.json`, JSON.stringify(d,null,1)); }).catch(e=>toast(errMsg(e))); }
  else if(a==="avbrytProdukt"){ S.nyttProdukt=false; render(); }
}
const finnOrdre = id => S.ordre.find(o=>o.id===id);

/* Tolker det som skrives i antallsfeltet og gir antall i salgsenheten.
   9 · 1,5dp · 1dp+3 · 1 dp 3 · 3b (3 løse). I D-pak-visning betyr et rent tall D-pak. */
function parseAntall(s,p,dpModus){
  s=String(s).trim().toLowerCase().replace(/,/g,".").replace(/d-?pak|dpk/g,"dp").replace(/\s+/g," ");
  if(!s) return 0;
  const per=p?.perDP>1 ? p.perDP : 1; let n;
  let m=s.match(/^(\d+(?:\.\d+)?) ?(?:dp|d) ?\+? ?(\d+(?:\.\d+)?) ?(?:b|stk|beger|løse)?$/);
  if(m) n=parseFloat(m[1])*per+parseFloat(m[2]);
  else if((m=s.match(/^(\d+(?:\.\d+)?) ?(dp|d|b|stk|beger|løse)?$/))){
    n=parseFloat(m[1]);
    const enhet=m[2]; if(enhet==="dp"||enhet==="d"||(!enhet && dpModus)) n*=per;
  } else return null;
  return per>1 ? Math.round(n) : Math.round(n*10)/10;
}
function visAntall(a,o,p){ if(!a) return ""; return iDP(o,p) ? String(Math.round(a/p.perDP*100)/100).replace(".",",") : String(a).replace(".",","); }
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
  if(t.id==="wordFil" && t.files && t.files[0]){
    const fil=t.files[0]; t.value="";
    Word.tolk(fil, S.kunder, prod).then(w=>{ S.wordImport=w; render(); document.getElementById("wordImport")?.scrollIntoView({block:"start",behavior:"smooth"}); if(!w.rader.length) toast("Fant ingen bestillinger i Word-lista."); })
      .catch(err=>toast("Klarte ikke å lese Word-fila: "+(err.message||err)));
    return;
  }
  if(a==="wiKunde"){ const r=S.wordImport?.rader[+t.dataset.i]; if(r){ r.kundeId=t.value; if(t.value) Word.lagreValg(r.tekst,t.value); render(); } return; }
  if(a==="wiMed"){ const r=S.wordImport?.rader[+t.dataset.i]; if(r){ r.med=t.checked; render(); } return; }
  if(a==="qty"){
    const o=finnOrdre(id); if(!o) return; const p=prod(t.dataset.v); const n=parseAntall(t.value,p,iDP(o,p));
    if(n===null){ toast("Skriv et tall. Eksempler: 9, 1,5dp, 1dp+3 eller 3b for tre løse."); t.value=visAntall(o.linjer?.[t.dataset.v]?.antall,o,p); return; }
    settAntall(o,t.dataset.v,n);
    t.dataset.keep="0"; t.value=visAntall(n,o,p);
  } else if(a==="dpModus"){
    const o=finnOrdre(id); if(!o) return; S.visDP.set(o.id, t.value==="dp"); render();
  } else if(a==="tp"){
    const tr=S.transport.find(x=>x.id===t.dataset.id); if(!tr) return; const f=t.dataset.f; const v=t.type==="checkbox"?t.checked:(f==="sort"?Number(t.value)||0:t.value.trim());
    if(f==="navn"){ if(!v){ toast("Transportøren må ha et navn."); t.value=tr.navn; return; } if(S.transport.some(x=>x.id!==tr.id&&x.navn.toLowerCase()===v.toLowerCase())){ toast("Det finnes allerede en transportør med dette navnet."); t.value=tr.navn; return; } }
    lagreTransport({...tr,[f]:v});
    if(f==="navn" && tr.navn!==v) toast(`Endret navn. Kunder og ordre som har «${tr.navn}», beholder det gamle navnet til de endres.`);
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

function settAntall(o,v,n){
  const lin={...(o.linjer||{})};
  if(n>0) lin[v]={...(lin[v]||{}), antall:n}; else delete lin[v];
  lagreOrdre({...o, linjer:lin});
}
function stegAntall(o,v,retning){
  const p=prod(v); const a=o.linjer?.[v]?.antall||0; const steg=iDP(o,p)?p.perDP:1;
  const ny=Math.max(0, iDP(o,p) ? (retning>0 ? (Math.floor(a/steg+1e-9)+1)*steg : (Math.ceil(a/steg-1e-9)-1)*steg) : Math.round((a+retning*steg)*10)/10);
  settAntall(o,v,ny);
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
async function kopierForrige(){
  const u=uke(); const d=new Date(mandag(S.ukeDato)); d.setDate(d.getDate()-7); const forrige=isoUke(d);
  let liste;
  try{ liste=await Store.hent("ordre",forrige.key); }catch(e){ toast(errMsg(e)); return; }
  const har=new Set(S.ordre.map(o=>o.kundeId));
  const kandidater=liste.filter(o=>linjer(o).length && !har.has(o.kundeId));
  if(!liste.length){ toast(`Fant ingen ordre i uke ${forrige.w}.`); return; }
  if(!kandidater.length){ toast(`Alle kundene fra uke ${forrige.w} har allerede ordre denne uka.`); return; }
  for(const o of kandidater){
    const lin={}; linjer(o).forEach(l=>{ lin[l.varenr]={antall:l.antall}; });
    const id=`${u.key}-${slug(o.kundeNavn||"kunde")||"kunde"}-${Math.random().toString(36).slice(2,6)}`;
    lagreOrdre({id, uke:u.key, kundeId:o.kundeId, kundeNavn:o.kundeNavn, linjer:lin, transport:o.transport||"", sendes:"", prodDato:"", bestnr:"", merknad:o.merknad||"", status:"ny", kilde:"kopi", opprettet:new Date().toISOString()});
  }
  const hoppet=liste.filter(o=>har.has(o.kundeId)).length;
  toast(`${kandidater.length} ordre er kopiert fra uke ${forrige.w}${hoppet?`. ${hoppet} kunder hadde allerede ordre og ble hoppet over`:""}. Husk å sjekke antallene.`);
}
function tomUke(){
  const u=uke(); const ids=S.ordre.map(o=>o.id); if(!ids.length) return;
  const sl=(S.slettet=S.slettet||new Set()); ids.forEach(id=>{ sl.add(id); S.override.delete(id); });
  S.ordre=[]; render();
  Promise.all(ids.map(id=>skriv("ordre/"+id,null).catch(()=>null))).then(()=>toast(`Uke ${u.w} er tømt. ${ids.length} ordre er slettet.`));
}
async function slettAlle(){
  const felt=document.getElementById("slettBekreft");
  if(!S.backupTatt){ toast("Last ned en sikkerhetskopi først."); return; }
  if(!felt || felt.value.trim()!=="SLETT"){ toast("Skriv SLETT i feltet for å bekrefte."); return; }
  try{
    const alle=await Store.hent("ordre"); let n=0;
    S.importStatus=`Sletter ${alle.length} ordre …`; render();
    for(const o of alle){ await Store.slett("ordre",o.id); n++; if(n%10===0){ S.importStatus=`Sletter … ${n} av ${alle.length}`; planRender(); } }
    S.ordre=[]; S.importStatus=`Alle ordre er slettet (${n}). Kunder og produkter er beholdt.`; render();
  }catch(e){ S.importStatus="Slettingen stoppet: "+(e.message||e); render(); }
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
  const pr=lesPriser(k.priser); if(pr===null){ toast("Kundepriser skrives som varenr=pris, f.eks. 0032=140; 0001=38."); return; } k.priser=pr;
  S.utkast.delete("kundeForm|"+(form.dataset.id||""));
  if(!k.navn){ toast("Kunden må ha et navn."); return; }
  if(!gammel){ k.id = k.kundenr || ("k-"+slug(k.navn)+"-"+Math.random().toString(36).slice(2,5)); if(kunde(k.id)){ toast("Det finnes allerede en kunde med dette kundenummeret."); return; } S.kunder=[...S.kunder,k].sort((a,b)=>a.navn.localeCompare(b.navn,"nb")); S.kundeSel=k.id; }
  oppdaterKunde(k).then(()=>toast(`${k.navn} er lagret.`));
}
function leggTilProdukt(){
  const g=id=>document.getElementById(id).value.trim();
  const varenr=g("np-varenr").replace(/^#/,"");
  if(!/^[A-Za-z0-9_-]+$/.test(varenr)){ toast("Varenummeret kan bare inneholde tall og bokstaver."); return; }
  if(prod(varenr)){ toast("Det finnes allerede et produkt med dette varenummeret."); return; }
  const p={varenr, kort:g("np-kort"), navn:g("np-kort"), gruppe:g("np-gruppe"), pakning:g("np-pakning"), perDP:Number(g("np-perDP"))||1, dpBruttoG:Number(g("np-dpBruttoG"))||0, stkNettoG:Number(g("np-stkNettoG"))||0, stkBruttoG:0, pris:Number(g("np-pris").replace(",","."))||0, enhet:"stk", varetype:"Fersk ost", sort:50, aktiv:true};
  S.produkter=[...S.produkter,p]; S.nyttProdukt=false; render(); const {varenr:v,...d}=p; skriv("produkter/"+v,d).then(()=>toast(`${p.kort} er lagt til.`));
}

/* ---------- PDF og filer ---------- */
/* PDF-er lastes ned med riktig filnavn (f.eks. «Ysteriliste uke 43 2026.pdf») i stedet for å åpnes i ny fane,
   der nettleseren ville gitt fila et tilfeldig navn. */
function forberedVindu(){ S.pdfVindu=null; }
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
function wrap(font,size,text,maxW){ const words=String(text).split(/[ \t\n]+/); const lines=[]; let cur=""; for(const w of words){ const t=cur?cur+" "+w:w; if(font.widthOfTextAtSize(t,size)<=maxW) cur=t; else { if(cur) lines.push(cur); cur=w; while(font.widthOfTextAtSize(cur,size)>maxW && cur.length>1){ let i=cur.length-1; while(i>1 && font.widthOfTextAtSize(cur.slice(0,i),size)>maxW) i--; lines.push(cur.slice(0,i)); cur=cur.slice(i); } } } if(cur) lines.push(cur); return lines.length?lines:[""]; }
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
  set("Carrier",transportAdresse(o.transport),9);
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

/* Korte kolonneoverskrifter til ysterilista, f.eks. «Mozz. 125 g», «Stracc. spann 5 kg». */
function kortEtikett(p){ return String(p.kort||p.varenr).replace(/\s*\(.*?\)\s*/g," ").replace(/Mozzarella/i,"Mozz.").replace(/Stracciatella/i,"Stracc.").replace(/Revet cubetti/i,"Revet").replace(/\s+/g," ").trim().replace(/(\d) (g|kg)\b/g,"$1 $2"); }
function kommentarFull(o){ const k=kunde(o.kundeId); return [o.prodDato?`Prod. ${dmy(o.prodDato)}`:"", o.sendes?`Sendes ${dmy(o.sendes)}`:"", o.transport, o.bestnr?`Best.nr ${o.bestnr}`:"", o.merknad||k?.merknad].filter(Boolean).join(" · "); }
async function ysteriPdf(){
  const {PDFDocument,StandardFonts,rgb}=PDFLib; const u=uke(); const {ordre,kol,sum}=ysteriData();
  const doc=await PDFDocument.create(); const f=await doc.embedFont(StandardFonts.Helvetica), fb=await doc.embedFont(StandardFonts.HelveticaBold);
  const W=841.89,H=595.28, M=28; const ink=rgb(0.17,0.16,0.15), grey=rgb(0.40,0.38,0.36), band=rgb(0.95,0.93,0.88), stripe=rgb(0.985,0.978,0.96), strek=rgb(0.85,0.84,0.82);
  const kundeW=165; const cw=Math.min(62,(W-2*M-kundeW)/Math.max(1,kol.length)); const tabW=kundeW+cw*kol.length;
  const xKol=i=>M+kundeW+i*cw;
  const høyre=(s,font,size,i,yy)=>{ s=safe(font,s); const w=font.widthOfTextAtSize(s,size); page.drawText(s,{x:xKol(i)+cw-5-w,y:yy,size,font,color:ink}); };
  let page, y, sideNr=0; const m=mandag(S.ukeDato), sø=new Date(m); sø.setDate(sø.getDate()+6);
  const ny=()=>{ page=doc.addPage([W,H]); sideNr++;
    page.drawText(safe(fb,`Ysteriliste uke ${u.w}`),{x:M,y:H-M-8,size:17,font:fb,color:ink});
    page.drawText(safe(f,`${fmtD(m)}–${fmtD(sø)}.${sø.getFullYear()} · ${ordre.length} ordre · skrevet ut ${dmyLang(iDag())}`),{x:M,y:H-M-24,size:9,font:f,color:grey});
    y=H-M-40;
    const etik=kol.map(p=>wrap(fb,7.5,safe(fb,kortEtikett(p)),cw-8));
    const hh=Math.max(...etik.map(e=>e.length),1)*9+10;
    page.drawRectangle({x:M,y:y-hh,width:tabW,height:hh,color:band});
    page.drawText("Kunde",{x:M+5,y:y-13,size:8.5,font:fb,color:ink});
    etik.forEach((ln,i)=>ln.forEach((t,j)=>{ const w=fb.widthOfTextAtSize(t,7.5); page.drawText(t,{x:xKol(i)+cw-5-w,y:y-12-j*9,size:7.5,font:fb,color:ink}); }));
    y-=hh+14; };
  ny();
  ordre.forEach((o,ri)=>{
    const nL=wrap(f,9.5,safe(f,o.kundeNavn),kundeW-10).slice(0,2);
    const kom=kommentarFull(o); const kL=kom?wrap(f,8,safe(f,kom),tabW-24):[];
    const h=nL.length*11+kL.length*10+8;
    if(y-h<M+24) ny();
    if(ri%2===1) page.drawRectangle({x:M,y:y-h+11,width:tabW,height:h,color:stripe});
    nL.forEach((t,j)=>page.drawText(t,{x:M+5,y:y-j*11,size:9.5,font:f,color:ink}));
    kol.forEach((p,i)=>{ const l=(o.linjer||{})[p.varenr]; if(l&&l.antall) høyre(nf(l.antall),fb,10.5,i,y); });
    kL.forEach((t,j)=>page.drawText(t,{x:M+14,y:y-nL.length*11-j*10,size:8,font:f,color:grey}));
    y-=h; page.drawLine({start:{x:M,y:y+11},end:{x:M+tabW,y:y+11},thickness:0.4,color:strek});
  });
  /* Summer */
  const tot=kol.reduce((s,p)=>s+kgNetto(sum[p.varenr],p),0);
  const dpL=kol.map(p=>{ if(!(p.perDP>2)) return []; const dp=Math.floor(sum[p.varenr]/p.perDP), r=sum[p.varenr]%p.perDP; return [dp?`${dp} DP`:"", r?`+ ${nf(r)}`:""].filter(Boolean); });
  const dpH=Math.max(1,...dpL.map(d=>d.length))*10;
  if(y-(16+dpH+16+20)<M+24) ny();
  y-=4; page.drawLine({start:{x:M,y:y+12},end:{x:M+tabW,y:y+12},thickness:1,color:ink});
  page.drawText("Sum enheter",{x:M+5,y,size:9.5,font:fb,color:ink}); kol.forEach((p,i)=>høyre(nf(sum[p.varenr]),fb,10.5,i,y)); y-=16;
  page.drawText("Detaljistpakninger (DP)",{x:M+5,y,size:9.5,font:fb,color:ink}); dpL.forEach((d,i)=>d.forEach((t,j)=>høyre(t,f,8.5,i,y-j*10))); y-=dpH+6;
  page.drawText("Kg ost (netto)",{x:M+5,y,size:9.5,font:fb,color:ink}); kol.forEach((p,i)=>høyre(nf(kgNetto(sum[p.varenr],p),1),f,9,i,y)); y-=18;
  page.drawText(safe(fb,`Totalt ${nf(tot,1)} kg ost (netto)`),{x:M+5,y,size:10,font:fb,color:ink});
  const sider=doc.getPages(); sider.forEach((pg,i)=>{ const t=`Side ${i+1} av ${sider.length} · Ysteriliste uke ${u.w}`; pg.drawText(t,{x:W-M-f.widthOfTextAtSize(t,7.5),y:M-8,size:7.5,font:f,color:grey}); });
  const bytes=await doc.save();
  await lagre(`Ysteriliste uke ${u.w} ${u.y}.pdf`, new Blob([bytes],{type:"application/pdf"}));
}
function ysteriTekst(){
  const u=uke(), {ordre,kol,sum}=ysteriData();
  let t=`YSTERILISTE UKE ${u.w}\n\nSUM\n`+kol.map(p=>`${p.kort}: ${sum[p.varenr]}${dpTekst(sum[p.varenr],p)?` (${dpTekst(sum[p.varenr],p)})`:""}`).join("\n");
  t+="\n\nPER KUNDE\n"+ordre.map(o=>`${o.kundeNavn}: `+linjer(o).map(l=>`${l.antall} ${l.p?.kort||l.varenr}`).join(", ")+(kommentarFull(o)?`\n  ${kommentarFull(o)}`:"")).join("\n");
  const ok=()=>toast("Ysterilista er kopiert. Lim den inn i SMS eller e-post.");
  try{ navigator.clipboard.writeText(t).then(ok,()=>fallbackKopi(t)); }catch(_){ fallbackKopi(t); }
}
function fallbackKopi(t){ const ta=document.createElement("textarea"); ta.value=t; ta.className="txt"; ta.style.height="240px"; ta.readOnly=true; $("#main").prepend(ta); ta.select(); toast("Teksten er markert. Kopier den med Ctrl+C eller ⌘C."); }
function priserTekst(o){ return Object.entries(o||{}).map(([v,p])=>`${v}=${String(p).replace(".",",")}`).join("; "); }
function lesPriser(t){ const o={}; for(const del of String(t||"").split(/[;\n]+/).map(x=>x.trim()).filter(Boolean)){ const m=del.match(/^#?([A-Za-z0-9_-]+)\s*[=:]\s*([0-9]+(?:[.,][0-9]+)?)$/); if(!m) return null; o[m[1]]=Number(m[2].replace(",",".")); } return o; }
function prisFor(k,varenr){ const kp=k?.priser?.[varenr]; if(kp>0) return kp; const p=prod(varenr); return p?.pris>0 ? p.pris : null; }
function lastSkript(src){ return new Promise((ok,nei)=>{ if(window.XLSX) return ok(); const s=document.createElement("script"); s.src=src; s.onload=ok; s.onerror=()=>nei(new Error("Fant ikke "+src)); document.head.appendChild(s); }); }
/* Ordre til Unimicro: én rad per varelinje, kolonner som i Unimicros importmal (OrderTemplate.xlsx).
   Ordrenr lages som ÅÅUU + løpenr (uke 43 i 2026 gir 264301, 264302 …). */
async function lastNedUnimicro(){
  const u=uke();
  const ordre=[...S.ordre].filter(o=>(o.status||"ny")!=="fakturert" && linjer(o).length).sort((a,b)=>(a.kundeNavn||"").localeCompare(b.kundeNavn||"","nb"));
  const utenNr=ordre.filter(o=>!kunde(o.kundeId)?.kundenr);
  if(utenNr.length){ toast(`Mangler Unimicro-kundenummer: ${utenNr.map(o=>o.kundeNavn).join(", ")}. Legg det inn på kunden først.`); return; }
  const utenPris=new Set(); ordre.forEach(o=>linjer(o).forEach(l=>{ if(prisFor(kunde(o.kundeId),l.varenr)===null) utenPris.add(l.p?.kort||l.varenr); }));
  if(utenPris.size){ toast(`Mangler pris på ${[...utenPris].join(", ")}. Legg inn pris under Produkter.`); return; }
  try { await lastSkript("vendor/xlsx.mini.min.js"); } catch(e){ toast("Klarte ikke å laste Excel-modulen. Sjekk nettforbindelsen."); return; }
  const H=["Ordrenr","Ordredato","Leveringsdato","Kundenr","Kundenavn","Vår referanse","Deres referanse","Hovedselger","Rekvisisjon","Leveringsadr.","Lev.postnr.","Lev.poststed","Varenr.","Varenavn","Antall","Pris eks. mva","Rabatt %","Prosjekt","Avdeling"];
  /* Unimicro krever datoformat ddMMyy, f.eks. 241026 */
  const ddmmyy=iso=>{ if(!iso) return ""; const [y,m,d]=iso.split("-"); return d+m+y.slice(2); };
  const idag=ddmmyy(iDag()); const rader=[H];
  ordre.forEach((o,i)=>{
    const k=kunde(o.kundeId); const nr=`${String(u.y).slice(2)}${pad(u.w)}${pad(i+1)}`;
    let adr=k.adresse||"", pnr=k.postnr||"", sted=k.poststed||"";
    if(k.levAdresse){ const l=k.levAdresse.split(/\n|,/).map(x=>x.trim()).filter(Boolean); const m=(l[l.length-1]||"").match(/^(\d{4})\s+(.+)$/); if(m){ pnr=m[1]; sted=m[2]; adr=l.slice(0,-1).join(", "); } else adr=l.join(", "); }
    linjer(o).forEach(l=>rader.push([nr, idag, ddmmyy(o.sendes||o.prodDato||iDag()), Number(k.kundenr), k.navn, AVSENDER.ref, k.kontakt||"", "", o.bestnr||"", adr, pnr, sted, l.varenr, l.p?.navn||l.p?.kort||l.varenr, l.antall, prisFor(k,l.varenr), "", "", ""]));
  });
  const ws=XLSX.utils.aoa_to_sheet(rader);
  for(let r=1;r<rader.length;r++){ for(const c of [0,1,2,12]){ const ref=XLSX.utils.encode_cell({r,c}); if(ws[ref]){ ws[ref].t="s"; ws[ref].v=String(ws[ref].v); ws[ref].z="@"; } } }
  ws["!cols"]=H.map((h,i)=>({wch:[10,11,11,9,30,16,16,10,12,28,9,14,8,32,8,11,8,8,8][i]}));
  const wb=XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb,ws,"Data");
  const data=XLSX.write(wb,{bookType:"xlsx",type:"array"});
  lagre(`Unimicro ordre uke ${u.w}.xlsx`, new Blob([data],{type:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"}));
  toast(`${ordre.length} ordre med ${rader.length-1} linjer er lagt i fila. Importer den i Unimicro.`);
}
function lastNedCSV(){
  const u=uke(); const rader=[["Uke","Kundenr","Kunde","Orgnr","Sendes","Bestillingsnr","Varenr","Vare","Antall","Status"]];
  [...S.ordre].sort((a,b)=>(a.kundeNavn||"").localeCompare(b.kundeNavn||"","nb")).forEach(o=>{ const k=kunde(o.kundeId); linjer(o).forEach(l=>rader.push([u.w,k?.kundenr||"",o.kundeNavn,k?.orgnr||"",dmyLang(o.sendes),o.bestnr||"",l.varenr,l.p?.kort||"",l.antall,o.status||"ny"])); });
  const csv="﻿"+rader.map(r=>r.map(c=>{ c=String(c??""); return /[;"\n]/.test(c)?`"${c.replace(/"/g,'""')}"`:c; }).join(";")).join("\r\n");
  lagre(`Fakturagrunnlag uke ${u.w}.csv`, csv);
}

start();
window.addEventListener("unhandledrejection",e=>{ if(S.pdfVindu){ try{S.pdfVindu.close();}catch(_){} S.pdfVindu=null; } });
if("serviceWorker" in navigator) window.addEventListener("load",()=>navigator.serviceWorker.register("sw.js").catch(()=>{}));
