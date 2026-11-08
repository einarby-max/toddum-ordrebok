/* Lager: beholdning i de to kjølerommene, ført som bevegelser i SharePoint-lista «Ordrebok Lager».
   Inn: ferdig ysting (produsert). Ut: «Pakk og send» på ordre, svinn. Telling og flytting føres for hånd.
   Beholdningen er summen av alle bevegelser per kjølerom og vare, i salgsenheten (beger, spann, kg …). */
const Lager = (() => {
  const L = () => S.lg || (S.lg = { rader: [], innst: null, loaded: false, laster: false, telling: null, flytt: false, bekreft: null });
  const ROM_STD = { r1: "Kjølerom 1", r2: "Kjølerom 2" };
  const romMap = () => ({ ...ROM_STD, ...((L().innst || {}).rom || {}) });
  const romIds = () => Object.keys(romMap());
  const romNavn = id => romMap()[id] || id || "–";
  const innRom = () => ((L().innst || {}).inn) || "r1";
  const tall = v => { const n = Number(String(v ?? "").replace(/\s/g, "").replace(",", ".")); return isFinite(n) ? n : 0; };
  const nå = () => new Date().toISOString();
  const bruker = () => (S.bruker && (S.bruker.navn || S.bruker.epost)) || "";
  const varer = () => prodSort().filter(p => p.gruppe !== "Melk");
  const TYPE = { produksjon: "Produksjon", sendt: "Sendt", svinn: "Svinn", telling: "Telling", flytt: "Flytting" };

  async function last() {
    const l = L(); if (l.laster || !Store.harListe("lager")) { l.loaded = true; return; }
    l.laster = true;
    try {
      const [r, i] = await Promise.all([Store.hent("lager"), Store.hent("innst").catch(() => [])]);
      l.rader = r; const s = i.find(x => x.id === "lager"); l.innst = s ? s.data : null; l.loaded = true;
    } catch (e) { console.warn("Lager", e); l.loaded = true; }
    l.laster = false; planRender();
  }
  function poll() { if (S.tab === "lager" && !L().telling) last(); }

  function beholdning() {
    const b = {}; for (const r of L().rader) { const v = r.varenr, rm = r.rom || "r1"; if (!v) continue; (b[v] = b[v] || {}); b[v][rm] = (b[v][rm] || 0) + tall(r.antall); }
    return b;
  }
  function skrivRad(r) {
    const l = L(); r = { tid: nå(), av: bruker(), ...r };
    const i = l.rader.findIndex(x => x.id === r.id); if (i >= 0) l.rader[i] = r; else l.rader.push(r);
    const { id, ...d } = r; return skriv("lager/" + id, d);
  }
  function slettRad(id) { const l = L(); if (!l.rader.some(x => x.id === id)) return Promise.resolve(); l.rader = l.rader.filter(x => x.id !== id); return skriv("lager/" + id, null); }
  function harLista() { if (Store.harListe("lager")) return true; toast("Lagerlista finnes ikke ennå. Last ordreboka på nytt, så opprettes den."); return false; }

  /* ---------- kalles fra ordre, produksjon og svinn ---------- */
  function registrerSendt(o) {
    if (!harLista()) return Promise.resolve();
    const rom = o.sendtInfo?.rom || "r1", dato = o.sendes || iDag(), jobber = [];
    for (const [v, l] of Object.entries(o.linjer || {})) {
      const id = `ut-${o.id}-${v}`, n = tall(l.sendt ?? l.antall);
      if (n > 0) jobber.push(skrivRad({ id, dato, varenr: v, antall: -n, rom, type: "sendt", ref: o.id, lot: l.batch || batchFra(o.prodDato) || "", merknad: o.kundeNavn || "" }));
      else jobber.push(slettRad(id));
    }
    planRender(); return Promise.all(jobber);
  }
  function angreSendt(o) { const ids = L().rader.filter(r => r.ref === o.id && r.type === "sendt").map(r => r.id); planRender(); return Promise.all(ids.map(slettRad)); }
  function produksjonInn(y) {
    if (!Store.harListe("lager")) return Promise.resolve();
    const rom = innRom(), jobber = [];
    for (const [v, n] of Object.entries(y.produsert || {})) if (tall(n) > 0) jobber.push(skrivRad({ id: `inn-${y.id}-${v}`, dato: y.dato || iDag(), varenr: v, antall: tall(n), rom, type: "produksjon", ref: y.id, lot: y.lot || "", merknad: `Ysting ${dmy(y.dato)}` }));
    return Promise.all(jobber).then(() => { if (jobber.length) toast(`Produksjonen er satt inn i ${romNavn(rom)}.`); });
  }
  function svinnUt(x, rom) { if (!Store.harListe("lager")) return Promise.resolve(); return skrivRad({ id: `sv-${x.id}`, dato: x.dato, varenr: x.varenr, antall: -tall(x.antall), rom: rom || "r1", type: "svinn", ref: x.id, lot: x.lot || "", merknad: x.arsak || "" }); }
  function slettRef(id) { return slettRad(id); }
  function romValg(valgt) { return romIds().map(r => `<option value="${r}" ${r === valgt ? "selected" : ""}>${esc(romNavn(r))}</option>`).join(""); }

  /* ---------- visning ---------- */
  function mengde(n, p) { const d = p && p.perDP > 1 && n > 0 ? dpTekst(Math.round(n), p) : ""; return `${nf(n)} ${esc(enhetNavn(p, n))}${d ? `<div class="hint">${esc(d)}</div>` : ""}`; }
  function vis() {
    const l = L();
    if (!Store.harListe("lager")) return `<div class="empty"><h3>Lageret er ikke satt opp ennå</h3>Last ordreboka på nytt. Da opprettes lista «Ordrebok Lager» i SharePoint.</div>`;
    if (!l.loaded) { last(); return `<div class="empty"><h3>Henter lagerbeholdning …</h3></div>`; }
    if (l.telling) return visTelling();
    const b = beholdning(), rom = romIds(), u = uke();
    const bestilt = {}; S.ordre.filter(o => !(o.sendtInfo && o.sendtInfo.tid) && (o.status || "ny") !== "sendt" && (o.status || "ny") !== "fakturert").forEach(o => linjer(o).forEach(x => bestilt[x.varenr] = (bestilt[x.varenr] || 0) + x.antall));
    const harTelling = l.rader.some(r => r.type === "telling");
    const rader = varer().map(p => { const r = b[p.varenr] || {}; const tot = rom.reduce((s, x) => s + (r[x] || 0), 0); const best = bestilt[p.varenr] || 0; return { p, r, tot, best, ledig: tot - best }; });
    const vist = rader.filter(x => x.tot || x.best || rom.some(r => x.r[r]));
    const siste = [...l.rader].sort((a, b) => (b.tid || "").localeCompare(a.tid || "")).slice(0, 30);
    return `${!harTelling ? `<p class="notice"><b>Start med å telle opp begge kjølerommene.</b> Beholdningen regnes fra bevegelsene som registreres i ordreboka. Den første opptellingen gir riktig utgangspunkt.</p>` : ""}
    <div class="toolbar">
      ${rom.map(r => `<button class="btn primary" data-act="lgTell" data-rom="${r}">Tell opp ${esc(romNavn(r))}</button>`).join("")}
      <button class="btn" data-act="lgFlytt">Flytt mellom kjølerom</button>
    </div>
    ${l.flytt ? visFlytt() : ""}
    <div class="tablewrap"><table class="grid lager"><thead><tr><th>Vare</th>${rom.map(r => `<th class="n">${esc(romNavn(r))}</th>`).join("")}<th class="n">Totalt</th><th class="n">Bestilt, ikke sendt (uke ${u.w})</th><th class="n">Ledig</th></tr></thead><tbody>
    ${vist.length ? vist.map(x => `<tr><td>${esc(x.p.kort)}<div class="hint">#${esc(x.p.varenr)}</div></td>${rom.map(r => `<td class="n ${(x.r[r] || 0) < 0 ? "neg" : ""}">${x.r[r] ? mengde(x.r[r], x.p) : "–"}</td>`).join("")}<td class="n"><b>${mengde(x.tot, x.p)}</b></td><td class="n">${x.best ? mengde(x.best, x.p) : "–"}</td><td class="n ${x.ledig < 0 ? "neg" : ""}">${mengde(x.ledig, x.p)}</td></tr>`).join("") : `<tr><td colspan="${rom.length + 4}" class="hint">Ingen varer på lager ennå.</td></tr>`}
    </tbody></table></div>
    <p class="hint">Inn: ferdige ystinger (settes i ${esc(romNavn(innRom()))}). Ut: ordre som er pakket og sendt, og svinn. Negativt tall betyr at noe ikke er registrert, og at rommet bør telles opp.</p>
    <div class="card form"><h3>Siste bevegelser</h3>
      <div class="tablewrap"><table class="grid"><thead><tr><th>Dato</th><th>Type</th><th>Vare</th><th class="n">Antall</th><th>Kjølerom</th><th>Lot</th><th>Merknad</th><th>Av</th><th></th></tr></thead><tbody>
      ${siste.length ? siste.map(r => { const p = prod(r.varenr); const manuell = r.type === "telling" || r.type === "flytt"; return `<tr><td>${esc(dmy(r.dato))}</td><td>${esc(TYPE[r.type] || r.type)}</td><td>${esc(p?.kort || r.varenr)}</td><td class="n">${tall(r.antall) > 0 ? "+" : ""}${nf(tall(r.antall))}</td><td>${esc(romNavn(r.rom))}</td><td>${esc(r.lot || "")}</td><td class="hint">${esc(r.merknad || "")}</td><td class="hint">${esc(r.av || "")}</td><td>${manuell ? `<button class="btn sm ghost danger" data-act="lgSlett" data-id="${esc(r.id)}" aria-label="Slett">${l.bekreft === r.id ? "Slett?" : "✕"}</button>` : ""}</td></tr>`; }).join("") : `<tr><td colspan="9" class="hint">Ingen bevegelser ennå.</td></tr>`}
      </tbody></table></div></div>
    <form class="card form" id="lgInnst">
      <h3>Kjølerommene</h3>
      <div class="row">${rom.map(r => `<label class="f">Navn<input class="txt" id="lgn-${r}" value="${esc(romNavn(r))}"></label>`).join("")}
        <label class="f">Ny produksjon settes i<select class="txt" id="lg-inn">${romValg(innRom())}</select></label></div>
      <div class="actions"><button class="btn" type="submit">Lagre</button></div>
    </form>`;
  }
  function visFlytt() {
    const rom = romIds();
    return `<form class="card form" id="lgFlyttForm" data-bevar="1"><h3>Flytt mellom kjølerom</h3>
      <div class="row"><label class="f">Vare<select class="txt" id="lf-vare">${varer().map(p => `<option value="${esc(p.varenr)}">${esc(p.kort)}</option>`).join("")}</select></label>
      <label class="f">Antall<input class="txt" id="lf-antall" inputmode="decimal" placeholder="f.eks. 12 eller 2dp" required></label>
      <label class="f">Fra<select class="txt" id="lf-fra">${romValg(rom[0])}</select></label>
      <label class="f">Til<select class="txt" id="lf-til">${romValg(rom[1])}</select></label></div>
      <div class="actions"><button class="btn primary" type="submit">Flytt</button><button class="btn ghost" type="button" data-act="lgFlyttAvbryt">Avbryt</button></div></form>`;
  }
  function visTelling() {
    const l = L(), rom = l.telling, b = beholdning();
    return `<form class="card form" id="lgTellForm" data-bevar="1" data-id="${rom}">
      <h3>Opptelling: ${esc(romNavn(rom))}</h3>
      <p class="hint">Skriv det som faktisk står i rommet. La feltet stå tomt for varer du ikke har telt. For D-pak kan du skrive f.eks. <b>3dp</b> eller <b>3dp+2</b>. Differansen føres som telling.</p>
      <div class="tablewrap"><table class="grid"><thead><tr><th>Vare</th><th class="n">Etter ordreboka</th><th>Telt</th></tr></thead><tbody>
      ${varer().map(p => { const n = (b[p.varenr] || {})[rom] || 0; return `<tr><td>${esc(p.kort)}<div class="hint">${esc(p.pakning || "")}</div></td><td class="n">${mengde(n, p)}</td><td><input class="txt" id="lt-${esc(p.varenr)}" inputmode="decimal" style="width:110px;font-size:16px" placeholder="${esc(enhetNavn(p, 2))}"></td></tr>`; }).join("")}
      </tbody></table></div>
      <div class="actions"><button class="btn primary" type="submit">Lagre opptelling</button><button class="btn ghost" type="button" data-act="lgTellAvbryt">Avbryt</button></div></form>`;
  }

  /* ---------- hendelser ---------- */
  function klikk(el) {
    const a = el.dataset.act, l = L(); if (!a || !a.startsWith("lg")) return false;
    if (a === "lgTell") { if (harLista()) { l.telling = el.dataset.rom; l.flytt = false; render(); window.scrollTo(0, 0); } }
    else if (a === "lgTellAvbryt") { S.utkast.delete("lgTellForm|" + l.telling); l.telling = null; render(); }
    else if (a === "lgFlytt") { l.flytt = !l.flytt; render(); }
    else if (a === "lgFlyttAvbryt") { l.flytt = false; S.utkast.delete("lgFlyttForm|"); render(); }
    else if (a === "lgSlett") { const id = el.dataset.id; if (l.bekreft === id) { l.bekreft = null; const r = l.rader.find(x => x.id === id); const par = r && r.type === "flytt" ? l.rader.filter(x => x.ref === r.ref && x.type === "flytt").map(x => x.id) : [id]; Promise.all(par.map(slettRad)).then(() => toast("Bevegelsen er slettet.")); render(); } else { l.bekreft = id; render(); setTimeout(() => { if (l.bekreft === id) { l.bekreft = null; render(); } }, 4000); } }
    return true;
  }
  function submit(form) {
    const l = L();
    if (form.id === "lgTellForm") {
      const rom = form.dataset.id, b = beholdning(), tid = Date.now().toString(36), jobber = []; let n = 0, feil = [];
      for (const p of varer()) {
        const el = document.getElementById("lt-" + p.varenr); if (!el || !el.value.trim()) continue;
        const telt = parseAntall(el.value, p, false); if (telt === null) { feil.push(p.kort); continue; }
        const før = (b[p.varenr] || {})[rom] || 0, diff = Math.round((telt - før) * 10) / 10; n++;
        if (diff) jobber.push(skrivRad({ id: `t-${tid}-${rom}-${p.varenr}`, dato: iDag(), varenr: p.varenr, antall: diff, rom, type: "telling", ref: `t-${tid}`, lot: "", merknad: `Telt ${nf(telt)}, var ${nf(før)}` }));
      }
      if (feil.length) { toast(`Forstår ikke antallet for ${feil.join(", ")}. Skriv f.eks. 12, 3dp eller 3dp+2.`); return true; }
      if (!n) { toast("Skriv antall for minst én vare."); return true; }
      S.utkast.delete("lgTellForm|" + rom); l.telling = null; render();
      Promise.all(jobber).then(() => toast(`✓ Opptellingen av ${romNavn(rom)} er lagret (${n} varer, ${jobber.length} rettet).`));
      return true;
    }
    if (form.id === "lgFlyttForm") {
      const v = $("#lf-vare").value, p = prod(v), n = parseAntall($("#lf-antall").value, p, false), fra = $("#lf-fra").value, til = $("#lf-til").value;
      if (!(n > 0)) { toast("Skriv antall, f.eks. 12 eller 2dp."); return true; }
      if (fra === til) { toast("Velg to forskjellige kjølerom."); return true; }
      const ref = "f-" + Date.now().toString(36), m = `${romNavn(fra)} → ${romNavn(til)}`;
      Promise.all([skrivRad({ id: ref + "-ut", dato: iDag(), varenr: v, antall: -n, rom: fra, type: "flytt", ref, lot: "", merknad: m }), skrivRad({ id: ref + "-inn", dato: iDag(), varenr: v, antall: n, rom: til, type: "flytt", ref, lot: "", merknad: m })])
        .then(() => toast(`✓ ${nf(n)} ${enhetNavn(p, n)} ${p?.kort || v} er flyttet til ${romNavn(til)}.`));
      l.flytt = false; S.utkast.delete("lgFlyttForm|"); render(); return true;
    }
    if (form.id === "lgInnst") {
      const rom = {}; romIds().forEach(r => { const v = ($("#lgn-" + r)?.value || "").trim(); if (v) rom[r] = v; });
      l.innst = { rom, inn: $("#lg-inn").value }; skriv("innst/lager", { data: l.innst }).then(() => toast("✓ Lagret.")); render(); return true;
    }
    return false;
  }

  return { last, poll, vis, klikk, submit, registrerSendt, angreSendt, produksjonInn, svinnUt, slettRef, romValg, romNavn, innRom };
})();
window.Lager = Lager;
