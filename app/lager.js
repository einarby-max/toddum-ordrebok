/* Lager: beholdning i kjølerommene og emballasjelageret, ført som bevegelser i SharePoint-lista «Ordrebok Lager».
   Ferdigvarer (produktene) og lagervarer (fløte, beger, esker, etiketter … i lista «Ordrebok Lagervarer»).
   Inn: ferdig ysting, varemottak. Ut: «Pakk og send», svinn, forbruk (fløte trekkes automatisk når en ysting avsluttes).
   Telling og flytting føres for hånd. Beholdningen er summen av alle bevegelser per rom og vare. */
const Lager = (() => {
  const L = () => S.lg || (S.lg = { rader: [], lv: [], innst: null, loaded: false, laster: false, telling: null, skjema: null, bekreft: null });
  const ROM_STD = { r1: "Kjølerom 1", r2: "Kjølerom 2", r3: "Emballasjelager" };
  const romMap = () => ({ ...ROM_STD, ...((L().innst || {}).rom || {}) });
  const romIds = () => Object.keys(romMap());
  const romNavn = id => romMap()[id] || id || "–";
  const innRom = () => ((L().innst || {}).inn) || "r1";
  const tall = v => { const n = Number(String(v ?? "").replace(/\s/g, "").replace(",", ".")); return isFinite(n) ? n : 0; };
  const nå = () => new Date().toISOString();
  const bruker = () => (S.bruker && (S.bruker.navn || S.bruker.epost)) || "";
  const varer = () => prodSort().filter(p => p.gruppe !== "Melk");
  const lagervarer = () => L().lv.filter(x => x.aktiv !== false).sort((a, b) => (a.kategori || "").localeCompare(b.kategori || "", "nb") || (a.sort ?? 99) - (b.sort ?? 99) || (a.navn || "").localeCompare(b.navn || "", "nb"));
  const lvKey = x => "lv-" + x.id;
  const lvFra = key => key && key.startsWith("lv-") ? L().lv.find(x => x.id === key.slice(3)) : null;
  const navnFor = key => { const lv = lvFra(key); if (lv) return lv.navn; const p = prod(key); return p ? p.kort : key; };
  const KAT = ["Råvare", "Emballasje", "Etiketter", "Annet"];
  const ENH = ["stk", "liter", "kg", "kartong", "pakke", "rull", "pall"];
  const TYPE = { produksjon: "Produksjon", sendt: "Sendt", svinn: "Svinn", telling: "Telling", flytt: "Flytting", mottak: "Varemottak", forbruk: "Forbruk" };

  async function last() {
    const l = L(); if (l.laster || !Store.harListe("lager")) { l.loaded = true; return; }
    l.laster = true;
    try {
      const [r, lv, i] = await Promise.all([Store.hent("lager"), Store.hent("lagervarer").catch(() => []), Store.hent("innst").catch(() => [])]);
      l.rader = r; l.lv = lv; const s = i.find(x => x.id === "lager"); l.innst = s ? s.data : null; l.loaded = true;
    } catch (e) { console.warn("Lager", e); l.loaded = true; }
    l.laster = false; planRender();
  }
  function poll() { const l = L(); if (S.tab === "lager" && !l.telling && !l.skjema) last(); }

  function beholdning() {
    const b = {}; for (const r of L().rader) { const v = r.varenr, rm = r.rom || "r1"; if (!v) continue; (b[v] = b[v] || {}); b[v][rm] = Math.round(((b[v][rm] || 0) + tall(r.antall)) * 1000) / 1000; }
    return b;
  }
  /* Snittforbruk per uke siste 4 uker: alt som har gått ut, unntatt flytting. */
  function forbrukUke(key) {
    const fra = new Date(Date.now() - 28 * 864e5).toISOString().slice(0, 10);
    const ut = L().rader.filter(r => r.varenr === key && r.type !== "flytt" && (r.dato || "") >= fra && tall(r.antall) < 0).reduce((s, r) => s - tall(r.antall), 0);
    return ut / 4;
  }
  function status(lv, tot) {
    const uke = forbrukUke(lvKey(lv)), uker = uke > 0 ? tot / uke : null, min = tall(lv.minimum), led = tall(lv.ledetidUker);
    if (min > 0 && tot <= min) return { k: "bestill", t: "Bestill nå", uke, uker };
    if (uker !== null && led > 0 && uker <= led + 1) return { k: "snart", t: "Bestill snart", uke, uker };
    return { k: "ok", t: "OK", uke, uker };
  }
  function skrivRad(r) {
    const l = L(); r = { tid: nå(), av: bruker(), ...r };
    const i = l.rader.findIndex(x => x.id === r.id); if (i >= 0) l.rader[i] = r; else l.rader.push(r);
    const { id, ...d } = r; return skriv("lager/" + id, d);
  }
  function slettRad(id) { const l = L(); if (!l.rader.some(x => x.id === id)) return Promise.resolve(); l.rader = l.rader.filter(x => x.id !== id); return skriv("lager/" + id, null); }
  function harLista() { if (Store.harListe("lager")) return true; toast("Lagerlista finnes ikke ennå. Last ordreboka på nytt, så opprettes den."); return false; }
  function lagreLv(x) { const l = L(); const i = l.lv.findIndex(y => y.id === x.id); if (i >= 0) l.lv[i] = x; else l.lv.push(x); const { id, ...d } = x; return skriv("lagervarer/" + id, d); }

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
    /* Fløte brukt i ystingen trekkes fra lagervaren som er koblet til fløte. */
    const fl = lagervarer().find(x => x.kobling === "flote");
    if (fl && tall(y.flote) > 0) { const n = Math.round(tall(y.flote) / (tall(fl.perEnhet) || 1) * 100) / 100; jobber.push(skrivRad({ id: `fl-${y.id}`, dato: y.dato || iDag(), varenr: lvKey(fl), antall: -n, rom: fl.rom || "r2", type: "forbruk", ref: y.id, lot: "", merknad: `${nf(tall(y.flote), 1)} l fløte i ysting ${dmy(y.dato)}` })); }
    return Promise.all(jobber).then(() => { if (jobber.length) toast(`Produksjonen er satt inn i ${romNavn(rom)}${fl && tall(y.flote) > 0 ? `, og fløten er trukket fra ${fl.navn}` : ""}.`); });
  }
  function svinnUt(x, rom) { if (!Store.harListe("lager")) return Promise.resolve(); return skrivRad({ id: `sv-${x.id}`, dato: x.dato, varenr: x.varenr, antall: -tall(x.antall), rom: rom || "r1", type: "svinn", ref: x.id, lot: x.lot || "", merknad: x.arsak || "" }); }
  function slettRef(id) { return slettRad(id); }
  function romValg(valgt) { return romIds().map(r => `<option value="${r}" ${r === valgt ? "selected" : ""}>${esc(romNavn(r))}</option>`).join(""); }

  /* ---------- visning ---------- */
  function mengde(n, p) { const d = p && p.perDP > 1 && n > 0 ? dpTekst(Math.round(n), p) : ""; return `${nf(n)} ${esc(enhetNavn(p, n))}${d ? `<div class="hint">${esc(d)}</div>` : ""}`; }
  const lvMengde = (n, lv) => `${nf(n, n % 1 ? 1 : 0)} ${esc(lv.enhet || "stk")}`;
  function vis() {
    const l = L();
    if (!Store.harListe("lager")) return `<div class="empty"><h3>Lageret er ikke satt opp ennå</h3>Last ordreboka på nytt. Da opprettes lista «Ordrebok Lager» i SharePoint.</div>`;
    if (!l.loaded) { last(); return `<div class="empty"><h3>Henter lagerbeholdning …</h3></div>`; }
    if (l.telling) return visTelling();
    const b = beholdning(), rom = romIds(), u = uke();
    const bestilt = {}; S.ordre.filter(o => !(o.sendtInfo && o.sendtInfo.tid) && !["sendt", "fakturert"].includes(o.status || "ny")).forEach(o => linjer(o).forEach(x => bestilt[x.varenr] = (bestilt[x.varenr] || 0) + x.antall));
    const harTelling = l.rader.some(r => r.type === "telling");
    const fv = varer().map(p => { const r = b[p.varenr] || {}; const tot = rom.reduce((s, x) => s + (r[x] || 0), 0); const best = bestilt[p.varenr] || 0; return { p, r, tot, best, ledig: tot - best }; }).filter(x => x.tot || x.best || rom.some(r => x.r[r]));
    const fRom = rom.filter(r => fv.some(x => x.r[r]) || r === "r1" || r === "r2");
    const lv = lagervarer().map(x => { const r = b[lvKey(x)] || {}; const tot = rom.reduce((s, y) => s + (r[y] || 0), 0); return { x, r, tot, st: status(x, tot) }; });
    const lRom = rom.filter(r => lv.some(y => y.r[r] || y.x.rom === r));
    const varsel = lv.filter(y => y.st.k !== "ok");
    const siste = [...l.rader].sort((a, b) => (b.tid || "").localeCompare(a.tid || "")).slice(0, 30);
    let kat = "";
    return `${!harTelling ? `<p class="notice"><b>Start med å telle opp alle rommene.</b> Beholdningen regnes fra bevegelsene som registreres i ordreboka. Den første opptellingen gir riktig utgangspunkt.</p>` : ""}
    ${varsel.length ? `<div class="card avvikkort"><h4>Må bestilles</h4>${varsel.map(y => `<div><b>${esc(y.x.navn)}</b>: ${lvMengde(y.tot, y.x)} igjen${y.st.uker !== null ? `, ca. ${nf(y.st.uker, 1)} uker med dagens forbruk` : ""}${tall(y.x.ledetidUker) ? `. Leveringstid ${nf(tall(y.x.ledetidUker))} uker` : ""}${y.x.leverandor ? `. Leverandør: ${esc(y.x.leverandor)}` : ""}. <span class="pill lg-${y.st.k}">${esc(y.st.t)}</span></div>`).join("")}</div>` : ""}
    <div class="toolbar">
      ${rom.map(r => `<button class="btn primary" data-act="lgTell" data-rom="${r}">Tell opp ${esc(romNavn(r))}</button>`).join("")}
    </div>
    <div class="toolbar">
      <button class="btn" data-act="lgSkjema" data-t="mottak">Varemottak</button>
      <button class="btn" data-act="lgSkjema" data-t="uttak">Uttak / forbruk</button>
      <button class="btn" data-act="lgSkjema" data-t="flytt">Flytt mellom rom</button>
      <button class="btn" data-act="lgSkjema" data-t="nyVare">Ny lagervare</button>
    </div>
    ${l.skjema ? visSkjema() : ""}
    <h3>Ferdigvarer</h3>
    <div class="tablewrap"><table class="grid lager"><thead><tr><th>Vare</th>${fRom.map(r => `<th class="n">${esc(romNavn(r))}</th>`).join("")}<th class="n">Totalt</th><th class="n">Bestilt, ikke sendt (uke ${u.w})</th><th class="n">Ledig</th></tr></thead><tbody>
    ${fv.length ? fv.map(x => `<tr><td>${esc(x.p.kort)}<div class="hint">#${esc(x.p.varenr)}</div></td>${fRom.map(r => `<td class="n ${(x.r[r] || 0) < 0 ? "neg" : ""}">${x.r[r] ? mengde(x.r[r], x.p) : "–"}</td>`).join("")}<td class="n"><b>${mengde(x.tot, x.p)}</b></td><td class="n">${x.best ? mengde(x.best, x.p) : "–"}</td><td class="n ${x.ledig < 0 ? "neg" : ""}">${mengde(x.ledig, x.p)}</td></tr>`).join("") : `<tr><td colspan="${fRom.length + 4}" class="hint">Ingen ferdigvarer på lager ennå.</td></tr>`}
    </tbody></table></div>
    <h3 style="margin-top:22px">Råvarer, emballasje og etiketter</h3>
    <div class="tablewrap"><table class="grid lager"><thead><tr><th>Vare</th>${lRom.map(r => `<th class="n">${esc(romNavn(r))}</th>`).join("")}<th class="n">Totalt</th><th class="n">Bestill under</th><th class="n">Forbruk/uke</th><th class="n">Uker igjen</th><th>Status</th><th></th></tr></thead><tbody>
    ${lv.length ? lv.map(y => { const h = y.x.kategori !== kat ? (kat = y.x.kategori, `<tr class="grp"><td colspan="${lRom.length + 7}"><b>${esc(kat || "Uten kategori")}</b></td></tr>`) : ""; return h + `<tr><td>${esc(y.x.navn)}${y.x.leverandor ? `<div class="hint">${esc(y.x.leverandor)}</div>` : ""}</td>${lRom.map(r => `<td class="n ${(y.r[r] || 0) < 0 ? "neg" : ""}">${y.r[r] ? lvMengde(y.r[r], y.x) : "–"}</td>`).join("")}<td class="n"><b>${lvMengde(y.tot, y.x)}</b></td><td class="n">${tall(y.x.minimum) ? lvMengde(tall(y.x.minimum), y.x) : "–"}</td><td class="n">${y.st.uke ? nf(y.st.uke, 1) : "–"}</td><td class="n">${y.st.uker !== null ? nf(y.st.uker, 1) : "–"}</td><td><span class="pill lg-${y.st.k}">${esc(y.st.t)}</span></td><td><button class="btn sm ghost" data-act="lgSkjema" data-t="endre" data-id="${esc(y.x.id)}">Endre</button></td></tr>`; }).join("") : `<tr><td colspan="${lRom.length + 7}" class="hint">Ingen lagervarer ennå. Trykk «Ny lagervare» for å legge inn fløte, beger, esker og etiketter.</td></tr>`}
    </tbody></table></div>
    <p class="hint">Ferdigvarer: inn fra ferdige ystinger (settes i ${esc(romNavn(innRom()))}), ut ved «Pakk og send» og svinn. Lagervarer: inn ved varemottak, ut ved uttak og opptelling. Fløte koblet til ystingen trekkes automatisk når ystingen avsluttes. Negativt tall betyr at noe ikke er registrert, og at rommet bør telles opp.</p>
    <div class="card form"><h3>Siste bevegelser</h3>
      <div class="tablewrap"><table class="grid"><thead><tr><th>Dato</th><th>Type</th><th>Vare</th><th class="n">Antall</th><th>Rom</th><th>Lot / best før</th><th>Merknad</th><th>Av</th><th></th></tr></thead><tbody>
      ${siste.length ? siste.map(r => { const manuell = ["telling", "flytt", "mottak", "forbruk"].includes(r.type) && !String(r.id).startsWith("fl-"); return `<tr><td>${esc(dmy(r.dato))}</td><td>${esc(TYPE[r.type] || r.type)}</td><td>${esc(navnFor(r.varenr))}</td><td class="n">${tall(r.antall) > 0 ? "+" : ""}${nf(tall(r.antall), tall(r.antall) % 1 ? 1 : 0)}</td><td>${esc(romNavn(r.rom))}</td><td>${esc(r.lot || "")}</td><td class="hint">${esc(r.merknad || "")}</td><td class="hint">${esc(r.av || "")}</td><td>${manuell ? `<button class="btn sm ghost danger" data-act="lgSlett" data-id="${esc(r.id)}" aria-label="Slett">${l.bekreft === r.id ? "Slett?" : "✕"}</button>` : ""}</td></tr>`; }).join("") : `<tr><td colspan="9" class="hint">Ingen bevegelser ennå.</td></tr>`}
      </tbody></table></div></div>
    <form class="card form" id="lgInnst">
      <h3>Rom</h3>
      <div class="row">${rom.map(r => `<label class="f">Navn<input class="txt" id="lgn-${r}" value="${esc(romNavn(r))}"></label>`).join("")}
        <label class="f">Nytt rom<input class="txt" id="lg-nytt" placeholder="f.eks. Fryserom"></label>
        <label class="f">Ny produksjon settes i<select class="txt" id="lg-inn">${romValg(innRom())}</select></label></div>
      <div class="actions"><button class="btn" type="submit">Lagre</button></div>
    </form>`;
  }
  function vareValg(medFerdig, valgt) {
    const lv = lagervarer().map(x => `<option value="${esc(lvKey(x))}" ${lvKey(x) === valgt ? "selected" : ""}>${esc(x.navn)} (${esc(x.enhet || "stk")})</option>`).join("");
    const fv = medFerdig ? varer().map(p => `<option value="${esc(p.varenr)}" ${p.varenr === valgt ? "selected" : ""}>${esc(p.kort)}</option>`).join("") : "";
    return medFerdig ? `<optgroup label="Lagervarer">${lv}</optgroup><optgroup label="Ferdigvarer">${fv}</optgroup>` : lv;
  }
  function visSkjema() {
    const l = L(), t = l.skjema.t, rom = romIds();
    if (t === "mottak" || t === "uttak") {
      if (!lagervarer().length) return `<div class="card form"><p class="hint">Legg inn lagervarene først med «Ny lagervare».</p><div class="actions"><button class="btn ghost" type="button" data-act="lgLukk">Lukk</button></div></div>`;
      const f = lagervarer()[0];
      return `<form class="card form" id="lgSkjema" data-id="${t}" data-bevar="1"><h3>${t === "mottak" ? "Varemottak" : "Uttak / forbruk"}</h3>
        <div class="row"><label class="f">Vare<select class="txt" id="ls-vare" data-act="lgVareRom">${vareValg(false)}</select></label>
        <label class="f">Antall<input class="txt" id="ls-antall" inputmode="decimal" required></label>
        <label class="f">Rom<select class="txt" id="ls-rom">${romValg(f.rom || "r3")}</select></label>
        ${t === "mottak" ? `<label class="f">Lot / best før<input class="txt" id="ls-lot" placeholder="valgfritt"></label>` : ""}</div>
        <label class="f">Merknad<input class="txt" id="ls-merk" placeholder="${t === "mottak" ? "f.eks. ordre fra Linea Lactis" : "f.eks. brukt i pakking"}"></label>
        <div class="actions"><button class="btn primary" type="submit">${t === "mottak" ? "Registrer mottak" : "Registrer uttak"}</button><button class="btn ghost" type="button" data-act="lgLukk">Avbryt</button></div></form>`;
    }
    if (t === "flytt") return `<form class="card form" id="lgSkjema" data-id="flytt" data-bevar="1"><h3>Flytt mellom rom</h3>
      <div class="row"><label class="f">Vare<select class="txt" id="ls-vare">${vareValg(true)}</select></label>
      <label class="f">Antall<input class="txt" id="ls-antall" inputmode="decimal" placeholder="f.eks. 12 eller 2dp" required></label>
      <label class="f">Fra<select class="txt" id="ls-fra">${romValg(rom[0])}</select></label>
      <label class="f">Til<select class="txt" id="ls-til">${romValg(rom[1])}</select></label></div>
      <div class="actions"><button class="btn primary" type="submit">Flytt</button><button class="btn ghost" type="button" data-act="lgLukk">Avbryt</button></div></form>`;
    const x = t === "endre" ? (l.lv.find(y => y.id === l.skjema.id) || {}) : {};
    const ny = t !== "endre";
    return `<form class="card form" id="lgSkjema" data-id="${ny ? "nyVare" : "endre"}" data-bevar="1"><h3>${ny ? "Ny lagervare" : "Endre " + esc(x.navn || "")}</h3>
      <div class="row"><label class="f">Navn<input class="txt" id="lv-navn" value="${esc(x.navn || "")}" placeholder="f.eks. Fløte 38 % 5 l" required></label>
      <label class="f">Kategori<input class="txt" id="lv-kat" list="lv-katl" value="${esc(x.kategori || "")}" placeholder="Råvare, Emballasje …"></label>
      <label class="f">Enhet<input class="txt" id="lv-enh" list="lv-enhl" value="${esc(x.enhet || "stk")}"></label>
      <label class="f">Står i<select class="txt" id="lv-rom">${romValg(x.rom || "r3")}</select></label></div>
      <datalist id="lv-katl">${KAT.map(k => `<option value="${k}">`).join("")}</datalist><datalist id="lv-enhl">${ENH.map(k => `<option value="${k}">`).join("")}</datalist>
      <div class="row"><label class="f">Bestill når det er under<input class="txt" id="lv-min" inputmode="decimal" value="${esc(x.minimum || "")}" placeholder="antall i enheten"></label>
      <label class="f">Leverandør<input class="txt" id="lv-lev" value="${esc(x.leverandor || "")}" placeholder="f.eks. Linea Lactis"></label>
      <label class="f">Leveringstid (uker)<input class="txt" id="lv-led" inputmode="decimal" value="${esc(x.ledetidUker || "")}"></label>
      ${ny ? `<label class="f">Antall på lager nå<input class="txt" id="lv-start" inputmode="decimal" placeholder="valgfritt"></label>` : ""}</div>
      <div class="row"><label class="f">Trekkes automatisk<select class="txt" id="lv-kob"><option value="">Nei, registreres ved uttak og telling</option><option value="flote" ${x.kobling === "flote" ? "selected" : ""}>Fløte brukt i ysting</option></select></label>
      <label class="f">Liter per enhet (for fløte)<input class="txt" id="lv-per" inputmode="decimal" value="${esc(x.perEnhet || "")}" placeholder="f.eks. 5 hvis enheten er 5 l kanne"></label></div>
      <label class="f">Merknad<input class="txt" id="lv-merk" value="${esc(x.merknad || "")}"></label>
      ${ny ? "" : `<label class="chk"><input type="checkbox" id="lv-aktiv" ${x.aktiv !== false ? "checked" : ""}> I bruk</label>`}
      <div class="actions"><button class="btn primary" type="submit">${ny ? "Legg til" : "Lagre"}</button><button class="btn ghost" type="button" data-act="lgLukk">Avbryt</button></div></form>`;
  }
  function visTelling() {
    const l = L(), rom = l.telling, b = beholdning();
    const lv = lagervarer().filter(x => (x.rom || "r3") === rom || (b[lvKey(x)] || {})[rom]);
    const fv = varer().filter(p => rom !== "r3" || (b[p.varenr] || {})[rom]);
    const rad = (key, navn, under, før, ph) => `<tr><td>${esc(navn)}<div class="hint">${esc(under || "")}</div></td><td class="n">${før}</td><td><input class="txt" id="lt-${esc(key)}" inputmode="decimal" style="width:110px;font-size:16px" placeholder="${esc(ph)}"></td></tr>`;
    return `<form class="card form" id="lgTellForm" data-bevar="1" data-id="${rom}">
      <h3>Opptelling: ${esc(romNavn(rom))}</h3>
      <p class="hint">Skriv det som faktisk står i rommet. La feltet stå tomt for varer du ikke har telt. Ferdigvarer i D-pak kan skrives som <b>3dp</b> eller <b>3dp+2</b>. Differansen føres som telling. Mangler en vare, legg den inn med «Ny lagervare» først.</p>
      <div class="tablewrap"><table class="grid"><thead><tr><th>Vare</th><th class="n">Etter ordreboka</th><th>Telt</th></tr></thead><tbody>
      ${lv.length ? `<tr class="grp"><td colspan="3"><b>Lagervarer</b></td></tr>` + lv.map(x => rad(lvKey(x), x.navn, x.kategori, lvMengde((b[lvKey(x)] || {})[rom] || 0, x), x.enhet || "stk")).join("") : ""}
      ${fv.length ? `<tr class="grp"><td colspan="3"><b>Ferdigvarer</b></td></tr>` + fv.map(p => rad(p.varenr, p.kort, p.pakning, mengde((b[p.varenr] || {})[rom] || 0, p), enhetNavn(p, 2))).join("") : ""}
      </tbody></table></div>
      <div class="actions"><button class="btn primary" type="submit">Lagre opptelling</button><button class="btn ghost" type="button" data-act="lgTellAvbryt">Avbryt</button></div></form>`;
  }

  /* ---------- hendelser ---------- */
  function klikk(el) {
    const a = el.dataset.act, l = L(); if (!a || !a.startsWith("lg")) return false;
    if (a === "lgVareRom") return true;
    if (a === "lgTell") { if (harLista()) { l.telling = el.dataset.rom; l.skjema = null; render(); window.scrollTo(0, 0); } }
    else if (a === "lgTellAvbryt") { const k = "lgTellForm|" + l.telling; l.telling = null; render(); S.utkast.delete(k); }
    else if (a === "lgSkjema") { if (!harLista()) return true; const t = el.dataset.t; l.skjema = l.skjema && l.skjema.t === t && l.skjema.id === el.dataset.id ? null : { t, id: el.dataset.id }; render(); setTimeout(() => document.getElementById("lgSkjema")?.scrollIntoView({ block: "nearest", behavior: "smooth" }), 30); }
    else if (a === "lgLukk") { const k = l.skjema ? "lgSkjema|" + l.skjema.t : ""; l.skjema = null; render(); if (k) S.utkast.delete(k); }
    else if (a === "lgSlett") { const id = el.dataset.id; if (l.bekreft === id) { l.bekreft = null; const r = l.rader.find(x => x.id === id); const par = r && r.type === "flytt" ? l.rader.filter(x => x.ref === r.ref && x.type === "flytt").map(x => x.id) : [id]; Promise.all(par.map(slettRad)).then(() => toast("Bevegelsen er slettet.")); render(); } else { l.bekreft = id; render(); setTimeout(() => { if (l.bekreft === id) { l.bekreft = null; render(); } }, 4000); } }
    return true;
  }
  function endring(e) {
    const t = e.target; if (t.id !== "ls-vare" || !t.closest("#lgSkjema")) return false;
    const lv = lvFra(t.value), r = document.getElementById("ls-rom"); if (lv && r && lv.rom) r.value = lv.rom; return true;
  }
  const antallFor = (key, s) => { const lv = lvFra(key); if (lv) { const v = String(s || "").trim(); if (!v) return 0; const n = Number(v.replace(",", ".")); return isFinite(n) ? n : null; } return parseAntall(s, prod(key), false); };
  /* Lukker skjemaet og fjerner utkastet etter at siden er tegnet på nytt (render lagrer ellers utkastet igjen). */
  function lukk(key) { const l = L(); l.skjema = null; l.telling = null; render(); S.utkast.delete(key); }
  function submit(form) {
    const l = L();
    if (form.id === "lgTellForm") {
      const rom = form.dataset.id, b = beholdning(), tid = Date.now().toString(36), jobber = []; let n = 0; const feil = [];
      for (const el of form.querySelectorAll("input[id^='lt-']")) {
        if (!el.value.trim()) continue; const key = el.id.slice(3);
        const telt = antallFor(key, el.value); if (telt === null) { feil.push(navnFor(key)); continue; }
        const før = (b[key] || {})[rom] || 0, diff = Math.round((telt - før) * 1000) / 1000; n++;
        if (diff) jobber.push(skrivRad({ id: `t-${tid}-${rom}-${key}`, dato: iDag(), varenr: key, antall: diff, rom, type: "telling", ref: `t-${tid}`, lot: "", merknad: `Telt ${nf(telt)}, var ${nf(før)}` }));
      }
      if (feil.length) { toast(`Forstår ikke antallet for ${feil.join(", ")}. Skriv f.eks. 12, 2,5, 3dp eller 3dp+2.`); return true; }
      if (!n) { toast("Skriv antall for minst én vare."); return true; }
      lukk("lgTellForm|" + rom);
      Promise.all(jobber).then(() => toast(`✓ Opptellingen av ${romNavn(rom)} er lagret (${n} varer, ${jobber.length} rettet).`));
      return true;
    }
    if (form.id === "lgSkjema") {
      const t = form.dataset.id;
      if (t === "mottak" || t === "uttak") {
        const key = $("#ls-vare").value, lv = lvFra(key), n = antallFor(key, $("#ls-antall").value), rom = $("#ls-rom").value;
        if (!(n > 0)) { toast("Skriv antall."); return true; }
        const id = (t === "mottak" ? "m-" : "u-") + Date.now().toString(36);
        skrivRad({ id, dato: iDag(), varenr: key, antall: t === "mottak" ? n : -n, rom, type: t === "mottak" ? "mottak" : "forbruk", ref: id, lot: ($("#ls-lot")?.value || "").trim(), merknad: $("#ls-merk").value.trim() })
          .then(() => toast(`✓ ${t === "mottak" ? "Mottak" : "Uttak"} registrert: ${nf(n)} ${lv?.enhet || ""} ${lv?.navn || ""}.`));
        lukk("lgSkjema|" + t); return true;
      }
      if (t === "flytt") {
        const key = $("#ls-vare").value, n = antallFor(key, $("#ls-antall").value), fra = $("#ls-fra").value, til = $("#ls-til").value;
        if (!(n > 0)) { toast("Skriv antall, f.eks. 12 eller 2dp."); return true; }
        if (fra === til) { toast("Velg to forskjellige rom."); return true; }
        const ref = "f-" + Date.now().toString(36), m = `${romNavn(fra)} → ${romNavn(til)}`;
        Promise.all([skrivRad({ id: ref + "-ut", dato: iDag(), varenr: key, antall: -n, rom: fra, type: "flytt", ref, lot: "", merknad: m }), skrivRad({ id: ref + "-inn", dato: iDag(), varenr: key, antall: n, rom: til, type: "flytt", ref, lot: "", merknad: m })])
          .then(() => toast(`✓ ${nf(n)} ${navnFor(key)} er flyttet til ${romNavn(til)}.`));
        lukk("lgSkjema|flytt"); return true;
      }
      /* ny eller endret lagervare */
      const navn = $("#lv-navn").value.trim(); if (!navn) { toast("Lagervaren må ha et navn."); return true; }
      const gammel = t === "endre" ? l.lv.find(y => y.id === l.skjema.id) : null;
      if (!gammel && l.lv.some(y => (y.navn || "").toLowerCase() === navn.toLowerCase())) { toast("Det finnes allerede en lagervare med dette navnet."); return true; }
      const x = { ...(gammel || { id: slug(navn).slice(0, 30) + "-" + Math.random().toString(36).slice(2, 5), sort: 50 }), navn, kategori: $("#lv-kat").value.trim(), enhet: $("#lv-enh").value.trim() || "stk", rom: $("#lv-rom").value, minimum: tall($("#lv-min").value), leverandor: $("#lv-lev").value.trim(), ledetidUker: tall($("#lv-led").value), kobling: $("#lv-kob").value, perEnhet: tall($("#lv-per").value), merknad: $("#lv-merk").value.trim(), aktiv: gammel ? !!$("#lv-aktiv")?.checked : true };
      const start = !gammel ? ($("#lv-start")?.value || "").trim() : "";
      lagreLv(x).then(() => { toast(`✓ ${navn} er ${gammel ? "lagret" : "lagt til"}.`); if (start) { const n = tall(start); if (n) skrivRad({ id: `t-${Date.now().toString(36)}-${x.rom}-${lvKey(x)}`, dato: iDag(), varenr: lvKey(x), antall: n, rom: x.rom, type: "telling", ref: "start", lot: "", merknad: `Startbeholdning ${nf(n)}` }); } });
      lukk("lgSkjema|" + t); return true;
    }
    if (form.id === "lgInnst") {
      const rom = {}; romIds().forEach(r => { const v = ($("#lgn-" + r)?.value || "").trim(); if (v) rom[r] = v; });
      const nytt = ($("#lg-nytt")?.value || "").trim(); if (nytt) { let i = 1; while (rom["r" + i] || ROM_STD["r" + i]) i++; rom["r" + i] = nytt; }
      l.innst = { rom, inn: $("#lg-inn").value }; skriv("innst/lager", { data: l.innst }).then(() => toast("✓ Lagret.")); render(); return true;
    }
    return false;
  }

  return { last, poll, vis, klikk, endring, submit, registrerSendt, angreSendt, produksjonInn, svinnUt, slettRef, romValg, romNavn, innRom };
})();
window.Lager = Lager;
