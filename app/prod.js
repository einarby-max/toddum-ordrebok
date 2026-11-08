/* Produksjon: melkehenting, ysting (tid, kontrollpunkter, resultat), kostnad per ysting, produsert mot bestilt og temperaturlogg fra kjølerommet.
   Data lagres i SharePoint-listene Ordrebok Melkehenting, Ordrebok Ysting, Ordrebok Temperatur og Ordrebok Innstillinger.
   Bilder av melketanken lagres i dokumentbiblioteket under Ysteri/Melkehenting. */
const Prod = (() => {
  /* Standardsatser fra Toddum Styringsmodell v10. Melkepris 7 kr/l bekreftet av Einar 08.11.2026. Kan endres under Produksjon → Satser. */
  const SATSER = {
    melkKrL: 7, floteKrL: 65.23, timeKr: 294.97,
    batchKr: 728.4,           // løpe 41,8 + kalsiumklorid 12,8 + kultur 448,8 + frakt 75 + vaskemiddel 150 (strøm og vann regnes for seg)
    kWh: 218, nettleieKr: 0.31, stromReserve: 1.1, vannKrM3: 0, omrade: "NO3",
    tempGrense: 4, melkTempGrense: 4,
    emb: { "0001":4.21, "0002":22.97, "0010":4.21, "0011":22.97, "0020":4.21, "0022":22.97, "0009":4.21, "0008":10.73, "0041":34.71, "0031":7.88, "0030":22.97, "0032":0, "0042":8.19 },
    modell: { "0001":13.6, "0002":213.72, "0010":15.65, "0011":175.71, "0020":43.59, "0022":426.38, "0030":268.37, "0031":30.12, "0009":46.27, "0008":88.61, "0041":651.51, "0032":81.56, "0042":131.59 },   // v10, melk 7 kr/l
  };
  /* Strøm per måned fra NTE-fakturaene (målepunkt 707057500065627076, eks. mva) og solcelleproduksjon avlest i SolisCloud 31.10.2026 (ca., lest av grafen).
     nett = kjøpt fra nettet (kWh), solgt = innmating (kWh), sol = produsert (kWh), kr = strøm + nettleie (kr eks. mva). Oppdateres når nye fakturaer kommer. */
  const ENERGI = [
    { mnd: "2025-12", nett: 9480, spot: 0.610, solgt: 0, sol: null, kap: "50–75", kapKr: 2856, kr: 12363 },
    { mnd: "2026-01", nett: 12157, spot: 1.293, solgt: 0, sol: 150, kap: "75–100", kapKr: 3750, kr: 23413 },
    { mnd: "2026-02", nett: 10461, spot: 1.254, solgt: 4, sol: 630, kap: "75–100", kapKr: 3750, kr: 20242 },
    { mnd: "2026-03", nett: 8817, spot: 0.815, solgt: 349, sol: 1780, kap: "75–100", kapKr: 3750, kr: 13499 },
    { mnd: "2026-04", nett: 7906, spot: 1.095, solgt: 732, sol: 3230, kap: "75–100", kapKr: 3750, kr: 14075 },
    { mnd: "2026-05", nett: 6576, spot: 0.826, solgt: 783, sol: 3320, kap: "75–100", kapKr: 3750, kr: 10603 },
    { mnd: "2026-06", nett: 6826, spot: 0.652, solgt: 836, sol: 2870, kap: "75–100", kapKr: 3750, kr: 9886 },
    { mnd: "2026-07", nett: 6698, spot: 0.763, solgt: 678, sol: 2790, kap: "75–100", kapKr: 3937, kr: 10629 },
    { mnd: "2026-08", nett: 6783, spot: 1.032, solgt: 411, sol: 2100, kap: "75–100", kapKr: 3937, kr: 12717 },
    { mnd: "2026-09", nett: 6417, spot: 1.084, solgt: 510, sol: 1880, kap: "100–150", kapKr: 5698, kr: 14229 },
    { mnd: "2026-10", nett: null, spot: null, solgt: null, sol: 1004, kap: "", kapKr: null, kr: null },
  ];
  const SOL = { kWp: 28.475, totalt: 64682, iÅr: 19918, fra: "desember 2023" };
  const ÅRSAK = ["Utgått dato", "Skadet eller lekk", "Kvalitet", "Smaking og prøver", "Annet"];
  const KONTROLL = ["pH", "Temperatur °C", "Pasteurisering °C", "Pasteurisering tid (s)", "Melketemperatur °C", "Kultur", "Kalsiumklorid (CaCl2)", "Løpe", "Ostemasse kg", "Annet"];
  /* Ystelogg etter papirskjemaet «Produksjon av mozzarella»: mengde, klokkeslett, temperatur og pH per trinn. */
  const LOGG = [
    ["kultur", "Kultur (type)"], ["kulturPoser", "Kultur – antall poser"], ["cacl2", "Kalsiumklorid (CaCl2)"], ["lope", "Løpe"],
    ["roring", "Røring"], ["koag", "Koagulering"], ["skj10", "Skjæring 10 %"], ["hvile", "Hvile"], ["skj45", "Skjæring 45 %"],
    ["hvileMyse", "Hvile i myse"], ["drenering", "Drenering"], ["strekking", "Strekking – vann"], ["forming", "Forming"], ["ferdig", "Ferdig produkt"],
  ];
  const OSTER = ["Mozzarella", "Burrata", "Stracciatella", "Fior di latte", "Ricotta", "Halloumi"];
  const batchNr = y => Math.max(1, tall(y && y.batchNr) || 1);
  const navnY = y => `${dmy(y.dato)}${batchNr(y) > 1 || P().ystinger.some(z => z.id !== y.id && z.dato === y.dato) ? ` · batch ${batchNr(y)}` : ""}${y.osteType ? ` · ${y.osteType}` : ""}`;
  const åpne = () => P().ystinger.filter(y => y.status !== "ferdig").sort((a, b) => (a.dato || "").localeCompare(b.dato || "") || batchNr(a) - batchNr(b));
  const P = () => S.pr || (S.pr = { pv: "oversikt", ystinger: [], melk: [], temp: [], svinn: [], spot: {}, spotRaw: {}, alleOrdre: null, innst: null, edit: null, loaded: false, laster: false, tempValg: 7, tempForslag: null, melkBilde: null, melkBildeUrl: "", lagrerT: null, feilTilgang: false });
  const sats = () => ({ ...SATSER, ...(P().innst || {}), emb: { ...SATSER.emb, ...((P().innst || {}).emb || {}) }, modell: { ...SATSER.modell, ...((P().innst || {}).modell || {}) } });
  const kr = n => nf(Math.round(n), 0) + " kr";
  const tall = v => { const n = Number(String(v ?? "").replace(/\s/g, "").replace(",", ".")); return isFinite(n) ? n : 0; };
  const nå = () => new Date().toISOString();
  const klokke = iso => { if (!iso) return ""; const d = new Date(iso); return `${pad(d.getHours())}.${pad(d.getMinutes())}`; };
  const datoKl = iso => { if (!iso) return ""; const d = new Date(iso); return `${pad(d.getDate())}.${pad(d.getMonth() + 1)} kl. ${pad(d.getHours())}.${pad(d.getMinutes())}`; };
  const brukerNavn = () => (S.bruker && (S.bruker.navn || S.bruker.epost)) || "";
  const nyId = pre => `${pre}-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 5)}`;
  function feilLagring(e) { if (e && e.status === 403) { P().feilTilgang = true; toast("Du har ikke skrivetilgang til produksjonslistene. Be Einar eller Silje gi deg tilgang."); } else toast(errMsg(e)); planRender(); }

  /* ---------- lasting ---------- */
  async function last() {
    const p = P(); if (p.laster) return; p.laster = true;
    try {
      const [y, m, i, sv] = await Promise.all([Store.hent("ysting"), Store.hent("melk"), Store.hent("innst"), Store.hent("svinn").catch(() => [])]);
      p.ystinger = y; p.melk = m; p.svinn = sv; const s = i.find(x => x.id === "satser"); p.innst = s ? s.data : null; p.loaded = true;
      if (p.edit) { const fersk = y.find(x => x.id === p.edit.id); if (fersk && !p.lagrerT && (fersk.endret || "") > (p.edit.endret || "")) p.edit = fersk; }
      planRender();
      Store.hent("temperatur").then(t => { p.temp = t.sort((a, b) => (a.fra || "").localeCompare(b.fra || "")); planRender(); }).catch(() => {});
    } catch (e) { p.loaded = true; planRender(); } finally { p.laster = false; }
  }
  function poll() { if (S.tab === "prod" && !P().lagrerT) last(); }

  /* ---------- beregning ---------- */
  const arr = v => Array.isArray(v) ? v : [];
  function minutter(y) { if (!y) return 0; if (y.minutterManuelt) return Math.round(tall(y.minutter)); let ms = 0; arr(y.tid).forEach(t => { if (t.fra) ms += (t.til ? new Date(t.til) : new Date()) - new Date(t.fra); }); return Math.round(ms / 60000); }
  const pågår = y => arr(y.tid).some(t => t.fra && !t.til);
  /* Strømpris (spot, uten mva) for prisområdet fra hvakosterstrommen.no. Hentes én gang per dato. */
  function hentSpot(dato) {
    const p = P(); if (!dato || dato in p.spot) return; p.spot[dato] = null;
    const [y, m, d] = dato.split("-");
    fetch(`https://www.hvakosterstrommen.no/api/v1/prices/${y}/${m}-${d}_${sats().omrade}.json`).then(r => r.ok ? r.json() : Promise.reject(r.status))
      .then(j => { p.spotRaw[dato] = j; p.spot[dato] = j.reduce((s, x) => s + x.NOK_per_kWh, 0) / j.length; planRender(); })
      .catch(() => { p.spot[dato] = 0; });
  }
  function spotPris(y) { if (tall(y.strompris) > 0) return tall(y.strompris); const v = P().spot[y.dato]; if (v === undefined) hentSpot(y.dato); return v > 0 ? v : sats().stromReserve; }
  const vannM3 = y => Math.max(0, tall(y.vannSlutt) - tall(y.vannStart));
  /* Råvarekostnaden ystingen selv har hatt for ostemassen: melk, kultur m.m., strøm og vann (ikke fløte og emballasje, som følger pakkingen). */
  function egenRaa(y) { const s = sats(), melk = tall(y.melkLiter); if (!melk) return 0; return melk * s.melkKrL + s.batchKr + s.kWh * (spotPris(y) + s.nettleieKr) + vannM3(y) * s.vannKrM3; }
  function overførtUtKr(y) { const om = tall(y.ostemasseKg), ut = tall(y.overfortUtKg); return om > 0 && ut > 0 ? egenRaa(y) * Math.min(1, ut / om) : 0; }
  function overførtInnKr(y) { return arr(y.overfortInn).reduce((sum, x) => { const src = P().ystinger.find(z => z.id === x.fra); const om = tall(src?.ostemasseKg); return sum + (src && om ? egenRaa(src) / om * tall(x.kg) : 0); }, 0); }
  function brukt(src, unntak) { return P().ystinger.filter(z => z.id !== unntak).reduce((s, z) => s + arr(z.overfortInn).filter(x => x.fra === src.id).reduce((a, x) => a + tall(x.kg), 0), 0); }
  const restUt = (src, unntak) => tall(src.overfortUtKg) - brukt(src, unntak);
  function beregn(y) {
    const s = sats(), prd = y.produsert || {};
    const melk = tall(y.melkLiter), flote = tall(y.flote), min = minutter(y), pers = Math.max(1, tall(y.personer) || 1);
    const timer = min / 60 * pers;
    let emb = 0, kg = 0, verdi = 0, modell = 0, enh = 0;
    Object.entries(prd).forEach(([v, n]) => { n = tall(n); if (!n) return; const p = prod(v); enh += n; emb += n * (s.emb[v] || 0); kg += kgNetto(n, p); verdi += n * (p?.pris || 0); modell += n * (s.modell[v] || 0); });
    const pris = melk ? spotPris(y) : 0;
    const k = { melk: melk * s.melkKrL, flote: flote * s.floteKrL, batch: melk ? s.batchKr : 0, strom: melk ? s.kWh * (pris + s.nettleieKr) : 0, vann: melk ? vannM3(y) * s.vannKrM3 : 0, emb, arbeid: timer * s.timeKr, ut: -overførtUtKr(y), inn: overførtInnKr(y) };
    k.sum = k.melk + k.flote + k.batch + k.strom + k.vann + k.emb + k.arbeid + k.ut + k.inn;
    const om = tall(y.ostemasseKg);
    return { melk, flote, min, pers, timer, kg, enh, verdi, modell, k, pris, om, perKg: kg ? k.sum / kg : 0, db: verdi - k.sum, dbPst: verdi ? (verdi - k.sum) / verdi : 0, utbytte: melk ? (om ? om : kg) / melk * 100 : 0, utbytteOm: !!om, timerPer100: kg ? timer / kg * 100 : 0 };
  }
  /* Batch 1 får vanlig lot (som ordrene), batch 2 og videre samme dag får -2, -3 … bak. */
  const lotFor = (dato, nr) => batchFra(dato) + (nr > 1 ? "-" + nr : "");
  const baseLot = lot => String(lot || "").replace(/-\d+$/, "");
  function bestiltUke(key) { const sum = {}; if (key !== uke().key) return null; S.ordre.forEach(o => linjer(o).forEach(l => sum[l.varenr] = (sum[l.varenr] || 0) + l.antall)); return sum; }
  function produsertUke(key) { const sum = {}; P().ystinger.filter(y => y.uke === key).forEach(y => Object.entries(y.produsert || {}).forEach(([v, n]) => sum[v] = (sum[v] || 0) + tall(n))); return sum; }
  function sporing(lot) { const ut = []; S.ordre.forEach(o => linjer(o).forEach(l => { const b = l.batch || batchFra(o.prodDato); if (b && b === baseLot(lot)) ut.push({ kunde: o.kundeNavn, varenr: l.varenr, antall: l.antall }); })); return ut; }
  const ledigMelk = () => P().melk.filter(m => !m.ystingId).sort((a, b) => (a.tid || "").localeCompare(b.tid || ""));

  /* ---------- lagring ---------- */
  function lagreYsting(y, straks) {
    const p = P(); y.endret = nå(); p.edit = y;
    const i = p.ystinger.findIndex(x => x.id === y.id); if (i >= 0) p.ystinger[i] = y; else p.ystinger.push(y);
    clearTimeout(p.lagrerT);
    const gjør = () => { p.lagrerT = null; const { id, ...d } = y; return skriv("ysting/" + id, d).catch(feilLagring); };
    if (straks) return gjør();
    p.lagrerT = setTimeout(gjør, 700); return Promise.resolve();
  }
  function lagreMelk(m) { const p = P(); const i = p.melk.findIndex(x => x.id === m.id); if (i >= 0) p.melk[i] = m; else p.melk.push(m); const { id, ...d } = m; return skriv("melk/" + id, d).catch(e => { feilLagring(e); throw e; }); }
  function nyYsting() {
    const d = iDag(), u = isoUke(new Date());
    const ml = ledigMelk(); const liter = ml.reduce((s, m) => s + tall(m.liter), 0);
    const nr = P().ystinger.filter(z => z.dato === d).reduce((m, z) => Math.max(m, batchNr(z)), 0) + 1;
    const y = { id: nyId("ysting"), dato: d, uke: u.key, batchNr: nr, lot: lotFor(d, nr), status: "pågår", melkLiter: liter, melkIds: ml.map(m => m.id), flote: 0, tid: [], personer: 2, minutter: 0, kontroller: [], produsert: {}, esker: 0, svinnKg: 0, merknad: "", registrertAv: brukerNavn(), ystetAv: brukerNavn(), osteType: "", logg: {} };
    ml.forEach(m => lagreMelk({ ...m, ystingId: y.id }).catch(() => {}));
    P().pv = "ysting"; lagreYsting(y, true); render();
  }

  /* ---------- visning ---------- */
  function vis() {
    const p = P(); if (!p.loaded) { last(); return `<div class="empty"><h3>Henter produksjonsdata …</h3></div>`; }
    const fane = (k, l) => `<button class="btn ${p.pv === k ? "primary" : ""}" data-act="prVis" data-v="${k}">${l}</button>`;
    const meny = `<div class="toolbar">${fane("oversikt", "Oversikt")}${fane("melk", "Hent melk")}${fane("ysting", p.edit ? "Ysting" : "Ny ysting")}${fane("svinn", "Svinn")}${fane("dash", "Dashbord")}${fane("temp", "Kjølerom")}${fane("satser", "Satser")}</div>`;
    const tilg = p.feilTilgang ? `<p class="notice"><b>Ikke lagret.</b> Kontoen din har ikke skrivetilgang til produksjonslistene i SharePoint. Be Einar eller Silje gi deg tilgang.</p>` : "";
    let h = "";
    if (p.pv === "melk") h = visMelk(); else if (p.pv === "ysting") h = visYsting(); else if (p.pv === "temp") h = visTemp(); else if (p.pv === "satser") h = visSatser(); else if (p.pv === "svinn") h = visSvinn(); else if (p.pv === "dash") h = visDashbord(); else h = visOversikt();
    return `<h2>Produksjon</h2>${meny}${tilg}${h}`;
  }

  function visOversikt() {
    const p = P(), u = uke(), ml = ledigMelk();
    const åp = åpne(), idag = iDag(), iDagN = p.ystinger.filter(y => y.dato === idag).length;
    const best = bestiltUke(u.key) || {}, prodd = produsertUke(u.key);
    const varer = prodAlle().filter(x => best[x.varenr] || prodd[x.varenr]);
    const siste = [...p.ystinger].sort((a, b) => (b.dato || "").localeCompare(a.dato || "") || batchNr(b) - batchNr(a)).slice(0, 12);
    const d30 = new Date(Date.now() - 30 * 864e5).toISOString().slice(0, 10);
    const ferdig30 = p.ystinger.filter(y => y.status === "ferdig" && y.dato >= d30).map(beregn);
    const sum = (f) => ferdig30.reduce((s, b) => s + f(b), 0);
    const t30 = { kg: sum(b => b.kg), melk: sum(b => b.melk), timer: sum(b => b.timer), kost: sum(b => b.k.sum), verdi: sum(b => b.verdi) };
    const sisteTemp = p.temp[p.temp.length - 1];
    return `
    <div class="toolbar">
      <button class="btn primary lagre" data-act="prVis" data-v="melk">Hent melk</button>
      <button class="btn primary lagre" data-act="prNy">${iDagN ? `Start ysting nr. ${iDagN + 1} i dag` : "Start ny ysting"}</button>
    </div>
    ${åp.length ? `<div class="card form"><h3>Ystinger som ikke er avsluttet</h3>${åp.map(y => `<div class="actions"><button class="btn ${y.dato === idag ? "primary" : ""}" data-act="prApne" data-id="${esc(y.id)}">Fortsett ${esc(navnY(y))}</button>${y.dato < idag ? `<span class="hint" style="color:var(--warn)">Fra en tidligere dag – avslutt den når den er ferdig, så dagens registreringer ikke havner her.</span>` : `<span class="hint">I dag · lot ${esc(y.lot)}</span>`}</div>`).join("")}</div>` : ""}
    ${ml.length ? `<p class="notice"><b>Melk klar i ysteriet:</b> ${ml.map(m => `${nf(tall(m.liter))} l (${esc(datoKl(m.tid))}, ${nf(tall(m.temp), 1)} °C)`).join(" · ")}. Den kobles til neste ysting.</p>` : ""}
    <div class="stats"><span>Siste 30 dager: <b>${ferdig30.length}</b> ystinger</span><span><b>${nf(t30.melk)}</b> l melk</span><span><b>${nf(t30.kg, 1)}</b> kg ost</span>
      <span>Utbytte <b>${t30.melk ? nf(t30.kg / t30.melk * 100, 1) : "–"}</b> kg/100 l</span><span><b>${nf(t30.timer, 1)}</b> timer</span>
      <span>Kost <b>${t30.kg ? nf(t30.kost / t30.kg, 0) : "–"}</b> kr/kg</span><span>DB <b>${t30.verdi ? nf((t30.verdi - t30.kost) / t30.verdi * 100, 0) + " %" : "–"}</b></span></div>
    <div class="two">
      <div class="card form"><h3>Uke ${u.w}: bestilt og produsert</h3>
        ${varer.length ? `<div class="tablewrap"><table class="grid"><thead><tr><th>Vare</th><th class="n">Bestilt</th><th class="n">Produsert</th><th class="n">Differanse</th></tr></thead><tbody>
        ${varer.map(x => { const b = best[x.varenr] || 0, pr = prodd[x.varenr] || 0, d = pr - b; return `<tr><td><span class="chip g-${esc(x.gruppe)}">${esc(x.kort)}</span></td><td class="n">${nf(b)}</td><td class="n">${nf(pr)}</td><td class="n" style="color:${d < 0 ? "var(--danger)" : d > 0 ? "var(--warn)" : "var(--ok)"};font-weight:700">${d > 0 ? "+" : ""}${nf(d)}</td></tr>`; }).join("")}
        </tbody></table></div><p class="hint">Rød = mangler til ordrene. Gul = mer produsert enn bestilt. Bytt uke øverst for å se andre uker.</p>` : `<p class="hint">Ingen ordre eller ystinger i uke ${u.w}.</p>`}
      </div>
      <div class="card form"><h3>Kjølerommet</h3>
        ${sisteTemp ? `${tempGraf(tempPunkter(7), 7, 160)}<p class="hint">Siste logg: ${esc(dmyLang((sisteTemp.til || "").slice(0, 10)))} · høyeste ${nf(tall(sisteTemp.max), 1)} °C · ${nf(tall(sisteTemp.over))} målinger over ${nf(sats().tempGrense)} °C</p>` : `<p class="hint">Ingen temperaturlogg lastet opp ennå. Gå til «Kjølerom» og last opp fila fra USB-loggeren.</p>`}
      </div>
    </div>
    <h3 style="font-family:var(--f-display);font-weight:400;margin:18px 0 8px">Siste ystinger</h3>
    ${siste.length ? `<div class="tablewrap"><table class="grid"><thead><tr><th>Dato</th><th>Ost</th><th>Lot</th><th class="n">Melk (l)</th><th class="n">Timer</th><th class="n">Kg ost</th><th class="n">Kg/100 l</th><th class="n">Kost</th><th class="n">Kr/kg</th><th class="n">DB</th><th>Status</th></tr></thead><tbody>
    ${siste.map(y => { const b = beregn(y); return `<tr data-act="prApne" data-id="${esc(y.id)}" style="cursor:pointer"><td>${esc(dmy(y.dato))}${batchNr(y) > 1 ? ` <span class="hint">#${batchNr(y)}</span>` : ""}</td><td>${esc(y.osteType || "")}</td><td>${esc(y.lot)}</td><td class="n">${nf(b.melk)}</td><td class="n">${nf(b.timer, 1)}</td><td class="n">${nf(b.kg, 1)}</td><td class="n">${b.melk ? nf(b.utbytte, 1) : "–"}</td><td class="n">${nf(Math.round(b.k.sum))}</td><td class="n">${b.kg ? nf(Math.round(b.perKg)) : "–"}</td><td class="n">${b.verdi ? nf(b.dbPst * 100) + " %" : "–"}</td><td><span class="pill ${y.status === "ferdig" ? "st-sendt" : "st-produsert"}">${y.status === "ferdig" ? "Ferdig" : "Pågår"}</span></td></tr>`; }).join("")}
    </tbody></table></div>` : `<div class="empty"><h3>Ingen ystinger registrert ennå</h3>Trykk «Start ny ysting» når dere begynner.</div>`}`;
  }


  /* Kontrollpunkter registrert en annen dag enn ystingens dato (typisk fordi gårsdagens ysting sto åpen). Kan flyttes uten at noe slettes. */
  function visFeilDag(y) {
    const andre = {}; arr(y.kontroller).forEach(k => { const d = (k.tid || "").slice(0, 10); if (d && d !== y.dato) andre[d] = (andre[d] || 0) + 1; });
    return Object.entries(andre).map(([d, n]) => {
      const mål = P().ystinger.filter(z => z.id !== y.id && z.dato === d).sort((a, b) => batchNr(a) - batchNr(b));
      return `<p class="notice">${n} kontrollpunkt${n > 1 ? "er" : ""} her er registrert ${esc(dmyLang(d))}, ikke ${esc(dmy(y.dato))}. ${mål.map(z => `<button type="button" class="btn sm" data-act="prFlyttK" data-d="${d}" data-til="${esc(z.id)}">Flytt til ${esc(navnY(z))}</button>`).join(" ")} <button type="button" class="btn sm" data-act="prFlyttK" data-d="${d}" data-til="">Flytt til en ny ysting ${esc(dmy(d))}</button></p>`;
    }).join("");
  }
  function visLogg(y) {
    const l = y.logg && typeof y.logg === "object" ? y.logg : {};
    const inp = (r, f, ph, mode) => `<input class="txt" id="pl-${r}-${f}" data-pl="${r}|${f}" ${mode ? `inputmode="${mode}"` : ""} placeholder="${ph}" value="${esc((l[r] || {})[f] || "")}" style="width:${f === "mengde" ? 52 : 42}px;font-size:15px">`;
    return `<h4>2. Ystelogg</h4>
      <div class="row">
        <label class="f">Type ost i dag<input class="txt" id="py-osteType" data-pf="osteType" list="pr-oster" value="${esc(y.osteType || "")}" placeholder="f.eks. Mozzarella"></label>
        <label class="f">Ystet av<input class="txt" id="py-ystetAv" data-pf="ystetAv" value="${esc(y.ystetAv || "")}"></label>
        <label class="f">Starttemperatur melk (°C)<input class="txt" id="py-startTemp" data-pf="startTemp" inputmode="decimal" value="${esc(y.startTemp || "")}"></label>
        <label class="f">Varmeelementer i bruk<input class="txt" id="py-varmeElementer" data-pf="varmeElementer" inputmode="numeric" value="${esc(y.varmeElementer || "")}"></label>
        <label class="f">Temp. når varmen slås av (°C)<input class="txt" id="py-tempVarmeAv" data-pf="tempVarmeAv" inputmode="decimal" value="${esc(y.tempVarmeAv || "")}"></label>
      </div>
      <datalist id="pr-oster">${OSTER.map(o => `<option value="${o}">`).join("")}</datalist>
      <div class="tablewrap"><table class="grid ylogg"><thead><tr><th>Trinn</th><th>Mengde</th><th>Kl.</th><th>Temp °C</th><th>pH</th></tr></thead><tbody>
      ${LOGG.map(([r, n]) => `<tr><td>${esc(n)}</td><td>${inp(r, "mengde", r === "kultur" ? "type" : "")}</td><td><div style="display:flex;gap:4px">${inp(r, "kl", "kl.", "decimal")}<button type="button" class="btn sm" data-act="prNaa" data-r="${r}" title="Sett klokkeslett nå" style="padding:4px 6px">Nå</button></div></td><td>${inp(r, "temp", "", "decimal")}</td><td>${inp(r, "ph", "", "decimal")}</td></tr>`).join("")}
      </tbody></table></div>
      <p class="hint">Fyll inn underveis. Trykk «Nå» for å sette klokkeslettet. Alt lagres med en gang.</p>`;
  }
  function visInn(y) {
    const kilder = P().ystinger.filter(z => z.id !== y.id && tall(z.overfortUtKg) > 0 && (restUt(z) > 0.05 || arr(y.overfortInn).some(x => x.fra === z.id)) && (z.dato || "") <= (y.dato || "9"));
    const inn = arr(y.overfortInn);
    if (!kilder.length && !inn.length) return "";
    return `<p class="hint" style="margin-bottom:0"><b>Ostemasse fra tidligere ysting:</b> ${inn.length ? inn.map((x, i) => { const z = P().ystinger.find(q => q.id === x.fra); return `${nf(tall(x.kg), 1)} kg fra ${esc(z ? dmy(z.dato) + " (lot " + z.lot + ")" : "ukjent")} <button type="button" class="linkbtn" data-act="prFjernInn" data-i="${i}">fjern</button>`; }).join(" · ") : "ingen"}</p>
      ${kilder.some(z => restUt(z) > 0.05) ? `<div class="row"><label class="f">Fra ysting<select class="txt" id="pi-fra">${kilder.filter(z => restUt(z) > 0.05).map(z => `<option value="${esc(z.id)}">${esc(dmy(z.dato))} · lot ${esc(z.lot)} · ${nf(restUt(z), 1)} kg igjen</option>`).join("")}</select></label>
        <label class="f">Kg brukt i dag<input class="txt" id="pi-kg" inputmode="decimal"></label></div>
        <div class="actions"><button type="button" class="btn" data-act="prLeggInn">Bruk ostemasse fra tidligere ysting</button></div>` : ""}`;
  }

  function visSvinn() {
    const p = P(), siste = [...p.svinn].sort((a, b) => (b.dato || "").localeCompare(a.dato || "") || (b.id || "").localeCompare(a.id || "")).slice(0, 25);
    const varer = prodSort().filter(x => x.gruppe !== "Melk");
    return `<form class="card form" id="prSvinnForm" data-bevar="1">
      <h3>Registrer svinn etter produksjon</h3>
      <p class="hint">Ferdigvarer som kastes eller ikke kan selges: utgått dato, skadet eller lekk emballasje, dårlig kvalitet eller brukt til smaking. Svinn under selve ystingen føres på ystingen.</p>
      <div class="row">
        <label class="f">Dato<input type="date" class="txt" id="sv-dato" value="${iDag()}"></label>
        <label class="f">Vare<select class="txt" id="sv-vare">${varer.map(x => `<option value="${esc(x.varenr)}">${esc(x.kort)}</option>`).join("")}</select></label>
        <label class="f">Antall<input class="txt" id="sv-antall" inputmode="decimal" required></label>
        <label class="f">Årsak<select class="txt" id="sv-arsak">${ÅRSAK.map(a => `<option>${a}</option>`).join("")}</select></label>
        <label class="f">Lot (valgfritt)<input class="txt" id="sv-lot" inputmode="numeric"></label>
      </div>
      <label class="f">Merknad<input class="txt" id="sv-merk"></label>
      <div class="actions"><button class="btn primary lagre" type="submit">Registrer svinn</button></div>
    </form>
    <h3 style="font-family:var(--f-display);font-weight:400;margin:18px 0 8px">Siste svinn</h3>
    ${siste.length ? `<div class="tablewrap"><table class="grid"><thead><tr><th>Dato</th><th>Vare</th><th class="n">Antall</th><th class="n">Kg</th><th class="n">Verdi</th><th>Årsak</th><th>Lot</th><th>Av</th><th></th></tr></thead><tbody>
    ${siste.map(x => { const v = prod(x.varenr); return `<tr><td>${esc(dmy(x.dato))}</td><td>${esc(v?.kort || x.varenr)}</td><td class="n">${nf(tall(x.antall))}</td><td class="n">${nf(kgNetto(tall(x.antall), v), 1)}</td><td class="n">${nf(Math.round(tall(x.antall) * (v?.pris || 0)))}</td><td>${esc(x.arsak)}${x.merknad ? `<div class="hint">${esc(x.merknad)}</div>` : ""}</td><td>${esc(x.lot || "")}</td><td class="hint">${esc(x.av || "")}</td><td><button class="btn sm ghost danger" data-act="prSlettSvinn" data-id="${esc(x.id)}" aria-label="Slett">✕</button></td></tr>`; }).join("")}
    </tbody></table></div>` : `<p class="hint">Ingen svinn registrert ennå.</p>`}`;
  }

  /* ---------- dashbord ---------- */
  function ukeNøkler(n) { const ut = [], d = new Date(); for (let i = n - 1; i >= 0; i--) { const x = new Date(d); x.setDate(x.getDate() - 7 * i); ut.push(isoUke(x)); } return ut; }
  function søyler(serier, etiketter, høyde = 180) {
    const W = 640, H = høyde, L = 40, B = 22, T = 8, n = etiketter.length, maks = Math.max(1, ...serier.flatMap(s => s.v));
    const bw = (W - L - 6) / n, gw = bw * 0.8 / serier.length;
    const Y = v => T + (1 - v / maks) * (H - T - B);
    const rå = maks / 4, mag = Math.pow(10, Math.floor(Math.log10(rå))), steg = [1, 2, 5, 10].find(m => m * mag >= rå) * mag;
    let g = ""; for (let v = 0; v <= maks; v += steg) g += `<line x1="${L}" x2="${W}" y1="${Y(v)}" y2="${Y(v)}" stroke="var(--line)"/><text x="${L - 6}" y="${Y(v) + 4}" font-size="11" text-anchor="end" fill="var(--muted)">${nf(v)}</text>`;
    etiketter.forEach((e, i) => { g += `<text x="${L + bw * i + bw / 2}" y="${H - 6}" font-size="11" text-anchor="middle" fill="var(--muted)">${esc(e)}</text>`; serier.forEach((s, j) => { const v = s.v[i] || 0; g += `<rect x="${L + bw * i + bw * 0.1 + gw * j}" y="${Y(v)}" width="${gw - 1}" height="${Math.max(0, H - B - Y(v))}" fill="${s.farge}" rx="2"><title>${esc(s.navn)}: ${nf(v, 1)}</title></rect>`; }); });
    return `<svg viewBox="0 0 ${W} ${H}" style="width:100%;height:auto;display:block" role="img">${g}</svg><div class="actions" style="gap:14px;margin-top:4px">${serier.map(s => `<span class="hint"><span style="display:inline-block;width:10px;height:10px;border-radius:2px;background:${s.farge};margin-right:5px"></span>${esc(s.navn)}</span>`).join("")}</div>`;
  }
  function visDashbord() {
    const p = P();
    if (!p.alleOrdre) { if (!p.henterOrdre) { p.henterOrdre = true; Store.hent("ordre").then(o => { p.alleOrdre = o; planRender(); }).catch(e => { p.alleOrdre = []; toast(errMsg(e)); }).finally(() => { p.henterOrdre = false; }); } return `<div class="empty"><h3>Henter ordre for alle uker …</h3></div>`; }
    const uker = ukeNøkler(12), key = uker.map(u => u.key);
    const rad = {}; key.forEach(k => rad[k] = { ordre: 0, kgBestilt: 0, verdiBestilt: 0, ystinger: 0, melk: 0, kg: 0, timer: 0, kost: 0, verdi: 0, svinnKg: 0, svinnKr: 0 });
    p.alleOrdre.forEach(o => { const r = rad[o.uke]; if (!r) return; const ls = linjer(o); if (!ls.length) return; r.ordre++; ls.forEach(l => { r.kgBestilt += kgNetto(l.antall, l.p); r.verdiBestilt += l.antall * (l.p?.pris || 0); }); });
    p.ystinger.forEach(y => { const r = rad[y.uke]; if (!r) return; const b = beregn(y); r.ystinger++; r.melk += b.melk; r.kg += b.kg; r.timer += b.timer; r.kost += b.k.sum; r.verdi += b.verdi; r.svinnKg += tall(y.svinnKg); });
    p.svinn.forEach(x => { const r = rad[isoUke(new Date((x.dato || iDag()) + "T12:00")).key]; if (!r) return; const v = prod(x.varenr); r.svinnKg += kgNetto(tall(x.antall), v); r.svinnKr += tall(x.antall) * (v?.pris || 0); });
    const R = key.map(k => rad[k]), sum = f => R.reduce((s, r) => s + f(r), 0);
    const iDagD = iDag(); hentSpot(iDagD); const raw = p.spotRaw[iDagD];
    const nåPris = raw ? (raw.find(x => new Date(x.time_start) <= new Date() && new Date(x.time_end) > new Date()) || {}).NOK_per_kWh : null;
    const s = sats();
    const kort = (tittel, verdi, under) => `<div class="card" style="padding:12px 14px"><div class="hint">${tittel}</div><div style="font-family:var(--f-display);font-size:26px">${verdi}</div>${under ? `<div class="hint">${under}</div>` : ""}</div>`;
    const siste4 = R.slice(-4), s4 = f => siste4.reduce((a, r) => a + f(r), 0);
    return `<p class="sub">Siste 12 uker. Tallene for produksjon gjelder ystinger registrert i appen.</p>
    <div style="display:grid;grid-template-columns:repeat(auto-fill,minmax(160px,1fr));gap:10px;margin-bottom:16px">
      ${kort("Ordre siste 4 uker", nf(s4(r => r.ordre)), `${nf(s4(r => r.kgBestilt))} kg bestilt`)}
      ${kort("Produsert siste 4 uker", nf(s4(r => r.kg)) + " kg", `${nf(s4(r => r.ystinger))} ystinger · ${nf(s4(r => r.melk))} l melk`)}
      ${kort("Kostnad per kg", s4(r => r.kg) ? nf(s4(r => r.kost) / s4(r => r.kg)) + " kr" : "–", s4(r => r.verdi) ? `DB ${nf((s4(r => r.verdi) - s4(r => r.kost)) / s4(r => r.verdi) * 100)} %` : "")}
      ${kort("Svinn siste 4 uker", nf(s4(r => r.svinnKg), 1) + " kg", s4(r => r.svinnKr) ? `${nf(Math.round(s4(r => r.svinnKr)))} kr i salgsverdi` : "")}
      ${kort("Strømpris nå (" + esc(s.omrade) + ")", nåPris != null ? nf(nåPris, 2) + " kr/kWh" : p.spot[iDagD] === 0 ? "ikke tilgjengelig" : "henter …", p.spot[iDagD] > 0 ? `snitt i dag ${nf(p.spot[iDagD], 2)} · strøm per ysting ca. ${kr(s.kWh * (p.spot[iDagD] + s.nettleieKr))}` : "")}
    </div>
    <div class="two">
      <div class="card form"><h3>Bestilt og produsert (kg)</h3>${søyler([{ navn: "Bestilt", farge: "var(--sage)", v: R.map(r => r.kgBestilt) }, { navn: "Produsert", farge: "var(--accent)", v: R.map(r => r.kg) }], uker.map(u => "u" + u.w))}</div>
      <div class="card form"><h3>Strømpris i dag (kr/kWh, uten mva)</h3>${raw ? strømGraf(raw) : `<p class="hint">${p.spot[iDagD] === 0 ? "Strømprisen kunne ikke hentes fra hvakosterstrommen.no. Kostnaden bruker reserveprisen under Satser." : "Henter …"}</p>`}</div>
    </div>
    <div class="tablewrap" style="margin-top:16px"><table class="grid"><thead><tr><th>Uke</th><th class="n">Ordre</th><th class="n">Kg bestilt</th><th class="n">Ystinger</th><th class="n">Melk (l)</th><th class="n">Kg produsert</th><th class="n">Timer</th><th class="n">Kr/kg</th><th class="n">Svinn (kg)</th></tr></thead><tbody>
    ${uker.map((u, i) => { const r = R[i]; return `<tr><td>Uke ${u.w}</td><td class="n">${nf(r.ordre)}</td><td class="n">${nf(r.kgBestilt)}</td><td class="n">${nf(r.ystinger)}</td><td class="n">${nf(r.melk)}</td><td class="n">${nf(r.kg)}</td><td class="n">${nf(r.timer, 1)}</td><td class="n">${r.kg ? nf(Math.round(r.kost / r.kg)) : "–"}</td><td class="n">${nf(r.svinnKg, 1)}</td></tr>`; }).reverse().join("")}
    </tbody><tfoot><tr><td>Sum</td><td class="n">${nf(sum(r => r.ordre))}</td><td class="n">${nf(sum(r => r.kgBestilt))}</td><td class="n">${nf(sum(r => r.ystinger))}</td><td class="n">${nf(sum(r => r.melk))}</td><td class="n">${nf(sum(r => r.kg))}</td><td class="n">${nf(sum(r => r.timer), 1)}</td><td class="n">${sum(r => r.kg) ? nf(Math.round(sum(r => r.kost) / sum(r => r.kg))) : "–"}</td><td class="n">${nf(sum(r => r.svinnKg), 1)}</td></tr></tfoot></table></div>
    ${visEnergi()}
    <p class="hint">Strømpris: <a href="https://www.hvakosterstrommen.no" target="_blank" rel="noopener">hvakosterstrommen.no</a>, spotpris for ${esc(s.omrade)} uten mva og nettleie.</p>`;
  }

  function visEnergi() {
    const s = sats(), M = ["jan","feb","mar","apr","mai","jun","jul","aug","sep","okt","nov","des"];
    const rader = ENERGI.map(e => { const egen = e.sol != null && e.solgt != null ? Math.max(0, e.sol - e.solgt) : null; const tot = e.nett != null && egen != null ? e.nett + egen : null; return { ...e, egen, tot, andel: tot ? egen / tot : null, spart: egen != null && e.spot != null ? egen * (e.spot + s.nettleieKr) + e.solgt * e.spot : null }; });
    const iÅr = rader.filter(r => r.mnd.startsWith("2026") && r.spart != null), sum = f => iÅr.reduce((a, r) => a + (f(r) || 0), 0);
    const andel = sum(r => r.egen) / sum(r => r.tot);
    const lab = r => { const [y, m] = r.mnd.split("-"); return M[+m - 1] + (m === "01" ? " " + y.slice(2) : ""); };
    const vis = rader.filter(r => r.nett != null);
    return `<div class="card form" style="margin-top:16px"><h3>Energi i ysteriet</h3>
      <div style="display:grid;grid-template-columns:repeat(auto-fill,minmax(170px,1fr));gap:10px">
        <div class="card" style="padding:12px 14px"><div class="hint">Solcellene i 2026</div><div style="font-family:var(--f-display);font-size:26px">${nf(SOL.iÅr)} kWh</div><div class="hint">${nf(SOL.kWp, 1)} kWp · ${nf(SOL.totalt / 1000, 1)} MWh siden ${SOL.fra}</div></div>
        <div class="card" style="padding:12px 14px"><div class="hint">Strøm vi lager selv</div><div style="font-family:var(--f-display);font-size:26px">ca. ${nf(andel * 100)} %</div><div class="hint">av alt ysteriet bruker, jan–sep 2026</div></div>
        <div class="card" style="padding:12px 14px"><div class="hint">Spart og solgt jan–sep</div><div style="font-family:var(--f-display);font-size:26px">ca. ${nf(Math.round(sum(r => r.spart) / 1000))} 000 kr</div><div class="hint">${nf(sum(r => r.egen))} kWh brukt selv · ${nf(sum(r => r.solgt))} kWh solgt</div></div>
        <div class="card" style="padding:12px 14px"><div class="hint">Strøm og nettleie jan–sep</div><div style="font-family:var(--f-display);font-size:26px">${nf(Math.round(sum(r => r.kr) / 1000))} 000 kr</div><div class="hint">eks. mva, fra NTE-fakturaene</div></div>
      </div>
      ${søyler([{ navn: "Kjøpt fra nettet (kWh)", farge: "var(--sage)", v: vis.map(r => r.nett) }, { navn: "Solceller brukt selv (kWh)", farge: "#E8B93A", v: vis.map(r => r.egen || 0) }, { navn: "Solgt til nettet (kWh)", farge: "var(--accent)", v: vis.map(r => r.solgt || 0) }], vis.map(lab), 200)}
      <div class="tablewrap" style="margin-top:10px"><table class="grid"><thead><tr><th>Måned</th><th class="n">Kjøpt (kWh)</th><th class="n">Sol produsert</th><th class="n">Brukt selv</th><th class="n">Solgt</th><th class="n">Andel sol</th><th class="n">Spot (kr/kWh)</th><th>Effekttrinn</th><th class="n">Strøm + nettleie</th></tr></thead><tbody>
      ${[...rader].reverse().map(r => `<tr><td>${lab(r)}</td><td class="n">${r.nett != null ? nf(r.nett) : "–"}</td><td class="n">${r.sol != null ? "ca. " + nf(r.sol) : "–"}</td><td class="n">${r.egen != null ? nf(r.egen) : "–"}</td><td class="n">${r.solgt != null ? nf(r.solgt) : "–"}</td><td class="n">${r.andel != null ? nf(r.andel * 100) + " %" : "–"}</td><td class="n">${r.spot != null ? nf(r.spot, 2) : "–"}</td><td>${r.kap ? r.kap + " kW · " + nf(r.kapKr) + " kr" : "–"}</td><td class="n">${r.kr != null ? nf(r.kr) + " kr" : "–"}</td></tr>`).join("")}
      </tbody></table></div>
      <p class="hint"><b>Effekttrinn:</b> Nettleien har et fast beløp etter den høyeste timen hver måned. I september passerte ysteriet 100 kW og havnet i trinnet 100–150 kW. Det kostet ca. 1 760 kr mer enn trinnet under. Unngå å kjøre pasteurisering, oppvask og varmtvann på full effekt samtidig.</p>
      <p class="hint">Kilder: NTE-fakturaer des. 2025–sep. 2026 (eks. mva) og SolisCloud (produksjon per måned lest av grafen 31.10.2026, derfor «ca.»). «Brukt selv» = produsert − solgt. «Spart» = brukt selv × (spot + nettleie og avgifter ${nf(s.nettleieKr, 2)} kr) + solgt × spot.</p>
    </div>`;
  }
  function strømGraf(raw) {
    const W = 640, H = 170, L = 40, B = 22, T = 8, v = raw.map(x => x.NOK_per_kWh), maks = Math.max(0.5, ...v) * 1.1, n = raw.length;
    const X = i => L + i / n * (W - L - 4), Y = p => T + (1 - p / maks) * (H - T - B), nå = new Date();
    let g = ""; for (let p = 0; p <= maks; p += maks > 2 ? 1 : 0.5) g += `<line x1="${L}" x2="${W}" y1="${Y(p)}" y2="${Y(p)}" stroke="var(--line)"/><text x="${L - 6}" y="${Y(p) + 4}" font-size="11" text-anchor="end" fill="var(--muted)">${nf(p, 1)}</text>`;
    raw.forEach((x, i) => { const t = new Date(x.time_start), aktiv = t <= nå && new Date(x.time_end) > nå; g += `<rect x="${X(i)}" y="${Y(x.NOK_per_kWh)}" width="${Math.max(1, (W - L - 4) / n - 1)}" height="${Math.max(0, H - B - Y(x.NOK_per_kWh))}" fill="${aktiv ? "var(--warn)" : "var(--accent)"}"><title>${pad(t.getHours())}.${pad(t.getMinutes())}: ${nf(x.NOK_per_kWh, 2)} kr</title></rect>`; if (t.getMinutes() === 0 && t.getHours() % 3 === 0) g += `<text x="${X(i)}" y="${H - 6}" font-size="11" fill="var(--muted)">${pad(t.getHours())}</text>`; });
    return `<svg viewBox="0 0 ${W} ${H}" style="width:100%;height:auto;display:block" role="img" aria-label="Strømpris i dag">${g}</svg><p class="hint">Gul = nå. Billigste time i dag: ${(() => { const m = raw.reduce((a, x) => x.NOK_per_kWh < a.NOK_per_kWh ? x : a); const t = new Date(m.time_start); return `kl. ${pad(t.getHours())}.${pad(t.getMinutes())} (${nf(m.NOK_per_kWh, 2)} kr)`; })()}</p>`;
  }

  function visMelk() {
    const p = P(), siste = [...p.melk].sort((a, b) => (b.tid || "").localeCompare(a.tid || "")).slice(0, 10);
    return `<form class="card form" id="prMelkForm" data-bevar="1">
      <h3>Hent melk fra fjøset</h3>
      <p class="hint">Ta bilde av displayet på melketanken, og skriv inn liter og temperatur. Melken kobles automatisk til neste ysting, og Silje ser den her i stedet for på sms.</p>
      <label class="btn prfil" style="justify-content:center;padding:14px">📷 ${p.melkBilde ? "Ta nytt bilde" : "Ta bilde av tankdisplayet"}<input type="file" data-prfoto="1" accept="image/*" capture="environment" hidden></label>
      ${p.melkBildeUrl ? `<img src="${p.melkBildeUrl}" alt="Bilde av tankdisplayet" style="max-height:220px;border-radius:8px;object-fit:contain;background:var(--sunk)">` : ""}
      <div class="row">
        <label class="f">Liter<input class="txt" id="prMelkLiter" inputmode="decimal" required placeholder="f.eks. 1800" style="font-size:22px"></label>
        <label class="f">Temperatur i tanken (°C)<input class="txt" id="prMelkTemp" inputmode="decimal" placeholder="f.eks. 3,8" style="font-size:22px"></label>
      </div>
      <label class="f">Merknad<input class="txt" id="prMelkMerk" placeholder="valgfritt"></label>
      <div class="actions"><button class="btn primary lagre" type="submit">Registrer melk</button></div>
    </form>
    <h3 style="font-family:var(--f-display);font-weight:400;margin:18px 0 8px">Siste melkehentinger</h3>
    ${siste.length ? `<div class="tablewrap"><table class="grid"><thead><tr><th>Hentet</th><th class="n">Liter</th><th class="n">°C</th><th>Hentet av</th><th>Bilde</th><th>Ysting</th></tr></thead><tbody>
    ${siste.map(m => { const y = p.ystinger.find(x => x.id === m.ystingId); const varm = tall(m.temp) > sats().melkTempGrense; return `<tr><td>${esc(datoKl(m.tid))}</td><td class="n">${nf(tall(m.liter))}</td><td class="n" style="${varm ? "color:var(--danger);font-weight:700" : ""}">${m.temp !== "" && m.temp != null ? nf(tall(m.temp), 1) : "–"}</td><td>${esc(m.hentetAv)}</td><td>${m.bilde ? `<a href="${esc(m.bilde)}" target="_blank" rel="noopener">Se bilde</a>` : "–"}</td><td>${y ? `Lot ${esc(y.lot)}` : `<span class="hint">Klar</span>`}</td></tr>`; }).join("")}
    </tbody></table></div>` : `<p class="hint">Ingen melkehentinger registrert ennå.</p>`}`;
  }

  function visYsting() {
    const p = P(); let y = p.edit;
    if (!y) {
      const ml = ledigMelk();
      const åp = åpne();
      return `<div class="empty"><h3>Ingen ysting er åpen</h3>${ml.length ? `${nf(ml.reduce((s, m) => s + tall(m.liter), 0))} liter melk er klar og blir koblet til ystingen.` : "Ingen melk er registrert ennå. Du kan likevel starte og skrive liter selv."}<div class="actions" style="justify-content:center;margin-top:12px"><button class="btn primary lagre" data-act="prNy">Start ny ysting</button></div>
        ${åp.length ? `<p class="hint" style="margin-top:14px">Eller fortsett en som ikke er avsluttet:</p><div class="actions" style="justify-content:center">${åp.map(z => `<button class="btn" data-act="prApne" data-id="${esc(z.id)}">${esc(navnY(z))}</button>`).join("")}</div>` : ""}</div>`;
    }
    const b = beregn(y), kjører = pågår(y), s = sats(), ro = y.status === "ferdig";
    const spor = sporing(y.lot);
    const ml = arr(y.melkIds).length ? p.melk.filter(m => arr(y.melkIds).includes(m.id)) : [];
    return `<form class="card form" id="prYstForm" data-id="${esc(y.id)}">
      <div class="thead"><h3>Ysting ${esc(navnY(y))} · lot ${esc(y.lot)}</h3><span class="pill ${ro ? "st-sendt" : "st-produsert"}">${ro ? "Ferdig" : "Pågår"}</span></div>
      ${!ro && y.dato < iDag() ? `<p class="notice">Denne ystingen er fra <b>${esc(dmyLang(y.dato))}</b>. Registrerer du for i dag, gå tilbake og trykk «Start ny ysting».</p>` : ""}
      ${visFeilDag(y)}
      ${!ro && ledigMelk().length ? `<p class="notice">${nf(ledigMelk().reduce((s, m) => s + tall(m.liter), 0))} liter melk er registrert, men ikke koblet til noen ysting. <button type="button" class="btn sm" data-act="prKobleMelk">Koble til denne ystingen</button></p>` : ""}
      <h4>1. Melk og råvarer</h4>
      <div class="row">
        <label class="f">Dato<input type="date" class="txt" id="py-dato" data-pf="dato" value="${esc(y.dato)}"></label>
        <label class="f">Melk (liter)<input class="txt" id="py-melkLiter" data-pf="melkLiter" inputmode="decimal" value="${esc(y.melkLiter || "")}"></label>
        <label class="f">Fløte (liter)<input class="txt" id="py-flote" data-pf="flote" inputmode="decimal" value="${esc(y.flote || "")}"></label>
        <label class="f">Vannmåler start (m³)<input class="txt" id="py-vannStart" data-pf="vannStart" inputmode="decimal" value="${esc(y.vannStart || "")}"></label>
        <label class="f">Vannmåler slutt (m³)<input class="txt" id="py-vannSlutt" data-pf="vannSlutt" inputmode="decimal" value="${esc(y.vannSlutt || "")}"></label>
      </div>
      ${ml.length ? `<p class="hint">Fra melkehenting: ${ml.map(m => `${nf(tall(m.liter))} l ${esc(datoKl(m.tid))}${m.temp !== "" && m.temp != null ? `, ${nf(tall(m.temp), 1)} °C` : ""}${m.bilde ? ` (<a href="${esc(m.bilde)}" target="_blank" rel="noopener">bilde</a>)` : ""}`).join(" · ")}</p>` : ""}
      <h4>Ostemasse</h4>
      <div class="row">
        <label class="f">Ostemasse fra denne melka (kg)<input class="txt" id="py-ostemasseKg" data-pf="ostemasseKg" inputmode="decimal" value="${esc(y.ostemasseKg || "")}"></label>
        <label class="f">Spart til senere ysting (kg)<input class="txt" id="py-overfortUtKg" data-pf="overfortUtKg" inputmode="decimal" value="${esc(y.overfortUtKg || "")}"></label>
      </div>
      ${tall(y.overfortUtKg) ? `<p class="hint">${nf(tall(y.overfortUtKg), 1)} kg ostemasse og ${kr(overførtUtKr(y))} av kostnaden flyttes til ystingen som bruker den. ${restUt(y) > 0.05 ? `${nf(restUt(y), 1)} kg er ikke brukt ennå.` : "Alt er brukt."}</p>` : ""}
      ${visInn(y)}
      ${visLogg(y)}
      <h4>3. Tid</h4>
      <div class="actions">
        ${kjører ? `<button type="button" class="btn danger lagre" data-act="prStopp">■ Stopp tid</button>` : `<button type="button" class="btn primary lagre" data-act="prStart">▶ Start tid</button>`}
        <span style="font-size:26px;font-family:var(--f-display)" id="prKlokke">${tidTekst(b.min)}</span>
        <label class="f" style="width:110px">Personer<input class="txt" id="py-personer" data-pf="personer" inputmode="numeric" value="${esc(y.personer || 1)}"></label>
        <label class="f" style="width:150px">Rett total (min)<input class="txt" id="py-minutter" data-pf="minutter" inputmode="numeric" placeholder="${b.min}" value="${y.minutterManuelt ? esc(y.minutter) : ""}"></label>
      </div>
      <p class="hint">Tiden teller fra «Start» til «Stopp», og dere kan starte og stoppe flere ganger. Arbeidstimer = tid × antall personer = <b>${nf(b.timer, 1)} timer</b>. Har dere glemt å starte, skriv riktig total i «Rett total».</p>
      <h4>4. Andre kontrollpunkter</h4>
      ${arr(y.kontroller).length ? `<div class="tablewrap"><table class="grid"><tbody>${y.kontroller.map((k, i) => `<tr><td>${esc((k.tid || "").slice(0, 10) === y.dato ? klokke(k.tid) : datoKl(k.tid))}</td><td>${esc(k.type)}</td><td class="n"><b>${esc(k.verdi)}</b></td><td class="hint">${esc(k.merknad || "")}</td><td><button type="button" class="btn sm ghost danger" data-act="prSlettKontroll" data-i="${i}" aria-label="Slett">✕</button></td></tr>`).join("")}</tbody></table></div>` : `<p class="hint">Ingen kontrollpunkter registrert.</p>`}
      <div class="row">
        <label class="f">Hva<select class="txt" id="pk-type">${KONTROLL.map(k => `<option ${k === (p.kType || KONTROLL[0]) ? "selected" : ""}>${k}</option>`).join("")}</select></label>
        <label class="f">Verdi<input class="txt" id="pk-verdi" inputmode="decimal" value="${esc(p.kVerdi || "")}"></label>
        <label class="f">Merknad<input class="txt" id="pk-merk" value="${esc(p.kMerk || "")}"></label>
      </div>
      <div class="actions"><button type="button" class="btn" data-act="prKontroll">Legg til kontrollpunkt</button></div>
      <h4>5. Produsert</h4>
      <div class="lines">${prodSort().filter(x => x.gruppe !== "Melk").map(x => { const n = tall((y.produsert || {})[x.varenr]); return `<div class="line bl-${esc(x.gruppe)} ${n ? "has" : ""}">
        <div class="pn">${esc(x.kort)}<small>#${esc(x.varenr)}</small>${n ? `<small class="ant">${nf(kgNetto(n, x), 1)} kg</small>` : ""}</div>
        <div class="qty"><button type="button" class="stp" data-act="prMinus" data-v="${esc(x.varenr)}" ${n ? "" : "disabled"} aria-label="Én mindre ${esc(x.kort)}">−</button><input id="pp-${esc(x.varenr)}" data-pp="${esc(x.varenr)}" inputmode="decimal" value="${n || ""}" placeholder="–" aria-label="Antall ${esc(x.kort)}"><button type="button" class="stp" data-act="prPluss" data-v="${esc(x.varenr)}" aria-label="Én mer ${esc(x.kort)}">+</button><span class="u">${esc(enhetNavn(x, 2))}</span></div></div>`; }).join("")}</div>
      <div class="row">
        <label class="f">Antall esker / kartonger<input class="txt" id="py-esker" data-pf="esker" inputmode="numeric" value="${esc(y.esker || "")}"></label>
        <label class="f">Svinn under produksjon (kg)<input class="txt" id="py-svinnKg" data-pf="svinnKg" inputmode="decimal" value="${esc(y.svinnKg || "")}"></label>
      </div>
      <label class="f">Merknad<textarea class="txt" id="py-merknad" data-pf="merknad" rows="2">${esc(y.merknad || "")}</textarea></label>
      <h4>Kostnad for denne ystingen</h4>
      <div class="tablewrap"><table class="grid"><tbody>
        <tr><td>Melk ${nf(b.melk)} l × ${nf(s.melkKrL, 2)} kr</td><td class="n">${kr(b.k.melk)}</td></tr>
        <tr><td>Fløte ${nf(b.flote, 1)} l × ${nf(s.floteKrL, 2)} kr</td><td class="n">${kr(b.k.flote)}</td></tr>
        <tr><td>Kultur, løpe, vaskemiddel (fast per ysting)</td><td class="n">${kr(b.k.batch)}</td></tr>
        ${b.melk ? `<tr><td>Strøm ${nf(s.kWh)} kWh × (${nf(b.pris, 2)} spot${tall(y.strompris) > 0 ? "" : P().spot[y.dato] > 0 ? "" : " (anslag)"} + ${nf(s.nettleieKr, 2)} nettleie)</td><td class="n">${kr(b.k.strom)}</td></tr>` : ""}
        ${b.k.vann ? `<tr><td>Vann ${nf(vannM3(y), 2)} m³ × ${nf(s.vannKrM3, 2)} kr</td><td class="n">${kr(b.k.vann)}</td></tr>` : ""}
        ${b.k.ut ? `<tr><td>Ostemasse spart til senere (${nf(tall(y.overfortUtKg), 1)} kg)</td><td class="n">${kr(b.k.ut)}</td></tr>` : ""}
        ${b.k.inn ? `<tr><td>Ostemasse fra tidligere ysting</td><td class="n">${kr(b.k.inn)}</td></tr>` : ""}
        <tr><td>Emballasje (${nf(b.enh)} enheter)</td><td class="n">${kr(b.k.emb)}</td></tr>
        <tr><td>Arbeid ${nf(b.timer, 1)} t × ${nf(s.timeKr, 0)} kr</td><td class="n">${kr(b.k.arbeid)}</td></tr>
      </tbody><tfoot>
        <tr><td>Sum kostnad · ${nf(b.kg, 1)} kg ost${b.melk ? ` · ${nf(b.utbytte, 1)} kg ${b.utbytteOm ? "ostemasse" : "ost"}/100 l` : ""}</td><td class="n">${kr(b.k.sum)}${b.kg ? ` <span class="hint">(${nf(b.perKg)} kr/kg)</span>` : ""}</td></tr>
        <tr><td>Salgsverdi til listepris</td><td class="n">${kr(b.verdi)}</td></tr>
        <tr><td>Dekningsbidrag</td><td class="n" style="color:${b.db < 0 ? "var(--danger)" : "var(--ok)"}">${kr(b.db)}${b.verdi ? ` (${nf(b.dbPst * 100)} %)` : ""}</td></tr>
        ${b.modell ? `<tr><td>Styringsmodellen forventet</td><td class="n">${kr(b.modell)}</td></tr>` : ""}
      </tfoot></table></div>
      ${spor.length ? `<h4>Sporbarhet: ordre med lot ${esc(y.lot)}</h4><p class="hint">${spor.map(x => `${esc(x.kunde)}: ${nf(x.antall)} × ${esc(prod(x.varenr)?.kort || x.varenr)}`).join(" · ")}</p>` : `<p class="hint">Lot ${esc(y.lot)} er batchnummeret ordrene får når produksjonsdatoen er ${esc(dmy(y.dato))}.</p>`}
      <div class="actions lagrerad">
        ${ro ? `<button type="button" class="btn" data-act="prGjenapne">Åpne igjen</button>` : `<button type="button" class="btn primary lagre" data-act="prFerdig">Ferdig med ystingen</button>`}
        <button type="button" class="btn" data-act="prLukk">Til oversikten</button>
        <span class="sp"></span>
        <button type="button" class="btn danger ghost" data-act="prSlett">${p.bekreftSlett === y.id ? "Trykk igjen for å slette" : "Slett ysting"}</button>
      </div>
      <p class="hint">Alt lagres av seg selv i SharePoint mens dere skriver.</p>
    </form>`;
  }
  const tidTekst = m => `${Math.floor(m / 60)} t ${pad(m % 60)} min`;

  /* ---------- temperatur ---------- */
  function tempPunkter(dager) {
    const fra = Date.now() - dager * 864e5, ut = [];
    P().temp.forEach(t => arr(t.data).forEach(([m, v]) => { const ms = m * 60000; if (!dager || ms >= fra) ut.push([ms, v / 10]); }));
    if (!ut.length) return ut;
    if (dager) return ut.sort((a, b) => a[0] - b[0]);
    return ut.sort((a, b) => a[0] - b[0]);
  }
  function tempGraf(pkt, dager, høyde = 220) {
    if (!pkt.length) return `<p class="hint">Ingen målinger ${dager ? `de siste ${dager} dagene` : ""}.</p>`;
    const g = sats().tempGrense, W = 640, H = høyde, L = 34, R = 8, T = 8, B = 22;
    const x0 = pkt[0][0], x1 = pkt[pkt.length - 1][0] || x0 + 1;
    let lo = pkt.reduce((m, p) => Math.min(m, p[1]), 0), hi = pkt.reduce((m, p) => Math.max(m, p[1]), g + 2); lo = Math.floor(lo); hi = Math.ceil(hi);
    const X = t => L + (t - x0) / Math.max(1, x1 - x0) * (W - L - R), Y = v => T + (hi - v) / Math.max(1, hi - lo) * (H - T - B);
    const step = Math.max(1, Math.ceil(pkt.length / 1500)); let d = "";
    for (let i = 0; i < pkt.length; i += step) d += (i ? "L" : "M") + X(pkt[i][0]).toFixed(1) + " " + Y(pkt[i][1]).toFixed(1);
    const over = pkt.filter(p => p[1] > g);
    const ticks = []; for (let v = lo; v <= hi; v += Math.max(1, Math.round((hi - lo) / 5))) ticks.push(v);
    const dagT = []; const dt = new Date(x0); dt.setHours(0, 0, 0, 0); for (let t = dt.getTime() + 864e5; t < x1; t += 864e5 * Math.max(1, Math.round((x1 - x0) / 864e5 / 8))) dagT.push(t);
    return `<svg viewBox="0 0 ${W} ${H}" style="width:100%;height:auto;display:block" role="img" aria-label="Temperatur i kjølerommet">
      ${ticks.map(v => `<line x1="${L}" x2="${W - R}" y1="${Y(v)}" y2="${Y(v)}" stroke="var(--line)"/><text x="${L - 6}" y="${Y(v) + 4}" font-size="11" text-anchor="end" fill="var(--muted)">${v}°</text>`).join("")}
      ${dagT.map(t => { const dd = new Date(t); return `<text x="${X(t)}" y="${H - 6}" font-size="11" text-anchor="middle" fill="var(--muted)">${pad(dd.getDate())}.${pad(dd.getMonth() + 1)}</text>`; }).join("")}
      <line x1="${L}" x2="${W - R}" y1="${Y(g)}" y2="${Y(g)}" stroke="var(--danger)" stroke-dasharray="5 4"/><text x="${W - R}" y="${Y(g) - 4}" font-size="11" text-anchor="end" fill="var(--danger)">${nf(g)} °C</text>
      <path d="${d}" fill="none" stroke="var(--accent)" stroke-width="1.6"/>
      ${over.length < 400 ? over.map(p => `<circle cx="${X(p[0])}" cy="${Y(p[1])}" r="2" fill="var(--danger)"/>`).join("") : ""}
    </svg>`;
  }
  function visTemp() {
    const p = P(), dager = p.tempValg, pkt = tempPunkter(dager), g = sats().tempGrense;
    const over = pkt.filter(x => x[1] > g);
    const f = p.tempForslag;
    return `<div class="card form">
      <h3>Kjølerommet</h3>
      <p class="hint">Last opp fila fra USB-loggeren (CSV, TXT eller Excel). Appen finner dato, klokkeslett og temperatur selv og viser hva den fant før noe lagres. Målinger som allerede er lastet opp, lagres ikke på nytt.</p>
      <label class="btn prfil" style="justify-content:center;padding:12px">Last opp loggfil<input type="file" id="prTempFil" accept=".csv,.txt,.xls,.xlsx,text/csv,text/plain" hidden></label>
      ${f ? `<div class="notice"><b>${esc(f.filnavn)}:</b> ${nf(f.data.length)} målinger fra ${esc(datoKl(new Date(f.data[0][0] * 60000).toISOString()))} til ${esc(datoKl(new Date(f.data[f.data.length - 1][0] * 60000).toISOString()))}, lavest ${nf(f.min, 1)} °C, høyest ${nf(f.max, 1)} °C, ${nf(f.over)} over ${nf(g)} °C.${f.nye < f.data.length ? ` ${nf(f.data.length - f.nye)} av dem er lastet opp før.` : ""}<br>Kolonner brukt: ${esc(f.kol)}.
        <div class="actions" style="margin-top:8px"><button class="btn primary" data-act="prTempLagre" ${f.nye ? "" : "disabled"}>Lagre ${nf(f.nye)} nye målinger</button><button class="btn ghost" data-act="prTempAvbryt">Avbryt</button></div></div>` : ""}
      <div class="actions">${[1, 7, 30, 0].map(d => `<button class="btn sm ${dager === d ? "primary" : ""}" data-act="prTempDager" data-v="${d}">${d ? d === 1 ? "Siste døgn" : `${d} dager` : "Alt"}</button>`).join("")}</div>
      ${tempGraf(pkt, dager)}
      ${pkt.length ? `<p class="hint">${nf(pkt.length)} målinger · lavest ${nf(pkt.reduce((m, x) => Math.min(m, x[1]), 99), 1)} °C · høyest ${nf(pkt.reduce((m, x) => Math.max(m, x[1]), -99), 1)} °C · <b style="color:${over.length ? "var(--danger)" : "var(--ok)"}">${nf(over.length)} over ${nf(g)} °C</b></p>` : ""}
      ${p.temp.length ? `<h4>Opplastede filer</h4><div class="tablewrap"><table class="grid"><thead><tr><th>Periode</th><th>Fil</th><th class="n">Målinger</th><th class="n">Høyest</th><th class="n">Over grense</th><th>Lastet opp</th></tr></thead><tbody>${[...p.temp].reverse().map(t => `<tr><td>${esc(dmy((t.fra || "").slice(0, 10)))}–${esc(dmy((t.til || "").slice(0, 10)))}</td><td>${esc(t.filnavn)}</td><td class="n">${nf(tall(t.antall))}</td><td class="n">${nf(tall(t.max), 1)}</td><td class="n">${nf(tall(t.over))}</td><td class="hint">${esc(datoKl(t.lastetOpp))} ${esc(t.av || "")}</td></tr>`).join("")}</tbody></table></div>` : ""}
    </div>`;
  }
  /* Tolker loggfila: finner kolonne med dato/tid og kolonne med temperatur. Tåler ; , og tab som skilletegn og komma som desimaltegn. */
  function tolkLogg(rader) {
    rader = rader.map(r => r.map(c => String(c ?? "").trim())).filter(r => r.some(Boolean));
    const datoRe = /(\d{4})[-./](\d{1,2})[-./](\d{1,2})|(\d{1,2})[-./](\d{1,2})[-./](\d{2,4})/;
    const tidRe = /(\d{1,2}):(\d{2})(?::(\d{2}))?\s*(AM|PM)?/i;
    const tallRe = /^-?\d+(?:[.,]\d+)?$/;
    const prøve = rader.slice(0, 200);
    const nK = Math.max(...prøve.map(r => r.length));
    let dK = -1, tK = -1, vK = -1, best = 0;
    for (let k = 0; k < nK; k++) { const n = prøve.filter(r => datoRe.test(r[k] || "")).length; if (n > best) { best = n; dK = k; } }
    if (dK < 0) throw new Error("Fant ingen kolonne med dato i fila.");
    if (!prøve.some(r => tidRe.test(r[dK] || ""))) { best = 0; for (let k = 0; k < nK; k++) { if (k === dK) continue; const n = prøve.filter(r => tidRe.test(r[k] || "") && !datoRe.test(r[k] || "")).length; if (n > best) { best = n; tK = k; } } }
    const hode = rader.find(r => !datoRe.test(r[dK] || "")) || [];
    let kand = []; for (let k = 0; k < nK; k++) { if (k === dK || k === tK) continue; const v = prøve.map(r => r[k] || "").filter(x => tallRe.test(x)).map(x => Number(x.replace(",", "."))); if (v.length > prøve.length * 0.5) kand.push({ k, v, navn: hode[k] || "" }); }
    const temp = kand.find(c => /temp|°c|celsius|\bc\b/i.test(c.navn)) || kand.filter(c => !/^(nr|no|#|index|id|rh|fukt|hum)/i.test(c.navn)).find(c => c.v.every(x => x > -40 && x < 60) && new Set(c.v).size > 1) || kand[0];
    if (!temp) throw new Error("Fant ingen kolonne med temperatur i fila.");
    vK = temp.k;
    const ut = [];
    rader.forEach(r => {
      const ds = r[dK] || "", m = ds.match(datoRe); if (!m) return;
      let y, mo, d; if (m[1]) { y = +m[1]; mo = +m[2]; d = +m[3]; } else { d = +m[4]; mo = +m[5]; y = +m[6]; if (y < 100) y += 2000; }
      const ts = (tK >= 0 ? r[tK] : ds).match(tidRe); let h = 0, mi = 0;
      if (ts) { h = +ts[1]; mi = +ts[2]; if (ts[4]) { const pm = /pm/i.test(ts[4]); if (pm && h < 12) h += 12; if (!pm && h === 12) h = 0; } }
      const v = r[vK]; if (!tallRe.test(v || "")) return;
      const t = new Date(y, mo - 1, d, h, mi).getTime(); if (!isFinite(t)) return;
      ut.push([Math.round(t / 60000), Math.round(Number(v.replace(",", ".")) * 10)]);
    });
    if (!ut.length) throw new Error("Fant ingen målinger i fila.");
    ut.sort((a, b) => a[0] - b[0]);
    return { data: ut, kol: `dato «${hode[dK] || dK + 1}»${tK >= 0 ? `, tid «${hode[tK] || tK + 1}»` : ""}, temperatur «${hode[vK] || vK + 1}»` };
  }
  async function lesLoggfil(fil) {
    let rader;
    if (/\.xlsx?$/i.test(fil.name)) {
      await lastSkript("vendor/xlsx.mini.min.js");
      const wb = XLSX.read(await fil.arrayBuffer(), { type: "array", cellDates: false, raw: false });
      rader = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, raw: false });
    } else {
      let txt = await fil.text(); if (/�/.test(txt)) txt = new TextDecoder("windows-1252").decode(await fil.arrayBuffer());
      const linjer = txt.split(/\r?\n/).filter(l => l.trim());
      const sep = [";", "\t", ","].map(c => [c, linjer.slice(0, 50).reduce((s, l) => s + l.split(c).length - 1, 0)]).sort((a, b) => b[1] - a[1])[0][0];
      rader = linjer.map(l => l.split(sep).map(c => c.replace(/^"|"$/g, "")));
    }
    const t = tolkLogg(rader);
    const har = new Set(); P().temp.forEach(x => arr(x.data).forEach(([m]) => har.add(m)));
    const nye = t.data.filter(([m]) => !har.has(m));
    const g = sats().tempGrense * 10;
    P().tempForslag = { filnavn: fil.name, data: t.data, nyeData: nye, nye: nye.length, kol: t.kol, min: t.data.reduce((m, x) => Math.min(m, x[1]), 9999) / 10, max: t.data.reduce((m, x) => Math.max(m, x[1]), -9999) / 10, over: t.data.filter(x => x[1] > g).length };
    render();
  }
  async function lagreTemp() {
    const p = P(), f = p.tempForslag; if (!f || !f.nye) return;
    const g = sats().tempGrense * 10, BIT = 2000; let n = 0;
    for (let i = 0; i < f.nyeData.length; i += BIT) {
      const del = f.nyeData.slice(i, i + BIT);
      const post = { id: nyId("temp"), fra: new Date(del[0][0] * 60000).toISOString(), til: new Date(del[del.length - 1][0] * 60000).toISOString(), filnavn: f.filnavn, antall: del.length, min: del.reduce((m, x) => Math.min(m, x[1]), 9999) / 10, max: del.reduce((m, x) => Math.max(m, x[1]), -9999) / 10, over: del.filter(x => x[1] > g).length, data: del, lastetOpp: nå(), av: brukerNavn() };
      const { id, ...d } = post;
      try { await skriv("temperatur/" + id, d); p.temp.push(post); n += del.length; } catch (e) { feilLagring(e); break; }
    }
    p.temp.sort((a, b) => (a.fra || "").localeCompare(b.fra || "")); p.tempForslag = null; render();
    if (n) toast(`✓ ${nf(n)} målinger fra kjølerommet er lagret.`);
  }

  function visSatser() {
    const s = sats(), fi = (k, l, v) => `<label class="f">${l}<input class="txt" id="ps-${k}" value="${esc(String(v).replace(".", ","))}" inputmode="decimal"></label>`;
    return `<form class="card form" id="prSatsForm">
      <h3>Satser for kostnadsberegning</h3>
      <p class="hint">Standardverdiene er hentet fra Toddum Styringsmodell v9. Endringer gjelder for alle ystinger, også de som er registrert før.</p>
      <div class="row">${fi("melkKrL", "Melk (kr/l)", s.melkKrL)}${fi("floteKrL", "Fløte (kr/l)", s.floteKrL)}${fi("timeKr", "Arbeid (kr/time)", s.timeKr)}${fi("batchKr", "Kultur, løpe, vaskemiddel per ysting (kr)", s.batchKr)}${fi("kWh", "Strøm per ysting (kWh)", s.kWh)}${fi("nettleieKr", "Nettleie og påslag (kr/kWh)", s.nettleieKr)}${fi("stromReserve", "Strømpris hvis spot mangler (kr/kWh)", s.stromReserve)}${fi("vannKrM3", "Vann (kr/m³)", s.vannKrM3)}${fi("tempGrense", "Grense kjølerom (°C)", s.tempGrense)}${fi("melkTempGrense", "Grense melketank (°C)", s.melkTempGrense)}</div>
      <h4>Emballasje per enhet (kr)</h4>
      <div class="row">${prodAlle().filter(x => x.gruppe !== "Melk").map(x => fi("emb-" + x.varenr, `${x.kort} (#${x.varenr})`, s.emb[x.varenr] ?? 0)).join("")}</div>
      <div class="actions"><button class="btn primary lagre" type="submit">Lagre satser</button><button class="btn ghost" type="button" data-act="prSatsStd">Tilbake til standard</button></div>
    </form>`;
  }

  /* ---------- handlinger ---------- */
  function flyttKontroller(d, tilId) {
    const p = P(), fra = p.edit; if (!fra) return;
    const flytt = arr(fra.kontroller).filter(k => (k.tid || "").slice(0, 10) === d); if (!flytt.length) return;
    let til = p.ystinger.find(z => z.id === tilId);
    if (!til) { const nr = p.ystinger.filter(z => z.dato === d).reduce((m, z) => Math.max(m, batchNr(z)), 0) + 1; til = { id: nyId("ysting"), dato: d, uke: isoUke(new Date(d + "T12:00")).key, batchNr: nr, lot: lotFor(d, nr), status: "pågår", melkLiter: 0, melkIds: [], flote: 0, tid: [], personer: 2, minutter: 0, kontroller: [], produsert: {}, esker: 0, svinnKg: 0, merknad: "", registrertAv: brukerNavn(), ystetAv: brukerNavn(), osteType: fra.osteType || "", logg: {} }; }
    const nyTil = { ...til, kontroller: [...arr(til.kontroller), ...flytt].sort((a, b) => (a.tid || "").localeCompare(b.tid || "")), endret: nå() };
    const { id, ...dTil } = nyTil;
    const i = p.ystinger.findIndex(x => x.id === id); if (i >= 0) p.ystinger[i] = nyTil; else p.ystinger.push(nyTil);
    /* Først lagres de flyttede punktene på mål-ystingen; først når det har gått bra, fjernes de fra denne. */
    skriv("ysting/" + id, dTil).then(() => {
      const n = { ...p.ystinger.find(x => x.id === fra.id) || fra }; n.kontroller = arr(n.kontroller).filter(k => (k.tid || "").slice(0, 10) !== d);
      return lagreYsting(n, true);
    }).then(() => { toast(`✓ ${flytt.length} kontrollpunkt flyttet til ${navnY(nyTil)}.`); render(); }).catch(feilLagring);
  }
  function endreY(f) { const y = P().edit; if (!y) return; const n = { ...y }; f(n); lagreYsting(n); planRender(); }
  function klikk(el) {
    const a = el.dataset.act; if (!a || !a.startsWith("pr")) return false;
    const p = P();
    if (a === "prVis") { p.pv = el.dataset.v; if (p.pv === "ysting" && !p.edit) { const akt = åpne().filter(y => y.dato === iDag()); if (akt.length === 1) p.edit = akt[0]; } render(); }
    else if (a === "prNaa") { const r = el.dataset.r, d = new Date(), kl = `${pad(d.getHours())}.${pad(d.getMinutes())}`; endreY(y => { const l = { ...(y.logg && typeof y.logg === "object" ? y.logg : {}) }; l[r] = { ...(l[r] || {}), kl }; y.logg = l; }); }
    else if (a === "prKobleMelk") { const ml = ledigMelk(); if (!ml.length || !p.edit) return true; const id = p.edit.id; endreY(y => { y.melkIds = [...arr(y.melkIds), ...ml.map(m => m.id)]; y.melkLiter = tall(y.melkLiter) + ml.reduce((s, m) => s + tall(m.liter), 0); }); ml.forEach(m => lagreMelk({ ...m, ystingId: id }).catch(() => {})); }
    else if (a === "prFlyttK") flyttKontroller(el.dataset.d, el.dataset.til);
    else if (a === "prNy") nyYsting();
    else if (a === "prApne") { p.edit = p.ystinger.find(y => y.id === el.dataset.id) || null; p.pv = "ysting"; render(); window.scrollTo(0, 0); }
    else if (a === "prLukk") { p.edit = null; p.pv = "oversikt"; render(); }
    else if (a === "prStart") endreY(y => { y.tid = [...arr(y.tid), { fra: nå() }]; y.minutterManuelt = false; });
    else if (a === "prStopp") endreY(y => { y.tid = arr(y.tid).map(t => t.til ? t : { ...t, til: nå() }); y.minutter = minutter(y); });
    else if (a === "prKontroll") { const v = $("#pk-verdi")?.value.trim(); if (!v) { toast("Skriv en verdi."); return true; } const k = { tid: nå(), type: $("#pk-type").value, verdi: v, merknad: $("#pk-merk").value.trim(), av: brukerNavn() }; p.kType = k.type; p.kVerdi = ""; p.kMerk = ""; endreY(y => { y.kontroller = [...arr(y.kontroller), k]; }); }
    else if (a === "prSlettKontroll") endreY(y => { y.kontroller = arr(y.kontroller).filter((_, i) => i !== +el.dataset.i); });
    else if (a === "prPluss" || a === "prMinus") { const v = el.dataset.v; endreY(y => { const pr = { ...(y.produsert || {}) }; const n = Math.max(0, tall(pr[v]) + (a === "prPluss" ? 1 : -1)); if (n) pr[v] = n; else delete pr[v]; y.produsert = pr; }); }
    else if (a === "prFerdig") { const y = P().edit; if (!y) return true; if (!Object.keys(y.produsert || {}).length) { toast("Registrer hva som er produsert før ystingen avsluttes."); return true; } const n = { ...y, status: "ferdig", tid: arr(y.tid).map(t => t.til ? t : { ...t, til: nå() }) }; if (tall(n.melkLiter) && !(tall(n.strompris) > 0) && P().spot[n.dato] > 0) n.strompris = Math.round(P().spot[n.dato] * 10000) / 10000; n.minutter = minutter(n); lagreYsting(n, true).then(() => toast(`✓ Ystingen ${dmy(n.dato)} er lagret. Lot ${n.lot}.`)); p.edit = null; p.pv = "oversikt"; render(); }
    else if (a === "prGjenapne") endreY(y => { y.status = "pågår"; });
    else if (a === "prSlett") { const y = p.edit; if (!y) return true; if (p.bekreftSlett === y.id) { p.bekreftSlett = null; p.ystinger = p.ystinger.filter(x => x.id !== y.id); p.melk.filter(m => m.ystingId === y.id).forEach(m => lagreMelk({ ...m, ystingId: "" }).catch(() => {})); p.edit = null; p.pv = "oversikt"; clearTimeout(p.lagrerT); p.lagrerT = null; skriv("ysting/" + y.id, null).then(() => toast("Ystingen er slettet.")).catch(feilLagring); render(); } else { p.bekreftSlett = y.id; render(); setTimeout(() => { if (p.bekreftSlett === y.id) { p.bekreftSlett = null; render(); } }, 4000); } }
    else if (a === "prTempDager") { p.tempValg = +el.dataset.v; render(); }
    else if (a === "prTempLagre") { el.disabled = true; lagreTemp(); }
    else if (a === "prTempAvbryt") { p.tempForslag = null; render(); }
    else if (a === "prSatsStd") { p.innst = null; skriv("innst/satser", { data: {} }).then(() => toast("Satsene er satt tilbake til standard.")).catch(feilLagring); render(); }
    else if (a === "prLeggInn") { const y = p.edit; if (!y) return true; const fra = $("#pi-fra")?.value, kg = tall($("#pi-kg")?.value); const src = p.ystinger.find(z => z.id === fra); if (!src || !(kg > 0)) { toast("Velg ysting og skriv antall kg."); return true; } if (kg > restUt(src) + 0.05) { toast(`Det er bare ${nf(Math.max(0, restUt(src)), 1)} kg igjen fra ${dmy(src.dato)}.`); return true; } endreY(n => { const inn = arr(n.overfortInn).filter(x => x.fra !== fra); const før = arr(n.overfortInn).find(x => x.fra === fra); inn.push({ fra, kg: kg + tall(før?.kg) }); n.overfortInn = inn; }); }
    else if (a === "prFjernInn") endreY(y => { y.overfortInn = arr(y.overfortInn).filter((_, i) => i !== +el.dataset.i); });
    else if (a === "prSlettSvinn") { const id = el.dataset.id; if (p.bekreftSvinn === id) { p.bekreftSvinn = null; p.svinn = p.svinn.filter(x => x.id !== id); skriv("svinn/" + id, null).then(() => toast("Svinnet er slettet.")).catch(feilLagring); render(); } else { p.bekreftSvinn = id; toast("Trykk ✕ igjen for å slette."); setTimeout(() => { if (p.bekreftSvinn === id) p.bekreftSvinn = null; }, 4000); } }
    else return false;
    return true;
  }
  function endring(e) {
    const t = e.target, p = P();
    if (t.dataset.prfoto && t.files && t.files[0]) { const f = t.files[0]; krympBilde(f).then(b => { p.melkBilde = b; if (p.melkBildeUrl) URL.revokeObjectURL(p.melkBildeUrl); p.melkBildeUrl = URL.createObjectURL(b); render(); }).catch(() => { p.melkBilde = f; p.melkBildeUrl = URL.createObjectURL(f); render(); }); return true; }
    if (t.id === "prTempFil" && t.files && t.files[0]) { const f = t.files[0]; t.value = ""; lesLoggfil(f).catch(err => toast("Klarte ikke å lese loggfila: " + (err.message || err))); return true; }
    if (t.id === "pk-type") { p.kType = t.value; return true; }
    if (t.id === "pk-verdi") { p.kVerdi = t.value; return true; }
    if (t.id === "pk-merk") { p.kMerk = t.value; return true; }
    if (t.dataset.pl) { const [r, f] = t.dataset.pl.split("|"), v = t.value.trim(); endreY(y => { const l = { ...(y.logg && typeof y.logg === "object" ? y.logg : {}) }; l[r] = { ...(l[r] || {}), [f]: v }; y.logg = l; }); return true; }
    if (t.dataset.pf) {
      const f = t.dataset.pf; let v = t.value.trim();
      endreY(y => {
        if (["melkLiter", "flote", "personer", "esker", "svinnKg", "ostemasseKg", "overfortUtKg", "vannStart", "vannSlutt"].includes(f)) y[f] = tall(v);
        else if (f === "minutter") { if (v === "") { y.minutterManuelt = false; y.minutter = 0; } else { y.minutter = Math.round(tall(v)); y.minutterManuelt = true; } }
        else if (f === "dato") { y.dato = v; y.uke = isoUke(new Date(v + "T12:00")).key; y.lot = lotFor(v, batchNr(y)); }
        else y[f] = v;
      });
      return true;
    }
    if (t.dataset.pp) { const v = t.dataset.pp; endreY(y => { const pr = { ...(y.produsert || {}) }; const n = tall(t.value); if (n > 0) pr[v] = n; else delete pr[v]; y.produsert = pr; }); return true; }
    return false;
  }
  function submit(form) {
    const p = P();
    if (form.id === "prMelkForm") {
      const liter = tall($("#prMelkLiter").value), tempS = $("#prMelkTemp").value.trim();
      if (!(liter > 0)) { toast("Skriv antall liter."); return true; }
      const knapp = form.querySelector("[type=submit]"); knapp.disabled = true; knapp.textContent = "Lagrer …";
      const tid = nå(), id = nyId("melk");
      const m = { id, tid, dato: tid.slice(0, 10), liter, temp: tempS === "" ? "" : tall(tempS), hentetAv: brukerNavn(), bilde: "", ystingId: "", merknad: $("#prMelkMerk").value.trim() };
      /* Kobles bare automatisk når det er nøyaktig én åpen ysting i dag og den ikke har melk ennå. Ellers står melken klar til neste ysting. */
      const iDagÅpne = åpne().filter(y => y.dato === iDag()); const aktiv = iDagÅpne.length === 1 && !tall(iDagÅpne[0].melkLiter) ? iDagÅpne[0] : null; if (aktiv) m.ystingId = aktiv.id;
      const opp = p.melkBilde ? Store.lastOppFil(`Ysteri/Melkehenting/${tid.slice(0, 10)} ${klokke(tid).replace(".", "")} ${liter} l.jpg`, p.melkBilde).then(r => r.url).catch(e => { toast("Bildet ble ikke lastet opp" + (e.status === 403 ? " (mangler tilgang)" : "") + ". Melken registreres uten bilde."); return ""; }) : Promise.resolve("");
      opp.then(url => { m.bilde = url; return lagreMelk(m); }).then(() => {
        if (aktiv) { const n = { ...aktiv, melkIds: [...arr(aktiv.melkIds), id], melkLiter: tall(aktiv.melkLiter) + liter }; lagreYsting(n, true); if (p.edit && p.edit.id === aktiv.id) p.edit = n; }
        p.melkBilde = null; if (p.melkBildeUrl) URL.revokeObjectURL(p.melkBildeUrl); p.melkBildeUrl = ""; form.reset(); S.utkast.delete("prMelkForm|");
        const varm = m.temp !== "" && m.temp > sats().melkTempGrense;
        toast(`✓ ${nf(liter)} liter er registrert${aktiv ? ` på ystingen ${dmy(aktiv.dato)}` : ""}.${varm ? ` OBS: ${nf(m.temp, 1)} °C er over ${nf(sats().melkTempGrense)} °C.` : ""}`);
        p.pv = "oversikt"; render();
      }).catch(() => { knapp.disabled = false; knapp.textContent = "Registrer melk"; });
      return true;
    }
    if (form.id === "prSvinnForm") {
      const antall = tall($("#sv-antall").value); if (!(antall > 0)) { toast("Skriv antall."); return true; }
      const x = { id: nyId("svinn"), dato: $("#sv-dato").value || iDag(), varenr: $("#sv-vare").value, antall, arsak: $("#sv-arsak").value, lot: $("#sv-lot").value.trim(), merknad: $("#sv-merk").value.trim(), av: brukerNavn() };
      const { id, ...d } = x; p.svinn.push(x);
      skriv("svinn/" + id, d).then(() => toast(`✓ Svinn registrert: ${nf(antall)} × ${prod(x.varenr)?.kort || x.varenr}.`)).catch(e => { p.svinn = p.svinn.filter(z => z.id !== id); feilLagring(e); });
      form.reset(); $("#sv-dato").value = iDag(); S.utkast.delete("prSvinnForm|"); render();
      return true;
    }
    if (form.id === "prSatsForm") {
      const v = id => tall($("#ps-" + id).value); const d = { melkKrL: v("melkKrL"), floteKrL: v("floteKrL"), timeKr: v("timeKr"), batchKr: v("batchKr"), kWh: v("kWh"), nettleieKr: v("nettleieKr"), stromReserve: v("stromReserve"), vannKrM3: v("vannKrM3"), omrade: sats().omrade, tempGrense: v("tempGrense"), melkTempGrense: v("melkTempGrense"), emb: {} };
      prodAlle().filter(x => x.gruppe !== "Melk").forEach(x => { const el = $("#ps-emb-" + x.varenr); if (el) d.emb[x.varenr] = tall(el.value); });
      p.innst = d; skriv("innst/satser", { data: d }).then(() => toast("✓ Satsene er lagret.")).catch(feilLagring); render();
      return true;
    }
    return false;
  }
  /* Krymper bildet til maks 1600 px før opplasting, så det går raskt også på mobilnett. */
  function krympBilde(fil) {
    return new Promise((ok, nei) => {
      const img = new Image(), url = URL.createObjectURL(fil);
      img.onload = () => { const maks = 1600, s = Math.min(1, maks / Math.max(img.width, img.height)); const c = document.createElement("canvas"); c.width = Math.round(img.width * s); c.height = Math.round(img.height * s); c.getContext("2d").drawImage(img, 0, 0, c.width, c.height); URL.revokeObjectURL(url); c.toBlob(b => b ? ok(b) : nei(new Error("bilde")), "image/jpeg", 0.82); };
      img.onerror = () => { URL.revokeObjectURL(url); nei(new Error("bilde")); };
      img.src = url;
    });
  }
  /* Oppdaterer stoppeklokka hvert 15. sekund uten å tegne hele siden. */
  setInterval(() => { const el = document.getElementById("prKlokke"); const y = S && S.pr && S.pr.edit; if (el && y && pågår(y)) el.textContent = tidTekst(minutter(y)); }, 15000);

  return { vis, klikk, endring, submit, last, poll };
})();
