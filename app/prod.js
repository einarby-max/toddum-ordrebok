/* Produksjon: melkehenting, ysting (tid, kontrollpunkter, resultat), kostnad per ysting, produsert mot bestilt og temperaturlogg fra kjølerommet.
   Data lagres i SharePoint-listene Ordrebok Melkehenting, Ordrebok Ysting, Ordrebok Temperatur og Ordrebok Innstillinger.
   Bilder av melketanken lagres i dokumentbiblioteket under Ysteri/Melkehenting. */
const Prod = (() => {
  /* Standardsatser fra Toddum Styringsmodell v9. Kan endres under Produksjon → Satser. */
  const SATSER = {
    melkKrL: 8, floteKrL: 65.23, timeKr: 294.97,
    batchKr: 968.2,           // løpe 41,8 + kalsiumklorid 12,8 + kultur 448,8 + frakt 75 + strøm 239,8 + vann/vask 150
    tempGrense: 4, melkTempGrense: 4,
    emb: { "0001":4.21, "0002":22.97, "0010":4.21, "0011":22.97, "0020":4.21, "0022":22.97, "0009":4.21, "0008":10.73, "0041":34.71, "0031":7.88, "0030":22.97, "0032":0, "0042":8.19 },
    modell: { "0001":14.52, "0002":232.2, "0010":16.76, "0011":190.5, "0020":44.73, "0022":437.76, "0030":290.55, "0031":31.97, "0009":46.7, "0008":89.46, "0041":660.05, "0032":88.96, "0042":133.29 },
  };
  const KONTROLL = ["Pasteurisering °C", "Pasteurisering tid (s)", "pH", "Melketemperatur °C", "Ostemasse kg", "Annet"];
  const P = () => S.pr || (S.pr = { pv: "oversikt", ystinger: [], melk: [], temp: [], innst: null, edit: null, loaded: false, laster: false, tempValg: 7, tempForslag: null, melkBilde: null, melkBildeUrl: "", lagrerT: null, feilTilgang: false });
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
      const [y, m, i] = await Promise.all([Store.hent("ysting"), Store.hent("melk"), Store.hent("innst")]);
      p.ystinger = y; p.melk = m; const s = i.find(x => x.id === "satser"); p.innst = s ? s.data : null; p.loaded = true;
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
  function beregn(y) {
    const s = sats(), prd = y.produsert || {};
    const melk = tall(y.melkLiter), flote = tall(y.flote), min = minutter(y), pers = Math.max(1, tall(y.personer) || 1);
    const timer = min / 60 * pers;
    let emb = 0, kg = 0, verdi = 0, modell = 0, enh = 0;
    Object.entries(prd).forEach(([v, n]) => { n = tall(n); if (!n) return; const p = prod(v); enh += n; emb += n * (s.emb[v] || 0); kg += kgNetto(n, p); verdi += n * (p?.pris || 0); modell += n * (s.modell[v] || 0); });
    const k = { melk: melk * s.melkKrL, flote: flote * s.floteKrL, batch: (melk || kg) ? s.batchKr : 0, emb, arbeid: timer * s.timeKr };
    k.sum = k.melk + k.flote + k.batch + k.emb + k.arbeid;
    return { melk, flote, min, pers, timer, kg, enh, verdi, modell, k, perKg: kg ? k.sum / kg : 0, db: verdi - k.sum, dbPst: verdi ? (verdi - k.sum) / verdi : 0, utbytte: melk ? kg / melk * 100 : 0, timerPer100: kg ? timer / kg * 100 : 0 };
  }
  const lotFor = dato => batchFra(dato);
  function bestiltUke(key) { const sum = {}; if (key !== uke().key) return null; S.ordre.forEach(o => linjer(o).forEach(l => sum[l.varenr] = (sum[l.varenr] || 0) + l.antall)); return sum; }
  function produsertUke(key) { const sum = {}; P().ystinger.filter(y => y.uke === key).forEach(y => Object.entries(y.produsert || {}).forEach(([v, n]) => sum[v] = (sum[v] || 0) + tall(n))); return sum; }
  function sporing(lot) { const ut = []; S.ordre.forEach(o => linjer(o).forEach(l => { const b = l.batch || batchFra(o.prodDato); if (b && b === lot) ut.push({ kunde: o.kundeNavn, varenr: l.varenr, antall: l.antall }); })); return ut; }
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
    const y = { id: nyId("ysting"), dato: d, uke: u.key, lot: lotFor(d), status: "pågår", melkLiter: liter, melkIds: ml.map(m => m.id), flote: 0, tid: [], personer: 2, minutter: 0, kontroller: [], produsert: {}, esker: 0, svinnKg: 0, merknad: "", registrertAv: brukerNavn() };
    ml.forEach(m => lagreMelk({ ...m, ystingId: y.id }).catch(() => {}));
    P().pv = "ysting"; lagreYsting(y, true); render();
  }

  /* ---------- visning ---------- */
  function vis() {
    const p = P(); if (!p.loaded) { last(); return `<div class="empty"><h3>Henter produksjonsdata …</h3></div>`; }
    const fane = (k, l) => `<button class="btn ${p.pv === k ? "primary" : ""}" data-act="prVis" data-v="${k}">${l}</button>`;
    const meny = `<div class="toolbar">${fane("oversikt", "Oversikt")}${fane("melk", "Hent melk")}${fane("ysting", p.edit ? "Ysting" : "Ny ysting")}${fane("temp", "Kjølerom")}${fane("satser", "Satser")}</div>`;
    const tilg = p.feilTilgang ? `<p class="notice"><b>Ikke lagret.</b> Kontoen din har ikke skrivetilgang til produksjonslistene i SharePoint. Be Einar eller Silje gi deg tilgang.</p>` : "";
    let h = "";
    if (p.pv === "melk") h = visMelk(); else if (p.pv === "ysting") h = visYsting(); else if (p.pv === "temp") h = visTemp(); else if (p.pv === "satser") h = visSatser(); else h = visOversikt();
    return `<h2>Produksjon</h2>${meny}${tilg}${h}`;
  }

  function visOversikt() {
    const p = P(), u = uke(), ml = ledigMelk();
    const aktiv = p.ystinger.find(y => y.status !== "ferdig");
    const best = bestiltUke(u.key) || {}, prodd = produsertUke(u.key);
    const varer = prodAlle().filter(x => best[x.varenr] || prodd[x.varenr]);
    const siste = [...p.ystinger].sort((a, b) => (b.dato || "").localeCompare(a.dato || "") || (b.endret || "").localeCompare(a.endret || "")).slice(0, 12);
    const d30 = new Date(Date.now() - 30 * 864e5).toISOString().slice(0, 10);
    const ferdig30 = p.ystinger.filter(y => y.status === "ferdig" && y.dato >= d30).map(beregn);
    const sum = (f) => ferdig30.reduce((s, b) => s + f(b), 0);
    const t30 = { kg: sum(b => b.kg), melk: sum(b => b.melk), timer: sum(b => b.timer), kost: sum(b => b.k.sum), verdi: sum(b => b.verdi) };
    const sisteTemp = p.temp[p.temp.length - 1];
    return `
    <div class="toolbar">
      <button class="btn primary lagre" data-act="prVis" data-v="melk">Hent melk</button>
      ${aktiv ? `<button class="btn primary lagre" data-act="prApne" data-id="${esc(aktiv.id)}">Fortsett ysting ${esc(dmy(aktiv.dato))}</button>` : `<button class="btn primary lagre" data-act="prNy">Start ny ysting</button>`}
    </div>
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
    ${siste.length ? `<div class="tablewrap"><table class="grid"><thead><tr><th>Dato</th><th>Lot</th><th class="n">Melk (l)</th><th class="n">Timer</th><th class="n">Kg ost</th><th class="n">Kg/100 l</th><th class="n">Kost</th><th class="n">Kr/kg</th><th class="n">DB</th><th>Status</th></tr></thead><tbody>
    ${siste.map(y => { const b = beregn(y); return `<tr data-act="prApne" data-id="${esc(y.id)}" style="cursor:pointer"><td>${esc(dmy(y.dato))}</td><td>${esc(y.lot)}</td><td class="n">${nf(b.melk)}</td><td class="n">${nf(b.timer, 1)}</td><td class="n">${nf(b.kg, 1)}</td><td class="n">${b.melk ? nf(b.utbytte, 1) : "–"}</td><td class="n">${nf(Math.round(b.k.sum))}</td><td class="n">${b.kg ? nf(Math.round(b.perKg)) : "–"}</td><td class="n">${b.verdi ? nf(b.dbPst * 100) + " %" : "–"}</td><td><span class="pill ${y.status === "ferdig" ? "st-sendt" : "st-produsert"}">${y.status === "ferdig" ? "Ferdig" : "Pågår"}</span></td></tr>`; }).join("")}
    </tbody></table></div>` : `<div class="empty"><h3>Ingen ystinger registrert ennå</h3>Trykk «Start ny ysting» når dere begynner.</div>`}`;
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
      return `<div class="empty"><h3>Ingen ysting er åpen</h3>${ml.length ? `${nf(ml.reduce((s, m) => s + tall(m.liter), 0))} liter melk er klar og blir koblet til ystingen.` : "Ingen melk er registrert ennå. Du kan likevel starte og skrive liter selv."}<div class="actions" style="justify-content:center;margin-top:12px"><button class="btn primary lagre" data-act="prNy">Start ny ysting</button></div></div>`;
    }
    const b = beregn(y), kjører = pågår(y), s = sats(), ro = y.status === "ferdig";
    const spor = sporing(y.lot);
    const ml = arr(y.melkIds).length ? p.melk.filter(m => arr(y.melkIds).includes(m.id)) : [];
    return `<form class="card form" id="prYstForm" data-id="${esc(y.id)}">
      <div class="thead"><h3>Ysting ${esc(dmy(y.dato))} · lot ${esc(y.lot)}</h3><span class="pill ${ro ? "st-sendt" : "st-produsert"}">${ro ? "Ferdig" : "Pågår"}</span></div>
      <h4>1. Melk og råvarer</h4>
      <div class="row">
        <label class="f">Dato<input type="date" class="txt" id="py-dato" data-pf="dato" value="${esc(y.dato)}"></label>
        <label class="f">Melk (liter)<input class="txt" id="py-melkLiter" data-pf="melkLiter" inputmode="decimal" value="${esc(y.melkLiter || "")}"></label>
        <label class="f">Fløte (liter)<input class="txt" id="py-flote" data-pf="flote" inputmode="decimal" value="${esc(y.flote || "")}"></label>
      </div>
      ${ml.length ? `<p class="hint">Fra melkehenting: ${ml.map(m => `${nf(tall(m.liter))} l ${esc(datoKl(m.tid))}${m.temp !== "" && m.temp != null ? `, ${nf(tall(m.temp), 1)} °C` : ""}${m.bilde ? ` (<a href="${esc(m.bilde)}" target="_blank" rel="noopener">bilde</a>)` : ""}`).join(" · ")}</p>` : ""}
      <h4>2. Tid</h4>
      <div class="actions">
        ${kjører ? `<button type="button" class="btn danger lagre" data-act="prStopp">■ Stopp tid</button>` : `<button type="button" class="btn primary lagre" data-act="prStart">▶ Start tid</button>`}
        <span style="font-size:26px;font-family:var(--f-display)" id="prKlokke">${tidTekst(b.min)}</span>
        <label class="f" style="width:110px">Personer<input class="txt" id="py-personer" data-pf="personer" inputmode="numeric" value="${esc(y.personer || 1)}"></label>
        <label class="f" style="width:150px">Rett total (min)<input class="txt" id="py-minutter" data-pf="minutter" inputmode="numeric" placeholder="${b.min}" value="${y.minutterManuelt ? esc(y.minutter) : ""}"></label>
      </div>
      <p class="hint">Tiden teller fra «Start» til «Stopp», og dere kan starte og stoppe flere ganger. Arbeidstimer = tid × antall personer = <b>${nf(b.timer, 1)} timer</b>. Har dere glemt å starte, skriv riktig total i «Rett total».</p>
      <h4>3. Kontrollpunkter</h4>
      ${arr(y.kontroller).length ? `<div class="tablewrap"><table class="grid"><tbody>${y.kontroller.map((k, i) => `<tr><td>${esc(klokke(k.tid))}</td><td>${esc(k.type)}</td><td class="n"><b>${esc(k.verdi)}</b></td><td class="hint">${esc(k.merknad || "")}</td><td><button type="button" class="btn sm ghost danger" data-act="prSlettKontroll" data-i="${i}" aria-label="Slett">✕</button></td></tr>`).join("")}</tbody></table></div>` : `<p class="hint">Ingen kontrollpunkter registrert.</p>`}
      <div class="row">
        <label class="f">Hva<select class="txt" id="pk-type">${KONTROLL.map(k => `<option>${k}</option>`).join("")}</select></label>
        <label class="f">Verdi<input class="txt" id="pk-verdi" inputmode="decimal"></label>
        <label class="f">Merknad<input class="txt" id="pk-merk"></label>
      </div>
      <div class="actions"><button type="button" class="btn" data-act="prKontroll">Legg til kontrollpunkt</button></div>
      <h4>4. Produsert</h4>
      <div class="lines">${prodSort().filter(x => x.gruppe !== "Melk").map(x => { const n = tall((y.produsert || {})[x.varenr]); return `<div class="line bl-${esc(x.gruppe)} ${n ? "has" : ""}">
        <div class="pn">${esc(x.kort)}<small>#${esc(x.varenr)}</small>${n ? `<small class="ant">${nf(kgNetto(n, x), 1)} kg</small>` : ""}</div>
        <div class="qty"><button type="button" class="stp" data-act="prMinus" data-v="${esc(x.varenr)}" ${n ? "" : "disabled"} aria-label="Én mindre ${esc(x.kort)}">−</button><input id="pp-${esc(x.varenr)}" data-pp="${esc(x.varenr)}" inputmode="decimal" value="${n || ""}" placeholder="–" aria-label="Antall ${esc(x.kort)}"><button type="button" class="stp" data-act="prPluss" data-v="${esc(x.varenr)}" aria-label="Én mer ${esc(x.kort)}">+</button><span class="u">${esc(enhetNavn(x, 2))}</span></div></div>`; }).join("")}</div>
      <div class="row">
        <label class="f">Antall esker / kartonger<input class="txt" id="py-esker" data-pf="esker" inputmode="numeric" value="${esc(y.esker || "")}"></label>
        <label class="f">Svinn (kg)<input class="txt" id="py-svinnKg" data-pf="svinnKg" inputmode="decimal" value="${esc(y.svinnKg || "")}"></label>
      </div>
      <label class="f">Merknad<textarea class="txt" id="py-merknad" data-pf="merknad" rows="2">${esc(y.merknad || "")}</textarea></label>
      <h4>Kostnad for denne ystingen</h4>
      <div class="tablewrap"><table class="grid"><tbody>
        <tr><td>Melk ${nf(b.melk)} l × ${nf(s.melkKrL, 2)} kr</td><td class="n">${kr(b.k.melk)}</td></tr>
        <tr><td>Fløte ${nf(b.flote, 1)} l × ${nf(s.floteKrL, 2)} kr</td><td class="n">${kr(b.k.flote)}</td></tr>
        <tr><td>Kultur, løpe, strøm, vask (fast per ysting)</td><td class="n">${kr(b.k.batch)}</td></tr>
        <tr><td>Emballasje (${nf(b.enh)} enheter)</td><td class="n">${kr(b.k.emb)}</td></tr>
        <tr><td>Arbeid ${nf(b.timer, 1)} t × ${nf(s.timeKr, 0)} kr</td><td class="n">${kr(b.k.arbeid)}</td></tr>
      </tbody><tfoot>
        <tr><td>Sum kostnad · ${nf(b.kg, 1)} kg ost${b.melk ? ` · ${nf(b.utbytte, 1)} kg/100 l` : ""}</td><td class="n">${kr(b.k.sum)}${b.kg ? ` <span class="hint">(${nf(b.perKg)} kr/kg)</span>` : ""}</td></tr>
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
      <div class="row">${fi("melkKrL", "Melk (kr/l)", s.melkKrL)}${fi("floteKrL", "Fløte (kr/l)", s.floteKrL)}${fi("timeKr", "Arbeid (kr/time)", s.timeKr)}${fi("batchKr", "Kultur, løpe, strøm, vask per ysting (kr)", s.batchKr)}${fi("tempGrense", "Grense kjølerom (°C)", s.tempGrense)}${fi("melkTempGrense", "Grense melketank (°C)", s.melkTempGrense)}</div>
      <h4>Emballasje per enhet (kr)</h4>
      <div class="row">${prodAlle().filter(x => x.gruppe !== "Melk").map(x => fi("emb-" + x.varenr, `${x.kort} (#${x.varenr})`, s.emb[x.varenr] ?? 0)).join("")}</div>
      <div class="actions"><button class="btn primary lagre" type="submit">Lagre satser</button><button class="btn ghost" type="button" data-act="prSatsStd">Tilbake til standard</button></div>
    </form>`;
  }

  /* ---------- handlinger ---------- */
  function endreY(f) { const y = P().edit; if (!y) return; const n = { ...y }; f(n); lagreYsting(n); planRender(); }
  function klikk(el) {
    const a = el.dataset.act; if (!a || !a.startsWith("pr")) return false;
    const p = P();
    if (a === "prVis") { p.pv = el.dataset.v; if (p.pv === "ysting" && !p.edit) { const akt = p.ystinger.find(y => y.status !== "ferdig"); if (akt) p.edit = akt; } render(); }
    else if (a === "prNy") nyYsting();
    else if (a === "prApne") { p.edit = p.ystinger.find(y => y.id === el.dataset.id) || null; p.pv = "ysting"; render(); window.scrollTo(0, 0); }
    else if (a === "prLukk") { p.edit = null; p.pv = "oversikt"; render(); }
    else if (a === "prStart") endreY(y => { y.tid = [...arr(y.tid), { fra: nå() }]; y.minutterManuelt = false; });
    else if (a === "prStopp") endreY(y => { y.tid = arr(y.tid).map(t => t.til ? t : { ...t, til: nå() }); y.minutter = minutter(y); });
    else if (a === "prKontroll") { const v = $("#pk-verdi")?.value.trim(); if (!v) { toast("Skriv en verdi."); return true; } const k = { tid: nå(), type: $("#pk-type").value, verdi: v, merknad: $("#pk-merk").value.trim(), av: brukerNavn() }; endreY(y => { y.kontroller = [...arr(y.kontroller), k]; }); }
    else if (a === "prSlettKontroll") endreY(y => { y.kontroller = arr(y.kontroller).filter((_, i) => i !== +el.dataset.i); });
    else if (a === "prPluss" || a === "prMinus") { const v = el.dataset.v; endreY(y => { const pr = { ...(y.produsert || {}) }; const n = Math.max(0, tall(pr[v]) + (a === "prPluss" ? 1 : -1)); if (n) pr[v] = n; else delete pr[v]; y.produsert = pr; }); }
    else if (a === "prFerdig") { const y = P().edit; if (!y) return true; if (!Object.keys(y.produsert || {}).length) { toast("Registrer hva som er produsert før ystingen avsluttes."); return true; } const n = { ...y, status: "ferdig", tid: arr(y.tid).map(t => t.til ? t : { ...t, til: nå() }) }; n.minutter = minutter(n); lagreYsting(n, true).then(() => toast(`✓ Ystingen ${dmy(n.dato)} er lagret. Lot ${n.lot}.`)); p.edit = null; p.pv = "oversikt"; render(); }
    else if (a === "prGjenapne") endreY(y => { y.status = "pågår"; });
    else if (a === "prSlett") { const y = p.edit; if (!y) return true; if (p.bekreftSlett === y.id) { p.bekreftSlett = null; p.ystinger = p.ystinger.filter(x => x.id !== y.id); p.melk.filter(m => m.ystingId === y.id).forEach(m => lagreMelk({ ...m, ystingId: "" }).catch(() => {})); p.edit = null; p.pv = "oversikt"; clearTimeout(p.lagrerT); p.lagrerT = null; skriv("ysting/" + y.id, null).then(() => toast("Ystingen er slettet.")).catch(feilLagring); render(); } else { p.bekreftSlett = y.id; render(); setTimeout(() => { if (p.bekreftSlett === y.id) { p.bekreftSlett = null; render(); } }, 4000); } }
    else if (a === "prTempDager") { p.tempValg = +el.dataset.v; render(); }
    else if (a === "prTempLagre") { el.disabled = true; lagreTemp(); }
    else if (a === "prTempAvbryt") { p.tempForslag = null; render(); }
    else if (a === "prSatsStd") { p.innst = null; skriv("innst/satser", { data: {} }).then(() => toast("Satsene er satt tilbake til standard.")).catch(feilLagring); render(); }
    else return false;
    return true;
  }
  function endring(e) {
    const t = e.target, p = P();
    if (t.dataset.prfoto && t.files && t.files[0]) { const f = t.files[0]; krympBilde(f).then(b => { p.melkBilde = b; if (p.melkBildeUrl) URL.revokeObjectURL(p.melkBildeUrl); p.melkBildeUrl = URL.createObjectURL(b); render(); }).catch(() => { p.melkBilde = f; p.melkBildeUrl = URL.createObjectURL(f); render(); }); return true; }
    if (t.id === "prTempFil" && t.files && t.files[0]) { const f = t.files[0]; t.value = ""; lesLoggfil(f).catch(err => toast("Klarte ikke å lese loggfila: " + (err.message || err))); return true; }
    if (t.dataset.pf) {
      const f = t.dataset.pf; let v = t.value.trim();
      endreY(y => {
        if (["melkLiter", "flote", "personer", "esker", "svinnKg"].includes(f)) y[f] = tall(v);
        else if (f === "minutter") { if (v === "") { y.minutterManuelt = false; y.minutter = 0; } else { y.minutter = Math.round(tall(v)); y.minutterManuelt = true; } }
        else if (f === "dato") { y.dato = v; y.uke = isoUke(new Date(v + "T12:00")).key; y.lot = lotFor(v); }
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
      const aktiv = p.ystinger.find(y => y.status !== "ferdig" && y.dato === iDag()); if (aktiv) m.ystingId = aktiv.id;
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
    if (form.id === "prSatsForm") {
      const v = id => tall($("#ps-" + id).value); const d = { melkKrL: v("melkKrL"), floteKrL: v("floteKrL"), timeKr: v("timeKr"), batchKr: v("batchKr"), tempGrense: v("tempGrense"), melkTempGrense: v("melkTempGrense"), emb: {} };
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
