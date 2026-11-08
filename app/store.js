/* Lagring i SharePoint-lister via Microsoft Graph, med innlogging via MSAL.
   Tre lister i Toddum-området: Ordrebok Produkter, Ordrebok Kunder og Ordrebok Ordre.
   Kolonnen Title holder nøkkelen (varenr for produkter, id for kunder og ordre). */
const Store = (() => {
  const T = "text", M = "multi", N = "number", B = "bool", J = "json";
  const LISTER = {
    produkter: { navn: "Ordrebok Produkter", nokkel: "varenr", kol: [
      ["kort",T],["navn",T],["gruppe",T],["enhet",T],["pakning",T],["perDP",N],["stkNettoG",N],["stkBruttoG",N],["dpBruttoG",N],
      ["gtin",T],["dpGtin",T],["epd",T],["varetype",T],["sort",N],["aktiv",B],["pris",N] ] },
    kunder: { navn: "Ordrebok Kunder", nokkel: "id", kol: [
      ["navn",T],["kundenr",T],["orgnr",T],["adresse",T],["adresse2",T],["postnr",T],["poststed",T],["land",T],["epost",T],["telefon",T],
      ["kontakt",T],["kontaktEpost",T],["kontaktTlf",T],["levNavn",T],["levAdresse",M],["transport",T],["cmr",B],["leveringsdag",T],
      ["merknad",M],["instruks",M],["fastFrekvens",T],["fast",J],["inaktiv",B],["priser",J],["antallI",T],["fakturering",T,{ny:true}] ] },
    transport: { navn: "Ordrebok Transport", nokkel: "id", valgfri: true, kol: [
      ["navn",T],["adresse",M],["kontakt",T],["telefon",T],["epost",T],["cmr",B],["merknad",M],["aktiv",B],["sort",N],["nettside",T] ] },
    melk: { navn: "Ordrebok Melkehenting", nokkel: "id", valgfri: true, kol: [
      ["tid",T],["dato",T],["liter",N],["temp",T],["hentetAv",T],["bilde",M],["ystingId",T],["merknad",M] ] },
    ysting: { navn: "Ordrebok Ysting", nokkel: "id", valgfri: true, kol: [
      ["dato",T],["uke",T],["lot",T],["status",T],["melkLiter",N],["melkIds",J],["flote",N],["tid",J],["personer",N],["minutter",N],["minutterManuelt",B],
      ["kontroller",J],["produsert",J],["esker",N],["svinnKg",N],["merknad",M],["registrertAv",T],["endret",T],
      ["ostemasseKg",N],["overfortUtKg",N],["overfortInn",J],["vannStart",N],["vannSlutt",N],["strompris",N],
      ["batchNr",N,{ny:true}],["osteType",T,{ny:true}],["ystetAv",T,{ny:true}],["startTemp",T,{ny:true}],["varmeElementer",T,{ny:true}],["tempVarmeAv",T,{ny:true}],["logg",J,{ny:true}] ], reserve: "produsert" },
    svinn: { navn: "Ordrebok Svinn", nokkel: "id", valgfri: true, kol: [
      ["dato",T],["varenr",T],["antall",N],["arsak",T],["lot",T],["merknad",M],["av",T] ] },
    temperatur: { navn: "Ordrebok Temperatur", nokkel: "id", valgfri: true, kol: [
      ["fra",T],["til",T],["filnavn",T],["antall",N],["min",N],["max",N],["over",N],["data",J],["lastetOpp",T],["av",T] ] },
    lager: { navn: "Ordrebok Lager", nokkel: "id", valgfri: true, kol: [
      ["tid",T],["dato",T],["varenr",T],["antall",N],["rom",T],["type",T],["ref",T],["lot",T],["av",T],["merknad",M] ] },
    innst: { navn: "Ordrebok Innstillinger", nokkel: "id", valgfri: true, kol: [ ["data",J] ] },
    ordre: { navn: "Ordrebok Ordre", nokkel: "id", kol: [
      ["uke",T,{indexed:true}],["kundeId",T],["kundeNavn",T],["linjer",J],["transport",T],["sendes",T],["prodDato",T],["bestnr",T],
      ["merknad",M],["status",T],["kilde",T],["opprettet",T],["endret",T],["fakturering",T,{ny:true}],["sendtInfo",J,{ny:true}] ] },
  };
  const SCOPES = ["Sites.Manage.All"];
  let app, konto, siteId;
  const listeId = {}, kjente = {}, spId = { produkter:{}, kunder:{}, ordre:{}, transport:{}, melk:{}, ysting:{}, temperatur:{}, innst:{}, svinn:{}, lager:{} }, kjeder = {};

  function basisUrl(){ return location.origin + location.pathname.replace(/[^/]*$/, ""); }
  function feil(melding, status, kode){ const e = new Error(melding); e.status = status; e.code = kode; return e; }

  async function init(){
    const C = window.CONFIG || {};
    if(!C.clientId || C.clientId.startsWith("SETT")) return { tilstand: "ikke-konfigurert" };
    app = new msal.PublicClientApplication({
      auth: { clientId: C.clientId, authority: `https://login.microsoftonline.com/${C.tenant}`, redirectUri: basisUrl() },
      cache: { cacheLocation: "localStorage" },
    });
    await app.initialize();
    const svar = await app.handleRedirectPromise();
    konto = (svar && svar.account) || app.getActiveAccount() || app.getAllAccounts()[0];
    if(!konto) return { tilstand: "logget-ut" };
    app.setActiveAccount(konto);
    const site = await g("GET", `/sites/${C.siteHost}:${C.sitePath}?$select=id,displayName`);
    siteId = site.id;
    const mangler = await finnLister();
    if(!mangler.length){
      /* Lister som kom i senere versjoner (f.eks. transport) opprettes automatisk. */
      for(const [k, def] of Object.entries(LISTER)) if(def.valgfri && !listeId[k]){
        try { const l = await g("POST", `/sites/${siteId}/lists`, { displayName: def.navn, list: { template: "genericList" }, columns: def.kol.map(kolonne) }); listeId[k] = l.id; }
        catch(e){ console.warn("Kunne ikke opprette " + def.navn, e); }
      }
      await sikreKolonner().catch(e => console.warn("Kunne ikke legge til nye kolonner", e));
    }
    return { tilstand: mangler.length ? "oppsett" : "klar", mangler, bruker: { navn: konto.name || konto.username, epost: konto.username } };
  }
  const loggInn = () => app.loginRedirect({ scopes: SCOPES, prompt: "select_account" });
  const loggUt = () => app.logoutRedirect({ account: konto, postLogoutRedirectUri: basisUrl() });

  async function token(){
    try { return (await app.acquireTokenSilent({ scopes: SCOPES, account: konto })).accessToken; }
    catch(e){
      if(e instanceof msal.InteractionRequiredAuthError || /interaction_required|login_required|consent_required/.test(e.errorCode||"")){
        await app.acquireTokenRedirect({ scopes: SCOPES, account: konto });
      }
      throw e;
    }
  }
  async function g(metode, sti, kropp, ekstra){
    let svar;
    for(let forsok = 0; forsok < 4; forsok++){
      const t = await token();
      try {
        svar = await fetch(sti.startsWith("http") ? sti : "https://graph.microsoft.com/v1.0" + sti, {
          method: metode,
          headers: { Authorization: "Bearer " + t, "Content-Type": "application/json", ...(ekstra||{}) },
          body: kropp ? JSON.stringify(kropp) : undefined,
        });
      } catch(_) { throw feil("Ingen kontakt med Microsoft 365. Sjekk nettforbindelsen.", 0, "nett"); }
      if(svar.status === 429 || svar.status === 503){ const vent = Number(svar.headers.get("Retry-After")) || (2 ** forsok); await new Promise(r => setTimeout(r, vent * 1000)); continue; }
      break;
    }
    if(svar.status === 204) return null;
    const j = await svar.json().catch(() => null);
    if(!svar.ok) throw feil(j?.error?.message || svar.statusText, svar.status, j?.error?.code);
    return j;
  }

  async function finnLister(){
    const j = await g("GET", `/sites/${siteId}/lists?$select=id,displayName&$top=999`);
    const mangler = [];
    for(const [k, def] of Object.entries(LISTER)){
      const l = j.value.find(x => x.displayName === def.navn);
      if(l) listeId[k] = l.id; else if(!def.valgfri) mangler.push(k);
    }
    return mangler;
  }
  /* Nye felt i senere versjoner (f.eks. pris) legges til i eksisterende lister. */
  /* Husker hvilke kolonner som faktisk finnes, slik at lagring aldri sender felt SharePoint ikke kjenner
     (det ville avvist hele lagringen). Mangler en kolonne, og brukeren ikke har lov til å opprette den,
     legges verdien i reservekolonnen (JSON) i stedet, så ingenting går tapt. */
  async function sikreKolonner(){
    for(const [k, def] of Object.entries(LISTER)){
      if(!listeId[k]) continue;
      let finnes;
      try { const j = await g("GET", `/sites/${siteId}/lists/${listeId[k]}/columns?$select=name&$top=999`); finnes = new Set(j.value.map(c => c.name)); }
      catch(e){ console.warn("Kunne ikke lese kolonnene i " + def.navn, e); continue; }
      kjente[k] = finnes;
      for(const kol of def.kol) if(!finnes.has(kol[0])){
        try { await g("POST", `/sites/${siteId}/lists/${listeId[k]}/columns`, kolonne(kol)); finnes.add(kol[0]); }
        catch(e){ console.warn(`Kunne ikke legge til kolonnen ${kol[0]} i ${def.navn}`, e); }
      }
    }
  }
  function harKol(k, navn){
    const kol = LISTER[k].kol.find(c => c[0] === navn);
    if(kjente[k]) return kjente[k].has(navn);
    return !(kol && kol[2] && kol[2].ny);
  }
  function kolonne([navn, type, valg]){
    const c = { name: navn };
    if(type === N) c.number = {};
    else if(type === B) c.boolean = {};
    else if(type === M || type === J) c.text = { allowMultipleLines: true, linesForEditing: 4, textType: "plain" };
    else c.text = { maxLength: 255 };
    if(valg && valg.indexed) c.indexed = true;
    return c;
  }
  async function opprettLister(){
    for(const [k, def] of Object.entries(LISTER)){
      if(listeId[k]) continue;
      const l = await g("POST", `/sites/${siteId}/lists`, { displayName: def.navn, list: { template: "genericList" }, columns: def.kol.map(kolonne) });
      listeId[k] = l.id;
    }
  }

  function fraFelter(k, f){
    const def = LISTER[k]; const r = { [def.nokkel]: f.Title };
    for(const [navn, type] of def.kol){
      const v = f[navn];
      if(type === J){ try { r[navn] = v ? JSON.parse(v) : {}; } catch(_) { r[navn] = {}; } }
      else if(type === B) r[navn] = v === true || v === "true" || v === 1;
      else if(type === N) r[navn] = typeof v === "number" ? v : (v == null || v === "" ? 0 : Number(v));
      else r[navn] = v == null ? "" : String(v);
    }
    /* Felt som ble lagret i reservekolonnen fordi kolonnen manglet, hentes fram igjen. */
    const res = def.reserve && r[def.reserve];
    if(res && typeof res === "object" && res.__ekstra){
      const x = res.__ekstra; delete res.__ekstra;
      for(const [navn, v] of Object.entries(x || {})){
        const kol = def.kol.find(c => c[0] === navn); if(!kol) continue;
        const tom = f[navn] == null || f[navn] === "";
        if(tom) r[navn] = v;
      }
    }
    return r;
  }
  function tilFelter(k, r){
    const def = LISTER[k]; const f = { Title: String(r[def.nokkel]) }, ekstra = {};
    for(const [navn, type] of def.kol){
      const v = r[navn];
      if(!harKol(k, navn)){ if(v !== undefined) ekstra[navn] = v; continue; }
      if(type === J) f[navn] = JSON.stringify(v ?? {});
      else if(type === B) f[navn] = !!v;
      else if(type === N) f[navn] = Number.isFinite(Number(v)) ? Number(v) : 0;
      else f[navn] = v == null ? "" : String(v).slice(0, type === T ? 255 : 60000);
    }
    if(def.reserve && Object.keys(ekstra).length && f[def.reserve] != null){
      const o = { ...(r[def.reserve] || {}), __ekstra: ekstra }; f[def.reserve] = JSON.stringify(o);
    }
    return f;
  }

  async function hent(k, uke){
    if(!listeId[k]) return [];
    let url = `/sites/${siteId}/lists/${listeId[k]}/items?$expand=fields&$top=999`;
    if(uke) url += `&$filter=fields/uke eq '${encodeURIComponent(uke)}'`;
    const ut = [];
    while(url){
      const j = await g("GET", url, null, { Prefer: "HonorNonIndexedQueriesWarningMayFailRandomly" });
      for(const it of j.value){ const r = fraFelter(k, it.fields); spId[k][r[LISTER[k].nokkel]] = it.id; ut.push(r); }
      url = j["@odata.nextLink"];
    }
    return ut;
  }
  function kjede(nokkel, fn){ const p = (kjeder[nokkel] || Promise.resolve()).then(fn); kjeder[nokkel] = p.catch(() => {}); return p; }
  function lagre(k, r){
    const id = String(r[LISTER[k].nokkel]);
    return kjede(k + "/" + id, async () => {
      const felt = tilFelter(k, r);
      const sp = spId[k][id];
      if(sp) await g("PATCH", `/sites/${siteId}/lists/${listeId[k]}/items/${sp}/fields`, felt);
      else { const it = await g("POST", `/sites/${siteId}/lists/${listeId[k]}/items`, { fields: felt }); spId[k][id] = it.id; }
    });
  }
  function slett(k, id){
    return kjede(k + "/" + id, async () => {
      const sp = spId[k][id]; if(!sp) return;
      await g("DELETE", `/sites/${siteId}/lists/${listeId[k]}/items/${sp}`);
      delete spId[k][id];
    });
  }
  async function importer(data, fremdrift){
    let ferdig = 0, hoppetOver = 0; const jobber = [];
    for(const k of ["produkter", "kunder", "ordre"]){
      const finnes = new Set((await hent(k)).map(r => String(r[LISTER[k].nokkel])));
      for(const r of (data[k] || [])){ if(finnes.has(String(r[LISTER[k].nokkel]))) { hoppetOver++; continue; } jobber.push([k, r]); }
    }
    const total = jobber.length; let i = 0;
    async function arbeider(){ while(i < jobber.length){ const [k, r] = jobber[i++]; await lagre(k, r); ferdig++; fremdrift && fremdrift(ferdig, total); } }
    await Promise.all([arbeider(), arbeider(), arbeider(), arbeider()]);
    return { ferdig, hoppetOver };
  }
  /* Laster opp en fil (f.eks. bilde av melketanken) til dokumentbiblioteket i Toddum-området. Mapper lages automatisk. */
  async function lastOppFil(sti, blob){
    const t = await token();
    const url = `https://graph.microsoft.com/v1.0/sites/${siteId}/drive/root:/${sti.split("/").map(encodeURIComponent).join("/")}:/content?@microsoft.graph.conflictBehavior=rename`;
    let svar;
    try { svar = await fetch(url, { method: "PUT", headers: { Authorization: "Bearer " + t, "Content-Type": blob.type || "application/octet-stream" }, body: blob }); }
    catch(_) { throw feil("Ingen kontakt med Microsoft 365. Sjekk nettforbindelsen.", 0, "nett"); }
    const j = await svar.json().catch(() => null);
    if(!svar.ok) throw feil(j?.error?.message || svar.statusText, svar.status, j?.error?.code);
    return { url: j.webUrl, id: j.id };
  }
  async function eksporter(){ const [produkter, kunder, ordre, transport] = await Promise.all([hent("produkter"), hent("kunder"), hent("ordre"), hent("transport")]); return { eksportert: new Date().toISOString(), produkter, kunder, ordre, transport }; }
  const harListe = k => !!listeId[k];

  return { init, loggInn, loggUt, opprettLister, hent, lagre, slett, importer, eksporter, harListe, lastOppFil, LISTER };
})();
