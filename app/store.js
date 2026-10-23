/* Lagring i SharePoint-lister via Microsoft Graph, med innlogging via MSAL.
   Tre lister i Toddum-området: Ordrebok Produkter, Ordrebok Kunder og Ordrebok Ordre.
   Kolonnen Title holder nøkkelen (varenr for produkter, id for kunder og ordre). */
const Store = (() => {
  const T = "text", M = "multi", N = "number", B = "bool", J = "json";
  const LISTER = {
    produkter: { navn: "Ordrebok Produkter", nokkel: "varenr", kol: [
      ["kort",T],["navn",T],["gruppe",T],["enhet",T],["pakning",T],["perDP",N],["stkNettoG",N],["stkBruttoG",N],["dpBruttoG",N],
      ["gtin",T],["dpGtin",T],["epd",T],["varetype",T],["sort",N],["aktiv",B] ] },
    kunder: { navn: "Ordrebok Kunder", nokkel: "id", kol: [
      ["navn",T],["kundenr",T],["orgnr",T],["adresse",T],["adresse2",T],["postnr",T],["poststed",T],["land",T],["epost",T],["telefon",T],
      ["kontakt",T],["kontaktEpost",T],["kontaktTlf",T],["levNavn",T],["levAdresse",M],["transport",T],["cmr",B],["leveringsdag",T],
      ["merknad",M],["instruks",M],["fastFrekvens",T],["fast",J],["inaktiv",B] ] },
    ordre: { navn: "Ordrebok Ordre", nokkel: "id", kol: [
      ["uke",T,{indexed:true}],["kundeId",T],["kundeNavn",T],["linjer",J],["transport",T],["sendes",T],["prodDato",T],["bestnr",T],
      ["merknad",M],["status",T],["kilde",T],["opprettet",T],["endret",T] ] },
  };
  const SCOPES = ["Sites.Manage.All"];
  let app, konto, siteId;
  const listeId = {}, spId = { produkter:{}, kunder:{}, ordre:{} }, kjeder = {};

  function feil(melding, status, kode){ const e = new Error(melding); e.status = status; e.code = kode; return e; }

  async function init(){
    const C = window.CONFIG || {};
    if(!C.clientId || C.clientId.startsWith("SETT")) return { tilstand: "ikke-konfigurert" };
    app = new msal.PublicClientApplication({
      auth: { clientId: C.clientId, authority: `https://login.microsoftonline.com/${C.tenant}`, redirectUri: location.origin + "/" },
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
    return { tilstand: mangler.length ? "oppsett" : "klar", mangler, bruker: { navn: konto.name || konto.username, epost: konto.username } };
  }
  const loggInn = () => app.loginRedirect({ scopes: SCOPES, prompt: "select_account" });
  const loggUt = () => app.logoutRedirect({ account: konto, postLogoutRedirectUri: location.origin + "/" });

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
      if(l) listeId[k] = l.id; else mangler.push(k);
    }
    return mangler;
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
    return r;
  }
  function tilFelter(k, r){
    const def = LISTER[k]; const f = { Title: String(r[def.nokkel]) };
    for(const [navn, type] of def.kol){
      const v = r[navn];
      if(type === J) f[navn] = JSON.stringify(v ?? {});
      else if(type === B) f[navn] = !!v;
      else if(type === N) f[navn] = Number.isFinite(Number(v)) ? Number(v) : 0;
      else f[navn] = v == null ? "" : String(v).slice(0, type === T ? 255 : 60000);
    }
    return f;
  }

  async function hent(k, uke){
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
  async function eksporter(){ const [produkter, kunder, ordre] = await Promise.all([hent("produkter"), hent("kunder"), hent("ordre")]); return { eksportert: new Date().toISOString(), produkter, kunder, ordre }; }

  return { init, loggInn, loggUt, opprettLister, hent, lagre, slett, importer, eksporter, LISTER };
})();
