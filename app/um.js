/* Kobling til Unimicro via API (innlogging med PKCE, appen «Toddum ordrebok» i Unimicros utviklerportal).
   Innloggingen skjer på unimicro.html, som legger token i sessionStorage og sender brukeren tilbake.
   Miljø: test (test.unimicro.no) til appen er godkjent av Unimicro; deretter prod (app.unimicro.no). */
const Um = (() => {
  const CLIENT_ID = "46b68e78-3649-4386-99fa-e5e05d41677c";
  const ENV = { test: { auth: "https://test-login.unimicro.no", api: "https://test.unimicro.no", navn: "testselskapet" }, prod: { auth: "https://login.unimicro.no", api: "https://app.unimicro.no", navn: "Unimicro" } };
  const SCOPE = "openid profile AppFramework";
  const KONTO = "3040", MVA = "31";
  const ss = { get: k => { try { return sessionStorage.getItem(k); } catch (e) { return null; } }, set: (k, v) => { try { sessionStorage.setItem(k, v); } catch (e) {} }, del: k => { try { sessionStorage.removeItem(k); } catch (e) {} } };
  const env = () => ss.get("um_env") || "test";
  const company = () => ss.get("um_company") || "";
  const token = () => { const t = ss.get("um_token"), e = Number(ss.get("um_exp") || 0); return t && e > Date.now() + 60000 ? t : null; };
  const b64u = buf => btoa(String.fromCharCode(...new Uint8Array(buf))).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");
  const tilfeldig = n => b64u(crypto.getRandomValues(new Uint8Array(n)));
  const callbackUrl = () => new URL("unimicro.html", location.href).href.split("?")[0].split("#")[0];

  async function loggInn() {
    const e = env(), d = await fetch(ENV[e].auth + "/.well-known/openid-configuration").then(r => r.json());
    const verifier = tilfeldig(48), state = tilfeldig(16);
    const challenge = b64u(await crypto.subtle.digest("SHA-256", new TextEncoder().encode(verifier)));
    ss.set("um_verifier", verifier); ss.set("um_state", state); ss.set("um_env", e); ss.set("um_scope", SCOPE);
    ss.set("um_return", location.pathname + location.search);
    location.href = d.authorization_endpoint + "?" + new URLSearchParams({ client_id: CLIENT_ID, redirect_uri: callbackUrl(), response_type: "code", scope: SCOPE, state, code_challenge: challenge, code_challenge_method: "S256" });
  }
  async function req(method, path, body) {
    const t = token(); if (!t) throw Object.assign(new Error("Ikke logget inn i Unimicro"), { login: true });
    const h = { Authorization: "Bearer " + t, Accept: "application/json" }; if (company()) h.CompanyKey = company(); if (body) h["Content-Type"] = "application/json";
    const r = await fetch(ENV[env()].api + path, { method, headers: h, body: body ? JSON.stringify(body) : undefined });
    const txt = await r.text(); let j; try { j = JSON.parse(txt); } catch (e) { j = txt; }
    if (r.status === 401) { ss.del("um_token"); throw Object.assign(new Error("Innloggingen i Unimicro har gått ut"), { login: true }); }
    if (!r.ok) { const m = (j && (j.Message || j.message || (j.Messages && j.Messages[0] && j.Messages[0].Message))) || txt.slice(0, 200); throw Object.assign(new Error(`Unimicro svarte ${r.status}: ${m}`), { svar: j }); }
    return j;
  }
  const get = p => req("GET", p), post = (p, b) => req("POST", p, b);
  const q = s => encodeURIComponent(s);

  async function velgSelskap() {
    const c = await get("/api/init/companies"); const liste = Array.isArray(c) ? c : [];
    if (!liste.length) throw new Error("Brukeren har ikke tilgang til noe selskap i Unimicro.");
    const toddum = liste.find(x => (x.OrganizationNumber || "").replace(/\s/g, "") === AVSENDER.orgnr) || liste[0];
    ss.set("um_company", toddum.Key); ss.set("um_companyName", toddum.Name); return toddum;
  }
  async function klar() {
    if (!token()) { await loggInn(); return false; }
    if (!company()) await velgSelskap();
    return true;
  }

  /* Oppslag: konto 3040, mva-kode 31, produkter per varenummer, kunder per kundenummer. */
  async function oppslag() {
    const [konto, mva, prod, kunder] = await Promise.all([
      get(`/api/biz/accounts?filter=${q("AccountNumber eq " + KONTO)}&select=ID`),
      get(`/api/biz/vattypes?filter=${q("VatCode eq '" + MVA + "'")}&select=ID`),
      get(`/api/biz/products?select=ID,PartName,Name,PriceExVat&top=1000`),
      get(`/api/biz/customers?select=ID,CustomerNumber&top=5000`)]);
    return { kontoId: konto[0]?.ID, mvaId: mva[0]?.ID, prod: new Map(prod.map(p => [String(p.PartName), p])), kunder: new Map(kunder.map(k => [String(k.CustomerNumber), k])) };
  }
  async function lagProdukt(o, p) {
    const n = await post("/api/biz/products", { PartName: p.varenr, Name: p.navn || p.kort, PriceExVat: Number(p.pris) || 0, AccountID: o.kontoId, VatTypeID: o.mvaId, Unit: p.enhet || "stk" });
    o.prod.set(String(p.varenr), n); return n;
  }
  async function lagKunde(o, k) {
    const n = await post("/api/biz/customers", { CustomerNumber: Number(k.kundenr), Info: { Name: k.navn, InvoiceAddress: { AddressLine1: k.adresse || "", PostalCode: k.postnr || "", City: k.poststed || "", CountryCode: "NO" } } });
    o.kunder.set(String(k.kundenr), n); return n;
  }

  /* Sender ukas ordre (samme utvalg og regler som Excel-fila) som salgsordre. Ordre som allerede er sendt til samme miljø hoppes over. */
  async function sendUke(grunnlag) {
    if (!(await klar())) return;
    const e = env(), o = await oppslag();
    if (!o.kontoId || !o.mvaId) throw new Error("Fant ikke konto 3040 eller mva-kode 31 i Unimicro.");
    const nye = grunnlag.filter(g => !(g.o.unimicro && g.o.unimicro.env === e && g.o.unimicro.id));
    const mangler = { prod: new Set(), kunder: new Set() };
    nye.forEach(g => { if (!o.kunder.has(String(g.b.kundenr))) mangler.kunder.add(g.b); g.linjer.forEach(l => { if (!o.prod.has(String(l.varenr))) mangler.prod.add(l.p || { varenr: l.varenr }); }); });
    if (mangler.prod.size || mangler.kunder.size) {
      if (e !== "test") throw new Error(`Mangler i Unimicro: ${[...mangler.kunder].map(k => k.navn).concat([...mangler.prod].map(p => "vare " + p.varenr)).join(", ")}. Legg dem inn i Unimicro først.`);
      for (const p of mangler.prod) await lagProdukt(o, p);
      for (const k of mangler.kunder) await lagKunde(o, k);
    }
    const res = { sendt: [], feil: [], hoppet: grunnlag.length - nye.length };
    for (const g of nye) {
      try {
        const body = { CustomerID: o.kunder.get(String(g.b.kundenr)).ID, OrderDate: iDag(), DeliveryDate: g.levDato, OurReference: g.varRef, YourReference: g.deresRef, Requisition: g.rekv,
          Items: g.linjer.map(l => ({ ProductID: o.prod.get(String(l.varenr)).ID, ItemText: l.navn, NumberOfItems: l.antall, PriceExVat: l.pris, DiscountPercent: l.rabatt || 0, VatTypeID: o.mvaId, AccountID: o.kontoId })) };
        if (g.levAdr) Object.assign(body, { ShippingAddressLine1: g.levAdr, ShippingPostalCode: g.levPnr, ShippingCity: g.levSted });
        const n = await post("/api/biz/orders", body);
        const info = { env: e, id: n.ID, nr: n.OrderNumber, sum: n.TaxExclusiveAmount, tid: new Date().toISOString(), av: S.bruker?.navn || "" };
        await lagreOrdre({ ...g.o, unimicro: info });
        res.sendt.push({ navn: g.o.kundeNavn, nr: n.OrderNumber, sum: n.TaxExclusiveAmount });
      } catch (err) { res.feil.push(`${g.o.kundeNavn}: ${err.message}`); }
    }
    return res;
  }
  /* Henter status for sendte ordre. En ordre regnes som fakturert først når fakturaen den er overført til, er fakturert
     (faktura-status 42002 eller høyere, ikke kladd 42001). Koblingen går via fakturalinjenes OrderItemId. */
  async function hentStatus(ordre) {
    if (!(await klar())) return;
    const e = env(), sendt = ordre.filter(o => o.unimicro && o.unimicro.env === e && o.unimicro.id && (o.status || "ny") !== "fakturert");
    const res = { fakturert: [], kladd: [], ikkeOverfort: 0, feil: [] };
    for (const o of sendt) {
      try {
        const u = await get(`/api/biz/orders/${o.unimicro.id}?expand=Items&select=ID,StatusCode`);
        const linje = (u.Items || [])[0];
        if (!linje || u.StatusCode < 41003) { res.ikkeOverfort++; continue; }
        const fl = await get(`/api/biz/invoiceitems?filter=${q("OrderItemId eq " + linje.ID)}&select=ID,CustomerInvoiceID`);
        const invId = fl[0]?.CustomerInvoiceID; if (!invId) { res.ikkeOverfort++; continue; }
        const inv = await get(`/api/biz/invoices/${invId}?select=ID,InvoiceNumber,StatusCode,TaxExclusiveAmount,InvoiceDate`);
        if (inv.StatusCode >= 42002 && inv.InvoiceNumber) {
          await lagreOrdre({ ...o, status: "fakturert", unimicro: { ...o.unimicro, fakturaId: inv.ID, fakturaNr: inv.InvoiceNumber, fakturaSum: inv.TaxExclusiveAmount, fakturaDato: inv.InvoiceDate } });
          res.fakturert.push(`${o.kundeNavn} (faktura ${inv.InvoiceNumber})`);
        } else res.kladd.push(o.kundeNavn);
      } catch (err) { res.feil.push(`${o.kundeNavn}: ${err.message}`); }
    }
    return res;
  }
  return { loggInn, klar, sendUke, hentStatus, env, envNavn: () => ENV[env()].navn, settEnv: v => { if (ENV[v] && v !== env()) { ss.set("um_env", v); ss.del("um_token"); ss.del("um_company"); } }, innlogget: () => !!token(), selskap: () => ss.get("um_companyName") || "" };
})();
window.Um = Um;
