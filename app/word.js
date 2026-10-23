/* Import av ukelista fra Word (.docx).
   Lista er en tabell med én rad per kunde og kolonnene Kunde, 125g, 10g, Burrata, Stracciatella, Revet og Kommentar.
   Cellene er fritekst, f.eks. «30-250 / 10-500», «2 spann», «1 dp – 250 gram», «2 x 5 kg» eller «12 kg».
   Alt tolkes i nettleseren. Ingenting lagres før brukeren har sett over forslaget og trykket «Legg inn». */
const Word = (() => {
  /* --- les .docx (zip) uten eksterne biblioteker --- */
  async function lesZipFil(buf, navn){
    const v = new DataView(buf), u8 = new Uint8Array(buf);
    let eocd = -1;
    for(let i = buf.byteLength - 22; i >= Math.max(0, buf.byteLength - 66000); i--) if(v.getUint32(i, true) === 0x06054b50){ eocd = i; break; }
    if(eocd < 0) throw new Error("Fila er ikke et Word-dokument (.docx).");
    const antall = v.getUint16(eocd + 10, true); let p = v.getUint32(eocd + 16, true);
    const dek = new TextDecoder();
    for(let n = 0; n < antall; n++){
      const metode = v.getUint16(p + 10, true), csize = v.getUint32(p + 20, true);
      const nlen = v.getUint16(p + 28, true), elen = v.getUint16(p + 30, true), clen = v.getUint16(p + 32, true), lok = v.getUint32(p + 42, true);
      const filnavn = dek.decode(u8.subarray(p + 46, p + 46 + nlen));
      if(filnavn === navn){
        const start = lok + 30 + v.getUint16(lok + 26, true) + v.getUint16(lok + 28, true);
        const data = u8.subarray(start, start + csize);
        if(metode === 0) return dek.decode(data);
        if(metode !== 8) throw new Error("Ukjent komprimering i Word-fila.");
        const strøm = new Blob([data]).stream().pipeThrough(new DecompressionStream("deflate-raw"));
        return await new Response(strøm).text();
      }
      p += 46 + nlen + elen + clen;
    }
    throw new Error("Fant ikke innholdet i Word-fila.");
  }
  async function lesTabell(fil){
    const xml = await lesZipFil(await fil.arrayBuffer(), "word/document.xml");
    const doc = new DOMParser().parseFromString(xml, "application/xml");
    const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    const tbl = doc.getElementsByTagNameNS(W, "tbl")[0];
    if(!tbl) throw new Error("Fant ingen tabell i Word-fila.");
    const tekst = el => Array.from(el.getElementsByTagNameNS(W, "t")).map(t => t.textContent).join("");
    return Array.from(tbl.childNodes).filter(n => n.localName === "tr").map(tr =>
      Array.from(tr.childNodes).filter(n => n.localName === "tc").map(tc =>
        Array.from(tc.getElementsByTagNameNS(W, "p")).map(tekst).map(s => s.trim()).filter(Boolean).join(" / ")));
  }

  /* --- tolking av celler --- */
  const GRUPPER = [["125g", /^125/], ["10g", /^10\s*g|småkul/], ["Burrata", /^burr/], ["Stracciatella", /^strac/], ["Revet", /^rev/]];
  const VARE = {
    "125g":          { stk: "0001", spann: "0002" },
    "10g":           { stk: "0010", spann: "0011" },
    "Burrata":       { stk: "0020", spann: "0022" },
    "Stracciatella": { stk: "0009", g250: "0009", g500: "0008", spann: "0041", kg: "0042" },
    "Revet":         { stk: "0031", g250: "0031", spann: "0030", kg: "0032" },
  };
  const tall = s => parseFloat(String(s).replace(",", "."));
  function vareFor(grp, type, prod){
    const v = VARE[grp] && VARE[grp][type];
    return v && prod(v) ? v : null;
  }
  function etterStørrelse(grp, gram, prod){
    if(grp === "10g" && gram === 150) return vareFor(grp, "stk", prod);
    if(grp === "125g" && gram === 125) return vareFor(grp, "stk", prod);
    if(grp === "Burrata" && gram === 100) return vareFor(grp, "stk", prod);
    return vareFor(grp, "g" + gram, prod);
  }
  function tolkCelle(grp, txt, prod){
    const linjer = {}, ukjent = [];
    let forrige = null;
    const legg = (v, n) => { if(!v || !(n > 0)) return false; linjer[v] = (linjer[v] || 0) + n; forrige = { v, n }; return true; };
    const deler = String(txt).toLowerCase().replace(/[–—]/g, "-").replace(/gram\b/g, "gr").replace(/d-?pak|dpk/g, "dp")
      .split(/\s*\/\s*/).map(s => s.trim()).filter(Boolean);
    for(const s of deler){
      let m;
      if(/^(dp|d)$/.test(s)) continue;                                                 // «dpak» alene
      if((m = s.match(/^(\d+(?:[.,]\d+)?)\s*x\s*(\d+(?:[.,]\d+)?)\s*kg$/))){            // 2 x 5 kg
        if(legg(vareFor(grp, "spann", prod), tall(m[1]))) continue;
      }
      if((m = s.match(/^(\d+)\s*-\s*(\d+)\s*(?:gr?)?\b/))){                             // 30-250, 24- 250 gr 4d
        if(legg(etterStørrelse(grp, +m[2], prod), +m[1])) continue;
      }
      if((m = s.match(/^(\d+)\s*gr?\s*x\s*(\d+)$/))){                                   // 500 gr x1
        if(legg(etterStørrelse(grp, +m[1], prod), +m[2])) continue;
      }
      if((m = s.match(/^(\d+(?:[.,]\d+)?)\s*spann$/))){                                 // 2 spann
        if(legg(vareFor(grp, "spann", prod), tall(m[1]))) continue;
      }
      if((m = s.match(/^(\d+(?:[.,]\d+)?)\s*kg$/))){                                    // 12 kg
        if(legg(vareFor(grp, "kg", prod), tall(m[1]))) continue;
      }
      if((m = s.match(/^(\d+(?:[.,]\d+)?)\s*(?:dp|d)\b\s*(?:-\s*(\d+)\s*(?:gr?)?)?$/))){ // 1 dp, 1 dp - 250 gr, 15 dp
        const v = m[2] ? etterStørrelse(grp, +m[2], prod) : (forrige ? forrige.v : vareFor(grp, "stk", prod));
        const p = v && prod(v); const n = tall(m[1]) * ((p && p.perDP) || 1);
        if(forrige && v === forrige.v && Math.abs(forrige.n - n) < 0.01) continue;      // «90 / 15 dp» – samme antall to ganger
        if(legg(v, n)) continue;
      }
      if((m = s.match(/^(\d+(?:[.,]\d+)?)(?:\s+\d+)?$/))){                              // 30, «30 5» (5 = antall D-pak)
        if(legg(vareFor(grp, "stk", prod), tall(m[1]))) continue;
      }
      ukjent.push(s);
    }
    return { linjer, ukjent };
  }

  /* --- kundeoppslag --- */
  const norm = s => String(s || "").toLowerCase().replace(/[’'`´]/g, "").replace(/\b(as|sa|da|ans|enk)\b/g, "").replace(/[^a-z0-9æøå]+/g, " ").trim();
  function husk(){ try { return JSON.parse(localStorage.getItem("wordKunder") || "{}"); } catch(_) { return {}; } }
  function lagreValg(tekst, id){ try { const h = husk(); h[norm(tekst)] = id; localStorage.setItem("wordKunder", JSON.stringify(h)); } catch(_) {} }
  function finnKunde(tekst, kunder){
    const t = norm(tekst); if(!t) return "";
    const h = husk()[t]; if(h && kunder.some(k => k.id === h)) return h;
    const aktive = kunder.filter(k => !k.inaktiv);
    const ord = t.split(" ").filter(o => o.length > 2);
    const kandidater = aktive.map(k => {
      const n = norm(k.navn), l = norm(k.levNavn);
      let p = 0;
      if(n === t || l === t) p = 100;
      else if(n.startsWith(t) || l.startsWith(t)) p = 80;
      else if((" " + n + " ").includes(" " + t + " ") || (" " + l + " ").includes(" " + t + " ")) p = 70;
      else if(ord.length && ord.every(o => n.includes(o) || l.includes(o))) p = 50;
      else if(t.length >= 4 && (n.includes(t.slice(0, Math.max(4, t.length - 1))))) p = 30;
      return { id: k.id, p, lengde: n.length };
    }).filter(x => x.p > 0).sort((a, b) => b.p - a.p || a.lengde - b.lengde);
    if(!kandidater.length) return "";
    if(kandidater.length > 1 && kandidater[0].p === kandidater[1].p && kandidater[0].p < 80) return "";   // uklart – la Silje velge
    return kandidater[0].id;
  }

  /* --- hele lista --- */
  async function tolk(fil, kunder, prod){
    const rader = await lesTabell(fil);
    if(!rader.length) throw new Error("Tabellen i Word-fila er tom.");
    const hode = rader[0].map(s => s.toLowerCase());
    const kol = {}; let komKol = -1;
    hode.forEach((h, i) => { if(i === 0) return; if(/komm|merkn/.test(h)) { komKol = i; return; } const g = GRUPPER.find(([, re]) => re.test(h)); if(g) kol[i] = g[0]; });
    if(!Object.keys(kol).length) throw new Error("Fant ikke kolonnene 125g, 10g, Burrata, Stracciatella og Revet i tabellen.");
    let uke = null; const ut = [];
    for(const r of rader.slice(1)){
      const navn = (r[0] || "").trim();
      const um = navn.match(/^uke\s*(\d{1,2})$/i); if(um){ uke = +um[1]; continue; }
      if(!navn) continue;
      const linjer = {}, ukjent = [], celler = {};
      for(const [i, grp] of Object.entries(kol)){
        const txt = (r[i] || "").trim(); if(!txt) continue;
        celler[grp] = txt;
        const t = tolkCelle(grp, txt, prod);
        for(const [v, n] of Object.entries(t.linjer)) linjer[v] = (linjer[v] || 0) + n;
        t.ukjent.forEach(u => ukjent.push(`${grp}: ${u}`));
      }
      if(!Object.keys(celler).length) continue;                                       // ingen bestilling denne uka
      ut.push({ tekst: navn, kundeId: finnKunde(navn, kunder), linjer, ukjent, celler, kommentar: komKol >= 0 ? (r[komKol] || "").trim() : "", med: true });
    }
    return { uke, filnavn: fil.name, rader: ut };
  }
  return { tolk, tolkCelle, finnKunde, lagreValg, norm };
})();
