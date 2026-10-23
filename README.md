# Toddum ordrebok

Med ordreboka legger Toddum Gårdsysteri AS inn ukeordre. Fra ordrene lager den ysteriliste, pakkseddel, CMR og fakturagrunnlag. Brukerne logger inn med Toddum-kontoen i Microsoft 365, og alle data lagres i tre SharePoint-lister i området Toddum (`toddum.sharepoint.com/sites/team`):

- **Ordrebok Produkter**: varenummer er nøkkelen (kolonnen Tittel). Fasiten for varenummer er *Produktregister Toddum.xlsx*.
- **Ordrebok Kunder**: kunde-ID er nøkkelen, vanligvis kundenummeret.
- **Ordrebok Ordre**: én rad per kunde per uke. Varelinjene ligger som JSON i kolonnen `linjer`.

Appen er en ren nettside uten server. Den publiseres til GitHub Pages hver gang det kommer en endring på `main`, til adressen `https://einarby-max.github.io/toddum-ordrebok/`. Kodelageret er offentlig. Det inneholder bare kode og det tomme CMR-skjemaet, ingen kunde- eller ordredata.

## Struktur

| Fil | Innhold |
|---|---|
| `app/index.html` | Oppsett og stil |
| `app/app.js` | Ukeordre, ysteriliste, kunder, produkter og PDF-er (pdf-lib) |
| `app/store.js` | Innlogging (MSAL) og lagring i SharePoint via Microsoft Graph |
| `app/config.js` | Klient-ID, tenant og SharePoint-område |
| `app/assets/cmr.pdf` | Tomt CMR-skjema som fylles ut |
| `app/sw.js`, `app/manifest.webmanifest` | Gjør appen installerbar på telefon |
| `.github/workflows/pages.yml` | Publiserer mappen `app` til GitHub Pages |

## Oppsett (gjøres én gang)

1. **GitHub Pages.** Gå til Settings › Pages og sett Source til «GitHub Actions».
2. **Entra.** Opprett appregistreringen «Toddum ordrebok» som Single-page application:
   - Redirect-URI er `https://einarby-max.github.io/toddum-ordrebok/`.
   - Legg til den delegerte Graph-tillatelsen `Sites.Manage.All` og gi administratorsamtykke.
3. **Konfigurasjon.** Skriv klient-ID-en inn i `app/config.js`.
4. **Første start.** Åpne appen og logg inn. Oppsettsiden oppretter listene og importerer startdata fra en JSON-eksport.
5. **Eget domene (valgfritt).** Opprett en CNAME-oppføring `ordre.toddum.com` som peker til `einarby-max.github.io`, og legg domenet inn under Settings › Pages. Legg til `https://ordre.toddum.com/` som redirect-URI i Entra.

## Bruk på telefon

- **iPhone:** Åpne adressen i Safari, trykk Del og velg «Legg til på Hjem-skjerm».
- **Android:** Åpne adressen i Chrome, trykk menyen og velg «Installer app».

## Regler i appen

- **Batch:** Produksjonsdatoen skrives som DDMMÅÅ og ganges med 2. Eksempel: 17.02.26 gir 340452. Batch kan overstyres per varelinje.
- **Antall:** Oppgis i salgsenheten, det vil si beger, spann, kg eller liter. Skriver man «4dp», blir det fire detaljistpakninger.
- **Faste ordre:** Kan settes til hver uke, partallsuker eller oddetallsuker. De lagres på kunden.

## Sikkerhetskopi

Gå til `https://<adresse>/#oppsett` og velg «Last ned sikkerhetskopi (JSON)». Ingen kundedata ligger i dette kodelageret.
