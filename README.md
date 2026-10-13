# Actieregister — iPhone-app

Een app om het **actieregister** (het Excel-bestand `Actieregister_Marsdijk.xlsx`) als
overzichtelijke to-do-lijst op je iPhone te gebruiken: afvinken, bewerken, nieuwe acties
toevoegen, en de wijzigingen weer terugzetten in hetzelfde Excel-bestand.

Het is een *webapp* (PWA): je opent hem één keer in Safari en zet hem op je beginscherm.
Daarna werkt hij als een gewone app, ook offline. Er is geen App Store, Mac of
Apple-ontwikkelaarsaccount voor nodig.

## Privacy

Het register bevat vertrouwelijke (personeels)gegevens. Daarom:

- staat er **geen registerdata in deze repository** (alleen `tests/fixture.xlsx` met verzonnen data);
- opent de app het Excel-bestand **op de telefoon zelf** en bewaart het alleen daar (IndexedDB);
- stuurt de app **niets naar internet** — er is geen server en geen account.

## Wat kan de app

| Tabblad | Functie |
|---|---|
| **Acties** | Kaartjes per actie met prioriteit, status, deadline (“3 dagen te laat”, “morgen”) en eigenaar. Snelle weergaven: *Mijn acties*, *Top 10* (kolom Rang), *Actief*, *Achterstallig*, *Deze week*, *Wacht op*, *Te verifiëren*, *Ideeën & plannen*, *Klaar*, *Alles*. Zoeken, filteren (prioriteit, eigenaar, domein, status) en groeperen (prioriteit, deadline, domein, eigenaar, status). Rondje aantikken = afvinken (met *Ongedaan maken*). |
| **Actie openen** | Alle kolommen: Waarom nu, Aanpak, Toelichting, Bron, enz. Status, prioriteit, deadline, eigenaar en teksten zijn direct te bewerken. |
| **+** | Nieuwe actie toevoegen (volgend nummer, besluitdatum vandaag). |
| **Overzicht** | Dezelfde tellingen als het tabblad *Overzicht* in Excel; tik op een getal of balk om die acties te zien. |
| **Bestand** | Je naam kiezen (voor *Mijn acties*), wijzigingen opslaan als Excel, een nieuwere versie inlezen, gegevens wissen. |

*Achterstallig* en *Dagen open* rekent de app zelf uit, net als de formules in Excel.

## Werkwijze

1. **Inlezen** — open in de app *Bestand → Nieuwere versie inlezen* (of bij eerste gebruik
   *Excel-bestand openen*) en kies `Actieregister_Marsdijk.xlsx`. In de iOS-kiezer vind je
   SharePoint/OneDrive onder *Bladeren* (zet zo nodig OneDrive aan via *Wijzig*).
2. **Werken** — vink af en bewerk. Gewijzigde acties krijgen een blauw stipje; het tabblad
   *Bestand* krijgt een rode stip zolang er iets niet in Excel staat.
3. **Opslaan** — *Bestand → Opslaan als Excel*. De app maakt een nieuwe versie van het
   **originele** bestand: alleen de gewijzigde cellen en nieuwe rijen veranderen; opmaak,
   formules, dropdowns, de Excel-tabel en de tabbladen *Overzicht* en *Toelichting* blijven
   behouden. Kies *Delen / bewaren… → Bewaar in Bestanden* en vervang het bestand in SharePoint.
4. **Na een overleg** — lees de nieuwste versie opnieuw in. Heb je nog niet-opgeslagen
   wijzigingen, dan vraagt de app of je die wilt meenemen (gekoppeld op actienummer).

> Tip: bewerk het register niet tegelijk in Excel én in de app. Lees na een Excel-sessie
> eerst de nieuwe versie in.

## Formaat van de Excel

De app herkent kolommen **op de koptekst**, niet op positie. Kolommen mogen dus van volgorde
wisselen en er mogen kolommen bij komen (die verschijnen onder *Meer velden*). Herkend worden:

`Nr` · `Actie` · `Domein` · `Eigenaar` · `Prioriteit` · `Deadline` · `Status` · `Achterstallig` ·
`Dagen open` · `Rang` · `Waarom nu` · `Aanpak — eerste stap` · `Bron` · `Besluitdatum` · `Toelichting`

Afspraken die de app het best laten werken:

- de koptekstrij staat in de eerste 30 rijen; titelregels erboven worden als titel getoond;
- **Deadline** en **Besluitdatum** als echte Excel-datum (vrije tekst als “Q4” mag, maar telt dan
  niet mee voor *Achterstallig*);
- meerdere eigenaren scheiden met `/` (bijv. `Anna / Bert`) — dan verschijnt de actie bij
  *Mijn acties* van beiden;
- de keuzelijsten voor Status/Prioriteit/Domein komen uit de dropdowns (gegevensvalidatie) in
  Excel; nieuwe statussen of domeinen verschijnen dus vanzelf in de app;
- de statussen *Klaar*, *Idee* en *Plan* tellen niet als “actief” (zelfde logica als het tabblad Overzicht).

## Installeren op de iPhone

De app moet ergens via **https** bereikbaar zijn. Twee opties:

**A. GitHub Pages (automatisch)** — deze repository bevat een workflow die de app bij elke
wijziging op `main` publiceert. Zet in GitHub *Settings → Pages → Source* op **GitHub Actions**.
Let op: voor GitHub Pages bij een *privé*-repository is een betaald GitHub-abonnement nodig;
anders kun je de repository openbaar maken (er staat geen registerdata in) of optie B gebruiken.
Het adres wordt `https://yoeri-devmans.github.io/Actieregister-/`.

**B. Elke andere statische host** (bijv. Netlify Drop, Cloudflare Pages, of een webserver van
de praktijk): upload de map met `index.html`, `manifest.webmanifest`, `sw.js`, `css/`, `js/`,
`icons/` en `vendor/`.

Daarna op de iPhone: open het adres in **Safari** → deelknop → **Zet op beginscherm**.

## Ontwikkelen

```bash
npm install
npm test        # maakt tests/fixture.xlsx (python3 + openpyxl) en test inlezen/terugschrijven
npm start       # lokale server op http://localhost:8080
```

| Bestand | Inhoud |
|---|---|
| `js/register.js` | Excel inlezen (kolomherkenning) en terugschrijven in het originele bestand |
| `js/app.js` | De interface |
| `js/store.js` | Lokale opslag op het toestel |
| `sw.js` | Offline werken — verhoog `VERSION` bij een nieuwe versie |
| `vendor/exceljs.min.js` | [ExcelJS](https://github.com/exceljs/exceljs) 4.4.0 (MIT) |
