# CRES Artist Finder

Internal portal + installable app for the **CRES artist master database** (Google Sheet).

- **Artist cards** (dark theme): photo, name, cost, number, 1–2 line bio, manager, email, and portfolio links (profile slide, Instagram, YouTube, website)
- **Photos** taken from each artist's own profile slide; click a photo to enlarge it
- **Search** artists across all categories — name, band, city, act type, remarks, bio
- **Filter** by category, city, budget, status (cost updated / follow-up), phone/link, has photo / no photo
- **Shortlist** artists and copy the list for WhatsApp / mail (with or without contacts)
- **Add artist** from a form — written to the right tab, right below the last entry, with the next Sr no
- **✎ Edit** opens the artist in a form inside the portal; only the fields you change are written to that row (with an **Edit Log** tab of old → new values). **Sheet ↗** opens the row in the sheet
- Phone validation (Indian + international), email check, city name clean-up, duplicate check

Data is always **live from the sheet** — nothing is stored in this repo or in the app.

App link: **https://mis0303.github.io/Cres-Artist-Database/**

---

## How it fits together

```
Google Sheet (Master sheet Artist database)
        ▲  read / add
        │
Apps Script web app  ← apps-script/Code.gs + apps-script/Index.html
        ▲  shown inside
        │
Installable app (GitHub Pages) ← index.html + manifest.json + sw.js + icons/
```

| Path | What it is |
|---|---|
| `apps-script/Code.gs` | Backend: reads the tabs, maps columns by header name, adds artists |
| `apps-script/Index.html` | The portal UI (search, filters, cards, shortlist, add form) |
| `index.html` | App wrapper — opens the Apps Script web app full-screen |
| `manifest.json`, `sw.js`, `icons/` | What makes it installable as an app |

---

## Setup / updating the Apps Script

1. Open the sheet → **Extensions → Apps Script**.
2. `Code.gs` → paste `apps-script/Code.gs`.
   `SHEET_ID` is left empty in this repo on purpose: a script created from inside the sheet uses that sheet automatically. For a standalone script, paste the sheet ID there.
3. HTML file named exactly **`Index`** → paste `apps-script/Index.html` (first line must be `<!DOCTYPE html>`).
4. **Deploy → Manage deployments → ✏️ Edit → Version: New version → Deploy.**
   - Execute as: **Me**
   - Who has access: **Anyone within creativerays.in**

Saving alone does **not** update the live link — always deploy a *New version*.
Using "New version" keeps the same URL, so the app needs no change.

## Profile columns

Each category tab has these columns at the end, filled by the team: `Email`, `Instagram`, `YouTube`, `Website`, `Bio`
(and `Manager contact no & name` on tabs that did not have one). The portal finds them by header name.
Instagram can be a link or an `@handle`.

## Photos

1. In `Code.gs`, set `PHOTO_FOLDER_ID` to the Drive folder for artist photos.
2. Reload the sheet and use the **Artist Finder** menu → **Extract photos from slides**
   (or run `extractPhotos` from the Apps Script editor). Allow the permissions it asks for.
3. It opens each artist's own profile slide, saves the photo in the folder, and keeps going in
   ~5-minute batches by itself. You get an email when it is finished. If the sheet shows
   "Running script", press **Dismiss**, not Cancel.
4. The **Photos** tab has one row per artist:
   - `Status`: `OK`, `NO SLIDE` (no profile slide; add a photo by hand), `NO PHOTO ON SLIDE`,
     `SLIDE NOT FOUND`, `NO ACCESS TO DECK`.
   - `Manual photo`: paste a Drive link (or file ID) to set or replace a photo. A manual photo always wins.
   - YouTube / Instagram / other links found on the slide are kept here too, and fill the card's
     buttons when the artist's own columns are empty.
5. **Update photo list (new artists)** adds new artists; **Retry photos that failed** tries again.

Photos are sent to the portal by the script itself (cached for 6 hours), so they show even where a browser blocks Google's cookies. If a photo can't load, the card shows the initials instead.

## Speed

- The artist list is cached on the server for 10 minutes (`CACHE_MINUTES` in `Code.gs`).
- Any edit in the sheet clears that cache (`onEdit`, works when the script is created from inside the sheet), and so does adding an artist from the portal.
- **↻ Refresh** always reads the sheet fresh.
- Each device also remembers the last list, so the portal shows it instantly and swaps in fresh data a moment later.

## Sheet structure the portal expects

Tabs are listed in `TABS` at the top of `Code.gs`:

| Tab | Category in portal |
|---|---|
| Emcee with Cost and contact | Emcee |
| Live singer and band with cost and  contact | Singer / Band |
| Standup | Standup |
| Motivational Speakers | Motivational Speaker |
| Dance Perfomance act | Dance Act |
| Mix various  Acts for Events | Other Acts |
| Cordinater And Agences | Coordinator / Agency |

"All in one" is skipped (it is a dropdown view of the same data).

- **New tab** → add one line to `TABS`.
- **Renamed tab** → update its line (the portal header shows "⚠ tab not found" otherwise).
- **Columns** are matched by header words (name, city/location, contact/number, manager, cost, link, remark, language, company/band/group/agency, type/category/provider), so columns can be moved freely. A header without any of these words is not shown.

## Installing the app

- **Android / Chrome / Edge:** open the app link → menu → **Install app** (or the install icon in the address bar).
- **iPhone (Safari):** Share → **Add to Home Screen**.
- Users must be signed in to Google with their **creativerays.in** account.
- The app uses the `script.google.com/a/macros/creativerays.in/...` link, so it works even when a phone or browser is signed in to several Google accounts (personal Gmail + work). If the app keeps loading, tap **Open in browser**, sign in once, then reopen the app.

## Changing the app wrapper

- New Apps Script deployment URL → update `APP_URL` in `index.html`.
- Changed `index.html` or icons → bump `VERSION` in `sw.js` so installed apps refresh.
