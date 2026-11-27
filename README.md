# CRES Artist Finder

Internal portal + installable app for the **CRES artist master database** (Google Sheet).

- **Search** artists across all categories — name, band, city, act type, remarks
- **Filter** by category, city, budget, status (cost updated / follow-up), phone / link / photo / video
- **Shortlist** artists and copy the list for WhatsApp / mail (with or without contacts)
- **Add artist** from a form — written to the right tab, right below the last entry, with the next Sr no
- **✎ Edit** an artist right in the portal (only changed fields are written; every edit is noted on the cell and in an `Edit Log` tab) — **Sheet ↗** opens the exact cell
- **Photos** on every card, taken from each artist's own profile slide (click to zoom)
- **Videos** at the bottom of the card: ★ Best work plays in a popup inside the portal, plus *More videos* and YouTube channel; "No YouTube videos in sheet" when there are none
- Phone validation (Indian + international), city name clean-up, duplicate check

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

## Photos and videos

Set `PHOTO_FOLDER_ID` in `Code.gs` (a Drive folder for the photos), then use the sheet menu **🎤 Artist Finder**:

| Menu item | What it does |
|---|---|
| Extract photos from slides | Saves the main photo from each artist's slide into the folder and lists it in the `Photos` tab |
| Update photo list | Adds new artists to the `Photos` tab |
| Retry failed photos | Tries the failed ones again |
| Get videos from slides | Reads the YouTube links on each slide (link text = video title) and checks if each video can play inside the portal |
| Re-check all videos | Checks again for removed / blocked videos |

In the `Photos` tab:
- **Manual photo** — paste a Drive link to use your own photo.
- **Best work** — paste one video link to make it the ★ Best work. Empty = first video.
- A YouTube *video* link in the sheet's YouTube column shows as **Showreel**; a *channel* link shows as **YouTube channel**.

Videos whose owner has blocked embedding open on YouTube instead of the popup. Video titles and checks use YouTube's free oEmbed — no API key, no cost.

Long runs continue by themselves (time trigger) and send an email when done.

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
