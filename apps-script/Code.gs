/**
 * CRES Artist Finder — search + "Add artist" portal on top of "Master sheet database"
 * ---------------------------------------------------------------------
 * - Every tab has different headers, so columns are matched by header NAME
 *   (not position). Reordering columns in the sheet will not break the portal.
 * - Add artist: writes to the right tab, right below the last entry, with the next Sr no.
 *   No existing row is ever overwritten. Who added it -> note on the Name cell.
 * - Editing still happens directly in the sheet ("Edit ↗" on each card).
 */

// EMPTY   = uses the sheet this script is bound to (Extensions → Apps Script).
// SHEET ID = reads/writes that sheet (needed for a standalone script).
const SHEET_ID = '';   // Paste the sheet ID here, or leave empty if this script is bound to the sheet

function getSS_() {
  return SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
}

// Tab name (exactly as in the sheet)  ->  Category label shown in the portal.
// "All in one" is skipped on purpose (it is a dropdown view that duplicates data).
// New tab in the sheet? Just add one line here.
const TABS = [
  { tab: 'Emcee with Cost and contact',                 cat: 'Emcee' },
  { tab: 'Live singer and band with cost and  contact', cat: 'Singer / Band' },
  { tab: 'Standup',                                     cat: 'Standup' },
  { tab: 'Motivational Speakers',                       cat: 'Motivational Speaker' },
  { tab: 'Dance Perfomance act',                        cat: 'Dance Act' },
  { tab: 'Mix various  Acts for Events',                cat: 'Other Acts' },
  { tab: 'Cordinater And Agences',                      cat: 'Coordinator / Agency' },
];

// Header keywords -> field. First match wins, so the order matters.
const FIELD_RULES = [
  ['manager',  /manager/i],
  ['link',     /link/i],
  ['remarks',  /remark/i],
  ['cost',     /cost/i],
  ['city',     /city|location|loaction/i],
  ['language', /language/i],
  ['company',  /company|band|group|agency/i],
  ['type',     /type|category|provider|act type/i],
  ['phone',    /contant|contact|number|mobile|phone/i],
  ['name',     /name/i],
  ['srno',     /^\s*(sr|ff|s\.?no)/i],
];

/* ============================ WEB APP ============================ */

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('CRES Artist Finder')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** Called by the frontend. force=true skips the cache (Refresh button). */
function getArtists(force) {
  if (!force) {
    const hit = cacheGet_();
    if (hit) { hit.cached = true; return hit; }
  }
  const ss = getSS_();
  const out = [];
  const missing = [];
  TABS.forEach(cfg => {
    const sh = ss.getSheetByName(cfg.tab);
    if (!sh) { missing.push(cfg.tab); return; }
    const values = sh.getDataRange().getDisplayValues();
    const links = readRichLinks_(sh, values);
    out.push.apply(out, normalizeTab_(values, cfg, sh.getSheetId(), links));
  });
  const res = {
    artists: out,
    missingTabs: missing,
    sheetName: ss.getName(),
    sheetUrl: ss.getUrl(),
    loadedAt: new Date().toISOString(),
  };
  cachePut_(res);
  return res;
}

/**
 * Hyperlinks hidden behind cell text (chips / rich text).
 * Reading rich text is slow, so: only the Link column, and only if some
 * link cell has text that isn't already a URL.
 */
function readRichLinks_(sh, values) {
  const map = {};
  try {
    const hm = mapHeader_(values);
    if (!hm || hm.col.link === undefined) return map;
    const c = hm.col.link;
    let needed = false;
    for (let r = hm.h + 1; r < values.length; r++) {
      const v = String(values[r][c] || '').trim();
      if (v && !/^https?:\/\//i.test(v)) { needed = true; break; }
    }
    if (!needed) return map;
    const rt = sh.getRange(1, c + 1, values.length, 1).getRichTextValues();
    rt.forEach((row, r) => {
      const u = row[0] && row[0].getLinkUrl && row[0].getLinkUrl();
      if (u) map[r + ':' + c] = u;
    });
  } catch (e) { /* ignore — plain text links still work */ }
  return map;
}

/* ============================ CACHE ============================ */
// Artist list is kept in the script cache for CACHE_MINUTES so the portal opens fast.
// It is cleared automatically when someone edits the sheet (onEdit below)
// or adds an artist from the portal. The Refresh button always reads fresh.
const CACHE_MINUTES = 10;
const CACHE_KEY = 'artists_v2';
const CHUNK = 90000; // CacheService allows ~100KB per key

function cachePut_(obj) {
  try {
    const s = JSON.stringify(obj);
    const parts = {};
    const n = Math.ceil(s.length / CHUNK);
    for (let i = 0; i < n; i++) parts[CACHE_KEY + '_' + i] = s.slice(i * CHUNK, (i + 1) * CHUNK);
    parts[CACHE_KEY + '_n'] = String(n);
    CacheService.getScriptCache().putAll(parts, CACHE_MINUTES * 60);
  } catch (e) { /* cache is only a speed-up */ }
}

function cacheGet_() {
  try {
    const cache = CacheService.getScriptCache();
    const n = parseInt(cache.get(CACHE_KEY + '_n'), 10);
    if (!n) return null;
    const keys = [];
    for (let i = 0; i < n; i++) keys.push(CACHE_KEY + '_' + i);
    const got = cache.getAll(keys);
    let s = '';
    for (const k of keys) { if (got[k] == null) return null; s += got[k]; }
    return JSON.parse(s);
  } catch (e) { return null; }
}

function clearArtistCache_() {
  try { CacheService.getScriptCache().remove(CACHE_KEY + '_n'); } catch (e) {}
}

/**
 * Simple trigger: any edit in the sheet clears the cache, so the next
 * portal load shows the change. Works when this script is created from
 * inside the sheet (Extensions → Apps Script). For a standalone script the
 * cache simply expires after CACHE_MINUTES.
 */
function onEdit(e) { clearArtistCache_(); }

/* ===================== NORMALISATION (pure JS) ===================== */

/** Find the header row and map each column to a field (by name, not position). */
function mapHeader_(values) {
  let h = -1;
  for (let i = 0; i < Math.min(values.length, 6); i++) {
    if (values[i].some(v => /name|agency/i.test(v))) { h = i; break; }
  }
  if (h < 0) return null;
  const header = values[h];
  const col = {};
  header.forEach((txt, c) => {
    if (!String(txt).trim()) return;
    for (const [field, re] of FIELD_RULES) {
      if (re.test(txt)) {
        if (field === 'phone' && col.phone !== undefined) { col.phone2 = c; break; }
        if (col[field] === undefined) col[field] = c;
        break;
      }
    }
  });
  // First column matched no field (e.g. "Dance Perfomance act") -> it is the Sr no
  if (col.srno === undefined && Object.keys(col).every(k => col[k] !== 0)) col.srno = 0;
  return { h: h, header: header, col: col };
}

function normalizeTab_(values, cfg, gid, richLinks) {
  richLinks = richLinks || {};
  const hm = mapHeader_(values);
  if (!hm) return [];
  const h = hm.h, header = hm.header, col = hm.col;

  const get = (row, f) => (col[f] === undefined ? '' : String(row[col[f]] || '').trim());
  const rows = [];
  for (let r = h + 1; r < values.length; r++) {
    const row = values[r];
    const name = get(row, 'name');
    const company = get(row, 'company');
    if (!name && !company) continue;

    // Anything written in columns without a header -> "extra note"
    const extra = row
      .map((v, c) => (!String(header[c] || '').trim() && String(v).trim() ? String(v).trim() : ''))
      .filter(Boolean).join(' · ');

    let link = get(row, 'link');
    if (!/^https?:/i.test(link)) link = richLinks[r + ':' + col.link] || '';
    // Sometimes the link was pasted into the Name / another column by mistake
    if (!link) {
      const any = row.find(v => /^https?:\/\//i.test(String(v).trim()));
      if (any) link = String(any).trim();
    }

    const phoneRaw = [get(row, 'phone'), get(row, 'phone2')].filter(Boolean).join(' | ');
    const managerRaw = get(row, 'manager');
    const costRaw = get(row, 'cost');
    const remarks = get(row, 'remarks');
    const cost = parseCost_(costRaw);

    rows.push({
      id: gid + '-' + (r + 1),
      cat: cfg.cat,
      tab: cfg.tab,
      gid: gid,
      row: r + 1,
      nameCol: (name ? col.name : col.company) + 1,   // 1-based column holding the name (for the Edit link)
      name: name || company,
      company: name ? company : '',
      city: tidyCity_(get(row, 'city')),
      cities: cityKeys_(get(row, 'city')),
      language: get(row, 'language'),
      type: get(row, 'type'),
      phones: extractPhones_(phoneRaw),
      phoneRaw: phoneRaw,
      managerRaw: managerRaw,
      managerPhones: extractPhones_(managerRaw),
      costRaw: costRaw,
      costMin: cost.min,
      costMax: cost.max,
      link: link,
      linkType: /docs\.google\.com\/presentation/.test(link) ? 'slides'
              : /instagram\.com/.test(link) ? 'instagram'
              : /youtu/.test(link) ? 'youtube' : (link ? 'web' : ''),
      remarks: remarks,
      status: statusOf_(remarks, costRaw),
      extra: extra,
    });
  }
  return rows;
}

/** "20k", "15 to 20 k", "1.5 L", "25/35K", "Mumbai - 45k (4 to 5 hr) Outside - 55k", "2L" */
function parseCost_(s) {
  const res = { min: null, max: null };
  if (!s) return res;
  const t = String(s).toLowerCase().replace(/,/g, '')
    // Drop non-money numbers like "4 to 5 hr", "2 hrs", "6 pieces"
    .replace(/\d+(?:\.\d+)?\s*(?:to|-)\s*\d+(?:\.\d+)?\s*(?:hrs?|hours?|mins?|pieces?|pcs?|days?|members?)/g, ' ');
  const re = /(\d+(?:\.\d+)?)\s*(k|lakhs?|lacs?|l|cr)?\s*(hrs?|hours?|mins?|pieces?|pcs?|days?|sets?|members?|people)?/g;
  const raw = [];
  let m;
  while ((m = re.exec(t)) !== null) {
    if (m[3]) continue;                       // "4 to 5 hr", "6 pieces" -> not a cost
    raw.push([parseFloat(m[1]), m[2] || null]);
  }
  // Missing unit -> borrow the next number's unit ("15 to 20 k" -> 15k-20k)
  let lastUnit = null;
  for (let i = raw.length - 1; i >= 0; i--) {
    if (raw[i][1]) lastUnit = raw[i][1]; else raw[i][1] = lastUnit;
  }
  const nums = [];
  raw.forEach(([n, u]) => {
    let v = n;
    if (u === 'k') v = n * 1000;
    else if (u === 'cr') v = n * 10000000;
    else if (u && /^l/.test(u)) v = n * 100000;
    else if (n < 1000) return;
    nums.push(v);
  });
  if (!nums.length) return res;
  res.min = Math.min.apply(null, nums);
  res.max = Math.max.apply(null, nums);
  return res;
}

/**
 * Pull phone numbers out of text. Returns digits including country code
 *   "98192 69029"      -> "919819269029"
 *   "+44 7911 123456"  -> "447911123456"
 */
function extractPhones_(s) {
  if (!s) return [];
  const found = [];
  const re = /\+?\d[\d\s\-().]{6,18}\d/g;
  let m;
  while ((m = re.exec(s)) !== null) {
    const v = checkPhone_(m[0]);
    if (v.ok && found.indexOf(v.digits) < 0) found.push(v.digits);
  }
  return found;
}

/**
 * Validate one number.
 *  India   : 10 digits starting with 6/7/8/9. "+91", "91" or "0" prefix allowed.
 *  Abroad  : must start with "+" and country code (e.g. +971 50 123 4567), 8–15 digits.
 *  Rejected: wrong length, letters, dummy numbers like 1111111111 / 1234567890.
 * Returns { ok, digits (with country code), pretty (as written to the sheet), err }
 */
function checkPhone_(raw) {
  const s = String(raw || '').trim();
  if (!s) return { ok: false, err: 'Number is empty' };
  if (/[a-z]/i.test(s)) return { ok: false, err: 'Number cannot contain letters' };
  const plus = /^\s*(\+|00)/.test(s);
  let d = s.replace(/\D/g, '');
  if (/^00/.test(s.trim())) d = d.slice(2);
  const dummy = x => /^(\d)\1+$/.test(x) || ['0123456789', '1234567890', '9876543210', '0987654321'].indexOf(x) >= 0;

  // Indian
  let ind = null;
  if (plus && /^91\d{10}$/.test(d)) ind = d.slice(2);
  else if (!plus && /^\d{10}$/.test(d)) ind = d;
  else if (!plus && /^0\d{10}$/.test(d)) ind = d.slice(1);
  else if (!plus && /^91\d{10}$/.test(d)) ind = d.slice(2);
  if (ind !== null) {
    if (!/^[6-9]/.test(ind)) return { ok: false, err: 'Indian mobile numbers start with 6, 7, 8 or 9' };
    if (dummy(ind)) return { ok: false, err: 'This looks like a dummy number' };
    return { ok: true, digits: '91' + ind, pretty: ind.slice(0, 5) + ' ' + ind.slice(5) };
  }
  if (plus && /^91/.test(d)) return { ok: false, err: '+91 must be followed by exactly 10 digits' };
  // International
  if (plus) {
    if (d.length < 8 || d.length > 15) return { ok: false, err: 'International numbers need 8–15 digits including country code' };
    if (/^(\d)\1+$/.test(d.slice(-8))) return { ok: false, err: 'This looks like a dummy number' };
    return { ok: true, digits: d, pretty: '+' + d };
  }
  return { ok: false, err: 'Indian numbers have 10 digits. For a number outside India, start with + and the country code (e.g. +971…)' };
}

// Spelling / area variants -> main city (used by the filter). Add new variants here.
const CITY_ALIASES = {
  mumbai: /^(mum|mumbai|mumabi|bombay|andheri|bandra|malad|mulund|dahisar|kandivali|kandiwali|borivali|goregaon|mira road|bhayand[ae]r|navi mumbai|thane|vashi)/,
  Pune: /^pune/, Delhi: /^(new delhi|delhi)/, Bangalore: /^(bangalore|banglore|bengaluru)/,
  Ahmedabad: /^(ahmedabad|ahemdabad)/, Vadodara: /^(baroda|vadodara|vadodra)/,
  Hyderabad: /^hyderabad/, Kolkata: /^(kolkata|calcutta)/, Chennai: /^chennai/, Goa: /^goa/,
};
function tidyCity_(c) {
  if (!c) return '';
  return c.replace(/\s*[\\]\s*/g, ' / ').replace(/mumabi/i, 'Mumbai').replace(/\s+/g, ' ').trim();
}
/**
 * Clean a city typed in the form to match the sheet's style:
 *  "  banglore " -> "Bangalore",  "MUMBAI" -> "Mumbai",  "malad west" -> "Malad West",
 *  "mumbai/pune" -> "Mumbai / Pune".  A new city is kept as typed, in Title Case.
 */
function cleanCity_(c) {
  c = String(c || '').replace(/\s+/g, ' ').trim();
  if (!c) return '';
  return c.split(/\s*[\/\\,|]\s*/).filter(Boolean).map(part => {
    const p = part.toLowerCase();
    for (const k in CITY_ALIASES) {
      // Only replace when the whole name is an alias ("banglore" yes, "malad west" no)
      const full = new RegExp(CITY_ALIASES[k].source + '$');
      if (full.test(p) && !/^(andheri|bandra|malad|mulund|dahisar|kandivali|kandiwali|borivali|goregaon|mira road|bhayand[ae]r|navi mumbai|thane|vashi)$/.test(p)) {
        return k === 'mumbai' ? 'Mumbai' : (p === 'new delhi' ? 'New Delhi' : k);
      }
    }
    return p.replace(/\b[a-z]/g, x => x.toUpperCase()).replace(/\bNcr\b/, 'NCR');
  }).join(' / ');
}

/** "Mum - pune" -> ["Mumbai","Pune"]; "Malad west" -> ["Mumbai"] */
function cityKeys_(c) {
  if (!c) return [];
  const keys = [];
  c.split(/\s*(?:\/|,|\||-|&|\band\b)\s*/i).forEach(part => {
    const p = part.trim().toLowerCase();
    if (!p || /^(english|hindi|marathi|gujarati)$/.test(p)) return;
    let hit = null;
    for (const k in CITY_ALIASES) if (CITY_ALIASES[k].test(p)) { hit = k === 'mumbai' ? 'Mumbai' : k; break; }
    if (!hit) hit = p.replace(/\b\w/g, x => x.toUpperCase()).replace(/ Ncr$/, ' NCR');
    if (keys.indexOf(hit) < 0) keys.push(hit);
  });
  return keys;
}

/** Status from remarks: verified / followup / unverified (cost given, no date) / nocost */
function statusOf_(remarks, costRaw) {
  const r = (remarks || '').toLowerCase();
  if (/need|to check|recheck|re-check|call|connect|check with|get contact/.test(r)) return 'followup';
  if (/updated|2026/.test(r)) return 'verified';
  return costRaw ? 'unverified' : 'nocost';
}


/* ======================= ADD ARTIST (frontend -> sheet) ======================= */

// Input type for each form field
const FIELD_KIND = { link: 'url', cost: 'cost', remarks: 'textarea', phone: 'phone', phone2: 'phone', manager: 'text' };

/** Form per category: that tab's real headers + suggestions. */
function getFormSchema() {
  const ss = getSS_();
  const all = TABS.map(cfg => {
    const sh = ss.getSheetByName(cfg.tab);
    return sh ? { cfg: cfg, values: sh.getDataRange().getDisplayValues() } : null;
  }).filter(Boolean);
  const cities = allCities_(all.map(x => x.values));
  return all.map(x => {
    const s = buildSchema_(x.values, x.cfg);
    if (s) s.fields.forEach(f => { if (f.key === 'city') f.options = cities; });
    return s;
  }).filter(Boolean);
}

/** Cities from all tabs in one list (cleaned names, most used first) */
function allCities_(tabsValues) {
  const count = {};
  tabsValues.forEach(values => {
    const hm = mapHeader_(values);
    if (!hm || hm.col.city === undefined) return;
    for (let r = hm.h + 1; r < values.length; r++) {
      const v = cleanCity_(values[r][hm.col.city]);
      if (v && !/^(english|hindi|marathi|gujarati)/i.test(v)) count[v] = (count[v] || 0) + 1;
    }
  });
  return Object.keys(count).sort((a, b) => count[b] - count[a] || a.localeCompare(b));
}

function buildSchema_(values, cfg) {
  const hm = mapHeader_(values);
  if (!hm) return null;
  const fields = [];
  Object.keys(hm.col).forEach(key => {
    if (key === 'srno') return;
    const c = hm.col[key];
    // Existing values -> suggestions (useful for city / type / language)
    let options = [];
    if (key === 'city' || key === 'type' || key === 'language') {
      const seen = {};
      for (let r = hm.h + 1; r < values.length; r++) {
        const v = String(values[r][c] || '').trim();
        if (v && !/^https?:/i.test(v)) seen[v] = (seen[v] || 0) + 1;
      }
      options = Object.keys(seen).sort((a, b) => seen[b] - seen[a]).slice(0, 40);
    }
    fields.push({ key: key, col: c, label: String(hm.header[c]).trim(), kind: FIELD_KIND[key] || 'text', options: options });
  });
  fields.sort((a, b) => a.col - b.col);
  return { cat: cfg.cat, tab: cfg.tab, fields: fields };
}

/**
 * Called from the frontend: addArtist({cat:'Emcee', data:{name:'..', city:'..', ...}, force:false})
 * - Writes to the right tab and the right columns
 * - Right below the last entry, with the next Sr no
 * - Warns if the same name already exists (force=true adds anyway)
 */
function addArtist(payload) {
  const cfg = TABS.filter(t => t.cat === payload.cat)[0];
  if (!cfg) throw new Error('Category not found: ' + payload.cat);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000); // prevents overwrites when two people add at the same time
  try {
    const sh = getSS_().getSheetByName(cfg.tab);
    if (!sh) throw new Error('Tab not found: ' + cfg.tab);
    const values = sh.getDataRange().getDisplayValues();
    const plan = planInsert_(values, payload.data || {});
    if (plan.error) throw new Error(plan.error);
    if (plan.duplicateRow && !payload.force) {
      return { duplicate: true, row: plan.duplicateRow, tab: cfg.tab, gid: sh.getSheetId(), nameCol: plan.nameCol + 1 };
    }

    const width = Math.max(sh.getLastColumn(), plan.rowValues.length);
    // If the target row is not empty, insert a new row (never overwrite data)
    if (plan.reuseRow) {
      // empty row / row with only a Sr no — write straight into it
    } else if (plan.targetRow <= sh.getMaxRows()) {
      const existing = sh.getRange(plan.targetRow, 1, 1, width).getDisplayValues()[0];
      if (existing.some(v => String(v).trim())) sh.insertRowAfter(plan.targetRow - 1);
    } else {
      sh.insertRowAfter(sh.getMaxRows());
    }
    // Copy formatting from the row above (font, borders, colours)
    if (plan.targetRow > 2) {
      sh.getRange(plan.targetRow - 1, 1, 1, width)
        .copyTo(sh.getRange(plan.targetRow, 1, 1, width), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
    }
    // Keep phone cells as plain text, otherwise they can show as "9.8E+09"
    plan.textCols.forEach(c => sh.getRange(plan.targetRow, c + 1).setNumberFormat('@'));
    sh.getRange(plan.targetRow, 1, 1, plan.rowValues.length).setValues([plan.rowValues]);

    // Who added it — as a note on the Name cell, no extra column
    let who = '';
    try { who = Session.getActiveUser().getEmail(); } catch (e) {}
    const when = Utilities.formatDate(new Date(), 'Asia/Kolkata', 'dd/MM/yyyy HH:mm');
    sh.getRange(plan.targetRow, plan.nameCol + 1).setNote('Added via Artist Finder' + (who ? ' by ' + who : '') + ' on ' + when);

    clearArtistCache_();
    return { ok: true, tab: cfg.tab, row: plan.targetRow, gid: sh.getSheetId(), srno: plan.srno, nameCol: plan.nameCol + 1 };
  } finally {
    lock.releaseLock();
  }
}

/** Pure logic (testable): which row, which Sr no, which values. */
function planInsert_(values, data) {
  const hm = mapHeader_(values);
  if (!hm) return { error: 'No header row found in this tab' };
  const col = hm.col;
  const clean = {};
  Object.keys(data).forEach(k => { const v = String(data[k] == null ? '' : data[k]).trim(); if (v) clean[k] = v; });
  if (!clean.name && !clean.company) return { error: 'Name (or Company/Band) is required' };
  if (clean.link && !/^https?:\/\//i.test(clean.link)) clean.link = 'https://' + clean.link;

  // Phone: a wrong number must never reach the sheet
  for (const k of ['phone', 'phone2']) {
    if (!clean[k]) continue;
    const v = checkPhone_(clean[k]);
    if (!v.ok) return { error: (k === 'phone2' ? 'Second number: ' : 'Contact number: ') + v.err, field: k };
    clean[k] = v.pretty;
  }
  // Manager field holds name + number — the number in it must be valid too
  if (clean.manager) {
    const nums = clean.manager.match(/\+?\d[\d\s\-().]*\d/g) || [];
    for (const n of nums) {
      if (n.replace(/\D/g, '').length < 4) continue;   // ignore text like "2 managers"
      const v = checkPhone_(n);
      if (!v.ok) return { error: 'Manager number: ' + v.err, field: 'manager' };
      clean.manager = clean.manager.replace(n, v.pretty);
    }
  }
  if (clean.city) clean.city = cleanCity_(clean.city);

  // Last entry = last row that has a name / company
  let last = hm.h, maxSr = 0, dup = 0;
  const target = (clean.name || clean.company).toLowerCase().replace(/\s+/g, ' ');
  for (let r = hm.h + 1; r < values.length; r++) {
    const nm = col.name !== undefined ? String(values[r][col.name] || '').trim() : '';
    const co = col.company !== undefined ? String(values[r][col.company] || '').trim() : '';
    if (nm || co) last = r;
    if (col.srno !== undefined) {
      const n = parseInt(String(values[r][col.srno]).replace(/\D/g, ''), 10);
      if (!isNaN(n) && n > maxSr && n < 100000) maxSr = n;
    }
    [nm, co].forEach(v => { if (v && v.toLowerCase().replace(/\s+/g, ' ') === target && !dup) dup = r + 1; });
  }

  const width = Math.max(hm.header.length, 1);
  const rowValues = new Array(width).fill('');
  const textCols = [];
  Object.keys(clean).forEach(k => {
    if (col[k] === undefined || k === 'srno') return;
    rowValues[col[k]] = clean[k];
    if (k === 'phone' || k === 'phone2' || k === 'manager') textCols.push(col[k]);
  });
  // Tab has no separate Name column (only company) -> put the name there
  if (clean.name && col.name === undefined && col.company !== undefined && !clean.company) rowValues[col.company] = clean.name;
  let srno = col.srno !== undefined ? maxSr + 1 : null;
  // Next row has only a pre-filled Sr no (rest empty) -> reuse that row and its Sr no
  let reuseRow = false;
  const next = values[last + 1];
  if (next) {
    const others = next.filter((v, c) => c !== col.srno && String(v).trim());
    const preSr = col.srno !== undefined ? parseInt(String(next[col.srno]).replace(/\D/g, ''), 10) : NaN;
    if (!others.length) { reuseRow = true; if (!isNaN(preSr)) srno = preSr; }
  }
  if (srno !== null) rowValues[col.srno] = srno;

  return {
    reuseRow: reuseRow,
    targetRow: last + 2,                 // 1-based: right below the last data row
    rowValues: rowValues,
    srno: srno,
    textCols: textCols,
    nameCol: col.name !== undefined ? col.name : col.company,
    duplicateRow: dup,
  };
}

// For local Node tests (ignored by Apps Script)
if (typeof module !== 'undefined') module.exports = { checkPhone_, cleanCity_, allCities_, normalizeTab_, parseCost_, extractPhones_, TABS, buildSchema_, planInsert_, mapHeader_ };
