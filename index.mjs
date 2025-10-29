import { chromium, webkit, firefox } from "playwright";
import { google } from "googleapis";

// ====== 設定 ======
const SHEET_ID = process.env.SHEET_ID;
const CONFIG_SHEET = "Config";
const VIEWS_SHEET  = "Views";
const DEBUG_SHEET  = "Debug";
const TAIPEI_TZ = "Asia/Taipei";

// ====== 共通 ======
function todayInTaipei() {
  const now = new Date();
  const taipei = new Date(now.toLocaleString("en-US", { timeZone: TAIPEI_TZ }));
  const y = taipei.getFullYear();
  const m = String(taipei.getMonth() + 1).padStart(2, "0");
  const d = String(taipei.getDate()).padStart(2, "0");
  return `${y}-${m}-${d}`;
}
function normalizeCount(v) {
  if (typeof v === "number") return v;
  if (typeof v === "string") {
    const u = v.trim().toUpperCase().replace(/,/g, "");
    const m = u.match(/^(\d+(?:\.\d+)?)([KM])$/);
    if (m) return Math.round(parseFloat(m[1]) * (m[2] === "K" ? 1e3 : 1e6));
    if (/^\d+$/.test(u)) return parseInt(u, 10);
  }
  return null;
}
function colLetter(n){let s="";while(n>0){const m=(n-1)%26;s=String.fromCharCode(65+m)+s;n=Math.floor((n-1)/26);}return s;}

// ====== Sheets ======
async function sheetsClient() {
  const creds = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON);
  const jwt = new google.auth.JWT(
    creds.client_email, null, creds.private_key,
    ["https://www.googleapis.com/auth/spreadsheets"]
  );
  await jwt.authorize();
  return google.sheets({ version: "v4", auth: jwt });
}
async function ensureSheet(sheets, title){
  const meta = await sheets.spreadsheets.get({ spreadsheetId: SHEET_ID });
  const ok = meta.data.sheets?.some(s => s.properties?.title === title);
  if (!ok){
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: SHEET_ID,
      requestBody: { requests: [{ addSheet: { properties: { title } } }] }
    });
  }
}
async function readUrls(sheets){
  await ensureSheet(sheets, CONFIG_SHEET);
  const r = await sheets.spreadsheets.values.get({
    spreadsheetId: SHEET_ID, range: `${CONFIG_SHEET}!A2:A`
  });
  const rows = r.data.values || [];
  return rows.map(x => (x[0]||"").toString().trim()).filter(Boolean);
}
async function ensureHeader(sheets, dateStr){
  await ensureSheet(sheets, VIEWS_SHEET);
  const r = await sheets.spreadsheets.values.get({
    spreadsheetId: SHEET_ID, range: `${VIEWS_SHEET}!1:1`
  });
  let header = (r.data.values && r.data.values[0]) || [];
  if (!header.length) header = ["URL"];
  let idx = header.indexOf(dateStr);
  if (idx === -1){
    header.push(dateStr);
    await sheets.spreadsheets.values.update({
      spreadsheetId: SHEET_ID, range: `${VIEWS_SHEET}!1:1`,
      valueInputOption: "RAW", requestBody: { values: [header] }
    });
    idx = header.length - 1;
  }
  return idx;
}
async function readRows(sheets){
  const r = await sheets.spreadsheets.values.get({
    spreadsheetId: SHEET_ID, range: `${VIEWS_SHEET}!A2:Z100000`
  });
  return r.data.values || [];
}
async function upsert(sheets, results, dateColIndex, existing){
  const map = new Map();
  existing.forEach((r,i)=>{ const u=(r[0]||"").toString().trim(); if(u) map.set(u, i+2); });
  const updates=[], appends=[]; const C = dateColIndex+1;
  for (const {url, value} of results){
    const row = map.get(url);
    if (row){
      const range = `${VIEWS_SHEET}!${colLetter(C)}${row}:${colLetter(C)}${row}`;
      updates.push({ range, values: [[value]] });
    } else {
      const arr = []; arr[0]=url; for(let i=1;i<C-1;i++) arr[i]=""; arr[C-1]=value;
      appends.push(arr);
    }
  }
  if (updates.length){
    await sheets.spreadsheets.values.batchUpdate({
      spreadsheetId: SHEET_ID, requestBody: { valueInputOption:"RAW", data: updates }
    });
  }
  if (appends.length){
    await sheets.spreadsheets.values.append({
      spreadsheetId: SHEET_ID, range: `${VIEWS_SHEET}!A:A`,
      valueInputOption:"RAW", insertDataOption:"INSERT_ROWS",
      requestBody: { values: appends }
    });
  }
}
async function appendDebug(sheets, rows){
  await ensureSheet(sheets, DEBUG_SHEET);
  if (!rows.length) return;
  await sheets.spreadsheets.values.append({
    spreadsheetId: SHEET_ID, range: `${DEBUG_SHEET}!A:A`,
    valueInputOption:"RAW", insertDataOption:"INSERT_ROWS",
    requestBody: { values: rows.map(x=>[new Date().toISOString(), x.url, x.engine, x.status||"", x.htmlLen||"", x.reason]) }
  });
}

// ====== TikTok 取得 ======
async function tryOneEngine(engineName, browserType, url){
  // エンジンごとに適切な起動オプションに分岐（ChromiumだけAutomationControlledを無効化）
  const launchArgs = [];
  if (engineName === "chromium") {
    launchArgs.push("--disable-blink-features=AutomationControlled");
    launchArgs.push("--no-sandbox","--disable-dev-shm-usage");
  }
  const browser = await browserType.launch({ headless: true, args: launchArgs });
  const context = await browser.newContext({
    locale: 'en-US',
    userAgent: 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0 Safari/537.36',
    viewport: { width: 1366, height: 900 },
    timezoneId: TAIPEI_TZ,
    extraHTTPHeaders: { 'Accept-Language': 'en-US,en;q=0.9,ja;q=0.8' }
  });
  const page = await context.newPage();
  let status = 0, htmlLen = 0, reason = "";

  // 正規化
  let target = url.replace("m.tiktok.com/", "www.tiktok.com/");
  target += (target.includes("?") ? "&" : "?") + "lang=en";

  try {
    const resp = await page.goto(target, { waitUntil: "domcontentloaded", timeout: 60000 });
    status = resp?.status() || 0;

    // Cookie同意押し（いくつかのラベルに対応）
    const selectors = [
      'button:has-text("Accept all")',
      'button:has-text("I agree")',
      'button:has-text("同意する")',
      '[data-e2e="cookie-banner-accept-button"]'
    ];
    for (const s of selectors){ const b = await page.$(s).catch(()=>null); if (b){ await b.click().catch(()=>{}); break; } }

    // 少し待ってスクロール
    await page.waitForTimeout(1500);
    await page.mouse.wheel(0, 1200);
    await page.waitForTimeout(800);

    // JSON取得
    const data = await page.evaluate(() => {
      const parse = (id)=>{ const el=document.querySelector(id); if(!el) return null; try{ return JSON.parse(el.textContent);}catch(_){return null;} };
      return { sigi: parse('#SIGI_STATE'), next: parse('#__NEXT_DATA__') };
    });

    if (!data.sigi && !data.next){
      const html = await page.content(); htmlLen = html.length;
      reason = "no SIGI_STATE / __NEXT_DATA__";
      return { ok:false, engine: engineName, status, htmlLen, reason };
    }

    const json = data.sigi || data.next;
    const id = target.split('/').filter(Boolean).pop()?.replace(/\?.*$/,"");
    let play = null;
    if (json?.ItemModule){
      const k = (/^\d+$/.test(id) && json.ItemModule[id]) ? id : Object.keys(json.ItemModule)[0];
      play = json.ItemModule[k]?.stats?.playCount ?? null;
    }
    if (play==null){
      const stack=[json];
      while(stack.length){
        const cur=stack.pop();
        if (!cur) continue;
        if (Array.isArray(cur)) { cur.forEach(x=>stack.push(x)); continue; }
        if (typeof cur==='object'){
          for (const key of Object.keys(cur)){
            if (/^play[_]?count(v2)?$/i.test(key)){
              const n = normalizeCount(cur[key]);
              if (n!=null){ play=n; break; }
            }
            const v=cur[key]; if (v && typeof v==='object') stack.push(v);
          }
          if (play!=null) break;
        }
      }
    }
    if (play==null){
      reason = "playCount not found";
      return { ok:false, engine: engineName, status, htmlLen: 0, reason };
    }
    return { ok:true, engine: engineName, count: normalizeCount(play), status };

  } catch(e){
    reason = "exception: " + (e?.message || e);
    return { ok:false, engine: engineName, status, htmlLen, reason };
  } finally {
    await context.close(); await browser.close();
  }
}

async function fetchPlayCountMulti(url){
  const engines = [
    ["chromium", chromium],
    ["webkit",   webkit],
    ["firefox",  firefox],
  ];
  for (const [name, type] of engines){
    const r = await tryOneEngine(name, type, url);
    if (r.ok) return r;
    if (name !== "firefox") await new Promise(r => setTimeout(r, 1200));
  }
  return { ok:false, engine:"all", status:0, htmlLen:0, reason:"all engines failed" };
}

// ====== Main ======
(async () => {
  try {
    const sheets = await sheetsClient();
    const urls = await readUrls(sheets);
    if (!urls.length){ console.log("Config!A2:A にURLを入れてください"); return; }

    const dateStr = todayInTaipei();
    await ensureSheet(sheets, VIEWS_SHEET);
    await sheets.spreadsheets.values.update({
      spreadsheetId: SHEET_ID, range: `${VIEWS_SHEET}!A1`,
      valueInputOption:"RAW", requestBody:{ values:[["URL"]] }
    });
    const colIndex = await ensureHeader(sheets, dateStr);
    const existing = await readRows(sheets);

    const results = [], debugs = [];
    for (const u of urls){
      const r = await fetchPlayCountMulti(u);
      if (r.ok){
        results.push({ url:u, value:r.count });
        console.log(`✅ ${u} -> ${r.count} (${r.engine})`);
      } else {
        results.push({ url:u, value:"ERROR" });
        debugs.push({ url:u, engine:r.engine, status:r.status, htmlLen:r.htmlLen, reason:r.reason });
        console.log(`❌ ${u} -> ${r.reason} (${r.engine})`);
      }
      await new Promise(res=>setTimeout(res, 1200));
    }

    await upsert(sheets, results, colIndex, existing);
    if (debugs.length) await appendDebug(sheets, debugs);
    console.log("✅ 完了:", dateStr);

  } catch(e){
    console.error("FATAL:", e?.stack || e);
    process.exit(1);
  }
})();
