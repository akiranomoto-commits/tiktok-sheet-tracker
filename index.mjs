import { chromium } from "playwright";
import { google } from "googleapis";

// ====== 設定 ======
const SHEET_ID = process.env.SHEET_ID;     // GitHub Secret
const CONFIG_SHEET = "Config";
const VIEWS_SHEET  = "Views";
const DEBUG_SHEET  = "Debug";
const TAIPEI_TZ = "Asia/Taipei";

// ====== 共通関数 ======
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

function columnLetter(n) {
  let s = "";
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

// ====== Google Sheets ======
async function getSheets() {
  const creds = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON);
  const jwt = new google.auth.JWT(
    creds.client_email,
    null,
    creds.private_key,
    ["https://www.googleapis.com/auth/spreadsheets"]
  );
  await jwt.authorize();
  return google.sheets({ version: "v4", auth: jwt });
}

async function ensureSheetExists(sheets, title) {
  const meta = await sheets.spreadsheets.get({ spreadsheetId: SHEET_ID });
  const exists = meta.data.sheets?.some(s => s.properties?.title === title);
  if (!exists) {
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: SHEET_ID,
      requestBody: { requests: [{ addSheet: { properties: { title } } }] }
    });
  }
}

async function readConfigUrls(sheets) {
  await ensureSheetExists(sheets, CONFIG_SHEET);
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SHEET_ID,
    range: `${CONFIG_SHEET}!A2:A`,
    valueRenderOption: "UNFORMATTED_VALUE"
  });
  const rows = res.data.values || [];
  return rows.map(r => (r[0] || "").toString().trim()).filter(Boolean);
}

async function ensureViewsHeader(sheets, dateStr) {
  await ensureSheetExists(sheets, VIEWS_SHEET);
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SHEET_ID,
    range: `${VIEWS_SHEET}!1:1`
  });
  let header = (res.data.values && res.data.values[0]) || [];
  if (header.length === 0) header = ["URL"];
  let colIndex = header.indexOf(dateStr);
  if (colIndex === -1) {
    header.push(dateStr);
    await sheets.spreadsheets.values.update({
      spreadsheetId: SHEET_ID,
      range: `${VIEWS_SHEET}!1:1`,
      valueInputOption: "RAW",
      requestBody: { values: [header] }
    });
    colIndex = header.length - 1;
  }
  return { header, colIndex };
}

async function readAllRows(sheets) {
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SHEET_ID,
    range: `${VIEWS_SHEET}!A2:Z100000`
  });
  return res.data.values || [];
}

async function upsertViews(sheets, rowsResult, dateColIndex, existingRows) {
  const map = new Map(); // URL -> rowNumber
  existingRows.forEach((r, i) => {
    const url = (r[0] || "").toString().trim();
    if (url) map.set(url, i + 2);
  });

  const updates = [];
  const appends = [];
  const sheetCol = dateColIndex + 1;

  rowsResult.forEach(({ url, value }) => {
    const rowNum = map.get(url);
    if (rowNum) {
      const range = `${VIEWS_SHEET}!${columnLetter(sheetCol)}${rowNum}:${columnLetter(sheetCol)}${rowNum}`;
      updates.push({ range, values: [[value]] });
    } else {
      const row = [];
      row[0] = url;
      for (let i = 1; i < sheetCol - 1; i++) row[i] = "";
      row[sheetCol - 1] = value;
      appends.push(row);
    }
  });

  if (updates.length) {
    await sheets.spreadsheets.values.batchUpdate({
      spreadsheetId: SHEET_ID,
      requestBody: { valueInputOption: "RAW", data: updates }
    });
  }
  if (appends.length) {
    await sheets.spreadsheets.values.append({
      spreadsheetId: SHEET_ID,
      range: `${VIEWS_SHEET}!A:A`,
      valueInputOption: "RAW",
      insertDataOption: "INSERT_ROWS",
      requestBody: { values: appends }
    });
  }
}

// Debugシート： [ISO時刻, URL, 状態/理由]
async function appendDebug(sheets, items) {
  await ensureSheetExists(sheets, DEBUG_SHEET);
  if (!items.length) return;
  await sheets.spreadsheets.values.append({
    spreadsheetId: SHEET_ID,
    range: `${DEBUG_SHEET}!A:A`,
    valueInputOption: "RAW",
    insertDataOption: "INSERT_ROWS",
    requestBody: {
      values: items.map(x => [new Date().toISOString(), x.url, x.reason])
    }
  });
}

// ====== TikTok 取得 ======
async function fetchPlayCount(page, url) {
  const result = { count: null, reason: "" };

  // URL 正規化（m. → www.、lang=en 付与）
  let target = url.replace("m.tiktok.com/", "www.tiktok.com/");
  target += (target.includes("?") ? "&" : "?") + "lang=en";

  try {
    const resp = await page.goto(target, {
      waitUntil: "networkidle",
      timeout: 60000
    });

    // 403/404など
    const status = resp?.status();
    if (status && status >= 400) {
      result.reason = `HTTP ${status}`;
      return result;
    }

    // Cookie同意を片付け（パターン複数）
    const consentSelectors = [
      'button:has-text("Accept all")',
      'button:has-text("I agree")',
      'button:has-text("同意する")',
      'button:has-text("同意")',
      '[data-e2e="cookie-banner-accept-button"]',
    ];
    for (const sel of consentSelectors) {
      const btn = await page.$(sel).catch(() => null);
      if (btn) { await btn.click().catch(() => {}); break; }
    }

    // まず SIGI_STATE を待つ → 無ければ __NEXT_DATA__ を待つ
    const hasSigi = await page.$('#SIGI_STATE');
    if (!hasSigi) {
      await page.waitForTimeout(1500);
    }

    // 取得ロジック
    const data = await page.evaluate(() => {
      const sigi = document.querySelector('#SIGI_STATE');
      if (sigi) {
        try { return { type: 'SIGI_STATE', json: JSON.parse(sigi.textContent) }; } catch (_e) {}
      }
      const next = document.querySelector('#__NEXT_DATA__');
      if (next) {
        try { return { type: '__NEXT_DATA__', json: JSON.parse(next.textContent) }; } catch (_e) {}
      }
      return null;
    });

    if (!data || !data.json) {
      result.reason = 'no SIGI_STATE / __NEXT_DATA__';
      return result;
    }

    // JSONからplayCountを抽出
    const json = data.json;
    const idMatch = target.split('/').filter(Boolean).pop()?.replace(/\?.*$/, "");
    let play = null;

    if (json.ItemModule) {
      const k = (/^\d+$/.test(idMatch) && json.ItemModule[idMatch]) ? idMatch : Object.keys(json.ItemModule)[0];
      play = json.ItemModule[k]?.stats?.playCount ?? null;
    }

    // さらにフォールバック（deep search）
    if (play == null) {
      const stack = [json];
      while (stack.length) {
        const cur = stack.pop();
        if (!cur) continue;
        if (typeof cur === 'object') {
          for (const key of Object.keys(cur)) {
            if (/^play[_]?count(v2)?$/i.test(key)) {
              const v = cur[key];
              const n = normalizeCount(v);
              if (n != null) { play = n; break; }
            }
            const val = cur[key];
            if (val && typeof val === 'object') stack.push(val);
          }
        } else if (Array.isArray(cur)) {
          cur.forEach(x => stack.push(x));
        }
        if (play != null) break;
      }
    }

    if (play == null) {
      result.reason = 'playCount not found';
      return result;
    }

    result.count = normalizeCount(play);
    result.reason = 'OK';
    return result;

  } catch (e) {
    result.reason = 'exception: ' + (e?.message || e);
    return result;
  }
}

// ====== メイン ======
(async () => {
  try {
    const sheets = await getSheets();
    const urls = await readConfigUrls(sheets);
    if (urls.length === 0) {
      console.log("ConfigにURLがありません。A2以降にURLを入れてください。");
      process.exit(0);
    }

    const dateStr = todayInTaipei();
    await ensureSheetExists(sheets, VIEWS_SHEET);
    await sheets.spreadsheets.values.update({
      spreadsheetId: SHEET_ID,
      range: `${VIEWS_SHEET}!A1`,
      valueInputOption: "RAW",
      requestBody: { values: [["URL"]] }
    });
    const { colIndex } = await ensureViewsHeader(sheets, dateStr);
    const existingRows = await readAllRows(sheets);

    // ブラウザ起動（UA/言語を上書き）
    const browser = await chromium.launch({
      headless: true,
      args: ['--disable-blink-features=AutomationControlled']
    });
    const context = await browser.newContext({
      locale: 'en-US',
      userAgent:
        'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
      extraHTTPHeaders: {
        'Accept-Language': 'en-US,en;q=0.9,ja;q=0.8'
      },
      viewport: { width: 1366, height: 900 },
      timezoneId: TAIPEI_TZ
    });
    const page = await context.newPage();

    const results = [];
    const debugs = [];

    for (const url of urls) {
      const { count, reason } = await fetchPlayCount(page, url);
      results.push({ url, value: count ?? 'ERROR' });
      if (count == null) {
        debugs.push({ url, reason }); // Debugシートに理由を残す
        console.log(`❌ ${url} -> ${reason}`);
      } else {
        console.log(`✅ ${url} -> ${count}`);
      }
      await page.waitForTimeout(1200 + Math.floor(Math.random() * 800));
    }

    await browser.close();

    await upsertViews(sheets, results, colIndex, existingRows);
    if (debugs.length) await appendDebug(sheets, debugs);
    console.log("✅ 更新完了:", dateStr);

  } catch (e) {
    console.error('FATAL:', e?.stack || e);
    process.exit(1);
  }
})();
