import { chromium } from "playwright";
import { google } from "googleapis";

// ====== 設定 ======
const SHEET_ID = process.env.SHEET_ID;     // GitHub Secret で渡す
const CONFIG_SHEET = "Config";
const VIEWS_SHEET  = "Views";
const TAIPEI_TZ = "Asia/Taipei";

// ====== 日付（YYYY-MM-DD, Asia/Taipei） ======
function todayInTaipei() {
  const now = new Date();
  const taipei = new Date(now.toLocaleString("en-US", { timeZone: TAIPEI_TZ }));
  const y = taipei.getFullYear();
  const m = String(taipei.getMonth() + 1).padStart(2, "0");
  const d = String(taipei.getDate()).padStart(2, "0");
  return `${y}-${m}-${d}`;
}

// K/Mや"12,345"を数値化
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

// TikTokの再生数を取得（実ブラウザ）
async function fetchPlayCount(page, url) {
  await page.goto(url, { waitUntil: "domcontentloaded", timeout: 45000 });
  const data = await page.evaluate(() => {
    const el = document.querySelector("#SIGI_STATE");
    return el ? JSON.parse(el.textContent) : null;
  });

  let playCount = null;
  if (data?.ItemModule) {
    const idMatch = url.split("/").filter(Boolean).pop()?.replace(/\?.*$/, "");
    const keys = Object.keys(data.ItemModule);
    const key = /^\d+$/.test(idMatch) && data.ItemModule[idMatch] ? idMatch : (keys[0] || null);
    if (key) playCount = data.ItemModule[key]?.stats?.playCount ?? null;
  }
  return normalizeCount(playCount);
}

// Google Sheets クライアント
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

async function readConfigUrls(sheets) {
  // Config!A2:A を取得
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SHEET_ID,
    range: `${CONFIG_SHEET}!A2:A`,
    valueRenderOption: "UNFORMATTED_VALUE"
  });
  const rows = res.data.values || [];
  return rows.map(r => (r[0] || "").toString().trim()).filter(Boolean);
}

async function ensureViewsHeader(sheets, dateStr) {
  // 1行目のヘッダを読み込み
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SHEET_ID,
    range: `${VIEWS_SHEET}!1:1`
  });
  let header = (res.data.values && res.data.values[0]) || [];
  if (header.length === 0) header = ["URL"]; // 初期化

  // 日付列がなければ追加
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
  return { header, colIndex }; // colIndex は0始まり（A=0, B=1, …）
}

async function readAllRows(sheets) {
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SHEET_ID,
    range: `${VIEWS_SHEET}!A2:Z100000` // 十分大きく
  });
  const values = res.data.values || [];
  return values; // 2行目以降
}

async function upsertViews(sheets, urls, dateColIndex, rows) {
  // rows: 2行目以降の配列（先頭列がURL）
  const existingMap = new Map(); // URL -> 行番号（シート上の実際の行番号）
  rows.forEach((r, i) => {
    const url = (r[0] || "").toString().trim();
    if (url) existingMap.set(url, i + 2); // +2: 1行目がヘッダ、配列は0始まり
  });

  // まず既存URLの行を更新、無いURLは新規行を追加
  const updates = [];
  const appends = [];

  // dateColIndex は0始まり。シートの列番号（A=1…）に直す
  const sheetCol = dateColIndex + 1;

  for (const { url, view } of urls) {
    const rowNum = existingMap.get(url);
    if (rowNum) {
      // 既存行の該当日付列だけ更新
      const range = `${VIEWS_SHEET}!${columnLetter(sheetCol)}${rowNum}:${columnLetter(sheetCol)}${rowNum}`;
      updates.push({ range, values: [[view]] });
    } else {
      // 新規行：URLをA列、日付列にview、それより左の空白を埋める
      const row = [];
      // A列 = URL
      row[0] = url;
      // 日付列位置に view を入れる（足りない分は空白を埋める）
      for (let i = 1; i < sheetCol - 1; i++) row[i] = "";
      row[sheetCol - 1] = view;
      appends.push(row);
    }
  }

  // バッチ更新
  if (updates.length) {
    await sheets.spreadsheets.values.batchUpdate({
      spreadsheetId: SHEET_ID,
      requestBody: {
        valueInputOption: "RAW",
        data: updates
      }
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

function columnLetter(n) {
  // 1 -> A, 2 -> B ...
  let s = "";
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

(async () => {
  const sheets = await getSheets();
  const urls = await readConfigUrls(sheets);
  if (urls.length === 0) {
    console.log("ConfigにURLがありません。A2以降にURLを入れてください。");
    process.exit(0);
  }

  const dateStr = todayInTaipei();
  await sheets.spreadsheets.values.update({
    spreadsheetId: SHEET_ID,
    range: `${VIEWS_SHEET}!A1`,
    valueInputOption: "RAW",
    requestBody: { values: [["URL"]] }
  });
  const { colIndex } = await ensureViewsHeader(sheets, dateStr);
  const existingRows = await readAllRows(sheets);

  const browser = await chromium.launch({ headless: true });
  const page = await browser.newPage();

  const results = [];
  for (const url of urls) {
    try {
      const view = await fetchPlayCount(page, url);
      results.push({ url, view: view ?? "ERROR" });
      await new Promise(r => setTimeout(r, 1500));
    } catch (e) {
      results.push({ url, view: "ERROR" });
    }
  }

  await browser.close();

  await upsertViews(sheets, results, colIndex, existingRows);
  console.log("✅ 更新完了:", dateStr);
})();
