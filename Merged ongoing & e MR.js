// merged_all_reports.js
// Usage: node merged_all_reports.js
// Before running: npm install puppeteer exceljs

const puppeteer = require("puppeteer");
const fs = require("fs");
const path = require("path");
const ExcelJS = require("exceljs");

// ------------ Configs ------------
const downloadPath = "H:\\scrapping";
const mappingFile = "H:\\Vibhor\\WORK CATEGORY.xlsx";
const subEngineerMapFile = "H:\\Vibhor\\PANCHAYAT MAPING.xlsx"; // <-- Sub-Engineer mapping source

// URLs
const ONGOING_URL = "https://nreganarep.nic.in/netnrega/dynamic_work_details.aspx?page=S&lflag=eng&state_name=MADHYA%20PRADESH&state_code=17&fin_year=2024-2025&source=national&Digest=lCBUfC2x7DqcoUCfIB4vPg";
const MUSTER_URL = "https://mnregaweb4.nic.in/netnrega/state_html/emuster_wage_rep1.aspx?typ=y&type=9b&exe_agency_code=3&month_name=All&lflag=eng&state_name=MADHYA%20PRADESH&state_code=17&district_name=KHARGONE&block_name=ZIRANYA&block_code=1724009&fin_year=2025-2026&Digest=7Tx14AU7o/B/eJif0694Bw";

// ------------ Headers & widths ------------
const TARGET_HEADERS = [
  "S. No.","Panchayat Name","Work Start Fin Year","Work Status","Work Code","Work Name","Agency Name",
  "Sanction Amount (in Rs) Wages","Sanction Amount (in Rs) Material","Sanction Amount (in Rs)",
  "Amount paid in current Fin Year (in Rs) Wages","Amount paid in current Fin Year (in Rs) Material",
  "Amount booked since inception (in Rs) Wages","Amount booked since inception (in Rs) Material",
  "Amount paid since inception (in Rs) Wages","Amount paid since inception (in Rs) Material",
  "Total Mandays","Mandays generated in current Fin Year","Work Started Date","Individual Beneficiary JCN (In case of Cat IV Works)","Work Category"
];

const ONGOING_HEADERS = [
  "S. No.","Panchayat Name","Work Start Fin Year","Work Status","Work Code","Work Name","Agency Name",
  "Sanction Amount (in Rs) Wages","Sanction Amount (in Rs) Material","Sanction Amount (in Rs)",
  "Amount paid in current Fin Year (in Rs) Wages","Amount paid in current Fin Year (in Rs) Material",
  "Total Amount paid in current Fin Year (in Rs)","Amount booked since inception (in Rs) Wages","Amount booked since inception (in Rs) Material",
  "Total Amount booked since inception (in Rs)","Amount paid since inception (in Rs) Wages","Amount paid since inception (in Rs) Material",
  "Total Amount paid since inception (in Rs)","% of Total Amount paid","Total Mandays","Balnce Mandays","Mandays generated in current Fin Year",
  "Work Started Date","Individual Beneficiary JCN (In case of Cat IV Works)","Work Category","Muster Issued Status"
];

const COLUMN_WIDTHS = {
  "S. No.": 8, "Panchayat Name": 20, "Work Start Fin Year": 12, "Work Status": 15, "Work Code": 30,
  "Work Name": 55, "Agency Name": 20, "Sanction Amount (in Rs) Wages": 12, "Sanction Amount (in Rs) Material": 12,
  "Sanction Amount (in Rs)": 12, "Amount paid in current Fin Year (in Rs) Wages": 12, "Amount paid in current Fin Year (in Rs) Material": 12,
  "Total Amount paid in current Fin Year (in Rs)": 14, "Amount booked since inception (in Rs) Wages": 12, "Amount booked since inception (in Rs) Material": 12,
  "Total Amount booked since inception (in Rs)": 14, "Amount paid since inception (in Rs) Wages": 12, "Amount paid since inception (in Rs) Material": 12,
  "Total Amount paid since inception (in Rs)": 14, "% of Total Amount paid": 10, "Total Mandays": 10, "Balnce Mandays": 12,
  "Mandays generated in current Fin Year": 12, "Work Started Date": 14, "Individual Beneficiary JCN (In case of Cat IV Works)": 25,
  "Work Category": 25, "Muster Issued Status": 20
};

const PMAY_HEADERS = [
  "SNo.","Panchayat Name","Work Start Fin Year","Work Status","Work Code","Work Name","Agency Name",
  "Sanction Amount (in Rs) Wages","Sanction Amount (in Rs) Material","Sanction Amount (in Rs)",
  "Amount paid in current Fin Year (in Rs) Wages","Amount paid in current Fin Year (in Rs) Material",
  "Amount booked since inception (in Rs) Wages","Amount booked since inception (in Rs) Material",
  "Amount paid since inception (in Rs) Wages","Amount paid since inception (in Rs) Material",
  "Total Mandays","Balance Mandays","Mandays generated in current Fin Year","Work Started Date","Individual Beneficiary JCN (In case of Cat IV Works)",
  "Muster Issued Status"
];
const PMAY_COLUMN_WIDTHS = [8,18,10,10,30,40,18,12,12,12,12,12,12,12,12,12,10,12,12,25,20];

// ------------ Utils ------------
function getUniqueFilePath(basePath, baseName, ext) {
  let counter = 0;
  let finalPath = path.join(basePath, `${baseName}${ext}`);
  while (fs.existsSync(finalPath)) {
    counter++;
    finalPath = path.join(basePath, `${baseName} (${counter})${ext}`);
  }
  return finalPath;
}
function parseNumericIfPossible(cell) {
  if (cell === null || cell === undefined) return null;
  const s = String(cell).trim();
  if (s.length === 0) return null;
  const cleaned = s.replace(/[,\s₹]/g, "").replace(/—/g, "-");
  const parenNeg = cleaned.match(/^\(([-\d.]+)\)$/);
  const candidate = parenNeg ? `-${parenNeg[1]}` : cleaned;
  if (/^-?\d+(\.\d+)?$/.test(candidate)) {
    const num = Number(candidate);
    if (!Number.isNaN(num) && Number.isFinite(num)) return Math.round(num);
  }
  return null;
}
function toNum(v){
  if (v === null || v === undefined || v === "") return 0;
  if (typeof v === "number") return v;
  const s = String(v).replace(/[,\s₹]/g, "").replace(/—/g, "-").trim();
  if (s.length === 0) return 0;
  const n = Number(s);
  return Number.isFinite(n) ? n : 0;
}
function colNumToLetter(num){
  let s="", n=num;
  while(n>0){ const m=(n-1)%26; s=String.fromCharCode(65+m)+s; n=Math.floor((n-1)/26); }
  return s;
}
function parseDateDMYToUTC(s){ if(!s||typeof s!=="string")return null; const m=s.match(/^\s*(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{4})\s*$/); if(m)return new Date(Date.UTC(+m[3],+m[2]-1,+m[1])); const m2=s.match(/^\s*(\d{4})[\/\-.](\d{1,2})[\/\-.](\d{1,2})\s*$/); if(m2)return new Date(Date.UTC(+m2[1],+m2[2]-1,+m2[3])); return null; }

// ------------ Mapping loaders ------------
async function loadMapping() {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(mappingFile);
  const ws = wb.worksheets[0];
  const map = {};
  ws.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return;
    const workCode = String(row.getCell(1).value || "").trim();
    const category = String(row.getCell(3).value || "").trim();
    if (workCode) map[workCode] = category || "UNKNOWN";
  });
  return map;
}

async function loadSubEngineerMap() {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(subEngineerMapFile);
  const ws = wb.worksheets[0];
  const map = {};
  ws.eachRow((row, rn) => {
    if (rn === 1) return;
    const pName = String(row.getCell(2).value || "").trim().toUpperCase();
    const seName = String(row.getCell(5).value || "").trim();
    if (pName) map[pName] = seName || "";
  });
  return map;
}

// ------------ Panchayat map (fallback) ------------
function loadPanchayatMap() {
  const mapFile = path.join(process.cwd(), "panchayat_map.json");
  if (fs.existsSync(mapFile)) {
    try { return JSON.parse(fs.readFileSync(mapFile, "utf8")); } catch (e) { /* ignore and use hardcoded */ }
  }
  return {
"1724009001":"ABHAPURI","1724009002":"AMBADOCHAR","1724009003":"ARDLA","1724009004":"BADI",
"1724009005":"BAYKHEDA","1724009006":"BEDCHHA","1724009007":"BETHANYA BUJURG","1724009073":"BHADALEN",
"1724009008":"BHAVSINGPURA","1724009009":"BILKHED","1724009010":"BORWAL (F)","1724009011":"BUNDA (F)",
"1724009012":"CHAINPUR","1724009013":"CHHAINDIYA","1724009014":"CHIRIYA","1724009015":"CHOPALI",
"1724009016":"DAMKHEDA","1724009017":"DEHRIYA","1724009018":"DEVIT BUJURG","1724009019":"DHASLGAON",
"1724009020":"DHUPA BUJURG (F)","1724009021":"DHUPI KHURD","1724009022":"GADGYAAM (F)","1724009023":"GAVLA",
"1724009024":"GHODI BUJURG","1724009025":"GORKHPUR","1724009072":"GUVADA","1724009026":"HARNKUNDIYA (F)",
"1724009027":"JAMLI","1724009028":"KAKODA","1724009029":"KAKRIYA","1724009030":"KHARWA",
"1724009031":"KHOI","1724009032":"KOTA BUJURG","1724009033":"KOTBEDA","1724009034":"KOTHDA",
"1724009035":"KUNDI","1724009036":"LAKHAPUR","1724009037":"MALGAON","1724009038":"MANDVI",
"1724009039":"MANIKERA","1724009040":"MARUGAD","1724009071":"MENDAGAD","1724009041":"MITAWAL",
"1724009074":"MORDAD","1724009042":"MORWA","1724009043":"MUNDIYA","1724009044":"NAHALDARI",
"1724009045":"NANKOUDI","1724009046":"NARWAT","1724009047":"NEEMSETHI","1724009048":"PADLYA",
"1724009049":"PAKHALYA","1724009050":"PALONA (F)","1724009051":"PIDIJAMLI","1724009052":"PIPARI",
"1724009053":"PIPRKHEDA NAKA","1724009054":"PUNASLA","1724009055":"PUTALA","1724009056":"RAJPURA",
"1724009057":"RATLIPURA","1724009058":"RATNPUR","1724009059":"REHTFAL","1724009060":"RUNDA (F)",
"1724009061":"SAIKHEDA","1724009062":"SAKA","1724009063":"SAPATIYA","1724009064":"SEMALKHUT",
"1724009065":"SHIVNA","1724009066":"SONKHEDI","1724009075":"TED","1724009076":"THOKANBEDA",
"1724009067":"TIGRIYA","1724009068":"TINSYA","1724009069":"TITRANYA (F)","1724009070":"ZIRANYA",
"1724":"ZP LEVEL WORK","1724009":"BP LEVEL WORK"
  };
}

// ------------ Scrapers ------------
const ORIGINAL_HEADERS = [
  "S. No.","District Name","Block Name","Panchayat Name","Work Start Fin Year","Work Status","Work Code","Work Name",
  "Master Work Category Name","Work Category Name","WORK TYPE as per new work creation Module","Work Type","Agency Name",
  "Sanction Amount (in Rs) Wages","Sanction Amount (in Rs) Material","Sanction Amount (in Rs)",
  "Amount paid in current Fin Year (in Rs) Wages","Amount paid in current Fin Year (in Rs) Material",
  "Amount booked since inception (in Rs) Wages","Amount booked since inception (in Rs) Material",
  "Amount paid since inception (in Rs) Wages","Amount paid since inception (in Rs) Material",
  "Total Mandays","Mandays generated in current Fin Year","No. Of Units","Is Exempted Ongoing Community Works(20) within GP","Is Secure",
  "Is Convergence","Work Started Date","Individual Beneficiary JCN (In case of Cat IV Works)"
];
const TARGET_HEADERS_WITHOUT_CAT = TARGET_HEADERS.filter(h => h !== "Work Category");
const KEEP_INDEXES = TARGET_HEADERS_WITHOUT_CAT.map(h => ORIGINAL_HEADERS.indexOf(h));

async function scrapeOngoingWorks(browser, webLogger) {
  const page = await browser.newPage();
  try {
    await webLogger("Starting scrapeOngoingWorks...");
    await page.goto(ONGOING_URL, { waitUntil: "domcontentloaded", timeout: 0 });

    async function waitForAjaxLocal() {
      try {
        await page.waitForSelector("#ContentPlaceHolder1_PageUpdateProgress", { visible: true, timeout: 10000 }).catch(()=>{});
        await page.waitForSelector("#ContentPlaceHolder1_PageUpdateProgress", { hidden: true, timeout: 180000 });
      } catch {}
    }

    await page.waitForSelector("#ddl_dist");
    await page.evaluate(()=>{ const el=document.querySelector("#ddl_dist"); if(el){ el.value="1724"; el.dispatchEvent(new Event("change")); } });
    await waitForAjaxLocal();

    await page.waitForSelector("#ddl_blk", { timeout: 60000 });
    await page.evaluate(()=>{ const el=document.querySelector("#ddl_blk"); if(el){ el.value="1724009"; el.dispatchEvent(new Event("change")); } });
    await waitForAjaxLocal();

    await page.waitForSelector("#ddlFin_year");
    await page.evaluate(()=>{ const el=document.querySelector("#ddlFin_year"); if(el){ el.value="ALL"; el.dispatchEvent(new Event("change")); } });
    await waitForAjaxLocal();

    await page.waitForSelector("#ContentPlaceHolder1_Button1");
    await page.evaluate(()=>{ const b=document.querySelector("#ContentPlaceHolder1_Button1"); if(b) b.click(); });
    await waitForAjaxLocal();

    await page.waitForFunction(() => document.querySelectorAll("table").length > 0, { polling: "mutation", timeout: 600000 });

    const tableInfo = await page.evaluate(() => {
      const tables = Array.from(document.querySelectorAll("table")).map((t,i)=>({ idx:i, rows: t.querySelectorAll("tr").length }));
      tables.sort((a,b)=>b.rows-a.rows);
      return tables.length ? tables[0] : null;
    });

    if (!tableInfo) throw new Error("No table found for ongoing works");

    let rawRows = await page.evaluate((tidx)=>{
      const tables = Array.from(document.querySelectorAll("table"));
      const t = tables[tidx];
      if(!t) return [];
      return Array.from(t.querySelectorAll("tr")).map(tr=>Array.from(tr.querySelectorAll("th, td")).map(td=>td.innerText.replace(/\u00A0/g," ").trim()));
    }, tableInfo.idx);

    rawRows = rawRows.filter(r => r.some(c => String(c||"").trim().length > 0));
    if (rawRows.length === 0) throw new Error("No data rows after cleanup in ongoing works");

    const startIndex = rawRows.length > 3 ? 3 : 0;
    const dataRows = rawRows.slice(startIndex);

    const filteredData = dataRows.map((row) =>
      KEEP_INDEXES.map((idx) => {
        const val = row[idx] !== undefined ? row[idx] : "";
        const parsed = parseNumericIfPossible(val);
        return parsed !== null ? parsed : val;
      })
    );

    const mapping = await loadMapping();
    const objects = filteredData.map(r => {
      const workCode = String(r[4] || "").trim();
      const category = mapping[workCode] || "UNKNOWN";
      const obj = {};
      TARGET_HEADERS.forEach((h,i) => {
        if (h === "Work Category") obj[h] = category;
        else obj[h] = r[i];
      });
      return obj;
    });

    await page.close();
    await webLogger("Completed scrapeOngoingWorks.");
    return objects;
  } catch (err) {
    try { await page.close(); } catch {}
    await webLogger("Error in scrapeOngoingWorks: " + (err && err.message ? err.message : err));
    throw err;
  }
}

// Muster helpers
function parseSNo(cell){ const m=(cell||"").match(/\d+/); return m?parseInt(m[0],10):null; }
function splitDateCell(cell){ const p=(cell||"").split(/-{2,}|–|—|-/).map(s=>s.trim()).filter(Boolean); return {from:p[0]||"",to:p[1]||""}; }
function detectProjectCodeRaw(rowCells, emrIndex, emrNo, workName){
  for(let i=0;i<rowCells.length;i++){
    if(i===0) continue;
    const c=(rowCells[i]||"").trim();
    if(!c) continue;
    if(/^\d{4,}\/.+/.test(c)) return c;
    if(/^\d{6,}$/.test(c)) return c;
    if(/^\d{4,}$/.test(c) && String(c)!==String(emrNo)) return c;
  }
  if(workName){
    const m=workName.match(/\(([^)]+)\)/);
    if(m&&m[1]) return m[1].trim();
  }
  return "";
}
function detectMappingCode(projectCodeRaw){
  if(!projectCodeRaw) return "";
  if(projectCodeRaw.includes("/")) return projectCodeRaw.split("/")[0].trim();
  const m=projectCodeRaw.match(/(\d{4,})/);
  if(m) return m[1];
  if(/^\d+$/.test(projectCodeRaw)) return projectCodeRaw;
  return "";
}

async function fetchTableRowsForMusterWithBrowser(browser, url){
  const page = await browser.newPage();
  try {
    await page.setDefaultNavigationTimeout(180000);
    await page.goto(url, { waitUntil: "networkidle2" }).catch(()=>page.goto(url,{waitUntil:"domcontentloaded"}));
    await page.waitForTimeout(1500);
    await page.waitForSelector("table", { timeout: 120000 });
    const rows = await page.evaluate(()=>{
      const tables = Array.from(document.querySelectorAll("table"));
      let best=null, bestCount=0;
      for(const t of tables){ const cnt = t.querySelectorAll("tr").length; if(cnt>bestCount){ best=t; bestCount=cnt; } }
      const tab = best || document.querySelector("table");
      if(!tab) return [];
      return Array.from(tab.querySelectorAll("tr")).map(tr=>Array.from(tr.querySelectorAll("td")).map(td=>td.innerText.replace(/\r/g," ").replace(/\n+/g," ").replace(/\s+/g," ").trim()));
    });
    await page.close();
    return rows;
  } catch (e) {
    try { await page.close(); } catch {}
    return [];
  }
}

function processRowsToObjects(rows, panchayatMap){
  const dataRows = rows.filter(r=>Array.isArray(r) && r.length>0);
  return dataRows.map(r=>{
    const sNo = parseSNo(r[0]) || "";
    let emrIndex = -1;
    for(let i=1;i<r.length;i++){ const c=(r[i]||"").trim(); if(/^\d{3,8}$/.test(c)){ emrIndex=i; break; } }
    if(emrIndex===-1){ for(let i=1;i<r.length;i++){ const c=(r[i]||"").trim(); if(/^\d+$/.test(c)){ emrIndex=i; break; } } }
    const workName = (emrIndex>0 ? r[emrIndex-1] : (r[4]||r[3]||r.slice(1).join(" | "))) || "";
    const emrNo = emrIndex>-1 ? (r[emrIndex]||"") : "";
    const rawDateCell = (emrIndex>-1 && r[emrIndex+1]) ? r[emrIndex+1] : (r[6]||"");
    const { from: dateFrom, to: dateTo } = splitDateCell(rawDateCell);
    const projectCodeRaw = detectProjectCodeRaw(r, emrIndex, emrNo, workName) || "";
    const firstPart = (projectCodeRaw || "").toString().split("/")[0].trim();
    const mappingCode = detectMappingCode(firstPart);
    const panchayat = panchayatMap[String(mappingCode)] || "NA";
    const projectNameClean = (""+workName).replace(/\(.*\)/,"").trim();
    return { SNo: sNo, ProjectName: projectNameClean, ProjectCode: projectCodeRaw, EMRNo: emrNo, DateFrom: dateFrom, DateTo: dateTo, Panchayat: panchayat };
  });
}

async function scrapeMusterData(browser, startSNo = 0, targetDate = null, webLogger) {
  if (!targetDate) {
    const t = new Date();
    targetDate = new Date(Date.UTC(t.getFullYear(), t.getMonth(), t.getDate()));
  }
  await webLogger("Starting scrapeMusterData...");
  const panchayatMap = loadPanchayatMap();
  const rawRows = await fetchTableRowsForMusterWithBrowser(browser, MUSTER_URL);
  const objects = processRowsToObjects(rawRows, panchayatMap);
  const filtered = objects.filter(o => { const s = parseInt(o.SNo); if(isNaN(s)) return false; return s >= startSNo; });
  if (!filtered.length) { await webLogger("No muster objects found"); return []; }
  const finalData = filtered.map(o=>{
    const dt = parseDateDMYToUTC(o.DateTo);
    let status = "Unknown";
    if (dt) {
      const dt0 = new Date(Date.UTC(dt.getUTCFullYear(), dt.getUTCMonth(), dt.getUTCDate()));
      status = dt0 < targetDate ? "Closed" : "Open";
    }
    let emrNo = o.EMRNo ? Math.round(Number(o.EMRNo)) : "";
    return { ...o, EMRNo: emrNo, Status: status };
  });
  await webLogger("Completed scrapeMusterData. Rows: " + finalData.length);
  return finalData;
}

// PMAY helpers
const delay = ms => new Promise(r => setTimeout(r, ms));
async function safeSelect(page, selector, value, opts = {}) {
  await page.waitForSelector(selector, { timeout: 20000 });
  await page.select(selector, value);
  await delay(250);
  await page.waitForSelector("#ContentPlaceHolder1_PageUpdateProgress", { visible: true, timeout: 2500 }).catch(()=>{});
  await page.waitForSelector("#ContentPlaceHolder1_PageUpdateProgress", { hidden: true, timeout: 10000 }).catch(()=>{});
  if (opts.waitForOption) {
    const start = Date.now();
    const tout = opts.waitForOption.timeout || 15000;
    while (Date.now() - start < tout) {
      const ok = await page.evaluate((sel,val) => {
        const el = document.querySelector(sel);
        if (!el) return false;
        return Array.from(el.options).some(o => o.value === val);
      }, opts.waitForOption.selector, opts.waitForOption.value);
      if (ok) break;
      await delay(300);
    }
  }
  await delay(200);
}
async function clickSubmit(page) {
  const selectors = [
    "#ContentPlaceHolder1_Button1",
    "#ctl00_ContentPlaceHolder1_Button1",
    "input[id*='Button1']",
    "button[id*='Button1']"
  ];
  for (const sel of selectors) {
    const el = await page.$(sel);
    if (el) {
      await page.evaluate(s => document.querySelector(s).click(), sel);
      return true;
    }
  }
  return false;
}
async function waitForTable(page, timeoutMs = 90000) {
  const start = Date.now();
  while (Date.now() - start < timeoutMs) {
    const ok = await page.evaluate(() => {
      const tables = Array.from(document.querySelectorAll("table"));
      for (const t of tables) if (t.rows && t.rows.length > 3) return true;
      return false;
    });
    if (ok) { await delay(250); return; }
    await delay(600);
  }
  throw new Error("Timeout waiting for table/export");
}
async function scrapeGrid(page) {
  const rawRows = await page.evaluate(() => {
    const tables = Array.from(document.querySelectorAll("table"));
    if (!tables.length) return [];
    let table = tables.find(t => t.id && t.id.includes("grdWorks"));
    if (!table) table = tables.sort((a,b) => (b.rows.length||0)-(a.rows.length||0))[0];
    if (!table) return [];
    const trs = Array.from(table.querySelectorAll("tr")).slice(3);
    return trs.map(tr => Array.from(tr.querySelectorAll("td")).map(td => td.innerText.replace(/\u00A0/g,' ').trim()));
  });

  const cleaned = rawRows.map(r => {
    const newRow = [];
    for (let i = 0; i < r.length; i++) {
      if (![1,2,8,9,10,11,24,25,26,27].includes(i)) newRow.push(r[i]);
    }
    if (newRow.length > PMAY_HEADERS.length) newRow.length = PMAY_HEADERS.length;
    while (newRow.length < PMAY_HEADERS.length) newRow.push('');
    return newRow;
  });
  return cleaned;
}
function roundNumericColumns(rows) {
  const numericCols = [7,8,9,10,11,12,13,14,15,16];
  return rows.map(r => r.map((val, idx) => {
    if (numericCols.includes(idx)) {
      const num = parseFloat(String(val).replace(/,/g,'').trim());
      if (!isNaN(num)) return Math.round(num);
    }
    return val;
  }));
}

// PMAY run
async function runPMAYStatus(browser, statusVal, statusName, webLogger) {
  const page = await browser.newPage();
  page.setDefaultNavigationTimeout(0);
  await webLogger(`Starting PMAY status ${statusName} (${statusVal})...`);
  await page.goto(ONGOING_URL, { waitUntil: "domcontentloaded", timeout: 0 });
  await safeSelect(page, "#ddl_dist", "1724", { waitForOption: { selector: "#ddl_blk", value: "1724009", timeout: 15000 } });
  await safeSelect(page, "#ddl_blk", "1724009");
  await safeSelect(page, "#Ddlworkcategory", "IF");
  try { await safeSelect(page, "#ddlprostatus", "Constr of PMAY-G House for Individuals"); } catch(e){}
  await safeSelect(page, "#ddlFin_year", "ALL");
  await safeSelect(page, "#Ddlwork_status", statusVal);
  await clickSubmit(page);
  await waitForTable(page, 90000);
  const rows = await scrapeGrid(page);
  await page.close();
  await webLogger(`Completed PMAY status ${statusName}. Rows: ${rows.length}`);
  return rows;
}

// ------------ Bucketing ------------
function bucketIndex(mandays) {
  const n = Number(mandays) || 0;
  if (n === 0) return 0;
  if (n >= 1 && n <= 25) return 1;
  if (n >= 26 && n <= 50) return 2;
  if (n >= 51 && n <= 75) return 3;
  if (n >= 76 && n <= 84) return 4;
  if (n >= 85 && n <= 89) return 5;
  if (n >= 90) return 6;
  return 0;
}
const BUCKET_LABELS = ["0","1-25","26-50","51-75","76-84","85-89","90 (पूर्ण आवास)"];
function safeNumFromRow(row, idx){ if (!Array.isArray(row)) return 0; return toNum(row[idx]); }

// ------------ Excel helpers ------------
function getUniqueSheetName(workbook, baseName){
  if(!workbook.getWorksheet(baseName)) return baseName;
  let i=2;
  while(true){
    const candidate = `${baseName} (${i})`;
    if(!workbook.getWorksheet(candidate)) return candidate;
    i++;
  }
}
function applyThreeLineHeader(ws, headerRowsArr, headersArray) {
  const lastColIdx = headersArray.length;
  const lastColLetterLocal = colNumToLetter(lastColIdx);
  headerRowsArr.forEach((h) => {
    const rowNumber = ws.lastRow ? ws.lastRow.number + 1 : 1;
    const row = ws.getRow(rowNumber);
    row.getCell(1).value = h.text;
    row.getCell(1).font = { bold: true, size: h.size };
    row.getCell(1).alignment = { horizontal: 'center', vertical: 'middle' };
    ws.mergeCells(`${colNumToLetter(1)}${rowNumber}:${lastColLetterLocal}${rowNumber}`);
    const mergedCell = ws.getCell(`${colNumToLetter(1)}${rowNumber}`);
    mergedCell.fill = { type:'pattern', pattern:'solid', fgColor:{argb:'FFCCFFCC'} };
    mergedCell.alignment = { horizontal:'center', vertical:'middle' };
    row.height = 25;
  });
  const headerRowNum = ws.lastRow.number + 1;
  const headerRow = ws.getRow(headerRowNum);
  for (let i=0;i<headersArray.length;i++){
    headerRow.getCell(i+1).value = headersArray[i] || "";
  }
  headerRow.height = 95;
  headerRow.font = { bold:true };
  headerRow.alignment = { horizontal:'center', vertical:'middle', wrapText:true };
  for (let i=1;i<=headersArray.length;i++){
    const cell = headerRow.getCell(i);
    cell.fill = { type:'pattern', pattern:'solid', fgColor:{argb:'FFADD8E6'} };
    cell.border = { top:{style:'medium'}, left:{style:'medium'}, bottom:{style:'medium'}, right:{style:'medium'} };
  }
  ws.autoFilter = { from: { row: headerRowNum, column: 1 }, to: { row: headerRowNum, column: headersArray.length } };
  ws.views = [{ state: "frozen", ySplit: headerRowNum }];
  return headerRowNum;
}
function addGrandTotalRow(ws, headerRowNum) {
  const firstDataRow = (typeof headerRowNum === "number") ? headerRowNum + 1 : 2;
  const lastDataRow = ws.lastRow ? ws.lastRow.number : firstDataRow;
  if (lastDataRow < firstDataRow) return;
  const totalRowNum = lastDataRow + 1;
  const totalRow = ws.getRow(totalRowNum);
  totalRow.getCell(1).value = "GRAND TOTAL";
  totalRow.getCell(1).font = { bold: true };
  totalRow.getCell(1).alignment = { horizontal: "center", vertical: "middle" };
  const lastCol = ws.columnCount || ws.actualColumnCount || ws.columns.length;
  for (let c = 2; c <= lastCol; c++) {
    const colLetter = colNumToLetter(c);
    const formula = `SUM(${colLetter}${firstDataRow}:${colLetter}${lastDataRow})`;
    totalRow.getCell(c).value = { formula };
    totalRow.getCell(c).numFmt = "#,##0";
  }
  totalRow.eachCell({ includeEmpty: true }, (cell) => {
    cell.font = { bold: true };
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFDDDDDD" } };
    cell.border = { top:{style:"medium"}, left:{style:"medium"}, bottom:{style:"medium"}, right:{style:"medium"} };
    cell.alignment = { horizontal: "center", vertical: "middle" };
  });
  ws.addRow([]);
}

// ------------ Sub-Engineer insertion helper ------------
function addSubEngineerColumnAtEnd(ws, panchayatHeaders, subEngMap) {
  try {
    // Find header row & Panchayat column
    let headerRowNum = 0, pCol = 0;
    const tryRows = Math.min(12, ws.rowCount || 12);

    for (let r = 1; r <= tryRows; r++) {
      const row = ws.getRow(r);
      if (!row) continue;
      for (let c = 1; c <= (row.actualCellCount || row.cellCount || ws.columnCount || 50); c++) {
        const val = String(row.getCell(c).value || "").trim().toUpperCase();
        if (panchayatHeaders.includes(val)) {
          headerRowNum = r;
          pCol = c;
          break;
        }
      }
      if (headerRowNum) break;
    }
    if (!headerRowNum || !pCol) return; // Sheet में GP/Panchayat ही नहीं → skip

    // ✅ नया कॉलम हमेशा END में जोड़ें
    const newCol = (ws.columnCount || ws.actualColumnCount) + 1;

    // Header
    const headerCell = ws.getRow(headerRowNum).getCell(newCol);
    headerCell.value = "Sub-Engineer";
    headerCell.font = { bold: true };
    headerCell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };

    // Rows
    const last = ws.lastRow ? ws.lastRow.number : ws.rowCount;
    for (let r = headerRowNum + 1; r <= last; r++) {
      const row = ws.getRow(r);
      const gpName = String(row.getCell(pCol).value || "").trim().toUpperCase();
      const seName = subEngMap[gpName] || "";
      const cell = row.getCell(newCol);
      cell.value = seName;

      // Formatting
      cell.alignment = { horizontal:'center', vertical:'middle', wrapText:true };
      cell.border = { top:{style:'thin'}, left:{style:'thin'}, bottom:{style:'thin'}, right:{style:'thin'} };
    }

    ws.getColumn(newCol).width = 22;

  } catch (e) {
    console.log("Sub-Engineer insertion error:", e.message);
  }
}

// ------------ Excel writer (with Sub-Engineer insertion) ------------
async function writeMergedExcel(ongoingObjects, musterObjects, pmayRows, webLogger) {
  if (!fs.existsSync(downloadPath)) fs.mkdirSync(downloadPath, { recursive: true });
  const today = new Date();
  const yyyy = today.getFullYear();
  const mm = String(today.getMonth() + 1).padStart(2, "0");
  const dd = String(today.getDate()).padStart(2, "0");
  const baseFileName = `MergedReport_${yyyy}${mm}${dd}`;
  const finalFilePath = getUniqueFilePath(downloadPath, baseFileName, ".xlsx");

  const workbook = new ExcelJS.Workbook();

  // --------------- Ongoing Works ---------------
  const ws1 = workbook.addWorksheet("Ongoing Works");
  const headerRow = ws1.addRow(ONGOING_HEADERS);
  headerRow.eachCell((cell) => {
    cell.font = { bold: true };
    cell.alignment = { vertical: "middle", horizontal: "center", wrapText: true };
    cell.border = { top: { style: "thick" }, left: { style: "thick" }, bottom: { style: "thick" }, right: { style: "thick" } };
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFD9D9D9" } };
  });
  ws1.getRow(1).height = 100;
  ws1.views = [{ state: "frozen", ySplit: 1 }];
  ws1.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: ONGOING_HEADERS.length } };

  const unknownRows = [];
  ongoingObjects.forEach((obj, idx) => {
    const mapObj = {};
    TARGET_HEADERS.forEach(h => { mapObj[h] = obj[h] !== undefined ? obj[h] : ""; });

    const sancW = toNum(mapObj["Sanction Amount (in Rs) Wages"]);
    const sancM = toNum(mapObj["Sanction Amount (in Rs) Material"]);
    const sancTotalFromField = toNum(mapObj["Sanction Amount (in Rs)"]);
    const sancTotal = sancTotalFromField || (sancW + sancM);

    const paidCurW = toNum(mapObj["Amount paid in current Fin Year (in Rs) Wages"]);
    const paidCurM = toNum(mapObj["Amount paid in current Fin Year (in Rs) Material"]);
    const totalPaidCurrentFY = paidCurW + paidCurM;

    const bookedW = toNum(mapObj["Amount booked since inception (in Rs) Wages"]);
    const bookedM = toNum(mapObj["Amount booked since inception (in Rs) Material"]);
    const totalBookedSince = bookedW + bookedM;

    const paidSinceW = toNum(mapObj["Amount paid since inception (in Rs) Wages"]);
    const paidSinceM = toNum(mapObj["Amount paid since inception (in Rs) Material"]);
    const totalPaidSince = paidSinceW + paidSinceM;

    const pct = sancTotal === 0 ? 0 : (totalPaidSince / sancTotal);
    const pctRounded = Math.round(pct * 10000) / 10000;

    const balMandaysRaw = Math.ceil((sancW - paidSinceW) / 261);
    const balanceMandays = balMandaysRaw > 0 ? balMandaysRaw : 0;

    const rowVals = [];
    for (const h of ONGOING_HEADERS) {
      switch (h) {
        case "Total Amount paid in current Fin Year (in Rs)": rowVals.push(totalPaidCurrentFY); break;
        case "Total Amount booked since inception (in Rs)": rowVals.push(totalBookedSince); break;
        case "Total Amount paid since inception (in Rs)": rowVals.push(totalPaidSince); break;
        case "% of Total Amount paid": rowVals.push(pctRounded); break;
        case "Balnce Mandays": rowVals.push(balanceMandays); break;
        case "Muster Issued Status": rowVals.push(""); break; // fill later
        default: rowVals.push(mapObj[h] ?? "");
      }
    }
    const r = ws1.addRow(rowVals);
    r.height = 28;
    r.eachCell((cell) => {
      cell.border = { top: { style: "medium" }, left: { style: "medium" }, bottom: { style: "medium" }, right: { style: "medium" } };
      cell.alignment = { vertical: "middle", horizontal: "center", shrinkToFit: true, wrapText: true };
    });
    if (idx % 2 === 1) r.eachCell(cell => cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF2F2F2" } });

    if (String(mapObj["Work Category"] || "").toUpperCase() === "UNKNOWN") {
      unknownRows.push({ obj: mapObj, ws1Row: r.number });
    }
  });

  ws1.columns.forEach((col, i) => {
    const header = ONGOING_HEADERS[i];
    const w = COLUMN_WIDTHS[header] || 15;
    col.width = Math.min(w, 40);
    col.alignment = { wrapText: true, horizontal: "center", vertical: "middle" };
  });

  await webLogger("Ongoing Works sheet written.");

  // Build musterOpenMap: (WorkCode → EMR) for Open
  const musterOpenMap = {};
  for (const m of musterObjects) {
    const code = (m.ProjectCode || "").toString().trim();
    const status = (m.Status || "").toString().trim();
    const emr = m.EMRNo !== undefined && m.EMRNo !== null ? String(m.EMRNo) : "";
    if (code && status.toLowerCase() === "open" && !musterOpenMap[code]) {
      musterOpenMap[code] = emr || "";
    }
  }

  // Fill Muster Issued Status in Ongoing Works sheet by Work Code
  const ongoingWorkCodeCol = ONGOING_HEADERS.indexOf("Work Code") + 1;
  const ongoingWorkCategoryCol = ONGOING_HEADERS.indexOf("Work Category") + 1;
  const pctColIndex = ONGOING_HEADERS.indexOf("% of Total Amount paid") + 1;
  for (let ri = 2; ri <= ws1.rowCount; ri++) {
    const workCode = String(ws1.getRow(ri).getCell(ongoingWorkCodeCol).value || "").trim();
    const musterStatusVal = musterOpenMap[workCode] ? musterOpenMap[workCode] : "No e-MR Issued";
    ws1.getRow(ri).getCell(ONGOING_HEADERS.indexOf("Muster Issued Status")+1).value = musterStatusVal;
    if (pctColIndex > 0) {
      const pctCell = ws1.getRow(ri).getCell(pctColIndex);
      pctCell.numFmt = "0.00%";
    }
  }
  await webLogger("Muster status filled into Ongoing Works sheet.");

  // ---------------- UNKNOWN sheet ----------------
  const wsUnknown = workbook.addWorksheet("UNKNOWN");
  const headerUnknown = wsUnknown.addRow(TARGET_HEADERS);
  headerUnknown.eachCell(cell => {
    cell.font = { bold: true }; cell.alignment = { vertical: "middle", horizontal: "center", wrapText: true };
    cell.border = { top:{style:"thick"}, left:{style:"thick"}, bottom:{style:"thick"}, right:{style:"thick"} };
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFF2CC" } };
  });
  wsUnknown.getRow(1).height = 100;
  wsUnknown.views = [{ state: "frozen", ySplit: 1 }];
  wsUnknown.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: TARGET_HEADERS.length } };

  unknownRows.forEach((u) => {
    const vals = TARGET_HEADERS.map(h => (h === "Work Category" ? "" : (u.obj[h] !== undefined ? u.obj[h] : "")));
    const r = wsUnknown.addRow(vals);
    r.height = 28;
    r.eachCell(cell => { cell.alignment = { vertical: "middle", horizontal: "center", wrapText: true }; });
  });
  wsUnknown.columns.forEach((col, i) => {
    const header = TARGET_HEADERS[i];
    col.width = Math.min(COLUMN_WIDTHS[header] || 15, 40);
    col.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
  });
  await webLogger("UNKNOWN sheet written.");

  // Put INDEX formula placeholders in Ongoing Works Work Category to read from UNKNOWN
  const unknownWorkCodeCol = TARGET_HEADERS.indexOf("Work Code") + 1;
  const unknownWorkCatCol = TARGET_HEADERS.indexOf("Work Category") + 1;
  const unknownWorkCodeColLetter = colNumToLetter(unknownWorkCodeCol);
  const unknownWorkCatColLetter = colNumToLetter(unknownWorkCatCol);
  const unknownStartRow = 2;
  const unknownEndRow = Math.max(2, 1 + unknownRows.length);

  const ongoingWorkCodeColLetter = colNumToLetter(ongoingWorkCodeCol);
  const ongoingWorkCatColLetter = colNumToLetter(ongoingWorkCategoryCol);
  const ongoingEndRow = Math.max(2, ws1.rowCount);

  for (let ri = 2; ri <= ws1.rowCount; ri++) {
    const workCodeCellA1 = `${colNumToLetter(ongoingWorkCodeCol)}${ri}`;
    const indexRange = `UNKNOWN!$${unknownWorkCatColLetter}$${unknownStartRow}:$${unknownWorkCatColLetter}$${unknownEndRow}`;
    const matchRange = `UNKNOWN!$${unknownWorkCodeColLetter}$${unknownStartRow}:$${unknownWorkCodeColLetter}$${unknownEndRow}`;
    const origCell = ws1.getRow(ri).getCell(ongoingWorkCategoryCol);
    const origValue = origCell.value !== null && origCell.value !== undefined ? String(origCell.value).replace(/"/g,'""') : "";
    const formula = `IFERROR(INDEX(${indexRange}, MATCH(${workCodeCellA1}, ${matchRange}, 0)), "${origValue}")`;
    ws1.getRow(ri).getCell(ongoingWorkCategoryCol).value = { formula };
    ws1.getRow(ri).getCell(ongoingWorkCategoryCol).alignment = { horizontal: "center", vertical: "middle", wrapText: true };
  }
  await webLogger("INDEX formulas set in Ongoing Works.");

  // ---------------- Muster Data sheet ----------------
  const ws2 = workbook.addWorksheet("Muster Data", { properties: { defaultRowHeight: 28 } });
  ws2.columns = [
    { header: "S No", key: "SNo", width: 8 },
    { header: "Project Name", key: "ProjectName", width: 40 },
    { header: "Project Code", key: "ProjectCode", width: 30 },
    { header: "E-MR No", key: "EMRNo", width: 10 },
    { header: "Date From", key: "DateFrom", width: 15 },
    { header: "Date To", key: "DateTo", width: 15 },
    { header: "Panchayat", key: "Panchayat", width: 20 },
    { header: "Status", key: "Status", width: 10 },
    { header: "Work Category", key: "WorkCategory", width: 25 }
  ];
  ws2.views = [{ state: "frozen", ySplit: 4 }];
  ws2.autoFilter = { from: { row: 4, column: 1 }, to: { row: 4, column: 9 } };

  const todayStr = new Date().toISOString().slice(0,10).split("-").reverse().join("/");
  ws2.mergeCells("A1:I1"); ws2.getCell("A1").value="कार्यालय महात्मा गांधी नरेगा, जनपद पंचायत झिरन्या, जिला खरगोन"; ws2.getCell("A1").alignment={horizontal:"center",vertical:"middle"}; ws2.getCell("A1").font={bold:true,size:14}; ws2.getRow(1).height=30;
  ws2.mergeCells("A2:I2"); ws2.getCell("A2").value="जारी मस्टर 2025-26 की जानकारी"; ws2.getCell("A2").alignment={horizontal:"center",vertical:"middle"}; ws2.getCell("A2").font={bold:true,size:14}; ws2.getRow(2).height=30;
  ws2.mergeCells("A3:I3"); ws2.getCell("A3").value=todayStr; ws2.getCell("A3").alignment={horizontal:"center",vertical:"middle"}; ws2.getCell("A3").font={bold:true,size:12}; ws2.getRow(3).height=25;

  const headerRow2 = ws2.getRow(4);
  ["S No","Project Name","Project Code","E-MR No","Date From","Date To","Panchayat","Status","Work Category"].forEach((h,i)=>{
    const cell = headerRow2.getCell(i+1);
    cell.value = h; cell.font = { bold:true }; cell.alignment = { horizontal:"center", vertical:"middle", wrapText:true };
    cell.border = { top:{style:"medium"}, left:{style:"medium"}, bottom:{style:"medium"}, right:{style:"medium"} };
  });
  ws2.getRow(4).height = 100;

  musterObjects.forEach((o, idx) => {
    const excelRowIndex = 5 + idx;
    const r = ws2.getRow(excelRowIndex);
    r.values = [o.SNo, o.ProjectName, o.ProjectCode, o.EMRNo, o.DateFrom, o.DateTo, o.Panchayat, o.Status, ""];
    r.height = 28;
    r.eachCell(cell => {
      cell.alignment = { horizontal:"center", vertical:"middle", wrapText:true };
      cell.border = { top:{style:"medium"}, left:{style:"medium"}, bottom:{style:"medium"}, right:{style:"medium"} };
    });
    const dateFromCell = r.getCell(5);
    const dateToCell = r.getCell(6);
    const d1 = parseDateDMYToUTC(dateFromCell.value);
    if (d1) { dateFromCell.value = d1; dateFromCell.numFmt = "dd/mm/yyyy"; }
    const d2 = parseDateDMYToUTC(dateToCell.value);
    if (d2) { dateToCell.value = d2; dateToCell.numFmt = "dd/mm/yyyy"; }

    // Map Work Category by WorkCode from 'Ongoing Works'
    const projectCodeCellRef = `C${excelRowIndex}`;
    const indexRange = `'Ongoing Works'!$${ongoingWorkCatColLetter}$2:$${ongoingWorkCatColLetter}$${ongoingEndRow}`;
    const matchRange = `'Ongoing Works'!$${ongoingWorkCodeColLetter}$2:$${ongoingWorkCodeColLetter}$${ongoingEndRow}`;
    const safeFallback = "Works completed/ Approved";
    const formula = `IFERROR(INDEX(${indexRange}, MATCH(TRIM(${projectCodeCellRef}), ${matchRange}, 0)), "${safeFallback}")`;
    r.getCell(9).value = { formula };
    r.getCell(9).alignment = { horizontal: "center", vertical: "middle", wrapText: true };
  });
  await webLogger("Muster Data sheet written.");

  // ---------------- PMAY Merged sheet ----------------
  const wsPMAY = workbook.addWorksheet('PMAY Merged', { views:[{state:'normal'}] });
  const now = new Date();
  const dateStr = `${String(now.getDate()).padStart(2,'0')}/${String(now.getMonth()+1).padStart(2,'0')}/${now.getFullYear()}`;
  const headerRows = [
    { text: "कार्यालय महात्मा गांधी नरेगा, जनपद पंचायत झिरन्या, जिला खरगोन", size: 18 },
    { text: "प्रधानमंत्री आवास (ग्रामीण) योजनांतर्गत कार्यो की जानकारी", size: 16 },
    { text: dateStr, size: 14 }
  ];
  const lastColIndex = PMAY_HEADERS.length;
  const lastColLetter = colNumToLetter(lastColIndex);

  headerRows.forEach((h) => {
    const rowNumber = wsPMAY.lastRow ? wsPMAY.lastRow.number + 1 : 1;
    const row = wsPMAY.getRow(rowNumber);
    row.getCell(1).value = h.text;
    row.getCell(1).font = { bold: true, size: h.size };
    row.getCell(1).alignment = { horizontal: 'center', vertical: 'middle' };
    wsPMAY.mergeCells(`${colNumToLetter(1)}${rowNumber}:${lastColLetter}${rowNumber}`);
    const mergedCell = wsPMAY.getCell(`${colNumToLetter(1)}${rowNumber}`);
    mergedCell.fill = { type:'pattern', pattern:'solid', fgColor:{argb:'FFCCFFCC'} };
    mergedCell.alignment = { horizontal:'center', vertical:'middle' };
    row.height = 25;
  });

  const headerRowNumber = wsPMAY.lastRow.number + 1;
  const headerRowPM = wsPMAY.getRow(headerRowNumber);
  for (let i = 0; i < lastColIndex; i++) headerRowPM.getCell(i+1).value = PMAY_HEADERS[i] || "";
  headerRowPM.height = 95;
  headerRowPM.font = { bold:true };
  headerRowPM.alignment = { horizontal:'center', vertical:'middle', wrapText:true };
  for (let i=1;i<=lastColIndex;i++){
    const cell = headerRowPM.getCell(i);
    cell.fill = { type:'pattern', pattern:'solid', fgColor:{argb:'FFADD8E6'} };
    cell.border = { top:{style:'medium'}, left:{style:'medium'}, bottom:{style:'medium'}, right:{style:'medium'} };
  }
  wsPMAY.autoFilter = { from: { row: headerRowNumber, column: 1 }, to: { row: headerRowNumber, column: lastColIndex } };

  // <-- Freeze header for PMAY Merged sheet
  wsPMAY.views = [{ state: 'frozen', ySplit: headerRowNumber }];

  await webLogger("PMAY Merged header created.");

  const pmayPrepared = roundNumericColumns(pmayRows || []);
  for (let i=0;i<pmayPrepared.length;i++){
    if (!Array.isArray(pmayPrepared[i])) pmayPrepared[i]=[];
    if (pmayPrepared[i].length < lastColIndex - 1) {
      while (pmayPrepared[i].length < lastColIndex - 1) pmayPrepared[i].push("");
    } else if (pmayPrepared[i].length > lastColIndex - 1) {
      pmayPrepared[i].length = lastColIndex - 1;
    }
  }

  const totalMandaysIndex = PMAY_HEADERS.indexOf("Total Mandays");
  for (let idx = 0; idx < pmayPrepared.length; idx++) {
    const rowVals = pmayPrepared[idx].slice();
    rowVals[0] = (idx + 1).toString(); // SNo
    const sancW = safeNumFromRow(rowVals, 7);
    const paidSinceW = safeNumFromRow(rowVals, 14);
    const balMandaysRaw = Math.ceil((sancW - paidSinceW) / 261);
    const balanceMandays = balMandaysRaw > 0 ? balMandaysRaw : 0;
    const insertPos = totalMandaysIndex + 1;
    rowVals.splice(insertPos, 0, balanceMandays);

    // Muster Issued Status
    const projectCode = String(rowVals[4] || "").trim();
    const musterVal = musterOpenMap[projectCode] ? musterOpenMap[projectCode] : "No e-MR Issued";
    rowVals[lastColIndex - 1] = musterVal;

    const newRow = wsPMAY.addRow(rowVals);
    newRow.height = 30;
    newRow.eachCell({ includeEmpty: true }, (cell) => {
      cell.alignment = { horizontal:'center', vertical:'middle', wrapText:true };
      cell.border = { top:{style:'thin'}, left:{style:'thin'}, bottom:{style:'thin'}, right:{style:'thin'} };
    });
    const dataRowIndex = headerRowNumber + 1 + idx;
    if (idx % 2 === 1) {
      for (let c = 1; c <= lastColIndex; c++) wsPMAY.getRow(dataRowIndex).getCell(c).fill = { type:'pattern', pattern:'solid', fgColor:{argb:'FFF2F2F2'} };
    }
  }
  wsPMAY.columns.forEach((col, idx) => { col.width = Math.min(PMAY_COLUMN_WIDTHS[idx] || 15, 40); col.alignment = { wrapText: true, horizontal: "center", vertical: "middle" }; });
  await webLogger("PMAY Merged sheet populated. Rows: " + pmayPrepared.length);

  // ---------------- GP Distribution (Overall) ----------------
  const wsGP = workbook.addWorksheet("GP Distribution", { views:[{state:'frozen', ySplit:1}] });
  const gpHeaders = ["GP"].concat(BUCKET_LABELS).concat(["Total Works","Total muster issued","Muster Not issued"]);
  const gpHeaderRowNum = applyThreeLineHeader(wsGP, [
    { text: "कार्यालय महात्मा गांधी नरेगा, जनपद पंचायत झिरन्या, जिला खरगोन", size: 18 },
    { text: "प्रधानमंत्री आवास (ग्रामीण) योजनांतर्गत कार्यो की जानकारी (GP Distribution - Overall)", size: 16 },
    { text: dateStr, size: 14 }
  ], gpHeaders);

  const gpStats = {};
  pmayPrepared.forEach((r) => {
    const gp = String(r[1] || "UNKNOWN").trim() || "UNKNOWN";
    const manDays = safeNumFromRow(r, 16);
    const bidx = bucketIndex(manDays);
    const workCode = String(r[4] || "").trim();
    const isIssued = musterOpenMap[workCode] ? 1 : 0;
    if (!gpStats[gp]) gpStats[gp] = { buckets: Array(7).fill(0), total:0, issued:0 };
    gpStats[gp].buckets[bidx] += 1;
    gpStats[gp].total += 1;
    gpStats[gp].issued += isIssued;
  });
  const gpNames = Object.keys(gpStats).sort((a,b)=>a.localeCompare(b));
  gpNames.forEach(gp => {
    const s = gpStats[gp];
    const notIssued = s.total - s.issued;
    const r = wsGP.addRow([gp].concat(s.buckets).concat([s.total, s.issued, notIssued]));
    r.eachCell(cell => {
      cell.alignment = { horizontal:'center', vertical:'middle', wrapText:true };
      cell.border = { top:{style:'thin'}, left:{style:'thin'}, bottom:{style:'thin'}, right:{style:'thin'} };
    });
  });
  const gpColWidths = [25,8,10,10,10,10,10,14,10,14,14];
  wsGP.columns.forEach((col, i) => { col.width = Math.min(gpColWidths[i] || 12, 40); col.alignment = { horizontal:'center', vertical:'middle', wrapText:true }; });
  addGrandTotalRow(wsGP, gpHeaderRowNum);
  await webLogger("GP Distribution sheet written (overall).");

  // ---------------- GP Distribution: Ongoing & Approved ----------------
  const wsGPOngoingName = getUniqueSheetName(workbook, "Ongoing Works");
  const wsGPOngoing = workbook.addWorksheet(wsGPOngoingName);
  const gpOngoingHeaderRowNum = applyThreeLineHeader(wsGPOngoing, [
    { text: "कार्यालय महात्मा गांधी नरेगा, जनपद पंचायत झिरन्या, जिला खरगोन", size: 18 },
    { text: "PMAY GP Distribution - Ongoing", size: 16 },
    { text: dateStr, size: 14 }
  ], gpHeaders);

  const wsGPApprovedName = getUniqueSheetName(workbook, "Approved Works");
  const wsGPApproved = workbook.addWorksheet(wsGPApprovedName);
  const gpApprovedHeaderRowNum = applyThreeLineHeader(wsGPApproved, [
    { text: "कार्यालय महात्मा गांधी नरेगा, जनपद पंचायत झिरन्या, जिला खरगोन", size: 18 },
    { text: "PMAY GP Distribution - Approved", size: 16 },
    { text: dateStr, size: 14 }
  ], gpHeaders);

  const gpStatsOngoing = {};
  const gpStatsApproved = {};
  pmayPrepared.forEach((r) => {
    const gp = String(r[1] || "UNKNOWN").trim() || "UNKNOWN";
    const manDays = safeNumFromRow(r, 16);
    const bidx = bucketIndex(manDays);
    const workCode = String(r[4] || "").trim();
    const isIssued = musterOpenMap[workCode] ? 1 : 0;
    const status = String(r[3] || "").toLowerCase();

    if (status.includes("ongoing")) {
      if (!gpStatsOngoing[gp]) gpStatsOngoing[gp] = { buckets: Array(7).fill(0), total:0, issued:0 };
      gpStatsOngoing[gp].buckets[bidx] += 1;
      gpStatsOngoing[gp].total += 1;
      gpStatsOngoing[gp].issued += isIssued;
    }
    if (status.includes("approved")) {
      if (!gpStatsApproved[gp]) gpStatsApproved[gp] = { buckets: Array(7).fill(0), total:0, issued:0 };
      gpStatsApproved[gp].buckets[bidx] += 1;
      gpStatsApproved[gp].total += 1;
      gpStatsApproved[gp].issued += isIssued;
    }
  });

  const gpOngoingNames = Object.keys(gpStatsOngoing).sort((a,b)=>a.localeCompare(b));
  gpOngoingNames.forEach(gp => {
    const s = gpStatsOngoing[gp];
    const notIssued = s.total - s.issued;
    const r = wsGPOngoing.addRow([gp].concat(s.buckets).concat([s.total, s.issued, notIssued]));
    r.eachCell(cell => {
      cell.alignment = { horizontal:'center', vertical:'middle', wrapText:true };
      cell.border = { top:{style:'thin'}, left:{style:'thin'}, bottom:{style:'thin'}, right:{style:'thin'} };
    });
  });
  wsGPOngoing.columns.forEach((col, i) => { col.width = Math.min(gpColWidths[i] || 12, 40); col.alignment = { horizontal:'center', vertical:'middle', wrapText:true }; });
  addGrandTotalRow(wsGPOngoing, gpOngoingHeaderRowNum);
  await webLogger(`GP-level Ongoing sheet written. GPs: ${gpOngoingNames.length}`);

  const gpApprovedNames = Object.keys(gpStatsApproved).sort((a,b)=>a.localeCompare(b));
  gpApprovedNames.forEach(gp => {
    const s = gpStatsApproved[gp];
    const notIssued = s.total - s.issued;
    const r = wsGPApproved.addRow([gp].concat(s.buckets).concat([s.total, s.issued, notIssued]));
    r.eachCell(cell => {
      cell.alignment = { horizontal:'center', vertical:'middle', wrapText:true };
      cell.border = { top:{style:'thin'}, left:{style:'thin'}, bottom:{style:'thin'}, right:{style:'thin'} };
    });
  });
  wsGPApproved.columns.forEach((col, i) => { col.width = Math.min(gpColWidths[i] || 12, 40); col.alignment = { horizontal:'center', vertical:'middle', wrapText:true }; });
  addGrandTotalRow(wsGPApproved, gpApprovedHeaderRowNum);
  await webLogger(`GP-level Approved sheet written. GPs: ${gpApprovedNames.length}`);

  // ---------------- Formatted PMAY (FIXED YES/NO + AutoFilter) ----------------
  const wsForm = workbook.addWorksheet("Formatted PMAY", { views:[{state:'frozen', ySplit:1}] });
  const FORM_HEADERS = ["S No.","Work Name","Work Code","Work Status","Financial Year","Man days","Estimate Cost","Exps","Balance Man days","MSR ISSUED","GP"];
  const headerF = wsForm.addRow(FORM_HEADERS);
  headerF.eachCell(cell=>{
    cell.font={bold:true}; cell.alignment={horizontal:'center', vertical:'middle', wrapText:true};
    cell.border={ top:{style:'medium'}, left:{style:'medium'}, bottom:{style:'medium'}, right:{style:'medium'} };
    cell.fill = { type:'pattern', pattern:'solid', fgColor:{argb:'FFEEECE1'} };
  });
  wsForm.getRow(1).height = 60;
  wsForm.columns = [
    {width:6},{width:50},{width:28},{width:12},{width:12},{width:10},{width:14},{width:14},{width:14},{width:12},{width:20}
  ];

  pmayPrepared.forEach((r, idx) => {
    const sNo = idx + 1;
    const workName = String(r[5] || "").trim();
    const workCode = String(r[4] || "").trim(); // ✅ Correct index
    const workStatus = String(r[3] || "").trim();
    const finYear = String(r[2] || "").trim();
    const manDays = safeNumFromRow(r, 16);
    const sancW = safeNumFromRow(r, 7);
    const sancM = safeNumFromRow(r, 8);
    const sancTotalField = safeNumFromRow(r, 9);
    const estimateCost = sancTotalField || (sancW + sancM);
    const paidSinceW = safeNumFromRow(r, 14);
    const paidSinceM = safeNumFromRow(r, 15);
    const exps = paidSinceW + paidSinceM;
    const balMandaysRaw = Math.ceil((sancW - paidSinceW) / 261);
    const balanceMandays = balMandaysRaw > 0 ? balMandaysRaw : 0;

    // ✅ Muster YES/NO using Open map
    const msrIssued = musterOpenMap[workCode] ? "YES" : "NO";

    const gp = String(r[1] || "").trim();
    const rowVals = [sNo, workName, workCode, workStatus, finYear, manDays, estimateCost, exps, balanceMandays, msrIssued, gp];
    const newR = wsForm.addRow(rowVals);
    newR.height = 22;
    newR.eachCell(cell => {
      cell.alignment = { horizontal:'center', vertical:'middle', wrapText:true };
      cell.border = { top:{style:'thin'}, left:{style:'thin'}, bottom:{style:'thin'}, right:{style:'thin'} };
    });
  });

  // ✅ AutoFilter ON for Formatted PMAY
  wsForm.autoFilter = {
    from: { row: 1, column: 1 },
    to: { row: 1, column: FORM_HEADERS.length }
  };

  await webLogger("Formatted PMAY sheet written.");

  // ---------------- Sub-Engineer mapping applied to all sheets ----------------
  const subEngMap = await loadSubEngineerMap();
  // header keywords we'll look for (uppercase)
  const panchayatHeaders = ["PANCHAYAT NAME", "PANCHAYAT", "GP"];

  // Apply to every worksheet (auto-detect header row & panchayat column)
  workbook.worksheets.forEach(ws => {
    addSubEngineerColumnAtEnd(ws, panchayatHeaders, subEngMap);
  });
// ✅ Fix Sub-Engineer column formatting + Auto-Filter in all sheets
workbook.worksheets.forEach(ws => {
  if (!ws) return;

  const headerRow = ws.getRow(1);
  if (!headerRow) return;

  const lastCol = ws.columnCount;
  if (lastCol < 2) return;

  // ---- Format Header Cell ----
  const h = headerRow.getCell(lastCol);
  h.font = { bold: true };
  h.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
  h.border = { top:{style:"medium"}, left:{style:"medium"}, bottom:{style:"medium"}, right:{style:"medium"} };
  ws.getColumn(lastCol).width = 22;

  // ---- Format All Data Cells ----
  for (let r = 2; r <= ws.rowCount; r++) {
    const cell = ws.getRow(r).getCell(lastCol);

    cell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
    cell.border = { top:{style:"thin"}, left:{style:"thin"}, bottom:{style:"thin"}, right:{style:"thin"} };

    // 🟦 Zebra striping (same pattern as sheet)
    if (r % 2 === 0) {
      cell.fill = { type:'pattern', pattern:'solid', fgColor:{ argb:'FFF2F2F2' } };
    }
  }

  // ✅ Extend Auto-Filter to include Sub-Engineer column
  if (ws.autoFilter && ws.autoFilter.from && ws.autoFilter.to) {
    ws.autoFilter.to.column = lastCol;
  }
});
  // ---------------- finalize workbook ----------------
  await workbook.xlsx.writeFile(finalFilePath);
  await webLogger("Excel file written: " + finalFilePath);
  return finalFilePath;
}

// ------------ MAIN ------------
(async () => {
  if (!fs.existsSync(downloadPath)) fs.mkdirSync(downloadPath, { recursive: true });

  const browser = await puppeteer.launch({
    headless: false, // per your choice
    defaultViewport: null,
    args: ["--no-sandbox"]
  });

  // Progress UI
  const progressPage = await browser.newPage();
  await progressPage.goto('about:blank');
  await progressPage.evaluate(() => {
    document.title = "Scrape Progress";
    document.body.style.fontFamily = "Arial, sans-serif";
    document.body.style.margin = "10px";
    const h = document.createElement('h2'); h.innerText = "Scrape Progress"; document.body.appendChild(h);
    const pre = document.createElement('pre'); pre.id = "log"; pre.style.whiteSpace = "pre-wrap"; pre.style.fontSize = "12px";
    pre.innerText = "";
    document.body.appendChild(pre);
  });

  async function webLogger(msg){
    try {
      console.log(msg);
      await progressPage.evaluate(m => {
        const p = document.getElementById('log');
        if (p) {
          p.innerText = p.innerText + "[" + new Date().toLocaleTimeString() + "] " + m + "\n";
          window.scrollTo(0, document.body.scrollHeight);
        }
      }, msg);
    } catch (e){
      console.log("webLogger error:", e && e.message ? e.message : e);
    }
  }

  try {
    await webLogger("Launching tasks...");

    const ongoingPromise = scrapeOngoingWorks(browser, webLogger);
    const musterPromise = (async ()=> {
      const startSNo = 0;
      const t = new Date();
      const targetDate = new Date(Date.UTC(t.getFullYear(), t.getMonth(), t.getDate()));
      return await scrapeMusterData(browser, startSNo, targetDate, webLogger);
    })();

    const pmayTask = (async () => {
      const statuses = [
        { val: "03", name: "Ongoing" },
        { val: "02", name: "Approved" }
      ];
      const results = await Promise.all(statuses.map(s => runPMAYStatus(browser, s.val, s.name, webLogger)));
      const merged = [];
      for (const r of results) merged.push(...(r || []));
      return merged;
    })();

    const [ongoingObjects, musterObjects, pmayRows] = await Promise.all([ongoingPromise, musterPromise, pmayTask]);

    await webLogger("All scrapers completed. Now writing Excel...");

    const outPath = await writeMergedExcel(ongoingObjects || [], musterObjects || [], pmayRows || [], webLogger);
    await webLogger("Excel saved: " + outPath);
    console.log("Excel saved:", outPath);
  } catch (err) {
    console.error("Error:", err && err.stack ? err.stack : err);
    await webLogger("Error: " + (err && err.message ? err.message : err));
    process.exitCode = 1;
  } finally {
    try { await browser.close(); } catch {}
    await webLogger("Browser closed.");
    console.log("Browser closed.");
  }
})();
