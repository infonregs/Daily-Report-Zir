/**
 * ✅ FINAL PERFECT CODE 
 * → Sub-Engineer + Work Name/Category
 * → Delay in Days (automatic current date से)
 * → तीनों शीट्स + yellow highlight
 * → Large page + slow NREGA site के लिए fixes
 */

const puppeteer = require("puppeteer");
const XLSX = require("xlsx-js-style");
const fs = require("fs/promises");
const path = require("path");

const DOWNLOAD_PATH = "H:\\scrapping";
const PROCESSED_FILE_NAME = "Tracking Report 25-26 zir.xlsx";
const MAPPING_FILE_PATH = "H:\\scrapping\\Muster Report zir 25-26.xlsx";

// 🔥 हर रोज़ automatic आज की date लेगा
const TODAY = new Date();
console.log(`📅 आज की तारीख (Delay calculation के लिए): ${TODAY.toLocaleDateString('en-GB')}`);

const normalizeText = s => (s || "").toString().replace(/\u200B/g, "").replace(/\s+/g, " ").replace(/[^\w\s\/-]/g, "").trim().toUpperCase();

function toDate(str) {
  if (!str) return null;
  const [d, m, y] = str.split("-").map(Number);
  if (!d || !m || !y) return null;
  const date = new Date(y, m - 1, d);
  return date.getFullYear() === y ? date : null;
}

function loadMappingFile() {
  console.log("📖 Reading mapping file...");
  const workbook = XLSX.readFile(MAPPING_FILE_PATH);
  const sheetName = workbook.SheetNames[0];
  const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1, defval: "" });

  let headerRowIndex = -1;
  const expected = ["PANCHAYAT", "SUB-ENGINEER", "PROJECT CODE", "PROJECT NAME", "WORK CATEGORY"];
  for (let i = 0; i < Math.min(rows.length, 10); i++) {
    const r = rows[i].map(c => (c || "").toString().toUpperCase());
    if (expected.filter(k => r.some(x => x.includes(k))).length >= 2) { headerRowIndex = i; break; }
  }
  if (headerRowIndex === -1) throw new Error("Header row not detected.");

  const headers = rows[headerRowIndex].map(h => (h || "").toString().trim());
  const dataRows = rows.slice(headerRowIndex + 1);

  const panchayatToEngineer = {};
  const workCodeToDetails = {};
  dataRows.forEach(r => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = r[i] || "");
    const panch = normalizeText(obj["Panchayat"] || obj["PANCHAYAT"]);
    const eng = (obj["Sub-Engineer"] || obj["SUB-ENGINEER"] || "").trim();
    const code = normalizeText(obj["Project Code"] || obj["PROJECT CODE"]);
    const name = (obj["Project Name"] || obj["PROJECT NAME"] || "").trim();
    const cat = (obj["Work Category"] || obj["WORK CATEGORY"] || "").trim();
    if (panch && eng) panchayatToEngineer[panch] = eng;
    if (code) workCodeToDetails[code] = { workName: name, workCat: cat };
  });

  console.log(`✅ Mapping loaded: ${Object.keys(panchayatToEngineer).length} Panchayats, ${Object.keys(workCodeToDetails).length} Works`);
  return { panchayatToEngineer, workCodeToDetails };
}

function createStyledSheet(header, data) {
  const ws = XLSX.utils.aoa_to_sheet([header, ...data]);
  const range = ws["!ref"] ? XLSX.utils.decode_range(ws["!ref"]) : null;
  if (!range) return ws;

  const border = { top:{style:"thin"}, bottom:{style:"thin"}, left:{style:"thin"}, right:{style:"thin"} };
  const headerStyle = { alignment:{horizontal:"center",vertical:"center",wrapText:true}, border, font:{bold:true,sz:12}, fill:{fgColor:{rgb:"FFD9D9D9"}} };
  const cellStyle = { alignment:{horizontal:"center",vertical:"center",wrapText:true}, border, font:{sz:11} };
  const altFill = { fgColor:{rgb:"FFF2F2F2"} };
  const yellowFill = { fgColor:{rgb:"FFFFE599"} };

  ws["!rows"] = [];
  for (let R = range.s.r; R <= range.e.r; R++) {
    for (let C = range.s.c; C <= range.e.c; C++) {
      const ref = XLSX.utils.encode_cell({r:R,c:C});
      if (!ws[ref]) continue;
      let style = R === 0 ? headerStyle : {...cellStyle, ...(R%2===1 ? {fill:altFill} : {})};
      if (ws[ref].v && ws[ref].v.toString().trim() === "Other Agency Work") style = { ...style, fill: yellowFill };
      ws[ref].s = style;
    }
    ws["!rows"][R] = {hpt: R===0?60:38};
  }

  const colWidths = [];
  for (let C = range.s.c; C <= range.e.c; C++) {
    let max = 10;
    for (let R = range.s.r + 1; R <= range.e.r; R++) {
      const cell = ws[XLSX.utils.encode_cell({r:R,c:C})];
      if (cell && cell.v) max = Math.max(max, cell.v.toString().length);
    }
    colWidths.push({wch: Math.min(max+2,38)});
  }
  ws["!cols"] = colWidths;
  ws["!autofilter"] = {ref: `A1:${XLSX.utils.encode_cell({r:0,c:range.e.c})}`};
  ws["!freeze"] = {r:1,c:0};
  return ws;
}

async function processExcelData(tableData, mapping) {
  console.log("⚙️ Processing + adding Delay in Days...");
  if (tableData.length < 2) { console.log("❌ No data"); return; }

  const header = tableData[0].slice();
  const data = tableData.slice(1);

  const panchayatIndex = header.findIndex(h => h && h.toLowerCase().includes("panchayat"));
  const workCodeIndex = header.findIndex(h => h && h.toLowerCase().includes("work code"));

  header.splice(panchayatIndex + 1, 0, "Sub-Engineer");
  let adjWorkCodeIndex = workCodeIndex + (workCodeIndex >= panchayatIndex + 1 ? 1 : 0);
  header.splice(adjWorkCodeIndex + 1, 0, "Work Name", "Work Category");

  let engCount = 0, workCount = 0;
  const updated = data.map(row => {
    const newRow = [...row];

    const panch = normalizeText(row[panchayatIndex]);
    const eng = mapping.panchayatToEngineer[panch] || "";
    if (eng) engCount++;
    newRow.splice(panchayatIndex + 1, 0, eng);

    let code = (row[workCodeIndex] || "").toString().trim().toUpperCase();
    const short = code.includes("/") ? code.split("/")[0] : code;
    let det = mapping.workCodeToDetails[code] || mapping.workCodeToDetails[short] || {workName: "Other Agency Work", workCat: "Other Agency Work"};
    if (det.workName !== "Other Agency Work") workCount++;
    newRow.splice(adjWorkCodeIndex + 1, 0, det.workName, det.workCat);

    // Delay in Days - automatic current date
    const closureStr = row[6];
    const closureDate = toDate(closureStr);
    let delay = "";
    if (closureDate) {
      const diff = Math.floor((TODAY - closureDate) / (1000 * 60 * 60 * 24));
      delay = diff >= 0 ? diff : "";
    }
    newRow.push(delay);

    return newRow;
  });

  header.push("Delay in Days");

  console.log(`✅ Mapping Done → Engineers: ${engCount}, Works: ${workCount}`);

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, createStyledSheet(header, updated), "Report");

  const d1 = header.findIndex(h => h && h.toLowerCase().includes("date of 1st sign"));
  const d2 = header.findIndex(h => h && h.toLowerCase().includes("date of 2nd sign"));

  const pending = updated.filter(r => !r[d1] || String(r[d1]).trim() === "");
  XLSX.utils.book_append_sheet(wb, createStyledSheet(header, pending), "Pending Musters");

  if (d1 !== -1 && d2 !== -1) {
    const aao = updated.filter(r => String(r[d1]).trim() !== "" && String(r[d2]).trim() === "");
    XLSX.utils.book_append_sheet(wb, createStyledSheet(header, aao), "Muster Pending at AAO Login");
  }

  const outPath = path.join(DOWNLOAD_PATH, PROCESSED_FILE_NAME);
  XLSX.writeFile(wb, outPath);
  console.log(`✅ REPORT READY WITH CURRENT DATE DELAY → ${outPath}`);
}

// ===================== MAIN =====================
(async () => {
  await fs.mkdir(DOWNLOAD_PATH, { recursive: true });
  const browser = await puppeteer.launch({ headless: false, defaultViewport: null });
  const page = await browser.newPage();

  const URL = "https://nregastrep.nic.in/netnrega/dynamic_muster_track.aspx?lflag=eng&state_code=17&fin_year=2025-2026&state_name=%u092e%u0927%u094d%u092f+%u092a%u094d%u0930%u0926%u0947%u0936+&Digest=%2f0dclwkJQM2w4GAt8GjFPw";

  try {
    console.log("🚀 Opening page...");
    await page.goto(URL, { waitUntil: "domcontentloaded", timeout: 90000 });

    await page.select("#ctl00_ContentPlaceHolder1_ddl_state", "17");
    await page.waitForTimeout(3000);

    await Promise.all([ page.waitForNavigation({ waitUntil: "networkidle2", timeout: 300000 }), page.select("#ctl00_ContentPlaceHolder1_ddl_dist", "1724") ]);
    await page.waitForTimeout(5000);

    await Promise.all([ page.waitForNavigation({ waitUntil: "networkidle2", timeout: 300000 }), page.select("#ctl00_ContentPlaceHolder1_ddl_blk", "1724009") ]);
    await page.waitForTimeout(5000);

    await page.select("#ctl00_ContentPlaceHolder1_ddl_pan", "ALL");
    await page.waitForTimeout(3000);

    const radio = (await page.$x("//label[contains(., 'Where payment is pending')]"))[0];
    if (radio) await radio.click();
    await page.waitForTimeout(2000);

    console.log("🔄 Submitting... (बड़ा डेटा है, 3-5 मिनट लग सकता है)");
    await Promise.all([
      page.waitForNavigation({ waitUntil: "networkidle2", timeout: 300000 }),
      page.evaluate(() => { const b = document.querySelector('input[type="submit"]'); if (b) b.click(); })
    ]).catch(() => console.log("⏳ अभी लोड हो रहा है, इंतज़ार करो..."));

    console.log("⏳ टेबल का इंतज़ार (5 मिनट तक)...");
    await page.waitForFunction(
      () => Array.from(document.querySelectorAll("table")).some(t => t.innerText.includes("S.No") || t.innerText.includes("SNo")),
      { timeout: 300000 }
    );

    const tableData = await page.evaluate(() => {
      const tables = [...document.querySelectorAll("table")];
      const target = tables.find(t => t.innerText.includes("S.No") || t.innerText.includes("SNo"));
      if (!target) return [];
      return [...target.querySelectorAll("tr")]
        .map(r => [...r.querySelectorAll("th,td")].map(c => c.innerText.trim()))
        .filter(r => r.length > 5);
    });

    console.log(`✅ ${tableData.length} rows निकाले`);

    const mapping = loadMappingFile();
    await processExcelData(tableData, mapping);

  } catch (e) {
    console.error("❌ Error:", e.message);
  } finally {
    await browser.close();
    console.log("🚪 Browser बंद");
  }
})();