const ExcelJS = require("exceljs");

const ALLOW_ORIGIN = "*";
const MAX_ROWS = 2000;

const MOTHER_HEADERS = [
  "Health Facility",
  "SubCentre",
  "Village",
  "S.N.",
  "Name",
  "Husband Name",
  "Mobile Number",
  "Mother DOB",
  "Mother Weight",
  "LMP",
  "TT1",
  "High Risk",
  "ASHA",
  "ANM",
  "RG Number",
  "Address",
];

const CHILD_HEADERS = [
  "Health Facility",
  "SubCentre",
  "Village",
  "S.N.",
  "Child Name",
  "Gender",
  "Mother Name",
  "Father Name",
  "Mobile",
  "RgDate",
  "DOB",
  "BCG",
  "Weight",
  "Address",
  "BirthPlaceName",
  "HEP_B-0",
  "OPV-0",
  "HBIGDate",
];

/* =========================================================
   1. STRING & RCH FORMATTER UTILITIES
   ========================================================= */

function clean(value) {
  return String(value ?? "")
    .replace(/\s+/g, " ")
    .trim();
}

// Cell me display hone wala text format
function formatRCHName(value) {
  if (value === null || value === undefined) return "";
  
  return String(value)
    .trim()
    .replace(/\s+/g, "_")
    .replace(/\(/g, "__")
    .replace(/\)/g, "__");
}

// EXCEL NAMED RANGE SAFE IDENTIFIER (Fixes "Removed Records" Error)
function sanitizeNamedRangeName(name) {
  if (!name) return "_EMPTY_";

  let sanitized = String(name)
    .replace(/[^a-zA-Z0-9_]/g, "_") // Asterisk (*) aur baaki non-alphanumeric chars ko _ banata hai
    .replace(/^[^a-zA-Z_]/, "_$&");  // Ensure karta hai starting character valid ho

  return sanitized;
}

function normalizeObject(value) {
  if (Array.isArray(value)) {
    return value.map(normalizeObject);
  }
  if (value && typeof value === "object") {
    const out = {};
    for (const [key, val] of Object.entries(value)) {
      out[clean(key)] = normalizeObject(val);
    }
    return out;
  }
  return value;
}

/* =========================================================
   2. DECODE BASE64 DATA
   ========================================================= */

function decodeLocationData(baseData) {
  if (!baseData) throw new Error("baseData is missing.");

  let decoded;
  try {
    decoded = Buffer.from(String(baseData), "base64").toString("utf8");
  } catch {
    throw new Error("baseData is not valid Base64.");
  }

  let data;
  try {
    data = JSON.parse(decoded);
  } catch {
    throw new Error("Decoded baseData is not valid JSON.");
  }

  return normalizeObject(data);
}

/* =========================================================
   3. EXCEL COLUMN HELPER
   ========================================================= */

function getColumnLetter(colIndex) {
  let temp;
  let letter = "";
  while (colIndex > 0) {
    temp = (colIndex - 1) % 26;
    letter = String.fromCharCode(65 + temp) + letter;
    colIndex = (colIndex - temp - 1) / 26;
  }
  return letter;
}

/* =========================================================
   4. WORKBOOK BUILDER (FIXED DEFINED NAMES)
   ========================================================= */

function createWorkbook(type, locationData) {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = "SHREE RCH";

  const main = workbook.addWorksheet("main");
  const subcenter = workbook.addWorksheet("subcenter");
  const village = workbook.addWorksheet("Village");

  const headers = type === "child" ? CHILD_HEADERS : MOTHER_HEADERS;
  
  // Format 'main' header row
  main.views = [{ state: "frozen", ySplit: 1 }];
  main.getRow(1).values = headers;
  main.getRow(1).height = 24;

  for (let c = 1; c <= headers.length; c++) {
    main.getColumn(c).width = Math.max(16, headers[c - 1].length + 4);
  }

  // Build Subcenter & Village Structure
  const rawPhcs = Object.keys(locationData);
  const formattedPhcList = [];

  subcenter.getCell(1, 1).value = "__ALL__";
  village.getCell(1, 1).value = "__ALL__";
  
  let subCol = 2;
  let villageCol = 2;

  rawPhcs.forEach(rawPhc => {
    const phcFormatted = formatRCHName(rawPhc);
    const phcSafeName = sanitizeNamedRangeName(phcFormatted);
    if (!phcFormatted) return;

    formattedPhcList.push(phcFormatted);
    subcenter.getCell(1, subCol).value = phcFormatted;

    const rawSubs = locationData[rawPhc] || {};
    const rawSubKeys = Object.keys(rawSubs);

    let subRow = 2;
    subcenter.getCell(subRow++, subCol).value = "__ALL__";

    rawSubKeys.forEach(rawSub => {
      const subFormatted = formatRCHName(rawSub);
      const subSafeName = sanitizeNamedRangeName(subFormatted);
      if (!subFormatted) return;

      subcenter.getCell(subRow, subCol).value = subFormatted;
      subRow++;

      // Process Villages for this Subcentre
      village.getCell(1, villageCol).value = subFormatted;
      const rawVillages = rawSubs[rawSub] || [];

      let vRow = 2;
      if (Array.isArray(rawVillages) && rawVillages.length > 0) {
        rawVillages.forEach(rawV => {
          const vFormatted = formatRCHName(rawV);
          if (vFormatted) {
            village.getCell(vRow++, villageCol).value = vFormatted;
          }
        });
      }
      village.getCell(vRow, villageCol).value = "__ALL__";

      // Add SAFE Named Range for Subcentre -> Village List
      const vColLetter = getColumnLetter(villageCol);
      const vMaxRow = Math.max(vRow, 2);
      
      try {
        workbook.definedNames.add(`'Village'!$${vColLetter}$1:$${vColLetter}$${vMaxRow}`, subSafeName);
      } catch (e) {
        console.error("Named range error:", subSafeName, e);
      }

      villageCol++;
    });

    // Add SAFE Named Range for PHC -> Subcentre List
    const sColLetter = getColumnLetter(subCol);
    const sMaxRow = Math.max(subRow - 1, 2);
    
    try {
      workbook.definedNames.add(`'subcenter'!$${sColLetter}$2:$${sColLetter}$${sMaxRow}`, phcSafeName);
    } catch (e) {
      console.error("Named range error:", phcSafeName, e);
    }

    subCol++;
  });

  // Dynamic Data Validation
  if (formattedPhcList.length > 0) {
    const phcFormula = `"${["__ALL__", ...formattedPhcList].join(",")}"`;

    for (let row = 2; row <= MAX_ROWS + 1; row++) {
      // Health Facility Dropdown
      main.getCell(`A${row}`).dataValidation = {
        type: "list",
        allowBlank: true,
        formulae: [phcFormula]
      };

      // SubCentre Dropdown
      main.getCell(`B${row}`).dataValidation = {
        type: "list",
        allowBlank: true,
        formulae: [`INDIRECT(A${row})`]
      };

      // Village Dropdown
      main.getCell(`C${row}`).dataValidation = {
        type: "list",
        allowBlank: true,
        formulae: [`INDIRECT(B${row})`]
      };
    }
  }

  return workbook;
}

/* =========================================================
   5. REQUEST PARSER & HANDLER
   ========================================================= */

async function readBody(req) {
  if (req.body && typeof req.body === "object") return req.body;

  const chunks = [];
  for await (const chunk of req) chunks.push(chunk);

  const raw = Buffer.concat(chunks).toString("utf8");
  if (!raw) return {};

  if (String(req.headers["content-type"] || "").includes("application/json")) {
    return JSON.parse(raw);
  }
  return Object.fromEntries(new URLSearchParams(raw).entries());
}

module.exports = async function handler(req, res) {
  res.setHeader("Access-Control-Allow-Origin", ALLOW_ORIGIN);
  res.setHeader("Access-Control-Allow-Methods", "POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");
  res.setHeader("Cache-Control", "no-store");

  if (req.method === "OPTIONS") return res.status(204).end();
  if (req.method !== "POST") return res.status(405).json({ ok: false, error: "Use POST method." });

  try {
    const body = await readBody(req);
    const type = clean(body.typData).toLowerCase();

    if (type !== "mother" && type !== "child") {
      throw new Error("typData must be 'mother' or 'child'.");
    }

    const locationData = decodeLocationData(body.baseData);
    const workbook = createWorkbook(type, locationData);

    const buffer = await workbook.xlsx.writeBuffer();
    const filename = type === "child" ? "SHREE_RCH_Child_Formate.xlsx" : "SHREE_RCH_Mother_Formate.xlsx";

    res.status(200);
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", `attachment; filename="${filename}"`);
    res.send(Buffer.from(buffer));
  } catch (error) {
    console.error("RCH API Error:", error);
    res.status(400).json({ ok: false, error: error.message });
  }
};
