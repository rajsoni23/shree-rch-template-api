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
   1. RCH PARSER & FORMATTERS
   ========================================================= */

function clean(value) {
  return String(value ?? "")
    .replace(/\s+/g, " ")
    .trim();
}

function extractIdFormat(segment) {
  if (!segment) return "";
  let str = String(segment).trim().replace(/[\s\-]+/g, "_");

  const digitMatches = str.match(/\d+/g);
  if (!digitMatches) {
    return str.replace(/[\(\)\*]+/g, "").replace(/_+/g, "_").replace(/^|_$/g, "");
  }

  const id = digitMatches[digitMatches.length - 1];
  const idPos = str.lastIndexOf(id);
  const rawText = str.substring(0, idPos);

  let cleanText = rawText
    .replace(/[\(\)\*\d]+/g, " ")
    .replace(/_+/g, " ")
    .trim()
    .replace(/\s+/g, "_");

  if (!cleanText) {
    return `${id}__`;
  }

  return `${cleanText}__${id}__`;
}

function formatRCHName(value) {
  if (value === null || value === undefined) return "";
  let str = String(value).trim();

  if (str === "ALL" || str === "--ALL--" || str === "__ALL__") return "__ALL__";

  str = str.replace(/[\s\-]+/g, "_");

  if (str.includes("*")) {
    const parts = str.split("*");
    const rawLeft = parts[0].replace(/_+/g, "_").replace(/_$/, "");

    const rightDigits = parts[1].match(/\d+/);
    const rightId = rightDigits 
      ? rightDigits[0] 
      : parts[1].replace(/[()]/g, "").replace(/_+/g, "_").replace(/^|_$/g, "");

    const leftFormatted = extractIdFormat(rawLeft);

    if (leftFormatted.endsWith("__")) {
      return `${leftFormatted.slice(0, -1)}*__${rightId}_`;
    }

    return `${leftFormatted}___*___${rightId}__`;
  }

  return extractIdFormat(str);
}

function sanitizeNamedRange(name) {
  if (!name) return "_EMPTY_";

  return String(name)
    .replace(/[^a-zA-Z0-9_]/g, "_")
    .replace(/^[^a-zA-Z_]/, "_$&");
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
   2. DATA DECODER & HELPERS
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
   3. WORKBOOK GENERATOR
   ========================================================= */

function createWorkbook(type, locationData) {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = "SHREE RCH";

  const main = workbook.addWorksheet("main");
  const subcenter = workbook.addWorksheet("subcenter");
  const village = workbook.addWorksheet("Village");

  const headers = type === "child" ? CHILD_HEADERS : MOTHER_HEADERS;

  main.views = [{ state: "frozen", ySplit: 1 }];
  main.getRow(1).values = headers;
  main.getRow(1).height = 24;

  for (let c = 1; c <= headers.length; c++) {
    main.getColumn(c).width = Math.max(16, headers[c - 1].length + 4);
  }

  const rawPhcs = Object.keys(locationData);
  
  // FIX #1: PHC Master List in subcenter Column A to prevent 255-char cutoff limit
  subcenter.getCell(1, 1).value = "__ALL__";
  let phcMasterRow = 2;

  let subCol = 2;
  let villageCol = 2;

  rawPhcs.forEach((rawPhc) => {
    const phcFormatted = formatRCHName(rawPhc);
    const phcSafeName = sanitizeNamedRange(phcFormatted);
    if (!phcFormatted) return;

    // Add PHC to Column A Master List
    subcenter.getCell(phcMasterRow++, 1).value = phcFormatted;

    // Header for Subcentre columns
    subcenter.getCell(1, subCol).value = phcFormatted;

    const rawSubs = locationData[rawPhc] || {};
    const rawSubKeys = Object.keys(rawSubs);

    let subRow = 2;
    subcenter.getCell(subRow++, subCol).value = "__ALL__";

    rawSubKeys.forEach((rawSub) => {
      const subFormatted = formatRCHName(rawSub);
      const subSafeName = sanitizeNamedRange(subFormatted);
      if (!subFormatted) return;

      subcenter.getCell(subRow++, subCol).value = subFormatted;

      // Header for Village columns
      village.getCell(1, villageCol).value = subFormatted;
      
      // FIX #2: Handle both Array and Object formats for Village list
      const rawVillages = rawSubs[rawSub] || [];
      let villageList = [];
      if (Array.isArray(rawVillages)) {
        villageList = rawVillages;
      } else if (rawVillages && typeof rawVillages === "object") {
        villageList = Object.entries(rawVillages).map(([k, v]) => 
          (typeof v === "string" && v.trim() ? v : k)
        );
      }

      let vRow = 2;
      village.getCell(vRow++, villageCol).value = "__ALL__"; // "__ALL__" placed at top

      if (villageList.length > 0) {
        villageList.forEach((rawV) => {
          const vFormatted = formatRCHName(rawV);
          if (vFormatted && vFormatted !== "__ALL__") {
            village.getCell(vRow++, villageCol).value = vFormatted;
          }
        });
      }

      // FIX #3: Defined Range starts at Row 2 (skips header)
      const vColLetter = getColumnLetter(villageCol);
      const vMaxRow = Math.max(vRow - 1, 2);

      try {
        workbook.definedNames.add(
          `'Village'!$${vColLetter}$2:$${vColLetter}$${vMaxRow}`,
          subSafeName
        );
      } catch (e) {
        console.error("Named range error (Village):", subSafeName, e);
      }

      villageCol++;
    });

    const sColLetter = getColumnLetter(subCol);
    const sMaxRow = Math.max(subRow - 1, 2);

    try {
      workbook.definedNames.add(
        `'subcenter'!$${sColLetter}$2:$${sColLetter}$${sMaxRow}`,
        phcSafeName
      );
    } catch (e) {
      console.error("Named range error (PHC):", phcSafeName, e);
    }

    subCol++;
  });

  // PHC List Named Range for Column A Master List
  const phcMaxRow = Math.max(phcMasterRow - 1, 1);
  try {
    workbook.definedNames.add(
      `'subcenter'!$A$1:$A$${phcMaxRow}`,
      "PHC_LIST"
    );
  } catch (e) {
    console.error("Named range error (PHC_LIST):", e);
  }

  // Dynamic Data Validation Rules
  for (let row = 2; row <= MAX_ROWS + 1; row++) {
    // Health Facility (Uses Range instead of Inline String)
    main.getCell(`A${row}`).dataValidation = {
      type: "list",
      allowBlank: true,
      formulae: ["PHC_LIST"],
    };

    // SubCentre (Dependent Dropdown)
    main.getCell(`B${row}`).dataValidation = {
      type: "list",
      allowBlank: true,
      formulae: [`INDIRECT(SUBSTITUTE(A${row}, "*", "_"))`],
    };

    // Village (Dependent Dropdown)
    main.getCell(`C${row}`).dataValidation = {
      type: "list",
      allowBlank: true,
      formulae: [`INDIRECT(SUBSTITUTE(B${row}, "*", "_"))`],
    };
  }

  return workbook;
}

/* =========================================================
   4. API HANDLER
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
    console.error("RCH Backend Error:", error);
    res.status(400).json({ ok: false, error: error.message });
  }
};
