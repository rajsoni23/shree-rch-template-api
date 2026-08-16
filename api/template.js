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
   BASIC HELPERS
   ========================================================= */

function clean(value) {
  return String(value ?? "")
    .replace(/\s+/g, " ")
    .trim();
}


/*
 * BAGHAURA (3181)
 *       ↓
 * BAGHAURA__3181__
 */
function convertLocationFormat(value) {

  if (
    value === null ||
    value === undefined
  ) {
    return "";
  }

  return String(value)
    .trim()
    .replace(/\(/g, "__")
    .replace(/\)/g, "__");
}


/* =========================================================
   NORMALIZE OBJECT
   ========================================================= */

function normalizeObject(value) {

  if (Array.isArray(value)) {
    return value.map(normalizeObject);
  }

  if (
    value &&
    typeof value === "object"
  ) {

    const out = {};

    for (
      const [key, val]
      of Object.entries(value)
    ) {

      out[clean(key)] =
        normalizeObject(val);

    }

    return out;
  }

  return value;
}


/* =========================================================
   SAFE EXCEL NAME
   ========================================================= */

function safeName(
  value,
  fallback = "ITEM"
) {

  let name =
    clean(value)
      .replace(
        /[^A-Za-z0-9_]/g,
        "_"
      )
      .replace(
        /_+/g,
        "_"
      );

  if (!name) {
    name = fallback;
  }

  if (
    /^[0-9]/.test(name)
  ) {
    name = "_" + name;
  }

  return name.slice(
    0,
    200
  );
}


/* =========================================================
   REQUEST PARSER
   ========================================================= */

function parseRequest(req) {

  const type =
    req.body?.typData;

  const baseData =
    req.body?.baseData;

  return {
    type,
    baseData
  };
}


/* =========================================================
   DECODE LOCATION DATA
   ========================================================= */

function decodeLocationData(
  baseData
) {

  if (!baseData) {

    throw new Error(
      "baseData is missing."
    );

  }

  let decoded;

  try {

    decoded =
      Buffer
        .from(
          String(baseData),
          "base64"
        )
        .toString("utf8");

  }
  catch {

    throw new Error(
      "baseData is not valid Base64."
    );

  }

  let data;

  try {

    data =
      JSON.parse(
        decoded
      );

  }
  catch {

    throw new Error(
      "Decoded baseData is not valid JSON."
    );

  }

  if (
    !data ||
    typeof data !== "object" ||
    Array.isArray(data)
  ) {

    throw new Error(
      "Location data must be an object."
    );

  }

  return normalizeObject(
    data
  );
}


/* =========================================================
   LOCATION MODEL
   ========================================================= */

function buildLocationModel(
  locationData
) {

  const phcs = [];

  const subByPhc =
    new Map();

  const villagesByPair =
    new Map();

  for (
    const [
      rawPhc,
      rawSubs
    ]
    of Object.entries(
      locationData
    )
  ) {

    const phc =
      convertLocationFormat(
        rawPhc
      );

    if (
      !phc ||
      !rawSubs ||
      typeof rawSubs !== "object" ||
      Array.isArray(rawSubs)
    ) {

      continue;

    }

    if (
      !phcs.includes(phc)
    ) {

      phcs.push(phc);

    }

    if (
      !subByPhc.has(phc)
    ) {

      subByPhc.set(
        phc,
        []
      );

    }

    for (
      const [
        rawSub,
        rawVillages
      ]
      of Object.entries(
        rawSubs
      )
    ) {

      const sub =
        convertLocationFormat(
          rawSub
        );

      if (!sub) {
        continue;
      }

      const subs =
        subByPhc.get(phc);

      if (
        !subs.includes(sub)
      ) {

        subs.push(sub);

      }

      const pairKey =
        `${phc}|${sub}`;

      if (
        !villagesByPair.has(
          pairKey
        )
      ) {

        villagesByPair.set(
          pairKey,
          []
        );

      }

      if (
        Array.isArray(
          rawVillages
        )
      ) {

        for (
          const rawVillage
          of rawVillages
        ) {

          const village =
            convertLocationFormat(
              rawVillage
            );

          if (
            village &&
            !villagesByPair
              .get(pairKey)
              .includes(village)
          ) {

            villagesByPair
              .get(pairKey)
              .push(
                village
              );

          }

        }

      }

    }

  }

  if (
    !phcs.length
  ) {

    throw new Error(
      "No Health Facility/PHC data was supplied."
    );

  }

  if (
    ![
      ...villagesByPair.values()
    ].some(
      list => list.length > 0
    )
  ) {

    throw new Error(
      "No Village data was supplied."
    );

  }

  return {
    phcs,
    subByPhc,
    villagesByPair
  };
}


/* =========================================================
   EXCEL COLUMN LETTER
   ========================================================= */

function columnLetter(
  column
) {

  let result = "";

  while (
    column > 0
  ) {

    const remainder =
      (column - 1) % 26;

    result =
      String.fromCharCode(
        65 + remainder
      ) + result;

    column =
      Math.floor(
        (column - 1) / 26
      );

  }

  return result;
}


/* =========================================================
   NAMED RANGE
   ========================================================= */

function addNamedRange(
  workbook,
  name,
  sheetName,
  range
) {

  /*
   * ExcelJS:
   *
   * first = range
   * second = name
   */

  workbook.definedNames.add(
    `'${sheetName}'!${range}`,
    name
  );

}


/* =========================================================
   LIST VALIDATION
   ========================================================= */

function addListValidation(
  cell,
  formula,
  errorTitle,
  error
) {

  cell.dataValidation = {

    type: "list",

    allowBlank: true,

    showInputMessage: true,

    showErrorMessage: true,

    errorStyle: "stop",

    errorTitle,

    error,

    formulae: [
      formula
    ]

  };

}


/* =========================================================
   HEADER STYLE
   ========================================================= */

function styleHeader(
  row
) {

  row.height = 24;

  row.eachCell(
    cell => {

      cell.font = {

        bold: true,

        color: {
          argb:
            "FFFFFFFF"
        }

      };

      cell.fill = {

        type: "pattern",

        pattern: "solid",

        fgColor: {
          argb:
            "FF1F4E78"
        }

      };

      cell.alignment = {

        vertical:
          "middle",

        horizontal:
          "center",

        wrapText:
          true

      };

      cell.border = {

        top: {
          style: "thin"
        },

        left: {
          style: "thin"
        },

        bottom: {
          style: "thin"
        },

        right: {
          style: "thin"
        }

      };

    }
  );

}


/* =========================================================
   MAIN SHEET
   ========================================================= */

function configureMainSheet(
  ws,
  headers
) {

  ws.views = [
    {
      state: "frozen",
      ySplit: 1
    }
  ];

  ws.getRow(
    1
  ).values =
    headers;

  styleHeader(
    ws.getRow(1)
  );

  ws.getColumn(1).width =
    28;

  ws.getColumn(2).width =
    28;

  ws.getColumn(3).width =
    32;

  for (
    let c = 4;
    c <= headers.length;
    c++
  ) {

    ws.getColumn(c).width =
      Math.max(
        14,
        Math.min(
          28,
          headers[c - 1].length + 4
        )
      );

  }

}


/* =========================================================
   CREATE WORKBOOK
   ========================================================= */

function createWorkbook(
  type,
  locationData
) {

  const model =
    buildLocationModel(
      locationData
    );

  const workbook =
    new ExcelJS.Workbook();

  workbook.creator =
    "SHREE RCH";

  workbook.created =
    new Date();

  workbook.modified =
    new Date();


  /* =======================================================
     MAIN
     ======================================================= */

  const main =
    workbook.addWorksheet(
      "Main"
    );


  /* =======================================================
     SUBCENTER
     ======================================================= */

  const subcenter =
    workbook.addWorksheet(
      "Subcenter"
    );


  /* =======================================================
     VILLAGE
     ======================================================= */

  const village =
    workbook.addWorksheet(
      "Village"
    );


  const headers =
    type === "child"
      ? CHILD_HEADERS
      : MOTHER_HEADERS;


  configureMainSheet(
    main,
    headers
  );


  /* =======================================================
     SUBCENTER DATA
     ======================================================= */

  subcenter.getRow(
    1
  ).values = [
    "PHC",
    "SubCentre"
  ];

  styleHeader(
    subcenter.getRow(1)
  );

  subcenter.getColumn(1).width =
    30;

  subcenter.getColumn(2).width =
    30;

  subcenter.getColumn(4).width =
    2;

  subcenter.getColumn(4).hidden =
    true;


  /*
   * PHC → row range
   */

  let subRow = 2;

  const phcRangeRows = [];


  for (
    const phc
    of model.phcs
  ) {

    const start =
      subRow;

    const subs =
      model.subByPhc.get(
        phc
      ) || [];


    for (
      const sub
      of subs
    ) {

      subcenter.getCell(
        subRow,
        1
      ).value =
        phc;

      subcenter.getCell(
        subRow,
        2
      ).value =
        sub;

      subRow++;

    }


    if (
      subRow > start
    ) {

      phcRangeRows.push({

        phc,

        start,

        end:
          subRow - 1

      });

    }

  }


  /*
   * Hidden PHC list
   */

  const phcListStart =
    2;


  model.phcs.forEach(
    (
      phc,
      index
    ) => {

      subcenter.getCell(
        phcListStart + index,
        4
      ).value =
        phc;

    }
  );


  /*
   * PHC_LIST
   */

  addNamedRange(

    workbook,

    "PHC_LIST",

    "Subcenter",

    `$D$${phcListStart}:$D$${phcListStart + model.phcs.length - 1}`

  );


  /*
   * SC_1, SC_2...
   */

  phcRangeRows.forEach(
    (
      item,
      index
    ) => {

      addNamedRange(

        workbook,

        `SC_${index + 1}`,

        "Subcenter",

        `$B$${item.start}:$B$${item.end}`

      );

    }
  );


  /* =======================================================
     VILLAGE DATA
     ======================================================= */

  village.getRow(
    1
  ).values = [
    "PHC",
    "SubCentre",
    "Village",
    "PAIR_KEY"
  ];

  styleHeader(
    village.getRow(1)
  );

  village.getColumn(1).width =
    30;

  village.getColumn(2).width =
    30;

  village.getColumn(3).width =
    36;

  village.getColumn(4).width =
    2;

  village.getColumn(4).hidden =
    true;


  let villageRow = 2;

  const pairRows = [];

  const pairList = [];


  for (
    const phc
    of model.phcs
  ) {

    const subs =
      model.subByPhc.get(
        phc
      ) || [];


    for (
      const sub
      of subs
    ) {

      const pairKey =
        `${phc}|${sub}`;

      const villages =
        model.villagesByPair
          .get(pairKey) || [];


      if (
        !villages.length
      ) {

        continue;

      }


      const start =
        villageRow;


      for (
        const villageName
        of villages
      ) {

        village.getCell(
          villageRow,
          1
        ).value =
          phc;

        village.getCell(
          villageRow,
          2
        ).value =
          sub;

        village.getCell(
          villageRow,
          3
        ).value =
          villageName;

        village.getCell(
          villageRow,
          4
        ).value =
          pairKey;

        villageRow++;

      }


      pairRows.push({

        pairKey,

        start,

        end:
          villageRow - 1

      });


      pairList.push(
        pairKey
      );

    }

  }


  /*
   * Hidden pair list
   */

  village.getColumn(5).width =
    2;

  village.getColumn(5).hidden =
    true;


  pairList.forEach(
    (
      pair,
      index
    ) => {

      village.getCell(
        index + 2,
        5
      ).value =
        pair;

    }
  );


  /*
   * PAIR_KEYS
   */

  if (
    pairList.length
  ) {

    addNamedRange(

      workbook,

      "PAIR_KEYS",

      "Village",

      `$E$2:$E$${pairList.length + 1}`

    );

  }


  /*
   * V_1, V_2...
   */

  pairRows.forEach(
    (
      item,
      index
    ) => {

      addNamedRange(

        workbook,

        `V_${index + 1}`,

        "Village",

        `$C$${item.start}:$C$${item.end}`

      );

    }
  );


  /* =======================================================
     DEPENDENT DROPDOWNS
     ======================================================= */

  for (
    let row = 2;
    row <= MAX_ROWS + 1;
    row++
  ) {


    /* =====================================================
       A = HEALTH FACILITY
       ===================================================== */

    addListValidation(

      main.getCell(
        row,
        1
      ),

      "=PHC_LIST",

      "Invalid Health Facility",

      "Select a Health Facility from the dropdown."

    );


    /* =====================================================
       B = SUBCENTRE
       =====================================================

       IMPORTANT:

       We use OFFSET + MATCH.

       This avoids:
       INDIRECT("SC_"&MATCH(...))

       which was the unreliable part.
       ===================================================== */

    const subCenterFormula =

      `=IFERROR(` +

      `OFFSET(` +

      `Subcenter!$B$1,` +

      `MATCH(` +

      `A${row},` +

      `Subcenter!$D$2:$D$${model.phcs.length + 1},` +

      `0` +

      `),` +

      `0,` +

      `COUNTIF(` +

      `Subcenter!$A$2:$A$${Math.max(
        2,
        subRow - 1
      )},` +

      `A${row}` +

      `),` +

      `1` +

      `)` +

      `,"")`;


    addListValidation(

      main.getCell(
        row,
        2
      ),

      subCenterFormula,

      "Invalid SubCentre",

      "Select a SubCentre belonging to the selected Health Facility."

    );


    /* =====================================================
       C = VILLAGE
       =====================================================

       A = PHC
       B = SubCentre

       We find:
       PHC|SubCentre

       Then use that pair's village range.
       ===================================================== */

    const villageFormula =

      `=IFERROR(` +

      `OFFSET(` +

      `Village!$C$1,` +

      `MATCH(` +

      `A${row}&"|"&B${row},` +

      `Village!$E$2:$E$${pairList.length + 1},` +

      `0` +

      `),` +

      `0,` +

      `COUNTIF(` +

      `Village!$D$2:$D$${Math.max(
        2,
        villageRow - 1
      )},` +

      `A${row}&"|"&B${row}` +

      `),` +

      `1` +

      `)` +

      `,"")`;


    addListValidation(

      main.getCell(
        row,
        3
      ),

      villageFormula,

      "Invalid Village",

      "Select a Village belonging to the selected Health Facility and SubCentre."

    );

  }


  /* =======================================================
     FINISH
     ======================================================= */

  main.autoFilter =
    undefined;

  subcenter.autoFilter =
    undefined;

  village.autoFilter =
    undefined;


  return workbook;

}


/* =========================================================
   READ REQUEST BODY
   ========================================================= */

async function readBody(
  req
) {

  if (
    req.body &&
    typeof req.body === "object"
  ) {

    return req.body;

  }


  const chunks = [];


  for await (
    const chunk
    of req
  ) {

    chunks.push(
      chunk
    );

  }


  const raw =
    Buffer
      .concat(chunks)
      .toString("utf8");


  if (!raw) {

    return {};

  }


  const contentType =
    String(
      req.headers[
        "content-type"
      ] || ""
    );


  if (
    contentType.includes(
      "application/json"
    )
  ) {

    return JSON.parse(
      raw
    );

  }


  const params =
    new URLSearchParams(
      raw
    );


  return Object.fromEntries(
    params.entries()
  );

}


/* =========================================================
   VERCEL HANDLER
   ========================================================= */

module.exports =
  async function handler(
    req,
    res
  ) {

    res.setHeader(
      "Access-Control-Allow-Origin",
      ALLOW_ORIGIN
    );

    res.setHeader(
      "Access-Control-Allow-Methods",
      "POST, OPTIONS"
    );

    res.setHeader(
      "Access-Control-Allow-Headers",
      "Content-Type"
    );

    res.setHeader(
      "Cache-Control",
      "no-store"
    );


    /* =====================================================
       OPTIONS
       ===================================================== */

    if (
      req.method === "OPTIONS"
    ) {

      res
        .status(204)
        .end();

      return;

    }


    /* =====================================================
       GET
       ===================================================== */

    if (
      req.method === "GET"
    ) {

      res
        .status(200)
        .json({

          ok: true,

          service:
            "SHREE RCH Template API",

          endpoint:
            "/api/template",

          methods: [
            "POST"
          ]

        });

      return;

    }


    /* =====================================================
       POST ONLY
       ===================================================== */

    if (
      req.method !== "POST"
    ) {

      res
        .status(405)
        .json({

          ok: false,

          error:
            "Method not allowed. Use POST."

        });

      return;

    }


    try {

      /* ===================================================
         BODY
         =================================================== */

      const body =
        await readBody(
          req
        );


      /* ===================================================
         TYPE
         =================================================== */

      const type =
        clean(
          body.typData
        )
        .toLowerCase();


      if (
        type !== "mother" &&
        type !== "child"
      ) {

        throw new Error(
          "typData must be 'mother' or 'child'."
        );

      }


      /* ===================================================
         LOCATION DATA
         =================================================== */

      const locationData =
        decodeLocationData(
          body.baseData
        );


      /* ===================================================
         WORKBOOK
         =================================================== */

      const workbook =
        createWorkbook(

          type,

          locationData

        );


      /* ===================================================
         XLSX
         =================================================== */

      const buffer =
        await workbook.xlsx
          .writeBuffer();


      const filename =

        type === "child"

          ? "SHREE_RCH_Child_Formate.xlsx"

          : "SHREE_RCH_Mother_Formate.xlsx";


      /* ===================================================
         RESPONSE
         =================================================== */

      res.status(200);


      res.setHeader(

        "Content-Type",

        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

      );


      res.setHeader(

        "Content-Disposition",

        `attachment; filename="${filename}"`

      );


      res.send(

        Buffer.from(
          buffer
        )

      );

    }
    catch (
      error
    ) {

      console.error(
        "SHREE RCH Template API:",
        error
      );


      res
        .status(400)
        .json({

          ok: false,

          error:
            error.message ||
            "Template generation failed."

        });

    }

  };
