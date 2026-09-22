import JSZip from "jszip";

const HEADERS = [
  "Referencia de fila",
  "ID HUB",
  "CÃ³digo del BPIP",
  "Estado",
  "Estado (PREVIO, REVISAR)",
  "Etapa",
  "Nombre corto",
  "Nombre del proyecto",
  "Â¿Pertenece a algÃºn programa?",
  "Programa vinculado",
  "Planes estratÃ©gicos vinculados (siglas)",
  "InstituciÃ³n",
  "UbicaciÃ³n polÃ­tico-administrativa (provincia, cantÃ³n)",
  "Ruta Nacional",
  "InversiÃ³n por proyecto (Millones USD)",
  "Coordenada Este (CRTM 05)",
  "Coordenada Norte (CRTM 05) ",
  "Tipo de Obra",
  "Sub tipo de obra",
  "IdentificaciÃ³n de Fuentes y Modalidades de Financiamiento",
  "Otras posibles formas de financiamiento",
  "Comentarios adicionales sobre el financiamiento",
  "Estado revisiÃ³n",
  "Fecha actualizaciÃ³n",
  "Observaciones",
  "Completo",
  "Estado HUB"
];

const INSTITUTION_CODES = {
  MOPT: "M",
  CONAVI: "V",
  COSEVI: "S",
  INCOFER: "F",
  DGAC: "D",
  INCOP: "I",
  JAPDEVA: "J",
  AYA: "A",
  CTP: "T",
  CNC: "C"
};

const HUB_ALLOWED_ORIGIN = "https://cmariam10.github.io";

function corsHeaders() {
  return {
    "Access-Control-Allow-Origin": HUB_ALLOWED_ORIGIN,
    "Access-Control-Allow-Methods": "GET, POST, PUT, DELETE, OPTIONS",
    "Access-Control-Allow-Headers": "Authorization, Content-Type, X-HUB-User, X-HUB-Role",
    "Access-Control-Max-Age": "86400",
    "Vary": "Origin"
  };
}

function json(data, status = 200) {
  return new Response(JSON.stringify(data, null, 2), {
    status,
    headers: {
      "Content-Type": "application/json; charset=utf-8",
      ...corsHeaders()
    }
  });
}

function decodeXml(text = "") {
  return String(text)
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'");
}

function encodeXml(text = "") {
  return String(text)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

async function getGraphToken(env) {
  // El Tenant ID y Client ID no son secretos. Se fijan aquí para evitar que
  // caracteres invisibles en bindings de producción alteren la URL OAuth.
  const tenantId = "b3eacfd6-2fae-4231-8cc6-ad611b3240f1";
  const clientId = "3a5ead05-e9e6-477b-9a63-24ef1f9fefd5";
  const clientSecret = String(env.MS_CLIENT_SECRET || "").trim();

  if (!clientSecret) {
    const error = new Error("MS_CLIENT_SECRET no está configurado en el Worker.");
    error.statusCode = 500;
    throw error;
  }

  const tokenUrl = new URL(
    `/b3eacfd6-2fae-4231-8cc6-ad611b3240f1/oauth2/v2.0/token`,
    "https://login.microsoftonline.com"
  );

  const body = new URLSearchParams();
  body.set("client_id", clientId);
  body.set("client_secret", clientSecret);
  body.set("scope", "https://graph.microsoft.com/.default");
  body.set("grant_type", "client_credentials");

  const response = await fetch(tokenUrl.toString(), {
    method: "POST",
    headers: {
      "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8"
    },
    body: body.toString()
  });

  const responseText = await response.text();
  let data = {};

  try {
    data = responseText ? JSON.parse(responseText) : {};
  } catch {
    const error = new Error(
      `Microsoft OAuth devolvió una respuesta no JSON. HTTP ${response.status}. ` +
      `URL usada: ${tokenUrl.toString()}. Respuesta: ${responseText.slice(0, 500)}`
    );
    error.statusCode = 502;
    throw error;
  }

  if (!response.ok || !data.access_token) {
    const error = new Error(
      data.error_description ||
      data.error ||
      `No fue posible obtener el token de Microsoft. HTTP ${response.status}.`
    );
    error.statusCode = response.status || 502;
    throw error;
  }

  return data.access_token;
}

function itemBaseUrl(env) {
  return (
    `https://graph.microsoft.com/v1.0/users/` +
    `35c36022-9fcb-4c9a-99a3-43f56441e7b6` +
    `/drive/items/01GCSG2KWKCADGQAQZBVHLDEWIQHZSMGUE`
  );
}

async function getMetadata(env, token) {
  const response = await fetch(
    itemBaseUrl(env) +
      "?$select=id,name,size,eTag,lastModifiedDateTime",
    {
      headers: {
        Authorization: `Bearer ${token}`
      }
    }
  );

  const text = await response.text();

  if (!response.ok) {
    throw new Error(
      `No se pudo obtener metadata. Graph ${response.status}: ${text}`
    );
  }

  return JSON.parse(text);
}

async function downloadExcel(env, token) {
  const response = await fetch(
    itemBaseUrl(env) + "/content",
    {
      headers: {
        Authorization: `Bearer ${token}`
      },
      redirect: "follow"
    }
  );

  if (!response.ok) {
    throw new Error(
      `No se pudo descargar el Excel. Graph ${response.status}`
    );
  }

  return await response.arrayBuffer();
}

function readSharedStrings(xml) {
  const strings = [];
  const matches = xml.match(/<si[\s\S]*?<\/si>/g) || [];

  for (const si of matches) {
    const parts = [];
    const regex = /<t(?:\s[^>]*)?>([\s\S]*?)<\/t>/g;
    let match;

    while ((match = regex.exec(si)) !== null) {
      parts.push(decodeXml(match[1]));
    }

    strings.push(parts.join(""));
  }

  return strings;
}

function columnNumber(reference) {
  const letters =
    (String(reference).match(/[A-Z]+/i) || ["A"])[0].toUpperCase();

  let result = 0;

  for (const letter of letters) {
    result = result * 26 + letter.charCodeAt(0) - 64;
  }

  return result;
}

function columnLetters(number) {
  let result = "";

  while (number > 0) {
    number--;
    result =
      String.fromCharCode(65 + (number % 26)) +
      result;
    number = Math.floor(number / 26);
  }

  return result;
}

function cellValue(attributes, content, sharedStrings) {
  const type =
    attributes.match(/\bt="([^"]+)"/)?.[1] || "";

  if (type === "inlineStr") {
    const parts = [];
    const regex = /<t(?:\s[^>]*)?>([\s\S]*?)<\/t>/g;
    let match;

    while ((match = regex.exec(content)) !== null) {
      parts.push(decodeXml(match[1]));
    }

    return parts.join("");
  }

  const value =
    content.match(/<v>([\s\S]*?)<\/v>/)?.[1];

  if (value === undefined) {
    return "";
  }

  const decoded = decodeXml(value);

  return type === "s"
    ? sharedStrings[Number(decoded)] ?? decoded
    : decoded;
}

function readRows(sheetXml, sharedStrings) {
  const rows = [];
  const matches =
    sheetXml.match(/<row\b[\s\S]*?<\/row>/g) || [];

  for (const rowXml of matches) {
    const rowNumber =
      Number(
        rowXml.match(/<row[^>]*\br="(\d+)"/)?.[1]
      ) || rows.length + 1;

    const cells = {};

    const regex =
      /<c\b([^>]*)>([\s\S]*?)<\/c>/g;

    let match;

    while ((match = regex.exec(rowXml)) !== null) {
      const reference =
        match[1].match(/\br="([^"]+)"/)?.[1];

      if (!reference) continue;

      cells[columnNumber(reference)] =
        cellValue(
          match[1],
          match[2],
          sharedStrings
        );
    }

    rows.push({
      rowNumber,
      cells
    });
  }

  return rows;
}

function rowsToProjects(rows) {
  if (!rows.length) return [];

  const headerRow = rows[0];
  const maxColumn =
    Math.max(
      0,
      ...Object.keys(headerRow.cells).map(Number)
    );

  const headers = [];

  for (let c = 1; c <= maxColumn; c++) {
    headers.push(
      String(headerRow.cells[c] ?? "").trim()
    );
  }

  return rows.slice(1).map(row => {
    const project = {
      excelRow: row.rowNumber
    };

    let hasData = false;

    for (let c = 1; c <= headers.length; c++) {
      if (!headers[c - 1]) continue;

      const value =
        row.cells[c] ?? "";

      project[headers[c - 1]] = value;

      if (String(value).trim()) {
        hasData = true;
      }
    }

    return hasData ? project : null;
  }).filter(Boolean);
}

function generateId(projects, institution) {
  const normalized =
    String(institution || "")
      .trim()
      .toUpperCase();

  const code =
    INSTITUTION_CODES[normalized];

  if (!code) {
    throw new Error(
      `InstituciÃ³n no reconocida para generar ID HUB: ${institution}`
    );
  }

  const year =
    String(new Date().getFullYear()).slice(-2);

  const prefix =
    `${code}-${year}-`;

  let maximum = 0;

  for (const project of projects) {
    const id =
      String(project["ID HUB"] || "").trim();

    if (!id.startsWith(prefix)) continue;

    const consecutive =
      Number(id.substring(prefix.length));

    if (
      Number.isInteger(consecutive) &&
      consecutive > maximum
    ) {
      maximum = consecutive;
    }
  }

  return (
    prefix +
    String(maximum + 1).padStart(3, "0")
  );
}

function buildInlineCell(
  column,
  row,
  value,
  style = ""
) {
  const reference =
    `${column}${row}`;

  const styleAttribute =
    style ? ` s="${style}"` : "";

  return (
    `<c r="${reference}"${styleAttribute} t="inlineStr">` +
    `<is><t xml:space="preserve">${encodeXml(value)}</t></is>` +
    `</c>`
  );
}

function getStylesFromTemplateRow(
  sheetXml,
  templateRowNumber,
  numberColumns
) {
  const rowRegex =
    new RegExp(
      `<row\\b[^>]*\\br="${templateRowNumber}"[^>]*>[\\s\\S]*?<\\/row>`,
      "i"
    );

  const rowXml =
    sheetXml.match(rowRegex)?.[0] || "";

  const styles = {};

  for (let c = 1; c <= numberColumns; c++) {
    const column =
      columnLetters(c);

    const cellRegex =
      new RegExp(
        `<c\\b([^>]*)\\br="${column}${templateRowNumber}"[^>]*>`,
        "i"
      );

    const cell =
      rowXml.match(cellRegex);

    if (!cell) continue;

    const style =
      cell[1].match(/\bs="([^"]+)"/)?.[1];

    if (style) {
      styles[c] = style;
    }
  }

  return styles;
}

function updateSheetDimension(
  sheetXml,
  lastRow
) {
  return sheetXml.replace(
    /<dimension\s+ref="([^"]+)"/i,
    (full, ref) => {
      const first =
        ref.includes(":")
          ? ref.split(":")[0]
          : "A1";

      return `<dimension ref="${first}:AA${lastRow}"`;
    }
  );
}

function updateTableRange(
  tableXml,
  lastRow
) {
  return tableXml
    .replace(
      /\bref="A1:AA\d+"/i,
      `ref="A1:AA${lastRow}"`
    )
    .replace(
      /(<autoFilter\b[^>]*\bref=")A1:AA\d+(")/i,
      `$1A1:AA${lastRow}$2`
    );
}

async function loadWorkbook(env, token) {
  const buffer =
    await downloadExcel(env, token);

  const zip =
    await JSZip.loadAsync(buffer);

  let sharedStrings = [];

  const sharedFile =
    zip.file("xl/sharedStrings.xml");

  if (sharedFile) {
    sharedStrings =
      readSharedStrings(
        await sharedFile.async("text")
      );
  }

  const sheetFile =
    zip.file("xl/worksheets/sheet1.xml");

  if (!sheetFile) {
    throw new Error(
      "No se encontrÃ³ sheet1.xml."
    );
  }

  const sheetXml =
    await sheetFile.async("text");

  const rows =
    readRows(
      sheetXml,
      sharedStrings
    );

  return {
    buffer,
    zip,
    sharedStrings,
    sheetXml,
    rows,
    projects:
      rowsToProjects(rows)
  };
}


function rootChildrenUrl(env) {
  return (
    `https://graph.microsoft.com/v1.0/users/` +
    `35c36022-9fcb-4c9a-99a3-43f56441e7b6` +
    `/drive/items/01GCSG2KRKAQWQFNVDUZF27DQ7K7VHC23W/children`
  );
}

function sanitizeOneDriveName(value = "") {
  return String(value)
    .replace(/[~"#%&*:<>?/\\{|}]/g, "-")
    .replace(/\s+/g, " ")
    .replace(/[. ]+$/g, "")
    .trim();
}

async function findProjectFolder(env, token, folderName) {
  let url =
    rootChildrenUrl(env) +
    "?$select=id,name,folder,webUrl&$top=200";

  while (url) {
    const response = await fetch(url, {
      headers: { Authorization: `Bearer ${token}` }
    });

    const responseText = await response.text();
    let data = {};

    try {
      data = responseText ? JSON.parse(responseText) : {};
    } catch {
      const error = new Error(
        `Microsoft Graph devolvió una respuesta no JSON al consultar las carpetas. HTTP ${response.status}: ${responseText.slice(0, 500)}`
      );
      error.statusCode = 502;
      throw error;
    }

    if (!response.ok) {
      const error = new Error(
        data?.error?.message ||
        `No se pudieron consultar las carpetas. Graph ${response.status}: ${responseText}`
      );
      error.statusCode = response.status;
      throw error;
    }

    const found = (data.value || []).find(
      item =>
        item.folder &&
        String(item.name || "").toLowerCase() ===
          folderName.toLowerCase()
    );

    if (found) return found;
    url = data["@odata.nextLink"] || "";
  }

  return null;
}

async function createProjectFolder(env, token, idHub, projectName) {
  const safeProjectName = sanitizeOneDriveName(projectName);

  if (!safeProjectName) {
    throw new Error(
      "No se puede crear la carpeta porque el Nombre del proyecto está vacío."
    );
  }

  const folderName =
    `${sanitizeOneDriveName(idHub)} - ${safeProjectName}`;

  const existing =
    await findProjectFolder(env, token, folderName);

  if (existing) {
    return {
      created: false,
      id: existing.id,
      name: existing.name,
      webUrl: existing.webUrl || ""
    };
  }

  const response = await fetch(
    rootChildrenUrl(env),
    {
      method: "POST",
      headers: {
        Authorization: `Bearer ${token}`,
        "Content-Type": "application/json"
      },
      body: JSON.stringify({
        name: folderName,
        folder: {},
        "@microsoft.graph.conflictBehavior": "fail"
      })
    }
  );

  const responseText = await response.text();
  let data = {};

  try {
    data = responseText ? JSON.parse(responseText) : {};
  } catch {}

  if (!response.ok) {
    if (response.status === 409) {
      const existingAfterConflict =
        await findProjectFolder(env, token, folderName);

      if (existingAfterConflict) {
        return {
          created: false,
          id: existingAfterConflict.id,
          name: existingAfterConflict.name,
          webUrl: existingAfterConflict.webUrl || ""
        };
      }
    }

    throw new Error(
      data?.error?.message ||
      `No se pudo crear la carpeta del proyecto. Graph ${response.status}: ${responseText}`
    );
  }

  return {
    created: true,
    id: data.id || "",
    name: data.name || folderName,
    webUrl: data.webUrl || ""
  };
}


function repairMojibake(value = "") {
  let text = String(value);

  // Corrige variantes típicas UTF-8 interpretadas como Windows-1252/Latin-1.
  const replacements = {
    "Ã¡": "á", "Ã©": "é", "Ã­": "í", "Ã³": "ó", "Ãº": "ú",
    "Ã": "Á", "Ã‰": "É", "Ã": "Í", "Ã“": "Ó", "Ãš": "Ú",
    "Ã±": "ñ", "Ã‘": "Ñ", "Â¿": "¿", "Â¡": "¡"
  };

  for (const [bad, good] of Object.entries(replacements)) {
    text = text.split(bad).join(good);
  }

  return text;
}

function normalizeFieldName(value = "") {
  return repairMojibake(value)
    .normalize("NFC")
    .trim()
    .toLowerCase();
}

function getInputValue(input, expectedHeader) {
  if (
    Object.prototype.hasOwnProperty.call(
      input,
      expectedHeader
    )
  ) {
    return input[expectedHeader];
  }

  const expected =
    normalizeFieldName(expectedHeader);

  for (const [key, value] of Object.entries(input || {})) {
    if (normalizeFieldName(key) === expected) {
      return value;
    }
  }

  return undefined;
}

function hasInputField(input, expectedHeader) {
  if (
    Object.prototype.hasOwnProperty.call(
      input,
      expectedHeader
    )
  ) {
    return true;
  }

  const expected =
    normalizeFieldName(expectedHeader);

  return Object.keys(input || {}).some(
    key => normalizeFieldName(key) === expected
  );
}

async function registerProjectOnce(
  env,
  input
) {
  const token =
    await getGraphToken(env);

  const metadata =
    await getMetadata(env, token);

  const workbook =
    await loadWorkbook(env, token);

  const institution =
    String(
      getInputValue(input, "InstituciÃ³n") || ""
    ).trim();

  if (!institution) {
    throw new Error(
      "InstituciÃ³n es obligatoria."
    );
  }

  const projectName =
    String(
      getInputValue(input, "Nombre del proyecto") || ""
    ).trim();

  if (!projectName) {
    throw new Error(
      "Nombre del proyecto es obligatorio."
    );
  }

  const idHub =
    generateId(
      workbook.projects,
      institution
    );

  const existingRows =
    workbook.rows
      .map(row => row.rowNumber)
      .filter(Number.isFinite);

  const lastRow =
    Math.max(...existingRows);

  const newRowNumber =
    lastRow + 1;

  // El consecutivo de "Referencia de fila" se calcula directamente
  // desde la columna A de TODAS las filas físicas del Excel. Así también
  // cuenta proyectos Eliminados y evita reutilizar referencias.
  const maxReference =
    workbook.rows.reduce(
      (max, row) => {
        // La fila 1 es el encabezado.
        if (row.rowNumber === 1) {
          return max;
        }

        const value =
          Number(
            String(row.cells[1] ?? "").trim()
          );

        return Number.isFinite(value)
          ? Math.max(max, value)
          : max;
      },
      0
    );

  const now =
    new Date().toISOString();

  const newProject = {};

  for (const header of HEADERS) {
    const value =
      getInputValue(input, header);

    newProject[header] =
      value ?? "";
  }

  newProject["Referencia de fila"] =
    String(maxReference + 1);

  newProject["ID HUB"] =
    idHub;

  newProject["Fecha actualizaciÃ³n"] =
    now;

  newProject["Estado HUB"] =
    "Incluido";

  const styles =
    getStylesFromTemplateRow(
      workbook.sheetXml,
      lastRow,
      HEADERS.length
    );

  let cellsXml = "";

  for (
    let c = 1;
    c <= HEADERS.length;
    c++
  ) {
    const header =
      HEADERS[c - 1];

    cellsXml +=
      buildInlineCell(
        columnLetters(c),
        newRowNumber,
        String(
          newProject[header] ?? ""
        ),
        styles[c] || ""
      );
  }

  const newRowXml =
    `<row r="${newRowNumber}">` +
    cellsXml +
    `</row>`;

  let modifiedSheetXml =
    workbook.sheetXml.replace(
      /<\/sheetData>/i,
      `${newRowXml}</sheetData>`
    );

  modifiedSheetXml =
    updateSheetDimension(
      modifiedSheetXml,
      newRowNumber
    );

  workbook.zip.file(
    "xl/worksheets/sheet1.xml",
    modifiedSheetXml
  );

  // Si existe una tabla de Excel,
  // ampliamos su rango para incluir la fila nueva.
  const tableFiles =
    Object.keys(workbook.zip.files)
      .filter(name =>
        /^xl\/tables\/table\d+\.xml$/i
          .test(name)
      );

  for (const tablePath of tableFiles) {
    const tableFile =
      workbook.zip.file(tablePath);

    if (!tableFile) continue;

    let tableXml =
      await tableFile.async("text");

    if (/ref="A1:AA\d+"/i.test(tableXml)) {
      tableXml =
        updateTableRange(
          tableXml,
          newRowNumber
        );

      workbook.zip.file(
        tablePath,
        tableXml
      );
    }
  }

  const output =
    await workbook.zip.generateAsync({
      type: "arraybuffer",
      compression: "DEFLATE"
    });

  const headers = {
    Authorization:
      `Bearer ${token}`,

    "Content-Type":
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  };

  if (metadata.eTag) {
    headers["If-Match"] =
      metadata.eTag;
  }

  const upload =
    await fetch(
      itemBaseUrl(env) + "/content",
      {
        method: "PUT",
        headers,
        body: output
      }
    );

  const uploadText =
    await upload.text();

  if (!upload.ok) {
    const error = new Error(
      `No se pudo guardar el Excel. Graph ${upload.status}: ${uploadText}`
    );
    error.statusCode = upload.status;
    throw error;
  }

  let carpeta = null;
  let advertenciaCarpeta = "";

  try {
    carpeta =
      await createProjectFolder(
        env,
        token,
        idHub,
        projectName
      );
  } catch (folderError) {
    // El proyecto ya quedó registrado en Excel.
    // No lo duplicamos ni lo revertimos: informamos el fallo
    // para poder reintentar únicamente la carpeta.
    advertenciaCarpeta =
      folderError?.message ||
      String(folderError);
  }

  return {
    ok: true,
    message:
      advertenciaCarpeta
        ? "Proyecto registrado en Excel, pero la carpeta de OneDrive requiere reintento."
        : "Proyecto registrado correctamente.",
    idHub,
    excelRow:
      newRowNumber,
    referenciaFila:
      newProject[
        "Referencia de fila"
      ],
    estadoHub:
      "Incluido",
    fechaActualizacion:
      now,
    carpetaOneDrive:
      carpeta,
    advertenciaCarpeta:
      advertenciaCarpeta || null,
    proyecto:
      newProject
  };
}


// Reintento seguro ante conflictos de concurrencia del Excel (HTTP 412).
async function registerProject(env, input) {
  const maxAttempts = 4;
  let lastError = null;
  for (let attempt = 1; attempt <= maxAttempts; attempt++) {
    try {
      return await registerProjectOnce(env, input);
    } catch (error) {
      lastError = error;
      const status = Number(error?.statusCode || 0);
      const conflict = status === 412 || /Graph\s+412|ETag does not match/i.test(String(error?.message || error));
      if (!conflict || attempt === maxAttempts) throw error;
      await new Promise(resolve => setTimeout(resolve, 250 * attempt));
    }
  }
  throw lastError || new Error("No fue posible registrar el proyecto después de reintentar el Excel.");
}


function findRowXmlByNumber(sheetXml, rowNumber) {
  const rowRegex = new RegExp(
    `<row\\b[^>]*\\br="${rowNumber}"[^>]*>[\\s\\S]*?<\\/row>`,
    "i"
  );
  return sheetXml.match(rowRegex)?.[0] || "";
}

function getCellStyle(cellXml) {
  if (!cellXml) return "";
  const openTag = cellXml.match(/^<c\b([^>]*)>/i);
  const attributes = openTag ? openTag[1] : "";
  return attributes.match(/\bs="([^"]+)"/i)?.[1] || "";
}

function replaceOrAddInlineCell(rowXml, column, rowNumber, value) {
  const reference = `${column}${rowNumber}`;
  const cellRegex = new RegExp(
    `<c\\b[^>]*\\br="${reference}"[^>]*>[\\s\\S]*?<\\/c>`,
    "i"
  );

  const existingCell = rowXml.match(cellRegex)?.[0] || "";
  const style = getCellStyle(existingCell);
  const newCell = buildInlineCell(
    column,
    rowNumber,
    String(value ?? ""),
    style
  );

  if (existingCell) {
    return rowXml.replace(cellRegex, newCell);
  }

  const targetColumnNumber = columnNumber(column);
  const cellRegexGlobal =
    /<c\b[^>]*\br="([A-Z]+)\d+"[^>]*>[\s\S]*?<\/c>/gi;

  let match;
  let insertBefore = null;

  while ((match = cellRegexGlobal.exec(rowXml)) !== null) {
    if (columnNumber(match[1]) > targetColumnNumber) {
      insertBefore = match[0];
      break;
    }
  }

  if (insertBefore) {
    return rowXml.replace(insertBefore, newCell + insertBefore);
  }

  return rowXml.replace(/<\/row>\s*$/i, `${newCell}</row>`);
}

async function updateProject(env, idHub, input) {
  const cleanId = String(idHub || "").trim();

  if (!cleanId) {
    const error = new Error("ID HUB es obligatorio.");
    error.statusCode = 400;
    throw error;
  }

  const token = await getGraphToken(env);
  const metadata = await getMetadata(env, token);
  const workbook = await loadWorkbook(env, token);

  const project = workbook.projects.find(
    item => String(item["ID HUB"] || "").trim() === cleanId
  );

  if (!project) {
    const error = new Error(`No se encontró el proyecto ${cleanId}.`);
    error.statusCode = 404;
    throw error;
  }

  const rowNumber = Number(project.excelRow);

  if (!Number.isFinite(rowNumber)) {
    throw new Error(
      `El proyecto ${cleanId} no tiene una fila de Excel válida.`
    );
  }

  const originalRowXml = findRowXmlByNumber(
    workbook.sheetXml,
    rowNumber
  );

  if (!originalRowXml) {
    throw new Error(
      `No se encontró la fila ${rowNumber} en sheet1.xml.`
    );
  }

  const protectedFields = new Set([
    "Referencia de fila",
    "ID HUB",
    "Estado HUB",
    "Fecha actualización"
  ]);

  const changes = [];
  let modifiedRowXml = originalRowXml;

  for (const header of HEADERS) {
    if (protectedFields.has(header)) continue;

    if (!hasInputField(input, header)) {
      continue;
    }

    const oldValue = String(project[header] ?? "");
    const newValue = String(
      getInputValue(input, header) ?? ""
    );

    if (oldValue === newValue) continue;

    const columnIndex = HEADERS.indexOf(header) + 1;
    const column = columnLetters(columnIndex);

    modifiedRowXml = replaceOrAddInlineCell(
      modifiedRowXml,
      column,
      rowNumber,
      newValue
    );

    changes.push({
      campo: header,
      anterior: oldValue,
      nuevo: newValue
    });
  }

  if (changes.length === 0) {
    return {
      ok: true,
      message: "No había cambios para guardar.",
      idHub: cleanId,
      excelRow: rowNumber,
      cambios: []
    };
  }

  const now = new Date().toISOString();
  const dateColumn = columnLetters(
    HEADERS.indexOf("Fecha actualización") + 1
  );

  modifiedRowXml = replaceOrAddInlineCell(
    modifiedRowXml,
    dateColumn,
    rowNumber,
    now
  );

  const modifiedSheetXml = workbook.sheetXml.replace(
    originalRowXml,
    modifiedRowXml
  );

  workbook.zip.file(
    "xl/worksheets/sheet1.xml",
    modifiedSheetXml
  );

  const output = await workbook.zip.generateAsync({
    type: "arraybuffer",
    compression: "DEFLATE"
  });

  const headers = {
    Authorization: `Bearer ${token}`,
    "Content-Type":
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  };

  if (metadata.eTag) {
    headers["If-Match"] = metadata.eTag;
  }

  const upload = await fetch(
    itemBaseUrl(env) + "/content",
    {
      method: "PUT",
      headers,
      body: output
    }
  );

  const uploadText = await upload.text();

  if (!upload.ok) {
    const error = new Error(
      `No se pudo guardar el Excel. Graph ${upload.status}: ${uploadText}`
    );
    error.statusCode = upload.status === 423 ? 423 : 500;
    throw error;
  }

  return {
    ok: true,
    message: "Proyecto actualizado correctamente.",
    idHub: cleanId,
    excelRow: rowNumber,
    fechaActualizacion: now,
    cambios: changes,
    camposProtegidos: [
      "Referencia de fila",
      "ID HUB",
      "Estado HUB"
    ]
  };
}


async function deleteProject(env, idHub) {
  const cleanId = String(idHub || "").trim();

  if (!cleanId) {
    const error = new Error("ID HUB es obligatorio.");
    error.statusCode = 400;
    throw error;
  }

  const token = await getGraphToken(env);
  const metadata = await getMetadata(env, token);
  const workbook = await loadWorkbook(env, token);

  const project = workbook.projects.find(
    item => String(item["ID HUB"] || "").trim() === cleanId
  );

  if (!project) {
    const error = new Error(`No se encontró el proyecto ${cleanId}.`);
    error.statusCode = 404;
    throw error;
  }

  const currentStatus =
    String(project["Estado HUB"] || "").trim();

  if (currentStatus.toLowerCase() === "eliminado") {
    return {
      ok: true,
      message: "El proyecto ya estaba marcado como Eliminado.",
      idHub: cleanId,
      excelRow: Number(project.excelRow),
      estadoHub: "Eliminado",
      cambios: []
    };
  }

  const rowNumber = Number(project.excelRow);

  if (!Number.isFinite(rowNumber)) {
    throw new Error(
      `El proyecto ${cleanId} no tiene una fila de Excel válida.`
    );
  }

  const originalRowXml = findRowXmlByNumber(
    workbook.sheetXml,
    rowNumber
  );

  if (!originalRowXml) {
    throw new Error(
      `No se encontró la fila ${rowNumber} en sheet1.xml.`
    );
  }

  let modifiedRowXml = originalRowXml;

  const statusColumn = columnLetters(
    HEADERS.indexOf("Estado HUB") + 1
  );

  modifiedRowXml = replaceOrAddInlineCell(
    modifiedRowXml,
    statusColumn,
    rowNumber,
    "Eliminado"
  );

  const now = new Date().toISOString();

  const dateColumn = columnLetters(
    HEADERS.indexOf("Fecha actualización") + 1
  );

  modifiedRowXml = replaceOrAddInlineCell(
    modifiedRowXml,
    dateColumn,
    rowNumber,
    now
  );

  const modifiedSheetXml = workbook.sheetXml.replace(
    originalRowXml,
    modifiedRowXml
  );

  workbook.zip.file(
    "xl/worksheets/sheet1.xml",
    modifiedSheetXml
  );

  const output = await workbook.zip.generateAsync({
    type: "arraybuffer",
    compression: "DEFLATE"
  });

  const headers = {
    Authorization: `Bearer ${token}`,
    "Content-Type":
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  };

  if (metadata.eTag) {
    headers["If-Match"] = metadata.eTag;
  }

  const upload = await fetch(
    itemBaseUrl(env) + "/content",
    {
      method: "PUT",
      headers,
      body: output
    }
  );

  const uploadText = await upload.text();

  if (!upload.ok) {
    const error = new Error(
      `No se pudo guardar el Excel. Graph ${upload.status}: ${uploadText}`
    );

    if (upload.status === 423) {
      error.statusCode = 423;
    } else if (upload.status === 412) {
      error.statusCode = 409;
    } else {
      error.statusCode = 500;
    }

    throw error;
  }

  return {
    ok: true,
    message:
      "Proyecto marcado como Eliminado correctamente.",
    idHub: cleanId,
    excelRow: rowNumber,
    estadoAnterior: currentStatus,
    estadoHub: "Eliminado",
    fechaActualizacion: now,
    cambios: [
      {
        campo: "Estado HUB",
        anterior: currentStatus,
        nuevo: "Eliminado"
      }
    ]
  };
}


async function repairReferenceFila(env, idHub) {
  const cleanId = String(idHub || "").trim();

  const token = await getGraphToken(env);
  const metadata = await getMetadata(env, token);
  const workbook = await loadWorkbook(env, token);

  const project = workbook.projects.find(
    item => String(item["ID HUB"] || "").trim() === cleanId
  );

  if (!project) {
    const error = new Error(`No se encontró el proyecto ${cleanId}.`);
    error.statusCode = 404;
    throw error;
  }

  const rowNumber = Number(project.excelRow);
  const originalRowXml =
    findRowXmlByNumber(workbook.sheetXml, rowNumber);

  if (!originalRowXml) {
    throw new Error(`No se encontró la fila ${rowNumber}.`);
  }

  const otherReferences = workbook.rows
    .filter(row => row.rowNumber !== 1 && row.rowNumber !== rowNumber)
    .map(row => Number(String(row.cells[1] ?? "").trim()))
    .filter(Number.isFinite);

  const maxOther =
    otherReferences.length
      ? Math.max(...otherReferences)
      : 0;

  const current =
    Number(String(project["Referencia de fila"] ?? "").trim());

  // Para esta reparación, asignamos un valor nuevo estrictamente mayor
  // que cualquier otra referencia existente.
  const newReference = String(
    Math.max(maxOther, Number.isFinite(current) ? current : 0) + 1
  );

  const refColumn =
    columnLetters(HEADERS.indexOf("Referencia de fila") + 1);

  let modifiedRowXml =
    replaceOrAddInlineCell(
      originalRowXml,
      refColumn,
      rowNumber,
      newReference
    );

  const now = new Date().toISOString();
  const dateColumn =
    columnLetters(HEADERS.indexOf("Fecha actualización") + 1);

  modifiedRowXml =
    replaceOrAddInlineCell(
      modifiedRowXml,
      dateColumn,
      rowNumber,
      now
    );

  const modifiedSheetXml =
    workbook.sheetXml.replace(
      originalRowXml,
      modifiedRowXml
    );

  workbook.zip.file(
    "xl/worksheets/sheet1.xml",
    modifiedSheetXml
  );

  const output =
    await workbook.zip.generateAsync({
      type: "arraybuffer",
      compression: "DEFLATE"
    });

  const headers = {
    Authorization: `Bearer ${token}`,
    "Content-Type":
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  };

  if (metadata.eTag) {
    headers["If-Match"] = metadata.eTag;
  }

  const upload =
    await fetch(
      itemBaseUrl(env) + "/content",
      {
        method: "PUT",
        headers,
        body: output
      }
    );

  const uploadText = await upload.text();

  if (!upload.ok) {
    const error = new Error(
      `No se pudo guardar el Excel. Graph ${upload.status}: ${uploadText}`
    );
    error.statusCode = upload.status === 423 ? 423 : 500;
    throw error;
  }

  return {
    ok: true,
    message: "Referencia de fila reparada correctamente.",
    idHub: cleanId,
    excelRow: rowNumber,
    referenciaAnterior:
      String(project["Referencia de fila"] ?? ""),
    referenciaNueva: newReference,
    fechaActualizacion: now
  };
}



function safeFlowFileName(idHub) {
  return String(idHub || "").trim().replace(/[~"#%&*:<>?/\\{|}]/g, "-") + ".json";
}

async function ensureWorkflowFolderApi(env, token) {
  // V20: NO usar MS_ROOT_ITEM_ID para el módulo de flujo.
  // El diagnóstico V19 demostró que ese binding llega como un carácter de control
  // (encodeURIComponent -> %16), generando /drive/items/%16/children y Graph 400.
  // Este es el ID real y ya verificado de HUB_Proyectos_MOPT.
  const workflowRootItemId = "01GCSG2KRKAQWQFNVDUZF27DQ7K7VHC23W";
  console.log("[WF_V20_ROOT]", JSON.stringify({ rootItemId: workflowRootItemId }));
  return await ensureChildFolder(
    token,
    workflowRootItemId,
    "_HUB_FLUJO"
  );
}

// V22: persistencia del workflow con nombres cortos y estables.
// El estado completo vive DENTRO de un JSON (__WF22__<ID>.json).
// Para GET /workflow se usa un marcador-resumen corto que evita descargar cientos
// de JSON y evita URLs excesivamente largas.
const WF_LEGACY_MARKER_PREFIX = "__WF18__";
const WF_DATA_PREFIX = "__WF22__";
const WF_SUMMARY_PREFIX = "__WF22S__";
// Owner fijo y verificado para workflow.
const WORKFLOW_OWNER_USER_ID = "35c36022-9fcb-4c9a-99a3-43f56441e7b6";

function wfB64Decode(value = "") {
  try {
    let s = String(value).replace(/-/g, "+").replace(/_/g, "/");
    while (s.length % 4) s += "=";
    const binary = atob(s);
    const bytes = Uint8Array.from(binary, c => c.charCodeAt(0));
    return new TextDecoder().decode(bytes);
  } catch { return ""; }
}

function workflowSafeId(projectId) {
  return String(projectId || "").trim().replace(/[^A-Za-z0-9._-]/g, "_");
}

function workflowDataName(projectId) {
  return `${WF_DATA_PREFIX}${workflowSafeId(projectId)}.json`;
}

function workflowSummaryName(payload) {
  const id = workflowSafeId(payload?.projectId);
  const step = Math.max(1, Number(payload?.step) || 1);
  const returned = payload?.returned ? 1 : 0;
  const rejected = payload?.rejected ? 1 : 0;
  return `${WF_SUMMARY_PREFIX}${id}__S${step}__R${returned}__X${rejected}.json`;
}

const WORKFLOW_STEP_OWNERS = {
  1:"Técnico / Profesional",2:"Jefatura Técnica",3:"Planeamiento",4:"Dirección / Gerencia",
  5:"Coordinador HUB",6:"Secretaría Técnica",7:"Áreas Técnicas",8:"Sistema HUB / Secretaría Técnica HUB",
  9:"Comité HUB Central",10:"Institución ejecutora",11:"Registro HUB",12:"Institución ejecutora",
  13:"Institución ejecutora",14:"Institución ejecutora",15:"Supervisor",16:"Supervisor",17:"Comité / ejecutor"
};

function workflowDerivedState(step, returned, rejected) {
  if (rejected) return "Rechazado / archivado";
  if (returned) return "Devuelto para ajustes";
  if (Number(step) === 1) return "Borrador / ajuste en curso";
  if (Number(step) === 2) return "En revisión técnica";
  if (Number(step) >= 14) return "En ejecución";
  return "En revisión";
}

function parseLegacyWorkflowMarker(item) {
  const name = String(item?.name || "");
  if (!name.startsWith(WF_LEGACY_MARKER_PREFIX) || !name.toLowerCase().endsWith(".json")) return null;
  try {
    const encoded = name.slice(WF_LEGACY_MARKER_PREFIX.length, -5);
    const c = JSON.parse(wfB64Decode(encoded));
    if (!c?.p) return null;
    return {
      projectId:String(c.p), step:Number(c.s)||1, state:String(c.t||""), owner:String(c.o||""),
      returned:!!c.r, rejected:!!c.x, updatedAt:String(c.u||item?.lastModifiedDateTime||""),
      updatedBy:String(c.b||""), historyCount:Number(c.h||0),
      _markerId:String(item?.id||""), _markerName:name, _kind:"legacy"
    };
  } catch { return null; }
}

function parseWorkflowSummary(item) {
  const name = String(item?.name || "");
  if (!name.startsWith(WF_SUMMARY_PREFIX) || !name.toLowerCase().endsWith(".json")) return null;
  const body = name.slice(WF_SUMMARY_PREFIX.length, -5);
  const m = body.match(/^(.*?)__S(\d+)__R([01])__X([01])$/);
  if (!m) return null;
  const projectId = m[1];
  const step = Number(m[2]) || 1;
  const returned = m[3] === "1";
  const rejected = m[4] === "1";
  return {
    projectId, step, returned, rejected,
    state:workflowDerivedState(step, returned, rejected),
    owner:WORKFLOW_STEP_OWNERS[step] || "",
    updatedAt:String(item?.lastModifiedDateTime||""), updatedBy:"", historyCount:0,
    _markerId:String(item?.id||""), _markerName:name, _kind:"summary"
  };
}

async function listWorkflowChildren(token, folder) {
  const owner = WORKFLOW_OWNER_USER_ID;
  console.log("[WF_V22_OWNER]", JSON.stringify({ ownerUserId:owner, folderId:String(folder?.id||"") }));
  let url = `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(owner)}/drive/items/${encodeURIComponent(folder.id)}/children?$select=id,name,file,lastModifiedDateTime&$top=200`;
  const items=[];
  while (url) {
    const data=await graphJson(token,url);
    items.push(...(data.value||[]));
    url=data["@odata.nextLink"]||"";
  }
  return items;
}

async function listWorkflowMarkers(env, token, folder) {
  const items=await listWorkflowChildren(token,folder);
  const summaries=[];
  const legacy=[];
  for (const item of items) {
    if (!item.file) continue;
    const s=parseWorkflowSummary(item); if (s) { summaries.push(s); continue; }
    const l=parseLegacyWorkflowMarker(item); if (l) legacy.push(l);
  }
  // Si ya existe resumen V22 para un proyecto, ignora su marcador V18 antiguo.
  const hasV22=new Set(summaries.map(x=>String(x.projectId)));
  return summaries.concat(legacy.filter(x=>!hasV22.has(String(x.projectId))));
}

function latestWorkflowByProject(markers) {
  const map=new Map();
  for (const wf of markers||[]) {
    const key=String(wf.projectId||"").trim(); if(!key) continue;
    const old=map.get(key);
    if(!old || String(wf.updatedAt||"") >= String(old.updatedAt||"")) map.set(key,wf);
  }
  return map;
}

async function readWorkflowDataByPath(token, folder, projectId) {
  const owner=WORKFLOW_OWNER_USER_ID;
  const name=workflowDataName(projectId);
  const url=`https://graph.microsoft.com/v1.0/users/${encodeURIComponent(owner)}/drive/items/${encodeURIComponent(folder.id)}:/${encodeURIComponent(name)}:/content`;
  const r=await fetch(url,{headers:{Authorization:`Bearer ${token}`}});
  if(r.status===404) return null;
  const t=await r.text();
  if(!r.ok){
    let d={}; try{d=t?JSON.parse(t):{}}catch{}
    const e=new Error(d?.error?.message||`No se pudo leer el flujo. Graph ${r.status}: ${t}`); e.statusCode=r.status; throw e;
  }
  try{return t?JSON.parse(t):null}catch{return null}
}

async function getProjectWorkflowApi(env,idHub){
  const cleanId=String(idHub||"").trim();
  if(!cleanId){const e=new Error("ID HUB es obligatorio.");e.statusCode=400;throw e;}
  const token=await getGraphToken(env);
  const folder=await ensureWorkflowFolderApi(env,token);
  let workflow=await readWorkflowDataByPath(token,folder,cleanId);
  if(!workflow){
    const markers=await listWorkflowMarkers(env,token,folder);
    workflow=latestWorkflowByProject(markers).get(cleanId)||null;
  }
  console.log("[WF_V22_GET]",JSON.stringify({projectId:cleanId,found:!!workflow,step:workflow?.step||null,state:workflow?.state||""}));
  return {ok:true,projectId:cleanId,workflow};
}

async function listAllWorkflowApi(env){
  const token=await getGraphToken(env);
  const folder=await ensureWorkflowFolderApi(env,token);
  const markers=await listWorkflowMarkers(env,token,folder);
  const workflows=Array.from(latestWorkflowByProject(markers).values()).map(wf=>{
    const {_markerId,_markerName,_kind,...clean}=wf; return clean;
  });
  return {ok:true,source:"worker-app-only-workflow-v22-summary",version:"workflow-v22",total:workflows.length,workflows};
}

async function saveProjectWorkflowApi(env,idHub,input){
  const cleanId=String(idHub||"").trim();
  if(!cleanId){const e=new Error("ID HUB es obligatorio.");e.statusCode=400;throw e;}
  console.log("[WF_V22_ENTER]",JSON.stringify({projectId:cleanId}));
  const token=await getGraphToken(env);
  const folder=await ensureWorkflowFolderApi(env,token);
  const owner=WORKFLOW_OWNER_USER_ID;

  // Estado previo: primero JSON V22; si aún no existe, migra desde marcador V18.
  let prior=await readWorkflowDataByPath(token,folder,cleanId);
  const existing=await listWorkflowMarkers(env,token,folder);
  if(!prior) prior=latestWorkflowByProject(existing).get(cleanId)||null;

  const movement=input?.movement||null;
  const oldHistory=Array.isArray(prior?.history)?prior.history:[];
  const history=movement?[...oldHistory,movement]:oldHistory;
  const historyCount=Math.max(Number(prior?.historyCount||0)+(movement?1:0),history.length);
  const payload={
    projectId:cleanId, step:Number(input?.step)||1, state:String(input?.state||""), owner:String(input?.owner||""),
    returned:!!input?.returned, rejected:!!input?.rejected,
    updatedAt:input?.updatedAt||new Date().toISOString(), updatedBy:String(input?.updatedBy||""),
    historyCount, history
  };

  const dataName=workflowDataName(cleanId);
  const dataUrl=`https://graph.microsoft.com/v1.0/users/${encodeURIComponent(owner)}/drive/items/${encodeURIComponent(folder.id)}:/${encodeURIComponent(dataName)}:/content`;
  console.log("[WF_V22_DATA_SAVE_START]",JSON.stringify({projectId:cleanId,fileName:dataName,step:payload.step,state:payload.state}));
  let r=await fetch(dataUrl,{method:"PUT",headers:{Authorization:`Bearer ${token}`,"Content-Type":"application/json; charset=utf-8"},body:JSON.stringify(payload,null,2)});
  let t=await r.text();
  if(!r.ok){let d={};try{d=t?JSON.parse(t):{}}catch{};const e=new Error(d?.error?.message||`No se pudo guardar el flujo. Graph ${r.status}: ${t}`);e.statusCode=r.status;throw e;}
  let saved={};try{saved=t?JSON.parse(t):{}}catch{}
  console.log("[WF_V22_DATA_SAVE_OK]",JSON.stringify({projectId:cleanId,itemId:String(saved?.id||"")}));

  // Resumen corto para sincronización masiva GET /workflow.
  const summaryName=workflowSummaryName(payload);
  const summaryUrl=`https://graph.microsoft.com/v1.0/users/${encodeURIComponent(owner)}/drive/items/${encodeURIComponent(folder.id)}:/${encodeURIComponent(summaryName)}:/content`;
  console.log("[WF_V22_SUMMARY_SAVE_START]",JSON.stringify({projectId:cleanId,summaryName}));
  r=await fetch(summaryUrl,{method:"PUT",headers:{Authorization:`Bearer ${token}`,"Content-Type":"application/json; charset=utf-8"},body:"{}"});
  t=await r.text();
  if(!r.ok){let d={};try{d=t?JSON.parse(t):{}}catch{};const e=new Error(d?.error?.message||`No se pudo guardar el resumen del flujo. Graph ${r.status}: ${t}`);e.statusCode=r.status;throw e;}
  let summarySaved={};try{summarySaved=t?JSON.parse(t):{}}catch{}
  console.log("[WF_V22_SUMMARY_SAVE_OK]",JSON.stringify({projectId:cleanId,itemId:String(summarySaved?.id||""),summaryName}));

  // Elimina resúmenes V22 anteriores del mismo proyecto y marcadores V18 heredados.
  for(const old of existing){
    if(String(old.projectId||"")!==cleanId || !old._markerId) continue;
    if(String(old._markerName||"")===summaryName) continue;
    try{
      const delUrl=`https://graph.microsoft.com/v1.0/users/${encodeURIComponent(owner)}/drive/items/${encodeURIComponent(old._markerId)}`;
      const dr=await fetch(delUrl,{method:"DELETE",headers:{Authorization:`Bearer ${token}`}});
      if(!dr.ok&&dr.status!==404) console.warn("[WF_V22_CLEAN_WARN]",cleanId,dr.status);
    }catch(e){console.warn("[WF_V22_CLEAN_WARN]",cleanId,e?.message||e);}
  }

  console.log("[WF_V22_SAVE_OK]",JSON.stringify({projectId:cleanId,step:payload.step,state:payload.state,historyCount}));
  return {ok:true,projectId:cleanId,workflow:payload};
}

async function findProjectFolderByIdHub(env, token, idHub) {
  const cleanId = String(idHub || "").trim();

  if (!cleanId) {
    const error = new Error("ID HUB es obligatorio.");
    error.statusCode = 400;
    throw error;
  }

  let url =
    rootChildrenUrl(env) +
    "?$select=id,name,folder,webUrl,lastModifiedDateTime&$top=200";

  while (url) {
    const response = await fetch(url, {
      headers: {
        Authorization: `Bearer ${token}`
      }
    });

    const responseText = await response.text();
    let data = {};

    try {
      data = responseText ? JSON.parse(responseText) : {};
    } catch {}

    if (!response.ok) {
      const error = new Error(
        data?.error?.message ||
        `No se pudieron consultar las carpetas. Graph ${response.status}: ${responseText}`
      );
      error.statusCode = response.status;
      throw error;
    }

    const prefix = `${cleanId} - `.toLowerCase();

    const found = (data.value || []).find(
      item =>
        item.folder &&
        String(item.name || "")
          .toLowerCase()
          .startsWith(prefix)
    );

    if (found) return found;

    url = data["@odata.nextLink"] || "";
  }

  return null;
}

async function listProjectDocuments(env, idHub) {
  const cleanId = String(idHub || "").trim();

  const token = await getGraphToken(env);

  const folder =
    await findProjectFolderByIdHub(
      env,
      token,
      cleanId
    );

  if (!folder) {
    const error = new Error(
      `No se encontró la carpeta de OneDrive para el proyecto ${cleanId}.`
    );
    error.statusCode = 404;
    throw error;
  }

  let url =
    `https://graph.microsoft.com/v1.0/users/` +
    `35c36022-9fcb-4c9a-99a3-43f56441e7b6` +
    `/drive/items/${encodeURIComponent(folder.id)}/children` +
    "?$select=id,name,size,file,folder,webUrl,eTag,lastModifiedDateTime,createdDateTime,parentReference,@microsoft.graph.downloadUrl&$top=200";

  const documents = [];

  while (url) {
    const response = await fetch(url, {
      headers: {
        Authorization: `Bearer ${token}`
      }
    });

    const responseText = await response.text();
    let data = {};

    try {
      data = responseText ? JSON.parse(responseText) : {};
    } catch {}

    if (!response.ok) {
      const error = new Error(
        data?.error?.message ||
        `No se pudieron consultar los documentos. Graph ${response.status}: ${responseText}`
      );
      error.statusCode = response.status;
      throw error;
    }

    for (const item of data.value || []) {
      console.log("[DOC_DIAG]", JSON.stringify({
        projectId: cleanId,
        id: String(item?.id || ""),
        name: String(item?.name || ""),
        hasFile: !!item?.file,
        hasFolder: !!item?.folder,
        hasWebUrl: !!item?.webUrl,
        hasDownloadUrl: !!item?.["@microsoft.graph.downloadUrl"],
        hasParentReference: !!item?.parentReference,
        parentDriveId: String(item?.parentReference?.driveId || ""),
        parentItemId: String(item?.parentReference?.id || "")
      }));
      // Este endpoint documental devuelve archivos.
      // Las subcarpetas se omiten para no mezclarlas con documentos.
      if (!item.file) continue;

      documents.push({
        id: item.id || "",
        name: item.name || "",
        size: Number(item.size || 0),
        mimeType: item.file?.mimeType || "",
        eTag: item.eTag || "",
        createdDateTime: item.createdDateTime || "",
        lastModifiedDateTime: item.lastModifiedDateTime || "",
        webUrl: item.webUrl || ""
      });
    }

    url = data["@odata.nextLink"] || "";
  }

  return {
    ok: true,
    idHub: cleanId,
    carpeta: {
      id: folder.id,
      name: folder.name,
      webUrl: folder.webUrl || ""
    },
    totalDocuments: documents.length,
    documents
  };
}




async function getProjectDocumentContent(env, idHub, documentId) {
  const cleanId = String(idHub || "").trim();
  const cleanDocumentId = String(documentId || "").trim();

  if (!cleanId || !cleanDocumentId) {
    const error = new Error("ID HUB y documentId son obligatorios.");
    error.statusCode = 400;
    throw error;
  }

  /*
   * V14: NO reconstruir una segunda llamada /children.
   *
   * Los logs demuestran que listProjectDocuments() funciona y que la llamada
   * duplicada construida dentro de /content es la que falla. Por tanto usamos
   * directamente la función ya probada en producción para resolver carpeta,
   * listar documentos y obtener nombre + id.
   */
  const listing = await listProjectDocuments(env, cleanId);

  const found = (listing?.documents || []).find(
    item => String(item?.id || "") === cleanDocumentId
  ) || null;

  if (!found) {
    const error = new Error(
      `No se encontró el documento ${cleanDocumentId} dentro de ${cleanId}.`
    );
    error.statusCode = 404;
    throw error;
  }

  const folderId = String(listing?.carpeta?.id || "").trim();
  if (!folderId) {
    const error = new Error("El listado no devolvió el ID de la carpeta del proyecto.");
    error.statusCode = 502;
    throw error;
  }

  const token = await getGraphToken(env);

  // Mismo propietario literal usado por listProjectDocuments(), que ya funciona.
  const owner = "35c36022-9fcb-4c9a-99a3-43f56441e7b6";
  const encodedName = encodeURIComponent(String(found.name || ""));

  const graphContentUrl =
    `https://graph.microsoft.com/v1.0/users/` +
    owner +
    `/drive/items/${encodeURIComponent(folderId)}:/${encodedName}:/content`;

  console.log("[DOC_V14_REUSE_LIST]", JSON.stringify({
    projectId: cleanId,
    documentId: cleanDocumentId,
    folderId,
    name: String(found.name || ""),
    listingReused: true
  }));

  const graphResponse = await fetch(graphContentUrl, {
    method: "GET",
    headers: { Authorization: `Bearer ${token}` },
    redirect: "manual"
  });

  if (graphResponse.status >= 300 && graphResponse.status < 400) {
    const location = String(graphResponse.headers.get("Location") || "").trim();

    console.log("[DOC_V14_REDIRECT]", JSON.stringify({
      status: graphResponse.status,
      hasLocation: !!location
    }));

    if (!location) {
      const error = new Error(
        `Graph respondió ${graphResponse.status} sin Location para el documento.`
      );
      error.statusCode = 502;
      throw error;
    }

    const contentResponse = await fetch(location, {
      method: "GET",
      redirect: "follow"
    });

    if (!contentResponse.ok) {
      const responseText = await contentResponse.text();
      const error = new Error(
        `OneDrive no pudo entregar el archivo. HTTP ${contentResponse.status}: ${responseText.slice(0,700)}`
      );
      error.statusCode = contentResponse.status;
      throw error;
    }

    const contentType =
      contentResponse.headers.get("Content-Type") ||
      found?.mimeType ||
      "application/octet-stream";

    const safeName = String(found.name || "documento").replace(/[\r\n"]/g, "_");

    return new Response(contentResponse.body, {
      status: 200,
      headers: {
        "Content-Type": contentType,
        "Content-Disposition": `inline; filename="${safeName}"`,
        "Cache-Control": "private, no-store",
        ...corsHeaders()
      }
    });
  }

  if (!graphResponse.ok) {
    const responseText = await graphResponse.text();
    const error = new Error(
      `La lista funcionó, pero Graph rechazó la descarga por ruta. HTTP ${graphResponse.status}: ${responseText.slice(0,700)}`
    );
    error.statusCode = graphResponse.status;
    throw error;
  }

  const contentType =
    graphResponse.headers.get("Content-Type") ||
    found?.mimeType ||
    "application/octet-stream";
  const safeName = String(found.name || "documento").replace(/[\r\n"]/g, "_");

  return new Response(graphResponse.body, {
    status: 200,
    headers: {
      "Content-Type": contentType,
      "Content-Disposition": `inline; filename="${safeName}"`,
      "Cache-Control": "private, no-store",
      ...corsHeaders()
    }
  });
}


async function createProjectDocumentEditLink(env, idHub, documentId, request) {
  const cleanId = String(idHub || "").trim();
  const cleanDocumentId = String(documentId || "").trim();

  if (!cleanId || !cleanDocumentId) {
    const error = new Error("ID HUB y documentId son obligatorios.");
    error.statusCode = 400;
    throw error;
  }

  // Reutiliza la función de listado que ya está comprobada en producción.
  const listing = await listProjectDocuments(env, cleanId);
  const found = (listing?.documents || []).find(
    item => String(item?.id || "") === cleanDocumentId
  ) || null;

  if (!found) {
    const error = new Error(`No se encontró el documento ${cleanDocumentId} dentro de ${cleanId}.`);
    error.statusCode = 404;
    throw error;
  }

  const folderId = String(listing?.carpeta?.id || "").trim();
  if (!folderId) {
    const error = new Error("El listado no devolvió el ID de la carpeta del proyecto.");
    error.statusCode = 502;
    throw error;
  }

  const token = await getGraphToken(env);
  const owner = "35c36022-9fcb-4c9a-99a3-43f56441e7b6";
  const encodedName = encodeURIComponent(String(found.name || ""));

  // Direccionamiento por ruta relativa a la carpeta: evita la consulta
  // individual del DriveItem que falló anteriormente en este entorno.
  const createLinkUrl =
    `https://graph.microsoft.com/v1.0/users/${owner}` +
    `/drive/items/${encodeURIComponent(folderId)}:/${encodedName}:/createLink`;

  const expiration = new Date(Date.now() + 60 * 60 * 1000).toISOString();

  const graphResponse = await fetch(createLinkUrl, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json"
    },
    body: JSON.stringify({
      type: "edit",
      scope: "anonymous",
      expirationDateTime: expiration
    })
  });

  const responseText = await graphResponse.text();
  let data = {};
  try { data = responseText ? JSON.parse(responseText) : {}; } catch {}

  if (!graphResponse.ok) {
    const graphMessage =
      data?.error?.message ||
      responseText.slice(0, 900) ||
      `HTTP ${graphResponse.status}`;

    const error = new Error(
      `No se pudo crear el vínculo de edición en Microsoft 365. Graph ${graphResponse.status}: ${graphMessage}`
    );
    error.statusCode = graphResponse.status;
    throw error;
  }

  const webUrl = String(data?.link?.webUrl || "").trim();
  if (!webUrl) {
    const error = new Error("Microsoft Graph creó el permiso, pero no devolvió link.webUrl.");
    error.statusCode = 502;
    throw error;
  }

  console.log("[DOC_EDIT_V15]", JSON.stringify({
    projectId: cleanId,
    documentId: cleanDocumentId,
    name: String(found.name || ""),
    type: String(data?.link?.type || "edit"),
    scope: String(data?.link?.scope || "anonymous"),
    hasWebUrl: true
  }));

  // Registrar la apertura de edición sin guardar el vínculo temporal.
  try {
    await safeWriteProjectTrace(
      env, token, request, {
        idHub: cleanId,
        action: "edit-online",
        document: {
          id: cleanDocumentId,
          name: String(found.name || "")
        },
        details: {
          provider: "Microsoft 365",
          mode: "edit",
          scope: String(data?.link?.scope || "anonymous")
        }
      }
    );
  } catch (traceError) {
    console.warn("No se pudo registrar trazabilidad de edición:", traceError?.message || traceError);
  }

  return {
    ok: true,
    idHub: cleanId,
    documentId: cleanDocumentId,
    name: String(found.name || ""),
    editUrl: webUrl,
    expiresAt: expiration
  };
}

function getTraceActor(request) {
  const user =
    String(request?.headers?.get("X-HUB-User") || "").trim() ||
    "No identificado";

  const role =
    String(request?.headers?.get("X-HUB-Role") || "").trim() ||
    "No informado";

  return {
    user,
    role,
    identityStatus:
      user === "No identificado"
        ? "pending-authentication"
        : "reported-by-client"
  };
}

async function graphJson(token, url, options = {}) {
  const response = await fetch(url, {
    ...options,
    headers: {
      Authorization: `Bearer ${token}`,
      ...(options.headers || {})
    }
  });

  const responseText = await response.text();
  let data = {};

  try {
    data = responseText ? JSON.parse(responseText) : {};
  } catch {}

  if (!response.ok) {
    console.error("[GRAPH_ERROR_V19]", JSON.stringify({ method: String(options?.method || "GET"), status: response.status, url: String(url), body: responseText.slice(0, 700) }));
    const error = new Error(
      data?.error?.message ||
      `Microsoft Graph ${response.status}: ${responseText}`
    );
    error.statusCode = response.status;
    throw error;
  }

  return data;
}

async function ensureChildFolder(token, parentId, folderName) {
  let childrenUrl =
    `https://graph.microsoft.com/v1.0/users/` +
    `35c36022-9fcb-4c9a-99a3-43f56441e7b6` +
    `/drive/items/${encodeURIComponent(parentId)}/children` +
    "?$select=id,name,folder&$top=200";

  while (childrenUrl) {
    const data = await graphJson(token, childrenUrl);

    const found = (data.value || []).find(
      item =>
        item.folder &&
        String(item.name || "").toLowerCase() ===
          String(folderName || "").toLowerCase()
    );

    if (found) return found;

    childrenUrl = data["@odata.nextLink"] || "";
  }

  const createUrl =
    `https://graph.microsoft.com/v1.0/users/` +
    `35c36022-9fcb-4c9a-99a3-43f56441e7b6` +
    `/drive/items/${encodeURIComponent(parentId)}/children`;

  return await graphJson(token, createUrl, {
    method: "POST",
    headers: {
      "Content-Type": "application/json; charset=utf-8"
    },
    body: JSON.stringify({
      name: folderName,
      folder: {},
      "@microsoft.graph.conflictBehavior": "fail"
    })
  });
}

async function writeDocumentTrace(
  env,
  token,
  request,
  {
    idHub,
    action,
    folder,
    document,
    versioning = null
  }
) {
  const traceRoot = await ensureChildFolder(
    token,
    "01GCSG2KRKAQWQFNVDUZF27DQ7K7VHC23W",
    "_HUB_TRAZABILIDAD"
  );

  let projectTraceFolder;

  try {
    projectTraceFolder = await ensureChildFolder(
      token,
      traceRoot.id,
      String(idHub || "").trim()
    );
  } catch (error) {
    // Si dos solicitudes crean la misma carpeta al mismo tiempo,
    // volver a consultarla una vez.
    if (error.statusCode === 409) {
      projectTraceFolder = await ensureChildFolder(
        token,
        traceRoot.id,
        String(idHub || "").trim()
      );
    } else {
      throw error;
    }
  }

  const now = new Date();
  const iso = now.toISOString();
  const safeTimestamp = iso.replace(/[:.]/g, "-");
  const eventId =
    typeof crypto !== "undefined" && crypto.randomUUID
      ? crypto.randomUUID()
      : `${Date.now()}-${Math.random().toString(16).slice(2)}`;

  const actor = getTraceActor(request);

  const trace = {
    schemaVersion: 1,
    eventId,
    eventType: "document",
    action,
    idHub: String(idHub || "").trim(),
    occurredAt: iso,
    actor,
    projectFolder: {
      id: folder?.id || "",
      name: folder?.name || ""
    },
    document: {
      id: document?.id || "",
      name: document?.name || "",
      size: Number(document?.size || 0),
      mimeType: document?.mimeType || "",
      eTag: document?.eTag || "",
      webUrl: document?.webUrl || ""
    },
    versioning: versioning || null,
    source: "hub-proyectos-mopt-api"
  };

  const fileName =
    `${safeTimestamp}_${action}_${eventId}.json`;

  const uploadUrl =
    `https://graph.microsoft.com/v1.0/users/` +
    `35c36022-9fcb-4c9a-99a3-43f56441e7b6` +
    `/drive/items/${encodeURIComponent(projectTraceFolder.id)}:/` +
    `${encodeURIComponent(fileName)}:/content`;

  const response = await fetch(uploadUrl, {
    method: "PUT",
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json; charset=utf-8"
    },
    body: JSON.stringify(trace, null, 2)
  });

  const responseText = await response.text();
  let data = {};

  try {
    data = responseText ? JSON.parse(responseText) : {};
  } catch {}

  if (!response.ok) {
    const error = new Error(
      data?.error?.message ||
      `No se pudo guardar la trazabilidad. Graph ${response.status}: ${responseText}`
    );
    error.statusCode = response.status;
    throw error;
  }

  return {
    saved: true,
    eventId,
    fileName,
    traceFolderId: projectTraceFolder.id,
    traceFileId: data.id || ""
  };
}

async function safeWriteDocumentTrace(
  env,
  token,
  request,
  payload
) {
  try {
    return await writeDocumentTrace(
      env,
      token,
      request,
      payload
    );
  } catch (error) {
    return {
      saved: false,
      warning:
        error?.message ||
        "No fue posible guardar la trazabilidad documental."
    };
  }
}


async function getVerifiedTraceActor(request) {
  const identity = await getVerifiedMicrosoftIdentity(request);
  const hubUser = String(request?.headers?.get("X-HUB-User") || "").trim() || "No identificado";
  const hubRole = String(request?.headers?.get("X-HUB-Role") || "").trim() || "No informado";

  return {
    user: hubUser,
    role: hubRole,
    identityStatus: "verified-microsoft",
    microsoft: {
      email: identity.email || "",
      displayName: identity.displayName || "",
      objectId: identity.microsoftObjectId || "",
      subject: identity.microsoftSubject || "",
      tenantId: identity.tenantId || "",
      userPrincipalName: identity.userPrincipalName || "",
      verifiedBy: identity.verifiedBy || "Microsoft Entra ID JWT"
    }
  };
}

async function writeProjectTrace(env, token, request, {
  idHub,
  action,
  project = null,
  changes = [],
  excelRow = null,
  details = null
}) {
  const cleanId = String(idHub || "").trim();
  const traceRoot = await ensureChildFolder(
    token,
    "01GCSG2KRKAQWQFNVDUZF27DQ7K7VHC23W",
    "_HUB_TRAZABILIDAD"
  );

  let projectTraceFolder;
  try {
    projectTraceFolder = await ensureChildFolder(token, traceRoot.id, cleanId);
  } catch (error) {
    if (error.statusCode === 409) {
      projectTraceFolder = await ensureChildFolder(token, traceRoot.id, cleanId);
    } else {
      throw error;
    }
  }

  const iso = new Date().toISOString();
  const safeTimestamp = iso.replace(/[:.]/g, "-");
  const eventId =
    typeof crypto !== "undefined" && crypto.randomUUID
      ? crypto.randomUUID()
      : `${Date.now()}-${Math.random().toString(16).slice(2)}`;
  const actor = await getVerifiedTraceActor(request);

  const trace = {
    schemaVersion: 1,
    eventId,
    eventType: "project",
    action,
    idHub: cleanId,
    occurredAt: iso,
    actor,
    excelRow: Number.isFinite(Number(excelRow)) ? Number(excelRow) : null,
    changes: Array.isArray(changes) ? changes : [],
    project: project || null,
    details: details || null,
    source: "hub-proyectos-mopt-api"
  };

  const fileName = `${safeTimestamp}_project-${action}_${eventId}.json`;
  const uploadUrl =
    `https://graph.microsoft.com/v1.0/users/` +
    `35c36022-9fcb-4c9a-99a3-43f56441e7b6` +
    `/drive/items/${encodeURIComponent(projectTraceFolder.id)}:/` +
    `${encodeURIComponent(fileName)}:/content`;

  const response = await fetch(uploadUrl, {
    method: "PUT",
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json; charset=utf-8"
    },
    body: JSON.stringify(trace, null, 2)
  });

  const responseText = await response.text();
  let data = {};
  try { data = responseText ? JSON.parse(responseText) : {}; } catch {}

  if (!response.ok) {
    const error = new Error(
      data?.error?.message ||
      `No se pudo guardar la trazabilidad del proyecto. Graph ${response.status}: ${responseText}`
    );
    error.statusCode = response.status;
    throw error;
  }

  return {
    saved: true,
    eventId,
    fileName,
    traceFolderId: projectTraceFolder.id,
    traceFileId: data.id || ""
  };
}

async function safeWriteProjectTrace(env, token, request, payload) {
  try {
    return await writeProjectTrace(env, token, request, payload);
  } catch (error) {
    return {
      saved: false,
      warning: error?.message || "No fue posible guardar la trazabilidad del proyecto."
    };
  }
}


async function findChildFolderReadOnly(token, parentId, folderName) {
  let childrenUrl =
    `https://graph.microsoft.com/v1.0/users/` +
    `35c36022-9fcb-4c9a-99a3-43f56441e7b6` +
    `/drive/items/${encodeURIComponent(parentId)}/children` +
    "?$select=id,name,folder&$top=200";

  while (childrenUrl) {
    const data = await graphJson(token, childrenUrl);

    const found = (data.value || []).find(
      item =>
        item.folder &&
        String(item.name || "").toLowerCase() ===
          String(folderName || "").toLowerCase()
    );

    if (found) return found;
    childrenUrl = data["@odata.nextLink"] || "";
  }

  return null;
}

async function getProjectDocumentTraceability(env, idHub) {
  const cleanId = String(idHub || "").trim();

  if (!cleanId) {
    const error = new Error("ID HUB es obligatorio.");
    error.statusCode = 400;
    throw error;
  }

  const token = await getGraphToken(env);

  const traceRoot = await findChildFolderReadOnly(
    token,
    "01GCSG2KRKAQWQFNVDUZF27DQ7K7VHC23W",
    "_HUB_TRAZABILIDAD"
  );

  if (!traceRoot) {
    return {
      ok: true,
      idHub: cleanId,
      totalEvents: 0,
      events: []
    };
  }

  const projectFolder = await findChildFolderReadOnly(
    token,
    traceRoot.id,
    cleanId
  );

  if (!projectFolder) {
    return {
      ok: true,
      idHub: cleanId,
      totalEvents: 0,
      events: []
    };
  }

  let childrenUrl =
    `https://graph.microsoft.com/v1.0/users/` +
    `35c36022-9fcb-4c9a-99a3-43f56441e7b6` +
    `/drive/items/${encodeURIComponent(projectFolder.id)}/children` +
    "?$select=id,name,file,size,lastModifiedDateTime&$top=200";

  const traceFiles = [];

  while (childrenUrl) {
    const data = await graphJson(token, childrenUrl);

    for (const item of data.value || []) {
      if (
        item.file &&
        String(item.name || "").toLowerCase().endsWith(".json")
      ) {
        traceFiles.push(item);
      }
    }

    childrenUrl = data["@odata.nextLink"] || "";
  }

  const events = [];

  for (const item of traceFiles) {
    const contentUrl =
      `https://graph.microsoft.com/v1.0/users/` +
      `35c36022-9fcb-4c9a-99a3-43f56441e7b6` +
      `/drive/items/${encodeURIComponent(item.id)}/content`;

    const response = await fetch(contentUrl, {
      headers: {
        Authorization: `Bearer ${token}`
      }
    });

    if (!response.ok) {
      events.push({
        eventId: "",
        eventType: "document",
        action: "unknown",
        idHub: cleanId,
        occurredAt: item.lastModifiedDateTime || "",
        actor: {
          user: "No identificado",
          role: "No informado",
          identityStatus: "unavailable"
        },
        document: {
          id: "",
          name: item.name || "",
          size: Number(item.size || 0),
          mimeType: "",
          eTag: "",
          webUrl: ""
        },
        versioning: null,
        traceFile: {
          id: item.id || "",
          name: item.name || ""
        },
        warning: `No se pudo leer el registro de trazabilidad. Graph ${response.status}.`
      });
      continue;
    }

    const raw = await response.text();

    try {
      const event = JSON.parse(raw);
      events.push({
        ...event,
        traceFile: {
          id: item.id || "",
          name: item.name || ""
        }
      });
    } catch {
      events.push({
        eventId: "",
        eventType: "document",
        action: "unknown",
        idHub: cleanId,
        occurredAt: item.lastModifiedDateTime || "",
        actor: {
          user: "No identificado",
          role: "No informado",
          identityStatus: "unavailable"
        },
        document: {
          id: "",
          name: item.name || "",
          size: Number(item.size || 0),
          mimeType: "",
          eTag: "",
          webUrl: ""
        },
        versioning: null,
        traceFile: {
          id: item.id || "",
          name: item.name || ""
        },
        warning: "El archivo de trazabilidad no contiene JSON válido."
      });
    }
  }

  events.sort((a, b) =>
    String(b.occurredAt || "").localeCompare(
      String(a.occurredAt || "")
    )
  );

  return {
    ok: true,
    idHub: cleanId,
    totalEvents: events.length,
    events
  };
}

function safeDocumentName(value = "") {
  return String(value)
    .replace(/[~"#%&*:<>?/\\{|}]/g, "-")
    .replace(/\s+/g, " ")
    .replace(/[. ]+$/g, "")
    .trim();
}

async function uploadProjectDocument(env, idHub, request) {
  const cleanId = String(idHub || "").trim();

  const contentType =
    String(request.headers.get("Content-Type") || "");

  if (!contentType.toLowerCase().includes("multipart/form-data")) {
    const error = new Error(
      "La carga debe enviarse como multipart/form-data."
    );
    error.statusCode = 400;
    throw error;
  }

  const formData = await request.formData();
  const file = formData.get("file");

  if (
    !file ||
    typeof file.arrayBuffer !== "function"
  ) {
    const error = new Error(
      'Debe incluir un archivo en el campo "file".'
    );
    error.statusCode = 400;
    throw error;
  }

  const MAX_FILE_SIZE = 5 * 1024 * 1024;

  if (file.size <= 0) {
    const error = new Error(
      "El archivo está vacío."
    );
    error.statusCode = 400;
    throw error;
  }

  if (file.size > MAX_FILE_SIZE) {
    const error = new Error(
      "El archivo supera el límite permitido de 5 MB."
    );
    error.statusCode = 413;
    throw error;
  }

  const fileName =
    safeDocumentName(file.name || "documento");

  if (!fileName) {
    const error = new Error(
      "El nombre del archivo no es válido."
    );
    error.statusCode = 400;
    throw error;
  }

  const token = await getGraphToken(env);

  const folder =
    await findProjectFolderByIdHub(
      env,
      token,
      cleanId
    );

  if (!folder) {
    const error = new Error(
      `No se encontró la carpeta de OneDrive para el proyecto ${cleanId}.`
    );
    error.statusCode = 404;
    throw error;
  }

  // Primera versión: si ya existe un archivo con el mismo nombre,
  // no lo sobrescribimos accidentalmente. El reemplazo/versionado
  // se implementará y probará como operación separada.
  let childrenUrl =
    `https://graph.microsoft.com/v1.0/users/` +
    `35c36022-9fcb-4c9a-99a3-43f56441e7b6` +
    `/drive/items/${encodeURIComponent(folder.id)}/children` +
    "?$select=id,name,file&$top=200";

  while (childrenUrl) {
    const checkResponse = await fetch(childrenUrl, {
      headers: {
        Authorization: `Bearer ${token}`
      }
    });

    const checkText = await checkResponse.text();
    let checkData = {};

    try {
      checkData = checkText ? JSON.parse(checkText) : {};
    } catch {}

    if (!checkResponse.ok) {
      const error = new Error(
        checkData?.error?.message ||
        `No se pudo comprobar si el documento ya existe. Graph ${checkResponse.status}: ${checkText}`
      );
      error.statusCode = checkResponse.status;
      throw error;
    }

    const duplicate = (checkData.value || []).find(
      item =>
        item.file &&
        String(item.name || "").toLowerCase() ===
          fileName.toLowerCase()
    );

    if (duplicate) {
      const error = new Error(
        `Ya existe un documento llamado "${fileName}" en este proyecto. Use la operación de reemplazo para conservar el versionado.`
      );
      error.statusCode = 409;
      throw error;
    }

    childrenUrl =
      checkData["@odata.nextLink"] || "";
  }

  const bytes = await file.arrayBuffer();

  const uploadUrl =
    `https://graph.microsoft.com/v1.0/users/` +
    `35c36022-9fcb-4c9a-99a3-43f56441e7b6` +
    `/drive/items/${encodeURIComponent(folder.id)}:/` +
    `${encodeURIComponent(fileName)}:/content`;

  const response = await fetch(
    uploadUrl,
    {
      method: "PUT",
      headers: {
        Authorization: `Bearer ${token}`,
        "Content-Type":
          file.type || "application/octet-stream"
      },
      body: bytes
    }
  );

  const responseText = await response.text();
  let data = {};

  try {
    data = responseText ? JSON.parse(responseText) : {};
  } catch {}

  if (!response.ok) {
    const error = new Error(
      data?.error?.message ||
      `No se pudo subir el documento. Graph ${response.status}: ${responseText}`
    );
    error.statusCode = response.status;
    throw error;
  }

  const documentResult = {
    id: data.id || "",
    name: data.name || fileName,
    size: Number(data.size || file.size || 0),
    mimeType:
      data.file?.mimeType ||
      file.type ||
      "application/octet-stream",
    eTag: data.eTag || "",
    createdDateTime: data.createdDateTime || "",
    lastModifiedDateTime:
      data.lastModifiedDateTime || "",
    webUrl: data.webUrl || ""
  };

  const traceability = await safeWriteDocumentTrace(
    env,
    token,
    request,
    {
      idHub: cleanId,
      action: "upload",
      folder,
      document: documentResult
    }
  );

  return {
    ok: true,
    message: "Documento cargado correctamente.",
    idHub: cleanId,
    carpeta: {
      id: folder.id,
      name: folder.name
    },
    document: documentResult,
    traceability
  };
}


async function replaceProjectDocument(env, idHub, request) {
  const cleanId = String(idHub || "").trim();

  const contentType =
    String(request.headers.get("Content-Type") || "");

  if (!contentType.toLowerCase().includes("multipart/form-data")) {
    const error = new Error(
      "El reemplazo debe enviarse como multipart/form-data."
    );
    error.statusCode = 400;
    throw error;
  }

  const formData = await request.formData();
  const file = formData.get("file");

  if (!file || typeof file.arrayBuffer !== "function") {
    const error = new Error(
      'Debe incluir un archivo en el campo "file".'
    );
    error.statusCode = 400;
    throw error;
  }

  const MAX_FILE_SIZE = 5 * 1024 * 1024;

  if (file.size <= 0) {
    const error = new Error("El archivo está vacío.");
    error.statusCode = 400;
    throw error;
  }

  if (file.size > MAX_FILE_SIZE) {
    const error = new Error(
      "El archivo supera el límite permitido de 5 MB."
    );
    error.statusCode = 413;
    throw error;
  }

  const fileName = safeDocumentName(file.name || "");

  if (!fileName) {
    const error = new Error("El nombre del archivo no es válido.");
    error.statusCode = 400;
    throw error;
  }

  const token = await getGraphToken(env);

  const folder = await findProjectFolderByIdHub(
    env,
    token,
    cleanId
  );

  if (!folder) {
    const error = new Error(
      `No se encontró la carpeta de OneDrive para el proyecto ${cleanId}.`
    );
    error.statusCode = 404;
    throw error;
  }

  // Buscar el archivo existente dentro de la carpeta.
  let childrenUrl =
    `https://graph.microsoft.com/v1.0/users/` +
    `35c36022-9fcb-4c9a-99a3-43f56441e7b6` +
    `/drive/items/${encodeURIComponent(folder.id)}/children` +
    "?$select=id,name,file,eTag,size,webUrl,lastModifiedDateTime&$top=200";

  let existing = null;

  while (childrenUrl) {
    const response = await fetch(childrenUrl, {
      headers: {
        Authorization: `Bearer ${token}`
      }
    });

    const responseText = await response.text();
    let data = {};

    try {
      data = responseText ? JSON.parse(responseText) : {};
    } catch {}

    if (!response.ok) {
      const error = new Error(
        data?.error?.message ||
        `No se pudo consultar el contenido de la carpeta. Graph ${response.status}: ${responseText}`
      );
      error.statusCode = response.status;
      throw error;
    }

    existing = (data.value || []).find(
      item =>
        item.file &&
        String(item.name || "").toLowerCase() ===
          fileName.toLowerCase()
    );

    if (existing) break;

    childrenUrl = data["@odata.nextLink"] || "";
  }

  if (!existing) {
    const error = new Error(
      `No existe "${fileName}" en el proyecto ${cleanId}. Para crear un documento use POST /projects/{ID HUB}/documents.`
    );
    error.statusCode = 404;
    throw error;
  }

  const bytes = await file.arrayBuffer();

  // PUT sobre el contenido del mismo DriveItem:
  // conserva el mismo archivo y permite que OneDrive/SharePoint
  // gestione el historial de versiones.
  const uploadUrl =
    `https://graph.microsoft.com/v1.0/users/` +
    `35c36022-9fcb-4c9a-99a3-43f56441e7b6` +
    `/drive/items/${encodeURIComponent(existing.id)}/content`;

  const response = await fetch(uploadUrl, {
    method: "PUT",
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type":
        file.type || "application/octet-stream"
    },
    body: bytes
  });

  const responseText = await response.text();
  let data = {};

  try {
    data = responseText ? JSON.parse(responseText) : {};
  } catch {}

  if (!response.ok) {
    const error = new Error(
      data?.error?.message ||
      `No se pudo reemplazar el documento. Graph ${response.status}: ${responseText}`
    );
    error.statusCode = response.status;
    throw error;
  }

  const documentResult = {
    id: data.id || existing.id,
    name: data.name || existing.name,
    size: Number(data.size || file.size || 0),
    mimeType:
      data.file?.mimeType ||
      file.type ||
      "application/octet-stream",
    eTag: data.eTag || "",
    lastModifiedDateTime:
      data.lastModifiedDateTime || "",
    webUrl: data.webUrl || ""
  };

  const versioning = {
    sameDriveItem: true,
    previousETag: existing.eTag || "",
    newETag: data.eTag || ""
  };

  console.log("[DOC_REPLACE_V16]", JSON.stringify({
    projectId: cleanId,
    documentId: String(documentResult.id || ""),
    name: String(documentResult.name || ""),
    uploadedBytes: Number(file.size || 0),
    previousETag: String(existing.eTag || ""),
    newETag: String(data.eTag || ""),
    changedETag: !!data.eTag && String(data.eTag) !== String(existing.eTag || ""),
    lastModifiedDateTime: String(documentResult.lastModifiedDateTime || "")
  }));

  const traceability = await safeWriteDocumentTrace(
    env,
    token,
    request,
    {
      idHub: cleanId,
      action: "replace",
      folder,
      document: documentResult,
      versioning
    }
  );

  return {
    ok: true,
    message:
      "Documento reemplazado correctamente. Se mantuvo el mismo archivo para conservar el historial de versiones.",
    idHub: cleanId,
    carpeta: {
      id: folder.id,
      name: folder.name
    },
    document: documentResult,
    versioning,
    traceability
  };
}


async function deleteProjectDocument(env, idHub, documentId, request) {
  const cleanId = String(idHub || "").trim();
  const cleanDocumentId = String(documentId || "").trim();

  if (!cleanId || !cleanDocumentId) {
    const error = new Error("ID HUB y documentId son obligatorios.");
    error.statusCode = 400;
    throw error;
  }

  const token = await getGraphToken(env);

  const folder = await findProjectFolderByIdHub(
    env,
    token,
    cleanId
  );

  if (!folder) {
    const error = new Error(
      `No se encontró la carpeta de OneDrive para el proyecto ${cleanId}.`
    );
    error.statusCode = 404;
    throw error;
  }

  // Primero verificamos que el documento pertenezca realmente
  // a la carpeta del proyecto. Así el endpoint no puede usarse
  // para eliminar arbitrariamente otros elementos del OneDrive.
  let childrenUrl =
    `https://graph.microsoft.com/v1.0/users/` +
    `35c36022-9fcb-4c9a-99a3-43f56441e7b6` +
    `/drive/items/${encodeURIComponent(folder.id)}/children` +
    "?$select=id,name,file,eTag,size,webUrl,lastModifiedDateTime&$top=200";

  let existing = null;

  while (childrenUrl) {
    const response = await fetch(childrenUrl, {
      headers: {
        Authorization: `Bearer ${token}`
      }
    });

    const responseText = await response.text();
    let data = {};

    try {
      data = responseText ? JSON.parse(responseText) : {};
    } catch {}

    if (!response.ok) {
      const error = new Error(
        data?.error?.message ||
        `No se pudo consultar el contenido de la carpeta. Graph ${response.status}: ${responseText}`
      );
      error.statusCode = response.status;
      throw error;
    }

    existing = (data.value || []).find(
      item =>
        item.file &&
        String(item.id || "") === cleanDocumentId
    );

    if (existing) break;

    childrenUrl = data["@odata.nextLink"] || "";
  }

  if (!existing) {
    const error = new Error(
      `No se encontró el documento solicitado dentro del proyecto ${cleanId}.`
    );
    error.statusCode = 404;
    throw error;
  }

  const deleteUrl =
    `https://graph.microsoft.com/v1.0/users/` +
    `35c36022-9fcb-4c9a-99a3-43f56441e7b6` +
    `/drive/items/${encodeURIComponent(existing.id)}`;

  const response = await fetch(deleteUrl, {
    method: "DELETE",
    headers: {
      Authorization: `Bearer ${token}`
    }
  });

  const responseText = await response.text();

  if (!response.ok) {
    let data = {};

    try {
      data = responseText ? JSON.parse(responseText) : {};
    } catch {}

    const graphMessage =
      data?.error?.message ||
      responseText ||
      `Graph ${response.status}`;

    const error = new Error(
      response.status === 423
        ? `El documento "${existing.name}" está abierto o bloqueado en Microsoft 365. Ciérrelo y vuelva a intentar la eliminación.`
        : `No se pudo eliminar el documento "${existing.name}". Graph ${response.status}: ${graphMessage}`
    );

    error.statusCode =
      response.status === 423
        ? 423
        : response.status === 404
          ? 404
          : response.status;

    throw error;
  }

  const documentResult = {
    id: existing.id,
    name: existing.name || "",
    size: Number(existing.size || 0),
    mimeType: existing.file?.mimeType || "",
    eTag: existing.eTag || "",
    webUrl: existing.webUrl || ""
  };

  const traceability = await safeWriteDocumentTrace(
    env,
    token,
    request,
    {
      idHub: cleanId,
      action: "delete",
      folder,
      document: documentResult
    }
  );

  return {
    ok: true,
    message: "Documento eliminado correctamente.",
    idHub: cleanId,
    carpeta: {
      id: folder.id,
      name: folder.name
    },
    document: documentResult,
    traceability
  };
}

async function getProjects(env) {
  const token =
    await getGraphToken(env);

  const workbook =
    await loadWorkbook(env, token);

  const active =
    workbook.projects.filter(
      project =>
        String(
          project["Estado HUB"] || ""
        )
          .trim()
          .toLowerCase() !==
        "eliminado"
    );

  return {
    ok: true,
    totalRecords:
      workbook.projects.length,
    activeRecords:
      active.length,
    excludedDeleted:
      workbook.projects.length -
      active.length,
    projects:
      active
  };
}


// ============================================================
// IDENTIDAD MICROSOFT DEL USUARIO DEL HUB
// Este token es el token delegado del usuario (no el token app-only
// que usa el Worker para operar sobre el OneDrive central).
// Microsoft Graph valida el token al resolver /me.
// ============================================================
const HUB_API_CLIENT_ID = "3a5ead05-e9e6-477b-9a63-24ef1f9fefd5";
const HUB_API_AUDIENCES = new Set([
  HUB_API_CLIENT_ID,
  `api://${HUB_API_CLIENT_ID}`
]);
const HUB_REQUIRED_SCOPE = "access_as_user";
let microsoftJwksCache = { expiresAt: 0, keys: [] };

function base64UrlToBytes(value) {
  let s = String(value || "").replace(/-/g, "+").replace(/_/g, "/");
  while (s.length % 4) s += "=";
  const binary = atob(s);
  return Uint8Array.from(binary, ch => ch.charCodeAt(0));
}

function decodeJwtJson(value) {
  const bytes = base64UrlToBytes(value);
  return JSON.parse(new TextDecoder().decode(bytes));
}

async function getMicrosoftSigningKeys() {
  const now = Date.now();
  if (microsoftJwksCache.keys.length && microsoftJwksCache.expiresAt > now) {
    return microsoftJwksCache.keys;
  }

  const response = await fetch(
    "https://login.microsoftonline.com/common/discovery/v2.0/keys",
    { headers: { Accept: "application/json" } }
  );
  if (!response.ok) {
    const error = new Error("No fue posible obtener las claves públicas de Microsoft.");
    error.statusCode = 503;
    throw error;
  }
  const data = await response.json();
  const keys = Array.isArray(data?.keys) ? data.keys : [];
  microsoftJwksCache = {
    keys,
    expiresAt: now + 6 * 60 * 60 * 1000
  };
  return keys;
}

async function verifyMicrosoftJwtSignature(token, header) {
  if (header?.alg !== "RS256" || !header?.kid) {
    const error = new Error("El token Microsoft usa un algoritmo o clave no admitidos.");
    error.statusCode = 401;
    throw error;
  }

  const keys = await getMicrosoftSigningKeys();
  const jwk = keys.find(k => k.kid === header.kid);
  if (!jwk) {
    microsoftJwksCache = { expiresAt: 0, keys: [] };
    const refreshed = await getMicrosoftSigningKeys();
    const retryJwk = refreshed.find(k => k.kid === header.kid);
    if (!retryJwk) {
      const error = new Error("No se encontró la clave pública usada para firmar el token Microsoft.");
      error.statusCode = 401;
      throw error;
    }
    return verifyWithJwk(token, retryJwk);
  }
  return verifyWithJwk(token, jwk);
}

async function verifyWithJwk(token, jwk) {
  const parts = token.split(".");
  const key = await crypto.subtle.importKey(
    "jwk",
    jwk,
    { name: "RSASSA-PKCS1-v1_5", hash: "SHA-256" },
    false,
    ["verify"]
  );
  const signed = new TextEncoder().encode(`${parts[0]}.${parts[1]}`);
  const signature = base64UrlToBytes(parts[2]);
  const ok = await crypto.subtle.verify(
    { name: "RSASSA-PKCS1-v1_5" },
    key,
    signature,
    signed
  );
  if (!ok) {
    const error = new Error("La firma del token Microsoft no es válida.");
    error.statusCode = 401;
    throw error;
  }
}

async function getVerifiedMicrosoftIdentity(request) {
  const authorization = String(request.headers.get("Authorization") || "").trim();
  if (!authorization.toLowerCase().startsWith("bearer ")) {
    const error = new Error("Falta el token de identidad Microsoft del usuario.");
    error.statusCode = 401;
    throw error;
  }

  const token = authorization.slice(7).trim();
  const parts = token.split(".");
  if (parts.length !== 3) {
    const error = new Error("El token Microsoft no tiene formato JWT válido.");
    error.statusCode = 401;
    throw error;
  }

  let header, claims;
  try {
    header = decodeJwtJson(parts[0]);
    claims = decodeJwtJson(parts[1]);
  } catch {
    const error = new Error("No fue posible interpretar el token Microsoft.");
    error.statusCode = 401;
    throw error;
  }

  await verifyMicrosoftJwtSignature(token, header);

  const now = Math.floor(Date.now() / 1000);
  if (!claims.exp || Number(claims.exp) <= now - 60) {
    const error = new Error("El token Microsoft expiró.");
    error.statusCode = 401;
    throw error;
  }
  if (claims.nbf && Number(claims.nbf) > now + 60) {
    const error = new Error("El token Microsoft todavía no es válido.");
    error.statusCode = 401;
    throw error;
  }

  const audience = String(claims.aud || "");
  if (!HUB_API_AUDIENCES.has(audience)) {
    const error = new Error("El token Microsoft no fue emitido para la API del HUB.");
    error.statusCode = 401;
    throw error;
  }

  const scopes = String(claims.scp || "").split(/\s+/).filter(Boolean);
  if (!scopes.includes(HUB_REQUIRED_SCOPE)) {
    const error = new Error("El token Microsoft no incluye el permiso access_as_user.");
    error.statusCode = 403;
    throw error;
  }

  const tid = String(claims.tid || "").trim();
  const issuer = String(claims.iss || "").replace(/\/+$/, "");
  if (!tid || issuer !== `https://login.microsoftonline.com/${tid}/v2.0`) {
    const error = new Error("El emisor del token Microsoft no es válido.");
    error.statusCode = 401;
    throw error;
  }

  const email = String(
    claims.email || claims.preferred_username || claims.upn || ""
  ).trim();

  return {
    verified: true,
    microsoftObjectId: String(claims.oid || ""),
    microsoftSubject: String(claims.sub || ""),
    tenantId: tid,
    displayName: String(claims.name || ""),
    email,
    userPrincipalName: String(claims.preferred_username || claims.upn || ""),
    verifiedBy: "Microsoft Entra ID JWT",
    audience,
    scope: HUB_REQUIRED_SCOPE
  };
}



async function getProjectGeometryApi(env, idHub) {
  const cleanId = String(idHub || "").trim();
  if (!cleanId) { const error = new Error("ID HUB es obligatorio."); error.statusCode = 400; throw error; }
  const token = await getGraphToken(env);
  const geometryFolder = await findChildFolderReadOnly(token, "01GCSG2KRKAQWQFNVDUZF27DQ7K7VHC23W", "_HUB_GEOMETRIA");
  if (!geometryFolder) return { ok: true, idHub: cleanId, geometry: null };
  const fileName = `${cleanId.replace(/[~"#%&*:<>?/\\{|}]/g, "-")}.json`;
  let childrenUrl = `https://graph.microsoft.com/v1.0/users/35c36022-9fcb-4c9a-99a3-43f56441e7b6/drive/items/${encodeURIComponent(geometryFolder.id)}/children?$select=id,name,file,lastModifiedDateTime,eTag&$top=200`;
  let found = null;
  while (childrenUrl && !found) {
    const data = await graphJson(token, childrenUrl);
    found = (data.value || []).find(item => item.file && String(item.name || "").toLowerCase() === fileName.toLowerCase()) || null;
    childrenUrl = found ? "" : (data["@odata.nextLink"] || "");
  }
  if (!found) return { ok: true, idHub: cleanId, geometry: null };
  const contentUrl = `https://graph.microsoft.com/v1.0/users/35c36022-9fcb-4c9a-99a3-43f56441e7b6/drive/items/${encodeURIComponent(geometryFolder.id)}:/${encodeURIComponent(found.name)}:/content`;
  let response = await fetch(contentUrl, { method:"GET", headers:{Authorization:`Bearer ${token}`}, redirect:"manual" });
  if (response.status >= 300 && response.status < 400) {
    const location = String(response.headers.get("Location") || "").trim();
    if (!location) { const error = new Error(`Graph respondió ${response.status} sin Location para la geometría.`); error.statusCode=502; throw error; }
    response = await fetch(location, {method:"GET",redirect:"follow"});
  }
  const text = await response.text();
  if (!response.ok) { const error=new Error(`No se pudo leer la geometría. HTTP ${response.status}: ${text.slice(0,700)}`); error.statusCode=response.status; throw error; }
  let geometry=null;
  try { geometry=text?JSON.parse(text):null; } catch { const error=new Error("El archivo de geometría central no contiene JSON válido."); error.statusCode=502; throw error; }
  return {ok:true,idHub:cleanId,geometry,file:{id:found.id||"",name:found.name||fileName,eTag:found.eTag||""}};
}

async function saveProjectGeometryApi(env, idHub, geometry) {
  const cleanId = String(idHub || "").trim();
  if (!cleanId) {
    const error = new Error("ID HUB es obligatorio.");
    error.statusCode = 400;
    throw error;
  }

  const type = String(geometry?.type || "").trim();
  const referencePoints = Array.isArray(geometry?.referencePoints)
    ? geometry.referencePoints
    : [];

  if (!["point", "segment"].includes(type) || !referencePoints.length) {
    const error = new Error("La geometría enviada no es válida.");
    error.statusCode = 400;
    throw error;
  }
  if (type === "segment" && referencePoints.length < 2) {
    const error = new Error("Un tramo requiere al menos dos puntos.");
    error.statusCode = 400;
    throw error;
  }

  const token = await getGraphToken(env);
  const geometryFolder = await ensureChildFolder(
    token,
    "01GCSG2KRKAQWQFNVDUZF27DQ7K7VHC23W",
    "_HUB_GEOMETRIA"
  );

  const payload = {
    projectId: cleanId,
    updatedAt: new Date().toISOString(),
    ...geometry,
    source: geometry?.source || "user"
  };

  const fileName = `${cleanId.replace(/[~"#%&*:<>?/\\{|}]/g, "-")}.json`;
  const uploadUrl =
    `https://graph.microsoft.com/v1.0/users/35c36022-9fcb-4c9a-99a3-43f56441e7b6` +
    `/drive/items/${encodeURIComponent(geometryFolder.id)}:/` +
    `${encodeURIComponent(fileName)}:/content`;

  const response = await fetch(uploadUrl, {
    method: "PUT",
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json; charset=utf-8"
    },
    body: JSON.stringify(payload, null, 2)
  });

  const responseText = await response.text();
  let data = {};
  try { data = responseText ? JSON.parse(responseText) : {}; } catch {}

  if (!response.ok) {
    const error = new Error(
      data?.error?.message ||
      `No se pudo guardar la geometría. Graph ${response.status}: ${responseText}`
    );
    error.statusCode = response.status;
    throw error;
  }

  return {
    ok: true,
    idHub: cleanId,
    geometry: payload,
    file: {
      id: data.id || "",
      name: data.name || fileName,
      eTag: data.eTag || ""
    }
  };
}


// ===== Gobernanza central V25: Administrador HUB + plazos por paso =====
const GOV24_DATA_PREFIX="__GOV24_DATA__";
const GOV24_MARKER_PREFIX="__GOV24__";
function govSafeId(v){return String(v||"").trim().replace(/[^A-Za-z0-9._-]+/g,"-").slice(0,90);}
function govB64Encode(obj){const bytes=new TextEncoder().encode(JSON.stringify(obj));let bin="";for(const b of bytes)bin+=String.fromCharCode(b);return btoa(bin).replace(/\+/g,"-").replace(/\//g,"_").replace(/=+$/g,"");}
function govB64Decode(s){try{let x=s.replace(/-/g,"+").replace(/_/g,"/");while(x.length%4)x+="=";const bin=atob(x),bytes=Uint8Array.from(bin,c=>c.charCodeAt(0));return JSON.parse(new TextDecoder().decode(bytes));}catch{return null;}}
async function resolveGovernanceFolder(env,token){return await ensureChildFolder(token,"01GCSG2KRKAQWQFNVDUZF27DQ7K7VHC23W","_HUB_GOBERNANZA");}
function governanceMarkerName(g){const a=g?.assignment||{},e=a.extension||{},q=g?.updateRequest||{};const compact={p:String(g.projectId||""),l:String(g.lifecycleState||"").slice(0,24),v:String(g.lastValidatedAt||"").slice(0,24),u:String(g.updatedAt||"").slice(0,24),a:String(a.activity||"").slice(0,32),r:String(a.responsible||a.responsibleRole||"").slice(0,24),d:String(a.dueAt||"").slice(0,24),b:Number(a.businessDays||0),s:String(a.status||"").slice(0,14),x:String(e.status||"").slice(0,12),z:Number(e.extraBusinessDays||0),o:String(e.originalDueAt||a.originalDueAt||"").slice(0,24),q:String(q.status||"").slice(0,12),qd:String(q.dueAt||"").slice(0,24),qr:String(q.responsible||"").slice(0,20)};return `${GOV24_MARKER_PREFIX}${govB64Encode(compact)}.json`;}
function parseGovernanceMarker(name){if(!String(name).startsWith(GOV24_MARKER_PREFIX)||String(name).startsWith(GOV24_DATA_PREFIX))return null;const raw=String(name).slice(GOV24_MARKER_PREFIX.length).replace(/\.json$/i,"");const c=govB64Decode(raw);if(!c?.p)return null;const assignment=c.a||c.r||c.d?{activity:c.a||"",responsible:c.r||"",dueAt:c.d||"",businessDays:Number(c.b||0),status:c.s||"",originalDueAt:c.o||"",extension:c.x?{status:c.x,extraBusinessDays:Number(c.z||0),originalDueAt:c.o||""}:null}:null;const updateRequest=c.q?{status:c.q,dueAt:c.qd||"",responsible:c.qr||""}:null;return {projectId:c.p,lifecycleState:c.l||"",lastValidatedAt:c.v||"",updatedAt:c.u||"",assignment,updateRequest};}
async function listGovernanceChildren(token,folder){let url=`https://graph.microsoft.com/v1.0/users/${WF_OWNER_USER_ID}/drive/items/${encodeURIComponent(folder.id)}/children?$select=id,name,file&$top=999`,out=[];while(url){const r=await fetch(url,{headers:{Authorization:`Bearer ${token}`}}),t=await r.text();if(!r.ok)throw new Error(`No se pudo listar gobernanza. Graph ${r.status}: ${t}`);const d=JSON.parse(t||"{}");out.push(...(d.value||[]));url=d["@odata.nextLink"]||"";}return out;}

const GOV25_CONFIG_NAME="__GOV25_CONFIG__.json";
const GOV25_DEFAULT_CONFIG={stepDeadlines:{1:5,2:5,3:5,4:5,5:3,6:5,7:10,8:2,9:10,10:30,11:5,12:5,13:15,14:60,15:10,16:10,17:15},maxBusinessDays:180,warningBusinessDays:30,updateRequestBusinessDays:5,warningBeforeDays:3};
function requireGovernanceAdmin(request){const role=String(request.headers.get("X-HUB-Role")||"").trim().toLowerCase();if(role!=="administrador")throw new Error("Esta operación requiere el perfil Administrador HUB / Gobernanza.");}
async function getGovernanceConfigApi(env){const token=await getGraphToken(env),folder=await resolveGovernanceFolder(env,token),url=`https://graph.microsoft.com/v1.0/users/${WF_OWNER_USER_ID}/drive/items/${encodeURIComponent(folder.id)}:/${encodeURIComponent(GOV25_CONFIG_NAME)}:/content`;const r=await fetch(url,{headers:{Authorization:`Bearer ${token}`},redirect:"follow"});if(r.status===404)return {ok:true,version:"governance-v25",config:GOV25_DEFAULT_CONFIG};const t=await r.text();if(!r.ok)throw new Error(`No se pudo leer configuración de gobernanza. Graph ${r.status}: ${t}`);return {ok:true,version:"governance-v25",config:{...GOV25_DEFAULT_CONFIG,...JSON.parse(t),stepDeadlines:{...GOV25_DEFAULT_CONFIG.stepDeadlines,...(JSON.parse(t).stepDeadlines||{})}}};}
async function saveGovernanceConfigApi(env,input,request){requireGovernanceAdmin(request);const token=await getGraphToken(env),folder=await resolveGovernanceFolder(env,token),previous=(await getGovernanceConfigApi(env)).config,now=new Date().toISOString();const clean={...previous,...(input||{}),stepDeadlines:{...previous.stepDeadlines,...((input||{}).stepDeadlines||{})},updatedAt:now,updatedBy:String((input||{}).updatedBy||"").slice(0,180)};for(let i=1;i<=17;i++){const n=Number(clean.stepDeadlines[i]);if(!Number.isFinite(n)||n<1||n>365)throw new Error(`Plazo inválido para el paso ${i}.`);clean.stepDeadlines[i]=Math.round(n);}const url=`https://graph.microsoft.com/v1.0/users/${WF_OWNER_USER_ID}/drive/items/${encodeURIComponent(folder.id)}:/${encodeURIComponent(GOV25_CONFIG_NAME)}:/content`;const r=await fetch(url,{method:"PUT",headers:{Authorization:`Bearer ${token}`,"Content-Type":"application/json; charset=utf-8"},body:JSON.stringify(clean,null,2)});if(!r.ok)throw new Error(`No se pudo guardar configuración de gobernanza. Graph ${r.status}: ${await r.text()}`);await safeWriteProjectTrace(env,token,request,{idHub:"_HUB_GOBERNANZA",action:"governance_config_update",project:{"ID HUB":"_HUB_GOBERNANZA"},changes:[],details:{stepDeadlines:clean.stepDeadlines,updatedBy:clean.updatedBy}});return {ok:true,version:"governance-v25",config:clean};}

async function listAllGovernanceApi(env){const token=await getGraphToken(env),folder=await resolveGovernanceFolder(env,token),items=await listGovernanceChildren(token,folder),m=new Map();for(const x of items){const g=parseGovernanceMarker(x.name);if(!g)continue;const prev=m.get(g.projectId);if(!prev||String(g.updatedAt)>String(prev.updatedAt))m.set(g.projectId,g);}return {ok:true,version:"governance-v24",total:m.size,governance:[...m.values()]};}
async function getProjectGovernanceApi(env,idHub){const id=String(idHub||"").trim(),token=await getGraphToken(env),folder=await resolveGovernanceFolder(env,token),name=`${GOV24_DATA_PREFIX}${govSafeId(id)}.json`,url=`https://graph.microsoft.com/v1.0/users/${WF_OWNER_USER_ID}/drive/items/${encodeURIComponent(folder.id)}:/${encodeURIComponent(name)}:/content`;const r=await fetch(url,{headers:{Authorization:`Bearer ${token}`},redirect:"follow"});if(r.status===404)return {ok:true,projectId:id,governance:null};const t=await r.text();if(!r.ok)throw new Error(`No se pudo leer gobernanza. Graph ${r.status}: ${t}`);return {ok:true,projectId:id,governance:JSON.parse(t)};}
async function saveProjectGovernanceApi(env,idHub,input,request){const id=String(idHub||"").trim();if(!id)throw new Error("ID HUB obligatorio.");const token=await getGraphToken(env),folder=await resolveGovernanceFolder(env,token);let previous=null;try{previous=(await getProjectGovernanceApi(env,id)).governance;}catch{}const now=new Date().toISOString();const payload={...(previous||{}),...(input||{}),projectId:id,updatedAt:now};const dataName=`${GOV24_DATA_PREFIX}${govSafeId(id)}.json`,dataUrl=`https://graph.microsoft.com/v1.0/users/${WF_OWNER_USER_ID}/drive/items/${encodeURIComponent(folder.id)}:/${encodeURIComponent(dataName)}:/content`;let r=await fetch(dataUrl,{method:"PUT",headers:{Authorization:`Bearer ${token}`,"Content-Type":"application/json; charset=utf-8"},body:JSON.stringify(payload,null,2)});let t=await r.text();if(!r.ok)throw new Error(`No se pudo guardar gobernanza. Graph ${r.status}: ${t}`);
 const items=await listGovernanceChildren(token,folder);for(const x of items.filter(x=>{const g=parseGovernanceMarker(x.name);return g?.projectId===id;})){await fetch(`https://graph.microsoft.com/v1.0/users/${WF_OWNER_USER_ID}/drive/items/${encodeURIComponent(x.id)}`,{method:"DELETE",headers:{Authorization:`Bearer ${token}`}});}
 const marker=governanceMarkerName(payload),markerUrl=`https://graph.microsoft.com/v1.0/users/${WF_OWNER_USER_ID}/drive/items/${encodeURIComponent(folder.id)}:/${encodeURIComponent(marker)}:/content`;r=await fetch(markerUrl,{method:"PUT",headers:{Authorization:`Bearer ${token}`,"Content-Type":"application/json"},body:"{}"});if(!r.ok)throw new Error(`No se pudo actualizar índice de gobernanza. Graph ${r.status}: ${await r.text()}`);
 await safeWriteProjectTrace(env,token,request,{idHub:id,action:"governance_update",project:{"ID HUB":id},changes:[],details:{lifecycleState:payload.lifecycleState||"",lastValidatedAt:payload.lastValidatedAt||"",assignment:payload.assignment||null}});
 return {ok:true,projectId:id,governance:payload};}

export default {
  async fetch(request, env) {
    try {
      const url =
        new URL(request.url);

      // CORS: GitHub Pages del HUB. El navegador envía OPTIONS antes de
      // solicitudes con Authorization/JSON; debe responderse sin autenticación.
      if (request.method === "OPTIONS") {
        const origin = String(request.headers.get("Origin") || "");
        if (origin !== HUB_ALLOWED_ORIGIN) {
          return json({ ok: false, error: "Origen no autorizado." }, 403);
        }
        return new Response(null, {
          status: 204,
          headers: corsHeaders()
        });
      }

      // PRIMERA PRUEBA DE IDENTIDAD: permite comprobar que el navegador
      // obtuvo un token delegado válido de Microsoft antes de proteger
      // el resto de endpoints del HUB.
      if (
        url.pathname === "/auth/me" &&
        request.method === "GET"
      ) {
        return json({
          ok: true,
          identity: await getVerifiedMicrosoftIdentity(request)
        });
      }

      const traceabilityMatch =
        url.pathname.match(
          /^\/projects\/([^/]+)\/traceability$/
        );

      if (
        traceabilityMatch &&
        request.method === "GET"
      ) {
        return json(
          await getProjectDocumentTraceability(
            env,
            decodeURIComponent(traceabilityMatch[1])
          )
        );
      }

      const documentEditMatch =
        url.pathname.match(
          /^\/projects\/([^/]+)\/documents\/([^/]+)\/edit-link$/
        );

      if (
        documentEditMatch &&
        request.method === "POST"
      ) {
        await getVerifiedMicrosoftIdentity(request);
        return json(
          await createProjectDocumentEditLink(
            env,
            decodeURIComponent(documentEditMatch[1]),
            decodeURIComponent(documentEditMatch[2]),
            request
          )
        );
      }

      const documentContentMatch =
        url.pathname.match(
          /^\/projects\/([^/]+)\/documents\/([^/]+)\/content$/
        );

      if (
        documentContentMatch &&
        request.method === "GET"
      ) {
        await getVerifiedMicrosoftIdentity(request);
        return await getProjectDocumentContent(
          env,
          decodeURIComponent(documentContentMatch[1]),
          decodeURIComponent(documentContentMatch[2])
        );
      }

      const deleteDocumentMatch =
        url.pathname.match(
          /^\/projects\/([^/]+)\/documents\/([^/]+)$/
        );

      if (
        deleteDocumentMatch &&
        request.method === "DELETE"
      ) {
        return json(
          await deleteProjectDocument(
            env,
            decodeURIComponent(deleteDocumentMatch[1]),
            decodeURIComponent(deleteDocumentMatch[2]),
            request
          )
        );
      }

      if (
        url.pathname === "/projects" &&
        request.method === "GET"
      ) {
        return json(
          await getProjects(env)
        );
      }

      if (
        url.pathname === "/projects" &&
        request.method === "POST"
      ) {
        const input =
          await request.json();

        const identity = await getVerifiedMicrosoftIdentity(request);
        const result = await registerProject(env, input);
        const graphToken = await getGraphToken(env);
        result.traceability = await safeWriteProjectTrace(
          env, graphToken, request, {
            idHub: result.idHub,
            action: "create",
            project: result.proyecto || input,
            excelRow: result.excelRow,
            changes: [],
            details: { estadoHub: result.estadoHub || "Incluido" }
          }
        );
        result.verifiedIdentity = {
          email: identity.email || "",
          displayName: identity.displayName || ""
        };
        return json(result, 201);
      }

      const documentsMatch =
        url.pathname.match(/^\/projects\/([^/]+)\/documents$/);

      if (
        documentsMatch &&
        request.method === "GET"
      ) {
        const idHub =
          decodeURIComponent(documentsMatch[1]);

        return json(
          await listProjectDocuments(env, idHub)
        );
      }

      if (
        documentsMatch &&
        request.method === "POST"
      ) {
        const idHub =
          decodeURIComponent(documentsMatch[1]);

        return json(
          await uploadProjectDocument(
            env,
            idHub,
            request
          ),
          201
        );
      }

      const replaceDocumentMatch =
        url.pathname.match(/^\/projects\/([^/]+)\/documents\/replace$/);

      if (
        replaceDocumentMatch &&
        request.method === "PUT"
      ) {
        const idHub =
          decodeURIComponent(replaceDocumentMatch[1]);

        return json(
          await replaceProjectDocument(
            env,
            idHub,
            request
          )
        );
      }

      const repairMatch =
        url.pathname.match(/^\/maintenance\/repair-reference\/([^/]+)$/);

      if (
        repairMatch &&
        request.method === "POST"
      ) {
        const idHub =
          decodeURIComponent(repairMatch[1]);

        return json(
          await repairReferenceFila(env, idHub)
        );
      }

      if (url.pathname === "/workflow" && request.method === "GET") {
        await getVerifiedMicrosoftIdentity(request);
        return json(await listAllWorkflowApi(env));
      }
      const workflowMatch = url.pathname.match(/^\/projects\/([^/]+)\/workflow$/);
      if (workflowMatch && request.method === "GET") {
        await getVerifiedMicrosoftIdentity(request);
        return json(await getProjectWorkflowApi(env, decodeURIComponent(workflowMatch[1])));
      }
      if (workflowMatch && request.method === "PUT") {
        await getVerifiedMicrosoftIdentity(request);
        return json(await saveProjectWorkflowApi(env, decodeURIComponent(workflowMatch[1]), await request.json()));
      }

      if (url.pathname === "/governance/config" && request.method === "GET") {
        await getVerifiedMicrosoftIdentity(request);
        return json(await getGovernanceConfigApi(env));
      }
      if (url.pathname === "/governance/config" && request.method === "PUT") {
        await getVerifiedMicrosoftIdentity(request);
        return json(await saveGovernanceConfigApi(env, await request.json(), request));
      }
      if (url.pathname === "/governance" && request.method === "GET") {
        await getVerifiedMicrosoftIdentity(request);
        return json(await listAllGovernanceApi(env));
      }
      const governanceMatch = url.pathname.match(/^\/projects\/([^/]+)\/governance$/);
      if (governanceMatch && request.method === "GET") {
        await getVerifiedMicrosoftIdentity(request);
        return json(await getProjectGovernanceApi(env, decodeURIComponent(governanceMatch[1])));
      }
      if (governanceMatch && request.method === "PUT") {
        await getVerifiedMicrosoftIdentity(request);
        return json(await saveProjectGovernanceApi(env, decodeURIComponent(governanceMatch[1]), await request.json(), request));
      }

      const geometryMatch =
        url.pathname.match(/^\/projects\/([^/]+)\/geometry$/);

      if (geometryMatch && request.method === "GET") {
        await getVerifiedMicrosoftIdentity(request);
        return json(await getProjectGeometryApi(env, decodeURIComponent(geometryMatch[1])));
      }

      if (geometryMatch && request.method === "PUT") {
        const identity = await getVerifiedMicrosoftIdentity(request);
        const idHub = decodeURIComponent(geometryMatch[1]);
        const inputGeometry = await request.json();
        const result = await saveProjectGeometryApi(env, idHub, inputGeometry);
        const traceToken = await getGraphToken(env);
        result.traceability = await safeWriteProjectTrace(env, traceToken, request, {
          idHub,
          action: "GEOMETRY_UPDATED",
          details: {
            geometryType: result?.geometry?.type || "",
            referencePoints: Array.isArray(result?.geometry?.referencePoints) ? result.geometry.referencePoints.length : 0,
            updatedAt: result?.geometry?.updatedAt || "",
            source: result?.geometry?.source || "user",
            microsoftIdentity: identity?.email || identity?.userPrincipalName || ""
          }
        });
        return json(result);
      }

      const projectMatch =
        url.pathname.match(/^\/projects\/([^/]+)$/);

      if (
        projectMatch &&
        request.method === "PUT"
      ) {
        const idHub =
          decodeURIComponent(projectMatch[1]);

        const input =
          await request.json();

        const identity = await getVerifiedMicrosoftIdentity(request);
        const result = await updateProject(env, idHub, input);
        if (Array.isArray(result.cambios) && result.cambios.length > 0) {
          const graphToken = await getGraphToken(env);
          result.traceability = await safeWriteProjectTrace(
            env, graphToken, request, {
              idHub,
              action: "update",
              project: input,
              excelRow: result.excelRow,
              changes: result.cambios
            }
          );
        } else {
          result.traceability = { saved: false, skipped: true, reason: "no-changes" };
        }
        result.verifiedIdentity = {
          email: identity.email || "",
          displayName: identity.displayName || ""
        };
        return json(result);
      }

      if (
        projectMatch &&
        request.method === "DELETE"
      ) {
        const idHub =
          decodeURIComponent(projectMatch[1]);

        const identity = await getVerifiedMicrosoftIdentity(request);
        const result = await deleteProject(env, idHub);
        if (Array.isArray(result.cambios) && result.cambios.length > 0) {
          const graphToken = await getGraphToken(env);
          result.traceability = await safeWriteProjectTrace(
            env, graphToken, request, {
              idHub,
              action: "delete",
              excelRow: result.excelRow,
              changes: result.cambios,
              details: {
                logicalDelete: true,
                oneDrivePreserved: true,
                estadoAnterior: result.estadoAnterior || "",
                estadoHub: result.estadoHub || "Eliminado"
              }
            }
          );
        } else {
          result.traceability = { saved: false, skipped: true, reason: "already-deleted" };
        }
        result.verifiedIdentity = {
          email: identity.email || "",
          displayName: identity.displayName || ""
        };
        return json(result);
      }

      return json({
        ok: true,
        servicio:
          "HUB Proyectos MOPT API",
        endpoints: {
          leer:
            "GET /projects",
          registrar:
            "POST /projects",
          editar:
            "PUT /projects/{ID HUB}",
          eliminar:
            "DELETE /projects/{ID HUB}",
          documentos:
            "GET /projects/{ID HUB}/documents",
          subirDocumento:
            "POST /projects/{ID HUB}/documents"
        }
      });

    } catch (error) {
      return json(
        {
          ok: false,
          error:
            error?.message ||
            String(error)
        },
        Number(error?.statusCode) || 500
      );
    }
  }
};