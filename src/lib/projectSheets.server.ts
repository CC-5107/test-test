// Server-only helpers for keeping the "Sponsored Projects (Lovable)"
// spreadsheet in sync with live project data. Calls go directly to the
// Google Sheets API using a server-side service account.

import { SHEETS_API_BASE, sheetsAuthHeaders, invalidateSheetsAccessToken, getServerEnv } from "./googleSheetsAuth.server";

export const PROJECTS_SPREADSHEET_ID =
  getServerEnv("GOOGLE_SPREADSHEET_ID") ||
  "17Y47rJ8alS9tiAUvKYR5ADlhWoGZrHo_2FFJdvTLnYs";
export const SUMMARY_SHEET = "SRL_OnePage_Summary";
const SHEETS_API = SHEETS_API_BASE;
const META_CACHE_MS = 5 * 60 * 1000;

type SheetMeta = { title: string; sheetId: number };
let sheetMetaCache: { expiresAt: number; value: SheetMeta[] } | null = null;
let sheetMetaRequest: Promise<SheetMeta[]> | null = null;
let summaryOperation = Promise.resolve();
let summaryHeaderVerified = false;

export const SUMMARY_HEADER = [
  "Project ID",
  "Company",
  "PI",
  "RFS/NTI",
  "Current SRL",
  "SRL1 Date",
  "SRL2 Date",
  "SRL3 Date",
  "SRL4 Date",
  "SRL5 Date",
  "SRL6 Date",
  "SRL7 Date",
  "Financial Value",
];

export const ORIGINAL_SUMMARY_HEADER = [
  "",
  "Company",
  "PI",
  "RFS/NTI",
  "Current SRL",
  "SRL1",
  "SRL2",
  "SRL3",
  "SRL4",
  "SRL5",
  "SRL6",
  "SRL7",
  "Financial Value",
];

export async function sheetsFetch(
  path: string,
  init?: { method?: string; body?: unknown },
): Promise<any> {
  for (let attempt = 0; attempt < 3; attempt++) {
    const controller = new AbortController();
    // A Google API request should never be allowed to stall the entire website
    // for minutes. If an upstream call hangs, retry promptly and surface a
    // normal sync error instead of leaving every dashboard section waiting.
    const timeout = setTimeout(() => controller.abort(), 20_000);
    try {
      const res = await fetch(`${SHEETS_API}${path}`, {
        method: init?.method ?? "GET",
        headers: await sheetsAuthHeaders(),
        signal: controller.signal,
        ...(init?.body ? { body: JSON.stringify(init.body) } : {}),
      });
      if (res.ok) return res.json();

      const body = await res.text();
      if (res.status === 401 && attempt < 2) {
        invalidateSheetsAccessToken();
        continue;
      }
      const retryable = res.status === 429 || res.status >= 500;
      if (retryable && attempt < 2) {
        const retryAfter = Number(res.headers.get("Retry-After"));
        // Respect a short server hint, but don't let a huge Retry-After freeze
        // the application for several minutes.
        const delayMs = Number.isFinite(retryAfter) && retryAfter > 0
          ? Math.min(retryAfter * 1000, 5_000)
          : 1000 * 2 ** attempt;
        await new Promise((resolve) => setTimeout(resolve, delayMs));
        continue;
      }
      console.error(`Sheets request failed [${res.status}] ${path}: ${body}`);
      throw new Error(`Sheets request failed [${res.status}]: ${body}`);
    } catch (error: any) {
      if (error?.name === "AbortError" && attempt < 2) {
        await new Promise((resolve) => setTimeout(resolve, 500 * (attempt + 1)));
        continue;
      }
      throw error;
    } finally {
      clearTimeout(timeout);
    }
  }
  throw new Error("Sheets request failed after retries");
}

function invalidateSheetMeta() {
  sheetMetaCache = null;
}

export function runSheetOperation<T>(operation: () => Promise<T>): Promise<T> {
  return runSummaryOperation(operation);
}

function runSummaryOperation<T>(operation: () => Promise<T>): Promise<T> {
  const result = summaryOperation.then(operation, operation);
  summaryOperation = result.then(() => undefined, () => undefined);
  return result;
}

export async function getSheetMeta(forceRefresh = false): Promise<
  Array<{ title: string; sheetId: number }>
> {
  if (!forceRefresh && sheetMetaCache && sheetMetaCache.expiresAt > Date.now()) {
    return sheetMetaCache.value;
  }
  if (forceRefresh) {
    sheetMetaCache = null;
  }
  if (!sheetMetaRequest) {
    sheetMetaRequest = sheetsFetch(
      `/spreadsheets/${PROJECTS_SPREADSHEET_ID}?fields=sheets.properties`,
    ).then((data) => {
      const value = (data.sheets || []).map((s: any) => ({
        title: s.properties.title as string,
        sheetId: s.properties.sheetId as number,
      }));
      sheetMetaCache = { expiresAt: Date.now() + META_CACHE_MS, value };
      return value;
    }).finally(() => {
      sheetMetaRequest = null;
    });
  }
  return sheetMetaRequest;
}

// Fast read path used by the website's initial load. The workbook is small
// (one compact summary tab plus ~31 A:C project tabs), so asking the Sheets
// API for value-only grid data returns every project in a single upstream
// request. This avoids the old metadata request + 31-range batchGet path,
// which could be slow when performed as many separate API calls. If Google changes
// the response shape or Google rejects this request, callers can fall
// back to the existing batchGet implementation.
const GRID_CACHE_MS = 15 * 1000;
let gridValueCache: { expiresAt: number; value: Array<{ title: string; sheetId: number; values: unknown[][] }> } | null = null;
let gridValueRequest: Promise<Array<{ title: string; sheetId: number; values: unknown[][] }>> | null = null;

export async function getAllSheetValuesFast(): Promise<
  Array<{ title: string; sheetId: number; values: unknown[][] }>
> {
  if (gridValueCache && gridValueCache.expiresAt > Date.now()) return gridValueCache.value;
  if (gridValueRequest) return gridValueRequest;

  const fields =
    "sheets(properties(title,sheetId),data(startRow,startColumn,rowData.values(formattedValue)))";
  gridValueRequest = sheetsFetch(
    `/spreadsheets/${PROJECTS_SPREADSHEET_ID}?includeGridData=true&fields=${encodeURIComponent(fields)}`,
  )
    .then((data) => {
      const result = (data.sheets || []).map((sheet: any) => {
        const title = String(sheet?.properties?.title || "");
        const sheetId = Number(sheet?.properties?.sheetId);
        const values: unknown[][] = [];

        for (const grid of sheet?.data || []) {
          const startRow = Number(grid?.startRow || 0);
          const startColumn = Number(grid?.startColumn || 0);
          const rows = grid?.rowData || [];
          rows.forEach((row: any, rowOffset: number) => {
            const targetRow = startRow + rowOffset;
            if (!values[targetRow]) values[targetRow] = [];
            (row?.values || []).forEach((cell: any, colOffset: number) => {
              const targetCol = startColumn + colOffset;
              values[targetRow][targetCol] = cell?.formattedValue ?? "";
            });
          });
        }
        return { title, sheetId, values };
      });

      // Reuse the same response for metadata lookups made by later edits.
      const meta = result.map(({ title, sheetId }) => ({ title, sheetId }));
      sheetMetaCache = { expiresAt: Date.now() + META_CACHE_MS, value: meta };
      gridValueCache = { expiresAt: Date.now() + GRID_CACHE_MS, value: result };
      return result;
    })
    .finally(() => {
      gridValueRequest = null;
    });

  return gridValueRequest;
}

export function quoteSheetTitle(title: string): string {
  return `'${title.replace(/'/g, "''")}'`;
}

export async function findProjectSheet(
  sheets: SheetMeta[],
  projectId: string,
  projectName: string,
  allowProjectNameMatch = true,
): Promise<SheetMeta | undefined> {
  const detailSheets = sheets.filter((sheet) => sheet.title !== SUMMARY_SHEET);
  if (!detailSheets.length) return undefined;

  const query = detailSheets
    .map((sheet) => `ranges=${encodeURIComponent(`${quoteSheetTitle(sheet.title)}!A1:G100`)}`)
    .join("&");
  const data = await sheetsFetch(
    `/spreadsheets/${PROJECTS_SPREADSHEET_ID}/values:batchGet?${query}`,
  );
  const requestedId = normalized(projectId);
  const requestedName = normalized(projectName);
  const inspected = detailSheets.map((sheet, index) => ({
    sheet,
    values: (data.valueRanges?.[index]?.values || []) as unknown[][],
  }));

  const byProjectId = inspected.find(({ values }) =>
    values.some((row) =>
      [0, 4].some(
        (labelColumn) =>
          normalized(row?.[labelColumn]) === "project id" &&
          normalized(row?.[labelColumn + 1]) === requestedId,
      ),
    ),
  );
  if (byProjectId) return byProjectId.sheet;

  if (!allowProjectNameMatch) return undefined;

  const byProjectName = inspected.filter(({ values }) =>
    values.some((row) =>
      [0, 4].some(
        (labelColumn) =>
          normalized(row?.[labelColumn]) === "title of the project" &&
          normalized(row?.[labelColumn + 1]) === requestedName,
      ),
    ),
  );
  if (byProjectName.length === 1) return byProjectName[0].sheet;
  // A tab title alone is not project identity. In particular, legacy tabs may
  // be named after a company, so reusing a title can overwrite another project
  // for that company. Only content-level ID/title matches are safe to update.
  return undefined;
}

export async function batchUpdate(requests: unknown[]) {
  if (!requests.length) return null;
  return sheetsFetch(`/spreadsheets/${PROJECTS_SPREADSHEET_ID}:batchUpdate`, {
    method: "POST",
    body: { requests },
  });
}

export async function ensureSummarySheet(): Promise<number> {
  const sheets = await getSheetMeta();
  let summary = sheets.find((s) => s.title === SUMMARY_SHEET);
  if (!summary) {
    const res = await batchUpdate([
      { addSheet: { properties: { title: SUMMARY_SHEET } } },
    ]);
    const sheetId = res.replies[0].addSheet.properties.sheetId as number;
    summary = { title: SUMMARY_SHEET, sheetId };
    invalidateSheetMeta();
    // New sheets use the same original layout as the supplied workbook.
    await sheetsFetch(
      `/spreadsheets/${PROJECTS_SPREADSHEET_ID}/values/${SUMMARY_SHEET}!A1:M1?valueInputOption=RAW`,
      { method: "PUT", body: { values: [ORIGINAL_SUMMARY_HEADER] } },
    );
    return sheetId;
  }

  // IMPORTANT: never rewrite an existing summary header. The original file's
  // column A is a row number, not Project ID. Replacing that header was one of
  // the causes of the website/sheet mismatch and duplicate projects.
  return summary.sheetId;
}


function isModernSummary(values: unknown[][]): boolean {
  return normalized(values?.[0]?.[0]) === "project id";
}

function summaryDataRows(values: unknown[][]): number[] {
  const rows: number[] = [];
  const modern = isModernSummary(values);
  for (let index = 1; index < values.length; index++) {
    const row = values[index] || [];
    const hasData = modern
      ? Boolean(String(row[0] ?? "").trim())
      : Boolean(String(row[1] ?? "").trim());
    if (!hasData) continue;
    if (String(row[11] ?? "").trim() === "Total Revenue") continue;
    rows.push(index + 1);
  }
  return rows;
}



async function readSummaryValues(): Promise<unknown[][]> {
  const data = await sheetsFetch(
    `/spreadsheets/${PROJECTS_SPREADSHEET_ID}/values/${SUMMARY_SHEET}!A:M`,
  );
  return data.values || [];
}

async function clearOldRevenueTotals(values: unknown[][]): Promise<void> {
  const totalRows = values.flatMap((row, index) =>
    String(row?.[11] ?? "").trim() === "Total Revenue" ? [index + 1] : [],
  );
  if (!totalRows.length) return;
  const ranges = totalRows.map((row) => `${SUMMARY_SHEET}!L${row}:M${row}`);
  await sheetsFetch(`/spreadsheets/${PROJECTS_SPREADSHEET_ID}/values:batchClear`, {
    method: "POST",
    body: { ranges },
  });
}

type SummaryIdentity = {
  company: string;
  pi: string;
  rfsNti: string;
};

function normalized(value: unknown): string {
  return String(value ?? "").trim().toLowerCase();
}

function isNumericProjectId(value: unknown): boolean {
  return /^\d+$/.test(String(value ?? "").trim());
}

export async function upsertSummaryValues(
  row: unknown[],
  identity: SummaryIdentity,
  createIfMissing: boolean,
): Promise<{ updated: boolean; row: number; projectId: string } | { updated: false; row: 0; projectId: string; missing: true }> {
  return runSummaryOperation(async () => {
    await ensureSummarySheet();
    const values = await readSummaryValues();
    const modern = isModernSummary(values);
    const dataRows = summaryDataRows(values);
    const requestedId = String(row[0] ?? "").trim();

    const exactMatches = modern
      ? dataRows.filter((rowNumber) => String(values[rowNumber - 1]?.[0] ?? "").trim() === requestedId)
      : [];
    const identityMatches = dataRows.filter((rowNumber) => {
      const existingRow = values[rowNumber - 1] || [];
      const companyMatches = normalized(existingRow[1]) === normalized(identity.company);
      const piMatches = normalized(existingRow[2]) === normalized(identity.pi);
      const requestedRfs = normalized(identity.rfsNti);
      const rfsMatches = requestedRfs ? normalized(existingRow[3]) === requestedRfs : true;
      return companyMatches && piMatches && rfsMatches;
    });

    // For the original summary schema there is no true Project ID column.
    // A newly-created website project must therefore APPEND a new row instead
    // of reusing a row merely because company/PI/RFS happen to match. Reusing
    // by identity is exactly how multiple projects under one company can vanish.
    const existing = modern
      ? (exactMatches[0] ?? (identityMatches.length === 1 ? identityMatches[0] : undefined))
      : (createIfMissing ? undefined : (identityMatches.length === 1 ? identityMatches[0] : undefined));
    if (!existing && !createIfMissing) {
      return { updated: false, row: 0, projectId: requestedId, missing: true as const };
    }

    const targetRow = existing ?? ((dataRows.at(-1) ?? 1) + 1);
    const lastDataRow = Math.max(targetRow, dataRows.at(-1) ?? 1);
    let valuesToWrite: unknown[];
    let returnedProjectId = requestedId;

    if (modern) {
      const canonicalProjectId = existing
        ? String(values[existing - 1]?.[0] ?? requestedId).trim() || requestedId
        : requestedId;
      valuesToWrite = [...row];
      valuesToWrite[0] = canonicalProjectId;
      returnedProjectId = canonicalProjectId;
    } else {
      // Preserve the original workbook: A remains its visible row number and
      // B:M keep the original Company..Financial Value schema.
      const priorSerial = existing ? String(values[existing - 1]?.[0] ?? "").trim() : "";
      const serial = priorSerial || String(dataRows.length + 1);
      valuesToWrite = [serial, ...row.slice(1, 13)];
    }

    await clearOldRevenueTotals(values);
    await sheetsFetch(`/spreadsheets/${PROJECTS_SPREADSHEET_ID}/values:batchUpdate`, {
      method: "POST",
      body: {
        valueInputOption: "USER_ENTERED",
        data: [
          { range: `${SUMMARY_SHEET}!A${targetRow}:M${targetRow}`, values: [valuesToWrite] },
          {
            range: `${SUMMARY_SHEET}!L${lastDataRow + 2}:M${lastDataRow + 2}`,
            values: [["Total Revenue", `=SUM(M2:M${lastDataRow})`]],
          },
        ],
      },
    });
    return { updated: Boolean(existing), row: targetRow, projectId: returnedProjectId };
  });
}


export async function deleteSummaryValues(projectId: string): Promise<boolean> {
  return runSummaryOperation(async () => {
    const sheetId = await ensureSummarySheet();
    const values = await readSummaryValues();
    if (!isModernSummary(values)) return false;
    const dataRows = summaryDataRows(values);
    const existing = dataRows.find(
      (rowNumber) => String(values[rowNumber - 1]?.[0] ?? "").trim() === projectId.trim(),
    );
    if (!existing) return false;

    await clearOldRevenueTotals(values);
    await batchUpdate([
      {
        deleteDimension: {
          range: { sheetId, dimension: "ROWS", startIndex: existing - 1, endIndex: existing },
        },
      },
    ]);
    const lastDataRow = Math.max(1, (dataRows.at(-1) ?? 1) - 1);
    if (lastDataRow >= 2) {
      await sheetsFetch(`/spreadsheets/${PROJECTS_SPREADSHEET_ID}/values:batchUpdate`, {
        method: "POST",
        body: {
          valueInputOption: "USER_ENTERED",
          data: [{
            range: `${SUMMARY_SHEET}!L${lastDataRow + 2}:M${lastDataRow + 2}`,
            values: [["Total Revenue", `=SUM(M2:M${lastDataRow})`]],
          }],
        },
      });
    }
    return true;
  });
}

export async function findSummaryRow(projectId: string): Promise<number | null> {
  const data = await sheetsFetch(
    `/spreadsheets/${PROJECTS_SPREADSHEET_ID}/values/${SUMMARY_SHEET}!A:A`,
  );
  const values: string[][] = data.values || [];
  for (let i = 0; i < values.length; i++) {
    if ((values[i]?.[0] || "") === projectId) return i + 1; // 1-based row number
  }
  return null;
}

// Keeps the "Total Revenue" cell just below the data block, always summing the
// full Financial Value column (M2:M<last data row>).
export async function refreshRevenueTotal(): Promise<void> {
  const data = await sheetsFetch(
    `/spreadsheets/${PROJECTS_SPREADSHEET_ID}/values/${SUMMARY_SHEET}!A:A`,
  );
  const values: string[][] = data.values || [];
  let last = 1;
  for (let i = 1; i < values.length; i++) {
    if ((values[i]?.[0] || "").trim() !== "") last = i + 1;
  }
  if (last < 2) return;
  // Clear any previous total placed further down, then write the new one.
  await sheetsFetch(
    `/spreadsheets/${PROJECTS_SPREADSHEET_ID}/values/${SUMMARY_SHEET}!L${last + 1}:M${last + 12}:clear`,
    { method: "POST", body: {} },
  );
  await sheetsFetch(
    `/spreadsheets/${PROJECTS_SPREADSHEET_ID}/values/${SUMMARY_SHEET}!L${last + 2}:M${last + 2}?valueInputOption=USER_ENTERED`,
    { method: "PUT", body: { values: [["Total Revenue", `=SUM(M2:M${last})`]] } },
  );
}

// Applies the shared look of the existing per-project tabs: bold italic labels
// in column A, a bold header row for the SRL table, borders around the table,
// and consistent column widths.
export async function formatProjectSheet(sheetId: number, rowCount: number, tableStartRow?: number) {
  const tableStart = typeof tableStartRow === "number" && tableStartRow > 0 ? tableStartRow : 12; // 0-based row index of "SRL Progression Data"
  const tableEnd = tableStart + 9; // exclusive: through SRL7
  const fundingRow = tableStart - 3; // "Financial Value"
  const borderStyle = { style: "SOLID", width: 1, color: { red: 0, green: 0, blue: 0 } };
  await batchUpdate([
    {
      repeatCell: {
        range: { sheetId, startRowIndex: 0, endRowIndex: tableStart, startColumnIndex: 0, endColumnIndex: 1 },
        cell: { userEnteredFormat: { textFormat: { bold: true, italic: true } } },
        fields: "userEnteredFormat.textFormat",
      },
    },
    {
      repeatCell: {
        range: { sheetId, startRowIndex: tableStart, endRowIndex: rowCount, startColumnIndex: 0, endColumnIndex: 3 },
        cell: { userEnteredFormat: { textFormat: { bold: true, italic: false } } },
        fields: "userEnteredFormat.textFormat",
      },
    },
    {
      repeatCell: {
        range: { sheetId, startRowIndex: tableStart, endRowIndex: tableStart + 1, startColumnIndex: 0, endColumnIndex: 3 },
        cell: { userEnteredFormat: { horizontalAlignment: "CENTER" } },
        fields: "userEnteredFormat.horizontalAlignment",
      },
    },
    {
      repeatCell: {
        range: { sheetId, startRowIndex: tableStart + 1, endRowIndex: tableEnd, startColumnIndex: 1, endColumnIndex: 3 },
        cell: { userEnteredFormat: { horizontalAlignment: "CENTER", textFormat: { bold: false } } },
        fields: "userEnteredFormat.horizontalAlignment,userEnteredFormat.textFormat.bold",
      },
    },
    {
      updateBorders: {
        range: { sheetId, startRowIndex: tableStart, endRowIndex: tableEnd, startColumnIndex: 0, endColumnIndex: 3 },
        top: borderStyle,
        bottom: borderStyle,
        left: borderStyle,
        right: borderStyle,
        innerHorizontal: borderStyle,
        innerVertical: borderStyle,
      },
    },
    {
      repeatCell: {
        range: { sheetId, startRowIndex: fundingRow, endRowIndex: fundingRow + 1, startColumnIndex: 1, endColumnIndex: 2 },
        cell: { userEnteredFormat: { numberFormat: { type: "CURRENCY", pattern: "$#,##0" } } },
        fields: "userEnteredFormat.numberFormat",
      },
    },
    {
      updateDimensionProperties: {
        range: { sheetId, dimension: "COLUMNS", startIndex: 0, endIndex: 1 },
        properties: { pixelSize: 240 },
        fields: "pixelSize",
      },
    },
    {
      updateDimensionProperties: {
        range: { sheetId, dimension: "COLUMNS", startIndex: 1, endIndex: 3 },
        properties: { pixelSize: 210 },
        fields: "pixelSize",
      },
    },
  ]);
}


// Mirrors the client-side sanitizeSheetName: strip illegal chars, cap at 31,
// then de-duplicate against titles already in the spreadsheet.
export function sanitizeTabName(name: string, used: Set<string>): string {
  const base =
    String(name || "Project")
      .replace(/[\\/?*[\]:]/g, "")
      .trim()
      .slice(0, 31) || "Project";
  let candidate = base;
  let n = 2;
  while (used.has(candidate.toLowerCase())) {
    const suffix = " " + n;
    candidate = (base.slice(0, 31 - suffix.length) + suffix).trim();
    n++;
  }
  return candidate;
}

export function noteSheetAdded(title: string, sheetId: number) {
  gridValueCache = null;
  if (!sheetMetaCache) return;
  const withoutTitle = sheetMetaCache.value.filter((sheet) => sheet.title !== title);
  sheetMetaCache = {
    expiresAt: Date.now() + META_CACHE_MS,
    value: [...withoutTitle, { title, sheetId }],
  };
}

export function noteSheetDeleted(sheetId: number) {
  gridValueCache = null;
  if (!sheetMetaCache) return;
  sheetMetaCache = {
    expiresAt: Date.now() + META_CACHE_MS,
    value: sheetMetaCache.value.filter((sheet) => sheet.sheetId !== sheetId),
  };
}
