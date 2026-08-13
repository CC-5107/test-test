import { createServerFn } from "@tanstack/react-start";
import {
  PROJECTS_SPREADSHEET_ID,
  SUMMARY_SHEET,
  batchUpdate,
  deleteSummaryValues,
  findProjectSheet,
  formatProjectSheet,
  getSheetMeta,
  noteSheetAdded,
  noteSheetDeleted,
  quoteSheetTitle,
  runSheetOperation,
  sanitizeTabName,
  sheetsFetch,
  upsertSummaryValues,
} from "./projectSheets.server";

// ---------------------------------------------------------------------------
// Polling read-back (sheet -> app), part 1: the SRL_OnePage_Summary tab
// (company, PI, RFS/NTI, SRL level + dates, financial value). Cheap enough
// to poll every 30s. See getProjectDetailSnapshots below for the funding,
// contacts, and comments that only live on each project's own detail tab.
export type SummarySnapshotRow = {
  projectId: string;
  sheetTabName: string; // exact detail tab inferred from the summary row formulas
  company: string;
  pi: string;
  rfsNti: string;
  currentSrl: string;
  srlDates: string[]; // 7 entries, SRL1..SRL7, raw cell text (may be empty)
  financialValue: string;
};

// The original summary does not have a project ID. Its Company column is the
// canonical company name, so it is NOT unique for companies with multiple
// projects (SABIC, Aramco, Maaden, Saudi Diesel, ...). The cells in D/F:L/M
// are formulas that point to the real detail tab, though. Reading the same
// range once with valueRenderOption=FORMULA gives us a stable project key
// without changing the spreadsheet layout or connection.
function sheetNameFromFormula(value: unknown): string {
  const formula = String(value ?? "").trim();
  if (!formula.startsWith("=")) return "";

  // Quoted Google Sheets tab name, e.g. ='Saudi Diesel'!$B$10. Doubled
  // apostrophes are the spreadsheet escape for a literal apostrophe.
  const quoted = formula.match(/^='((?:[^']|'')+)'!/);
  if (quoted) return quoted[1].replace(/''/g, "'");

  // Unquoted tab name, e.g. =SABIC!$B$10.
  const plain = formula.match(/^=([^!]+)!/);
  return plain ? plain[1].trim() : "";
}

function inferSummaryDetailTab(formulaRow: unknown[]): string {
  // Financial Value is the best single identity reference in the original
  // workbook. Then prefer SRL date links; use RFS last because a few legacy
  // rows contain copy/paste mistakes in that particular formula.
  const preferredColumns = [12, 5, 6, 7, 8, 9, 10, 11, 3];
  for (const index of preferredColumns) {
    const name = sheetNameFromFormula(formulaRow?.[index]);
    if (name) return name;
  }
  return "";
}

export const getSummarySnapshot = createServerFn({ method: "GET" }).handler(async () => {
  try {
    // Read the header too. The original workbook has a row-number column in A
    // (blank/# header), while newer app-written versions may have "Project ID"
    // in A. Treat column A as an ID only when the header explicitly says so.
    const [data, formulaData] = await Promise.all([
      sheetsFetch(`/spreadsheets/${PROJECTS_SPREADSHEET_ID}/values/${SUMMARY_SHEET}!A1:M`),
      sheetsFetch(
        `/spreadsheets/${PROJECTS_SPREADSHEET_ID}/values/${SUMMARY_SHEET}!A1:M?valueRenderOption=FORMULA`,
      ),
    ]);
    const values: unknown[][] = data.values || [];
    const formulaValues: unknown[][] = formulaData.values || [];
    const header = values[0] || [];
    const hasProjectIdColumn = normalizeLabel(header[0]) === "project id";
    const rows: SummarySnapshotRow[] = values
      .slice(1)
      .map((row, index) => ({ row, formulaRow: formulaValues[index + 1] || [] }))
      .filter(({ row }) => String(row?.[1] ?? "").trim() && String(row?.[11] ?? "").trim() !== "Total Revenue")
      .map(({ row, formulaRow }) => ({
        projectId: hasProjectIdColumn ? String(row[0] ?? "").trim() : "",
        sheetTabName: inferSummaryDetailTab(formulaRow),
        company: String(row[1] ?? "").trim(),
        pi: String(row[2] ?? "").trim(),
        rfsNti: String(row[3] ?? "").trim(),
        currentSrl: String(row[4] ?? "").trim(),
        srlDates: [5, 6, 7, 8, 9, 10, 11].map((i) => String(row[i] ?? "").trim()),
        financialValue: String(row[12] ?? "").trim(),
      }));
    return { ok: true as const, rows };
  } catch (error) {
    console.error("Sheet poll failed:", error);
    return { ok: false as const, rows: [] as SummarySnapshotRow[], error: "Could not reach Google Sheets." };
  }
});

function parseSummaryValues(values: unknown[][]): SummarySnapshotRow[] {
  if (!values.length) return [];
  const header = values[0] || [];
  const labels = header.map((v) => normalizeLabel(v));
  const findColumn = (predicate: (label: string) => boolean, fallback: number) => {
    const idx = labels.findIndex(predicate);
    return idx >= 0 ? idx : fallback;
  };

  const idCol = labels.findIndex((label) => label.includes("project") && label.includes("id"));
  const companyCol = findColumn((label) => label === "company" || label.includes("company name"), 1);
  const piCol = findColumn((label) => label === "pi" || label.includes("lead pi"), 2);
  const rfsCol = findColumn((label) => label.includes("rfs") || label.includes("nti"), 3);
  const currentSrlCol = findColumn((label) => label.includes("current") && label.includes("srl"), 4);
  const financialCol = findColumn((label) => label.includes("financial") && label.includes("value"), 12);
  const srlCols = Array.from({ length: 7 }, (_, i) => {
    const level = i + 1;
    return findColumn((label) => new RegExp(`^srl\\s*0*${level}(?:\\s|$)`).test(label), 5 + i);
  });

  return values.slice(1).flatMap((row, index) => {
    const company = String(row?.[companyCol] ?? "").trim();
    const financial = String(row?.[financialCol] ?? "").trim();
    const rowText = row.map((v) => String(v ?? "").trim().toLowerCase()).join(" | ");
    if (!company || rowText.includes("total revenue")) return [];
    const rawId = idCol >= 0 ? String(row?.[idCol] ?? "").trim() : "";
    return [{
      projectId: rawId || String(index + 1),
      sheetTabName: "",
      company,
      pi: String(row?.[piCol] ?? "").trim(),
      rfsNti: String(row?.[rfsCol] ?? "").trim(),
      currentSrl: String(row?.[currentSrlCol] ?? "").trim(),
      srlDates: srlCols.map((col) => String(row?.[col] ?? "").trim()),
      financialValue: financial,
    }];
  });
}

// ---------------------------------------------------------------------------
// Detail-tab read-back (sheet -> app), part 2. Each project's own tab is a
// label/value/comment sheet (columns A/B/C) — this is where funding, notes,
// and extra contacts actually live, since the summary tab only has room for
// the compact SRL_OnePage_Summary columns. Polled on a slower cadence than
// the summary tab (see SrlApp.jsx) since it reads every detail tab in one
// batchGet, which costs more than a single-range read.
export type ContactDetail = {
  name: string;
  jobTitle: string;
  email: string;
  phone: string;
};

export type ProjectDetailSnapshot = {
  sheetTabName: string;
  projectId: string;
  companyName: string;
  leadPi: string;
  title: string;
  rfsNumber: string;
  financialValue: string;
  startDate: string; // raw text from the "Start Date" row (DD/MM/YYYY, as written by the app)
  deadline: string; // raw text from the "Deadline" row
  potential: string;
  status: string;
  primaryContact: ContactDetail;
  extraContacts: ContactDetail[];
  srlDates: string[]; // 7 entries, read directly from rows SRL1..SRL7
  srlComments: Record<string, string>; // level ("1".."7") -> comment text
  poolComment: string;
};

// Normalizes a label for matching: lowercase, trim, collapse whitespace.
function normalizeLabel(value: unknown): string {
  return String(value ?? "")
    .trim()
    .toLowerCase()
    .replace(/[_–—-]+/g, " ")
    .replace(/[():]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}


function splitOriginalContactCells(nameRaw: string, emailRaw: string): ContactDetail[] {
  let name = String(nameRaw || "").trim();
  let email = String(emailRaw || "").trim();
  const looksLikeEmail = (value: string) => /[^\s@]+@[^\s@]+\.[^\s@]+/.test(value);

  // Be forgiving of a legacy row where a person's name was accidentally put
  // in the email cell (or vice versa). The import-spreadsheet path is tolerant
  // of these human data-entry inconsistencies, so live sync should be too.
  if (!name && email && !looksLikeEmail(email)) {
    name = email;
    email = "";
  } else if (looksLikeEmail(name) && email && !looksLikeEmail(email)) {
    const tmp = name;
    name = email;
    email = tmp;
  }

  const emails = email
    .split(/[;\n]+|,\s*(?=[^,;\s]+@)/)
    .map((v) => v.trim())
    .filter(Boolean);

  // Split only on separators that clearly represent another person. A plain
  // comma remains part of names/affiliations such as "Samuel, SDEC" and
  // "Dr. Heechang Oh, Hyundai". A comma before another honorific, semicolon,
  // newline, or ampersand can represent a second person.
  let names = name
    .split(/;|\n|\s+&\s+|,\s*(?=(?:dr\.?|doctor|mr\.?|mrs\.?|ms\.?|prof\.?|professor)\s+)/i)
    .map((v) => v.trim())
    .filter(Boolean);

  if (emails.length > 1) {
    if (names.length !== emails.length) {
      names = Array.from({ length: emails.length }, (_, i) => names[i] || (i === 0 ? name : ""));
    }
    return emails.map((value, i) => ({ name: names[i] || "", jobTitle: "", email: value, phone: "" }));
  }

  if (names.length > 1) {
    return names.map((person, i) => ({
      name: person,
      jobTitle: "",
      email: i === 0 ? (emails[0] || "") : "",
      phone: "",
    }));
  }

  if (!name && !email) return [];
  return [{ name, jobTitle: "", email: emails[0] || email, phone: "" }];
}

function parseDetailTab(sheetTitle: string, values: unknown[][]): ProjectDetailSnapshot {
  const map = new Map<string, { b: string; c: string }>();

  // Read values only — never formatting. Font, bold/italic, fill/highlight,
  // alignment, borders, and number-format styling have no effect on parsing.
  // A:C is the original workbook block. E:G is reserved by the website for
  // additional contacts so the original 23-row layout never has to move.
  const looksLikeProjectLabel = (label: string) => {
    if (!label) return false;
    return (
      label.includes("project id") ||
      label.includes("company") ||
      label === "lead pi" ||
      (label.includes("lead") && label.includes("pi")) ||
      (label.includes("title") && label.includes("project")) ||
      label.includes("industry contact") ||
      label.includes("rfs") ||
      (label.includes("financial") && label.includes("value")) ||
      (label.includes("expected") && (label.includes("start") || label.includes("strat") || label.includes("end") || label.includes("ebd"))) ||
      label === "start date" ||
      label === "deadline" ||
      label === "potential" ||
      label === "status" ||
      label === "pool" ||
      /^srl\s*0*[1-7]\b/.test(label)
    );
  };

  // Do not depend on a label being in column A or E. Website-created sheets,
  // manually adjusted sheets, and copied templates can move cells around. We
  // scan every populated cell for a recognized semantic label and treat the
  // next two cells on that row as value/comment. Formatting (font, boldness,
  // fill, alignment, borders, number format) is never requested or inspected.
  values.forEach((row) => {
    for (let labelColumn = 0; labelColumn < row.length; labelColumn++) {
      const label = normalizeLabel(row?.[labelColumn]);
      if (!looksLikeProjectLabel(label)) continue;
      const b = String(row?.[labelColumn + 1] ?? "").trim();
      const c = String(row?.[labelColumn + 2] ?? "").trim();
      const existing = map.get(label);
      if (!existing || b || c) map.set(label, { b, c });
    }
  });
  const get = (label: string) => map.get(normalizeLabel(label))?.b ?? "";
  const getComment = (label: string) => map.get(normalizeLabel(label))?.c ?? "";
  const findCell = (predicate: (label: string) => boolean) => {
    for (const [label, cell] of map.entries()) if (predicate(label)) return cell;
    return undefined;
  };
  const getLoose = (exact: string, predicate: (label: string) => boolean) =>
    get(exact) || findCell(predicate)?.b || "";

  const srlDates = Array.from({ length: 7 }, () => "");
  const srlComments: Record<string, string> = {};
  for (const [label, cell] of map.entries()) {
    const m = label.match(/^srl\s*0*([1-7])\b/);
    if (!m) continue;
    const levelIndex = Number(m[1]) - 1;
    srlDates[levelIndex] = cell.b;
    if (cell.c) srlComments[m[1]] = cell.c;
  }

  const isNumberedContactLabel = (label: string) => /\bindustry contact\s+\d+\b/.test(label);
  const contactName = getLoose(
    "industry contact",
    (label) => label.includes("industry contact") && !label.includes("email") && !label.includes("phone") && !label.includes("job") && !isNumberedContactLabel(label),
  );
  const contactEmail = getLoose(
    "industry contact email",
    (label) => label.includes("industry contact") && label.includes("email") && !isNumberedContactLabel(label),
  );
  const originalContacts = splitOriginalContactCells(contactName, contactEmail);
  const primaryContact = originalContacts[0] || { name: "", jobTitle: "", email: "", phone: "" };
  primaryContact.jobTitle = getLoose(
    "industry contact job title",
    (label) => label.includes("industry contact") && label.includes("job") && label.includes("title") && !isNumberedContactLabel(label),
  );
  primaryContact.phone = getLoose(
    "industry contact phone",
    (label) => label.includes("industry contact") && (label.includes("phone") || label.includes("number")) && !isNumberedContactLabel(label),
  );

  const extraContacts: ContactDetail[] = originalContacts.slice(1);
  for (let n = 2; n <= 100; n++) {
    const name = get(`industry contact ${n}`);
    const jobTitle = get(`industry contact ${n} job title`);
    const email = get(`industry contact ${n} email`);
    const phone = get(`industry contact ${n} phone`);
    // Do not stop at a gap; manually-edited sheets sometimes skip a contact
    // number. Later numbered contacts should still be imported.
    if (!name && !jobTitle && !email && !phone) continue;
    extraContacts.push({ name, jobTitle, email, phone });
  }

  return {
    sheetTabName: sheetTitle,
    projectId: getLoose("project id", (label) => label.includes("project") && label.includes("id")),
    companyName: getLoose("company name", (label) => label.includes("company") && label.includes("name")),
    leadPi: getLoose("lead pi", (label) => label.includes("lead") && label.includes("pi")),
    title: getLoose("title of the project", (label) => label.includes("title") && label.includes("project")),
    rfsNumber: getLoose("if yes, rfs number", (label) => label.includes("rfs") && label.includes("number")),
    financialValue: getLoose("financial value", (label) => label.includes("financial") && label.includes("value")),
    // Support both the app's modern labels and the untouched original
    // workbook labels (including its historical Strat/Ebd typos), plus small
    // wording differences such as "Expected Start" / "Expected Ebd".
    startDate:
      get("start date") ||
      get("expected start date") ||
      get("expected strat date") ||
      findCell((label) => label.includes("expected") && (label.includes("start") || label.includes("strat")))?.b ||
      "",
    deadline:
      get("deadline") ||
      get("expected end date") ||
      get("expected ebd date") ||
      findCell((label) => label.includes("expected") && (label.includes("end") || label.includes("ebd")))?.b ||
      "",
    potential: get("potential"),
    status: get("status"),
    primaryContact,
    extraContacts,
    srlDates,
    srlComments,
    poolComment: getComment("pool"),
  };
}

export const getProjectDetailSnapshots = createServerFn({ method: "GET" }).handler(async () => {
  try {
    // Fast/live equivalent of Import Spreadsheet: sheet metadata tells us the
    // complete project inventory and one values.batchGet reads every detail tab
    // in a single request. One non-summary tab is always one project.
    const sheets = await getSheetMeta(true);
    const detailSheets = sheets.filter((s) => s.title !== SUMMARY_SHEET);
    if (!detailSheets.length) {
      return { ok: true as const, snapshots: [] as ProjectDetailSnapshot[] };
    }

    const query = detailSheets
      .map((s) => `ranges=${encodeURIComponent(`${quoteSheetTitle(s.title)}!A1:Z500`)}`)
      .join("&");
    const data = await sheetsFetch(
      `/spreadsheets/${PROJECTS_SPREADSHEET_ID}/values:batchGet?majorDimension=ROWS&${query}`,
    );
    const valueRanges = data.valueRanges || [];

    // Map by the metadata list so even an empty detail tab remains a distinct
    // project instead of shifting/misaligning the following tabs.
    const snapshots = detailSheets.map((sheet, i) =>
      parseDetailTab(sheet.title, (valueRanges[i]?.values || []) as unknown[][]),
    );
    return { ok: true as const, snapshots };
  } catch (error) {
    console.error("Detail sheet poll failed:", error);
    return {
      ok: false as const,
      snapshots: [] as ProjectDetailSnapshot[],
      error: "Could not reach Google Sheets.",
    };
  }
});

type SummaryPayload = {
  projectId: string;
  company: string;
  pi: string;
  rfsNti: string;
  currentSrl: string;
  srlDates: string[]; // 7 entries, SRL1..SRL7
  financialValue: string;
  sheetTabName?: string;
  createIfMissing?: boolean;
};

function buildSummaryWriteRow(data: SummaryPayload, tabOverride?: string): unknown[] {
  const dates = Array.from({ length: 7 }, (_, i) => data.srlDates[i] || "");
  const tab = String(tabOverride || data.sheetTabName || "").trim();
  if (tab) {
    const q = quoteSheetTitle(tab);
    const currentSrlFormula =
      `=IF(${q}!B21<>"",7,IF(${q}!B20<>"",6,IF(${q}!B19<>"",5,IF(${q}!B18<>"",4,IF(${q}!B17<>"",3,IF(${q}!B16<>"",2,IF(${q}!B15<>"",1,"")))))))`;
    return [
      data.projectId,
      `=${q}!B1`,
      `=${q}!B2`,
      `=${q}!B9`,
      currentSrlFormula,
      ...Array.from({ length: 7 }, (_, i) => `=${q}!B${15 + i}`),
      `=${q}!B10`,
    ];
  }
  return [
    data.projectId,
    data.company,
    data.pi,
    data.rfsNti,
    data.currentSrl,
    ...dates,
    data.financialValue,
  ];
}

export const upsertSummaryRow = createServerFn({ method: "POST" })
  .inputValidator((data: SummaryPayload) => data)
  .handler(async ({ data }) => {
    try {
      const result = await upsertSummaryValues(
        buildSummaryWriteRow(data),
        { company: data.company, pi: data.pi, rfsNti: data.rfsNti },
        data.createIfMissing === true,
      );
      return { ok: true as const, ...result };
    } catch (error: any) {
      console.error("Summary sheet sync failed:", error);
      return { ok: false as const, error: String(error?.message || error || "Could not write the Google Sheet summary.") };
    }
  });

export const deleteSummaryRow = createServerFn({ method: "POST" })
  .inputValidator((data: { projectId: string }) => data)
  .handler(async ({ data }) => {
    try {
      return { ok: true, deleted: await deleteSummaryValues(data.projectId) };
    } catch (error) {
      console.error("Summary sheet delete deferred:", error);
      return { ok: false, deleted: false, error: "Google Sheets is busy; deletion sync can be retried later." };
    }
  });

type ProjectSheetPayload = {
  projectId: string;
  projectName: string;
  preferredName: string;
  alternateNames?: string[];
  sheetTabName?: string;
  rows: string[][];
  primaryContact?: ContactDetail;
  extraContacts?: ContactDetail[];
  potential?: string;
  tableStart?: number;
  createIfMissing?: boolean;
};

function additionalContactBlock(
  projectId: string,
  primaryContact: ContactDetail | undefined,
  contacts: ContactDetail[] | undefined,
  potential: string | undefined,
): string[][] {
  const primary = primaryContact || { name: "", jobTitle: "", email: "", phone: "" };
  const people = (contacts || []).filter((person) =>
    Boolean(
      String(person?.name || "").trim() ||
      String(person?.jobTitle || "").trim() ||
      String(person?.email || "").trim() ||
      String(person?.phone || "").trim()
    ),
  );

  // Keep website-only metadata outside the original A:C template. This makes
  // the detail tab round-trip all website fields without moving the original
  // fixed rows that SRL_OnePage_Summary references.
  const rows: string[][] = [["Project ID", String(projectId || "").trim(), ""]];
  if (String(primary.jobTitle || "").trim()) {
    rows.push(["Industry Contact Job Title", String(primary.jobTitle || "").trim(), ""]);
  }
  if (String(primary.phone || "").trim()) {
    rows.push(["Industry Contact Phone", String(primary.phone || "").trim(), ""]);
  }
  if (String(potential || "").trim()) {
    rows.push(["Potential", String(potential || "").trim(), ""]);
  }
  if (people.length) rows.push(["Additional Contacts", "", ""]);
  people.forEach((person, index) => {
    const n = index + 2; // primary contact is the unnumbered Industry Contact in A6/B6.
    rows.push(
      [`Industry Contact ${n}`, String(person.name || "").trim(), ""],
      [`Industry Contact ${n} Job Title`, String(person.jobTitle || "").trim(), ""],
      [`Industry Contact ${n} Email`, String(person.email || "").trim(), ""],
      [`Industry Contact ${n} Phone`, String(person.phone || "").trim(), ""],
    );
  });
  return rows;
}

async function writeAdditionalContactBlock(
  title: string,
  projectId: string,
  primaryContact: ContactDetail | undefined,
  contacts: ContactDetail[] | undefined,
  potential: string | undefined,
) {
  const blockRange = `${quoteSheetTitle(title)}!E1:G500`;

  // E:G is owned by the website specifically for metadata/additional contacts.
  // Clearing this block is safe and, unlike clearing A:C, never destroys the
  // original project-sheet fields or their fixed row numbers.
  await sheetsFetch(`/spreadsheets/${PROJECTS_SPREADSHEET_ID}/values:batchClear`, {
    method: "POST",
    body: { ranges: [blockRange] },
  });

  const rows = additionalContactBlock(projectId, primaryContact, contacts, potential);
  if (!rows.length) return;

  await sheetsFetch(`/spreadsheets/${PROJECTS_SPREADSHEET_ID}/values:batchUpdate`, {
    method: "POST",
    body: {
      valueInputOption: "USER_ENTERED",
      data: [{
        range: `${quoteSheetTitle(title)}!E1:G${rows.length}`,
        values: rows,
      }],
    },
  });
}

async function writeVisibleAdditionalContacts(
  title: string,
  contacts: ContactDetail[] | undefined,
) {
  // Rows 25+ are the website-owned extension area. Clear VALUES only (never
  // formatting) so deleting/editing additional contacts cannot leave stale
  // contact text behind in Google Sheets.
  await sheetsFetch(`/spreadsheets/${PROJECTS_SPREADSHEET_ID}/values:batchClear`, {
    method: "POST",
    body: { ranges: [`${quoteSheetTitle(title)}!A25:C500`] },
  });

  const people = (contacts || []).filter((person) =>
    Boolean(
      String(person?.name || "").trim() ||
      String(person?.jobTitle || "").trim() ||
      String(person?.email || "").trim() ||
      String(person?.phone || "").trim()
    ),
  );
  if (!people.length) return;

  const rows: string[][] = [["Additional Industry Contacts", "", ""]];
  people.forEach((person, index) => {
    const n = index + 2;
    rows.push(
      [`Industry Contact ${n}`, String(person.name || "").trim(), ""],
      [`Industry Contact ${n} Email`, String(person.email || "").trim(), ""],
      [`Industry Contact ${n} Job Title`, String(person.jobTitle || "").trim(), ""],
      [`Industry Contact ${n} Phone`, String(person.phone || "").trim(), ""],
    );
  });

  // Rows 1-23 remain the untouched legacy template. The visible additional
  // contact section begins at row 25, so no original summary formulas move.
  await sheetsFetch(`/spreadsheets/${PROJECTS_SPREADSHEET_ID}/values:batchUpdate`, {
    method: "POST",
    body: {
      valueInputOption: "USER_ENTERED",
      data: [{
        range: `${quoteSheetTitle(title)}!A25:C${24 + rows.length}`,
        values: rows,
      }],
    },
  });
}

async function upsertProjectSheetData(data: ProjectSheetPayload) {
  return runSheetOperation(async () => {
    const sheets = await getSheetMeta();
    const titles = new Set(sheets.map((s) => s.title.toLowerCase()));
    const exactKnownTab = data.sheetTabName
      ? sheets.find((s) => s.title === data.sheetTabName)
      : undefined;
    const resolvedMatch = exactKnownTab ?? await findProjectSheet(
      sheets,
      data.projectId,
      data.projectName,
      true,
    );
    let title = resolvedMatch?.title;
    let sheetId = resolvedMatch?.sheetId;
    let created = false;

    if (!title) {
      if (data.createIfMissing !== true) {
        return { ok: false as const, missing: true as const, error: "Existing project tab could not be identified; no duplicate tab was created." };
      }
      title = sanitizeTabName(data.preferredName, titles);
      const res = await batchUpdate([{ addSheet: { properties: { title } } }]);
      sheetId = res?.replies?.[0]?.addSheet?.properties?.sheetId as number | undefined;
      if (!title || typeof sheetId !== "number") {
        throw new Error("Google Sheets did not confirm creation of the individual project tab.");
      }
      noteSheetAdded(title, sheetId);
      created = true;
    }

    if (created) {
      await sheetsFetch(`/spreadsheets/${PROJECTS_SPREADSHEET_ID}/values:batchUpdate`, {
        method: "POST",
        body: {
          valueInputOption: "USER_ENTERED",
          data: [{
            range: `${quoteSheetTitle(title)}!A1:C${data.rows.length}`,
            values: data.rows,
          }],
        },
      });
      await writeAdditionalContactBlock(
        title,
        data.projectId,
        data.primaryContact,
        data.extraContacts,
        data.potential,
      );
      await writeVisibleAdditionalContacts(title, data.extraContacts);
      if (typeof sheetId === "number") await formatProjectSheet(sheetId, Math.max(data.rows.length, 24), data.tableStart);
    } else {
      const rowValue = (oneBasedRow: number) => data.rows[oneBasedRow - 1] || ["", "", ""];
      const controlledRows = [1, 2, 5, 6, 7, 8, 9, 10, 11, 12, 15, 16, 17, 18, 19, 20, 21, 22];
      await sheetsFetch(`/spreadsheets/${PROJECTS_SPREADSHEET_ID}/values:batchUpdate`, {
        method: "POST",
        body: {
          valueInputOption: "USER_ENTERED",
          data: controlledRows.map((rowNumber) => ({
            range: `${quoteSheetTitle(title!)}!B${rowNumber}:C${rowNumber}`,
            values: [[rowValue(rowNumber)[1] ?? "", rowValue(rowNumber)[2] ?? ""]],
          })),
        },
      });
      await writeAdditionalContactBlock(
        title!,
        data.projectId,
        data.primaryContact,
        data.extraContacts,
        data.potential,
      );
      await writeVisibleAdditionalContacts(title!, data.extraContacts);
    }

    // Read-after-write verification. A save is not reported as successful just
    // because the HTTP write returned 200; the exact tab must be immediately
    // readable and contain the project/company values we just wrote.
    const verify = await sheetsFetch(
      `/spreadsheets/${PROJECTS_SPREADSHEET_ID}/values/${encodeURIComponent(`${quoteSheetTitle(title!)}!A1:G500`)}`,
    ).catch(() => null);
    const verifyValues = (verify?.values || []) as unknown[][];
    const parsed = parseDetailTab(title!, verifyValues);
    const expectedCompany = String(data.rows?.[0]?.[1] ?? "").trim();
    const expectedTitle = String(data.rows?.[4]?.[1] ?? data.projectName ?? "").trim();
    if (!verify || (expectedCompany && normalizeLabel(parsed.companyName) !== normalizeLabel(expectedCompany))) {
      throw new Error("Google Sheets created the request but the individual project tab could not be verified.");
    }
    if (expectedTitle && parsed.title && normalizeLabel(parsed.title) !== normalizeLabel(expectedTitle)) {
      throw new Error("Google Sheets returned different project data during save verification.");
    }

    return { ok: true as const, sheetTabName: title!, created };
  });
}

export const upsertProjectSheet = createServerFn({ method: "POST" })
  .inputValidator((data: ProjectSheetPayload) => data)
  .handler(async ({ data }) => {
    try {
      return await upsertProjectSheetData(data);
    } catch (error: any) {
      console.error("Project sheet sync failed:", error);
      return { ok: false as const, error: String(error?.message || error || "Could not write the individual project sheet.") };
    }
  });

type PersistProjectPayload = {
  project: ProjectSheetPayload;
  summary: Omit<SummaryPayload, "sheetTabName" | "createIfMissing">;
};

// New project creation goes through one server call. The browser does not move
// away from the form until this function has created + verified the detail tab
// AND appended + verified the summary row. This prevents phantom projects that
// exist only in React state and disappear on refresh.
export const persistNewProject = createServerFn({ method: "POST" })
  .inputValidator((data: PersistProjectPayload) => data)
  .handler(async ({ data }) => {
    try {
      const projectResult = await upsertProjectSheetData({ ...data.project, createIfMissing: true });
      if (!projectResult.ok || !projectResult.sheetTabName) {
        return { ok: false as const, stage: "detail", error: projectResult.error || "Could not create the individual project tab." };
      }

      const summaryPayload: SummaryPayload = {
        ...data.summary,
        sheetTabName: projectResult.sheetTabName,
        createIfMissing: true,
      };
      const expectedTab = normalizeLabel(projectResult.sheetTabName);
      const formulaRefersToTab = (row: unknown[]) =>
        [1, 2, 3, 5, 6, 7, 8, 9, 10, 11, 12].some((column) =>
          normalizeLabel(sheetNameFromFormula(row?.[column])) === expectedTab,
        );

      // Idempotency: if a previous click created the detail tab and appended
      // the summary row but the browser lost the response, do not append a
      // second summary row on retry.
      let formulaData = await sheetsFetch(
        `/spreadsheets/${PROJECTS_SPREADSHEET_ID}/values/${SUMMARY_SHEET}!A1:M?valueRenderOption=FORMULA`,
      );
      let formulaRows = (formulaData.values || []) as unknown[][];
      let existingSummaryIndex = formulaRows.slice(1).findIndex(formulaRefersToTab);
      let summaryRowNumber = existingSummaryIndex >= 0 ? existingSummaryIndex + 2 : 0;

      if (!summaryRowNumber) {
        const summaryResult = await upsertSummaryValues(
          buildSummaryWriteRow(summaryPayload, projectResult.sheetTabName),
          { company: summaryPayload.company, pi: summaryPayload.pi, rfsNti: summaryPayload.rfsNti },
          true,
        );
        summaryRowNumber = summaryResult.row;
      }

      // Verify that the row is readable after the write and points to the exact
      // detail tab. Duplicate company names are therefore irrelevant.
      formulaData = await sheetsFetch(
        `/spreadsheets/${PROJECTS_SPREADSHEET_ID}/values/${SUMMARY_SHEET}!A1:M?valueRenderOption=FORMULA`,
      );
      formulaRows = (formulaData.values || []) as unknown[][];
      existingSummaryIndex = formulaRows.slice(1).findIndex(formulaRefersToTab);
      if (existingSummaryIndex < 0) {
        throw new Error("The individual project tab was created, but SRL_OnePage_Summary did not confirm the new project row.");
      }
      summaryRowNumber = existingSummaryIndex + 2;

      return {
        ok: true as const,
        sheetTabName: projectResult.sheetTabName,
        summaryRow: summaryRowNumber,
      };
    } catch (error: any) {
      console.error("New project persistence failed:", error);
      return {
        ok: false as const,
        error: String(error?.message || error || "Could not save the project to Google Sheets."),
      };
    }
  });

export const deleteProjectSheet = createServerFn({ method: "POST" })
  .inputValidator((data: { sheetTabName: string }) => data)
  .handler(async ({ data }) => {
    try {
      if (!data.sheetTabName) return { ok: true, deleted: false };
      const sheets = await getSheetMeta();
      const match = sheets.find((s) => s.title === data.sheetTabName);
      if (!match) return { ok: true, deleted: false };
      await batchUpdate([{ deleteSheet: { sheetId: match.sheetId } }]);
      noteSheetDeleted(match.sheetId);
      return { ok: true, deleted: true };
    } catch (error) {
      console.error("Project sheet delete deferred:", error);
      return { ok: false, deleted: false, error: "Google Sheets is busy; deletion sync can be retried later." };
    }
  });
