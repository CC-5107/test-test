import { createServerFn } from "@tanstack/react-start";
import { SHEETS_API_BASE, sheetsAuthHeaders, getServerEnv } from "./googleSheetsAuth.server";

const SPREADSHEET_ID =
  getServerEnv("CRM_SPREADSHEET_ID") ||
  "1lwLc4FepNQDAyRCWEbq5EDZYLjyE7SAGdzYO99YJHy8";
const RANGE = "Active!A1:Z1";
const SHEETS_API = SHEETS_API_BASE;

type RowPayload = {
  kind: "contact" | "project";
  companyName: string;
  industry?: string;
  hqLocation?: string;
  sbuLocation?: string;
  contactName?: string;
  jobTitle?: string;
  phone?: string;
  email?: string;
  srlLevel?: string;
  contractValue?: string;
  comments?: string;
  source?: string;
};

function splitName(full: string): { first: string; last: string } {
  const parts = (full || "").trim().split(/\s+/);
  if (parts.length === 0 || !parts[0]) return { first: "", last: "" };
  if (parts.length === 1) return { first: parts[0], last: "" };
  return { first: parts[0], last: parts.slice(1).join(" ") };
}

function entryDate(): string {
  const d = new Date();
  const month = d.toLocaleString("en-US", { month: "short" });
  const year = String(d.getFullYear()).slice(-2);
  return `${month}-${year}`;
}

export const appendActiveRow = createServerFn({ method: "POST" })
  .inputValidator((data: RowPayload) => data)
  .handler(async ({ data }) => {
    let authHeaders: Record<string, string>;
    try {
      authHeaders = await sheetsAuthHeaders();
    } catch (error) {
      console.error("Google Sheets service account not configured", error);
      return { ok: false, kind: data.kind, error: "Google Sheets service account not configured" };
    }

    const { first, last } = splitName(data.contactName || "");
    // Columns A..Z (26 total)
    const row = [
      "",                       // A (blank/#)
      data.companyName || "",   // B Company Name
      data.industry || "",      // C Industry
      "",                       // D HQ Country
      data.hqLocation || "",    // E HQ Location
      data.sbuLocation || "",   // F SBU Location
      "",                       // G SBU Country
      "",                       // H Title
      first,                    // I First Name
      last,                     // J Last Name
      data.jobTitle || "",      // K Job Title
      data.phone || "",         // L Phone
      data.email || "",         // M Email
      "",                       // N Primary contact
      "",                       // O Follow Up
      "",                       // P Method of Contact
      data.srlLevel || "",      // Q Status (SRL)
      "",                       // R Faculty 1
      "",                       // S Faculty 2
      "",                       // T Faculty 3
      "",                       // U Research Interests
      data.contractValue || "", // V Contract Value
      data.comments || "",      // W Comments
      "",                       // X blank
      data.source || "",        // Y Source (Initials)
      entryDate(),              // Z Entry Date
    ];

    const url =
      `${SHEETS_API}/spreadsheets/${SPREADSHEET_ID}/values/${RANGE}:append` +
      `?valueInputOption=USER_ENTERED&insertDataOption=INSERT_ROWS`;

    const res = await fetch(url, {
      method: "POST",
      headers: authHeaders,
      body: JSON.stringify({ values: [row] }),
    });

    if (!res.ok) {
      const body = await res.text();
      console.error(`Sheets append failed [${res.status}]: ${body}`);
      // Don't throw: a CRM-sheet permission problem must not break saving.
      return { ok: false, kind: data.kind, error: `Sheets append failed [${res.status}]: ${body}` };
    }
    return { ok: true, kind: data.kind };
  });
