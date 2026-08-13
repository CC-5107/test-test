import { useState, useEffect, useRef } from "react";
import { LogOut, Mail, Lock, User, ArrowLeft, Check, Plus, Phone, Pencil, Upload, Download, Trash2 } from "lucide-react";
import * as XLSX from "xlsx";
import { PieChart, Pie, Cell, Tooltip, ResponsiveContainer } from "recharts";
import { auth, db } from "./firebase";
import {
  onAuthStateChanged,
  signInWithEmailAndPassword,
  createUserWithEmailAndPassword,
  updateProfile,
  signOut,
  sendPasswordResetEmail,
} from "firebase/auth";
import { doc, getDoc, setDoc } from "firebase/firestore";
import { appendActiveRow } from "./lib/sheets.functions";
import {
  deleteSummaryRow,
  upsertProjectSheet,
  deleteProjectSheet,
  getProjectDetailSnapshots,
  getSummarySnapshot,
  persistNewProject,
} from "./lib/projectSheets.functions";
import { loadAppData, saveAppData, deleteAppDataRecords } from "./lib/appData";

// ---------------------------------------------------------------------------
// Preview-only stand-ins
// ---------------------------------------------------------------------------
// This file is the real, deployed version — it talks to the live Google
// Sheets sync backend above (direct Sheets API via a service account; see
// lib/googleSheetsAuth.server.ts) and to real Firebase Authentication +
// Firestore for accounts (see the auth handlers near the top of App(), and
// ./firebase.js for the project config). jsPDF is replaced with a small
// dependency-free shim that implements the handful of methods this file
// actually calls, and renders the same report through the browser's native
// print-to-PDF instead of writing PDF bytes directly.
//
// The standalone Claude-preview copy of this file (SrlApp.jsx at the repo
// root, used to run this app inside Claude) keeps no-op stand-ins for the
// Sheets calls above, and falls back to in-memory auth instead of Firebase,
// since that sandbox has no network access — everything else is identical.

// Minimal jsPDF-compatible shim covering exactly the methods this file uses:
// setFont, setFontSize, setTextColor, setDrawColor, text, line, splitTextToSize,
// addImage, addPage, internal.pageSize.getWidth/getHeight, and save. Instead of
// producing PDF bytes, it lays the same content out as print-ready HTML pages
// (US Letter size) and opens the browser print dialog, where "Save as PDF"
// produces an equivalent file.
class MockPdf {
  constructor() {
    this.pageW = 612; // Letter, points
    this.pageH = 792;
    this.font = { family: "helvetica", style: "normal", size: 12 };
    this.textColor = "rgb(20,24,31)";
    this.drawColor = "rgb(215,218,222)";
    this.pages = [[]];
    this._measureCtx = document.createElement("canvas").getContext("2d");
  }
  get internal() {
    return { pageSize: { getWidth: () => this.pageW, getHeight: () => this.pageH } };
  }
  _cur() {
    return this.pages[this.pages.length - 1];
  }
  _fontCss(sizeOverride) {
    const style = this.font.style === "italic" ? "italic " : "";
    const weight = this.font.style === "bold" ? "bold " : "";
    return style + weight + (sizeOverride || this.font.size) + "pt Helvetica, Arial, sans-serif";
  }
  setFont(family, style) {
    this.font.family = family || "helvetica";
    this.font.style = style || "normal";
  }
  setFontSize(size) {
    this.font.size = size;
  }
  setTextColor(r, g, b) {
    this.textColor = "rgb(" + r + "," + g + "," + b + ")";
  }
  setDrawColor(r, g, b) {
    this.drawColor = "rgb(" + r + "," + g + "," + b + ")";
  }
  setFillColor(r, g, b) {
    this.fillColor = "rgb(" + r + "," + g + "," + b + ")";
  }
  getTextWidth(text) {
    this._measureCtx.font = this._fontCss();
    return this._measureCtx.measureText(String(text)).width;
  }
  text(str, x, y) {
    this._cur().push({ type: "text", str, x, y, font: this._fontCss(), color: this.textColor });
  }
  line(x1, y1, x2, y2) {
    this._cur().push({ type: "line", x1, y1, x2, y2, color: this.drawColor });
  }
  rect(x, y, w, h, style) {
    this._cur().push({
      type: "rect",
      x, y, w, h,
      style: style || "S",
      fillColor: this.fillColor,
      drawColor: this.drawColor,
    });
  }
  splitTextToSize(text, maxWidth) {
    this._measureCtx.font = this._fontCss();
    const words = String(text).split(/\s+/);
    const lines = [];
    let current = "";
    words.forEach((word) => {
      const candidate = current ? current + " " + word : word;
      if (this._measureCtx.measureText(candidate).width > maxWidth && current) {
        lines.push(current);
        current = word;
      } else {
        current = candidate;
      }
    });
    if (current) lines.push(current);
    return lines.length ? lines : [""];
  }
  addImage(dataUrl, format, x, y, w, h) {
    this._cur().push({ type: "image", dataUrl, x, y, w, h });
  }
  addPage() {
    this.pages.push([]);
  }
  save(filename) {
    const pageHtml = this.pages
      .map((elements) => {
        const inner = elements
          .map((el) => {
            if (el.type === "text") {
              return (
                '<div style="position:absolute; left:' + el.x + "pt; top:" + (el.y - 10) +
                'pt; font:' + el.font + "; color:" + el.color + '; white-space:pre;">' +
                escapeHtml(el.str) + "</div>"
              );
            }
            if (el.type === "line") {
              const len = Math.hypot(el.x2 - el.x1, el.y2 - el.y1);
              const angle = Math.atan2(el.y2 - el.y1, el.x2 - el.x1) * (180 / Math.PI);
              return (
                '<div style="position:absolute; left:' + el.x1 + "pt; top:" + el.y1 +
                "pt; width:" + len + "pt; border-top:1pt solid " + el.color +
                "; transform-origin: 0 0; transform: rotate(" + angle + 'deg);"></div>'
              );
            }
            if (el.type === "image") {
              return (
                '<img src="' + el.dataUrl + '" style="position:absolute; left:' + el.x +
                "pt; top:" + el.y + "pt; width:" + el.w + "pt; height:" + el.h + 'pt;" />'
              );
            }
            if (el.type === "rect") {
              const styleParts = [
                "position:absolute",
                "left:" + el.x + "pt",
                "top:" + el.y + "pt",
                "width:" + el.w + "pt",
                "height:" + el.h + "pt",
              ];
              if (el.style === "F" || el.style === "FD" || el.style === "DF") {
                styleParts.push("background:" + (el.fillColor || "transparent"));
              }
              if (el.style === "D" || el.style === "S" || el.style === "FD" || el.style === "DF") {
                styleParts.push("border:1pt solid " + el.drawColor);
              }
              return '<div style="' + styleParts.join("; ") + ';"></div>';
            }
            return "";
          })
          .join("");
        return (
          '<div style="position:relative; width:' + this.pageW + "pt; height:" + this.pageH +
          'pt; page-break-after: always; overflow:hidden;">' + inner + "</div>"
        );
      })
      .join("");

    const html =
      "<!DOCTYPE html><html><head><title>" + escapeHtml(filename) +
      '</title><meta charset="utf-8"><style>' +
      "@page { size: letter; margin: 0; } body { margin: 0; } " +
      "img { display:block; }</style></head><body>" +
      pageHtml +
      "</body></html>";

    const win = window.open("", "_blank");
    if (!win) {
      console.error("Report export: pop-up blocked. Allow pop-ups to download the report.");
      return;
    }
    win.document.open();
    win.document.write(html);
    win.document.close();
    win.onload = () => {
      win.focus();
      win.print();
    };
  }
}

function escapeHtml(str) {
  return String(str)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

// ---------------------------------------------------------------------------
// SRL Project Pipelines — project management app shell
// Login / Sign up / Forgot password / Reset password / Dashboard
// Accounts run on Firebase Authentication (username maps to a real email
// under the hood via the "usernames" Firestore collection). Projects,
// companies, and contacts are persisted in the shared Firestore appData/main
// document for fast, durable refreshes, while the existing Google Sheet remains
// synchronized as the spreadsheet/reporting copy.
// ---------------------------------------------------------------------------

const EMAIL_RE = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

// Projects are stored with status "Finished"/"Unfinished" (unchanged, so
// existing saved data and all the p.status === "Finished" comparisons keep
// working). This only controls what the word looks like on screen — the
// website now calls it "Finalized" everywhere it's shown to the user.
function displayStatus(status) {
  return status === "Finished" ? "Finalized" : status;
}

// Same PI, entered with or without an honorific (e.g. "Mani Sarathy" and
// "Prof. Mani Sarathy") — strip common titles/degrees for comparison so
// reports treat them as the same person instead of two separate entries.
const LEAD_TITLE_RE = /^(prof(?:essor)?|dr|doctor|mr|mrs|ms|mx)\.?\s+/i;

function normalizePersonKey(name) {
  let value = String(name || "").trim();
  // Strip repeated honorifics ("Prof. Dr.") and common degree suffixes so
  // "Mani Sarathy" and "Prof. Mani Sarathy" resolve to the same person.
  while (LEAD_TITLE_RE.test(value)) value = value.replace(LEAD_TITLE_RE, "").trim();
  return value
    .replace(/,?\s+(ph\.?d\.?|md|m\.?sc\.?|mba)\.?$/i, "")
    .replace(/[^a-z0-9]+/gi, " ")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();
}

function normalizeLeadKey(name) {
  return normalizePersonKey(name);
}

// Groups a list of projects by PI, merging title variants of the same name
// into one group (so their projects/funding combine, with no duplicates).
// The display name shown for the group prefers whichever variant includes
// the title, since that's the more complete version of the name.
function buildLeadGroups(projectList) {
  const groups = new Map();
  projectList.forEach((p) => {
    const raw = String(p.lead || "").trim();
    if (!raw) return;
    const key = normalizeLeadKey(raw);
    if (!key) return;
    if (!groups.has(key)) {
      groups.set(key, { key, displayName: raw, projects: [] });
    }
    const group = groups.get(key);
    if (LEAD_TITLE_RE.test(raw) && !LEAD_TITLE_RE.test(group.displayName)) {
      group.displayName = raw;
    }
    group.projects.push(p);
  });
  return groups;
}

function formatMoney(raw) {
  const num = parseMoneyCell(raw);
  if (!Number.isFinite(num)) return raw || "0";
  return num.toLocaleString("en-US", {
    minimumFractionDigits: num % 1 !== 0 ? 2 : 0,
    maximumFractionDigits: 2,
  });
}

function fundingNumber(raw) {
  const num = parseMoneyCell(raw);
  return Number.isFinite(num) ? num : NaN;
}

function hasFundingValue(raw) {
  return Number.isFinite(fundingNumber(raw));
}

const MDY_RE = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/;

function maskDateInput(value) {
  const digits = value.replace(/\D/g, "").slice(0, 8);
  let out = digits.slice(0, 2);
  if (digits.length > 2) out += "/" + digits.slice(2, 4);
  if (digits.length > 4) out += "/" + digits.slice(4, 8);
  return out;
}

function maskHashPrefix(value) {
  const stripped = String(value || "").replace(/#/g, "");
  return stripped ? "#" + stripped : "";
}

function stripCommas(value) {
  return String(value || "").replace(/,/g, "");
}

// Parses a Financial Value cell read back from Google Sheets into a plain
// number. The Sheets API returns cells as *displayed* (currency-formatted)
// text by default — e.g. "$ 100,000.00" — so a bare parseFloat() stops at
// the "$" and yields NaN. Strip everything except digits/dot/minus first.
function parseMoneyCell(value) {
  const cleaned = String(value ?? "").replace(/[^0-9.-]/g, "");
  const num = parseFloat(cleaned);
  return Number.isFinite(num) ? num : NaN;
}

function maskFundingInput(value) {
  let cleaned = String(value || "").replace(/[^\d.]/g, "");
  const firstDot = cleaned.indexOf(".");
  if (firstDot !== -1) {
    cleaned = cleaned.slice(0, firstDot + 1) + cleaned.slice(firstDot + 1).replace(/\./g, "");
  }
  let [intPart, decPart] = cleaned.split(".");
  const withCommas = intPart ? Number(intPart).toLocaleString("en-US") : "";
  return decPart !== undefined ? withCommas + "." + decPart.slice(0, 2) : withCommas;
}

// Parses a "DD/MM/YYYY" (British-format) date string. Function name kept as
// parseMDY for historical reasons, but it now reads day first, month second.
function parseMDY(str) {
  if (!str) return null;
  const m = String(str).trim().match(MDY_RE);
  if (!m) return null;
  const day = parseInt(m[1], 10);
  const month = parseInt(m[2], 10);
  const year = parseInt(m[3], 10);
  if (month < 1 || month > 12) return null;
  const date = new Date(year, month - 1, day);
  if (date.getFullYear() !== year || date.getMonth() !== month - 1 || date.getDate() !== day) {
    return null;
  }
  return date;
}

function formatDateShort(value) {
  if (!value) return "—";
  const date = value instanceof Date ? value : new Date(value);
  if (Number.isNaN(date.getTime())) return "—";
  return date.toLocaleDateString("en-GB", { month: "short", day: "numeric", year: "numeric" });
}

// ---------------------------------------------------------------------------
// Phone number display formatting: just prefixes "+" (no grouping/spacing).
// ---------------------------------------------------------------------------
function formatPhoneDisplay(raw) {
  const trimmed = String(raw || "").trim();
  if (!trimmed) return "";
  const digits = trimmed.replace(/\D/g, "");
  if (!digits) return trimmed;
  return "+" + digits;
}

// ---------------------------------------------------------------------------
// Email display formatting: strips stray "<" / ">" characters.
// ---------------------------------------------------------------------------
function formatEmailDisplay(raw) {
  return String(raw || "").replace(/[<>]/g, "");
}

function daysBetween(startStr, endStr) {
  const start = parseMDY(startStr);
  const end = parseMDY(endStr);
  if (!start || !end) return null;
  const startMidnight = new Date(start.getFullYear(), start.getMonth(), start.getDate());
  const endMidnight = new Date(end.getFullYear(), end.getMonth(), end.getDate());
  return Math.round((endMidnight - startMidnight) / (1000 * 60 * 60 * 24));
}

function parseFlexibleDate(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) return value;
  if (typeof value === "number" && Number.isFinite(value)) {
    return new Date(Math.round((value - 25569) * 86400 * 1000));
  }
  if (typeof value === "string") {
    const trimmed = value.trim();
    if (!trimmed) return null;
    const mdy = parseMDY(trimmed);
    if (mdy) return mdy;
    const cleaned = trimmed.replace(/(\d+)(st|nd|rd|th)\b/gi, "$1");
    const d = new Date(cleaned);
    if (!Number.isNaN(d.getTime())) return d;
  }
  return null;
}

// Formats a Date as "DD/MM/YYYY" (British format). Function name kept as
// dateToMDYString for historical reasons, but day comes first now.
function dateToMDYString(date) {
  if (!(date instanceof Date) || Number.isNaN(date.getTime())) return "";
  const mm = String(date.getMonth() + 1).padStart(2, "0");
  const dd = String(date.getDate()).padStart(2, "0");
  const yyyy = date.getFullYear();
  return dd + "/" + mm + "/" + yyyy;
}

function normalizeLabel(label) {
  return String(label || "").trim().toLowerCase();
}

// Loose comparison key for matching a sheet tab title against a project's
// name/company (letters+digits only, lowercased) — tab titles for legacy
// tabs were derived from the company name but may differ in spacing,
// underscores, or punctuation (e.g. "Hyundai_OSP" tab vs "HyundaiOSP" text).
function normalizeCompareKey(value) {
  return String(value || "").toLowerCase().replace(/[^a-z0-9]/g, "");
}

function timeUntil(deadlineStr) {
  const date = parseMDY(deadlineStr);
  if (!date) return "—";
  const now = new Date();
  const nowMidnight = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  const deadlineMidnight = new Date(date.getFullYear(), date.getMonth(), date.getDate());
  const diffDays = Math.round((deadlineMidnight - nowMidnight) / (1000 * 60 * 60 * 24));
  if (diffDays > 0) return diffDays + (diffDays === 1 ? " day left" : " days left");
  if (diffDays === 0) return "Due today";
  const overdue = Math.abs(diffDays);
  return "Overdue by " + overdue + (overdue === 1 ? " day" : " days");
}

function formatDuration(startMs, endMs) {
  const diffMs = Math.max(0, endMs - startMs);
  const totalDays = Math.floor(diffMs / (1000 * 60 * 60 * 24));
  const months = Math.floor(totalDays / 30);
  const days = totalDays % 30;
  if (months === 0) return days + (days === 1 ? " day" : " days");
  if (days === 0) return months + (months === 1 ? " month" : " months");
  return (
    months + (months === 1 ? " month, " : " months, ") + days + (days === 1 ? " day" : " days")
  );
}

const APP_NAME = "SRL Project Pipelines";
const LOGO_SRC = "data:image/png;base64,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";

function BrandLogo({ size = 22 }) {
  return (
    <img
      src={LOGO_SRC}
      alt="SRL Project Pipelines"
      width={size}
      height={size}
      style={{ display: "block", objectFit: "contain" }}
    />
  );
}

function Field({ label, id, error, children, full }) {
  return (
    <div className={"field" + (full ? " field-full" : "")}>
      <label className="label" htmlFor={id}>{label}</label>
      {children}
      {error ? <div className="error-text">{error}</div> : null}
    </div>
  );
}

// Potential is optional, so this renders "—" instead of an empty chip when
// a project has no potentialLevel set.
function PotentialChip({ level, large }) {
  if (!level) return <span>—</span>;
  return (
    <span className={"chip " + (large ? "chip-lg " : "") + "chip-" + level.toLowerCase()}>
      {level}
    </span>
  );
}

// Custom pie-chart tooltip: funding amount first, project name second (in
// that order), and rendered smaller than recharts' default tooltip text.
function PieFundingTooltip({ active, payload }) {
  if (!active || !payload || !payload.length) return null;
  const d = payload[0];
  return (
    <div className="pie-tooltip">
      <div className="pie-tooltip-funding">${formatMoney(d.value)}</div>
      <div className="pie-tooltip-project">{d.name}</div>
    </div>
  );
}

function TextInput({ id, icon, prefix, error, ...props }) {
  const Icon = icon;
  return (
    <div
      className={
        "input-wrap" +
        (error ? " input-wrap-error" : "") +
        (prefix ? " input-wrap-prefixed" : "")
      }
    >
      {Icon ? <Icon size={16} className="input-icon" strokeWidth={1.75} /> : null}
      {prefix ? <span className="input-prefix">{prefix}</span> : null}
      <input id={id} className="input" {...props} />
    </div>
  );
}

// ---------------------------------------------------------------------------
// Contact name field with a suggestions dropdown pulled from the existing
// contacts list. Typing still works like a normal text field — picking a
// suggestion just auto-fills the other contact fields (title/email/number)
// so the same person doesn't get re-entered as a duplicate.
// ---------------------------------------------------------------------------
function ContactNameCombo({ id, value, error, contacts, onTextChange, onSelectContact }) {
  const [open, setOpen] = useState(false);
  const wrapRef = useRef(null);

  useEffect(() => {
    function handleClickOutside(e) {
      if (wrapRef.current && !wrapRef.current.contains(e.target)) setOpen(false);
    }
    document.addEventListener("mousedown", handleClickOutside);
    return () => document.removeEventListener("mousedown", handleClickOutside);
  }, []);

  const query = String(value || "").trim().toLowerCase();
  const matches = query
    ? contacts
        .filter((c) => !c.archived && (c.contactName || "").toLowerCase().includes(query))
        .slice(0, 8)
    : [];

  return (
    <div className="contact-combo" ref={wrapRef}>
      <TextInput
        id={id}
        error={error}
        value={value}
        autoComplete="off"
        onChange={(e) => {
          onTextChange(e.target.value);
          setOpen(true);
        }}
        onFocus={() => setOpen(true)}
      />
      {open && matches.length > 0 && (
        <div className="contact-combo-dropdown">
          {matches.map((c) => (
            <button
              type="button"
              key={c.id}
              className="contact-combo-option"
              onMouseDown={(e) => {
                e.preventDefault();
                onSelectContact(c);
                setOpen(false);
              }}
            >
              <span className="contact-combo-option-name">{c.contactName || "Unnamed contact"}</span>
              <span className="contact-combo-option-sub">
                {[c.jobTitle, c.contactEmail].filter(Boolean).join(" · ") || "No additional details on file"}
              </span>
            </button>
          ))}
        </div>
      )}
    </div>
  );
}

function CompanyCombo({ id, value, error, companies, onTextChange, onSelectCompany }) {
  const [open, setOpen] = useState(false);
  const wrapRef = useRef(null);

  useEffect(() => {
    function handleClickOutside(e) {
      if (wrapRef.current && !wrapRef.current.contains(e.target)) setOpen(false);
    }
    document.addEventListener("mousedown", handleClickOutside);
    return () => document.removeEventListener("mousedown", handleClickOutside);
  }, []);

  const query = String(value || "").trim().toLowerCase();
  const activeCompanies = companies.filter((c) => !c.archived);
  const matches = query
    ? activeCompanies.filter((c) => (c.name || "").toLowerCase().includes(query)).slice(0, 8)
    : activeCompanies.slice(0, 8);

  return (
    <div className="contact-combo" ref={wrapRef}>
      <TextInput
        id={id}
        error={error}
        value={value}
        autoComplete="off"
        placeholder={activeCompanies.length ? "Start typing a company name" : "No companies yet"}
        onChange={(e) => {
          onTextChange(e.target.value);
          setOpen(true);
        }}
        onFocus={() => setOpen(true)}
      />
      {open && matches.length > 0 && (
        <div className="contact-combo-dropdown">
          {matches.map((c) => (
            <button
              type="button"
              key={c.id}
              className="contact-combo-option"
              onMouseDown={(e) => {
                e.preventDefault();
                onSelectCompany(c);
                setOpen(false);
              }}
            >
              <span className="contact-combo-option-name">{c.name}</span>
              <span className="contact-combo-option-sub">{c.industry || "No industry on file"}</span>
            </button>
          ))}
        </div>
      )}
    </div>
  );
}

function SkeletonCards({ count = 6 }) {
  return (
    <div className="skeleton-grid" aria-busy="true" aria-label="Loading projects">
      {Array.from({ length: count }).map((_, i) => (
        <div className="skeleton-card" key={i}>
          <div className="skeleton-line skeleton-line-lg" />
          <div className="skeleton-line skeleton-line-md" />
          <div className="skeleton-line skeleton-line-sm" />
        </div>
      ))}
    </div>
  );
}

function SelectInput({ id, error, children, ...props }) {
  return (
    <select id={id} className={"select" + (error ? " select-error" : "")} {...props}>
      {children}
    </select>
  );
}

function Textarea({ id, error, ...props }) {
  return <textarea id={id} className={"textarea" + (error ? " select-error" : "")} {...props} />;
}

function SrlNoteCell({ level, note, onSave }) {
  const [text, setText] = useState(note ? note.text : "");
  return (
    <div className="srl-note-cell">
      <textarea
        className="srl-note-input"
        rows={4}
        placeholder="Add a note…"
        value={text}
        onChange={(e) => setText(e.target.value)}
        onBlur={() => onSave(level, text)}
      />
      {note && note.text ? (
        <p className="srl-note-meta">
          — {note.author}
          {note.updatedAt ? ", " + formatDateShort(note.updatedAt) : ""}
        </p>
      ) : null}
    </div>
  );
}

function AuthedLayout({ currentUser, onSignOut, storageError, children }) {
  return (
    <div className="dashboard">
      <header className="dashboard-header">
        <div className="brand">
          <BrandLogo />
          <span className="brand-name">{APP_NAME}</span>
        </div>
        <div className="dashboard-user">
          <span className="dashboard-user-name">{currentUser.username}</span>
          <span className="badge-mono">{currentUser.email}</span>
        </div>
      </header>

      <main className="dashboard-main">
        {storageError ? <div className="form-error" style={{ marginBottom: 16 }}>{storageError}</div> : null}
        {children}
      </main>

      <footer className="dashboard-footer">
        <button type="button" className="btn-secondary" onClick={onSignOut}>
          <LogOut size={16} strokeWidth={1.75} />
          Sign out
        </button>
      </footer>
    </div>
  );
}

export default function App() {
  // Inject fonts once.
  useEffect(() => {
    const link = document.createElement("link");
    link.rel = "stylesheet";
    link.href =
      "https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=JetBrains+Mono:wght@500&display=swap";
    document.head.appendChild(link);
    return () => document.head.removeChild(link);
  }, []);

  // Ensure a mobile-friendly viewport is set (some hosts omit this by default).
  useEffect(() => {
    let meta = document.querySelector('meta[name="viewport"]');
    const hadMeta = !!meta;
    const previousContent = meta ? meta.getAttribute("content") : null;
    if (!meta) {
      meta = document.createElement("meta");
      meta.name = "viewport";
      document.head.appendChild(meta);
    }
    meta.setAttribute("content", "width=device-width, initial-scale=1, maximum-scale=1");
    return () => {
      if (!hadMeta && meta.parentNode) {
        meta.parentNode.removeChild(meta);
      } else if (hadMeta && previousContent != null) {
        meta.setAttribute("content", previousContent);
      }
    };
  }, []);

  // "Database" — Firebase Auth (accounts) + a small Firestore "usernames"
  // collection that maps each username to its real email under the hood.
  const [currentUser, setCurrentUser] = useState(null);
  const [authChecked, setAuthChecked] = useState(false);
  const [pendingReset, setPendingReset] = useState(null); // { email, token } — see handleForgot
  const [projects, setProjects] = useState([]);
  // True until the first Google Sheets fetch settles, so the dashboard and
  // project directory can show skeleton cards instead of an empty page.
  const [sheetLoading, setSheetLoading] = useState(true);
  const [selectedProjectId, setSelectedProjectId] = useState(null);
  const [projectReturnView, setProjectReturnView] = useState("projects");
  const [companies, setCompanies] = useState([]);
  const [selectedCompanyId, setSelectedCompanyId] = useState(null);
  const [companyReturnView, setCompanyReturnView] = useState("companies");
  const [contacts, setContacts] = useState([]);

  // Keep synchronous references to the latest sheet-backed state. The old
  // implementation nested setProjects inside setContacts inside setCompanies;
  // React can defer those nested updater callbacks, which made projects,
  // companies and contacts appear minutes apart. A sheet poll now reconciles
  // all three collections in one pass and commits them together.
  const projectsRef = useRef([]);
  const companiesRef = useRef([]);
  const contactsRef = useRef([]);
  useEffect(() => { projectsRef.current = projects; }, [projects]);
  useEffect(() => { companiesRef.current = companies; }, [companies]);
  useEffect(() => { contactsRef.current = contacts; }, [contacts]);

  const [selectedContactId, setSelectedContactId] = useState(null);
  const [contactReturnView, setContactReturnView] = useState("contacts");

  const [storageError, setStorageError] = useState("");

  // Shared Firestore app-data document. Firebase Authentication already exists
  // in this project; appData/main gives projects, companies and contacts a
  // durable home that survives refresh independently of Google Sheets.
  const [appDataReady, setAppDataReady] = useState(false);
  const appDataLoadedForUid = useRef("");
  const appDataSaveTimer = useRef(null);
  // Sheet reconciliation is a READ operation. Never let a sheet-driven React
  // update trigger the generic Firestore autosave and overwrite a user edit.
  const suppressNextAppDataAutosave = useRef(false);

  // Track sign-in state. currentUser here mirrors the Firebase Auth user,
  // with `username` read from the display name we set at signup.
  useEffect(() => {
    const unsub = onAuthStateChanged(auth, (fbUser) => {
      if (fbUser) {
        setCurrentUser({
          uid: fbUser.uid,
          username: fbUser.displayName || fbUser.email.split("@")[0],
          email: fbUser.email,
        });
      } else {
        setCurrentUser(null);
      }
      setAuthChecked(true);
    });
    return () => unsub();
  }, []);

  // Load Firestore first on every authenticated browser session. This makes a
  // refresh fast and, more importantly, prevents website-created projects from
  // disappearing just because Google Sheets is delayed or temporarily rejects
  // a write. Google Sheets still reconciles in the background afterwards.
  useEffect(() => {
    let cancelled = false;
    const uid = currentUser?.uid || "";

    if (!authChecked || !uid) {
      appDataLoadedForUid.current = "";
      setAppDataReady(false);
      return () => { cancelled = true; };
    }
    if (appDataLoadedForUid.current === uid) return () => { cancelled = true; };

    setAppDataReady(false);
    (async () => {
      const result = await loadAppData();
      if (cancelled) return;

      if (result.ok && result.exists) {
        const loadedProjects = Array.isArray(result.projects) ? result.projects : [];
        const loadedCompanies = Array.isArray(result.companies) ? result.companies : [];
        const loadedContacts = Array.isArray(result.contacts) ? result.contacts : [];
        projectsRef.current = loadedProjects;
        companiesRef.current = loadedCompanies;
        contactsRef.current = loadedContacts;
        setProjects(loadedProjects);
        setCompanies(loadedCompanies);
        setContacts(loadedContacts);
      } else if (!result.ok) {
        console.error("Firestore app-data load failed:", result.error);
      }

      appDataLoadedForUid.current = uid;
      setAppDataReady(true);
    })();

    return () => { cancelled = true; };
  }, [authChecked, currentUser?.uid]);

  // Persist every app-state change to Firestore with a short debounce. This
  // covers project/company/contact adds, edits, imports, archive actions, SRL
  // notes, and other UI changes without needing a separate save call in every
  // handler. The initial empty React state is never written before hydration.
  useEffect(() => {
    if (!appDataReady || !currentUser?.uid) return;
    if (suppressNextAppDataAutosave.current) {
      suppressNextAppDataAutosave.current = false;
      return;
    }
    if (appDataSaveTimer.current) clearTimeout(appDataSaveTimer.current);

    appDataSaveTimer.current = setTimeout(async () => {
      const result = await saveAppData({
        projects: projectsRef.current,
        companies: companiesRef.current,
        contacts: contactsRef.current,
        updatedBy: currentUser.username || currentUser.email || currentUser.uid,
      });
      if (!result.ok) console.error("Firestore app-data save failed:", result.error);
    }, 50);

    return () => {
      if (appDataSaveTimer.current) clearTimeout(appDataSaveTimer.current);
    };
  }, [projects, companies, contacts, appDataReady, currentUser?.uid]);

  const [view, setView] = useState("login"); // login | signup | forgot | forgotSent | reset | dashboard | projects | addProject | viewProject | companies | addCompany | contacts | viewContact
  const [banner, setBanner] = useState(""); // one-line success message shown on login screen

  // ---- Login form -----------------------------------------------------
  const [loginUsername, setLoginUsername] = useState("");
  const [loginPassword, setLoginPassword] = useState("");
  const [loginErrors, setLoginErrors] = useState({});

  function handleLogin(e) {
    e.preventDefault();
    const errors = {};
    if (!loginUsername.trim()) errors.username = "Enter your username.";
    if (!loginPassword) errors.password = "Enter your password.";
    if (Object.keys(errors).length) return setLoginErrors(errors);

    (async () => {
      try {
        const unameKey = loginUsername.trim().toLowerCase();
        const lookup = await getDoc(doc(db, "usernames", unameKey));
        if (!lookup.exists()) {
          setLoginErrors({ form: "Username or password is incorrect." });
          return;
        }
        const { email } = lookup.data();
        await signInWithEmailAndPassword(auth, email, loginPassword.trim());
        setLoginErrors({});
        setBanner("");
        setLoginUsername("");
        setLoginPassword("");
        setView("dashboard");
      } catch (err) {
        setLoginErrors({ form: "Username or password is incorrect." });
      }
    })();
  }

  // ---- Sign up form -----------------------------------------------------
  const [suUsername, setSuUsername] = useState("");
  const [suEmail, setSuEmail] = useState("");
  const [suPassword, setSuPassword] = useState("");
  const [suConfirm, setSuConfirm] = useState("");
  const [suErrors, setSuErrors] = useState({});

  function handleSignup(e) {
    e.preventDefault();
    const errors = {};
    const uname = suUsername.trim();
    const email = suEmail.trim();

    if (uname.length < 3) errors.username = "Username must be at least 3 characters.";
    if (!EMAIL_RE.test(email)) errors.email = "Enter a valid email address.";
    if (suPassword.length < 6) errors.password = "Password must be at least 6 characters.";
    if (suConfirm !== suPassword) errors.confirm = "Passwords don't match.";

    if (Object.keys(errors).length) return setSuErrors(errors);

    (async () => {
      try {
        const unameKey = uname.toLowerCase();
        const existing = await getDoc(doc(db, "usernames", unameKey));
        if (existing.exists()) {
          setSuErrors({ username: "That username is already taken." });
          return;
        }

        const cred = await createUserWithEmailAndPassword(auth, email, suPassword);
        await updateProfile(cred.user, { displayName: uname });
        await setDoc(doc(db, "usernames", unameKey), { email, uid: cred.user.uid });

        setSuErrors({});
        setSuUsername("");
        setSuEmail("");
        setSuPassword("");
        setSuConfirm("");
        setBanner("Account created. Sign in with your new username and password.");
        await signOut(auth); // don't auto-login; send them to the sign-in screen
        setView("login");
      } catch (err) {
        if (err.code === "auth/email-already-in-use") {
          setSuErrors({ email: "An account with that email already exists." });
        } else {
          setSuErrors({ form: "Something went wrong creating your account. Try again." });
        }
      }
    })();
  }

  // ---- Forgot password ---------------------------------------------------
  const [fpEmail, setFpEmail] = useState("");
  const [fpError, setFpError] = useState("");

  function handleForgot(e) {
    e.preventDefault();
    const email = fpEmail.trim();
    if (!EMAIL_RE.test(email)) {
      setFpError("Enter a valid email address.");
      return;
    }
    setFpError("");
    (async () => {
      try {
        await sendPasswordResetEmail(auth, email);
      } catch (err) {
        // Intentionally don't reveal whether the email exists — same screen either way.
      }
      setFpEmail("");
      setView("forgotSent");
    })();
  }

  // ---- Reset password -----------------------------------------------------
  const [rpPassword, setRpPassword] = useState("");
  const [rpConfirm, setRpConfirm] = useState("");
  const [rpErrors, setRpErrors] = useState({});

  // With real Firebase Auth, password resets happen through the emailed link
  // (Firebase's own hosted reset page), not this in-app screen — this stays
  // only so the "reset" route doesn't dangle. It's unreachable in normal use
  // since pendingReset is never set by handleForgot above.
  function handleReset(e) {
    e.preventDefault();
    setRpErrors({});
    setRpPassword("");
    setRpConfirm("");
    setPendingReset(null);
    setView("login");
  }

  function handleSignOut() {
    signOut(auth);
    setBanner("");
    setView("login");
  }

  function goTo(next) {
    setLoginErrors({});
    setSuErrors({});
    setFpError("");
    setRpErrors({});
    setBanner("");
    setView(next);
  }

  // ---- Archive / restore / permanent delete -----------------------------
  function archiveProject(id) {
    const now = Date.now();
    setProjects((prev) => prev.map((p) => (p.id === id ? { ...p, archived: true, updatedAt: now } : p)));
  }
  function restoreProject(id) {
    const now = Date.now();
    setProjects((prev) => prev.map((p) => (p.id === id ? { ...p, archived: false, updatedAt: now } : p)));
  }
  function deleteProjectForever(id) {
    const target = projectsRef.current.find((p) => p.id === id);
    if (target) removeProjectFromSheets(target);
    const next = projectsRef.current.filter((p) => p.id !== id);
    projectsRef.current = next;
    setProjects(next);
    void deleteAppDataRecords({
      projectIds: [id],
      updatedBy: currentUser?.username || currentUser?.email || currentUser?.uid || "",
    }).then((result) => {
      if (!result.ok) setStorageError(`Firebase delete failed: ${result.error}`);
    });
  }

  function archiveCompany(id) {
    const now = Date.now();
    setCompanies((prev) => prev.map((c) => (c.id === id ? { ...c, archived: true, updatedAt: now } : c)));
  }
  function restoreCompany(id) {
    const now = Date.now();
    setCompanies((prev) => prev.map((c) => (c.id === id ? { ...c, archived: false, updatedAt: now } : c)));
  }
  function deleteCompanyForever(id) {
    const next = companiesRef.current.filter((c) => c.id !== id);
    companiesRef.current = next;
    setCompanies(next);
    void deleteAppDataRecords({
      companyIds: [id],
      updatedBy: currentUser?.username || currentUser?.email || currentUser?.uid || "",
    }).then((result) => {
      if (!result.ok) setStorageError(`Firebase delete failed: ${result.error}`);
    });
  }

  function archiveContact(id) {
    const now = Date.now();
    setContacts((prev) => prev.map((c) => (c.id === id ? { ...c, archived: true, updatedAt: now } : c)));
  }
  function restoreContact(id) {
    const now = Date.now();
    setContacts((prev) => prev.map((c) => (c.id === id ? { ...c, archived: false, updatedAt: now } : c)));
  }
  function deleteContactForever(id) {
    const next = contactsRef.current.filter((c) => c.id !== id);
    contactsRef.current = next;
    setContacts(next);
    void deleteAppDataRecords({
      contactIds: [id],
      updatedBy: currentUser?.username || currentUser?.email || currentUser?.uid || "",
    }).then((result) => {
      if (!result.ok) setStorageError(`Firebase delete failed: ${result.error}`);
    });
  }

  // Permanently clear a whole archive in one go.
  function deleteAllArchivedProjects() {
    const ids = projectsRef.current.filter((p) => p.archived).map((p) => p.id);
    projectsRef.current.filter((p) => p.archived).forEach((p) => removeProjectFromSheets(p));
    const next = projectsRef.current.filter((p) => !p.archived);
    projectsRef.current = next;
    setProjects(next);
    if (ids.length) void deleteAppDataRecords({
      projectIds: ids,
      updatedBy: currentUser?.username || currentUser?.email || currentUser?.uid || "",
    });
  }
  function deleteAllArchivedCompanies() {
    const ids = companiesRef.current.filter((c) => c.archived).map((c) => c.id);
    const next = companiesRef.current.filter((c) => !c.archived);
    companiesRef.current = next;
    setCompanies(next);
    if (ids.length) void deleteAppDataRecords({
      companyIds: ids,
      updatedBy: currentUser?.username || currentUser?.email || currentUser?.uid || "",
    });
  }
  function deleteAllArchivedContacts() {
    const ids = contactsRef.current.filter((c) => c.archived).map((c) => c.id);
    const next = contactsRef.current.filter((c) => !c.archived);
    contactsRef.current = next;
    setContacts(next);
    if (ids.length) void deleteAppDataRecords({
      contactIds: ids,
      updatedBy: currentUser?.username || currentUser?.email || currentUser?.uid || "",
    });
  }

  // Contact rows typed into a project are matched against the existing
  // contact list (email first, then name) so an already-known person is
  // linked instead of duplicated — and crucially keeps their own company.
  function findContactMatch(row, pool) {
    const email = String(row.email || "").trim().toLowerCase();
    const name = String(row.name || "").trim().toLowerCase();
    if (email) {
      const byEmail = pool.find((c) => String(c.contactEmail || "").trim().toLowerCase() === email);
      if (byEmail) return byEmail;
    }
    if (name) {
      return pool.find((c) => String(c.contactName || "").trim().toLowerCase() === name) || null;
    }
    return null;
  }
  function applyContactRow(contact, row, at) {
    return {
      ...contact,
      contactName: row.name.trim() || contact.contactName,
      jobTitle: row.title.trim() || contact.jobTitle,
      contactEmail: row.email.trim() || contact.contactEmail,
      contactNumber: row.number.trim() || contact.contactNumber,
      updatedAt: at,
    };
  }

  // Click-to-confirm state for permanent deletes (avoids relying on window.confirm,
  // which some embedded/sandboxed environments block silently).
  const [confirmForeverId, setConfirmForeverId] = useState(null);

  // ---- Add project form -----------------------------------------------------
  const SRL_LEVELS = ["1", "2", "3", "4", "5", "6", "7"];
  const SRL_KEY = [
    { lvl: "1", name: "Prospecting", desc: "Sourcing potential leads" },
    { lvl: "2", name: "Lead Qualification", desc: "Assessing lead potential" },
    { lvl: "3", name: "Needs Assessment", desc: "Gathering prospect requirements" },
    { lvl: "4", name: "Proposal", desc: "Presenting services, pricing & terms" },
    { lvl: "5", name: "Evaluate", desc: "Prospect reviewing the proposal" },
    { lvl: "6", name: "Negotiate", desc: "Finalizing terms with the prospect" },
    { lvl: "7", name: "Sign", desc: "Contract signed & confirmed" },
  ];
  const POTENTIAL_LEVELS = ["Low", "Medium", "High"];

  // ---- Live Google Sheet sync -----------------------------------------------
  // The one-page summary is the authoritative PROJECT INVENTORY. If it has 31
  // project rows, the website has 31 projects. Detail tabs enrich those rows
  // with the real project title, full PI, contacts, comments, start/end dates,
  // and any SRL dates that are missing from the summary.
  //
  // This deliberately mirrors the Excel import behavior while avoiding the two
  // regressions that previously reduced 31 projects to 26:
  //   1) company name is never a project ID;
  //   2) detail tabs never replace/collapse summary rows.
  function contactKey(c, companyId = "") {
    const email = String(c?.email || c?.contactEmail || "").trim().toLowerCase();
    if (email && email.includes("@")) return "e:" + email;
    const name = normalizePersonKey(c?.name || c?.contactName || "");
    return name ? "n:" + normalizeCompareKey(companyId) + ":" + name : "";
  }

  function validSummarySrl(raw) {
    const match = String(raw || "").match(/(?:^|\D)([1-7])(?:\D|$)/);
    return match ? match[1] : "";
  }

  function detailDerivedSrl(detail) {
    let level = "";
    (detail?.srlDates || []).forEach((raw, index) => {
      if (parseFlexibleDate(raw)) level = String(index + 1);
    });
    return level;
  }

  function sameDateValue(a, b) {
    const da = parseFlexibleDate(a);
    const db = parseFlexibleDate(b);
    return Boolean(da && db && da.getTime() === db.getTime());
  }

  // Pair detail tabs to summary rows one-to-one. Exact formula/tab references
  // win. When an old summary row has no formula reference, use several fields
  // together (company + PI/RFS/funding/SRL dates). A detail tab can be consumed
  // only once, so two SABIC/Aramco/Maaden/Saudi Diesel rows can never collapse
  // into one project.
  function matchDetailsToSummaryRows(summaryRows, detailSnapshots) {
    const details = Array.isArray(detailSnapshots) ? detailSnapshots : [];
    const matches = Array.from({ length: summaryRows.length }, () => null);
    const used = new Set();

    // 1. Exact summary formula -> detail-tab name.
    summaryRows.forEach((row, rowIndex) => {
      const tabKey = normalizeCompareKey(row.sheetTabName);
      if (!tabKey) return;
      const detailIndex = details.findIndex(
        (detail, index) => !used.has(index) && normalizeCompareKey(detail.sheetTabName) === tabKey,
      );
      if (detailIndex !== -1) {
        matches[rowIndex] = details[detailIndex];
        used.add(detailIndex);
      }
    });

    // 2. Exact Project ID when a newer detail sheet has one.
    summaryRows.forEach((row, rowIndex) => {
      if (matches[rowIndex]) return;
      const id = String(row.projectId || "").trim();
      if (!id) return;
      const detailIndex = details.findIndex(
        (detail, index) => !used.has(index) && String(detail.projectId || "").trim() === id,
      );
      if (detailIndex !== -1) {
        matches[rowIndex] = details[detailIndex];
        used.add(detailIndex);
      }
    });

    // 3. Tolerant one-to-one matching for untouched legacy rows.
    summaryRows.forEach((row, rowIndex) => {
      if (matches[rowIndex]) return;
      const companyKey = normalizeCompareKey(row.company);
      const piKey = normalizeLeadKey(row.pi);
      const rfsKey = normalizeCompareKey(String(row.rfsNti || "").replace(/^#/, ""));
      const summaryFunding = parseMoneyCell(row.financialValue);
      const summarySrl = validSummarySrl(row.currentSrl);

      let bestIndex = -1;
      let bestScore = -Infinity;
      details.forEach((detail, detailIndex) => {
        if (used.has(detailIndex)) return;
        const detailCompanyKey = normalizeCompareKey(detail.companyName);
        const detailTabKey = normalizeCompareKey(detail.sheetTabName);
        const detailPiKey = normalizeLeadKey(detail.leadPi);
        const detailRfsKey = normalizeCompareKey(String(detail.rfsNumber || "").replace(/^#/, ""));
        const detailFunding = parseMoneyCell(detail.financialValue);
        const detailSrl = detailDerivedSrl(detail);

        let score = 0;
        if (companyKey && detailCompanyKey === companyKey) score += 1200;
        // Legacy tab names often include a surname suffix: AramcoBassam,
        // MaadenBill1, SaudiDieselShehab, etc. Treat that only as supporting
        // evidence, never as company identity.
        if (companyKey && detailTabKey && (detailTabKey.startsWith(companyKey) || companyKey.startsWith(detailTabKey))) {
          score += 250;
        }
        if (piKey && detailPiKey && piKey === detailPiKey) score += 900;
        if (rfsKey && detailRfsKey && rfsKey === detailRfsKey) score += 1300;
        if (Number.isFinite(summaryFunding) && Number.isFinite(detailFunding)) {
          if (Math.abs(summaryFunding - detailFunding) < 0.01) score += 1000;
          else score -= 250;
        }
        if (summarySrl && detailSrl && summarySrl === detailSrl) score += 250;
        (row.srlDates || []).forEach((raw, dateIndex) => {
          if (raw && sameDateValue(raw, detail.srlDates?.[dateIndex])) score += 120;
        });

        if (score > bestScore) {
          bestScore = score;
          bestIndex = detailIndex;
        }
      });

      // A company match alone is sufficient only because this assignment is
      // one-to-one and the SUMMARY ROW remains the project identity. Stronger
      // fields above disambiguate companies with multiple projects.
      if (bestIndex !== -1 && bestScore >= 1000) {
        matches[rowIndex] = details[bestIndex];
        used.add(bestIndex);
      }
    });

    return matches;
  }

  function applyCanonicalSheetSnapshot(summaryRows, detailSnapshots = null) {
    if (!Array.isArray(summaryRows)) return;
    const hasDetails = Array.isArray(detailSnapshots) && detailSnapshots.length > 0;
    if (summaryRows.length === 0 && !hasDetails) return;
    const now = Date.now();
    const detailsAvailable = Array.isArray(detailSnapshots) && detailSnapshots.length > 0;
    let inventoryRows = [...summaryRows];
    let detailMatches = detailsAvailable
      ? matchDetailsToSummaryRows(inventoryRows, detailSnapshots)
      : Array.from({ length: inventoryRows.length }, () => null);

    // A website-created detail tab carries a stable Project ID in its metadata.
    // If the detail tab exists but its summary append was delayed/interrupted,
    // keep it in the website inventory instead of deleting it on refresh. This
    // fallback is intentionally limited to tabs with an explicit Project ID so
    // legacy company tabs can never be double-counted.
    if (detailsAvailable) {
      const matchedTabs = new Set(
        detailMatches
          .filter(Boolean)
          .map((detail) => normalizeCompareKey(detail.sheetTabName))
          .filter(Boolean),
      );
      const representedIds = new Set(
        inventoryRows.map((row, index) => String(row.projectId || index + 1).trim()).filter(Boolean),
      );
      detailSnapshots.forEach((detail) => {
        const tabKey = normalizeCompareKey(detail.sheetTabName);
        const explicitId = String(detail.projectId || "").trim();
        if (!explicitId || matchedTabs.has(tabKey) || representedIds.has(explicitId)) return;
        inventoryRows.push({
          projectId: explicitId,
          sheetTabName: detail.sheetTabName || "",
          company: detail.companyName || "",
          pi: detail.leadPi || "",
          rfsNti: detail.rfsNumber || "",
          currentSrl: detailDerivedSrl(detail),
          srlDates: Array.from({ length: 7 }, (_, i) => detail.srlDates?.[i] || ""),
          financialValue: detail.financialValue || "",
        });
        detailMatches.push(detail);
        matchedTabs.add(tabKey);
        representedIds.add(explicitId);
      });
    }

    const previousById = new Map(
      projectsRef.current.map((project) => [String(project.sheetProjectId || project.id || "").trim(), project]),
    );
    const previousByTab = new Map(
      projectsRef.current
        .filter((project) => String(project.sheetTabName || "").trim())
        .map((project) => [normalizeCompareKey(project.sheetTabName), project]),
    );

    // Preserve user-created/non-sheet company records, but rebuild all
    // sheet-derived companies from the current 31-row inventory so stale
    // aliases such as "AramcoBassam" cannot survive as separate companies.
    const nextCompanies = companiesRef.current
      .filter((company) => !String(company.id || "").startsWith("sheet-company:"))
      .map((company) => ({ ...company }));
    const companyByName = new Map(
      nextCompanies
        .map((company) => [normalizeCompareKey(company.name), company.id])
        .filter(([key]) => Boolean(key)),
    );

    function ensureCompany(rawName) {
      const name = String(rawName || "").trim();
      if (!name) return "";
      const key = normalizeCompareKey(name);
      if (companyByName.has(key)) return companyByName.get(key);
      let id = "sheet-company:" + (key || String(nextCompanies.length + 1));
      if (nextCompanies.some((company) => company.id === id)) id += ":" + nextCompanies.length;
      companyByName.set(key, id);
      nextCompanies.push({
        id,
        name,
        about: "",
        industry: "",
        hqCountry: "",
        hqLocation: "",
        sbuLocation: "",
        updates: [],
        archived: false,
      });
      return id;
    }

    // Preserve manual contacts, rebuild sheet contacts from the project tabs.
    let nextContacts = contactsRef.current
      .filter((contact) => !String(contact.id || "").startsWith("sheet-contact:"))
      .map((contact) => ({ ...contact }));
    const contactByKey = new Map();
    nextContacts.forEach((contact) => {
      const key = contactKey(contact, contact.companyId);
      if (key && !contactByKey.has(key)) contactByKey.set(key, contact.id);
    });

    function ensureContact(person, companyId, projectId, contactIndex) {
      const incoming = {
        name: String(person?.name || "").trim(),
        jobTitle: String(person?.jobTitle || "").trim(),
        email: String(person?.email || "").trim(),
        phone: String(person?.phone || "").trim(),
      };
      if (!incoming.name && !incoming.email && !incoming.phone) return "";

      const key = contactKey(incoming, companyId);
      let id = key ? contactByKey.get(key) : "";
      let index = id ? nextContacts.findIndex((contact) => contact.id === id) : -1;

      if (index === -1) {
        const stable = normalizeCompareKey(incoming.email || normalizePersonKey(incoming.name) || `${projectId}-${contactIndex}`);
        id = "sheet-contact:" + (stable || `${projectId}-${contactIndex}`);
        while (nextContacts.some((contact) => contact.id === id)) id += "x";
        nextContacts.push({
          id,
          companyId,
          contactName: incoming.name,
          jobTitle: incoming.jobTitle,
          contactEmail: incoming.email,
          contactNumber: incoming.phone,
          createdAt: now,
          updatedAt: now,
          archived: false,
        });
        index = nextContacts.length - 1;
      } else {
        const prior = nextContacts[index];
        const preferredName =
          incoming.name && (LEAD_TITLE_RE.test(incoming.name) || !prior.contactName)
            ? incoming.name
            : prior.contactName;
        nextContacts[index] = {
          ...prior,
          companyId: companyId || prior.companyId || "",
          contactName: preferredName || incoming.name,
          jobTitle: incoming.jobTitle || prior.jobTitle || "",
          contactEmail: incoming.email || prior.contactEmail || "",
          contactNumber: incoming.phone || prior.contactNumber || "",
          updatedAt: now,
          archived: false,
        };
      }

      const finalContact = nextContacts[index];
      const finalKey = contactKey(finalContact, finalContact.companyId);
      if (finalKey) contactByKey.set(finalKey, finalContact.id);
      return finalContact.id;
    }

    const finalProjects = inventoryRows.map((row, rowIndex) => {
      const detail = detailMatches[rowIndex];
      // Website-created tabs carry a stable Project ID in their metadata.
      // During the fast summary-only phase, the detail metadata is not loaded
      // yet, but the summary formulas still expose the exact tab name. Reuse
      // the existing project by tab before falling back to a legacy row number.
      // This prevents the summary-first poll from temporarily changing an app
      // ID into "32" and then autosaving that rollback to Firestore.
      const explicitProjectId = String(detail?.projectId || row.projectId || "").trim();
      const priorByTab = row.sheetTabName
        ? previousByTab.get(normalizeCompareKey(row.sheetTabName)) || null
        : null;
      const prior = (explicitProjectId ? previousById.get(explicitProjectId) : null) || priorByTab || null;
      const projectId = String(
        explicitProjectId || prior?.sheetProjectId || prior?.id || rowIndex + 1
      ).trim();

      // The fast summary request is inventory-only. It must never overwrite an
      // already-loaded project's editable values while the slower detail tabs
      // are still loading/recalculating. This removes the visible "2 -> 1"
      // rollback that occurred immediately after an edit.
      if (!detailsAvailable && prior) {
        return {
          ...prior,
          id: projectId,
          sheetProjectId: prior.sheetProjectId || projectId,
          sheetTabName: String(row.sheetTabName || prior.sheetTabName || "").trim(),
        };
      }

      // A website edit/addition is authoritative until Google Sheets confirms
      // the write. Without this guard, the 60-second sheet poll can read the
      // old spreadsheet value, overwrite the fresh React/Firestore value, and
      // then the Firestore autosave makes that rollback permanent.
      if (prior?.sheetSyncPending === true) {
        return {
          ...prior,
          id: projectId,
          sheetProjectId: prior.sheetProjectId || projectId,
          sheetTabName: String(detail?.sheetTabName || row.sheetTabName || prior.sheetTabName || "").trim(),
        };
      }

      // The individual sheet's `Company Name` is authoritative for company
      // identity. The one-page summary in older copies may contain tab-style
      // aliases such as AramcoBassam/AramcoXuLu; those must both resolve to the
      // detail sheet's canonical `Aramco`. Use the summary only as a fallback
      // while detail data is unavailable.
      const companyNameValue = String(detail?.companyName || row.company || "").trim();
      const companyId = ensureCompany(companyNameValue);

      // Project display title comes from "Title of the project". Preserve a
      // previously resolved title while detail tabs are still loading; otherwise
      // fall back to the real tab name rather than the generic word "Project".
      const title = String(
        detail?.title ||
        prior?.name ||
        detail?.sheetTabName ||
        row.sheetTabName ||
        companyNameValue ||
        `#${projectId}`
      ).trim();

      // Summary funding wins because it is the complete 31-row financial
      // inventory. Detail funding is only a fallback. This guarantees ABB,
      // AVL, the $3.5M Aramco contract, and all 16 funded projects are counted
      // even if a legacy detail label is unusual or a detail request is late.
      const summaryFunding = parseMoneyCell(row.financialValue);
      const detailFunding = parseMoneyCell(detail?.financialValue);
      // Once the detail tab is available it is the authoritative editable
      // record. Summary cells are formulas and may lag briefly after a write.
      const funding = Number.isFinite(detailFunding)
        ? String(detailFunding)
        : Number.isFinite(summaryFunding)
          ? String(summaryFunding)
          : "";

      const combinedDates = Array.from({ length: 7 }, (_, index) => {
        const detailValue = detail?.srlDates?.[index] || "";
        if (parseFlexibleDate(detailValue)) return detailValue;
        return row.srlDates?.[index] || "";
      });
      const srlDatesByLevel = {};
      combinedDates.forEach((raw, index) => {
        const parsed = parseFlexibleDate(raw);
        if (parsed) srlDatesByLevel[String(index + 1)] = parsed.getTime();
      });
      const srlHistory = SRL_LEVELS
        .filter((level) => srlDatesByLevel[level])
        .map((level) => ({ level, date: srlDatesByLevel[level] }));
      const summarySrl = validSummarySrl(row.currentSrl);
      const detailSrl = detail ? detailDerivedSrl(detail) : "";
      const srlLevel = detailSrl || summarySrl || (srlHistory.length ? srlHistory[srlHistory.length - 1].level : "");

      const notes = { ...(prior?.srlNotes || {}) };
      Object.entries(detail?.srlComments || {}).forEach(([level, text]) => {
        if (String(text || "").trim()) {
          notes[level] = { text: String(text).trim(), author: "Google Sheet", updatedAt: now };
        }
      });

      const people = detail
        ? [detail.primaryContact, ...(detail.extraContacts || [])].filter(
            (person) => person && (person.name || person.email || person.phone),
          )
        : [];
      const contactIds = [];
      people.forEach((person, contactIndex) => {
        const id = ensureContact(person, companyId, projectId, contactIndex);
        if (id && !contactIds.includes(id)) contactIds.push(id);
      });
      const primary = people[0] || null;

      const startParsed = parseFlexibleDate(detail?.startDate);
      let deadlineParsed = parseFlexibleDate(detail?.deadline);
      if (srlDatesByLevel["7"]) deadlineParsed = new Date(srlDatesByLevel["7"]);

      return {
        ...(prior || {}),
        id: projectId,
        sheetProjectId: projectId,
        sheetTabName: String(detail?.sheetTabName || row.sheetTabName || prior?.sheetTabName || "").trim(),
        name: title,
        companyId,
        lead: String(detail?.leadPi || row.pi || prior?.lead || "").trim(),
        contactName: primary ? String(primary.name || "").trim() : (prior?.contactName || ""),
        contactTitle: primary ? String(primary.jobTitle || "").trim() : (prior?.contactTitle || ""),
        contactEmail: primary ? String(primary.email || "").trim() : (prior?.contactEmail || ""),
        contactNumber: primary ? String(primary.phone || "").trim() : (prior?.contactNumber || ""),
        contactIds: detailsAvailable ? contactIds : (prior?.contactIds || []),
        funding,
        srlLevel,
        potentialLevel: String(detail?.potential || prior?.potentialLevel || "").trim(),
        startDate: startParsed ? dateToMDYString(startParsed) : (prior?.startDate || ""),
        deadline: deadlineParsed ? dateToMDYString(deadlineParsed) : (prior?.deadline || ""),
        rfsNti: detail?.rfsNumber
          ? maskHashPrefix(detail.rfsNumber)
          : row.rfsNti
            ? maskHashPrefix(row.rfsNti)
            : (prior?.rfsNti || ""),
        status: srlLevel === "7" ? "Finished" : "Unfinished",
        updates: prior?.updates || [],
        createdAt: prior?.createdAt || now + rowIndex,
        updatedAt: now,
        srlHistory,
        srlNotes: notes,
        // A row that is present in the authoritative spreadsheet is active.
        // Old reconciliation versions must not leave valid source projects
        // hidden in Archive.
        archived: false,
        dataSource: prior?.dataSource || "sheet",
        sheetSyncPending: false,
      };
    });

    // Keep Firestore-backed website projects that have not reached Google
    // Sheets yet. They remain visible after refresh and are retried in the
    // background instead of being deleted when the 31-row summary is read.
    const representedProjectIds = new Set(
      finalProjects.map((project) => String(project.sheetProjectId || project.id || "").trim()),
    );
    const pendingAppProjects = projectsRef.current.filter((project) => {
      const id = String(project.sheetProjectId || project.id || "").trim();
      return project?.sheetSyncPending === true && id && !representedProjectIds.has(id);
    });
    const mergedProjects = [...finalProjects, ...pendingAppProjects.map((project) => ({ ...project }))];

    // If detail data was not available yet, retain the existing sheet contacts
    // until the detail request finishes instead of making the contacts page
    // flash empty. The 31-project/funding inventory still updates immediately.
    if (!detailsAvailable) {
      nextContacts = contactsRef.current.map((contact) => ({ ...contact }));
    }

    const usedCompanyIds = new Set([
      ...mergedProjects.map((project) => project.companyId).filter(Boolean),
      ...nextContacts.map((contact) => contact.companyId).filter(Boolean),
    ]);
    const finalCompanies = !detailsAvailable
      ? companiesRef.current.map((company) => ({ ...company }))
      : nextCompanies.filter(
          (company) => !String(company.id || "").startsWith("sheet-company:") || usedCompanyIds.has(company.id),
        );

    projectsRef.current = mergedProjects;
    companiesRef.current = finalCompanies;
    contactsRef.current = nextContacts;
    suppressNextAppDataAutosave.current = true;
    setProjects(mergedProjects);
    setCompanies(finalCompanies);
    setContacts(nextContacts);

    const fundedCount = mergedProjects.filter((project) => Number.isFinite(parseMoneyCell(project.funding))).length;
    const fundingTotal = mergedProjects.reduce((sum, project) => {
      const value = parseMoneyCell(project.funding);
      return sum + (Number.isFinite(value) ? value : 0);
    }, 0);
    console.info(
      `[Google Sheets sync] ${mergedProjects.length} projects; ${fundedCount} funded; $${fundingTotal.toLocaleString("en-US", { maximumFractionDigits: 2 })}`,
    );
  }

  const sheetPollInFlight = useRef(false);
  async function pollGoogleSheet() {
    if (sheetPollInFlight.current) return;
    sheetPollInFlight.current = true;

    const summaryPromise = getSummarySnapshot();
    const detailsPromise = getProjectDetailSnapshots();

    try {
      const summaryResult = await summaryPromise;
      const summaryRows = summaryResult && summaryResult.ok && Array.isArray(summaryResult.rows)
        ? summaryResult.rows
        : [];

      // Render the compact summary immediately when available.
      if (summaryRows.length) {
        applyCanonicalSheetSnapshot(summaryRows, null);
        setSheetLoading(false);
      } else if (summaryResult?.error) {
        console.error("Summary sheet sync failed:", summaryResult.error);
      }

      try {
        const detailResult = await detailsPromise;
        if (detailResult && detailResult.ok) {
          const details = detailResult.snapshots || [];
          // Normal case: enrich the summary. Recovery case: if a website-created
          // detail tab exists with a Project ID but the summary append was
          // interrupted, applyCanonicalSheetSnapshot adds it as an unmatched
          // explicit-ID project instead of deleting it on refresh.
          if (summaryRows.length || details.some((detail) => String(detail.projectId || "").trim())) {
            applyCanonicalSheetSnapshot(summaryRows, details);
          }
        } else if (detailResult?.error) {
          console.error("Detail sheet sync failed:", detailResult.error);
        }
      } catch (detailError) {
        console.error("Detail sheet sync failed:", detailError);
      }
    } catch (summaryError) {
      console.error("Summary sheet sync failed:", summaryError);
      try {
        const detailResult = await detailsPromise;
        if (detailResult?.ok) {
          const details = detailResult.snapshots || [];
          if (details.some((detail) => String(detail.projectId || "").trim())) {
            applyCanonicalSheetSnapshot([], details);
          }
        }
      } catch (_) {}
    } finally {
      sheetPollInFlight.current = false;
      setSheetLoading(false);
      // Firestore-backed projects are never lost if Google was temporarily
      // unavailable. Retry any pending Sheet writes after each reconciliation.
      setTimeout(() => retryPendingSheetProjects(), 0);
    }
  }

  useEffect(() => {
    if (!authChecked || !currentUser || !appDataReady) return;
    pollGoogleSheet();
    const id = setInterval(pollGoogleSheet, 60000);
    return () => clearInterval(id);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [authChecked, currentUser, appDataReady]);

  const [apName, setApName] = useState("");
  const [apCompanyId, setApCompanyId] = useState("");
  const [apCompanySearch, setApCompanySearch] = useState("");
  const [apCompanyMode, setApCompanyMode] = useState("select"); // "select" | "create"
  const [apNewCompanyName, setApNewCompanyName] = useState("");
  const [apNewCompanyIndustry, setApNewCompanyIndustry] = useState("");
  const [apNewCompanyHqLocation, setApNewCompanyHqLocation] = useState("");
  const [apNewCompanySbuLocation, setApNewCompanySbuLocation] = useState("");
  const [apLead, setApLead] = useState("");
  const [apFunding, setApFunding] = useState("");
  const [apSrl, setApSrl] = useState("");
  const [apSrlDate, setApSrlDate] = useState("");
  const [apPotential, setApPotential] = useState("");
  const [apDeadline, setApDeadline] = useState("");
  const [apStartDate, setApStartDate] = useState("");
  const [apRfsNti, setApRfsNti] = useState("");
  const [apErrors, setApErrors] = useState({});
  const [apSaving, setApSaving] = useState(false);

  // ---- Add Project: repeatable contact rows ------------------------------
  // The first row is kept as the project's own embedded contact fields
  // (contactName/contactTitle/contactEmail/contactNumber), same as before —
  // any additional rows become their own entries in the shared contacts list.
  const MAX_PROJECT_CONTACTS = 100;
  const [apContacts, setApContacts] = useState([blankContactRow()]);

  function addProjectContactRow() {
    setApContacts((prev) => (prev.length >= MAX_PROJECT_CONTACTS ? prev : [...prev, blankContactRow()]));
  }
  function removeProjectContactRow(key) {
    setApContacts((prev) => (prev.length <= 1 ? prev : prev.filter((row) => row.key !== key)));
  }
  function updateProjectContactRow(key, field, value) {
    setApContacts((prev) => prev.map((row) => (row.key === key ? { ...row, [field]: value } : row)));
  }
  function fillProjectContactRow(key, contact) {
    setApContacts((prev) =>
      prev.map((row) =>
        row.key === key
          ? {
              ...row,
              name: contact.contactName || "",
              title: contact.jobTitle || "",
              email: contact.contactEmail || "",
              number: contact.contactNumber || "",
            }
          : row
      )
    );
  }

  function handleApCompanyTextChange(text) {
    setApCompanySearch(text);
    setApCompanyId("");
  }
  function handleApSelectCompany(company) {
    setApCompanySearch(company.name);
    setApCompanyId(company.id);
  }

  function switchApCompanyMode(mode) {
    setApCompanyMode(mode);
    setApErrors((prev) => ({ ...prev, companyId: undefined, newCompanyName: undefined }));
    if (mode === "select") {
      setApNewCompanyName("");
      setApNewCompanyIndustry("");
      setApNewCompanyHqLocation("");
      setApNewCompanySbuLocation("");
    } else {
      setApCompanyId("");
      setApCompanySearch("");
    }
  }

  function resetAddProjectForm() {
    setApName("");
    setApCompanyId("");
    setApCompanySearch("");
    setApCompanyMode("select");
    setApNewCompanyName("");
    setApNewCompanyIndustry("");
    setApNewCompanyHqLocation("");
    setApNewCompanySbuLocation("");
    setApLead("");
    setApContacts([blankContactRow()]);
    setApFunding("");
    setApSrl("");
    setApSrlDate("");
    setApPotential("");
    setApDeadline("");
    setApStartDate("");
    setApRfsNti("");
    setApErrors({});
  }

  const FUNDING_RE = /^\d+(\.\d{1,2})?$/;
  const PHONE_RE = /^[0-9+\-() ]{7,}$/;

  async function handleSaveProject(e) {
    e.preventDefault();
    if (apSaving) return;
    const errors = {};
    if (!apName.trim()) errors.name = "Enter a project name.";
    if (apCompanyMode === "create") {
      const trimmedNewName = apNewCompanyName.trim();
      if (!trimmedNewName) errors.newCompanyName = "Enter a company name.";
      else if (companies.some((c) => c.name.toLowerCase() === trimmedNewName.toLowerCase()))
        errors.newCompanyName = "A company with that name already exists — select it instead.";
    } else if (!apCompanyId) {
      errors.companyId = "Select a company.";
    }
    if (!apLead.trim()) errors.lead = "Enter a PI name.";

    const primaryContact = apContacts[0] || { name: "", title: "", email: "", number: "" };
    if (primaryContact.email.trim() && !EMAIL_RE.test(primaryContact.email.trim()))
      errors.contactEmail = "Enter a valid email address.";
    if (primaryContact.number.trim() && !PHONE_RE.test(primaryContact.number.trim()))
      errors.contactNumber = "Enter a valid contact number.";

    const contactErrors = {};
    apContacts.slice(1).forEach((row) => {
      const rowErrors = {};
      if (row.email.trim() && !EMAIL_RE.test(row.email.trim()))
        rowErrors.email = "Enter a valid email address.";
      if (row.number.trim() && !PHONE_RE.test(row.number.trim()))
        rowErrors.number = "Enter a valid contact number.";
      if (Object.keys(rowErrors).length) contactErrors[row.key] = rowErrors;
    });
    if (Object.keys(contactErrors).length) errors.contacts = contactErrors;

    if (apFunding.trim() && !FUNDING_RE.test(stripCommas(apFunding).trim()))
      errors.funding = "Enter an amount, e.g. 25000 or 25000.50 (or 0 if not decided).";
    if (!apSrl) errors.srl = "Select a current SRL.";
    if (apSrlDate.trim() && !parseMDY(apSrlDate.trim()))
      errors.srlDate = "Enter a valid date as DD/MM/YYYY.";
    if (apStartDate.trim() && !parseMDY(apStartDate.trim()))
      errors.startDate = "Enter a valid date as DD/MM/YYYY.";
    if (apDeadline.trim() && !parseMDY(apDeadline.trim())) {
      errors.deadline = "Enter a valid date as DD/MM/YYYY.";
    } else if (
      apStartDate.trim() &&
      apDeadline.trim() &&
      parseMDY(apStartDate.trim()) &&
      parseMDY(apDeadline.trim()) &&
      daysBetween(apStartDate.trim(), apDeadline.trim()) < 0
    ) {
      errors.deadline = "Deadline must be on or after the start date.";
    }
    if (Object.keys(errors).length) return setApErrors(errors);

    const now = Date.now();
    let resolvedCompanyId = apCompanyId;
    let resolvedCompanyName = "";
    let createdCompany = null;
    if (apCompanyMode === "create") {
      createdCompany = {
        id: now.toString(36) + Math.random().toString(36).slice(2, 6) + "co",
        name: apNewCompanyName.trim(),
        about: "",
        industry: apNewCompanyIndustry.trim(),
        hqLocation: apNewCompanyHqLocation.trim(),
        sbuLocation: apNewCompanySbuLocation.trim(),
        updates: [],
        createdAt: now,
        updatedAt: now,
        archived: false,
      };
      resolvedCompanyId = createdCompany.id;
      resolvedCompanyName = createdCompany.name;
    } else {
      resolvedCompanyName = companies.find((c) => c.id === resolvedCompanyId)?.name || "";
    }

    const apExtraRows = apContacts
      .slice(1)
      .filter((row) => row.name.trim() || row.title.trim() || row.email.trim() || row.number.trim());
    const apMatchedContacts = [];
    const apFreshRows = [];
    apExtraRows.forEach((row) => {
      const match = findContactMatch(row, contacts);
      if (match) apMatchedContacts.push(applyContactRow(match, row, now));
      else apFreshRows.push(row);
    });
    const extraContacts = apFreshRows.map((row, idx) => ({
      id: (now + idx + 1).toString(36) + Math.random().toString(36).slice(2, 6),
      companyId: resolvedCompanyId,
      contactName: row.name.trim(),
      jobTitle: row.title.trim(),
      contactEmail: row.email.trim(),
      contactNumber: row.number.trim(),
      createdAt: now,
      updatedAt: now,
      archived: false,
    }));
    const apLinkedContacts = [...apMatchedContacts, ...extraContacts];
    const parsedApSrlDate = apSrlDate.trim() ? parseMDY(apSrlDate.trim()) : null;
    const apSrlReachedAt = parsedApSrlDate ? parsedApSrlDate.getTime() : now;
    const parsedApStartDate = apStartDate.trim() ? parseMDY(apStartDate.trim()) : null;
    const apSrlHistory = [{ level: apSrl, date: apSrlReachedAt }];
    if (apSrl !== "1" && parsedApStartDate) {
      apSrlHistory.unshift({ level: "1", date: parsedApStartDate.getTime() });
    }

    const newProject = {
      id: nextProjectId(),
      name: apName.trim(),
      companyId: resolvedCompanyId,
      lead: apLead.trim(),
      contactName: primaryContact.name.trim(),
      contactTitle: primaryContact.title.trim(),
      contactEmail: primaryContact.email.trim(),
      contactNumber: primaryContact.number.trim(),
      contactIds: apLinkedContacts.map((c) => c.id),
      funding: stripCommas(apFunding).trim(),
      srlLevel: apSrl,
      potentialLevel: apPotential,
      deadline: apDeadline.trim(),
      startDate: apStartDate.trim(),
      rfsNti: apRfsNti.trim(),
      status: apSrl === "7" ? "Finished" : "Unfinished",
      updates: [],
      createdAt: now,
      updatedAt: now,
      srlHistory: apSrlHistory,
      srlNotes: {},
      archived: false,
      dataSource: "app",
      sheetSyncPending: true,
    };

    // Firestore is the durable app store. Save there first so a browser refresh
    // can never erase a project that was successfully submitted in the UI.
    // Google Sheets is synchronized immediately afterwards; a failed Sheet
    // write remains pending and is retried in the background.
    setApSaving(true);
    setApErrors((prev) => ({ ...prev, sync: undefined }));

    const nextCompaniesForSave = createdCompany ? [...companiesRef.current, createdCompany] : [...companiesRef.current];
    const matchedByIdForSave = new Map(apMatchedContacts.map((c) => [c.id, c]));
    const nextContactsForSave = [
      ...contactsRef.current.map((c) => matchedByIdForSave.get(c.id) || c),
      ...extraContacts,
    ];
    const nextProjectsForSave = [...projectsRef.current, newProject];

    // Make the new project visible/pending immediately so a Sheet poll that
    // completes during the Firestore request cannot remove it from the UI.
    const previousProjectsForAdd = projectsRef.current;
    const previousCompaniesForAdd = companiesRef.current;
    const previousContactsForAdd = contactsRef.current;
    projectsRef.current = nextProjectsForSave;
    companiesRef.current = nextCompaniesForSave;
    contactsRef.current = nextContactsForSave;
    setProjects(nextProjectsForSave);
    setCompanies(nextCompaniesForSave);
    setContacts(nextContactsForSave);

    const firestoreResult = await saveAppData({
      projects: nextProjectsForSave,
      companies: nextCompaniesForSave,
      contacts: nextContactsForSave,
      updatedBy: currentUser?.username || currentUser?.email || currentUser?.uid || "",
    });

    if (!firestoreResult.ok) {
      projectsRef.current = previousProjectsForAdd;
      companiesRef.current = previousCompaniesForAdd;
      contactsRef.current = previousContactsForAdd;
      setProjects(previousProjectsForAdd);
      setCompanies(previousCompaniesForAdd);
      setContacts(previousContactsForAdd);
      setApSaving(false);
      setApErrors((prev) => ({
        ...prev,
        sync: `Firebase save failed: ${firestoreResult.error || "The app database did not confirm the new project."}`,
      }));
      return;
    }

    const mergedProjectsAfterAdd = Array.isArray(firestoreResult.projects) ? firestoreResult.projects : nextProjectsForSave;
    const mergedCompaniesAfterAdd = Array.isArray(firestoreResult.companies) ? firestoreResult.companies : nextCompaniesForSave;
    const mergedContactsAfterAdd = Array.isArray(firestoreResult.contacts) ? firestoreResult.contacts : nextContactsForSave;
    projectsRef.current = mergedProjectsAfterAdd;
    companiesRef.current = mergedCompaniesAfterAdd;
    contactsRef.current = mergedContactsAfterAdd;
    setProjects(mergedProjectsAfterAdd);
    setCompanies(mergedCompaniesAfterAdd);
    setContacts(mergedContactsAfterAdd);

    // Firestore is the save boundary for the UI. Once Firebase confirms the
    // project, return control to the user immediately. Google Sheets is a
    // background synchronization target and must never hold the form open.
    setApSaving(false);
    resetAddProjectForm();
    goTo("projects");

    const company = createdCompany || companies.find((c) => c.id === resolvedCompanyId) || null;

    // Start Google Sheets synchronization without awaiting it. The project is
    // already durable in Firestore with sheetSyncPending=true, so a slow or
    // temporarily unavailable Google API cannot make the Save button hang.
    void (async () => {
      const syncResult = await syncProjectToSheets(newProject, {
        createIfMissing: true,
        companyNameOverride: resolvedCompanyName,
        extraContacts: apLinkedContacts,
        refreshAfterSync: false,
      });

      let persistedProject = newProject;
      if (syncResult?.ok) {
        persistedProject = {
          ...newProject,
          sheetProjectId: newProject.id,
          sheetTabName: syncResult.sheetTabName || "",
          sheetSyncPending: false,
        };
        const syncedProjects = projectsRef.current.map((project) =>
          project.id === newProject.id ? persistedProject : project,
        );
        projectsRef.current = syncedProjects;
        setProjects(syncedProjects);
        const postSyncSave = await saveAppData({
          projects: syncedProjects,
          companies: companiesRef.current,
          contacts: contactsRef.current,
          updatedBy: currentUser?.username || currentUser?.email || currentUser?.uid || "",
        });
        if (!postSyncSave.ok) console.error("Firestore post-Sheets save failed:", postSyncSave.error);
        setStorageError("");
      } else {
        const syncError = syncResult?.error || "Unknown Google Sheets error";
        console.error("Project saved to Firebase; Google Sheets sync remains pending:", syncError);
        setStorageError(`Saved to Firebase, but Google Sheets sync is pending: ${syncError}`);
      }

      // The separate CRM activity sheet is also background-only.
      appendActiveRow({
        data: {
          kind: "project",
          companyName: company ? company.name : resolvedCompanyName,
          industry: company ? company.industry || "" : "",
          hqLocation: company ? company.hqLocation || "" : "",
          sbuLocation: company ? company.sbuLocation || "" : "",
          contactName: persistedProject.contactName,
          jobTitle: persistedProject.contactTitle,
          phone: persistedProject.contactNumber,
          email: persistedProject.contactEmail,
          srlLevel: persistedProject.srlLevel,
          contractValue: persistedProject.funding,
          comments: persistedProject.name,
          source: currentUser ? currentUser.username : "",
        },
      }).catch((err) => console.error("Sheets sync (project) failed:", err));

      extraContacts.forEach((contact) => {
        appendActiveRow({
          data: {
            kind: "contact",
            companyName: company ? company.name : resolvedCompanyName,
            industry: company ? company.industry || "" : "",
            hqLocation: company ? company.hqLocation || "" : "",
            sbuLocation: company ? company.sbuLocation || "" : "",
            contactName: contact.contactName,
            jobTitle: contact.jobTitle,
            phone: contact.contactNumber,
            email: contact.contactEmail,
            source: currentUser ? currentUser.username : "",
          },
        }).catch((err) => console.error("Sheets sync (contact) failed:", err));
      });

      if (syncResult?.ok) setTimeout(() => pollGoogleSheet(), 250);
    })();
  }

  function handleDiscardProject() {
    resetAddProjectForm();
    goTo("projects");
  }

  function openProject(id, returnView) {
    setSelectedProjectId(id);
    setProjectReturnView(returnView || "projects");
    goTo("viewProject");
  }

  const selectedProject = projects.find((p) => p.id === selectedProjectId) || null;

  async function persistProjectChange(updatedProject, options = {}) {
    if (!updatedProject) return { ok: false, error: "No project was supplied." };

    const pendingProject = {
      ...updatedProject,
      sheetSyncPending: true,
      // Keep the original source marker when present. sheetSyncPending is the
      // actual authority flag used by the reconciliation guard above.
      dataSource: updatedProject.dataSource || "app",
    };
    const nextProjects = projectsRef.current.map((project) =>
      project.id === pendingProject.id ? pendingProject : project,
    );
    const nextCompanies = Array.isArray(options.companiesOverride)
      ? options.companiesOverride
      : companiesRef.current;
    const nextContacts = Array.isArray(options.contactsOverride)
      ? options.contactsOverride
      : contactsRef.current;

    // Make the user edit authoritative in memory BEFORE any network await.
    // This ensures a Sheet poll that finishes while Firebase is saving sees
    // sheetSyncPending=true and cannot roll the form back to the old value.
    const previousProjects = projectsRef.current;
    const previousCompanies = companiesRef.current;
    const previousContacts = contactsRef.current;
    projectsRef.current = nextProjects;
    companiesRef.current = nextCompanies;
    contactsRef.current = nextContacts;
    setProjects(nextProjects);
    setCompanies([...nextCompanies]);
    setContacts([...nextContacts]);

    // Explicitly commit Firebase BEFORE leaving the edit action. Do not rely
    // only on the debounce because a refresh/navigation can happen first.
    const firestoreResult = await saveAppData({
      projects: nextProjects,
      companies: nextCompanies,
      contacts: nextContacts,
      updatedBy: currentUser?.username || currentUser?.email || currentUser?.uid || "",
    });
    if (!firestoreResult.ok) {
      projectsRef.current = previousProjects;
      companiesRef.current = previousCompanies;
      contactsRef.current = previousContacts;
      setProjects(previousProjects);
      setCompanies([...previousCompanies]);
      setContacts([...previousContacts]);
      return {
        ok: false,
        error: `Firebase save failed: ${firestoreResult.error || "The app database did not confirm the change."}`,
      };
    }

    // A transaction may have preserved records saved by another browser. Use
    // the merged server result locally too instead of hiding those records.
    const mergedProjectsAfterEdit = Array.isArray(firestoreResult.projects) ? firestoreResult.projects : nextProjects;
    const mergedCompaniesAfterEdit = Array.isArray(firestoreResult.companies) ? firestoreResult.companies : nextCompanies;
    const mergedContactsAfterEdit = Array.isArray(firestoreResult.contacts) ? firestoreResult.contacts : nextContacts;
    projectsRef.current = mergedProjectsAfterEdit;
    companiesRef.current = mergedCompaniesAfterEdit;
    contactsRef.current = mergedContactsAfterEdit;
    setProjects(mergedProjectsAfterEdit);
    setCompanies([...mergedCompaniesAfterEdit]);
    setContacts([...mergedContactsAfterEdit]);
    setStorageError("");

    // Sheets is synchronized immediately after the durable Firebase commit.
    // Keep sheetSyncPending=true on failure so sheet polling cannot roll the
    // edit back and retryPendingSheetProjects can try again later.
    void (async () => {
      const syncResult = await syncProjectToSheets(pendingProject, {
        extraContacts: options.extraContacts,
        companyNameOverride: options.companyNameOverride,
        refreshAfterSync: false,
      });
      if (!syncResult?.ok) {
        setStorageError(`Saved to Firebase, but Google Sheets sync is pending: ${syncResult?.error || "Unknown Google Sheets error"}`);
        return;
      }

      const latest = projectsRef.current.find((project) => project.id === pendingProject.id) || pendingProject;
      const syncedProject = {
        ...latest,
        sheetProjectId: latest.sheetProjectId || latest.id,
        sheetTabName: syncResult.sheetTabName || latest.sheetTabName || "",
        sheetSyncPending: false,
      };
      const syncedProjects = projectsRef.current.map((project) =>
        project.id === syncedProject.id ? syncedProject : project,
      );
      projectsRef.current = syncedProjects;
      setProjects(syncedProjects);

      const postSyncSave = await saveAppData({
        projects: syncedProjects,
        companies: companiesRef.current,
        contacts: contactsRef.current,
        updatedBy: currentUser?.username || currentUser?.email || currentUser?.uid || "",
      });
      if (!postSyncSave.ok) {
        console.error("Firestore post-Sheets edit save failed:", postSyncSave.error);
      }
      setStorageError("");
      setTimeout(() => pollGoogleSheet(), 1500);
    })();

    return { ok: true };
  }

  function toggleProjectFinished(project) {
    const nextStatus = project.status === "Finished" ? "Unfinished" : "Finished";
    const now = Date.now();
    const logEntries = [
      {
        id: Date.now().toString(36) + Math.random().toString(36).slice(2, 6) + "status",
        user: currentUser ? currentUser.username : "unknown",
        field: "Status",
        from: project.status,
        to: nextStatus,
        at: now,
      },
    ];

    // Marking a project finished should always mean SRL 7 — bump the level
    // and progression log automatically so it shows up correctly everywhere
    // (including the home page's finalized-projects list, which is keyed off
    // SRL level rather than status).
    let nextSrlLevel = project.srlLevel;
    let nextSrlHistory = project.srlHistory || [];
    if (nextStatus === "Finished" && project.srlLevel !== "7") {
      logEntries.push({
        id: Date.now().toString(36) + Math.random().toString(36).slice(2, 6) + "srl",
        user: currentUser ? currentUser.username : "unknown",
        field: "Current SRL",
        from: "SRL " + project.srlLevel,
        to: "SRL 7",
        at: now,
      });
      nextSrlLevel = "7";
      const hasLevel7 = nextSrlHistory.some((h) => h.level === "7");
      nextSrlHistory = hasLevel7 ? nextSrlHistory : [...nextSrlHistory, { level: "7", date: now }];
    }

    const updatedProject = {
      ...project,
      status: nextStatus,
      srlLevel: nextSrlLevel,
      srlHistory: nextSrlHistory,
      updates: [...(project.updates || []), ...logEntries],
      updatedAt: now,
    };
    void persistProjectChange(updatedProject);

  }

  // ---- Edit project form -----------------------------------------------------
  // Mirrors the Add project form field-for-field (including RFS/NTI, start
  // date, and the repeatable contact rows) so editing offers the same inputs
  // as creating a project.
  const [epName, setEpName] = useState("");
  const [epCompanyId, setEpCompanyId] = useState("");
  const [epLead, setEpLead] = useState("");
  const [epFunding, setEpFunding] = useState("");
  const [epSrl, setEpSrl] = useState("");
  const [epSrlDates, setEpSrlDates] = useState({});
  const [epPotential, setEpPotential] = useState("");
  const [epStartDate, setEpStartDate] = useState("");
  const [epDeadline, setEpDeadline] = useState("");
  const [epRfsNti, setEpRfsNti] = useState("");
  const [epErrors, setEpErrors] = useState({});
  const [epSaving, setEpSaving] = useState(false);

  // ---- Edit Project: repeatable contact rows (mirrors Add project) ------
  // Row 0 is the project's own embedded contact fields (contactName/
  // contactTitle/contactEmail/contactNumber), same as the Add form. Extra
  // rows correspond to entries in the shared contacts list, linked through
  // the project's contactIds — each carries `existingId` so saving can tell
  // an edited contact apart from a brand-new one.
  const [epContacts, setEpContacts] = useState([blankContactRow()]);

  function addEditContactRow() {
    setEpContacts((prev) => (prev.length >= MAX_PROJECT_CONTACTS ? prev : [...prev, blankContactRow()]));
  }
  function removeEditContactRow(key) {
    setEpContacts((prev) => (prev.length <= 1 ? prev : prev.filter((row) => row.key !== key)));
  }
  function updateEditContactRow(key, field, value) {
    setEpContacts((prev) => prev.map((row) => (row.key === key ? { ...row, [field]: value } : row)));
  }
  function fillEditContactRow(key, contact) {
    setEpContacts((prev) =>
      prev.map((row) =>
        row.key === key
          ? {
              ...row,
              name: contact.contactName || "",
              title: contact.jobTitle || "",
              email: contact.contactEmail || "",
              number: contact.contactNumber || "",
            }
          : row
      )
    );
  }

  function updateEpSrlDate(level, value) {
    setEpSrlDates((prev) => ({ ...prev, [level]: value }));
  }

  function startEditProject() {
    if (!selectedProject) return;
    setEpName(selectedProject.name);
    setEpCompanyId(selectedProject.companyId);
    setEpLead(selectedProject.lead);
    setEpFunding(maskFundingInput(selectedProject.funding));
    setEpSrl(selectedProject.srlLevel);
    const historyDates = {};
    (selectedProject.srlHistory || []).forEach((entry) => {
      if (entry && entry.level) historyDates[entry.level] = dateToMDYString(new Date(entry.date));
    });
    setEpSrlDates(historyDates);
    setEpPotential(selectedProject.potentialLevel);
    setEpStartDate(selectedProject.startDate || "");
    setEpDeadline(selectedProject.deadline || "");
    setEpRfsNti(selectedProject.rfsNti || "");

    const primaryRow = {
      key: makeContactRowKey(),
      name: selectedProject.contactName || "",
      title: selectedProject.contactTitle || "",
      email: selectedProject.contactEmail || "",
      number: selectedProject.contactNumber || "",
    };
    const extraRows = (selectedProject.contactIds || [])
      .map((id) => contacts.find((c) => c.id === id))
      .filter(Boolean)
      .map((c) => ({
        key: makeContactRowKey(),
        existingId: c.id,
        name: c.contactName || "",
        title: c.jobTitle || "",
        email: c.contactEmail || "",
        number: c.contactNumber || "",
      }));
    setEpContacts([primaryRow, ...extraRows]);

    setEpErrors({});
    goTo("editProject");
  }

  function handleCancelEditProject() {
    goTo("viewProject");
  }

  const EDIT_FIELD_LABELS = {
    name: "Project name",
    companyId: "Company",
    lead: "PI name",
    contactName: "Contact name",
    contactTitle: "Contact job title",
    contactEmail: "Contact email",
    contactNumber: "Contact number",
    funding: "Funding",
    srlLevel: "Current SRL",
    potentialLevel: "Potential",
    startDate: "Project Start Date",
    deadline: "Project End Date",
    rfsNti: "RFS/NTI number",
  };

  function editDisplayValue(key, value) {
    if (key === "companyId") return companyName(value);
    if (key === "funding") return "$" + formatMoney(value);
    if (key === "srlLevel") return "SRL " + value;
    return value;
  }

  async function handleSaveEditProject(e) {
    e.preventDefault();
    if (epSaving) return;
    const errors = {};
    if (!epName.trim()) errors.name = "Enter a project name.";
    if (!epCompanyId) errors.companyId = "Select a company.";
    if (!epLead.trim()) errors.lead = "Enter a PI name.";

    const primaryContact = epContacts[0] || { name: "", title: "", email: "", number: "" };
    if (primaryContact.email.trim() && !EMAIL_RE.test(primaryContact.email.trim()))
      errors.contactEmail = "Enter a valid email address.";
    if (primaryContact.number.trim() && !PHONE_RE.test(primaryContact.number.trim()))
      errors.contactNumber = "Enter a valid contact number.";

    const contactErrors = {};
    epContacts.slice(1).forEach((row) => {
      const rowErrors = {};
      if (row.email.trim() && !EMAIL_RE.test(row.email.trim()))
        rowErrors.email = "Enter a valid email address.";
      if (row.number.trim() && !PHONE_RE.test(row.number.trim()))
        rowErrors.number = "Enter a valid contact number.";
      if (Object.keys(rowErrors).length) contactErrors[row.key] = rowErrors;
    });
    if (Object.keys(contactErrors).length) errors.contacts = contactErrors;

    if (epFunding.trim() && !FUNDING_RE.test(stripCommas(epFunding).trim()))
      errors.funding = "Enter an amount, e.g. 25000 or 25000.50 (or 0 if not decided).";
    if (!epSrl) errors.srl = "Select a current SRL.";
    const srlDateErrors = {};
    SRL_LEVELS.forEach((lvl) => {
      const str = (epSrlDates[lvl] || "").trim();
      if (str && !parseMDY(str)) {
        srlDateErrors[lvl] = "Enter a valid date as DD/MM/YYYY.";
      }
    });
    if (Object.keys(srlDateErrors).length) errors.srlDates = srlDateErrors;
    if (epStartDate.trim() && !parseMDY(epStartDate.trim()))
      errors.startDate = "Enter a valid date as DD/MM/YYYY.";
    if (epDeadline.trim() && !parseMDY(epDeadline.trim())) {
      errors.deadline = "Enter a valid date as DD/MM/YYYY.";
    } else if (
      epStartDate.trim() &&
      epDeadline.trim() &&
      parseMDY(epStartDate.trim()) &&
      parseMDY(epDeadline.trim()) &&
      daysBetween(epStartDate.trim(), epDeadline.trim()) < 0
    ) {
      errors.deadline = "Deadline must be on or after the start date.";
    }
    if (Object.keys(errors).length) return setEpErrors(errors);

    const original = projects.find((p) => p.id === selectedProjectId);
    if (!original) return;

    const now = Date.now();

    // Reconcile the repeatable contact rows against the shared contacts
    // list, the same way Add project links extra rows to contactIds: rows
    // with an existingId are updated in place, rows without one become new
    // contact records, and a previously-linked contact whose row was
    // removed is simply unlinked (its record is left intact elsewhere).
    const extraRows = epContacts
      .slice(1)
      .filter((row) => row.name.trim() || row.title.trim() || row.email.trim() || row.number.trim());
    const epUnlinkedRows = extraRows.filter((row) => !row.existingId);
    const epMatchedPairs = [];
    const epFreshRows = [];
    epUnlinkedRows.forEach((row) => {
      const match = findContactMatch(row, contacts);
      if (match) epMatchedPairs.push([match.id, row]);
      else epFreshRows.push(row);
    });
    const newContacts = epFreshRows.map((row, idx) => ({
      id: (now + idx + 1).toString(36) + Math.random().toString(36).slice(2, 6),
      companyId: epCompanyId,
      contactName: row.name.trim(),
      jobTitle: row.title.trim(),
      contactEmail: row.email.trim(),
      contactNumber: row.number.trim(),
      createdAt: now,
      updatedAt: now,
      archived: false,
    }));
    const editedRowsById = new Map([
      ...extraRows.filter((row) => row.existingId).map((row) => [row.existingId, row]),
      ...epMatchedPairs,
    ]);
    const nextContactIds = [...editedRowsById.keys(), ...newContacts.map((c) => c.id)];

    const nextContactsForSave = [
      ...contactsRef.current.map((c) => {
        const row = editedRowsById.get(c.id);
        if (!row) return c;
        return {
          ...c,
          contactName: row.name.trim(),
          jobTitle: row.title.trim(),
          contactEmail: row.email.trim(),
          contactNumber: row.number.trim(),
          updatedAt: now,
        };
      }),
      ...newContacts,
    ];

    const nextValues = {
      name: epName.trim(),
      companyId: epCompanyId,
      lead: epLead.trim(),
      contactName: primaryContact.name.trim(),
      contactTitle: primaryContact.title.trim(),
      contactEmail: primaryContact.email.trim(),
      contactNumber: primaryContact.number.trim(),
      funding: stripCommas(epFunding).trim(),
      srlLevel: epSrl,
      potentialLevel: epPotential,
      startDate: epStartDate.trim(),
      deadline: epDeadline.trim(),
      rfsNti: epRfsNti.trim(),
    };

    const logEntries = [];
    Object.keys(nextValues).forEach((key) => {
      if (nextValues[key] !== original[key]) {
        logEntries.push({
          id: Date.now().toString(36) + Math.random().toString(36).slice(2, 6) + key,
          user: currentUser ? currentUser.username : "unknown",
          field: EDIT_FIELD_LABELS[key],
          from: editDisplayValue(key, original[key]),
          to: editDisplayValue(key, nextValues[key]),
          at: now,
        });
      }
    });

    const srlChanged = nextValues.srlLevel !== original.srlLevel;

    const manualSrlDates = {};
    SRL_LEVELS.forEach((lvl) => {
      const str = (epSrlDates[lvl] || "").trim();
      if (str) {
        const d = parseMDY(str);
        if (d) manualSrlDates[lvl] = d.getTime();
      }
    });
    // If the current SRL was bumped but the user didn't manually set a date
    // for that level, default it to right now (same behavior as before this
    // section was made editable).
    if (nextValues.srlLevel && !manualSrlDates[nextValues.srlLevel]) {
      manualSrlDates[nextValues.srlLevel] = now;
    }
    const nextSrlHistory = SRL_LEVELS.filter((lvl) => manualSrlDates[lvl]).map((lvl) => ({
      level: lvl,
      date: manualSrlDates[lvl],
    }));

    const originalHistoryKey = JSON.stringify(
      (original.srlHistory || []).map((h) => [h.level, h.date])
    );
    const nextHistoryKey = JSON.stringify(nextSrlHistory.map((h) => [h.level, h.date]));
    if (originalHistoryKey !== nextHistoryKey) {
      logEntries.push({
        id: Date.now().toString(36) + Math.random().toString(36).slice(2, 6) + "srlProgression",
        user: currentUser ? currentUser.username : "unknown",
        field: "SRL Progression",
        from: (original.srlHistory || []).map((h) => "SRL " + h.level + " (" + formatDateShort(h.date) + ")").join(", ") || "—",
        to: nextSrlHistory.map((h) => "SRL " + h.level + " (" + formatDateShort(h.date) + ")").join(", ") || "—",
        at: now,
      });
    }

    const reachedSrl7 = srlChanged && nextValues.srlLevel === "7";
    const nextStatus = reachedSrl7 ? "Finished" : original.status;
    if (reachedSrl7 && original.status !== "Finished") {
      logEntries.push({
        id: Date.now().toString(36) + Math.random().toString(36).slice(2, 6) + "status",
        user: currentUser ? currentUser.username : "unknown",
        field: "Status",
        from: original.status,
        to: "Finished",
        at: now,
      });
    }

    const updatedProject = {
      ...original,
      ...nextValues,
      contactIds: nextContactIds,
      status: nextStatus,
      srlHistory: nextSrlHistory,
      updates: [...(original.updates || []), ...logEntries],
      updatedAt: logEntries.length ? now : original.updatedAt,
    };
    setEpSaving(true);
    setEpErrors((prev) => ({ ...prev, sync: undefined }));
    const saveResult = await persistProjectChange(updatedProject, {
      contactsOverride: nextContactsForSave,
      extraContacts: extraRows.map((row) => ({
        name: row.name.trim(),
        jobTitle: row.title.trim(),
        email: row.email.trim(),
        phone: row.number.trim(),
      })),
    });
    setEpSaving(false);
    if (!saveResult.ok) {
      setEpErrors((prev) => ({ ...prev, sync: saveResult.error || "Could not save the project." }));
      return;
    }

    goTo("viewProject");
  }

  function saveSrlNote(level, text) {
    if (!selectedProject) return;
    const existing = selectedProject.srlNotes && selectedProject.srlNotes[level];
    const prevText = existing ? existing.text : "";
    if (text === prevText) return;
    const now = Date.now();
    const updatedProject = {
      ...selectedProject,
      srlNotes: {
        ...(selectedProject.srlNotes || {}),
        [level]: {
          text,
          author: currentUser ? currentUser.username : "unknown",
          updatedAt: now,
        },
      },
      updatedAt: now,
    };
    void persistProjectChange(updatedProject);
  }



  // ---- Add company form -----------------------------------------------------
  const [acName, setAcName] = useState("");
  const [acAbout, setAcAbout] = useState("");
  const [acIndustry, setAcIndustry] = useState("");
  const [acHqLocation, setAcHqLocation] = useState("");
  const [acSbuLocation, setAcSbuLocation] = useState("");
  const [acErrors, setAcErrors] = useState({});

  // ---- Add Company: repeatable contact rows -----------------------------
  const MAX_COMPANY_CONTACTS = 100;
  function makeContactRowKey() {
    return Date.now().toString(36) + Math.random().toString(36).slice(2, 6);
  }
  function blankContactRow() {
    return { key: makeContactRowKey(), name: "", title: "", email: "", number: "" };
  }
  const [acContacts, setAcContacts] = useState([blankContactRow()]);

  function addCompanyContactRow() {
    setAcContacts((prev) => (prev.length >= MAX_COMPANY_CONTACTS ? prev : [...prev, blankContactRow()]));
  }
  function removeCompanyContactRow(key) {
    setAcContacts((prev) => (prev.length <= 1 ? prev : prev.filter((row) => row.key !== key)));
  }
  function updateCompanyContactRow(key, field, value) {
    setAcContacts((prev) => prev.map((row) => (row.key === key ? { ...row, [field]: value } : row)));
  }
  function fillCompanyContactRow(key, contact) {
    setAcContacts((prev) =>
      prev.map((row) =>
        row.key === key
          ? {
              ...row,
              name: contact.contactName || "",
              title: contact.jobTitle || "",
              email: contact.contactEmail || "",
              number: contact.contactNumber || "",
            }
          : row
      )
    );
  }

  function resetAddCompanyForm() {
    setAcName("");
    setAcAbout("");
    setAcIndustry("");
    setAcHqLocation("");
    setAcSbuLocation("");
    setAcContacts([blankContactRow()]);
    setAcErrors({});
  }

  function handleSaveCompany(e) {
    e.preventDefault();
    const errors = {};
    if (!acName.trim()) errors.name = "Enter a company name.";
    else if (companies.some((c) => c.name.toLowerCase() === acName.trim().toLowerCase()))
      errors.name = "A company with that name already exists.";

    const contactErrors = {};
    acContacts.forEach((row) => {
      const rowErrors = {};
      if (row.email.trim() && !EMAIL_RE.test(row.email.trim()))
        rowErrors.email = "Enter a valid email address.";
      if (row.number.trim() && !PHONE_RE.test(row.number.trim()))
        rowErrors.number = "Enter a valid contact number.";
      if (Object.keys(rowErrors).length) contactErrors[row.key] = rowErrors;
    });
    if (Object.keys(contactErrors).length) errors.contacts = contactErrors;

    if (Object.keys(errors).length) return setAcErrors(errors);

    const newCompany = {
      id: Date.now().toString(36) + Math.random().toString(36).slice(2, 6),
      name: acName.trim(),
      about: acAbout.trim(),
      industry: acIndustry.trim(),
      hqLocation: acHqLocation.trim(),
      sbuLocation: acSbuLocation.trim(),
      updates: [],
      createdAt: Date.now(),
      updatedAt: Date.now(),
      archived: false,
    };
    setCompanies((prev) => [...prev, newCompany]);

    const now = Date.now();
    const newContacts = acContacts
      .filter((row) => row.name.trim() || row.title.trim() || row.email.trim() || row.number.trim())
      .map((row, idx) => ({
        id: (now + idx).toString(36) + Math.random().toString(36).slice(2, 6),
        companyId: newCompany.id,
        contactName: row.name.trim(),
        jobTitle: row.title.trim(),
        contactEmail: row.email.trim(),
        contactNumber: row.number.trim(),
        createdAt: now,
        updatedAt: now,
        archived: false,
      }));
    if (newContacts.length) {
      setContacts((prev) => [...prev, ...newContacts]);
      newContacts.forEach((contact) => {
        appendActiveRow({
          data: {
            kind: "contact",
            companyName: newCompany.name,
            industry: newCompany.industry,
            hqLocation: newCompany.hqLocation,
            sbuLocation: newCompany.sbuLocation,
            contactName: contact.contactName,
            jobTitle: contact.jobTitle,
            phone: contact.contactNumber,
            email: contact.contactEmail,
            source: currentUser ? currentUser.username : "",
          },
        }).catch((err) => console.error("Sheets sync (contact) failed:", err));
      });
    }

    resetAddCompanyForm();
    goTo("companies");
  }

  function handleDiscardCompany() {
    resetAddCompanyForm();
    goTo("companies");
  }

  function openCompany(id, returnView) {
    setSelectedCompanyId(id);
    setCompanyReturnView(returnView || "companies");
    goTo("viewCompany");
  }

  const selectedCompany = companies.find((c) => c.id === selectedCompanyId) || null;
  function companyName(id) {
    const match = companies.find((c) => c.id === id);
    return match ? match.name : "—";
  }

  function openContact(id, returnView) {
    setSelectedContactId(id);
    setContactReturnView(returnView || "contacts");
    goTo("viewContact");
  }

  const selectedContact = contacts.find((c) => c.id === selectedContactId) || null;

  const companyProjects = selectedCompany
    ? projects.filter((p) => p.companyId === selectedCompany.id && !p.archived)
    : [];
  const selectedCompanyProjectCount = companyProjects.length;
  const companyFinishedCount = companyProjects.filter((p) => p.status === "Finished").length;
  const companyCompletionRate =
    selectedCompanyProjectCount > 0 ? Math.round((companyFinishedCount / selectedCompanyProjectCount) * 100) : 0;
  const companyFundedProjects = companyProjects.filter((p) => hasFundingValue(p.funding));
  const companyTotalFunding = companyFundedProjects.reduce((sum, p) => sum + fundingNumber(p.funding), 0);
  const companyContacts = companyProjects.filter((p) => p.contactName || p.contactEmail || p.contactNumber);
  const companyGeneralContacts = selectedCompany
    ? contacts.filter((c) => c.companyId === selectedCompany.id)
    : [];

  // ---- Bulk import from Excel -------------------------------------------------
  const importInputRef = useRef(null);
  const [importStatus, setImportStatus] = useState("");

  const CERP_SRL_ROW_LABELS = {
    "1": "SRL1 (Prospect)",
    "2": "SRL2 (Lead Qualification)",
    "3": "SRL3 (Needs Assesment)",
    "4": "SRL4 (Proposal)",
    "5": "SRL5 (Evaluate)",
    "6": "SRL6 (Negotiate)",
    "7": "SRL7 (Sign)",
  };

  function handleImportButtonClick() {
    if (importInputRef.current) importInputRef.current.click();
  }

  // Detail sheets are laid out as label/value pairs in columns A/B (and C for
  // SRL progression comments). Build a lookup from normalized label -> {b, c}.
  function sheetToLabelMap(sheet) {
    const aoa = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
    const map = {};
    aoa.forEach((row) => {
      const label = normalizeLabel(row[0]);
      if (label) map[label] = { b: row[1], c: row[2] };
    });
    return map;
  }

  function splitProjectContactCells(nameRaw, emailRaw) {
    let name = String(nameRaw || "").trim();
    let email = String(emailRaw || "").trim();
    const looksLikeEmail = (value) => /[^\s@]+@[^\s@]+\.[^\s@]+/.test(String(value || ""));
    if (!name && email && !looksLikeEmail(email)) {
      name = email;
      email = "";
    } else if (looksLikeEmail(name) && email && !looksLikeEmail(email)) {
      [name, email] = [email, name];
    }
    const emails = email
      .split(/[;\n]+|,\s*(?=[^,;\s]+@)/)
      .map((v) => v.trim())
      .filter(Boolean);
    let names = name
      .split(/;|\n|\s+&\s+|,\s*(?=(?:dr\.?|doctor|mr\.?|mrs\.?|ms\.?|prof\.?|professor)\s+)/i)
      .map((v) => v.trim())
      .filter(Boolean);
    if (emails.length > 1) {
      if (names.length !== emails.length) {
        names = Array.from({ length: emails.length }, (_, i) => names[i] || (i === 0 ? name : ""));
      }
      return emails.map((value, i) => ({ name: names[i] || "", email: value }));
    }
    if (names.length > 1) {
      return names.map((person, i) => ({ name: person, email: i === 0 ? (emails[0] || "") : "" }));
    }
    return name || email ? [{ name, email: emails[0] || email }] : [];
  }

  function parseCerpProjectSheet(sheet, sheetName) {
    const map = sheetToLabelMap(sheet);
    let companyName = "";
    let leadPi = "";
    let title = "";
    let contactName = "";
    let contactEmail = "";
    let rfsNumberRaw = "";
    let financialRaw = "";
    let expectedStartRaw = "";
    let expectedEndRaw = "";
    const srlDates = {};
    const srlNotes = {};

    Object.keys(map).forEach((label) => {
      const { b, c } = map[label];
      if (label.includes("company")) companyName = String(b || "").trim();
      else if (label === "lead pi") leadPi = String(b || "").trim();
      else if (label.includes("title")) title = String(b || "").trim();
      else if (label.includes("email")) contactEmail = String(b || "").trim();
      else if (label.includes("industry contact")) contactName = String(b || "").trim();
      else if (label.includes("rfs number")) rfsNumberRaw = String(b || "").trim();
      else if (label.includes("financial value")) financialRaw = b;
      else if (label.includes("expected") && (label.includes("start") || label.includes("strat"))) {
        expectedStartRaw = b;
      } else if (label.includes("expected") && (label.includes("end") || label.includes("ebd"))) {
        expectedEndRaw = b;
      } else {
        const srlMatch = label.match(/^srl\s*0*([1-7])\b/);
        if (srlMatch) {
          const lvl = srlMatch[1];
          const d = parseFlexibleDate(b);
          if (d) srlDates[lvl] = d;
          const comment = String(c || "").trim();
          if (comment) srlNotes[lvl] = comment;
        }
      }
    });

    if (!companyName) companyName = sheetName;
    const projectName = title || sheetName;

    const srlHistory = SRL_LEVELS.filter((lvl) => srlDates[lvl]).map((lvl) => ({
      level: lvl,
      date: srlDates[lvl].getTime(),
    }));
    const srlLevel = srlHistory.length ? srlHistory[srlHistory.length - 1].level : "";

    const notes = {};
    Object.keys(srlNotes).forEach((lvl) => {
      notes[lvl] = { text: srlNotes[lvl], author: "Imported", updatedAt: Date.now() };
    });

    let fundingNum = typeof financialRaw === "number" ? financialRaw : parseFloat(String(financialRaw).replace(/,/g, ""));
    const funding = Number.isFinite(fundingNum) ? String(fundingNum) : "";

    const rfsNti = rfsNumberRaw ? maskHashPrefix(rfsNumberRaw) : "";

    const startParsed = parseFlexibleDate(expectedStartRaw);
    const startDate = startParsed ? dateToMDYString(startParsed) : "";

    let deadlineDate = parseFlexibleDate(expectedEndRaw);
    if (srlDates["7"]) deadlineDate = srlDates["7"];
    const deadline = deadlineDate ? dateToMDYString(deadlineDate) : "";

    const parsedContacts = splitProjectContactCells(contactName, contactEmail);
    const primaryContact = parsedContacts[0] || { name: contactName, email: contactEmail };

    return {
      companyName,
      name: projectName,
      lead: leadPi,
      contactName: primaryContact.name || "",
      contactEmail: primaryContact.email || "",
      contacts: parsedContacts,
      funding,
      srlLevel,
      startDate,
      deadline,
      rfsNti,
      srlHistory,
      srlNotes: notes,
    };
  }

  // ---- "Company / contact list" import format ---------------------------
  // A single sheet with one row per person: Company Name, Industry, HQ
  // Country, HQ Location, SBU Location, First Name, Last Name, Job Title,
  // Phone, Email, Status (SRL), Research Interests, Contract Value, etc.
  function mapHqCountry(raw) {
    const v = String(raw || "").trim().toLowerCase();
    if (!v) return "";
    if (v === "ik") return "IK";
    if (v === "ook") return "OOK";
    if (v === "sa" || v.includes("saudi")) return "IK";
    return "OOK";
  }

  function findHeaderIndex(headerRow, matchers) {
    for (let i = 0; i < headerRow.length; i++) {
      const label = normalizeLabel(headerRow[i]);
      if (!label) continue;
      if (matchers.some((m) => label.includes(m))) return i;
    }
    return -1;
  }

  function contactsListHeaderIndex(headerRow) {
    return {
      company: findHeaderIndex(headerRow, ["company name", "company"]),
      industry: findHeaderIndex(headerRow, ["industry"]),
      hqCountry: findHeaderIndex(headerRow, ["hq country"]),
      hqLocation: findHeaderIndex(headerRow, ["hq location"]),
      sbuLocation: findHeaderIndex(headerRow, ["sbu location"]),
      firstName: findHeaderIndex(headerRow, ["first name"]),
      lastName: findHeaderIndex(headerRow, ["last name"]),
      jobTitle: findHeaderIndex(headerRow, ["job title"]),
      phone: findHeaderIndex(headerRow, ["phone"]),
      email: findHeaderIndex(headerRow, ["email"]),
      srl: findHeaderIndex(headerRow, ["status (srl)", "status(srl)", "srl"]),
      researchInterests: findHeaderIndex(headerRow, ["research interests", "research interest"]),
      contractValue: findHeaderIndex(headerRow, ["contract value"]),
    };
  }

  function isContactsListSheet(headerRow) {
    const idx = contactsListHeaderIndex(headerRow);
    return idx.company !== -1 && idx.firstName !== -1 && idx.lastName !== -1;
  }

  function parseContactsListSheet(sheet) {
    const aoa = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
    if (!aoa.length) return [];
    const idx = contactsListHeaderIndex(aoa[0]);
    const cell = (row, key) => (idx[key] === -1 ? "" : row[idx[key]]);

    return aoa
      .slice(1)
      .map((row) => ({
        companyName: String(cell(row, "company") || "").trim(),
        industry: String(cell(row, "industry") || "").trim(),
        hqCountry: mapHqCountry(cell(row, "hqCountry")),
        hqLocation: String(cell(row, "hqLocation") || "").trim(),
        sbuLocation: String(cell(row, "sbuLocation") || "").trim(),
        firstName: String(cell(row, "firstName") || "").trim(),
        lastName: String(cell(row, "lastName") || "").trim(),
        jobTitle: String(cell(row, "jobTitle") || "").trim(),
        phone: String(cell(row, "phone") || "").trim(),
        email: String(cell(row, "email") || "").trim(),
        srlRaw: cell(row, "srl"),
        researchInterests: String(cell(row, "researchInterests") || "").trim(),
        contractValueRaw: cell(row, "contractValue"),
      }))
      .filter((r) => r.companyName || r.firstName || r.lastName);
  }

  function handleImportFileChange(e) {
    const file = e.target.files && e.target.files[0];
    if (!file) return;
    setImportStatus("Importing…");

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const data = new Uint8Array(evt.target.result);
        const workbook = XLSX.read(data, { type: "array", cellDates: true });

        let contactsListSheetName = "";
        for (const sn of workbook.SheetNames) {
          const headerRow = XLSX.utils.sheet_to_json(workbook.Sheets[sn], { header: 1, defval: "" })[0] || [];
          if (isContactsListSheet(headerRow)) {
            contactsListSheetName = sn;
            break;
          }
        }

        const isCerpFormat =
          !contactsListSheetName &&
          workbook.SheetNames.length > 1 &&
          workbook.SheetNames.some((n) => /summary/i.test(n));

        function projectKey(companyId, name) {
          return companyId + "::" + String(name || "").trim().toLowerCase();
        }
        const existingProjectKeys = new Set(projects.map((p) => projectKey(p.companyId, p.name)));
        function contactKey(companyId, name, email) {
          return companyId + "::" + String(name || "").trim().toLowerCase() + "::" + String(email || "").trim().toLowerCase();
        }
        const existingContactKeys = new Set(contacts.map((c) => contactKey(c.companyId, c.contactName, c.contactEmail)));
        let duplicateCount = 0;

        // NOTE: this import work must NOT live inside a setState updater —
        // React invokes updaters more than once (StrictMode / re-renders),
        // which previously appended every imported project twice.
        {
          const prevCompanies = companies;
          const companiesByName = new Map(
            prevCompanies.map((c) => [c.name.trim().toLowerCase(), c.id])
          );
          const newCompanies = [];

          function getOrCreateCompanyId(rawName, extraFields) {
            const trimmed = String(rawName || "").trim();
            if (!trimmed) return "";
            const key = trimmed.toLowerCase();
            if (companiesByName.has(key)) return companiesByName.get(key);
            const id = Date.now().toString(36) + Math.random().toString(36).slice(2, 6) + newCompanies.length;
            companiesByName.set(key, id);
            newCompanies.push({
              id,
              name: trimmed,
              about: "",
              industry: "",
              hqCountry: "",
              hqLocation: "",
              sbuLocation: "",
              updates: [],
              ...(extraFields || {}),
            });
            return id;
          }

          const newProjects = [];
          const newContacts = [];
          let importedCount = 0;
          let importedContactCount = 0;

          if (contactsListSheetName) {
            const rows = parseContactsListSheet(workbook.Sheets[contactsListSheetName]);

            rows.forEach((r, i) => {
              const companyId = getOrCreateCompanyId(r.companyName, {
                industry: r.industry,
                hqCountry: r.hqCountry,
                hqLocation: r.hqLocation,
                sbuLocation: r.sbuLocation,
              });
              if (!companyId) return;

              const contactName = [r.firstName, r.lastName].filter(Boolean).join(" ").trim();
              const now = Date.now() + i;

              if (r.researchInterests) {
                const key = projectKey(companyId, r.researchInterests);
                if (existingProjectKeys.has(key)) {
                  duplicateCount++;
                  return;
                }
                existingProjectKeys.add(key);

                const srlNum = parseInt(String(r.srlRaw).trim(), 10);
                const srlLevel = SRL_LEVELS.includes(String(srlNum)) ? String(srlNum) : "";
                const fundingNum =
                  typeof r.contractValueRaw === "number" ? r.contractValueRaw : parseFloat(r.contractValueRaw);
                const funding = Number.isFinite(fundingNum) ? String(fundingNum) : "";

                newProjects.push({
                  id: now.toString(36) + Math.random().toString(36).slice(2, 6),
                  name: r.researchInterests,
                  companyId,
                  lead: "",
                  contactName,
                  contactTitle: r.jobTitle,
                  contactEmail: r.email,
                  contactNumber: r.phone,
                  funding,
                  srlLevel,
                  potentialLevel: "",
                  startDate: "",
                  deadline: "",
                  rfsNti: "",
                  status: srlLevel === "7" ? "Finished" : "Unfinished",
                  updates: [],
                  createdAt: now,
                  updatedAt: now,
                  srlHistory: srlLevel ? [{ level: srlLevel, date: now }] : [],
                  srlNotes: {},
                  archived: false,
                });
                importedCount++;
              } else {
                if (!contactName && !r.email && !r.phone) return;
                const key = contactKey(companyId, contactName, r.email);
                if (existingContactKeys.has(key)) {
                  duplicateCount++;
                  return;
                }
                existingContactKeys.add(key);

                newContacts.push({
                  id: now.toString(36) + Math.random().toString(36).slice(2, 6),
                  companyId,
                  contactName,
                  jobTitle: r.jobTitle,
                  contactEmail: r.email,
                  contactNumber: r.phone,
                  createdAt: now,
                  updatedAt: now,
                  archived: false,
                });
                importedContactCount++;
              }
            });
          } else if (isCerpFormat) {
            const detailSheetNames = workbook.SheetNames.filter((n) => !/summary/i.test(n));
            detailSheetNames.forEach((sheetName, i) => {
              const parsed = parseCerpProjectSheet(workbook.Sheets[sheetName], sheetName);
              if (!parsed.companyName && !parsed.name) return;

              const companyId = getOrCreateCompanyId(parsed.companyName);
              const key = projectKey(companyId, parsed.name);
              if (existingProjectKeys.has(key)) {
                duplicateCount++;
                return;
              }
              existingProjectKeys.add(key);
              const now = Date.now() + i;

              newProjects.push({
                id: now.toString(36) + Math.random().toString(36).slice(2, 6),
                name: parsed.name,
                companyId,
                lead: parsed.lead,
                contactName: parsed.contactName,
                contactEmail: parsed.contactEmail,
                contactNumber: "",
                funding: parsed.funding,
                srlLevel: parsed.srlLevel,
                potentialLevel: "",
                startDate: parsed.startDate,
                deadline: parsed.deadline,
                rfsNti: parsed.rfsNti,
                status: parsed.srlLevel === "7" ? "Finished" : "Unfinished",
                updates: [],
                createdAt: now,
                updatedAt: now,
                srlHistory: parsed.srlHistory,
                srlNotes: parsed.srlNotes,
                archived: false,
              });
              importedCount++;

              // Every Industry Contact / Industry Contact Email pair becomes
              // a real entry in "List of all contacts". Multiple semicolon-
              // separated emails are split into separate people when possible.
              (parsed.contacts || []).forEach((person, contactIndex) => {
                const cKey = contactKey(companyId, person.name, person.email);
                if (!existingContactKeys.has(cKey)) {
                  existingContactKeys.add(cKey);
                  newContacts.push({
                    id: (now + contactIndex + 1).toString(36) + Math.random().toString(36).slice(2, 6),
                    companyId,
                    contactName: person.name || "",
                    jobTitle: "",
                    contactEmail: person.email || "",
                    contactNumber: "",
                    createdAt: now,
                    updatedAt: now,
                    archived: false,
                  });
                  importedContactCount++;
                } else {
                  duplicateCount++;
                }
              });
            });
          } else {
            const firstSheetName = workbook.SheetNames[0];
            const sheet = workbook.Sheets[firstSheetName];
            const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });

            const startingEmpty = projects.length === 0;
            const usedIds = new Set(projects.map((p) => p.id));
            let nextImportedProjectId = projects.reduce((max, project) => {
              const value = String(project.id || "").trim();
              return /^\d+$/.test(value) ? Math.max(max, Number(value)) : max;
            }, 31) + 1;

            rows.forEach((row, i) => {
              const companyId = getOrCreateCompanyId(row.company);
              const projectName = String(row.name || "").trim();
              if (!projectName) return;

              const key = projectKey(companyId, projectName);
              if (existingProjectKeys.has(key)) {
                duplicateCount++;
                return;
              }
              existingProjectKeys.add(key);

              const now = Date.now() + i;

              const rawIdValue =
                row.id !== undefined && row.id !== null && String(row.id).trim() !== ""
                  ? String(row.id).trim()
                  : "";
              const validImportedId = /^\d+$/.test(rawIdValue) && !usedIds.has(rawIdValue);
              let projectId = startingEmpty && validImportedId ? rawIdValue : "";
              if (!projectId) {
                while (usedIds.has(String(nextImportedProjectId))) nextImportedProjectId += 1;
                projectId = String(nextImportedProjectId);
                nextImportedProjectId += 1;
              }
              usedIds.add(projectId);

              const srlRaw = row.srl;
              const srlNum = typeof srlRaw === "number" ? srlRaw : parseInt(String(srlRaw).trim(), 10);
              const srlLevel = SRL_LEVELS.includes(String(srlNum)) ? String(srlNum) : "";

              const potentialRaw = String(row.potential || "").trim().toLowerCase();
              const potentialLevel =
                POTENTIAL_LEVELS.find((lvl) => lvl.toLowerCase() === potentialRaw) || "";

              const fundingRaw = row.cost;
              const fundingNum = typeof fundingRaw === "number" ? fundingRaw : parseFloat(fundingRaw);
              const funding =
                fundingRaw === "" || fundingRaw === undefined || Number.isNaN(fundingNum)
                  ? ""
                  : String(fundingRaw).trim();

              newProjects.push({
                id: projectId,
                name: projectName,
                companyId,
                lead: String(row.lead || "").trim(),
                contactName: String(row.contactPerson || "").trim(),
                contactEmail: String(row.contactEmail || "").trim(),
                contactNumber: "",
                funding,
                srlLevel,
                potentialLevel,
                startDate: "",
                deadline: "",
                rfsNti: "",
                status: srlLevel === "7" ? "Finished" : "Unfinished",
                updates: [],
                createdAt: now,
                updatedAt: now,
                srlHistory: srlLevel ? [{ level: srlLevel, date: now }] : [],
                srlNotes: {},
                archived: false,
              });
              importedCount++;
            });
          }

          setProjects((prev) => [...prev, ...newProjects]);
          setContacts((prev) => [...prev, ...newContacts]);

          if (contactsListSheetName) {
            const parts = [];
            if (importedCount) parts.push(importedCount + " project" + (importedCount === 1 ? "" : "s"));
            if (importedContactCount) parts.push(importedContactCount + " contact" + (importedContactCount === 1 ? "" : "s"));
            setImportStatus(
              parts.length === 0
                ? duplicateCount > 0
                  ? "No new records — " + duplicateCount + " already existed and " + (duplicateCount === 1 ? "was" : "were") + " skipped."
                  : "No matching rows found in that file."
                : "Imported " + parts.join(" and ") +
                  (newCompanies.length ? " and " + newCompanies.length + " new compan" + (newCompanies.length === 1 ? "y" : "ies") : "") +
                  (duplicateCount > 0 ? " (" + duplicateCount + " duplicate" + (duplicateCount === 1 ? "" : "s") + " skipped)" : "") +
                  "."
            );
          } else {
            setImportStatus(
              importedCount === 0 && importedContactCount === 0
                ? duplicateCount > 0
                  ? "No new records — " + duplicateCount + " already existed and " + (duplicateCount === 1 ? "was" : "were") + " skipped."
                  : "No project rows found in that file."
                : "Imported " + importedCount + " project" + (importedCount === 1 ? "" : "s") +
                  (importedContactCount ? " and " + importedContactCount + " contact" + (importedContactCount === 1 ? "" : "s") : "") +
                  (newCompanies.length ? " and " + newCompanies.length + " new compan" + (newCompanies.length === 1 ? "y" : "ies") : "") +
                  (duplicateCount > 0 ? " (" + duplicateCount + " duplicate" + (duplicateCount === 1 ? "" : "s") + " skipped)" : "") +
                  "."
            );
          }

          if (newCompanies.length) {
            setCompanies((prev) => [...prev, ...newCompanies]);
          }
        }
      } catch (err) {
        setImportStatus("Could not read that file. Please upload a valid .xlsx spreadsheet.");
      }
    };
    reader.readAsArrayBuffer(file);
    e.target.value = "";
  }

  function toISODate(ms) {
    if (!ms) return "";
    const d = new Date(ms);
    if (Number.isNaN(d.getTime())) return "";
    return d.toISOString().slice(0, 10);
  }

  // Stable collision-proof IDs for website-created projects. The original
  // spreadsheet's visible column A remains its normal 1,2,3... serial number;
  // this ID lives only in app metadata/detail tabs. Two browsers can therefore
  // create projects at the same time without both choosing "32".
  function nextProjectId() {
    return `app-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 10)}`;
  }

  // ---- Live Google Sheets sync -------------------------------------------
  // Keeps the "Sponsered Projects_SRL movement_April2026" spreadsheet in step
  // with the app: one row per project on SRL_OnePage_Summary, plus one detail
  // tab per project mirroring the manual export layout.
  function cellToText(value) {
    if (value instanceof Date) {
      return Number.isNaN(value.getTime()) ? "" : value.toLocaleDateString("en-GB");
    }
    if (value === null || value === undefined) return "";
    return String(value);
  }

  async function syncProjectToSheets(project, options = {}) {
    if (!project) return { ok: false, error: "No project data was supplied." };

    const stateCompanyName = companyName(project.companyId);
    const cName = String(
      options.companyNameOverride || (stateCompanyName === "—" ? "" : stateCompanyName) || ""
    ).trim();
    const sheetProjectId = project.sheetProjectId || project.id;
    const rows = buildStandardProjectSheetAoa(
      { ...project, id: sheetProjectId },
      cName,
    ).map((row) => row.map(cellToText));

    const explicitExtras = Array.isArray(options.extraContacts) ? options.extraContacts : null;
    const extraContacts = (explicitExtras || (project.contactIds || [])
      .map((id) => contacts.find((c) => c.id === id))
      .filter(Boolean))
      .map((c) => ({
        name: String(c.name ?? c.contactName ?? "").trim(),
        jobTitle: String(c.jobTitle ?? c.title ?? "").trim(),
        email: String(c.email ?? c.contactEmail ?? "").trim(),
        phone: String(c.phone ?? c.number ?? c.contactNumber ?? "").trim(),
      }))
      .filter((c) => c.name || c.jobTitle || c.email || c.phone);

    const projectPayload = {
      projectId: sheetProjectId,
      projectName: project.name || "",
      preferredName: cName || project.name || "Project",
      alternateNames: [project.name, cName].filter(Boolean),
      sheetTabName: project.sheetTabName || "",
      rows,
      primaryContact: {
        name: String(project.contactName || "").trim(),
        jobTitle: String(project.contactTitle || "").trim(),
        email: String(project.contactEmail || "").trim(),
        phone: String(project.contactNumber || "").trim(),
      },
      extraContacts,
      potential: project.potentialLevel || "",
      createIfMissing: options.createIfMissing === true,
    };

    try {
      let res;
      if (options.createIfMissing === true) {
        const history = project.srlHistory || [];
        const dateForLevel = (lvl) => {
          const entry = history.find((h) => h.level === lvl);
          return entry ? "'" + new Date(entry.date).toLocaleDateString("en-GB") : "";
        };
        res = await persistNewProject({
          data: {
            project: projectPayload,
            summary: {
              projectId: sheetProjectId,
              company: cName,
              pi: project.lead || "",
              rfsNti: project.rfsNti || "",
              currentSrl: project.srlLevel || "",
              srlDates: SRL_LEVELS.map((lvl) => dateForLevel(lvl)),
              financialValue: project.funding ? String(project.funding) : "TBD",
            },
          },
        });
      } else {
        res = await upsertProjectSheet({ data: projectPayload });
      }

      if (!res || res.ok !== true) {
        return {
          ok: false,
          error: String(res?.error || "Google Sheets did not confirm the project save."),
          sheetTabName: res?.sheetTabName || project.sheetTabName || "",
        };
      }

      const resolvedTab = res.sheetTabName || project.sheetTabName || "";
      if (options.refreshAfterSync === true) {
        // Let Google's read endpoint observe the committed write before
        // reconciling state. The detail-tab fallback also protects the project
        // if the summary row is temporarily delayed.
        setTimeout(() => pollGoogleSheet(), 1500);
      }
      return { ok: true, sheetTabName: resolvedTab };
    } catch (err) {
      console.error("Sheets project sync failed:", err);
      return { ok: false, error: String(err?.message || err || "Could not save to Google Sheets.") };
    }
  }

  async function retryPendingSheetProjects() {
    const pending = projectsRef.current.filter(
      (project) => project?.sheetSyncPending === true,
    );
    if (!pending.length) return;

    for (const project of pending) {
      const company = companiesRef.current.find((item) => item.id === project.companyId);
      const linkedContacts = (project.contactIds || [])
        .map((id) => contactsRef.current.find((contact) => contact.id === id))
        .filter(Boolean);
      const result = await syncProjectToSheets(project, {
        // A pending EDIT must update its existing tab, not run the new-project
        // creation path and append another summary row. Only projects that have
        // never acquired a tab are creation retries.
        createIfMissing: !String(project.sheetTabName || "").trim(),
        companyNameOverride: company?.name || "",
        extraContacts: linkedContacts,
        refreshAfterSync: false,
      });
      if (!result?.ok) continue;

      const synced = {
        ...project,
        sheetProjectId: project.sheetProjectId || project.id,
        sheetTabName: result.sheetTabName || project.sheetTabName || "",
        sheetSyncPending: false,
      };
      const nextProjects = projectsRef.current.map((item) =>
        item.id === project.id ? synced : item,
      );
      projectsRef.current = nextProjects;
      setProjects(nextProjects);
      const saveResult = await saveAppData({
        projects: nextProjects,
        companies: companiesRef.current,
        contacts: contactsRef.current,
        updatedBy: currentUser?.username || currentUser?.email || currentUser?.uid || "",
      });
      if (!saveResult.ok) console.error("Firestore pending-sync update failed:", saveResult.error);
      else setStorageError("");
    }
  }

  function removeProjectFromSheets(project) {
    if (!project) return;
    deleteSummaryRow({ data: { projectId: project.id } }).catch((err) =>
      console.error("Sheets summary delete failed:", err)
    );
    if (project.sheetTabName) {
      deleteProjectSheet({ data: { sheetTabName: project.sheetTabName } }).catch((err) =>
        console.error("Sheets tab delete failed:", err)
      );
    }
  }


  function buildCerpProjectSheetAoa(p) {
    const history = p.srlHistory || [];
    const dateForLevel = (lvl) => {
      const entry = history.find((h) => h.level === lvl);
      return entry ? new Date(entry.date) : "";
    };
    const noteForLevel = (lvl) => {
      const note = p.srlNotes ? p.srlNotes[lvl] : null;
      return note && note.text ? note.text : "";
    };
    const fundingValue = p.funding ? Number(p.funding) : "TBD";
    const startDateValue = parseMDY(p.startDate) || "";
    const endDateValue = parseMDY(p.deadline) || "";

    const aoa = [
      ["Company Name", companyName(p.companyId) === "—" ? "" : companyName(p.companyId), ""],
      ["Lead PI", p.lead || "", ""],
      ["Sub PI", "", ""],
      ["Lead RS/PostDoc", "", ""],
      ["Title of the project", p.name || "", ""],
      ["Industry Contact", p.contactName || "", ""],
      ["Industry Contact Job Title", p.contactTitle || "", ""],
      ["Industry Contact Email", p.contactEmail || "", ""],
      ["Industry Contact Phone", p.contactNumber || "", ""],
    ];

    // Additional contacts beyond the primary one (added via the repeatable
    // "Contacts" rows on the project form, and editable the same way from
    // the edit project page) each get their own numbered set of rows.
    const extraContacts = (p.contactIds || [])
      .map((id) => contacts.find((c) => c.id === id))
      .filter(Boolean);
    extraContacts.forEach((c, idx) => {
      const n = idx + 2; // primary contact is #1 (unnumbered), extras start at 2
      aoa.push(["Industry Contact " + n, c.contactName || "", ""]);
      aoa.push(["Industry Contact " + n + " Job Title", c.jobTitle || "", ""]);
      aoa.push(["Industry Contact " + n + " Email", c.contactEmail || "", ""]);
      aoa.push(["Industry Contact " + n + " Phone", c.contactNumber || "", ""]);
    });

    aoa.push(
      ["Proposal Submitted to RFS", p.rfsNti ? "Yes" : "No", ""],
      ["If Yes, RFS number", p.rfsNti || "", ""],
      ["Financial Value (USD)", fundingValue, ""],
      ["Potential", p.potentialLevel || "", ""],
      ["Expected Start Date", startDateValue, ""],
      ["Expected End Date", endDateValue, ""],
      ["SRL Progression Data", "Date", "Comments"],
      ["Pool", "", ""]
    );
    SRL_LEVELS.forEach((lvl) => {
      aoa.push([CERP_SRL_ROW_LABELS[lvl], dateForLevel(lvl), noteForLevel(lvl)]);
    });
    aoa.push(["Date last modified", p.updatedAt ? new Date(p.updatedAt) : "", ""]);
    aoa.push(["Lead from CERP Business team", "", ""]);
    return aoa;
  }

  // Fixed layout used for the live Google Sheet tabs so every project tab
  // matches the existing sheets exactly (row 9 = RFS number, row 10 = financial
  // value, rows 15-21 = SRL1..SRL7 — the rows the summary tab points at).
  function buildStandardProjectSheetAoa(p, companyNameOverride = "") {
    const history = p.srlHistory || [];
    const dmy = (ms) => {
      const d = new Date(ms);
      if (Number.isNaN(d.getTime())) return "";
      const m = d.toLocaleString("en-US", { month: "short" });
      return `${String(d.getDate()).padStart(2, "0")}-${m}-${String(d.getFullYear()).slice(-2)}`;
    };
    const dateForLevel = (lvl) => {
      const entry = history.find((h) => h.level === lvl);
      return entry ? dmy(entry.date) : "";
    };
    const noteForLevel = (lvl) => {
      const note = p.srlNotes ? p.srlNotes[lvl] : null;
      return note && note.text ? note.text : "";
    };
    const monthYear = (str) => {
      const d = parseMDY(str);
      if (!d) return "";
      return `${d.toLocaleString("en-US", { month: "short" })}-${String(d.getFullYear()).slice(-2)}`;
    };
    const stateCompanyName = companyName(p.companyId);
    const cName = String(
      companyNameOverride || (stateCompanyName === "—" ? "" : stateCompanyName) || ""
    ).trim();
    const funding = p.funding ? Number(stripCommas(String(p.funding))) : "";

    // Keep A:C in the individual Google Sheet tabs in the original workbook
    // layout: exactly the same 23-row structure as the supplied XLSX. Website-
    // only metadata/additional contacts are stored separately in E:G.
    const aoa = [
      ["Company Name", cName, ""],
      ["Lead PI", p.lead || "", ""],
      ["Sub PI", "", ""],
      ["Lead RS/PostDoc", "", ""],
      ["Title of the project", p.name || "", ""],
      ["Industry Contact", p.contactName || "", ""],
      ["Industry Contact Email", p.contactEmail || "", ""],
      ["Proposal Submitted to RFS", p.rfsNti ? "Yes" : "No", ""],
      ["If Yes, RFS number", p.rfsNti || "", ""],
      ["Financial Value", Number.isFinite(funding) && funding !== "" ? funding : "TBD", ""],
      ["Expected Strat Date", monthYear(p.startDate), ""],
      ["Expected Ebd Date", monthYear(p.deadline), ""],
      ["SRL Progression Data", "Date", "Comments"],
      ["Pool", "", ""],
    ];
    SRL_LEVELS.forEach((lvl) => {
      aoa.push([CERP_SRL_ROW_LABELS[lvl], dateForLevel(lvl), noteForLevel(lvl)]);
    });
    aoa.push(["Date last modified", p.updatedAt ? dmy(p.updatedAt) : "", ""]);
    aoa.push(["Lead from CERP Business team", "", ""]);
    return aoa;
  }


  function sanitizeSheetName(name, usedNames) {
    let base = String(name || "Project").replace(/[\\/?*[\]:]/g, "").trim().slice(0, 31) || "Project";
    let candidate = base;
    let n = 2;
    while (usedNames.has(candidate.toLowerCase())) {
      const suffix = " " + n;
      candidate = (base.slice(0, 31 - suffix.length) + suffix).trim();
      n++;
    }
    usedNames.add(candidate.toLowerCase());
    return candidate;
  }

  function computeColWidths(aoa, maxWidth) {
    const cap = maxWidth || 50;
    const colCount = aoa.reduce((max, row) => Math.max(max, row.length), 0);
    const widths = [];
    for (let col = 0; col < colCount; col++) {
      let maxLen = 8;
      aoa.forEach((row) => {
        const cell = row[col];
        let text = "";
        if (cell instanceof Date) text = cell.toLocaleDateString("en-US");
        else if (cell === null || cell === undefined) text = "";
        else text = String(cell);
        if (text.length > maxLen) maxLen = text.length;
      });
      widths.push({ wch: Math.min(cap, maxLen + 2) });
    }
    return widths;
  }

  function handleExportExcel() {
    const usedNames = new Set(["srl_onepage_summary"]);
    const workbook = XLSX.utils.book_new();

    const summaryRows = projects.map((p) => {
      const history = p.srlHistory || [];
      const dateForLevel = (lvl) => {
        const entry = history.find((h) => h.level === lvl);
        return entry ? new Date(entry.date) : "";
      };
      return [
        companyName(p.companyId) === "—" ? "" : companyName(p.companyId),
        p.lead || "",
        p.rfsNti || "",
        p.srlLevel ? Number(p.srlLevel) : "",
        p.potentialLevel || "",
        dateForLevel("1"),
        dateForLevel("2"),
        dateForLevel("3"),
        dateForLevel("4"),
        dateForLevel("5"),
        dateForLevel("6"),
        dateForLevel("7"),
        p.funding ? Number(p.funding) : "TBD",
      ];
    });
    const summaryAoa = [
      ["Company", "PI", "RFS/NTI", "Current SRL", "Potential", "SRL1", "SRL2", "SRL3", "SRL4", "SRL5", "SRL6", "SRL7", "Financial Value"],
      ...summaryRows,
    ];
    const summarySheet = XLSX.utils.aoa_to_sheet(summaryAoa);
    summarySheet["!cols"] = computeColWidths(summaryAoa, 24);
    XLSX.utils.book_append_sheet(workbook, summarySheet, "SRL_OnePage_Summary");

    projects.forEach((p) => {
      const cName = companyName(p.companyId) === "—" ? "Project" : companyName(p.companyId);
      const sheetName = sanitizeSheetName(cName, usedNames);
      const aoa = buildCerpProjectSheetAoa(p);
      const sheet = XLSX.utils.aoa_to_sheet(aoa);
      sheet["!cols"] = computeColWidths(aoa, 50);
      XLSX.utils.book_append_sheet(workbook, sheet, sheetName);
    });

    XLSX.writeFile(workbook, "SRL_Lists_CERP.xlsx");
  }

  function handleExportContacts() {
    const sortedContacts = [...contacts].sort((a, b) =>
      (a.contactName || "").localeCompare(b.contactName || "")
    );
    const aoa = [
      ["Contact name", "Job title", "Company", "Email", "Phone number"],
      ...sortedContacts.map((c) => [
        c.contactName || "",
        c.jobTitle || "",
        companyName(c.companyId) === "—" ? "" : companyName(c.companyId),
        c.contactEmail || "",
        c.contactNumber ? formatPhoneDisplay(c.contactNumber) : "",
      ]),
    ];
    const sheet = XLSX.utils.aoa_to_sheet(aoa);
    sheet["!cols"] = computeColWidths(aoa, 30);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, sheet, "Contacts");
    XLSX.writeFile(workbook, "SRL_Contacts.xlsx");
  }

  // ---- Generate Report ---------------------------------------------------
  const REPORT_SECTIONS = [
    { key: "srl", label: "By SRL" },
    { key: "funding", label: "By funding" },
    { key: "company", label: "By company" },
    { key: "completion", label: "By completion" },
    { key: "lead", label: "By PI" },
    { key: "total", label: "By total number of projects" },
  ];

  const [rpSections, setRpSections] = useState(new Set());
  const [rpSrlLevels, setRpSrlLevels] = useState(new Set());
  const [rpCompanyIds, setRpCompanyIds] = useState(new Set());
  const [rpLeads, setRpLeads] = useState(new Set());
  const [rpGenerating, setRpGenerating] = useState(false);
  const [rpError, setRpError] = useState("");

  // Grouped PI options for the "PIs to include" checklist — title variants
  // of the same person (e.g. "Mani Sarathy" / "Prof. Mani Sarathy") collapse
  // into a single entry here.
  const reportLeadOptions = [...buildLeadGroups(projects.filter((p) => !p.archived)).values()].sort((a, b) =>
    a.displayName.localeCompare(b.displayName)
  );

  function toggleReportSection(key) {
    setRpSections((prev) => {
      const next = new Set(prev);
      if (next.has(key)) next.delete(key);
      else next.add(key);
      return next;
    });
  }
  function toggleAllReportSections() {
    setRpSections((prev) =>
      prev.size === REPORT_SECTIONS.length ? new Set() : new Set(REPORT_SECTIONS.map((s) => s.key))
    );
  }
  function toggleReportSrlLevel(lvl) {
    setRpSrlLevels((prev) => {
      const next = new Set(prev);
      if (next.has(lvl)) next.delete(lvl);
      else next.add(lvl);
      return next;
    });
  }
  function toggleAllReportSrlLevels() {
    setRpSrlLevels((prev) => (prev.size === SRL_LEVELS.length ? new Set() : new Set(SRL_LEVELS)));
  }
  function toggleReportCompany(id) {
    setRpCompanyIds((prev) => {
      const next = new Set(prev);
      if (next.has(id)) next.delete(id);
      else next.add(id);
      return next;
    });
  }
  function toggleAllReportCompanies() {
    const activeCompanyIds = companies.filter((c) => !c.archived).map((c) => c.id);
    setRpCompanyIds((prev) => (prev.size === activeCompanyIds.length ? new Set() : new Set(activeCompanyIds)));
  }
  function toggleReportLead(name) {
    setRpLeads((prev) => {
      const next = new Set(prev);
      if (next.has(name)) next.delete(name);
      else next.add(name);
      return next;
    });
  }
  function toggleAllReportLeads() {
    setRpLeads((prev) =>
      prev.size === reportLeadOptions.length ? new Set() : new Set(reportLeadOptions.map((g) => g.key))
    );
  }

  // Draws a simple bar chart on an offscreen canvas and returns a PNG data
  // URL ready for jsPDF's addImage, along with the pixel dimensions used.
  function drawBarChart(labels, values, title) {
    const width = 640;
    const height = 360;
    const canvas = document.createElement("canvas");
    canvas.width = width;
    canvas.height = height;
    const ctx = canvas.getContext("2d");

    ctx.fillStyle = "#ffffff";
    ctx.fillRect(0, 0, width, height);

    ctx.fillStyle = "#14181f";
    ctx.font = "bold 16px Arial";
    ctx.textAlign = "center";
    ctx.fillText(title, width / 2, 26);

    const padLeft = 60;
    const padRight = 24;
    const padTop = 48;
    const padBottom = 70;
    const chartW = width - padLeft - padRight;
    const chartH = height - padTop - padBottom;
    const maxVal = Math.max(1, ...values);

    // axis
    ctx.strokeStyle = "#d7dade";
    ctx.lineWidth = 1;
    ctx.beginPath();
    ctx.moveTo(padLeft, padTop);
    ctx.lineTo(padLeft, padTop + chartH);
    ctx.lineTo(padLeft + chartW, padTop + chartH);
    ctx.stroke();

    const n = Math.max(labels.length, 1);
    const barGap = 14;
    const barW = Math.max(8, (chartW - barGap * (n - 1)) / n);

    labels.forEach((label, i) => {
      const val = values[i] || 0;
      const barH = maxVal > 0 ? (val / maxVal) * chartH : 0;
      const x = padLeft + i * (barW + barGap);
      const y = padTop + chartH - barH;

      ctx.fillStyle = "#0f6b5c";
      ctx.fillRect(x, y, barW, barH);

      // value on top of bar
      ctx.fillStyle = "#14181f";
      ctx.font = "11px Arial";
      ctx.textAlign = "center";
      const valLabel = val >= 1000 ? Math.round(val).toLocaleString() : String(val);
      ctx.fillText(valLabel, x + barW / 2, Math.max(14, y - 6));

      // x-axis label, truncated + rotated slightly if long
      ctx.save();
      ctx.translate(x + barW / 2, padTop + chartH + 16);
      const shortLabel = String(label).length > 12 ? String(label).slice(0, 11) + "…" : String(label);
      if (n > 6) {
        ctx.rotate(-Math.PI / 5);
        ctx.textAlign = "right";
      } else {
        ctx.textAlign = "center";
      }
      ctx.fillStyle = "#6b7280";
      ctx.font = "11px Arial";
      ctx.fillText(shortLabel, 0, 0);
      ctx.restore();
    });

    return { dataUrl: canvas.toDataURL("image/png"), width, height };
  }

  async function handleGenerateReport() {
    if (rpSections.size === 0) {
      setRpError("Select at least one section to include.");
      return;
    }
    setRpError("");
    setRpGenerating(true);

    try {
      const activeProjects = projects.filter((p) => !p.archived);
      const activeCompanies = companies.filter((c) => !c.archived);
      const activeContacts = contacts.filter((c) => !c.archived);

      const pdf = new MockPdf();
      const PAGE_W = pdf.internal.pageSize.getWidth();
      const PAGE_H = pdf.internal.pageSize.getHeight();
      const MARGIN = 50;
      const CONTENT_W = PAGE_W - MARGIN * 2;
      let y = MARGIN;

      function ensureSpace(needed) {
        if (y + needed > PAGE_H - MARGIN) {
          pdf.addPage();
          y = MARGIN;
        }
      }

      pdf.setFont("helvetica", "bold");
      pdf.setFontSize(20);
      pdf.text("SRL Project Pipelines — Report", MARGIN, y);
      y += 26;

      pdf.setFont("helvetica", "italic");
      pdf.setFontSize(10);
      pdf.setTextColor(107, 114, 128);
      pdf.text(
        "Generated " + new Date().toLocaleDateString("en-GB") + (currentUser ? " by " + currentUser.username : ""),
        MARGIN,
        y
      );
      pdf.setTextColor(20, 24, 31);
      y += 26;

      function addHeading(text) {
        y += 14;
        ensureSpace(30);
        pdf.setFont("helvetica", "bold");
        pdf.setFontSize(15);
        pdf.text(text, MARGIN, y);
        y += 8;
        pdf.setDrawColor(215, 218, 222);
        pdf.line(MARGIN, y, PAGE_W - MARGIN, y);
        y += 16;
      }
      function addSubheading(text) {
        ensureSpace(22);
        pdf.setFont("helvetica", "bold");
        pdf.setFontSize(12);
        const lines = pdf.splitTextToSize(text, CONTENT_W);
        lines.forEach((line) => {
          ensureSpace(16);
          pdf.text(line, MARGIN, y);
          y += 16;
        });
        y += 2;
      }
      function addBullet(text) {
        pdf.setFont("helvetica", "normal");
        pdf.setFontSize(10);
        const lines = pdf.splitTextToSize(text, CONTENT_W - 14);
        lines.forEach((line, i) => {
          ensureSpace(13);
          pdf.text((i === 0 ? "•  " : "   ") + line, MARGIN, y);
          y += 13;
        });
      }
      function addSubBullet(text) {
        pdf.setFont("helvetica", "normal");
        pdf.setFontSize(10);
        const INDENT = 18;
        const lines = pdf.splitTextToSize(text, CONTENT_W - INDENT - 14);
        lines.forEach((line, i) => {
          ensureSpace(13);
          pdf.text((i === 0 ? "–  " : "   ") + line, MARGIN + INDENT, y);
          y += 13;
        });
      }
      function addNote(text) {
        pdf.setFont("helvetica", "italic");
        pdf.setFontSize(10);
        pdf.setTextColor(107, 114, 128);
        const lines = pdf.splitTextToSize(text, CONTENT_W);
        lines.forEach((line) => {
          ensureSpace(13);
          pdf.text(line, MARGIN, y);
          y += 13;
        });
        pdf.setTextColor(20, 24, 31);
      }
      function addChart(labels, values, title) {
        if (!labels.length || values.every((v) => !v)) return;
        const chart = drawBarChart(labels, values, title);
        const imgW = 380;
        const imgH = imgW * (chart.height / chart.width);
        ensureSpace(imgH + 20);
        y += 6;
        pdf.addImage(chart.dataUrl, "PNG", MARGIN, y, imgW, imgH);
        y += imgH + 20;
      }
      function truncateToWidth(text, maxWidth) {
        const str = String(text ?? "");
        if (pdf.getTextWidth(str) <= maxWidth) return str;
        let lo = 0;
        let hi = str.length;
        while (lo < hi) {
          const mid = Math.ceil((lo + hi) / 2);
          if (pdf.getTextWidth(str.slice(0, mid) + "…") <= maxWidth) lo = mid;
          else hi = mid - 1;
        }
        return lo > 0 ? str.slice(0, lo) + "…" : "…";
      }
      // Draws a bordered table with a shaded, bold header row. Repeats the
      // header on every new page a table spills onto.
      function addTable(columns, rows) {
        const HEADER_H = 20;
        const ROW_H = 16;
        const CELL_PAD = 6;

        function drawHeaderRow() {
          pdf.setFillColor(243, 244, 246);
          pdf.rect(MARGIN, y, CONTENT_W, HEADER_H, "F");
          pdf.setDrawColor(209, 213, 219);
          pdf.rect(MARGIN, y, CONTENT_W, HEADER_H, "D");
          pdf.setFont("helvetica", "bold");
          pdf.setFontSize(9);
          pdf.setTextColor(20, 24, 31);
          let cx = MARGIN;
          columns.forEach((col) => {
            const label = truncateToWidth(col.label, col.width - CELL_PAD * 2);
            pdf.text(label, cx + CELL_PAD, y + HEADER_H - 7);
            cx += col.width;
          });
          y += HEADER_H;
        }

        ensureSpace(HEADER_H + ROW_H);
        drawHeaderRow();
        pdf.setFont("helvetica", "normal");
        pdf.setFontSize(9);
        pdf.setTextColor(20, 24, 31);

        rows.forEach((row) => {
          if (y + ROW_H > PAGE_H - MARGIN) {
            pdf.addPage();
            y = MARGIN;
            drawHeaderRow();
            pdf.setFont("helvetica", "normal");
            pdf.setFontSize(9);
            pdf.setTextColor(20, 24, 31);
          }
          let cx = MARGIN;
          columns.forEach((col, i) => {
            const cellText = truncateToWidth(row[i], col.width - CELL_PAD * 2);
            pdf.text(cellText, cx + CELL_PAD, y + ROW_H - 5);
            cx += col.width;
          });
          pdf.setDrawColor(230, 231, 233);
          pdf.line(MARGIN, y + ROW_H, MARGIN + CONTENT_W, y + ROW_H);
          y += ROW_H;
        });
        y += 14;
      }
      function companyIndustry(id) {
        const match = companies.find((c) => c.id === id);
        return match && match.industry ? match.industry : "—";
      }
      function projectDuration(p) {
        const start = parseMDY(p.startDate);
        const end = parseMDY(p.deadline);
        if (!start || !end) return "—";
        return formatDuration(start.getTime(), end.getTime());
      }

      // ---- By SRL --------------------------------------------------
      if (rpSections.has("srl")) {
        addHeading("By SRL");
        const levels = rpSrlLevels.size ? SRL_LEVELS.filter((l) => rpSrlLevels.has(l)) : SRL_LEVELS;
        const levelCounts = [];
        const levelFunding = [];
        levels.forEach((lvl) => {
          const lvlProjects = activeProjects.filter((p) => p.srlLevel === lvl);
          addSubheading("SRL " + lvl + " — " + lvlProjects.length + " project" + (lvlProjects.length === 1 ? "" : "s"));
          if (lvlProjects.length === 0) {
            addNote("No projects currently at this level.");
          } else {
            lvlProjects.forEach((p) => {
              addBullet(
                p.name +
                  " — " +
                  companyName(p.companyId) +
                  " · PI: " +
                  (p.lead || "—") +
                  " · Contact: " +
                  (p.contactName || "—") +
                  " · Funding: $" +
                  formatMoney(p.funding)
              );
            });
          }
          levelCounts.push(lvlProjects.length);
          levelFunding.push(lvlProjects.reduce((s, p) => s + (Number(p.funding) || 0), 0));
        });
        if (levels.length > 1) {
          addChart(levels.map((l) => "SRL " + l), levelCounts, "Number of projects per SRL");
          addChart(levels.map((l) => "SRL " + l), levelFunding, "Funding per SRL ($)");
        }
      }

      // ---- By funding ----------------------------------------------------
      if (rpSections.has("funding")) {
        addHeading("By Funding");
        const sorted = [...activeProjects].sort((a, b) => (Number(b.funding) || 0) - (Number(a.funding) || 0));
        const totalFunding = activeProjects.reduce((s, p) => s + (Number(p.funding) || 0), 0);
        addBullet("Total funding across all active projects: $" + formatMoney(totalFunding));
        sorted.forEach((p) => {
          addBullet(p.name + " — " + companyName(p.companyId) + " · $" + formatMoney(p.funding));
        });
        const top10ByFunding = sorted.slice(0, 10);
        const top10Ascending = [...top10ByFunding].reverse();
        addChart(
          top10Ascending.map((p) => p.name),
          top10Ascending.map((p) => Number(p.funding) || 0),
          "Top projects by funding ($)"
        );
      }

      // ---- By company ------------------------------------------------
      if (rpSections.has("company")) {
        addHeading("By Company");
        const selectedCompanies = rpCompanyIds.size
          ? activeCompanies.filter((c) => rpCompanyIds.has(c.id))
          : activeCompanies;
        selectedCompanies.forEach((c) => {
          const compProjects = activeProjects.filter((p) => p.companyId === c.id);
          const compContacts = activeContacts.filter((ct) => ct.companyId === c.id);
          const compLeads = [...buildLeadGroups(compProjects).values()].map((g) => g.displayName);
          const compFunding = compProjects.reduce((s, p) => s + (Number(p.funding) || 0), 0);
          addSubheading(c.name);
          if (c.about) addNote(c.about);
          addBullet("Projects (" + compProjects.length + "):" + (compProjects.length ? "" : " None"));
          compProjects.forEach((p) => {
            addSubBullet(
              p.name +
                " — SRL " +
                (p.srlLevel || "—") +
                " · PI: " +
                (p.lead || "—") +
                " · Funding: $" +
                formatMoney(p.funding)
            );
          });
          addBullet("PIs (" + compLeads.length + "):" + (compLeads.length ? "" : " —"));
          compLeads.forEach((name) => addSubBullet(name));
          const contactNames = compContacts.map((ct) => ct.contactName).filter(Boolean);
          addBullet("Contacts (" + contactNames.length + "):" + (contactNames.length ? "" : " —"));
          contactNames.forEach((name) => addSubBullet(name));
          addBullet("Total funding: $" + formatMoney(compFunding));
        });
        const companiesByFunding = [...selectedCompanies].sort((a, b) => {
          const fa = activeProjects.filter((p) => p.companyId === a.id).reduce((s, p) => s + (Number(p.funding) || 0), 0);
          const fb = activeProjects.filter((p) => p.companyId === b.id).reduce((s, p) => s + (Number(p.funding) || 0), 0);
          return fa - fb;
        });
        addChart(
          companiesByFunding.map((c) => c.name),
          companiesByFunding.map((c) =>
            activeProjects.filter((p) => p.companyId === c.id).reduce((s, p) => s + (Number(p.funding) || 0), 0)
          ),
          "Funding by company ($)"
        );
      }

      // ---- By lead ---------------------------------------------------
      if (rpSections.has("lead")) {
        addHeading("By PI");
        const leadGroups = buildLeadGroups(activeProjects);
        const allGroups = [...leadGroups.values()];
        const selectedGroups = rpLeads.size ? allGroups.filter((g) => rpLeads.has(g.key)) : allGroups;
        if (allGroups.length === 0) {
          addNote("No projects have a PI assigned yet.");
        }
        selectedGroups.forEach((g) => {
          const leadFunding = g.projects.reduce((s, p) => s + (Number(p.funding) || 0), 0);
          addSubheading(g.displayName);
          addTable(
            [
              { label: "No.", width: 32 },
              { label: "Brief Project Title", width: 144 },
              { label: "Industry", width: 64 },
              { label: "Current SRL", width: 82 },
              { label: "Project Cost (USD)", width: 126 },
              { label: "Duration", width: 64 },
            ],
            g.projects.map((p, idx) => [
              String(idx + 1),
              p.name || "—",
              companyIndustry(p.companyId),
              "SRL " + (p.srlLevel || "—"),
              "$" + formatMoney(p.funding),
              projectDuration(p),
            ])
          );
          addBullet("Total funding: $" + formatMoney(leadFunding));
        });
        const groupsByFunding = [...selectedGroups].sort((a, b) => {
          const fa = a.projects.reduce((s, p) => s + (Number(p.funding) || 0), 0);
          const fb = b.projects.reduce((s, p) => s + (Number(p.funding) || 0), 0);
          return fa - fb;
        });
        addChart(
          groupsByFunding.map((g) => g.displayName),
          groupsByFunding.map((g) => g.projects.reduce((s, p) => s + (Number(p.funding) || 0), 0)),
          "Funding by PI ($)"
        );
      }

      // ---- By completion -----------------------------------------------
      if (rpSections.has("completion")) {
        addHeading("By Completion");
        const finished = activeProjects.filter((p) => p.srlLevel === "7");
        const unfinished = activeProjects.filter((p) => p.srlLevel !== "7");
        const finishedFunding = finished.reduce((s, p) => s + (Number(p.funding) || 0), 0);
        const pct = activeProjects.length ? Math.round((finished.length / activeProjects.length) * 100) : 0;
        addBullet(
          "Finalized projects (SRL 7): " + finished.length + " of " + activeProjects.length + " (" + pct + "%)"
        );
        addBullet("Total funding tied to finalized projects: $" + formatMoney(finishedFunding));
        if (finished.length === 0) {
          addNote("No projects have reached SRL 7 yet.");
        } else {
          finished.forEach((p) => {
            addBullet(
              p.name +
                " — " +
                companyName(p.companyId) +
                " · Funding: $" +
                formatMoney(p.funding) +
                " · Contact: " +
                (p.contactName || "—")
            );
          });
        }
        addChart(["Finalized (SRL 7)", "Unfinished"], [finished.length, unfinished.length], "Project completion status");
      }

      // ---- Total number of projects -----------------------------------
      if (rpSections.has("total")) {
        addHeading("Overview");
        const total = activeProjects.length;
        const totalFunding = activeProjects.reduce((s, p) => s + (Number(p.funding) || 0), 0);
        const finishedCount = activeProjects.filter((p) => p.srlLevel === "7").length;
        const completionRate = total ? Math.round((finishedCount / total) * 100) : 0;
        addBullet("Total active projects: " + total);
        addBullet("Total funding: $" + formatMoney(totalFunding));
        addBullet("Completion rate: " + completionRate + "% (" + finishedCount + " of " + total + " finalized)");
        addChart(
          SRL_LEVELS.map((l) => "SRL " + l),
          SRL_LEVELS.map((l) => activeProjects.filter((p) => p.srlLevel === l).length),
          "Project distribution by SRL"
        );
      }

      pdf.save("SRL_Report_" + new Date().toISOString().slice(0, 10) + ".pdf");
    } catch (err) {
      console.error("Report generation failed:", err);
      setRpError("Something went wrong generating the report. Try again.");
    } finally {
      setRpGenerating(false);
    }
  }

  // ---- Edit company form -----------------------------------------------------
  const [ecName, setEcName] = useState("");
  const [ecAbout, setEcAbout] = useState("");
  const [ecIndustry, setEcIndustry] = useState("");
  const [ecHqLocation, setEcHqLocation] = useState("");
  const [ecSbuLocation, setEcSbuLocation] = useState("");
  const [ecErrors, setEcErrors] = useState({});

  function startEditCompany() {
    if (!selectedCompany) return;
    setEcName(selectedCompany.name);
    setEcAbout(selectedCompany.about || "");
    setEcIndustry(selectedCompany.industry || "");
    setEcHqLocation(selectedCompany.hqLocation || "");
    setEcSbuLocation(selectedCompany.sbuLocation || "");
    setEcErrors({});
    goTo("editCompany");
  }

  function handleCancelEditCompany() {
    goTo("viewCompany");
  }

  const EDIT_COMPANY_FIELD_LABELS = {
    name: "Company name",
    about: "About",
    industry: "Industry",
    hqLocation: "HQ Location",
    sbuLocation: "SBU Location",
  };

  function handleSaveEditCompany(e) {
    e.preventDefault();
    const errors = {};
    const trimmedName = ecName.trim();
    if (!trimmedName) errors.name = "Enter a company name.";
    else if (
      companies.some(
        (c) => c.id !== selectedCompanyId && c.name.toLowerCase() === trimmedName.toLowerCase()
      )
    )
      errors.name = "A company with that name already exists.";
    if (Object.keys(errors).length) return setEcErrors(errors);

    const original = companies.find((c) => c.id === selectedCompanyId);
    if (!original) return;

    const nextValues = {
      name: trimmedName,
      about: ecAbout.trim(),
      industry: ecIndustry.trim(),
      hqLocation: ecHqLocation.trim(),
      sbuLocation: ecSbuLocation.trim(),
      updatedAt: Date.now(),
    };

    const logEntries = [];
    Object.keys(nextValues).forEach((key) => {
      if (nextValues[key] !== original[key]) {
        logEntries.push({
          id: Date.now().toString(36) + Math.random().toString(36).slice(2, 6) + key,
          user: currentUser ? currentUser.username : "unknown",
          field: EDIT_COMPANY_FIELD_LABELS[key],
          from: original[key] || "—",
          to: nextValues[key] || "—",
          at: Date.now(),
        });
      }
    });

    setCompanies((prev) =>
      prev.map((c) =>
        c.id === selectedCompanyId
          ? { ...c, ...nextValues, updates: [...(c.updates || []), ...logEntries] }
          : c
      )
    );

    goTo("viewCompany");
  }

  // -------------------------------------------------------------------------
  const activeDashboardProjects = projects.filter((p) => !p.archived);
  const funnelCounts = SRL_LEVELS.map((lvl) => activeDashboardProjects.filter((p) => p.srlLevel === lvl).length);
  const funnelMax = Math.max(1, ...funnelCounts);
  const funnelTotal = activeDashboardProjects.length;

  const finishedCount = activeDashboardProjects.filter((p) => p.status === "Finished").length;
  const completionRate = funnelTotal > 0 ? Math.round((finishedCount / funnelTotal) * 100) : 0;
  const fundedProjects = activeDashboardProjects.filter((p) => hasFundingValue(p.funding));
  const totalFunding = fundedProjects.reduce((sum, p) => sum + fundingNumber(p.funding), 0);

  const SRL7_FUNDING_COLORS = [
    "#0f6b5c", "#1a8f7b", "#3aab97", "#6cc4b0", "#98a2b3",
    "#e8a33d", "#c97a2b", "#7a5cff", "#4f8ff0", "#e0637a",
  ];
  const srl7FundingChartData = activeDashboardProjects
    .filter(
      (p) =>
        Number(p.srlLevel) === 7 &&
        hasFundingValue(p.funding) &&
        fundingNumber(p.funding) > 0
    )
    .map((p) => ({ name: p.name, value: fundingNumber(p.funding) }))
    .sort((a, b) => b.value - a.value)
    .map((d, i) => ({ ...d, fill: SRL7_FUNDING_COLORS[i % SRL7_FUNDING_COLORS.length] }));
  const srl7TotalFunding = srl7FundingChartData.reduce((sum, d) => sum + d.value, 0);

  const highSrlProjects = projects
    .filter((p) => !p.archived && Number(p.srlLevel) >= 4 && Number(p.srlLevel) <= 6)
    .sort((a, b) => {
      const lvlDiff = Number(b.srlLevel) - Number(a.srlLevel);
      if (lvlDiff !== 0) return lvlDiff;
      return parseFloat(b.funding || 0) - parseFloat(a.funding || 0);
    });

  const finalizedProjects = projects
    .filter((p) => !p.archived && Number(p.srlLevel) === 7)
    .sort((a, b) => parseFloat(b.funding || 0) - parseFloat(a.funding || 0));

  // ---- Projects list: search + sort -----------------------------------------
  const [projectSearch, setProjectSearch] = useState("");
  const [projectSort, setProjectSort] = useState("");

  const POTENTIAL_RANK = { Low: 0, Medium: 1, High: 2 };

  const archivedProjectsList = projects
    .filter((p) => p.archived)
    .sort((a, b) => (b.updatedAt || b.createdAt || 0) - (a.updatedAt || a.createdAt || 0));

  const visibleProjects = (() => {
    const q = projectSearch.trim().toLowerCase();
    let list = projects.filter((p) => !p.archived);
    if (q) {
      list = list.filter((p) => {
        const company = companyName(p.companyId).toLowerCase();
        return (
          p.name.toLowerCase().includes(q) ||
          company.includes(q) ||
          p.lead.toLowerCase().includes(q) ||
          p.srlLevel.toLowerCase().includes(q) ||
          ("srl " + p.srlLevel).toLowerCase().includes(q)
        );
      });
    }
    list = [...list];
    switch (projectSort) {
      case "srl-desc":
        list.sort((a, b) => Number(b.srlLevel) - Number(a.srlLevel));
        break;
      case "srl-asc":
        list.sort((a, b) => Number(a.srlLevel) - Number(b.srlLevel));
        break;
      case "funding-desc":
        list.sort((a, b) => parseFloat(b.funding || 0) - parseFloat(a.funding || 0));
        break;
      case "funding-asc":
        list.sort((a, b) => parseFloat(a.funding || 0) - parseFloat(b.funding || 0));
        break;
      case "potential-desc":
        list.sort((a, b) => (POTENTIAL_RANK[b.potentialLevel] ?? -1) - (POTENTIAL_RANK[a.potentialLevel] ?? -1));
        break;
      case "potential-asc":
        list.sort((a, b) => (POTENTIAL_RANK[a.potentialLevel] ?? -1) - (POTENTIAL_RANK[b.potentialLevel] ?? -1));
        break;
      case "date-asc":
        list.sort((a, b) => (a.createdAt || 0) - (b.createdAt || 0));
        break;
      case "date-desc":
        list.sort((a, b) => (b.createdAt || 0) - (a.createdAt || 0));
        break;
      default:
        // No sort chosen — newest projects first, so a project you just
        // created shows up at the top of the list.
        list.sort((a, b) => (b.createdAt || 0) - (a.createdAt || 0));
        break;
    }
    return list;
  })();

  const unfinishedProjectsList = projects
    .filter((p) => p.status !== "Finished" && !p.archived)
    .sort((a, b) => (b.createdAt || 0) - (a.createdAt || 0));
  const finishedProjectsList = projects
    .filter((p) => p.status === "Finished" && !p.archived)
    .sort((a, b) => (b.createdAt || 0) - (a.createdAt || 0));

  // ---- Companies list: search + sort -----------------------------------------
  const [companySearch, setCompanySearch] = useState("");
  const [companySort, setCompanySort] = useState("name-asc");

  function companyProjectCount(c) {
    return projects.filter((p) => p.companyId === c.id).length;
  }

  const archivedCompaniesList = companies.filter((c) => c.archived).sort((a, b) => a.name.localeCompare(b.name));

  const visibleCompanies = (() => {
    const q = companySearch.trim().toLowerCase();
    let list = companies.filter((c) => !c.archived);
    if (q) {
      list = list.filter(
        (c) => c.name.toLowerCase().includes(q) || (c.about || "").toLowerCase().includes(q)
      );
    }
    list = [...list];
    switch (companySort) {
      case "name-desc":
        list.sort((a, b) => b.name.localeCompare(a.name));
        break;
      case "projects-desc":
        list.sort((a, b) => companyProjectCount(b) - companyProjectCount(a) || a.name.localeCompare(b.name));
        break;
      case "projects-asc":
        list.sort((a, b) => companyProjectCount(a) - companyProjectCount(b) || a.name.localeCompare(b.name));
        break;
      case "name-asc":
      default:
        list.sort((a, b) => a.name.localeCompare(b.name));
        break;
    }
    return list;
  })();

  // ---- Contacts list: search + sort -------------------------------------------
  const [contactSearch, setContactSearch] = useState("");
  const [contactSort, setContactSort] = useState("name-asc");

  const archivedContactsList = contacts
    .filter((c) => c.archived)
    .sort((a, b) => (a.contactName || "").localeCompare(b.contactName || ""));

  const visibleContacts = (() => {
    const q = contactSearch.trim().toLowerCase();
    let list = contacts.filter((c) => !c.archived);
    if (q) {
      list = list.filter((c) => {
        const company = companyName(c.companyId).toLowerCase();
        return (
          (c.contactName || "").toLowerCase().includes(q) ||
          (c.jobTitle || "").toLowerCase().includes(q) ||
          (c.contactEmail || "").toLowerCase().includes(q) ||
          (c.contactNumber || "").toLowerCase().includes(q) ||
          company.includes(q)
        );
      });
    }
    list = [...list];
    switch (contactSort) {
      case "name-desc":
        list.sort((a, b) => (b.contactName || "").localeCompare(a.contactName || ""));
        break;
      case "company-asc":
        list.sort((a, b) => companyName(a.companyId).localeCompare(companyName(b.companyId)));
        break;
      case "name-asc":
      default:
        list.sort((a, b) => (a.contactName || "").localeCompare(b.contactName || ""));
        break;
    }
    return list;
  })();

  // ---- Add contact form -------------------------------------------------------
  const [ncCompanyId, setNcCompanyId] = useState("");
  const [ncCompanySearch, setNcCompanySearch] = useState("");
  const [ncCompanyMode, setNcCompanyMode] = useState("select"); // "select" | "create"
  const [ncNewCompanyName, setNcNewCompanyName] = useState("");
  const [ncNewCompanyIndustry, setNcNewCompanyIndustry] = useState("");
  const [ncNewCompanyHqLocation, setNcNewCompanyHqLocation] = useState("");
  const [ncNewCompanySbuLocation, setNcNewCompanySbuLocation] = useState("");
  const [ncContactName, setNcContactName] = useState("");
  const [ncContactTitle, setNcContactTitle] = useState("");
  const [ncContactEmail, setNcContactEmail] = useState("");
  const [ncContactNumber, setNcContactNumber] = useState("");
  const [ncErrors, setNcErrors] = useState({});

  function handleNcCompanyTextChange(text) {
    setNcCompanySearch(text);
    setNcCompanyId("");
  }
  function handleNcSelectCompany(company) {
    setNcCompanySearch(company.name);
    setNcCompanyId(company.id);
  }

  function switchNcCompanyMode(mode) {
    setNcCompanyMode(mode);
    setNcErrors((prev) => ({ ...prev, companyId: undefined, newCompanyName: undefined }));
    if (mode === "select") {
      setNcNewCompanyName("");
      setNcNewCompanyIndustry("");
      setNcNewCompanyHqLocation("");
      setNcNewCompanySbuLocation("");
    } else {
      setNcCompanyId("");
      setNcCompanySearch("");
    }
  }

  function resetAddContactForm() {
    setNcCompanyId("");
    setNcCompanySearch("");
    setNcCompanyMode("select");
    setNcNewCompanyName("");
    setNcNewCompanyIndustry("");
    setNcNewCompanyHqLocation("");
    setNcNewCompanySbuLocation("");
    setNcContactName("");
    setNcContactTitle("");
    setNcContactEmail("");
    setNcContactNumber("");
    setNcErrors({});
  }

  function handleSaveContact(e) {
    e.preventDefault();
    const errors = {};
    if (!ncContactName.trim()) errors.contactName = "Enter a contact name.";
    if (ncCompanyMode === "create") {
      const trimmedNewName = ncNewCompanyName.trim();
      if (!trimmedNewName) errors.newCompanyName = "Enter a company name.";
      else if (companies.some((c) => c.name.toLowerCase() === trimmedNewName.toLowerCase()))
        errors.newCompanyName = "A company with that name already exists — select it instead.";
    } else if (!ncCompanyId) {
      errors.companyId = "Select a company.";
    }
    if (ncContactEmail.trim() && !EMAIL_RE.test(ncContactEmail.trim()))
      errors.contactEmail = "Enter a valid email address.";
    if (ncContactNumber.trim() && !PHONE_RE.test(ncContactNumber.trim()))
      errors.contactNumber = "Enter a valid contact number.";
    if (Object.keys(errors).length) return setNcErrors(errors);

    const now = Date.now();

    let resolvedCompanyId = ncCompanyId;
    if (ncCompanyMode === "create") {
      const newCompany = {
        id: now.toString(36) + Math.random().toString(36).slice(2, 6) + "co",
        name: ncNewCompanyName.trim(),
        about: "",
        industry: ncNewCompanyIndustry.trim(),
        hqLocation: ncNewCompanyHqLocation.trim(),
        sbuLocation: ncNewCompanySbuLocation.trim(),
        updates: [],
        createdAt: now,
        updatedAt: now,
        archived: false,
      };
      setCompanies((prev) => [...prev, newCompany]);
      resolvedCompanyId = newCompany.id;
    }

    const newContact = {
      id: now.toString(36) + Math.random().toString(36).slice(2, 6),
      companyId: resolvedCompanyId,
      contactName: ncContactName.trim(),
      jobTitle: ncContactTitle.trim(),
      contactEmail: ncContactEmail.trim(),
      contactNumber: ncContactNumber.trim(),
      createdAt: now,
      updatedAt: now,
      archived: false,
    };
    setContacts((prev) => [...prev, newContact]);
    const company =
      companies.find((c) => c.id === resolvedCompanyId) ||
      (ncCompanyMode === "create"
        ? {
            name: ncNewCompanyName.trim(),
            industry: ncNewCompanyIndustry.trim(),
            hqLocation: ncNewCompanyHqLocation.trim(),
            sbuLocation: ncNewCompanySbuLocation.trim(),
          }
        : null);
    appendActiveRow({
      data: {
        kind: "contact",
        companyName: company ? company.name : "",
        industry: company ? company.industry || "" : "",
        hqLocation: company ? company.hqLocation || "" : "",
        sbuLocation: company ? company.sbuLocation || "" : "",
        contactName: newContact.contactName,
        jobTitle: newContact.jobTitle,
        phone: newContact.contactNumber,
        email: newContact.contactEmail,
        source: currentUser ? currentUser.username : "",
      },
    }).catch((err) => console.error("Sheets sync (contact) failed:", err));
    resetAddContactForm();
    goTo("contacts");
  }

  function handleDiscardContact() {
    resetAddContactForm();
    goTo("contacts");
  }

  // ---- Edit contact form -------------------------------------------------
  const [edcCompanyId, setEdcCompanyId] = useState("");
  const [edcContactName, setEdcContactName] = useState("");
  const [edcContactTitle, setEdcContactTitle] = useState("");
  const [edcContactEmail, setEdcContactEmail] = useState("");
  const [edcContactNumber, setEdcContactNumber] = useState("");
  const [edcErrors, setEdcErrors] = useState({});

  function startEditContact() {
    if (!selectedContact) return;
    openEditContact(selectedContact.id, contactReturnView);
  }

  function openEditContact(id, returnView) {
    const c = contacts.find((x) => x.id === id);
    if (!c) return;
    setSelectedContactId(id);
    setContactReturnView(returnView || "contacts");
    setEdcCompanyId(c.companyId || "");
    setEdcContactName(c.contactName || "");
    setEdcContactTitle(c.jobTitle || "");
    setEdcContactEmail(c.contactEmail || "");
    setEdcContactNumber(c.contactNumber || "");
    setEdcErrors({});
    goTo("editContact");
  }

  function handleCancelEditContact() {
    goTo("viewContact");
  }

  function handleSaveEditContact(e) {
    e.preventDefault();
    const errors = {};
    if (!edcContactName.trim()) errors.contactName = "Enter a contact name.";
    if (!edcCompanyId) errors.companyId = "Select a company.";
    if (edcContactEmail.trim() && !EMAIL_RE.test(edcContactEmail.trim()))
      errors.contactEmail = "Enter a valid email address.";
    if (edcContactNumber.trim() && !PHONE_RE.test(edcContactNumber.trim()))
      errors.contactNumber = "Enter a valid contact number.";
    if (Object.keys(errors).length) return setEdcErrors(errors);

    const nextValues = {
      companyId: edcCompanyId,
      contactName: edcContactName.trim(),
      jobTitle: edcContactTitle.trim(),
      contactEmail: edcContactEmail.trim(),
      contactNumber: edcContactNumber.trim(),
      updatedAt: Date.now(),
    };

    setContacts((prev) =>
      prev.map((c) => (c.id === selectedContactId ? { ...c, ...nextValues } : c))
    );

    goTo("viewContact");
  }

  const AUTHED_VIEWS = ["dashboard", "projects", "unfinishedProjects", "finishedProjects", "companies", "addProject", "addCompany", "addContact", "viewProject", "viewCompany", "editProject", "editCompany", "contacts", "viewContact", "editContact", "archivedProjects", "archivedCompanies", "archivedContacts", "generateReport"];

  if (!authChecked) {
    return (
      <Shell>
        <div className="auth-loading">Loading…</div>
      </Shell>
    );
  }

  if (currentUser && AUTHED_VIEWS.includes(view)) {
    return (
      <Shell>
        <AuthedLayout currentUser={currentUser} onSignOut={handleSignOut} storageError={storageError}>
          {view === "dashboard" && sheetLoading && projects.length === 0 && (
            <div className="home-content">
              <SkeletonCards count={6} />
            </div>
          )}
          {view === "dashboard" && !(sheetLoading && projects.length === 0) && (
            <div className="home-content">
              <div className="home-welcome">
                <div className="empty-state-mark">
                  <BrandLogo size={28} />
                </div>
                <h1 className="empty-state-title">Welcome, {currentUser.username}.</h1>
              </div>

              <div className="home-overview">
                <div className="funnel-chart">
                  {SRL_LEVELS.map((lvl, i) => {
                    const count = funnelCounts[i];
                    const pct = funnelTotal > 0 ? Math.round((count / funnelTotal) * 100) : 0;
                    const widthPct = Math.max(count > 0 ? 10 : 4, (count / funnelMax) * 100);
                    return (
                      <div className="funnel-row" key={lvl}>
                        <span className="funnel-label">SRL {lvl}</span>
                        <div className="funnel-track">
                          <div
                            className={"funnel-bar funnel-bar-srl-" + lvl + (count === 0 ? " funnel-bar-empty" : "")}
                            style={{ width: widthPct + "%" }}
                          >
                            <span className="funnel-count">{count}</span>
                          </div>
                        </div>
                        <span className="funnel-percent">{pct}%</span>
                      </div>
                    );
                  })}

                  <div className="funnel-key">
                    <p className="funnel-key-title">SRL Key</p>
                    <ul className="funnel-key-list">
                      {SRL_KEY.map((item) => (
                        <li className="funnel-key-item" key={item.lvl}>
                          <span className={"funnel-key-dot funnel-bar-srl-" + item.lvl}></span>
                          <span className="funnel-key-text">
                            <strong>SRL {item.lvl} · {item.name}</strong>
                            <span className="funnel-key-divider" aria-hidden="true">|</span>
                            <span className="funnel-key-desc">{item.desc}</span>
                          </span>
                        </li>
                      ))}
                    </ul>
                  </div>
                </div>

                <div className="home-stats">
                  <div className="stat-card">
                    <span className="stat-label">Total number of projects</span>
                    <span className="stat-value">{funnelTotal}</span>
                    <span className="stat-sub">across all companies</span>
                  </div>
                  <div className="stat-card">
                    <span className="stat-label">Completion rate</span>
                    <span className="stat-value">{completionRate}%</span>
                    <span className="stat-sub">{finishedCount} of {funnelTotal} projects finalized</span>
                  </div>
                  <div className="stat-card">
                    <span className="stat-label">Funding - Signed Contracts</span>
                    <span className="stat-value">${formatMoney(srl7TotalFunding)}</span>
                    <span className="stat-sub">across {srl7FundingChartData.length} SRL 7 project{srl7FundingChartData.length === 1 ? "" : "s"}</span>
                  </div>
                  <div className="stat-card">
                    <span className="stat-label">Funding in Pipeline</span>
                    <span className="stat-value">${formatMoney(totalFunding)}</span>
                    <span className="stat-sub">across {fundedProjects.length} funded project{fundedProjects.length === 1 ? "" : "s"}</span>
                  </div>
                </div>

                <div className="potential-column">
                  {srl7FundingChartData.length > 0 && (
                    <div className="potential-chart-card">
                      <span className="stat-label">SRL 7 funding by project</span>
                      <div className="potential-chart-wrap">
                        <ResponsiveContainer width="100%" height={160}>
                          <PieChart>
                            <Pie
                              data={srl7FundingChartData}
                              dataKey="value"
                              nameKey="name"
                              innerRadius={40}
                              outerRadius={70}
                              paddingAngle={2}
                            >
                              {srl7FundingChartData.map((d) => (
                                <Cell key={d.name} fill={d.fill} />
                              ))}
                            </Pie>
                            <Tooltip content={<PieFundingTooltip />} />
                          </PieChart>
                        </ResponsiveContainer>
                      </div>
                      {/* Custom scrollable legend — recharts' built-in <Legend>
                          doesn't clip itself when there are many finalized
                          projects, so it used to spill out of the card. This
                          caps the visible height and scrolls instead. */}
                      <div className="potential-chart-legend">
                        {srl7FundingChartData.map((d) => (
                          <div key={d.name} className="potential-chart-legend-item">
                            <span className="potential-chart-legend-dot" style={{ backgroundColor: d.fill }} />
                            <span className="potential-chart-legend-label">{d.name}</span>
                          </div>
                        ))}
                      </div>
                    </div>
                  )}
                  <div className="list-buttons-card">
                    <button type="button" className="btn-choice btn-list-companies" onClick={() => goTo("companies")}>
                      List of companies
                    </button>
                    <button type="button" className="btn-choice btn-list-contacts" onClick={() => goTo("contacts")}>
                      List of all contacts
                    </button>
                    <button type="button" className="btn-choice btn-list-projects" onClick={() => goTo("projects")}>
                      List of all projects
                    </button>
                  </div>
                </div>
              </div>

              <div className="home-top-projects">
                <h2 className="section-title">High SRL Projects (Unfinished)</h2>
                {highSrlProjects.length === 0 ? (
                  <p className="empty-table-note">No projects at SRL 4–6 yet.</p>
                ) : (
                  <div className="table-wrap table-wrap-scroll">
                    <table className="data-table">
                      <thead>
                        <tr>
                          <th>SRL</th>
                          <th>Project</th>
                          <th>Company</th>
                          <th>PI</th>
                          <th>Funding</th>
                          <th>Potential</th>
                        </tr>
                      </thead>
                      <tbody>
                        {highSrlProjects.map((p) => (
                          <tr key={p.id} className="data-row" onClick={() => openProject(p.id, "dashboard")}>
                            <td>{p.srlLevel}</td>
                            <td>{p.name}</td>
                            <td>{companyName(p.companyId)}</td>
                            <td>{p.lead}</td>
                            <td>${formatMoney(p.funding)}</td>
                            <td>
                              <PotentialChip level={p.potentialLevel} />
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                )}
              </div>

              <div className="home-top-projects">
                <h2 className="section-title">Finalized Projects</h2>
                {finalizedProjects.length === 0 ? (
                  <p className="empty-table-note">No finalized (SRL 7) projects yet.</p>
                ) : (
                  <div className="table-wrap table-wrap-scroll">
                    <table className="data-table">
                      <thead>
                        <tr>
                          <th>SRL</th>
                          <th>Project</th>
                          <th>Company</th>
                          <th>PI</th>
                          <th>Funding</th>
                          <th>Potential</th>
                        </tr>
                      </thead>
                      <tbody>
                        {finalizedProjects.map((p) => (
                          <tr key={p.id} className="data-row" onClick={() => openProject(p.id, "dashboard")}>
                            <td>{p.srlLevel}</td>
                            <td>{p.name}</td>
                            <td>{companyName(p.companyId)}</td>
                            <td>{p.lead}</td>
                            <td>${formatMoney(p.funding)}</td>
                            <td>
                              <PotentialChip level={p.potentialLevel} />
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                )}
              </div>

              <div className="button-row">
                <button type="button" className="btn-choice btn-choice-report" onClick={() => goTo("generateReport")}>
                  Generate Report
                </button>
              </div>

              <div className="import-row">
                <button type="button" className="btn-import" onClick={handleImportButtonClick}>
                  <Upload size={14} strokeWidth={1.75} />
                  Import Spreadsheet
                </button>
                <input
                  ref={importInputRef}
                  type="file"
                  accept=".xlsx,.xls"
                  className="import-file-input"
                  onChange={handleImportFileChange}
                />
                <button type="button" className="btn-import" onClick={handleExportExcel}>
                  <Download size={14} strokeWidth={1.75} />
                  Export Spreadsheet
                </button>
                {importStatus ? <span className="import-status">{importStatus}</span> : null}
              </div>

              <div className="button-row dashboard-bottom-row">
                <a
                  className="btn-choice"
                  href="https://docs.google.com/spreadsheets/d/17Y47rJ8alS9tiAUvKYR5ADlhWoGZrHo_2FFJdvTLnYs/edit?usp=sharing"
                  target="_blank"
                  rel="noopener noreferrer"
                >
                  Live Google Spreadsheet
                </a>
              </div>
            </div>
          )}

          {view === "projects" && sheetLoading && projects.length === 0 && (
            <div className="page-content">
              <h1 className="page-title">All projects</h1>
              <SkeletonCards count={6} />
            </div>
          )}
          {view === "projects" && !(sheetLoading && projects.length === 0) && (
            <div className="page-content">
              <div className="page-content-header">
                <button type="button" className="btn-back" onClick={() => goTo("dashboard")}>
                  <ArrowLeft size={15} strokeWidth={1.75} />
                  Back to home
                </button>
                <h1 className="page-title">All projects</h1>
              </div>

              <div className="list-toolbar">
                <div className="search-wrap">
                  <input
                    type="text"
                    className="search-input"
                    placeholder="Search by project, company, PI, or SRL"
                    value={projectSearch}
                    onChange={(e) => setProjectSearch(e.target.value)}
                  />
                </div>
                <div className="sort-select-wrap">
                  <SelectInput
                    id="project-sort"
                    value={projectSort}
                    onChange={(e) => setProjectSort(e.target.value)}
                  >
                    <option value="">Sort by…</option>
                    <option value="date-desc">Date (newest to oldest)</option>
                    <option value="date-asc">Date (oldest to newest)</option>
                    <option value="srl-desc">SRL (7 → 1)</option>
                    <option value="srl-asc">SRL (1 → 7)</option>
                    <option value="funding-desc">Funding (highest to lowest)</option>
                    <option value="funding-asc">Funding (lowest to highest)</option>
                    <option value="potential-desc">Potential (high to low)</option>
                    <option value="potential-asc">Potential (low to high)</option>
                  </SelectInput>
                </div>
              </div>

              <div className="button-row">
                <button type="button" className="btn-choice" onClick={() => goTo("unfinishedProjects")}>
                  Unfinished Projects
                </button>
                <button type="button" className="btn-choice" onClick={() => goTo("finishedProjects")}>
                  Finalized Projects
                </button>
                <button type="button" className="btn-choice" onClick={() => goTo("archivedProjects")}>
                  Archive{archivedProjectsList.length ? " (" + archivedProjectsList.length + ")" : ""}
                </button>
              </div>

              <div className="page-body">
                <button type="button" className="add-tile" onClick={() => goTo("addProject")}>
                  <span className="add-tile-icon">
                    <Plus size={20} strokeWidth={2} />
                  </span>
                  <span className="add-tile-label">Add project</span>
                </button>

                {visibleProjects.map((p) => (
                  <div
                    role="button"
                    tabIndex={0}
                    key={p.id}
                    className="add-tile project-tile"
                    onClick={() => openProject(p.id, "projects")}
                    onKeyDown={(e) => {
                      if (e.key === "Enter" || e.key === " ") {
                        e.preventDefault();
                        openProject(p.id, "projects");
                      }
                    }}
                  >
                    <div className="project-tile-status-row">
                      <label
                        className="project-tile-status-check"
                        onClick={(e) => e.stopPropagation()}
                      >
                        <input
                          type="checkbox"
                          className="row-checkbox"
                          checked={p.status === "Finished"}
                          onChange={() => toggleProjectFinished(p)}
                        />
                      </label>
                      <span
                        className={
                          "status-label " +
                          (p.status === "Finished" ? "status-label-finished" : "status-label-unfinished")
                        }
                      >
                        {displayStatus(p.status)}
                      </span>
                      <button
                        type="button"
                        className="tile-delete-btn"
                        title="Delete"
                        onClick={(e) => {
                          e.stopPropagation();
                                                      archiveProject(p.id);
                        }}
                      >
                        <Trash2 size={14} strokeWidth={1.9} />
                      </button>
                    </div>

                    <span className="project-tile-name">{p.name}</span>
                    <span className="project-tile-company">{companyName(p.companyId)}</span>
                    <span className="project-tile-chips">
                      <span className="chip chip-lg">SRL {p.srlLevel}</span>
                      <PotentialChip level={p.potentialLevel} large />
                      <span className="chip chip-lg chip-funding">${formatMoney(p.funding)}</span>
                    </span>
                  </div>
                ))}

                {projects.length > 0 && visibleProjects.length === 0 ? (
                  <p className="empty-table-note">No projects match your search.</p>
                ) : null}
              </div>
            </div>
          )}

          {view === "unfinishedProjects" && (
            <div className="page-content">
              <div className="page-content-header">
                <button type="button" className="btn-back" onClick={() => goTo("projects")}>
                  <ArrowLeft size={15} strokeWidth={1.75} />
                  Back to all projects
                </button>
                <h1 className="page-title">Unfinished projects</h1>
                <p className="page-count">
                  {unfinishedProjectsList.length} {unfinishedProjectsList.length === 1 ? "project" : "projects"}
                </p>
              </div>

              <div className="page-body">
                {unfinishedProjectsList.length === 0 ? (
                  <p className="empty-table-note">No unfinished projects right now.</p>
                ) : (
                  unfinishedProjectsList.map((p) => (
                    <div
                      role="button"
                      tabIndex={0}
                      key={p.id}
                      className="add-tile project-tile"
                      onClick={() => openProject(p.id, "unfinishedProjects")}
                      onKeyDown={(e) => {
                        if (e.key === "Enter" || e.key === " ") {
                          e.preventDefault();
                          openProject(p.id, "unfinishedProjects");
                        }
                      }}
                    >
                      <div className="project-tile-status-row">
                        <label
                          className="project-tile-status-check"
                          onClick={(e) => e.stopPropagation()}
                        >
                          <input
                            type="checkbox"
                            className="row-checkbox"
                            checked={p.status === "Finished"}
                            onChange={() => toggleProjectFinished(p)}
                          />
                        </label>
                        <span className="status-label status-label-unfinished">{displayStatus(p.status)}</span>
                        <button
                          type="button"
                          className="tile-delete-btn"
                          title="Delete"
                          onClick={(e) => {
                            e.stopPropagation();
                                                          archiveProject(p.id);
                          }}
                        >
                          <Trash2 size={14} strokeWidth={1.9} />
                        </button>
                      </div>

                      <span className="project-tile-name">{p.name}</span>
                      <span className="project-tile-company">{companyName(p.companyId)}</span>
                      <span className="project-tile-chips">
                        <span className="chip chip-lg">SRL {p.srlLevel}</span>
                        <PotentialChip level={p.potentialLevel} large />
                        <span className="chip chip-lg chip-funding">${formatMoney(p.funding)}</span>
                      </span>
                    </div>
                  ))
                )}
              </div>
            </div>
          )}

          {view === "finishedProjects" && (
            <div className="page-content">
              <div className="page-content-header">
                <button type="button" className="btn-back" onClick={() => goTo("projects")}>
                  <ArrowLeft size={15} strokeWidth={1.75} />
                  Back to all projects
                </button>
                <h1 className="page-title">Finalized projects</h1>
                <p className="page-count">
                  {finishedProjectsList.length} {finishedProjectsList.length === 1 ? "project" : "projects"}
                </p>
              </div>

              <div className="page-body">
                {finishedProjectsList.length === 0 ? (
                  <p className="empty-table-note">No finalized projects yet.</p>
                ) : (
                  finishedProjectsList.map((p) => (
                    <div
                      role="button"
                      tabIndex={0}
                      key={p.id}
                      className="add-tile project-tile"
                      onClick={() => openProject(p.id, "finishedProjects")}
                      onKeyDown={(e) => {
                        if (e.key === "Enter" || e.key === " ") {
                          e.preventDefault();
                          openProject(p.id, "finishedProjects");
                        }
                      }}
                    >
                      <div className="project-tile-status-row">
                        <label
                          className="project-tile-status-check"
                          onClick={(e) => e.stopPropagation()}
                        >
                          <input
                            type="checkbox"
                            className="row-checkbox"
                            checked={p.status === "Finished"}
                            onChange={() => toggleProjectFinished(p)}
                          />
                        </label>
                        <span className="status-label status-label-finished">{displayStatus(p.status)}</span>
                        <button
                          type="button"
                          className="tile-delete-btn"
                          title="Delete"
                          onClick={(e) => {
                            e.stopPropagation();
                                                          archiveProject(p.id);
                          }}
                        >
                          <Trash2 size={14} strokeWidth={1.9} />
                        </button>
                      </div>

                      <span className="project-tile-name">{p.name}</span>
                      <span className="project-tile-company">{companyName(p.companyId)}</span>
                      <span className="project-tile-chips">
                        <span className="chip chip-lg">SRL {p.srlLevel}</span>
                        <PotentialChip level={p.potentialLevel} large />
                        <span className="chip chip-lg chip-funding">${formatMoney(p.funding)}</span>
                      </span>
                    </div>
                  ))
                )}
              </div>
            </div>
          )}

          {view === "archivedProjects" && (
            <div className="page-content">
              <div className="page-content-header">
                <button type="button" className="btn-back" onClick={() => goTo("projects")}>
                  <ArrowLeft size={15} strokeWidth={1.75} />
                  Back to all projects
                </button>
                <h1 className="page-title">Archived projects</h1>
                <p className="page-count">
                  {archivedProjectsList.length} {archivedProjectsList.length === 1 ? "project" : "projects"}
                </p>
              </div>
              {archivedProjectsList.length > 0 ? (
                <div className="button-row">
                  <button
                    type="button"
                    className="btn-delete"
                    onClick={() => {
                      if (confirmForeverId === "all-projects") {
                        deleteAllArchivedProjects();
                        setConfirmForeverId(null);
                      } else {
                        setConfirmForeverId("all-projects");
                      }
                    }}
                  >
                    {confirmForeverId === "all-projects" ? "Click again to delete all" : "Delete all"}
                  </button>
                </div>
              ) : null}

              <div className="page-body">

                {archivedProjectsList.length === 0 ? (
                  <p className="empty-table-note">No archived projects.</p>
                ) : (
                  archivedProjectsList.map((p) => (
                    <div key={p.id} className="add-tile project-tile">
                      <span className="project-tile-name">{p.name}</span>
                      <span className="project-tile-company">{companyName(p.companyId)}</span>
                      <span className="project-tile-chips">
                        <span className="chip chip-lg">SRL {p.srlLevel}</span>
                        <span className="chip chip-lg chip-funding">${formatMoney(p.funding)}</span>
                      </span>
                      <div className="form-actions">
                        <button type="button" className="btn-secondary" onClick={() => restoreProject(p.id)}>
                          Unarchive
                        </button>
                        <button
                          type="button"
                          className="btn-delete"
                          onClick={() => {
                            if (confirmForeverId === p.id) {
                              deleteProjectForever(p.id);
                              setConfirmForeverId(null);
                            } else {
                              setConfirmForeverId(p.id);
                            }
                          }}
                        >
                          {confirmForeverId === p.id ? "Click again to confirm" : "Delete forever"}
                        </button>
                      </div>
                    </div>
                  ))
                )}
              </div>
            </div>
          )}

          {view === "companies" && (
            <div className="page-content">
              <div className="page-content-header">
                <button type="button" className="btn-back" onClick={() => goTo("dashboard")}>
                  <ArrowLeft size={15} strokeWidth={1.75} />
                  Back to home
                </button>
                <h1 className="page-title">Companies</h1>
                <p className="page-count">
                  {visibleCompanies.length} {visibleCompanies.length === 1 ? "company" : "companies"}
                </p>
              </div>

              <div className="button-row">
                <button type="button" className="btn-choice" onClick={() => goTo("archivedCompanies")}>
                  Archive{archivedCompaniesList.length ? " (" + archivedCompaniesList.length + ")" : ""}
                </button>
              </div>

              <div className="list-toolbar">
                <div className="search-wrap">
                  <input
                    type="text"
                    className="search-input"
                    placeholder="Search by company name"
                    value={companySearch}
                    onChange={(e) => setCompanySearch(e.target.value)}
                  />
                </div>
                <div className="sort-select-wrap">
                  <SelectInput
                    id="company-sort"
                    value={companySort}
                    onChange={(e) => setCompanySort(e.target.value)}
                  >
                    <option value="name-asc">Name (A → Z)</option>
                    <option value="name-desc">Name (Z → A)</option>
                    <option value="projects-desc">Most projects</option>
                    <option value="projects-asc">Least projects</option>
                  </SelectInput>
                </div>
              </div>

              <div className="page-body page-body-companies">
                <button type="button" className="add-tile" onClick={() => goTo("addCompany")}>
                  <span className="add-tile-icon">
                    <Plus size={20} strokeWidth={2} />
                  </span>
                  <span className="add-tile-label">Add company</span>
                </button>

                {visibleCompanies.map((c) => (
                  <div
                    role="button"
                    tabIndex={0}
                    key={c.id}
                    className="add-tile project-tile"
                    onClick={() => openCompany(c.id)}
                    onKeyDown={(e) => {
                      if (e.key === "Enter" || e.key === " ") {
                        e.preventDefault();
                        openCompany(c.id);
                      }
                    }}
                  >
                    <div className="project-tile-status-row">
                      <span className="project-tile-name">{c.name}</span>
                      <button
                        type="button"
                        className="tile-delete-btn"
                        title="Delete"
                        onClick={(e) => {
                          e.stopPropagation();
                                                      archiveCompany(c.id);
                        }}
                      >
                        <Trash2 size={14} strokeWidth={1.9} />
                      </button>
                    </div>
                    <span className="project-tile-company">
                      {c.about
                        ? c.about.length > 90
                          ? c.about.slice(0, 90) + "…"
                          : c.about
                        : "No description yet."}
                    </span>

                    {(() => {
                      const poc = contacts
                        .filter((ct) => ct.companyId === c.id && !ct.archived)
                        .sort((a, b) => (a.createdAt || 0) - (b.createdAt || 0))[0];
                      return (
                        <div className="poc-block">
                          <span className="poc-label">Point of Contact</span>
                          {poc ? (
                            <>
                              <span className="poc-name">{poc.contactName || "—"}</span>
                              {poc.jobTitle ? <span className="poc-title">{poc.jobTitle}</span> : null}
                              <span className="poc-detail">
                                <Mail size={12} strokeWidth={1.75} />
                                {poc.contactEmail || "—"}
                              </span>
                              <span className="poc-detail">
                                <Phone size={12} strokeWidth={1.75} />
                                {poc.contactNumber || "—"}
                              </span>
                            </>
                          ) : (
                            <span className="poc-empty">No contact on file</span>
                          )}
                        </div>
                      );
                    })()}
                  </div>
                ))}

                {companies.length > 0 && visibleCompanies.length === 0 ? (
                  <p className="empty-table-note">No companies match your search.</p>
                ) : null}
              </div>
            </div>
          )}

          {view === "archivedCompanies" && (
            <div className="page-content">
              <div className="page-content-header">
                <button type="button" className="btn-back" onClick={() => goTo("companies")}>
                  <ArrowLeft size={15} strokeWidth={1.75} />
                  Back to companies
                </button>
                <h1 className="page-title">Archived companies</h1>
                <p className="page-count">
                  {archivedCompaniesList.length} {archivedCompaniesList.length === 1 ? "company" : "companies"}
                </p>
              </div>
              {archivedCompaniesList.length > 0 ? (
                <div className="button-row">
                  <button
                    type="button"
                    className="btn-delete"
                    onClick={() => {
                      if (confirmForeverId === "all-companies") {
                        deleteAllArchivedCompanies();
                        setConfirmForeverId(null);
                      } else {
                        setConfirmForeverId("all-companies");
                      }
                    }}
                  >
                    {confirmForeverId === "all-companies" ? "Click again to delete all" : "Delete all"}
                  </button>
                </div>
              ) : null}

              <div className="page-body page-body-companies">

                {archivedCompaniesList.length === 0 ? (
                  <p className="empty-table-note">No archived companies.</p>
                ) : (
                  archivedCompaniesList.map((c) => (
                    <div key={c.id} className="add-tile project-tile">
                      <span className="project-tile-name">{c.name}</span>
                      <span className="project-tile-company">{c.about || "No description yet."}</span>
                      <div className="form-actions">
                        <button type="button" className="btn-secondary" onClick={() => restoreCompany(c.id)}>
                          Unarchive
                        </button>
                        <button
                          type="button"
                          className="btn-delete"
                          onClick={() => {
                            if (confirmForeverId === c.id) {
                              deleteCompanyForever(c.id);
                              setConfirmForeverId(null);
                            } else {
                              setConfirmForeverId(c.id);
                            }
                          }}
                        >
                          {confirmForeverId === c.id ? "Click again to confirm" : "Delete forever"}
                        </button>
                      </div>
                    </div>
                  ))
                )}
              </div>
            </div>
          )}

          {view === "contacts" && (
            <div className="page-content">
              <div className="page-content-header">
                <button type="button" className="btn-back" onClick={() => goTo("dashboard")}>
                  <ArrowLeft size={15} strokeWidth={1.75} />
                  Back to home
                </button>
                <h1 className="page-title">All contacts</h1>
                <p className="page-count">
                  {visibleContacts.length} {visibleContacts.length === 1 ? "contact" : "contacts"}
                </p>
              </div>

              <div className="button-row">
                <button type="button" className="btn-choice" onClick={() => goTo("archivedContacts")}>
                  Archive{archivedContactsList.length ? " (" + archivedContactsList.length + ")" : ""}
                </button>
              </div>

              <div className="list-toolbar">
                <div className="search-wrap">
                  <input
                    type="text"
                    className="search-input"
                    placeholder="Search by name, company, job title, email, or phone"
                    value={contactSearch}
                    onChange={(e) => setContactSearch(e.target.value)}
                  />
                </div>
                <div className="sort-select-wrap">
                  <SelectInput id="contact-sort" value={contactSort} onChange={(e) => setContactSort(e.target.value)}>
                    <option value="name-asc">Name (A → Z)</option>
                    <option value="name-desc">Name (Z → A)</option>
                    <option value="company-asc">Company (A → Z)</option>
                  </SelectInput>
                </div>
                <button type="button" className="btn-import" onClick={handleExportContacts}>
                  <Download size={14} strokeWidth={1.75} />
                  Export contacts
                </button>
                <button type="button" className="btn-add-contact" onClick={() => goTo("addContact")}>
                  <Plus size={16} strokeWidth={2} />
                  Add contact
                </button>
              </div>

              <div className="page-body">
                {contacts.length === 0 ? (
                  <p className="empty-table-note">
                    No contacts on file yet. Import a contact list from Excel to get started.
                  </p>
                ) : visibleContacts.length === 0 ? (
                  <p className="empty-table-note">No contacts match your search.</p>
                ) : (
                  <div className="table-wrap">
                    <table className="data-table">
                      <thead>
                        <tr>
                          <th>Contact name</th>
                          <th>Job title</th>
                          <th>Company</th>
                          <th>Email</th>
                          <th>Phone</th>
                          <th></th>
                        </tr>
                      </thead>
                      <tbody>
                        {visibleContacts.map((c) => (
                          <tr key={c.id} className="data-row" onClick={() => openContact(c.id, "contacts")}>
                            <td>{c.contactName || "—"}</td>
                            <td>{c.jobTitle || "—"}</td>
                            <td>{companyName(c.companyId)}</td>
                            <td>{c.contactEmail || "—"}</td>
                            <td>{c.contactNumber ? formatPhoneDisplay(c.contactNumber) : "—"}</td>
                            <td>
                              <div className="row-action-group">
                                <button
                                  type="button"
                                  className="tile-edit-btn"
                                  title="Edit"
                                  onClick={(e) => {
                                    e.stopPropagation();
                                    openEditContact(c.id, "contacts");
                                  }}
                                >
                                  <Pencil size={14} strokeWidth={1.9} />
                                </button>
                                <button
                                  type="button"
                                  className="tile-delete-btn"
                                  title="Delete"
                                  onClick={(e) => {
                                    e.stopPropagation();
                                    archiveContact(c.id);
                                  }}
                                >
                                  <Trash2 size={14} strokeWidth={1.9} />
                                </button>
                              </div>
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                )}
              </div>
            </div>
          )}

          {view === "archivedContacts" && (
            <div className="page-content">
              <div className="page-content-header">
                <button type="button" className="btn-back" onClick={() => goTo("contacts")}>
                  <ArrowLeft size={15} strokeWidth={1.75} />
                  Back to all contacts
                </button>
                <h1 className="page-title">Archived contacts</h1>
                <p className="page-count">
                  {archivedContactsList.length} {archivedContactsList.length === 1 ? "contact" : "contacts"}
                </p>
              </div>
              {archivedContactsList.length > 0 ? (
                <div className="button-row">
                  <button
                    type="button"
                    className="btn-delete"
                    onClick={() => {
                      if (confirmForeverId === "all-contacts") {
                        deleteAllArchivedContacts();
                        setConfirmForeverId(null);
                      } else {
                        setConfirmForeverId("all-contacts");
                      }
                    }}
                  >
                    {confirmForeverId === "all-contacts" ? "Click again to delete all" : "Delete all"}
                  </button>
                </div>
              ) : null}

              <div className="page-body">

                {archivedContactsList.length === 0 ? (
                  <p className="empty-table-note">No archived contacts.</p>
                ) : (
                  <div className="table-wrap">
                    <table className="data-table">
                      <thead>
                        <tr>
                          <th>Contact name</th>
                          <th>Company</th>
                          <th>Email</th>
                          <th></th>
                        </tr>
                      </thead>
                      <tbody>
                        {archivedContactsList.map((c) => (
                          <tr key={c.id} className="data-row">
                            <td>{c.contactName || "—"}</td>
                            <td>{companyName(c.companyId)}</td>
                            <td>{c.contactEmail || "—"}</td>
                            <td>
                              <div className="form-actions">
                                <button type="button" className="btn-secondary" onClick={() => restoreContact(c.id)}>
                                  Unarchive
                                </button>
                                <button
                                  type="button"
                                  className="btn-delete"
                                  onClick={() => {
                                    if (confirmForeverId === c.id) {
                                      deleteContactForever(c.id);
                                      setConfirmForeverId(null);
                                    } else {
                                      setConfirmForeverId(c.id);
                                    }
                                  }}
                                >
                                  {confirmForeverId === c.id ? "Click again to confirm" : "Delete forever"}
                                </button>
                              </div>
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                )}
              </div>
            </div>
          )}

          {view === "generateReport" && (
            <div className="page-content">
              <div className="page-content-header">
                <button type="button" className="btn-back" onClick={() => goTo("dashboard")}>
                  <ArrowLeft size={15} strokeWidth={1.75} />
                  Back to home
                </button>
                <h1 className="page-title">Generate Report</h1>
                <p className="page-count">Choose what to include, then generate a Word document.</p>
              </div>

              <div className="report-builder">
                <div className="report-section-block">
                  <div className="report-section-block-header">
                    <h2 className="form-section-title">Include sections</h2>
                    <button type="button" className="home-list-viewall" onClick={toggleAllReportSections}>
                      {rpSections.size === REPORT_SECTIONS.length ? "Clear all" : "Select all"}
                    </button>
                  </div>
                  <div className="report-checkbox-grid">
                    {REPORT_SECTIONS.map((s) => (
                      <label key={s.key} className="report-checkbox">
                        <input
                          type="checkbox"
                          checked={rpSections.has(s.key)}
                          onChange={() => toggleReportSection(s.key)}
                        />
                        <span>{s.label}</span>
                      </label>
                    ))}
                  </div>
                </div>

                {rpSections.has("srl") && (
                  <div className="report-section-block">
                    <div className="report-section-block-header">
                      <h2 className="form-section-title">SRLs to include</h2>
                      <button type="button" className="home-list-viewall" onClick={toggleAllReportSrlLevels}>
                        {rpSrlLevels.size === SRL_LEVELS.length ? "Clear all" : "Select all"}
                      </button>
                    </div>
                    <p className="form-section-subtitle">Leave none selected to include every level.</p>
                    <div className="report-checkbox-grid report-checkbox-grid-levels">
                      {SRL_LEVELS.map((lvl) => (
                        <label key={lvl} className="report-checkbox">
                          <input
                            type="checkbox"
                            checked={rpSrlLevels.has(lvl)}
                            onChange={() => toggleReportSrlLevel(lvl)}
                          />
                          <span>SRL {lvl}</span>
                        </label>
                      ))}
                    </div>
                  </div>
                )}

                {rpSections.has("company") && (
                  <div className="report-section-block">
                    <div className="report-section-block-header">
                      <h2 className="form-section-title">Companies to include</h2>
                      <button type="button" className="home-list-viewall" onClick={toggleAllReportCompanies}>
                        {rpCompanyIds.size === companies.filter((c) => !c.archived).length ? "Clear all" : "Select all"}
                      </button>
                    </div>
                    <p className="form-section-subtitle">Leave none selected to include every company.</p>
                    {companies.filter((c) => !c.archived).length === 0 ? (
                      <p className="empty-table-note">No companies yet.</p>
                    ) : (
                      <div className="report-checkbox-grid">
                        {companies
                          .filter((c) => !c.archived)
                          .map((c) => (
                            <label key={c.id} className="report-checkbox">
                              <input
                                type="checkbox"
                                checked={rpCompanyIds.has(c.id)}
                                onChange={() => toggleReportCompany(c.id)}
                              />
                              <span>{c.name}</span>
                            </label>
                          ))}
                      </div>
                    )}
                  </div>
                )}

                {rpSections.has("lead") && (
                  <div className="report-section-block">
                    <div className="report-section-block-header">
                      <h2 className="form-section-title">PIs to include</h2>
                      <button type="button" className="home-list-viewall" onClick={toggleAllReportLeads}>
                        {rpLeads.size === reportLeadOptions.length ? "Clear all" : "Select all"}
                      </button>
                    </div>
                    <p className="form-section-subtitle">Leave none selected to include every PI.</p>
                    {reportLeadOptions.length === 0 ? (
                      <p className="empty-table-note">No PIs assigned to any project yet.</p>
                    ) : (
                      <div className="report-checkbox-grid">
                        {reportLeadOptions.map((g) => (
                          <label key={g.key} className="report-checkbox">
                            <input
                              type="checkbox"
                              checked={rpLeads.has(g.key)}
                              onChange={() => toggleReportLead(g.key)}
                            />
                            <span>{g.displayName}</span>
                          </label>
                        ))}
                      </div>
                    )}
                  </div>
                )}

                {rpError ? <div className="form-error">{rpError}</div> : null}

                <div className="form-actions">
                  <button
                    type="button"
                    className="btn-primary"
                    disabled={rpGenerating}
                    onClick={handleGenerateReport}
                  >
                    {rpGenerating ? "Generating…" : "Generate Report"}
                  </button>
                </div>
              </div>
            </div>
          )}

          {view === "viewContact" && (
            <div className="detail-page">
              <button
                type="button"
                className="btn-back"
                onClick={() => goTo(contactReturnView)}
              >
                <ArrowLeft size={15} strokeWidth={1.75} />
                {contactReturnView === "viewCompany" ? "Back to company" : "Back to all contacts"}
              </button>

              {selectedContact ? (
                <>
                  <h1 className="page-title">{selectedContact.contactName || "Unnamed contact"}</h1>
                  <p className="page-subtitle page-subtitle-tight">
                    {selectedContact.jobTitle || "No job title on file"}
                  </p>
                  <p className="page-meta">{companyName(selectedContact.companyId)}</p>

                  <div className="detail-grid">
                    <div className="detail-row">
                      <span className="detail-label">Company</span>
                      <span className="detail-value">{companyName(selectedContact.companyId)}</span>
                    </div>
                    <div className="detail-row">
                      <span className="detail-label">Job title</span>
                      <span className="detail-value">{selectedContact.jobTitle || "—"}</span>
                    </div>
                    <div className="detail-row">
                      <span className="detail-label">Email</span>
                      <span className="detail-value">{selectedContact.contactEmail || "—"}</span>
                    </div>
                    <div className="detail-row">
                      <span className="detail-label">Phone</span>
                      <span className="detail-value">
                        {selectedContact.contactNumber ? formatPhoneDisplay(selectedContact.contactNumber) : "—"}
                      </span>
                    </div>
                  </div>

                  <div className="form-actions">
                    {selectedContact.companyId ? (
                      <button
                        type="button"
                        className="btn-secondary"
                        onClick={() => openCompany(selectedContact.companyId, "viewContact")}
                      >
                        View company
                      </button>
                    ) : null}
                    <button type="button" className="btn-edit" onClick={startEditContact}>
                      <Pencil size={14} strokeWidth={1.9} />
                      Edit
                    </button>
                    <button
                      type="button"
                      className="btn-delete"
                      onClick={() => {
                        archiveContact(selectedContact.id);
                        goTo(contactReturnView);
                      }}
                    >
                      Delete
                    </button>
                  </div>
                </>
              ) : (
                <p className="empty-state-text">Contact not found.</p>
              )}
            </div>
          )}

          {view === "addContact" && (
            <div className="detail-page">
              <button type="button" className="btn-back" onClick={() => goTo("contacts")}>
                <ArrowLeft size={15} strokeWidth={1.75} />
                Back to all contacts
              </button>
              <h1 className="page-title">Add contact</h1>
              <p className="page-subtitle">Enter the contact's details below.</p>

              <div className="form-grid">
                <Field label="Contact name" id="nc-contact-name" error={ncErrors.contactName} full>
                  <TextInput
                    id="nc-contact-name"
                    error={ncErrors.contactName}
                    value={ncContactName}
                    onChange={(e) => setNcContactName(e.target.value)}
                  />
                </Field>

                <Field label="Company" id="nc-company" full>
                  <div className="company-mode-toggle">
                    <button
                      type="button"
                      className={"company-mode-btn" + (ncCompanyMode === "select" ? " company-mode-btn-active" : "")}
                      onClick={() => switchNcCompanyMode("select")}
                    >
                      Select existing company
                    </button>
                    <button
                      type="button"
                      className={"company-mode-btn" + (ncCompanyMode === "create" ? " company-mode-btn-active" : "")}
                      onClick={() => switchNcCompanyMode("create")}
                    >
                      Create new company
                    </button>
                  </div>

                  {ncCompanyMode === "select" ? (
                    <>
                      <CompanyCombo
                        id="nc-company"
                        error={ncErrors.companyId}
                        value={ncCompanySearch}
                        companies={companies}
                        onTextChange={handleNcCompanyTextChange}
                        onSelectCompany={handleNcSelectCompany}
                      />
                      {ncErrors.companyId ? <div className="error-text">{ncErrors.companyId}</div> : null}
                      {companies.length === 0 ? (
                        <p className="field-hint">Or switch to "Create new company" above.</p>
                      ) : null}
                    </>
                  ) : (
                    <div className="new-company-fields">
                      <Field label="New company name" id="nc-new-company-name" error={ncErrors.newCompanyName}>
                        <TextInput
                          id="nc-new-company-name"
                          error={ncErrors.newCompanyName}
                          value={ncNewCompanyName}
                          onChange={(e) => setNcNewCompanyName(e.target.value)}
                        />
                      </Field>
                      <Field label="Industry (optional)" id="nc-new-company-industry">
                        <TextInput
                          id="nc-new-company-industry"
                          value={ncNewCompanyIndustry}
                          onChange={(e) => setNcNewCompanyIndustry(e.target.value)}
                        />
                      </Field>
                      <Field label="HQ Location (optional)" id="nc-new-company-hq">
                        <TextInput
                          id="nc-new-company-hq"
                          value={ncNewCompanyHqLocation}
                          onChange={(e) => setNcNewCompanyHqLocation(e.target.value)}
                        />
                      </Field>
                      <Field label="SBU Location (optional)" id="nc-new-company-sbu">
                        <TextInput
                          id="nc-new-company-sbu"
                          value={ncNewCompanySbuLocation}
                          onChange={(e) => setNcNewCompanySbuLocation(e.target.value)}
                        />
                      </Field>
                      <p className="field-hint">This will create a new company along with the contact.</p>
                    </div>
                  )}
                </Field>

                <Field label="Job title" id="nc-contact-title" error={ncErrors.contactTitle}>
                  <TextInput
                    id="nc-contact-title"
                    error={ncErrors.contactTitle}
                    value={ncContactTitle}
                    onChange={(e) => setNcContactTitle(e.target.value)}
                  />
                </Field>

                <Field label="Contact email" id="nc-contact-email" error={ncErrors.contactEmail}>
                  <TextInput
                    id="nc-contact-email"
                    type="email"
                    error={ncErrors.contactEmail}
                    value={ncContactEmail}
                    onChange={(e) => setNcContactEmail(e.target.value)}
                  />
                </Field>

                <Field label="Contact number" id="nc-contact-number" error={ncErrors.contactNumber}>
                  <TextInput
                    id="nc-contact-number"
                    icon={Phone}
                    type="tel"
                    error={ncErrors.contactNumber}
                    value={ncContactNumber}
                    onChange={(e) => setNcContactNumber(e.target.value)}
                  />
                </Field>
              </div>

              <div className="form-actions">
                <button type="button" className="btn-primary" onClick={handleSaveContact}>Save</button>
                <button type="button" className="btn-secondary" onClick={handleDiscardContact}>Discard</button>
              </div>
            </div>
          )}

          {view === "editContact" && selectedContact && (
            <div className="detail-page">
              <button type="button" className="btn-back" onClick={handleCancelEditContact}>
                <ArrowLeft size={15} strokeWidth={1.75} />
                Back to contact
              </button>
              <h1 className="page-title">Edit contact</h1>
              <p className="page-subtitle">Update the contact's details below.</p>

              <div className="form-grid">
                <Field label="Contact name" id="edc-contact-name" error={edcErrors.contactName} full>
                  <TextInput
                    id="edc-contact-name"
                    error={edcErrors.contactName}
                    value={edcContactName}
                    onChange={(e) => setEdcContactName(e.target.value)}
                  />
                </Field>

                <Field label="Company" id="edc-company" error={edcErrors.companyId} full>
                  <SelectInput
                    id="edc-company"
                    error={edcErrors.companyId}
                    value={edcCompanyId}
                    onChange={(e) => setEdcCompanyId(e.target.value)}
                  >
                    <option value="">{companies.length ? "Select company" : "No companies yet"}</option>
                    {[...companies]
                      .sort((a, b) => a.name.localeCompare(b.name))
                      .map((c) => (
                        <option key={c.id} value={c.id}>{c.name}</option>
                      ))}
                  </SelectInput>
                </Field>

                <Field label="Job title" id="edc-contact-title" error={edcErrors.contactTitle}>
                  <TextInput
                    id="edc-contact-title"
                    error={edcErrors.contactTitle}
                    value={edcContactTitle}
                    onChange={(e) => setEdcContactTitle(e.target.value)}
                  />
                </Field>

                <Field label="Contact email" id="edc-contact-email" error={edcErrors.contactEmail}>
                  <TextInput
                    id="edc-contact-email"
                    type="email"
                    error={edcErrors.contactEmail}
                    value={edcContactEmail}
                    onChange={(e) => setEdcContactEmail(e.target.value)}
                  />
                </Field>

                <Field label="Contact number" id="edc-contact-number" error={edcErrors.contactNumber}>
                  <TextInput
                    id="edc-contact-number"
                    icon={Phone}
                    type="tel"
                    error={edcErrors.contactNumber}
                    value={edcContactNumber}
                    onChange={(e) => setEdcContactNumber(e.target.value)}
                  />
                </Field>
              </div>

              <div className="form-actions">
                <button type="button" className="btn-primary" onClick={handleSaveEditContact}>
                  Save changes
                </button>
                <button type="button" className="btn-secondary" onClick={handleCancelEditContact}>
                  Cancel
                </button>
              </div>
            </div>
          )}

          {view === "addProject" && (
            <div className="detail-page">
              <button type="button" className="btn-back" onClick={() => goTo("projects")}>
                <ArrowLeft size={15} strokeWidth={1.75} />
                Back to all projects
              </button>
              <h1 className="page-title">Add project</h1>
              <p className="page-subtitle">Enter the project details below.</p>

              <div className="form-grid">
                <Field label="Project name" id="ap-name" error={apErrors.name} full>
                  <TextInput
                    id="ap-name"
                    error={apErrors.name}
                    value={apName}
                    onChange={(e) => setApName(e.target.value)}
                  />
                </Field>

                <Field label="Company" id="ap-company" full>
                  <div className="company-mode-toggle">
                    <button
                      type="button"
                      className={"company-mode-btn" + (apCompanyMode === "select" ? " company-mode-btn-active" : "")}
                      onClick={() => switchApCompanyMode("select")}
                    >
                      Select existing company
                    </button>
                    <button
                      type="button"
                      className={"company-mode-btn" + (apCompanyMode === "create" ? " company-mode-btn-active" : "")}
                      onClick={() => switchApCompanyMode("create")}
                    >
                      Create new company
                    </button>
                  </div>

                  {apCompanyMode === "select" ? (
                    <>
                      <CompanyCombo
                        id="ap-company"
                        error={apErrors.companyId}
                        value={apCompanySearch}
                        companies={companies}
                        onTextChange={handleApCompanyTextChange}
                        onSelectCompany={handleApSelectCompany}
                      />
                      {apErrors.companyId ? <div className="error-text">{apErrors.companyId}</div> : null}
                      {companies.length === 0 ? (
                        <p className="field-hint">Or switch to "Create new company" above.</p>
                      ) : null}
                    </>
                  ) : (
                    <div className="new-company-fields">
                      <Field label="New company name" id="ap-new-company-name" error={apErrors.newCompanyName}>
                        <TextInput
                          id="ap-new-company-name"
                          error={apErrors.newCompanyName}
                          value={apNewCompanyName}
                          onChange={(e) => setApNewCompanyName(e.target.value)}
                        />
                      </Field>
                      <Field label="Industry (optional)" id="ap-new-company-industry">
                        <TextInput
                          id="ap-new-company-industry"
                          value={apNewCompanyIndustry}
                          onChange={(e) => setApNewCompanyIndustry(e.target.value)}
                        />
                      </Field>
                      <Field label="HQ Location (optional)" id="ap-new-company-hq">
                        <TextInput
                          id="ap-new-company-hq"
                          value={apNewCompanyHqLocation}
                          onChange={(e) => setApNewCompanyHqLocation(e.target.value)}
                        />
                      </Field>
                      <Field label="SBU Location (optional)" id="ap-new-company-sbu">
                        <TextInput
                          id="ap-new-company-sbu"
                          value={apNewCompanySbuLocation}
                          onChange={(e) => setApNewCompanySbuLocation(e.target.value)}
                        />
                      </Field>
                      <p className="field-hint">This will create a new company along with the project.</p>
                    </div>
                  )}
                </Field>

                <Field label="RFS/NTI number (optional)" id="ap-rfs-nti" full>
                  <TextInput
                    id="ap-rfs-nti"
                    placeholder="#12345"
                    value={apRfsNti}
                    onChange={(e) => setApRfsNti(maskHashPrefix(e.target.value))}
                  />
                </Field>

                <Field label="PI name" id="ap-lead" error={apErrors.lead}>
                  <TextInput
                    id="ap-lead"
                    error={apErrors.lead}
                    value={apLead}
                    onChange={(e) => setApLead(e.target.value)}
                  />
                </Field>

                <Field label="Funding (enter 0 if not decided)" id="ap-funding" error={apErrors.funding}>
                  <TextInput
                    id="ap-funding"
                    prefix="$"
                    inputMode="decimal"
                    placeholder="0.00"
                    error={apErrors.funding}
                    value={apFunding}
                    onChange={(e) => setApFunding(maskFundingInput(e.target.value))}
                  />
                </Field>

                <Field label="Current SRL" id="ap-srl" error={apErrors.srl}>
                  <SelectInput
                    id="ap-srl"
                    error={apErrors.srl}
                    value={apSrl}
                    onChange={(e) => setApSrl(e.target.value)}
                  >
                    <option value="">Select level</option>
                    {SRL_LEVELS.map((lvl) => (
                      <option key={lvl} value={lvl}>{lvl}</option>
                    ))}
                  </SelectInput>
                </Field>

                <Field
                  label={"Date SRL " + (apSrl || "—") + " Reached (Optional)"}
                  id="ap-srl-date"
                  error={apErrors.srlDate}
                >
                  <TextInput
                    id="ap-srl-date"
                    placeholder="DD/MM/YYYY"
                    error={apErrors.srlDate}
                    value={apSrlDate}
                    onChange={(e) => setApSrlDate(maskDateInput(e.target.value))}
                  />
                  <p className="field-hint">
                    Leave blank to use today's date. This updates the SRL progression log, separately from the project's start date.
                  </p>
                </Field>

                <Field label="Potential (optional)" id="ap-potential" error={apErrors.potential}>
                  <SelectInput
                    id="ap-potential"
                    error={apErrors.potential}
                    value={apPotential}
                    onChange={(e) => setApPotential(e.target.value)}
                  >
                    <option value="">Select potential</option>
                    {POTENTIAL_LEVELS.map((lvl) => (
                      <option key={lvl} value={lvl}>{lvl}</option>
                    ))}
                  </SelectInput>
                </Field>

                <div className="form-grid-pair field-full">
                  <Field label="Project Start Date" id="ap-start-date" error={apErrors.startDate}>
                    <TextInput
                      id="ap-start-date"
                      placeholder="DD/MM/YYYY"
                      error={apErrors.startDate}
                      value={apStartDate}
                      onChange={(e) => setApStartDate(maskDateInput(e.target.value))}
                    />
                  </Field>

                  <Field label="Project End Date (Optional)" id="ap-deadline" error={apErrors.deadline}>
                    <TextInput
                      id="ap-deadline"
                      placeholder="DD/MM/YYYY"
                      error={apErrors.deadline}
                      value={apDeadline}
                      onChange={(e) => setApDeadline(maskDateInput(e.target.value))}
                    />
                    {(() => {
                      const days = daysBetween(apStartDate, apDeadline);
                      if (days === null) return null;
                      return (
                        <p className="field-hint">
                          {days < 0
                            ? "Deadline is before the start date."
                            : days === 0
                            ? "Same-day deadline."
                            : days + (days === 1 ? " day" : " days") + " from start to deadline."}
                        </p>
                      );
                    })()}
                  </Field>
                </div>
              </div>

              <div className="form-section">
                <h2 className="form-section-title">Contact details</h2>
                <p className="form-section-subtitle">Who should we reach out to about this project? All fields optional.</p>

                {apContacts.map((row, idx) => (
                  <div className="form-grid" key={row.key} style={{ marginBottom: 16 }}>
                    <Field
                      label="Contact name"
                      id={`ap-contact-name-${row.key}`}
                      error={idx === 0 ? apErrors.contactName : apErrors.contacts?.[row.key]?.name}
                    >
                      <ContactNameCombo
                        id={`ap-contact-name-${row.key}`}
                        error={idx === 0 ? apErrors.contactName : apErrors.contacts?.[row.key]?.name}
                        value={row.name}
                        contacts={contacts}
                        onTextChange={(val) => updateProjectContactRow(row.key, "name", val)}
                        onSelectContact={(c) => fillProjectContactRow(row.key, c)}
                      />
                    </Field>

                    <Field label="Job title" id={`ap-contact-title-${row.key}`}>
                      <TextInput
                        id={`ap-contact-title-${row.key}`}
                        value={row.title}
                        onChange={(e) => updateProjectContactRow(row.key, "title", e.target.value)}
                      />
                    </Field>

                    <Field
                      label="Contact email"
                      id={`ap-contact-email-${row.key}`}
                      error={idx === 0 ? apErrors.contactEmail : apErrors.contacts?.[row.key]?.email}
                    >
                      <TextInput
                        id={`ap-contact-email-${row.key}`}
                        type="email"
                        error={idx === 0 ? apErrors.contactEmail : apErrors.contacts?.[row.key]?.email}
                        value={row.email}
                        onChange={(e) => updateProjectContactRow(row.key, "email", e.target.value)}
                      />
                    </Field>

                    <Field
                      label="Contact number"
                      id={`ap-contact-number-${row.key}`}
                      error={idx === 0 ? apErrors.contactNumber : apErrors.contacts?.[row.key]?.number}
                    >
                      <TextInput
                        id={`ap-contact-number-${row.key}`}
                        icon={Phone}
                        type="tel"
                        error={idx === 0 ? apErrors.contactNumber : apErrors.contacts?.[row.key]?.number}
                        value={row.number}
                        onChange={(e) => updateProjectContactRow(row.key, "number", e.target.value)}
                      />
                    </Field>

                    {apContacts.length > 1 && (
                      <button
                        type="button"
                        className="btn-secondary"
                        onClick={() => removeProjectContactRow(row.key)}
                      >
                        Remove this contact
                      </button>
                    )}
                  </div>
                ))}

                {apContacts.length < MAX_PROJECT_CONTACTS && (
                  <button type="button" className="btn-secondary" onClick={addProjectContactRow}>
                    <Plus size={14} strokeWidth={2} />
                    Add another contact
                  </button>
                )}
              </div>

              {apErrors.sync ? <div className="form-error">{apErrors.sync}</div> : null}
              <div className="form-actions">
                <button type="button" className="btn-primary" onClick={handleSaveProject} disabled={apSaving}>
                  {apSaving ? "Saving…" : "Save"}
                </button>
                <button type="button" className="btn-secondary" onClick={handleDiscardProject} disabled={apSaving}>Discard</button>
              </div>
            </div>
          )}

          {view === "viewProject" && (
            <div className="detail-page">
              <div className="detail-page-header-row">
                <button
                  type="button"
                  className="btn-back"
                  onClick={() => goTo(projectReturnView)}
                >
                  <ArrowLeft size={15} strokeWidth={1.75} />
                  {projectReturnView === "dashboard"
                    ? "Back to home"
                    : projectReturnView === "viewCompany"
                    ? "Back to company"
                    : projectReturnView === "unfinishedProjects"
                    ? "Back to unfinished projects"
                    : projectReturnView === "finishedProjects"
                    ? "Back to finalized projects"
                    : "Back to all projects"}
                </button>
                {selectedProject ? (
                  <div className="detail-header-actions">
                    <button type="button" className="btn-edit" onClick={startEditProject}>
                      <Pencil size={14} strokeWidth={1.9} />
                      Edit
                    </button>
                    <button
                      type="button"
                      className="btn-delete"
                      onClick={() => {
                        archiveProject(selectedProject.id);
                        goTo(projectReturnView);
                      }}
                    >
                      Delete
                    </button>
                  </div>
                ) : null}
              </div>

              {selectedProject ? (
                <>
                  <h1 className="page-title">{selectedProject.name}</h1>
                  <button
                    type="button"
                    className="page-subtitle page-subtitle-tight company-link"
                    onClick={() => openCompany(selectedProject.companyId, "viewProject")}
                  >
                    {companyName(selectedProject.companyId)}
                  </button>
                  <p className="page-meta">
                    Last updated {formatDateShort(selectedProject.updatedAt || selectedProject.createdAt)}
                  </p>

                  <div className="srl-progress" role="img" aria-label={"SRL " + selectedProject.srlLevel + " of 7"}>
                    {SRL_LEVELS.map((lvl) => {
                      const isFilled = Number(lvl) <= Number(selectedProject.srlLevel);
                      const isCurrent = lvl === selectedProject.srlLevel;
                      return (
                        <div
                          key={lvl}
                          className={
                            "srl-progress-segment" +
                            (isFilled ? " srl-progress-segment-filled" : "") +
                            (isCurrent ? " srl-progress-segment-current" : "")
                          }
                        >
                          {isCurrent ? (
                            <span className="srl-progress-label">SRL {selectedProject.srlLevel}</span>
                          ) : null}
                        </div>
                      );
                    })}
                  </div>

                  <div className="detail-grid">
                    <div className="detail-row">
                      <span className="detail-label">Company</span>
                      <button
                        type="button"
                        className="detail-value company-link"
                        onClick={() => openCompany(selectedProject.companyId, "viewProject")}
                      >
                        {companyName(selectedProject.companyId)}
                      </button>
                    </div>
                    <div className="detail-row">
                      <span className="detail-label">PI name</span>
                      <span className="detail-value">{selectedProject.lead}</span>
                    </div>
                    <div className="detail-row">
                      <span className="detail-label">Contact name</span>
                      <span className="detail-value">{selectedProject.contactName}</span>
                    </div>
                    <div className="detail-row">
                      <span className="detail-label">Contact email</span>
                      <span className="detail-value">{selectedProject.contactEmail}</span>
                    </div>
                    <div className="detail-row">
                      <span className="detail-label">Contact number</span>
                      <span className="detail-value">
                        {selectedProject.contactNumber ? formatPhoneDisplay(selectedProject.contactNumber) : ""}
                      </span>
                    </div>
                    <div className="detail-row">
                      <span className="detail-label">Funding</span>
                      <span className="detail-value">${formatMoney(selectedProject.funding)}</span>
                    </div>
                    <div className="detail-row">
                      <span className="detail-label">Current SRL</span>
                      <span className="detail-value">SRL {selectedProject.srlLevel}</span>
                    </div>
                    <div className="detail-row">
                      <span className="detail-label">Potential</span>
                      <span className="detail-value">{selectedProject.potentialLevel || "—"}</span>
                    </div>
                    <div className="detail-row">
                      <span className="detail-label">Status</span>
                      <span className="detail-value">{displayStatus(selectedProject.status)}</span>
                    </div>
                    {selectedProject.rfsNti ? (
                      <div className="detail-row">
                        <span className="detail-label">RFS/NTI number</span>
                        <span className="detail-value">{selectedProject.rfsNti}</span>
                      </div>
                    ) : null}
                    <div className="detail-row">
                      <span className="detail-label">Project Start Date</span>
                      <span className="detail-value">
                        {parseMDY(selectedProject.startDate) ? selectedProject.startDate : "—"}
                      </span>
                    </div>
                    <div className="detail-row">
                      <span className="detail-label">Project End Date</span>
                      <span className="detail-value">
                        {parseMDY(selectedProject.deadline) ? selectedProject.deadline : "—"}
                      </span>
                    </div>
                    <div className="detail-row">
                      <span className="detail-label">Time to complete</span>
                      <span className="detail-value">{timeUntil(selectedProject.deadline)}</span>
                    </div>
                  </div>

                  {(selectedProject.contactIds || []).length > 0 ? (
                    <div className="table-block">
                      <h2 className="section-title">Additional contacts</h2>
                      <div className="additional-contacts-list">
                        {selectedProject.contactIds
                          .map((id) => contacts.find((c) => c.id === id))
                          .filter(Boolean)
                          .map((c) => (
                            <div className="additional-contact-card" key={c.id}>
                              <div className="detail-row">
                                <span className="detail-label">Contact name</span>
                                <span className="detail-value">{c.contactName || "—"}</span>
                              </div>
                              <div className="detail-row">
                                <span className="detail-label">Job title</span>
                                <span className="detail-value">{c.jobTitle || "—"}</span>
                              </div>
                              <div className="detail-row">
                                <span className="detail-label">Contact email</span>
                                <span className="detail-value">{c.contactEmail || "—"}</span>
                              </div>
                              <div className="detail-row">
                                <span className="detail-label">Contact number</span>
                                <span className="detail-value">
                                  {c.contactNumber ? formatPhoneDisplay(c.contactNumber) : "—"}
                                </span>
                              </div>
                            </div>
                          ))}
                      </div>
                    </div>
                  ) : null}

                  <div className="table-block">
                    <h2 className="section-title">SRL Progression</h2>
                    <div className="table-wrap table-wrap-scroll srl-progression-scroll">
                      <table className="data-table srl-progression-table">
                        <thead>
                          <tr>
                            <th>SRL</th>
                            <th>Date reached</th>
                            <th>Time in level</th>
                            <th>Notes</th>
                          </tr>
                        </thead>
                        <tbody>
                          {SRL_LEVELS.map((lvl) => {
                            const history = selectedProject.srlHistory || [];
                            const entryIndex = history.findIndex((h) => h.level === lvl);
                            const entry = entryIndex >= 0 ? history[entryIndex] : null;
                            const nextEntry = entry ? history[entryIndex + 1] : null;

                            let dateReached = "—";
                            let timeInLevel = "—";
                            if (entry) {
                              dateReached = formatDateShort(entry.date);
                              const endMs = nextEntry ? nextEntry.date : Date.now();
                              const endLabel = nextEntry ? formatDateShort(nextEntry.date) : "Present";
                              const months = formatDuration(entry.date, endMs);
                              timeInLevel = `${formatDateShort(entry.date)} – ${endLabel} (${months})`;
                            }

                            const note = selectedProject.srlNotes ? selectedProject.srlNotes[lvl] : null;

                            return (
                              <tr key={lvl}>
                                <td>SRL {lvl}</td>
                                <td>{dateReached}</td>
                                <td>{timeInLevel}</td>
                                <td>
                                  <SrlNoteCell
                                    level={lvl}
                                    note={note}
                                    onSave={saveSrlNote}
                                    key={selectedProject.id + "-" + lvl}
                                  />
                                </td>
                              </tr>
                            );
                          })}
                        </tbody>
                      </table>
                    </div>
                  </div>

                  <div className="update-log">
                    <h2 className="section-title">Update log</h2>
                    {!selectedProject.updates || selectedProject.updates.length === 0 ? (
                      <p className="empty-table-note">No edits yet.</p>
                    ) : (
                      <div className="table-wrap">
                        <table className="data-table update-log-table">
                          <thead>
                            <tr>
                              <th>User</th>
                              <th>Field</th>
                              <th>Change</th>
                              <th>Date</th>
                            </tr>
                          </thead>
                          <tbody>
                            {[...selectedProject.updates]
                              .reverse()
                              .map((u, i) => (
                                <tr
                                  key={u.id}
                                  className={"update-log-row " + (i % 2 === 0 ? "update-log-row-a" : "update-log-row-b")}
                                >
                                  <td>{u.user}</td>
                                  <td>{u.field}</td>
                                  <td>
                                    "{displayStatus(u.from) || "—"}" → "{displayStatus(u.to) || "—"}"
                                  </td>
                                  <td className="update-log-time">{new Date(u.at).toLocaleString()}</td>
                                </tr>
                              ))}
                          </tbody>
                        </table>
                      </div>
                    )}
                  </div>
                </>
              ) : (
                <p className="empty-state-text">Project not found.</p>
              )}
            </div>
          )}

          {view === "editProject" && selectedProject && (
            <div className="detail-page">
              <button type="button" className="btn-back" onClick={handleCancelEditProject}>
                <ArrowLeft size={15} strokeWidth={1.75} />
                Back to project
              </button>
              <h1 className="page-title">Edit project</h1>
              <p className="page-subtitle">Update the project details below.</p>

              <div className="form-grid">
                <Field label="Project name" id="ep-name" error={epErrors.name} full>
                  <TextInput
                    id="ep-name"
                    error={epErrors.name}
                    value={epName}
                    onChange={(e) => setEpName(e.target.value)}
                  />
                </Field>

                <Field label="Company" id="ep-company" error={epErrors.companyId} full>
                  <SelectInput
                    id="ep-company"
                    error={epErrors.companyId}
                    value={epCompanyId}
                    onChange={(e) => setEpCompanyId(e.target.value)}
                  >
                    <option value="">{companies.length ? "Select company" : "No companies yet"}</option>
                    {[...companies]
                      .sort((a, b) => a.name.localeCompare(b.name))
                      .map((c) => (
                        <option key={c.id} value={c.id}>{c.name}</option>
                      ))}
                  </SelectInput>
                </Field>

                <Field label="RFS/NTI number (optional)" id="ep-rfs-nti" full>
                  <TextInput
                    id="ep-rfs-nti"
                    placeholder="#12345"
                    value={epRfsNti}
                    onChange={(e) => setEpRfsNti(maskHashPrefix(e.target.value))}
                  />
                </Field>

                <Field label="PI name" id="ep-lead" error={epErrors.lead}>
                  <TextInput
                    id="ep-lead"
                    error={epErrors.lead}
                    value={epLead}
                    onChange={(e) => setEpLead(e.target.value)}
                  />
                </Field>

                <Field label="Funding (enter 0 if not decided)" id="ep-funding" error={epErrors.funding}>
                  <TextInput
                    id="ep-funding"
                    prefix="$"
                    inputMode="decimal"
                    placeholder="0.00"
                    error={epErrors.funding}
                    value={epFunding}
                    onChange={(e) => setEpFunding(maskFundingInput(e.target.value))}
                  />
                </Field>

                <Field label="Current SRL" id="ep-srl" error={epErrors.srl}>
                  <SelectInput
                    id="ep-srl"
                    error={epErrors.srl}
                    value={epSrl}
                    onChange={(e) => setEpSrl(e.target.value)}
                  >
                    <option value="">Select level</option>
                    {SRL_LEVELS.map((lvl) => (
                      <option key={lvl} value={lvl}>{lvl}</option>
                    ))}
                  </SelectInput>
                </Field>

                <Field label="Potential (optional)" id="ep-potential" error={epErrors.potential}>
                  <SelectInput
                    id="ep-potential"
                    error={epErrors.potential}
                    value={epPotential}
                    onChange={(e) => setEpPotential(e.target.value)}
                  >
                    <option value="">Select potential</option>
                    {POTENTIAL_LEVELS.map((lvl) => (
                      <option key={lvl} value={lvl}>{lvl}</option>
                    ))}
                  </SelectInput>
                </Field>

                <Field label="Project Start Date" id="ep-start-date" error={epErrors.startDate}>
                  <TextInput
                    id="ep-start-date"
                    placeholder="DD/MM/YYYY"
                    error={epErrors.startDate}
                    value={epStartDate}
                    onChange={(e) => setEpStartDate(maskDateInput(e.target.value))}
                  />
                </Field>

                <Field label="Project End Date (Optional)" id="ep-deadline" error={epErrors.deadline}>
                  <TextInput
                    id="ep-deadline"
                    placeholder="DD/MM/YYYY"
                    error={epErrors.deadline}
                    value={epDeadline}
                    onChange={(e) => setEpDeadline(maskDateInput(e.target.value))}
                  />
                  {(() => {
                    const days = daysBetween(epStartDate, epDeadline);
                    if (days === null) return null;
                    return (
                      <p className="field-hint">
                        {days < 0
                          ? "Deadline is before the start date."
                          : days === 0
                          ? "Same-day deadline."
                          : days + (days === 1 ? " day" : " days") + " from start to deadline."}
                      </p>
                    );
                  })()}
                </Field>
              </div>

              <div className="form-section">
                <h2 className="form-section-title">Contact details</h2>
                <p className="form-section-subtitle">Who should we reach out to about this project? All fields optional.</p>

                {epContacts.map((row, idx) => (
                  <div className="form-grid" key={row.key} style={{ marginBottom: 16 }}>
                    <Field
                      label="Contact name"
                      id={`ep-contact-name-${row.key}`}
                      error={idx === 0 ? epErrors.contactName : epErrors.contacts?.[row.key]?.name}
                    >
                      <ContactNameCombo
                        id={`ep-contact-name-${row.key}`}
                        error={idx === 0 ? epErrors.contactName : epErrors.contacts?.[row.key]?.name}
                        value={row.name}
                        contacts={contacts}
                        onTextChange={(val) => updateEditContactRow(row.key, "name", val)}
                        onSelectContact={(c) => fillEditContactRow(row.key, c)}
                      />
                    </Field>

                    <Field label="Job title" id={`ep-contact-title-${row.key}`}>
                      <TextInput
                        id={`ep-contact-title-${row.key}`}
                        value={row.title}
                        onChange={(e) => updateEditContactRow(row.key, "title", e.target.value)}
                      />
                    </Field>

                    <Field
                      label="Contact email"
                      id={`ep-contact-email-${row.key}`}
                      error={idx === 0 ? epErrors.contactEmail : epErrors.contacts?.[row.key]?.email}
                    >
                      <TextInput
                        id={`ep-contact-email-${row.key}`}
                        type="email"
                        error={idx === 0 ? epErrors.contactEmail : epErrors.contacts?.[row.key]?.email}
                        value={row.email}
                        onChange={(e) => updateEditContactRow(row.key, "email", e.target.value)}
                      />
                    </Field>

                    <Field
                      label="Contact number"
                      id={`ep-contact-number-${row.key}`}
                      error={idx === 0 ? epErrors.contactNumber : epErrors.contacts?.[row.key]?.number}
                    >
                      <TextInput
                        id={`ep-contact-number-${row.key}`}
                        icon={Phone}
                        type="tel"
                        error={idx === 0 ? epErrors.contactNumber : epErrors.contacts?.[row.key]?.number}
                        value={row.number}
                        onChange={(e) => updateEditContactRow(row.key, "number", e.target.value)}
                      />
                    </Field>

                    {epContacts.length > 1 && (
                      <button
                        type="button"
                        className="btn-secondary"
                        onClick={() => removeEditContactRow(row.key)}
                      >
                        Remove this contact
                      </button>
                    )}
                  </div>
                ))}

                {epContacts.length < MAX_PROJECT_CONTACTS && (
                  <button type="button" className="btn-secondary" onClick={addEditContactRow}>
                    <Plus size={14} strokeWidth={2} />
                    Add another contact
                  </button>
                )}
              </div>

              <div className="form-section">
                <h2 className="form-section-title">SRL Progression</h2>
                <p className="form-section-subtitle">
                  Set the date each SRL level was reached. Leave a level blank if the project hasn't reached it yet.
                </p>
                <div className="srl-date-grid">
                  {SRL_LEVELS.map((lvl) => (
                    <Field
                      key={lvl}
                      label={"SRL " + lvl + " date reached"}
                      id={`ep-srl-date-${lvl}`}
                      error={epErrors.srlDates?.[lvl]}
                    >
                      <TextInput
                        id={`ep-srl-date-${lvl}`}
                        placeholder="DD/MM/YYYY"
                        error={epErrors.srlDates?.[lvl]}
                        value={epSrlDates[lvl] || ""}
                        onChange={(e) => updateEpSrlDate(lvl, maskDateInput(e.target.value))}
                      />
                    </Field>
                  ))}
                </div>
              </div>

              {epErrors.sync ? <div className="form-error">{epErrors.sync}</div> : null}
              <div className="form-actions">
                <button type="button" className="btn-primary" onClick={handleSaveEditProject} disabled={epSaving}>
                  {epSaving ? "Saving…" : "Save changes"}
                </button>
                <button type="button" className="btn-secondary" onClick={handleCancelEditProject} disabled={epSaving}>
                  Cancel
                </button>
              </div>
            </div>
          )}

          {view === "addCompany" && (
            <div className="detail-page">
              <button type="button" className="btn-back" onClick={() => goTo("companies")}>
                <ArrowLeft size={15} strokeWidth={1.75} />
                Back to companies
              </button>
              <h1 className="page-title">Add company</h1>
              <p className="page-subtitle">Enter the company details below.</p>

              <div className="form-grid">
                <Field label="Company name" id="ac-name" error={acErrors.name} full>
                  <TextInput
                    id="ac-name"
                    error={acErrors.name}
                    value={acName}
                    onChange={(e) => setAcName(e.target.value)}
                  />
                </Field>

                <Field label="About" id="ac-about" full>
                  <Textarea
                    id="ac-about"
                    rows={5}
                    value={acAbout}
                    onChange={(e) => setAcAbout(e.target.value)}
                  />
                </Field>

                <Field label="Industry" id="ac-industry">
                  <TextInput
                    id="ac-industry"
                    value={acIndustry}
                    onChange={(e) => setAcIndustry(e.target.value)}
                  />
                </Field>

                <Field label="HQ Location" id="ac-hq-location">
                  <TextInput
                    id="ac-hq-location"
                    value={acHqLocation}
                    onChange={(e) => setAcHqLocation(e.target.value)}
                  />
                </Field>

                <Field label="SBU Location" id="ac-sbu-location">
                  <TextInput
                    id="ac-sbu-location"
                    value={acSbuLocation}
                    onChange={(e) => setAcSbuLocation(e.target.value)}
                  />
                </Field>
              </div>

              <div className="form-section">
                <h2 className="form-section-title">Contacts</h2>
                <p className="form-section-subtitle">
                  Add anyone we should reach out to at this company. All fields optional.
                </p>

                {acContacts.map((row, idx) => (
                  <div className="form-grid" key={row.key} style={{ marginBottom: 16 }}>
                    <Field
                      label="Contact name"
                      id={`ac-contact-name-${row.key}`}
                      error={acErrors.contacts?.[row.key]?.name}
                    >
                      <ContactNameCombo
                        id={`ac-contact-name-${row.key}`}
                        error={acErrors.contacts?.[row.key]?.name}
                        value={row.name}
                        contacts={contacts}
                        onTextChange={(val) => updateCompanyContactRow(row.key, "name", val)}
                        onSelectContact={(c) => fillCompanyContactRow(row.key, c)}
                      />
                    </Field>

                    <Field label="Job title" id={`ac-contact-title-${row.key}`}>
                      <TextInput
                        id={`ac-contact-title-${row.key}`}
                        value={row.title}
                        onChange={(e) => updateCompanyContactRow(row.key, "title", e.target.value)}
                      />
                    </Field>

                    <Field
                      label="Contact email"
                      id={`ac-contact-email-${row.key}`}
                      error={acErrors.contacts?.[row.key]?.email}
                    >
                      <TextInput
                        id={`ac-contact-email-${row.key}`}
                        type="email"
                        error={acErrors.contacts?.[row.key]?.email}
                        value={row.email}
                        onChange={(e) => updateCompanyContactRow(row.key, "email", e.target.value)}
                      />
                    </Field>

                    <Field
                      label="Contact number"
                      id={`ac-contact-number-${row.key}`}
                      error={acErrors.contacts?.[row.key]?.number}
                    >
                      <TextInput
                        id={`ac-contact-number-${row.key}`}
                        icon={Phone}
                        type="tel"
                        error={acErrors.contacts?.[row.key]?.number}
                        value={row.number}
                        onChange={(e) => updateCompanyContactRow(row.key, "number", e.target.value)}
                      />
                    </Field>

                    {acContacts.length > 1 && (
                      <button
                        type="button"
                        className="btn-secondary"
                        onClick={() => removeCompanyContactRow(row.key)}
                      >
                        Remove this contact
                      </button>
                    )}
                  </div>
                ))}

                {acContacts.length < MAX_COMPANY_CONTACTS && (
                  <button type="button" className="btn-secondary" onClick={addCompanyContactRow}>
                    <Plus size={14} strokeWidth={2} />
                    Add another contact
                  </button>
                )}
              </div>

              <div className="form-actions">
                <button type="button" className="btn-primary" onClick={handleSaveCompany}>Save</button>
                <button type="button" className="btn-secondary" onClick={handleDiscardCompany}>Discard</button>
              </div>
            </div>
          )}

          {view === "viewCompany" && (
            <div className="detail-page">
              <div className="detail-page-header-row">
                <button type="button" className="btn-back" onClick={() => goTo(companyReturnView)}>
                  <ArrowLeft size={15} strokeWidth={1.75} />
                  {companyReturnView === "viewProject"
                    ? "Back to project"
                    : companyReturnView === "viewContact"
                    ? "Back to contact"
                    : "Back to companies"}
                </button>
                {selectedCompany ? (
                  <div className="detail-header-actions">
                    <button type="button" className="btn-edit" onClick={startEditCompany}>
                      <Pencil size={14} strokeWidth={1.9} />
                      Edit
                    </button>
                    <button
                      type="button"
                      className="btn-delete"
                      onClick={() => {
                        archiveCompany(selectedCompany.id);
                        goTo(companyReturnView);
                      }}
                    >
                      Delete
                    </button>
                  </div>
                ) : null}
              </div>

              {selectedCompany ? (
                <>
                  <h1 className="page-title">{selectedCompany.name}</h1>
                  <p className="page-subtitle">{selectedCompany.about || "No description yet."}</p>

                  <div className="detail-grid">
                    <div className="detail-row">
                      <span className="detail-label">Industry</span>
                      <span className="detail-value">{selectedCompany.industry || "—"}</span>
                    </div>
                    <div className="detail-row">
                      <span className="detail-label">HQ Location</span>
                      <span className="detail-value">{selectedCompany.hqLocation || "—"}</span>
                    </div>
                    <div className="detail-row">
                      <span className="detail-label">SBU Location</span>
                      <span className="detail-value">{selectedCompany.sbuLocation || "—"}</span>
                    </div>
                  </div>

                  <div className="home-stats">
                    <div className="stat-card">
                      <span className="stat-label">Total number of projects</span>
                      <span className="stat-value">{selectedCompanyProjectCount}</span>
                      <span className="stat-sub">for {selectedCompany.name}</span>
                    </div>
                    <div className="stat-card">
                      <span className="stat-label">Completion rate</span>
                      <span className="stat-value">{companyCompletionRate}%</span>
                      <span className="stat-sub">
                        {companyFinishedCount} of {selectedCompanyProjectCount} finalized
                      </span>
                    </div>
                    <div className="stat-card">
                      <span className="stat-label">Total funding</span>
                      <span className="stat-value">${formatMoney(companyTotalFunding)}</span>
                      <span className="stat-sub">across {companyFundedProjects.length} funded project{companyFundedProjects.length === 1 ? "" : "s"}</span>
                    </div>
                  </div>

                  <div className="table-block">
                    <h2 className="section-title">Contacts</h2>
                    {companyContacts.length === 0 ? (
                      <p className="empty-table-note">No contacts on file yet.</p>
                    ) : (
                      <div className="table-wrap">
                        <table className="data-table">
                          <thead>
                            <tr>
                              <th>Contact name</th>
                              <th>Job title</th>
                              <th>Project</th>
                              <th>Email</th>
                              <th>Mobile number</th>
                            </tr>
                          </thead>
                          <tbody>
                            {companyContacts.map((p) => (
                              <tr key={p.id} className="data-row" onClick={() => openProject(p.id, "viewCompany")}>
                                <td>{p.contactName || "—"}</td>
                                <td>{p.contactTitle || "—"}</td>
                                <td>{p.name}</td>
                                <td>{p.contactEmail || "—"}</td>
                                <td>{p.contactNumber ? formatPhoneDisplay(p.contactNumber) : "—"}</td>
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      </div>
                    )}
                  </div>

                  <div className="table-block">
                    <h2 className="section-title">General contacts</h2>
                    {companyGeneralContacts.length === 0 ? (
                      <p className="empty-table-note">No general contacts on file yet.</p>
                    ) : (
                      <div className="table-wrap">
                        <table className="data-table">
                          <thead>
                            <tr>
                              <th>Contact name</th>
                              <th>Job title</th>
                              <th>Email</th>
                              <th>Mobile number</th>
                            </tr>
                          </thead>
                          <tbody>
                            {companyGeneralContacts.map((c) => (
                              <tr key={c.id} className="data-row" onClick={() => openContact(c.id, "viewCompany")}>
                                <td>{c.contactName || "—"}</td>
                                <td>{c.jobTitle || "—"}</td>
                                <td>{c.contactEmail || "—"}</td>
                                <td>{c.contactNumber ? formatPhoneDisplay(c.contactNumber) : "—"}</td>
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      </div>
                    )}
                  </div>

                  {["Unfinished", "Finished"].map((status) => {
                    const rows = projects
                      .filter((p) => p.companyId === selectedCompany.id && p.status === status)
                      .sort((a, b) => {
                        const lvlDiff = Number(a.srlLevel) - Number(b.srlLevel);
                        if (lvlDiff !== 0) return lvlDiff;
                        return parseFloat(b.funding || 0) - parseFloat(a.funding || 0);
                      });
                    return (
                      <div className="table-block" key={status}>
                        <h2 className="section-title">{displayStatus(status)} projects</h2>
                        {rows.length === 0 ? (
                          <p className="empty-table-note">No {displayStatus(status).toLowerCase()} projects yet.</p>
                        ) : (
                          <div className="table-wrap">
                            <table className="data-table">
                              <thead>
                                <tr>
                                  <th>SRL</th>
                                  <th>Project name</th>
                                  <th>PI name</th>
                                  <th>Funding</th>
                                  <th>Potential</th>
                                  <th>Done</th>
                                </tr>
                              </thead>
                              <tbody>
                                {rows.map((p) => (
                                  <tr key={p.id} className="data-row" onClick={() => openProject(p.id, "viewCompany")}>
                                    <td>SRL {p.srlLevel}</td>
                                    <td>{p.name}</td>
                                    <td>{p.lead}</td>
                                    <td>${formatMoney(p.funding)}</td>
                                    <td>
                                      <PotentialChip level={p.potentialLevel} />
                                    </td>
                                    <td>
                                      <input
                                        type="checkbox"
                                        className="row-checkbox"
                                        checked={p.status === "Finished"}
                                        onClick={(e) => e.stopPropagation()}
                                        onChange={() => toggleProjectFinished(p)}
                                      />
                                    </td>
                                  </tr>
                                ))}
                              </tbody>
                            </table>
                          </div>
                        )}
                      </div>
                    );
                  })}

                  <div className="update-log">
                    <h2 className="section-title">Update log</h2>
                    {!selectedCompany.updates || selectedCompany.updates.length === 0 ? (
                      <p className="empty-table-note">No edits yet.</p>
                    ) : (
                      <div className="table-wrap">
                        <table className="data-table update-log-table">
                          <thead>
                            <tr>
                              <th>User</th>
                              <th>Field</th>
                              <th>Change</th>
                              <th>Date</th>
                            </tr>
                          </thead>
                          <tbody>
                            {[...selectedCompany.updates]
                              .reverse()
                              .map((u, i) => (
                                <tr
                                  key={u.id}
                                  className={"update-log-row " + (i % 2 === 0 ? "update-log-row-a" : "update-log-row-b")}
                                >
                                  <td>{u.user}</td>
                                  <td>{u.field}</td>
                                  <td>
                                    "{displayStatus(u.from) || "—"}" → "{displayStatus(u.to) || "—"}"
                                  </td>
                                  <td className="update-log-time">{new Date(u.at).toLocaleString()}</td>
                                </tr>
                              ))}
                          </tbody>
                        </table>
                      </div>
                    )}
                  </div>
                </>
              ) : (
                <p className="empty-state-text">Company not found.</p>
              )}
            </div>
          )}

          {view === "editCompany" && selectedCompany && (
            <div className="detail-page">
              <button type="button" className="btn-back" onClick={handleCancelEditCompany}>
                <ArrowLeft size={15} strokeWidth={1.75} />
                Back to company
              </button>
              <h1 className="page-title">Edit company</h1>
              <p className="page-subtitle">Update the company details below.</p>

              <div className="form-grid">
                <Field label="Company name" id="ec-name" error={ecErrors.name} full>
                  <TextInput
                    id="ec-name"
                    error={ecErrors.name}
                    value={ecName}
                    onChange={(e) => setEcName(e.target.value)}
                  />
                </Field>

                <Field label="About" id="ec-about" full>
                  <Textarea
                    id="ec-about"
                    rows={5}
                    value={ecAbout}
                    onChange={(e) => setEcAbout(e.target.value)}
                  />
                </Field>

                <Field label="Industry" id="ec-industry">
                  <TextInput
                    id="ec-industry"
                    value={ecIndustry}
                    onChange={(e) => setEcIndustry(e.target.value)}
                  />
                </Field>

                <Field label="HQ Location" id="ec-hq-location">
                  <TextInput
                    id="ec-hq-location"
                    value={ecHqLocation}
                    onChange={(e) => setEcHqLocation(e.target.value)}
                  />
                </Field>

                <Field label="SBU Location" id="ec-sbu-location">
                  <TextInput
                    id="ec-sbu-location"
                    value={ecSbuLocation}
                    onChange={(e) => setEcSbuLocation(e.target.value)}
                  />
                </Field>
              </div>

              <div className="form-actions">
                <button type="button" className="btn-primary" onClick={handleSaveEditCompany}>
                  Save changes
                </button>
                <button type="button" className="btn-secondary" onClick={handleCancelEditCompany}>
                  Cancel
                </button>
              </div>
            </div>
          )}
        </AuthedLayout>
      </Shell>
    );
  }


  return (
    <Shell>
      <div className="auth-page">
        <div className="auth-card">
          <div className="brand brand-centered">
            <BrandLogo size={26} />
            <span className="brand-name">{APP_NAME}</span>
          </div>

          {view === "login" && (
            <>
              <h1 className="auth-title">Sign in</h1>
              <p className="auth-subtitle">Enter your username and password to continue.</p>

              {banner ? <div className="success-banner">{banner}</div> : null}
              {storageError ? <div className="form-error">{storageError}</div> : null}
              {loginErrors.form ? <div className="form-error">{loginErrors.form}</div> : null}

              <div>
                <Field label="Username" id="login-username" error={loginErrors.username}>
                  <TextInput
                    id="login-username"
                    icon={User}
                    error={loginErrors.username}
                    autoComplete="username"
                    autoCapitalize="none"
                    autoCorrect="off"
                    spellCheck={false}
                    value={loginUsername}
                    onChange={(e) => setLoginUsername(e.target.value)}
                  />
                </Field>

                <Field label="Password" id="login-password" error={loginErrors.password}>
                  <TextInput
                    id="login-password"
                    icon={Lock}
                    type="password"
                    error={loginErrors.password}
                    autoComplete="current-password"
                    value={loginPassword}
                    onChange={(e) => setLoginPassword(e.target.value)}
                    onKeyDown={(e) => e.key === "Enter" && handleLogin(e)}
                  />
                </Field>

                <div className="row-between">
                  <button type="button" className="btn-link" onClick={() => goTo("forgot")}>
                    Forgot password?
                  </button>
                </div>

                <button type="button" className="btn-primary btn-block" onClick={handleLogin}>Sign in</button>
              </div>

              <p className="auth-footer-text">
                Don't have an account?{" "}
                <button type="button" className="btn-link" onClick={() => goTo("signup")}>
                  Create one
                </button>
              </p>
            </>
          )}

          {view === "signup" && (
            <>
              <h1 className="auth-title">Create your account</h1>
              <p className="auth-subtitle">Set up a username, password, and email to get started.</p>

              <div>
                <Field label="Username" id="su-username" error={suErrors.username}>
                  <TextInput
                    id="su-username"
                    icon={User}
                    error={suErrors.username}
                    autoComplete="username"
                    value={suUsername}
                    onChange={(e) => setSuUsername(e.target.value)}
                  />
                </Field>

                <Field label="Email" id="su-email" error={suErrors.email}>
                  <TextInput
                    id="su-email"
                    icon={Mail}
                    type="email"
                    error={suErrors.email}
                    autoComplete="email"
                    value={suEmail}
                    onChange={(e) => setSuEmail(e.target.value)}
                  />
                </Field>

                <Field label="Password" id="su-password" error={suErrors.password}>
                  <TextInput
                    id="su-password"
                    icon={Lock}
                    type="password"
                    error={suErrors.password}
                    autoComplete="new-password"
                    value={suPassword}
                    onChange={(e) => setSuPassword(e.target.value)}
                  />
                </Field>

                <Field label="Confirm password" id="su-confirm" error={suErrors.confirm}>
                  <TextInput
                    id="su-confirm"
                    icon={Lock}
                    type="password"
                    error={suErrors.confirm}
                    autoComplete="new-password"
                    value={suConfirm}
                    onChange={(e) => setSuConfirm(e.target.value)}
                    onKeyDown={(e) => e.key === "Enter" && handleSignup(e)}
                  />
                </Field>

                <button type="button" className="btn-primary btn-block" onClick={handleSignup}>Create account</button>
              </div>

              <p className="auth-footer-text">
                Already have an account?{" "}
                <button type="button" className="btn-link" onClick={() => goTo("login")}>
                  Sign in
                </button>
              </p>
            </>
          )}

          {view === "forgot" && (
            <>
              <button type="button" className="btn-back" onClick={() => goTo("login")}>
                <ArrowLeft size={15} strokeWidth={1.75} />
                Back to sign in
              </button>
              <h1 className="auth-title">Reset your password</h1>
              <p className="auth-subtitle">
                Enter the email on your account and we'll send a link to reset your password.
              </p>

              <div>
                <Field label="Email" id="fp-email" error={fpError}>
                  <TextInput
                    id="fp-email"
                    icon={Mail}
                    type="email"
                    error={fpError}
                    autoComplete="email"
                    value={fpEmail}
                    onChange={(e) => setFpEmail(e.target.value)}
                    onKeyDown={(e) => e.key === "Enter" && handleForgot(e)}
                  />
                </Field>

                <button type="button" className="btn-primary btn-block" onClick={handleForgot}>Send reset link</button>
              </div>
            </>
          )}

          {view === "forgotSent" && (
            <>
              <button type="button" className="btn-back" onClick={() => goTo("login")}>
                <ArrowLeft size={15} strokeWidth={1.75} />
                Back to sign in
              </button>
              <div className="check-mark-circle">
                <Check size={18} strokeWidth={2} />
              </div>
              <h1 className="auth-title">Check your email</h1>
              <p className="auth-subtitle">
                If an account exists for that email, a link to reset your password is on its way.
              </p>

              {pendingReset ? (
                <div className="demo-callout">
                  <p className="demo-callout-label">Demo email preview</p>
                  <p className="demo-callout-text">
                    Since this project has no email server connected, use the button below in
                    place of clicking the link in the email.
                  </p>
                  <button type="button" className="btn-primary btn-block" onClick={() => goTo("reset")}>
                    Open reset link
                  </button>
                </div>
              ) : null}
            </>
          )}

          {view === "reset" && pendingReset && (
            <>
              <h1 className="auth-title">Set a new password</h1>
              <p className="auth-subtitle">Choose a new password for {pendingReset.email}.</p>

              <div>
                <Field label="New password" id="rp-password" error={rpErrors.password}>
                  <TextInput
                    id="rp-password"
                    icon={Lock}
                    type="password"
                    error={rpErrors.password}
                    autoComplete="new-password"
                    value={rpPassword}
                    onChange={(e) => setRpPassword(e.target.value)}
                  />
                </Field>

                <Field label="Confirm new password" id="rp-confirm" error={rpErrors.confirm}>
                  <TextInput
                    id="rp-confirm"
                    icon={Lock}
                    type="password"
                    error={rpErrors.confirm}
                    autoComplete="new-password"
                    value={rpConfirm}
                    onChange={(e) => setRpConfirm(e.target.value)}
                    onKeyDown={(e) => e.key === "Enter" && handleReset(e)}
                  />
                </Field>

                <button type="button" className="btn-primary btn-block" onClick={handleReset}>Reset password</button>
              </div>
            </>
          )}
        </div>
      </div>
    </Shell>
  );
}

function Shell({ children }) {
  return (
    <div className="meridian-root">
      <style>{CSS}</style>
      {children}
    </div>
  );
}

const CSS = `
.meridian-root {
  --bg: #ffffff;
  --surface: #fafafa;
  --border: #e4e6e9;
  --border-strong: #d7dade;
  --ink: #14181f;
  --text-primary: #14181f;
  --text-secondary: #6b7280;
  --text-tertiary: #98a2b3;
  --accent: #0f6b5c;
  --accent-dark: #0b5147;
  --danger: #b42318;
  --danger-bg: #fef3f2;
  --success-bg: #f0f9f6;
  --success-text: #0b5147;
  font-family: "Inter", -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
  color: var(--text-primary);
  background: var(--bg);
  min-height: 100vh;
  width: 100%;
  -webkit-font-smoothing: antialiased;
}

.meridian-root * { box-sizing: border-box; }

.auth-loading {
  display: flex;
  align-items: center;
  justify-content: center;
  min-height: 100vh;
  color: var(--text-secondary);
  font-size: 14px;
}

/* ---------- Auth pages ---------- */
.auth-page {
  min-height: 100vh;
  display: flex;
  align-items: center;
  justify-content: center;
  padding: 32px 20px;
  background: var(--bg);
}

.auth-card {
  width: 100%;
  max-width: 380px;
  padding: 36px 32px 32px;
  background: var(--bg);
  border: 1px solid var(--border);
  border-radius: 10px;
}

.brand {
  display: flex;
  align-items: center;
  gap: 8px;
  margin-bottom: 28px;
}

.brand-centered { justify-content: flex-start; }

.brand-name {
  font-weight: 600;
  font-size: 15px;
  letter-spacing: -0.01em;
  color: var(--ink);
}

.auth-title {
  font-size: 20px;
  font-weight: 600;
  letter-spacing: -0.01em;
  margin: 0 0 6px;
  color: var(--ink);
}

.auth-subtitle {
  font-size: 13.5px;
  line-height: 1.5;
  color: var(--text-secondary);
  margin: 0 0 24px;
}

.auth-footer-text {
  margin: 20px 0 0;
  font-size: 13px;
  color: var(--text-secondary);
  text-align: center;
}

.btn-back {
  display: inline-flex;
  align-items: center;
  gap: 8px;
  background: none;
  border: 1px solid var(--border);
  border-radius: 8px;
  padding: 8px 14px;
  margin-bottom: 18px;
  font-family: inherit;
  font-size: 14px;
  font-weight: 500;
  color: var(--text-secondary);
  cursor: pointer;
  transition: background 120ms ease, color 120ms ease, border-color 120ms ease;
}
.btn-back:hover { color: var(--ink); background: var(--surface); border-color: var(--border-strong); }

/* ---------- Form fields ---------- */
.field { margin-bottom: 16px; }

.label {
  display: block;
  font-size: 12.5px;
  font-weight: 500;
  color: var(--text-primary);
  margin-bottom: 6px;
}

.input-wrap {
  position: relative;
  display: flex;
  align-items: center;
}

.input-icon {
  position: absolute;
  left: 12px;
  color: var(--text-tertiary);
  pointer-events: none;
}

.input {
  width: 100%;
  font-family: inherit;
  font-size: 14px;
  color: var(--text-primary);
  background: var(--bg);
  border: 1px solid var(--border-strong);
  border-radius: 7px;
  padding: 10px 12px 10px 36px;
  outline: none;
  transition: border-color 120ms ease, box-shadow 120ms ease;
}

.input::placeholder { color: var(--text-tertiary); }

.input:focus {
  border-color: var(--accent);
  box-shadow: 0 0 0 3px rgba(15, 107, 92, 0.12);
}

.input-wrap-error .input {
  border-color: var(--danger);
}
.input-wrap-error .input:focus {
  box-shadow: 0 0 0 3px rgba(180, 35, 24, 0.1);
}

.error-text {
  margin-top: 6px;
  font-size: 12px;
  color: var(--danger);
}

/* ---------- Contact name combo (suggestions dropdown) ---------- */
.contact-combo {
  position: relative;
}

.contact-combo-dropdown {
  position: absolute;
  top: calc(100% + 4px);
  left: 0;
  right: 0;
  z-index: 20;
  background: var(--bg);
  border: 1px solid var(--border-strong);
  border-radius: 8px;
  box-shadow: 0 8px 24px rgba(15, 23, 42, 0.12);
  padding: 4px;
  max-height: 240px;
  overflow-y: auto;
  display: flex;
  flex-direction: column;
  gap: 1px;
}

.contact-combo-option {
  width: 100%;
  display: flex;
  flex-direction: column;
  align-items: flex-start;
  gap: 1px;
  text-align: left;
  font-family: inherit;
  background: none;
  border: none;
  border-radius: 6px;
  padding: 8px 10px;
  cursor: pointer;
  transition: background 120ms ease;
}
.contact-combo-option:hover { background: #eaf4ff; }
.contact-combo-option:active { background: #dcecff; }

.contact-combo-option-name {
  font-size: 13.5px;
  font-weight: 600;
  color: var(--text-primary);
}

.contact-combo-option-sub {
  font-size: 12px;
  color: var(--text-secondary);
  white-space: nowrap;
  overflow: hidden;
  text-overflow: ellipsis;
  max-width: 100%;
}

.form-error {
  background: var(--danger-bg);
  color: var(--danger);
  font-size: 13px;
  border-radius: 7px;
  padding: 10px 12px;
  margin-bottom: 16px;
}

.success-banner {
  background: var(--success-bg);
  color: var(--success-text);
  font-size: 13px;
  border-radius: 7px;
  padding: 10px 12px;
  margin-bottom: 16px;
}

.row-between {
  display: flex;
  justify-content: flex-end;
  margin: -6px 0 18px;
}

/* ---------- Buttons ---------- */
.btn-primary {
  width: auto;
  font-family: inherit;
  font-size: 14px;
  font-weight: 600;
  color: #ffffff;
  background: var(--ink);
  border: 1px solid var(--ink);
  border-radius: 7px;
  padding: 10px 16px;
  cursor: pointer;
  transition: background 120ms ease, transform 60ms ease;
}
.btn-primary:hover { background: #262b34; }
.btn-primary:active { transform: translateY(1px); }

.btn-block { width: 100%; }

.btn-secondary {
  display: inline-flex;
  align-items: center;
  gap: 8px;
  font-family: inherit;
  font-size: 13.5px;
  font-weight: 500;
  color: var(--text-primary);
  background: var(--bg);
  border: 1px solid var(--border-strong);
  border-radius: 7px;
  padding: 8px 14px;
  cursor: pointer;
  transition: background 120ms ease;
}
.btn-secondary:hover { background: var(--surface); }

.btn-link {
  font-family: inherit;
  font-size: 12.5px;
  font-weight: 500;
  color: var(--accent);
  background: none;
  border: none;
  padding: 0;
  cursor: pointer;
}
.btn-link:hover { color: var(--accent-dark); text-decoration: underline; }

.btn-edit {
  display: inline-flex;
  align-items: center;
  gap: 6px;
  font-family: inherit;
  font-size: 13px;
  font-weight: 500;
  color: var(--accent-dark);
  background: var(--success-bg);
  border: 1px solid #bfe3d8;
  border-radius: 7px;
  padding: 7px 13px;
  cursor: pointer;
  transition: background 120ms ease, border-color 120ms ease;
}
.btn-edit:hover { background: #e3f2ec; border-color: var(--accent); }
.btn-edit:active { transform: translateY(1px); }

.detail-header-actions {
  display: inline-flex;
  align-items: center;
  gap: 8px;
}

.btn-delete {
  display: inline-flex;
  align-items: center;
  gap: 6px;
  font-family: inherit;
  font-size: 13px;
  font-weight: 500;
  color: #b3261e;
  background: #fdecea;
  border: 1px solid #f3c2bd;
  border-radius: 7px;
  padding: 7px 13px;
  cursor: pointer;
  transition: background 120ms ease, border-color 120ms ease;
}
.btn-delete:hover { background: #fbdbd7; border-color: #e2897f; }
.btn-delete:active { transform: translateY(1px); }

/* ---------- Demo callout (forgot-password mock) ---------- */
.check-mark-circle {
  width: 34px;
  height: 34px;
  border-radius: 50%;
  background: var(--success-bg);
  color: var(--success-text);
  display: flex;
  align-items: center;
  justify-content: center;
  margin-bottom: 14px;
}

.demo-callout {
  margin-top: 20px;
  padding: 14px;
  border: 1px dashed var(--border-strong);
  border-radius: 8px;
  background: var(--surface);
}

.demo-callout-label {
  font-family: "JetBrains Mono", monospace;
  font-size: 10.5px;
  letter-spacing: 0.04em;
  text-transform: uppercase;
  color: var(--text-tertiary);
  margin: 0 0 6px;
}

.demo-callout-text {
  font-size: 12.5px;
  line-height: 1.5;
  color: var(--text-secondary);
  margin: 0 0 12px;
}

/* ---------- Dashboard ---------- */
.dashboard {
  min-height: 100vh;
  display: flex;
  flex-direction: column;
  background: var(--bg);
}

.dashboard-header {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 10px;
  flex-wrap: wrap;
  padding: 16px 28px;
  border-bottom: 1px solid var(--border);
}

.dashboard-user {
  display: flex;
  align-items: center;
  gap: 10px;
  min-width: 0;
}

.dashboard-user-name {
  font-size: 13.5px;
  font-weight: 600;
  color: var(--text-primary);
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
  max-width: 40vw;
}

.badge-mono {
  font-family: "JetBrains Mono", monospace;
  font-size: 11px;
  color: var(--text-tertiary);
  background: var(--surface);
  border: 1px solid var(--border);
  border-radius: 5px;
  padding: 3px 7px;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
  max-width: 45vw;
}

.dashboard-main {
  flex: 1;
  display: flex;
  align-items: center;
  justify-content: center;
  padding: 40px 20px;
}

.empty-state {
  text-align: center;
  max-width: 320px;
}

.empty-state-mark {
  display: flex;
  justify-content: center;
  margin-bottom: 16px;
  opacity: 0.9;
}

.empty-state-title {
  font-size: 19px;
  font-weight: 600;
  letter-spacing: -0.01em;
  color: var(--ink);
  margin: 0 0 6px;
}

.empty-state-text {
  font-size: 13.5px;
  color: var(--text-secondary);
  margin: 0;
}

.dashboard-footer {
  display: flex;
  justify-content: center;
  padding: 18px 28px 24px;
  border-top: 1px solid var(--border);
}

/* ---------- Sub-pages with a corner action tile ---------- */
.page-content {
  width: 100%;
  align-self: stretch;
  display: flex;
  flex-direction: column;
}

.page-content-header {
  padding: 0 4px;
}

.skeleton-grid {
  display: grid;
  grid-template-columns: repeat(auto-fill, minmax(260px, 1fr));
  gap: 14px;
  margin-top: 18px;
}

.skeleton-card {
  border: 1px solid var(--line, rgba(0, 0, 0, 0.08));
  border-radius: 12px;
  padding: 18px;
  background: var(--surface, #fff);
}

.skeleton-line {
  height: 11px;
  border-radius: 6px;
  background: linear-gradient(90deg, rgba(0,0,0,0.06), rgba(0,0,0,0.12), rgba(0,0,0,0.06));
  background-size: 200% 100%;
  animation: skeleton-shimmer 1.3s ease-in-out infinite;
  margin-bottom: 10px;
}

.skeleton-line-lg { width: 70%; height: 14px; }
.skeleton-line-md { width: 90%; }
.skeleton-line-sm { width: 45%; margin-bottom: 0; }

@keyframes skeleton-shimmer {
  0% { background-position: 200% 0; }
  100% { background-position: -200% 0; }
}

.page-title {
  font-size: 19px;
  font-weight: 600;
  letter-spacing: -0.01em;
  color: var(--ink);
  margin: 6px 0 0;
}

.page-body {
  flex: 1;
  display: flex;
  flex-wrap: wrap;
  align-items: flex-start;
  justify-content: flex-start;
  gap: 20px;
  padding: 28px 4px 4px;
}

.list-toolbar {
  display: flex;
  align-items: center;
  gap: 10px;
  margin-top: 18px;
  padding: 0 4px;
  flex-wrap: wrap;
}

.search-wrap {
  flex: 1;
  min-width: 220px;
}

.search-input {
  width: 100%;
  font-family: inherit;
  font-size: 13.5px;
  color: var(--text-primary);
  background: var(--bg);
  border: 1px solid var(--border-strong);
  border-radius: 7px;
  padding: 9px 12px;
  outline: none;
  transition: border-color 120ms ease, box-shadow 120ms ease;
}
.search-input::placeholder { color: var(--text-tertiary); }
.search-input:focus {
  border-color: var(--accent);
  box-shadow: 0 0 0 3px rgba(15, 107, 92, 0.12);
}

.sort-select-wrap {
  min-width: 200px;
}
.sort-select-wrap .select {
  width: 100%;
}

.add-tile {
  width: 31%;
  height: 150px;
  min-width: 190px;
  min-height: 150px;
  max-width: 300px;
  max-height: 150px;
  display: flex;
  flex-direction: column;
  align-items: center;
  justify-content: center;
  gap: 12px;
  background: var(--surface);
  border: 1.5px dashed var(--border-strong);
  border-radius: 12px;
  cursor: pointer;
  transition: border-color 120ms ease, background 120ms ease;
}
.add-tile:hover {
  border-color: var(--accent);
  background: var(--success-bg);
}
.add-tile:focus-visible {
  outline: 2px solid var(--accent);
  outline-offset: 2px;
}

.page-body-companies .add-tile {
  width: 32%;
  height: 300px;
  min-width: 240px;
  min-height: 300px;
  max-width: 340px;
  max-height: 300px;
  padding: 20px;
}
.page-body-companies .project-tile {
  gap: 6px;
  overflow: hidden;
}

.poc-block {
  margin-top: auto;
  padding-top: 12px;
  border-top: 1px solid var(--border);
  display: flex;
  flex-direction: column;
  gap: 3px;
  width: 100%;
}
.poc-label {
  font-size: 10.5px;
  font-weight: 700;
  letter-spacing: 0.04em;
  text-transform: uppercase;
  color: var(--text-tertiary);
}
.poc-name {
  font-size: 13px;
  font-weight: 600;
  color: var(--ink);
  white-space: nowrap;
  overflow: hidden;
  text-overflow: ellipsis;
  max-width: 100%;
}
.poc-title {
  font-size: 11.5px;
  color: var(--text-secondary);
  white-space: nowrap;
  overflow: hidden;
  text-overflow: ellipsis;
  max-width: 100%;
}
.poc-detail {
  display: flex;
  align-items: center;
  gap: 5px;
  font-size: 11.5px;
  color: var(--text-secondary);
  white-space: nowrap;
  overflow: hidden;
  text-overflow: ellipsis;
  max-width: 100%;
}
.poc-detail svg {
  flex-shrink: 0;
  color: var(--text-tertiary);
}
.poc-empty {
  font-size: 11.5px;
  color: var(--text-tertiary);
  font-style: italic;
}
.page-body-companies .add-tile-icon {
  width: 34px;
  height: 34px;
}

.add-tile-icon {
  width: 42px;
  height: 42px;
  border-radius: 50%;
  display: flex;
  align-items: center;
  justify-content: center;
  background: var(--bg);
  border: 1px solid var(--border-strong);
  color: var(--accent);
}

.add-tile-label {
  font-size: 14px;
  font-weight: 600;
  color: var(--text-primary);
}

/* ---------- Project tiles (saved projects, same size as add-tile) ---------- */
.project-tile {
  border-style: solid;
  border-color: var(--border);
  background: var(--bg);
  align-items: flex-start;
  justify-content: flex-start;
  text-align: left;
  padding: 14px 16px;
  gap: 4px;
}
.project-tile:hover {
  border-color: var(--accent);
  background: var(--surface);
}

.project-tile-name {
  width: 100%;
  font-size: 14px;
  font-weight: 600;
  color: var(--ink);
  letter-spacing: -0.005em;
  white-space: nowrap;
  overflow: hidden;
  text-overflow: ellipsis;
}

.project-tile-company {
  width: 100%;
  font-size: 12px;
  color: var(--text-secondary);
  white-space: nowrap;
  overflow: hidden;
  text-overflow: ellipsis;
}

.project-tile-chips {
  display: flex;
  flex-wrap: wrap;
  gap: 6px;
  margin-top: auto;
  padding-top: 8px;
}

.project-tile-status-row {
  display: flex;
  align-items: center;
  gap: 8px;
  width: 100%;
}

.project-tile-status-check {
  display: inline-flex;
  align-items: center;
  cursor: pointer;
}

.tile-delete-btn {
  margin-left: auto;
  display: inline-flex;
  align-items: center;
  justify-content: center;
  width: 26px;
  height: 26px;
  border-radius: 6px;
  border: 1px solid transparent;
  background: transparent;
  color: var(--text-secondary);
  cursor: pointer;
  transition: background 120ms ease, border-color 120ms ease, color 120ms ease;
}
.tile-delete-btn:hover {
  background: #fdecea;
  border-color: #f3c2bd;
  color: #b3261e;
}
.tile-delete-btn:active { transform: translateY(1px); }

.row-action-group {
  display: inline-flex;
  align-items: center;
  gap: 4px;
}

.tile-edit-btn {
  display: inline-flex;
  align-items: center;
  justify-content: center;
  width: 26px;
  height: 26px;
  border-radius: 6px;
  border: 1px solid transparent;
  background: transparent;
  color: var(--text-secondary);
  cursor: pointer;
  transition: background 120ms ease, border-color 120ms ease, color 120ms ease;
}
.tile-edit-btn:hover {
  background: var(--surface);
  border-color: var(--accent);
  color: var(--accent-dark);
}
.tile-edit-btn:active { transform: translateY(1px); }

.row-checkbox {
  width: 15px;
  height: 15px;
  cursor: pointer;
  accent-color: var(--accent);
}

.status-label {
  font-size: 11px;
  font-weight: 600;
  text-transform: uppercase;
  letter-spacing: 0.04em;
  padding: 3px 9px;
  border-radius: 999px;
}

.status-label-finished {
  color: var(--success-text);
  background: var(--success-bg);
  border: 1px solid #bfe3d8;
}

.status-label-unfinished {
  color: var(--text-secondary);
  background: var(--surface);
  border: 1px solid var(--border);
}

.chip {
  font-family: "JetBrains Mono", monospace;
  font-size: 10.5px;
  padding: 3px 8px;
  border-radius: 999px;
  background: var(--surface);
  border: 1px solid var(--border);
  color: var(--text-secondary);
}

.chip-low { color: var(--text-secondary); }
.chip-medium { color: #9a6700; background: #fff8e6; border-color: #f2dca0; }
.chip-high { color: var(--success-text); background: var(--success-bg); border-color: #bfe3d8; }

.chip-lg {
  font-size: 12.5px;
  font-weight: 600;
  padding: 5px 11px;
}

.chip-funding {
  font-family: "Inter", -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
  color: var(--ink);
  background: var(--bg);
  border-color: var(--border-strong);
}

/* ---------- Detail / form pages (Add project, View project) ---------- */
.detail-page {
  width: 100%;
  max-width: 640px;
  margin: 0 auto;
  padding: 4px 4px 40px;
}

.detail-page-header-row {
  display: flex;
  align-items: center;
  justify-content: space-between;
  margin-bottom: 18px;
}
.detail-page-header-row .btn-back { margin-bottom: 0; }

.page-subtitle {
  font-size: 13.5px;
  color: var(--text-secondary);
  margin: 4px 0 24px;
}

.page-count {
  font-size: 12.5px;
  font-weight: 600;
  color: var(--text-tertiary);
  margin: 4px 0 0;
}

.page-subtitle-tight {
  margin-bottom: 4px;
}

.page-meta {
  font-size: 12px;
  color: var(--text-tertiary);
  margin: 0 0 20px;
}

.form-grid {
  display: grid;
  grid-template-columns: 1fr 1fr;
  gap: 4px 20px;
  margin-bottom: 8px;
}

.srl-date-grid {
  display: grid;
  grid-template-columns: repeat(auto-fill, minmax(180px, 1fr));
  gap: 4px 20px;
}

.field-full {
  grid-column: 1 / -1;
}

.form-grid-pair {
  display: grid;
  grid-template-columns: 1fr 1fr;
  gap: 4px 20px;
}

@media (max-width: 480px) {
  .form-grid-pair {
    grid-template-columns: 1fr;
  }
}

@media (max-width: 600px) {
  .dashboard-header {
    padding: 14px 16px;
  }
  .dashboard-main {
    padding: 24px 14px;
  }
  .dashboard-footer {
    padding: 16px 16px 20px;
  }
  .page-body,
  .page-body-companies {
    gap: 14px;
  }
  .add-tile,
  .page-body-companies .add-tile {
    width: 100%;
    max-width: none;
  }
  .home-overview {
    flex-direction: column;
  }
  .funnel-chart,
  .home-stats,
  .potential-column {
    flex-basis: auto;
    min-width: 0;
  }
  .form-actions {
    flex-wrap: wrap;
  }
  .form-actions .btn-primary,
  .form-actions .btn-secondary {
    flex: 1 1 auto;
  }
}

@media (max-width: 480px) {
  .form-grid {
    grid-template-columns: 1fr;
  }
}

.select {
  width: 100%;
  font-family: inherit;
  font-size: 14px;
  color: var(--text-primary);
  background: var(--bg);
  border: 1px solid var(--border-strong);
  border-radius: 7px;
  padding: 10px 12px;
  outline: none;
  cursor: pointer;
  transition: border-color 120ms ease, box-shadow 120ms ease;
}
.select:focus {
  border-color: var(--accent);
  box-shadow: 0 0 0 3px rgba(15, 107, 92, 0.12);
}
.select-error {
  border-color: var(--danger);
}

.form-actions {
  display: flex;
  gap: 12px;
  margin-top: 12px;
}

.form-section {
  margin-top: 28px;
  padding-top: 20px;
  border-top: 1px solid var(--border);
}

.form-section-title {
  font-size: 15px;
  font-weight: 650;
  color: var(--text-primary);
  margin: 0 0 4px;
}

.form-section-subtitle {
  font-size: 12.5px;
  color: var(--text-tertiary);
  margin: 0 0 14px;
}

.srl-progress {
  display: flex;
  gap: 3px;
  margin-top: 16px;
  height: 30px;
}

.srl-progress-segment {
  flex: 1;
  display: flex;
  align-items: center;
  justify-content: center;
  background: var(--surface);
  border: 1px solid var(--border);
  border-radius: 4px;
  transition: background 120ms ease, border-color 120ms ease;
}

.srl-progress-segment-filled {
  background: var(--accent);
  border-color: var(--accent);
}

.srl-progress-segment-current {
  position: relative;
}

.srl-progress-label {
  font-family: "JetBrains Mono", monospace;
  font-size: 10.5px;
  font-weight: 600;
  color: #ffffff;
  white-space: nowrap;
}

.detail-grid {
  display: grid;
  grid-template-columns: 1fr 1fr;
  gap: 18px 24px;
  margin-top: 24px;
}

@media (max-width: 480px) {
  .detail-grid {
    grid-template-columns: 1fr;
  }
}

.additional-contacts-list {
  display: flex;
  flex-direction: column;
  gap: 14px;
}

.additional-contact-card {
  display: grid;
  grid-template-columns: 1fr 1fr;
  gap: 14px 24px;
  padding: 16px 18px;
  border: 1px solid var(--border);
  border-radius: 10px;
  background: var(--surface);
}

.additional-contact-card .detail-row {
  padding-bottom: 0;
  border-bottom: none;
}

@media (max-width: 480px) {
  .additional-contact-card {
    grid-template-columns: 1fr;
  }
}

.detail-row {
  display: flex;
  flex-direction: column;
  gap: 4px;
  padding-bottom: 14px;
  border-bottom: 1px solid var(--border);
}

.detail-label {
  font-size: 11px;
  text-transform: uppercase;
  letter-spacing: 0.04em;
  color: var(--text-tertiary);
  font-weight: 500;
}

.detail-value {
  font-size: 14px;
  color: var(--text-primary);
  font-weight: 500;
}

.company-link {
  display: inline-block;
  width: auto;
  font-family: inherit;
  background: none;
  border: none;
  padding: 0;
  text-align: left;
  cursor: pointer;
  color: var(--accent-dark);
  transition: color 120ms ease;
}
.company-link:hover {
  text-decoration: underline;
}

.button-row {
  display: flex;
  gap: 12px;
  margin-top: 24px;
  justify-content: center;
  flex-wrap: wrap;
}

.dashboard-bottom-row {
  margin-top: 32px;
  padding-top: 24px;
}

.home-top-projects {
  margin-top: 40px;
}

.btn-choice {
  font-family: inherit;
  font-size: 13.5px;
  font-weight: 600;
  color: var(--text-primary);
  background: var(--bg);
  border: 1px solid var(--border-strong);
  border-radius: 7px;
  padding: 10px 18px;
  cursor: pointer;
  text-decoration: none;
  transition: background 120ms ease, border-color 120ms ease, color 120ms ease;
}
.btn-choice:hover {
  background: var(--surface);
  border-color: var(--accent);
  color: var(--accent-dark);
}

.btn-choice-report {
  background: var(--accent);
  border-color: var(--accent);
  color: #ffffff;
}
.btn-choice-report:hover {
  background: var(--accent-dark);
  border-color: var(--accent-dark);
  color: #ffffff;
}

/* ---------- Generate Report ---------- */
.report-builder {
  display: flex;
  flex-direction: column;
  gap: 20px;
  max-width: 640px;
}

.report-section-block {
  background: #ffffff;
  border: 1px solid var(--border);
  border-radius: 12px;
  padding: 18px 20px;
}

.report-section-block-header {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 12px;
  margin-bottom: 4px;
}

.report-checkbox-grid {
  display: grid;
  grid-template-columns: repeat(2, 1fr);
  gap: 10px 16px;
  margin-top: 12px;
}

.report-checkbox-grid-levels {
  grid-template-columns: repeat(4, 1fr);
}

.report-checkbox {
  display: flex;
  align-items: center;
  gap: 8px;
  font-size: 13.5px;
  color: var(--text-primary);
  cursor: pointer;
}
.report-checkbox input {
  width: 16px;
  height: 16px;
  accent-color: var(--accent);
  cursor: pointer;
}

.import-row {
  display: flex;
  align-items: center;
  gap: 12px;
  margin-top: 28px;
  padding-top: 20px;
  border-top: 1px solid var(--border);
}

.btn-import {
  display: inline-flex;
  align-items: center;
  gap: 6px;
  font-family: inherit;
  font-size: 12.5px;
  font-weight: 600;
  color: var(--text-secondary);
  background: var(--bg);
  border: 1px solid var(--border);
  border-radius: 7px;
  padding: 7px 12px;
  cursor: pointer;
  transition: background 120ms ease, border-color 120ms ease, color 120ms ease;
}
.btn-import:hover {
  background: var(--surface);
  border-color: var(--accent);
  color: var(--accent-dark);
}

.btn-add-contact {
  display: inline-flex;
  align-items: center;
  gap: 6px;
  flex-shrink: 0;
  white-space: nowrap;
  font-family: inherit;
  font-size: 13px;
  font-weight: 600;
  color: #ffffff;
  background: var(--ink);
  border: 1px solid var(--ink);
  border-radius: 7px;
  padding: 9px 14px;
  cursor: pointer;
  transition: background 120ms ease, transform 60ms ease;
}
.btn-add-contact:hover { background: #262b34; }
.btn-add-contact:active { transform: translateY(1px); }

.import-file-input {
  display: none;
}

.import-status {
  font-size: 12.5px;
  color: var(--text-secondary);
}

/* ---------- Debug panel (dev only) ---------- */
.debug-panel {
  position: fixed;
  right: 16px;
  bottom: 16px;
  z-index: 50;
  display: flex;
  flex-direction: column;
  align-items: flex-end;
  font-family: "JetBrains Mono", monospace;
}

.debug-toggle {
  font-family: inherit;
  font-size: 11px;
  font-weight: 500;
  color: var(--text-secondary);
  background: var(--surface);
  border: 1px solid var(--border-strong);
  border-radius: 999px;
  padding: 7px 14px;
  cursor: pointer;
  box-shadow: 0 1px 3px rgba(0, 0, 0, 0.08);
}
.debug-toggle:hover { color: var(--ink); border-color: var(--accent); }

.debug-body {
  margin-bottom: 8px;
  width: 320px;
  max-height: 260px;
  overflow: auto;
  background: var(--bg);
  border: 1px solid var(--border-strong);
  border-radius: 8px;
  padding: 12px;
  box-shadow: 0 8px 24px rgba(0, 0, 0, 0.12);
}

.debug-note {
  font-size: 10px;
  color: var(--text-tertiary);
  text-transform: uppercase;
  letter-spacing: 0.03em;
  margin: 0 0 10px;
}

.debug-empty {
  font-size: 11px;
  color: var(--text-tertiary);
  margin: 0;
}

.debug-table {
  width: 100%;
  border-collapse: collapse;
  font-size: 10.5px;
}
.debug-table th {
  text-align: left;
  font-weight: 500;
  color: var(--text-tertiary);
  padding: 4px 6px;
  border-bottom: 1px solid var(--border);
}
.debug-table td {
  padding: 4px 6px;
  border-bottom: 1px solid var(--border);
  color: var(--text-primary);
  word-break: break-all;
}

/* ---------- Prefixed inputs (e.g. Funding "$") ---------- */
.input-wrap-prefixed .input-prefix {
  position: absolute;
  left: 12px;
  font-size: 14px;
  font-weight: 500;
  color: var(--text-secondary);
  pointer-events: none;
}
.input-wrap-prefixed .input {
  padding-left: 26px;
}

.field-hint {
  margin-top: 6px;
  font-size: 12px;
  color: var(--text-tertiary);
}

.company-mode-toggle {
  display: inline-flex;
  gap: 4px;
  padding: 3px;
  background: var(--surface);
  border: 1px solid var(--border);
  border-radius: 9px;
  margin-bottom: 12px;
}

.company-mode-btn {
  font-family: inherit;
  font-size: 12.5px;
  font-weight: 600;
  color: var(--text-secondary);
  background: transparent;
  border: none;
  border-radius: 6px;
  padding: 7px 14px;
  cursor: pointer;
  transition: background 120ms ease, color 120ms ease;
}
.company-mode-btn:hover { color: var(--text-primary); }
.company-mode-btn-active {
  background: #ffffff;
  color: var(--accent-dark);
  box-shadow: 0 1px 2px rgba(15, 65, 130, 0.08);
}

.new-company-fields {
  display: flex;
  flex-direction: column;
  gap: 14px;
  padding: 16px;
  background: var(--surface);
  border: 1px solid var(--border);
  border-radius: 10px;
}

.textarea {
  width: 100%;
  font-family: inherit;
  font-size: 14px;
  color: var(--text-primary);
  background: var(--bg);
  border: 1px solid var(--border-strong);
  border-radius: 7px;
  padding: 10px 12px;
  outline: none;
  resize: vertical;
  min-height: 100px;
  transition: border-color 120ms ease, box-shadow 120ms ease;
}
.textarea::placeholder { color: var(--text-tertiary); }
.textarea:focus {
  border-color: var(--accent);
  box-shadow: 0 0 0 3px rgba(15, 107, 92, 0.12);
}

/* ---------- Company detail: project tables ---------- */
.table-block { margin-top: 28px; }

.update-log { margin-top: 32px; }

.update-log-table td { font-size: 12.5px; }
.update-log-row-a { background: #ffffff; }
.update-log-row-b { background: var(--surface); }
.update-log-time { color: var(--text-tertiary); white-space: nowrap; }

.srl-progression-table td {
  font-size: 12.5px;
  vertical-align: top;
}
.srl-progression-table td:nth-child(3) {
  white-space: nowrap;
  color: var(--text-secondary);
}

.srl-note-cell {
  min-width: 280px;
}

.srl-note-input {
  width: 100%;
  font-family: inherit;
  font-size: 12.5px;
  color: var(--text-primary);
  background: var(--bg);
  border: 1px solid var(--border);
  border-radius: 6px;
  padding: 8px 10px;
  outline: none;
  resize: vertical;
  min-height: 84px;
  line-height: 1.45;
  transition: border-color 120ms ease, box-shadow 120ms ease;
}
.srl-note-input::placeholder { color: var(--text-tertiary); }
.srl-note-input:focus {
  border-color: var(--accent);
  box-shadow: 0 0 0 3px rgba(15, 107, 92, 0.12);
}

.srl-note-meta {
  margin: 4px 0 0;
  font-size: 11px;
  color: var(--text-tertiary);
}

.section-title {
  font-size: 14px;
  font-weight: 600;
  color: var(--ink);
  margin: 0 0 10px;
}

.empty-table-note {
  font-size: 12.5px;
  color: var(--text-tertiary);
  margin: 0;
}

.table-wrap {
  border: 1px solid var(--border);
  border-radius: 8px;
  overflow-x: auto;
  overflow-y: hidden;
  -webkit-overflow-scrolling: touch;
}

.table-wrap-scroll {
  max-height: 300px;
  overflow-y: auto;
  overflow-x: auto;
  -webkit-overflow-scrolling: touch;
}

.srl-progression-scroll {
  max-height: 420px;
}

.table-wrap-scroll .data-table thead th {
  position: sticky;
  top: 0;
  z-index: 1;
}

.data-table {
  width: 100%;
  min-width: 520px;
  border-collapse: collapse;
  font-size: 13px;
}

.data-table th {
  text-align: left;
  font-size: 11px;
  font-weight: 600;
  text-transform: uppercase;
  letter-spacing: 0.03em;
  color: var(--text-tertiary);
  background: var(--surface);
  padding: 9px 14px;
  border-bottom: 1px solid var(--border);
}

.data-table td {
  padding: 10px 14px;
  border-bottom: 1px solid var(--border);
  color: var(--text-primary);
}

.data-row { cursor: pointer; transition: background 120ms ease; }
.data-row:hover { background: var(--surface); }
.data-row:last-child td { border-bottom: none; }

/* ---------- Home page ---------- */
.home-content {
  width: 100%;
  background: linear-gradient(180deg, #eaf4ff 0%, #f5faff 100%);
  border: 1px solid #d7ebff;
  border-radius: 16px;
  padding: 32px;
  margin: 0 0 24px;
}

.home-welcome {
  text-align: center;
}

.potential-column {
  flex: 1 1 260px;
  display: flex;
  flex-direction: column;
  gap: 20px;
}

.potential-chart-card {
  display: flex;
  flex-direction: column;
  gap: 10px;
  padding: 20px;
  border: 1px solid var(--border);
  border-radius: 10px;
  background: var(--bg);
}

.potential-chart-wrap {
  width: 100%;
}

.potential-chart-legend {
  max-height: 110px;
  overflow-y: auto;
  display: flex;
  flex-direction: column;
  gap: 5px;
  padding-right: 4px;
  -webkit-overflow-scrolling: touch;
}

.potential-chart-legend-item {
  display: flex;
  align-items: flex-start;
  gap: 6px;
  font-size: 10.5px;
  line-height: 1.3;
  color: var(--text-secondary);
}

.potential-chart-legend-dot {
  flex: 0 0 auto;
  width: 7px;
  height: 7px;
  border-radius: 50%;
  margin-top: 3px;
}

.potential-chart-legend-label {
  overflow-wrap: anywhere;
}

.pie-tooltip {
  padding: 6px 10px;
  border: 1px solid var(--border);
  border-radius: 6px;
  background: var(--bg);
  box-shadow: 0 2px 8px rgba(0, 0, 0, 0.08);
  font-size: 11px;
  line-height: 1.4;
}

.pie-tooltip-funding {
  font-weight: 700;
  color: var(--accent-dark);
}

.pie-tooltip-project {
  color: var(--text-secondary);
}

.list-buttons-card {
  flex: 1;
  display: flex;
  flex-direction: column;
  gap: 10px;
  padding: 12px;
  border: 1px solid var(--border);
  border-radius: 10px;
  background: var(--bg);
}

.list-buttons-card .btn-choice {
  flex: 1;
  display: flex;
  align-items: center;
  justify-content: center;
  width: 100%;
  min-height: 44px;
}

.btn-list-companies {
  background: #0f467f;
  border-color: #0f467f;
  color: #ffffff;
}
.btn-list-companies:hover {
  background: #0a3560;
  border-color: #0a3560;
  color: #ffffff;
}

.btn-list-contacts {
  background: #bb8311;
  border-color: #bb8311;
  color: #ffffff;
}
.btn-list-contacts:hover {
  background: #96690d;
  border-color: #96690d;
  color: #ffffff;
}

.btn-list-projects {
  background: var(--accent);
  border-color: var(--accent);
  color: #ffffff;
}
.btn-list-projects:hover {
  background: var(--accent-dark);
  border-color: var(--accent-dark);
  color: #ffffff;
}

/* ---------- Home overview: funnel chart + summary cards ---------- */
.home-overview {
  display: flex;
  align-items: stretch;
  gap: 20px;
  margin: 28px 0 8px;
  flex-wrap: wrap;
}

.funnel-chart {
  flex: 1.4 1 340px;
  display: flex;
  flex-direction: column;
  gap: 10px;
  margin: 0;
  padding: 20px;
  border: 1px solid var(--border);
  border-radius: 10px;
  background: var(--bg);
}

.home-stats {
  flex: 1 1 220px;
  display: flex;
  flex-direction: column;
  gap: 12px;
}

.stat-card {
  flex: 1;
  display: flex;
  flex-direction: column;
  justify-content: center;
  gap: 4px;
  padding: 16px 18px;
  border: 1px solid var(--border);
  border-left: 3px solid var(--accent);
  border-radius: 10px;
  background: var(--success-bg);
}

.stat-label {
  font-size: 11.5px;
  font-weight: 600;
  text-transform: uppercase;
  letter-spacing: 0.03em;
  color: var(--text-secondary);
}

.stat-value {
  font-size: 24px;
  font-weight: 700;
  color: var(--accent-dark);
  letter-spacing: -0.01em;
}

.stat-sub {
  font-size: 11.5px;
  color: var(--text-tertiary);
}

.funnel-row {
  display: flex;
  align-items: center;
  gap: 14px;
}

.funnel-label {
  width: 54px;
  flex-shrink: 0;
  font-size: 12.5px;
  font-weight: 500;
  color: var(--text-secondary);
}

.funnel-track {
  flex: 1;
  min-width: 0;
  display: flex;
  justify-content: center;
  background: var(--surface);
  border-radius: 6px;
}

.funnel-bar {
  display: flex;
  align-items: center;
  justify-content: center;
  min-width: 34px;
  height: 32px;
  background: #6fcf97;
  border-radius: 6px;
  transition: width 300ms ease;
}

.funnel-bar-srl-1 { background: #2f9e5c; }
.funnel-bar-srl-2 { background: #2f9e5c; }
.funnel-bar-srl-3 { background: #2f9e5c; }
.funnel-bar-srl-4 { background: #2f9e5c; }
.funnel-bar-srl-5 { background: #2f9e5c; }
.funnel-bar-srl-6 { background: #2f9e5c; }
.funnel-bar-srl-7 { background: #2f9e5c; }

.funnel-bar-empty {
  background: var(--surface);
  border: 1px dashed var(--border-strong);
}

.funnel-count {
  font-size: 13px;
  font-weight: 600;
  color: #ffffff;
}

.funnel-bar-empty .funnel-count {
  color: var(--text-tertiary);
}

.funnel-percent {
  width: 36px;
  flex-shrink: 0;
  text-align: right;
  font-size: 12px;
  color: var(--text-tertiary);
}

.funnel-key {
  margin-top: 16px;
  padding-top: 14px;
  border-top: 1px solid var(--border);
}

.funnel-key-title {
  margin: 0 0 8px;
  font-size: 12px;
  font-weight: 700;
  text-transform: uppercase;
  letter-spacing: 0.03em;
  color: var(--text-tertiary);
}

.funnel-key-list {
  list-style: none;
  margin: 0;
  padding: 0;
  display: flex;
  flex-direction: column;
  gap: 6px;
}

.funnel-key-item {
  display: flex;
  align-items: flex-start;
  gap: 8px;
  font-size: 12.5px;
  line-height: 1.5;
  color: var(--text-secondary);
}

.funnel-key-dot {
  flex-shrink: 0;
  width: 10px;
  height: 10px;
  margin-top: 4px;
  border-radius: 50%;
}

.funnel-key-text strong {
  color: var(--text-primary);
  font-weight: 600;
}

.funnel-key-divider {
  margin: 0 6px;
  color: var(--border-strong);
  font-weight: 400;
}

.funnel-key-desc {
  font-weight: 400;
  color: var(--text-secondary);
}
`;
