import { doc, getDoc, runTransaction } from "firebase/firestore";
import { db } from "../firebase";

const APP_DATA_COLLECTION = "appData";
const APP_DATA_DOCUMENT = "main";
const APP_DATA_SCHEMA_VERSION = 2;

function cleanForFirestore(value) {
  // Firestore rejects `undefined`. The app state is JSON-shaped, so a JSON
  // round-trip is a deterministic sanitizer for projects/companies/contacts.
  return JSON.parse(JSON.stringify(value ?? null));
}

function recordTimestamp(record) {
  const updated = Number(record?.updatedAt || 0);
  const created = Number(record?.createdAt || 0);
  return Number.isFinite(updated) && updated > 0
    ? updated
    : Number.isFinite(created) && created > 0
      ? created
      : 0;
}

function mergeRecords(serverRecords, incomingRecords) {
  const merged = new Map();
  (Array.isArray(serverRecords) ? serverRecords : []).forEach((record) => {
    if (record && record.id != null) merged.set(String(record.id), record);
  });

  (Array.isArray(incomingRecords) ? incomingRecords : []).forEach((record) => {
    if (!record || record.id == null) return;
    const key = String(record.id);
    const existing = merged.get(key);
    if (!existing || recordTimestamp(record) >= recordTimestamp(existing)) {
      merged.set(key, record);
    }
  });

  return Array.from(merged.values());
}

export async function loadAppData() {
  try {
    const snapshot = await getDoc(doc(db, APP_DATA_COLLECTION, APP_DATA_DOCUMENT));
    if (!snapshot.exists()) {
      return {
        ok: true,
        exists: false,
        projects: [],
        companies: [],
        contacts: [],
      };
    }

    const data = snapshot.data() || {};
    return {
      ok: true,
      exists: true,
      projects: Array.isArray(data.projects) ? data.projects : [],
      companies: Array.isArray(data.companies) ? data.companies : [],
      contacts: Array.isArray(data.contacts) ? data.contacts : [],
      updatedAt: data.updatedAt || 0,
      updatedBy: data.updatedBy || "",
      schemaVersion: data.schemaVersion || 0,
    };
  } catch (error) {
    return {
      ok: false,
      exists: false,
      projects: [],
      companies: [],
      contacts: [],
      error: String(error?.message || error || "Could not load Firestore app data."),
    };
  }
}

// Merge by stable record ID inside a Firestore transaction. This is important
// when two browsers are open: a stale browser can no longer replace the whole
// projects array and accidentally erase a project saved by another browser.
export async function saveAppData({ projects, companies, contacts, updatedBy = "" }) {
  try {
    const ref = doc(db, APP_DATA_COLLECTION, APP_DATA_DOCUMENT);
    const result = await runTransaction(db, async (transaction) => {
      const snapshot = await transaction.get(ref);
      const existing = snapshot.exists() ? (snapshot.data() || {}) : {};

      const mergedProjects = mergeRecords(existing.projects, projects);
      const mergedCompanies = mergeRecords(existing.companies, companies);
      const mergedContacts = mergeRecords(existing.contacts, contacts);
      const payload = cleanForFirestore({
        schemaVersion: APP_DATA_SCHEMA_VERSION,
        projects: mergedProjects,
        companies: mergedCompanies,
        contacts: mergedContacts,
        updatedAt: Date.now(),
        updatedBy: String(updatedBy || ""),
      });

      transaction.set(ref, payload, { merge: false });
      return payload;
    });

    return {
      ok: true,
      projects: result.projects || [],
      companies: result.companies || [],
      contacts: result.contacts || [],
      updatedAt: result.updatedAt || 0,
    };
  } catch (error) {
    return {
      ok: false,
      error: String(error?.message || error || "Could not save Firestore app data."),
    };
  }
}

// Deletions are explicit. Normal saves intentionally preserve server-only
// records so another device cannot delete data merely because its local copy is
// old. Only these explicit IDs are removed.
export async function deleteAppDataRecords({ projectIds = [], companyIds = [], contactIds = [], updatedBy = "" }) {
  try {
    const projectSet = new Set((projectIds || []).map(String));
    const companySet = new Set((companyIds || []).map(String));
    const contactSet = new Set((contactIds || []).map(String));
    const ref = doc(db, APP_DATA_COLLECTION, APP_DATA_DOCUMENT);

    const result = await runTransaction(db, async (transaction) => {
      const snapshot = await transaction.get(ref);
      const existing = snapshot.exists() ? (snapshot.data() || {}) : {};
      const payload = cleanForFirestore({
        schemaVersion: APP_DATA_SCHEMA_VERSION,
        projects: (Array.isArray(existing.projects) ? existing.projects : []).filter(
          (record) => !projectSet.has(String(record?.id ?? "")),
        ),
        companies: (Array.isArray(existing.companies) ? existing.companies : []).filter(
          (record) => !companySet.has(String(record?.id ?? "")),
        ),
        contacts: (Array.isArray(existing.contacts) ? existing.contacts : []).filter(
          (record) => !contactSet.has(String(record?.id ?? "")),
        ),
        updatedAt: Date.now(),
        updatedBy: String(updatedBy || ""),
      });
      transaction.set(ref, payload, { merge: false });
      return payload;
    });

    return { ok: true, ...result };
  } catch (error) {
    return {
      ok: false,
      error: String(error?.message || error || "Could not delete Firestore app data."),
    };
  }
}
