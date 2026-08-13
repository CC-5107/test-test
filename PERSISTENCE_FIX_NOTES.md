# Persistence fix

This build fixes the write/reconciliation race without changing the original Google Sheet layout or formatting.

- User edits are marked pending in local state before any network await.
- Firestore writes merge records transactionally by stable ID, so another browser cannot erase newer records by saving a stale full-array snapshot.
- Google Sheet summary polling is inventory-only while detail tabs are loading; it cannot roll edited values back.
- When detail tabs are available, editable detail-tab values are authoritative over lagging summary formulas (SRL, funding, RFS, etc.).
- Google Sheet reads do not trigger a Firestore autosave of sheet-reconciled state.
- Pending edit retries update an existing tab; they do not use the new-project creation path.
- Additional contacts are updated in the website-owned rows 25+ without changing formatting or original row/column positions.
- Firebase deletions are explicit so transaction merging does not resurrect deleted app records.
- Authenticated pages now show Firebase/Google sync errors instead of hiding them until logout.

No changes were made to projectSheets.server.ts, Google service-account authentication, spreadsheet IDs, or spreadsheet formatting commands.
