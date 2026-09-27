// Same function names/signatures App.jsx already calls everywhere
// (loadProfile, saveProfile, ...), so App.jsx's *usage* of them doesn't
// change at all. Only their storage backend changes: IndexedDB via
// store.js, instead of localStorage. See MIGRATION_GUIDE.md for the two
// small edits App.jsx itself needs (delete the old definitions of these
// same names, add one import line).
export {
  loadProfile, saveProfile, clearProfileStorage,
  loadCrewNames, saveCrewNames,
  loadThemePrefs, saveThemePrefs,
  loadDailyLogEntries, saveDailyLogEntries,
  loadDailyLogAccess, saveDailyLogAccess,
  loadLastFile, saveLastFile, clearLastFileStorage,
  loadSoundSettings, saveSoundSettings,
  loadLang, saveLang,
  ready, isReadOnly, readOnlyReasonCode, getInitReport, subscribe,
} from "./store.js";

// loadAdminSession/saveAdminSession are NOT re-exported here on purpose —
// the admin-unlock flag is deliberately sessionStorage (see
// ADMIN_SESSION_KEY in App.jsx) so it clears itself when the tab closes.
// It isn't user data and was never part of the migration; leave App.jsx's
// existing loadAdminSession/saveAdminSession functions exactly as they are.
