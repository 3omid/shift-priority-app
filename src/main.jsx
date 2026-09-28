import React from "react";
import ReactDOM from "react-dom/client";
import App from "./App.jsx";
import { init } from "./data/store.js";
import { checkPeriodicSnapshot } from "./data/backup.js";

// The app's own components read profile/crewNames/dailyLog/etc. synchronously
// from an in-memory cache (see store.js) — that cache has to be hydrated
// from IndexedDB (and, on someone's very first run of this version, migrated
// from their old localStorage data) before React ever mounts, or the app's
// first render would flash empty defaults.
init().then((report) => {
  if (report?.migrationError) {
    // Migration failed — store.js has already fallen back to showing this
    // person's existing data read-only. Surface it in the console for now;
    // App.jsx can read Storage.getInitReport() to show a banner if desired.
    console.error("Shift Priority: data migration failed, running read-only.", report.migrationError);
  }
  ReactDOM.createRoot(document.getElementById("root")).render(
    <React.StrictMode>
      <App />
    </React.StrictMode>
  );
  checkPeriodicSnapshot();
});
