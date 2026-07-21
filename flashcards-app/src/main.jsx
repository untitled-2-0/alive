import React from "react";
import { createRoot } from "react-dom/client";
import FlashcardsApp from "./FlashcardsApp.jsx";
import { initCloud, refreshSession, isSignedIn, pullToLocal, mergeCloud } from "./cloud.js";

/*
 * window.storage shim — backs the app's storage with localStorage.
 * On a claude.ai artifact window.storage is provided by the runtime; here (and
 * on the deployed site) we back it with localStorage so data persists locally.
 */
if (!window.storage) {
  const backing = window.localStorage;
  window.storage = {
    async getItem(key) { return backing.getItem(`recall:${key}`); },
    async setItem(key, value) { backing.setItem(`recall:${key}`, value); },
    async removeItem(key) { backing.removeItem(`recall:${key}`); },
  };
}

const render = () =>
  createRoot(document.getElementById("root")).render(
    <React.StrictMode>
      <FlashcardsApp />
    </React.StrictMode>
  );

// Boot: if a Supabase session exists, pull cloud → local before first paint so
// the app renders the latest synced data. Always render, even if cloud fails.
(async () => {
  // Did we arrive from a magic-link? (token in the URL hash, before the client clears it)
  const fromMagicLink = typeof window !== "undefined" && /[#&](access_token|error)=/.test(window.location.hash || "");
  try {
    initCloud();
    await refreshSession();
    if (isSignedIn()) {
      try {
        // First sign-in via magic-link → two-way merge (upload local + pull cloud),
        // so nothing on this device is stranded. Normal boot → just pull.
        if (fromMagicLink) await mergeCloud(); else await pullToLocal();
      } catch (e) { console.warn("[cloud pull]", e?.message || e); }
    }
  } catch (e) {
    console.warn("[cloud init]", e?.message || e);
  }
  render();
})();
