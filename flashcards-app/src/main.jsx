import React from "react";
import { createRoot } from "react-dom/client";
import FlashcardsApp from "./FlashcardsApp.jsx";

/*
 * window.storage shim — LOCAL PREVIEW ONLY.
 * In a claude.ai artifact, window.storage is provided by the runtime, so this
 * block is skipped. Here we back it with localStorage so persistence works
 * across reloads while you test.
 */
if (!window.storage) {
  const backing = window.localStorage;
  window.storage = {
    async getItem(key) {
      return backing.getItem(`recall:${key}`);
    },
    async setItem(key, value) {
      backing.setItem(`recall:${key}`, value);
    },
    async removeItem(key) {
      backing.removeItem(`recall:${key}`);
    },
  };
}

createRoot(document.getElementById("root")).render(
  <React.StrictMode>
    <FlashcardsApp />
  </React.StrictMode>
);
