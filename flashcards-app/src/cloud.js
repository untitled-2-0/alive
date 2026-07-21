// Cloud sync via Supabase — optional, opt-in, key-value mirror of window.storage.
// The anon key is public by design (protected by Row-Level-Security on the kv table).
import { createClient } from "@supabase/supabase-js";

const SUPABASE_URL = "https://nczwmbezbxmvwrabmrju.supabase.co";
const SUPABASE_ANON = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Im5jendtYmV6YnhtdndyYWJtcmp1Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODQ2MzcxNzIsImV4cCI6MjEwMDIxMzE3Mn0.NkjwuQVL4rbgjcI3KZ6p8AIs9Np-dRK7HyEcqqVp5ww";
const PREFIX = "recall:"; // window.storage shim prefixes local keys with this

let supabase = null;
let user = null;
let accessToken = null; // cached so the unload flush can fire without an async getSession()

export function initCloud() {
  if (!supabase) {
    supabase = createClient(SUPABASE_URL, SUPABASE_ANON, {
      // detectSessionInUrl: true → if a magic-link lands back on the app with
      // #access_token=… in the URL, the client parses it and establishes the session.
      auth: { persistSession: true, autoRefreshToken: true, detectSessionInUrl: true },
    });
    // Keep user + token fresh across token refreshes so the keepalive unload
    // flush always has a valid bearer token to send.
    supabase.auth.onAuthStateChange((_event, session) => {
      user = session?.user || null;
      accessToken = session?.access_token || null;
    });
  }
  return supabase;
}

export async function refreshSession() {
  if (!supabase) initCloud();
  try {
    const { data } = await supabase.auth.getSession();
    user = data?.session?.user || null;
    accessToken = data?.session?.access_token || null;
  } catch { user = null; accessToken = null; }
  return user;
}

export const isSignedIn = () => !!user;
export const currentEmail = () => user?.email || null;

/* ---- pull cloud → local (writes straight into localStorage) ---- */
export async function pullToLocal() {
  if (!user) return 0;
  const { data, error } = await supabase.from("kv").select("key,value");
  if (error) throw error;
  for (const row of data || []) {
    try { localStorage.setItem(PREFIX + row.key, JSON.stringify(row.value)); } catch (e) { /* ignore */ }
  }
  return (data || []).length;
}

/* ---- push all local recall:* keys → cloud ---- */
export async function pushAllLocal() {
  if (!user) return 0;
  const rows = collectLocalRows();
  // upsert in chunks to stay well within limits
  for (let i = 0; i < rows.length; i += 200) {
    const { error } = await supabase.from("kv").upsert(rows.slice(i, i + 200));
    if (error) throw error;
  }
  return rows.length;
}

function collectLocalRows(skipKeys) {
  const rows = [];
  for (let i = 0; i < localStorage.length; i++) {
    const k = localStorage.key(i);
    if (!k || !k.startsWith(PREFIX)) continue;
    const key = k.slice(PREFIX.length);
    if (skipKeys && skipKeys.has(key)) continue;
    let value;
    try { value = JSON.parse(localStorage.getItem(k)); } catch { value = localStorage.getItem(k); }
    rows.push({ user_id: user.id, key, value });
  }
  return rows;
}

/* ---- debounced dual-write, with requeue-on-failure so a transient network
       error never silently drops a write from the sync queue ---- */
const pending = new Map();
let flushTimer = null;
let backoff = 1200;
function schedule(delay) { if (!flushTimer) flushTimer = setTimeout(flush, delay ?? 1200); }
async function flush() {
  flushTimer = null;
  if (!user || !supabase || !pending.size) return;
  // Snapshot then clear so concurrent cloudPush() calls during the await land
  // in a fresh queue; only truly-failed keys get requeued below.
  const snapshot = new Map(pending);
  pending.clear();
  const upserts = [], deletes = [];
  for (const [k, v] of snapshot) { if (v && v.__deleted) deletes.push(k); else upserts.push({ user_id: user.id, key: k, value: v }); }
  try {
    if (upserts.length) { const { error } = await supabase.from("kv").upsert(upserts); if (error) throw error; }
    for (const k of deletes) { const { error } = await supabase.from("kv").delete().eq("key", k); if (error) throw error; }
    backoff = 1200; // success → reset backoff
  } catch (e) {
    console.warn("[cloud flush]", e?.message || e);
    // Requeue every key that wasn't superseded by a newer write while we awaited,
    // then retry with exponential backoff (capped) instead of losing the writes.
    for (const [k, v] of snapshot) { if (!pending.has(k)) pending.set(k, v); }
    backoff = Math.min(backoff * 2, 30000);
    schedule(backoff);
  }
}
export function cloudPush(key, value) { if (!user) return; pending.set(key, value); schedule(); }
export function cloudRemove(key) { if (!user) return; pending.set(key, { __deleted: true }); schedule(); }

/* ---- last-ditch flush on tab close / backgrounding ----
   The normal flush() uses supabase-js fetch, which the browser cancels on
   unload. keepalive:true survives unload so the last ~1.2s of edits reach the
   DB. visibilitychange('hidden') + pagehide are the reliable mobile signals. */
function flushKeepalive() {
  if (!user || !accessToken || !pending.size) return;
  const snapshot = new Map(pending);
  pending.clear();
  const upserts = [], deletes = [];
  for (const [k, v] of snapshot) { if (v && v.__deleted) deletes.push(k); else upserts.push({ user_id: user.id, key: k, value: v }); }
  const base = `${SUPABASE_URL}/rest/v1/kv`;
  const headers = { apikey: SUPABASE_ANON, Authorization: `Bearer ${accessToken}`, "Content-Type": "application/json" };
  try {
    if (upserts.length) {
      fetch(base, { method: "POST", keepalive: true, headers: { ...headers, Prefer: "resolution=merge-duplicates,return=minimal" }, body: JSON.stringify(upserts) })
        .catch(() => { for (const [k, v] of snapshot) if (!pending.has(k)) pending.set(k, v); });
    }
    for (const k of deletes) {
      fetch(`${base}?key=eq.${encodeURIComponent(k)}`, { method: "DELETE", keepalive: true, headers }).catch(() => {});
    }
  } catch { for (const [k, v] of snapshot) if (!pending.has(k)) pending.set(k, v); }
}
if (typeof window !== "undefined") {
  window.addEventListener("pagehide", flushKeepalive);
  window.addEventListener("beforeunload", flushKeepalive);
  window.addEventListener("visibilitychange", () => { if (document.visibilityState === "hidden") flushKeepalive(); });
}

/* ---- auth (email OTP code) ---- */
export async function sendCode(email) {
  if (!supabase) initCloud();
  // emailRedirectTo makes the magic-link in the email come back to THIS app (not
  // whatever the Supabase "Site URL" happens to be) — as long as the origin is an
  // allowed redirect URL in the Supabase project.
  const redirect = typeof window !== "undefined" ? window.location.origin : undefined;
  const { error } = await supabase.auth.signInWithOtp({ email: email.trim(), options: { shouldCreateUser: true, emailRedirectTo: redirect } });
  if (error) throw error;
}

/* Two-way sign-in merge: pull the cloud down (cloud wins on shared keys), and
   upload every local-only key so nothing on this device is stranded off the DB.
   For a brand-new account the cloud is empty, so this uploads everything. */
export async function mergeCloud() {
  if (!user) return { pulled: 0, pushed: 0 };
  const { data, error } = await supabase.from("kv").select("key,value");
  if (error) throw error;
  const cloudKeys = new Set();
  for (const row of data || []) {
    cloudKeys.add(row.key);
    try { localStorage.setItem(PREFIX + row.key, JSON.stringify(row.value)); } catch { /* ignore */ }
  }
  const rows = collectLocalRows(cloudKeys); // local keys the cloud lacks
  for (let i = 0; i < rows.length; i += 200) {
    const { error: e2 } = await supabase.from("kv").upsert(rows.slice(i, i + 200));
    if (e2) throw e2;
  }
  return { pulled: cloudKeys.size, pushed: rows.length };
}

// After verifying the code, merge local ⇄ cloud. Then the caller reloads.
export async function verifyCode(email, token) {
  const { data, error } = await supabase.auth.verifyOtp({ email: email.trim(), token: token.trim(), type: "email" });
  if (error) throw error;
  user = data.user;
  accessToken = data.session?.access_token || accessToken;
  const r = await mergeCloud();
  return { email: user.email, merged: r.pulled ? "merged" : "pushed" };
}

export async function signOutCloud() {
  if (!supabase) return;
  try { await supabase.auth.signOut(); } catch { /* ignore */ }
  user = null;
  accessToken = null;
}

// Force a full two-way sync: push local up, then pull remote down.
export async function syncNow() {
  if (!user) return { ok: false };
  await flush();
  await pushAllLocal();
  await pullToLocal();
  return { ok: true };
}
