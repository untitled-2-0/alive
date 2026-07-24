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

/* ---- upsert rows in byte-sized batches; continue past a failed batch so one
       oversized value (e.g. a card photo) can't strand everything else ---- */
async function upsertRows(rows) {
  const MAX_BYTES = 700000; // keep each request comfortably under API limits
  let ok = 0, batch = [], size = 0;
  const send = async () => {
    if (!batch.length) return;
    try { const { error } = await supabase.from("kv").upsert(batch); if (error) throw error; ok += batch.length; }
    catch (e) { console.warn("[cloud upsert batch]", e?.message || e); }
    batch = []; size = 0;
  };
  for (const r of rows) {
    let vlen = 0; try { vlen = JSON.stringify(r.value ?? "").length; } catch { vlen = 0; }
    const bytes = (r.key.length + vlen) * 2;
    if (batch.length && size + bytes > MAX_BYTES) await send();
    batch.push(r); size += bytes;
    if (bytes > MAX_BYTES) await send(); // an oversized single value: send it alone
  }
  await send();
  return ok;
}

/* ---- push all local recall:* keys (incl. image:* photos) → cloud ---- */
export async function pushAllLocal() {
  if (!user) return 0;
  return await upsertRows(collectLocalRows());
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
  const rows = collectLocalRows(cloudKeys); // local keys the cloud lacks (incl. photos)
  const pushed = await upsertRows(rows);
  return { pulled: cloudKeys.size, pushed };
}

// After verifying the code, merge local ⇄ cloud. Then the caller reloads.
export async function verifyCode(email, token) {
  const e = email.trim(), t = token.trim();
  // Existing users verify with type "email"; a first-ever sign-in (new user)
  // needs type "signup". Try both so a valid code is never wrongly rejected.
  let { data, error } = await supabase.auth.verifyOtp({ email: e, token: t, type: "email" });
  if (error) {
    const retry = await supabase.auth.verifyOtp({ email: e, token: t, type: "signup" });
    if (!retry.error) { data = retry.data; error = null; }
  }
  if (error) throw error;
  user = data.user;
  accessToken = data.session?.access_token || accessToken;
  const r = await mergeCloud();
  return { email: user.email, merged: r.pulled ? "merged" : "pushed" };
}

// Fallback sign-in: the user pastes the magic-link URL (or its hash) from the
// email. We extract the token and set the session directly — this works even
// when the link redirected to the wrong site, because we never rely on the
// redirect destination, only on the token it carries.
export async function signInWithLink(urlOrHash) {
  if (!supabase) initCloud();
  const raw = String(urlOrHash || "").trim();
  if (!raw) throw new Error("Встав посилання з листа.");
  let frag = raw;
  const h = raw.indexOf("#"), q = raw.indexOf("?");
  if (h >= 0) frag = raw.slice(h + 1);
  else if (q >= 0) frag = raw.slice(q + 1);
  const params = new URLSearchParams(frag);
  const access_token = params.get("access_token");
  const refresh_token = params.get("refresh_token");
  const code = params.get("code");
  let res;
  if (access_token && refresh_token) {
    res = await supabase.auth.setSession({ access_token, refresh_token });
  } else if (code) {
    res = await supabase.auth.exchangeCodeForSession(code);
  } else {
    throw new Error("У посиланні немає токена. Скопіюй увесь URL зі стрічки адреси браузера (після кліку на посилання) — там має бути «access_token=…».");
  }
  if (res.error) throw res.error;
  user = res.data.user;
  accessToken = res.data.session?.access_token || access_token || accessToken;
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
