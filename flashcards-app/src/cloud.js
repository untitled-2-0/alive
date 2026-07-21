// Cloud sync via Supabase — optional, opt-in, key-value mirror of window.storage.
// The anon key is public by design (protected by Row-Level-Security on the kv table).
import { createClient } from "@supabase/supabase-js";

const SUPABASE_URL = "https://nczwmbezbxmvwrabmrju.supabase.co";
const SUPABASE_ANON = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Im5jendtYmV6YnhtdndyYWJtcmp1Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODQ2MzcxNzIsImV4cCI6MjEwMDIxMzE3Mn0.NkjwuQVL4rbgjcI3KZ6p8AIs9Np-dRK7HyEcqqVp5ww";
const PREFIX = "recall:"; // window.storage shim prefixes local keys with this

let supabase = null;
let user = null;

export function initCloud() {
  if (!supabase) {
    supabase = createClient(SUPABASE_URL, SUPABASE_ANON, {
      auth: { persistSession: true, autoRefreshToken: true, detectSessionInUrl: false },
    });
  }
  return supabase;
}

export async function refreshSession() {
  if (!supabase) initCloud();
  try {
    const { data } = await supabase.auth.getSession();
    user = data?.session?.user || null;
  } catch { user = null; }
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
  const rows = [];
  for (let i = 0; i < localStorage.length; i++) {
    const k = localStorage.key(i);
    if (!k || !k.startsWith(PREFIX)) continue;
    const key = k.slice(PREFIX.length);
    let value;
    try { value = JSON.parse(localStorage.getItem(k)); } catch { value = localStorage.getItem(k); }
    rows.push({ user_id: user.id, key, value });
  }
  // upsert in chunks to stay well within limits
  for (let i = 0; i < rows.length; i += 200) {
    const { error } = await supabase.from("kv").upsert(rows.slice(i, i + 200));
    if (error) throw error;
  }
  return rows.length;
}

/* ---- debounced dual-write ---- */
const pending = new Map();
let flushTimer = null;
function schedule() { if (!flushTimer) flushTimer = setTimeout(flush, 1200); }
async function flush() {
  flushTimer = null;
  if (!user || !supabase || !pending.size) return;
  const upserts = [], deletes = [];
  for (const [k, v] of pending) { if (v && v.__deleted) deletes.push(k); else upserts.push({ user_id: user.id, key: k, value: v }); }
  pending.clear();
  try {
    if (upserts.length) await supabase.from("kv").upsert(upserts);
    for (const k of deletes) await supabase.from("kv").delete().eq("key", k);
  } catch (e) { console.warn("[cloud flush]", e?.message || e); }
}
export function cloudPush(key, value) { if (!user) return; pending.set(key, value); schedule(); }
export function cloudRemove(key) { if (!user) return; pending.set(key, { __deleted: true }); schedule(); }
if (typeof window !== "undefined") window.addEventListener("beforeunload", () => { if (pending.size) flush(); });

/* ---- auth (email OTP code) ---- */
export async function sendCode(email) {
  if (!supabase) initCloud();
  const { error } = await supabase.auth.signInWithOtp({ email: email.trim(), options: { shouldCreateUser: true } });
  if (error) throw error;
}
// After verifying, merge: if the cloud is empty, upload local; otherwise download cloud. Then caller reloads.
export async function verifyCode(email, token) {
  const { data, error } = await supabase.auth.verifyOtp({ email: email.trim(), token: token.trim(), type: "email" });
  if (error) throw error;
  user = data.user;
  const { count } = await supabase.from("kv").select("*", { count: "exact", head: true });
  if (!count) await pushAllLocal(); else await pullToLocal();
  return { email: user.email, merged: count ? "pulled" : "pushed" };
}
export async function signOutCloud() {
  if (!supabase) return;
  try { await supabase.auth.signOut(); } catch { /* ignore */ }
  user = null;
}
