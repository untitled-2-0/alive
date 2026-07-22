import React, { useState, useEffect, useMemo, useRef, useCallback } from "react";
import * as XLSX from "xlsx";
import Papa from "papaparse";
import {
  BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer, CartesianGrid, Cell,
  LineChart, Line,
} from "recharts";
import {
  Brain, Upload, BarChart3, Layers, Plus, Trash2, Download, RotateCcw,
  ChevronRight, Flame, Clock, FileSpreadsheet, ClipboardPaste, X, Sparkles,
  GraduationCap, Play, Keyboard, Check, ArrowLeft, Target, Inbox,
  Pencil, Filter, Zap, Shuffle, Tag, FolderPlus, BookOpen, ChevronDown, Layers3,
  Image as ImageIcon, ImagePlus, Folder, FolderOpen, Volume2, CalendarClock,
  Timer, Maximize2, Settings, ChevronsUpDown, CalendarDays,
  PanelLeftClose, PanelLeft, CheckCircle2, Circle, Sun, Repeat, ListChecks,
  Trophy, Smile, Menu, GripVertical, ArrowRight, Sunrise,
  Wind, Waves, Anchor, HeartPulse, TrendingUp, NotebookPen, Hourglass,
  Leaf, Pause, SkipForward, ListTree, Heart, Sparkle,
  Coffee, Droplet, Scale, ShieldAlert, Info, Square, TrendingDown,
  Utensils, GlassWater, LineChart as LineChartIcon, Cloud, CloudOff, LogOut, Mail,
  Briefcase, Lightbulb, Compass, BookMarked, ChevronLeft, RefreshCw,
  Wrench, Star, Users, Sparkles as SparklesIcon, Scale as ScaleIcon, ArrowLeftRight, Home,
  HandHeart, ShoppingCart, Wallet, ShoppingBasket, Search,
  Package, Lock, HelpCircle,
} from "lucide-react";
import { cloudPush, cloudRemove, isSignedIn as cloudSignedIn, currentEmail, sendCode, verifyCode, signInWithLink, signOutCloud, refreshSession, syncNow } from "./cloud.js";

/* ------------------------------------------------------------------ */
/* Constants + tiny utils                                              */
/* ------------------------------------------------------------------ */
const DAY = 86_400_000;
const MIN = 60_000;
const LEARN_STEPS = [1, 10]; // minutes
const GRADUATE_GOOD = 1; // days
const GRADUATE_EASY = 4; // days
const DEFAULT_NEW_PER_DAY = 20;
const SESSION_REQUEUE_WINDOW = 20 * MIN; // "Again" re-appears within the session

const uid = (p = "id") =>
  `${p}_${Date.now().toString(36)}_${Math.random().toString(36).slice(2, 7)}`;

const dateKey = (ms) => {
  const d = new Date(ms);
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(
    d.getDate()
  ).padStart(2, "0")}`;
};

const clampEF = (ef) => Math.max(1.3, ef);

function formatDelta(ms) {
  if (ms <= 0) return "now";
  if (ms < DAY) {
    const m = Math.round(ms / MIN);
    if (m < 60) return `${m}m`;
    return `${Math.round(m / 60)}h`;
  }
  const days = Math.round(ms / DAY);
  if (days < 30) return `${days}d`;
  if (days < 365) return `${Math.round(days / 30)}mo`;
  return `${(days / 365).toFixed(1)}y`;
}

/* ------------------------------------------------------------------ */
/* SM-2 (Anki-flavoured) scheduler                                     */
/* ------------------------------------------------------------------ */
// deckOpts (optional): { goal: "longterm" | "deadline", deadline: ms }
function schedule(card, grade, now, deckOpts) {
  let ef = clampEF(card.ef || 2.5);
  let interval = card.interval || 0;
  let reps = card.reps || 0;
  let lapses = card.lapses || 0;
  let stepIdx = card.stepIdx || 0;
  let state = card.state || "new";
  let due;

  if (state === "new" || state === "learning") {
    if (grade === "again") {
      state = "learning";
      stepIdx = 0;
      interval = 0;
      due = now + LEARN_STEPS[0] * MIN;
    } else if (grade === "hard") {
      state = "learning";
      due = now + LEARN_STEPS[Math.min(stepIdx, LEARN_STEPS.length - 1)] * MIN;
    } else if (grade === "good") {
      if (stepIdx >= LEARN_STEPS.length - 1) {
        state = "review";
        interval = GRADUATE_GOOD;
        reps = 1;
        stepIdx = 0;
        due = now + interval * DAY;
      } else {
        stepIdx += 1;
        state = "learning";
        due = now + LEARN_STEPS[stepIdx] * MIN;
      }
    } else {
      // easy — graduate immediately
      state = "review";
      interval = GRADUATE_EASY;
      reps = 1;
      stepIdx = 0;
      due = now + interval * DAY;
    }
  } else {
    // review state
    if (grade === "again") {
      lapses += 1;
      ef = clampEF(ef - 0.2);
      reps = 0;
      state = "learning";
      stepIdx = 0;
      interval = 0;
      due = now + LEARN_STEPS[0] * MIN;
    } else if (grade === "hard") {
      ef = clampEF(ef - 0.15);
      interval = Math.max(1, Math.round(interval * 1.2));
      due = now + interval * DAY;
    } else if (grade === "good") {
      interval = Math.max(1, Math.round(interval * ef));
      reps += 1;
      due = now + interval * DAY;
    } else {
      ef = ef + 0.15;
      interval = Math.max(1, Math.round(interval * ef * 1.3));
      reps += 1;
      due = now + interval * DAY;
    }
  }

  // Deadline goal: compress review intervals so every card cycles several more
  // times before the target date instead of being pushed months out.
  if (state === "review" && deckOpts && deckOpts.goal === "deadline" && deckOpts.deadline) {
    const daysLeft = (deckOpts.deadline - now) / DAY;
    const cap = daysLeft <= 1 ? 1 : Math.max(1, Math.floor(daysLeft / 3));
    interval = Math.min(interval, cap);
    due = Math.min(now + interval * DAY, deckOpts.deadline);
    if (due <= now) due = now + Math.max(1, interval) * DAY;
  }

  // record the interval progression (in days) each time the card lands in review,
  // so the card can show its schedule in plain words ("1 → 3 → 8 d")
  let ivls = card.ivls || [];
  if (state === "review") ivls = [...ivls, interval];

  return { ...card, ef, interval, reps, lapses, stepIdx, due, state, lastReviewed: now, ivls };
}

// Plain-words spaced-repetition schedule for one card.
function cardScheduleText(card, now = Date.now()) {
  if (card.state === "new") return { line: "Нова картка — ще не в графіку", prog: "", next: "" };
  const ivls = card.ivls || [];
  const prog = ivls.length ? ivls.map((d) => `${d}`).join(" → ") + " дн" : "";
  const cur = card.state === "learning" ? "вивчення" : `${card.interval || 1} дн`;
  const diff = card.due - now;
  let next;
  if (diff <= 0) next = "готова зараз";
  else if (diff < DAY) next = `за ${Math.max(1, Math.round(diff / (60 * 60 * 1000)))} год`;
  else { const days = Math.round(diff / DAY); next = `за ${days} ${days === 1 ? "день" : days < 5 ? "дні" : "днів"}`; }
  return { prog, cur, next };
}

/* ------------------------------------------------------------------ */
/* Interval / scheduling visualisation                                 */
/* ------------------------------------------------------------------ */
// Ordered stages from short (hot) to long (cool). `max` is in days.
const INTERVAL_STAGES = [
  { id: "learning", label: "Learning", max: 0, dot: "#e11d48", bg: "bg-red-100", text: "text-red-700" },
  { id: "1d", label: "1 day", max: 1, dot: "#f97316", bg: "bg-orange-100", text: "text-orange-700" },
  { id: "3d", label: "3 days", max: 3, dot: "#f59e0b", bg: "bg-amber-100", text: "text-amber-700" },
  { id: "1w", label: "1 week", max: 7, dot: "#eab308", bg: "bg-yellow-100", text: "text-yellow-700" },
  { id: "2w", label: "2 weeks", max: 14, dot: "#84cc16", bg: "bg-lime-100", text: "text-lime-700" },
  { id: "1mo", label: "1 month", max: 30, dot: "#22c55e", bg: "bg-green-100", text: "text-green-700" },
  { id: "3mo", label: "3 months", max: 90, dot: "#14b8a6", bg: "bg-teal-100", text: "text-teal-700" },
  { id: "long", label: "Long-term", max: Infinity, dot: "#2563eb", bg: "bg-blue-100", text: "text-blue-700" },
];

function stageForCard(card) {
  if (card.state === "new") return { id: "new", label: "New", max: 0, dot: "#94a3b8", bg: "bg-slate-100", text: "text-slate-600" };
  if (card.state === "learning") return INTERVAL_STAGES[0];
  const d = card.interval || 0;
  for (const s of INTERVAL_STAGES) if (d <= s.max) return s;
  return INTERVAL_STAGES[INTERVAL_STAGES.length - 1];
}

// human label for a card's interval, e.g. "3d", "2w", "Learning"
function intervalLabel(card) {
  if (card.state === "new") return "New";
  if (card.state === "learning") return "Learning";
  return formatDelta((card.interval || 1) * DAY);
}

// "in 3d", "tomorrow", "today", "2h" — relative next-due
function dueLabel(card, now = Date.now()) {
  const diff = card.due - now;
  if (diff <= 0) return "due now";
  if (diff < DAY) {
    const h = Math.round(diff / (60 * MIN));
    return h <= 1 ? "< 1h" : `in ${h}h`;
  }
  const days = Math.round(diff / DAY);
  if (days === 1) return "tomorrow";
  if (days < 30) return `in ${days}d`;
  if (days < 365) return `in ${Math.round(days / 30)}mo`;
  return `in ${(days / 365).toFixed(1)}y`;
}

const SCHED_GOALS = {
  longterm: { id: "longterm", label: "Long-term retention", short: "Long-term", icon: Layers3, desc: "Standard SM-2 — intervals grow into months and years." },
  deadline: { id: "deadline", label: "Short-term / deadline", short: "Deadline", icon: CalendarClock, desc: "Compress intervals to fit before a target date, cycling each card several times." },
};

function daysUntil(ms, now = Date.now()) {
  return Math.ceil((ms - now) / DAY);
}

const GRADES = [
  { key: "again", label: "Again", hint: "1", cls: "bg-red-600 hover:bg-red-700", ring: "ring-red-300" },
  { key: "hard", label: "Hard", hint: "2", cls: "bg-amber-500 hover:bg-amber-600", ring: "ring-amber-300" },
  { key: "good", label: "Good", hint: "3", cls: "bg-green-600 hover:bg-green-700", ring: "ring-green-300" },
  { key: "easy", label: "Easy", hint: "4", cls: "bg-blue-600 hover:bg-blue-700", ring: "ring-blue-300" },
];

/* ------------------------------------------------------------------ */
/* Deck cosmetics + taxonomy                                           */
/* ------------------------------------------------------------------ */
const DECK_COLORS = [
  { id: "indigo", dot: "#4f46e5", bg: "bg-indigo-50", text: "text-indigo-600" },
  { id: "violet", dot: "#7c3aed", bg: "bg-violet-50", text: "text-violet-600" },
  { id: "blue", dot: "#2563eb", bg: "bg-blue-50", text: "text-blue-600" },
  { id: "teal", dot: "#0d9488", bg: "bg-teal-50", text: "text-teal-600" },
  { id: "green", dot: "#16a34a", bg: "bg-green-50", text: "text-green-600" },
  { id: "amber", dot: "#d97706", bg: "bg-amber-50", text: "text-amber-600" },
  { id: "orange", dot: "#ea580c", bg: "bg-orange-50", text: "text-orange-600" },
  { id: "rose", dot: "#e11d48", bg: "bg-rose-50", text: "text-rose-600" },
  { id: "pink", dot: "#db2777", bg: "bg-pink-50", text: "text-pink-600" },
  { id: "slate", dot: "#475569", bg: "bg-slate-100", text: "text-slate-600" },
];
const getColor = (id) => DECK_COLORS.find((c) => c.id === id) || DECK_COLORS[0];

const TOPIC_PRESETS = [
  "Languages", "Biology", "History", "Geography", "Science", "Medicine",
  "Law", "Business", "Technology", "Math", "Art", "Music", "Exam prep", "Other",
];

const DECK_EMOJIS = [
  "📚", "🗣️", "🧬", "🏛️", "🌍", "⚗️", "💊", "⚖️", "💼", "💻",
  "➗", "🎨", "🎵", "🍳", "⚽", "✈️", "🧠", "📝", "🔬", "🌱", "⭐", "🔥",
];

const STUDY_MODES = [
  { id: "due", label: "Due today", icon: Target, desc: "Only cards the schedule says are ready now." },
  { id: "custom", label: "A set number", icon: Layers3, desc: "Study a fixed number of cards, due ones first." },
  { id: "all", label: "All cards", icon: BookOpen, desc: "Every card in scope, scheduling as normal." },
  { id: "new", label: "Only new", icon: Sparkles, desc: "Cards you haven't started learning yet." },
  { id: "review", label: "Only review", icon: RotateCcw, desc: "Cards you've already started — drill them ahead of time." },
  { id: "cram", label: "Cram", icon: Zap, desc: "Ignore the schedule and drill. Won't change your due dates." },
];

/* ------------------------------------------------------------------ */
/* Persistence layer (window.storage) with in-memory fallback          */
/* ------------------------------------------------------------------ */
const memFallback = new Map();
// Resolve the backend lazily on every call: in a claude.ai artifact window.storage
// is present up front, but a local shim may install it after this module loads, so
// caching the check once would silently strand every write in memory.
const backend = () =>
  typeof window !== "undefined" && window.storage && typeof window.storage.getItem === "function"
    ? window.storage
    : null;

const store = {
  async get(key, fallback) {
    try {
      const be = backend();
      const raw = be ? await be.getItem(key) : memFallback.has(key) ? memFallback.get(key) : null;
      if (raw == null) return fallback;
      return typeof raw === "string" ? JSON.parse(raw) : raw;
    } catch (e) {
      console.error("[storage.get]", key, e);
      return fallback;
    }
  },
  async set(key, value) {
    try {
      const raw = JSON.stringify(value);
      const be = backend();
      if (be) await be.setItem(key, raw);
      else memFallback.set(key, raw);
      cloudPush(key, value); // mirror to cloud when signed in (no-op otherwise)
      return true;
    } catch (e) {
      console.error("[storage.set]", key, e);
      return false;
    }
  },
  async remove(key) {
    try {
      const be = backend();
      if (be) await be.removeItem(key);
      else memFallback.delete(key);
      cloudRemove(key);
    } catch (e) {
      console.error("[storage.remove]", key, e);
    }
  },
};

/* ------------------------------------------------------------------ */
/* Card factory + column auto-detection                                */
/* ------------------------------------------------------------------ */
function makeCard(front, back, tags = "", notes = "") {
  const now = Date.now();
  return {
    id: uid("c"),
    front: String(front ?? "").trim(),
    back: String(back ?? "").trim(),
    tags: String(tags ?? "").trim(),
    notes: String(notes ?? "").trim(),
    ef: 2.5,
    interval: 0,
    reps: 0,
    lapses: 0,
    stepIdx: 0,
    state: "new",
    due: now,
    created: now,
    lastReviewed: null,
  };
}

const SYNONYMS = {
  front: ["front", "question", "q", "term", "word", "prompt", "en", "english", "kanji", "spanish"],
  back: ["back", "answer", "a", "definition", "meaning", "translation", "uk", "ukrainian", "es", "reverse"],
  deck: ["deck", "category", "topic", "group", "set", "subject", "chapter"],
  tags: ["tags", "tag", "labels", "label"],
  notes: ["notes", "note", "extra", "hint", "example", "context"],
};

function autoMapColumns(headers) {
  const map = { front: "", back: "", deck: "", tags: "", notes: "" };
  const lower = headers.map((h) => ({ h, l: String(h).toLowerCase().trim() }));
  for (const field of Object.keys(SYNONYMS)) {
    const found = lower.find(({ l }) => SYNONYMS[field].includes(l));
    if (found) map[field] = found.h;
  }
  // fallback: first two unmapped columns become front/back
  if (!map.front && headers[0]) map.front = headers[0];
  if (!map.back) {
    const back = headers.find((h) => h !== map.front);
    if (back) map.back = back;
  }
  return map;
}

/* ------------------------------------------------------------------ */
/* Multi-sheet extraction — for messy real-world workbooks             */
/* Finds the Front/Back columns per sheet even when there are title    */
/* rows, __EMPTY headers, or extra columns, and works left-to-right    */
/* by fullness when no header row is present.                          */
/* ------------------------------------------------------------------ */
const FRONT_HEADER = /^(front|question|q|term|word|phrase|english|prompt|фраза|англ)/i;
const BACK_HEADER = /^(back|answer|a|definition|meaning|translation|reverse|переклад|українськ|ukrainian|значенн)/i;

function cleanDeckName(name) {
  const n = String(name || "").trim();
  return n.replace(/\s+v?\d+(\.\d+)?\s*$/i, "").trim() || n;
}

function headersLookMessy(headers) {
  return headers.some((h) => /^__EMPTY/.test(String(h))) || !autoMapColumns(headers).front;
}

// rows = array-of-arrays (XLSX sheet_to_json with header:1)
function extractSheetCards(rows) {
  if (!rows || !rows.length) return { cards: [] };
  const width = rows.reduce((w, r) => Math.max(w, r.length), 0);
  const cell = (r, c) => String((r && r[c]) ?? "").trim();

  // 1) look for a header row in the first few rows
  let hRow = -1, fCol = -1, bCol = -1;
  for (let i = 0; i < Math.min(8, rows.length); i++) {
    let f = -1, b = -1;
    for (let c = 0; c < width; c++) {
      const v = cell(rows[i], c);
      if (!v || v.length > 40) continue;
      if (f < 0 && FRONT_HEADER.test(v)) f = c;
      else if (b < 0 && BACK_HEADER.test(v)) b = c;
    }
    if (f >= 0 && b >= 0) { hRow = i; fCol = f; bCol = b; break; }
  }

  // 2) fallback: the two most-filled columns, left-to-right
  if (fCol < 0) {
    const fill = Array(width).fill(0);
    for (const r of rows) for (let c = 0; c < width; c++) if (cell(r, c)) fill[c] += 1;
    const order = [...fill.keys()].filter((c) => fill[c] > 0).sort((a, b) => fill[b] - fill[a]);
    const two = order.slice(0, 2).sort((a, b) => a - b);
    fCol = two[0] ?? 0;
    bCol = two[1] ?? 1;
  }

  const cards = [];
  for (let i = hRow >= 0 ? hRow + 1 : 0; i < rows.length; i++) {
    const f = cell(rows[i], fCol);
    const b = cell(rows[i], bCol);
    if (!f || !b) continue; // skip blanks, section titles, date-only rows
    if (FRONT_HEADER.test(f) && BACK_HEADER.test(b)) continue; // stray header row
    cards.push([f, b]);
  }
  return { cards, frontCol: fCol, backCol: bCol };
}

/* ------------------------------------------------------------------ */
/* Stats helpers                                                       */
/* ------------------------------------------------------------------ */
const emptyDay = () => ({
  studied: 0, again: 0, hard: 0, good: 0, easy: 0,
  newIntroduced: 0, matureAns: 0, maturePass: 0,
});

function bumpStats(stats, { grade, wasNew, mature }) {
  const key = dateKey(Date.now());
  const history = { ...(stats.history || {}) };
  const d = { ...emptyDay(), ...(history[key] || {}) };
  d.studied += 1;
  d[grade] += 1;
  if (wasNew) d.newIntroduced += 1;
  if (mature) {
    d.matureAns += 1;
    if (grade !== "again") d.maturePass += 1;
  }
  history[key] = d;
  return { ...stats, history };
}

function computeStreak(history) {
  if (!history) return 0;
  let streak = 0;
  const day = new Date();
  if (!(history[dateKey(day.getTime())]?.studied > 0)) {
    day.setDate(day.getDate() - 1); // today not done yet — count from yesterday
  }
  while (history[dateKey(day.getTime())]?.studied > 0) {
    streak += 1;
    day.setDate(day.getDate() - 1);
  }
  return streak;
}

function retention30(history) {
  if (!history) return null;
  const now = Date.now();
  let ans = 0, pass = 0;
  for (let i = 0; i < 30; i++) {
    const k = dateKey(now - i * DAY);
    const d = history[k];
    if (d) { ans += d.matureAns || 0; pass += d.maturePass || 0; }
  }
  if (!ans) return null;
  return Math.round((pass / ans) * 100);
}

/* ------------------------------------------------------------------ */
/* Images — stored one-per-key so the batched cards blob stays small   */
/* ------------------------------------------------------------------ */
const imgKey = (cardId, side) => `image:${cardId}:${side}`;
const loadImage = (cardId, side) => store.get(imgKey(cardId, side), null);
const saveImage = (cardId, side, dataUrl) => store.set(imgKey(cardId, side), dataUrl);
const removeImage = (cardId, side) => store.remove(imgKey(cardId, side));

// Downscale + compress to a small JPEG data URL. Accepts a File or a data URL.
function compressImage(source, maxW = 1000, quality = 0.8) {
  return new Promise((resolve, reject) => {
    const img = new window.Image();
    img.onload = () => {
      try {
        const scale = Math.min(1, maxW / img.width);
        const w = Math.max(1, Math.round(img.width * scale));
        const h = Math.max(1, Math.round(img.height * scale));
        const canvas = document.createElement("canvas");
        canvas.width = w;
        canvas.height = h;
        const ctx = canvas.getContext("2d");
        ctx.fillStyle = "#ffffff"; // flatten transparency for JPEG
        ctx.fillRect(0, 0, w, h);
        ctx.drawImage(img, 0, 0, w, h);
        resolve(canvas.toDataURL("image/jpeg", quality));
      } catch (e) {
        reject(e);
      }
    };
    img.onerror = () => reject(new Error("Could not load that image."));
    if (typeof source === "string") {
      img.src = source;
    } else {
      const reader = new FileReader();
      reader.onload = (e) => { img.src = e.target.result; };
      reader.onerror = () => reject(new Error("Could not read that file."));
      reader.readAsDataURL(source);
    }
  });
}

// Pull the first image out of a clipboard paste event, if any.
function imageFromClipboard(e) {
  const items = e.clipboardData?.items || [];
  for (const it of items) {
    if (it.type && it.type.startsWith("image/")) return it.getAsFile();
  }
  return null;
}

/* ------------------------------------------------------------------ */
/* Text-to-speech (Web Speech API) — generated on the fly, never stored */
/* ------------------------------------------------------------------ */
const LANGUAGES = [
  { code: "", label: "Deck default" },
  { code: "en-US", label: "English (US)" },
  { code: "en-GB", label: "English (UK)" },
  { code: "uk-UA", label: "Українська" },
  { code: "de-DE", label: "Deutsch" },
  { code: "fr-FR", label: "Français" },
  { code: "es-ES", label: "Español" },
  { code: "it-IT", label: "Italiano" },
  { code: "pt-BR", label: "Português (BR)" },
  { code: "pl-PL", label: "Polski" },
  { code: "ru-RU", label: "Русский" },
  { code: "ja-JP", label: "日本語" },
  { code: "ko-KR", label: "한국어" },
  { code: "zh-CN", label: "中文" },
];
const DECK_LANGUAGES = LANGUAGES.filter((l) => l.code); // deck picker excludes "default"

const ttsSupported = () => typeof window !== "undefined" && "speechSynthesis" in window;

function getVoices() {
  if (!ttsSupported()) return [];
  return window.speechSynthesis.getVoices() || [];
}

// Best voice for a language tag: exact match → same base language → null.
function pickVoice(lang, voices = getVoices()) {
  if (!lang || !voices.length) return null;
  const lower = lang.toLowerCase();
  const base = lower.split("-")[0];
  return (
    voices.find((v) => v.lang && v.lang.toLowerCase() === lower) ||
    voices.find((v) => v.lang && v.lang.toLowerCase().startsWith(base)) ||
    null
  );
}

// Returns { ok } — ok=false means we fell back to the default voice (no match).
function speak(text, lang) {
  if (!ttsSupported() || !text) return { ok: true };
  try {
    window.speechSynthesis.cancel();
    const u = new SpeechSynthesisUtterance(String(text));
    const voice = pickVoice(lang);
    let ok = true;
    if (voice) {
      u.voice = voice;
      u.lang = voice.lang;
    } else if (lang) {
      u.lang = lang;
      ok = false; // no exact voice — browser default will approximate
    }
    window.speechSynthesis.speak(u);
    return { ok };
  } catch (e) {
    console.error("[tts]", e);
    return { ok: false };
  }
}

/* ================================================================== */
/* MY ROUTINE — Me+ style planner                                     */
/* Tasks (with subtasks = routine steps), categories, timed goals,    */
/* daily streak, mood. Stored under routine:* keys, separate from     */
/* every flashcard key.                                               */
/* ================================================================== */
const RKEYS = {
  tasks: "routine:tasks",
  categories: "routine:categories",
  cindex: "routine:completions:index",
  streak: "routine:streak",
  mood: "routine:mood",
  seeded: "routine:seeded",
  xp: "routine:xp",         // { xp: number }
  rewards: "routine:rewards", // [{ id, label, cost, unlockedAt }]
  pomo: "routine:pomo",      // { work, break } Pomodoro settings
  movement: "routine:movement",      // { [date]: count }
  movementCfg: "routine:movement:cfg", // { remindersOn, snoozeUntil }
  meds: "routine:meds",              // [{ id, name, dose, perDay, supply, refillAt, taper:[{date,dose,note}] }]
  medsLog: "routine:meds:wellbeing", // { [date]: { taken:{medId:bool}, wellbeing, sideEffects, note } }
  gratitude: "mood:gratitude",       // { [date]: [items] }
  activation: "mood:activation",     // [{ id, text }]
  moodNotes: "mood:notes",           // { [date]: { note, factors:[] } }
};
const cKey = (date) => `routine:completions:${date}`;

/* ---- ADHD gamification helpers ---- */
const ENERGY = {
  low: { label: "Легко", emoji: "🟢", xp: 0 },
  med: { label: "Середнє", emoji: "🟡", xp: 5 },
  high: { label: "Складне", emoji: "🔴", xp: 15 },
};
const EST_CHIPS = [2, 5, 15, 30, 60];
// XP earned for finishing a task: base + effort + a little for longer estimates
function xpForTask(task) {
  let xp = 10;
  if (task?.energy && ENERGY[task.energy]) xp += ENERGY[task.energy].xp;
  if (task?.estMin) xp += Math.min(20, Math.floor(task.estMin / 5) * 2);
  return xp;
}
// Level curve: level N starts at 50*N*(N-1) XP (0,100,300,600,1000,…)
function levelFromXp(xp) {
  let lvl = 1;
  while (50 * (lvl + 1) * lvl <= xp) lvl += 1;
  return lvl;
}
const xpForLevel = (lvl) => 50 * lvl * (lvl - 1);
function levelProgress(xp) {
  const lvl = levelFromXp(xp);
  const cur = xpForLevel(lvl), next = xpForLevel(lvl + 1);
  return { lvl, cur, next, into: xp - cur, span: next - cur, pct: Math.min(1, (xp - cur) / (next - cur)) };
}
const fmtEst = (min) => { min = Math.round(min || 0); if (min < 60) return `${min} хв`; const h = Math.floor(min / 60), m = min % 60; return m ? `${h} год ${m} хв` : `${h} год`; };

// Daily challenges — rotating pool. Each has a check(ctx) -> bool.
const CHALLENGE_POOL = [
  { id: "close3", emoji: "✅", label: "Закрий 3 справи сьогодні", xp: 20, check: (c) => c.doneCount >= 3 },
  { id: "close5", emoji: "🏆", label: "Закрий 5 справ сьогодні", xp: 35, check: (c) => c.doneCount >= 5 },
  { id: "frog", emoji: "🐸", label: "З'їж жабу: закрий одну «складну» справу", xp: 25, check: (c) => c.tasks.some((t) => t.energy === "high" && c.isDone(t)) },
  { id: "anytime", emoji: "🎈", label: "Закрий одну справу «будь-коли»", xp: 15, check: (c) => c.tasks.some((t) => !t.time && c.isDone(t)) },
  { id: "beatEst", emoji: "⏱️", label: "Вклади в свою оцінку часу на одній справі", xp: 25, check: (c) => c.tasks.some((t) => t.estMin && c.actualMin(t) != null && c.actualMin(t) <= t.estMin && c.isDone(t)) },
  { id: "twoQuick", emoji: "⚡", label: "Закрий 2 швидкі перемоги (≤5 хв)", xp: 20, check: (c) => c.tasks.filter((t) => (t.estMin && t.estMin <= 5) && c.isDone(t)).length >= 2 },
  { id: "firstThing", emoji: "🌅", label: "Закрий першу справу дня зранку", xp: 15, check: (c) => c.doneCount >= 1 },
  { id: "half", emoji: "🌤️", label: "Закрий половину сьогоднішніх справ", xp: 30, check: (c) => c.tasks.length > 0 && c.doneCount >= Math.ceil(c.tasks.length / 2) },
];
// Deterministic 3 challenges per day
function pickChallenges(dateStr) {
  let h = 0; for (let i = 0; i < dateStr.length; i++) h = (h * 31 + dateStr.charCodeAt(i)) >>> 0;
  const pool = [...CHALLENGE_POOL];
  const out = [];
  for (let i = 0; i < 3 && pool.length; i++) { const idx = h % pool.length; out.push(pool.splice(idx, 1)[0]); h = (h * 1103515245 + 12345) >>> 0; }
  return out;
}
const ruid = (p) => uid(p);

// Me+ pastel palette
const PASTELS = [
  { id: "pink", card: "#fde7ef", chip: "#fbcfe0", ink: "#be185d", dot: "#ec4899" },
  { id: "orange", card: "#ffe9d6", chip: "#fed7aa", ink: "#c2410c", dot: "#fb923c" },
  { id: "yellow", card: "#fef7c8", chip: "#fde68a", ink: "#a16207", dot: "#f5b800" },
  { id: "green", card: "#d8f6e3", chip: "#bbf7d0", ink: "#15803d", dot: "#34d399" },
  { id: "teal", card: "#cdf5f6", chip: "#a5f3fc", ink: "#0e7490", dot: "#22d3ee" },
  { id: "purple", card: "#ece9fe", chip: "#ddd6fe", ink: "#6d28d9", dot: "#a78bfa" },
];
const getPastel = (id) => PASTELS.find((p) => p.id === id) || PASTELS[0];

const TASK_EMOJIS = [
  "☀️", "🌅", "🌙", "😴", "💧", "💊", "🧘", "🏃", "🚿", "🍳", "🍲", "🍽️",
  "🧹", "🧺", "💻", "📚", "✍️", "🎧", "🎉", "🛒", "🐾", "🌱", "🙏", "🧠",
  "☕", "🦷", "🚶", "📵", "✋", "❤️", "⭐", "🔥",
];

const MOODS = [
  { score: 1, emoji: "😞", label: "Rough", color: "#f472b6" },
  { score: 2, emoji: "😕", label: "Meh", color: "#fb923c" },
  { score: 3, emoji: "😐", label: "Okay", color: "#facc15" },
  { score: 4, emoji: "🙂", label: "Good", color: "#4ade80" },
  { score: 5, emoji: "😄", label: "Great", color: "#22d3ee" },
];

const WEEKDAYS = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];
const WEEKDAYS_MON = [1, 2, 3, 4, 5, 6, 0]; // Mon-first order of getDay() values
const WD_LETTER = { 0: "S", 1: "M", 2: "T", 3: "W", 4: "T", 5: "F", 6: "S" };

/* ---- persistence ---- */
async function loadRoutineData() {
  const tasks = await store.get(RKEYS.tasks, null);
  const categories = await store.get(RKEYS.categories, null);
  const cindex = await store.get(RKEYS.cindex, []);
  const streak = await store.get(RKEYS.streak, { best: 0, lastCelebrated: "" });
  const moods = await store.get(RKEYS.mood, {});
  const xp = await store.get(RKEYS.xp, { xp: 0 });
  const rewards = await store.get(RKEYS.rewards, []);
  const pomo = await store.get(RKEYS.pomo, { work: 25, break: 5 });
  const completions = {};
  for (const d of cindex) {
    const doc = await store.get(cKey(d), null);
    if (doc) completions[d] = doc;
  }
  return { tasks, categories, cindex, completions, streak, moods, xp, rewards, pomo };
}

async function collectRoutineExport() {
  const d = await loadRoutineData();
  const extra = {};
  for (const k of ["movement", "movementCfg", "meds", "medsLog", "gratitude", "activation", "moodNotes"]) extra[k] = await store.get(RKEYS[k], null);
  return { tasks: d.tasks || [], categories: d.categories || [], completions: d.completions, streak: d.streak, moods: d.moods, xp: d.xp, rewards: d.rewards, pomo: d.pomo, ...extra };
}

async function clearRoutineData() {
  const cindex = await store.get(RKEYS.cindex, []);
  for (const d of cindex) await store.remove(cKey(d));
  for (const k of Object.values(RKEYS)) await store.remove(k);
  await store.remove("routine:mig:shave"); // one-time migration flag
}

/* ---- scheduling / queries ---- */
function taskOccursOn(task, ds) {
  if (ds < task.date) return false;
  const rep = task.repeat || { type: "off" };
  if (rep.type === "off") return ds === task.date;
  if (rep.type === "daily") return true;
  if (rep.type === "times") return true; // shown daily; weekly target tracked separately
  if (rep.type === "weekdays") {
    const wd = new Date(ds + "T00:00:00").getDay();
    return (rep.days || []).includes(wd);
  }
  return ds === task.date;
}

function repeatWords(task) {
  const rep = task.repeat || { type: "off" };
  let base;
  if (rep.type === "daily") base = "Repeats every day";
  else if (rep.type === "times") base = `Repeats ${rep.times || 3}× per week`;
  else if (rep.type === "weekdays") {
    const ds = (rep.days || []).slice().sort();
    base = ds.length === 7 ? "Repeats every day" : ds.length === 0 ? "No repeat" : "Repeats every " + ds.map((d) => WEEKDAYS[d]).join(", ");
  } else base = "Doesn't repeat";
  const timeStr = task.time ? `. At ${task.time}` : ". Anytime";
  const rem = task.reminder ? `. Remind me at ${task.reminder}` : "";
  return base + timeStr + rem + ".";
}

const dayHasCompletion = (completions, ds) => {
  const doc = completions[ds];
  return !!(doc && doc.tasks && Object.values(doc.tasks).some(Boolean));
};

function computeTaskStreak(completions, today) {
  let current = 0;
  const d = new Date(today + "T00:00:00");
  if (!dayHasCompletion(completions, today)) d.setDate(d.getDate() - 1);
  for (let i = 0; i < 800; i++) {
    if (dayHasCompletion(completions, dateKey(d.getTime()))) { current += 1; d.setDate(d.getDate() - 1); }
    else break;
  }
  let best = 0, run = 0;
  for (let i = 365; i >= 0; i--) {
    if (dayHasCompletion(completions, dateKey(Date.now() - i * DAY))) { run += 1; best = Math.max(best, run); }
    else run = 0;
  }
  return { current, best: Math.max(best, current) };
}

function totalTasksCompleted(completions) {
  let n = 0;
  for (const doc of Object.values(completions)) if (doc?.tasks) n += Object.values(doc.tasks).filter(Boolean).length;
  return n;
}

// how many of a date's occurring tasks are done (for the daily ring / heatmap)
function dayTaskProgress(tasks, completions, ds) {
  const doc = completions[ds] || {};
  const occurring = (tasks || []).filter((t) => taskOccursOn(t, ds));
  const done = occurring.filter((t) => doc.tasks?.[t.id]).length;
  return { done, total: occurring.length, pct: occurring.length ? done / occurring.length : 0 };
}

/* ---- first-run seed: the user's real routine from their plan ---- */
function seededRoutine() {
  const cat = (name, color) => ({ id: ruid("cat"), name, color });
  const morning = cat("Ранок", "pink");
  const work = cat("Робота", "teal");
  const evening = cat("Вечір", "purple");
  const categories = [morning, work, evening];

  const T = (o) => ({
    id: ruid("t"), note: "", time: null, reminder: null, categoryId: "",
    repeat: { type: "daily" }, goal: { type: "off" }, subtasks: [], created: Date.now(),
    date: dateKey(Date.now()), ...o,
  });
  const wd = { type: "weekdays", days: [1, 2, 3, 4, 5] };

  const tasks = [
    T({ emoji: "☀️", image: "/routine/podyom.jpg", title: "Підйом", note: "Тонізуючий напій + вітаміни.", time: "07:00", color: "orange", categoryId: morning.id }),
    T({ emoji: "💊", image: "/routine/lomeksin-face.jpg", title: "Ломексин на обличчя (1-й раз)", note: "Прив'яжи до ранкового ритуалу — одразу після підйому.", time: "07:10", color: "pink", categoryId: morning.id }),
    T({ emoji: "🧘", image: "/routine/kalanetyka.jpg", title: "Каланетика", note: "Поки Міша спить, поки тихо. Таймер — і понеслі.", time: "07:15", color: "purple", categoryId: morning.id, goal: { type: "timed", minutes: 30 } }),
    T({ emoji: "🚿", image: "/routine/dush.jpg", title: "Душ", note: "Кетоконазол шампунь — вт/пт, залишити на 5 хв.", time: "07:45", color: "teal", categoryId: morning.id }),
    T({ emoji: "🍳", image: "/routine/snidanok.jpg", title: "Сніданок + Силібор + ліки від каменю", note: "Їжа вдома — не доставка. Щось просте і швидке.", time: "08:05", color: "yellow", categoryId: morning.id }),
    T({ emoji: "🧹", image: "/routine/prybyrannia.jpg", title: "Прибирання — одна зона", note: "Кухня → ванна → коридор → вітальня → спальня. Таймер і стоп.", time: "08:25", color: "green", categoryId: morning.id, goal: { type: "timed", minutes: 20 } }),

    T({ emoji: "💻", image: "/routine/robota.jpg", title: "Початок роботи", note: "Перші 25 хв — найважливіша задача без відволікань. Телефон вбік.", time: "09:00", color: "teal", categoryId: work.id, repeat: wd }),
    T({ emoji: "🍽️", image: "/routine/lomeksin-lunch.jpg", title: "Обід + Ломексин (2-й раз)", note: "Прив'яжи другий прийом ломексину до обіду — легше пам'ятати.", time: "13:00", color: "orange", categoryId: work.id, repeat: wd }),
    T({ emoji: "✋", image: "/routine/kinets-roboty.jpg", title: "Кінець роботи — закрити ноут", note: "Не тягнути роботу у вечір. Межа важлива.", time: "18:00", color: "pink", categoryId: work.id, repeat: wd }),

    T({ emoji: "📚", title: "Англійська", note: "Одразу після роботи. Duolingo, відео — регулярне.", time: "18:00", color: "purple", categoryId: evening.id, goal: { type: "timed", minutes: 15 } }),
    T({ emoji: "🍲", image: "/routine/vecheria.jpg", title: "Приготування вечері + вечеря", note: "Готуємо вдома. Смачно і без доставки.", time: "18:30", color: "green", categoryId: evening.id }),
    T({ emoji: "🎉", image: "/routine/vilnyi-chas.jpg", title: "Вільний час", note: "Без почуття провини. Ти вже зробила все головне за день.", time: "19:30", color: "pink", categoryId: evening.id }),
    T({ emoji: "💊", title: "Ліки від пролактину", note: "Тільки в понеділок.", time: null, color: "yellow", categoryId: evening.id, repeat: { type: "weekdays", days: [1] } }),
    T({ emoji: "😴", title: "Сон", note: "Телефон геть. Книга або щось легке. До 00:00 вже спати.", time: "23:00", reminder: "23:00", color: "purple", categoryId: evening.id }),

    T({ emoji: "💧", title: "Вода 1.5–2 л", note: "Протягом усього дня, не залпом.", time: null, color: "teal", categoryId: morning.id }),
    T({ emoji: "💊", title: "Золофт", note: "На тумбочці з вечора — щоб не пропустити.", time: "21:00", reminder: "21:00", color: "pink", categoryId: evening.id }),
  ];
  return { categories, tasks };
}

/* ================================================================== */
/* CALM — anti-anxiety practices. Stored under calm:* keys.           */
/* ================================================================== */
const CKEYS = {
  fears: "calm:fears",
  thoughts: "calm:thoughtRecords",
  sessions: "calm:sessions",
  settings: "calm:settings",
  plan: "calm:planDone", // { [date]: { [itemId]: true } } — supportive-plan daily check-off
};

const BREATH_PATTERNS = [
  { id: "box", name: "Квадратне дихання", desc: "Вдих 4 · Затримка 4 · Видих 4 · Затримка 4", phases: [["in", 4], ["hold", 4], ["out", 4], ["hold", 4]] },
  { id: "478", name: "4-7-8", desc: "Вдих 4 · Затримка 7 · Видих 8", phases: [["in", 4], ["hold", 7], ["out", 8]] },
  { id: "relax", name: "Розслаблення", desc: "Вдих 4 · Видих 6", phases: [["in", 4], ["out", 6]] },
];
const PHASE_TEXT = { in: "Вдих", hold: "Затримай", out: "Видих" };

const MUSCLE_GROUPS = [
  "Кисті", "Передпліччя", "Плечі (руки)", "Плечі", "Обличчя", "Шия",
  "Груди й спина", "Живіт", "Сідниці", "Стегна", "Литки", "Стопи",
];
const PMR_TENSE = 10, PMR_RELEASE = 15;

const CALM_TECHNIQUES = {
  breath: { label: "Дихання", icon: Wind },
  pmr: { label: "Розслаблення м'язів", icon: HeartPulse },
  ground: { label: "Заземлення 5-4-3-2-1", icon: Anchor },
  thought: { label: "Журнал думок", icon: NotebookPen },
  fear: { label: "Сходинки страху", icon: TrendingUp },
  focus: { label: "Таймер фокусу", icon: Timer },
  worry: { label: "Час для тривоги", icon: Hourglass },
  beforework: { label: "Перед роботою", icon: Sparkle },
};

async function loadCalmData() {
  const fears = await store.get(CKEYS.fears, []);
  const thoughts = await store.get(CKEYS.thoughts, []);
  const sessions = await store.get(CKEYS.sessions, []);
  const settings = await store.get(CKEYS.settings, { name: "Спокій" });
  return { fears, thoughts, sessions, settings };
}
// Recovery lives inside the Calm tab; its data rides along in Calm's export/reset.
const RECKEYS = { alcohol: "recovery:alcohol", smoke: "recovery:smoke", triggers: "recovery:triggers", reason: "recovery:reason", noteSeen: "recovery:noteSeen" };
async function collectCalmExport() {
  const d = await loadCalmData();
  const recovery = {};
  for (const [k, key] of Object.entries(RECKEYS)) recovery[k] = await store.get(key, null);
  return { fears: d.fears, thoughts: d.thoughts, sessions: d.sessions, settings: d.settings, recovery, plan: await store.get(CKEYS.plan, {}) };
}
async function clearCalmData() {
  for (const k of Object.values(CKEYS)) await store.remove(k);
  for (const k of Object.values(RECKEYS)) await store.remove(k);
  await store.remove("calm:trGuideOpen"); // "how it works" toggle
}

// soft optional tick using Web Audio (no asset, no storage)
let _actx = null;
function calmTick(freq = 440, dur = 0.09) {
  try {
    _actx = _actx || new (window.AudioContext || window.webkitAudioContext)();
    const o = _actx.createOscillator(), g = _actx.createGain();
    o.type = "sine"; o.frequency.value = freq;
    g.gain.setValueAtTime(0.0001, _actx.currentTime);
    g.gain.exponentialRampToValueAtTime(0.06, _actx.currentTime + 0.01);
    g.gain.exponentialRampToValueAtTime(0.0001, _actx.currentTime + dur);
    o.connect(g); g.connect(_actx.destination);
    o.start(); o.stop(_actx.currentTime + dur);
  } catch (e) { /* ignore */ }
}

function calmStreak(sessions, today = dateKey(Date.now())) {
  const days = new Set(sessions.map((s) => s.date));
  let cur = 0;
  const d = new Date(today + "T00:00:00");
  if (!days.has(today)) d.setDate(d.getDate() - 1);
  for (let i = 0; i < 800; i++) {
    if (days.has(dateKey(d.getTime()))) { cur += 1; d.setDate(d.getDate() - 1); } else break;
  }
  return cur;
}
const calmMinutes = (sessions) => Math.round(sessions.reduce((s, x) => s + (x.durationSec || 0), 0) / 60);
const fmtClock = (sec) => `${Math.floor(sec / 60)}:${String(Math.max(0, Math.round(sec % 60))).padStart(2, "0")}`;

/* ================================================================== */
/* FASTING — Dr. Fung intermittent-fasting tracker (calm:*→fasting:*) */
/* ================================================================== */
const FKEYS = {
  goals: "fasting:goals",
  diary: "fasting:diary",
  current: "fasting:currentFast",
  settings: "fasting:settings",
};

// Protocol ladder, gentlest → hardest (level drives intensity color)
const PROTOCOLS = [
  { id: "16:8", label: "16:8", hrs: 16, window: 8, freq: "Щодня", level: 1, note: "Старт для початківців. Пропускаєш сніданок, їси з 12:00 до 20:00. М'яко і стало." },
  { id: "18:6", label: "18:6", hrs: 18, window: 6, freq: "Щодня", level: 2, note: "Наступний крок після 16:8. Вужче вікно їжі, сильніший ефект." },
  { id: "20:4", label: "20:4", hrs: 20, window: 4, freq: "Щодня", level: 3, note: "«Дієта воїна». Один-два прийоми їжі у 4-годинному вікні." },
  { id: "OMAD", label: "OMAD (23:1)", hrs: 23, window: 1, freq: "Щодня / кілька разів на тиждень", level: 4, note: "Один прийом їжі на день. Простий графік, вимагає повноцінного прийому." },
  { id: "24h", label: "24 год", hrs: 24, window: 0, freq: "2–3 рази/тиждень", level: 5, note: "Від вечері до вечері. Один день без їжі, наступний — звичайно." },
  { id: "36h", label: "36 год", hrs: 36, window: 0, freq: "Через день", level: 6, note: "Сильніший вплив на інсулін і вагу. Часто в програмах Фанга при діабеті 2 типу." },
  { id: "42h", label: "42 год", hrs: 42, window: 0, freq: "2–3 рази/тиждень", level: 7, note: "Пропускаєш вечерю, весь наступний день і снідаєш через день." },
  { id: "48h+", label: "Тривале (>48 год)", hrs: 48, window: 0, freq: "Рідко / під наглядом", level: 8, note: "3–7+ днів. Лише з електролітами й бажано під наглядом лікаря." },
];
const getProtocol = (id) => PROTOCOLS.find((p) => p.id === id) || PROTOCOLS[0];
const protocolColor = (level) => (level <= 1 ? "#22c55e" : level <= 2 ? "#84cc16" : level <= 3 ? "#eab308" : level <= 4 ? "#f59e0b" : level <= 5 ? "#f97316" : level <= 6 ? "#ef4444" : "#dc2626");

// Gentle, educational stage timeline (not medical claims)
const FAST_STAGES = [
  { from: 0, to: 4, title: "Ситість", desc: "Тіло перетравлює їжу, цукор у крові зростає.", color: "#f59e0b" },
  { from: 4, to: 12, title: "Цукор спадає", desc: "Рівень цукру стабілізується, витрачається запас глікогену.", color: "#f97316" },
  { from: 12, to: 16, title: "Перехід на жир", desc: "Глікоген вичерпується — тіло починає брати енергію з жиру.", color: "#14b8a6" },
  { from: 16, to: 24, title: "Спалювання жиру", desc: "Кетоз зростає, запускається рання автофагія — м'яке очищення клітин.", color: "#0ea5e9" },
  { from: 24, to: 999, title: "Глибша автофагія", desc: "Довше голодування — глибші відновні процеси в клітинах.", color: "#6366f1" },
];
const stageForHours = (h) => FAST_STAGES.find((s) => h >= s.from && h < s.to) || FAST_STAGES[FAST_STAGES.length - 1];

const WD_UA = ["Нд", "Пн", "Вт", "Ср", "Чт", "Пт", "Сб"];

async function loadFastingData() {
  const goals = await store.get(FKEYS.goals, { startWeight: null, targetWeight: null, protocol: "16:8", startDate: dateKey(Date.now()) });
  const diary = await store.get(FKEYS.diary, []);
  const current = await store.get(FKEYS.current, null);
  const settings = await store.get(FKEYS.settings, { name: "Fasting" });
  return { goals, diary, current, settings };
}
async function collectFastingExport() {
  const d = await loadFastingData();
  return { goals: d.goals, diary: d.diary, current: d.current, settings: d.settings };
}
async function clearFastingData() {
  for (const k of Object.values(FKEYS)) await store.remove(k);
}

// diary sorted by date asc, with auto weightChange vs previous weighed entry
function diarySorted(diary) {
  const rows = [...diary].sort((a, b) => (a.date < b.date ? -1 : a.date > b.date ? 1 : (a.ts || 0) - (b.ts || 0)));
  let prevW = null;
  return rows.map((r) => {
    const change = r.weight != null && prevW != null ? Math.round((r.weight - prevW) * 10) / 10 : null;
    if (r.weight != null) prevW = r.weight;
    return { ...r, weightChange: change };
  });
}
function fastingMetrics(goals, diary) {
  const rows = diarySorted(diary);
  const weighed = rows.filter((r) => r.weight != null);
  const currentWeight = weighed.length ? weighed[weighed.length - 1].weight : (goals.startWeight ?? null);
  const start = goals.startWeight ?? currentWeight;
  const withHrs = diary.filter((r) => r.actualHrs != null && r.actualHrs > 0);
  const totalHrs = withHrs.reduce((s, r) => s + r.actualHrs, 0);
  return {
    currentWeight,
    weightChange: start != null && currentWeight != null ? Math.round((currentWeight - start) * 10) / 10 : null,
    remaining: currentWeight != null && goals.targetWeight != null ? Math.round((currentWeight - goals.targetWeight) * 10) / 10 : null,
    totalFasts: withHrs.length,
    avgHrs: withHrs.length ? Math.round((totalHrs / withHrs.length) * 10) / 10 : 0,
    longestHrs: withHrs.length ? Math.max(...withHrs.map((r) => r.actualHrs)) : 0,
    totalHrs: Math.round(totalHrs),
  };
}
function fastingStreak(diary, today = dateKey(Date.now())) {
  const met = new Set(diary.filter((r) => r.goalMet).map((r) => r.date));
  let cur = 0; const d = new Date(today + "T00:00:00");
  if (!met.has(today)) d.setDate(d.getDate() - 1);
  for (let i = 0; i < 800; i++) { if (met.has(dateKey(d.getTime()))) { cur += 1; d.setDate(d.getDate() - 1); } else break; }
  return cur;
}
const fmtHM = (ms) => { const m = Math.max(0, Math.floor(ms / 60000)); return `${Math.floor(m / 60)}г ${m % 60}хв`; };
const fmtHMS = (ms) => { const t = Math.max(0, Math.floor(ms / 1000)); const h = Math.floor(t / 3600), m = Math.floor((t % 3600) / 60), s = t % 60; return `${h}г ${m}хв ${s}с`; };

/* ================================================================== */
/* TOOLKIT — Anxiety toolkit library + Chore Splitter wizard          */
/* ================================================================== */
const TKEYS = {
  settings: "toolkit:settings",
  anxFav: "toolkit:anxiety:favorites",
  anxTried: "toolkit:anxiety:tried",
  anxWeek: "toolkit:anxiety:weekFocus",
  members: "toolkit:chores:members",
  list: "toolkit:chores:list",
  ratings: "toolkit:chores:ratings",
  assignments: "toolkit:chores:assignments",
  meta: "toolkit:chores:meta",
};

// Anxiety toolkit — a calm library of techniques (interactive versions live in Calm)
const ANX_TECHNIQUES = [
  { id: "breath", emoji: "🌬️", name: "Дихання (box / 4-7-8)", what: "Повільний ритм вдих-затримка-видих.", why: "Активує парасимпатику — тіло розуміє, що можна розслабитись." },
  { id: "ground", emoji: "⚓", name: "Заземлення 5-4-3-2-1", what: "Назви 5 що бачиш, 4 чуєш, 3 торкаєшся, 2 нюхаєш, 1 смакуєш.", why: "Повертає з тривожних думок у теперішній момент." },
  { id: "pmr", emoji: "💪", name: "Розслаблення м'язів", what: "Напруж і відпусти групи м'язів по черзі.", why: "Знімає фізичну напругу, яку тримає тіло під час тривоги." },
  { id: "thought", emoji: "📝", name: "Запис думки (CBT)", what: "Ситуація → думка → докази за/проти → збалансована думка.", why: "Розплутує автоматичну тривожну думку й повертає перспективу." },
  { id: "fear", emoji: "🪜", name: "Сходинки страху", what: "Список страхів від легкого до важкого, крок за кроком.", why: "Поступова експозиція м'яко зменшує уникання й тривогу." },
  { id: "worry", emoji: "⏳", name: "Час для тривоги", what: "Виділи 10 хв, щоб хвилюватись навмисно — потім стоп.", why: "Контейнує хвилювання, замість того щоб воно розтікалось на весь день." },
  { id: "focus", emoji: "⏱️", name: "Фокус-таймер", what: "Спокійний вдих, потім блок роботи й коротка перерва.", why: "Знижує тривогу «нічого не встигаю» через маленькі кроки." },
  { id: "self", emoji: "💛", name: "Самоспівчуття", what: "Скажи собі те, що сказала б другові в такій ситуації.", why: "Пом'якшує внутрішнього критика, який підживлює тривогу." },
];

const CHORE_ROOMS = [
  { room: "Спальня", items: ["Застелити ліжко", "Змінити білизну", "Пропилососити", "Розкласти одяг", "Витерти пил"] },
  { room: "Ванна", items: ["Помити раковину", "Помити унітаз", "Помити душ/ванну", "Дзеркало", "Поміняти рушники", "Помити підлогу"] },
  { room: "Вітальня/їдальня", items: ["Пропилососити", "Витерти пил", "Прибрати зі столу", "Помити підлогу", "Скласти речі"] },
  { room: "Кухня", items: ["Помити посуд", "Витерти столи", "Помити плиту", "Розібрати продукти", "Помити підлогу"] },
  { room: "Інше", items: ["Винести сміття", "Посуд у машину", "Коридор", "Полити рослини", "Погодувати тварин", "Прання: завантажити", "Прання: розвісити", "Прання: скласти", "Прасування"] },
];

const ATT = { like: { label: "Подобається", emoji: "🙂", weight: 2, color: "#22c55e" }, tolerable: { label: "Терпимо", emoji: "😐", weight: 5, color: "#eab308" }, hate: { label: "Ненавиджу", emoji: "😖", weight: 9, color: "#ef4444" } };
const ATT_ORDER = ["like", "tolerable", "hate"];

async function loadToolkitData() {
  const settings = await store.get(TKEYS.settings, { name: "Toolkit" });
  const favorites = await store.get(TKEYS.anxFav, []);
  const tried = await store.get(TKEYS.anxTried, []);
  const weekFocus = await store.get(TKEYS.anxWeek, null);
  const members = await store.get(TKEYS.members, null);
  const list = await store.get(TKEYS.list, []);
  const ratings = await store.get(TKEYS.ratings, {});
  const assignments = await store.get(TKEYS.assignments, {});
  const meta = await store.get(TKEYS.meta, { step: 1, useScores: false, finished: false });
  return { settings, favorites, tried, weekFocus, members, list, ratings, assignments, meta };
}
async function collectToolkitExport() {
  const d = await loadToolkitData();
  return { settings: d.settings, anxiety: { favorites: d.favorites, tried: d.tried, weekFocus: d.weekFocus }, chores: { members: d.members, list: d.list, ratings: d.ratings, assignments: d.assignments, meta: d.meta } };
}
async function clearToolkitData() {
  for (const k of Object.values(TKEYS)) await store.remove(k);
}
// effective difficulty score for member on chore (score if using 1-10, else attitude weight)
function choreScore(ratings, choreId, memberId, useScores) {
  const r = ratings[choreId]?.[memberId];
  if (!r) return null;
  if (useScores && r.score != null) return r.score;
  return ATT[r.attitude]?.weight ?? null;
}

/* ------------------------------------------------------------------ */
/* Small presentational pieces                                         */
/* ------------------------------------------------------------------ */
function StatTile({ icon: Icon, label, value, sub, tint = "text-slate-700" }) {
  return (
    <div className="flex-1 min-w-[140px] rounded-xl border border-slate-200 bg-white p-4 shadow-sm">
      <div className="flex items-center gap-2 text-slate-400">
        <Icon className="h-4 w-4" />
        <span className="text-xs font-medium uppercase tracking-wide">{label}</span>
      </div>
      <div className={`mt-2 text-3xl font-bold tabular-nums ${tint}`}>{value}</div>
      {sub && <div className="mt-0.5 text-xs text-slate-400">{sub}</div>}
    </div>
  );
}

function CountPill({ n, cls }) {
  if (!n) return null;
  return (
    <span className={`inline-flex min-w-[20px] justify-center rounded-full px-1.5 py-0.5 text-[11px] font-semibold tabular-nums ${cls}`}>
      {n}
    </span>
  );
}

/* ------------------------------------------------------------------ */
/* Main component                                                      */
/* ------------------------------------------------------------------ */
export default function FlashcardsApp() {
  const [loading, setLoading] = useState(true);
  const [loadMsg, setLoadMsg] = useState("Opening your library…");
  const [decks, setDecks] = useState([]); // [{id,name,created,groupId,language,goal,deadline,autoPlay,...}]
  const [groups, setGroups] = useState([]); // [{id,name,emoji,color,collapsed}]
  const [cardsByDeck, setCardsByDeck] = useState({}); // {deckId: [cards]}
  const [stats, setStats] = useState({ history: {}, settings: { newPerDay: DEFAULT_NEW_PER_DAY } });
  const [view, setView] = useState("home"); // home | deck | study | setup | import | stats
  const [section, setSection] = useState("studying"); // studying | routine | calm | fasting
  const [sidebarCollapsed, setSidebarCollapsed] = useState(false);
  const [calmName, setCalmName] = useState("Спокій");
  const [fastingName, setFastingName] = useState("Fasting");
  const [mgmtName, setMgmtName] = useState("Менеджмент");
  const [toolkitName, setToolkitName] = useState("Toolkit");
  const [budgetName, setBudgetName] = useState("Budget");
  const [inventoryName, setInventoryName] = useState("Inventory");
  const [financeName, setFinanceName] = useState("Finance");
  const [cloudState, setCloudState] = useState({ signedIn: false, email: null, syncing: false });
  const [toast, setToast] = useState(null);
  const [deckEditor, setDeckEditor] = useState(null); // null | { deck } (deck=null → create)
  const [groupEditor, setGroupEditor] = useState(null); // null | { group } (group=null → create)
  const [cardEditor, setCardEditor] = useState(null); // null | { deckId, card } (card=null → new)
  const [deckDetailId, setDeckDetailId] = useState(null); // deck being browsed

  const newPerDay = stats.settings?.newPerDay ?? DEFAULT_NEW_PER_DAY;

  /* ---------- initial progressive load ---------- */
  useEffect(() => {
    let alive = true;
    (async () => {
      const idx = await store.get("decks:index", { decks: [] });
      const gi = await store.get("groups:index", { groups: [] });
      // one-time: seed FR/PL/ES starter decks under a "Мови" group
      if (!(await store.get("langs:seeded", false))) {
        try {
          const seed = await (await fetch("/lang-starters.json")).json();
          const gid = uid("g");
          const newDecks = [];
          for (const d of seed.decks) {
            const did = uid("d");
            newDecks.push({ id: did, name: d.name, created: Date.now(), topic: "", description: "", emoji: d.emoji, color: d.color || "blue", groupId: gid, language: d.language, autoPlay: false, goal: "longterm", deadline: null });
            await store.set(`cards:${did}`, d.cards.map(([f, b]) => makeCard(f, b)));
          }
          gi.groups = [...(gi.groups || []), { id: gid, name: "Мови", emoji: "🌍", color: "blue", collapsed: false }];
          idx.decks = [...(idx.decks || []), ...newDecks];
          await store.set("groups:index", { groups: gi.groups });
          await store.set("decks:index", { decks: idx.decks });
          await store.set("langs:seeded", true);
        } catch (e) { /* offline: skip seeding */ }
      }
      const ui = await store.get("ui:prefs", { section: "review", sidebarCollapsed: false });
      const calmSettings = await store.get(CKEYS.settings, { name: "Спокій" });
      const fastingSettings = await store.get(FKEYS.settings, { name: "Fasting" });
      const mgmtSettings = await store.get("mgmt:settings", { name: "Менеджмент" });
      const toolkitSettings = await store.get(TKEYS.settings, { name: "Toolkit" });
      const budgetSettings = await store.get(BKEYS.settings, { name: "Budget" });
      const inventorySettings = await store.get(IKEYS.settings, { name: "Inventory" });
      const financeSettings = await store.get(FNKEYS.settings, { name: "Finance" });
      const st = await store.get("stats", {
        history: {},
        settings: { newPerDay: DEFAULT_NEW_PER_DAY },
      });
      if (!alive) return;
      setSection(["review", "routine", "calm", "fasting", "management", "inventory", "money"].includes(ui.section) ? ui.section : "studying");
      setSidebarCollapsed(!!ui.sidebarCollapsed);
      setCalmName(calmSettings?.name && calmSettings.name !== "Calm" ? calmSettings.name : "Спокій");
      setFastingName(fastingSettings?.name || "Fasting");
      setMgmtName(mgmtSettings?.name || "Менеджмент");
      setToolkitName(toolkitSettings?.name || "Toolkit");
      setBudgetName(budgetSettings?.name || "Budget");
      setInventoryName(inventorySettings?.name || "Inventory");
      setFinanceName(financeSettings?.name || "Finance");
      setStats(st);
      setGroups(gi.groups || []);
      setDecks(idx.decks || []);
      const map = {};
      for (const d of idx.decks || []) {
        setLoadMsg(`Loading “${d.name}”…`);
        map[d.id] = await store.get(`cards:${d.id}`, []);
        if (!alive) return;
        setCardsByDeck({ ...map });
      }
      setLoading(false);
    })();
    return () => { alive = false; };
  }, []);

  const flash = useCallback((msg) => {
    setToast(msg);
    window.clearTimeout(flash._t);
    flash._t = window.setTimeout(() => setToast(null), 2200);
  }, []);

  /* ---------- persistence helpers ---------- */
  const persistIndex = useCallback(async (nextDecks) => {
    await store.set("decks:index", { decks: nextDecks });
  }, []);
  const persistDeckCards = useCallback(async (deckId, cards) => {
    await store.set(`cards:${deckId}`, cards);
  }, []);
  const persistStats = useCallback(async (nextStats) => {
    await store.set("stats", nextStats);
  }, []);

  const changeSection = useCallback((next) => {
    setSection(next);
    store.set("ui:prefs", { section: next, sidebarCollapsed });
  }, [sidebarCollapsed]);

  const toggleSidebar = useCallback(() => {
    setSidebarCollapsed((c) => {
      const next = !c;
      store.set("ui:prefs", { section, sidebarCollapsed: next });
      return next;
    });
  }, [section]);

  // cloud sync status for the sidebar indicator
  useEffect(() => {
    (async () => { await refreshSession(); setCloudState((s) => ({ ...s, signedIn: cloudSignedIn(), email: currentEmail() })); })();
  }, []);

  const doSyncNow = useCallback(async () => {
    if (!cloudSignedIn()) { changeSection("studying"); setView("stats"); return; }
    setCloudState((s) => ({ ...s, syncing: true }));
    try { await syncNow(); location.reload(); }
    catch (e) { setCloudState((s) => ({ ...s, syncing: false })); flash("Помилка синхронізації"); }
  }, [changeSection, flash]);

  const renameCalm = useCallback(async (name) => {
    const clean = (name || "").trim() || "Спокій";
    setCalmName(clean);
    const prev = await store.get(CKEYS.settings, { name: "Спокій" });
    await store.set(CKEYS.settings, { ...prev, name: clean });
  }, []);

  const renameFasting = useCallback(async (name) => {
    const clean = (name || "").trim() || "Fasting";
    setFastingName(clean);
    const prev = await store.get(FKEYS.settings, { name: "Fasting" });
    await store.set(FKEYS.settings, { ...prev, name: clean });
  }, []);

  const renameMgmt = useCallback(async (name) => {
    const clean = (name || "").trim() || "Менеджмент";
    setMgmtName(clean);
    const prev = await store.get("mgmt:settings", { name: "Менеджмент" });
    await store.set("mgmt:settings", { ...prev, name: clean });
  }, []);

  const renameToolkit = useCallback(async (name) => {
    const clean = (name || "").trim() || "Toolkit";
    setToolkitName(clean);
    const prev = await store.get(TKEYS.settings, { name: "Toolkit" });
    await store.set(TKEYS.settings, { ...prev, name: clean });
  }, []);

  const renameBudget = useCallback(async (name) => {
    const clean = (name || "").trim() || "Budget";
    setBudgetName(clean);
    const prev = await store.get(BKEYS.settings, { name: "Budget" });
    await store.set(BKEYS.settings, { ...prev, name: clean });
  }, []);

  const renameInventory = useCallback(async (name) => {
    const clean = (name || "").trim() || "Inventory";
    setInventoryName(clean);
    const prev = await store.get(IKEYS.settings, { name: "Inventory" });
    await store.set(IKEYS.settings, { ...prev, name: clean });
  }, []);

  const renameFinance = useCallback(async (name) => {
    const clean = (name || "").trim() || "Finance";
    setFinanceName(clean);
    const prev = await store.get(FNKEYS.settings, { name: "Finance" });
    await store.set(FNKEYS.settings, { ...prev, name: clean });
  }, []);

  /* ---------- derived: per-deck due summary ---------- */
  const deckSummary = useMemo(() => {
    const now = Date.now();
    const endToday = new Date(now); endToday.setHours(23, 59, 59, 999);
    const endTs = endToday.getTime();
    const out = {};
    for (const d of decks) {
      const cards = cardsByDeck[d.id] || [];
      let learn = 0, review = 0, newTotal = 0, learned = 0;
      const stages = {}; // stageId -> count, across ALL cards (interval breakdown)
      for (const c of cards) {
        if (c.state === "new") newTotal += 1;
        else if (c.state === "learning" && c.due <= endTs) learn += 1;
        else if (c.state === "review" && c.due <= endTs) review += 1;
        // "learned" = mature card: reviewed and on a long (≥21d) interval
        if (c.state === "review" && (c.interval || 0) >= 21) learned += 1;
        const sid = stageForCard(c).id;
        stages[sid] = (stages[sid] || 0) + 1;
      }
      // Strict "due today" = only cards the schedule actually placed today or earlier.
      // New cards are NOT padded in — they're studied via the "only new"/"all" modes.
      out[d.id] = { learn, review, newTotal, newDue: newTotal, learned, total: cards.length, due: learn + review, stages };
    }
    return out;
  }, [decks, cardsByDeck, stats]);

  // roll up due counts + card totals per group
  const groupSummary = useMemo(() => {
    const out = {};
    for (const g of groups) out[g.id] = { due: 0, total: 0, deckCount: 0 };
    for (const d of decks) {
      const gid = d.groupId || "";
      if (!out[gid]) continue;
      const s = deckSummary[d.id] || { due: 0, total: 0 };
      out[gid].due += s.due;
      out[gid].total += s.total;
      out[gid].deckCount += 1;
    }
    return out;
  }, [groups, decks, deckSummary]);

  const totalDue = useMemo(
    () => Object.values(deckSummary).reduce((s, x) => s + x.due, 0),
    [deckSummary]
  );

  /* ---------- build a study queue across decks for a given mode ---------- */
  // Returns an array of { deckId, id } refs so a session can span many decks.
  const buildSessionQueue = useCallback(
    (deckIds, mode, count) => {
      const now = Date.now();
      const todayKey = dateKey(now);
      const introducedToday = stats.history?.[todayKey]?.newIntroduced || 0;
      const newRemaining = Math.max(0, newPerDay - introducedToday);

      const all = [];
      for (const did of deckIds)
        for (const c of cardsByDeck[did] || []) all.push({ deckId: did, card: c });

      const isNew = (x) => x.card.state === "new";
      const isLearn = (x) => x.card.state === "learning";
      const isReview = (x) => x.card.state === "review";
      const dueNow = (x) => x.card.due <= now;
      const toRef = (x) => ({ deckId: x.deckId, id: x.card.id });

      let out = [];
      if (mode === "new") {
        out = all.filter(isNew);
      } else if (mode === "review") {
        out = all.filter((x) => isReview(x) || isLearn(x)); // already-started cards, drilled ahead
      } else if (mode === "all" || mode === "cram") {
        const due = all.filter((x) => (isReview(x) || isLearn(x)) && dueNow(x));
        const news = all.filter(isNew);
        const rest = all.filter((x) => !due.includes(x) && !news.includes(x));
        out = [...due, ...news, ...rest]; // everything, due-first
      } else if (mode === "custom") {
        const review = all.filter((x) => isReview(x) && dueNow(x));
        const learn = all.filter((x) => isLearn(x) && dueNow(x));
        const news = all.filter(isNew);
        const future = all.filter((x) => isReview(x) && !dueNow(x));
        out = [...review, ...learn, ...news, ...future].slice(0, Math.max(1, count || 20));
      } else {
        // "due" — strictly the scheduled queue: cards actually due today or earlier,
        // however many, with NO new-card padding and NO cap. New cards are a separate
        // opt-in ("only new" / "all" / custom).
        const endToday = new Date(now); endToday.setHours(23, 59, 59, 999);
        const endTs = endToday.getTime();
        const dueToday = (x) => x.card.due <= endTs;
        const review = all.filter((x) => isReview(x) && dueToday(x));
        const learn = all.filter((x) => isLearn(x) && dueToday(x));
        out = [...review, ...learn];
      }
      return out.map(toRef);
    },
    [cardsByDeck, stats, newPerDay]
  );

  // How many cards a given mode would queue right now (for the setup screen)
  const countForMode = useCallback(
    (deckIds, mode, count) => buildSessionQueue(deckIds, mode, count).length,
    [buildSessionQueue]
  );

  /* ------------------------------------------------------------------ */
  /* Deck / card mutations                                              */
  /* ------------------------------------------------------------------ */
  const newDeck = (name, meta = {}) => ({
    id: uid("d"),
    name: (name || "Untitled deck").trim() || "Untitled deck",
    created: Date.now(),
    topic: (meta.topic || "").trim(),
    description: (meta.description || "").trim(),
    emoji: meta.emoji || "",
    color: meta.color || "indigo",
    groupId: meta.groupId || "",
    language: meta.language || "en-US",
    autoPlay: !!meta.autoPlay,
    goal: meta.goal || "longterm",
    deadline: meta.deadline || null,
  });

  const findOrCreateDeck = useCallback((name, workingDecks, meta) => {
    const trimmed = (name || "Untitled deck").trim() || "Untitled deck";
    const existing = workingDecks.find((d) => d.name.toLowerCase() === trimmed.toLowerCase());
    if (existing) return { deck: existing, decks: workingDecks, created: false };
    const deck = newDeck(trimmed, meta);
    return { deck, decks: [...workingDecks, deck], created: true };
  }, []);

  const createDeck = useCallback(
    async (meta) => {
      const deck = newDeck(meta.name, meta);
      const next = [...decks, deck];
      setDecks(next);
      setCardsByDeck((m) => ({ ...m, [deck.id]: [] }));
      await persistIndex(next);
      await persistDeckCards(deck.id, []);
      flash(`Deck “${deck.name}” created`);
      return deck;
    },
    [decks, persistIndex, persistDeckCards, flash]
  );

  const updateDeck = useCallback(
    async (deckId, patch) => {
      const next = decks.map((d) => (d.id === deckId ? { ...d, ...patch } : d));
      setDecks(next);
      await persistIndex(next);
      flash("Deck updated");
    },
    [decks, persistIndex, flash]
  );

  // groups: { deckName: [cardObjects] }.  opts.targetDeckId forces every card
  // into one existing deck (append). Importing NEVER replaces existing cards.
  const importCards = useCallback(
    async (groups, opts = {}) => {
      let workingDecks = [...decks];
      const nextByDeck = { ...cardsByDeck };
      let added = 0;
      for (const [deckName, cards] of Object.entries(groups)) {
        if (!cards.length) continue;
        let deck = opts.targetDeckId ? workingDecks.find((d) => d.id === opts.targetDeckId) : null;
        if (!deck) {
          const res = findOrCreateDeck(deckName, workingDecks, opts.newDeckMeta);
          workingDecks = res.decks;
          deck = res.deck;
        }
        const merged = [...(nextByDeck[deck.id] || []), ...cards];
        nextByDeck[deck.id] = merged;
        await persistDeckCards(deck.id, merged);
        added += cards.length;
      }
      setDecks(workingDecks);
      setCardsByDeck(nextByDeck);
      await persistIndex(workingDecks);
      flash(`Imported ${added.toLocaleString()} card${added === 1 ? "" : "s"}`);
      setView("home");
    },
    [decks, cardsByDeck, findOrCreateDeck, persistDeckCards, persistIndex, flash]
  );

  const deleteDeck = useCallback(
    async (deckId) => {
      const next = decks.filter((d) => d.id !== deckId);
      const nextByDeck = { ...cardsByDeck };
      // clean up any per-card images
      for (const c of cardsByDeck[deckId] || []) {
        if (c.imgFront) await removeImage(c.id, "front");
        if (c.imgBack) await removeImage(c.id, "back");
      }
      delete nextByDeck[deckId];
      setDecks(next);
      setCardsByDeck(nextByDeck);
      await persistIndex(next);
      await store.remove(`cards:${deckId}`);
      if (deckDetailId === deckId) { setDeckDetailId(null); setView("home"); }
      flash("Deck deleted");
    },
    [decks, cardsByDeck, persistIndex, flash, deckDetailId]
  );

  /* ---------- groups (folders) ---------- */
  const persistGroups = useCallback(async (g) => {
    await store.set("groups:index", { groups: g });
  }, []);

  const createGroup = useCallback(
    async (meta) => {
      const g = { id: uid("g"), name: (meta.name || "New group").trim() || "New group", emoji: meta.emoji || "", color: meta.color || "slate", collapsed: false };
      const next = [...groups, g];
      setGroups(next);
      await persistGroups(next);
      flash(`Group “${g.name}” created`);
      return g;
    },
    [groups, persistGroups, flash]
  );

  const updateGroup = useCallback(
    async (id, patch) => {
      const next = groups.map((g) => (g.id === id ? { ...g, ...patch } : g));
      setGroups(next);
      await persistGroups(next);
    },
    [groups, persistGroups]
  );

  const deleteGroup = useCallback(
    async (id) => {
      const nextDecks = decks.map((d) => (d.groupId === id ? { ...d, groupId: "" } : d));
      setDecks(nextDecks);
      await persistIndex(nextDecks);
      const next = groups.filter((g) => g.id !== id);
      setGroups(next);
      await persistGroups(next);
      flash("Group deleted — its decks were kept");
    },
    [decks, groups, persistIndex, persistGroups, flash]
  );

  const toggleGroup = useCallback((id) => {
    setGroups((gs) => {
      const next = gs.map((g) => (g.id === id ? { ...g, collapsed: !g.collapsed } : g));
      store.set("groups:index", { groups: next });
      return next;
    });
  }, []);

  /* ---------- individual cards (create / edit / delete, with images) ---------- */
  const saveCard = useCallback(
    async (deckId, cardId, fields, images) => {
      let deckCards = [...(cardsByDeck[deckId] || [])];
      let id = cardId;
      const isNew = !id;
      if (isNew) {
        const c = makeCard(fields.front, fields.back);
        id = c.id;
        deckCards = [...deckCards, c];
      }
      const patch = { front: (fields.front || "").trim(), back: (fields.back || "").trim(), lang: fields.lang || "" };
      // images: {front|back: dataUrl(new) | null(remove) | undefined(unchanged)}
      for (const side of ["front", "back"]) {
        const img = images?.[side];
        if (img === undefined) continue;
        const flag = side === "front" ? "imgFront" : "imgBack";
        if (img === null) { await removeImage(id, side); patch[flag] = false; }
        else { await saveImage(id, side, img); patch[flag] = true; }
      }
      deckCards = deckCards.map((c) => (c.id === id ? { ...c, ...patch } : c));
      const nextByDeck = { ...cardsByDeck, [deckId]: deckCards };
      setCardsByDeck(nextByDeck);
      await persistDeckCards(deckId, deckCards);
      flash(isNew ? "Card added" : "Card saved");
    },
    [cardsByDeck, persistDeckCards, flash]
  );

  const deleteCard = useCallback(
    async (deckId, cardId) => {
      const card = (cardsByDeck[deckId] || []).find((c) => c.id === cardId);
      if (card?.imgFront) await removeImage(cardId, "front");
      if (card?.imgBack) await removeImage(cardId, "back");
      const deckCards = (cardsByDeck[deckId] || []).filter((c) => c.id !== cardId);
      setCardsByDeck((m) => ({ ...m, [deckId]: deckCards }));
      await persistDeckCards(deckId, deckCards);
      flash("Card deleted");
    },
    [cardsByDeck, persistDeckCards, flash]
  );

  /* ---------- study scope resolution (all / group / single deck) ---------- */
  const scopeToDeckIds = useCallback(
    (scope) => {
      if (scope === "all") return decks.map((d) => d.id);
      if (typeof scope === "string" && scope.startsWith("group:")) {
        const gid = scope.slice(6);
        return decks.filter((d) => (d.groupId || "") === gid).map((d) => d.id);
      }
      return [scope];
    },
    [decks]
  );

  const scopeName = useCallback(
    (scope) => {
      if (scope === "all") return "All decks";
      if (typeof scope === "string" && scope.startsWith("group:")) {
        return groups.find((g) => g.id === scope.slice(6))?.name || "Group";
      }
      return decks.find((d) => d.id === scope)?.name || "Deck";
    },
    [decks, groups]
  );

  const deckOptsFor = useCallback(
    (deckId) => {
      const d = decks.find((x) => x.id === deckId);
      return d ? { goal: d.goal || "longterm", deadline: d.deadline || null } : null;
    },
    [decks]
  );

  const resetAll = useCallback(async () => {
    for (const d of decks) {
      await store.remove(`cards:${d.id}`);
      for (const c of cardsByDeck[d.id] || []) {
        if (c.imgFront) await removeImage(c.id, "front");
        if (c.imgBack) await removeImage(c.id, "back");
      }
    }
    await store.remove("decks:index");
    await store.remove("groups:index");
    await store.remove("langs:seeded");
    await store.remove("ui:prefs");
    await store.remove("stats");
    await clearRoutineData();
    await clearCalmData();
    await clearFastingData();
    await store.remove("mgmt:settings");
    await clearCareerData();
    await clearToolkitData();
    await clearBudgetData();
    await clearInventoryData();
    await clearReviewData();
    await clearFinanceData();
    setDecks([]);
    setGroups([]);
    setCardsByDeck({});
    setStats({ history: {}, settings: { newPerDay: DEFAULT_NEW_PER_DAY } });
    setView("home");
    setCalmName("Спокій");
    setFastingName("Fasting");
    setMgmtName("Менеджмент");
    setToolkitName("Toolkit");
    setBudgetName("Budget");
    setInventoryName("Inventory");
    setFinanceName("Finance");
    flash("All data reset");
    window.dispatchEvent(new CustomEvent("routine-reset"));
    window.dispatchEvent(new CustomEvent("calm-reset"));
    window.dispatchEvent(new CustomEvent("fasting-reset"));
    window.dispatchEvent(new CustomEvent("toolkit-reset"));
    window.dispatchEvent(new CustomEvent("budget-reset"));
    window.dispatchEvent(new CustomEvent("inventory-reset"));
    window.dispatchEvent(new CustomEvent("review-reset"));
    window.dispatchEvent(new CustomEvent("finance-reset"));
  }, [decks, cardsByDeck, flash]);

  const exportAll = useCallback(async () => {
    const routine = await collectRoutineExport();
    const calm = await collectCalmExport();
    const fasting = await collectFastingExport();
    const mgmt = await store.get("mgmt:settings", { name: "Менеджмент" });
    const career = await collectCareerExport();
    const toolkit = await collectToolkitExport();
    const budget = await collectBudgetExport();
    const inventory = await collectInventoryExport();
    const review = await collectReviewExport();
    const finance = await collectFinanceExport();
    const payload = {
      exportedAt: new Date().toISOString(),
      version: 11,
      decks,
      groups,
      cards: cardsByDeck,
      stats,
      routine,
      calm,
      fasting,
      mgmt,
      career,
      toolkit,
      budget,
      inventory,
      review,
      finance,
    };
    const blob = new Blob([JSON.stringify(payload, null, 2)], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `flashcards-backup-${dateKey(Date.now())}.json`;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
    flash("Backup downloaded");
  }, [decks, groups, cardsByDeck, stats, flash]);

  const loadSample = useCallback(async () => {
    const groups = {
      "World Capitals": [
        makeCard("Capital of Japan", "Tokyo", "geography"),
        makeCard("Capital of Australia", "Canberra", "geography"),
        makeCard("Capital of Canada", "Ottawa", "geography"),
        makeCard("Capital of Brazil", "Brasília", "geography"),
        makeCard("Capital of Egypt", "Cairo", "geography"),
      ],
      "Spanish Basics": [
        makeCard("hello", "hola", "greetings"),
        makeCard("thank you", "gracias", "greetings"),
        makeCard("water", "agua", "nouns"),
        makeCard("to eat", "comer", "verbs"),
        makeCard("beautiful", "hermoso / hermosa", "adjectives"),
      ],
    };
    await importCards(groups);
  }, [importCards]);

  const [loadingEnglish, setLoadingEnglish] = useState(false);
  const importEnglishDecks = useCallback(async () => {
    setLoadingEnglish(true);
    try {
      const res = await fetch("/english-decks.json");
      if (!res.ok) throw new Error("not found");
      const data = await res.json();
      const g = await createGroup({ name: "English", emoji: "🇬🇧", color: "blue" });
      const groups = {};
      for (const [name, cards] of Object.entries(data)) groups[name] = cards.map(([f, b]) => makeCard(f, b));
      await importCards(groups, { newDeckMeta: { groupId: g.id } });
    } catch (e) {
      flash("Не вдалося завантажити колоди");
    } finally {
      setLoadingEnglish(false);
    }
  }, [createGroup, importCards, flash]);

  /* ------------------------------------------------------------------ */
  /* Study session                                                      */
  /* ------------------------------------------------------------------ */
  const [session, setSession] = useState(null);
  // session = { title, mode, cram, queue:[{deckId,id}], total, flipped, done, again, correct }
  const [setup, setSetup] = useState(null); // { deckScope:'all'|deckId, mode, count }

  const openSetup = useCallback((deckScope) => {
    setSetup({ deckScope: deckScope || "all", mode: "due", count: 50 });
    setView("setup");
  }, []);

  const startSession = useCallback(
    (config) => {
      const deckIds = scopeToDeckIds(config.deckScope);
      const queue = buildSessionQueue(deckIds, config.mode, config.count);
      if (!queue.length) {
        flash("No cards match that mode right now");
        return;
      }
      const modeLabel = STUDY_MODES.find((m) => m.id === config.mode)?.label || "Study";
      setSession({
        title: scopeName(config.deckScope),
        subtitle: modeLabel,
        mode: config.mode,
        cram: config.mode === "cram",
        reverse: !!config.reverse,
        queue,
        total: queue.length,
        flipped: false,
        done: 0,
        again: 0,
        correct: 0,
      });
      setView("study");
    },
    [scopeToDeckIds, scopeName, buildSessionQueue, flash]
  );

  const currentRef = session && session.queue.length ? session.queue[0] : null;

  const currentCard = useMemo(() => {
    if (!currentRef) return null;
    return (cardsByDeck[currentRef.deckId] || []).find((c) => c.id === currentRef.id) || null;
  }, [currentRef, cardsByDeck]);

  const gradePreviews = useMemo(() => {
    if (!currentCard || session?.cram) return {};
    const now = Date.now();
    const opts = currentRef ? deckOptsFor(currentRef.deckId) : null;
    const out = {};
    for (const g of GRADES) {
      const r = schedule(currentCard, g.key, now, opts);
      out[g.key] = formatDelta(r.due - now);
    }
    return out;
  }, [currentCard, session, currentRef, deckOptsFor]);

  const answer = useCallback(
    async (grade) => {
      if (!session || !currentCard) return;
      const now = Date.now();
      const ref = session.queue[0];

      const advance = (rest) =>
        setSession((s) => ({
          ...s,
          queue: rest,
          flipped: false,
          done: s.done + 1,
          again: s.again + (grade === "again" ? 1 : 0),
          correct: s.correct + (grade !== "again" ? 1 : 0),
        }));

      // Cram: drill only — never touches SM-2 state, due dates, or stats.
      if (session.cram) {
        const rest = session.queue.slice(1);
        if (grade === "again" || grade === "hard") rest.push(ref);
        advance(rest);
        return;
      }

      // Normal grading — SM-2 updates even when reviewed ahead of schedule.
      const wasNew = currentCard.state === "new";
      const mature = currentCard.state === "review";
      const updated = schedule(currentCard, grade, now, deckOptsFor(ref.deckId));
      const deckCards = cardsByDeck[ref.deckId] || [];
      const nextCards = deckCards.map((c) => (c.id === updated.id ? updated : c));
      const nextByDeck = { ...cardsByDeck, [ref.deckId]: nextCards };
      setCardsByDeck(nextByDeck);

      const nextStats = bumpStats(stats, { grade, wasNew, mature });
      setStats(nextStats);

      const rest = session.queue.slice(1);
      if (updated.due - now <= SESSION_REQUEUE_WINDOW) rest.push(ref);
      advance(rest);

      await persistDeckCards(ref.deckId, nextCards);
      await persistStats(nextStats);
    },
    [session, currentCard, cardsByDeck, stats, persistDeckCards, persistStats, deckOptsFor]
  );

  /* ---------- keyboard shortcuts during study ---------- */
  useEffect(() => {
    if (view !== "study" || !session) return;
    const onKey = (e) => {
      if (e.repeat) return;
      const flippedNow = session.flipped;
      if (e.code === "Space" || e.key === "Enter") {
        e.preventDefault();
        if (!flippedNow) setSession((s) => ({ ...s, flipped: true }));
        return;
      }
      if (flippedNow && ["1", "2", "3", "4"].includes(e.key)) {
        e.preventDefault();
        answer(GRADES[Number(e.key) - 1].key);
      }
    };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, [view, session, answer]);

  const currentDeck = useMemo(
    () => (currentRef ? decks.find((d) => d.id === currentRef.deckId) || null : null),
    [decks, currentRef]
  );

  /* end session when the queue drains */
  const sessionFinished = session && session.queue.length === 0;

  const detailDeck = deckDetailId ? decks.find((d) => d.id === deckDetailId) : null;

  /* ------------------------------------------------------------------ */
  /* Render: loading                                                    */
  /* ------------------------------------------------------------------ */
  if (loading) {
    return (
      <div className="flex min-h-screen items-center justify-center bg-slate-50 text-slate-500">
        <div className="flex flex-col items-center gap-3">
          <Brain className="h-8 w-8 animate-pulse text-rose-500" />
          <div className="text-sm">{loadMsg}</div>
        </div>
      </div>
    );
  }

  /* ------------------------------------------------------------------ */
  /* Render                                                             */
  /* ------------------------------------------------------------------ */
  return (
    <div className="flex min-h-screen bg-slate-50 font-sans text-slate-800 antialiased">
      <Sidebar
        section={section}
        collapsed={sidebarCollapsed}
        onSection={changeSection}
        onToggle={toggleSidebar}
        studyingDue={totalDue}
        calmName={calmName}
        fastingName={fastingName}
        mgmtName={mgmtName}
        toolkitName={toolkitName}
        budgetName={budgetName}
        inventoryName={inventoryName}
        financeName={financeName}
        cloud={cloudState}
        onSyncNow={doSyncNow}
      />

      <div className="flex min-h-screen min-w-0 flex-1 flex-col pb-16 lg:pb-0">
        {section === "review" ? (
          <ReviewSection onGo={changeSection} />
        ) : section === "routine" ? (
          <RoutineSection />
        ) : section === "calm" ? (
          <CalmSection name={calmName} onRename={renameCalm} />
        ) : section === "fasting" ? (
          <FastingSection name={fastingName} onRename={renameFasting} />
        ) : section === "management" ? (
          <ManagementSection name={mgmtName} onRename={renameMgmt} />
        ) : section === "money" ? (
          <MoneySection budgetName={budgetName} renameBudget={renameBudget} financeName={financeName} renameFinance={renameFinance} onGo={changeSection} />
        ) : section === "inventory" ? (
          <InventorySection name={inventoryName} onRename={renameInventory} />
        ) : (
        <>
      {/* studying top bar */}
      <header className="sticky top-0 z-20 border-b border-slate-200 bg-white/90 backdrop-blur">
        <div className="mx-auto flex h-14 w-full max-w-5xl items-center gap-2 px-4">
          <button
            onClick={() => { setView("home"); setSession(null); }}
            className="mr-auto flex items-center gap-2 font-semibold tracking-tight text-slate-900"
          >
            <span className="text-base">Studying</span>
            {totalDue > 0 && view === "home" && (
              <CountPill n={totalDue} cls="bg-rose-100 text-rose-700 ml-1" />
            )}
          </button>

          <NavButton active={view === "home"} onClick={() => { setView("home"); setSession(null); }} icon={Layers}>
            Decks
          </NavButton>
          <NavButton active={view === "stats"} onClick={() => setView("stats")} icon={BarChart3}>
            Stats
          </NavButton>
          <NavButton active={view === "import"} onClick={() => setView("import")} icon={Upload}>
            Import
          </NavButton>
        </div>
      </header>

      <main className="mx-auto w-full max-w-5xl px-4 py-6">
        {view === "home" && (
          <HomeView
            decks={decks}
            groups={groups}
            summary={deckSummary}
            groupSummary={groupSummary}
            totalDue={totalDue}
            stats={stats}
            onStudy={openSetup}
            onStudyAll={() => openSetup("all")}
            onStudyGroup={(gid) => openSetup(`group:${gid}`)}
            onOpenDeck={(id) => { setDeckDetailId(id); setView("deck"); }}
            onDelete={deleteDeck}
            onEdit={(deck) => setDeckEditor({ deck })}
            onNewDeck={() => setDeckEditor({ deck: null })}
            onNewGroup={() => setGroupEditor({ group: null })}
            onEditGroup={(group) => setGroupEditor({ group })}
            onDeleteGroup={deleteGroup}
            onToggleGroup={toggleGroup}
            onMoveDeck={(deckId, groupId) => updateDeck(deckId, { groupId })}
            onImport={() => setView("import")}
            onSample={loadSample}
            onLoadEnglish={importEnglishDecks}
            loadingEnglish={loadingEnglish}
          />
        )}
        {view === "deck" && detailDeck && (
          <DeckDetailView
            deck={detailDeck}
            cards={cardsByDeck[detailDeck.id] || []}
            summary={deckSummary[detailDeck.id]}
            onBack={() => { setView("home"); setDeckDetailId(null); }}
            onStudy={() => openSetup(detailDeck.id)}
            onEditDeck={() => setDeckEditor({ deck: detailDeck })}
            onAddCard={() => setCardEditor({ deckId: detailDeck.id, card: null })}
            onEditCard={(card) => setCardEditor({ deckId: detailDeck.id, card })}
            onDeleteCard={(cardId) => deleteCard(detailDeck.id, cardId)}
          />
        )}
        {view === "setup" && setup && (
          <SetupView
            decks={decks}
            groups={groups}
            summary={deckSummary}
            setup={setup}
            countForMode={countForMode}
            scopeToDeckIds={scopeToDeckIds}
            onChange={(patch) => setSetup((s) => ({ ...s, ...patch }))}
            onStart={startSession}
            onCancel={() => { setView("home"); setSetup(null); }}
          />
        )}
        {view === "import" && (
          <ImportView decks={decks} onImport={importCards} onCancel={() => setView("home")} onLoadEnglish={importEnglishDecks} loadingEnglish={loadingEnglish} />
        )}
        {view === "stats" && (
          <StatsView
            stats={stats}
            decks={decks}
            cardsByDeck={cardsByDeck}
            totalDue={totalDue}
            onExport={exportAll}
            onReset={resetAll}
            onChangeNewPerDay={async (n) => {
              const next = { ...stats, settings: { ...stats.settings, newPerDay: n } };
              setStats(next);
              await persistStats(next);
            }}
          />
        )}
        {view === "study" && session && (
          <StudyView
            session={session}
            card={currentCard}
            deck={currentDeck}
            finished={sessionFinished}
            previews={gradePreviews}
            onFlip={() => setSession((s) => ({ ...s, flipped: true }))}
            onAnswer={answer}
            onExit={() => { setView("home"); setSession(null); }}
          />
        )}
      </main>
        </>
        )}
      </div>

      <MobileNav
        section={section}
        onSection={changeSection}
        studyingDue={totalDue}
        calmName={calmName}
        fastingName={fastingName}
        mgmtName={mgmtName}
        toolkitName={toolkitName}
        budgetName={budgetName}
        inventoryName={inventoryName}
        financeName={financeName}
        cloud={cloudState}
        onSyncNow={doSyncNow}
      />

      {deckEditor && (
        <DeckEditor
          deck={deckEditor.deck}
          groups={groups}
          topics={[...new Set(decks.map((d) => d.topic).filter(Boolean))]}
          onClose={() => setDeckEditor(null)}
          onSave={async (meta) => {
            if (deckEditor.deck) await updateDeck(deckEditor.deck.id, meta);
            else await createDeck(meta);
            setDeckEditor(null);
          }}
        />
      )}

      {groupEditor && (
        <GroupEditor
          group={groupEditor.group}
          onClose={() => setGroupEditor(null)}
          onSave={async (meta) => {
            if (groupEditor.group) await updateGroup(groupEditor.group.id, meta);
            else await createGroup(meta);
            setGroupEditor(null);
          }}
        />
      )}

      {cardEditor && (
        <CardEditor
          deck={decks.find((d) => d.id === cardEditor.deckId)}
          card={cardEditor.card}
          onClose={() => setCardEditor(null)}
          onSave={async (fields, images) => {
            await saveCard(cardEditor.deckId, cardEditor.card?.id || null, fields, images);
            setCardEditor(null);
          }}
        />
      )}

      {toast && (
        <div className="fixed bottom-6 left-1/2 z-40 -translate-x-1/2 rounded-full bg-slate-900 px-4 py-2 text-sm text-white shadow-lg">
          {toast}
        </div>
      )}
    </div>
  );
}

/* ------------------------------------------------------------------ */
/* Nav button                                                          */
/* ------------------------------------------------------------------ */
/* Mobile bottom tab bar — shown below lg, replaces the left rail on phones */
function MobileNav({ section, onSection, studyingDue, calmName, fastingName, mgmtName, toolkitName, budgetName, inventoryName, financeName, cloud, onSyncNow }) {
  const items = [
    { id: "review", label: "Огляд", icon: Sunrise, badge: 0 },
    { id: "studying", label: "Навчання", icon: GraduationCap, badge: studyingDue },
    { id: "routine", label: "Рутина", icon: Sun, badge: 0 },
    { id: "calm", label: calmName || "Спокій", icon: Leaf, badge: 0 },
    { id: "fasting", label: fastingName || "Fasting", icon: Hourglass, badge: 0 },
    { id: "management", label: mgmtName || "Менеджмент", icon: Briefcase, badge: 0 },
    { id: "money", label: "Гроші", icon: Wallet, badge: 0 },
    { id: "inventory", label: inventoryName || "Inventory", icon: Home, badge: 0 },
  ];
  return (
    <nav className="fixed inset-x-0 bottom-0 z-40 border-t border-slate-200 bg-white/95 backdrop-blur lg:hidden" style={{ paddingBottom: "env(safe-area-inset-bottom)" }}>
      {cloud && (
        <button onClick={onSyncNow} disabled={cloud.syncing}
          className={`flex w-full items-center justify-center gap-1.5 border-b border-slate-100 py-1 text-[11px] font-medium ${cloud.signedIn ? "text-green-600" : "text-slate-400"}`}>
          {cloud.syncing ? <RefreshCw className="h-3 w-3 animate-spin" /> : cloud.signedIn ? <Cloud className="h-3 w-3" /> : <CloudOff className="h-3 w-3" />}
          <span>{cloud.syncing ? "Синхронізація…" : cloud.signedIn ? "Синхронізовано" : "Офлайн"}</span>
        </button>
      )}
      <div className="grid grid-cols-8">
        {items.map((it) => {
          const active = section === it.id;
          return (
            <button key={it.id} onClick={() => onSection(it.id)} title={it.label}
              className={`relative flex flex-col items-center gap-0.5 py-1.5 text-[10px] font-medium transition ${active ? "text-rose-600" : "text-slate-400"}`}>
              <span className="relative">
                <it.icon className="h-5 w-5" />
                {it.badge > 0 && <span className="absolute -right-1.5 -top-1 min-w-[14px] rounded-full bg-rose-500 px-1 text-center text-[9px] font-bold leading-[14px] text-white">{it.badge > 99 ? "99+" : it.badge}</span>}
              </span>
              <span className="max-w-full truncate px-0.5">{it.label}</span>
            </button>
          );
        })}
      </div>
    </nav>
  );
}

function NavButton({ active, onClick, icon: Icon, children }) {
  return (
    <button
      onClick={onClick}
      className={`flex items-center gap-1.5 rounded-lg px-3 py-1.5 text-sm font-medium transition ${
        active ? "bg-rose-50 text-rose-700" : "text-slate-500 hover:bg-slate-100 hover:text-slate-800"
      }`}
    >
      <Icon className="h-4 w-4" />
      <span className="hidden sm:inline">{children}</span>
    </button>
  );
}

/* ------------------------------------------------------------------ */
/* Sidebar — top-level section navigation                              */
/* ------------------------------------------------------------------ */
function Sidebar({ section, collapsed, onSection, onToggle, studyingDue, calmName, fastingName, mgmtName, toolkitName, budgetName, inventoryName, financeName, cloud, onSyncNow }) {
  const items = [
    { id: "review", label: "Огляд", icon: Sunrise, badge: 0 },
    { id: "studying", label: "Studying", icon: GraduationCap, badge: studyingDue },
    { id: "routine", label: "My Routine", icon: Sun, badge: 0 },
    { id: "calm", label: calmName || "Спокій", icon: Leaf, badge: 0 },
    { id: "fasting", label: fastingName || "Fasting", icon: Hourglass, badge: 0 },
    { id: "management", label: mgmtName || "Менеджмент", icon: Briefcase, badge: 0 },
    { id: "money", label: "Гроші", icon: Wallet, badge: 0 },
    { id: "inventory", label: inventoryName || "Inventory", icon: Home, badge: 0 },
  ];
  const wide = !collapsed;
  return (
    <aside className={`sticky top-0 hidden h-screen shrink-0 flex-col border-r border-slate-200 bg-white transition-all lg:flex ${collapsed ? "lg:w-16" : "lg:w-60"}`}>
      <div className="flex h-14 items-center gap-2 px-4">
        <span className="grid h-8 w-8 shrink-0 place-items-center rounded-lg bg-rose-600 text-white">
          <Brain className="h-4 w-4" />
        </span>
        {wide && <span className="hidden text-lg font-bold tracking-tight text-slate-900 lg:inline">Recall</span>}
      </div>

      <nav className="flex flex-1 flex-col gap-1 p-2">
        {items.map((it) => {
          const active = section === it.id;
          return (
            <button
              key={it.id}
              onClick={() => onSection(it.id)}
              title={it.label}
              className={`flex items-center gap-3 rounded-lg px-2.5 py-2.5 text-sm font-medium transition ${
                active ? "bg-rose-50 text-rose-700" : "text-slate-500 hover:bg-slate-100 hover:text-slate-800"
              }`}
            >
              <span className="relative shrink-0">
                <it.icon className="h-5 w-5" />
                {it.badge > 0 && collapsed && <span className="absolute -right-1 -top-1 h-2 w-2 rounded-full bg-rose-500" />}
              </span>
              {wide && <span className="hidden flex-1 text-left lg:inline">{it.label}</span>}
              {wide && it.badge > 0 && <span className="hidden lg:inline"><CountPill n={it.badge} cls="bg-rose-100 text-rose-700" /></span>}
            </button>
          );
        })}
      </nav>

      {/* cloud sync status */}
      {cloud && (
        <button
          onClick={onSyncNow}
          disabled={cloud.syncing}
          title={cloud.signedIn ? `Синхронізовано: ${cloud.email} — натисни, щоб синхронізувати зараз` : "Офлайн — увімкни хмарну синхронізацію"}
          className={`mx-2 flex items-center gap-2 rounded-lg px-2.5 py-2 text-xs font-medium transition ${cloud.signedIn ? "text-green-600 hover:bg-green-50" : "text-slate-400 hover:bg-slate-100"}`}
        >
          <span className="relative shrink-0">
            {cloud.syncing ? <RefreshCw className="h-4 w-4 animate-spin" /> : cloud.signedIn ? <Cloud className="h-4 w-4" /> : <CloudOff className="h-4 w-4" />}
            {collapsed && cloud.signedIn && <span className="absolute -right-0.5 -top-0.5 h-2 w-2 rounded-full bg-green-500 ring-2 ring-white" />}
          </span>
          {wide && (
            <span className="hidden flex-1 items-center justify-between gap-1 lg:flex">
              <span className="truncate">{cloud.syncing ? "Синхронізація…" : cloud.signedIn ? "Синхронізовано" : "Офлайн"}</span>
              {cloud.signedIn && !cloud.syncing && <RefreshCw className="h-3.5 w-3.5 shrink-0 opacity-60" />}
            </span>
          )}
        </button>
      )}

      <button
        onClick={onToggle}
        className="m-2 mt-0 hidden items-center gap-2 rounded-lg px-2.5 py-2 text-xs font-medium text-slate-400 transition hover:bg-slate-100 hover:text-slate-600 lg:flex"
        title={collapsed ? "Expand" : "Collapse"}
      >
        {collapsed ? <PanelLeft className="h-4 w-4" /> : <PanelLeftClose className="h-4 w-4" />}
        {wide && <span className="hidden lg:inline">Collapse</span>}
      </button>
    </aside>
  );
}

/* ------------------------------------------------------------------ */
/* Home view                                                           */
/* ------------------------------------------------------------------ */
function HomeView({
  decks, groups, summary, groupSummary, totalDue, stats,
  onStudy, onStudyAll, onStudyGroup, onOpenDeck, onDelete, onEdit, onMoveDeck,
  onNewDeck, onNewGroup, onEditGroup, onDeleteGroup, onToggleGroup, onImport, onSample,
  onLoadEnglish, loadingEnglish,
}) {
  const streak = computeStreak(stats.history);
  const studiedToday = stats.history?.[dateKey(Date.now())]?.studied || 0;
  const [dragId, setDragId] = useState(null);
  const [overTarget, setOverTarget] = useState(null); // group id | "ungrouped"

  // learned (mature) vs remaining across all decks
  const mastery = useMemo(() => {
    let total = 0, learned = 0, fresh = 0;
    for (const d of decks) { const s = summary[d.id]; if (!s) continue; total += s.total || 0; learned += s.learned || 0; fresh += s.newTotal || 0; }
    return { total, learned, fresh, young: Math.max(0, total - learned - fresh), remaining: total - learned };
  }, [decks, summary]);

  const dropProps = (target) => ({
    onDragOver: (e) => { if (dragId) { e.preventDefault(); setOverTarget(target); } },
    onDragLeave: () => setOverTarget((t) => (t === target ? null : t)),
    onDrop: (e) => {
      e.preventDefault();
      if (dragId) onMoveDeck(dragId, target === "ungrouped" ? "" : target);
      setDragId(null); setOverTarget(null);
    },
  });

  if (!decks.length && !groups.length) {
    return (
      <div className="mx-auto max-w-md py-16 text-center">
        <div className="mx-auto grid h-16 w-16 place-items-center rounded-2xl bg-rose-600 text-white">
          <Brain className="h-8 w-8" />
        </div>
        <h1 className="mt-5 text-2xl font-bold text-slate-900">Study anything, remembered.</h1>
        <p className="mt-2 text-slate-500">
          Build a deck, import a spreadsheet, or paste your cards. Recall schedules each one with
          spaced repetition so reviews land right before you'd forget.
        </p>
        <div className="mt-6 flex flex-col items-center gap-3">
          {onLoadEnglish && (
            <button onClick={onLoadEnglish} disabled={loadingEnglish} className="inline-flex items-center gap-2 rounded-xl bg-blue-600 px-6 py-3 text-base font-bold text-white shadow-md shadow-blue-500/20 transition hover:bg-blue-700 disabled:opacity-60">
              {loadingEnglish ? <><RefreshCw className="h-5 w-5 animate-spin" /> Завантажую…</> : <>🇬🇧 Мої англійські колоди</>}
            </button>
          )}
          <div className="flex gap-3">
            <button onClick={onNewDeck} className="inline-flex items-center gap-2 rounded-lg bg-rose-600 px-5 py-2.5 font-semibold text-white shadow-sm transition hover:bg-rose-700">
              <Plus className="h-4 w-4" /> New deck
            </button>
            <button onClick={onImport} className="inline-flex items-center gap-2 rounded-lg border border-slate-300 bg-white px-5 py-2.5 font-semibold text-slate-700 shadow-sm transition hover:bg-slate-50">
              <Upload className="h-4 w-4" /> Import
            </button>
          </div>
          <button onClick={onSample} className="inline-flex items-center gap-2 text-sm font-medium text-slate-500 hover:text-rose-600">
            <Sparkles className="h-4 w-4" /> or load a sample deck
          </button>
        </div>
      </div>
    );
  }

  const ungrouped = decks.filter((d) => !d.groupId || !groups.some((g) => g.id === d.groupId));
  const deckCard = (d) => (
    <DeckCard
      key={d.id} deck={d} s={summary[d.id]} groups={groups}
      onOpen={onOpenDeck} onStudy={onStudy} onEdit={onEdit} onDelete={onDelete} onMove={onMoveDeck}
      dragging={dragId === d.id}
      onDragStart={() => setDragId(d.id)} onDragEnd={() => { setDragId(null); setOverTarget(null); }}
    />
  );

  return (
    <div className="space-y-6">
      {/* summary strip */}
      <div className="flex flex-wrap gap-3">
        <StatTile icon={Target} label="Due today" value={totalDue} tint={totalDue ? "text-rose-600" : "text-slate-400"} sub={totalDue ? "cards waiting" : "all caught up"} />
        <StatTile icon={Check} label="Studied today" value={studiedToday} tint="text-slate-700" sub="reviews done" />
        <StatTile icon={Flame} label="Streak" value={streak} tint={streak ? "text-orange-500" : "text-slate-400"} sub={streak === 1 ? "day" : "days"} />
        <StatTile icon={Layers} label="Decks" value={decks.length} tint="text-slate-700" />
      </div>

      {/* learned vs remaining */}
      {mastery.total > 0 && (
        <div className="rounded-2xl bg-white p-4 shadow-sm ring-1 ring-rose-50">
          <div className="flex items-center justify-between text-sm">
            <span className="font-semibold text-slate-700">Вивчено {mastery.learned.toLocaleString()} з {mastery.total.toLocaleString()} карток</span>
            <span className="font-bold text-green-600 tabular-nums">{Math.round((mastery.learned / mastery.total) * 100)}%</span>
          </div>
          <div className="mt-2 flex h-3 overflow-hidden rounded-full bg-slate-100">
            <div className="h-full bg-green-500 transition-all" style={{ width: `${(mastery.learned / mastery.total) * 100}%` }} />
            <div className="h-full bg-amber-400 transition-all" style={{ width: `${(mastery.young / mastery.total) * 100}%` }} />
            <div className="h-full bg-slate-300 transition-all" style={{ width: `${(mastery.fresh / mastery.total) * 100}%` }} />
          </div>
          <div className="mt-2 flex flex-wrap items-center gap-x-3 gap-y-1 text-[11px] font-medium text-slate-500">
            <span className="inline-flex items-center gap-1"><span className="h-2 w-2 rounded-full bg-green-500" /> вивчено {mastery.learned.toLocaleString()}</span>
            <span className="inline-flex items-center gap-1"><span className="h-2 w-2 rounded-full bg-amber-400" /> вчу {mastery.young.toLocaleString()}</span>
            <span className="inline-flex items-center gap-1"><span className="h-2 w-2 rounded-full bg-slate-300" /> ще не починала {mastery.fresh.toLocaleString()}</span>
            <span className="ml-auto font-semibold text-slate-600">залишилось {mastery.remaining.toLocaleString()}</span>
          </div>
        </div>
      )}

      {/* actions */}
      <div className="flex flex-wrap items-center gap-2">
        <button onClick={onStudyAll} className="inline-flex items-center gap-2 rounded-lg bg-rose-600 px-4 py-2 text-sm font-semibold text-white shadow-sm transition hover:bg-rose-700">
          <Play className="h-4 w-4" /> Study
        </button>
        <button onClick={onNewDeck} className="inline-flex items-center gap-2 rounded-lg border border-slate-300 bg-white px-4 py-2 text-sm font-medium text-slate-700 transition hover:bg-slate-50">
          <Plus className="h-4 w-4" /> New deck
        </button>
        <button onClick={onNewGroup} className="inline-flex items-center gap-2 rounded-lg border border-slate-300 bg-white px-4 py-2 text-sm font-medium text-slate-700 transition hover:bg-slate-50">
          <FolderPlus className="h-4 w-4" /> New group
        </button>
        <button onClick={onImport} className="inline-flex items-center gap-2 rounded-lg border border-slate-300 bg-white px-4 py-2 text-sm font-medium text-slate-700 transition hover:bg-slate-50">
          <Upload className="h-4 w-4" /> Import
        </button>
      </div>

      {groups.length > 0 && (
        <p className="flex items-center gap-1.5 text-xs text-slate-400">
          <GripVertical className="h-3.5 w-3.5" /> Tip: drag a deck onto a group to move it — or use the folder button on a deck.
        </p>
      )}

      {/* group folders */}
      <div className="space-y-3">
        {groups.map((g) => (
          <GroupFolder
            key={g.id}
            group={g}
            rollup={groupSummary[g.id]}
            decks={decks.filter((d) => d.groupId === g.id)}
            renderDeck={deckCard}
            onStudyGroup={onStudyGroup}
            onEditGroup={onEditGroup}
            onDeleteGroup={onDeleteGroup}
            onToggle={onToggleGroup}
            dropProps={dropProps(g.id)}
            highlight={overTarget === g.id}
          />
        ))}
      </div>

      {/* ungrouped decks */}
      {(ungrouped.length > 0 || (dragId && groups.length > 0)) && (
        <div {...dropProps("ungrouped")} className={`rounded-xl transition ${overTarget === "ungrouped" ? "ring-2 ring-rose-300 ring-offset-2" : ""}`}>
          {groups.length > 0 && (
            <h3 className="mb-2 flex items-center gap-1.5 text-xs font-semibold uppercase tracking-wide text-slate-400">
              <Layers className="h-3.5 w-3.5" /> Ungrouped <span className="text-slate-300">· {ungrouped.length}</span>
            </h3>
          )}
          {ungrouped.length > 0 ? (
            <div className="grid gap-3 sm:grid-cols-2">{ungrouped.map(deckCard)}</div>
          ) : (
            <p className="rounded-xl border border-dashed border-slate-300 py-4 text-center text-sm text-slate-400">Drop here to remove from a group</p>
          )}
        </div>
      )}

      <div className="flex items-center gap-4 pt-2 text-xs text-slate-400">
        <span className="inline-flex items-center gap-1"><span className="h-2 w-2 rounded-full bg-blue-400" /> new</span>
        <span className="inline-flex items-center gap-1"><span className="h-2 w-2 rounded-full bg-red-400" /> learning</span>
        <span className="inline-flex items-center gap-1"><span className="h-2 w-2 rounded-full bg-green-400" /> review</span>
      </div>
    </div>
  );
}

function GroupFolder({ group, rollup, decks, renderDeck, onStudyGroup, onEditGroup, onDeleteGroup, onToggle, dropProps = {}, highlight = false }) {
  const color = getColor(group.color);
  const r = rollup || { due: 0, total: 0, deckCount: decks.length };
  const open = !group.collapsed;
  return (
    <div {...dropProps} className={`overflow-hidden rounded-xl border bg-white shadow-sm transition ${highlight ? "border-rose-400 ring-2 ring-rose-200" : "border-slate-200"}`}>
      <div className="group flex items-center gap-2 px-3 py-2.5">
        <button onClick={() => onToggle(group.id)} className="flex min-w-0 flex-1 items-center gap-2.5 text-left">
          <ChevronRight className={`h-4 w-4 shrink-0 text-slate-400 transition-transform ${open ? "rotate-90" : ""}`} />
          <span className={`grid h-8 w-8 shrink-0 place-items-center rounded-lg text-base ${color.bg} ${color.text}`}>
            {group.emoji ? <span>{group.emoji}</span> : (open ? <FolderOpen className="h-4 w-4" /> : <Folder className="h-4 w-4" />)}
          </span>
          <span className="min-w-0">
            <span className="block truncate font-semibold text-slate-800">{group.name}</span>
            <span className="block text-xs text-slate-400">{r.deckCount} deck{r.deckCount === 1 ? "" : "s"} · {r.total} cards</span>
          </span>
        </button>
        {r.due > 0 && (
          <button onClick={() => onStudyGroup(group.id)} className="inline-flex shrink-0 items-center gap-1 rounded-lg bg-rose-600 px-3 py-1.5 text-sm font-semibold text-white transition hover:bg-rose-700">
            <Play className="h-3.5 w-3.5" /> {r.due}
          </button>
        )}
        <button onClick={() => onEditGroup(group)} className="rounded-md p-1.5 text-slate-300 transition hover:bg-slate-100 hover:text-slate-600" title="Edit group">
          <Pencil className="h-4 w-4" />
        </button>
        <button
          onClick={() => { if (confirm(`Delete group “${group.name}”? Its decks are kept and moved to Ungrouped.`)) onDeleteGroup(group.id); }}
          className="rounded-md p-1.5 text-slate-300 transition hover:bg-red-50 hover:text-red-500" title="Delete group"
        >
          <Trash2 className="h-4 w-4" />
        </button>
      </div>
      {open && (
        <div className="border-t border-slate-100 bg-slate-50/60 p-3">
          {decks.length ? (
            <div className="grid gap-3 sm:grid-cols-2">{decks.map(renderDeck)}</div>
          ) : (
            <p className="px-1 py-4 text-center text-sm text-slate-400">Empty — drag a deck here, or use a deck's folder button to move it in.</p>
          )}
        </div>
      )}
    </div>
  );
}

function DeckCard({ deck, s, groups = [], onOpen, onStudy, onEdit, onDelete, onMove, dragging = false, onDragStart, onDragEnd }) {
  const sum = s || { due: 0, newDue: 0, learn: 0, review: 0, total: 0 };
  const color = getColor(deck.color);
  const isDeadline = deck.goal === "deadline" && deck.deadline;
  const dleft = isDeadline ? daysUntil(deck.deadline) : null;
  const [menu, setMenu] = useState(false);
  return (
    <div
      draggable
      onDragStart={(e) => { e.dataTransfer.effectAllowed = "move"; e.dataTransfer.setData("text/plain", deck.id); onDragStart?.(); }}
      onDragEnd={() => onDragEnd?.()}
      className={`group relative flex items-center gap-3 rounded-xl border border-slate-200 bg-white p-4 shadow-sm transition hover:border-rose-200 hover:shadow ${dragging ? "opacity-40" : ""}`}
    >
      <button onClick={() => onOpen(deck.id)} className="flex min-w-0 flex-1 items-center gap-3 text-left">
        <div className={`grid h-11 w-11 shrink-0 place-items-center rounded-lg text-lg ${color.bg} ${color.text}`}>
          {deck.emoji ? <span>{deck.emoji}</span> : <GraduationCap className="h-5 w-5" />}
        </div>
        <div className="min-w-0 flex-1">
          <div className="flex items-center gap-1.5">
            <span className="truncate font-semibold text-slate-800">{deck.name}</span>
            {isDeadline && (
              <span className="inline-flex shrink-0 items-center gap-0.5 rounded-full bg-orange-100 px-1.5 py-0.5 text-[10px] font-semibold text-orange-700">
                <CalendarClock className="h-3 w-3" />{dleft <= 0 ? "due" : `${dleft}d`}
              </span>
            )}
          </div>
          <div className="mt-1 flex flex-wrap items-center gap-1.5 text-xs">
            <CountPill n={sum.newDue} cls="bg-blue-100 text-blue-700" />
            <CountPill n={sum.learn} cls="bg-red-100 text-red-700" />
            <CountPill n={sum.review} cls="bg-green-100 text-green-700" />
            <span className="text-slate-400">{sum.total} card{sum.total === 1 ? "" : "s"}</span>
          </div>
          {sum.total > 0 && (
            <div className="mt-1.5 flex items-center gap-2">
              <div className="h-1.5 flex-1 overflow-hidden rounded-full bg-slate-100"><div className="h-full rounded-full bg-green-500" style={{ width: `${((sum.learned || 0) / sum.total) * 100}%` }} /></div>
              <span className="shrink-0 text-[10px] font-medium tabular-nums text-slate-400">вивчено {sum.learned || 0}/{sum.total}</span>
            </div>
          )}
        </div>
      </button>
      <div className="flex shrink-0 items-center gap-1">
        {sum.due > 0 ? (
          <button onClick={() => onStudy(deck.id)} className="inline-flex items-center gap-1 rounded-lg bg-rose-600 px-3 py-1.5 text-sm font-semibold text-white transition hover:bg-rose-700" title="Study">
            <Play className="h-3.5 w-3.5" /> {sum.due}
          </button>
        ) : (
          <button onClick={() => onStudy(deck.id)} className="rounded-lg bg-slate-100 px-3 py-1.5 text-xs font-medium text-slate-500 transition hover:bg-slate-200" title="Study">study</button>
        )}
        {onMove && groups.length > 0 && (
          <button onClick={() => setMenu((v) => !v)} className="rounded-md p-1.5 text-slate-300 transition hover:bg-slate-100 hover:text-slate-600 sm:opacity-0 sm:group-hover:opacity-100" title="Move to group">
            <Folder className="h-4 w-4" />
          </button>
        )}
        <button onClick={() => onEdit(deck)} className="rounded-md p-1.5 text-slate-300 transition hover:bg-slate-100 hover:text-slate-600 sm:opacity-0 sm:group-hover:opacity-100" title="Edit deck">
          <Pencil className="h-4 w-4" />
        </button>
        <button
          onClick={() => { if (confirm(`Delete deck “${deck.name}” and its ${sum.total} cards?`)) onDelete(deck.id); }}
          className="rounded-md p-1.5 text-slate-300 transition hover:bg-red-50 hover:text-red-500 sm:opacity-0 sm:group-hover:opacity-100" title="Delete deck"
        >
          <Trash2 className="h-4 w-4" />
        </button>
      </div>

      {menu && (
        <>
          <div className="fixed inset-0 z-10" onClick={() => setMenu(false)} />
          <div className="absolute right-2 top-14 z-20 w-52 overflow-hidden rounded-xl border border-slate-200 bg-white py-1 shadow-lg">
            <div className="px-3 py-1.5 text-[11px] font-semibold uppercase tracking-wide text-slate-400">Move to group</div>
            {deck.groupId && (
              <button onClick={() => { onMove(deck.id, ""); setMenu(false); }} className="flex w-full items-center gap-2 px-3 py-2 text-left text-sm text-slate-600 hover:bg-slate-50">
                <X className="h-4 w-4 text-slate-400" /> Remove from group
              </button>
            )}
            {groups.map((g) => (
              <button key={g.id} onClick={() => { onMove(deck.id, g.id); setMenu(false); }}
                className={`flex w-full items-center gap-2 px-3 py-2 text-left text-sm hover:bg-slate-50 ${deck.groupId === g.id ? "font-semibold text-rose-600" : "text-slate-700"}`}>
                <span className="text-base">{g.emoji || "📁"}</span> <span className="truncate">{g.name}</span>
                {deck.groupId === g.id && <Check className="ml-auto h-4 w-4" />}
              </button>
            ))}
          </div>
        </>
      )}
    </div>
  );
}

/* ------------------------------------------------------------------ */
/* Interval / scheduling visual pieces                                 */
/* ------------------------------------------------------------------ */
function StageBadge({ card, showDue = true }) {
  const stage = stageForCard(card);
  return (
    <span className={`inline-flex items-center gap-1 rounded-full px-2 py-0.5 text-[11px] font-semibold ${stage.bg} ${stage.text}`}>
      <span className="h-1.5 w-1.5 rounded-full" style={{ backgroundColor: stage.dot }} />
      {intervalLabel(card)}
      {showDue && card.state !== "new" && <span className="font-normal opacity-70">· {dueLabel(card)}</span>}
    </span>
  );
}

// The Learning → 1d → 3d → 1w → 2w → 1mo → 3mo+ scale, active stage highlighted.
function IntervalTimeline({ activeStageId, className = "" }) {
  return (
    <div className={`flex items-center gap-1 overflow-x-auto ${className}`}>
      {INTERVAL_STAGES.map((s, i) => {
        const active = s.id === activeStageId;
        return (
          <div key={s.id} className="flex items-center gap-1">
            {i > 0 && <span className="h-px w-2 bg-slate-200" />}
            <span
              className={`whitespace-nowrap rounded-full px-2 py-0.5 text-[10px] font-semibold transition ${active ? "text-white" : "text-slate-400"}`}
              style={active ? { backgroundColor: s.dot } : { backgroundColor: "#f1f5f9" }}
            >
              {s.id === "learning" ? "Learn" : s.id === "long" ? "3mo+" : s.label.replace(" day", "d").replace(" days", "d").replace(" week", "w").replace(" weeks", "w").replace(" month", "mo").replace(" months", "mo")}
            </span>
          </div>
        );
      })}
    </div>
  );
}

// Horizontal stacked breakdown of how many cards sit at each interval stage.
function StageBreakdown({ stages, total }) {
  const order = [{ id: "new", label: "New", dot: "#94a3b8" }, ...INTERVAL_STAGES];
  const present = order.filter((s) => (stages[s.id] || 0) > 0);
  if (!total) return null;
  return (
    <div>
      <div className="flex h-3 w-full overflow-hidden rounded-full bg-slate-100">
        {present.map((s) => (
          <div key={s.id} style={{ width: `${((stages[s.id] || 0) / total) * 100}%`, backgroundColor: s.dot }} title={`${s.label}: ${stages[s.id]}`} />
        ))}
      </div>
      <div className="mt-2 flex flex-wrap gap-x-3 gap-y-1">
        {present.map((s) => (
          <span key={s.id} className="inline-flex items-center gap-1 text-[11px] text-slate-500">
            <span className="h-2 w-2 rounded-full" style={{ backgroundColor: s.dot }} />
            {s.id === "learning" ? "Learning" : s.label} <span className="font-semibold tabular-nums text-slate-700">{stages[s.id]}</span>
          </span>
        ))}
      </div>
    </div>
  );
}

function SpeakerButton({ text, lang, size = "sm", onFallback }) {
  if (!text || !ttsSupported()) return null;
  const dim = size === "lg" ? "h-9 w-9" : "h-7 w-7";
  return (
    <button
      onClick={(e) => { e.stopPropagation(); const r = speak(text, lang); if (!r.ok && onFallback) onFallback(); }}
      className={`grid ${dim} shrink-0 place-items-center rounded-full bg-slate-100 text-slate-500 transition hover:bg-rose-100 hover:text-rose-600`}
      title="Read aloud (R)"
    >
      <Volume2 className={size === "lg" ? "h-4.5 w-4.5" : "h-4 w-4"} style={{ width: size === "lg" ? 18 : 15, height: size === "lg" ? 18 : 15 }} />
    </button>
  );
}

/* ------------------------------------------------------------------ */
/* Study view                                                          */
/* ------------------------------------------------------------------ */
function StudyView({ session, card, deck, finished, previews, onFlip, onAnswer, onExit }) {
  const progress = session.total ? Math.round(((session.total - session.queue.length) / session.total) * 100) : 100;
  const lang = card?.lang || deck?.language || "";
  const [imgs, setImgs] = useState({ front: null, back: null });
  const [lightbox, setLightbox] = useState(null);
  const [voiceHint, setVoiceHint] = useState(false);
  const [drag, setDrag] = useState(0); // touch-swipe horizontal offset
  const dragRef = useRef(null);
  const dragging = !!dragRef.current;

  const say = useCallback((text) => {
    if (!text) return;
    const r = speak(text, lang);
    setVoiceHint(!r.ok);
  }, [lang]);

  // Lazily load this card's images from their own storage keys.
  useEffect(() => {
    let alive = true;
    setImgs({ front: null, back: null });
    setLightbox(null);
    if (!card) return;
    (async () => {
      const f = card.imgFront ? await loadImage(card.id, "front") : null;
      const b = card.imgBack ? await loadImage(card.id, "back") : null;
      if (alive) setImgs({ front: f, back: b });
    })();
    return () => { alive = false; };
  }, [card?.id]);

  // Auto-play: front when the card appears, back when it flips.
  useEffect(() => {
    if (card && deck?.autoPlay && !session.flipped) say(session.reverse ? card.back : card.front);
  }, [card?.id]); // eslint-disable-line react-hooks/exhaustive-deps
  useEffect(() => {
    if (card && deck?.autoPlay && session.flipped) say(session.reverse ? card.front : card.back);
  }, [session.flipped, card?.id]); // eslint-disable-line react-hooks/exhaustive-deps

  // "R" / "P" replays the currently visible side.
  useEffect(() => {
    const onKey = (e) => {
      if (e.repeat) return;
      if (["r", "R", "p", "P"].includes(e.key)) {
        e.preventDefault();
        say(session.flipped ? card?.back || card?.front : card?.front);
      }
    };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, [session.flipped, card, say]);

  // reset any swipe offset when the card changes
  useEffect(() => { setDrag(0); dragRef.current = null; }, [card?.id, session.flipped]);
  const onTouchStart = (e) => { const t = e.touches[0]; dragRef.current = { x0: t.clientX, y0: t.clientY, dx: 0, dy: 0 }; };
  const onTouchMove = (e) => {
    if (!dragRef.current) return;
    const t = e.touches[0];
    const dx = t.clientX - dragRef.current.x0, dy = t.clientY - dragRef.current.y0;
    dragRef.current.dx = dx; dragRef.current.dy = dy;
    // follow the finger as soon as the gesture is mostly horizontal — before OR after flip,
    // so the card visibly moves on the first swipe (no tap needed first).
    if (Math.abs(dx) > Math.abs(dy)) setDrag(dx);
  };
  const onTouchEnd = () => {
    const st = dragRef.current; dragRef.current = null;
    if (!st) return;
    if (!session.flipped) { if (Math.abs(st.dx) > 40 || Math.abs(st.dy) > 40) onFlip(); setDrag(0); return; }
    if (st.dx > 80) onAnswer("good");
    else if (st.dx < -80) onAnswer("again");
    else setDrag(0);
  };

  if (finished || !card) {
    const acc = session.done ? Math.round((session.correct / session.done) * 100) : 0;
    return (
      <div className="mx-auto max-w-md py-16 text-center">
        <div className="mx-auto grid h-16 w-16 place-items-center rounded-2xl bg-green-100 text-green-600">
          <Check className="h-8 w-8" />
        </div>
        <h2 className="mt-5 text-2xl font-bold text-slate-900">Session complete</h2>
        <p className="mt-1 text-slate-500">Nice work on {session.title}{session.cram ? " — no due dates changed." : "."}</p>
        <div className="mt-6 flex justify-center gap-3">
          <StatTile icon={Check} label="Reviewed" value={session.done} />
          <StatTile icon={Target} label="Recalled" value={`${acc}%`} tint="text-green-600" />
        </div>
        <button onClick={onExit} className="mt-8 inline-flex items-center gap-2 rounded-lg bg-rose-600 px-5 py-2.5 font-semibold text-white transition hover:bg-rose-700">
          <ArrowLeft className="h-4 w-4" /> Back to decks
        </button>
      </div>
    );
  }

  // Reverse mode: show the back (translation) as the prompt, recall the front.
  const rev = !!session.reverse;
  const frontText = rev ? card.back : card.front;
  const backText = rev ? card.front : card.back;
  const frontImg = rev ? imgs.back : imgs.front;
  const backImg = rev ? imgs.front : imgs.back;

  return (
    <div className="mx-auto max-w-2xl">
      {/* header row */}
      <div className="mb-4 flex items-center gap-3">
        <button onClick={onExit} className="rounded-md p-1.5 text-slate-400 hover:bg-slate-100 hover:text-slate-700" title="End session">
          <X className="h-5 w-5" />
        </button>
        <div className="flex-1">
          <div className="flex items-center justify-between text-xs text-slate-500">
            <span className="flex items-center gap-2 font-medium">
              {session.title}
              <span className={`rounded-full px-2 py-0.5 text-[10px] font-semibold uppercase tracking-wide ${session.cram ? "bg-amber-100 text-amber-700" : "bg-slate-100 text-slate-500"}`}>
                {session.cram ? "Cram" : session.subtitle}
              </span>
            </span>
            <span className="tabular-nums">{session.total - session.queue.length} / {session.total}</span>
          </div>
          <div className="mt-1.5 h-2 overflow-hidden rounded-full bg-slate-200">
            <div className="h-full rounded-full bg-rose-600 transition-all duration-300" style={{ width: `${progress}%` }} />
          </div>
        </div>
      </div>

      {/* card (div, not button — it contains speaker buttons) */}
      <div
        role="button"
        tabIndex={0}
        onClick={session.flipped ? undefined : onFlip}
        onTouchStart={onTouchStart}
        onTouchMove={onTouchMove}
        onTouchEnd={onTouchEnd}
        style={{ transform: drag ? `translateX(${drag}px) rotate(${drag / 22}deg)` : undefined, transition: dragging ? "none" : "transform .25s ease", touchAction: "pan-y" }}
        className={`relative flex min-h-[300px] w-full flex-col items-center justify-center overflow-hidden rounded-2xl border border-slate-200 bg-white p-8 text-center shadow-sm ${session.flipped ? "" : "cursor-pointer"}`}
      >
        {/* swipe hint overlay (mobile) */}
        {session.flipped && drag !== 0 && (
          <div className={`pointer-events-none absolute inset-0 flex items-center justify-center ${drag > 0 ? "bg-green-500/10" : "bg-red-500/10"}`} style={{ opacity: Math.min(1, Math.abs(drag) / 85) }}>
            <span className={`rounded-full px-4 py-2 text-lg font-extrabold text-white shadow ${drag > 0 ? "bg-green-500" : "bg-red-500"}`}>{drag > 0 ? "Знаю ✓" : "Ще раз"}</span>
          </div>
        )}
        {card.tags && (
          <span className="mb-3 rounded-full bg-slate-100 px-2.5 py-0.5 text-xs font-medium text-slate-500">{card.tags}</span>
        )}
        {frontImg && (
          <img
            src={frontImg}
            alt=""
            onClick={(e) => { e.stopPropagation(); setLightbox(frontImg); }}
            className="mb-4 max-h-56 w-auto cursor-zoom-in rounded-lg border border-slate-200 object-contain"
          />
        )}
        <div className="flex items-center gap-2">
          <div className="text-2xl font-semibold leading-snug text-slate-900" style={{ textWrap: "balance" }}>
            {frontText}
          </div>
          <SpeakerButton text={frontText} lang={lang} onFallback={() => setVoiceHint(true)} />
        </div>

        {session.flipped ? (
          <>
            <div className="my-6 h-px w-24 bg-slate-200" />
            {backImg && (
              <img
                src={backImg}
                alt=""
                onClick={(e) => { e.stopPropagation(); setLightbox(backImg); }}
                className="mb-4 max-h-56 w-auto cursor-zoom-in rounded-lg border border-slate-200 object-contain"
              />
            )}
            <div className="flex items-center gap-2">
              <div className="text-2xl font-medium text-rose-700" style={{ textWrap: "balance" }}>{backText}</div>
              <SpeakerButton text={backText} lang={lang} onFallback={() => setVoiceHint(true)} />
            </div>
            {card.notes && <div className="mt-4 max-w-md text-sm text-slate-500">{card.notes}</div>}
          </>
        ) : (
          <div className="mt-8 inline-flex items-center gap-2 text-sm text-slate-400">
            <Keyboard className="h-4 w-4" /> tap or press <kbd className="rounded border border-slate-300 bg-slate-50 px-1.5 py-0.5 text-xs">Space</kbd> to flip
          </div>
        )}
      </div>

      {/* current interval + fallback voice hint */}
      <div className="mt-3 flex items-center justify-between gap-2">
        {session.flipped && !session.cram ? <StageBadge card={card} /> : <span />}
        {voiceHint && ttsSupported() && (
          <span className="text-[11px] text-slate-400">No {lang || "matching"} voice installed — using the default.</span>
        )}
      </div>

      {/* mobile swipe hint */}
      {session.flipped && (
        <div className="mt-2 flex items-center justify-center gap-1.5 text-[11px] text-slate-400 sm:hidden">← свайп «Ще раз» · «Знаю» свайп →</div>
      )}

      {/* rating buttons */}
      {session.flipped && (
        <div className="mt-2 grid grid-cols-4 gap-2">
          {GRADES.map((g) => (
            <button
              key={g.key}
              onClick={() => onAnswer(g.key)}
              className={`flex flex-col items-center gap-1 rounded-xl px-2 py-3 font-semibold text-white shadow-sm transition ${g.cls}`}
            >
              <span>{g.label}</span>
              <span className="text-[11px] font-normal text-white/80 tabular-nums">
                {session.cram ? (g.key === "again" || g.key === "hard" ? "again" : "done") : previews[g.key]}
              </span>
              <span className="mt-0.5 rounded bg-white/20 px-1.5 text-[10px]">{g.hint}</span>
            </button>
          ))}
        </div>
      )}

      <div className="mt-4 flex items-center justify-center gap-4 text-xs text-slate-400">
        <span>Space = flip</span><span>1–4 = rate</span>
        {ttsSupported() && <span>R = replay audio</span>}
      </div>

      {/* image lightbox */}
      {lightbox && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/80 p-4" onClick={() => setLightbox(null)}>
          <img src={lightbox} alt="" className="max-h-[90vh] max-w-full rounded-lg object-contain" />
          <button className="absolute right-4 top-4 rounded-full bg-white/10 p-2 text-white hover:bg-white/20" onClick={() => setLightbox(null)}>
            <X className="h-5 w-5" />
          </button>
        </div>
      )}
    </div>
  );
}

/* ------------------------------------------------------------------ */
/* Study setup — pick deck scope + mode before a session               */
/* ------------------------------------------------------------------ */
function SetupView({ decks, groups, summary, setup, countForMode, scopeToDeckIds, onChange, onStart, onCancel }) {
  const deckIds = scopeToDeckIds(setup.deckScope);
  const scopeKey = deckIds.join(",");
  const cardsIn = (ids) => ids.reduce((n, id) => n + (summary[id]?.total || 0), 0);

  const counts = useMemo(() => {
    const c = {};
    for (const m of STUDY_MODES) c[m.id] = countForMode(deckIds, m.id, setup.count);
    return c;
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [scopeKey, setup.count, countForMode]);

  const startCount = counts[setup.mode] || 0;
  const totalInScope = deckIds.reduce((n, id) => n + (summary[id]?.total || 0), 0);

  return (
    <div className="mx-auto max-w-2xl space-y-6">
      <div className="flex items-center gap-3">
        <button onClick={onCancel} className="rounded-md p-1.5 text-slate-400 hover:bg-slate-100 hover:text-slate-700">
          <ArrowLeft className="h-5 w-5" />
        </button>
        <h1 className="text-xl font-bold text-slate-900">Study setup</h1>
      </div>

      {/* deck scope */}
      <div>
        <label className="mb-1.5 block text-xs font-semibold uppercase tracking-wide text-slate-400">Study from</label>
        <div className="relative">
          <select
            value={setup.deckScope}
            onChange={(e) => onChange({ deckScope: e.target.value })}
            className="w-full appearance-none rounded-lg border border-slate-300 bg-white px-3 py-2.5 pr-9 text-sm font-medium focus:border-rose-400 focus:outline-none focus:ring-2 focus:ring-rose-100"
          >
            <option value="all">All decks — {cardsIn(decks.map((d) => d.id)).toLocaleString()} cards</option>
            {groups.length > 0 && (
              <optgroup label="Groups">
                {groups.map((g) => (
                  <option key={g.id} value={`group:${g.id}`}>
                    {g.emoji ? `${g.emoji} ` : "📁 "}{g.name} — {cardsIn(decks.filter((d) => d.groupId === g.id).map((d) => d.id)).toLocaleString()} cards
                  </option>
                ))}
              </optgroup>
            )}
            <optgroup label="Decks">
              {decks.map((d) => (
                <option key={d.id} value={d.id}>
                  {d.emoji ? `${d.emoji} ` : ""}{d.name} — {(summary[d.id]?.total || 0).toLocaleString()} cards
                </option>
              ))}
            </optgroup>
          </select>
          <ChevronDown className="pointer-events-none absolute right-3 top-1/2 h-4 w-4 -translate-y-1/2 text-slate-400" />
        </div>
      </div>

      {/* mode */}
      <div>
        <label className="mb-1.5 block text-xs font-semibold uppercase tracking-wide text-slate-400">How much do you want to study?</label>
        <div className="grid gap-2.5 sm:grid-cols-2">
          {STUDY_MODES.map((m) => {
            const Icon = m.icon;
            const active = setup.mode === m.id;
            const n = counts[m.id];
            return (
              <button
                key={m.id}
                onClick={() => onChange({ mode: m.id })}
                className={`flex items-start gap-3 rounded-xl border p-3 text-left transition ${
                  active ? "border-rose-500 bg-rose-50 ring-2 ring-rose-100" : "border-slate-200 bg-white hover:border-slate-300"
                }`}
              >
                <div className={`grid h-9 w-9 shrink-0 place-items-center rounded-lg ${active ? "bg-rose-600 text-white" : "bg-slate-100 text-slate-500"}`}>
                  <Icon className="h-4.5 w-4.5" style={{ width: 18, height: 18 }} />
                </div>
                <div className="min-w-0 flex-1">
                  <div className="flex items-center justify-between gap-2">
                    <span className="font-semibold text-slate-800">{m.label}</span>
                    <span className={`shrink-0 text-xs font-semibold tabular-nums ${n ? "text-rose-600" : "text-slate-300"}`}>{n.toLocaleString()}</span>
                  </div>
                  <p className="mt-0.5 text-xs leading-snug text-slate-500">{m.desc}</p>
                </div>
              </button>
            );
          })}
        </div>
      </div>

      {/* custom count */}
      {setup.mode === "custom" && (
        <div className="flex items-center gap-3 rounded-xl border border-slate-200 bg-white p-4">
          <label className="text-sm font-medium text-slate-700">How many cards?</label>
          <input
            type="number"
            min={1}
            max={9999}
            value={setup.count}
            onChange={(e) => onChange({ count: Math.max(1, Math.min(9999, Number(e.target.value) || 1)) })}
            className="w-24 rounded-lg border border-slate-300 px-3 py-1.5 text-right text-sm tabular-nums focus:border-rose-400 focus:outline-none focus:ring-2 focus:ring-rose-100"
          />
          <span className="text-xs text-slate-400">due cards come first</span>
        </div>
      )}

      {/* reverse mode */}
      <button
        onClick={() => onChange({ reverse: !setup.reverse })}
        className={`flex w-full items-center gap-3 rounded-xl border p-3 text-left transition ${setup.reverse ? "border-rose-500 bg-rose-50 ring-2 ring-rose-100" : "border-slate-200 bg-white hover:border-slate-300"}`}
      >
        <div className={`grid h-9 w-9 shrink-0 place-items-center rounded-lg ${setup.reverse ? "bg-rose-600 text-white" : "bg-slate-100 text-slate-500"}`}><RefreshCw className="h-4 w-4" /></div>
        <div className="min-w-0 flex-1">
          <div className="text-sm font-semibold text-slate-800">Реверс (показувати переклад)</div>
          <div className="text-xs text-slate-400">Спершу бачиш зворот картки й пригадуєш оригінал — тренує в інший бік.</div>
        </div>
        <div className={`h-6 w-11 shrink-0 rounded-full p-0.5 transition ${setup.reverse ? "bg-rose-500" : "bg-slate-200"}`}><div className={`h-5 w-5 rounded-full bg-white shadow transition ${setup.reverse ? "translate-x-5" : ""}`} /></div>
      </button>

      <div className="flex items-center gap-3 border-t border-slate-100 pt-4">
        <button
          onClick={() => onStart(setup)}
          disabled={!startCount}
          className="inline-flex items-center gap-2 rounded-lg bg-rose-600 px-6 py-2.5 font-semibold text-white shadow-sm transition hover:bg-rose-700 disabled:cursor-not-allowed disabled:bg-slate-300"
        >
          <Play className="h-4 w-4" /> Start · {startCount.toLocaleString()} card{startCount === 1 ? "" : "s"}
        </button>
        {!startCount && (
          <span className="text-sm text-slate-400">
            {totalInScope ? "No cards match this mode right now." : "This deck has no cards yet."}
          </span>
        )}
      </div>
    </div>
  );
}

/* ------------------------------------------------------------------ */
/* Deck editor modal — create or edit a deck's identity                */
/* ------------------------------------------------------------------ */
function DeckEditor({ deck, groups = [], topics, onClose, onSave }) {
  const [name, setName] = useState(deck?.name || "");
  const [topic, setTopic] = useState(deck?.topic || "");
  const [description, setDescription] = useState(deck?.description || "");
  const [emoji, setEmoji] = useState(deck?.emoji || "");
  const [color, setColor] = useState(deck?.color || "indigo");
  const [groupId, setGroupId] = useState(deck?.groupId || "");
  const [language, setLanguage] = useState(deck?.language || "en-US");
  const [autoPlay, setAutoPlay] = useState(!!deck?.autoPlay);
  const [goal, setGoal] = useState(deck?.goal || "longterm");
  const [deadline, setDeadline] = useState(
    deck?.deadline ? new Date(deck.deadline).toISOString().slice(0, 10) : ""
  );

  const topicOptions = useMemo(
    () => [...new Set([...topics, ...TOPIC_PRESETS])],
    [topics]
  );

  const save = () => {
    if (!name.trim()) return;
    let deadlineMs = null;
    if (goal === "deadline" && deadline) {
      const d = new Date(deadline + "T23:59:59");
      if (!isNaN(d.getTime())) deadlineMs = d.getTime();
    }
    onSave({
      name: name.trim(), topic: topic.trim(), description: description.trim(), emoji, color,
      groupId, language, autoPlay, goal, deadline: deadlineMs,
    });
  };

  const dleft = deadline ? daysUntil(new Date(deadline + "T23:59:59").getTime()) : null;

  return (
    <div className="fixed inset-0 z-50 flex items-end justify-center bg-slate-900/40 p-0 backdrop-blur-sm sm:items-center sm:p-4" onClick={onClose}>
      <div
        className="max-h-[92vh] w-full max-w-lg overflow-y-auto rounded-t-2xl bg-white p-5 shadow-xl sm:rounded-2xl"
        onClick={(e) => e.stopPropagation()}
      >
        <div className="mb-4 flex items-center justify-between">
          <h2 className="text-lg font-bold text-slate-900">{deck ? "Edit deck" : "New deck"}</h2>
          <button onClick={onClose} className="rounded-md p-1 text-slate-400 hover:bg-slate-100 hover:text-slate-700">
            <X className="h-5 w-5" />
          </button>
        </div>

        {/* preview */}
        <div className="mb-4 flex items-center gap-3 rounded-xl border border-slate-200 bg-slate-50 p-3">
          <div className={`grid h-11 w-11 shrink-0 place-items-center rounded-lg text-lg ${getColor(color).bg} ${getColor(color).text}`}>
            {emoji ? <span>{emoji}</span> : <GraduationCap className="h-5 w-5" />}
          </div>
          <div className="min-w-0">
            <div className="truncate font-semibold text-slate-800">{name.trim() || "Untitled deck"}</div>
            <div className="truncate text-xs text-slate-400">{topic.trim() || "No topic"}{description.trim() ? ` · ${description.trim()}` : ""}</div>
          </div>
        </div>

        <div className="space-y-4">
          <label className="block">
            <span className="mb-1 block text-xs font-medium text-slate-500">Name <span className="text-red-500">*</span></span>
            <input
              autoFocus
              value={name}
              onChange={(e) => setName(e.target.value)}
              placeholder="e.g. Spanish verbs"
              className="w-full rounded-lg border border-slate-300 bg-white px-3 py-2 text-sm focus:border-rose-400 focus:outline-none focus:ring-2 focus:ring-rose-100"
            />
          </label>

          <label className="block">
            <span className="mb-1 block text-xs font-medium text-slate-500">Topic / category</span>
            <input
              value={topic}
              onChange={(e) => setTopic(e.target.value)}
              list="topic-options"
              placeholder="Type or pick — e.g. Languages"
              className="w-full rounded-lg border border-slate-300 bg-white px-3 py-2 text-sm focus:border-rose-400 focus:outline-none focus:ring-2 focus:ring-rose-100"
            />
            <datalist id="topic-options">
              {topicOptions.map((t) => <option key={t} value={t} />)}
            </datalist>
          </label>

          <label className="block">
            <span className="mb-1 block text-xs font-medium text-slate-500">Description (optional)</span>
            <textarea
              value={description}
              onChange={(e) => setDescription(e.target.value)}
              rows={2}
              placeholder="What's in this deck?"
              className="w-full rounded-lg border border-slate-300 bg-white px-3 py-2 text-sm focus:border-rose-400 focus:outline-none focus:ring-2 focus:ring-rose-100"
            />
          </label>

          <div>
            <span className="mb-1.5 block text-xs font-medium text-slate-500">Color</span>
            <div className="flex flex-wrap gap-2">
              {DECK_COLORS.map((c) => (
                <button
                  key={c.id}
                  onClick={() => setColor(c.id)}
                  className={`h-7 w-7 rounded-full transition ${color === c.id ? "ring-2 ring-slate-900 ring-offset-2" : ""}`}
                  style={{ backgroundColor: c.dot }}
                  title={c.id}
                />
              ))}
            </div>
          </div>

          <div>
            <span className="mb-1.5 block text-xs font-medium text-slate-500">Icon (optional)</span>
            <div className="flex flex-wrap gap-1.5">
              <button
                onClick={() => setEmoji("")}
                className={`grid h-9 w-9 place-items-center rounded-lg border text-slate-400 transition ${emoji === "" ? "border-rose-500 bg-rose-50" : "border-slate-200 hover:bg-slate-50"}`}
                title="No icon"
              >
                <GraduationCap className="h-4 w-4" />
              </button>
              {DECK_EMOJIS.map((e) => (
                <button
                  key={e}
                  onClick={() => setEmoji(e)}
                  className={`grid h-9 w-9 place-items-center rounded-lg border text-lg transition ${emoji === e ? "border-rose-500 bg-rose-50" : "border-slate-200 hover:bg-slate-50"}`}
                >
                  {e}
                </button>
              ))}
            </div>
          </div>

          {/* group + audio language */}
          <div className="grid gap-4 sm:grid-cols-2">
            <label className="block">
              <span className="mb-1 block text-xs font-medium text-slate-500">Group</span>
              <div className="relative">
                <select value={groupId} onChange={(e) => setGroupId(e.target.value)} className="w-full appearance-none rounded-lg border border-slate-300 bg-white px-3 py-2 pr-9 text-sm focus:border-rose-400 focus:outline-none focus:ring-2 focus:ring-rose-100">
                  <option value="">No group (ungrouped)</option>
                  {groups.map((g) => (
                    <option key={g.id} value={g.id}>{g.emoji ? `${g.emoji} ` : ""}{g.name}</option>
                  ))}
                </select>
                <ChevronDown className="pointer-events-none absolute right-3 top-1/2 h-4 w-4 -translate-y-1/2 text-slate-400" />
              </div>
            </label>
            <label className="block">
              <span className="mb-1 block text-xs font-medium text-slate-500">Audio language (TTS)</span>
              <div className="relative">
                <select value={language} onChange={(e) => setLanguage(e.target.value)} className="w-full appearance-none rounded-lg border border-slate-300 bg-white px-3 py-2 pr-9 text-sm focus:border-rose-400 focus:outline-none focus:ring-2 focus:ring-rose-100">
                  {DECK_LANGUAGES.map((l) => <option key={l.code} value={l.code}>{l.label}</option>)}
                </select>
                <ChevronDown className="pointer-events-none absolute right-3 top-1/2 h-4 w-4 -translate-y-1/2 text-slate-400" />
              </div>
            </label>
          </div>

          <label className="flex items-center justify-between gap-4 rounded-lg border border-slate-200 bg-slate-50 px-3 py-2.5">
            <span className="flex items-center gap-2 text-sm font-medium text-slate-700">
              <Volume2 className="h-4 w-4 text-slate-400" /> Auto-play audio when a card is shown
            </span>
            <input type="checkbox" checked={autoPlay} onChange={(e) => setAutoPlay(e.target.checked)} className="h-4 w-4 accent-rose-600" />
          </label>

          {/* scheduling goal */}
          <div>
            <span className="mb-1.5 block text-xs font-medium text-slate-500">Scheduling goal</span>
            <div className="grid gap-2 sm:grid-cols-2">
              {Object.values(SCHED_GOALS).map((g) => {
                const Icon = g.icon;
                const active = goal === g.id;
                return (
                  <button
                    key={g.id}
                    onClick={() => setGoal(g.id)}
                    className={`flex items-start gap-2.5 rounded-xl border p-3 text-left transition ${active ? "border-rose-500 bg-rose-50 ring-2 ring-rose-100" : "border-slate-200 bg-white hover:border-slate-300"}`}
                  >
                    <div className={`grid h-8 w-8 shrink-0 place-items-center rounded-lg ${active ? "bg-rose-600 text-white" : "bg-slate-100 text-slate-500"}`}>
                      <Icon className="h-4 w-4" />
                    </div>
                    <div>
                      <div className="text-sm font-semibold text-slate-800">{g.short}</div>
                      <p className="mt-0.5 text-[11px] leading-snug text-slate-500">{g.desc}</p>
                    </div>
                  </button>
                );
              })}
            </div>
            {goal === "deadline" && (
              <div className="mt-3 flex flex-wrap items-center gap-3 rounded-lg border border-orange-200 bg-orange-50 px-3 py-2.5">
                <label className="text-sm font-medium text-orange-800">Target date</label>
                <input type="date" value={deadline} onChange={(e) => setDeadline(e.target.value)} className="rounded-lg border border-orange-200 bg-white px-3 py-1.5 text-sm focus:border-orange-400 focus:outline-none" />
                {dleft != null && (
                  <span className="text-xs font-semibold text-orange-700">
                    {dleft <= 0 ? "date has passed" : `${dleft} day${dleft === 1 ? "" : "s"} left`}
                  </span>
                )}
              </div>
            )}
          </div>
        </div>

        <div className="mt-6 flex items-center justify-end gap-3">
          <button onClick={onClose} className="rounded-lg px-4 py-2 text-sm font-medium text-slate-500 hover:text-slate-800">
            Cancel
          </button>
          <button
            onClick={save}
            disabled={!name.trim()}
            className="inline-flex items-center gap-2 rounded-lg bg-rose-600 px-5 py-2 font-semibold text-white shadow-sm transition hover:bg-rose-700 disabled:cursor-not-allowed disabled:bg-slate-300"
          >
            <Check className="h-4 w-4" /> {deck ? "Save changes" : "Create deck"}
          </button>
        </div>
      </div>
    </div>
  );
}

/* ------------------------------------------------------------------ */
/* Group editor modal                                                  */
/* ------------------------------------------------------------------ */
function GroupEditor({ group, onClose, onSave }) {
  const [name, setName] = useState(group?.name || "");
  const [emoji, setEmoji] = useState(group?.emoji || "");
  const [color, setColor] = useState(group?.color || "slate");
  const save = () => { if (name.trim()) onSave({ name: name.trim(), emoji, color }); };

  return (
    <div className="fixed inset-0 z-50 flex items-end justify-center bg-slate-900/40 p-0 backdrop-blur-sm sm:items-center sm:p-4" onClick={onClose}>
      <div className="max-h-[92vh] w-full max-w-md overflow-y-auto rounded-t-2xl bg-white p-5 shadow-xl sm:rounded-2xl" onClick={(e) => e.stopPropagation()}>
        <div className="mb-4 flex items-center justify-between">
          <h2 className="text-lg font-bold text-slate-900">{group ? "Edit group" : "New group"}</h2>
          <button onClick={onClose} className="rounded-md p-1 text-slate-400 hover:bg-slate-100 hover:text-slate-700"><X className="h-5 w-5" /></button>
        </div>

        <div className="mb-4 flex items-center gap-3 rounded-xl border border-slate-200 bg-slate-50 p-3">
          <div className={`grid h-11 w-11 shrink-0 place-items-center rounded-lg text-lg ${getColor(color).bg} ${getColor(color).text}`}>
            {emoji ? <span>{emoji}</span> : <Folder className="h-5 w-5" />}
          </div>
          <div className="truncate font-semibold text-slate-800">{name.trim() || "Untitled group"}</div>
        </div>

        <div className="space-y-4">
          <label className="block">
            <span className="mb-1 block text-xs font-medium text-slate-500">Name <span className="text-red-500">*</span></span>
            <input autoFocus value={name} onChange={(e) => setName(e.target.value)} placeholder="e.g. Languages" className="w-full rounded-lg border border-slate-300 bg-white px-3 py-2 text-sm focus:border-rose-400 focus:outline-none focus:ring-2 focus:ring-rose-100" />
          </label>
          <div>
            <span className="mb-1.5 block text-xs font-medium text-slate-500">Color</span>
            <div className="flex flex-wrap gap-2">
              {DECK_COLORS.map((c) => (
                <button key={c.id} onClick={() => setColor(c.id)} className={`h-7 w-7 rounded-full transition ${color === c.id ? "ring-2 ring-slate-900 ring-offset-2" : ""}`} style={{ backgroundColor: c.dot }} title={c.id} />
              ))}
            </div>
          </div>
          <div>
            <span className="mb-1.5 block text-xs font-medium text-slate-500">Icon (optional)</span>
            <div className="flex flex-wrap gap-1.5">
              <button onClick={() => setEmoji("")} className={`grid h-9 w-9 place-items-center rounded-lg border text-slate-400 transition ${emoji === "" ? "border-rose-500 bg-rose-50" : "border-slate-200 hover:bg-slate-50"}`}><Folder className="h-4 w-4" /></button>
              {DECK_EMOJIS.map((e) => (
                <button key={e} onClick={() => setEmoji(e)} className={`grid h-9 w-9 place-items-center rounded-lg border text-lg transition ${emoji === e ? "border-rose-500 bg-rose-50" : "border-slate-200 hover:bg-slate-50"}`}>{e}</button>
              ))}
            </div>
          </div>
        </div>

        <div className="mt-6 flex items-center justify-end gap-3">
          <button onClick={onClose} className="rounded-lg px-4 py-2 text-sm font-medium text-slate-500 hover:text-slate-800">Cancel</button>
          <button onClick={save} disabled={!name.trim()} className="inline-flex items-center gap-2 rounded-lg bg-rose-600 px-5 py-2 font-semibold text-white shadow-sm transition hover:bg-rose-700 disabled:cursor-not-allowed disabled:bg-slate-300">
            <Check className="h-4 w-4" /> {group ? "Save changes" : "Create group"}
          </button>
        </div>
      </div>
    </div>
  );
}

/* ------------------------------------------------------------------ */
/* Deck detail — browse/edit cards, see intervals + scheduling goal    */
/* ------------------------------------------------------------------ */
// Plain-words spaced-repetition schedule shown under each card.
function CardScheduleLine({ card }) {
  if (card.state === "new") return <div className="mt-0.5 text-[11px] italic text-slate-300">Нова — ще не в графіку повторень</div>;
  const s = cardScheduleText(card);
  return (
    <div className="mt-0.5 flex flex-wrap items-center gap-x-1.5 gap-y-0.5 text-[11px] text-slate-400">
      {s.prog && <span className="rounded bg-slate-100 px-1.5 py-0.5 font-medium text-slate-500" title="Як росли інтервали повторень">{s.prog}</span>}
      <span>інтервал {s.cur}</span>
      <span className="text-slate-300">·</span>
      <span>наступний повтор {s.next}</span>
    </div>
  );
}

function DeckDetailView({ deck, cards, summary, onBack, onStudy, onEditDeck, onAddCard, onEditCard, onDeleteCard }) {
  const [q, setQ] = useState("");
  const color = getColor(deck.color);
  const s = summary || { total: cards.length, due: 0, stages: {} };
  const goal = SCHED_GOALS[deck.goal] || SCHED_GOALS.longterm;
  const isDeadline = deck.goal === "deadline" && deck.deadline;
  const dleft = isDeadline ? daysUntil(deck.deadline) : null;
  const lang = deck.language || "";

  const filtered = useMemo(() => {
    const needle = q.trim().toLowerCase();
    const list = needle
      ? cards.filter((c) => (c.front + " " + c.back).toLowerCase().includes(needle))
      : cards;
    return list;
  }, [cards, q]);
  const shown = filtered.slice(0, 300);

  return (
    <div className="space-y-6">
      {/* header */}
      <div className="flex items-start gap-3">
        <button onClick={onBack} className="mt-1 rounded-md p-1.5 text-slate-400 hover:bg-slate-100 hover:text-slate-700">
          <ArrowLeft className="h-5 w-5" />
        </button>
        <div className={`grid h-12 w-12 shrink-0 place-items-center rounded-xl text-xl ${color.bg} ${color.text}`}>
          {deck.emoji ? <span>{deck.emoji}</span> : <GraduationCap className="h-6 w-6" />}
        </div>
        <div className="min-w-0 flex-1">
          <h1 className="flex items-center gap-2 text-xl font-bold text-slate-900">
            <span className="truncate">{deck.name}</span>
          </h1>
          <div className="mt-1 flex flex-wrap items-center gap-2 text-xs text-slate-500">
            <span>{cards.length} cards</span>
            {deck.topic && <span className="inline-flex items-center gap-1"><Tag className="h-3 w-3" />{deck.topic}</span>}
            <span className={`inline-flex items-center gap-1 rounded-full px-2 py-0.5 font-semibold ${isDeadline ? "bg-orange-100 text-orange-700" : "bg-blue-100 text-blue-700"}`}>
              <goal.icon className="h-3 w-3" /> {goal.short}
              {isDeadline && (dleft <= 0 ? " · date passed" : ` · ${dleft}d left`)}
            </span>
            {ttsSupported() && lang && <span className="inline-flex items-center gap-1"><Volume2 className="h-3 w-3" />{DECK_LANGUAGES.find((l) => l.code === lang)?.label || lang}</span>}
          </div>
        </div>
        <div className="flex shrink-0 gap-2">
          <button onClick={onEditDeck} className="inline-flex items-center gap-1.5 rounded-lg border border-slate-300 bg-white px-3 py-2 text-sm font-medium text-slate-700 transition hover:bg-slate-50">
            <Settings className="h-4 w-4" /> <span className="hidden sm:inline">Settings</span>
          </button>
          <button onClick={onStudy} disabled={!cards.length} className="inline-flex items-center gap-1.5 rounded-lg bg-rose-600 px-4 py-2 text-sm font-semibold text-white shadow-sm transition hover:bg-rose-700 disabled:bg-slate-300">
            <Play className="h-4 w-4" /> Study
          </button>
        </div>
      </div>

      {/* interval breakdown */}
      {cards.length > 0 && (
        <div className="rounded-xl border border-slate-200 bg-white p-4 shadow-sm">
          <div className="mb-3 flex items-center gap-2">
            <Timer className="h-4 w-4 text-slate-400" />
            <h2 className="text-sm font-semibold text-slate-700">Where your cards sit</h2>
          </div>
          <StageBreakdown stages={s.stages || {}} total={cards.length} />
          <div className="mt-4">
            <IntervalTimeline activeStageId={null} />
            <p className="mt-1.5 text-[11px] text-slate-400">
              {isDeadline
                ? `Deadline mode: intervals are capped to fit before your target date${dleft > 0 ? `, ${dleft} days away` : ""}.`
                : "Long-term mode: intervals keep growing — short (red) to long (blue)."}
            </p>
          </div>
        </div>
      )}

      {/* card list */}
      <div>
        <div className="mb-3 flex flex-wrap items-center gap-2">
          <h2 className="mr-auto text-sm font-semibold uppercase tracking-wide text-slate-400">Cards</h2>
          <div className="relative">
            <input
              value={q}
              onChange={(e) => setQ(e.target.value)}
              placeholder="Search…"
              className="w-40 rounded-lg border border-slate-300 bg-white px-3 py-1.5 text-sm focus:border-rose-400 focus:outline-none focus:ring-2 focus:ring-rose-100 sm:w-56"
            />
          </div>
          <button onClick={onAddCard} className="inline-flex items-center gap-1.5 rounded-lg bg-rose-600 px-3 py-1.5 text-sm font-semibold text-white transition hover:bg-rose-700">
            <Plus className="h-4 w-4" /> Add card
          </button>
        </div>

        {cards.length === 0 ? (
          <div className="rounded-xl border border-dashed border-slate-300 bg-white py-12 text-center text-slate-500">
            <p className="font-medium">No cards yet</p>
            <button onClick={onAddCard} className="mt-3 inline-flex items-center gap-1.5 rounded-lg bg-rose-600 px-4 py-2 text-sm font-semibold text-white transition hover:bg-rose-700">
              <Plus className="h-4 w-4" /> Add your first card
            </button>
          </div>
        ) : (
          <div className="divide-y divide-slate-100 overflow-hidden rounded-xl border border-slate-200 bg-white shadow-sm">
            {shown.map((c) => (
              <div key={c.id} className="group flex items-center gap-3 px-4 py-2.5">
                {(c.imgFront || c.imgBack) && <ImageIcon className="h-4 w-4 shrink-0 text-slate-300" />}
                <div className="min-w-0 flex-1">
                  <div className="truncate text-sm font-medium text-slate-800">{c.front || <span className="italic text-slate-400">(image)</span>}</div>
                  <div className="truncate text-xs text-slate-400">{c.back}</div>
                  <CardScheduleLine card={c} />
                </div>
                <StageBadge card={c} />
                <SpeakerButton text={c.front} lang={c.lang || lang} />
                <button onClick={() => onEditCard(c)} className="rounded-md p-1.5 text-slate-300 transition hover:bg-slate-100 hover:text-slate-600 sm:opacity-0 sm:group-hover:opacity-100" title="Edit card">
                  <Pencil className="h-4 w-4" />
                </button>
                <button onClick={() => { if (confirm("Delete this card?")) onDeleteCard(c.id); }} className="rounded-md p-1.5 text-slate-300 transition hover:bg-red-50 hover:text-red-500 sm:opacity-0 sm:group-hover:opacity-100" title="Delete card">
                  <Trash2 className="h-4 w-4" />
                </button>
              </div>
            ))}
            {filtered.length > shown.length && (
              <div className="px-4 py-2.5 text-center text-xs text-slate-400">
                Showing first {shown.length.toLocaleString()} of {filtered.length.toLocaleString()} — search to narrow down.
              </div>
            )}
          </div>
        )}
      </div>
    </div>
  );
}

/* ------------------------------------------------------------------ */
/* Card editor modal — text + images (upload / paste) + audio           */
/* ------------------------------------------------------------------ */
function CardEditor({ deck, card, onClose, onSave }) {
  const [front, setFront] = useState(card?.front || "");
  const [back, setBack] = useState(card?.back || "");
  const [lang, setLang] = useState(card?.lang || "");
  const [img, setImg] = useState({ front: null, back: null }); // current dataUrl or null
  const [dirty, setDirty] = useState({ front: false, back: false });
  const [pasteTarget, setPasteTarget] = useState("front");
  const [busy, setBusy] = useState(false);
  const [err, setErr] = useState("");
  const frontFile = useRef(null);
  const backFile = useRef(null);

  const effLang = lang || deck?.language || "";

  // load existing images for preview
  useEffect(() => {
    let alive = true;
    (async () => {
      const f = card?.imgFront ? await loadImage(card.id, "front") : null;
      const b = card?.imgBack ? await loadImage(card.id, "back") : null;
      if (alive) setImg({ front: f, back: b });
    })();
    return () => { alive = false; };
  }, [card?.id]);

  const setSideImage = async (side, fileOrDataUrl) => {
    setErr("");
    setBusy(true);
    try {
      const dataUrl = await compressImage(fileOrDataUrl);
      setImg((m) => ({ ...m, [side]: dataUrl }));
      setDirty((d) => ({ ...d, [side]: true }));
    } catch (e) {
      setErr(e.message || "Couldn't process that image.");
    } finally {
      setBusy(false);
    }
  };

  const removeSide = (side) => {
    setImg((m) => ({ ...m, [side]: null }));
    setDirty((d) => ({ ...d, [side]: true }));
  };

  const onPaste = (e) => {
    const file = imageFromClipboard(e);
    if (file) { e.preventDefault(); setSideImage(pasteTarget, file); }
  };

  const save = () => {
    if (!front.trim() && !img.front && !back.trim() && !img.back) {
      setErr("Add some text or an image first.");
      return;
    }
    const images = {
      front: dirty.front ? img.front : undefined,
      back: dirty.back ? img.back : undefined,
    };
    onSave({ front, back, lang }, images);
  };

  const ImageSlot = ({ side }) => (
    <div
      tabIndex={0}
      onClick={() => setPasteTarget(side)}
      onFocus={() => setPasteTarget(side)}
      className={`rounded-lg border-2 border-dashed p-2 transition ${pasteTarget === side ? "border-rose-400 bg-rose-50/40" : "border-slate-200"}`}
    >
      {img[side] ? (
        <div className="relative">
          <img src={img[side]} alt="" className="mx-auto max-h-40 w-auto rounded object-contain" />
          <button onClick={() => removeSide(side)} className="absolute right-1 top-1 rounded-full bg-slate-900/70 p-1 text-white hover:bg-slate-900" title="Remove image">
            <X className="h-3.5 w-3.5" />
          </button>
        </div>
      ) : (
        <div className="flex flex-col items-center gap-1.5 py-3 text-center">
          <ImagePlus className="h-5 w-5 text-slate-400" />
          <div className="flex items-center gap-2 text-xs">
            <button
              onClick={() => (side === "front" ? frontFile : backFile).current?.click()}
              className="font-medium text-rose-600 hover:text-rose-700"
            >
              Upload
            </button>
            <span className="text-slate-300">·</span>
            <span className="text-slate-400">click, then paste (⌘/Ctrl+V)</span>
          </div>
        </div>
      )}
      <input ref={side === "front" ? frontFile : backFile} type="file" accept="image/*" className="hidden"
        onChange={(e) => { const f = e.target.files?.[0]; if (f) setSideImage(side, f); e.target.value = ""; }} />
    </div>
  );

  return (
    <div className="fixed inset-0 z-50 flex items-end justify-center bg-slate-900/40 p-0 backdrop-blur-sm sm:items-center sm:p-4" onClick={onClose}>
      <div className="max-h-[92vh] w-full max-w-2xl overflow-y-auto rounded-t-2xl bg-white p-5 shadow-xl sm:rounded-2xl" onClick={(e) => e.stopPropagation()} onPaste={onPaste}>
        <div className="mb-4 flex items-center justify-between">
          <h2 className="text-lg font-bold text-slate-900">{card ? "Edit card" : "New card"}</h2>
          <button onClick={onClose} className="rounded-md p-1 text-slate-400 hover:bg-slate-100 hover:text-slate-700"><X className="h-5 w-5" /></button>
        </div>

        {err && <div className="mb-3 rounded-lg border border-red-200 bg-red-50 px-3 py-2 text-sm text-red-700">{err}</div>}

        <div className="grid gap-5 sm:grid-cols-2">
          {["front", "back"].map((side) => (
            <div key={side} className="space-y-2">
              <div className="flex items-center justify-between">
                <span className="text-xs font-semibold uppercase tracking-wide text-slate-400">{side}</span>
                <SpeakerButton text={side === "front" ? front : back} lang={effLang} />
              </div>
              <textarea
                value={side === "front" ? front : back}
                onChange={(e) => (side === "front" ? setFront : setBack)(e.target.value)}
                rows={3}
                placeholder={side === "front" ? "Question / prompt" : "Answer"}
                className="w-full rounded-lg border border-slate-300 bg-white px-3 py-2 text-sm focus:border-rose-400 focus:outline-none focus:ring-2 focus:ring-rose-100"
              />
              <ImageSlot side={side} />
            </div>
          ))}
        </div>

        <label className="mt-4 block">
          <span className="mb-1 block text-xs font-medium text-slate-500">Audio language override (optional)</span>
          <div className="relative sm:w-64">
            <select value={lang} onChange={(e) => setLang(e.target.value)} className="w-full appearance-none rounded-lg border border-slate-300 bg-white px-3 py-2 pr-9 text-sm focus:border-rose-400 focus:outline-none focus:ring-2 focus:ring-rose-100">
              {LANGUAGES.map((l) => <option key={l.code || "default"} value={l.code}>{l.code ? l.label : `Deck default (${DECK_LANGUAGES.find((x) => x.code === deck?.language)?.label || deck?.language || "—"})`}</option>)}
            </select>
            <ChevronDown className="pointer-events-none absolute right-3 top-1/2 h-4 w-4 -translate-y-1/2 text-slate-400" />
          </div>
        </label>

        <div className="mt-6 flex items-center justify-end gap-3">
          {busy && <span className="mr-auto text-xs text-slate-400">Processing image…</span>}
          <button onClick={onClose} className="rounded-lg px-4 py-2 text-sm font-medium text-slate-500 hover:text-slate-800">Cancel</button>
          <button onClick={save} disabled={busy} className="inline-flex items-center gap-2 rounded-lg bg-rose-600 px-5 py-2 font-semibold text-white shadow-sm transition hover:bg-rose-700 disabled:bg-slate-300">
            <Check className="h-4 w-4" /> {card ? "Save card" : "Add card"}
          </button>
        </div>
      </div>
    </div>
  );
}

/* ------------------------------------------------------------------ */
/* Import view                                                         */
/* ------------------------------------------------------------------ */
function ImportView({ decks, onImport, onCancel, onLoadEnglish, loadingEnglish }) {
  const [mode, setMode] = useState("file"); // file | paste
  const [parsed, setParsed] = useState(null); // { headers, rows, source }
  const [sheets, setSheets] = useState(null); // [{name, deckName, cards:[[f,b]], include}] for multi-sheet workbooks
  const [sourceName, setSourceName] = useState("");
  const [mapping, setMapping] = useState({ front: "", back: "", deck: "", tags: "", notes: "" });
  // where imported cards land: deckId "" means "create a new deck named `name`"
  const [target, setTarget] = useState({ deckId: "", name: "" });
  const [error, setError] = useState("");
  const [pasteText, setPasteText] = useState("");
  const fileRef = useRef(null);

  const resetFile = () => { setParsed(null); setSheets(null); setError(""); };

  const handleFile = (e) => {
    setError("");
    const file = e.target.files?.[0];
    if (!file) return;
    const base = file.name.replace(/\.[^.]+$/, "");
    setTarget((t) => ({ ...t, name: base }));
    const lower = file.name.toLowerCase();
    if (lower.endsWith(".csv")) {
      Papa.parse(file, {
        header: true,
        skipEmptyLines: true,
        complete: (res) => {
          const headers = (res.meta.fields || []).filter(Boolean);
          if (!headers.length) return setError("Couldn't find any columns in that CSV.");
          setParsed({ headers, rows: res.data, source: file.name });
          setMapping(autoMapColumns(headers));
        },
        error: (err) => setError(`CSV error: ${err.message}`),
      });
    } else {
      const reader = new FileReader();
      reader.onload = (ev) => {
        try {
          const data = new Uint8Array(ev.target.result);
          const wb = XLSX.read(data, { type: "array" });
          setSourceName(file.name);

          // scan every sheet for Front/Back data
          const found = [];
          for (const sn of wb.SheetNames) {
            const rowsAoA = XLSX.utils.sheet_to_json(wb.Sheets[sn], { header: 1, defval: "" });
            const { cards } = extractSheetCards(rowsAoA);
            if (cards.length) found.push({ name: sn, deckName: cleanDeckName(sn), cards, include: true });
          }
          if (!found.length) return setError("Couldn't find any Front/Back data in that file.");

          // more than one usable sheet -> per-sheet import picker
          if (found.length > 1) {
            setSheets(found);
            return;
          }

          // single sheet: use column-mapping unless the headers are messy,
          // in which case fall back to the smart extractor as a 1-sheet import
          const only = found[0];
          const ws = wb.Sheets[only.name];
          const json = XLSX.utils.sheet_to_json(ws, { defval: "" });
          const headers = json.length ? Object.keys(json[0]) : [];
          if (!headers.length || headersLookMessy(headers)) {
            setSheets(found);
          } else {
            setParsed({ headers, rows: json, source: file.name });
            setTarget((t) => ({ ...t, name: base }));
            setMapping(autoMapColumns(headers));
          }
        } catch (err) {
          setError(`Couldn't read that file: ${err.message}`);
        }
      };
      reader.readAsArrayBuffer(file);
    }
    e.target.value = ""; // allow re-selecting the same file
  };

  const previewCards = useMemo(() => {
    if (!parsed || !mapping.front) return [];
    return parsed.rows.slice(0, 6).map((r) => ({
      front: r[mapping.front],
      back: mapping.back ? r[mapping.back] : "",
      deck: mapping.deck ? r[mapping.deck] : "",
      tags: mapping.tags ? r[mapping.tags] : "",
    }));
  }, [parsed, mapping]);

  const validCount = useMemo(() => {
    if (!parsed || !mapping.front) return 0;
    return parsed.rows.filter((r) => String(r[mapping.front] ?? "").trim()).length;
  }, [parsed, mapping]);

  const targetDeck = decks.find((d) => d.id === target.deckId);
  const targetLabel = target.deckId ? targetDeck?.name : (target.name.trim() || "New deck");

  const commitFile = () => {
    if (!parsed || !mapping.front) return;

    // A "Deck" column splits rows across many decks (target picker ignored).
    if (mapping.deck) {
      const groups = {};
      for (const r of parsed.rows) {
        const front = String(r[mapping.front] ?? "").trim();
        if (!front) continue;
        const dn = String(r[mapping.deck] ?? "").trim() || target.name.trim() || "Imported";
        (groups[dn] ||= []).push(
          makeCard(front, mapping.back ? r[mapping.back] : "", mapping.tags ? r[mapping.tags] : "", mapping.notes ? r[mapping.notes] : "")
        );
      }
      if (!Object.keys(groups).length) return setError("No rows had a Front value.");
      return onImport(groups);
    }

    // Otherwise everything goes into the chosen target deck.
    const cards = [];
    for (const r of parsed.rows) {
      const front = String(r[mapping.front] ?? "").trim();
      if (!front) continue;
      cards.push(makeCard(front, mapping.back ? r[mapping.back] : "", mapping.tags ? r[mapping.tags] : "", mapping.notes ? r[mapping.notes] : ""));
    }
    if (!cards.length) return setError("No rows had a Front value.");
    if (target.deckId) onImport({ [targetDeck.name]: cards }, { targetDeckId: target.deckId });
    else onImport({ [target.name.trim() || "Imported"]: cards });
  };

  const includedSheets = sheets ? sheets.filter((s) => s.include) : [];
  const includedCount = includedSheets.reduce((n, s) => n + s.cards.length, 0);

  const setSheet = (i, patch) =>
    setSheets((arr) => arr.map((s, idx) => (idx === i ? { ...s, ...patch } : s)));
  const setAllSheets = (include) => setSheets((arr) => arr.map((s) => ({ ...s, include })));

  const commitSheets = () => {
    const groups = {};
    for (const s of includedSheets) {
      const dn = (s.deckName || "").trim() || s.name;
      (groups[dn] ||= []).push(...s.cards.map(([f, b]) => makeCard(f, b)));
    }
    if (!Object.keys(groups).length) return setError("Select at least one sheet to import.");
    onImport(groups);
  };

  const commitPaste = () => {
    const lines = pasteText.split("\n").map((l) => l.trim()).filter(Boolean);
    const cards = [];
    for (const line of lines) {
      let parts;
      if (line.includes("\t")) parts = line.split("\t");
      else if (line.includes("|")) parts = line.split("|");
      else parts = line.split(/,(.+)/); // split on first comma only
      const front = (parts[0] || "").trim();
      const back = (parts[1] || "").trim();
      if (front) cards.push(makeCard(front, back));
    }
    if (!cards.length) return setError("Type at least one line as  Front | Back");
    if (target.deckId) onImport({ [targetDeck.name]: cards }, { targetDeckId: target.deckId });
    else onImport({ [target.name.trim() || "Pasted cards"]: cards });
  };

  return (
    <div className="mx-auto max-w-2xl space-y-5">
      <div className="flex items-center gap-3">
        <button onClick={onCancel} className="rounded-md p-1.5 text-slate-400 hover:bg-slate-100 hover:text-slate-700">
          <ArrowLeft className="h-5 w-5" />
        </button>
        <h1 className="text-xl font-bold text-slate-900">Import cards</h1>
      </div>

      {onLoadEnglish && (
        <div className="flex flex-wrap items-center gap-3 rounded-2xl border border-blue-200 bg-blue-50 p-4">
          <span className="text-2xl">🇬🇧</span>
          <div className="min-w-0 flex-1">
            <div className="font-bold text-slate-800">Мої англійські колоди</div>
            <div className="text-xs text-slate-500">15 готових колод · ~9 500 карток. Один тап — і вони у тебе.</div>
          </div>
          <button onClick={onLoadEnglish} disabled={loadingEnglish} className="inline-flex shrink-0 items-center gap-2 rounded-xl bg-blue-600 px-5 py-2.5 font-bold text-white transition hover:bg-blue-700 disabled:opacity-60">
            {loadingEnglish ? <><RefreshCw className="h-4 w-4 animate-spin" /> Завантажую…</> : "Завантажити"}
          </button>
        </div>
      )}

      {/* mode toggle */}
      <div className="inline-flex rounded-lg border border-slate-200 bg-white p-1">
        <ModeTab active={mode === "file"} onClick={() => { setMode("file"); setError(""); }} icon={FileSpreadsheet}>Spreadsheet</ModeTab>
        <ModeTab active={mode === "paste"} onClick={() => { setMode("paste"); setError(""); }} icon={ClipboardPaste}>Paste text</ModeTab>
      </div>

      {error && (
        <div className="flex items-start gap-2 rounded-lg border border-red-200 bg-red-50 px-3 py-2 text-sm text-red-700">
          <X className="mt-0.5 h-4 w-4 shrink-0" /> {error}
        </div>
      )}

      {mode === "file" && !parsed && !sheets && (
        <div>
          <button
            onClick={() => fileRef.current?.click()}
            className="flex w-full flex-col items-center gap-2 rounded-xl border-2 border-dashed border-slate-300 bg-white py-12 text-slate-500 transition hover:border-rose-400 hover:text-rose-600"
          >
            <Upload className="h-8 w-8" />
            <span className="font-medium">Choose a .xlsx, .xls or .csv file</span>
            <span className="text-xs text-slate-400">Columns: Front, Back, and optional Deck, Tags, Notes — every sheet becomes a deck</span>
          </button>
          <input ref={fileRef} type="file" accept=".xlsx,.xls,.csv" onChange={handleFile} className="hidden" />
        </div>
      )}

      {mode === "file" && sheets && (
        <div className="space-y-4">
          <div className="flex items-center justify-between rounded-lg border border-slate-200 bg-white px-4 py-2.5 text-sm">
            <span className="flex items-center gap-2 text-slate-600">
              <FileSpreadsheet className="h-4 w-4 text-green-600" />
              <span className="font-medium">{sourceName}</span>
              <span className="text-slate-400">· {sheets.length} sheet{sheets.length === 1 ? "" : "s"} with cards</span>
            </span>
            <button onClick={resetFile} className="text-slate-400 hover:text-slate-700">
              <X className="h-4 w-4" />
            </button>
          </div>

          <div className="flex items-center justify-between">
            <p className="text-sm text-slate-500">Each sheet imports as its own deck. Untick any you don't want, or rename them.</p>
            <div className="flex shrink-0 gap-3 text-xs font-medium">
              <button onClick={() => setAllSheets(true)} className="text-rose-600 hover:text-rose-700">All</button>
              <button onClick={() => setAllSheets(false)} className="text-slate-400 hover:text-slate-600">None</button>
            </div>
          </div>

          <div className="max-h-[22rem] space-y-1.5 overflow-y-auto rounded-xl border border-slate-200 bg-white p-2">
            {sheets.map((s, i) => (
              <div
                key={s.name}
                className={`flex items-center gap-3 rounded-lg px-2 py-2 transition ${s.include ? "" : "opacity-50"}`}
              >
                <input
                  type="checkbox"
                  checked={s.include}
                  onChange={(e) => setSheet(i, { include: e.target.checked })}
                  className="h-4 w-4 shrink-0 accent-rose-600"
                />
                <input
                  value={s.deckName}
                  onChange={(e) => setSheet(i, { deckName: e.target.value })}
                  className="min-w-0 flex-1 rounded-md border border-transparent bg-transparent px-2 py-1 text-sm font-medium text-slate-800 hover:border-slate-200 focus:border-rose-400 focus:bg-white focus:outline-none focus:ring-2 focus:ring-rose-100"
                />
                <span className="shrink-0 rounded-full bg-slate-100 px-2 py-0.5 text-xs font-semibold tabular-nums text-slate-500">
                  {s.cards.length} cards
                </span>
              </div>
            ))}
          </div>

          <div className="flex items-center gap-3">
            <button
              onClick={commitSheets}
              disabled={!includedCount}
              className="inline-flex items-center gap-2 rounded-lg bg-rose-600 px-5 py-2.5 font-semibold text-white shadow-sm transition hover:bg-rose-700 disabled:cursor-not-allowed disabled:bg-slate-300"
            >
              <Check className="h-4 w-4" /> Import {includedCount.toLocaleString()} card{includedCount === 1 ? "" : "s"}
              <span className="opacity-80">· {includedSheets.length} deck{includedSheets.length === 1 ? "" : "s"}</span>
            </button>
            <button onClick={resetFile} className="text-sm font-medium text-slate-500 hover:text-slate-800">
              Choose another file
            </button>
          </div>
        </div>
      )}

      {mode === "file" && parsed && (
        <div className="space-y-5">
          <div className="flex items-center justify-between rounded-lg border border-slate-200 bg-white px-4 py-2.5 text-sm">
            <span className="flex items-center gap-2 text-slate-600">
              <FileSpreadsheet className="h-4 w-4 text-green-600" />
              <span className="font-medium">{parsed.source}</span>
              <span className="text-slate-400">· {parsed.rows.length} rows</span>
            </span>
            <button onClick={() => { setParsed(null); setError(""); }} className="text-slate-400 hover:text-slate-700">
              <X className="h-4 w-4" />
            </button>
          </div>

          {/* column mapping */}
          <div>
            <h3 className="mb-2 text-sm font-semibold text-slate-700">Map your columns</h3>
            <div className="grid gap-3 sm:grid-cols-2">
              {[
                { field: "front", label: "Front", required: true },
                { field: "back", label: "Back", required: false },
                { field: "deck", label: "Deck (optional)", required: false },
                { field: "tags", label: "Tags (optional)", required: false },
                { field: "notes", label: "Notes (optional)", required: false },
              ].map(({ field, label, required }) => (
                <label key={field} className="block">
                  <span className="mb-1 block text-xs font-medium text-slate-500">
                    {label} {required && <span className="text-red-500">*</span>}
                  </span>
                  <select
                    value={mapping[field]}
                    onChange={(e) => setMapping((m) => ({ ...m, [field]: e.target.value }))}
                    className="w-full rounded-lg border border-slate-300 bg-white px-3 py-2 text-sm focus:border-rose-400 focus:outline-none focus:ring-2 focus:ring-rose-100"
                  >
                    <option value="">— none —</option>
                    {parsed.headers.map((h) => (
                      <option key={h} value={h}>{h}</option>
                    ))}
                  </select>
                </label>
              ))}
            </div>
          </div>

          {!mapping.deck ? (
            <DeckTargetPicker decks={decks} target={target} onChange={setTarget} defaultName="Imported" />
          ) : (
            <p className="rounded-lg bg-slate-50 px-3 py-2 text-xs text-slate-500">
              Rows are split into decks by the “{mapping.deck}” column. Matching existing decks get the new cards appended.
            </p>
          )}

          {/* preview */}
          {previewCards.length > 0 && (
            <div>
              <h3 className="mb-2 text-sm font-semibold text-slate-700">Preview</h3>
              <div className="overflow-x-auto rounded-lg border border-slate-200">
                <table className="w-full text-left text-sm">
                  <thead className="bg-slate-50 text-xs uppercase tracking-wide text-slate-400">
                    <tr>
                      <th className="px-3 py-2 font-medium">Front</th>
                      <th className="px-3 py-2 font-medium">Back</th>
                      {mapping.deck && <th className="px-3 py-2 font-medium">Deck</th>}
                      {mapping.tags && <th className="px-3 py-2 font-medium">Tags</th>}
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100">
                    {previewCards.map((c, i) => (
                      <tr key={i}>
                        <td className="px-3 py-2 text-slate-800">{String(c.front)}</td>
                        <td className="px-3 py-2 text-slate-500">{String(c.back)}</td>
                        {mapping.deck && <td className="px-3 py-2 text-slate-500">{String(c.deck)}</td>}
                        {mapping.tags && <td className="px-3 py-2 text-slate-400">{String(c.tags)}</td>}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          )}

          <div className="flex items-center gap-3 pt-1">
            <button
              onClick={commitFile}
              disabled={!mapping.front || !validCount}
              className="inline-flex items-center gap-2 rounded-lg bg-rose-600 px-5 py-2.5 font-semibold text-white shadow-sm transition hover:bg-rose-700 disabled:cursor-not-allowed disabled:bg-slate-300"
            >
              <Check className="h-4 w-4" /> Import {validCount} card{validCount === 1 ? "" : "s"}
            </button>
            <button onClick={() => { setParsed(null); setError(""); }} className="text-sm font-medium text-slate-500 hover:text-slate-800">
              Choose another file
            </button>
          </div>
        </div>
      )}

      {mode === "paste" && (
        <div className="space-y-4">
          <DeckTargetPicker decks={decks} target={target} onChange={setTarget} defaultName="Pasted cards" />
          <div>
            <label className="mb-1 block text-xs font-medium text-slate-500">
              One card per line — <span className="font-mono">Front | Back</span> (also accepts tab or comma)
            </label>
            <textarea
              value={pasteText}
              onChange={(e) => setPasteText(e.target.value)}
              rows={9}
              placeholder={"bonjour | hello\nmerci | thank you\nau revoir | goodbye"}
              className="w-full rounded-lg border border-slate-300 bg-white px-3 py-2 font-mono text-sm focus:border-rose-400 focus:outline-none focus:ring-2 focus:ring-rose-100"
            />
          </div>
          <button
            onClick={commitPaste}
            disabled={!pasteText.trim()}
            className="inline-flex items-center gap-2 rounded-lg bg-rose-600 px-5 py-2.5 font-semibold text-white shadow-sm transition hover:bg-rose-700 disabled:cursor-not-allowed disabled:bg-slate-300"
          >
            <Check className="h-4 w-4" /> Add cards
          </button>
        </div>
      )}
    </div>
  );
}

function ModeTab({ active, onClick, icon: Icon, children }) {
  return (
    <button
      onClick={onClick}
      className={`flex items-center gap-1.5 rounded-md px-3 py-1.5 text-sm font-medium transition ${
        active ? "bg-rose-600 text-white" : "text-slate-500 hover:text-slate-800"
      }`}
    >
      <Icon className="h-4 w-4" /> {children}
    </button>
  );
}

// Choose where imported cards go: an existing deck (append) or a brand-new one.
function DeckTargetPicker({ decks, target, onChange, defaultName }) {
  return (
    <div className="space-y-2">
      <label className="block text-xs font-medium text-slate-500">Add cards to</label>
      <div className="flex flex-col gap-2 sm:flex-row sm:items-center">
        <div className="relative sm:w-64">
          <select
            value={target.deckId}
            onChange={(e) => onChange({ ...target, deckId: e.target.value })}
            className="w-full appearance-none rounded-lg border border-slate-300 bg-white px-3 py-2 pr-9 text-sm focus:border-rose-400 focus:outline-none focus:ring-2 focus:ring-rose-100"
          >
            <option value="">➕ Create a new deck…</option>
            {decks.length > 0 && <option disabled>──────────</option>}
            {decks.map((d) => (
              <option key={d.id} value={d.id}>{d.emoji ? `${d.emoji} ` : ""}{d.name}</option>
            ))}
          </select>
          <ChevronDown className="pointer-events-none absolute right-3 top-1/2 h-4 w-4 -translate-y-1/2 text-slate-400" />
        </div>
        {!target.deckId && (
          <input
            value={target.name}
            onChange={(e) => onChange({ ...target, name: e.target.value })}
            placeholder={defaultName || "New deck name"}
            className="flex-1 rounded-lg border border-slate-300 bg-white px-3 py-2 text-sm focus:border-rose-400 focus:outline-none focus:ring-2 focus:ring-rose-100"
          />
        )}
      </div>
      {target.deckId && (
        <p className="text-xs text-slate-400">New cards will be appended — nothing already in the deck is removed.</p>
      )}
    </div>
  );
}

/* ------------------------------------------------------------------ */
/* Stats view                                                          */
/* ------------------------------------------------------------------ */
function StatsView({ stats, decks, cardsByDeck, totalDue, onExport, onReset, onChangeNewPerDay }) {
  const streak = computeStreak(stats.history);
  const studiedToday = stats.history?.[dateKey(Date.now())]?.studied || 0;
  const retention = retention30(stats.history);
  const newPerDay = stats.settings?.newPerDay ?? DEFAULT_NEW_PER_DAY;

  const totalCards = useMemo(
    () => Object.values(cardsByDeck).reduce((s, arr) => s + arr.length, 0),
    [cardsByDeck]
  );

  // upcoming reviews (next 14 days)
  const upcoming = useMemo(() => {
    const now = Date.now();
    const startOfToday = new Date(); startOfToday.setHours(0, 0, 0, 0);
    const buckets = Array.from({ length: 14 }, (_, i) => {
      const day = new Date(startOfToday.getTime() + i * DAY);
      return { label: i === 0 ? "Today" : day.toLocaleDateString(undefined, { weekday: "short", day: "numeric" }), count: 0, i };
    });
    for (const arr of Object.values(cardsByDeck)) {
      for (const c of arr) {
        if (c.state !== "review") continue;
        const idx = Math.floor((c.due - startOfToday.getTime()) / DAY);
        if (idx >= 0 && idx < 14) buckets[idx].count += 1;
        else if (idx < 0) buckets[0].count += 1; // overdue rolls into today
      }
    }
    return buckets;
  }, [cardsByDeck]);

  // last 7 days studied
  const last7 = useMemo(() => {
    const now = Date.now();
    return Array.from({ length: 7 }, (_, i) => {
      const ms = now - (6 - i) * DAY;
      const d = new Date(ms);
      return {
        label: d.toLocaleDateString(undefined, { weekday: "short" }),
        studied: stats.history?.[dateKey(ms)]?.studied || 0,
      };
    });
  }, [stats]);

  const maxUpcoming = Math.max(1, ...upcoming.map((b) => b.count));

  return (
    <div className="space-y-6">
      <h1 className="text-xl font-bold text-slate-900">Your progress</h1>

      <div className="flex flex-wrap gap-3">
        <StatTile icon={Check} label="Studied today" value={studiedToday} tint="text-rose-600" />
        <StatTile icon={Flame} label="Streak" value={streak} tint={streak ? "text-orange-500" : "text-slate-400"} sub={streak === 1 ? "day" : "days"} />
        <StatTile icon={Target} label="Retention" value={retention == null ? "—" : `${retention}%`} tint="text-green-600" sub="last 30 days" />
        <StatTile icon={Inbox} label="Due now" value={totalDue} tint={totalDue ? "text-slate-700" : "text-slate-400"} />
        <StatTile icon={Layers} label="Total cards" value={totalCards} tint="text-slate-700" />
      </div>

      {/* upcoming chart */}
      <div className="rounded-xl border border-slate-200 bg-white p-5 shadow-sm">
        <div className="mb-1 flex items-center gap-2">
          <Clock className="h-4 w-4 text-slate-400" />
          <h2 className="text-sm font-semibold text-slate-700">Upcoming reviews · next 14 days</h2>
        </div>
        <p className="mb-4 text-xs text-slate-400">When your review cards are next scheduled to come back.</p>
        <div className="h-56 w-full">
          <ResponsiveContainer width="100%" height="100%">
            <BarChart data={upcoming} margin={{ top: 4, right: 4, left: -20, bottom: 0 }}>
              <CartesianGrid strokeDasharray="3 3" stroke="#eef2f7" vertical={false} />
              <XAxis dataKey="label" tick={{ fontSize: 11, fill: "#94a3b8" }} axisLine={false} tickLine={false} interval={0} angle={-30} textAnchor="end" height={44} />
              <YAxis allowDecimals={false} tick={{ fontSize: 11, fill: "#94a3b8" }} axisLine={false} tickLine={false} />
              <Tooltip
                cursor={{ fill: "#eef2ff" }}
                contentStyle={{ borderRadius: 10, border: "1px solid #e2e8f0", fontSize: 12 }}
                labelStyle={{ color: "#475569", fontWeight: 600 }}
              />
              <Bar dataKey="count" radius={[4, 4, 0, 0]}>
                {upcoming.map((b, i) => (
                  <Cell key={i} fill={i === 0 ? "#4f46e5" : "#c7d2fe"} />
                ))}
              </Bar>
            </BarChart>
          </ResponsiveContainer>
        </div>
      </div>

      {/* last 7 days */}
      <div className="rounded-xl border border-slate-200 bg-white p-5 shadow-sm">
        <h2 className="mb-4 flex items-center gap-2 text-sm font-semibold text-slate-700">
          <BarChart3 className="h-4 w-4 text-slate-400" /> Cards studied · last 7 days
        </h2>
        <div className="flex items-end justify-between gap-2" style={{ height: 120 }}>
          {last7.map((d, i) => {
            const max = Math.max(1, ...last7.map((x) => x.studied));
            const h = Math.round((d.studied / max) * 100);
            return (
              <div key={i} className="flex flex-1 flex-col items-center gap-1">
                <div className="flex w-full flex-1 items-end">
                  <div
                    className="w-full rounded-t bg-rose-500 transition-all"
                    style={{ height: `${d.studied ? Math.max(6, h) : 0}%`, minHeight: d.studied ? 6 : 0 }}
                    title={`${d.studied} cards`}
                  />
                </div>
                <span className="text-[11px] font-medium tabular-nums text-slate-600">{d.studied || ""}</span>
                <span className="text-[11px] text-slate-400">{d.label}</span>
              </div>
            );
          })}
        </div>
      </div>

      {/* cloud sync */}
      <CloudSyncPanel />

      {/* settings + data */}
      <div className="rounded-xl border border-slate-200 bg-white p-5 shadow-sm">
        <h2 className="mb-4 text-sm font-semibold text-slate-700">Settings & data</h2>
        <label className="flex items-center justify-between gap-4 border-b border-slate-100 pb-4">
          <span>
            <span className="block text-sm font-medium text-slate-700">New cards per day</span>
            <span className="block text-xs text-slate-400">How many brand-new cards to introduce daily.</span>
          </span>
          <input
            type="number" min={0} max={999} value={newPerDay}
            onChange={(e) => onChangeNewPerDay(Math.max(0, Math.min(999, Number(e.target.value) || 0)))}
            className="w-20 rounded-lg border border-slate-300 px-3 py-1.5 text-right text-sm tabular-nums focus:border-rose-400 focus:outline-none focus:ring-2 focus:ring-rose-100"
          />
        </label>
        <div className="mt-4 flex flex-wrap gap-3">
          <button onClick={onExport} className="inline-flex items-center gap-2 rounded-lg border border-slate-300 bg-white px-4 py-2 text-sm font-medium text-slate-700 transition hover:border-slate-400 hover:bg-slate-50">
            <Download className="h-4 w-4" /> Export backup (JSON)
          </button>
          <button
            onClick={() => { if (confirm("Reset ALL data — decks, cards, stats AND your routine/habits? This cannot be undone.")) onReset(); }}
            className="inline-flex items-center gap-2 rounded-lg border border-red-200 bg-white px-4 py-2 text-sm font-medium text-red-600 transition hover:bg-red-50"
          >
            <RotateCcw className="h-4 w-4" /> Reset everything
          </button>
        </div>
      </div>
    </div>
  );
}

/* ------------------------------------------------------------------ */
/* Cloud sync (Supabase) — optional cross-device sync                  */
/* ------------------------------------------------------------------ */
function CloudSyncPanel() {
  const [signed, setSigned] = useState(false);
  const [email, setEmail] = useState(null);
  const [step, setStep] = useState("idle"); // idle | email | code
  const [input, setInput] = useState("");
  const [code, setCode] = useState("");
  const [link, setLink] = useState("");
  const [busy, setBusy] = useState(false);
  const [err, setErr] = useState("");

  useEffect(() => { (async () => { await refreshSession(); setSigned(cloudSignedIn()); setEmail(currentEmail()); })(); }, []);

  const doSend = async () => {
    setErr(""); if (!input.trim()) return;
    setBusy(true);
    try { await sendCode(input); setStep("code"); } catch (e) { setErr(e?.message || "Не вдалося надіслати код."); } finally { setBusy(false); }
  };
  const doVerify = async () => {
    setErr(""); if (!code.trim()) return;
    setBusy(true);
    try { await verifyCode(input, code); location.reload(); }
    catch (e) {
      const msg = e?.message || "";
      if (/kv|schema cache|relation|does not exist/i.test(msg)) setErr("Майже готово — створи таблицю kv у Supabase (SQL з інструкції), і синхронізація запрацює.");
      else setErr(msg || "Невірний або застарілий код.");
      setBusy(false);
    }
  };
  const doLink = async () => {
    setErr(""); if (!link.trim()) return;
    setBusy(true);
    try { await signInWithLink(link); location.reload(); }
    catch (e) { setErr(e?.message || "Не вдалося увійти за посиланням."); setBusy(false); }
  };
  const doSignOut = async () => { setBusy(true); await signOutCloud(); setSigned(false); setEmail(null); setStep("idle"); setBusy(false); };

  return (
    <div className="rounded-xl border border-slate-200 bg-white p-5 shadow-sm">
      <div className="mb-1 flex items-center gap-2">
        {signed ? <Cloud className="h-4 w-4 text-green-500" /> : <CloudOff className="h-4 w-4 text-slate-400" />}
        <h2 className="text-sm font-semibold text-slate-700">Хмарна синхронізація (між пристроями)</h2>
      </div>

      {signed ? (
        <div className="mt-2 flex items-center justify-between gap-3">
          <p className="text-sm text-slate-500">Синхронізовано як <span className="font-semibold text-slate-700">{email}</span>. Дані збережено в хмарі й синхронізуються між пристроями.</p>
          <button onClick={doSignOut} disabled={busy} className="inline-flex shrink-0 items-center gap-1.5 rounded-lg border border-slate-300 px-3 py-1.5 text-sm font-medium text-slate-600 hover:bg-slate-50"><LogOut className="h-4 w-4" /> Вийти</button>
        </div>
      ) : (
        <>
          <p className="mb-3 text-sm text-slate-500">Щоб <b>усе зберігалося в базі даних</b> і синхронізувалося між телефоном і компʼютером — увійди своєю поштою. У листі буде код і посилання для входу. Без входу дані живуть лише на цьому пристрої.</p>
          {step !== "code" ? (
            <div className="flex flex-wrap items-center gap-2">
              <div className="relative flex-1 min-w-[200px]">
                <Mail className="pointer-events-none absolute left-3 top-1/2 h-4 w-4 -translate-y-1/2 text-slate-400" />
                <input type="email" value={input} onChange={(e) => setInput(e.target.value)} onKeyDown={(e) => { if (e.key === "Enter") doSend(); }} placeholder="you@email.com" className="w-full rounded-lg border border-slate-300 py-2 pl-9 pr-3 text-sm focus:border-rose-400 focus:outline-none focus:ring-2 focus:ring-rose-100" />
              </div>
              <button onClick={doSend} disabled={busy || !input.trim()} className="rounded-lg bg-rose-600 px-4 py-2 text-sm font-semibold text-white transition hover:bg-rose-700 disabled:bg-slate-300">{busy ? "…" : "Надіслати код"}</button>
            </div>
          ) : (
            <div className="flex flex-wrap items-center gap-2">
              <input value={code} onChange={(e) => setCode(e.target.value.replace(/\D/g, "").slice(0, 6))} onKeyDown={(e) => { if (e.key === "Enter") doVerify(); }} placeholder="6-значний код" inputMode="numeric" className="w-40 rounded-lg border border-slate-300 px-3 py-2 text-center text-sm tracking-widest focus:border-rose-400 focus:outline-none focus:ring-2 focus:ring-rose-100" />
              <button onClick={doVerify} disabled={busy || code.length < 6} className="rounded-lg bg-rose-600 px-4 py-2 text-sm font-semibold text-white transition hover:bg-rose-700 disabled:bg-slate-300">{busy ? "…" : "Підтвердити й синхронізувати"}</button>
              <button onClick={() => { setStep("idle"); setCode(""); setErr(""); }} className="text-sm font-medium text-slate-400 hover:text-slate-600">Змінити пошту</button>
              <span className="w-full text-xs text-slate-400">Лист надіслано на {input} (перевір і спам). Введи код — або скористайся посиланням нижче.</span>
            </div>
          )}

          <details className="mt-3 rounded-lg bg-slate-50 p-3">
            <summary className="cursor-pointer text-xs font-semibold text-slate-600">Прийшло тільки посилання, без коду? Натисни сюди 👇</summary>
            <p className="mt-2 text-xs leading-relaxed text-slate-500">1) Клікни посилання в листі. 2) Куди б воно не привело — <b>скопіюй увесь URL зі стрічки адреси браузера</b> (там усередині є «access_token=…»). 3) Встав його сюди:</p>
            <div className="mt-2 flex flex-wrap items-center gap-2">
              <input value={link} onChange={(e) => setLink(e.target.value)} placeholder="https://…#access_token=…" className="min-w-[200px] flex-1 rounded-lg border border-slate-300 px-3 py-2 text-xs focus:border-rose-400 focus:outline-none focus:ring-2 focus:ring-rose-100" />
              <button onClick={doLink} disabled={busy || !link.trim()} className="rounded-lg bg-rose-600 px-4 py-2 text-sm font-semibold text-white transition hover:bg-rose-700 disabled:bg-slate-300">{busy ? "…" : "Увійти за посиланням"}</button>
            </div>
          </details>
          {err && <p className="mt-2 text-sm text-red-600">{err}</p>}
        </>
      )}
    </div>
  );
}

/* ================================================================== */
/* MY ROUTINE — Me+ style UI                                          */
/* ================================================================== */
function ProgressRing({ pct, size = 64, stroke = 6, color = "#ec4899", track = "#f7dceb", children }) {
  const r = (size - stroke) / 2;
  const c = 2 * Math.PI * r;
  const off = c * (1 - Math.max(0, Math.min(1, pct)));
  return (
    <div className="relative shrink-0" style={{ width: size, height: size }}>
      <svg width={size} height={size} className="-rotate-90">
        <circle cx={size / 2} cy={size / 2} r={r} fill="none" stroke={track} strokeWidth={stroke} />
        <circle cx={size / 2} cy={size / 2} r={r} fill="none" stroke={color} strokeWidth={stroke} strokeDasharray={c} strokeDashoffset={off} strokeLinecap="round" style={{ transition: "stroke-dashoffset .5s ease" }} />
      </svg>
      <div className="absolute inset-0 flex flex-col items-center justify-center">{children}</div>
    </div>
  );
}

const weekDaysOf = (anchorDs) => {
  const base = new Date(anchorDs + "T00:00:00");
  const dow = (base.getDay() + 6) % 7; // Mon=0
  const mon = new Date(base); mon.setDate(base.getDate() - dow);
  return Array.from({ length: 7 }, (_, i) => { const d = new Date(mon); d.setDate(mon.getDate() + i); return dateKey(d.getTime()); });
};
const prettyDate = (ds) => new Date(ds + "T00:00:00").toLocaleDateString(undefined, { weekday: "long", month: "long", day: "numeric" });

/* ---------- ADHD gamification: XP bar, challenges, wheel, focus, rewards, recap ---------- */
function GamifyBar({ xp, onRewards }) {
  const lp = levelProgress(xp);
  return (
    <button onClick={onRewards} className="mt-4 flex w-full items-center gap-3 rounded-2xl bg-white/80 p-3 text-left shadow-sm ring-1 ring-red-100 transition hover:ring-red-200">
      <span className="grid h-11 w-11 shrink-0 place-items-center rounded-2xl bg-gradient-to-br from-pink-400 to-fuchsia-400 text-white shadow-sm">
        <span className="text-xs font-black leading-none">LVL</span>
      </span>
      <div className="min-w-0 flex-1">
        <div className="flex items-baseline justify-between">
          <span className="text-sm font-extrabold text-slate-800">Рівень {lp.lvl}</span>
          <span className="text-[11px] font-semibold tabular-nums text-slate-400">{xp} XP</span>
        </div>
        <div className="mt-1 h-2.5 overflow-hidden rounded-full bg-pink-100">
          <div className="h-full rounded-full bg-gradient-to-r from-pink-400 to-fuchsia-400 transition-all" style={{ width: `${lp.pct * 100}%` }} />
        </div>
        <div className="mt-0.5 text-[10px] text-slate-400">{lp.next - xp} XP до рівня {lp.lvl + 1} · нагороди 🎁</div>
      </div>
    </button>
  );
}

function ChallengesCard({ challenges, ctx, chDoc, onDismiss }) {
  const visible = challenges.filter((c) => !chDoc.dismissed?.[c.id]);
  if (!visible.length) return null;
  return (
    <div className="mt-4 rounded-2xl bg-white/80 p-3.5 shadow-sm ring-1 ring-red-100">
      <div className="mb-2 flex items-center gap-1.5 text-sm font-extrabold text-slate-800"><span>🎯</span> Челенджі дня</div>
      <div className="space-y-2">
        {visible.map((c) => {
          const done = c.check(ctx);
          return (
            <div key={c.id} className={`flex items-center gap-2.5 rounded-xl px-3 py-2 ${done ? "bg-green-50 ring-1 ring-green-200" : "bg-slate-50"}`}>
              <span className="text-lg">{c.emoji}</span>
              <span className={`min-w-0 flex-1 text-sm ${done ? "font-semibold text-green-700 line-through" : "text-slate-600"}`}>{c.label}</span>
              <span className="shrink-0 rounded-full bg-white px-2 py-0.5 text-[10px] font-bold text-pink-500 ring-1 ring-pink-100">+{c.xp}</span>
              {done ? <Check className="h-4 w-4 shrink-0 text-green-500" /> : <button onClick={() => onDismiss(c.id)} className="shrink-0 rounded-full p-0.5 text-slate-300 hover:text-slate-500"><X className="h-4 w-4" /></button>}
            </div>
          );
        })}
      </div>
    </div>
  );
}

function WheelSpin({ tasks, onClose, onPick }) {
  const [picked, setPicked] = useState(null);
  const [spinning, setSpinning] = useState(false);
  const spin = () => {
    if (!tasks.length) return;
    setSpinning(true);
    let n = 0;
    const iv = setInterval(() => {
      setPicked(tasks[Math.floor((n * 7) % tasks.length)]);
      n += 1;
      if (n > 14) { clearInterval(iv); const final = tasks[Math.floor((n * 7 + 3) % tasks.length)]; setPicked(final); setSpinning(false); }
    }, 90);
  };
  return (
    <div className="fixed inset-0 z-50 flex items-end justify-center bg-slate-900/40 p-0 backdrop-blur-sm sm:items-center sm:p-4" onClick={onClose}>
      <div className="w-full max-w-sm rounded-t-3xl bg-white p-6 text-center shadow-xl sm:rounded-3xl" onClick={(e) => e.stopPropagation()}>
        <div className="text-4xl">🎡</div>
        <h3 className="mt-2 text-lg font-extrabold text-slate-900">Колесо задач</h3>
        <p className="mt-1 text-sm text-slate-400">Не знаєш, з чого почати? Хай вирішить колесо.</p>
        {tasks.length === 0 ? (
          <p className="mt-6 text-sm text-slate-400">Немає незавершених справ на сьогодні 🎉</p>
        ) : (
          <>
            <div className="my-5 grid min-h-[64px] place-items-center rounded-2xl bg-pink-50 px-4 py-4 ring-1 ring-pink-100">
              {picked ? <div className="flex items-center gap-2 text-lg font-bold text-slate-800"><span className="text-2xl">{picked.emoji || "⭐"}</span>{picked.title}</div> : <span className="text-sm text-slate-400">Крути, щоб обрати</span>}
            </div>
            <div className="flex gap-2">
              <button onClick={spin} disabled={spinning} className="flex-1 rounded-2xl bg-pink-500 py-3 font-bold text-white shadow-lg shadow-pink-500/20 hover:bg-pink-600 disabled:opacity-60">{picked ? "Ще раз" : "Крутити"}</button>
              {picked && !spinning && <button onClick={() => onPick(picked)} className="flex-1 rounded-2xl bg-slate-800 py-3 font-bold text-white hover:bg-slate-900">Робити це</button>}
            </div>
          </>
        )}
      </div>
    </div>
  );
}

function FocusMode({ tasks, doc, onClose, onDone, onSkip }) {
  const [idx, setIdx] = useState(0);
  const task = tasks[idx];
  if (!task) return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-gradient-to-b from-pink-500 to-fuchsia-500 p-6 text-center text-white" onClick={onClose}>
      <div><div className="text-5xl">🎉</div><h2 className="mt-3 text-2xl font-extrabold">Усе на зараз закрито!</h2><p className="mt-1 text-white/80">Можеш видихнути.</p><button onClick={onClose} className="mt-6 rounded-2xl bg-white px-8 py-3 font-bold text-pink-600">Готово</button></div>
    </div>
  );
  const p = getPastel(task.color);
  return (
    <div className="fixed inset-0 z-50 flex flex-col bg-gradient-to-b from-red-50 to-pink-100 p-6">
      <div className="flex items-center justify-between">
        <span className="text-sm font-bold text-slate-400">Зараз · {idx + 1}/{tasks.length}</span>
        <button onClick={onClose} className="rounded-full bg-white/70 p-2 text-slate-500"><X className="h-5 w-5" /></button>
      </div>
      <div className="flex flex-1 flex-col items-center justify-center text-center">
        <div className="grid h-28 w-28 place-items-center rounded-3xl text-5xl shadow-lg" style={{ backgroundColor: p.card }}>{task.emoji || "⭐"}</div>
        <h1 className="mt-5 max-w-md text-3xl font-extrabold text-slate-900">{task.title}</h1>
        {task.note && <p className="mt-2 max-w-sm text-sm text-slate-500">{task.note}</p>}
        <div className="mt-3 flex flex-wrap justify-center gap-2 text-xs">
          {task.estMin && <span className="rounded-full bg-white px-3 py-1 font-semibold text-slate-500 ring-1 ring-slate-200">≈ {fmtEst(task.estMin)}</span>}
          {task.energy && ENERGY[task.energy] && <span className="rounded-full bg-white px-3 py-1 font-semibold text-slate-500 ring-1 ring-slate-200">{ENERGY[task.energy].emoji} {ENERGY[task.energy].label}</span>}
        </div>
      </div>
      <div className="flex gap-3">
        <button onClick={() => setIdx((i) => i + 1)} className="flex-1 rounded-2xl bg-white py-4 font-bold text-slate-500 shadow-sm ring-1 ring-slate-200">Пропустити →</button>
        <button onClick={() => { onDone(task); setIdx((i) => i); }} className="flex-[2] rounded-2xl bg-pink-500 py-4 font-bold text-white shadow-lg shadow-pink-500/25 hover:bg-pink-600">Готово ✓</button>
      </div>
    </div>
  );
}

function RewardsPanel({ xp, rewards, onClose, onAdd, onUnlock, onDelete }) {
  const [label, setLabel] = useState("");
  const [cost, setCost] = useState(300);
  const lp = levelProgress(xp);
  return (
    <div className="fixed inset-0 z-50 flex items-end justify-center bg-slate-900/40 p-0 backdrop-blur-sm sm:items-center sm:p-4" onClick={onClose}>
      <div className="max-h-[88vh] w-full max-w-md overflow-y-auto rounded-t-3xl bg-white p-5 shadow-xl sm:rounded-3xl" onClick={(e) => e.stopPropagation()}>
        <div className="mb-1 flex items-center justify-between">
          <h3 className="text-lg font-extrabold text-slate-900">🎁 Мої нагороди</h3>
          <button onClick={onClose} className="rounded-full p-1 text-slate-400 hover:text-slate-600"><X className="h-5 w-5" /></button>
        </div>
        <div className="mb-3 rounded-2xl bg-pink-50 p-3">
          <div className="flex items-baseline justify-between text-sm"><span className="font-bold text-slate-700">Рівень {lp.lvl}</span><span className="font-semibold tabular-nums text-pink-500">{xp} XP</span></div>
          <div className="mt-1.5 h-2 overflow-hidden rounded-full bg-pink-100"><div className="h-full rounded-full bg-gradient-to-r from-pink-400 to-fuchsia-400" style={{ width: `${lp.pct * 100}%` }} /></div>
        </div>
        <p className="mb-3 text-xs text-slate-400">Придумай собі маленькі приємності й відмикай їх, коли назбираєш XP. Це твоє особисте «меню дофаміну».</p>
        <div className="space-y-2">
          {rewards.length === 0 && <div className="rounded-xl bg-slate-50 py-6 text-center text-sm text-slate-400">Ще немає нагород. Додай першу нижче 👇</div>}
          {rewards.map((r) => {
            const unlocked = !!r.unlockedAt;
            const can = xp >= r.cost;
            return (
              <div key={r.id} className={`flex items-center gap-2 rounded-xl px-3 py-2.5 ${unlocked ? "bg-green-50 ring-1 ring-green-200" : "bg-slate-50"}`}>
                <span className="text-lg">{unlocked ? "🎉" : can ? "🔓" : "🔒"}</span>
                <div className="min-w-0 flex-1"><div className={`truncate text-sm font-semibold ${unlocked ? "text-green-700" : "text-slate-700"}`}>{r.label}</div><div className="text-[11px] text-slate-400">{r.cost} XP{unlocked ? " · відкрито" : ""}</div></div>
                {!unlocked && <button onClick={() => onUnlock(r.id)} disabled={!can} className="shrink-0 rounded-full bg-pink-500 px-3 py-1 text-xs font-bold text-white disabled:bg-slate-200 disabled:text-slate-400">Відкрити</button>}
                <button onClick={() => onDelete(r.id)} className="shrink-0 rounded-full p-1 text-slate-300 hover:text-red-500"><Trash2 className="h-4 w-4" /></button>
              </div>
            );
          })}
        </div>
        <div className="mt-3 rounded-2xl border border-dashed border-slate-200 p-3">
          <input value={label} onChange={(e) => setLabel(e.target.value)} placeholder="Напр. улюблений снек, серія серіалу…" className="w-full rounded-lg border border-slate-300 px-3 py-2 text-sm focus:border-pink-400 focus:outline-none" />
          <div className="mt-2 flex items-center gap-2">
            <span className="text-xs text-slate-400">Ціна:</span>
            <input type="number" min={50} step={50} value={cost} onChange={(e) => setCost(Math.max(50, +e.target.value || 50))} className="w-24 rounded-lg border border-slate-300 px-2 py-1.5 text-right text-sm" />
            <span className="text-xs text-slate-400">XP</span>
            <button onClick={() => { if (label.trim()) { onAdd(label.trim(), cost); setLabel(""); } }} className="ml-auto rounded-full bg-slate-800 px-4 py-1.5 text-sm font-semibold text-white hover:bg-slate-900">Додати</button>
          </div>
        </div>
      </div>
    </div>
  );
}

function DayRecap({ tasks, doc, xpToday, xp, streak, onClose, onCarryOver, carryCount }) {
  const doneTasks = tasks.filter((t) => doc.tasks?.[t.id]);
  const lp = levelProgress(xp);
  const totalMin = doneTasks.reduce((s, t) => s + (t.estMin || 0), 0);
  return (
    <div className="fixed inset-0 z-50 flex items-end justify-center bg-slate-900/40 p-0 backdrop-blur-sm sm:items-center sm:p-4" onClick={onClose}>
      <div className="max-h-[88vh] w-full max-w-md overflow-y-auto rounded-t-3xl bg-white p-6 text-center shadow-xl sm:rounded-3xl" onClick={(e) => e.stopPropagation()}>
        <div className="text-4xl">🌙</div>
        <h3 className="mt-2 text-xl font-extrabold text-slate-900">Підсумок дня</h3>
        <div className="my-4 grid grid-cols-3 gap-2">
          <div className="rounded-2xl bg-pink-50 p-3"><div className="text-2xl font-extrabold text-pink-500">{doneTasks.length}</div><div className="text-[11px] text-slate-400">закрито</div></div>
          <div className="rounded-2xl bg-fuchsia-50 p-3"><div className="text-2xl font-extrabold text-fuchsia-500">+{xpToday}</div><div className="text-[11px] text-slate-400">XP сьогодні</div></div>
          <div className="rounded-2xl bg-orange-50 p-3"><div className="text-2xl font-extrabold text-orange-500">{streak.current}🔥</div><div className="text-[11px] text-slate-400">днів поспіль</div></div>
        </div>
        <div className="mb-1 text-sm font-semibold text-slate-600">Рівень {lp.lvl}{totalMin ? ` · ≈ ${fmtEst(totalMin)} роботи` : ""}</div>
        {doneTasks.length > 0 ? (
          <div className="mt-3 rounded-2xl bg-slate-50 p-3 text-left">
            <div className="mb-1.5 text-xs font-bold uppercase tracking-wide text-slate-400">Що зроблено</div>
            <div className="space-y-1">{doneTasks.map((t) => <div key={t.id} className="flex items-center gap-2 text-sm text-slate-700"><span>{t.emoji || "✅"}</span><span className="truncate">{t.title}</span></div>)}</div>
          </div>
        ) : <p className="mt-3 text-sm text-slate-400">Сьогодні нічого не закрито — і це теж нормально. Завтра новий день 💛</p>}
        {carryCount > 0 && <button onClick={onCarryOver} className="mt-4 w-full rounded-2xl bg-white py-3 text-sm font-semibold text-slate-500 ring-1 ring-slate-200 hover:bg-slate-50">Перенести {carryCount} незавершену(і) на завтра — без провини 💛</button>}
        <button onClick={onClose} className="mt-2 w-full rounded-2xl bg-pink-500 py-3 font-bold text-white hover:bg-pink-600">Гарного вечора ✨</button>
      </div>
    </div>
  );
}

function LevelUp({ level, onClose }) {
  return (
    <div className="fixed inset-0 z-[60] flex items-center justify-center bg-gradient-to-b from-fuchsia-500 to-pink-600 p-6 text-center text-white" onClick={onClose}>
      <div>
        <div className="text-6xl">🎉</div>
        <div className="mt-3 text-sm font-bold uppercase tracking-widest text-white/80">Новий рівень</div>
        <h2 className="text-5xl font-black">Рівень {level}</h2>
        <p className="mt-2 text-white/80">Так тримати! Кожна закрита справа — це крок.</p>
        <button onClick={onClose} className="mt-6 rounded-2xl bg-white px-8 py-3 font-bold text-pink-600">Далі</button>
      </div>
    </div>
  );
}

function RoutineSection() {
  const [loading, setLoading] = useState(true);
  const [rview, setRview] = useState("today"); // today | stats
  const [tasks, setTasks] = useState([]);
  const [categories, setCategories] = useState([]);
  const [completions, setCompletions] = useState({});
  const [cindex, setCindex] = useState([]);
  const [streakMeta, setStreakMeta] = useState({ best: 0, lastCelebrated: "" });
  const [moods, setMoods] = useState({});
  const [selDate, setSelDate] = useState(dateKey(Date.now()));
  const [selCat, setSelCat] = useState("all");
  const [taskEditor, setTaskEditor] = useState(null); // {task}
  const [detailId, setDetailId] = useState(null);
  const [moodOpen, setMoodOpen] = useState(false);
  const [bannerDismissed, setBannerDismissed] = useState(false);
  const [celebration, setCelebration] = useState(null);
  const [catManager, setCatManager] = useState(false);
  const [menuOpen, setMenuOpen] = useState(false);
  const [timer, setTimer] = useState(null); // {taskId, elapsed, target, running}
  const [toast, setToast] = useState(null);
  const [xp, setXp] = useState(0);
  const [rewards, setRewards] = useState([]);
  const [wheelOpen, setWheelOpen] = useState(false);
  const [focusOpen, setFocusOpen] = useState(false);
  const [rewardsOpen, setRewardsOpen] = useState(false);
  const [recapOpen, setRecapOpen] = useState(false);
  const [levelUp, setLevelUp] = useState(null);
  const [energyFilter, setEnergyFilter] = useState(null); // null | "quick" | "low"
  const [pomo, setPomo] = useState(null); // { task, mode: "2min" | "pomodoro" }
  const [pomoSettings, setPomoSettings] = useState({ work: 25, break: 5 });
  const timerRef = useRef(null);

  const today = dateKey(Date.now());
  const isToday = selDate === today;

  const flash = useCallback((m) => { setToast(m); window.clearTimeout(flash._t); flash._t = window.setTimeout(() => setToast(null), 2000); }, []);

  const reload = useCallback(async () => {
    let d = await loadRoutineData();
    if (d.tasks === null && d.categories === null && d.cindex.length === 0) {
      const seed = seededRoutine();
      await store.set(RKEYS.categories, seed.categories);
      await store.set(RKEYS.tasks, seed.tasks);
      await store.set(RKEYS.seeded, true);
      d = await loadRoutineData();
    }
    // one-time: add the "shave" task on Tue + Fri
    if (!(await store.get("routine:mig:shave", false))) {
      const list = [...(d.tasks || []), { id: ruid("t"), emoji: "🪒", title: "Поголитися", note: "", time: null, color: "teal", categoryId: "", repeat: { type: "weekdays", days: [2, 5] }, goal: { type: "off" }, subtasks: [], created: Date.now(), date: dateKey(Date.now()) }];
      await store.set(RKEYS.tasks, list); await store.set("routine:mig:shave", true);
      d = { ...d, tasks: list };
    }
    setTasks(d.tasks || []); setCategories(d.categories || []); setCompletions(d.completions);
    setCindex(d.cindex); setStreakMeta(d.streak || { best: 0, lastCelebrated: "" }); setMoods(d.moods || {});
    setXp((d.xp && d.xp.xp) || 0); setRewards(d.rewards || []);
    setPomoSettings(d.pomo || { work: 25, break: 5 });
    setLoading(false);
  }, []);

  useEffect(() => {
    reload();
    const onReset = () => reload();
    window.addEventListener("routine-reset", onReset);
    return () => window.removeEventListener("routine-reset", onReset);
  }, [reload]);

  /* ---- persistence ---- */
  const saveTasks = useCallback(async (next) => { setTasks(next); await store.set(RKEYS.tasks, next); }, []);
  const saveCategories = useCallback(async (next) => { setCategories(next); await store.set(RKEYS.categories, next); }, []);
  const saveMoods = useCallback(async (next) => { setMoods(next); await store.set(RKEYS.mood, next); }, []);
  const completionsRef = useRef(completions);
  completionsRef.current = completions;
  const persistCompletion = useCallback(async (date, doc) => {
    setCompletions((m) => ({ ...m, [date]: doc }));
    completionsRef.current = { ...completionsRef.current, [date]: doc };
    await store.set(cKey(date), doc);
    setCindex((prev) => { if (prev.includes(date)) return prev; const next = [...prev, date].sort(); store.set(RKEYS.cindex, next); return next; });
  }, []);

  // Award XP for any freshly-completed tasks + newly-satisfied challenges (idempotent, never removes).
  const settleDoc = useCallback((doc, dateStr) => {
    const occurring = tasks.filter((t) => taskOccursOn(t, dateStr));
    doc.xpAwarded = { ...(doc.xpAwarded || {}) };
    let delta = 0;
    for (const t of occurring) {
      if (doc.tasks?.[t.id] && doc.xpAwarded[t.id] == null) { const amt = xpForTask(t); doc.xpAwarded[t.id] = amt; delta += amt; }
    }
    const ctx = {
      tasks: occurring,
      doneCount: occurring.filter((t) => doc.tasks?.[t.id]).length,
      isDone: (t) => !!doc.tasks?.[t.id],
      actualMin: (t) => (doc.goal?.[t.id] != null ? Math.round(doc.goal[t.id] / 60) : null),
    };
    doc.ch = { dismissed: { ...(doc.ch?.dismissed || {}) }, awarded: { ...(doc.ch?.awarded || {}) } };
    for (const c of pickChallenges(dateStr)) {
      if (!doc.ch.awarded[c.id] && c.check(ctx)) { doc.ch.awarded[c.id] = c.xp; delta += c.xp; }
    }
    return delta;
  }, [tasks]);

  // Persist a completion doc AND grant any earned XP, with a level-up moment.
  const commitDoc = useCallback(async (doc) => {
    const delta = settleDoc(doc, selDate);
    await persistCompletion(selDate, doc);
    if (delta > 0) {
      setXp((prev) => { const nx = prev + delta; store.set(RKEYS.xp, { xp: nx }); if (levelFromXp(nx) > levelFromXp(prev)) setLevelUp(levelFromXp(nx)); return nx; });
      flash(`+${delta} XP ✨`);
    }
  }, [settleDoc, selDate, persistCompletion, flash]);

  const maybeCelebrate = useCallback((nextCompletions) => {
    if (selDate !== today) return;
    const after = computeTaskStreak(nextCompletions, today).current;
    const before = computeTaskStreak(completions, today).current;
    if (before === 0 && after === 1 && streakMeta.lastCelebrated !== today) {
      setCelebration({ streak: after });
      const nextMeta = { ...streakMeta, lastCelebrated: today, best: Math.max(streakMeta.best || 0, after) };
      setStreakMeta(nextMeta); store.set(RKEYS.streak, nextMeta);
    }
  }, [selDate, today, completions, streakMeta]);

  const toggleTask = useCallback(async (taskId) => {
    const doc = { ...(completionsRef.current[selDate] || {}) };
    doc.tasks = { ...(doc.tasks || {}) };
    const willComplete = !doc.tasks[taskId];
    if (willComplete) doc.tasks[taskId] = true; else delete doc.tasks[taskId];
    const nextCompletions = { ...completions, [selDate]: doc };
    await commitDoc(doc);
    if (willComplete) maybeCelebrate(nextCompletions);
  }, [completions, selDate, commitDoc, maybeCelebrate]);

  const toggleSubtask = useCallback(async (taskId, subId) => {
    const task = tasks.find((t) => t.id === taskId);
    const doc = { ...(completionsRef.current[selDate] || {}) };
    doc.subtasks = { ...(doc.subtasks || {}) };
    doc.subtasks[taskId] = { ...(doc.subtasks[taskId] || {}) };
    if (doc.subtasks[taskId][subId]) delete doc.subtasks[taskId][subId]; else doc.subtasks[taskId][subId] = true;
    // auto-complete parent when all subtasks done
    doc.tasks = { ...(doc.tasks || {}) };
    const total = (task?.subtasks || []).length;
    const done = Object.keys(doc.subtasks[taskId]).length;
    let willComplete = false;
    if (total > 0 && done >= total) { if (!doc.tasks[taskId]) willComplete = true; doc.tasks[taskId] = true; }
    else if (total > 0) delete doc.tasks[taskId];
    const nextCompletions = { ...completions, [selDate]: doc };
    await commitDoc(doc);
    if (willComplete) maybeCelebrate(nextCompletions);
  }, [tasks, completions, selDate, commitDoc, maybeCelebrate]);

  const setMood = useCallback(async (score) => { await saveMoods({ ...moods, [today]: score }); setMoodOpen(false); flash("Mood saved"); }, [moods, today, saveMoods, flash]);

  const saveTask = useCallback(async (meta, id) => {
    if (id) await saveTasks(tasks.map((t) => (t.id === id ? { ...t, ...meta } : t)));
    else await saveTasks([...tasks, { id: ruid("t"), created: Date.now(), ...meta }]);
    flash(id ? "Task saved" : "Task added");
  }, [tasks, saveTasks, flash]);
  const deleteTask = useCallback(async (id) => { await saveTasks(tasks.filter((t) => t.id !== id)); flash("Task deleted"); }, [tasks, saveTasks, flash]);

  const addCategory = useCallback(async (name, color) => { await saveCategories([...categories, { id: ruid("cat"), name: name.trim(), color }]); }, [categories, saveCategories]);
  const renameCategory = useCallback(async (id, name) => { await saveCategories(categories.map((c) => (c.id === id ? { ...c, name } : c))); }, [categories, saveCategories]);
  const recolorCategory = useCallback(async (id, color) => { await saveCategories(categories.map((c) => (c.id === id ? { ...c, color } : c))); }, [categories, saveCategories]);
  const deleteCategory = useCallback(async (id) => {
    await saveCategories(categories.filter((c) => c.id !== id));
    await saveTasks(tasks.map((t) => (t.categoryId === id ? { ...t, categoryId: "" } : t)));
    if (selCat === id) setSelCat("all");
  }, [categories, tasks, saveCategories, saveTasks, selCat]);

  /* ---- timed-goal timer ---- */
  const stopTimer = useCallback(async (markProgress = true) => {
    if (timerRef.current) { window.clearInterval(timerRef.current); timerRef.current = null; }
    setTimer((cur) => {
      if (cur && markProgress) {
        const doc = { ...(completionsRef.current[selDate] || {}) };
        doc.goal = { ...(doc.goal || {}) };
        doc.goal[cur.taskId] = cur.elapsed;
        doc.tasks = { ...(doc.tasks || {}) };
        let willComplete = false;
        if (cur.elapsed >= cur.target && !doc.tasks[cur.taskId]) { doc.tasks[cur.taskId] = true; willComplete = true; }
        commitDoc(doc);
        if (willComplete) maybeCelebrate({ ...completions, [selDate]: doc });
      }
      return null;
    });
  }, [completions, selDate, commitDoc, maybeCelebrate]);

  const startTimer = useCallback((task) => {
    if (timerRef.current) window.clearInterval(timerRef.current);
    const target = (task.goal?.minutes || 1) * 60;
    const startElapsed = completions[selDate]?.goal?.[task.id] || 0;
    setTimer({ taskId: task.id, elapsed: startElapsed, target, running: true });
    timerRef.current = window.setInterval(() => {
      setTimer((cur) => {
        if (!cur) return cur;
        const elapsed = cur.elapsed + 1;
        if (elapsed >= cur.target) {
          // finish
          window.clearInterval(timerRef.current); timerRef.current = null;
          const doc = { ...(completionsRef.current[selDate] || {}) };
          doc.goal = { ...(doc.goal || {}), [cur.taskId]: elapsed };
          doc.tasks = { ...(doc.tasks || {}), [cur.taskId]: true };
          commitDoc(doc);
          maybeCelebrate({ ...completions, [selDate]: doc });
          flash("Ціль виконано! 🎉");
          return null;
        }
        return { ...cur, elapsed };
      });
    }, 1000);
  }, [completions, selDate, commitDoc, maybeCelebrate, flash]);

  useEffect(() => () => { if (timerRef.current) window.clearInterval(timerRef.current); }, []);

  /* ---- derived ---- */
  const doc = completions[selDate] || {};
  const streak = useMemo(() => computeTaskStreak(completions, today), [completions, today]);
  const dayTasks = useMemo(() => {
    const list = tasks.filter((t) => taskOccursOn(t, selDate) && (selCat === "all" || t.categoryId === selCat));
    return list.slice().sort((a, b) => {
      const at = a.time || "99:99", bt = b.time || "99:98"; // Anytime after timed
      return at.localeCompare(bt);
    });
  }, [tasks, selDate, selCat]);
  const dayDone = dayTasks.filter((t) => doc.tasks?.[t.id]).length;
  const dayPct = dayTasks.length ? dayDone / dayTasks.length : 0;

  /* ---- gamification derived ---- */
  const occurringAll = useMemo(() => tasks.filter((t) => taskOccursOn(t, selDate)), [tasks, selDate]);
  const shownTasks = useMemo(() => {
    if (energyFilter === "quick") return dayTasks.filter((t) => t.estMin && t.estMin <= 5);
    if (energyFilter === "low") return dayTasks.filter((t) => t.energy === "low");
    return dayTasks;
  }, [dayTasks, energyFilter]);
  const notDone = dayTasks.filter((t) => !doc.tasks?.[t.id]);
  const dayTotalMin = occurringAll.reduce((s, t) => s + (t.estMin || 0), 0);
  const challenges = useMemo(() => pickChallenges(selDate), [selDate]);
  const chCtx = { tasks: occurringAll, doneCount: occurringAll.filter((t) => doc.tasks?.[t.id]).length, isDone: (t) => !!doc.tasks?.[t.id], actualMin: (t) => (doc.goal?.[t.id] != null ? Math.round(doc.goal[t.id] / 60) : null) };
  const xpToday = Object.values(doc.xpAwarded || {}).reduce((s, v) => s + v, 0) + Object.values(doc.ch?.awarded || {}).reduce((s, v) => s + v, 0);
  const carryList = occurringAll.filter((t) => !doc.tasks?.[t.id] && (t.repeat?.type || "off") === "off");

  const saveRewards = useCallback(async (next) => { setRewards(next); await store.set(RKEYS.rewards, next); }, []);
  const addReward = (label, cost) => saveRewards([...rewards, { id: ruid("rw"), label, cost, unlockedAt: null }]);
  const unlockReward = (id) => { saveRewards(rewards.map((r) => (r.id === id ? { ...r, unlockedAt: Date.now() } : r))); flash("Нагороду відкрито! 🎉"); };
  const deleteReward = (id) => saveRewards(rewards.filter((r) => r.id !== id));
  const dismissChallenge = useCallback(async (cid) => { const d = { ...(completionsRef.current[selDate] || {}) }; d.ch = { dismissed: { ...(d.ch?.dismissed || {}), [cid]: true }, awarded: { ...(d.ch?.awarded || {}) } }; await persistCompletion(selDate, d); }, [completions, selDate, persistCompletion]);
  const carryOver = useCallback(async () => { const tomorrow = dateKey(Date.now() + 86400000); const ids = new Set(carryList.map((c) => c.id)); await saveTasks(tasks.map((t) => (ids.has(t.id) ? { ...t, date: tomorrow } : t))); setRecapOpen(false); flash("Перенесено на завтра 💛"); }, [tasks, carryList, saveTasks, flash]);
  const focusDone = useCallback(async (task) => { const d = { ...(completionsRef.current[selDate] || {}) }; d.tasks = { ...(d.tasks || {}), [task.id]: true }; await commitDoc(d); maybeCelebrate({ ...completions, [selDate]: d }); }, [completions, selDate, commitDoc, maybeCelebrate]);
  const logPomo = useCallback((taskId, secs) => { const d = { ...(completionsRef.current[selDate] || {}) }; d.pomo = { ...(d.pomo || {}) }; d.pomo[taskId] = (d.pomo[taskId] || 0) + secs; persistCompletion(selDate, d); flash(`Записано ${Math.round(secs / 60)} хв роботи 🍅`); }, [selDate, persistCompletion, flash]);
  const savePomoSettings = useCallback((s) => { setPomoSettings(s); store.set(RKEYS.pomo, s); }, []);

  if (loading) {
    return <div className="flex flex-1 items-center justify-center text-pink-400"><div className="flex flex-col items-center gap-3"><Sun className="h-8 w-8 animate-pulse" /><span className="text-sm">Loading your routine…</span></div></div>;
  }

  const detailTask = detailId ? tasks.find((t) => t.id === detailId) : null;

  return (
    <div className="min-h-screen flex-1 bg-gradient-to-b from-red-50 to-pink-50/40" style={{ fontFamily: "inherit" }}>
      {rview === "today" && (
        <div className="mx-auto w-full max-w-2xl px-4 pb-28 pt-5">
          {/* header */}
          <div className="mb-4 flex items-center justify-between">
            <div>
              <h1 className="text-2xl font-extrabold text-slate-900">{isToday ? "Today" : prettyDate(selDate).split(",")[0]}</h1>
              <p className="text-xs font-medium text-slate-400">{prettyDate(selDate)}</p>
            </div>
            <div className="flex items-center gap-2">
              <div className="flex items-center gap-1 rounded-full bg-white px-3 py-1.5 shadow-sm ring-1 ring-red-100">
                <span className="text-lg">🔥</span>
                <span className="text-sm font-bold tabular-nums text-orange-500">{streak.current}</span>
              </div>
              <div className="relative">
                <button onClick={() => setMenuOpen((v) => !v)} className="grid h-9 w-9 place-items-center rounded-full bg-white text-slate-500 shadow-sm ring-1 ring-red-100 hover:text-slate-700"><Menu className="h-4 w-4" /></button>
                {menuOpen && (<>
                  <div className="fixed inset-0 z-10" onClick={() => setMenuOpen(false)} />
                  <div className="absolute right-0 top-11 z-20 w-48 overflow-hidden rounded-xl border border-slate-200 bg-white py-1 shadow-lg">
                    <button onClick={() => { setRview("wellbeing"); setMenuOpen(false); }} className="flex w-full items-center gap-2 px-3 py-2 text-left text-sm text-slate-700 hover:bg-slate-50"><Smile className="h-4 w-4 text-slate-400" /> Настрій і вдячність</button>
                    <button onClick={() => { setRview("meds"); setMenuOpen(false); }} className="flex w-full items-center gap-2 px-3 py-2 text-left text-sm text-slate-700 hover:bg-slate-50"><HeartPulse className="h-4 w-4 text-slate-400" /> Ліки й самопочуття</button>
                    <button onClick={() => { setRview("stats"); setMenuOpen(false); }} className="flex w-full items-center gap-2 px-3 py-2 text-left text-sm text-slate-700 hover:bg-slate-50"><BarChart3 className="h-4 w-4 text-slate-400" /> Stats & profile</button>
                    <button onClick={() => { setCatManager(true); setMenuOpen(false); }} className="flex w-full items-center gap-2 px-3 py-2 text-left text-sm text-slate-700 hover:bg-slate-50"><Tag className="h-4 w-4 text-slate-400" /> Manage categories</button>
                  </div>
                </>)}
              </div>
            </div>
          </div>

          {/* week strip */}
          <WeekStrip selDate={selDate} today={today} completions={completions} tasks={tasks} onPick={setSelDate} />

          {/* XP / level */}
          <GamifyBar xp={xp} onRewards={() => setRewardsOpen(true)} />

          {/* daily challenges */}
          {isToday && <ChallengesCard challenges={challenges} ctx={chCtx} chDoc={doc.ch || {}} onDismiss={dismissChallenge} />}

          {/* quick actions */}
          {isToday && dayTasks.length > 0 && (
            <div className="mt-4 flex flex-wrap items-center gap-2">
              <button onClick={() => notDone.length ? setWheelOpen(true) : flash("Усе на сьогодні закрито 🎉")} className="inline-flex items-center gap-1 rounded-full bg-white px-3.5 py-2 text-sm font-bold text-slate-700 shadow-sm ring-1 ring-red-100 hover:ring-red-200">🎡 Колесо</button>
              <button onClick={() => setFocusOpen(true)} className="inline-flex items-center gap-1 rounded-full bg-white px-3.5 py-2 text-sm font-bold text-slate-700 shadow-sm ring-1 ring-red-100 hover:ring-red-200">🎯 Зараз</button>
              <button onClick={() => setRecapOpen(true)} className="inline-flex items-center gap-1 rounded-full bg-white px-3.5 py-2 text-sm font-bold text-slate-700 shadow-sm ring-1 ring-red-100 hover:ring-red-200">🌙 Підсумок</button>
              {dayTotalMin > 0 && <span className="ml-auto text-xs font-semibold text-slate-400">Сьогодні ≈ {fmtEst(dayTotalMin)}</span>}
            </div>
          )}
          {isToday && dayTasks.length > 0 && (
            <div className="mt-2 flex flex-wrap gap-2">
              {[["quick", "⚡ Швидкі перемоги"], ["low", "🟢 Мало енергії"]].map(([k, label]) => (
                <button key={k} onClick={() => setEnergyFilter((f) => (f === k ? null : k))} className={`rounded-full px-3 py-1.5 text-xs font-semibold ring-1 transition ${energyFilter === k ? "bg-pink-500 text-white ring-pink-500" : "bg-white text-slate-500 ring-slate-200 hover:ring-pink-200"}`}>{label}</button>
              ))}
            </div>
          )}

          {/* mood banner */}
          {isToday && !bannerDismissed && moods[today] == null && (
            <div className="mt-4 flex items-center gap-3 rounded-2xl bg-white/80 p-3 shadow-sm ring-1 ring-red-100">
              <span className="text-2xl">🌸</span>
              <div className="flex-1">
                <div className="text-sm font-semibold text-slate-800">How is your day?</div>
                <div className="text-xs text-slate-400">Take a second to check in.</div>
              </div>
              <button onClick={() => setMoodOpen(true)} className="rounded-full bg-pink-500 px-3 py-1.5 text-xs font-semibold text-white hover:bg-pink-600">Check in</button>
              <button onClick={() => setBannerDismissed(true)} className="rounded-full p-1 text-slate-300 hover:text-slate-500"><X className="h-4 w-4" /></button>
            </div>
          )}
          {isToday && moods[today] != null && (
            <button onClick={() => setMoodOpen(true)} className="mt-4 flex w-full items-center gap-3 rounded-2xl bg-white/80 p-3 text-left shadow-sm ring-1 ring-red-100">
              <span className="text-2xl">{MOODS.find((m) => m.score === moods[today])?.emoji}</span>
              <div className="flex-1"><div className="text-sm font-semibold text-slate-800">Feeling {MOODS.find((m) => m.score === moods[today])?.label.toLowerCase()}</div><div className="text-xs text-slate-400">Tap to change</div></div>
            </button>
          )}

          {/* category chips */}
          <div className="mt-4 flex flex-wrap items-center gap-2">
            <Chip active={selCat === "all"} onClick={() => setSelCat("all")}>All</Chip>
            {categories.map((c) => (
              <Chip key={c.id} active={selCat === c.id} color={c.color} onClick={() => setSelCat(c.id)}>{c.name}</Chip>
            ))}
            <button onClick={() => setCatManager(true)} className="grid h-7 w-7 place-items-center rounded-full bg-white text-slate-400 shadow-sm ring-1 ring-slate-200 hover:text-slate-600"><Plus className="h-4 w-4" /></button>
          </div>

          {/* progress line */}
          {dayTasks.length > 0 && (
            <div className="mt-4 flex items-center gap-3 rounded-2xl bg-white/70 p-3 ring-1 ring-red-100">
              <ProgressRing pct={dayPct} size={44} stroke={5}><span className="text-[11px] font-bold text-pink-600">{Math.round(dayPct * 100)}%</span></ProgressRing>
              <div className="text-sm font-medium text-slate-600">{dayDone} of {dayTasks.length} done{dayPct >= 1 ? " — all clear! 🎉" : ""}</div>
            </div>
          )}

          {/* movement (block 5) */}
          {isToday && <MovementCard flash={flash} />}

          {/* task list */}
          <div className="mt-4 space-y-2.5">
            {dayTasks.length === 0 ? (
              <div className="rounded-2xl bg-white/70 py-12 text-center text-sm text-slate-400">Nothing here yet — tap the + to add a task.</div>
            ) : shownTasks.length === 0 ? (
              <div className="rounded-2xl bg-white/70 py-10 text-center text-sm text-slate-400">Нема справ під цей фільтр. <button onClick={() => setEnergyFilter(null)} className="font-semibold text-pink-500 underline">Показати всі</button></div>
            ) : shownTasks.map((t) => (
              <TaskCard key={t.id} task={t} done={!!doc.tasks?.[t.id]} doc={doc}
                timer={timer?.taskId === t.id ? timer : null}
                onToggle={() => toggleTask(t.id)} onOpen={() => setDetailId(t.id)} onStartTimer={() => startTimer(t)} onStopTimer={() => stopTimer(true)} onTwoMin={() => setPomo({ task: t, mode: "2min" })} />
            ))}
          </div>
        </div>
      )}

      {rview === "stats" && (
        <RoutineStats tasks={tasks} completions={completions} moods={moods} streak={streak} best={Math.max(streakMeta.best || 0, streak.best)} onBack={() => setRview("today")} />
      )}
      {rview === "wellbeing" && <WellbeingView onExit={() => setRview("today")} moods={moods} onMood={setMood} />}
      {rview === "meds" && <MedsView onExit={() => setRview("today")} />}

      {/* floating add */}
      {rview === "today" && (
        <button onClick={() => setTaskEditor({ task: null })} className="fixed bottom-6 right-6 z-30 grid h-14 w-14 place-items-center rounded-full bg-pink-500 text-white shadow-lg shadow-pink-500/30 transition hover:bg-pink-600 hover:scale-105">
          <Plus className="h-7 w-7" />
        </button>
      )}

      {/* mood picker */}
      {moodOpen && (
        <div className="fixed inset-0 z-50 flex items-end justify-center bg-slate-900/30 p-0 backdrop-blur-sm sm:items-center sm:p-4" onClick={() => setMoodOpen(false)}>
          <div className="w-full max-w-sm rounded-t-3xl bg-white p-6 text-center shadow-xl sm:rounded-3xl" onClick={(e) => e.stopPropagation()}>
            <h3 className="mb-1 text-lg font-bold text-slate-900">How is your day?</h3>
            <p className="mb-4 text-sm text-slate-400">{prettyDate(today)}</p>
            <div className="flex justify-between">
              {MOODS.map((m) => (
                <button key={m.score} onClick={() => setMood(m.score)} className={`flex flex-col items-center gap-1 rounded-2xl px-2 py-2 transition hover:scale-110 ${moods[today] === m.score ? "bg-pink-50 ring-2 ring-pink-300" : ""}`}>
                  <span className="text-3xl">{m.emoji}</span><span className="text-[10px] font-medium text-slate-400">{m.label}</span>
                </button>
              ))}
            </div>
          </div>
        </div>
      )}

      {/* task detail */}
      {detailTask && (
        <TaskDetail task={detailTask} doc={completions[selDate] || {}} category={categories.find((c) => c.id === detailTask.categoryId)}
          timer={timer?.taskId === detailTask.id ? timer : null}
          onClose={() => setDetailId(null)}
          onEdit={() => { setTaskEditor({ task: detailTask }); setDetailId(null); }}
          onDelete={() => { if (confirm("Delete this task?")) { deleteTask(detailTask.id); setDetailId(null); } }}
          onToggle={() => toggleTask(detailTask.id)}
          onToggleSub={(sid) => toggleSubtask(detailTask.id, sid)}
          onStartTimer={() => startTimer(detailTask)} onStopTimer={() => stopTimer(true)}
          onTwoMin={() => { setPomo({ task: detailTask, mode: "2min" }); setDetailId(null); }}
          onPomodoro={() => { setPomo({ task: detailTask, mode: "pomodoro" }); setDetailId(null); }} />
      )}

      {pomo && <RoutinePomodoro task={pomo.task} mode={pomo.mode} settings={pomoSettings} onClose={() => setPomo(null)} onLog={logPomo} onSaveSettings={savePomoSettings} />}

      {/* task editor */}
      {taskEditor && (
        <TaskEditor task={taskEditor.task} categories={categories} defaultDate={selDate}
          onClose={() => setTaskEditor(null)}
          onSave={async (meta) => { await saveTask(meta, taskEditor.task?.id); setTaskEditor(null); }} />
      )}

      {/* category manager */}
      {catManager && (
        <CategoryManager categories={categories} onClose={() => setCatManager(false)}
          onAdd={addCategory} onRename={renameCategory} onRecolor={recolorCategory} onDelete={deleteCategory} />
      )}

      {/* streak born celebration */}
      {celebration && <StreakBorn streak={celebration.streak} onClose={() => setCelebration(null)} />}

      {/* gamification modals */}
      {wheelOpen && <WheelSpin tasks={notDone} onClose={() => setWheelOpen(false)} onPick={(t) => { setWheelOpen(false); setDetailId(t.id); }} />}
      {focusOpen && <FocusMode tasks={notDone} doc={doc} onClose={() => setFocusOpen(false)} onDone={focusDone} onSkip={() => {}} />}
      {rewardsOpen && <RewardsPanel xp={xp} rewards={rewards} onClose={() => setRewardsOpen(false)} onAdd={addReward} onUnlock={unlockReward} onDelete={deleteReward} />}
      {recapOpen && <DayRecap tasks={occurringAll} doc={doc} xpToday={xpToday} xp={xp} streak={streak} onClose={() => setRecapOpen(false)} onCarryOver={carryOver} carryCount={carryList.length} />}
      {levelUp && <LevelUp level={levelUp} onClose={() => setLevelUp(null)} />}

      {toast && <div className="fixed bottom-24 left-1/2 z-40 -translate-x-1/2 rounded-full bg-slate-900 px-4 py-2 text-sm text-white shadow-lg">{toast}</div>}
    </div>
  );
}

function Chip({ active, color, onClick, children }) {
  const p = color ? getPastel(color) : null;
  return (
    <button onClick={onClick}
      className={`rounded-full px-3.5 py-1.5 text-sm font-semibold transition ${active ? "text-white shadow-sm" : "bg-white text-slate-500 ring-1 ring-slate-200 hover:bg-slate-50"}`}
      style={active ? { backgroundColor: p ? p.dot : "#ec4899" } : {}}>
      {children}
    </button>
  );
}

function WeekStrip({ selDate, today, completions, tasks, onPick }) {
  const days = weekDaysOf(selDate);
  return (
    <div className="flex justify-between gap-1">
      {days.map((ds) => {
        const d = new Date(ds + "T00:00:00");
        const sel = ds === selDate;
        const isToday = ds === today;
        const prog = dayTaskProgress(tasks, completions, ds);
        return (
          <button key={ds} onClick={() => onPick(ds)} className={`flex flex-1 flex-col items-center gap-1 rounded-2xl py-2 transition ${sel ? "bg-pink-500 text-white shadow-md shadow-pink-500/20" : "text-slate-500 hover:bg-white/60"}`}>
            <span className={`text-[10px] font-semibold uppercase ${sel ? "text-white/80" : "text-slate-400"}`}>{WD_LETTER[d.getDay()]}</span>
            <span className={`text-sm font-bold ${isToday && !sel ? "text-pink-500" : ""}`}>{d.getDate()}</span>
            <span className={`h-1.5 w-1.5 rounded-full ${prog.total && prog.pct >= 1 ? (sel ? "bg-white" : "bg-green-400") : prog.done ? (sel ? "bg-white/60" : "bg-pink-300") : "bg-transparent"}`} />
          </button>
        );
      })}
    </div>
  );
}

function TaskCard({ task, done, doc, timer, onToggle, onOpen, onStartTimer, onStopTimer, onTwoMin }) {
  const p = getPastel(task.color);
  const timed = task.goal?.type === "timed";
  const goalSecs = timer ? timer.elapsed : (doc.goal?.[task.id] || 0);
  const subDone = task.subtasks?.length ? Object.keys(doc.subtasks?.[task.id] || {}).length : 0;
  const fmt = (s) => `${Math.floor(s / 60)}:${String(s % 60).padStart(2, "0")}`;
  return (
    <div className="flex items-center gap-3 rounded-2xl p-3.5 shadow-sm transition" style={{ backgroundColor: p.card }}>
      <button onClick={onOpen} className="grid h-16 w-16 shrink-0 place-items-center overflow-hidden rounded-2xl bg-white/70 text-2xl">
        {task.image ? <img src={task.image} alt="" loading="lazy" className="h-full w-full object-cover" /> : (task.emoji || "⭐")}
      </button>
      <button onClick={onOpen} className="min-w-0 flex-1 text-left">
        <div className="text-[11px] font-semibold" style={{ color: p.ink }}>{task.time || "Anytime"}</div>
        <div className={`truncate font-bold ${done ? "text-slate-400 line-through" : "text-slate-800"}`}>{task.title}</div>
        <div className="mt-0.5 flex flex-wrap items-center gap-2 text-[11px]" style={{ color: p.ink }}>
          {timed && <span className="font-semibold">Goal: {Math.floor(goalSecs / 60)}/{task.goal.minutes} min</span>}
          {task.subtasks?.length > 0 && <span className="font-semibold">{subDone}/{task.subtasks.length} steps</span>}
          {task.estMin > 0 && <span className="font-semibold opacity-80">≈ {fmtEst(task.estMin)}</span>}
          {task.energy && ENERGY[task.energy] && <span className="opacity-80">{ENERGY[task.energy].emoji}</span>}
          {task.reminder && <span className="inline-flex items-center gap-0.5 opacity-70"><Clock className="h-3 w-3" />{task.reminder}</span>}
          {done && task.estMin > 0 && goalSecs > 0 && <span className="rounded-full bg-white/70 px-1.5 font-medium opacity-90">оцінка {task.estMin}хв · факт {Math.max(1, Math.round(goalSecs / 60))}хв</span>}
        </div>
      </button>
      {!done && !timer && onTwoMin && <button onClick={onTwoMin} title="Почни з 2 хвилин" className="grid h-9 shrink-0 place-items-center rounded-full bg-white/80 px-2.5 text-[11px] font-bold" style={{ color: p.ink }}>▶ 2хв</button>}
      {timed && !done && (
        timer ? (
          <button onClick={onStopTimer} className="grid h-9 w-16 shrink-0 place-items-center rounded-full bg-white/80 text-xs font-bold tabular-nums" style={{ color: p.ink }}>{fmt(timer.target - timer.elapsed)}</button>
        ) : (
          <button onClick={onStartTimer} className="grid h-9 w-9 shrink-0 place-items-center rounded-full bg-white/80" style={{ color: p.ink }} title="Start timer"><Play className="h-4 w-4" /></button>
        )
      )}
      <button onClick={onToggle} title="Done" className={`grid h-8 w-8 shrink-0 place-items-center rounded-full border-2 transition-all ${done ? "scale-100 border-transparent text-white" : "border-current bg-white/40"}`} style={done ? { backgroundColor: p.dot } : { color: p.dot }}>
        <Check className={`h-4 w-4 transition-transform ${done ? "scale-100" : "scale-0"}`} />
      </button>
    </div>
  );
}

function TaskDetail({ task, doc, category, timer, onClose, onEdit, onDelete, onToggle, onToggleSub, onStartTimer, onStopTimer, onTwoMin, onPomodoro }) {
  const p = getPastel(task.color);
  const done = !!doc.tasks?.[task.id];
  const timed = task.goal?.type === "timed";
  const goalSecs = timer ? timer.elapsed : (doc.goal?.[task.id] || 0);
  const pomoSecs = doc.pomo?.[task.id] || 0;
  const noSteps = !(task.subtasks?.length > 0);
  const fmt = (s) => `${Math.floor(s / 60)}:${String(s % 60).padStart(2, "0")}`;
  return (
    <div className="fixed inset-0 z-50 flex items-end justify-center bg-slate-900/40 p-0 backdrop-blur-sm sm:items-center sm:p-4" onClick={onClose}>
      <div className="max-h-[92vh] w-full max-w-md overflow-y-auto rounded-t-3xl bg-white p-5 shadow-xl sm:rounded-3xl" onClick={(e) => e.stopPropagation()}>
        <div className="mb-3 flex items-start justify-between">
          <div className="flex items-center gap-3">
            <span className="grid h-12 w-12 shrink-0 place-items-center overflow-hidden rounded-2xl text-2xl" style={{ backgroundColor: p.card }}>
              {task.image ? <img src={task.image} alt="" className="h-full w-full object-cover" /> : (task.emoji || "⭐")}
            </span>
            <div>
              <h2 className="text-lg font-bold text-slate-900">{task.title}</h2>
              {category && <span className="rounded-full px-2 py-0.5 text-[11px] font-semibold text-white" style={{ backgroundColor: getPastel(category.color).dot }}>{category.name}</span>}
            </div>
          </div>
          <button onClick={onClose} className="rounded-md p-1 text-slate-400 hover:bg-slate-100"><X className="h-5 w-5" /></button>
        </div>

        {task.image && <img src={task.image} alt="" className="mb-3 max-h-80 w-full rounded-2xl object-contain bg-slate-50" />}
        {task.note && <p className="mb-3 rounded-xl bg-slate-50 px-3 py-2 text-sm text-slate-600">{task.note}</p>}
        <p className="mb-3 text-sm text-slate-500">{repeatWords(task)}</p>

        {/* ADHD low-barrier start */}
        {!done && (
          <div className="mb-3 rounded-2xl bg-red-50/70 p-3">
            {task.twoMin && <div className="mb-2 text-sm text-slate-600"><span className="font-semibold text-slate-800">Версія на 2 хв:</span> {task.twoMin}</div>}
            <div className="flex gap-2">
              <button onClick={onTwoMin} className="flex flex-1 items-center justify-center gap-1.5 rounded-xl bg-pink-500 py-2.5 text-sm font-bold text-white hover:bg-pink-600">🌱 Почни з 2 хв</button>
              <button onClick={onPomodoro} className="flex items-center justify-center gap-1.5 rounded-xl bg-white px-4 py-2.5 text-sm font-bold text-slate-600 ring-1 ring-slate-200 hover:ring-pink-200">🍅 Помодоро</button>
            </div>
            <p className="mt-1.5 text-center text-[11px] text-slate-400">Не мусиш робити все — лише 2 хвилини. Далі саме піде.</p>
          </div>
        )}
        {pomoSecs > 0 && <p className="mb-3 text-xs font-medium text-pink-500">🍅 відпрацьовано {Math.round(pomoSecs / 60)} хв</p>}

        {noSteps && (
          <button onClick={onEdit} className="mb-3 flex w-full items-center gap-2 rounded-2xl border border-dashed border-slate-300 px-3 py-2.5 text-left text-sm text-slate-500 hover:bg-slate-50">
            <span className="text-lg">🧩</span><span className="flex-1">Велика чи розмита задача? Розбий на маленькі кроки — так легше почати.</span><ChevronRight className="h-4 w-4 text-slate-300" />
          </button>
        )}

        {timed && (
          <div className="mb-3 flex items-center gap-3 rounded-2xl p-3" style={{ backgroundColor: p.card }}>
            <ProgressRing pct={Math.min(1, goalSecs / (task.goal.minutes * 60))} size={52} stroke={6} color={p.dot} track="#ffffff88">
              <span className="text-[10px] font-bold" style={{ color: p.ink }}>{Math.floor(goalSecs / 60)}m</span>
            </ProgressRing>
            <div className="flex-1"><div className="font-bold text-slate-800">Goal: {Math.floor(goalSecs / 60)}/{task.goal.minutes} minutes</div><div className="text-xs text-slate-500">{timer ? `Running — ${fmt(timer.target - timer.elapsed)} left` : "Tap play to start a timer"}</div></div>
            {!done && (timer
              ? <button onClick={onStopTimer} className="rounded-full bg-white px-4 py-2 text-sm font-bold" style={{ color: p.ink }}>Pause</button>
              : <button onClick={onStartTimer} className="grid h-11 w-11 place-items-center rounded-full text-white" style={{ backgroundColor: p.dot }}><Play className="h-5 w-5" /></button>)}
          </div>
        )}

        {task.subtasks?.length > 0 && (
          <div className="mb-3">
            <div className="mb-1.5 text-xs font-semibold uppercase tracking-wide text-slate-400">Steps</div>
            <div className="space-y-1">
              {task.subtasks.map((s) => {
                const sd = !!doc.subtasks?.[task.id]?.[s.id];
                return (
                  <button key={s.id} onClick={() => onToggleSub(s.id)} className="flex w-full items-center gap-2 rounded-lg px-1 py-1.5 text-left text-sm hover:bg-slate-50">
                    {sd ? <CheckCircle2 className="h-5 w-5 shrink-0 text-green-500" /> : <Circle className="h-5 w-5 shrink-0 text-slate-300" />}
                    <span className={sd ? "text-slate-400 line-through" : "text-slate-700"}>{s.text}</span>
                  </button>
                );
              })}
            </div>
          </div>
        )}

        <div className="mt-4 flex items-center gap-2">
          <button onClick={onToggle} className={`flex flex-1 items-center justify-center gap-2 rounded-xl py-2.5 font-bold text-white transition ${done ? "bg-slate-400" : ""}`} style={done ? {} : { backgroundColor: p.dot }}>
            {done ? <><RotateCcw className="h-4 w-4" /> Mark not done</> : <><Check className="h-4 w-4" /> Mark done</>}
          </button>
          <button onClick={onEdit} className="rounded-xl border border-slate-300 px-4 py-2.5 text-sm font-semibold text-slate-700 hover:bg-slate-50">Edit Task</button>
          <button onClick={onDelete} className="rounded-xl border border-red-200 px-3 py-2.5 text-red-500 hover:bg-red-50"><Trash2 className="h-4 w-4" /></button>
        </div>
      </div>
    </div>
  );
}

function TaskEditor({ task, categories, defaultDate, onClose, onSave }) {
  const [emoji, setEmoji] = useState(task?.emoji || "⭐");
  const [image, setImage] = useState(task?.image || "");
  const [title, setTitle] = useState(task?.title || "");
  const [note, setNote] = useState(task?.note || "");
  const [color, setColor] = useState(task?.color || "pink");
  const [date, setDate] = useState(task?.date || defaultDate);
  const [repType, setRepType] = useState(task?.repeat?.type || "daily");
  const [days, setDays] = useState(task?.repeat?.days || [1, 2, 3, 4, 5]);
  const [times, setTimes] = useState(task?.repeat?.times || 3);
  const [anytime, setAnytime] = useState(task ? !task.time : false);
  const [time, setTime] = useState(task?.time || "09:00");
  const [reminderOn, setReminderOn] = useState(!!task?.reminder);
  const [reminder, setReminder] = useState(task?.reminder || "09:00");
  const [categoryId, setCategoryId] = useState(task?.categoryId || "");
  const [goalOn, setGoalOn] = useState(task?.goal?.type === "timed");
  const [goalMin, setGoalMin] = useState(task?.goal?.minutes || 20);
  const [subs, setSubs] = useState(task?.subtasks?.map((s) => ({ ...s })) || []);
  const [subText, setSubText] = useState("");
  const [estMin, setEstMin] = useState(task?.estMin || 0);
  const [twoMin, setTwoMin] = useState(task?.twoMin || "");
  const [energy, setEnergy] = useState(task?.energy || "");

  const toggleDay = (d) => setDays((ds) => (ds.includes(d) ? ds.filter((x) => x !== d) : [...ds, d]));
  const addSub = () => { if (subText.trim()) { setSubs((s) => [...s, { id: ruid("s"), text: subText.trim() }]); setSubText(""); } };

  const save = () => {
    if (!title.trim()) return;
    const repeat = repType === "off" ? { type: "off" } : repType === "daily" ? { type: "daily" }
      : repType === "weekdays" ? { type: "weekdays", days: days.slice().sort() } : { type: "times", times };
    onSave({
      emoji, image: image || "", title: title.trim().slice(0, 50), note: note.trim(), color, date,
      repeat, time: anytime ? null : time, reminder: reminderOn ? reminder : null, categoryId,
      goal: goalOn ? { type: "timed", minutes: goalMin } : { type: "off" },
      subtasks: subs.filter((s) => s.text.trim()),
      estMin: estMin || 0, energy: energy || null, twoMin: twoMin.trim(),
    });
  };

  const p = getPastel(color);
  return (
    <div className="fixed inset-0 z-50 flex items-end justify-center bg-slate-900/40 p-0 backdrop-blur-sm sm:items-center sm:p-4" onClick={onClose}>
      <div className="max-h-[94vh] w-full max-w-lg overflow-y-auto rounded-t-3xl bg-white p-5 shadow-xl sm:rounded-3xl" onClick={(e) => e.stopPropagation()}>
        <div className="mb-4 flex items-center justify-between">
          <h2 className="text-lg font-bold text-slate-900">{task ? "Edit task" : "New task"}</h2>
          <button onClick={onClose} className="rounded-md p-1 text-slate-400 hover:bg-slate-100"><X className="h-5 w-5" /></button>
        </div>

        {/* emoji/photo + name */}
        <div className="mb-3 flex items-center gap-3">
          <span className="relative grid h-14 w-14 shrink-0 place-items-center overflow-hidden rounded-2xl text-3xl" style={{ backgroundColor: p.card }}>
            {image ? <img src={image} alt="" className="h-full w-full object-cover" /> : emoji}
            {image && <button onClick={() => setImage("")} title="Remove photo" className="absolute right-0.5 top-0.5 grid h-5 w-5 place-items-center rounded-full bg-slate-900/60 text-white"><X className="h-3 w-3" /></button>}
          </span>
          <div className="flex-1">
            <input value={title} onChange={(e) => setTitle(e.target.value.slice(0, 50))} placeholder="Task name" className="w-full rounded-xl border border-slate-300 bg-white px-3 py-2 text-sm font-semibold focus:border-pink-400 focus:outline-none focus:ring-2 focus:ring-pink-100" />
            <div className="mt-0.5 text-right text-[10px] text-slate-400">{title.length}/50</div>
          </div>
        </div>

        {/* emoji picker */}
        <div className="mb-3 flex flex-wrap gap-1.5">
          {TASK_EMOJIS.map((e) => <button key={e} onClick={() => setEmoji(e)} className={`grid h-8 w-8 place-items-center rounded-lg text-lg transition ${emoji === e ? "bg-pink-100 ring-2 ring-pink-300" : "hover:bg-slate-100"}`}>{e}</button>)}
        </div>

        {/* color */}
        <div className="mb-3">
          <div className="mb-1.5 text-xs font-medium text-slate-500">Color</div>
          <div className="flex gap-2">
            {PASTELS.map((c) => <button key={c.id} onClick={() => setColor(c.id)} className={`h-8 w-8 rounded-full transition ${color === c.id ? "ring-2 ring-slate-900 ring-offset-2" : ""}`} style={{ backgroundColor: c.dot }} />)}
          </div>
        </div>

        {/* note */}
        <textarea value={note} onChange={(e) => setNote(e.target.value)} rows={2} placeholder="Note (optional)" className="mb-3 w-full rounded-xl border border-slate-300 bg-white px-3 py-2 text-sm focus:border-pink-400 focus:outline-none focus:ring-2 focus:ring-pink-100" />

        {/* time estimate */}
        <div className="mb-3">
          <div className="mb-1.5 text-xs font-medium text-slate-500">Скільки часу займе? (для планування дня)</div>
          <div className="flex flex-wrap items-center gap-2">
            {EST_CHIPS.map((m) => <button key={m} onClick={() => setEstMin(estMin === m ? 0 : m)} className={`rounded-full px-3 py-1.5 text-xs font-semibold ring-1 transition ${estMin === m ? "bg-pink-500 text-white ring-pink-500" : "bg-white text-slate-500 ring-slate-200 hover:ring-pink-200"}`}>{m} хв</button>)}
            <input type="number" min={0} value={estMin && !EST_CHIPS.includes(estMin) ? estMin : ""} onChange={(e) => setEstMin(Math.max(0, +e.target.value || 0))} placeholder="інше" className="w-16 rounded-full border border-slate-300 px-2 py-1.5 text-center text-xs focus:border-pink-400 focus:outline-none" />
          </div>
        </div>

        {/* energy / effort */}
        <div className="mb-3">
          <div className="mb-1.5 text-xs font-medium text-slate-500">Скільки енергії треба?</div>
          <div className="flex gap-2">
            {Object.entries(ENERGY).map(([k, v]) => <button key={k} onClick={() => setEnergy(energy === k ? "" : k)} className={`flex-1 rounded-xl px-2 py-2 text-xs font-semibold ring-1 transition ${energy === k ? "bg-pink-50 text-slate-800 ring-pink-300" : "bg-white text-slate-500 ring-slate-200 hover:ring-pink-200"}`}>{v.emoji} {v.label}</button>)}
          </div>
        </div>

        {/* 2-minute starter */}
        <div className="mb-3">
          <div className="mb-1.5 text-xs font-medium text-slate-500">Версія на 2 хвилини <span className="text-slate-400">(крихітний перший крок, щоб зрушити)</span></div>
          <input value={twoMin} onChange={(e) => setTwoMin(e.target.value)} placeholder="напр. «Прибрати кухню» → просто звільнити раковину" className="w-full rounded-xl border border-slate-300 px-3 py-2 text-sm focus:border-pink-400 focus:outline-none" />
        </div>

        <div className="space-y-3">
          <Row label="Date">
            <input type="date" value={date} onChange={(e) => setDate(e.target.value)} className="rounded-lg border border-slate-300 px-2 py-1.5 text-sm focus:border-pink-400 focus:outline-none" />
          </Row>

          <div>
            <div className="mb-1.5 text-xs font-medium text-slate-500">Repeat</div>
            <div className="flex flex-wrap gap-1.5">
              {[{ id: "off", l: "Off" }, { id: "daily", l: "Every day" }, { id: "weekdays", l: "Days" }, { id: "times", l: "X / week" }].map((o) => (
                <button key={o.id} onClick={() => setRepType(o.id)} className={`rounded-lg px-3 py-1.5 text-sm font-medium transition ${repType === o.id ? "bg-pink-500 text-white" : "bg-slate-100 text-slate-500 hover:bg-slate-200"}`}>{o.l}</button>
              ))}
            </div>
            {repType === "weekdays" && (
              <div className="mt-2 flex gap-1.5">
                {[1, 2, 3, 4, 5, 6, 0].map((i) => <button key={i} onClick={() => toggleDay(i)} className={`h-9 flex-1 rounded-lg text-xs font-semibold transition ${days.includes(i) ? "bg-pink-500 text-white" : "bg-slate-100 text-slate-500"}`}>{WD_LETTER[i]}</button>)}
              </div>
            )}
            {repType === "times" && (
              <div className="mt-2 flex items-center gap-2"><input type="number" min={1} max={7} value={times} onChange={(e) => setTimes(Math.max(1, Math.min(7, +e.target.value || 1)))} className="w-16 rounded-lg border border-slate-300 px-2 py-1.5 text-right text-sm" /><span className="text-sm text-slate-500">times per week</span></div>
            )}
          </div>

          <Row label="Time">
            <div className="flex items-center gap-2">
              <button onClick={() => setAnytime((v) => !v)} className={`rounded-lg px-3 py-1.5 text-sm font-medium ${anytime ? "bg-pink-500 text-white" : "bg-slate-100 text-slate-500"}`}>Anytime</button>
              {!anytime && <input type="time" value={time} onChange={(e) => setTime(e.target.value)} className="rounded-lg border border-slate-300 px-2 py-1.5 text-sm" />}
            </div>
          </Row>

          <Row label="Reminder">
            <div className="flex items-center gap-2">
              <button onClick={() => setReminderOn((v) => !v)} className={`rounded-lg px-3 py-1.5 text-sm font-medium ${reminderOn ? "bg-pink-500 text-white" : "bg-slate-100 text-slate-500"}`}>{reminderOn ? "On" : "Off"}</button>
              {reminderOn && <input type="time" value={reminder} onChange={(e) => setReminder(e.target.value)} className="rounded-lg border border-slate-300 px-2 py-1.5 text-sm" />}
            </div>
          </Row>
          {reminderOn && <p className="-mt-1 text-[11px] text-slate-400">Shown on the card as text. In-app only — an artifact can't send phone notifications.</p>}

          <Row label="Category">
            <select value={categoryId} onChange={(e) => setCategoryId(e.target.value)} className="rounded-lg border border-slate-300 px-2 py-1.5 text-sm focus:border-pink-400 focus:outline-none">
              <option value="">None</option>
              {categories.map((c) => <option key={c.id} value={c.id}>{c.name}</option>)}
            </select>
          </Row>

          <div>
            <div className="flex items-center justify-between">
              <div className="text-xs font-medium text-slate-500">Timed goal</div>
              <button onClick={() => setGoalOn((v) => !v)} className={`rounded-lg px-3 py-1 text-sm font-medium ${goalOn ? "bg-pink-500 text-white" : "bg-slate-100 text-slate-500"}`}>{goalOn ? "On" : "Off"}</button>
            </div>
            {goalOn && <div className="mt-2 flex items-center gap-2"><input type="number" min={1} max={240} value={goalMin} onChange={(e) => setGoalMin(Math.max(1, Math.min(240, +e.target.value || 1)))} className="w-20 rounded-lg border border-slate-300 px-2 py-1.5 text-right text-sm" /><span className="text-sm text-slate-500">minutes — shows a timer on the card</span></div>}
          </div>

          <div>
            <div className="mb-1.5 text-xs font-medium text-slate-500">Subtasks (a routine's steps)</div>
            <div className="space-y-1.5">
              {subs.map((s, i) => (
                <div key={s.id} className="flex items-center gap-2">
                  <input value={s.text} onChange={(e) => setSubs((arr) => arr.map((x) => (x.id === s.id ? { ...x, text: e.target.value } : x)))} className="flex-1 rounded-lg border border-slate-300 px-3 py-1.5 text-sm" placeholder={`Step ${i + 1}`} />
                  <button onClick={() => setSubs((arr) => arr.filter((x) => x.id !== s.id))} className="rounded p-1 text-slate-300 hover:text-red-500"><X className="h-4 w-4" /></button>
                </div>
              ))}
            </div>
            <div className="mt-2 flex items-center gap-2">
              <input value={subText} onChange={(e) => setSubText(e.target.value)} onKeyDown={(e) => { if (e.key === "Enter") addSub(); }} placeholder="Add a step…" className="flex-1 rounded-lg border border-slate-300 px-3 py-1.5 text-sm" />
              <button onClick={addSub} className="rounded-lg bg-slate-100 px-3 py-1.5 text-sm font-medium text-slate-600 hover:bg-slate-200"><Plus className="h-4 w-4" /></button>
            </div>
          </div>
        </div>

        <div className="sticky bottom-0 -mx-5 -mb-5 mt-6 flex items-center justify-end gap-3 border-t border-slate-100 bg-white/95 px-5 py-3 backdrop-blur">
          <button onClick={onClose} className="rounded-lg px-4 py-2 text-sm font-medium text-slate-500 hover:text-slate-800">Скасувати</button>
          <button onClick={save} disabled={!title.trim()} className="inline-flex items-center gap-2 rounded-xl bg-pink-500 px-5 py-2.5 font-semibold text-white shadow-sm transition hover:bg-pink-600 disabled:bg-slate-300">
            <Check className="h-4 w-4" /> {task ? "Зберегти" : "Додати"}
          </button>
        </div>
      </div>
    </div>
  );
}

function Row({ label, children }) {
  return <div className="flex items-center justify-between gap-3"><span className="text-sm font-medium text-slate-600">{label}</span><div>{children}</div></div>;
}

function CategoryManager({ categories, onClose, onAdd, onRename, onRecolor, onDelete }) {
  const [name, setName] = useState("");
  const [color, setColor] = useState("pink");
  return (
    <div className="fixed inset-0 z-50 flex items-end justify-center bg-slate-900/40 p-0 backdrop-blur-sm sm:items-center sm:p-4" onClick={onClose}>
      <div className="max-h-[92vh] w-full max-w-md overflow-y-auto rounded-t-3xl bg-white p-5 shadow-xl sm:rounded-3xl" onClick={(e) => e.stopPropagation()}>
        <div className="mb-4 flex items-center justify-between"><h2 className="text-lg font-bold text-slate-900">Categories</h2><button onClick={onClose} className="rounded-md p-1 text-slate-400 hover:bg-slate-100"><X className="h-5 w-5" /></button></div>
        <div className="space-y-2">
          {categories.map((c) => (
            <div key={c.id} className="flex items-center gap-2">
              <span className="h-4 w-4 shrink-0 rounded-full" style={{ backgroundColor: getPastel(c.color).dot }} />
              <input value={c.name} onChange={(e) => onRename(c.id, e.target.value)} className="flex-1 rounded-lg border border-slate-200 px-2 py-1.5 text-sm focus:border-pink-400 focus:outline-none" />
              <select value={c.color} onChange={(e) => onRecolor(c.id, e.target.value)} className="rounded-lg border border-slate-200 px-1 py-1.5 text-xs">
                {PASTELS.map((p) => <option key={p.id} value={p.id}>{p.id}</option>)}
              </select>
              <button onClick={() => { if (confirm(`Delete “${c.name}”? Its tasks stay, just uncategorized.`)) onDelete(c.id); }} className="rounded p-1 text-slate-300 hover:text-red-500"><Trash2 className="h-4 w-4" /></button>
            </div>
          ))}
        </div>
        <div className="mt-4 border-t border-slate-100 pt-4">
          <div className="mb-1.5 text-xs font-medium text-slate-500">New category</div>
          <div className="flex items-center gap-2">
            <div className="flex gap-1">{PASTELS.map((p) => <button key={p.id} onClick={() => setColor(p.id)} className={`h-6 w-6 rounded-full ${color === p.id ? "ring-2 ring-slate-900 ring-offset-1" : ""}`} style={{ backgroundColor: p.dot }} />)}</div>
          </div>
          <div className="mt-2 flex items-center gap-2">
            <input value={name} onChange={(e) => setName(e.target.value)} onKeyDown={(e) => { if (e.key === "Enter" && name.trim()) { onAdd(name, color); setName(""); } }} placeholder="Category name" className="flex-1 rounded-lg border border-slate-300 px-3 py-2 text-sm focus:border-pink-400 focus:outline-none" />
            <button onClick={() => { if (name.trim()) { onAdd(name, color); setName(""); } }} className="rounded-lg bg-pink-500 px-4 py-2 text-sm font-semibold text-white hover:bg-pink-600">Add</button>
          </div>
        </div>
      </div>
    </div>
  );
}

function StreakBorn({ streak, onClose }) {
  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/50 p-4 backdrop-blur-sm" onClick={onClose}>
      <div className="w-full max-w-sm rounded-3xl bg-white p-7 text-center shadow-2xl" onClick={(e) => e.stopPropagation()}>
        <div className="mx-auto mb-2 text-6xl">🔥</div>
        <div className="text-4xl font-extrabold tabular-nums text-orange-500">{streak}</div>
        <h2 className="mt-2 text-xl font-bold text-slate-900">A streak is born!</h2>
        <p className="mt-1 text-sm text-slate-500">Keep it up every day to help it grow.</p>
        <div className="mx-auto mt-4 flex max-w-[220px] justify-between">
          {[0, 1, 2, 3, 4, 5, 6].map((i) => <span key={i} className={`h-2.5 w-2.5 rounded-full ${i === 0 ? "bg-orange-400" : "bg-orange-100"}`} />)}
        </div>
        <button onClick={onClose} className="mt-6 w-full rounded-2xl bg-pink-500 py-3 font-bold text-white transition hover:bg-pink-600">I'm committed 💪</button>
      </div>
    </div>
  );
}

function RoutineStats({ tasks, completions, moods, streak, best, onBack }) {
  const [monthOffset, setMonthOffset] = useState(0);
  const total = totalTasksCompleted(completions);
  const now = new Date();
  const view = new Date(now.getFullYear(), now.getMonth() + monthOffset, 1);
  const year = view.getFullYear(), month = view.getMonth();
  const daysInMonth = new Date(year, month + 1, 0).getDate();
  const startPad = new Date(year, month, 1).getDay();

  const cells = [];
  for (let i = 0; i < startPad; i++) cells.push(null);
  for (let d = 1; d <= daysInMonth; d++) {
    const ds = dateKey(new Date(year, month, d).getTime());
    cells.push({ d, ds, prog: dayTaskProgress(tasks, completions, ds), mood: moods[ds] });
  }
  const heat = (pct, total) => !total ? "#f6e6ef" : pct >= 1 ? "#22c55e" : pct >= 0.6 ? "#86efac" : pct > 0 ? "#f9a8d4" : "#fbcfe0";

  const week = Array.from({ length: 7 }, (_, i) => {
    const ms = Date.now() - (6 - i) * DAY; const ds = dateKey(ms);
    const p = dayTaskProgress(tasks, completions, ds);
    return { label: WD_LETTER[new Date(ms).getDay()], pct: Math.round(p.pct * 100) };
  });

  return (
    <div className="mx-auto w-full max-w-2xl px-4 pb-16 pt-5">
      <div className="mb-4 flex items-center gap-2">
        <button onClick={onBack} className="grid h-9 w-9 place-items-center rounded-full bg-white text-slate-500 shadow-sm ring-1 ring-red-100"><ArrowLeft className="h-4 w-4" /></button>
        <h1 className="text-2xl font-extrabold text-slate-900">Stats</h1>
      </div>

      <div className="grid grid-cols-3 gap-3">
        <div className="rounded-2xl bg-white p-4 text-center shadow-sm ring-1 ring-red-100"><div className="text-3xl">🔥</div><div className="text-2xl font-extrabold tabular-nums text-orange-500">{streak.current}</div><div className="text-[11px] text-slate-400">day streak</div></div>
        <div className="rounded-2xl bg-white p-4 text-center shadow-sm ring-1 ring-red-100"><div className="text-3xl">🏆</div><div className="text-2xl font-extrabold tabular-nums text-amber-500">{best}</div><div className="text-[11px] text-slate-400">best streak</div></div>
        <div className="rounded-2xl bg-white p-4 text-center shadow-sm ring-1 ring-red-100"><div className="text-3xl">✅</div><div className="text-2xl font-extrabold tabular-nums text-green-500">{total}</div><div className="text-[11px] text-slate-400">completed</div></div>
      </div>

      {/* completion calendar */}
      <div className="mt-4 rounded-2xl bg-white p-5 shadow-sm ring-1 ring-red-100">
        <div className="mb-3 flex items-center justify-between">
          <h2 className="text-sm font-bold text-slate-700">{view.toLocaleDateString(undefined, { month: "long", year: "numeric" })}</h2>
          <div className="flex gap-1">
            <button onClick={() => setMonthOffset((m) => m - 1)} className="rounded-md p-1.5 text-slate-400 hover:bg-slate-100"><ArrowLeft className="h-4 w-4" /></button>
            <button onClick={() => setMonthOffset(0)} disabled={monthOffset === 0} className="rounded-md px-2 py-1 text-xs font-medium text-slate-500 hover:bg-slate-100 disabled:opacity-40">Today</button>
            <button onClick={() => setMonthOffset((m) => Math.min(0, m + 1))} disabled={monthOffset === 0} className="rounded-md p-1.5 text-slate-400 hover:bg-slate-100 disabled:opacity-40"><ArrowRight className="h-4 w-4" /></button>
          </div>
        </div>
        <div className="grid grid-cols-7 gap-1.5">
          {["M", "T", "W", "T", "F", "S", "S"].map((w, i) => <div key={i} className="pb-1 text-center text-[10px] font-semibold text-slate-400">{w}</div>)}
          {/* re-pad for Mon-first */}
          {(() => {
            const monPad = (startPad + 6) % 7;
            const out = [];
            for (let i = 0; i < monPad; i++) out.push(<div key={"p" + i} />);
            for (const c of cells.filter(Boolean)) out.push(
              <div key={c.ds} className="aspect-square rounded-lg" style={{ backgroundColor: heat(c.prog.pct, c.prog.total) }} title={`${c.ds}: ${Math.round(c.prog.pct * 100)}%`}>
                <div className="flex h-full items-center justify-center text-[10px] font-medium text-slate-600/70">{c.d}</div>
              </div>
            );
            return out;
          })()}
        </div>
      </div>

      {/* mood calendar */}
      <div className="mt-4 rounded-2xl bg-white p-5 shadow-sm ring-1 ring-red-100">
        <h2 className="mb-3 text-sm font-bold text-slate-700">Mood check-ins</h2>
        <div className="grid grid-cols-7 gap-1.5">
          {(() => {
            const monPad = (startPad + 6) % 7;
            const out = [];
            for (let i = 0; i < monPad; i++) out.push(<div key={"mp" + i} />);
            for (const c of cells.filter(Boolean)) {
              const m = MOODS.find((x) => x.score === c.mood);
              out.push(<div key={c.ds} className="grid aspect-square place-items-center rounded-lg text-sm" style={{ backgroundColor: m ? m.color + "44" : "#f1f5f9" }} title={c.ds}>{m ? m.emoji : ""}</div>);
            }
            return out;
          })()}
        </div>
      </div>

      {/* weekly chart */}
      <div className="mt-4 rounded-2xl bg-white p-5 shadow-sm ring-1 ring-red-100">
        <h2 className="mb-4 text-sm font-bold text-slate-700">This week</h2>
        <div className="h-48 w-full">
          <ResponsiveContainer width="100%" height="100%">
            <BarChart data={week} margin={{ top: 4, right: 4, left: -24, bottom: 0 }}>
              <CartesianGrid strokeDasharray="3 3" stroke="#fce7f3" vertical={false} />
              <XAxis dataKey="label" tick={{ fontSize: 11, fill: "#c084a8" }} axisLine={false} tickLine={false} />
              <YAxis domain={[0, 100]} tick={{ fontSize: 11, fill: "#c084a8" }} axisLine={false} tickLine={false} unit="%" />
              <Tooltip cursor={{ fill: "#fdf2f8" }} contentStyle={{ borderRadius: 10, border: "1px solid #fbcfe0", fontSize: 12 }} formatter={(v) => [`${v}%`, "done"]} />
              <Bar dataKey="pct" radius={[6, 6, 0, 0]}>{week.map((w, i) => <Cell key={i} fill={w.pct >= 100 ? "#22c55e" : "#ec4899"} />)}</Bar>
            </BarChart>
          </ResponsiveContainer>
        </div>
      </div>
    </div>
  );
}

/* ---------- Supportive daily plan for living with GAD ---------- */
const CALM_PLAN_DAILY = [
  { id: "morning", emoji: "🌅", title: "Ранкове заземлення", when: "щойно прокинулась", what: "2–3 хв дихання або вправа 5-4-3-2-1", why: "почати день у тілі, а не в потоці тривожних думок", go: "breath" },
  { id: "move", emoji: "🚶‍♀️", title: "Рух", when: "будь-коли вдень", what: "20–30 хв ходьби (чи будь-який рух, що подобається)", why: "регулярний рух знижує базовий рівень тривоги — це доведено" },
  { id: "worry", emoji: "⏳", title: "«Час для тривоги»", when: "пополудні, НЕ перед сном", what: "10–15 хв випиши всі «а що як», а тоді свідомо стоп", why: "збирає хвилювання в одне вікно, щоб воно не текло на весь день", go: "worry" },
  { id: "journal", emoji: "📓", title: "Вечірній розбір", when: "перед сном", what: "розбери 1 тривожну думку (факти за/проти) + запиши 3 хороші речі дня", why: "тренує реалістичне мислення й перемикає фокус із загрози на ресурс", go: "thought" },
  { id: "sleep", emoji: "😴", title: "Сон за розкладом", when: "щовечора", what: "лягай і вставай в той самий час; екран убік за 30 хв до сну", why: "недосип напряму підсилює тривогу — це фундамент усього" },
];
const CALM_PLAN_WEEKLY = [
  { emoji: "🪜", text: "Одна сходинка драбини страху — маленький крок назустріч тому, чого уникаєш." },
  { emoji: "🔎", text: "Тижневий огляд: що цього тижня заспокоювало, а що розганяло тривогу." },
  { emoji: "☕", text: "Тримай кофеїн і алкоголь у межах — обидва фізично підсилюють тривожність." },
  { emoji: "💬", text: "Хоча б одна тепла розмова чи зустріч — ізоляція годує тривогу." },
];
const CALM_PLAN_PILLARS = [
  { emoji: "😴", label: "Сон", note: "7–9 год, стабільно" },
  { emoji: "🏃‍♀️", label: "Рух", note: "щодня потроху" },
  { emoji: "☕", label: "Кофеїн", note: "менше = спокійніше" },
  { emoji: "💬", label: "Зв'язок", note: "люди, не ізоляція" },
];

function CalmPlan({ onExit, onGo }) {
  const [done, setDone] = useState({});
  const today = dateKey(Date.now());
  useEffect(() => { let on = true; store.get(CKEYS.plan, {}).then((v) => { if (on) setDone(v || {}); }); return () => { on = false; }; }, []);
  const todayMap = done[today] || {};
  const toggle = (id) => setDone((prev) => {
    const day = { ...(prev[today] || {}) };
    if (day[id]) delete day[id]; else day[id] = true;
    const next = { ...prev, [today]: day };
    store.set(CKEYS.plan, next);
    return next;
  });
  const doneCount = Object.values(todayMap).filter(Boolean).length;
  let streak = 0;
  for (let i = 0; i < 400; i++) {
    const c = Object.values(done[dateKey(Date.now() - i * DAY)] || {}).filter(Boolean).length;
    if (c >= 3) streak++; else if (i === 0) continue; else break;
  }

  return (
    <div className="mx-auto w-full max-w-lg px-4 pb-16 pt-6">
      <CalmHeader title="Мій план спокою" onExit={onExit} />

      <div className="rounded-3xl bg-gradient-to-br from-teal-400 to-sky-400 p-5 text-white shadow-sm">
        <div className="text-sm font-semibold text-white/90">Підтримувальний план при тривозі (ГТР)</div>
        <div className="mt-1 flex items-end gap-3">
          <div><div className="text-3xl font-extrabold tabular-nums">{doneCount}/{CALM_PLAN_DAILY.length}</div><div className="text-xs text-white/80">сьогодні</div></div>
          <div className="ml-auto text-right"><div className="text-3xl font-extrabold tabular-nums">🔥 {streak}</div><div className="text-xs text-white/80">днів поспіль</div></div>
        </div>
        <p className="mt-2 text-xs leading-relaxed text-white/90">Не мусиш робити все ідеально. Навіть 3 пункти на день — це вже турбота про себе. Пропустила день — просто повернись завтра, без провини. 💛</p>
      </div>

      {/* daily */}
      <div className="mt-4 mb-2 text-sm font-bold text-slate-700">Щодня</div>
      <div className="space-y-2">
        {CALM_PLAN_DAILY.map((it) => {
          const on = !!todayMap[it.id];
          return (
            <div key={it.id} className={`rounded-2xl p-4 shadow-sm ring-1 transition ${on ? "bg-teal-50 ring-teal-200" : "bg-white ring-teal-50"}`}>
              <div className="flex items-start gap-3">
                <button onClick={() => toggle(it.id)} className={`mt-0.5 grid h-7 w-7 shrink-0 place-items-center rounded-full border-2 transition ${on ? "border-transparent bg-teal-500 text-white" : "border-slate-300 hover:border-teal-400"}`}>{on && <Check className="h-4 w-4" />}</button>
                <div className="min-w-0 flex-1">
                  <div className="flex items-center gap-2"><span>{it.emoji}</span><span className={`font-bold ${on ? "text-slate-500" : "text-slate-800"}`}>{it.title}</span><span className="rounded-full bg-slate-100 px-1.5 py-0.5 text-[10px] font-semibold text-slate-500">{it.when}</span></div>
                  <div className="mt-1 text-sm text-slate-600"><b className="font-semibold text-slate-700">Що робити:</b> {it.what}</div>
                  <div className="mt-0.5 text-xs text-slate-400">чому: {it.why}</div>
                  {it.go && onGo && <button onClick={() => onGo(it.go)} className="mt-1.5 inline-flex items-center gap-1 text-xs font-semibold text-teal-600 hover:text-teal-700">Відкрити вправу <ArrowRight className="h-3 w-3" /></button>}
                </div>
              </div>
            </div>
          );
        })}
      </div>

      {/* weekly */}
      <div className="mt-5 mb-2 text-sm font-bold text-slate-700">Щотижня</div>
      <div className="space-y-2">
        {CALM_PLAN_WEEKLY.map((it, i) => (
          <div key={i} className="flex items-start gap-3 rounded-2xl bg-white p-3 shadow-sm ring-1 ring-teal-50"><span className="text-lg">{it.emoji}</span><span className="text-sm text-slate-600">{it.text}</span></div>
        ))}
      </div>

      {/* foundation pillars */}
      <div className="mt-5 mb-2 text-sm font-bold text-slate-700">Фундамент (на ньому все тримається)</div>
      <div className="grid grid-cols-4 gap-2">
        {CALM_PLAN_PILLARS.map((p) => (
          <div key={p.label} className="rounded-2xl bg-white p-3 text-center shadow-sm ring-1 ring-teal-50"><div className="text-2xl">{p.emoji}</div><div className="mt-1 text-xs font-bold text-slate-700">{p.label}</div><div className="text-[10px] text-slate-400">{p.note}</div></div>
        ))}
      </div>

      {/* how it helps */}
      <div className="mt-5 rounded-2xl bg-teal-50/70 p-4 text-sm leading-relaxed text-teal-900 ring-1 ring-teal-100">
        <div className="mb-1 font-bold">Чому саме так</div>
        При ГТР мозок «застрягає» в режимі загрози. Цей план не «прибирає» тривогу силою, а щодня потроху вчить нервову систему, що можна бути в безпеці: рух і сон знижують фізичний фон, «час для тривоги» й журнал розбирають думки, а маленькі кроки назустріч страху показують мозку, що небезпеки нема. Працює саме <b>сталість</b>, а не інтенсивність.
      </div>

      {/* safety */}
      <div className="mt-3 rounded-2xl bg-amber-50 p-4 text-xs leading-relaxed text-amber-900 ring-1 ring-amber-100">
        <div className="mb-1 flex items-center gap-1.5 font-bold"><ShieldAlert className="h-4 w-4" /> Важливо</div>
        <p>Це підтримка для щодення, а <b>не заміна</b> терапії. Найкраще працює <b>разом</b> із психотерапевтом (КПТ) і, якщо призначив лікар, медикаментами. Ліки не починай і не відміняй сама — тільки з лікарем.</p>
        <p className="mt-2">Звернись по допомогу швидше, якщо: тривога зриває сон/їжу тижнями, накрила паніка, що не минає, або з'являються думки нашкодити собі. В Україні цілодобово й безкоштовно — <b>Lifeline Ukraine 7333</b> (лінія емоційної підтримки та запобігання суїцидам). Ти не сама. 💛</p>
      </div>
    </div>
  );
}

/* ================================================================== */
/* CALM — section UI                                                  */
/* ================================================================== */
function CalmSection({ name, onRename }) {
  const [loading, setLoading] = useState(true);
  const [cview, setCview] = useState("hub");
  const [anxOpen, setAnxOpen] = useState(false);
  const [fears, setFears] = useState([]);
  const [thoughts, setThoughts] = useState([]);
  const [sessions, setSessions] = useState([]);
  const [settings, setSettings] = useState({ name: "Спокій", tick: true, pattern: "box" });
  const [toast, setToast] = useState(null);
  const [renaming, setRenaming] = useState(false);
  const [nameDraft, setNameDraft] = useState(name);

  const today = dateKey(Date.now());
  const flash = useCallback((m) => { setToast(m); window.clearTimeout(flash._t); flash._t = window.setTimeout(() => setToast(null), 2200); }, []);

  const reload = useCallback(async () => {
    const d = await loadCalmData();
    setFears(d.fears); setThoughts(d.thoughts); setSessions(d.sessions);
    setSettings({ tick: true, pattern: "box", ...d.settings });
    setLoading(false);
  }, []);
  useEffect(() => {
    reload();
    const onReset = () => { setFears([]); setThoughts([]); setSessions([]); setCview("hub"); };
    window.addEventListener("calm-reset", onReset);
    return () => window.removeEventListener("calm-reset", onReset);
  }, [reload]);

  const saveSettings = useCallback(async (patch) => {
    const next = { ...settings, ...patch };
    setSettings(next); await store.set(CKEYS.settings, next);
  }, [settings]);

  const log = useCallback(async (type, durationSec, meta = {}) => {
    const s = { id: ruid("cs"), type, date: dateKey(Date.now()), durationSec: Math.round(durationSec || 0), meta, ts: Date.now() };
    const next = [...sessions, s];
    setSessions(next); await store.set(CKEYS.sessions, next);
  }, [sessions]);

  const saveFears = useCallback(async (next) => { setFears(next); await store.set(CKEYS.fears, next); }, []);
  const saveThoughts = useCallback(async (next) => { setThoughts(next); await store.set(CKEYS.thoughts, next); }, []);

  const streak = useMemo(() => calmStreak(sessions, today), [sessions, today]);
  const minutes = useMemo(() => calmMinutes(sessions), [sessions]);

  if (loading) return <div className="flex flex-1 items-center justify-center text-teal-400"><div className="flex flex-col items-center gap-3"><Leaf className="h-8 w-8 animate-pulse" /><span className="text-sm">Завантаження…</span></div></div>;

  const back = () => setCview("hub");
  const done = (type) => (sec, meta) => { log(type, sec, meta); flash("Молодець — записано 🌿"); back(); };

  const PRACTICES = [
    { id: "breath", label: "Дихання з підказкою", desc: "Анімоване дихання, щоб сповільнитися", icon: Wind, color: "#0ea5e9" },
    { id: "pmr", label: "Розслаблення м'язів", desc: "Напруж і відпусти — від голови до п'ят", icon: HeartPulse, color: "#14b8a6" },
    { id: "ground", label: "Заземлення 5-4-3-2-1", desc: "Повернися до своїх відчуттів", icon: Anchor, color: "#6366f1" },
    { id: "thought", label: "Журнал думок", desc: "Розплутати тривожну думку", icon: NotebookPen, color: "#8b5cf6" },
    { id: "fear", label: "Сходинки страху", desc: "Назустріч страху — по одній м'якій сходинці", icon: TrendingUp, color: "#f59e0b" },
    { id: "focus", label: "Таймер фокусу", desc: "Спокійний вдих, потім зосереджена робота", icon: Timer, color: "#10b981" },
    { id: "worry", label: "Час для тривоги", desc: "Виділений час потривожитись — і відпустити", icon: Hourglass, color: "#f472b6" },
    { id: "recovery", label: "Відновлення", desc: "Тверезість, тригери й підтримка в мить пориву", icon: HandHeart, color: "#0d9488" },
  ];

  return (
    <div className="min-h-screen flex-1 bg-gradient-to-b from-teal-50 via-sky-50/50 to-white">
      {cview === "hub" && (
        <div className="mx-auto w-full max-w-2xl px-4 pb-20 pt-6">
          {/* header */}
          <div className="mb-5 flex items-center justify-between">
            <div>
              {renaming ? (
                <input autoFocus value={nameDraft} onChange={(e) => setNameDraft(e.target.value)}
                  onBlur={() => { onRename(nameDraft); setRenaming(false); }}
                  onKeyDown={(e) => { if (e.key === "Enter") { onRename(nameDraft); setRenaming(false); } }}
                  className="rounded-lg border border-teal-200 px-2 py-1 text-2xl font-extrabold text-slate-900 focus:outline-none" />
              ) : (
                <button onClick={() => { setNameDraft(name); setRenaming(true); }} className="text-left" title="Натисни, щоб перейменувати">
                  <h1 className="text-2xl font-extrabold text-slate-900">{name} <Pencil className="ml-1 inline h-4 w-4 text-slate-300" /></h1>
                </button>
              )}
              <p className="text-sm text-slate-500">Зроби вдих. Усе гаразд.</p>
            </div>
            <div className="flex items-center gap-2">
              <div className="flex items-center gap-1 rounded-full bg-white px-3 py-1.5 shadow-sm ring-1 ring-teal-100"><Leaf className="h-4 w-4 text-teal-500" /><span className="text-sm font-bold tabular-nums text-teal-600">{streak}</span></div>
              <button onClick={() => setCview("stats")} className="grid h-9 w-9 place-items-center rounded-full bg-white text-slate-500 shadow-sm ring-1 ring-teal-100 hover:text-slate-700"><BarChart3 className="h-4 w-4" /></button>
            </div>
          </div>

          {/* front door — I'm anxious right now */}
          <button onClick={() => setAnxOpen(true)} className="mb-4 flex w-full items-center gap-3 rounded-3xl bg-gradient-to-r from-red-300 via-teal-300 to-sky-300 p-4 text-left text-white shadow-lg shadow-teal-500/20 transition hover:brightness-105">
            <span className="grid h-12 w-12 shrink-0 place-items-center rounded-2xl bg-white/25"><HandHeart className="h-6 w-6" /></span>
            <span className="flex-1"><span className="block text-lg font-bold">Мені зараз тривожно</span><span className="block text-sm text-white/90">Тисни — і я підкажу, з чого почати.</span></span>
            <ArrowRight className="h-5 w-5" />
          </button>

          {/* before work */}
          <button onClick={() => setCview("beforework")} className="mb-3 flex w-full items-center gap-3 rounded-3xl bg-gradient-to-r from-teal-400 to-sky-400 p-4 text-left text-white shadow-lg shadow-teal-500/20 transition hover:from-teal-500 hover:to-sky-500">
            <span className="grid h-12 w-12 shrink-0 place-items-center rounded-2xl bg-white/20"><Sparkle className="h-6 w-6" /></span>
            <span className="flex-1"><span className="block text-lg font-bold">Перед роботою</span><span className="block text-sm text-white/90">2 хв дихання → заземлення. Одне натискання — і ти в ресурсі.</span></span>
            <ArrowRight className="h-5 w-5" />
          </button>

          <button onClick={() => setCview("plan")} className="mb-4 flex w-full items-center gap-3 rounded-3xl bg-white p-4 text-left shadow-sm ring-1 ring-teal-100 transition hover:shadow-md hover:ring-teal-200">
            <span className="grid h-12 w-12 shrink-0 place-items-center rounded-2xl bg-teal-100 text-teal-600"><ListChecks className="h-6 w-6" /></span>
            <span className="flex-1"><span className="block text-lg font-bold text-slate-800">Мій план спокою</span><span className="block text-sm text-slate-400">Щоденні кроки при тривозі (ГТР) — що, скільки й коли робити.</span></span>
            <ArrowRight className="h-5 w-5 text-slate-300" />
          </button>

          {/* stat strip */}
          <div className="mb-4 grid grid-cols-3 gap-3">
            <div className="rounded-2xl bg-white p-3 text-center shadow-sm ring-1 ring-teal-100"><div className="text-2xl font-extrabold tabular-nums text-teal-600">{minutes}</div><div className="text-[11px] text-slate-400">хвилин</div></div>
            <div className="rounded-2xl bg-white p-3 text-center shadow-sm ring-1 ring-teal-100"><div className="text-2xl font-extrabold tabular-nums text-sky-600">{sessions.length}</div><div className="text-[11px] text-slate-400">сесій</div></div>
            <div className="rounded-2xl bg-white p-3 text-center shadow-sm ring-1 ring-teal-100"><div className="text-2xl font-extrabold tabular-nums text-emerald-500">{streak}</div><div className="text-[11px] text-slate-400">днів поспіль</div></div>
          </div>

          {/* practices */}
          <div className="grid gap-3 sm:grid-cols-2">
            {PRACTICES.map((p) => (
              <button key={p.id} onClick={() => setCview(p.id)} className="flex items-center gap-3 rounded-2xl bg-white p-4 text-left shadow-sm ring-1 ring-teal-50 transition hover:shadow-md hover:ring-teal-200">
                <span className="grid h-11 w-11 shrink-0 place-items-center rounded-xl text-white" style={{ backgroundColor: p.color }}><p.icon className="h-5 w-5" /></span>
                <span className="min-w-0"><span className="block font-bold text-slate-800">{p.label}</span><span className="block truncate text-xs text-slate-400">{p.desc}</span></span>
              </button>
            ))}
          </div>

          <p className="mt-6 text-center text-xs leading-relaxed text-slate-400">
            Це інструменти самодопомоги, а не заміна професійної підтримки. Якщо тривога сильна або триває довго — розмова з фахівцем справді може допомогти. 💛
          </p>
        </div>
      )}

      {anxOpen && (
        <div className="fixed inset-0 z-50 flex items-end justify-center bg-slate-900/40 p-4 backdrop-blur-sm sm:items-center" onClick={() => setAnxOpen(false)}>
          <div className="w-full max-w-md rounded-3xl bg-white p-5 shadow-2xl" onClick={(e) => e.stopPropagation()}>
            <div className="mb-1 flex items-center gap-2"><span className="grid h-9 w-9 place-items-center rounded-2xl bg-teal-50 text-teal-500"><HandHeart className="h-5 w-5" /></span><h2 className="text-lg font-extrabold text-slate-800">Я поруч. Що зараз потрібно?</h2></div>
            <p className="mb-4 text-sm text-slate-500">Нема правильної відповіді. Обери те, що легше.</p>
            <div className="space-y-2.5">
              {[
                { icon: Wind, color: "#0ea5e9", label: "Просто подихати", desc: "М'яке дихання, щоб сповільнитися", go: () => setCview("breath") },
                { icon: Anchor, color: "#6366f1", label: "Повернутися в тіло", desc: "Заземлення 5-4-3-2-1", go: () => setCview("ground") },
                { icon: HeartPulse, color: "#14b8a6", label: "Відпустити напругу", desc: "Розслаблення м'язів", go: () => setCview("pmr") },
                { icon: NotebookPen, color: "#8b5cf6", label: "Виписати думку", desc: "Розплутати тривожну думку на папері", go: () => setCview("thought") },
              ].map((o) => (
                <button key={o.label} onClick={() => { setAnxOpen(false); o.go(); }} className="flex w-full items-center gap-3 rounded-2xl bg-white p-3 text-left ring-1 ring-slate-100 transition hover:ring-teal-200">
                  <span className="grid h-10 w-10 shrink-0 place-items-center rounded-xl text-white" style={{ backgroundColor: o.color }}><o.icon className="h-5 w-5" /></span>
                  <span className="min-w-0 flex-1"><span className="block font-bold text-slate-800">{o.label}</span><span className="block text-xs text-slate-400">{o.desc}</span></span>
                  <ArrowRight className="h-4 w-4 text-slate-300" />
                </button>
              ))}
            </div>
            <button onClick={() => setAnxOpen(false)} className="mt-4 w-full rounded-2xl py-2.5 text-sm font-semibold text-slate-400 hover:text-slate-600">Просто побути тут</button>
          </div>
        </div>
      )}

      {cview === "breath" && <BreathPractice settings={settings} saveSettings={saveSettings} onExit={back} onDone={done("breath")} />}
      {cview === "pmr" && <PMRPractice onExit={back} onDone={done("pmr")} />}
      {cview === "ground" && <GroundingPractice onExit={back} onDone={done("ground")} />}
      {cview === "thought" && <ThoughtRecord thoughts={thoughts} onExit={back} onSave={async (entry, sec) => { await saveThoughts([entry, ...thoughts]); log("thought", sec); flash("Збережено 🌿"); }} onDelete={async (id) => saveThoughts(thoughts.filter((t) => t.id !== id))} />}
      {cview === "fear" && <FearLadder fears={fears} onExit={back} onSave={saveFears} onLog={(sec, meta) => log("fear", sec, meta)} flash={flash} />}
      {cview === "focus" && <FocusTimer settings={settings} onExit={back} onDone={done("focus")} />}
      {cview === "worry" && <WorryTimer onExit={back} onDone={done("worry")} />}
      {cview === "beforework" && <BeforeWork onExit={back} onDone={(sec) => { log("beforework", sec); flash("Готово — у тебе все вийде ✨"); back(); }} />}
      {cview === "stats" && <CalmStats sessions={sessions} onExit={back} />}
      {cview === "plan" && <CalmPlan onExit={back} onGo={(v) => setCview(v)} />}
      {cview === "recovery" && <RecoveryView onExit={back} onQuickCalm={(v) => setCview(v)} />}

      {toast && <div className="fixed bottom-6 left-1/2 z-40 -translate-x-1/2 rounded-full bg-slate-900 px-4 py-2 text-sm text-white shadow-lg">{toast}</div>}
    </div>
  );
}


function CalmHeader({ title, onExit, right }) {
  return (
    <div className="mb-5 flex items-center gap-3">
      <button onClick={onExit} className="grid h-9 w-9 place-items-center rounded-full bg-white text-slate-500 shadow-sm ring-1 ring-teal-100"><ArrowLeft className="h-4 w-4" /></button>
      <h1 className="flex-1 text-xl font-extrabold text-slate-900">{title}</h1>
      {right}
    </div>
  );
}

/* ---------- Guided breathing ---------- */
function BreathPractice({ settings, saveSettings, onExit, onDone, auto }) {
  const [patternId, setPatternId] = useState(auto?.patternId || settings.pattern || "box");
  const [mode, setMode] = useState(auto?.mode || "cycles");
  const [cycles, setCycles] = useState(auto?.cycles || 6);
  const [minutes, setMinutes] = useState(auto?.minutes || 3);
  const [tick, setTick] = useState(auto ? false : settings.tick !== false);
  const [run, setRun] = useState(null); // {phaseIdx, remaining, cyclesDone, elapsed, scale}
  const raf = useRef(null);
  const started = !!run;

  const pattern = BREATH_PATTERNS.find((p) => p.id === patternId) || BREATH_PATTERNS[0];
  const cycleSecs = pattern.phases.reduce((s, [, sec]) => s + sec, 0);
  const target = mode === "cycles" ? cycles : Math.max(1, Math.round((minutes * 60) / cycleSecs));

  useEffect(() => { if (auto) start(); return () => clearInterval(raf.current); /* eslint-disable-next-line */ }, []);

  const start = () => {
    const first = pattern.phases[0];
    setRun({ phaseIdx: 0, remaining: first[1], cyclesDone: 0, elapsed: 0, scale: first[0] === "in" ? 1 : 0.5, dur: first[1] });
    if (tick) calmTick(520);
    clearInterval(raf.current);
    raf.current = setInterval(() => {
      setRun((r) => {
        if (!r) return r;
        let remaining = r.remaining - 0.1;
        let elapsed = r.elapsed + 0.1;
        if (remaining > 0.001) return { ...r, remaining, elapsed };
        // advance phase
        let idx = r.phaseIdx + 1, cyclesDone = r.cyclesDone;
        if (idx >= pattern.phases.length) { idx = 0; cyclesDone += 1; }
        if (cyclesDone >= target) { clearInterval(raf.current); setTimeout(() => onDone(Math.round(elapsed)), 10); return null; }
        const [key, sec] = pattern.phases[idx];
        if (tick) calmTick(key === "in" ? 520 : key === "out" ? 400 : 460);
        const scale = key === "in" ? 1 : key === "out" ? 0.5 : r.scale;
        return { phaseIdx: idx, remaining: sec, elapsed, cyclesDone, scale, dur: key === "hold" ? 0 : sec };
      });
    }, 100);
  };
  const stop = () => { clearInterval(raf.current); if (run && run.elapsed > 5) onDone(Math.round(run.elapsed)); else { setRun(null); onExit(); } };

  if (!started) {
    return (
      <div className="mx-auto w-full max-w-md px-4 pb-16 pt-6">
        <CalmHeader title="Дихання з підказкою" onExit={onExit} />
        <div className="space-y-4">
          <div>
            <div className="mb-1.5 text-xs font-semibold uppercase tracking-wide text-slate-400">Патерн</div>
            <div className="space-y-2">
              {BREATH_PATTERNS.map((p) => (
                <button key={p.id} onClick={() => setPatternId(p.id)} className={`flex w-full items-center justify-between rounded-2xl border p-3 text-left transition ${patternId === p.id ? "border-sky-400 bg-sky-50 ring-2 ring-sky-100" : "border-slate-200 bg-white hover:border-slate-300"}`}>
                  <span><span className="block font-bold text-slate-800">{p.name}</span><span className="block text-xs text-slate-400">{p.desc}</span></span>
                  {patternId === p.id && <CheckCircle2 className="h-5 w-5 text-sky-500" />}
                </button>
              ))}
            </div>
          </div>
          <div>
            <div className="mb-1.5 text-xs font-semibold uppercase tracking-wide text-slate-400">Тривалість</div>
            <div className="mb-2 flex gap-2">
              <button onClick={() => setMode("cycles")} className={`rounded-lg px-3 py-1.5 text-sm font-medium ${mode === "cycles" ? "bg-sky-500 text-white" : "bg-slate-100 text-slate-500"}`}>Цикли</button>
              <button onClick={() => setMode("minutes")} className={`rounded-lg px-3 py-1.5 text-sm font-medium ${mode === "minutes" ? "bg-sky-500 text-white" : "bg-slate-100 text-slate-500"}`}>Хвилини</button>
            </div>
            {mode === "cycles"
              ? <div className="flex items-center gap-2"><input type="number" min={1} max={50} value={cycles} onChange={(e) => setCycles(Math.max(1, Math.min(50, +e.target.value || 1)))} className="w-20 rounded-lg border border-slate-300 px-3 py-1.5 text-right text-sm" /><span className="text-sm text-slate-500">циклів (~{Math.round(cycles * cycleSecs / 60 * 10) / 10} хв)</span></div>
              : <div className="flex items-center gap-2"><input type="number" min={1} max={60} value={minutes} onChange={(e) => setMinutes(Math.max(1, Math.min(60, +e.target.value || 1)))} className="w-20 rounded-lg border border-slate-300 px-3 py-1.5 text-right text-sm" /><span className="text-sm text-slate-500">хвилин</span></div>}
          </div>
          <label className="flex items-center justify-between rounded-2xl border border-slate-200 bg-white px-3 py-2.5">
            <span className="text-sm font-medium text-slate-700">М'який звук-підказка</span>
            <input type="checkbox" checked={tick} onChange={(e) => { setTick(e.target.checked); saveSettings({ tick: e.target.checked }); }} className="h-4 w-4 accent-sky-500" />
          </label>
          <button onClick={() => { saveSettings({ pattern: patternId }); start(); }} className="w-full rounded-2xl bg-sky-500 py-3.5 font-bold text-white shadow-lg shadow-sky-500/20 transition hover:bg-sky-600">Почати</button>
        </div>
      </div>
    );
  }

  const [phaseKey] = pattern.phases[run.phaseIdx];
  return (
    <div className="mx-auto flex w-full max-w-md flex-col items-center px-4 pb-16 pt-6">
      <div className="mb-8 w-full"><CalmHeader title="Дихання" onExit={stop} right={<span className="text-sm font-semibold text-slate-400 tabular-nums">{run.cyclesDone + 1}/{target}</span>} /></div>
      <div className="relative my-6 grid h-72 w-72 place-items-center">
        <div className="absolute rounded-full bg-gradient-to-br from-sky-300 to-teal-300 opacity-70"
          style={{ width: 260, height: 260, transform: `scale(${run.scale})`, transition: `transform ${run.dur}s ease-in-out` }} />
        <div className="absolute rounded-full bg-gradient-to-br from-sky-400 to-teal-400"
          style={{ width: 180, height: 180, transform: `scale(${run.scale})`, transition: `transform ${run.dur}s ease-in-out` }} />
        <div className="relative z-10 text-center text-white">
          <div className="text-2xl font-bold drop-shadow">{PHASE_TEXT[phaseKey]}</div>
          <div className="text-5xl font-extrabold tabular-nums drop-shadow">{Math.ceil(run.remaining)}</div>
        </div>
      </div>
      <button onClick={stop} className="mt-6 inline-flex items-center gap-2 rounded-full bg-white px-6 py-3 font-bold text-slate-600 shadow-sm ring-1 ring-slate-200 hover:bg-slate-50"><Pause className="h-4 w-4" /> Завершити</button>
    </div>
  );
}

/* ---------- Progressive muscle relaxation ---------- */
function PMRPractice({ onExit, onDone }) {
  const [run, setRun] = useState(null); // {idx, phase:'tense'|'release', remaining, elapsed}
  const timer = useRef(null);
  const totalSec = MUSCLE_GROUPS.length * (PMR_TENSE + PMR_RELEASE);

  const start = () => {
    setRun({ idx: 0, phase: "tense", remaining: PMR_TENSE, elapsed: 0 });
    calmTick(500);
    clearInterval(timer.current);
    timer.current = setInterval(() => {
      setRun((r) => {
        if (!r) return r;
        let remaining = r.remaining - 0.1, elapsed = r.elapsed + 0.1;
        if (remaining > 0.001) return { ...r, remaining, elapsed };
        if (r.phase === "tense") { calmTick(400); return { ...r, phase: "release", remaining: PMR_RELEASE, elapsed }; }
        // release done -> next group
        const idx = r.idx + 1;
        if (idx >= MUSCLE_GROUPS.length) { clearInterval(timer.current); setTimeout(() => onDone(Math.round(elapsed)), 10); return null; }
        calmTick(520);
        return { idx, phase: "tense", remaining: PMR_TENSE, elapsed };
      });
    }, 100);
  };
  const stop = () => { clearInterval(timer.current); if (run && run.elapsed > 8) onDone(Math.round(run.elapsed)); else onExit(); };
  useEffect(() => () => clearInterval(timer.current), []);

  if (!run) return (
    <div className="mx-auto w-full max-w-md px-4 pb-16 pt-6">
      <CalmHeader title="Розслаблення м'язів" onExit={onExit} />
      <div className="rounded-3xl bg-white p-5 text-center shadow-sm ring-1 ring-teal-100">
        <HeartPulse className="mx-auto h-10 w-10 text-teal-500" />
        <p className="mt-3 text-sm text-slate-500">Пройдемося по 12 групах м'язів — від голови до п'ят. Для кожної: <b>напруж ~10 с</b>, тоді <b>відпусти ~15 с</b>. Сядь зручно і повністю відпускай на видиху.</p>
        <p className="mt-2 text-xs text-slate-400">Приблизно {Math.round(totalSec / 60)} хв.</p>
        <button onClick={start} className="mt-4 w-full rounded-2xl bg-teal-500 py-3.5 font-bold text-white shadow-lg shadow-teal-500/20 hover:bg-teal-600">Почати</button>
      </div>
    </div>
  );

  const isTense = run.phase === "tense";
  const overall = (run.elapsed / totalSec) * 100;
  return (
    <div className="mx-auto flex w-full max-w-md flex-col px-4 pb-16 pt-6">
      <CalmHeader title="Розслаблення м'язів" onExit={stop} right={<span className="text-sm font-semibold text-slate-400 tabular-nums">{run.idx + 1}/{MUSCLE_GROUPS.length}</span>} />
      <div className="mb-4 h-2 overflow-hidden rounded-full bg-slate-200"><div className="h-full rounded-full bg-teal-500 transition-all" style={{ width: `${overall}%` }} /></div>
      <div className={`grid place-items-center rounded-3xl p-8 text-center transition-colors ${isTense ? "bg-orange-50" : "bg-teal-50"}`}>
        <div className="text-xs font-semibold uppercase tracking-widest text-slate-400">Зараз</div>
        <div className="mt-1 text-3xl font-extrabold text-slate-900">{MUSCLE_GROUPS[run.idx]}</div>
        <div className={`mt-4 grid h-32 w-32 place-items-center rounded-full text-white ${isTense ? "bg-orange-400" : "bg-teal-400"}`} style={{ transform: `scale(${isTense ? 1 : 0.9})`, transition: "transform .4s" }}>
          <div><div className="text-lg font-bold">{isTense ? "Напруж…" : "Відпусти…"}</div><div className="text-3xl font-extrabold tabular-nums">{Math.ceil(run.remaining)}</div></div>
        </div>
        <div className="mt-4 text-sm text-slate-500">{isTense ? "Стисни цю групу м'язів — не до болю, просто відчутно." : "Повністю відпусти. Відчуй, яка вона мʼяка."}</div>
      </div>
      <button onClick={stop} className="mx-auto mt-6 inline-flex items-center gap-2 rounded-full bg-white px-6 py-3 font-bold text-slate-600 shadow-sm ring-1 ring-slate-200 hover:bg-slate-50"><Pause className="h-4 w-4" /> Завершити</button>
    </div>
  );
}

/* ---------- Grounding 5-4-3-2-1 ---------- */
const GROUND_SENSES = [
  { key: "see", n: 5, label: "речей, які бачиш", emoji: "👀" },
  { key: "hear", n: 4, label: "звуки, які чуєш", emoji: "👂" },
  { key: "touch", n: 3, label: "речі, яких торкаєшся", emoji: "✋" },
  { key: "smell", n: 2, label: "запахи, які відчуваєш", emoji: "👃" },
  { key: "taste", n: 1, label: "смак, який відчуваєш", emoji: "👅" },
];
function GroundingPractice({ onExit, onDone }) {
  const [filled, setFilled] = useState({});
  const startRef = useRef(Date.now());
  const total = GROUND_SENSES.reduce((s, x) => s + x.n, 0);
  const doneCount = Object.values(filled).reduce((s, arr) => s + (arr?.length || 0), 0);
  const tap = (key, i) => { calmTick(480); setFilled((f) => { const arr = f[key] || []; const has = arr.includes(i); return { ...f, [key]: has ? arr.filter((x) => x !== i) : [...arr, i] }; }); };
  const complete = doneCount >= total;

  return (
    <div className="mx-auto w-full max-w-md px-4 pb-16 pt-6">
      <CalmHeader title="Заземлення 5-4-3-2-1" onExit={onExit} right={<span className="text-sm font-semibold text-slate-400 tabular-nums">{doneCount}/{total}</span>} />
      <p className="mb-4 text-sm text-slate-500">Помічай кожне поволі. Тапай, коли знайшла. Без поспіху.</p>
      <div className="space-y-4">
        {GROUND_SENSES.map((s) => (
          <div key={s.key} className="rounded-2xl bg-white p-4 shadow-sm ring-1 ring-rose-50">
            <div className="mb-2 flex items-center gap-2"><span className="text-xl">{s.emoji}</span><span className="font-bold text-slate-800">{s.n} {s.label}</span></div>
            <div className="flex flex-wrap gap-2">
              {Array.from({ length: s.n }).map((_, i) => {
                const on = (filled[s.key] || []).includes(i);
                return <button key={i} onClick={() => tap(s.key, i)} className={`grid h-10 w-10 place-items-center rounded-full border-2 transition ${on ? "border-transparent bg-rose-500 text-white" : "border-rose-200 text-rose-300 hover:border-rose-400"}`}>{on ? <Check className="h-5 w-5" /> : <Circle className="h-4 w-4" />}</button>;
              })}
            </div>
          </div>
        ))}
      </div>
      {complete && (
        <button onClick={() => onDone(Math.round((Date.now() - startRef.current) / 1000))} className="mt-5 w-full rounded-2xl bg-rose-500 py-3.5 font-bold text-white shadow-lg shadow-rose-500/20 hover:bg-rose-600">Готово — я більше тут 🌿</button>
      )}
    </div>
  );
}

/* ---------- Thought record (CBT) ---------- */
// Stable, module-level field so the textarea keeps focus while typing whole sentences.
function TRField({ label, hint, rows = 2, value, onChange }) {
  return (
    <label className="block"><span className="mb-1 block text-sm font-semibold text-slate-700">{label}</span>{hint && <span className="mb-1 block text-xs text-slate-400">{hint}</span>}
      <textarea value={value} onChange={onChange} rows={rows} className="w-full rounded-xl border border-slate-300 bg-white px-3 py-2 text-sm focus:border-pink-400 focus:outline-none focus:ring-2 focus:ring-pink-100" /></label>
  );
}

// "How it works" guide — collapsible, calm, verbatim CBT walkthrough.
function ThoughtGuide() {
  const rows = [
    ["Ситуація", "сухо й конкретно, що сталося, як зняла б камера. Не «все жахливо», а «написала колезі о 14:00, він не відповів дві години». Тільки факти, без тлумачень."],
    ["Автоматична думка", "що миттєво промайнуло в голові, дослівно: «він на мене злий», «я всіх підвела». Саме цю думку будемо перевіряти. Часто вона категорична — «завжди», «ніколи», «всі»."],
    ["Емоція + %", "назви почуття (тривога, сором, злість) і постав інтенсивність повзунком. Це точка «до». Наприкінці порівняєш — і майже завжди відсоток падає, навіть якщо думка не зникла повністю."],
    ["Докази за", "чесно: що реально підтверджує думку? Тільки факти, не здогади. Часто виявляється, що «за» — це сама тривога, а твердих фактів обмаль."],
    ["Докази проти", "серце техніки, тут не лінуйся. Що суперечить думці? Що я забуваю, коли панікую? Опори: чи є інші пояснення? чи бувало інакше й обійшлося? що б я сказала подрузі з такою думкою? найгірший сценарій справді ймовірний — чи просто можливий?"],
    ["Врівноважена думка", "не «все чудово!» (це фальш, мозок не повірить), а тверезіший, добріший погляд, що враховує і «за», і «проти»: не «він мене ненавидить», а «можливо, він зайнятий; якщо є проблема — з'ясую, коли відповість»."],
  ];
  return (
    <div className="mt-3 rounded-2xl bg-white p-5 shadow-sm ring-1 ring-pink-100">
      <h2 className="text-lg font-extrabold text-slate-900">Як працювати з журналом думок</h2>
      <p className="mt-2 text-sm leading-relaxed text-slate-600">Це техніка з когнітивно-поведінкової терапії. Тривожна думка здається фактом, а насправді це лише одна з версій. Коли ти виносиш її на папір і розкладаєш по поличках — вона втрачає владу.</p>
      <p className="mt-3 text-sm leading-relaxed text-slate-600"><b className="font-semibold text-slate-800">Головне:</b> заповнюй у момент, коли накрило, або одразу після — не «колись увечері». Свіжа емоція і є той матеріал, з яким працюєш. Іди полями зверху вниз.</p>
      <div className="mt-3 space-y-2.5">
        {rows.map(([name, body]) => (
          <div key={name} className="rounded-xl bg-pink-50/60 p-3">
            <span className="text-sm font-bold text-pink-800">{name}</span>
            <span className="text-sm leading-relaxed text-slate-600"> — {body}</span>
          </div>
        ))}
      </div>
      <p className="mt-3 text-sm leading-relaxed text-slate-600"><b className="font-semibold text-slate-800">Наприкінці:</b> ще раз глянь на емоцію — скільки % тепер? Це зниження і є результат.</p>
      <p className="mt-3 text-sm leading-relaxed text-slate-600"><b className="font-semibold text-slate-800">Два моменти:</b> не пиши ідеально — криві формулювання нормально, сенс у тому, щоб витягти думку з голови, а не скласти твір. І веди регулярно — на дистанції побачиш, які думки й ситуації запускають тебе найчастіше.</p>
      <p className="mt-4 border-t border-slate-100 pt-3 text-xs leading-relaxed text-slate-400">Журнал думок добре доповнює терапію, але не замінює її. Якщо тривога тримається тижнями або заважає функціонувати — це сигнал звернутися до фахівця, а не слабкість.</p>
    </div>
  );
}

function ThoughtRecord({ thoughts, onExit, onSave, onDelete }) {
  const [open, setOpen] = useState(false);
  const [f, setF] = useState({ situation: "", thought: "", emotion: "", intensity: 60, forEv: "", against: "", balanced: "" });
  const [guideOpen, setGuideOpen] = useState(false);
  const startRef = useRef(Date.now());
  useEffect(() => { let on = true; store.get("calm:trGuideOpen", false).then((v) => { if (on) setGuideOpen(!!v); }); return () => { on = false; }; }, []);
  const toggleGuide = () => setGuideOpen((v) => { const nv = !v; store.set("calm:trGuideOpen", nv); return nv; });
  const setField = (k) => (e) => setF((s) => ({ ...s, [k]: e.target.value }));
  const save = () => {
    if (!f.thought.trim() && !f.situation.trim()) return;
    onSave({ id: ruid("tr"), date: dateKey(Date.now()), ...f }, Math.round((Date.now() - startRef.current) / 1000));
    setOpen(false); setF({ situation: "", thought: "", emotion: "", intensity: 60, forEv: "", against: "", balanced: "" });
  };

  if (open) return (
    <div className="mx-auto w-full max-w-lg px-4 pb-16 pt-6">
      <CalmHeader title="Журнал думок" onExit={() => setOpen(false)} />
      <div className="space-y-3">
        <TRField label="Ситуація" hint="Що відбувалося?" value={f.situation} onChange={setField("situation")} />
        <TRField label="Автоматична думка" hint="Що промайнуло в голові?" value={f.thought} onChange={setField("thought")} />
        <div className="rounded-xl bg-white p-3 ring-1 ring-slate-100">
          <div className="mb-1 flex items-center justify-between"><span className="text-sm font-semibold text-slate-700">Емоція</span><span className="text-sm font-bold text-pink-600 tabular-nums">{f.intensity}%</span></div>
          <input value={f.emotion} onChange={(e) => setF((s) => ({ ...s, emotion: e.target.value }))} placeholder="напр. тривога, сум" className="mb-2 w-full rounded-lg border border-slate-300 px-3 py-1.5 text-sm focus:border-pink-400 focus:outline-none" />
          <input type="range" min={0} max={100} value={f.intensity} onChange={(e) => setF((s) => ({ ...s, intensity: +e.target.value }))} className="w-full accent-pink-500" />
        </div>
        <TRField label="Докази за цю думку" value={f.forEv} onChange={setField("forEv")} />
        <TRField label="Докази проти неї" value={f.against} onChange={setField("against")} />
        <TRField label="Врівноважена думка" hint="Добріший, реалістичніший погляд" value={f.balanced} onChange={setField("balanced")} />
        <button onClick={save} className="w-full rounded-2xl bg-pink-500 py-3 font-bold text-white hover:bg-pink-600">Зберегти запис</button>
      </div>
    </div>
  );

  return (
    <div className="mx-auto w-full max-w-lg px-4 pb-16 pt-6">
      <CalmHeader title="Журнал думок" onExit={onExit} right={<button onClick={() => { startRef.current = Date.now(); setOpen(true); }} className="inline-flex items-center gap-1 rounded-full bg-pink-500 px-3 py-1.5 text-sm font-semibold text-white hover:bg-pink-600"><Plus className="h-4 w-4" /> Новий</button>} />
      <button onClick={toggleGuide} className="mb-1 inline-flex items-center gap-1 text-sm font-semibold text-pink-600 hover:text-pink-700">
        <Info className="h-4 w-4" /> Як це працює {guideOpen ? <ChevronDown className="h-3.5 w-3.5" /> : <ChevronRight className="h-3.5 w-3.5" />}
      </button>
      {guideOpen && <ThoughtGuide />}
      <div className="mt-3" />
      {thoughts.length === 0 ? (
        <div className="rounded-2xl bg-white py-12 text-center text-sm text-slate-400 ring-1 ring-pink-50">Записів ще немає. Злови тривожну думку і розплутай її.</div>
      ) : (
        <div className="space-y-3">
          {thoughts.map((t) => (
            <div key={t.id} className="rounded-2xl bg-white p-4 shadow-sm ring-1 ring-pink-50">
              <div className="mb-1 flex items-center justify-between"><span className="text-xs font-medium text-slate-400">{t.date}{t.emotion ? ` · ${t.emotion} ${t.intensity}%` : ""}</span><button onClick={() => onDelete(t.id)} className="text-slate-300 hover:text-red-500"><Trash2 className="h-4 w-4" /></button></div>
              {t.thought && <div className="font-semibold text-slate-800">“{t.thought}”</div>}
              {t.balanced && <div className="mt-1 rounded-lg bg-pink-50 px-3 py-2 text-sm text-pink-800">↪ {t.balanced}</div>}
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

/* ---------- Fear ladder ---------- */
function FearLadder({ fears, onExit, onSave, onLog, flash }) {
  const [adding, setAdding] = useState(false);
  const [title, setTitle] = useState("");
  const [intensity, setIntensity] = useState(50);
  const [logId, setLogId] = useState(null);
  const [before, setBefore] = useState(50);
  const [after, setAfter] = useState(30);
  const [note, setNote] = useState("");
  const [guideOpen, setGuideOpen] = useState(false);

  const sorted = [...fears].sort((a, b) => a.intensity - b.intensity);
  const nowLevel = (f) => { const a = f.attempts || []; return a.length ? a[a.length - 1].after : f.intensity; };
  const isMastered = (f) => (f.attempts || []).length > 0 && nowLevel(f) <= 20;
  // "твоя сходинка зараз" = найлегша ще не приборкана
  const currentStepId = sorted.find((f) => !isMastered(f))?.id;
  const addFear = () => { if (!title.trim()) return; onSave([...fears, { id: ruid("f"), title: title.trim(), intensity, created: Date.now(), attempts: [] }]); setTitle(""); setIntensity(50); setAdding(false); };
  const saveAttempt = () => {
    const next = fears.map((f) => f.id === logId ? { ...f, attempts: [...(f.attempts || []), { date: dateKey(Date.now()), before, after, note: note.trim() }] } : f);
    onSave(next); onLog(120, { fearId: logId });
    if (after < before) flash("Крок уперед — молодець 🌱"); else flash("Записано. Будь до себе лагідною.");
    setLogId(null); setNote(""); setBefore(50); setAfter(30);
  };
  const removeFear = (id) => onSave(fears.filter((f) => f.id !== id));

  const logFear = fears.find((f) => f.id === logId);

  return (
    <div className="mx-auto w-full max-w-lg px-4 pb-16 pt-6">
      <CalmHeader title="Сходинки страху" onExit={onExit} right={<button onClick={() => setAdding(true)} className="inline-flex items-center gap-1 rounded-full bg-amber-500 px-3 py-1.5 text-sm font-semibold text-white hover:bg-amber-600"><Plus className="h-4 w-4" /> Додати</button>} />

      {/* how exposure works */}
      <div className="mb-4 overflow-hidden rounded-2xl bg-amber-50 ring-1 ring-amber-100">
        <button onClick={() => setGuideOpen((v) => !v)} className="flex w-full items-center gap-2 px-3 py-2.5 text-left text-sm font-bold text-amber-800">
          <HelpCircle className="h-4 w-4 shrink-0" /> Як страх «розчиняється» (2 хв читання)
          {guideOpen ? <ChevronDown className="ml-auto h-4 w-4" /> : <ChevronRight className="ml-auto h-4 w-4" />}
        </button>
        {guideOpen && (
          <div className="space-y-3 px-3 pb-3 text-xs leading-relaxed text-amber-900">
            <p><b>Уникання годує страх.</b> Коли ми тікаємо від того, що лякає, мозок «запам'ятовує»: це було небезпечно. Страх міцнішає. Експозиція робить навпаки — ти лишаєшся в ситуації, тривога піднімається і сама спадає, і мозок вчиться: нічого страшного не сталося.</p>
            <p><b>Тривога має піднятися — це і є робота.</b> «У спокійному стані» не означає бути розслабленою перед спробою — інакше це й не була б експозиція. Це означає підходити при ясній голові: не в гострій паніці, не під алкоголем чи заспокійливими, не в момент, коли й так усе валиться. А під час спроби тривога росте — і має рости. Ось твоя крива: пік, потім спад, і наступного разу пік нижчий.</p>
            <div className="rounded-xl bg-white/70 p-2">
              <div className="mb-1 text-[11px] font-bold text-amber-700">Крива тривоги за кілька спроб ↓</div>
              <svg viewBox="0 0 220 60" className="w-full">
                <path d="M0,50 C18,10 30,10 45,34 C60,50 62,50 70,52" fill="none" stroke="#f59e0b" strokeWidth="2.5" />
                <path d="M70,52 C86,26 96,26 108,42 C120,52 122,52 130,53" fill="none" stroke="#fb923c" strokeWidth="2.5" opacity="0.8" />
                <path d="M130,53 C144,40 152,40 162,49 C172,55 176,55 220,55" fill="none" stroke="#34d399" strokeWidth="2.5" opacity="0.85" />
              </svg>
            </div>
            <div>
              <p className="font-bold">Головне — не панікувати й не тікати, а не «бути спокійною».</p>
              <p className="mt-1">Є різниця між двома речами:</p>
              <ul className="mt-1.5 space-y-1.5">
                <li className="flex gap-2"><span className="shrink-0">✅</span><span><b>Працює:</b> ти спокійно вирішуєш зайти на сходинку → тривога всередині росте → ти лишаєшся й даєш їй піднятись і спасти, не тікаючи. Мозок вчиться: «я витримала».</span></li>
                <li className="flex gap-2"><span className="shrink-0">⛔</span><span><b>Не працює:</b> ти зненацька, на піку паніки опиняєшся в найстрашнішому → тривога зашкалює → ти тікаєш. Мозок запам'ятовує «було жахливо, добре що втекла» — і страх закріплюється.</span></li>
              </ul>
              <p className="mt-1.5">Тому не кидайся в найстрашніше раптово й на піку паніки. Драбина для того й потрібна: щоб на кожній сходинці тривога була керована — досить сильна, щоб вчитися, але не така, що ти вилітаєш.</p>
            </div>
            <div>
              <p className="font-bold">Три ознаки, що спроба та сама, що треба:</p>
              <ul className="mt-1.5 space-y-1">
                <li className="flex gap-2"><span className="mt-px shrink-0 text-amber-500">•</span><span>ти сама обрала зайти, а не тебе загнало в кут зненацька;</span></li>
                <li className="flex gap-2"><span className="mt-px shrink-0 text-amber-500">•</span><span>тривога піднімається, але ти лишаєшся й даєш їй спасти, а не тікаєш на піку;</span></li>
                <li className="flex gap-2"><span className="mt-px shrink-0 text-amber-500">•</span><span>після можеш сказати «було важко, але я витримала», а не «це був жах, більше ніколи».</span></li>
              </ul>
            </div>
            <p><b>Якщо накрило панікою і втримати не вийшло</b> — це не провал і не «я роблю неправильно». Це означає, що сходинка завелика. Розбий її на дрібнішу: не «виступ перед залою», а «сказати одну фразу вголос при одній людині». Драбина працює, коли кроки достатньо малі, щоб тривога лишалася в зоні «важко, але терпимо».</p>
            <p><b>Навіщо тоді дихання і заземлення?</b> Не щоб зробити спробу непотрібно спокійною. А щоб зайти при ясній голові (а не на піку паніки) і щоб лишитися на сходинці, коли тривога росте, замість тікати. Не прибрати тривогу — а витримати її. 🧡</p>
          </div>
        )}
      </div>

      {adding && (
        <div className="mb-4 rounded-2xl bg-white p-4 shadow-sm ring-1 ring-amber-100">
          <input value={title} onChange={(e) => setTitle(e.target.value)} placeholder="Страх чи тривога…" className="mb-3 w-full rounded-lg border border-slate-300 px-3 py-2 text-sm focus:border-amber-400 focus:outline-none" />
          <div className="mb-1 flex justify-between text-sm"><span className="font-medium text-slate-600">Наскільки страшно зараз?</span><span className="font-bold text-amber-600 tabular-nums">{intensity}</span></div>
          <input type="range" min={0} max={100} value={intensity} onChange={(e) => setIntensity(+e.target.value)} className="w-full accent-amber-500" />
          <div className="mt-3 flex gap-2"><button onClick={addFear} className="flex-1 rounded-xl bg-amber-500 py-2 font-semibold text-white hover:bg-amber-600">Додати до сходинок</button><button onClick={() => setAdding(false)} className="rounded-xl px-4 py-2 text-sm text-slate-500">Скасувати</button></div>
        </div>
      )}

      {sorted.length === 0 && !adding ? (
        <div className="rounded-2xl bg-white py-12 text-center text-sm text-slate-400 ring-1 ring-amber-50">Додай кілька тривог — ми розкладемо їх від найлегшої до найважчої.</div>
      ) : (
        <div className="space-y-2">
          <div className="mb-1 flex items-center justify-between px-1 text-[11px] font-semibold uppercase tracking-wide text-slate-400"><span>↑ важче</span><span>рухайся знизу вгору</span></div>
          {[...sorted].reverse().map((f, ri) => {
            const step = sorted.length - ri;
            const atts = f.attempts || [];
            const now = nowLevel(f);
            const start = f.intensity;
            const drop = Math.max(0, start - now);
            const mastered = isMastered(f);
            const isCurrent = f.id === currentStepId;
            const pct = Math.max(0, Math.min(100, now));
            const startPct = Math.max(0, Math.min(100, start));
            const barColor = now <= 20 ? "bg-green-500" : now <= 45 ? "bg-amber-400" : "bg-orange-500";
            return (
              <div key={f.id} className={`rounded-2xl bg-white p-4 shadow-sm ring-1 transition ${mastered ? "ring-green-100" : isCurrent ? "ring-2 ring-amber-300" : "ring-amber-50"}`}>
                <div className="flex items-center gap-2">
                  <span className={`grid h-8 w-8 shrink-0 place-items-center rounded-lg text-sm font-bold ${mastered ? "bg-green-100 text-green-700" : "bg-amber-100 text-amber-700"}`}>{mastered ? <Check className="h-4 w-4" /> : step}</span>
                  <div className="min-w-0 flex-1">
                    <div className="flex items-center gap-1.5"><span className="truncate font-bold text-slate-800">{f.title}</span>
                      {mastered ? <span className="shrink-0 rounded-full bg-green-100 px-1.5 py-0.5 text-[10px] font-bold text-green-700">приборкано</span>
                        : isCurrent ? <span className="shrink-0 rounded-full bg-amber-400 px-1.5 py-0.5 text-[10px] font-bold text-white">твоя сходинка</span>
                        : atts.length ? <span className="shrink-0 rounded-full bg-amber-50 px-1.5 py-0.5 text-[10px] font-bold text-amber-600">в роботі</span>
                        : <span className="shrink-0 rounded-full bg-slate-100 px-1.5 py-0.5 text-[10px] font-bold text-slate-400">попереду</span>}
                    </div>
                    <div className="text-xs text-slate-400">{atts.length ? `${atts.length} ${atts.length === 1 ? "спроба" : atts.length < 5 ? "спроби" : "спроб"}` : "ще не пробувала"}{drop > 0 ? ` · впала на ${drop}` : ""}</div>
                  </div>
                  <button onClick={() => { setLogId(f.id); setBefore(now); }} className="shrink-0 rounded-full bg-amber-500 px-3 py-1.5 text-xs font-semibold text-white hover:bg-amber-600">Записати спробу</button>
                  <button onClick={() => { if (confirm("Прибрати цей страх?")) removeFear(f.id); }} className="shrink-0 text-slate-300 hover:text-red-500"><Trash2 className="h-4 w-4" /></button>
                </div>
                {/* vector: anxiety from start → now, aiming for 0 */}
                <div className="mt-2.5">
                  <div className="relative h-2 rounded-full bg-slate-100">
                    <div className={`h-full rounded-full transition-all ${barColor}`} style={{ width: `${pct}%` }} />
                    {atts.length > 0 && startPct > pct && <div className="absolute top-1/2 h-3 w-0.5 -translate-y-1/2 bg-slate-300" style={{ left: `${startPct}%` }} title={`старт ${start}`} />}
                  </div>
                  <div className="mt-1 flex items-center justify-between text-[10px] font-medium text-slate-400"><span>тривога зараз <b className="text-slate-600">{now}</b>/100</span><span>ціль ≤ 20</span></div>
                </div>
                {atts.length > 1 && (
                  <div className="mt-2 flex items-end gap-1" title="кожна спроба: тривога після">
                    {atts.slice(-14).map((a, i) => <div key={i} title={`${a.date}: ${a.before}→${a.after}`} className={`flex-1 rounded-t ${a.after <= 20 ? "bg-green-400" : "bg-amber-300"}`} style={{ height: Math.max(4, (a.after / 100) * 32) }} />)}
                  </div>
                )}
              </div>
            );
          })}
        </div>
      )}

      {logFear && (
        <div className="fixed inset-0 z-50 flex items-end justify-center bg-slate-900/40 p-0 backdrop-blur-sm sm:items-center sm:p-4" onClick={() => setLogId(null)}>
          <div className="w-full max-w-md rounded-t-3xl bg-white p-5 shadow-xl sm:rounded-3xl" onClick={(e) => e.stopPropagation()}>
            <h3 className="mb-1 text-lg font-bold text-slate-900">Записати спробу</h3>
            <p className="mb-4 text-sm text-slate-500">“{logFear.title}”</p>
            <div className="mb-3"><div className="mb-1 flex justify-between text-sm"><span className="font-medium text-slate-600">Тривога до</span><span className="font-bold text-slate-700 tabular-nums">{before}</span></div><input type="range" min={0} max={100} value={before} onChange={(e) => setBefore(+e.target.value)} className="w-full accent-amber-500" /></div>
            <div className="mb-3"><div className="mb-1 flex justify-between text-sm"><span className="font-medium text-slate-600">Тривога після</span><span className="font-bold text-teal-600 tabular-nums">{after}</span></div><input type="range" min={0} max={100} value={after} onChange={(e) => setAfter(+e.target.value)} className="w-full accent-teal-500" /></div>
            <textarea value={note} onChange={(e) => setNote(e.target.value)} rows={2} placeholder="Як усе пройшло? (необов'язково)" className="mb-3 w-full rounded-xl border border-slate-300 px-3 py-2 text-sm focus:border-amber-400 focus:outline-none" />
            <button onClick={saveAttempt} className="w-full rounded-2xl bg-amber-500 py-3 font-bold text-white hover:bg-amber-600">Зберегти спробу</button>
          </div>
        </div>
      )}
    </div>
  );
}

/* ---------- Focus timer (Pomodoro w/ calming breath) ---------- */
function FocusTimer({ settings, onExit, onDone }) {
  const [workMin, setWorkMin] = useState(25);
  const [breakMin, setBreakMin] = useState(5);
  const [stage, setStage] = useState("config"); // config | breathe | work | break | done
  const [remaining, setRemaining] = useState(0);
  const [breatheLeft, setBreatheLeft] = useState(3);
  const timer = useRef(null);
  const workedRef = useRef(0);

  useEffect(() => () => clearInterval(timer.current), []);

  const startBreath = () => { setStage("breathe"); setBreatheLeft(3); clearInterval(timer.current); calmTick(520);
    timer.current = setInterval(() => setBreatheLeft((b) => { if (b <= 1) { clearInterval(timer.current); startWork(); return 0; } calmTick(480); return b - 1; }), 4000);
  };
  const startWork = () => { setStage("work"); setRemaining(workMin * 60); tickDown("work"); };
  const startBreak = () => { setStage("break"); setRemaining(breakMin * 60); tickDown("break"); };
  const tickDown = (which) => {
    clearInterval(timer.current);
    timer.current = setInterval(() => setRemaining((r) => {
      if (r <= 1) {
        clearInterval(timer.current);
        if (which === "work") { workedRef.current += workMin * 60; calmTick(560); setStage("break"); setRemaining(breakMin * 60); tickDown("break"); return breakMin * 60; }
        else { calmTick(520); onDone(workedRef.current || workMin * 60); return 0; }
      }
      return r - 1;
    }), 1000);
  };
  const stop = () => { clearInterval(timer.current); const worked = workedRef.current + (stage === "work" ? (workMin * 60 - remaining) : 0); if (worked > 20) onDone(worked); else onExit(); };

  if (stage === "config") return (
    <div className="mx-auto w-full max-w-md px-4 pb-16 pt-6">
      <CalmHeader title="Таймер фокусу" onExit={onExit} />
      <div className="rounded-3xl bg-white p-5 shadow-sm ring-1 ring-emerald-100">
        <p className="mb-4 text-sm text-slate-500">Почнемо з трьох повільних вдихів, тоді — блок зосередженої роботи, а потім коротка перерва.</p>
        <div className="mb-3 flex items-center justify-between"><span className="text-sm font-medium text-slate-700">Робота</span><span className="flex items-center gap-2"><input type="number" min={1} max={120} value={workMin} onChange={(e) => setWorkMin(Math.max(1, Math.min(120, +e.target.value || 1)))} className="w-16 rounded-lg border border-slate-300 px-2 py-1 text-right text-sm" /><span className="text-sm text-slate-400">хв</span></span></div>
        <div className="mb-4 flex items-center justify-between"><span className="text-sm font-medium text-slate-700">Перерва</span><span className="flex items-center gap-2"><input type="number" min={1} max={60} value={breakMin} onChange={(e) => setBreakMin(Math.max(1, Math.min(60, +e.target.value || 1)))} className="w-16 rounded-lg border border-slate-300 px-2 py-1 text-right text-sm" /><span className="text-sm text-slate-400">хв</span></span></div>
        <button onClick={startBreath} className="w-full rounded-2xl bg-emerald-500 py-3.5 font-bold text-white shadow-lg shadow-emerald-500/20 hover:bg-emerald-600">Почати</button>
      </div>
    </div>
  );

  if (stage === "breathe") return (
    <div className="mx-auto flex w-full max-w-md flex-col items-center px-4 pb-16 pt-6">
      <div className="w-full"><CalmHeader title="Налаштуйся" onExit={stop} /></div>
      <div className="my-10 grid h-56 w-56 place-items-center rounded-full bg-gradient-to-br from-emerald-300 to-teal-300 text-white">
        <div className="text-center"><div className="text-lg font-bold">Дихай повільно</div><div className="text-5xl font-extrabold tabular-nums">{breatheLeft}</div></div>
      </div>
      <p className="text-sm text-slate-500">Кілька спокійних вдихів перед фокусом…</p>
    </div>
  );

  const isWork = stage === "work";
  return (
    <div className="mx-auto flex w-full max-w-md flex-col items-center px-4 pb-16 pt-6">
      <div className="w-full"><CalmHeader title={isWork ? "Фокус" : "Перерва"} onExit={stop} /></div>
      <div className={`my-10 grid h-64 w-64 place-items-center rounded-full text-white ${isWork ? "bg-gradient-to-br from-emerald-400 to-teal-400" : "bg-gradient-to-br from-sky-300 to-teal-300"}`}>
        <div className="text-center"><div className="text-sm font-semibold uppercase tracking-widest text-white/80">{isWork ? "Фокус" : "Відпочинок"}</div><div className="text-6xl font-extrabold tabular-nums">{fmtClock(remaining)}</div></div>
      </div>
      <button onClick={stop} className="inline-flex items-center gap-2 rounded-full bg-white px-6 py-3 font-bold text-slate-600 shadow-sm ring-1 ring-slate-200 hover:bg-slate-50"><Pause className="h-4 w-4" /> Завершити</button>
    </div>
  );
}

/* ---------- Worry time ---------- */
function WorryTimer({ onExit, onDone }) {
  const [min, setMin] = useState(10);
  const [remaining, setRemaining] = useState(0);
  const [running, setRunning] = useState(false);
  const timer = useRef(null);
  useEffect(() => () => clearInterval(timer.current), []);
  const start = () => { setRunning(true); setRemaining(min * 60); clearInterval(timer.current); timer.current = setInterval(() => setRemaining((r) => { if (r <= 1) { clearInterval(timer.current); calmTick(440); onDone(min * 60); return 0; } return r - 1; }), 1000); };
  const stop = () => { clearInterval(timer.current); const spent = min * 60 - remaining; if (spent > 20) onDone(spent); else onExit(); };

  if (!running) return (
    <div className="mx-auto w-full max-w-md px-4 pb-16 pt-6">
      <CalmHeader title="Час для тривоги" onExit={onExit} />
      <div className="rounded-3xl bg-white p-5 shadow-sm ring-1 ring-pink-100">
        <p className="mb-3 text-sm leading-relaxed text-slate-600">Це <b>свідомо виділений, обмежений у часі</b> проміжок, щоб дозволити собі потривожитись навмисне. Постав таймер, тривожся вільно — пиши, думай, відчувай це — а коли час вийде, м'яко відклади це вбік. Це техніка, щоб <i>вмістити</i> тривогу, а не занурюватись у неї.</p>
        <div className="mb-4 flex items-center justify-between"><span className="text-sm font-medium text-slate-700">Скільки часу?</span><span className="flex items-center gap-2"><input type="number" min={1} max={60} value={min} onChange={(e) => setMin(Math.max(1, Math.min(60, +e.target.value || 1)))} className="w-16 rounded-lg border border-slate-300 px-2 py-1 text-right text-sm" /><span className="text-sm text-slate-400">хв</span></span></div>
        <button onClick={start} className="w-full rounded-2xl bg-pink-500 py-3.5 font-bold text-white shadow-lg shadow-pink-500/20 hover:bg-pink-600">Почати час для тривоги</button>
      </div>
    </div>
  );
  return (
    <div className="mx-auto flex w-full max-w-md flex-col items-center px-4 pb-16 pt-6">
      <div className="w-full"><CalmHeader title="Час для тривоги" onExit={stop} /></div>
      <div className="my-10 grid h-64 w-64 place-items-center rounded-full bg-gradient-to-br from-pink-300 to-red-300 text-white"><div className="text-6xl font-extrabold tabular-nums">{fmtClock(remaining)}</div></div>
      <p className="mb-4 max-w-xs text-center text-sm text-slate-500">Хай усе підніметься зараз. Коли таймер закінчиться — ми зачинимо ці двері. Поки що.</p>
      <button onClick={stop} className="inline-flex items-center gap-2 rounded-full bg-white px-6 py-3 font-bold text-slate-600 shadow-sm ring-1 ring-slate-200 hover:bg-slate-50"><Check className="h-4 w-4" /> Я закінчила</button>
    </div>
  );
}

/* ---------- Before-work chain ---------- */
function BeforeWork({ onExit, onDone }) {
  const [step, setStep] = useState("intro"); // intro | breath | ground | done
  const acc = useRef(0);
  return (
    <div className="min-h-full">
      {step === "intro" && (
        <div className="mx-auto w-full max-w-md px-4 pb-16 pt-6">
          <CalmHeader title="Перед роботою" onExit={onExit} />
          <div className="rounded-3xl bg-white p-6 text-center shadow-sm ring-1 ring-teal-100">
            <span className="mx-auto grid h-14 w-14 place-items-center rounded-2xl bg-teal-100 text-teal-600"><Sparkle className="h-7 w-7" /></span>
            <h2 className="mt-3 text-xl font-bold text-slate-900">М'яке перезавантаження</h2>
            <p className="mt-2 text-sm text-slate-500">Дві хвилини дихання, потім швидке заземлення. І ти будеш готова почати — спокійно й ясно.</p>
            <button onClick={() => setStep("breath")} className="mt-5 w-full rounded-2xl bg-teal-500 py-3.5 font-bold text-white shadow-lg shadow-teal-500/20 hover:bg-teal-600">Почати</button>
          </div>
        </div>
      )}
      {step === "breath" && <BreathPractice auto={{ patternId: "relax", mode: "minutes", minutes: 2 }} settings={{ tick: false }} saveSettings={() => {}} onExit={onExit} onDone={(sec) => { acc.current += sec; setStep("ground"); }} />}
      {step === "ground" && <GroundingPractice onExit={onExit} onDone={(sec) => { acc.current += sec; setStep("done"); }} />}
      {step === "done" && (
        <div className="mx-auto w-full max-w-md px-4 pb-16 pt-16 text-center">
          <div className="mx-auto grid h-16 w-16 place-items-center rounded-2xl bg-teal-100 text-teal-600"><Check className="h-8 w-8" /></div>
          <h2 className="mt-4 text-2xl font-bold text-slate-900">Ти готова ✨</h2>
          <p className="mt-1 text-sm text-slate-500">Спокійно й ясно. Крок за кроком, по одній справі.</p>
          <button onClick={() => onDone(acc.current)} className="mt-6 rounded-2xl bg-teal-500 px-8 py-3 font-bold text-white hover:bg-teal-600">Почати мій день</button>
        </div>
      )}
    </div>
  );
}

/* ---------- Calm stats ---------- */
function CalmStats({ sessions, onExit }) {
  const [monthOffset, setMonthOffset] = useState(0);
  const minutes = calmMinutes(sessions);
  const streak = calmStreak(sessions);
  const byType = {};
  for (const s of sessions) byType[s.type] = (byType[s.type] || 0) + 1;

  const now = new Date();
  const view = new Date(now.getFullYear(), now.getMonth() + monthOffset, 1);
  const year = view.getFullYear(), month = view.getMonth();
  const daysInMonth = new Date(year, month + 1, 0).getDate();
  const startPad = (new Date(year, month, 1).getDay() + 6) % 7;
  const dayCount = {};
  for (const s of sessions) dayCount[s.date] = (dayCount[s.date] || 0) + 1;

  return (
    <div className="mx-auto w-full max-w-2xl px-4 pb-16 pt-6">
      <CalmHeader title="Твоя практика спокою" onExit={onExit} />
      <div className="grid grid-cols-3 gap-3">
        <div className="rounded-2xl bg-white p-4 text-center shadow-sm ring-1 ring-teal-100"><div className="text-3xl">🌿</div><div className="text-2xl font-extrabold tabular-nums text-teal-600">{streak}</div><div className="text-[11px] text-slate-400">днів поспіль</div></div>
        <div className="rounded-2xl bg-white p-4 text-center shadow-sm ring-1 ring-teal-100"><div className="text-3xl">⏱️</div><div className="text-2xl font-extrabold tabular-nums text-sky-600">{minutes}</div><div className="text-[11px] text-slate-400">хвилин</div></div>
        <div className="rounded-2xl bg-white p-4 text-center shadow-sm ring-1 ring-teal-100"><div className="text-3xl">🧘</div><div className="text-2xl font-extrabold tabular-nums text-emerald-500">{sessions.length}</div><div className="text-[11px] text-slate-400">сесій</div></div>
      </div>

      <div className="mt-4 rounded-2xl bg-white p-5 shadow-sm ring-1 ring-teal-100">
        <h2 className="mb-3 text-sm font-bold text-slate-700">За практикою</h2>
        {Object.keys(byType).length === 0 ? <p className="text-sm text-slate-400">Сесій ще немає.</p> : (
          <div className="space-y-2">
            {Object.entries(byType).sort((a, b) => b[1] - a[1]).map(([type, n]) => {
              const T = CALM_TECHNIQUES[type] || { label: type, icon: Leaf };
              const max = Math.max(...Object.values(byType));
              return <div key={type} className="flex items-center gap-2"><T.icon className="h-4 w-4 shrink-0 text-teal-500" /><span className="w-36 shrink-0 truncate text-sm text-slate-600">{T.label}</span><div className="h-2.5 flex-1 overflow-hidden rounded-full bg-slate-100"><div className="h-full rounded-full bg-teal-400" style={{ width: `${(n / max) * 100}%` }} /></div><span className="w-6 text-right text-xs font-semibold tabular-nums text-slate-500">{n}</span></div>;
            })}
          </div>
        )}
      </div>

      <div className="mt-4 rounded-2xl bg-white p-5 shadow-sm ring-1 ring-teal-100">
        <div className="mb-3 flex items-center justify-between">
          <h2 className="text-sm font-bold text-slate-700">{view.toLocaleDateString("uk-UA", { month: "long", year: "numeric" })}</h2>
          <div className="flex gap-1">
            <button onClick={() => setMonthOffset((m) => m - 1)} className="rounded-md p-1.5 text-slate-400 hover:bg-slate-100"><ArrowLeft className="h-4 w-4" /></button>
            <button onClick={() => setMonthOffset(0)} disabled={monthOffset === 0} className="rounded-md px-2 py-1 text-xs font-medium text-slate-500 hover:bg-slate-100 disabled:opacity-40">Сьогодні</button>
            <button onClick={() => setMonthOffset((m) => Math.min(0, m + 1))} disabled={monthOffset === 0} className="rounded-md p-1.5 text-slate-400 hover:bg-slate-100 disabled:opacity-40"><ArrowRight className="h-4 w-4" /></button>
          </div>
        </div>
        <div className="grid grid-cols-7 gap-1.5">
          {["Пн", "Вт", "Ср", "Чт", "Пт", "Сб", "Нд"].map((w, i) => <div key={i} className="pb-1 text-center text-[10px] font-semibold text-slate-400">{w}</div>)}
          {Array.from({ length: startPad }).map((_, i) => <div key={"p" + i} />)}
          {Array.from({ length: daysInMonth }).map((_, i) => {
            const d = i + 1; const ds = dateKey(new Date(year, month, d).getTime()); const c = dayCount[ds] || 0;
            const bg = !c ? "#eef2f5" : c >= 3 ? "#0d9488" : c === 2 ? "#2dd4bf" : "#99f6e4";
            return <div key={ds} className="grid aspect-square place-items-center rounded-lg text-[10px] font-medium" style={{ backgroundColor: bg, color: c ? "#fff" : "#94a3b8" }} title={`${ds}: ${c}`}>{d}</div>;
          })}
        </div>
      </div>

      <p className="mt-6 text-center text-xs leading-relaxed text-slate-400">Інструменти самодопомоги, а не заміна професійної підтримки. Якщо стає важко — звернутися до когось справді може допомогти. 💛</p>
    </div>
  );
}

/* ================================================================== */
/* FASTING — section UI                                               */
/* ================================================================== */
function FastingSection({ name, onRename }) {
  const [loading, setLoading] = useState(true);
  const [fview, setFview] = useState("timer"); // timer | ladder | diary | overview | reference
  const [goals, setGoals] = useState({ startWeight: null, targetWeight: null, protocol: "16:8", startDate: dateKey(Date.now()) });
  const [diary, setDiary] = useState([]);
  const [current, setCurrent] = useState(null);
  const [diaryEditor, setDiaryEditor] = useState(null); // {entry} | {entry:null}
  const [renaming, setRenaming] = useState(false);
  const [nameDraft, setNameDraft] = useState(name);
  const [toast, setToast] = useState(null);
  const [now, setNow] = useState(Date.now());

  const flash = useCallback((m) => { setToast(m); window.clearTimeout(flash._t); flash._t = window.setTimeout(() => setToast(null), 2400); }, []);

  const reload = useCallback(async () => {
    const d = await loadFastingData();
    setGoals(d.goals); setDiary(d.diary); setCurrent(d.current); setLoading(false);
  }, []);
  useEffect(() => {
    reload();
    const onReset = () => { setGoals({ startWeight: null, targetWeight: null, protocol: "16:8", startDate: dateKey(Date.now()) }); setDiary([]); setCurrent(null); setFview("timer"); };
    window.addEventListener("fasting-reset", onReset);
    return () => window.removeEventListener("fasting-reset", onReset);
  }, [reload]);

  // tick while a fast is running
  useEffect(() => {
    if (!current) return;
    const t = setInterval(() => setNow(Date.now()), 1000);
    setNow(Date.now());
    return () => clearInterval(t);
  }, [current]);

  const saveGoals = useCallback(async (patch) => { const next = { ...goals, ...patch }; setGoals(next); await store.set(FKEYS.goals, next); }, [goals]);
  const saveDiary = useCallback(async (next) => { setDiary(next); await store.set(FKEYS.diary, next); }, []);
  const saveCurrent = useCallback(async (val) => { setCurrent(val); await store.set(FKEYS.current, val); }, []);

  const protocol = getProtocol(goals.protocol);

  const startFast = useCallback(async (customTs) => {
    const startTs = (typeof customTs === "number" && !isNaN(customTs)) ? Math.min(customTs, Date.now()) : Date.now();
    await saveCurrent({ startTs, targetHrs: protocol.hrs, protocol: protocol.id });
    flash(`Пішов відлік — ціль ${protocol.hrs} год 💪`);
  }, [saveCurrent, protocol, flash]);

  const setStartTs = useCallback(async (ts) => { if (current) await saveCurrent({ ...current, startTs: ts }); }, [current, saveCurrent]);

  const endFast = useCallback(async () => {
    if (!current) return;
    const actualHrs = Math.round(((Date.now() - current.startTs) / 3600000) * 10) / 10;
    const entry = {
      id: ruid("d"), date: dateKey(current.startTs), ts: current.startTs,
      protocol: current.protocol, targetHrs: current.targetHrs, actualHrs,
      goalMet: actualHrs + 0.05 >= current.targetHrs,
      weight: null, waist: null, energy: null, hunger: null, wellbeing: "", notes: "",
    };
    await saveDiary([...diary, entry]);
    await saveCurrent(null);
    setDiaryEditor({ entry });
    setFview("diary");
    flash(entry.goalMet ? "Ціль досягнута! Запис у щоденнику ✓" : "Голодування записано. Будь-який крок — це прогрес.");
  }, [current, diary, saveDiary, saveCurrent, flash]);

  const saveEntry = useCallback(async (entry, id) => {
    if (id) await saveDiary(diary.map((r) => (r.id === id ? { ...r, ...entry } : r)));
    else await saveDiary([...diary, { id: ruid("d"), ts: Date.now(), ...entry }]);
    setDiaryEditor(null); flash("Збережено");
  }, [diary, saveDiary, flash]);
  const deleteEntry = useCallback(async (id) => { await saveDiary(diary.filter((r) => r.id !== id)); flash("Видалено"); }, [diary, saveDiary, flash]);

  if (loading) return <div className="flex flex-1 items-center justify-center text-orange-400"><div className="flex flex-col items-center gap-3"><Hourglass className="h-8 w-8 animate-pulse" /><span className="text-sm">Завантаження…</span></div></div>;

  const NAV = [
    { id: "timer", label: "Таймер", icon: Timer },
    { id: "ladder", label: "Драбина", icon: TrendingUp },
    { id: "plan", label: "План", icon: Target },
    { id: "diary", label: "Щоденник", icon: NotebookPen },
    { id: "overview", label: "Огляд", icon: Scale },
    { id: "reference", label: "Довідник", icon: Info },
  ];

  return (
    <div className="min-h-screen flex-1 bg-gradient-to-b from-amber-50 via-orange-50/40 to-white">
      <header className="sticky top-0 z-20 border-b border-amber-100 bg-white/85 backdrop-blur">
        <div className="mx-auto flex h-14 w-full max-w-3xl items-center gap-1 px-4">
          {renaming ? (
            <input autoFocus value={nameDraft} onChange={(e) => setNameDraft(e.target.value)} onBlur={() => { onRename(nameDraft); setRenaming(false); }} onKeyDown={(e) => { if (e.key === "Enter") { onRename(nameDraft); setRenaming(false); } }} className="mr-auto w-32 rounded-lg border border-amber-200 px-2 py-1 text-base font-semibold focus:outline-none" />
          ) : (
            <button onClick={() => { setNameDraft(name); setRenaming(true); }} className="mr-auto text-base font-semibold text-slate-900">{name} <Pencil className="ml-0.5 inline h-3.5 w-3.5 text-slate-300" /></button>
          )}
          {NAV.map((n) => <NavButton key={n.id} active={fview === n.id} onClick={() => setFview(n.id)} icon={n.icon}>{n.label}</NavButton>)}
        </div>
      </header>

      <main className="mx-auto w-full max-w-3xl px-4 py-6">
        {fview === "timer" && <FastTimer current={current} now={now} protocol={protocol} onStart={startFast} onEnd={endFast} onSetStart={setStartTs} onChangeProtocol={() => setFview("ladder")} />}
        {fview === "ladder" && <ProtocolLadder goals={goals} diary={diary} onSaveGoals={saveGoals} onSet={(id) => { saveGoals({ protocol: id, protocolSince: dateKey(Date.now()), stepUpDismissed: null }); flash(`Протокол: ${getProtocol(id).label}`); setFview("timer"); }} />}
        {fview === "plan" && <FastPlan goals={goals} diary={diary} onSaveGoals={saveGoals} onGoLog={() => setDiaryEditor({ entry: null })} />}
        {fview === "diary" && <FastDiary diary={diary} onNew={() => setDiaryEditor({ entry: null })} onEdit={(e) => setDiaryEditor({ entry: e })} onDelete={deleteEntry} />}
        {fview === "overview" && <FastOverview goals={goals} diary={diary} onSaveGoals={saveGoals} />}
        {fview === "reference" && <FastReference />}
      </main>

      {diaryEditor && <DiaryForm entry={diaryEditor.entry} defaultProtocol={goals.protocol} onClose={() => setDiaryEditor(null)} onSave={(e) => saveEntry(e, diaryEditor.entry?.id)} />}
      {toast && <div className="fixed bottom-6 left-1/2 z-40 -translate-x-1/2 rounded-full bg-slate-900 px-4 py-2 text-sm text-white shadow-lg">{toast}</div>}
    </div>
  );
}

/* ---------- Live timer ---------- */
function FastRing({ pct, size = 240, stroke = 16, color = "#f97316", children }) {
  const r = (size - stroke) / 2, c = 2 * Math.PI * r, off = c * (1 - Math.max(0, Math.min(1, pct)));
  return (
    <div className="relative" style={{ width: size, height: size }}>
      <svg width={size} height={size} className="-rotate-90">
        <circle cx={size / 2} cy={size / 2} r={r} fill="none" stroke="#fdead1" strokeWidth={stroke} />
        <circle cx={size / 2} cy={size / 2} r={r} fill="none" stroke={color} strokeWidth={stroke} strokeDasharray={c} strokeDashoffset={off} strokeLinecap="round" style={{ transition: "stroke-dashoffset 1s linear" }} />
      </svg>
      <div className="absolute inset-0 flex flex-col items-center justify-center text-center">{children}</div>
    </div>
  );
}
function FastTimer({ current, now, protocol, onStart, onEnd, onSetStart, onChangeProtocol }) {
  const [editing, setEditing] = useState(false);
  const [startPick, setStartPick] = useState(false);
  const [customStart, setCustomStart] = useState("");
  const elapsedMs = current ? now - current.startTs : 0;
  const elapsedH = elapsedMs / 3600000;
  const targetH = current ? current.targetHrs : protocol.hrs;
  const pct = current ? elapsedMs / (targetH * 3600000) : 0;
  const stage = stageForHours(elapsedH);
  const projEnd = current ? new Date(current.startTs + targetH * 3600000) : null;
  const startDate = current ? new Date(current.startTs) : null;
  const p = current ? getProtocol(current.protocol) : protocol;

  return (
    <div className="flex flex-col items-center">
      <div className="mb-2 flex items-center gap-2">
        <span className="rounded-full px-3 py-1 text-sm font-bold text-white" style={{ backgroundColor: protocolColor(p.level) }}>{p.label}</span>
        <button onClick={onChangeProtocol} className="rounded-full bg-white px-3 py-1 text-sm font-medium text-slate-500 ring-1 ring-slate-200 hover:bg-slate-50">змінити</button>
      </div>

      <FastRing pct={pct} color={current ? stage.color : "#fb923c"}>
        <div className="text-xs font-semibold uppercase tracking-wide text-slate-400">{current ? "Голодування" : "Готова почати"}</div>
        <div className="text-4xl font-extrabold tabular-nums text-slate-900">{current ? fmtHMS(elapsedMs) : fmtHM(elapsedMs)}</div>
        <div className="mt-1 text-sm text-slate-400">ціль {targetH} год{current ? ` · ${(pct * 100).toFixed(1)}%` : ""}</div>
      </FastRing>

      {current ? (
        <>
          <div className="mt-4 flex flex-wrap items-center justify-center gap-2 text-sm text-slate-500">
            <span>Старт: {startDate.toLocaleString(undefined, { day: "numeric", month: "short", hour: "2-digit", minute: "2-digit" })}</span>
            <span className="text-slate-300">·</span>
            <span>Ціль: {projEnd.toLocaleString(undefined, { day: "numeric", month: "short", hour: "2-digit", minute: "2-digit" })}</span>
            <button onClick={() => setEditing((v) => !v)} className="text-orange-500 hover:text-orange-600"><Pencil className="h-3.5 w-3.5" /></button>
          </div>
          {editing && (
            <div className="mt-2 flex items-center gap-2 rounded-xl bg-white px-3 py-2 shadow-sm ring-1 ring-amber-100">
              <span className="text-xs text-slate-500">Час старту</span>
              <input type="datetime-local" value={toLocalInput(current.startTs)} onChange={(e) => { const t = new Date(e.target.value).getTime(); if (!isNaN(t)) onSetStart(t); }} className="rounded-lg border border-slate-300 px-2 py-1 text-sm" />
            </div>
          )}
          <button onClick={onEnd} className="mt-6 rounded-2xl bg-slate-800 px-10 py-3.5 font-bold text-white shadow-lg transition hover:bg-slate-900">Завершити голодування</button>
        </>
      ) : (
        <>
          <button onClick={() => onStart(startPick && customStart ? new Date(customStart).getTime() : undefined)} className="mt-6 rounded-2xl bg-orange-500 px-12 py-3.5 font-bold text-white shadow-lg shadow-orange-500/25 transition hover:bg-orange-600">Почати голодування</button>
          {!startPick ? (
            <button onClick={() => { setStartPick(true); setCustomStart(toLocalInput(Date.now())); }} className="mt-3 text-sm font-medium text-orange-500 hover:text-orange-600">Почала раніше? Вказати час старту</button>
          ) : (
            <div className="mt-3 flex flex-wrap items-center justify-center gap-2 rounded-xl bg-white px-3 py-2 shadow-sm ring-1 ring-amber-100">
              <span className="text-xs text-slate-500">Старт о</span>
              <input type="datetime-local" max={toLocalInput(Date.now())} value={customStart} onChange={(e) => setCustomStart(e.target.value)} className="rounded-lg border border-slate-300 px-2 py-1 text-sm" />
              <button onClick={() => setStartPick(false)} className="text-xs font-medium text-slate-400 hover:text-slate-600">зараз</button>
            </div>
          )}
        </>
      )}

      {/* stage timeline */}
      <div className="mt-8 w-full">
        <div className="mb-2 text-sm font-semibold text-slate-600">Що відбувається в тілі</div>
        <div className="space-y-2">
          {FAST_STAGES.map((s) => {
            const active = current && elapsedH >= s.from && elapsedH < s.to;
            const past = current && elapsedH >= s.to;
            const label = s.to >= 999 ? `${s.from}+ год` : `${s.from}–${s.to} год`;
            return (
              <div key={s.title} className={`flex items-center gap-3 rounded-2xl border p-3 transition ${active ? "border-transparent shadow-md" : "border-slate-100 bg-white"}`} style={active ? { backgroundColor: s.color + "18", borderColor: s.color } : {}}>
                <span className="grid h-9 w-9 shrink-0 place-items-center rounded-full text-white" style={{ backgroundColor: past ? "#cbd5e1" : s.color }}>{past ? <Check className="h-4 w-4" /> : <span className="text-[11px] font-bold">{s.from}</span>}</span>
                <div className="min-w-0 flex-1">
                  <div className="flex items-center gap-2"><span className="font-bold text-slate-800">{s.title}</span><span className="text-[11px] font-medium text-slate-400">{label}</span>{active && <span className="rounded-full bg-white/70 px-2 text-[10px] font-bold" style={{ color: s.color }}>зараз</span>}</div>
                  <div className="text-xs text-slate-500">{s.desc}</div>
                </div>
              </div>
            );
          })}
        </div>
        <p className="mt-3 text-center text-[11px] text-slate-400">Орієнтовні етапи для розуміння процесу — не медичні твердження. Час індивідуальний.</p>
      </div>
    </div>
  );
}
function toLocalInput(ts) { const d = new Date(ts - new Date().getTimezoneOffset() * 60000); return d.toISOString().slice(0, 16); }

/* ---------- Protocol ladder ---------- */
function ProtocolLadder({ goals, diary, onSet, onSaveGoals }) {
  const cur = getProtocol(goals.protocol);
  const nextP = PROTOCOLS.find((p) => p.level === cur.level + 1);
  const stepUpDays = goals.stepUpDays || 14;
  const since = goals.protocolSince || goals.startDate || dateKey(Date.now());
  // consistency: distinct days you actually completed this protocol since you settled on it
  const daysDone = new Set(diary.filter((r) => r.protocol === goals.protocol && r.goalMet && r.date >= since).map((r) => r.date)).size;
  const remaining = Math.max(0, stepUpDays - daysDone);
  const reached = daysDone >= stepUpDays && !!nextP;
  const dismissed = goals.stepUpDismissed === goals.protocol;
  const pctToStep = Math.min(1, stepUpDays ? daysDone / stepUpDays : 0);
  const highlightNext = reached && !dismissed && nextP;
  const setStepDays = (d) => onSaveGoals({ stepUpDays: Math.max(3, Math.min(60, d)) });

  return (
    <div>
      <h1 className="text-xl font-extrabold text-slate-900">Драбина протоколів</h1>
      <p className="mt-1 rounded-2xl bg-amber-50 px-3 py-2 text-sm leading-relaxed text-amber-800">
        Починай з найм'якшого щабля й піднімайся <b>поступово</b> — тільки коли поточний рівень дається легко й приємно. Немає «єдино правильного» рівня; сталість важливіша за інтенсивність. Жодного поспіху.
      </p>

      {/* recommendation */}
      {nextP ? (
        highlightNext ? (
          <div className="mt-3 rounded-2xl border border-green-200 bg-green-50 p-4">
            <div className="flex items-start gap-3">
              <Sparkle className="mt-0.5 h-5 w-5 shrink-0 text-green-500" />
              <div className="flex-1">
                <p className="text-sm text-green-800">Ти освоїла <b>{cur.label}</b> — вже {daysDone} {daysDone === 1 ? "день" : daysDone < 5 ? "дні" : "днів"} на цьому рівні. Коли відчуєш готовність, можна спробувати <b>{nextP.label}</b>. Без поспіху 💛</p>
                <div className="mt-3 flex flex-wrap gap-2">
                  <button onClick={() => onSet(nextP.id)} className="rounded-full bg-green-600 px-4 py-1.5 text-xs font-semibold text-white hover:bg-green-700">Спробувати {nextP.label}</button>
                  <button onClick={() => onSaveGoals({ stepUpDismissed: goals.protocol })} className="rounded-full bg-white px-4 py-1.5 text-xs font-semibold text-slate-500 ring-1 ring-slate-200 hover:bg-slate-50">Поки лишити {cur.label}</button>
                </div>
              </div>
            </div>
          </div>
        ) : (
          <div className="mt-3 rounded-2xl border border-slate-200 bg-white p-4">
            <div className="flex items-center justify-between gap-2">
              <p className="text-sm text-slate-600">Ти на <b>{cur.label}</b> вже <b>{daysDone}</b> {daysDone === 1 ? "день" : daysDone < 5 ? "дні" : "днів"}.{dismissed ? " Лишаєшся на цьому рівні 👍" : ` Орієнтир: ще ~${remaining} дн на цьому рівні, перш ніж пробувати ${nextP.label}.`}</p>
            </div>
            <div className="mt-2 h-2 overflow-hidden rounded-full bg-slate-100"><div className="h-full rounded-full transition-all" style={{ width: `${pctToStep * 100}%`, backgroundColor: protocolColor(cur.level) }} /></div>
            <div className="mt-2 flex items-center gap-2 text-xs text-slate-400">
              <span>Орієнтир переходу:</span>
              <button onClick={() => setStepDays(stepUpDays - 1)} className="grid h-6 w-6 place-items-center rounded-full bg-slate-100 font-bold text-slate-500 hover:bg-slate-200">−</button>
              <span className="font-semibold tabular-nums text-slate-600">{stepUpDays} дн</span>
              <button onClick={() => setStepDays(stepUpDays + 1)} className="grid h-6 w-6 place-items-center rounded-full bg-slate-100 font-bold text-slate-500 hover:bg-slate-200">+</button>
              {dismissed && <button onClick={() => onSaveGoals({ stepUpDismissed: null })} className="ml-auto font-medium text-slate-400 underline hover:text-slate-600">повернути підказку</button>}
            </div>
          </div>
        )
      ) : (
        <div className="mt-3 rounded-2xl border border-slate-200 bg-white px-4 py-3 text-sm text-slate-500">Ти на найвищому щаблі драбини. Слухай своє тіло й лишайся на комфортному рівні 💛</div>
      )}
      <p className="mt-2 px-1 text-[11px] leading-relaxed text-slate-400">Це загальний орієнтир, а не медична порада. Немає єдино правильного темпу — рухайся вгору лише коли готова, і слухай своє тіло.</p>

      <div className="mt-4 space-y-2">
        {[...PROTOCOLS].reverse().map((p) => {
          const active = p.id === goals.protocol;
          const isNext = highlightNext && p.id === nextP.id;
          const color = protocolColor(p.level);
          return (
            <div key={p.id} className={`flex items-start gap-3 rounded-2xl border p-4 transition ${active ? "shadow-md" : isNext ? "border-green-300 bg-green-50/60 ring-2 ring-green-200" : "border-slate-100 bg-white"}`} style={active ? { borderColor: color, backgroundColor: color + "12" } : {}}>
              <span className="grid h-11 w-11 shrink-0 place-items-center rounded-xl text-sm font-extrabold text-white" style={{ backgroundColor: color }}>{p.hrs}г</span>
              <div className="min-w-0 flex-1">
                <div className="flex flex-wrap items-center gap-2">
                  <span className="font-bold text-slate-900">{p.label}</span>
                  <span className="text-xs text-slate-400">вікно {p.window} год · {p.freq}</span>
                  {active && <span className="rounded-full px-2 py-0.5 text-[10px] font-bold text-white" style={{ backgroundColor: color }}>мій зараз</span>}
                  {isNext && <span className="rounded-full bg-green-500 px-2 py-0.5 text-[10px] font-bold text-white">рекомендовано</span>}
                </div>
                <p className="mt-0.5 text-sm text-slate-500">{p.note}</p>
              </div>
              {!active && <button onClick={() => onSet(p.id)} className={`shrink-0 rounded-full px-3 py-1.5 text-xs font-semibold text-white ${isNext ? "bg-green-600 hover:bg-green-700" : "bg-slate-800 hover:bg-slate-900"}`}>Обрати</button>}
            </div>
          );
        })}
      </div>
    </div>
  );
}

/* ---------- Overview: goals + metrics + charts ---------- */
/* ---------- Fasting: gradual weight-loss plan ---------- */
function FastPlan({ goals, diary, onSaveGoals, onGoLog }) {
  const num = (v) => { const n = parseFloat(v); return isNaN(n) ? null : n; };
  const months = goals.planMonths || 12;
  const startW = goals.startWeight;
  const targetW = goals.targetWeight;
  const startDate = goals.planStart || goals.startDate || dateKey(Date.now());
  const weighed = (diary || []).filter((r) => r.weight != null).sort((a, b) => a.date.localeCompare(b.date));
  const currentW = weighed.length ? weighed[weighed.length - 1].weight : startW;

  const ready = startW != null && targetW != null && startW > targetW;
  const totalLoss = ready ? startW - targetW : 0;
  const perMonth = ready ? totalLoss / months : 0;
  const perWeek = ready ? totalLoss / (months * 4.345) : 0;
  const lostSoFar = (startW != null && currentW != null) ? Math.round((startW - currentW) * 10) / 10 : 0;
  const pctDone = totalLoss > 0 ? Math.max(0, Math.min(1, lostSoFar / totalLoss)) : 0;

  // elapsed months since start → expected weight now
  const elapsedMs = Date.now() - new Date(startDate + "T00:00:00").getTime();
  const elapsedMonths = Math.max(0, elapsedMs / (30.44 * 86400000));
  const expectedNow = ready ? Math.max(targetW, startW - perMonth * elapsedMonths) : null;
  const delta = (ready && currentW != null && expectedNow != null) ? Math.round((currentW - expectedNow) * 10) / 10 : null; // >0 = behind plan

  const milestones = [];
  if (ready) {
    const d0 = new Date(startDate + "T00:00:00");
    for (let i = 1; i <= months; i++) { const d = new Date(d0); d.setMonth(d.getMonth() + i); milestones.push({ i, w: Math.round((startW - perMonth * i) * 10) / 10, date: `${d.getMonth() + 1}.${d.getFullYear()}` }); }
    if (milestones.length) milestones[milestones.length - 1].w = targetW;
  }
  const paceOk = perWeek > 0 && perWeek <= 1;

  // ---- КБЖУ + activity calculator (Mifflin-St Jeor BMR → TDEE → deficit) ----
  const h = goals.heightCm, age = goals.age, sex = goals.sex || "f";
  const activity = goals.activity || 1.375;
  const w = currentW ?? startW;
  const canKcal = w != null && h != null && age != null;
  const bmr = canKcal ? Math.round(10 * w + 6.25 * h - 5 * age + (sex === "m" ? 5 : -161)) : null;
  const tdee = bmr != null ? Math.round(bmr * activity) : null;
  const dailyDeficit = ready ? Math.round((perWeek * 7700) / 7) : 0; // 1 кг жиру ≈ 7700 ккал
  const floorKcal = sex === "m" ? 1500 : 1200;
  const rawTarget = tdee != null ? tdee - dailyDeficit : null;
  const kcalTarget = rawTarget != null ? Math.max(floorKcal, rawTarget) : null;
  const belowFloor = rawTarget != null && rawTarget < floorKcal;
  const refW = targetW || w; // білок/жир рахуємо від цільової ваги
  const proteinG = refW != null ? Math.round(1.8 * refW) : null;
  const fatG = refW != null ? Math.max(40, Math.round(0.8 * refW)) : null;
  const carbsG = (kcalTarget != null && proteinG != null && fatG != null) ? Math.max(30, Math.round((kcalTarget - proteinG * 4 - fatG * 9) / 4)) : null;
  const waterL = w != null ? Math.round((w * 33) / 100) / 10 : null;
  const macro = (g, kcalPer) => (kcalTarget ? Math.round((g * kcalPer / kcalTarget) * 100) : 0);
  const ACTS = [[1.2, "Сидячий"], [1.375, "Легкий (1–3 трен.)"], [1.55, "Помірний (3–5)"], [1.725, "Активний (6+)"]];

  const TIPS = [
    { emoji: "⚖️", title: "Темп", body: "≈0.5–0.7 кг/тиждень — стало й безпечно. Швидше = більше втрати м'язів і гірше підтягується шкіра. Повільніше — краще для тіла й шкіри." },
    { emoji: "🔥", title: "Голодування", body: "Піднімайся драбиною протоколів поступово (16:8 → 18:6 …), без форсування. Щоденне 16:8 уже добре працює — сталість важливіша за екстрим." },
    { emoji: "💪", title: "Рух — головне для тонусу", body: "Силові 2–3 рази/тиждень — саме вони дають підкачане тіло й допомагають шкірі підтягнутись (м'язи заповнюють об'єм). Плюс 8–10 тис кроків на день." },
    { emoji: "🍽️", title: "Харчування", body: "Достатньо білка (≈1.6–2 г на кг ваги — зберігає м'язи), овочі, вода 2+ л. Дефіцит помірний, не «голод до нуля» — інакше тіло їсть м'язи." },
    { emoji: "🧴", title: "Щоб шкіра втягнулась", body: "Повільна втрата + білок + силові + вода + сон + час. Колаген/вітамін C. Нікотин руйнує колаген — це ще одна причина кидати (див. Відновлення). Еластичність повертається місяцями." },
  ];

  return (
    <div>
      <div className="rounded-3xl bg-gradient-to-br from-orange-400 to-red-400 p-5 text-white shadow-sm">
        <div className="text-sm font-semibold text-white/90">План — плавно й надовго</div>
        <div className="text-2xl font-extrabold">−{totalLoss || 30} кг за {months} міс</div>
        {ready ? <div className="text-sm text-white/90">≈ {perMonth.toFixed(1)} кг/міс · {perWeek.toFixed(2)} кг/тиждень {paceOk ? "✓ безпечний темп" : "⚠️ швидкувато"}</div>
          : <div className="text-sm text-white/90">Впиши стартову й цільову вагу нижче — і зʼявиться графік по місяцях.</div>}
      </div>

      {/* weight goals */}
      <div className="mt-3 grid grid-cols-3 gap-2">
        <label className="block rounded-2xl bg-white p-3 shadow-sm ring-1 ring-orange-50"><span className="mb-1 block text-[11px] text-slate-400">Старт, кг</span><input type="number" step="0.1" value={goals.startWeight ?? ""} onChange={(e) => onSaveGoals({ startWeight: num(e.target.value) })} className="w-full rounded-lg border border-slate-300 px-2 py-1.5 text-sm focus:border-orange-400 focus:outline-none" /></label>
        <label className="block rounded-2xl bg-white p-3 shadow-sm ring-1 ring-orange-50"><span className="mb-1 block text-[11px] text-slate-400">Ціль, кг</span><input type="number" step="0.1" value={goals.targetWeight ?? ""} onChange={(e) => onSaveGoals({ targetWeight: num(e.target.value) })} placeholder={startW != null ? String(Math.round(startW - 30)) : ""} className="w-full rounded-lg border border-slate-300 px-2 py-1.5 text-sm focus:border-orange-400 focus:outline-none" /></label>
        <label className="block rounded-2xl bg-white p-3 shadow-sm ring-1 ring-orange-50"><span className="mb-1 block text-[11px] text-slate-400">Місяців</span><input type="number" min={3} max={24} value={months} onChange={(e) => onSaveGoals({ planMonths: Math.max(3, Math.min(24, +e.target.value || 12)) })} className="w-full rounded-lg border border-slate-300 px-2 py-1.5 text-sm focus:border-orange-400 focus:outline-none" /></label>
      </div>

      {/* personal data for the calorie calc */}
      <div className="mt-3 rounded-2xl bg-white p-4 shadow-sm ring-1 ring-orange-50">
        <div className="mb-2 text-sm font-bold text-slate-700">Твої дані <span className="font-normal text-slate-400">— щоб порахувати КБЖУ</span></div>
        <div className="grid grid-cols-3 gap-2">
          <label className="block"><span className="mb-1 block text-[11px] text-slate-400">Зріст, см</span><input type="number" value={goals.heightCm ?? ""} onChange={(e) => onSaveGoals({ heightCm: num(e.target.value) })} placeholder="напр. 165" className="w-full rounded-lg border border-slate-300 px-2 py-1.5 text-sm focus:border-orange-400 focus:outline-none" /></label>
          <label className="block"><span className="mb-1 block text-[11px] text-slate-400">Вік</span><input type="number" value={goals.age ?? ""} onChange={(e) => onSaveGoals({ age: num(e.target.value) })} placeholder="напр. 30" className="w-full rounded-lg border border-slate-300 px-2 py-1.5 text-sm focus:border-orange-400 focus:outline-none" /></label>
          <div><span className="mb-1 block text-[11px] text-slate-400">Стать</span><div className="flex gap-1">{[["f", "Ж"], ["m", "Ч"]].map(([v, l]) => <button key={v} onClick={() => onSaveGoals({ sex: v })} className={`flex-1 rounded-lg py-1.5 text-sm font-semibold transition ${sex === v ? "bg-orange-500 text-white" : "bg-slate-100 text-slate-500"}`}>{l}</button>)}</div></div>
        </div>
        <div className="mt-2"><span className="mb-1 block text-[11px] text-slate-400">Активність</span><div className="flex flex-wrap gap-1.5">{ACTS.map(([v, l]) => <button key={v} onClick={() => onSaveGoals({ activity: v })} className={`rounded-full px-2.5 py-1 text-xs font-semibold ring-1 transition ${Math.abs(activity - v) < 0.01 ? "bg-orange-500 text-white ring-orange-500" : "bg-white text-slate-500 ring-slate-200"}`}>{l}</button>)}</div></div>
      </div>

      {/* КБЖУ per day */}
      {canKcal && ready ? (
        <div className="mt-3 rounded-2xl bg-white p-4 shadow-sm ring-1 ring-orange-100">
          <div className="flex items-end justify-between">
            <div><div className="text-sm font-bold text-slate-700">Скільки їсти щодня</div><div className="text-[11px] text-slate-400">щоб втрачати ≈ {perWeek.toFixed(2)} кг/тиждень</div></div>
            <div className="text-right"><div className="text-3xl font-extrabold tabular-nums text-orange-600">{kcalTarget}</div><div className="text-[11px] text-slate-400">ккал / день</div></div>
          </div>
          <div className="mt-3 grid grid-cols-3 gap-2 text-center">
            {[["Білки", proteinG, macro(proteinG, 4), "bg-rose-50 text-rose-600"], ["Жири", fatG, macro(fatG, 9), "bg-amber-50 text-amber-600"], ["Вуглеводи", carbsG, macro(carbsG, 4), "bg-sky-50 text-sky-600"]].map(([l, g, pct, cls]) => (
              <div key={l} className={`rounded-2xl p-3 ${cls}`}><div className="text-[11px] font-medium opacity-80">{l}</div><div className="text-lg font-extrabold tabular-nums">{g} г</div><div className="text-[10px] opacity-70">{pct}%</div></div>
            ))}
          </div>
          <div className="mt-3 flex flex-wrap gap-2 text-xs">
            <span className="rounded-full bg-slate-100 px-2.5 py-1 font-semibold text-slate-600">BMR ≈ {bmr}</span>
            <span className="rounded-full bg-slate-100 px-2.5 py-1 font-semibold text-slate-600">Витрата ≈ {tdee}</span>
            <span className="rounded-full bg-emerald-50 px-2.5 py-1 font-semibold text-emerald-700">Дефіцит −{dailyDeficit}</span>
            <span className="rounded-full bg-sky-50 px-2.5 py-1 font-semibold text-sky-700">💧 вода ≈ {waterL} л</span>
          </div>
          {belowFloor && <p className="mt-3 rounded-xl bg-amber-50 px-3 py-2 text-[11px] leading-relaxed text-amber-800">⚠️ Для такого темпу калорій вийшло б менше безпечного мінімуму ({floorKcal} ккал), тож я підняла до {floorKcal}. Щоб не голодувати — краще розтягнути план на більше місяців (втрата буде трохи повільніша, але здоровіша).</p>}
          <p className="mt-2 text-[11px] leading-relaxed text-slate-400">Білок високий, щоб зберегти м'язи; решта калорій — вуглеводи й жири. Овочі та клітковина — понад норму, їх не рахуємо жорстко.</p>
        </div>
      ) : ready ? (
        <div className="mt-3 rounded-2xl bg-orange-50/70 p-4 text-center text-sm text-orange-700 ring-1 ring-orange-100">Впиши зріст, вік і стать вище — і я порахую твої калорії та БЖУ на день. 🍽️</div>
      ) : null}

      {/* activity targets */}
      {ready && (
        <div className="mt-3 grid grid-cols-2 gap-2">
          <div className="rounded-2xl bg-white p-4 shadow-sm ring-1 ring-orange-50 text-center"><div className="text-2xl">🚶‍♀️</div><div className="mt-1 text-xl font-extrabold text-slate-800">8–10 тис</div><div className="text-[11px] text-slate-400">кроків на день</div></div>
          <div className="rounded-2xl bg-white p-4 shadow-sm ring-1 ring-orange-50 text-center"><div className="text-2xl">💪</div><div className="mt-1 text-xl font-extrabold text-slate-800">2–3</div><div className="text-[11px] text-slate-400">силові / тиждень (+ ходьба)</div></div>
        </div>
      )}

      {/* fasting integration */}
      {ready && (
        <div className="mt-3 rounded-2xl bg-gradient-to-br from-orange-50 to-white p-4 shadow-sm ring-1 ring-orange-100">
          <div className="flex items-center gap-2 text-sm font-bold text-slate-800"><Hourglass className="h-4 w-4 text-orange-500" /> Як вписати в голодування</div>
          <p className="mt-1 text-sm leading-relaxed text-slate-600">На 16:8 усі {kcalTarget || "твої"} ккал з'їдай у вікні 8 год — зазвичай 2–3 прийоми. Почни їжу з білка й овочів (ситніше). Поза вікном — вода, чай, кава без цукру. Дефіцит створюєш калоріями, а голодування лише допомагає легше в нього вкластись — не «голодуй + майже не їж», це забагато.</p>
        </div>
      )}

      {ready && (
        <div className="mt-3 rounded-2xl bg-white p-4 shadow-sm ring-1 ring-orange-100">
          <div className="flex items-center justify-between text-sm"><span className="font-semibold text-slate-700">Зараз {currentW} кг · скинуто {lostSoFar} кг з {totalLoss}</span>{delta != null && <span className={`font-bold ${delta <= 0.5 ? "text-green-600" : "text-amber-600"}`}>{delta <= 0.5 ? "у графіку ✓" : `+${delta} кг до плану`}</span>}</div>
          <div className="mt-2 h-2.5 overflow-hidden rounded-full bg-orange-50"><div className="h-full rounded-full bg-gradient-to-r from-orange-400 to-red-400 transition-all" style={{ width: `${pctDone * 100}%` }} /></div>
          <button onClick={onGoLog} className="mt-2 text-xs font-semibold text-orange-600">+ записати вагу сьогодні</button>
        </div>
      )}

      {ready && milestones.length > 0 && (
        <div className="mt-3 rounded-2xl bg-white p-4 shadow-sm ring-1 ring-orange-100">
          <div className="mb-2 text-sm font-bold text-slate-700">Орієнтири по місяцях</div>
          <div className="grid grid-cols-2 gap-1.5 sm:grid-cols-3">
            {milestones.map((m) => { const hit = currentW != null && currentW <= m.w; return (
              <div key={m.i} className={`rounded-xl px-3 py-2 text-center ${hit ? "bg-green-50 ring-1 ring-green-200" : "bg-slate-50"}`}>
                <div className="text-[10px] text-slate-400">міс {m.i} · {m.date}</div>
                <div className={`text-sm font-extrabold tabular-nums ${hit ? "text-green-600" : "text-slate-700"}`}>{m.w} кг</div>
              </div>
            ); })}
          </div>
        </div>
      )}

      {/* guidance */}
      <div className="mt-3 space-y-2">
        {TIPS.map((t) => (
          <div key={t.title} className="rounded-2xl bg-white p-4 shadow-sm ring-1 ring-orange-50">
            <div className="flex items-center gap-2 text-sm font-bold text-slate-800"><span className="text-lg">{t.emoji}</span> {t.title}</div>
            <p className="mt-1 text-sm leading-relaxed text-slate-600">{t.body}</p>
          </div>
        ))}
      </div>

      <p className="mt-4 rounded-2xl bg-amber-50 px-4 py-3 text-xs leading-relaxed text-amber-800">
        ⚠️ Це загальні орієнтири, а не медична порада. 30 кг — суттєва зміна, тож варто йти під наглядом лікаря чи дієтолога, особливо з голодуванням. Слухай тіло: запаморочення, слабкість, випадіння волосся — сигнал сповільнитись. Ціль не «швидко», а щоб було стало, здорово й із гарним самопочуттям. 💛
      </p>
    </div>
  );
}

function FastOverview({ goals, diary, onSaveGoals }) {
  const m = fastingMetrics(goals, diary);
  const rows = diarySorted(diary);
  const weightData = rows.filter((r) => r.weight != null).map((r) => ({ date: r.date.slice(5), weight: r.weight }));
  const hoursData = rows.filter((r) => r.actualHrs != null).slice(-14).map((r) => ({ date: r.date.slice(5), hrs: r.actualHrs, met: r.goalMet }));
  const streak = fastingStreak(diary);

  const start = goals.startWeight, target = goals.targetWeight, cur = m.currentWeight;
  let pct = 0;
  if (start != null && target != null && cur != null && start !== target) pct = Math.max(0, Math.min(1, (start - cur) / (start - target)));

  const num = (v) => (v == null || v === "" ? null : parseFloat(v));

  return (
    <div className="space-y-5">
      <h1 className="text-xl font-extrabold text-slate-900">Огляд</h1>

      {/* goals setup */}
      <div className="rounded-2xl bg-white p-4 shadow-sm ring-1 ring-amber-100">
        <div className="mb-3 text-sm font-bold text-slate-700">🎯 Мої цілі</div>
        <div className="grid grid-cols-2 gap-3">
          <label className="block"><span className="mb-1 block text-xs text-slate-500">Стартова вага, кг</span><input type="number" step="0.1" value={goals.startWeight ?? ""} onChange={(e) => onSaveGoals({ startWeight: num(e.target.value) })} className="w-full rounded-lg border border-slate-300 px-3 py-2 text-sm focus:border-orange-400 focus:outline-none" /></label>
          <label className="block"><span className="mb-1 block text-xs text-slate-500">Цільова вага, кг</span><input type="number" step="0.1" value={goals.targetWeight ?? ""} onChange={(e) => onSaveGoals({ targetWeight: num(e.target.value) })} className="w-full rounded-lg border border-slate-300 px-3 py-2 text-sm focus:border-orange-400 focus:outline-none" /></label>
          <label className="block"><span className="mb-1 block text-xs text-slate-500">Основний протокол</span>
            <select value={goals.protocol} onChange={(e) => onSaveGoals({ protocol: e.target.value })} className="w-full rounded-lg border border-slate-300 px-3 py-2 text-sm focus:border-orange-400 focus:outline-none">{PROTOCOLS.map((p) => <option key={p.id} value={p.id}>{p.label}</option>)}</select>
          </label>
          <label className="block"><span className="mb-1 block text-xs text-slate-500">Дата старту</span><input type="date" value={goals.startDate || ""} onChange={(e) => onSaveGoals({ startDate: e.target.value })} className="w-full rounded-lg border border-slate-300 px-3 py-2 text-sm focus:border-orange-400 focus:outline-none" /></label>
        </div>
        {start != null && target != null && (
          <div className="mt-4">
            <div className="mb-1 flex justify-between text-xs text-slate-500"><span>{start} кг</span><span className="font-semibold text-orange-600">{cur ?? start} кг</span><span>{target} кг</span></div>
            <div className="h-3 overflow-hidden rounded-full bg-slate-100"><div className="h-full rounded-full bg-gradient-to-r from-orange-400 to-green-400 transition-all" style={{ width: `${pct * 100}%` }} /></div>
            <div className="mt-1 text-center text-xs text-slate-400">{Math.round(pct * 100)}% шляху до цілі</div>
          </div>
        )}
      </div>

      {/* metric cards */}
      <div className="grid grid-cols-2 gap-3 sm:grid-cols-3">
        <Metric label="Поточна вага" value={m.currentWeight != null ? `${m.currentWeight}` : "—"} unit="кг" />
        <Metric label="Зміна ваги" value={m.weightChange != null ? `${m.weightChange > 0 ? "+" : ""}${m.weightChange}` : "—"} unit="кг" tint={m.weightChange != null && m.weightChange < 0 ? "text-green-600" : "text-slate-700"} />
        <Metric label="До цілі" value={m.remaining != null ? `${m.remaining}` : "—"} unit="кг" />
        <Metric label="Всього голодувань" value={m.totalFasts} />
        <Metric label="Середня тривалість" value={m.avgHrs} unit="год" />
        <Metric label="Найдовше" value={m.longestHrs} unit="год" />
        <Metric label="Всього годин" value={m.totalHrs} unit="год" />
        <Metric label="Серія" value={streak} unit="дн" tint="text-orange-500" />
      </div>

      {/* weight chart */}
      <div className="rounded-2xl bg-white p-5 shadow-sm ring-1 ring-amber-100">
        <h2 className="mb-3 text-sm font-bold text-slate-700">Вага в часі</h2>
        {weightData.length < 2 ? <p className="py-8 text-center text-sm text-slate-400">Додай кілька записів ваги у щоденнику.</p> : (
          <div className="h-56 w-full"><ResponsiveContainer width="100%" height="100%">
            <LineChart data={weightData} margin={{ top: 4, right: 8, left: -16, bottom: 0 }}>
              <CartesianGrid strokeDasharray="3 3" stroke="#fef3e2" vertical={false} />
              <XAxis dataKey="date" tick={{ fontSize: 11, fill: "#b08968" }} axisLine={false} tickLine={false} />
              <YAxis domain={["dataMin - 1", "dataMax + 1"]} tick={{ fontSize: 11, fill: "#b08968" }} axisLine={false} tickLine={false} />
              <Tooltip contentStyle={{ borderRadius: 10, border: "1px solid #fed7aa", fontSize: 12 }} />
              <Line type="monotone" dataKey="weight" stroke="#f97316" strokeWidth={2.5} dot={{ r: 3, fill: "#f97316" }} />
            </LineChart>
          </ResponsiveContainer></div>
        )}
      </div>

      {/* hours chart */}
      <div className="rounded-2xl bg-white p-5 shadow-sm ring-1 ring-amber-100">
        <h2 className="mb-3 text-sm font-bold text-slate-700">Години голодування</h2>
        {hoursData.length === 0 ? <p className="py-8 text-center text-sm text-slate-400">Записів ще немає.</p> : (
          <div className="h-52 w-full"><ResponsiveContainer width="100%" height="100%">
            <BarChart data={hoursData} margin={{ top: 4, right: 4, left: -20, bottom: 0 }}>
              <CartesianGrid strokeDasharray="3 3" stroke="#fef3e2" vertical={false} />
              <XAxis dataKey="date" tick={{ fontSize: 11, fill: "#b08968" }} axisLine={false} tickLine={false} />
              <YAxis tick={{ fontSize: 11, fill: "#b08968" }} axisLine={false} tickLine={false} />
              <Tooltip cursor={{ fill: "#fff7ed" }} contentStyle={{ borderRadius: 10, border: "1px solid #fed7aa", fontSize: 12 }} formatter={(v) => [`${v} год`, "факт"]} />
              <Bar dataKey="hrs" radius={[5, 5, 0, 0]}>{hoursData.map((d, i) => <Cell key={i} fill={d.met ? "#22c55e" : "#fb923c"} />)}</Bar>
            </BarChart>
          </ResponsiveContainer></div>
        )}
      </div>

      <FastHeatmap diary={diary} />
    </div>
  );
}
function Metric({ label, value, unit, tint = "text-slate-800" }) {
  return <div className="rounded-2xl bg-white p-3 text-center shadow-sm ring-1 ring-amber-100"><div className={`text-xl font-extrabold tabular-nums ${tint}`}>{value}{unit && <span className="ml-0.5 text-xs font-medium text-slate-400">{unit}</span>}</div><div className="mt-0.5 text-[11px] text-slate-400">{label}</div></div>;
}
function FastHeatmap({ diary }) {
  const [mo, setMo] = useState(0);
  const now = new Date(); const view = new Date(now.getFullYear(), now.getMonth() + mo, 1);
  const year = view.getFullYear(), month = view.getMonth();
  const days = new Date(year, month + 1, 0).getDate();
  const pad = (new Date(year, month, 1).getDay() + 6) % 7;
  const byDate = {}; for (const r of diary) if (r.actualHrs) byDate[r.date] = Math.max(byDate[r.date] || 0, r.goalMet ? 2 : 1);
  const col = (v) => (!v ? "#f5f0e8" : v === 2 ? "#22c55e" : "#fdba74");
  return (
    <div className="rounded-2xl bg-white p-5 shadow-sm ring-1 ring-amber-100">
      <div className="mb-3 flex items-center justify-between">
        <h2 className="text-sm font-bold text-slate-700">{view.toLocaleDateString(undefined, { month: "long", year: "numeric" })}</h2>
        <div className="flex gap-1"><button onClick={() => setMo((x) => x - 1)} className="rounded-md p-1.5 text-slate-400 hover:bg-slate-100"><ArrowLeft className="h-4 w-4" /></button><button onClick={() => setMo(0)} disabled={mo === 0} className="rounded-md px-2 py-1 text-xs text-slate-500 hover:bg-slate-100 disabled:opacity-40">Зараз</button><button onClick={() => setMo((x) => Math.min(0, x + 1))} disabled={mo === 0} className="rounded-md p-1.5 text-slate-400 hover:bg-slate-100 disabled:opacity-40"><ArrowRight className="h-4 w-4" /></button></div>
      </div>
      <div className="grid grid-cols-7 gap-1.5">
        {["Пн", "Вт", "Ср", "Чт", "Пт", "Сб", "Нд"].map((w, i) => <div key={i} className="pb-1 text-center text-[10px] font-semibold text-slate-400">{w}</div>)}
        {Array.from({ length: pad }).map((_, i) => <div key={"p" + i} />)}
        {Array.from({ length: days }).map((_, i) => { const d = i + 1; const ds = dateKey(new Date(year, month, d).getTime()); const v = byDate[ds] || 0; return <div key={ds} className="grid aspect-square place-items-center rounded-lg text-[10px] font-medium" style={{ backgroundColor: col(v), color: v ? "#fff" : "#94a3b8" }} title={ds}>{d}</div>; })}
      </div>
    </div>
  );
}

/* ---------- Diary ---------- */
function FastDiary({ diary, onNew, onEdit, onDelete }) {
  const rows = diarySorted(diary).reverse();
  const dow = (ds) => WD_UA[new Date(ds + "T00:00:00").getDay()];
  return (
    <div>
      <div className="mb-3 flex items-center justify-between">
        <h1 className="text-xl font-extrabold text-slate-900">Щоденник</h1>
        <button onClick={onNew} className="inline-flex items-center gap-1.5 rounded-full bg-orange-500 px-4 py-2 text-sm font-semibold text-white hover:bg-orange-600"><Plus className="h-4 w-4" /> Запис</button>
      </div>
      {rows.length === 0 ? (
        <div className="rounded-2xl bg-white py-12 text-center text-sm text-slate-400 ring-1 ring-amber-50">Записів ще немає. Заверши голодування на таймері — і рядок з'явиться сам.</div>
      ) : (
        <div className="space-y-2">
          {rows.map((r) => (
            <button key={r.id} onClick={() => onEdit(r)} className="flex w-full items-center gap-3 rounded-2xl bg-white p-3 text-left shadow-sm ring-1 ring-amber-50 transition hover:ring-amber-200">
              <div className="w-14 shrink-0 text-center"><div className="text-xs font-semibold text-slate-400">{dow(r.date)}</div><div className="text-sm font-bold text-slate-700">{r.date.slice(5)}</div></div>
              <span className="rounded-full px-2 py-0.5 text-[11px] font-bold text-white" style={{ backgroundColor: protocolColor(getProtocol(r.protocol).level) }}>{getProtocol(r.protocol).label}</span>
              <div className="min-w-0 flex-1 text-sm">
                <span className="font-semibold text-slate-800">{r.actualHrs ?? "—"}</span><span className="text-slate-400">/{r.targetHrs ?? "—"} год</span>
                {r.goalMet ? <Check className="ml-1 inline h-4 w-4 text-green-500" /> : <span className="ml-1 text-slate-300">✗</span>}
                {r.weight != null && <span className="ml-2 text-slate-500">{r.weight} кг{r.weightChange != null ? ` (${r.weightChange > 0 ? "+" : ""}${r.weightChange})` : ""}</span>}
                {r.notes && <div className="truncate text-xs text-slate-400">{r.notes}</div>}
              </div>
              <span onClick={(e) => { e.stopPropagation(); if (confirm("Видалити запис?")) onDelete(r.id); }} className="rounded p-1 text-slate-300 hover:text-red-500"><Trash2 className="h-4 w-4" /></span>
            </button>
          ))}
        </div>
      )}
    </div>
  );
}
function DiaryForm({ entry, defaultProtocol, onClose, onSave }) {
  const [f, setF] = useState({
    date: entry?.date || dateKey(Date.now()), protocol: entry?.protocol || defaultProtocol || "16:8",
    targetHrs: entry?.targetHrs ?? getProtocol(entry?.protocol || defaultProtocol || "16:8").hrs,
    actualHrs: entry?.actualHrs ?? "", weight: entry?.weight ?? "", waist: entry?.waist ?? "",
    energy: entry?.energy ?? "", hunger: entry?.hunger ?? "", wellbeing: entry?.wellbeing || "", notes: entry?.notes || "",
  });
  const set = (k, v) => setF((s) => ({ ...s, [k]: v }));
  const num = (v) => (v === "" || v == null ? null : parseFloat(v));
  const save = () => {
    const targetHrs = num(f.targetHrs), actualHrs = num(f.actualHrs);
    onSave({ date: f.date, protocol: f.protocol, targetHrs, actualHrs, goalMet: actualHrs != null && targetHrs != null && actualHrs + 0.05 >= targetHrs, weight: num(f.weight), waist: num(f.waist), energy: num(f.energy), hunger: num(f.hunger), wellbeing: f.wellbeing.trim(), notes: f.notes.trim() });
  };
  return (
    <div className="fixed inset-0 z-50 flex items-end justify-center bg-slate-900/40 p-0 backdrop-blur-sm sm:items-center sm:p-4" onClick={onClose}>
      <div className="max-h-[94vh] w-full max-w-lg overflow-y-auto rounded-t-3xl bg-white p-5 shadow-xl sm:rounded-3xl" onClick={(e) => e.stopPropagation()}>
        <div className="mb-4 flex items-center justify-between"><h2 className="text-lg font-bold text-slate-900">{entry ? "Запис щоденника" : "Новий запис"}</h2><button onClick={onClose} className="rounded-md p-1 text-slate-400 hover:bg-slate-100"><X className="h-5 w-5" /></button></div>
        <div className="grid grid-cols-2 gap-3">
          <label className="block"><span className="mb-1 block text-xs text-slate-500">Дата</span><input type="date" value={f.date} onChange={(e) => set("date", e.target.value)} className="w-full rounded-lg border border-slate-300 px-2 py-1.5 text-sm" /></label>
          <label className="block"><span className="mb-1 block text-xs text-slate-500">Протокол</span><select value={f.protocol} onChange={(e) => { set("protocol", e.target.value); set("targetHrs", getProtocol(e.target.value).hrs); }} className="w-full rounded-lg border border-slate-300 px-2 py-1.5 text-sm">{PROTOCOLS.map((p) => <option key={p.id} value={p.id}>{p.label}</option>)}</select></label>
          <label className="block"><span className="mb-1 block text-xs text-slate-500">Ціль, год</span><input type="number" step="0.5" value={f.targetHrs} onChange={(e) => set("targetHrs", e.target.value)} className="w-full rounded-lg border border-slate-300 px-2 py-1.5 text-sm" /></label>
          <label className="block"><span className="mb-1 block text-xs text-slate-500">Факт, год</span><input type="number" step="0.5" value={f.actualHrs} onChange={(e) => set("actualHrs", e.target.value)} className="w-full rounded-lg border border-slate-300 px-2 py-1.5 text-sm" /></label>
          <label className="block"><span className="mb-1 block text-xs text-slate-500">Вага, кг</span><input type="number" step="0.1" value={f.weight} onChange={(e) => set("weight", e.target.value)} className="w-full rounded-lg border border-slate-300 px-2 py-1.5 text-sm" /></label>
          <label className="block"><span className="mb-1 block text-xs text-slate-500">Талія, см</span><input type="number" step="0.5" value={f.waist} onChange={(e) => set("waist", e.target.value)} className="w-full rounded-lg border border-slate-300 px-2 py-1.5 text-sm" /></label>
          <label className="block"><span className="mb-1 block text-xs text-slate-500">Енергія 1-5</span><input type="number" min="1" max="5" value={f.energy} onChange={(e) => set("energy", e.target.value)} className="w-full rounded-lg border border-slate-300 px-2 py-1.5 text-sm" /></label>
          <label className="block"><span className="mb-1 block text-xs text-slate-500">Голод 1-5</span><input type="number" min="1" max="5" value={f.hunger} onChange={(e) => set("hunger", e.target.value)} className="w-full rounded-lg border border-slate-300 px-2 py-1.5 text-sm" /></label>
        </div>
        <label className="mt-3 block"><span className="mb-1 block text-xs text-slate-500">Сон / самопочуття</span><input value={f.wellbeing} onChange={(e) => set("wellbeing", e.target.value)} className="w-full rounded-lg border border-slate-300 px-3 py-2 text-sm" /></label>
        <label className="mt-3 block"><span className="mb-1 block text-xs text-slate-500">Нотатки</span><textarea rows={2} value={f.notes} onChange={(e) => set("notes", e.target.value)} className="w-full rounded-lg border border-slate-300 px-3 py-2 text-sm" /></label>
        <div className="mt-5 flex justify-end gap-3"><button onClick={onClose} className="rounded-lg px-4 py-2 text-sm font-medium text-slate-500 hover:text-slate-800">Скасувати</button><button onClick={save} className="inline-flex items-center gap-2 rounded-xl bg-orange-500 px-5 py-2.5 font-semibold text-white hover:bg-orange-600"><Check className="h-4 w-4" /> Зберегти</button></div>
      </div>
    </div>
  );
}

/* ---------- Reference: drinks, tips, safety ---------- */
const DRINKS = [
  { t: "Вода (звичайна, газована)", ok: "yes", c: "Основа. Пий багато. Можна дрібку морської солі для електролітів." },
  { t: "Чорна кава (без цукру)", ok: "yes", c: "Притупляє голод. Без цукру й молока." },
  { t: "Чай (зелений, чорний, трав'яний)", ok: "yes", c: "Без цукру. Гарячий чай добре відволікає від голоду." },
  { t: "Кістковий бульйон", ok: "yes", c: "Особливо при довгих голодуваннях — дає сіль і електроліти." },
  { t: "Морська / гімалайська сіль", ok: "yes", c: "Дрібка у воду проти головного болю й слабкості." },
  { t: "Трохи вершків / олії в каві", ok: "warn", c: "Технічно перериває «чисте» голодування; ок для довгих фаз, не для строгого." },
  { t: "Штучні підсолоджувачі", ok: "warn", c: "Можуть провокувати інсулін і апетит. Фанг радить уникати." },
  { t: "Соки, смузі, молоко", ok: "no", c: "Калорії/цукор — перериває голодування." },
  { t: "Будь-яка їжа, снек", ok: "no", c: "Навіть маленький перекус зупиняє голодування." },
];
const TIPS = [
  ["Пий воду", "Починай ранок з великої склянки води. Гідратація зменшує відчуття голоду."],
  ["Будь зайнятим", "Голод приходить хвилями. Зайнятий день відволікає — час минає непомітно."],
  ["Пий каву", "Чорна кава притупляє апетит і додає бадьорості."],
  ["Пам'ятай: голод минає", "Голод не наростає нескінченно, а йде хвилями. Перечекай хвилю — і вона спаде."],
  ["Не всім розповідай", "Менше зайвих порад і тиску оточення — легше дотримуватись плану."],
  ["Дай собі місяць", "Тілу потрібен час на адаптацію. Перші тижні найважчі, далі легшає."],
  ["Їж низьковуглеводно", "У вікні їжі менше цукру й рафінованих вуглеводів = стабільніший інсулін і менший голод."],
  ["Виходь з голоду м'яко", "Не переїдай одразу. Почни з невеликої порції, щоб не навантажити шлунок."],
  ["Не привід їсти сміття", "Голодування не скасовує якість їжі. Поєднуй його зі здоровим харчуванням."],
  ["Слухай своє тіло", "Якщо стало по-справжньому погано — зупинись і поїж. Здоров'я важливіше за план."],
];
const NO_FAST = [
  "Вагітні та жінки, що годують груддю",
  "Діти й підлітки, що ростуть",
  "Люди з дефіцитом ваги (ІМТ < 18.5)",
  "Розлади харчової поведінки (анорексія, булімія) в анамнезі",
  "Діабет 1 типу або прийом інсуліну / цукрознижувальних — лише під наглядом лікаря",
  "Будь-які хронічні хвороби чи регулярні ліки — спершу консультація лікаря",
];
const SIDE_FX = [
  ["Голод", "Вода, кава або чай. Перечекай хвилю — голод спаде."],
  ["Головний біль", "Часто через нестачу солі. Дрібка солі у воду + більше пити."],
  ["Запаморочення, слабкість", "Сіль і вода. Не вставай різко. Якщо не минає — поїж."],
  ["Судоми м'язів", "Магній (добавка або багата на магній їжа у вікні)."],
  ["Печія", "Не лягай одразу, пий воду, уникай гострого у вікні їжі."],
  ["Дратівливість («hangry»)", "Зазвичай минає з адаптацією за 2–4 тижні."],
  ["Запор", "Більше клітковини й води у вікні їжі."],
];
function FastReference() {
  const okIcon = { yes: "✅", warn: "⚠️", no: "❌" };
  const okColor = { yes: "bg-green-50", warn: "bg-amber-50", no: "bg-red-50" };
  return (
    <div className="space-y-5">
      <h1 className="text-xl font-extrabold text-slate-900">Довідник</h1>

      {/* SAFETY — prominent, first */}
      <div className="rounded-2xl border-2 border-red-200 bg-red-50/60 p-4">
        <div className="mb-2 flex items-center gap-2 text-red-700"><ShieldAlert className="h-5 w-5" /><h2 className="font-bold">Застереження та безпека</h2></div>
        <div className="text-sm font-semibold text-slate-700">Кому НЕ можна голодувати (або лише під наглядом лікаря):</div>
        <ul className="mt-1 space-y-1">{NO_FAST.map((x, i) => <li key={i} className="flex gap-2 text-sm text-slate-600"><span className="text-red-400">•</span>{x}</li>)}</ul>
        <div className="mt-4 overflow-hidden rounded-xl ring-1 ring-red-100">
          <div className="grid grid-cols-2 bg-red-100/70 px-3 py-1.5 text-xs font-bold text-red-700"><span>Симптом</span><span>Що робити</span></div>
          {SIDE_FX.map(([s, w], i) => <div key={i} className={`grid grid-cols-2 gap-2 px-3 py-2 text-sm ${i % 2 ? "bg-white" : "bg-red-50/40"}`}><span className="font-semibold text-slate-700">{s}</span><span className="text-slate-600">{w}</span></div>)}
        </div>
        <p className="mt-3 rounded-lg bg-red-600 px-3 py-2 text-sm font-semibold text-white">❗ Негайно припини голодування і звернись по допомогу при: сильній слабкості, плутанині свідомості, серцебитті, непритомності.</p>
        <p className="mt-2 text-xs text-slate-500">Це освітній матеріал за книгами д-ра Дж. Фанга, не медична порада. Перед голодуванням порадься з лікарем.</p>
      </div>

      {/* Drinks */}
      <div className="rounded-2xl bg-white p-4 shadow-sm ring-1 ring-amber-100">
        <div className="mb-2 flex items-center gap-2 text-slate-700"><GlassWater className="h-5 w-5 text-sky-500" /><h2 className="font-bold">Що можна пити під час голодування</h2></div>
        <div className="space-y-1.5">{DRINKS.map((d, i) => (
          <div key={i} className={`flex items-start gap-2 rounded-xl px-3 py-2 ${okColor[d.ok]}`}>
            <span>{okIcon[d.ok]}</span><div><div className="text-sm font-semibold text-slate-800">{d.t}</div><div className="text-xs text-slate-500">{d.c}</div></div>
          </div>
        ))}</div>
      </div>

      {/* Tips */}
      <div className="rounded-2xl bg-white p-4 shadow-sm ring-1 ring-amber-100">
        <div className="mb-2 flex items-center gap-2 text-slate-700"><Coffee className="h-5 w-5 text-amber-600" /><h2 className="font-bold">10 порад Фанга, як витримати голодування</h2></div>
        <div className="space-y-2">{TIPS.map(([t, e], i) => (
          <div key={i} className="flex gap-3"><span className="grid h-7 w-7 shrink-0 place-items-center rounded-full bg-amber-100 text-sm font-bold text-amber-700">{i + 1}</span><div><div className="text-sm font-semibold text-slate-800">{t}</div><div className="text-xs text-slate-500">{e}</div></div></div>
        ))}</div>
      </div>
    </div>
  );
}

/* ================================================================== */
/* MANAGEMENT — "Manage It!" book notes reader                        */
/* ================================================================== */
function mdInline(str, keyBase = "") {
  const out = [];
  const re = /\*\*(.+?)\*\*|`([^`]+)`|\[(.+?)\]\((.+?)\)/g;
  let last = 0, m, k = 0;
  while ((m = re.exec(str))) {
    if (m.index > last) out.push(str.slice(last, m.index));
    if (m[1] != null) out.push(<strong key={keyBase + k++} className="font-semibold text-slate-900">{m[1]}</strong>);
    else if (m[2] != null) out.push(<code key={keyBase + k++} className="rounded bg-slate-100 px-1 py-0.5 text-[0.85em] text-rose-700">{m[2]}</code>);
    else out.push(<span key={keyBase + k++} className="font-medium text-rose-600">{m[3]}</span>);
    last = re.lastIndex;
  }
  if (last < str.length) out.push(str.slice(last));
  return out;
}
function mdBlocks(text) {
  const lines = (text || "").replace(/\r/g, "").split("\n");
  const els = [];
  let para = [], list = null, quote = null;
  const flushPara = () => { if (para.length) { els.push(<p key={els.length} className="mb-3 leading-relaxed text-slate-600">{mdInline(para.join(" "), els.length + "p")}</p>); para = []; } };
  const flushList = () => {
    if (!list) return;
    if (list.type === "ul") els.push(<ul key={els.length} className="mb-3 space-y-1.5">{list.items.map((it, i) => <li key={i} className="flex gap-2.5 text-slate-600"><span className="mt-[9px] h-1.5 w-1.5 shrink-0 rounded-full bg-rose-300" /><span className="leading-relaxed">{mdInline(it, els.length + "u" + i)}</span></li>)}</ul>);
    else els.push(<ol key={els.length} className="mb-3 space-y-2">{list.items.map((it, i) => <li key={i} className="flex gap-2.5 text-slate-600"><span className="grid h-5 w-5 shrink-0 place-items-center rounded-full bg-rose-100 text-[11px] font-bold text-rose-600">{i + 1}</span><span className="leading-relaxed">{mdInline(it, els.length + "o" + i)}</span></li>)}</ol>);
    list = null;
  };
  const flushQuote = () => { if (quote) { els.push(<blockquote key={els.length} className="mb-3 rounded-r-xl border-l-4 border-amber-300 bg-amber-50/70 px-4 py-2.5 text-slate-700">{mdInline(quote.join(" "), els.length + "q")}</blockquote>); quote = null; } };
  const flushAll = () => { flushPara(); flushList(); flushQuote(); };
  for (const raw of lines) {
    const line = raw.trimEnd();
    if (!line.trim()) { flushAll(); continue; }
    if (/^---+$/.test(line.trim())) { flushAll(); continue; }
    const q = line.match(/^>\s?(.*)/); if (q) { flushPara(); flushList(); (quote = quote || []).push(q[1]); continue; }
    const ul = line.match(/^[-*]\s+(.*)/); if (ul) { flushPara(); flushQuote(); if (!list || list.type !== "ul") { flushList(); list = { type: "ul", items: [] }; } list.items.push(ul[1]); continue; }
    const ol = line.match(/^\d+\.\s+(.*)/); if (ol) { flushPara(); flushQuote(); if (!list || list.type !== "ol") { flushList(); list = { type: "ol", items: [] }; } list.items.push(ol[1]); continue; }
    flushList(); flushQuote(); para.push(line.trim());
  }
  flushAll();
  return els;
}
function parseSections(body) {
  const parts = (body || "").split(/\n### /);
  const lead = parts[0].trim();
  const sections = [];
  for (let i = 1; i < parts.length; i++) {
    const nl = parts[i].indexOf("\n");
    sections.push({ heading: parts[i].slice(0, nl).trim(), content: parts[i].slice(nl + 1) });
  }
  return { lead, sections };
}
function parseManageIt(md) {
  const blocks = md.replace(/\r/g, "").split(/\n## /);
  const intro = blocks[0];
  const chapters = [];
  for (let i = 1; i < blocks.length; i++) {
    const nl = blocks[i].indexOf("\n");
    const heading = blocks[i].slice(0, nl).trim();
    const body = blocks[i].slice(nl + 1);
    if (/^Зміст/.test(heading)) continue;
    const m = heading.match(/^Глава\s+(\d+)\.\s+(.*)/);
    const { lead, sections } = parseSections(body);
    let preview = "";
    const about = sections.find((s) => /Про що глава/i.test(s.heading));
    if (about) preview = about.content.replace(/\r/g, "").split("\n").find((l) => l.trim()) || "";
    chapters.push({ num: m ? m[1] : null, title: m ? m[2] : heading, heading, lead, sections, preview });
  }
  return { intro, chapters };
}

function ManagementSection({ name, onRename }) {
  const [doc, setDoc] = useState(null);
  const [error, setError] = useState("");
  const [view, setView] = useState("toc"); // toc | number index
  const [mtab, setMtab] = useState("book"); // book | career
  const [renaming, setRenaming] = useState(false);
  const [nameDraft, setNameDraft] = useState(name);

  useEffect(() => {
    let alive = true;
    fetch("/manage-it.md").then((r) => { if (!r.ok) throw new Error("not found"); return r.text(); })
      .then((t) => { if (alive) setDoc(parseManageIt(t)); })
      .catch(() => { if (alive) setError("Не вдалося завантажити конспект."); });
    return () => { alive = false; };
  }, []);
  useEffect(() => { document.querySelector("main")?.scrollTo?.(0, 0); window.scrollTo(0, 0); }, [view]);

  if (error) return <div className="flex flex-1 items-center justify-center text-slate-400">{error}</div>;
  if (!doc) return <div className="flex flex-1 items-center justify-center text-rose-400"><div className="flex flex-col items-center gap-3"><BookMarked className="h-8 w-8 animate-pulse" /><span className="text-sm">Завантаження конспекту…</span></div></div>;

  const chapters = doc.chapters;
  const open = (i) => setView(i);
  const chapter = typeof view === "number" ? chapters[view] : null;

  return (
    <div className="min-h-screen flex-1 bg-gradient-to-b from-rose-50/50 via-slate-50 to-white">
      <header className="sticky top-0 z-20 border-b border-slate-200 bg-white/85 backdrop-blur">
        <div className="mx-auto flex h-14 w-full max-w-3xl items-center gap-2 px-4">
          {renaming ? (
            <input autoFocus value={nameDraft} onChange={(e) => setNameDraft(e.target.value)} onBlur={() => { onRename(nameDraft); setRenaming(false); }} onKeyDown={(e) => { if (e.key === "Enter") { onRename(nameDraft); setRenaming(false); } }} className="mr-auto w-40 rounded-lg border border-rose-200 px-2 py-1 text-base font-semibold focus:outline-none" />
          ) : (
            <button onClick={() => { setNameDraft(name); setRenaming(true); }} className="mr-auto text-base font-semibold text-slate-900">{name} <Pencil className="ml-0.5 inline h-3.5 w-3.5 text-slate-300" /></button>
          )}
          {mtab === "book" && view !== "toc" && <button onClick={() => setView("toc")} className="inline-flex items-center gap-1.5 rounded-lg px-3 py-1.5 text-sm font-medium text-slate-500 hover:bg-slate-100 hover:text-slate-800"><ListTree className="h-4 w-4" /> Зміст</button>}
        </div>
      </header>

      <main className="mx-auto w-full max-w-3xl px-4 py-6">
        <div className="mb-4 flex gap-2 rounded-2xl bg-white p-1 shadow-sm ring-1 ring-rose-100">
          <button onClick={() => setMtab("book")} className={`flex flex-1 items-center justify-center gap-1.5 rounded-xl py-2 text-sm font-bold transition ${mtab === "book" ? "bg-rose-500 text-white shadow" : "text-slate-500 hover:text-slate-700"}`}><BookMarked className="h-4 w-4" /> Книга</button>
          <button onClick={() => setMtab("career")} className={`flex flex-1 items-center justify-center gap-1.5 rounded-xl py-2 text-sm font-bold transition ${mtab === "career" ? "bg-rose-500 text-white shadow" : "text-slate-500 hover:text-slate-700"}`}><TrendingUp className="h-4 w-4" /> Кар'єра</button>
        </div>
        {mtab === "career" ? <CareerView /> : view === "toc" ? (
          <>
            {/* book hero from intro */}
            <div className="mb-6 rounded-3xl border border-rose-100 bg-white p-6 shadow-sm">
              <div className="flex items-start gap-4">
                <span className="grid h-14 w-14 shrink-0 place-items-center rounded-2xl bg-rose-600 text-white"><BookMarked className="h-7 w-7" /></span>
                <div className="min-w-0">
                  <h1 className="text-2xl font-extrabold leading-tight text-slate-900">Manage It!</h1>
                  <p className="text-sm font-medium text-slate-500">Johanna Rothman · конспект і лайфхаки по кожній главі</p>
                </div>
              </div>
              <div className="mt-4 text-[15px]">{mdBlocks(introBody(doc.intro))}</div>
            </div>

            <h2 className="mb-3 flex items-center gap-2 text-sm font-semibold uppercase tracking-wide text-slate-400"><ListTree className="h-4 w-4" /> Зміст — {chapters.length} розділів</h2>
            <div className="space-y-2">
              {chapters.map((c, i) => {
                const special = !c.num;
                return (
                  <button key={i} onClick={() => open(i)} className={`flex w-full items-center gap-3 rounded-2xl border p-4 text-left shadow-sm transition hover:shadow-md ${special ? "border-amber-200 bg-amber-50/60 hover:border-amber-300" : "border-slate-100 bg-white hover:border-rose-200"}`}>
                    <span className={`grid h-10 w-10 shrink-0 place-items-center rounded-xl text-sm font-extrabold ${special ? "bg-amber-400 text-white" : "bg-rose-100 text-rose-700"}`}>{special ? "🎯" : c.num}</span>
                    <span className="min-w-0 flex-1">
                      <span className="block font-bold text-slate-800">{c.title}</span>
                      {c.preview && <span className="mt-0.5 line-clamp-2 block text-xs leading-relaxed text-slate-400">{c.preview.replace(/\*\*/g, "")}</span>}
                    </span>
                    <ChevronRight className="h-5 w-5 shrink-0 text-slate-300" />
                  </button>
                );
              })}
            </div>
          </>
        ) : (
          <ChapterReader chapter={chapter} index={view} total={chapters.length} onNav={setView} onToc={() => setView("toc")} />
        )}
      </main>
    </div>
  );
}

// pull the intro's description + idea blockquote (drop the "# title" line and author line)
function introBody(intro) {
  const lines = (intro || "").replace(/\r/g, "").split("\n");
  return lines.filter((l) => !/^#\s/.test(l) && !/^\*\*Джоанна|^\*\*Johanna/i.test(l.trim())).join("\n").trim();
}

function ChapterReader({ chapter, index, total, onNav, onToc }) {
  if (!chapter) return null;
  const secStyle = (h) => {
    if (/🔑/.test(h)) return { wrap: "rounded-2xl border border-amber-200 bg-amber-50/60 p-4", icon: Lightbulb, iconCls: "text-amber-500", title: "text-amber-900" };
    if (/🧭/.test(h)) return { wrap: "rounded-2xl border border-sky-200 bg-sky-50/50 p-4", icon: Compass, iconCls: "text-sky-500", title: "text-sky-900" };
    if (/Про що глава/i.test(h)) return { wrap: "rounded-2xl bg-slate-100/70 p-4", icon: Info, iconCls: "text-slate-400", title: "text-slate-700" };
    return { wrap: "", icon: null, iconCls: "", title: "text-slate-800" };
  };
  const cleanH = (h) => h.replace(/🔑|🧭/g, "").trim();
  return (
    <div>
      <div className="mb-5 flex items-center gap-3">
        {chapter.num && <span className="grid h-11 w-11 shrink-0 place-items-center rounded-2xl bg-rose-600 text-lg font-extrabold text-white">{chapter.num}</span>}
        <h1 className="text-2xl font-extrabold leading-tight text-slate-900">{chapter.title}</h1>
      </div>

      {chapter.lead && <div className="mb-4">{mdBlocks(chapter.lead)}</div>}

      <div className="space-y-4">
        {chapter.sections.map((s, i) => {
          const st = secStyle(s.heading);
          return (
            <section key={i} className={st.wrap}>
              <h3 className={`mb-2 flex items-center gap-2 text-base font-bold ${st.title}`}>{st.icon && <st.icon className={`h-5 w-5 ${st.iconCls}`} />}{cleanH(s.heading)}</h3>
              <div className="text-[15px]">{mdBlocks(s.content)}</div>
            </section>
          );
        })}
      </div>

      {/* prev / next */}
      <div className="mt-8 flex items-center justify-between gap-3 border-t border-slate-100 pt-5">
        <button disabled={index <= 0} onClick={() => onNav(index - 1)} className="inline-flex items-center gap-1.5 rounded-xl border border-slate-300 bg-white px-4 py-2 text-sm font-medium text-slate-700 transition hover:bg-slate-50 disabled:opacity-40"><ChevronLeft className="h-4 w-4" /> Назад</button>
        <button onClick={onToc} className="text-sm font-medium text-slate-400 hover:text-slate-600">Зміст</button>
        <button disabled={index >= total - 1} onClick={() => onNav(index + 1)} className="inline-flex items-center gap-1.5 rounded-xl bg-rose-600 px-4 py-2 text-sm font-semibold text-white transition hover:bg-rose-700 disabled:opacity-40">Далі <ChevronRight className="h-4 w-4" /></button>
      </div>
    </div>
  );
}

/* ================================================================== */
/* TOOLKIT — section UI                                               */
/* ================================================================== */
/* ================================================================== */
/* BUDGET — monthly shopping list + budget tracker (budget:* keys)     */
/* ================================================================== */
const BKEYS = {
  cats: "budget:categories",   // [{id, emoji, name}]
  items: "budget:items",       // [{id, catId, name, qty, unit, price, notes}]
  bought: "budget:bought",     // { [month]: { [itemId]: true } } — куплено цього місяця
  stock: "budget:stock",       // { [month]: { [itemId]: true } } — вже є в запасі, докуповувати не треба
  actuals: "budget:actuals",   // { [month]: { [itemId]: number } } — скільки насправді витрачено
  budgets: "budget:budgets",   // { [month]: number }
  month: "budget:month",       // "2026-07"
  history: "budget:history",   // [{ month, planned, spent }]
  settings: "budget:settings", // { name }
  seeded: "budget:seeded",
};
const MONTHS_UA = ["Січень", "Лютий", "Березень", "Квітень", "Травень", "Червень", "Липень", "Серпень", "Вересень", "Жовтень", "Листопад", "Грудень"];
const budMonthLabel = (m) => { const [y, mo] = (m || "").split("-").map(Number); return MONTHS_UA[(mo || 1) - 1] + " " + (y || ""); };
const budDefaultMonth = () => { const d = new Date(); return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`; };
const budShiftMonth = (m, delta) => { let [y, mo] = m.split("-").map(Number); mo += delta; while (mo > 12) { mo -= 12; y += 1; } while (mo < 1) { mo += 12; y -= 1; } return `${y}-${String(mo).padStart(2, "0")}`; };
const budFmt = (n) => { const r = Math.round((n + Number.EPSILON) * 100) / 100; const s = (Number.isInteger(r) ? r : r.toFixed(2)).toLocaleString ? (Number.isInteger(r) ? r.toLocaleString("uk-UA") : r.toFixed(2)) : String(r); return s + " ₴"; };
const budLineSum = (it) => (Number(it.qty) || 0) * (Number(it.price) || 0);
// stock value: legacy `true` = fully in stock (1); a number 0..1 = fraction already on hand
const budStockFrac = (v) => (v === true ? 1 : (typeof v === "number" ? Math.max(0, Math.min(1, v)) : 0));
const budFracLabel = (f) => (f >= 1 ? "усе" : Math.abs(f - 0.25) < 0.01 ? "¼" : Math.abs(f - 0.5) < 0.01 ? "½" : Math.abs(f - 0.75) < 0.01 ? "¾" : `${Math.round(f * 100)}%`);
// Human-readable "how often" for amortized fractional quantities (qty = share per month).
// 0.12/міс → куплю раз на ~8 місяців.
function budFreqHint(qty) {
  const q = Number(qty) || 0;
  if (q <= 0 || q >= 1) return "";
  const months = Math.round(1 / q);
  if (months <= 1) return "";
  if (months < 12) return `≈ раз на ${months} міс`;
  const years = Math.round(months / 12);
  if (years <= 1) return "≈ раз на рік";
  const w = years >= 2 && years <= 4 ? "роки" : "років";
  return `≈ раз на ${years} ${w}`;
}

async function loadBudgetData() {
  const cats = await store.get(BKEYS.cats, null);
  const items = await store.get(BKEYS.items, null);
  const bought = await store.get(BKEYS.bought, {});
  const stock = await store.get(BKEYS.stock, {});
  const actuals = await store.get(BKEYS.actuals, {});
  const budgets = await store.get(BKEYS.budgets, {});
  const month = await store.get(BKEYS.month, budDefaultMonth());
  const history = await store.get(BKEYS.history, []);
  const settings = await store.get(BKEYS.settings, { name: "Budget" });
  return { cats, items, bought, stock, actuals, budgets, month, history, settings };
}
async function collectBudgetExport() {
  const d = await loadBudgetData();
  return { cats: d.cats || [], items: d.items || [], bought: d.bought, stock: d.stock, actuals: d.actuals, budgets: d.budgets, month: d.month, history: d.history, settings: d.settings };
}
async function clearBudgetData() { for (const k of Object.values(BKEYS)) await store.remove(k); }

// Parse a shopping-list .xlsx following the category-block layout → { cats, items }
function parseBudgetWorkbook(wb) {
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: "" });
  const cats = [], items = [];
  let cur = null;
  const num = (v) => { const n = Number(v); return isNaN(n) ? 0 : n; };
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    const c0 = (r[0] ?? "").toString().trim();
    if (!c0 || c0 === "Товар" || c0.startsWith("Разом")) continue;
    const onlyC0 = (r[1] ?? "") === "" && (r[2] ?? "") === "" && (r[3] ?? "") === "" && (r[4] ?? "") === "";
    if (onlyC0 && /[\u{1F000}-\u{1FAFF}\u{2600}-\u{27BF}]/u.test(c0)) {
      if (/СПИСОК ПОКУПОК|ЗАГАЛОМ|Залишок/i.test(c0)) { cur = null; continue; }
      const sp = c0.indexOf(" ");
      const emoji = sp > 0 ? c0.slice(0, sp) : "🛒";
      const name = sp > 0 ? c0.slice(sp + 1).trim() : c0;
      cur = cats.find((x) => x.name === name);
      if (!cur) { cur = { id: ruid("bc"), emoji, name }; cats.push(cur); }
      continue;
    }
    if (cur && c0) {
      if (items.some((it) => it.catId === cur.id && it.name === c0)) continue;
      items.push({ id: ruid("bi"), catId: cur.id, name: c0, qty: num(r[1]), unit: (r[2] ?? "").toString().trim(), price: num(r[3]), notes: (r[6] ?? "").toString().trim() });
    }
  }
  return { cats, items };
}

function MoneyToggle({ active, onSet }) {
  return (
    <div className="flex gap-1 rounded-full bg-slate-100 p-0.5">
      <button onClick={() => onSet("budget")} className={`rounded-full px-2.5 py-1 text-xs font-bold transition ${active === "budget" ? "bg-emerald-500 text-white" : "text-slate-500"}`}>🛒 Покупки</button>
      <button onClick={() => onSet("finance")} className={`rounded-full px-2.5 py-1 text-xs font-bold transition ${active === "finance" ? "bg-emerald-500 text-white" : "text-slate-500"}`}>💳 Фінанси</button>
    </div>
  );
}
function MoneySection({ budgetName, renameBudget, financeName, renameFinance, onGo }) {
  const [mv, setMv] = useState("budget");
  return mv === "budget"
    ? <BudgetSection name={budgetName} onRename={renameBudget} moneyTab={mv} onMoneyTab={setMv} />
    : <FinanceSection name={financeName} onRename={renameFinance} onGo={onGo} moneyTab={mv} onMoneyTab={setMv} />;
}

function BudgetSection({ name, onRename, moneyTab, onMoneyTab }) {
  const [loading, setLoading] = useState(true);
  const [bview, setBview] = useState("list"); // list | shop | chart
  const [cats, setCats] = useState([]);
  const [items, setItems] = useState([]);
  const [bought, setBought] = useState({});
  const [stock, setStock] = useState({});
  const [actuals, setActuals] = useState({});
  const [budgets, setBudgets] = useState({});
  const [month, setMonth] = useState(budDefaultMonth());
  const [history, setHistory] = useState([]);
  const [collapsed, setCollapsed] = useState({});
  const [itemEditor, setItemEditor] = useState(null); // {item, catId}
  const [catMgr, setCatMgr] = useState(false);
  const [newMonthOpen, setNewMonthOpen] = useState(false);
  const [menuOpen, setMenuOpen] = useState(false);
  const [chartMode, setChartMode] = useState("planned"); // planned | spent
  const [shopFilter, setShopFilter] = useState(false); // hide bought
  const [shopFlat, setShopFlat] = useState(false);
  const [renaming, setRenaming] = useState(false);
  const [nameDraft, setNameDraft] = useState(name);
  const [toast, setToast] = useState(null);
  const fileRef = useRef(null);

  const flash = useCallback((m) => { setToast(m); window.clearTimeout(flash._t); flash._t = window.setTimeout(() => setToast(null), 2200); }, []);

  const reload = useCallback(async () => {
    let d = await loadBudgetData();
    if (d.cats === null && d.items === null) {
      try {
        const res = await fetch("/budget-seed.json");
        const seed = await res.json();
        const c = [], it = [];
        for (const sc of seed.categories) {
          const cid = ruid("bc"); c.push({ id: cid, emoji: sc.emoji, name: sc.name });
          for (const si of sc.items) it.push({ id: ruid("bi"), catId: cid, name: si.name, qty: si.qty, unit: si.unit, price: si.price, notes: si.notes || "" });
        }
        await store.set(BKEYS.cats, c); await store.set(BKEYS.items, it); await store.set(BKEYS.seeded, true);
        d = await loadBudgetData();
      } catch (e) { d.cats = d.cats || []; d.items = d.items || []; }
    }
    setCats(d.cats || []); setItems(d.items || []); setBought(d.bought || {}); setStock(d.stock || {}); setActuals(d.actuals || {}); setBudgets(d.budgets || {});
    setMonth(d.month || budDefaultMonth()); setHistory(d.history || []);
    setLoading(false);
  }, []);
  useEffect(() => {
    reload();
    const onReset = () => reload();
    window.addEventListener("budget-reset", onReset);
    return () => window.removeEventListener("budget-reset", onReset);
  }, [reload]);

  const saveCats = useCallback(async (next) => { setCats(next); await store.set(BKEYS.cats, next); }, []);
  const saveItems = useCallback(async (next) => { setItems(next); await store.set(BKEYS.items, next); }, []);
  const saveBought = useCallback(async (next) => { setBought(next); await store.set(BKEYS.bought, next); }, []);
  const saveBudgets = useCallback(async (next) => { setBudgets(next); await store.set(BKEYS.budgets, next); }, []);
  const saveMonth = useCallback(async (m) => { setMonth(m); await store.set(BKEYS.month, m); }, []);
  const saveHistory = useCallback(async (next) => { setHistory(next); await store.set(BKEYS.history, next); }, []);

  const boughtMap = bought[month] || {};
  const stockMap = stock[month] || {};
  const actualMap = actuals[month] || {};
  // фактично витрачено на позицію: введена сума, інакше планова (для купленого)
  const spentOf = (it) => (boughtMap[it.id] ? (actualMap[it.id] != null ? actualMap[it.id] : budLineSum(it)) : 0);
  const setActual = (id, raw) => setActuals((prev) => {
    const mm = { ...(prev[month] || {}) };
    const n = parseFloat(raw);
    if (raw === "" || isNaN(n)) delete mm[id]; else mm[id] = Math.max(0, n);
    const next = { ...prev, [month]: mm };
    store.set(BKEYS.actuals, next);
    return next;
  });
  const setBoughtItem = (id) => {
    const turningOn = !boughtMap[id];
    setBought((prev) => { const mm = { ...(prev[month] || {}) }; if (mm[id]) delete mm[id]; else mm[id] = true; const next = { ...prev, [month]: mm }; store.set(BKEYS.bought, next); return next; });
    // куплено ↔ в запасі взаємовиключні: якщо позначили купленим, знімаємо «в запасі»
    if (turningOn) setStock((prev) => { const mm = { ...(prev[month] || {}) }; if (!mm[id]) return prev; delete mm[id]; const next = { ...prev, [month]: mm }; store.set(BKEYS.stock, next); return next; });
  };
  const setStockItem = (id) => {
    const on = budStockFrac(stockMap[id]) > 0;
    setStock((prev) => { const mm = { ...(prev[month] || {}) }; if (on) delete mm[id]; else mm[id] = 1; const next = { ...prev, [month]: mm }; store.set(BKEYS.stock, next); return next; });
    if (!on) setBought((prev) => { const mm = { ...(prev[month] || {}) }; if (!mm[id]) return prev; delete mm[id]; const next = { ...prev, [month]: mm }; store.set(BKEYS.bought, next); return next; });
  };
  const setStockFrac = (id, frac) => {
    const f = Math.max(0, Math.min(1, frac));
    setStock((prev) => { const mm = { ...(prev[month] || {}) }; if (f <= 0) delete mm[id]; else mm[id] = f; const next = { ...prev, [month]: mm }; store.set(BKEYS.stock, next); return next; });
    if (f > 0) setBought((prev) => { const mm = { ...(prev[month] || {}) }; if (!mm[id]) return prev; delete mm[id]; const next = { ...prev, [month]: mm }; store.set(BKEYS.bought, next); return next; });
  };

  const itemsOf = (cid) => items.filter((it) => it.catId === cid);
  const planned = useMemo(() => items.reduce((s, it) => s + budLineSum(it), 0), [items]);
  const spent = useMemo(() => items.reduce((s, it) => s + spentOf(it), 0), [items, boughtMap, actualMap]);
  // partial stock counts proportionally: half in stock → half its cost is "already covered"
  const inStockSum = useMemo(() => items.reduce((s, it) => s + budStockFrac(stockMap[it.id]) * budLineSum(it), 0), [items, stockMap]);
  const toBuy = useMemo(() => items.reduce((s, it) => s + (boughtMap[it.id] ? 0 : (1 - budStockFrac(stockMap[it.id])) * budLineSum(it)), 0), [items, boughtMap, stockMap]);
  const budgetAmt = budgets[month] ?? 0;
  const remaining = budgetAmt - spent;

  const setBudgetAmt = (v) => saveBudgets({ ...budgets, [month]: Math.max(0, v) });

  // item CRUD
  const upsertItem = (meta, id, catId) => {
    if (id) saveItems(items.map((it) => (it.id === id ? { ...it, ...meta } : it)));
    else saveItems([...items, { id: ruid("bi"), catId, ...meta }]);
  };
  const deleteItem = (id) => saveItems(items.filter((it) => it.id !== id));

  // category CRUD
  const addCat = (emoji, nm) => saveCats([...cats, { id: ruid("bc"), emoji: emoji || "🛒", name: nm.trim() || "Нова категорія" }]);
  const renameCat = (id, patch) => saveCats(cats.map((c) => (c.id === id ? { ...c, ...patch } : c)));
  const deleteCat = (id) => { saveCats(cats.filter((c) => c.id !== id)); saveItems(items.filter((it) => it.catId !== id)); };
  const moveCat = (id, dir) => { const i = cats.findIndex((c) => c.id === id); const j = i + dir; if (i < 0 || j < 0 || j >= cats.length) return; const next = cats.slice(); [next[i], next[j]] = [next[j], next[i]]; saveCats(next); };

  const startNewMonth = async (keepBudget) => {
    const nm = budShiftMonth(month, 1);
    // snapshot current month into history
    const hist = history.filter((h) => h.month !== month);
    hist.push({ month, planned, spent });
    await saveHistory(hist.sort((a, b) => a.month.localeCompare(b.month)));
    if (keepBudget) await saveBudgets({ ...budgets, [nm]: budgetAmt });
    await saveMonth(nm); // bought[nm] is empty → all unchecked
    setNewMonthOpen(false); flash(`Новий місяць: ${budMonthLabel(nm)} — усі позначки скинуто`);
  };

  const onImportFile = async (file) => {
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: "array" });
      const parsed = parseBudgetWorkbook(wb);
      if (!parsed.cats.length) { flash("Не вдалося розпізнати файл"); return; }
      await saveCats(parsed.cats); await saveItems(parsed.items);
      flash(`Імпортовано: ${parsed.cats.length} категорій, ${parsed.items.length} товарів`);
    } catch (e) { flash("Помилка імпорту"); }
  };

  if (loading) return <div className="flex flex-1 items-center justify-center text-emerald-500"><div className="flex flex-col items-center gap-3"><ShoppingCart className="h-8 w-8 animate-pulse" /><span className="text-sm">Завантаження…</span></div></div>;

  const overBudget = budgetAmt > 0 && spent > budgetAmt;

  return (
    <div className="min-h-screen flex-1 bg-gradient-to-b from-emerald-50/60 to-white">
      <header className="sticky top-0 z-20 border-b border-slate-200 bg-white/85 backdrop-blur">
        <div className="mx-auto flex h-14 w-full max-w-3xl items-center gap-2 px-4">
          {renaming ? (
            <input autoFocus value={nameDraft} onChange={(e) => setNameDraft(e.target.value)} onBlur={() => { onRename(nameDraft); setRenaming(false); }} onKeyDown={(e) => { if (e.key === "Enter") { onRename(nameDraft); setRenaming(false); } }} className="mr-auto w-32 rounded-lg border border-emerald-200 px-2 py-1 text-base font-semibold focus:outline-none" />
          ) : (
            <button onClick={() => { setNameDraft(name); setRenaming(true); }} className="mr-auto text-base font-semibold text-slate-900">{name} <Pencil className="ml-0.5 inline h-3.5 w-3.5 text-slate-300" /></button>
          )}
          {onMoneyTab && <MoneyToggle active={moneyTab} onSet={onMoneyTab} />}
          <div className="relative">
            <button onClick={() => setMenuOpen((v) => !v)} className="grid h-9 w-9 place-items-center rounded-full text-slate-500 hover:bg-slate-100"><Settings className="h-4 w-4" /></button>
            {menuOpen && (<>
              <div className="fixed inset-0 z-10" onClick={() => setMenuOpen(false)} />
              <div className="absolute right-0 top-11 z-20 w-52 overflow-hidden rounded-xl border border-slate-200 bg-white py-1 shadow-lg">
                <button onClick={() => { setCatMgr(true); setMenuOpen(false); }} className="flex w-full items-center gap-2 px-3 py-2 text-left text-sm text-slate-700 hover:bg-slate-50"><ListTree className="h-4 w-4 text-slate-400" /> Категорії</button>
                <button onClick={() => { setNewMonthOpen(true); setMenuOpen(false); }} className="flex w-full items-center gap-2 px-3 py-2 text-left text-sm text-slate-700 hover:bg-slate-50"><CalendarDays className="h-4 w-4 text-slate-400" /> Почати новий місяць</button>
                <button onClick={() => { fileRef.current?.click(); setMenuOpen(false); }} className="flex w-full items-center gap-2 px-3 py-2 text-left text-sm text-slate-700 hover:bg-slate-50"><Upload className="h-4 w-4 text-slate-400" /> Імпорт з Excel</button>
              </div>
            </>)}
          </div>
        </div>
      </header>
      <input ref={fileRef} type="file" accept=".xlsx,.xls" className="hidden" onChange={(e) => { const f = e.target.files?.[0]; if (f) onImportFile(f); e.target.value = ""; }} />

      <main className="mx-auto w-full max-w-3xl px-4 py-5">
        {/* summary */}
        <BudgetSummary month={month} onMonth={(d) => saveMonth(budShiftMonth(month, d))} budgetAmt={budgetAmt} onBudget={setBudgetAmt} planned={planned} spent={spent} remaining={remaining} overBudget={overBudget} toBuy={toBuy} inStockSum={inStockSum} />

        {/* view nav */}
        <div className="mb-4 mt-4 flex gap-2 rounded-2xl bg-white p-1 shadow-sm ring-1 ring-emerald-100">
          {[["list", "Список", ListChecks], ["shop", "Покупки", ShoppingCart], ["chart", "Аналітика", BarChart3]].map(([k, label, Icon]) => (
            <button key={k} onClick={() => setBview(k)} className={`flex flex-1 items-center justify-center gap-1.5 rounded-xl py-2 text-sm font-bold transition ${bview === k ? "bg-emerald-500 text-white shadow" : "text-slate-500 hover:text-slate-700"}`}><Icon className="h-4 w-4" /> {label}</button>
          ))}
        </div>

        {bview === "list" && (
          <div className="space-y-3">
            {cats.length === 0 ? (
              <div className="rounded-2xl bg-white py-12 text-center text-sm text-slate-400 ring-1 ring-emerald-50">Список порожній. Додай категорію в меню ⚙️ або імпортуй Excel.</div>
            ) : cats.map((c) => (
              <CategoryBlock key={c.id} cat={c} items={itemsOf(c.id)} boughtMap={boughtMap} stockMap={stockMap} actualMap={actualMap} collapsed={!!collapsed[c.id]}
                onToggleCollapse={() => setCollapsed((m) => ({ ...m, [c.id]: !m[c.id] }))}
                onToggleBought={setBoughtItem} onToggleStock={setStockItem} onStockFrac={setStockFrac} onActual={setActual} onAddItem={() => setItemEditor({ item: null, catId: c.id })}
                onEditItem={(it) => setItemEditor({ item: it, catId: c.id })} onDeleteItem={deleteItem} />
            ))}
            <button onClick={() => setCatMgr(true)} className="flex w-full items-center justify-center gap-1.5 rounded-2xl border border-dashed border-emerald-300 py-3 text-sm font-semibold text-emerald-600 hover:bg-emerald-50"><Plus className="h-4 w-4" /> Категорія</button>
          </div>
        )}

        {bview === "shop" && (
          <ShoppingView cats={cats} items={items} boughtMap={boughtMap} stockMap={stockMap} actualMap={actualMap} onToggle={setBoughtItem} onToggleStock={setStockItem} onStockFrac={setStockFrac} onActual={setActual}
            filter={shopFilter} setFilter={setShopFilter} flat={shopFlat} setFlat={setShopFlat} />
        )}

        {bview === "chart" && (
          <BudgetChart cats={cats} items={items} boughtMap={boughtMap} actualMap={actualMap} mode={chartMode} setMode={setChartMode} planned={planned} spent={spent} history={history} />
        )}
      </main>

      {itemEditor && <BudgetItemEditor item={itemEditor.item} cats={cats} catId={itemEditor.catId}
        onClose={() => setItemEditor(null)} onDelete={itemEditor.item ? () => { deleteItem(itemEditor.item.id); setItemEditor(null); } : null}
        onSave={(meta, catId) => { upsertItem(meta, itemEditor.item?.id, catId); setItemEditor(null); }} />}
      {catMgr && <BudgetCatManager cats={cats} onClose={() => setCatMgr(false)} onAdd={addCat} onRename={renameCat} onDelete={deleteCat} onMove={moveCat} />}
      {newMonthOpen && <NewMonthModal month={month} next={budShiftMonth(month, 1)} planned={planned} spent={spent} onClose={() => setNewMonthOpen(false)} onConfirm={startNewMonth} />}

      {toast && <div className="fixed bottom-20 left-1/2 z-40 -translate-x-1/2 rounded-full bg-slate-900 px-4 py-2 text-sm text-white shadow-lg lg:bottom-6">{toast}</div>}
    </div>
  );
}

function BudgetSummary({ month, onMonth, budgetAmt, onBudget, planned, spent, remaining, overBudget, toBuy, inStockSum }) {
  return (
    <div className="rounded-3xl bg-white p-4 shadow-sm ring-1 ring-emerald-100">
      <div className="flex items-center justify-between">
        <button onClick={() => onMonth(-1)} className="grid h-8 w-8 place-items-center rounded-full text-slate-400 hover:bg-slate-100"><ChevronLeft className="h-4 w-4" /></button>
        <div className="text-center"><div className="text-lg font-extrabold text-slate-900">{budMonthLabel(month)}</div></div>
        <button onClick={() => onMonth(1)} className="grid h-8 w-8 place-items-center rounded-full text-slate-400 hover:bg-slate-100"><ChevronRight className="h-4 w-4" /></button>
      </div>
      <label className="mt-3 flex items-center justify-between gap-2 rounded-2xl bg-emerald-50/70 px-3 py-2">
        <span className="flex items-center gap-1.5 text-sm font-semibold text-emerald-800"><Wallet className="h-4 w-4" /> Бюджет</span>
        <span className="flex items-center gap-1"><input type="number" min={0} value={budgetAmt || ""} onChange={(e) => onBudget(Math.max(0, +e.target.value || 0))} placeholder="0" className="w-28 rounded-lg border border-emerald-200 bg-white px-2 py-1 text-right text-sm font-bold tabular-nums focus:border-emerald-400 focus:outline-none" /><span className="text-sm font-bold text-emerald-800">₴</span></span>
      </label>
      <div className="mt-3 grid grid-cols-3 gap-2 text-center">
        <div className="rounded-2xl bg-slate-50 p-3"><div className="text-[11px] font-medium text-slate-400">Заплановано</div><div className="mt-0.5 text-base font-extrabold tabular-nums text-slate-700">{budFmt(planned)}</div></div>
        <div className="rounded-2xl bg-sky-50 p-3"><div className="text-[11px] font-medium text-sky-500">Витрачено</div><div className="mt-0.5 text-base font-extrabold tabular-nums text-sky-600">{budFmt(spent)}</div></div>
        <div className={`rounded-2xl p-3 ${overBudget ? "bg-amber-50" : "bg-emerald-50"}`}><div className={`text-[11px] font-medium ${overBudget ? "text-amber-600" : "text-emerald-600"}`}>{overBudget ? "Понад бюджет" : "Залишок"}</div><div className={`mt-0.5 text-base font-extrabold tabular-nums ${overBudget ? "text-amber-600" : "text-emerald-600"}`}>{budFmt(Math.abs(remaining))}</div></div>
      </div>
      {(toBuy > 0 || inStockSum > 0) && (
        <div className="mt-2 flex flex-wrap gap-2 text-xs font-semibold">
          <span className="inline-flex items-center gap-1 rounded-full bg-slate-100 px-2.5 py-1 text-slate-600"><ShoppingCart className="h-3.5 w-3.5" /> Ще купити: <span className="tabular-nums text-slate-800">{budFmt(toBuy)}</span></span>
          {inStockSum > 0 && <span className="inline-flex items-center gap-1 rounded-full bg-amber-50 px-2.5 py-1 text-amber-700"><Package className="h-3.5 w-3.5" /> В запасі: <span className="tabular-nums">{budFmt(inStockSum)}</span></span>}
        </div>
      )}
    </div>
  );
}

function CategoryBlock({ cat, items, boughtMap, stockMap, actualMap, collapsed, onToggleCollapse, onToggleBought, onToggleStock, onStockFrac, onActual, onAddItem, onEditItem, onDeleteItem }) {
  const subtotal = items.reduce((s, it) => s + budLineSum(it), 0);
  const boughtCount = items.filter((it) => boughtMap[it.id]).length;
  const stockCount = items.filter((it) => budStockFrac(stockMap[it.id]) > 0).length;
  return (
    <div className="overflow-hidden rounded-2xl bg-white shadow-sm ring-1 ring-emerald-50">
      <button onClick={onToggleCollapse} className="flex w-full items-center gap-3 px-4 py-3 text-left hover:bg-slate-50">
        <span className="text-xl">{cat.emoji}</span>
        <span className="min-w-0 flex-1"><span className="block truncate font-bold text-slate-800">{cat.name}</span><span className="block text-xs text-slate-400">{boughtCount} з {items.length} куплено{stockCount ? ` · ${stockCount} в запасі` : ""}</span></span>
        <span className="text-sm font-bold tabular-nums text-slate-600">{budFmt(subtotal)}</span>
        {collapsed ? <ChevronRight className="h-4 w-4 text-slate-300" /> : <ChevronDown className="h-4 w-4 text-slate-300" />}
      </button>
      {!collapsed && (
        <div className="divide-y divide-slate-50 border-t border-slate-100">
          {items.map((it) => <BudgetItemRow key={it.id} item={it} bought={!!boughtMap[it.id]} stockFrac={budStockFrac(stockMap[it.id])} actual={actualMap[it.id]} onToggle={() => onToggleBought(it.id)} onToggleStock={() => onToggleStock(it.id)} onStockFrac={(f) => onStockFrac(it.id, f)} onActual={(v) => onActual(it.id, v)} onEdit={() => onEditItem(it)} onDelete={() => onDeleteItem(it.id)} />)}
          <button onClick={onAddItem} className="flex w-full items-center gap-1.5 px-4 py-2.5 text-left text-sm font-medium text-emerald-600 hover:bg-emerald-50"><Plus className="h-4 w-4" /> Товар</button>
        </div>
      )}
    </div>
  );
}

function BudgetItemRow({ item, bought, stockFrac = 0, actual, onToggle, onToggleStock, onStockFrac, onActual, onEdit, onDelete }) {
  const [showNotes, setShowNotes] = useState(false);
  const inStock = stockFrac > 0;
  const nameCls = bought ? "text-slate-400 line-through" : inStock ? "text-amber-600" : "text-slate-800";
  const planned = budLineSum(item);
  const delta = actual != null ? Math.round((actual - planned) * 100) / 100 : 0;
  return (
    <div className="group px-4 py-2.5">
      <div className="flex items-center gap-3">
        <button onClick={onToggle} className={`grid h-6 w-6 shrink-0 place-items-center rounded-md border-2 transition ${bought ? "border-transparent bg-emerald-500 text-white" : "border-slate-300 hover:border-emerald-400"}`}>{bought && <Check className="h-4 w-4" />}</button>
        <button onClick={onEdit} className="min-w-0 flex-1 text-left">
          <div className={`flex items-center gap-1.5 truncate text-sm font-medium ${nameCls}`}>{item.name}{inStock && <span className="shrink-0 rounded-full bg-amber-100 px-1.5 py-0.5 text-[10px] font-bold text-amber-700">в запасі {stockFrac < 1 ? budFracLabel(stockFrac) : ""}</span>}</div>
          <div className="text-xs text-slate-400 tabular-nums">{item.qty} {item.unit} × {budFmt(item.price)} = <span className="font-semibold text-slate-500">{budFmt(planned)}</span></div>
          {budFreqHint(item.qty) && <div className="text-[11px] text-slate-400">{budFreqHint(item.qty)}</div>}
        </button>
        <button onClick={onToggleStock} title="Вже є в запасі — докуповувати не треба" className={`shrink-0 rounded-md p-1 transition ${inStock ? "text-amber-500" : "text-slate-300 hover:text-amber-500"}`}><Package className="h-4 w-4" /></button>
        {item.notes && <button onClick={() => setShowNotes((v) => !v)} title="Нотатки" className={`shrink-0 rounded-md p-1 ${showNotes ? "text-emerald-500" : "text-slate-300 hover:text-slate-500"}`}><Info className="h-4 w-4" /></button>}
        <button onClick={onDelete} className="shrink-0 rounded-md p-1 text-slate-300 hover:text-red-500 sm:opacity-0 sm:group-hover:opacity-100"><Trash2 className="h-4 w-4" /></button>
      </div>
      {inStock && (
        <div className="mt-1.5 flex flex-wrap items-center gap-1.5 pl-9">
          <span className="text-xs font-medium text-slate-400">Скільки вже є:</span>
          {[[0.25, "¼"], [0.5, "½"], [0.75, "¾"], [1, "усе"]].map(([v, l]) => (
            <button key={v} onClick={() => onStockFrac(v)} className={`rounded-full px-2.5 py-0.5 text-xs font-semibold ring-1 transition ${Math.abs(stockFrac - v) < 0.01 ? "bg-amber-400 text-white ring-amber-400" : "bg-white text-slate-500 ring-slate-200 hover:ring-amber-300"}`}>{l}</button>
          ))}
          {stockFrac < 1 && <span className="text-[11px] text-slate-400">→ докупити {budFmt((1 - stockFrac) * planned)}</span>}
        </div>
      )}
      {bought && (
        <div className="mt-1.5 flex flex-wrap items-center gap-x-2 gap-y-1 pl-9">
          <span className="text-xs font-medium text-slate-400">Насправді витрачено:</span>
          <span className="relative inline-flex items-center">
            <input type="number" step="any" min={0} value={actual ?? ""} onChange={(e) => onActual(e.target.value)} placeholder={String(Math.round(planned))} className="w-24 rounded-lg border border-slate-200 bg-white py-1 pl-2 pr-5 text-right text-xs font-semibold tabular-nums text-emerald-700 focus:border-emerald-400 focus:outline-none" />
            <span className="pointer-events-none absolute right-2 text-xs text-slate-400">₴</span>
          </span>
          {actual != null && delta !== 0 && <span className={`text-[11px] font-semibold ${delta > 0 ? "text-red-500" : "text-emerald-600"}`}>{delta > 0 ? "+" : "−"}{budFmt(Math.abs(delta))} {delta > 0 ? "над план" : "менше плану"}</span>}
        </div>
      )}
      {showNotes && item.notes && <div className="mt-1 rounded-lg bg-slate-50 px-3 py-1.5 text-xs text-slate-500">{item.notes}</div>}
    </div>
  );
}

function ShoppingView({ cats, items, boughtMap, stockMap, actualMap, onToggle, onToggleStock, onStockFrac, onActual, filter, setFilter, flat, setFlat }) {
  const total = items.length;
  const done = items.filter((it) => boughtMap[it.id]).length;
  const stockCount = items.filter((it) => budStockFrac(stockMap[it.id]) > 0).length;
  const spentOf = (it) => (boughtMap[it.id] ? (actualMap[it.id] != null ? actualMap[it.id] : budLineSum(it)) : 0);
  const cartTotal = items.reduce((s, it) => s + spentOf(it), 0);
  // "Лишилось купити" ховає куплене й ПОВНІСТЮ наявне; часткове лишається (треба докупити решту)
  const visible = (list) => (filter ? list.filter((it) => !boughtMap[it.id] && budStockFrac(stockMap[it.id]) < 1) : list);
  const Row = (it) => {
    const isBought = !!boughtMap[it.id], sf = budStockFrac(stockMap[it.id]), isStock = sf > 0;
    const planned = budLineSum(it), actual = actualMap[it.id];
    return (
      <div key={it.id} className={`rounded-2xl p-4 shadow-sm ring-1 transition ${isBought ? "bg-emerald-50 ring-emerald-100" : sf >= 1 ? "bg-amber-50/70 ring-amber-100" : "bg-white ring-slate-100 hover:ring-emerald-200"}`}>
        <div className="flex items-center gap-3">
          <button onClick={() => onToggle(it.id)} className={`grid h-9 w-9 shrink-0 place-items-center rounded-full border-2 transition ${isBought ? "border-transparent bg-emerald-500 text-white" : "border-slate-300 hover:border-emerald-400"}`}>{isBought && <Check className="h-5 w-5" />}</button>
          <button onClick={() => onToggle(it.id)} className="min-w-0 flex-1 text-left"><span className={`flex items-center gap-1.5 truncate font-bold ${isBought ? "text-slate-400 line-through" : isStock ? "text-amber-600" : "text-slate-800"}`}>{it.name}{isStock && <span className="shrink-0 rounded-full bg-amber-100 px-1.5 py-0.5 text-[10px] font-bold text-amber-700">в запасі {sf < 1 ? budFracLabel(sf) : ""}</span>}</span><span className="block text-xs text-slate-400 tabular-nums">{it.qty} {it.unit} × {budFmt(it.price)}{sf > 0 && sf < 1 ? ` · докупити ${budFmt((1 - sf) * planned)}` : ""}</span></button>
          <button onClick={() => onToggleStock(it.id)} title="Вже є в запасі" className={`shrink-0 rounded-md p-1.5 transition ${isStock ? "text-amber-500" : "text-slate-300 hover:text-amber-500"}`}><Package className="h-4 w-4" /></button>
          <span className="shrink-0 text-sm font-bold tabular-nums text-slate-600">{budFmt(planned)}</span>
        </div>
        {isStock && sf < 1 && (
          <div className="mt-2 flex flex-wrap items-center gap-1.5 pl-12">
            <span className="text-xs font-medium text-slate-400">Вже є:</span>
            {[[0.25, "¼"], [0.5, "½"], [0.75, "¾"], [1, "усе"]].map(([v, l]) => (
              <button key={v} onClick={() => onStockFrac(it.id, v)} className={`rounded-full px-2.5 py-0.5 text-xs font-semibold ring-1 transition ${Math.abs(sf - v) < 0.01 ? "bg-amber-400 text-white ring-amber-400" : "bg-white text-slate-500 ring-slate-200 hover:ring-amber-300"}`}>{l}</button>
            ))}
          </div>
        )}
        {isBought && (
          <div className="mt-2 flex items-center gap-2 pl-12">
            <span className="text-xs font-medium text-slate-400">Насправді:</span>
            <span className="relative inline-flex items-center">
              <input type="number" step="any" min={0} value={actual ?? ""} onChange={(e) => onActual(it.id, e.target.value)} placeholder={String(Math.round(planned))} className="w-24 rounded-lg border border-slate-200 bg-white py-1 pl-2 pr-5 text-right text-xs font-semibold tabular-nums text-emerald-700 focus:border-emerald-400 focus:outline-none" />
              <span className="pointer-events-none absolute right-2 text-xs text-slate-400">₴</span>
            </span>
          </div>
        )}
      </div>
    );
  };
  return (
    <div>
      <div className="sticky top-14 z-10 -mx-4 mb-3 bg-gradient-to-b from-emerald-50/60 to-transparent px-4 pb-2 pt-1">
        <div className="rounded-2xl bg-white p-3 shadow-sm ring-1 ring-emerald-100">
          <div className="flex items-center justify-between text-sm"><span className="font-semibold text-slate-700">Куплено {done} з {total}{stockCount ? ` · ${stockCount} в запасі` : ""}</span><span className="font-bold tabular-nums text-emerald-600">У кошику: {budFmt(cartTotal)}</span></div>
          <div className="mt-2 h-2.5 overflow-hidden rounded-full bg-emerald-50"><div className="h-full rounded-full bg-emerald-500 transition-all" style={{ width: `${total ? (done / total) * 100 : 0}%` }} /></div>
          <div className="mt-2 flex gap-2">
            <button onClick={() => setFilter(!filter)} className={`rounded-full px-3 py-1 text-xs font-semibold ring-1 transition ${filter ? "bg-emerald-500 text-white ring-emerald-500" : "bg-white text-slate-500 ring-slate-200"}`}>Лишилось купити</button>
            <button onClick={() => setFlat(!flat)} className={`rounded-full px-3 py-1 text-xs font-semibold ring-1 transition ${flat ? "bg-slate-800 text-white ring-slate-800" : "bg-white text-slate-500 ring-slate-200"}`}>{flat ? "Одним списком" : "За категоріями"}</button>
          </div>
        </div>
      </div>
      {flat ? (
        <div className="space-y-2">{visible(items).map(Row)}</div>
      ) : (
        <div className="space-y-4">
          {cats.map((c) => { const list = visible(items.filter((it) => it.catId === c.id)); if (!list.length) return null; return (
            <div key={c.id}><div className="mb-1.5 flex items-center gap-1.5 px-1 text-sm font-bold text-slate-600"><span>{c.emoji}</span> {c.name}</div><div className="space-y-2">{list.map(Row)}</div></div>
          ); })}
        </div>
      )}
    </div>
  );
}

const BUD_COLORS = ["#10b981", "#0ea5e9", "#8b5cf6", "#f59e0b", "#ec4899", "#14b8a6", "#ef4444", "#84cc16", "#6366f1", "#f97316", "#06b6d4"];
function BudgetChart({ cats, items, boughtMap, actualMap, mode, setMode, planned, spent, history }) {
  const spentOf = (it) => (boughtMap[it.id] ? ((actualMap && actualMap[it.id] != null) ? actualMap[it.id] : budLineSum(it)) : 0);
  const data = cats.map((c, i) => {
    const list = items.filter((it) => it.catId === c.id);
    const val = mode === "planned" ? list.reduce((s, it) => s + budLineSum(it), 0) : list.reduce((s, it) => s + spentOf(it), 0);
    return { name: `${c.emoji} ${c.name}`, value: Math.round(val), color: BUD_COLORS[i % BUD_COLORS.length] };
  }).filter((d) => d.value > 0).sort((a, b) => b.value - a.value);
  const totalVal = mode === "planned" ? planned : spent;
  return (
    <div className="space-y-4">
      <div className="flex gap-2">
        {[["planned", "Заплановано"], ["spent", "Витрачено"]].map(([k, label]) => (
          <button key={k} onClick={() => setMode(k)} className={`flex-1 rounded-xl py-2 text-sm font-bold ring-1 transition ${mode === k ? "bg-emerald-500 text-white ring-emerald-500" : "bg-white text-slate-500 ring-slate-200"}`}>{label}</button>
        ))}
      </div>
      <div className="rounded-2xl bg-white p-4 shadow-sm ring-1 ring-emerald-100">
        <div className="mb-2 text-sm font-bold text-slate-700">{mode === "planned" ? "Куди йде бюджет" : "Куди пішли гроші"} · {budFmt(totalVal)}</div>
        {data.length === 0 ? <p className="py-8 text-center text-sm text-slate-400">Немає даних.</p> : (
          <div style={{ height: Math.max(200, data.length * 38) }}>
            <ResponsiveContainer width="100%" height="100%">
              <BarChart data={data} layout="vertical" margin={{ left: 8, right: 16, top: 4, bottom: 4 }}>
                <XAxis type="number" hide />
                <YAxis type="category" dataKey="name" width={130} tick={{ fontSize: 11, fill: "#64748b" }} />
                <Tooltip formatter={(v) => budFmt(v)} cursor={{ fill: "#f1f5f9" }} />
                <Bar dataKey="value" radius={[0, 6, 6, 0]}>{data.map((d, i) => <Cell key={i} fill={d.color} />)}</Bar>
              </BarChart>
            </ResponsiveContainer>
          </div>
        )}
      </div>
      {history.length > 0 && (
        <div className="rounded-2xl bg-white p-4 shadow-sm ring-1 ring-emerald-100">
          <div className="mb-2 text-sm font-bold text-slate-700">Минулі місяці</div>
          <div className="space-y-1.5">
            {history.slice().reverse().map((h) => (
              <div key={h.month} className="flex items-center justify-between text-sm"><span className="text-slate-500">{budMonthLabel(h.month)}</span><span className="tabular-nums text-slate-600">витрачено <b>{budFmt(h.spent)}</b> з {budFmt(h.planned)}</span></div>
            ))}
          </div>
        </div>
      )}
    </div>
  );
}

function BudgetItemEditor({ item, cats, catId, onClose, onSave, onDelete }) {
  const [name, setName] = useState(item?.name || "");
  const [qty, setQty] = useState(item?.qty ?? 1);
  const [unit, setUnit] = useState(item?.unit || "шт");
  const [price, setPrice] = useState(item?.price ?? 0);
  const [notes, setNotes] = useState(item?.notes || "");
  const [cid, setCid] = useState(catId || item?.catId || cats[0]?.id);
  const sum = (Number(qty) || 0) * (Number(price) || 0);
  return (
    <div className="fixed inset-0 z-50 flex items-end justify-center bg-slate-900/40 p-0 backdrop-blur-sm sm:items-center sm:p-4" onClick={onClose}>
      <div className="max-h-[92vh] w-full max-w-md overflow-y-auto rounded-t-3xl bg-white p-5 shadow-xl sm:rounded-3xl" onClick={(e) => e.stopPropagation()}>
        <div className="mb-3 flex items-center justify-between"><h3 className="text-lg font-bold text-slate-900">{item ? "Редагувати товар" : "Новий товар"}</h3><button onClick={onClose} className="rounded-md p-1 text-slate-400 hover:bg-slate-100"><X className="h-5 w-5" /></button></div>
        <input value={name} onChange={(e) => setName(e.target.value)} placeholder="Назва товару" className="mb-3 w-full rounded-xl border border-slate-300 px-3 py-2 text-sm font-semibold focus:border-emerald-400 focus:outline-none" />
        <div className="mb-3 grid grid-cols-3 gap-2">
          <label className="block"><span className="mb-1 block text-xs text-slate-500">К-сть</span><input type="number" min={0} step="any" value={qty} onChange={(e) => setQty(e.target.value)} className="w-full rounded-lg border border-slate-300 px-2 py-1.5 text-sm focus:border-emerald-400 focus:outline-none" /></label>
          <label className="block"><span className="mb-1 block text-xs text-slate-500">Од.</span><input value={unit} onChange={(e) => setUnit(e.target.value)} list="bud-units" className="w-full rounded-lg border border-slate-300 px-2 py-1.5 text-sm focus:border-emerald-400 focus:outline-none" /><datalist id="bud-units">{["шт", "уп", "кг", "л", "міс", "пара", "пачка", "компл", "поїздка"].map((u) => <option key={u} value={u} />)}</datalist></label>
          <label className="block"><span className="mb-1 block text-xs text-slate-500">Ціна ₴</span><input type="number" min={0} step="any" value={price} onChange={(e) => setPrice(e.target.value)} className="w-full rounded-lg border border-slate-300 px-2 py-1.5 text-sm focus:border-emerald-400 focus:outline-none" /></label>
        </div>
        <div className="mb-3 rounded-xl bg-emerald-50 px-3 py-2 text-sm"><div className="flex items-center justify-between"><span className="font-semibold text-emerald-800">Сума</span><span className="font-extrabold tabular-nums text-emerald-700">{budFmt(sum)}</span></div>{budFreqHint(qty) && <div className="mt-0.5 text-[11px] text-emerald-600">дробова к-сть — це «частка на місяць»: {budFreqHint(qty)}</div>}</div>
        <label className="mb-3 block"><span className="mb-1 block text-xs text-slate-500">Категорія</span><select value={cid} onChange={(e) => setCid(e.target.value)} className="w-full rounded-lg border border-slate-300 px-2 py-1.5 text-sm">{cats.map((c) => <option key={c.id} value={c.id}>{c.emoji} {c.name}</option>)}</select></label>
        <textarea value={notes} onChange={(e) => setNotes(e.target.value)} rows={2} placeholder="Нотатки (необов'язково)" className="mb-3 w-full rounded-xl border border-slate-300 px-3 py-2 text-sm focus:border-emerald-400 focus:outline-none" />
        <div className="flex gap-2">
          <button onClick={() => { if (name.trim()) onSave({ name: name.trim(), qty: Number(qty) || 0, unit: unit.trim(), price: Number(price) || 0, notes: notes.trim() }, cid); }} className="flex-1 rounded-2xl bg-emerald-500 py-3 font-bold text-white hover:bg-emerald-600">Зберегти</button>
          {onDelete && <button onClick={onDelete} className="rounded-2xl bg-red-50 px-4 py-3 font-semibold text-red-500 hover:bg-red-100"><Trash2 className="h-5 w-5" /></button>}
        </div>
      </div>
    </div>
  );
}

function BudgetCatManager({ cats, onClose, onAdd, onRename, onDelete, onMove }) {
  const [emoji, setEmoji] = useState("🛒");
  const [nm, setNm] = useState("");
  return (
    <div className="fixed inset-0 z-50 flex items-end justify-center bg-slate-900/40 p-0 backdrop-blur-sm sm:items-center sm:p-4" onClick={onClose}>
      <div className="max-h-[88vh] w-full max-w-md overflow-y-auto rounded-t-3xl bg-white p-5 shadow-xl sm:rounded-3xl" onClick={(e) => e.stopPropagation()}>
        <div className="mb-3 flex items-center justify-between"><h3 className="text-lg font-bold text-slate-900">Категорії</h3><button onClick={onClose} className="rounded-md p-1 text-slate-400 hover:bg-slate-100"><X className="h-5 w-5" /></button></div>
        <div className="space-y-2">
          {cats.map((c, i) => (
            <div key={c.id} className="flex items-center gap-2 rounded-xl bg-slate-50 px-2 py-2">
              <input value={c.emoji} onChange={(e) => onRename(c.id, { emoji: e.target.value.slice(0, 2) })} className="w-9 rounded-lg border border-slate-200 bg-white px-1 py-1 text-center text-lg" />
              <input value={c.name} onChange={(e) => onRename(c.id, { name: e.target.value })} className="min-w-0 flex-1 rounded-lg border border-slate-200 bg-white px-2 py-1.5 text-sm font-medium focus:border-emerald-400 focus:outline-none" />
              <div className="flex shrink-0 flex-col">
                <button onClick={() => onMove(c.id, -1)} disabled={i === 0} className="text-slate-300 hover:text-slate-600 disabled:opacity-30"><ChevronRight className="h-3.5 w-3.5 -rotate-90" /></button>
                <button onClick={() => onMove(c.id, 1)} disabled={i === cats.length - 1} className="text-slate-300 hover:text-slate-600 disabled:opacity-30"><ChevronRight className="h-3.5 w-3.5 rotate-90" /></button>
              </div>
              <button onClick={() => { if (confirm(`Видалити «${c.name}» і всі її товари?`)) onDelete(c.id); }} className="shrink-0 rounded-md p-1 text-slate-300 hover:text-red-500"><Trash2 className="h-4 w-4" /></button>
            </div>
          ))}
        </div>
        <div className="mt-3 flex items-center gap-2 rounded-xl border border-dashed border-slate-300 p-2">
          <input value={emoji} onChange={(e) => setEmoji(e.target.value.slice(0, 2))} className="w-9 rounded-lg border border-slate-200 px-1 py-1 text-center text-lg" />
          <input value={nm} onChange={(e) => setNm(e.target.value)} placeholder="Нова категорія" className="min-w-0 flex-1 rounded-lg border border-slate-200 px-2 py-1.5 text-sm focus:border-emerald-400 focus:outline-none" />
          <button onClick={() => { if (nm.trim()) { onAdd(emoji, nm); setNm(""); setEmoji("🛒"); } }} className="shrink-0 rounded-full bg-slate-800 px-3 py-1.5 text-sm font-semibold text-white hover:bg-slate-900">Додати</button>
        </div>
      </div>
    </div>
  );
}

function NewMonthModal({ month, next, planned, spent, onClose, onConfirm }) {
  const [keepBudget, setKeepBudget] = useState(true);
  return (
    <div className="fixed inset-0 z-50 flex items-end justify-center bg-slate-900/40 p-0 backdrop-blur-sm sm:items-center sm:p-4" onClick={onClose}>
      <div className="w-full max-w-sm rounded-t-3xl bg-white p-6 text-center shadow-xl sm:rounded-3xl" onClick={(e) => e.stopPropagation()}>
        <div className="text-4xl">🗓️</div>
        <h3 className="mt-2 text-lg font-extrabold text-slate-900">Почати {budMonthLabel(next)}?</h3>
        <p className="mt-1 text-sm text-slate-500">Список товарів, ціни й категорії лишаться. Усі позначки «куплено» скинуться — почнеш місяць з чистого аркуша.</p>
        <div className="my-3 rounded-2xl bg-slate-50 p-3 text-sm text-slate-600">{budMonthLabel(month)}: витрачено <b>{budFmt(spent)}</b> з {budFmt(planned)} — збережу в історію.</div>
        <label className="mb-4 flex items-center justify-center gap-2 text-sm text-slate-600"><input type="checkbox" checked={keepBudget} onChange={(e) => setKeepBudget(e.target.checked)} className="h-4 w-4 accent-emerald-500" /> Перенести суму бюджету</label>
        <div className="flex gap-2">
          <button onClick={onClose} className="flex-1 rounded-2xl bg-slate-100 py-3 font-semibold text-slate-500">Скасувати</button>
          <button onClick={() => onConfirm(keepBudget)} className="flex-1 rounded-2xl bg-emerald-500 py-3 font-bold text-white hover:bg-emerald-600">Почати</button>
        </div>
      </div>
    </div>
  );
}

/* ================================================================== */
/* INVENTORY — home checklist (rooms → sections → items)               */
/* ================================================================== */
const IKEYS = {
  rooms: "inventory:rooms",       // [{id, name, sections:[{id,name}]}]
  items: "inventory:items",       // [{id, roomId, secId, name, rec, where, status, notes}]
  settings: "inventory:settings", // {name}
  seeded: "inventory:seeded",
};
const INV_STATUSES = [
  { id: "Є", label: "Є", dot: "#22c55e", bg: "bg-green-50", ring: "ring-green-200", text: "text-green-700", handled: true },
  { id: "Купити", label: "Купити", dot: "#f59e0b", bg: "bg-amber-50", ring: "ring-amber-200", text: "text-amber-700" },
  { id: "Замінити за нагоди", label: "Замінити за нагоди", dot: "#0ea5e9", bg: "bg-sky-50", ring: "ring-sky-200", text: "text-sky-700" },
  { id: "Не потрібно", label: "Не потрібно", dot: "#94a3b8", bg: "bg-slate-100", ring: "ring-slate-200", text: "text-slate-500", handled: true },
  { id: "Не вирішено", label: "Не вирішено", dot: "#cbd5e1", bg: "bg-white", ring: "ring-slate-200", text: "text-slate-400" },
];
const invStatus = (id) => INV_STATUSES.find((s) => s.id === id) || INV_STATUSES[4];
const INV_DECIDED = (st) => st === "Є" || st === "Не потрібно"; // "handled" for readiness

async function loadInventoryData() {
  const rooms = await store.get(IKEYS.rooms, null);
  const items = await store.get(IKEYS.items, null);
  const settings = await store.get(IKEYS.settings, { name: "Inventory" });
  return { rooms, items, settings };
}
async function collectInventoryExport() { const d = await loadInventoryData(); return { rooms: d.rooms || [], items: d.items || [], settings: d.settings }; }
async function clearInventoryData() { for (const k of Object.values(IKEYS)) await store.remove(k); }

function InventorySection({ name, onRename }) {
  const [loading, setLoading] = useState(true);
  const [view, setView] = useState("dash"); // dash | browse | tobuy
  const [rooms, setRooms] = useState([]);
  const [items, setItems] = useState([]);
  const [collapsed, setCollapsed] = useState({});
  const [statusFilter, setStatusFilter] = useState("");
  const [roomFilter, setRoomFilter] = useState("");
  const [query, setQuery] = useState("");
  const [itemEditor, setItemEditor] = useState(null); // {item, roomId, secId}
  const [mgrOpen, setMgrOpen] = useState(false);
  const [menuOpen, setMenuOpen] = useState(false);
  const [renaming, setRenaming] = useState(false);
  const [nameDraft, setNameDraft] = useState(name);
  const [toast, setToast] = useState(null);

  const flash = useCallback((m) => { setToast(m); window.clearTimeout(flash._t); flash._t = window.setTimeout(() => setToast(null), 2200); }, []);

  const reload = useCallback(async () => {
    let d = await loadInventoryData();
    if (d.rooms === null && d.items === null) {
      try {
        const res = await fetch("/inventory-seed.json");
        const seed = await res.json();
        const rms = [], its = [];
        for (const sr of seed.rooms) {
          const rid = ruid("ir"); const secs = [];
          for (const ss of sr.sections) {
            const sid = ruid("is"); secs.push({ id: sid, name: ss.name });
            for (const si of ss.items) its.push({ id: ruid("ii"), roomId: rid, secId: sid, name: si.name, rec: si.rec || "", where: si.where || "", status: si.status || "Не вирішено", notes: si.notes || "" });
          }
          rms.push({ id: rid, name: sr.name, sections: secs });
        }
        await store.set(IKEYS.rooms, rms); await store.set(IKEYS.items, its); await store.set(IKEYS.seeded, true);
        d = await loadInventoryData();
      } catch (e) { d.rooms = d.rooms || []; d.items = d.items || []; }
    }
    setRooms(d.rooms || []); setItems(d.items || []); setLoading(false);
  }, []);
  useEffect(() => {
    reload();
    const onReset = () => reload();
    window.addEventListener("inventory-reset", onReset);
    return () => window.removeEventListener("inventory-reset", onReset);
  }, [reload]);

  const saveRooms = useCallback(async (next) => { setRooms(next); await store.set(IKEYS.rooms, next); }, []);
  const saveItems = useCallback(async (next) => { setItems(next); await store.set(IKEYS.items, next); }, []);

  const setStatus = (id, status) => setItems((prev) => { const next = prev.map((it) => (it.id === id ? { ...it, status } : it)); store.set(IKEYS.items, next); return next; });
  const upsertItem = (meta, id) => { if (id) saveItems(items.map((it) => (it.id === id ? { ...it, ...meta } : it))); else saveItems([...items, { id: ruid("ii"), status: "Не вирішено", ...meta }]); };
  const deleteItem = (id) => saveItems(items.filter((it) => it.id !== id));

  // room / section CRUD
  const addRoom = (nm) => saveRooms([...rooms, { id: ruid("ir"), name: nm.trim() || "Нова кімната", sections: [] }]);
  const renameRoom = (id, nm) => saveRooms(rooms.map((r) => (r.id === id ? { ...r, name: nm } : r)));
  const deleteRoom = (id) => { saveRooms(rooms.filter((r) => r.id !== id)); saveItems(items.filter((it) => it.roomId !== id)); };
  const moveRoom = (id, dir) => { const i = rooms.findIndex((r) => r.id === id); const j = i + dir; if (i < 0 || j < 0 || j >= rooms.length) return; const n = rooms.slice(); [n[i], n[j]] = [n[j], n[i]]; saveRooms(n); };
  const addSection = (roomId, nm) => saveRooms(rooms.map((r) => (r.id === roomId ? { ...r, sections: [...r.sections, { id: ruid("is"), name: nm.trim() || "Новий розділ" }] } : r)));
  const renameSection = (roomId, secId, nm) => saveRooms(rooms.map((r) => (r.id === roomId ? { ...r, sections: r.sections.map((s) => (s.id === secId ? { ...s, name: nm } : s)) } : r)));
  const deleteSection = (roomId, secId) => { saveRooms(rooms.map((r) => (r.id === roomId ? { ...r, sections: r.sections.filter((s) => s.id !== secId) } : r))); saveItems(items.filter((it) => it.secId !== secId)); };
  const moveSection = (roomId, secId, dir) => saveRooms(rooms.map((r) => { if (r.id !== roomId) return r; const i = r.sections.findIndex((s) => s.id === secId); const j = i + dir; if (i < 0 || j < 0 || j >= r.sections.length) return r; const n = r.sections.slice(); [n[i], n[j]] = [n[j], n[i]]; return { ...r, sections: n }; }));

  const itemsOfRoom = (rid) => items.filter((it) => it.roomId === rid);
  const roomStats = (rid) => {
    const list = itemsOfRoom(rid);
    const c = { total: list.length, є: 0, buy: 0, no: 0, undec: 0, repl: 0 };
    for (const it of list) { if (it.status === "Є") c.є++; else if (it.status === "Купити") c.buy++; else if (it.status === "Не потрібно") c.no++; else if (it.status === "Замінити за нагоди") c.repl++; else c.undec++; }
    c.readiness = c.total ? (c.є + c.no) / c.total : 0;
    return c;
  };
  const overall = useMemo(() => {
    const total = items.length; const handled = items.filter((it) => INV_DECIDED(it.status)).length;
    return { total, handled, pct: total ? handled / total : 0 };
  }, [items]);

  const toBuy = items.filter((it) => it.status === "Купити");

  const copyToBuy = () => {
    const byRoom = rooms.map((r) => { const list = toBuy.filter((it) => it.roomId === r.id); return list.length ? `${r.name}:\n` + list.map((it) => `  • ${it.name}${it.rec ? ` (${it.rec})` : ""}`).join("\n") : ""; }).filter(Boolean).join("\n\n");
    const text = "Купити для дому:\n\n" + byRoom;
    try { navigator.clipboard.writeText(text); flash("Скопійовано у буфер 📋"); } catch { flash("Не вдалося скопіювати"); }
  };
  const sendToBudget = async () => {
    if (!toBuy.length) return;
    const cats = await store.get(BKEYS.cats, []);
    const its = await store.get(BKEYS.items, []);
    let cat = cats.find((c) => c.name === "Дім (інвентар)");
    let nextCats = cats;
    if (!cat) { cat = { id: ruid("bc"), emoji: "🏠", name: "Дім (інвентар)" }; nextCats = [...cats, cat]; }
    const existing = new Set(its.filter((x) => x.catId === cat.id).map((x) => x.name));
    const add = toBuy.filter((it) => !existing.has(it.name)).map((it) => ({ id: ruid("bi"), catId: cat.id, name: it.name, qty: 1, unit: "шт", price: 0, notes: it.rec ? `реком.: ${it.rec}` : "" }));
    await store.set(BKEYS.cats, nextCats); await store.set(BKEYS.items, [...its, ...add]);
    flash(`Додано в Budget: ${add.length} товар(и) 🛒`);
  };

  if (loading) return <div className="flex flex-1 items-center justify-center text-rose-400"><div className="flex flex-col items-center gap-3"><Home className="h-8 w-8 animate-pulse" /><span className="text-sm">Завантаження…</span></div></div>;

  return (
    <div className="min-h-screen flex-1 bg-gradient-to-b from-rose-50/50 to-white">
      <header className="sticky top-0 z-20 border-b border-slate-200 bg-white/85 backdrop-blur">
        <div className="mx-auto flex h-14 w-full max-w-3xl items-center gap-2 px-4">
          {renaming ? (
            <input autoFocus value={nameDraft} onChange={(e) => setNameDraft(e.target.value)} onBlur={() => { onRename(nameDraft); setRenaming(false); }} onKeyDown={(e) => { if (e.key === "Enter") { onRename(nameDraft); setRenaming(false); } }} className="mr-auto w-32 rounded-lg border border-rose-200 px-2 py-1 text-base font-semibold focus:outline-none" />
          ) : (
            <button onClick={() => { setNameDraft(name); setRenaming(true); }} className="mr-auto text-base font-semibold text-slate-900">{name} <Pencil className="ml-0.5 inline h-3.5 w-3.5 text-slate-300" /></button>
          )}
          <div className="relative">
            <button onClick={() => setMenuOpen((v) => !v)} className="grid h-9 w-9 place-items-center rounded-full text-slate-500 hover:bg-slate-100"><Settings className="h-4 w-4" /></button>
            {menuOpen && (<>
              <div className="fixed inset-0 z-10" onClick={() => setMenuOpen(false)} />
              <div className="absolute right-0 top-11 z-20 w-52 overflow-hidden rounded-xl border border-slate-200 bg-white py-1 shadow-lg">
                <button onClick={() => { setMgrOpen(true); setMenuOpen(false); }} className="flex w-full items-center gap-2 px-3 py-2 text-left text-sm text-slate-700 hover:bg-slate-50"><ListTree className="h-4 w-4 text-slate-400" /> Кімнати й розділи</button>
              </div>
            </>)}
          </div>
        </div>
      </header>

      <main className="mx-auto w-full max-w-3xl px-4 py-5">
        <div className="mb-4 flex gap-2 rounded-2xl bg-white p-1 shadow-sm ring-1 ring-rose-100">
          {[["dash", "Готовність", BarChart3], ["browse", "Перелік", ListChecks], ["tobuy", "Купити", ShoppingCart]].map(([k, label, Icon]) => (
            <button key={k} onClick={() => setView(k)} className={`relative flex flex-1 items-center justify-center gap-1.5 rounded-xl py-2 text-sm font-bold transition ${view === k ? "bg-rose-500 text-white shadow" : "text-slate-500 hover:text-slate-700"}`}><Icon className="h-4 w-4" /> {label}{k === "tobuy" && toBuy.length > 0 && <span className={`ml-0.5 rounded-full px-1.5 text-[10px] font-bold ${view === k ? "bg-white/25" : "bg-amber-100 text-amber-700"}`}>{toBuy.length}</span>}</button>
          ))}
        </div>

        {view === "dash" && <InvDashboard rooms={rooms} overall={overall} roomStats={roomStats} onRoom={(rid) => { setRoomFilter(rid); setView("browse"); }} />}
        {view === "browse" && <InvBrowse rooms={rooms} items={items} collapsed={collapsed} setCollapsed={setCollapsed}
          statusFilter={statusFilter} setStatusFilter={setStatusFilter} roomFilter={roomFilter} setRoomFilter={setRoomFilter} query={query} setQuery={setQuery}
          onStatus={setStatus} onEdit={(it) => setItemEditor({ item: it })} onDelete={deleteItem} onAddItem={(roomId, secId) => setItemEditor({ item: null, roomId, secId })} roomStats={roomStats} />}
        {view === "tobuy" && <InvToBuy rooms={rooms} toBuy={toBuy} onAcquire={(id) => setStatus(id, "Є")} onCopy={copyToBuy} onSendBudget={sendToBudget} />}
      </main>

      {itemEditor && <InvItemEditor item={itemEditor.item} rooms={rooms} roomId={itemEditor.roomId} secId={itemEditor.secId}
        onClose={() => setItemEditor(null)} onDelete={itemEditor.item ? () => { deleteItem(itemEditor.item.id); setItemEditor(null); } : null}
        onSave={(meta) => { upsertItem(meta, itemEditor.item?.id); setItemEditor(null); }} />}
      {mgrOpen && <InvManager rooms={rooms} onClose={() => setMgrOpen(false)} onAddRoom={addRoom} onRenameRoom={renameRoom} onDeleteRoom={deleteRoom} onMoveRoom={moveRoom} onAddSection={addSection} onRenameSection={renameSection} onDeleteSection={deleteSection} onMoveSection={moveSection} />}

      {toast && <div className="fixed bottom-20 left-1/2 z-40 -translate-x-1/2 rounded-full bg-slate-900 px-4 py-2 text-sm text-white shadow-lg lg:bottom-6">{toast}</div>}
    </div>
  );
}

function InvDashboard({ rooms, overall, roomStats, onRoom }) {
  return (
    <div className="space-y-4">
      <div className="rounded-3xl bg-white p-5 shadow-sm ring-1 ring-rose-100">
        <div className="flex items-center gap-4">
          <ProgressRing pct={overall.pct} size={72} stroke={8}><span className="text-sm font-extrabold text-rose-600">{Math.round(overall.pct * 100)}%</span></ProgressRing>
          <div>
            <div className="text-lg font-extrabold text-slate-900">Готовність дому</div>
            <div className="text-sm text-slate-500">{overall.handled} з {overall.total} позицій вирішено</div>
          </div>
        </div>
        <p className="mt-3 rounded-xl bg-rose-50/60 px-3 py-2 text-xs leading-relaxed text-rose-800">Готовність = (Є + Не потрібно) ÷ усі позиції. Тобто скільки речей уже вирішено — куплено/є або свідомо не треба.</p>
      </div>
      <div className="space-y-2.5">
        {rooms.map((r) => { const c = roomStats(r.id); return (
          <button key={r.id} onClick={() => onRoom(r.id)} className="block w-full rounded-2xl bg-white p-4 text-left shadow-sm ring-1 ring-rose-50 hover:ring-rose-200">
            <div className="flex items-center justify-between">
              <span className="font-bold text-slate-800">{r.name}</span>
              <span className="text-sm font-bold tabular-nums text-rose-600">{Math.round(c.readiness * 100)}%</span>
            </div>
            <div className="mt-2 h-2 overflow-hidden rounded-full bg-rose-50"><div className="h-full rounded-full bg-rose-500 transition-all" style={{ width: `${c.readiness * 100}%` }} /></div>
            <div className="mt-2 flex flex-wrap gap-x-3 gap-y-0.5 text-[11px] text-slate-400">
              <span>усього {c.total}</span>
              <span className="text-green-600">Є {c.є}</span>
              <span className="text-amber-600">Купити {c.buy}</span>
              <span className="text-slate-500">Не потрібно {c.no}</span>
              <span>Не вирішено {c.undec}</span>
            </div>
          </button>
        ); })}
      </div>
    </div>
  );
}

function InvStatusSelect({ value, onChange }) {
  const s = invStatus(value);
  return (
    <div className="relative">
      <select value={value} onChange={(e) => onChange(e.target.value)} className={`cursor-pointer appearance-none rounded-full py-1 pl-2.5 pr-6 text-xs font-bold ring-1 ${s.bg} ${s.text} ${s.ring} focus:outline-none`}>
        {INV_STATUSES.map((o) => <option key={o.id} value={o.id}>{o.label}</option>)}
      </select>
      <ChevronDown className="pointer-events-none absolute right-1.5 top-1/2 h-3.5 w-3.5 -translate-y-1/2 opacity-50" />
    </div>
  );
}

function InvItemRow({ item, onStatus, onEdit, onDelete }) {
  const [showNotes, setShowNotes] = useState(false);
  return (
    <div className="group px-4 py-2.5">
      <div className="flex items-start gap-2">
        <button onClick={onEdit} className="min-w-0 flex-1 text-left">
          <div className="text-sm font-medium text-slate-800">{item.name}</div>
          <div className="mt-0.5 flex flex-wrap gap-x-2 text-[11px] text-slate-400">
            {item.rec && <span>реком.: <span className="text-slate-500">{item.rec}</span></span>}
            {item.where && <span>· вдома: <span className="text-slate-500">{item.where}</span></span>}
          </div>
        </button>
        {item.notes && <button onClick={() => setShowNotes((v) => !v)} title="Нотатки" className={`shrink-0 rounded-md p-1 ${showNotes ? "text-rose-500" : "text-slate-300 hover:text-slate-500"}`}><Info className="h-4 w-4" /></button>}
        <InvStatusSelect value={item.status} onChange={(st) => onStatus(item.id, st)} />
        <button onClick={onDelete} className="shrink-0 rounded-md p-1 text-slate-300 hover:text-red-500 sm:opacity-0 sm:group-hover:opacity-100"><Trash2 className="h-4 w-4" /></button>
      </div>
      {showNotes && item.notes && <div className="mt-1 rounded-lg bg-slate-50 px-3 py-1.5 text-xs text-slate-500">{item.notes}</div>}
    </div>
  );
}

function InvBrowse({ rooms, items, collapsed, setCollapsed, statusFilter, setStatusFilter, roomFilter, setRoomFilter, query, setQuery, onStatus, onEdit, onDelete, onAddItem, roomStats }) {
  const q = query.trim().toLowerCase();
  const match = (it) => (!statusFilter || it.status === statusFilter) && (!q || it.name.toLowerCase().includes(q));
  const visRooms = rooms.filter((r) => !roomFilter || r.id === roomFilter);
  return (
    <div className="space-y-3">
      <div className="relative">
        <Search className="pointer-events-none absolute left-3 top-1/2 h-4 w-4 -translate-y-1/2 text-slate-400" />
        <input value={query} onChange={(e) => setQuery(e.target.value)} placeholder="Пошук предмета по всьому дому…" className="w-full rounded-2xl border border-rose-100 bg-white py-2.5 pl-10 pr-9 text-sm shadow-sm focus:border-rose-300 focus:outline-none" />
        {query && <button onClick={() => setQuery("")} className="absolute right-3 top-1/2 -translate-y-1/2 text-slate-400"><X className="h-4 w-4" /></button>}
      </div>
      <div className="-mx-1 flex gap-2 overflow-x-auto px-1 pb-1">
        <select value={roomFilter} onChange={(e) => setRoomFilter(e.target.value)} className="shrink-0 rounded-full border border-slate-200 bg-white px-3 py-1.5 text-xs font-semibold text-slate-600 focus:outline-none"><option value="">Усі кімнати</option>{rooms.map((r) => <option key={r.id} value={r.id}>{r.name}</option>)}</select>
        <button onClick={() => setStatusFilter("")} className={`shrink-0 rounded-full px-3 py-1.5 text-xs font-semibold ring-1 ${!statusFilter ? "bg-slate-800 text-white ring-slate-800" : "bg-white text-slate-500 ring-slate-200"}`}>Усі статуси</button>
        {INV_STATUSES.map((s) => <button key={s.id} onClick={() => setStatusFilter(statusFilter === s.id ? "" : s.id)} className={`shrink-0 rounded-full px-3 py-1.5 text-xs font-semibold ring-1 transition ${statusFilter === s.id ? `${s.bg} ${s.text} ${s.ring}` : "bg-white text-slate-500 ring-slate-200"}`}>{s.label}</button>)}
      </div>
      {visRooms.map((r) => {
        const roomItems = items.filter((it) => it.roomId === r.id && match(it));
        if (!roomItems.length && (statusFilter || q)) return null;
        const c = roomStats(r.id); const decided = c.total - c.undec;
        return (
          <div key={r.id} className="overflow-hidden rounded-2xl bg-white shadow-sm ring-1 ring-rose-50">
            <button onClick={() => setCollapsed((m) => ({ ...m, [r.id]: !m[r.id] }))} className="flex w-full items-center gap-3 px-4 py-3 text-left hover:bg-slate-50">
              <span className="min-w-0 flex-1"><span className="block truncate font-extrabold text-slate-800">{r.name}</span><span className="block text-xs text-slate-400">{decided} з {c.total} вирішено</span></span>
              <span className="text-sm font-bold tabular-nums text-rose-600">{Math.round(c.readiness * 100)}%</span>
              {collapsed[r.id] ? <ChevronRight className="h-4 w-4 text-slate-300" /> : <ChevronDown className="h-4 w-4 text-slate-300" />}
            </button>
            {!collapsed[r.id] && (
              <div className="border-t border-slate-100">
                {r.sections.map((sec) => {
                  const secItems = items.filter((it) => it.secId === sec.id && match(it));
                  if (!secItems.length) return null;
                  const sTotal = items.filter((it) => it.secId === sec.id).length;
                  const sDecided = items.filter((it) => it.secId === sec.id && it.status !== "Не вирішено").length;
                  return (
                    <div key={sec.id}>
                      <div className="flex items-center justify-between bg-slate-50/70 px-4 py-1.5"><span className="text-xs font-bold text-slate-500">{sec.name}</span><span className="text-[11px] text-slate-400">{sDecided}/{sTotal}</span></div>
                      <div className="divide-y divide-slate-50">
                        {secItems.map((it) => <InvItemRow key={it.id} item={it} onStatus={onStatus} onEdit={() => onEdit(it)} onDelete={() => onDelete(it.id)} />)}
                        {!statusFilter && !q && <button onClick={() => onAddItem(r.id, sec.id)} className="flex w-full items-center gap-1.5 px-4 py-2 text-left text-xs font-medium text-rose-600 hover:bg-rose-50"><Plus className="h-3.5 w-3.5" /> Предмет</button>}
                      </div>
                    </div>
                  );
                })}
              </div>
            )}
          </div>
        );
      })}
    </div>
  );
}

function InvToBuy({ rooms, toBuy, onAcquire, onCopy, onSendBudget }) {
  return (
    <div className="space-y-4">
      <div className="flex items-center justify-between rounded-2xl bg-white p-4 shadow-sm ring-1 ring-rose-100">
        <div><div className="text-lg font-extrabold text-slate-900">Купити для дому</div><div className="text-sm text-slate-400">{toBuy.length} позицій · познач ✓ коли придбала</div></div>
        <ShoppingCart className="h-8 w-8 text-amber-400" />
      </div>
      {toBuy.length === 0 ? (
        <div className="rounded-2xl bg-white py-12 text-center text-sm text-slate-400 ring-1 ring-rose-50">Нічого купувати 🎉 Постав комусь статус «Купити» в переліку.</div>
      ) : (<>
        <div className="flex gap-2">
          <button onClick={onCopy} className="flex flex-1 items-center justify-center gap-1.5 rounded-2xl bg-white py-2.5 text-sm font-semibold text-slate-600 shadow-sm ring-1 ring-slate-200 hover:ring-rose-200"><ClipboardPaste className="h-4 w-4" /> Копіювати списком</button>
          <button onClick={onSendBudget} className="flex flex-1 items-center justify-center gap-1.5 rounded-2xl bg-white py-2.5 text-sm font-semibold text-slate-600 shadow-sm ring-1 ring-slate-200 hover:ring-emerald-200"><ShoppingBasket className="h-4 w-4" /> У Budget</button>
        </div>
        {rooms.map((r) => { const list = toBuy.filter((it) => it.roomId === r.id); if (!list.length) return null; return (
          <div key={r.id}>
            <div className="mb-1.5 px-1 text-sm font-bold text-slate-600">{r.name}</div>
            <div className="space-y-2">
              {list.map((it) => (
                <button key={it.id} onClick={() => onAcquire(it.id)} className="flex w-full items-center gap-3 rounded-2xl bg-white p-3.5 text-left shadow-sm ring-1 ring-slate-100 transition hover:ring-green-200">
                  <span className="grid h-8 w-8 shrink-0 place-items-center rounded-full border-2 border-slate-300 text-transparent hover:border-green-400"><Check className="h-4 w-4" /></span>
                  <span className="min-w-0 flex-1"><span className="block truncate font-semibold text-slate-800">{it.name}</span>{it.rec && <span className="block text-xs text-slate-400">реком.: {it.rec}</span>}</span>
                </button>
              ))}
            </div>
          </div>
        ); })}
      </>)}
    </div>
  );
}

function InvItemEditor({ item, rooms, roomId, secId, onClose, onSave, onDelete }) {
  const [name, setName] = useState(item?.name || "");
  const [rec, setRec] = useState(item?.rec || "");
  const [where, setWhere] = useState(item?.where || "");
  const [status, setStatus] = useState(item?.status || "Не вирішено");
  const [notes, setNotes] = useState(item?.notes || "");
  const [rid, setRid] = useState(item?.roomId || roomId || rooms[0]?.id);
  const room = rooms.find((r) => r.id === rid);
  const [sid, setSid] = useState(item?.secId || secId || room?.sections[0]?.id);
  return (
    <div className="fixed inset-0 z-50 flex items-end justify-center bg-slate-900/40 p-0 backdrop-blur-sm sm:items-center sm:p-4" onClick={onClose}>
      <div className="max-h-[92vh] w-full max-w-md overflow-y-auto rounded-t-3xl bg-white p-5 shadow-xl sm:rounded-3xl" onClick={(e) => e.stopPropagation()}>
        <div className="mb-3 flex items-center justify-between"><h3 className="text-lg font-bold text-slate-900">{item ? "Редагувати" : "Новий предмет"}</h3><button onClick={onClose} className="rounded-md p-1 text-slate-400 hover:bg-slate-100"><X className="h-5 w-5" /></button></div>
        <input value={name} onChange={(e) => setName(e.target.value)} placeholder="Назва предмета" className="mb-3 w-full rounded-xl border border-slate-300 px-3 py-2 text-sm font-semibold focus:border-rose-400 focus:outline-none" />
        <div className="mb-3 grid grid-cols-2 gap-2">
          <label className="block"><span className="mb-1 block text-xs text-slate-500">Рекомендовано</span><input value={rec} onChange={(e) => setRec(e.target.value)} placeholder="напр. 2–4" className="w-full rounded-lg border border-slate-300 px-2 py-1.5 text-sm focus:border-rose-400 focus:outline-none" /></label>
          <label className="block"><span className="mb-1 block text-xs text-slate-500">Статус</span><select value={status} onChange={(e) => setStatus(e.target.value)} className="w-full rounded-lg border border-slate-300 px-2 py-1.5 text-sm">{INV_STATUSES.map((s) => <option key={s.id} value={s.id}>{s.label}</option>)}</select></label>
        </div>
        <label className="mb-3 block"><span className="mb-1 block text-xs text-slate-500">Де є вдома</span><input value={where} onChange={(e) => setWhere(e.target.value)} placeholder="напр. в шафі на верхній полиці" className="w-full rounded-lg border border-slate-300 px-2 py-1.5 text-sm focus:border-rose-400 focus:outline-none" /></label>
        <div className="mb-3 grid grid-cols-2 gap-2">
          <label className="block"><span className="mb-1 block text-xs text-slate-500">Кімната</span><select value={rid} onChange={(e) => { setRid(e.target.value); const rr = rooms.find((x) => x.id === e.target.value); setSid(rr?.sections[0]?.id); }} className="w-full rounded-lg border border-slate-300 px-2 py-1.5 text-sm">{rooms.map((r) => <option key={r.id} value={r.id}>{r.name}</option>)}</select></label>
          <label className="block"><span className="mb-1 block text-xs text-slate-500">Розділ</span><select value={sid} onChange={(e) => setSid(e.target.value)} className="w-full rounded-lg border border-slate-300 px-2 py-1.5 text-sm">{(room?.sections || []).map((s) => <option key={s.id} value={s.id}>{s.name}</option>)}</select></label>
        </div>
        <textarea value={notes} onChange={(e) => setNotes(e.target.value)} rows={2} placeholder="Нотатки (необов'язково)" className="mb-3 w-full rounded-xl border border-slate-300 px-3 py-2 text-sm focus:border-rose-400 focus:outline-none" />
        <div className="flex gap-2">
          <button onClick={() => { if (name.trim() && rid && sid) onSave({ name: name.trim(), rec: rec.trim(), where: where.trim(), status, notes: notes.trim(), roomId: rid, secId: sid }); }} className="flex-1 rounded-2xl bg-rose-500 py-3 font-bold text-white hover:bg-rose-600">Зберегти</button>
          {onDelete && <button onClick={onDelete} className="rounded-2xl bg-red-50 px-4 py-3 font-semibold text-red-500 hover:bg-red-100"><Trash2 className="h-5 w-5" /></button>}
        </div>
      </div>
    </div>
  );
}

function InvManager({ rooms, onClose, onAddRoom, onRenameRoom, onDeleteRoom, onMoveRoom, onAddSection, onRenameSection, onDeleteSection, onMoveSection }) {
  const [newRoom, setNewRoom] = useState("");
  const [openRoom, setOpenRoom] = useState(null);
  const [newSec, setNewSec] = useState("");
  return (
    <div className="fixed inset-0 z-50 flex items-end justify-center bg-slate-900/40 p-0 backdrop-blur-sm sm:items-center sm:p-4" onClick={onClose}>
      <div className="max-h-[88vh] w-full max-w-md overflow-y-auto rounded-t-3xl bg-white p-5 shadow-xl sm:rounded-3xl" onClick={(e) => e.stopPropagation()}>
        <div className="mb-3 flex items-center justify-between"><h3 className="text-lg font-bold text-slate-900">Кімнати й розділи</h3><button onClick={onClose} className="rounded-md p-1 text-slate-400 hover:bg-slate-100"><X className="h-5 w-5" /></button></div>
        <div className="space-y-2">
          {rooms.map((r, i) => (
            <div key={r.id} className="rounded-xl bg-slate-50 p-2">
              <div className="flex items-center gap-2">
                <input value={r.name} onChange={(e) => onRenameRoom(r.id, e.target.value)} className="min-w-0 flex-1 rounded-lg border border-slate-200 bg-white px-2 py-1.5 text-sm font-semibold focus:border-rose-400 focus:outline-none" />
                <div className="flex shrink-0 flex-col"><button onClick={() => onMoveRoom(r.id, -1)} disabled={i === 0} className="text-slate-300 hover:text-slate-600 disabled:opacity-30"><ChevronRight className="h-3.5 w-3.5 -rotate-90" /></button><button onClick={() => onMoveRoom(r.id, 1)} disabled={i === rooms.length - 1} className="text-slate-300 hover:text-slate-600 disabled:opacity-30"><ChevronRight className="h-3.5 w-3.5 rotate-90" /></button></div>
                <button onClick={() => setOpenRoom(openRoom === r.id ? null : r.id)} className="shrink-0 rounded-md px-2 py-1 text-xs font-semibold text-rose-500 hover:bg-rose-50">розділи</button>
                <button onClick={() => { if (confirm(`Видалити «${r.name}» і всі її предмети?`)) onDeleteRoom(r.id); }} className="shrink-0 rounded-md p-1 text-slate-300 hover:text-red-500"><Trash2 className="h-4 w-4" /></button>
              </div>
              {openRoom === r.id && (
                <div className="mt-2 space-y-1.5 border-t border-slate-200 pt-2">
                  {r.sections.map((s, si) => (
                    <div key={s.id} className="flex items-center gap-2">
                      <input value={s.name} onChange={(e) => onRenameSection(r.id, s.id, e.target.value)} className="min-w-0 flex-1 rounded-lg border border-slate-200 bg-white px-2 py-1 text-xs focus:border-rose-400 focus:outline-none" />
                      <div className="flex shrink-0 flex-col"><button onClick={() => onMoveSection(r.id, s.id, -1)} disabled={si === 0} className="text-slate-300 hover:text-slate-600 disabled:opacity-30"><ChevronRight className="h-3 w-3 -rotate-90" /></button><button onClick={() => onMoveSection(r.id, s.id, 1)} disabled={si === r.sections.length - 1} className="text-slate-300 hover:text-slate-600 disabled:opacity-30"><ChevronRight className="h-3 w-3 rotate-90" /></button></div>
                      <button onClick={() => { if (confirm("Видалити розділ і його предмети?")) onDeleteSection(r.id, s.id); }} className="shrink-0 text-slate-300 hover:text-red-500"><Trash2 className="h-3.5 w-3.5" /></button>
                    </div>
                  ))}
                  <div className="flex items-center gap-2"><input value={openRoom === r.id ? newSec : ""} onChange={(e) => setNewSec(e.target.value)} placeholder="Новий розділ" className="min-w-0 flex-1 rounded-lg border border-slate-200 px-2 py-1 text-xs focus:border-rose-400 focus:outline-none" /><button onClick={() => { if (newSec.trim()) { onAddSection(r.id, newSec); setNewSec(""); } }} className="shrink-0 rounded-full bg-slate-700 px-2.5 py-1 text-xs font-semibold text-white">+</button></div>
                </div>
              )}
            </div>
          ))}
        </div>
        <div className="mt-3 flex items-center gap-2 rounded-xl border border-dashed border-slate-300 p-2">
          <input value={newRoom} onChange={(e) => setNewRoom(e.target.value)} placeholder="Нова кімната" className="min-w-0 flex-1 rounded-lg border border-slate-200 px-2 py-1.5 text-sm focus:border-rose-400 focus:outline-none" />
          <button onClick={() => { if (newRoom.trim()) { onAddRoom(newRoom); setNewRoom(""); } }} className="shrink-0 rounded-full bg-slate-800 px-3 py-1.5 text-sm font-semibold text-white hover:bg-slate-900">Додати</button>
        </div>
      </div>
    </div>
  );
}

/* ================================================================== */
/* DAILY REVIEW — one-screen morning+evening check-in (review:{date})  */
/* ================================================================== */
const REVKEYS = { index: "review:index" };
const rvKey = (d) => `review:${d}`;
async function collectReviewExport() {
  const index = await store.get(REVKEYS.index, []);
  const docs = {};
  for (const d of index) { const doc = await store.get(rvKey(d), null); if (doc) docs[d] = doc; }
  return { index, docs };
}
async function clearReviewData() {
  const index = await store.get(REVKEYS.index, []);
  for (const d of index) await store.remove(rvKey(d));
  await store.remove(REVKEYS.index);
}

function ReviewSection({ onGo }) {
  const today = dateKey(Date.now());
  const nowHour = new Date().getHours();
  const [loading, setLoading] = useState(true);
  const [mode, setMode] = useState(nowHour < 15 ? "morning" : "evening");
  const [doc, setDoc] = useState({});
  const [mood, setMood] = useState(null);
  const [routine, setRoutine] = useState({ tasks: [], completions: {}, streak: { current: 0 }, xp: 0 });
  const [fasting, setFasting] = useState(null);
  const [study, setStudy] = useState({ streak: 0, due: 0 });
  const [calmStreakVal, setCalmStreakVal] = useState(0);
  const [toast, setToast] = useState(null);
  const flash = useCallback((m) => { setToast(m); window.clearTimeout(flash._t); flash._t = window.setTimeout(() => setToast(null), 1800); }, []);

  const reload = useCallback(async () => {
    const d = await store.get(rvKey(today), {});
    setDoc(d || {});
    const r = await loadRoutineData();
    setMood((r.moods || {})[today] ?? null);
    setRoutine({ tasks: r.tasks || [], completions: r.completions || {}, streak: computeTaskStreak(r.completions || {}, today), xp: (r.xp && r.xp.xp) || 0 });
    const f = await loadFastingData();
    setFasting(f.current || null);
    const c = await loadCalmData();
    setCalmStreakVal(calmStreak(c.sessions || [], today));
    const stats = await store.get("stats", { history: {} });
    // strict due-today count across decks
    const idx = await store.get("decks:index", { decks: [] });
    const endToday = new Date(); endToday.setHours(23, 59, 59, 999); const endTs = endToday.getTime();
    let due = 0;
    for (const dk of idx.decks || []) {
      const cards = await store.get(`cards:${dk.id}`, []);
      for (const cd of cards) if ((cd.state === "learning" || cd.state === "review") && cd.due <= endTs) due += 1;
    }
    setStudy({ streak: computeStreak(stats.history || {}), due });
    setLoading(false);
  }, [today]);
  useEffect(() => {
    reload();
    const onReset = () => reload();
    window.addEventListener("review-reset", onReset);
    return () => window.removeEventListener("review-reset", onReset);
  }, [reload]);

  const saveDoc = useCallback((patch) => {
    setDoc((prev) => {
      const next = { ...prev, ...patch };
      store.set(rvKey(today), next);
      store.get(REVKEYS.index, []).then((idx) => { if (!idx.includes(today)) store.set(REVKEYS.index, [...idx, today].sort()); });
      return next;
    });
  }, [today]);
  const setMoodVal = useCallback(async (score) => {
    setMood(score);
    const prev = await store.get("routine:mood", {});
    await store.set("routine:mood", { ...prev, [today]: score });
  }, [today]);

  const occurring = useMemo(() => routine.tasks.filter((t) => taskOccursOn(t, today)), [routine.tasks, today]);
  const todayDoc = routine.completions[today] || {};
  const doneTasks = occurring.filter((t) => todayDoc.tasks?.[t.id]);
  const fastElapsedH = fasting ? (Date.now() - fasting.startTs) / 3600000 : 0;

  if (loading) return <div className="flex flex-1 items-center justify-center text-amber-400"><div className="flex flex-col items-center gap-3"><Sunrise className="h-8 w-8 animate-pulse" /><span className="text-sm">Завантаження…</span></div></div>;

  const Toggle = ({ label, value, onYes, onNo, goodWhenNo }) => (
    <div className="flex items-center justify-between rounded-2xl bg-white px-4 py-3 shadow-sm ring-1 ring-slate-100">
      <span className="text-sm font-semibold text-slate-700">{label}</span>
      <div className="flex gap-1.5">
        <button onClick={onNo} className={`rounded-full px-3 py-1 text-xs font-bold ring-1 transition ${value === false ? (goodWhenNo ? "bg-emerald-500 text-white ring-emerald-500" : "bg-slate-700 text-white ring-slate-700") : "bg-white text-slate-400 ring-slate-200"}`}>Ні</button>
        <button onClick={onYes} className={`rounded-full px-3 py-1 text-xs font-bold ring-1 transition ${value === true ? (goodWhenNo ? "bg-slate-700 text-white ring-slate-700" : "bg-emerald-500 text-white ring-emerald-500") : "bg-white text-slate-400 ring-slate-200"}`}>Так</button>
      </div>
    </div>
  );

  const MoodPicker = () => (
    <div className="rounded-2xl bg-white p-3 shadow-sm ring-1 ring-slate-100">
      <div className="mb-1.5 text-sm font-semibold text-slate-700">Як настрій?</div>
      <div className="flex justify-between">
        {MOODS.map((m) => (
          <button key={m.score} onClick={() => setMoodVal(m.score)} className={`flex flex-col items-center gap-0.5 rounded-2xl px-2 py-1.5 transition hover:scale-110 ${mood === m.score ? "bg-amber-50 ring-2 ring-amber-300" : ""}`}>
            <span className="text-2xl">{m.emoji}</span><span className="text-[9px] font-medium text-slate-400">{m.label}</span>
          </button>
        ))}
      </div>
    </div>
  );

  const StreakChips = () => (
    <div className="flex flex-wrap gap-2">
      <span className="inline-flex items-center gap-1 rounded-full bg-white px-3 py-1.5 text-xs font-bold shadow-sm ring-1 ring-slate-100"><span>🔥</span><span className="tabular-nums text-orange-500">{routine.streak.current}</span> рутина</span>
      <span className="inline-flex items-center gap-1 rounded-full bg-white px-3 py-1.5 text-xs font-bold shadow-sm ring-1 ring-slate-100"><GraduationCap className="h-3.5 w-3.5 text-rose-500" /><span className="tabular-nums text-rose-500">{study.streak}</span> навчання</span>
      <span className="inline-flex items-center gap-1 rounded-full bg-white px-3 py-1.5 text-xs font-bold shadow-sm ring-1 ring-slate-100"><Leaf className="h-3.5 w-3.5 text-teal-500" /><span className="tabular-nums text-teal-500">{calmStreakVal}</span> спокій</span>
    </div>
  );

  const Glance = () => (
    <div className="grid grid-cols-3 gap-2">
      <button onClick={() => onGo("routine")} className="rounded-2xl bg-white p-3 text-center shadow-sm ring-1 ring-slate-100 hover:ring-pink-200"><div className="text-2xl font-extrabold tabular-nums text-pink-500">{doneTasks.length}/{occurring.length}</div><div className="text-[11px] text-slate-400">справи</div></button>
      <button onClick={() => onGo("fasting")} className="rounded-2xl bg-white p-3 text-center shadow-sm ring-1 ring-slate-100 hover:ring-orange-200"><div className="text-2xl font-extrabold tabular-nums text-orange-500">{fasting ? `${Math.floor(fastElapsedH)}г` : "—"}</div><div className="text-[11px] text-slate-400">{fasting ? "голодування" : "не постишся"}</div></button>
      <button onClick={() => onGo("studying")} className="rounded-2xl bg-white p-3 text-center shadow-sm ring-1 ring-slate-100 hover:ring-rose-200"><div className="text-2xl font-extrabold tabular-nums text-rose-500">{study.due}</div><div className="text-[11px] text-slate-400">карток на сьогодні</div></button>
    </div>
  );

  return (
    <div className="min-h-screen flex-1 bg-gradient-to-b from-amber-50/50 via-red-50/30 to-white">
      <div className="mx-auto w-full max-w-2xl px-4 pb-24 pt-6">
        {/* header */}
        <div className="mb-4 flex items-center justify-between">
          <div>
            <h1 className="text-2xl font-extrabold text-slate-900">Огляд дня</h1>
            <p className="text-xs font-medium text-slate-400">{prettyDate(today)}</p>
          </div>
          <div className="flex gap-1 rounded-full bg-white p-1 shadow-sm ring-1 ring-slate-100">
            <button onClick={() => setMode("morning")} className={`rounded-full px-3 py-1.5 text-xs font-bold transition ${mode === "morning" ? "bg-amber-400 text-white" : "text-slate-400"}`}>🌅 Ранок</button>
            <button onClick={() => setMode("evening")} className={`rounded-full px-3 py-1.5 text-xs font-bold transition ${mode === "evening" ? "bg-rose-500 text-white" : "text-slate-400"}`}>🌙 Вечір</button>
          </div>
        </div>

        {mode === "morning" ? (
          <div className="space-y-3">
            <div className="rounded-3xl bg-gradient-to-r from-amber-300 to-red-300 p-4 text-white shadow-sm">
              <div className="text-lg font-extrabold">Доброго ранку 🌅</div>
              <div className="text-sm text-white/90">Хвилинка, щоб налаштуватись — без тиску.</div>
            </div>
            <StreakChips />
            <MoodPicker />
            <div className="rounded-2xl bg-white p-4 shadow-sm ring-1 ring-slate-100">
              <div className="mb-1.5 text-sm font-semibold text-slate-700">Мій намір на сьогодні</div>
              <textarea value={doc.intentions || ""} onChange={(e) => saveDoc({ intentions: e.target.value })} rows={2} placeholder="Одна річ, яка зробить день добрим…" className="w-full rounded-xl border border-slate-200 px-3 py-2 text-sm focus:border-amber-400 focus:outline-none" />
            </div>
            <div className="mb-1 mt-2 text-xs font-semibold uppercase tracking-wide text-slate-400">На сьогодні чекає</div>
            <Glance />
            <p className="pt-2 text-center text-xs text-slate-400">Не мусиш робити все. Обери одне — і почни з нього 💛</p>
          </div>
        ) : (
          <div className="space-y-3">
            <div className="rounded-3xl bg-gradient-to-r from-rose-400 to-pink-400 p-4 text-white shadow-sm">
              <div className="text-lg font-extrabold">Як пройшов день? 🌙</div>
              <div className="text-sm text-white/90">Відзначимо, що вдалося — решта зачекає.</div>
            </div>
            <StreakChips />
            <MoodPicker />

            {/* done celebration */}
            <div className="rounded-2xl bg-white p-4 shadow-sm ring-1 ring-slate-100">
              <div className="flex items-center justify-between"><span className="text-sm font-semibold text-slate-700">Закрито сьогодні</span><span className="text-sm font-extrabold text-pink-500">{doneTasks.length}{routine.xp ? ` · рівень ${levelProgress(routine.xp).lvl}` : ""}</span></div>
              {doneTasks.length > 0 ? (
                <div className="mt-2 space-y-1">{doneTasks.slice(0, 12).map((t) => <div key={t.id} className="flex items-center gap-2 text-sm text-slate-700"><span>{t.emoji || "✅"}</span><span className="truncate">{t.title}</span></div>)}</div>
              ) : <p className="mt-1 text-sm text-slate-400">Сьогодні нічого не закрито — і це ок. Завтра новий день 💛</p>}
            </div>

            {/* quick check-ins */}
            <div className="space-y-2">
              <Toggle label="Ліки прийняла?" value={doc.meds ?? null} onYes={() => saveDoc({ meds: true })} onNo={() => saveDoc({ meds: false })} />
              <Toggle label="Порухалась сьогодні?" value={doc.moved ?? null} onYes={() => saveDoc({ moved: true })} onNo={() => saveDoc({ moved: false })} />
              <Toggle label="Алкоголь сьогодні?" value={doc.alcohol ?? null} onYes={() => saveDoc({ alcohol: true })} onNo={() => saveDoc({ alcohol: false })} goodWhenNo />
              <Toggle label="Сигарети сьогодні?" value={doc.smoke ?? null} onYes={() => saveDoc({ smoke: true })} onNo={() => saveDoc({ smoke: false })} goodWhenNo />
              <Toggle label="Витрати поза планом?" value={doc.spentOver ?? null} onYes={() => saveDoc({ spentOver: true })} onNo={() => saveDoc({ spentOver: false })} goodWhenNo />
            </div>

            <div className="rounded-2xl bg-white p-4 shadow-sm ring-1 ring-slate-100">
              <div className="mb-1.5 text-sm font-semibold text-slate-700">Як був день? (одним рядком)</div>
              <textarea value={doc.note || ""} onChange={(e) => saveDoc({ note: e.target.value })} rows={2} placeholder="Що запам'яталось…" className="w-full rounded-xl border border-slate-200 px-3 py-2 text-sm focus:border-rose-400 focus:outline-none" />
            </div>

            <div className="mb-1 mt-2 text-xs font-semibold uppercase tracking-wide text-slate-400">Стан секцій</div>
            <Glance />
            <p className="pt-2 text-center text-xs text-slate-400">Ти зробила достатньо на сьогодні. Відпочинок — теж частина плану 💛</p>
          </div>
        )}
      </div>
      {toast && <div className="fixed bottom-20 left-1/2 z-40 -translate-x-1/2 rounded-full bg-slate-900 px-4 py-2 text-sm text-white shadow-lg lg:bottom-6">{toast}</div>}
    </div>
  );
}

/* ================================================================== */
/* FINANCE — debts + discretionary spending + impulse counter          */
/* ================================================================== */
const FNKEYS = {
  debts: "finance:debts",         // [{id, name, creditor, balance, start, rate, minPayment}]
  allowance: "finance:allowance", // { amount, period: "day"|"week" }
  expenses: "finance:expenses",   // [{id, date, amount, note, ts}]
  impulse: "finance:impulse",     // { since, best, resisted, slips:[dates] }
  settings: "finance:settings",   // { name, strategy }
};
const finFmt = (n) => `${Math.round((Number(n) || 0)).toLocaleString("uk-UA")} ₴`;
function finWeekStart(ds) { const d = new Date(ds + "T00:00:00"); const wd = (d.getDay() + 6) % 7; d.setDate(d.getDate() - wd); return dateKey(d.getTime()); }
function finDaysBetween(a, b) { return Math.max(0, Math.floor((new Date(b + "T00:00:00") - new Date(a + "T00:00:00")) / 86400000)); }

async function loadFinanceData() {
  const debts = await store.get(FNKEYS.debts, []);
  const allowance = await store.get(FNKEYS.allowance, { amount: 0, period: "day" });
  const expenses = await store.get(FNKEYS.expenses, []);
  const impulse = await store.get(FNKEYS.impulse, null);
  const settings = await store.get(FNKEYS.settings, { name: "Finance", strategy: "snowball" });
  return { debts, allowance, expenses, impulse, settings };
}
async function collectFinanceExport() { const d = await loadFinanceData(); return { debts: d.debts, allowance: d.allowance, expenses: d.expenses, impulse: d.impulse, settings: d.settings }; }
async function clearFinanceData() { for (const k of Object.values(FNKEYS)) await store.remove(k); }

function FinanceSection({ name, onRename, onGo, moneyTab, onMoneyTab }) {
  const today = dateKey(Date.now());
  const [loading, setLoading] = useState(true);
  const [fview, setFview] = useState("debts"); // debts | spend | impulse
  const [debts, setDebts] = useState([]);
  const [allowance, setAllowance] = useState({ amount: 0, period: "day" });
  const [expenses, setExpenses] = useState([]);
  const [impulse, setImpulse] = useState({ since: today, best: 0, resisted: 0, slips: [] });
  const [strategy, setStrategy] = useState("snowball");
  const [debtEditor, setDebtEditor] = useState(null);
  const [payFor, setPayFor] = useState(null);
  const [renaming, setRenaming] = useState(false);
  const [nameDraft, setNameDraft] = useState(name);
  const [expAmt, setExpAmt] = useState("");
  const [expNote, setExpNote] = useState("");
  const [toast, setToast] = useState(null);
  const flash = useCallback((m) => { setToast(m); window.clearTimeout(flash._t); flash._t = window.setTimeout(() => setToast(null), 2000); }, []);

  const reload = useCallback(async () => {
    const d = await loadFinanceData();
    setDebts(d.debts || []); setAllowance(d.allowance || { amount: 0, period: "day" }); setExpenses(d.expenses || []);
    setImpulse(d.impulse || { since: today, best: 0, resisted: 0, slips: [] });
    setStrategy(d.settings?.strategy || "snowball");
    if (!d.impulse) await store.set(FNKEYS.impulse, { since: today, best: 0, resisted: 0, slips: [] });
    setLoading(false);
  }, [today]);
  useEffect(() => { reload(); const onR = () => reload(); window.addEventListener("finance-reset", onR); return () => window.removeEventListener("finance-reset", onR); }, [reload]);

  const saveDebts = useCallback((n) => { setDebts(n); store.set(FNKEYS.debts, n); }, []);
  const saveAllowance = useCallback((n) => { setAllowance(n); store.set(FNKEYS.allowance, n); }, []);
  const saveExpenses = useCallback((n) => { setExpenses(n); store.set(FNKEYS.expenses, n); }, []);
  const saveImpulse = useCallback((n) => { setImpulse(n); store.set(FNKEYS.impulse, n); }, []);
  const saveStrategy = useCallback(async (st) => { setStrategy(st); const prev = await store.get(FNKEYS.settings, { name: "Finance" }); store.set(FNKEYS.settings, { ...prev, strategy: st }); }, []);

  // debts
  const upsertDebt = (meta, id) => { if (id) saveDebts(debts.map((x) => (x.id === id ? { ...x, ...meta } : x))); else saveDebts([...debts, { id: ruid("fd"), start: meta.balance, ...meta }]); };
  const deleteDebt = (id) => saveDebts(debts.filter((x) => x.id !== id));
  const logPayment = (id, amt) => saveDebts(debts.map((x) => (x.id === id ? { ...x, balance: Math.max(0, (Number(x.balance) || 0) - amt) } : x)));

  const totalDebt = debts.reduce((s, x) => s + (Number(x.balance) || 0), 0);
  const totalStart = debts.reduce((s, x) => s + (Number(x.start) || Number(x.balance) || 0), 0);
  const totalPaid = totalStart - totalDebt;
  const ordered = [...debts].sort((a, b) => strategy === "snowball" ? (a.balance - b.balance) : (b.rate - a.rate)).sort((a, b) => (a.balance <= 0) - (b.balance <= 0));
  const focus = ordered.find((x) => x.balance > 0);

  // allowance / expenses
  const periodStart = allowance.period === "week" ? finWeekStart(today) : today;
  const periodExp = expenses.filter((e) => e.date >= periodStart);
  const periodSpent = periodExp.reduce((s, e) => s + (Number(e.amount) || 0), 0);
  const periodRemaining = (Number(allowance.amount) || 0) - periodSpent;
  const logExpense = () => { const a = Number(expAmt); if (!a) return; saveExpenses([{ id: ruid("fe"), date: today, amount: a, note: expNote.trim(), ts: Date.now() }, ...expenses]); setExpAmt(""); setExpNote(""); flash("Записано"); };
  const delExpense = (id) => saveExpenses(expenses.filter((e) => e.id !== id));

  // impulse
  const impDays = finDaysBetween(impulse.since, today);
  const resisted = () => { saveImpulse({ ...impulse, resisted: (impulse.resisted || 0) + 1 }); flash("Молодець — це перемога 💪"); };
  const slipped = () => { saveImpulse({ ...impulse, since: today, best: Math.max(impulse.best || 0, impDays), slips: [...(impulse.slips || []), today] }); flash("Це трапляється. Завтра — новий день 💛"); };

  if (loading) return <div className="flex flex-1 items-center justify-center text-emerald-500"><div className="flex flex-col items-center gap-3"><Wallet className="h-8 w-8 animate-pulse" /><span className="text-sm">Завантаження…</span></div></div>;

  return (
    <div className="min-h-screen flex-1 bg-gradient-to-b from-emerald-50/40 to-white">
      <header className="sticky top-0 z-20 border-b border-slate-200 bg-white/85 backdrop-blur">
        <div className="mx-auto flex h-14 w-full max-w-3xl items-center gap-2 px-4">
          {renaming ? (
            <input autoFocus value={nameDraft} onChange={(e) => setNameDraft(e.target.value)} onBlur={() => { onRename(nameDraft); setRenaming(false); }} onKeyDown={(e) => { if (e.key === "Enter") { onRename(nameDraft); setRenaming(false); } }} className="mr-auto w-32 rounded-lg border border-emerald-200 px-2 py-1 text-base font-semibold focus:outline-none" />
          ) : (
            <button onClick={() => { setNameDraft(name); setRenaming(true); }} className="mr-auto text-base font-semibold text-slate-900">{name} <Pencil className="ml-0.5 inline h-3.5 w-3.5 text-slate-300" /></button>
          )}
          {onMoneyTab && <MoneyToggle active={moneyTab} onSet={onMoneyTab} />}
        </div>
      </header>

      <main className="mx-auto w-full max-w-3xl px-4 py-5">
        <p className="mb-3 rounded-2xl bg-emerald-50/70 px-3 py-2 text-xs leading-relaxed text-emerald-800">Гроші й тривога часто ходять поруч. Тут — спокійно, крок за кроком, без осуду. Мета не «ідеально», а трохи ясніше.</p>

        <div className="mb-4 flex gap-2 rounded-2xl bg-white p-1 shadow-sm ring-1 ring-emerald-100">
          {[["debts", "Борги", TrendingDown], ["spend", "Кишеня", Coffee], ["impulse", "Стійкість", ShieldAlert]].map(([k, label, Icon]) => (
            <button key={k} onClick={() => setFview(k)} className={`flex flex-1 items-center justify-center gap-1.5 rounded-xl py-2 text-sm font-bold transition ${fview === k ? "bg-emerald-500 text-white shadow" : "text-slate-500 hover:text-slate-700"}`}><Icon className="h-4 w-4" /> {label}</button>
          ))}
        </div>

        {fview === "debts" && (
          <div className="space-y-3">
            <div className="rounded-3xl bg-white p-5 shadow-sm ring-1 ring-emerald-100">
              <div className="text-sm text-slate-400">Загальний борг</div>
              <div className="text-3xl font-extrabold tabular-nums text-slate-900">{finFmt(totalDebt)}</div>
              {totalStart > 0 && <>
                <div className="mt-2 h-2 overflow-hidden rounded-full bg-emerald-50"><div className="h-full rounded-full bg-emerald-500 transition-all" style={{ width: `${Math.min(100, (totalPaid / totalStart) * 100)}%` }} /></div>
                <div className="mt-1 text-xs text-slate-400">погашено {finFmt(totalPaid)} з {finFmt(totalStart)}</div>
              </>}
            </div>

            {debts.length > 0 && (
              <div className="rounded-2xl bg-white p-3 shadow-sm ring-1 ring-emerald-50">
                <div className="mb-1.5 text-xs font-semibold text-slate-500">Стратегія погашення</div>
                <div className="flex gap-2">
                  {[["snowball", "Сніжинка", "спершу найменший борг — швидкі перемоги"], ["avalanche", "Лавина", "спершу найдорожчий % — менше переплати"]].map(([k, label, desc]) => (
                    <button key={k} onClick={() => saveStrategy(k)} className={`flex-1 rounded-xl p-2 text-left ring-1 transition ${strategy === k ? "bg-emerald-50 ring-emerald-300" : "bg-white ring-slate-200"}`}><div className="text-sm font-bold text-slate-800">{label}</div><div className="text-[11px] leading-tight text-slate-400">{desc}</div></button>
                  ))}
                </div>
              </div>
            )}

            {ordered.map((d) => { const paid = (Number(d.start) || d.balance) - d.balance; const pct = d.start > 0 ? Math.min(1, paid / d.start) : 0; const done = d.balance <= 0; const isFocus = focus && d.id === focus.id; return (
              <div key={d.id} className={`rounded-2xl bg-white p-4 shadow-sm ring-1 ${isFocus ? "ring-2 ring-emerald-300" : "ring-emerald-50"}`}>
                <div className="flex items-start justify-between gap-2">
                  <div className="min-w-0">
                    <div className="flex items-center gap-2"><span className="truncate font-bold text-slate-800">{d.name}</span>{isFocus && <span className="shrink-0 rounded-full bg-emerald-500 px-2 py-0.5 text-[10px] font-bold text-white">фокус зараз</span>}{done && <span className="shrink-0 rounded-full bg-green-100 px-2 py-0.5 text-[10px] font-bold text-green-700">погашено 🎉</span>}</div>
                    {d.creditor && <div className="text-xs text-slate-400">кому: {d.creditor}</div>}
                  </div>
                  <div className="shrink-0 text-right"><div className="font-extrabold tabular-nums text-slate-900">{finFmt(d.balance)}</div>{d.rate > 0 && <div className="text-[11px] text-slate-400">{d.rate}% · мін {finFmt(d.minPayment)}</div>}</div>
                </div>
                <div className="mt-2 h-1.5 overflow-hidden rounded-full bg-slate-100"><div className="h-full rounded-full bg-emerald-400" style={{ width: `${pct * 100}%` }} /></div>
                <div className="mt-2 flex items-center gap-2">
                  {!done && <button onClick={() => setPayFor(d)} className="rounded-full bg-emerald-500 px-3 py-1 text-xs font-semibold text-white hover:bg-emerald-600">Внести оплату</button>}
                  <button onClick={() => setDebtEditor({ debt: d })} className="rounded-full bg-white px-3 py-1 text-xs font-semibold text-slate-500 ring-1 ring-slate-200 hover:bg-slate-50">Редагувати</button>
                  <button onClick={() => { if (confirm("Видалити борг?")) deleteDebt(d.id); }} className="ml-auto rounded-md p-1 text-slate-300 hover:text-red-500"><Trash2 className="h-4 w-4" /></button>
                </div>
                {isFocus && !done && <div className="mt-2 rounded-lg bg-emerald-50 px-3 py-1.5 text-[11px] text-emerald-700">Плати мінімум по всіх, а <b>сюди</b> — усе, що зможеш зверху. Один фокус за раз.</div>}
              </div>
            ); })}

            <button onClick={() => setDebtEditor({ debt: null })} className="flex w-full items-center justify-center gap-1.5 rounded-2xl border border-dashed border-emerald-300 py-3 text-sm font-semibold text-emerald-600 hover:bg-emerald-50"><Plus className="h-4 w-4" /> Додати борг</button>
          </div>
        )}

        {fview === "spend" && (
          <div className="space-y-3">
            <div className="rounded-3xl bg-white p-5 shadow-sm ring-1 ring-emerald-100">
              <div className="mb-2 flex items-center justify-between">
                <span className="text-sm font-semibold text-slate-600">Ліміт на дрібні витрати</span>
                <div className="flex gap-1 rounded-full bg-slate-100 p-0.5">
                  {[["day", "день"], ["week", "тиждень"]].map(([k, l]) => <button key={k} onClick={() => saveAllowance({ ...allowance, period: k })} className={`rounded-full px-2.5 py-1 text-xs font-bold ${allowance.period === k ? "bg-emerald-500 text-white" : "text-slate-500"}`}>{l}</button>)}
                </div>
              </div>
              <div className="flex items-center gap-2"><input type="number" min={0} value={allowance.amount || ""} onChange={(e) => saveAllowance({ ...allowance, amount: Math.max(0, +e.target.value || 0) })} placeholder="0" className="w-28 rounded-lg border border-emerald-200 px-2 py-1 text-right text-sm font-bold tabular-nums focus:border-emerald-400 focus:outline-none" /><span className="text-sm font-bold text-emerald-700">₴ / {allowance.period === "week" ? "тиждень" : "день"}</span></div>
              <div className={`mt-3 rounded-2xl p-3 text-center ${periodRemaining < 0 ? "bg-amber-50" : "bg-emerald-50"}`}>
                <div className={`text-[11px] font-medium ${periodRemaining < 0 ? "text-amber-600" : "text-emerald-600"}`}>{periodRemaining < 0 ? "перевитрата" : "лишилось"} на {allowance.period === "week" ? "цей тиждень" : "сьогодні"}</div>
                <div className={`text-2xl font-extrabold tabular-nums ${periodRemaining < 0 ? "text-amber-600" : "text-emerald-600"}`}>{finFmt(Math.abs(periodRemaining))}</div>
                <div className="text-[11px] text-slate-400">витрачено {finFmt(periodSpent)}</div>
              </div>
            </div>
            <div className="rounded-2xl bg-white p-3 shadow-sm ring-1 ring-emerald-50">
              <div className="flex gap-2"><input type="number" min={0} value={expAmt} onChange={(e) => setExpAmt(e.target.value)} placeholder="Сума ₴" className="w-24 rounded-lg border border-slate-300 px-2 py-2 text-sm focus:border-emerald-400 focus:outline-none" /><input value={expNote} onChange={(e) => setExpNote(e.target.value)} placeholder="На що? (необов'язково)" className="min-w-0 flex-1 rounded-lg border border-slate-300 px-2 py-2 text-sm focus:border-emerald-400 focus:outline-none" onKeyDown={(e) => { if (e.key === "Enter") logExpense(); }} /><button onClick={logExpense} className="shrink-0 rounded-lg bg-emerald-500 px-3 text-sm font-semibold text-white hover:bg-emerald-600">+</button></div>
            </div>
            {periodExp.length > 0 && (
              <div className="rounded-2xl bg-white p-3 shadow-sm ring-1 ring-emerald-50">
                <div className="mb-1.5 text-xs font-semibold text-slate-500">Витрати за період</div>
                <div className="space-y-1">{periodExp.map((e) => <div key={e.id} className="group flex items-center gap-2 text-sm"><span className="tabular-nums font-semibold text-slate-700">{finFmt(e.amount)}</span><span className="min-w-0 flex-1 truncate text-slate-400">{e.note || "—"}</span><span className="text-[11px] text-slate-300">{e.date.slice(5)}</span><button onClick={() => delExpense(e.id)} className="text-slate-300 hover:text-red-500 sm:opacity-0 sm:group-hover:opacity-100"><X className="h-3.5 w-3.5" /></button></div>)}</div>
              </div>
            )}
            <button onClick={() => onGo && onGo("budget")} className="flex w-full items-center justify-center gap-1.5 rounded-2xl bg-white py-2.5 text-sm font-semibold text-slate-500 shadow-sm ring-1 ring-slate-200 hover:ring-emerald-200"><ShoppingCart className="h-4 w-4" /> Щомісячні покупки — у вкладці Budget</button>
          </div>
        )}

        {fview === "impulse" && (
          <div className="space-y-3">
            <div className="rounded-3xl bg-gradient-to-br from-emerald-400 to-teal-400 p-6 text-center text-white shadow-sm">
              <div className="text-5xl">🛡️</div>
              <div className="mt-2 text-5xl font-black tabular-nums">{impDays}</div>
              <div className="text-sm font-semibold text-white/90">{impDays === 1 ? "день" : "днів"} без імпульсивної покупки</div>
              {impulse.best > 0 && <div className="mt-1 text-xs text-white/70">рекорд: {impulse.best} дн</div>}
            </div>
            <div className="grid grid-cols-2 gap-2">
              <button onClick={resisted} className="rounded-2xl bg-white p-4 text-center shadow-sm ring-1 ring-emerald-100 hover:ring-emerald-300"><div className="text-2xl">💪</div><div className="mt-1 text-sm font-bold text-slate-700">Я втрималась</div><div className="text-[11px] text-slate-400">втримань: {impulse.resisted || 0}</div></button>
              <button onClick={() => { if (confirm("Позначити зрив? Стрік почнеться заново — без осуду.")) slipped(); }} className="rounded-2xl bg-white p-4 text-center shadow-sm ring-1 ring-slate-100 hover:ring-slate-300"><div className="text-2xl">🌱</div><div className="mt-1 text-sm font-bold text-slate-700">Був зрив</div><div className="text-[11px] text-slate-400">почати заново</div></button>
            </div>
            <p className="rounded-2xl bg-slate-50 px-4 py-3 text-center text-xs leading-relaxed text-slate-500">Зрив — не провал, а дані. Поміть, що передувало пориву (втома? нудьга? стрес?), і наступного разу буде легше. Ти вчишся, а не «не впоралась».</p>
          </div>
        )}
      </main>

      {debtEditor && <FinDebtEditor debt={debtEditor.debt} onClose={() => setDebtEditor(null)} onDelete={debtEditor.debt ? () => { deleteDebt(debtEditor.debt.id); setDebtEditor(null); } : null} onSave={(meta) => { upsertDebt(meta, debtEditor.debt?.id); setDebtEditor(null); }} />}
      {payFor && <FinPayModal debt={payFor} onClose={() => setPayFor(null)} onPay={(amt) => { logPayment(payFor.id, amt); setPayFor(null); flash("Оплату записано 👏"); }} />}

      {toast && <div className="fixed bottom-20 left-1/2 z-40 -translate-x-1/2 rounded-full bg-slate-900 px-4 py-2 text-sm text-white shadow-lg lg:bottom-6">{toast}</div>}
    </div>
  );
}

function FinDebtEditor({ debt, onClose, onSave, onDelete }) {
  const [name, setName] = useState(debt?.name || "");
  const [creditor, setCreditor] = useState(debt?.creditor || "");
  const [balance, setBalance] = useState(debt?.balance ?? "");
  const [rate, setRate] = useState(debt?.rate ?? "");
  const [minPayment, setMin] = useState(debt?.minPayment ?? "");
  return (
    <div className="fixed inset-0 z-50 flex items-end justify-center bg-slate-900/40 p-0 backdrop-blur-sm sm:items-center sm:p-4" onClick={onClose}>
      <div className="w-full max-w-md rounded-t-3xl bg-white p-5 shadow-xl sm:rounded-3xl" onClick={(e) => e.stopPropagation()}>
        <div className="mb-3 flex items-center justify-between"><h3 className="text-lg font-bold text-slate-900">{debt ? "Редагувати борг" : "Новий борг"}</h3><button onClick={onClose} className="rounded-md p-1 text-slate-400 hover:bg-slate-100"><X className="h-5 w-5" /></button></div>
        <input value={name} onChange={(e) => setName(e.target.value)} placeholder="Назва (напр. кредитка, позика від мами)" className="mb-2 w-full rounded-xl border border-slate-300 px-3 py-2 text-sm font-semibold focus:border-emerald-400 focus:outline-none" />
        <input value={creditor} onChange={(e) => setCreditor(e.target.value)} placeholder="Кому винна (необов'язково)" className="mb-2 w-full rounded-xl border border-slate-300 px-3 py-2 text-sm focus:border-emerald-400 focus:outline-none" />
        <div className="mb-3 grid grid-cols-3 gap-2">
          <label className="block"><span className="mb-1 block text-xs text-slate-500">Баланс ₴</span><input type="number" min={0} value={balance} onChange={(e) => setBalance(e.target.value)} className="w-full rounded-lg border border-slate-300 px-2 py-1.5 text-sm focus:border-emerald-400 focus:outline-none" /></label>
          <label className="block"><span className="mb-1 block text-xs text-slate-500">Ставка %</span><input type="number" min={0} step="any" value={rate} onChange={(e) => setRate(e.target.value)} className="w-full rounded-lg border border-slate-300 px-2 py-1.5 text-sm focus:border-emerald-400 focus:outline-none" /></label>
          <label className="block"><span className="mb-1 block text-xs text-slate-500">Мін. платіж ₴</span><input type="number" min={0} value={minPayment} onChange={(e) => setMin(e.target.value)} className="w-full rounded-lg border border-slate-300 px-2 py-1.5 text-sm focus:border-emerald-400 focus:outline-none" /></label>
        </div>
        <div className="flex gap-2">
          <button onClick={() => { if (name.trim()) onSave({ name: name.trim(), creditor: creditor.trim(), balance: Number(balance) || 0, rate: Number(rate) || 0, minPayment: Number(minPayment) || 0 }); }} className="flex-1 rounded-2xl bg-emerald-500 py-3 font-bold text-white hover:bg-emerald-600">Зберегти</button>
          {onDelete && <button onClick={onDelete} className="rounded-2xl bg-red-50 px-4 py-3 font-semibold text-red-500 hover:bg-red-100"><Trash2 className="h-5 w-5" /></button>}
        </div>
      </div>
    </div>
  );
}

function FinPayModal({ debt, onClose, onPay }) {
  const [amt, setAmt] = useState(debt.minPayment || "");
  return (
    <div className="fixed inset-0 z-50 flex items-end justify-center bg-slate-900/40 p-0 backdrop-blur-sm sm:items-center sm:p-4" onClick={onClose}>
      <div className="w-full max-w-sm rounded-t-3xl bg-white p-5 shadow-xl sm:rounded-3xl" onClick={(e) => e.stopPropagation()}>
        <h3 className="mb-1 text-lg font-bold text-slate-900">Оплата боргу</h3>
        <p className="mb-3 text-sm text-slate-500">«{debt.name}» · зараз {finFmt(debt.balance)}</p>
        <input autoFocus type="number" min={0} value={amt} onChange={(e) => setAmt(e.target.value)} placeholder="Сума ₴" className="mb-3 w-full rounded-xl border border-slate-300 px-3 py-2 text-sm font-semibold focus:border-emerald-400 focus:outline-none" onKeyDown={(e) => { if (e.key === "Enter" && Number(amt) > 0) onPay(Number(amt)); }} />
        <button onClick={() => { if (Number(amt) > 0) onPay(Number(amt)); }} className="w-full rounded-2xl bg-emerald-500 py-3 font-bold text-white hover:bg-emerald-600">Записати оплату</button>
      </div>
    </div>
  );
}

/* ---------- Recovery (inside Calm): sobriety + triggers + urge support ---------- */
const REC_MOODS = ["спокій", "тривога", "сум", "злість", "нудьга", "втома", "радість", "стрес"];
const REC_COMPANY = ["наодинці", "з друзями", "з родиною", "на людях", "на роботі"];
const REC_TIME = ["ранок", "день", "вечір", "ніч"];
const REC_SUBST = { alcohol: { label: "Алкоголь", emoji: "🍷" }, smoke: { label: "Нікотин", emoji: "🚬" }, other: { label: "Інше", emoji: "•" } };
function recDefault() { return { since: dateKey(Date.now()), best: 0, slips: [] }; }

function RecoveryView({ onExit, onQuickCalm }) {
  const today = dateKey(Date.now());
  const [alcohol, setAlcohol] = useState(recDefault());
  const [smoke, setSmoke] = useState(recDefault());
  const [triggers, setTriggers] = useState([]);
  const [reason, setReason] = useState("");
  const [noteSeen, setNoteSeen] = useState(false);
  const [urgeOpen, setUrgeOpen] = useState(false);
  const [logForm, setLogForm] = useState(null); // { substance, type }
  const [toast, setToast] = useState(null);
  const flash = (m) => { setToast(m); window.clearTimeout(flash._t); flash._t = window.setTimeout(() => setToast(null), 2200); };

  const load = useCallback(async () => {
    setAlcohol((await store.get(RECKEYS.alcohol, null)) || recDefault());
    setSmoke((await store.get(RECKEYS.smoke, null)) || recDefault());
    setTriggers(await store.get(RECKEYS.triggers, []));
    setReason(await store.get(RECKEYS.reason, ""));
    setNoteSeen(await store.get(RECKEYS.noteSeen, false));
  }, []);
  useEffect(() => { load(); const onR = () => load(); window.addEventListener("calm-reset", onR); return () => window.removeEventListener("calm-reset", onR); }, [load]);

  const saveAlcohol = (n) => { setAlcohol(n); store.set(RECKEYS.alcohol, n); };
  const saveSmoke = (n) => { setSmoke(n); store.set(RECKEYS.smoke, n); };
  const saveTriggers = (n) => { setTriggers(n); store.set(RECKEYS.triggers, n); };
  const saveReason = (v) => { setReason(v); store.set(RECKEYS.reason, v); };
  const dismissNote = () => { setNoteSeen(true); store.set(RECKEYS.noteSeen, true); };

  const resetCounter = (which) => {
    if (which === "alcohol") { const days = finDaysBetween(alcohol.since, today); saveAlcohol({ since: today, best: Math.max(alcohol.best || 0, days), slips: [...(alcohol.slips || []), today] }); }
    else { const days = finDaysBetween(smoke.since, today); saveSmoke({ since: today, best: Math.max(smoke.best || 0, days), slips: [...(smoke.slips || []), today] }); }
  };
  const saveLog = (entry) => {
    saveTriggers([{ id: ruid("rt"), date: today, ts: Date.now(), ...entry }, ...triggers]);
    if (entry.type === "slip" && (entry.substance === "alcohol" || entry.substance === "smoke")) resetCounter(entry.substance);
    setLogForm(null);
    flash(entry.type === "slip" ? "Записано. Завтра — новий день 💛" : "Записано. Дякую, що поставила паузу 💪");
  };

  // pattern summary
  const patterns = useMemo(() => {
    if (triggers.length < 2) return null;
    const top = (key) => { const c = {}; triggers.forEach((t) => { (t[key] || []).forEach ? (t[key] || []).forEach((v) => c[v] = (c[v] || 0) + 1) : (t[key] && (c[t[key]] = (c[t[key]] || 0) + 1)); }); const e = Object.entries(c).sort((a, b) => b[1] - a[1])[0]; return e ? e[0] : null; };
    return { time: top("time"), mood: top("moods"), company: top("company") };
  }, [triggers]);

  const Counter = ({ label, emoji, data, which, color }) => {
    const days = finDaysBetween(data.since, today);
    return (
      <div className="rounded-2xl bg-white p-4 text-center shadow-sm ring-1 ring-teal-50">
        <div className="text-2xl">{emoji}</div>
        <div className="text-3xl font-black tabular-nums" style={{ color }}>{days}</div>
        <div className="text-xs font-semibold text-slate-500">{label}</div>
        <div className="text-[11px] text-slate-400">{days === 1 ? "день" : "днів"}{data.best > 0 ? ` · рекорд ${data.best}` : ""}</div>
        <button onClick={() => setLogForm({ substance: which, type: "slip" })} className="mt-2 w-full rounded-full bg-slate-50 py-1.5 text-[11px] font-semibold text-slate-400 hover:bg-slate-100">був зрив</button>
      </div>
    );
  };

  return (
    <div className="mx-auto w-full max-w-lg px-4 pb-16 pt-6">
      <CalmHeader title="Відновлення" onExit={onExit} />
      <p className="mb-4 text-sm leading-relaxed text-slate-500">Підтримка, а не контроль. Тут без осуду — кожен день рахується, а зрив не перекреслює прогресу.</p>

      {/* urge button */}
      <button onClick={() => setUrgeOpen(true)} className="mb-4 flex w-full items-center gap-3 rounded-3xl bg-gradient-to-r from-teal-500 to-emerald-500 p-4 text-left text-white shadow-lg shadow-teal-500/20 transition hover:brightness-105">
        <span className="grid h-12 w-12 shrink-0 place-items-center rounded-2xl bg-white/20"><HandHeart className="h-6 w-6" /></span>
        <span className="flex-1"><span className="block text-lg font-bold">Мені хочеться вжити зараз</span><span className="block text-sm text-white/90">Натисни — перечекаємо разом.</span></span>
        <ArrowRight className="h-5 w-5" />
      </button>

      {/* counters */}
      <div className="grid grid-cols-2 gap-3">
        <Counter label="без алкоголю" emoji="🍷" data={alcohol} which="alcohol" color="#0d9488" />
        <Counter label="без нікотину" emoji="🚬" data={smoke} which="smoke" color="#0ea5e9" />
      </div>

      {/* reason */}
      <div className="mt-3 rounded-2xl bg-white p-4 shadow-sm ring-1 ring-teal-50">
        <div className="mb-1.5 text-sm font-semibold text-slate-700">Моя причина</div>
        <textarea value={reason} onChange={(e) => saveReason(e.target.value)} rows={2} placeholder="Заради чого я це роблю? (побачиш це в мить пориву)" className="w-full rounded-xl border border-slate-200 px-3 py-2 text-sm focus:border-teal-400 focus:outline-none" />
      </div>

      {/* trigger journal */}
      <div className="mt-4 flex items-center justify-between">
        <h2 className="text-sm font-bold text-slate-700">Тригер-журнал</h2>
        <button onClick={() => setLogForm({ substance: "alcohol", type: "urge" })} className="inline-flex items-center gap-1 rounded-full bg-teal-500 px-3 py-1.5 text-xs font-semibold text-white hover:bg-teal-600"><Plus className="h-3.5 w-3.5" /> Записати</button>
      </div>
      {patterns && (patterns.time || patterns.mood || patterns.company) && (
        <div className="mt-2 rounded-2xl bg-teal-50/70 px-3 py-2 text-xs text-teal-800">Найчастіше пориви: {[patterns.time, patterns.mood, patterns.company].filter(Boolean).join(" · ")}. Помічати — вже половина справи.</div>
      )}
      <div className="mt-2 space-y-2">
        {triggers.length === 0 ? (
          <div className="rounded-2xl bg-white py-8 text-center text-sm text-slate-400 ring-1 ring-teal-50">Порожньо. Після пориву чи зриву — запиши, що передувало. З часом побачиш свої патерни.</div>
        ) : triggers.slice(0, 30).map((t) => (
          <div key={t.id} className="rounded-2xl bg-white p-3 shadow-sm ring-1 ring-teal-50">
            <div className="flex items-center gap-2 text-sm">
              <span>{REC_SUBST[t.substance]?.emoji}</span>
              <span className={`rounded-full px-2 py-0.5 text-[10px] font-bold ${t.type === "slip" ? "bg-amber-100 text-amber-700" : "bg-teal-100 text-teal-700"}`}>{t.type === "slip" ? "зрив" : "порив"}</span>
              <span className="ml-auto text-[11px] text-slate-400">{t.date.slice(5)}</span>
            </div>
            {(t.moods?.length || t.company?.length || t.time || t.stress) && <div className="mt-1 flex flex-wrap gap-1 text-[11px] text-slate-500">{[t.time, t.stress && `стрес: ${t.stress}`, ...(t.company || []), ...(t.moods || [])].filter(Boolean).map((x, i) => <span key={i} className="rounded-full bg-slate-100 px-2 py-0.5">{x}</span>)}</div>}
            {t.note && <div className="mt-1 text-xs text-slate-500">{t.note}</div>}
          </div>
        ))}
      </div>

      {/* professional note */}
      {!noteSeen && (
        <div className="mt-4 flex items-start gap-2 rounded-2xl bg-slate-100/70 px-4 py-3 text-sm text-slate-500">
          <Info className="mt-0.5 h-4 w-4 shrink-0 text-slate-400" />
          <p className="flex-1">Ці інструменти підтримують, але не замінюють фахову допомогу — лікаря, нарколога чи групу. Звернутися по підтримку — це сила, а не слабкість. 💛</p>
          <button onClick={dismissNote} className="rounded-full p-0.5 text-slate-300 hover:text-slate-500"><X className="h-4 w-4" /></button>
        </div>
      )}

      {/* urge overlay */}
      {urgeOpen && (
        <div className="fixed inset-0 z-50 flex items-end justify-center bg-slate-900/50 p-4 backdrop-blur-sm sm:items-center" onClick={() => setUrgeOpen(false)}>
          <div className="w-full max-w-md rounded-3xl bg-white p-5 shadow-2xl" onClick={(e) => e.stopPropagation()}>
            <div className="text-center">
              <div className="text-3xl">🌊</div>
              <h2 className="mt-1 text-lg font-extrabold text-slate-800">Пориви минають хвилями</h2>
              <p className="mt-1 text-sm text-slate-500">Не мусиш боротися — просто перечекай хвилю. За кілька хвилин відпустить.</p>
            </div>
            {reason.trim() && <div className="mt-3 rounded-2xl bg-teal-50 px-4 py-3 text-center text-sm text-teal-800"><span className="text-[11px] font-semibold uppercase tracking-wide text-teal-500">Моя причина</span><div className="mt-0.5">{reason}</div></div>}
            <div className="mt-4 space-y-2">
              <button onClick={() => { setUrgeOpen(false); onQuickCalm("breath"); }} className="flex w-full items-center gap-3 rounded-2xl bg-white p-3 text-left ring-1 ring-slate-100 hover:ring-sky-200"><span className="grid h-10 w-10 place-items-center rounded-xl bg-sky-500 text-white"><Wind className="h-5 w-5" /></span><span className="flex-1"><span className="block font-bold text-slate-800">Подихати</span><span className="block text-xs text-slate-400">Сповільнити тіло за 2 хвилини</span></span><ArrowRight className="h-4 w-4 text-slate-300" /></button>
              <button onClick={() => { setUrgeOpen(false); onQuickCalm("ground"); }} className="flex w-full items-center gap-3 rounded-2xl bg-white p-3 text-left ring-1 ring-slate-100 hover:ring-rose-200"><span className="grid h-10 w-10 place-items-center rounded-xl bg-rose-500 text-white"><Anchor className="h-5 w-5" /></span><span className="flex-1"><span className="block font-bold text-slate-800">Заземлитися 5-4-3-2-1</span><span className="block text-xs text-slate-400">Повернутись у тіло й у момент</span></span><ArrowRight className="h-4 w-4 text-slate-300" /></button>
              <button onClick={() => { setUrgeOpen(false); setLogForm({ substance: "alcohol", type: "urge" }); }} className="flex w-full items-center gap-3 rounded-2xl bg-white p-3 text-left ring-1 ring-slate-100 hover:ring-teal-200"><span className="grid h-10 w-10 place-items-center rounded-xl bg-teal-500 text-white"><NotebookPen className="h-5 w-5" /></span><span className="flex-1"><span className="block font-bold text-slate-800">Записати цей порив</span><span className="block text-xs text-slate-400">Що зараз коїться — для патернів</span></span><ArrowRight className="h-4 w-4 text-slate-300" /></button>
            </div>
            <button onClick={() => setUrgeOpen(false)} className="mt-3 w-full rounded-2xl py-2.5 text-sm font-semibold text-slate-400 hover:text-slate-600">Мені вже легше</button>
          </div>
        </div>
      )}

      {logForm && <RecoveryLog init={logForm} onClose={() => setLogForm(null)} onSave={saveLog} />}
      {toast && <div className="fixed bottom-6 left-1/2 z-[60] -translate-x-1/2 rounded-full bg-slate-900 px-4 py-2 text-sm text-white shadow-lg">{toast}</div>}
    </div>
  );
}

function RecoveryLog({ init, onClose, onSave }) {
  const [substance, setSubstance] = useState(init.substance || "alcohol");
  const [type, setType] = useState(init.type || "urge");
  const [time, setTime] = useState("");
  const [stress, setStress] = useState("");
  const [company, setCompany] = useState([]);
  const [moods, setMoods] = useState([]);
  const [note, setNote] = useState("");
  const toggle = (arr, v, set) => set(arr.includes(v) ? arr.filter((x) => x !== v) : [...arr, v]);
  const Chip = ({ on, onClick, children }) => <button onClick={onClick} className={`rounded-full px-3 py-1 text-xs font-semibold ring-1 transition ${on ? "bg-teal-500 text-white ring-teal-500" : "bg-white text-slate-500 ring-slate-200"}`}>{children}</button>;
  return (
    <div className="fixed inset-0 z-[55] flex items-end justify-center bg-slate-900/40 p-0 backdrop-blur-sm sm:items-center sm:p-4" onClick={onClose}>
      <div className="max-h-[92vh] w-full max-w-md overflow-y-auto rounded-t-3xl bg-white p-5 shadow-xl sm:rounded-3xl" onClick={(e) => e.stopPropagation()}>
        <div className="mb-3 flex items-center justify-between"><h3 className="text-lg font-bold text-slate-900">Що передувало?</h3><button onClick={onClose} className="rounded-md p-1 text-slate-400 hover:bg-slate-100"><X className="h-5 w-5" /></button></div>
        <div className="mb-3 flex gap-2">
          {Object.entries(REC_SUBST).map(([k, v]) => <Chip key={k} on={substance === k} onClick={() => setSubstance(k)}>{v.emoji} {v.label}</Chip>)}
        </div>
        <div className="mb-3 flex gap-2">
          <Chip on={type === "urge"} onClick={() => setType("urge")}>Порив (втрималась)</Chip>
          <Chip on={type === "slip"} onClick={() => setType("slip")}>Зрив</Chip>
        </div>
        <div className="mb-2 text-xs font-semibold text-slate-500">Коли</div>
        <div className="mb-3 flex flex-wrap gap-2">{REC_TIME.map((t) => <Chip key={t} on={time === t} onClick={() => setTime(time === t ? "" : t)}>{t}</Chip>)}</div>
        <div className="mb-2 text-xs font-semibold text-slate-500">Стрес</div>
        <div className="mb-3 flex flex-wrap gap-2">{["низький", "середній", "високий"].map((s) => <Chip key={s} on={stress === s} onClick={() => setStress(stress === s ? "" : s)}>{s}</Chip>)}</div>
        <div className="mb-2 text-xs font-semibold text-slate-500">З ким</div>
        <div className="mb-3 flex flex-wrap gap-2">{REC_COMPANY.map((c) => <Chip key={c} on={company.includes(c)} onClick={() => toggle(company, c, setCompany)}>{c}</Chip>)}</div>
        <div className="mb-2 text-xs font-semibold text-slate-500">Настрій / стан</div>
        <div className="mb-3 flex flex-wrap gap-2">{REC_MOODS.map((m) => <Chip key={m} on={moods.includes(m)} onClick={() => toggle(moods, m, setMoods)}>{m}</Chip>)}</div>
        <textarea value={note} onChange={(e) => setNote(e.target.value)} rows={2} placeholder="Що ще? (необов'язково)" className="mb-3 w-full rounded-xl border border-slate-300 px-3 py-2 text-sm focus:border-teal-400 focus:outline-none" />
        <button onClick={() => onSave({ substance, type, time, stress, company, moods, note: note.trim() })} className="w-full rounded-2xl bg-teal-500 py-3 font-bold text-white hover:bg-teal-600">{type === "slip" ? "Записати зрив (стрік почнеться заново)" : "Записати"}</button>
      </div>
    </div>
  );
}

/* ---------- Routine Pomodoro + 2-minute starter (ADHD low-barrier start) ---------- */
function RoutinePomodoro({ task, mode, settings, onClose, onLog, onSaveSettings }) {
  const [stage, setStage] = useState(mode === "pomodoro" ? "setup" : "work"); // setup | work | break | done
  const [workMin, setWorkMin] = useState(settings.work || 25);
  const [breakMin, setBreakMin] = useState(settings.break || 5);
  const [left, setLeft] = useState(120);
  const [cycles, setCycles] = useState(0);
  const worked = useRef(0);
  const iv = useRef(null);
  const p = getPastel(task.color);
  const fmt = (s) => `${Math.floor(s / 60)}:${String(Math.max(0, s % 60)).padStart(2, "0")}`;

  const run = (secs, isBreak) => {
    clearInterval(iv.current);
    setLeft(secs); setStage(isBreak ? "break" : "work");
    iv.current = setInterval(() => setLeft((l) => {
      if (!isBreak) worked.current += 1;
      if (l <= 1) { clearInterval(iv.current); phaseEnd(isBreak); return 0; }
      return l - 1;
    }), 1000);
  };
  const phaseEnd = (wasBreak) => {
    if (mode === "2min") { setStage("done"); return; }
    if (!wasBreak) { setCycles((c) => c + 1); run(breakMin * 60, true); }
    else run(workMin * 60, false);
  };
  useEffect(() => { if (mode === "2min") run(120, false); return () => clearInterval(iv.current); /* eslint-disable-next-line */ }, []);
  const finish = () => { clearInterval(iv.current); if (worked.current >= 5) onLog(task.id, Math.round(worked.current)); onClose(); };
  const startPomodoro = () => { onSaveSettings({ work: workMin, break: breakMin }); run(workMin * 60, false); };

  if (mode === "pomodoro" && stage === "setup") {
    return (
      <div className="fixed inset-0 z-[55] flex items-center justify-center bg-slate-900/50 p-4 backdrop-blur-sm" onClick={onClose}>
        <div className="w-full max-w-sm rounded-3xl bg-white p-6 text-center shadow-2xl" onClick={(e) => e.stopPropagation()}>
          <div className="text-4xl">🍅</div>
          <h3 className="mt-2 text-lg font-extrabold text-slate-900">Помодоро</h3>
          <p className="mt-0.5 truncate text-sm text-slate-500">{task.title}</p>
          <div className="my-4 flex justify-center gap-4">
            <label className="text-sm"><div className="mb-1 text-xs text-slate-400">Робота</div><input type="number" min={1} max={90} value={workMin} onChange={(e) => setWorkMin(Math.max(1, Math.min(90, +e.target.value || 1)))} className="w-20 rounded-lg border border-slate-300 px-2 py-1.5 text-center text-sm" /> <span className="text-xs text-slate-400">хв</span></label>
            <label className="text-sm"><div className="mb-1 text-xs text-slate-400">Перерва</div><input type="number" min={1} max={30} value={breakMin} onChange={(e) => setBreakMin(Math.max(1, Math.min(30, +e.target.value || 1)))} className="w-20 rounded-lg border border-slate-300 px-2 py-1.5 text-center text-sm" /> <span className="text-xs text-slate-400">хв</span></label>
          </div>
          <button onClick={startPomodoro} className="w-full rounded-2xl bg-pink-500 py-3 font-bold text-white shadow-lg shadow-pink-500/20 hover:bg-pink-600">Почати</button>
          <button onClick={onClose} className="mt-2 w-full py-2 text-sm font-semibold text-slate-400">Скасувати</button>
        </div>
      </div>
    );
  }

  const isBreak = stage === "break";
  const done2 = mode === "2min" && stage === "done";
  return (
    <div className="fixed inset-0 z-[55] flex flex-col bg-gradient-to-b from-red-50 to-pink-100 p-6">
      <div className="flex items-center justify-between">
        <span className="text-sm font-bold text-slate-400">{mode === "2min" ? "Старт на 2 хвилини" : `Помодоро · коло ${cycles + 1}`}</span>
        <button onClick={finish} className="rounded-full bg-white/70 p-2 text-slate-500"><X className="h-5 w-5" /></button>
      </div>
      <div className="flex flex-1 flex-col items-center justify-center text-center">
        <div className="grid h-20 w-20 place-items-center rounded-3xl text-4xl shadow-lg" style={{ backgroundColor: p.card }}>{task.emoji || "⭐"}</div>
        <h1 className="mt-4 max-w-md text-2xl font-extrabold text-slate-900">{task.title}</h1>
        {done2 ? (
          <>
            <div className="mt-4 text-5xl">🎉</div>
            <p className="mt-2 max-w-xs text-lg font-bold text-slate-800">Ти почала — це найважче!</p>
            <p className="mt-1 max-w-xs text-sm text-slate-500">Початок зроблено. Хочеш проїхати ще трохи на цій хвилі?</p>
            <div className="mt-5 flex gap-2">
              <button onClick={() => run(300, false)} className="rounded-2xl bg-pink-500 px-5 py-3 font-bold text-white shadow-lg shadow-pink-500/20">Ще 5 хвилин</button>
              <button onClick={finish} className="rounded-2xl bg-white px-5 py-3 font-bold text-slate-500 ring-1 ring-slate-200">Досить на зараз</button>
            </div>
          </>
        ) : (
          <>
            {mode === "2min" && task.twoMin && <p className="mt-2 max-w-sm rounded-2xl bg-white/70 px-4 py-2 text-sm font-semibold text-slate-700">Тільки це: {task.twoMin}</p>}
            {mode === "2min" && !task.twoMin && <p className="mt-2 max-w-sm text-sm text-slate-500">Просто почни. Дозволено зробити абияк — головне рушити.</p>}
            <div className={`mt-6 grid h-56 w-56 place-items-center rounded-full text-white ${isBreak ? "bg-gradient-to-br from-sky-300 to-teal-300" : "bg-gradient-to-br from-pink-400 to-fuchsia-400"}`}>
              <div className="text-center"><div className="text-xs font-semibold uppercase tracking-widest text-white/80">{isBreak ? "Перерва" : "Робота"}</div><div className="text-6xl font-black tabular-nums">{fmt(left)}</div></div>
            </div>
            <button onClick={finish} className="mt-6 inline-flex items-center gap-2 rounded-full bg-white px-6 py-3 font-bold text-slate-600 shadow-sm ring-1 ring-slate-200"><Check className="h-4 w-4" /> Завершити{worked.current >= 60 ? ` (${Math.round(worked.current / 60)} хв)` : ""}</button>
          </>
        )}
      </div>
    </div>
  );
}

/* ---------- Block 6: Wellbeing — mood over time + gratitude + behavioral activation ---------- */
const ACTIVATION_SEED = ["одна улюблена пісня", "вийти на вулицю на 5 хв", "написати другу пару слів", "склянка води", "розтягнутися 2 хв", "відкрити вікно й подихати"];
function WellbeingView({ onExit, moods, onMood }) {
  const today = dateKey(Date.now());
  const [gratitude, setGratitude] = useState({});
  const [activation, setActivation] = useState([]);
  const [notes, setNotes] = useState({});
  const [gInputs, setGInputs] = useState(["", "", ""]);
  const [newAct, setNewAct] = useState("");
  const [picked, setPicked] = useState(null);
  const [range, setRange] = useState(30);

  useEffect(() => { (async () => {
    setGratitude(await store.get(RKEYS.gratitude, {}));
    let a = await store.get(RKEYS.activation, null);
    if (!a) { a = ACTIVATION_SEED.map((t) => ({ id: ruid("ba"), text: t })); await store.set(RKEYS.activation, a); }
    setActivation(a);
    setNotes(await store.get(RKEYS.moodNotes, {}));
  })(); }, []);
  useEffect(() => { setGInputs(((gratitude[today]) || ["", "", ""]).concat(["", "", ""]).slice(0, 3)); }, [gratitude, today]);

  const saveGratitude = () => { const items = gInputs.map((s) => s.trim()).filter(Boolean); const next = { ...gratitude, [today]: items }; if (!items.length) delete next[today]; setGratitude(next); store.set(RKEYS.gratitude, next); };
  const saveActivation = (n) => { setActivation(n); store.set(RKEYS.activation, n); };
  const saveNote = (patch) => { const next = { ...notes, [today]: { ...(notes[today] || {}), ...patch } }; setNotes(next); store.set(RKEYS.moodNotes, next); };
  const pickAction = () => { if (!activation.length) return; setPicked(activation[Math.floor(((Date.now() / 1000) % activation.length))] || activation[0]); };

  const chartData = useMemo(() => {
    const out = []; const now = new Date();
    for (let i = range - 1; i >= 0; i--) { const d = new Date(now.getFullYear(), now.getMonth(), now.getDate() - i); const ds = dateKey(d.getTime()); out.push({ day: `${d.getDate()}.${d.getMonth() + 1}`, score: moods[ds] ?? null }); }
    return out;
  }, [moods, range]);
  const rated = chartData.filter((d) => d.score != null);
  const avg = rated.length ? (rated.reduce((s, d) => s + d.score, 0) / rated.length) : null;
  const FACTORS = ["сон", "робота", "стрес", "самотність", "рух", "їжа", "погода", "люди", "здоров'я"];

  return (
    <div className="mx-auto w-full max-w-2xl px-4 pb-24 pt-6">
      <div className="mb-4 flex items-center gap-3">
        <button onClick={onExit} className="grid h-9 w-9 place-items-center rounded-full bg-white text-slate-500 shadow-sm ring-1 ring-red-100"><ArrowLeft className="h-4 w-4" /></button>
        <h1 className="text-xl font-extrabold text-slate-900">Настрій і вдячність</h1>
      </div>

      {/* mood over time */}
      <div className="rounded-2xl bg-white p-4 shadow-sm ring-1 ring-red-100">
        <div className="mb-2 flex items-center justify-between">
          <span className="text-sm font-bold text-slate-700">Настрій у часі{avg != null ? ` · середнє ${avg.toFixed(1)}` : ""}</span>
          <div className="flex gap-1 rounded-full bg-slate-100 p-0.5">{[14, 30, 90].map((r) => <button key={r} onClick={() => setRange(r)} className={`rounded-full px-2 py-0.5 text-xs font-bold ${range === r ? "bg-pink-500 text-white" : "text-slate-500"}`}>{r}д</button>)}</div>
        </div>
        {rated.length < 2 ? <p className="py-6 text-center text-sm text-slate-400">Відмічай настрій щодня — тут з'явиться графік, і побачиш, що впливає на кращі й гірші дні.</p> : (
          <div style={{ height: 180 }}>
            <ResponsiveContainer width="100%" height="100%">
              <LineChart data={chartData} margin={{ left: -20, right: 8, top: 4, bottom: 0 }}>
                <CartesianGrid strokeDasharray="3 3" stroke="#f1f5f9" />
                <XAxis dataKey="day" tick={{ fontSize: 10, fill: "#94a3b8" }} interval="preserveStartEnd" minTickGap={24} />
                <YAxis domain={[1, 5]} ticks={[1, 2, 3, 4, 5]} tick={{ fontSize: 10, fill: "#94a3b8" }} />
                <Tooltip formatter={(v) => MOODS.find((m) => m.score === v)?.label || v} />
                <Line type="monotone" dataKey="score" stroke="#ec4899" strokeWidth={2.5} dot={{ r: 3, fill: "#ec4899" }} connectNulls />
              </LineChart>
            </ResponsiveContainer>
          </div>
        )}
      </div>

      {/* today mood + tag */}
      <div className="mt-3 rounded-2xl bg-white p-4 shadow-sm ring-1 ring-red-100">
        <div className="mb-2 text-sm font-bold text-slate-700">Сьогодні</div>
        <div className="flex justify-between">{MOODS.map((m) => <button key={m.score} onClick={() => onMood(m.score)} className={`flex flex-col items-center gap-0.5 rounded-2xl px-2 py-1.5 transition hover:scale-110 ${moods[today] === m.score ? "bg-pink-50 ring-2 ring-pink-300" : ""}`}><span className="text-2xl">{m.emoji}</span><span className="text-[9px] text-slate-400">{m.label}</span></button>)}</div>
        <div className="mt-3 text-xs font-semibold text-slate-500">Що вплинуло?</div>
        <div className="mt-1.5 flex flex-wrap gap-1.5">{FACTORS.map((f) => { const on = (notes[today]?.factors || []).includes(f); return <button key={f} onClick={() => { const cur = notes[today]?.factors || []; saveNote({ factors: on ? cur.filter((x) => x !== f) : [...cur, f] }); }} className={`rounded-full px-2.5 py-1 text-xs font-semibold ring-1 ${on ? "bg-pink-500 text-white ring-pink-500" : "bg-white text-slate-500 ring-slate-200"}`}>{f}</button>; })}</div>
        <input value={notes[today]?.note || ""} onChange={(e) => saveNote({ note: e.target.value })} placeholder="Нотатка про день (необов'язково)" className="mt-2 w-full rounded-lg border border-slate-200 px-3 py-1.5 text-sm focus:border-pink-400 focus:outline-none" />
      </div>

      {/* gratitude */}
      <div className="mt-3 rounded-2xl bg-white p-4 shadow-sm ring-1 ring-red-100">
        <div className="mb-2 text-sm font-bold text-slate-700">🙏 3 речі, за які вдячна сьогодні</div>
        <div className="space-y-2">{[0, 1, 2].map((i) => <input key={i} value={gInputs[i] || ""} onChange={(e) => setGInputs((a) => { const n = [...a]; n[i] = e.target.value; return n; })} onBlur={saveGratitude} placeholder={`${i + 1}…`} className="w-full rounded-lg border border-slate-200 px-3 py-1.5 text-sm focus:border-pink-400 focus:outline-none" />)}</div>
        {Object.keys(gratitude).filter((d) => d !== today).length > 0 && <div className="mt-2 text-[11px] text-slate-400">записів вдячності: {Object.keys(gratitude).length}</div>}
      </div>

      {/* behavioral activation */}
      <div className="mt-3 rounded-2xl bg-white p-4 shadow-sm ring-1 ring-red-100">
        <div className="mb-1 text-sm font-bold text-slate-700">Маленькі дії на день без сил</div>
        <p className="mb-2 text-xs text-slate-400">Коли енергії нема, не обирай — хай застосунок підкаже одну маленьку дію.</p>
        <button onClick={pickAction} className="w-full rounded-2xl bg-gradient-to-r from-pink-400 to-fuchsia-400 py-3 font-bold text-white shadow-sm">🎲 Мало енергії — обери за мене</button>
        {picked && <div className="mt-2 rounded-2xl bg-pink-50 px-4 py-3 text-center"><div className="text-[11px] font-semibold uppercase tracking-wide text-pink-500">спробуй це</div><div className="mt-0.5 text-lg font-bold text-slate-800">{picked.text}</div></div>}
        <div className="mt-3 space-y-1.5">
          {activation.map((a) => (
            <div key={a.id} className="group flex items-center gap-2 rounded-lg bg-slate-50 px-3 py-1.5 text-sm text-slate-600"><span className="flex-1">{a.text}</span><button onClick={() => saveActivation(activation.filter((x) => x.id !== a.id))} className="text-slate-300 hover:text-red-500 sm:opacity-0 sm:group-hover:opacity-100"><X className="h-3.5 w-3.5" /></button></div>
          ))}
        </div>
        <div className="mt-2 flex gap-2"><input value={newAct} onChange={(e) => setNewAct(e.target.value)} placeholder="Додати свою дію…" onKeyDown={(e) => { if (e.key === "Enter" && newAct.trim()) { saveActivation([...activation, { id: ruid("ba"), text: newAct.trim() }]); setNewAct(""); } }} className="min-w-0 flex-1 rounded-lg border border-slate-200 px-3 py-1.5 text-sm focus:border-pink-400 focus:outline-none" /><button onClick={() => { if (newAct.trim()) { saveActivation([...activation, { id: ruid("ba"), text: newAct.trim() }]); setNewAct(""); } }} className="rounded-lg bg-slate-800 px-3 text-sm font-semibold text-white">+</button></div>
      </div>
      <p className="mt-4 text-center text-xs leading-relaxed text-slate-400">Якщо пригнічений настрій тримається тижнями — це вагома причина поговорити з фахівцем. Ти не маєш давати собі раду наодинці. 💛</p>
    </div>
  );
}

/* ---------- Block 5: Movement — counter + hourly nudge (on Today) ---------- */
const STRETCHES = ["встань і потягнись до стелі", "10 присідань", "пройдись до вікна й назад", "покрути плечима й шиєю", "налий води й випий стоячи", "походи хвилину на місці"];
function MovementCard({ flash }) {
  const today = dateKey(Date.now());
  const [count, setCount] = useState(0);
  const [cfg, setCfg] = useState({ remindersOn: true, snoozeUntil: 0, lastNudge: 0 });
  const [nudge, setNudge] = useState(false);
  const [tip, setTip] = useState(STRETCHES[0]);

  useEffect(() => { (async () => {
    const mv = await store.get(RKEYS.movement, {}); setCount(mv[today] || 0);
    setCfg(await store.get(RKEYS.movementCfg, { remindersOn: true, snoozeUntil: 0, lastNudge: 0 }));
  })(); }, [today]);
  const saveCount = (n) => { setCount(n); store.get(RKEYS.movement, {}).then((mv) => store.set(RKEYS.movement, { ...mv, [today]: n })); };
  const saveCfg = (c) => { setCfg(c); store.set(RKEYS.movementCfg, c); };

  useEffect(() => {
    const check = () => {
      if (!cfg.remindersOn) { setNudge(false); return; }
      const now = Date.now(); const h = new Date().getHours();
      if (h < 9 || h >= 18) { setNudge(false); return; }
      if (now < (cfg.snoozeUntil || 0)) { setNudge(false); return; }
      if (now - (cfg.lastNudge || 0) > 55 * 60000) setNudge(true);
    };
    check(); const iv = setInterval(check, 60000); return () => clearInterval(iv);
  }, [cfg]);

  const didMove = () => { saveCount(count + 1); saveCfg({ ...cfg, lastNudge: Date.now() }); setNudge(false); flash && flash("Красуня — тіло дякує 💪"); };
  const snooze = () => { saveCfg({ ...cfg, snoozeUntil: Date.now() + 15 * 60000 }); setNudge(false); };
  const turnOff = () => { saveCfg({ ...cfg, remindersOn: false }); setNudge(false); };

  return (
    <div className="mt-4">
      {nudge && (
        <div className="mb-3 flex items-center gap-3 rounded-2xl bg-gradient-to-r from-lime-400 to-emerald-400 p-3 text-white shadow-sm">
          <span className="text-2xl">🧍</span>
          <div className="flex-1"><div className="text-sm font-bold">Час встати й порухатись</div><div className="text-xs text-white/90">{tip}</div></div>
          <div className="flex flex-col gap-1">
            <button onClick={didMove} className="rounded-full bg-white/25 px-3 py-1 text-xs font-bold">Порухалась</button>
            <div className="flex gap-1"><button onClick={snooze} className="rounded-full bg-white/15 px-2 py-0.5 text-[10px]">+15хв</button><button onClick={turnOff} className="rounded-full bg-white/15 px-2 py-0.5 text-[10px]">вимк</button></div>
          </div>
        </div>
      )}
      <div className="flex items-center gap-3 rounded-2xl bg-white/80 p-3 shadow-sm ring-1 ring-red-100">
        <span className="grid h-10 w-10 shrink-0 place-items-center rounded-xl bg-lime-100 text-xl">🚶</span>
        <div className="flex-1"><div className="text-sm font-bold text-slate-700">Рух сьогодні: {count}</div><div className="text-[11px] text-slate-400">проти сидіння 9–18 · нагадування {cfg.remindersOn ? "увімкнені" : "вимкнені"}</div></div>
        <button onClick={() => saveCount(count + 1)} className="rounded-full bg-lime-500 px-3 py-1.5 text-xs font-bold text-white hover:bg-lime-600">+ Порухалась</button>
        <button onClick={() => saveCfg({ ...cfg, remindersOn: !cfg.remindersOn })} title="Нагадування" className={`grid h-8 w-8 place-items-center rounded-full ${cfg.remindersOn ? "bg-lime-100 text-lime-600" : "bg-slate-100 text-slate-400"}`}>{cfg.remindersOn ? <Clock className="h-4 w-4" /> : <Clock className="h-4 w-4 opacity-50" />}</button>
      </div>
    </div>
  );
}

/* ---------- Block 4: Meds — refill + wellbeing/side-effect log + doctor taper ---------- */
const MED_SIDE_FX = ["сонливість", "нудота", "головний біль", "безсоння", "апетит", "тривога", "сухість у роті", "запаморочення"];
function MedsView({ onExit }) {
  const today = dateKey(Date.now());
  const [meds, setMeds] = useState([]);
  const [log, setLog] = useState({});
  const [editor, setEditor] = useState(null);
  const [taperFor, setTaperFor] = useState(null);

  useEffect(() => { (async () => { setMeds(await store.get(RKEYS.meds, [])); setLog(await store.get(RKEYS.medsLog, {})); })(); }, []);
  const saveMeds = (n) => { setMeds(n); store.set(RKEYS.meds, n); };
  const saveLog = (n) => { setLog(n); store.set(RKEYS.medsLog, n); };
  const todayLog = log[today] || {};
  const setTaken = (medId, v) => { const d = { ...todayLog, taken: { ...(todayLog.taken || {}), [medId]: v } }; saveLog({ ...log, [today]: d }); };
  const setWell = (patch) => saveLog({ ...log, [today]: { ...todayLog, ...patch } });

  const upsertMed = (m, id) => { if (id) saveMeds(meds.map((x) => (x.id === id ? { ...x, ...m } : x))); else saveMeds([...meds, { id: ruid("med"), taper: [], ...m }]); };
  const delMed = (id) => saveMeds(meds.filter((x) => x.id !== id));
  const refill = (id, add) => saveMeds(meds.map((x) => (x.id === id ? { ...x, supply: (Number(x.supply) || 0) + add } : x)));

  const wkData = useMemo(() => { const out = []; const now = new Date(); for (let i = 13; i >= 0; i--) { const d = new Date(now.getFullYear(), now.getMonth(), now.getDate() - i); const ds = dateKey(d.getTime()); out.push({ day: `${d.getDate()}.${d.getMonth() + 1}`, w: log[ds]?.wellbeing ?? null }); } return out; }, [log]);
  const wkRated = wkData.filter((d) => d.w != null);

  return (
    <div className="mx-auto w-full max-w-2xl px-4 pb-24 pt-6">
      <div className="mb-3 flex items-center gap-3">
        <button onClick={onExit} className="grid h-9 w-9 place-items-center rounded-full bg-white text-slate-500 shadow-sm ring-1 ring-red-100"><ArrowLeft className="h-4 w-4" /></button>
        <h1 className="text-xl font-extrabold text-slate-900">Ліки й самопочуття</h1>
      </div>
      <p className="mb-3 rounded-2xl bg-amber-50 px-3 py-2 text-xs leading-relaxed text-amber-800">Зміни дози — лише разом із лікарем. Застосунок нічого не призначає й не радить схем: він лише веде облік того, що призначив лікар, і як ти почуваєшся.</p>

      {/* meds list */}
      <div className="space-y-2">
        {meds.length === 0 ? <div className="rounded-2xl bg-white py-8 text-center text-sm text-slate-400 ring-1 ring-red-50">Додай ліки, щоб бачити запас і вести журнал самопочуття.</div> : meds.map((m) => {
          const daysLeft = m.perDay > 0 ? Math.floor((Number(m.supply) || 0) / m.perDay) : null;
          const low = daysLeft != null && daysLeft <= 5;
          return (
            <div key={m.id} className={`rounded-2xl bg-white p-4 shadow-sm ring-1 ${low ? "ring-2 ring-amber-300" : "ring-red-50"}`}>
              <div className="flex items-start justify-between gap-2">
                <div className="min-w-0"><div className="font-bold text-slate-800">{m.name}</div><div className="text-xs text-slate-400">{m.dose}{m.perDay ? ` · ${m.perDay}×/день` : ""}</div></div>
                <button onClick={() => setTaken(m.id, !todayLog.taken?.[m.id])} className={`shrink-0 rounded-full px-3 py-1.5 text-xs font-bold ${todayLog.taken?.[m.id] ? "bg-green-500 text-white" : "bg-slate-100 text-slate-500"}`}>{todayLog.taken?.[m.id] ? "✓ прийнято" : "прийняти"}</button>
              </div>
              {daysLeft != null && (
                <div className={`mt-2 flex items-center justify-between rounded-xl px-3 py-1.5 text-xs ${low ? "bg-amber-50 text-amber-700" : "bg-slate-50 text-slate-500"}`}>
                  <span>{low ? "⚠️ " : ""}запас ~{daysLeft} дн ({m.supply} шт)</span>
                  <span className="flex gap-1"><button onClick={() => refill(m.id, 30)} className="rounded-full bg-white px-2 py-0.5 font-semibold ring-1 ring-slate-200">+30</button><button onClick={() => { const v = prompt("Скільки шт зараз у запасі?", m.supply); if (v != null) upsertMed({ supply: Math.max(0, +v || 0) }, m.id); }} className="rounded-full bg-white px-2 py-0.5 font-semibold ring-1 ring-slate-200">задати</button></span>
                </div>
              )}
              {m.taper?.length > 0 && (
                <div className="mt-2 rounded-xl bg-sky-50 px-3 py-2">
                  <div className="text-[11px] font-bold text-sky-700">Схема зниження (за призначенням лікаря)</div>
                  <div className="mt-1 space-y-0.5">{m.taper.map((t, i) => { const active = t.date <= today; return <div key={i} className={`flex justify-between text-xs ${active ? "text-sky-800 font-semibold" : "text-slate-400"}`}><span>{t.date}</span><span>{t.dose}{t.note ? ` · ${t.note}` : ""}</span></div>; })}</div>
                </div>
              )}
              <div className="mt-2 flex gap-2">
                <button onClick={() => setEditor({ med: m })} className="rounded-full bg-white px-3 py-1 text-xs font-semibold text-slate-500 ring-1 ring-slate-200">Редагувати</button>
                <button onClick={() => setTaperFor(m)} className="rounded-full bg-white px-3 py-1 text-xs font-semibold text-sky-600 ring-1 ring-sky-200">Схема лікаря</button>
                <button onClick={() => { if (confirm("Видалити?")) delMed(m.id); }} className="ml-auto text-slate-300 hover:text-red-500"><Trash2 className="h-4 w-4" /></button>
              </div>
            </div>
          );
        })}
        <button onClick={() => setEditor({ med: null })} className="flex w-full items-center justify-center gap-1.5 rounded-2xl border border-dashed border-red-300 py-3 text-sm font-semibold text-red-500 hover:bg-red-50"><Plus className="h-4 w-4" /> Додати ліки</button>
      </div>

      {/* wellbeing log */}
      <div className="mt-4 rounded-2xl bg-white p-4 shadow-sm ring-1 ring-red-100">
        <div className="mb-2 text-sm font-bold text-slate-700">Самопочуття сьогодні</div>
        <div className="flex justify-between">{MOODS.map((m) => <button key={m.score} onClick={() => setWell({ wellbeing: m.score })} className={`flex flex-col items-center gap-0.5 rounded-2xl px-2 py-1.5 hover:scale-110 ${todayLog.wellbeing === m.score ? "bg-pink-50 ring-2 ring-pink-300" : ""}`}><span className="text-2xl">{m.emoji}</span></button>)}</div>
        <div className="mt-2 text-xs font-semibold text-slate-500">Побічні ефекти</div>
        <div className="mt-1.5 flex flex-wrap gap-1.5">{MED_SIDE_FX.map((s) => { const on = (todayLog.sideEffects || []).includes(s); return <button key={s} onClick={() => { const cur = todayLog.sideEffects || []; setWell({ sideEffects: on ? cur.filter((x) => x !== s) : [...cur, s] }); }} className={`rounded-full px-2.5 py-1 text-xs font-semibold ring-1 ${on ? "bg-amber-500 text-white ring-amber-500" : "bg-white text-slate-500 ring-slate-200"}`}>{s}</button>; })}</div>
        <input value={todayLog.note || ""} onChange={(e) => setWell({ note: e.target.value })} placeholder="Нотатка для лікаря (необов'язково)" className="mt-2 w-full rounded-lg border border-slate-200 px-3 py-1.5 text-sm focus:border-pink-400 focus:outline-none" />
        {wkRated.length >= 2 && (
          <div className="mt-3" style={{ height: 120 }}>
            <div className="mb-1 text-[11px] font-semibold text-slate-400">Самопочуття за 2 тижні (покажи лікарю)</div>
            <ResponsiveContainer width="100%" height="100%">
              <LineChart data={wkData} margin={{ left: -28, right: 8, top: 4, bottom: 0 }}>
                <XAxis dataKey="day" tick={{ fontSize: 9, fill: "#94a3b8" }} interval="preserveStartEnd" /><YAxis domain={[1, 5]} ticks={[1, 3, 5]} tick={{ fontSize: 9, fill: "#94a3b8" }} />
                <Tooltip /><Line type="monotone" dataKey="w" stroke="#8b5cf6" strokeWidth={2} dot={{ r: 2 }} connectNulls />
              </LineChart>
            </ResponsiveContainer>
          </div>
        )}
      </div>

      {editor && <MedEditor med={editor.med} onClose={() => setEditor(null)} onSave={(m) => { upsertMed(m, editor.med?.id); setEditor(null); }} />}
      {taperFor && <TaperEditor med={taperFor} onClose={() => setTaperFor(null)} onSave={(taper) => { upsertMed({ taper }, taperFor.id); setTaperFor(null); }} />}
    </div>
  );
}

function MedEditor({ med, onClose, onSave }) {
  const [name, setName] = useState(med?.name || "");
  const [dose, setDose] = useState(med?.dose || "");
  const [perDay, setPerDay] = useState(med?.perDay ?? 1);
  const [supply, setSupply] = useState(med?.supply ?? "");
  return (
    <div className="fixed inset-0 z-[55] flex items-end justify-center bg-slate-900/40 p-0 backdrop-blur-sm sm:items-center sm:p-4" onClick={onClose}>
      <div className="w-full max-w-md rounded-t-3xl bg-white p-5 shadow-xl sm:rounded-3xl" onClick={(e) => e.stopPropagation()}>
        <div className="mb-3 flex items-center justify-between"><h3 className="text-lg font-bold text-slate-900">{med ? "Редагувати" : "Нові ліки"}</h3><button onClick={onClose} className="rounded-md p-1 text-slate-400"><X className="h-5 w-5" /></button></div>
        <input value={name} onChange={(e) => setName(e.target.value)} placeholder="Назва (напр. Золофт)" className="mb-2 w-full rounded-xl border border-slate-300 px-3 py-2 text-sm font-semibold focus:border-pink-400 focus:outline-none" />
        <input value={dose} onChange={(e) => setDose(e.target.value)} placeholder="Доза (напр. 50 мг)" className="mb-2 w-full rounded-xl border border-slate-300 px-3 py-2 text-sm focus:border-pink-400 focus:outline-none" />
        <div className="mb-3 grid grid-cols-2 gap-2">
          <label className="block"><span className="mb-1 block text-xs text-slate-500">Разів на день</span><input type="number" min={0} value={perDay} onChange={(e) => setPerDay(Math.max(0, +e.target.value || 0))} className="w-full rounded-lg border border-slate-300 px-2 py-1.5 text-sm" /></label>
          <label className="block"><span className="mb-1 block text-xs text-slate-500">Запас (шт)</span><input type="number" min={0} value={supply} onChange={(e) => setSupply(e.target.value)} className="w-full rounded-lg border border-slate-300 px-2 py-1.5 text-sm" /></label>
        </div>
        <button onClick={() => { if (name.trim()) onSave({ name: name.trim(), dose: dose.trim(), perDay: Number(perDay) || 0, supply: Number(supply) || 0 }); }} className="w-full rounded-2xl bg-pink-500 py-3 font-bold text-white hover:bg-pink-600">Зберегти</button>
      </div>
    </div>
  );
}

function TaperEditor({ med, onClose, onSave }) {
  const [rows, setRows] = useState(med.taper?.length ? med.taper.map((t) => ({ ...t })) : [{ date: dateKey(Date.now()), dose: "", note: "" }]);
  const upd = (i, k, v) => setRows((r) => r.map((x, j) => (j === i ? { ...x, [k]: v } : x)));
  return (
    <div className="fixed inset-0 z-[55] flex items-end justify-center bg-slate-900/40 p-0 backdrop-blur-sm sm:items-center sm:p-4" onClick={onClose}>
      <div className="max-h-[90vh] w-full max-w-md overflow-y-auto rounded-t-3xl bg-white p-5 shadow-xl sm:rounded-3xl" onClick={(e) => e.stopPropagation()}>
        <div className="mb-1 flex items-center justify-between"><h3 className="text-lg font-bold text-slate-900">Схема зниження</h3><button onClick={onClose} className="rounded-md p-1 text-slate-400"><X className="h-5 w-5" /></button></div>
        <p className="mb-3 rounded-xl bg-sky-50 px-3 py-2 text-xs text-sky-800">Введи саме те, що призначив <b>лікар</b>. Застосунок нічого не вигадує — лише показує й нагадує. Ніколи не змінюй дозу самостійно.</p>
        <div className="space-y-2">{rows.map((r, i) => (
          <div key={i} className="flex gap-2">
            <input type="date" value={r.date} onChange={(e) => upd(i, "date", e.target.value)} className="rounded-lg border border-slate-300 px-2 py-1.5 text-xs" />
            <input value={r.dose} onChange={(e) => upd(i, "dose", e.target.value)} placeholder="доза" className="w-20 rounded-lg border border-slate-300 px-2 py-1.5 text-sm" />
            <input value={r.note} onChange={(e) => upd(i, "note", e.target.value)} placeholder="нотатка" className="min-w-0 flex-1 rounded-lg border border-slate-300 px-2 py-1.5 text-sm" />
            <button onClick={() => setRows((rr) => rr.filter((_, j) => j !== i))} className="text-slate-300 hover:text-red-500"><X className="h-4 w-4" /></button>
          </div>
        ))}</div>
        <button onClick={() => setRows((r) => [...r, { date: dateKey(Date.now()), dose: "", note: "" }])} className="mt-2 text-sm font-semibold text-sky-600">+ Рядок</button>
        <button onClick={() => onSave(rows.filter((r) => r.date && r.dose))} className="mt-3 w-full rounded-2xl bg-sky-500 py-3 font-bold text-white hover:bg-sky-600">Зберегти схему</button>
      </div>
    </div>
  );
}

/* ---------- Block 7: Career growth (inside Management) ---------- */
const CAREERKEYS = { skills: "career:skills", achievements: "career:achievements", reviews: "career:reviews", path: "career:pathProgress", bookPath: "career:bookProgress" };
async function collectCareerExport() { return { skills: await store.get(CAREERKEYS.skills, []), achievements: await store.get(CAREERKEYS.achievements, []), reviews: await store.get(CAREERKEYS.reviews, []), path: await store.get(CAREERKEYS.path, {}), bookPath: await store.get(CAREERKEYS.bookPath, {}) }; }
async function clearCareerData() { for (const k of Object.values(CAREERKEYS)) await store.remove(k); await store.remove("career:seeded"); }

function CareerView() {
  const today = dateKey(Date.now());
  const week = finWeekStart(today);
  const [tab, setTab] = useState("path"); // path | skills | wins | review
  const [skills, setSkills] = useState([]);
  const [wins, setWins] = useState([]);
  const [reviews, setReviews] = useState([]);
  const [pathProg, setPathProg] = useState({});
  const [bookProg, setBookProg] = useState({});
  const [pathTab, setPathTab] = useState("roadmap"); // roadmap | book
  const [skillEd, setSkillEd] = useState(null);
  const [winText, setWinText] = useState("");
  const [reviewText, setReviewText] = useState("");

  useEffect(() => { (async () => {
    let sk = await store.get(CAREERKEYS.skills, []);
    if (!sk.length && !(await store.get("career:seeded", false))) {
      const D = (days) => dateKey(Date.now() + days * 86400000);
      sk = [
        { id: ruid("sk"), name: "Продуктові основи PM", target: "Roadmap, PRD, discovery, пріоритизація (RICE/MoSCoW), user stories", deadline: D(42), progress: 0 },
        { id: ruid("sk"), name: "Технічна глибина", target: "System design, API/REST, бази даних, як влаштований веб — щоб говорити з інженерами", deadline: D(90), progress: 0 },
        { id: ruid("sk"), name: "AI-грамотність", target: "LLM і токени, промпт-інжиніринг, embeddings, RAG, оцінка якості (evals), обмеження й безпека", deadline: D(120), progress: 0 },
        { id: ruid("sk"), name: "Будувати з AI", target: "OpenAI/Anthropic API, зібрати й показати одну AI-фічу (no-code + трохи коду)", deadline: D(150), progress: 0 },
        { id: ruid("sk"), name: "Аналітика й метрики", target: "SQL, north-star метрика, воронки, A/B-експерименти, читати дашборди", deadline: D(180), progress: 0 },
        { id: ruid("sk"), name: "Стейкхолдери й комунікація", target: "Презентувати roadmap, вирівнювати founder/eng/design, писати чіткі специфікації", deadline: "", progress: 0 },
        { id: ruid("sk"), name: "Портфоліо AI-PM", target: "2 pet-проєкти з AI + оформлені кейси (проблема → рішення → метрика)", deadline: D(270), progress: 0 },
        { id: ruid("sk"), name: "Підготовка до співбесід", target: "Product sense, system design для PM, AI-кейси, метрики, поведінкові — і подавати заявки", deadline: D(300), progress: 0 },
      ];
      await store.set(CAREERKEYS.skills, sk); await store.set("career:seeded", true);
    }
    setSkills(sk);
    setWins(await store.get(CAREERKEYS.achievements, []));
    setPathProg(await store.get(CAREERKEYS.path, {}));
    setBookProg(await store.get(CAREERKEYS.bookPath, {}));
    const rv = await store.get(CAREERKEYS.reviews, []); setReviews(rv);
    setReviewText((rv.find((r) => r.week === week) || {}).text || "");
  })(); }, [week]);
  const setStepDone = (stepId, done) => setPathProg((prev) => { const n = { ...prev }; if (done) n[stepId] = true; else delete n[stepId]; store.set(CAREERKEYS.path, n); return n; });
  const setBookStepDone = (stepId, done) => setBookProg((prev) => { const n = { ...prev }; if (done) n[stepId] = true; else delete n[stepId]; store.set(CAREERKEYS.bookPath, n); return n; });
  const saveSkills = (n) => { setSkills(n); store.set(CAREERKEYS.skills, n); };
  const saveWins = (n) => { setWins(n); store.set(CAREERKEYS.achievements, n); };
  const saveReviews = (n) => { setReviews(n); store.set(CAREERKEYS.reviews, n); };

  const upsertSkill = (m, id) => { if (id) saveSkills(skills.map((x) => (x.id === id ? { ...x, ...m } : x))); else saveSkills([...skills, { id: ruid("sk"), progress: 0, ...m }]); };
  const setProgress = (id, p) => saveSkills(skills.map((x) => (x.id === id ? { ...x, progress: p } : x)));
  const addWin = () => { if (!winText.trim()) return; saveWins([{ id: ruid("win"), date: today, text: winText.trim() }, ...wins]); setWinText(""); };
  const saveReview = () => { const others = reviews.filter((r) => r.week !== week); const next = reviewText.trim() ? [...others, { week, text: reviewText.trim(), ts: Date.now() }] : others; saveReviews(next.sort((a, b) => a.week.localeCompare(b.week))); };

  return (
    <div>
      <div className="mb-4 flex gap-1.5 rounded-2xl bg-white p-1 shadow-sm ring-1 ring-rose-100">
        {[["path", "Шлях", GraduationCap], ["skills", "Цілі", Target], ["wins", "Досягнення", Trophy], ["review", "Огляд", CalendarDays]].map(([k, label, Icon]) => (
          <button key={k} onClick={() => setTab(k)} className={`flex flex-1 items-center justify-center gap-1 rounded-xl py-2 text-xs font-bold transition sm:text-sm ${tab === k ? "bg-rose-500 text-white shadow" : "text-slate-500 hover:text-slate-700"}`}><Icon className="h-4 w-4 shrink-0" /> {label}</button>
        ))}
      </div>

      {tab === "path" && (
        <div>
          <div className="mb-3 flex gap-1.5 rounded-full bg-slate-100 p-1">
            <button onClick={() => setPathTab("roadmap")} className={`flex-1 rounded-full py-1.5 text-xs font-bold transition ${pathTab === "roadmap" ? "bg-white text-rose-600 shadow-sm" : "text-slate-500"}`}>🧭 Моя роадмапа</button>
            <button onClick={() => setPathTab("book")} className={`flex-1 rounded-full py-1.5 text-xs font-bold transition ${pathTab === "book" ? "bg-white text-rose-600 shadow-sm" : "text-slate-500"}`}>📚 Курс за книгою</button>
          </div>
          {pathTab === "roadmap"
            ? <CareerPath progress={pathProg} onDone={setStepDone} />
            : <CareerPath modules={BOOK_PATH} heading="Курс: Managing IT (по книзі)" progress={bookProg} onDone={setBookStepDone} />}
        </div>
      )}

      {tab === "skills" && (
        <div className="space-y-2">
          {skills.length === 0 && <div className="rounded-2xl bg-white py-8 text-center text-sm text-slate-400 ring-1 ring-rose-50">Що вчиш чи прокачуєш? Додай ціль із дедлайном і відстежуй прогрес.</div>}
          {skills.map((s) => (
            <div key={s.id} className="rounded-2xl bg-white p-4 shadow-sm ring-1 ring-rose-50">
              <div className="flex items-start justify-between gap-2">
                <div className="min-w-0"><div className="font-bold text-slate-800">{s.name}</div>{(s.target || s.deadline) && <div className="text-xs text-slate-400">{s.target}{s.deadline ? ` · до ${s.deadline}` : ""}</div>}</div>
                <span className="shrink-0 text-sm font-bold tabular-nums text-rose-600">{s.progress || 0}%</span>
              </div>
              <input type="range" min={0} max={100} step={5} value={s.progress || 0} onChange={(e) => setProgress(s.id, +e.target.value)} className="mt-2 w-full accent-rose-500" />
              <div className="mt-1 flex gap-2"><button onClick={() => setSkillEd({ skill: s })} className="rounded-full bg-white px-3 py-1 text-xs font-semibold text-slate-500 ring-1 ring-slate-200">Редагувати</button><button onClick={() => { if (confirm("Видалити ціль?")) saveSkills(skills.filter((x) => x.id !== s.id)); }} className="ml-auto text-slate-300 hover:text-red-500"><Trash2 className="h-4 w-4" /></button></div>
            </div>
          ))}
          <button onClick={() => setSkillEd({ skill: null })} className="flex w-full items-center justify-center gap-1.5 rounded-2xl border border-dashed border-rose-300 py-3 text-sm font-semibold text-rose-600 hover:bg-rose-50"><Plus className="h-4 w-4" /> Ціль / навичка</button>
        </div>
      )}

      {tab === "wins" && (
        <div className="space-y-3">
          <div className="rounded-2xl bg-rose-50/60 px-4 py-3 text-xs leading-relaxed text-rose-800">Занотовуй робочі перемоги — великі й малі. Це і для резюме, і щоб на важкий день згадати: ти багато можеш.</div>
          <div className="flex gap-2"><input value={winText} onChange={(e) => setWinText(e.target.value)} onKeyDown={(e) => { if (e.key === "Enter") addWin(); }} placeholder="Що вдалося? (напр. закрила складний баг)" className="min-w-0 flex-1 rounded-xl border border-slate-300 px-3 py-2 text-sm focus:border-rose-400 focus:outline-none" /><button onClick={addWin} className="shrink-0 rounded-xl bg-rose-500 px-4 text-sm font-bold text-white hover:bg-rose-600">+</button></div>
          {wins.length === 0 ? <div className="rounded-2xl bg-white py-8 text-center text-sm text-slate-400 ring-1 ring-rose-50">Ще порожньо. Перша перемога вже сьогодні? 🏆</div> : (
            <div className="space-y-2">{wins.map((w) => <div key={w.id} className="group flex items-start gap-2 rounded-2xl bg-white p-3 shadow-sm ring-1 ring-rose-50"><span className="text-lg">🏆</span><span className="min-w-0 flex-1"><span className="block text-sm text-slate-700">{w.text}</span><span className="text-[11px] text-slate-400">{w.date}</span></span><button onClick={() => saveWins(wins.filter((x) => x.id !== w.id))} className="text-slate-300 hover:text-red-500 sm:opacity-0 sm:group-hover:opacity-100"><X className="h-4 w-4" /></button></div>)}</div>
          )}
        </div>
      )}

      {tab === "review" && (
        <div className="space-y-3">
          <div className="rounded-2xl bg-white p-4 shadow-sm ring-1 ring-rose-100">
            <div className="mb-1 text-sm font-bold text-slate-700">Тиждень від {week}</div>
            <p className="mb-2 text-xs text-slate-400">Що вдалося на роботі цього тижня? Навіть одне речення рахується.</p>
            <textarea value={reviewText} onChange={(e) => setReviewText(e.target.value)} onBlur={saveReview} rows={3} placeholder="Три речі, якими пишаюся цього тижня…" className="w-full rounded-xl border border-slate-200 px-3 py-2 text-sm focus:border-rose-400 focus:outline-none" />
          </div>
          {reviews.filter((r) => r.week !== week).length > 0 && (
            <div className="space-y-2">
              <div className="text-xs font-semibold uppercase tracking-wide text-slate-400">Минулі тижні</div>
              {reviews.filter((r) => r.week !== week).slice().reverse().map((r) => <div key={r.week} className="rounded-2xl bg-white p-3 shadow-sm ring-1 ring-rose-50"><div className="text-[11px] font-semibold text-slate-400">від {r.week}</div><div className="mt-0.5 whitespace-pre-wrap text-sm text-slate-600">{r.text}</div></div>)}
            </div>
          )}
        </div>
      )}

      {skillEd && <SkillEditor skill={skillEd.skill} onClose={() => setSkillEd(null)} onSave={(m) => { upsertSkill(m, skillEd.skill?.id); setSkillEd(null); }} />}
    </div>
  );
}

function SkillEditor({ skill, onClose, onSave }) {
  const [name, setName] = useState(skill?.name || "");
  const [target, setTarget] = useState(skill?.target || "");
  const [deadline, setDeadline] = useState(skill?.deadline || "");
  return (
    <div className="fixed inset-0 z-50 flex items-end justify-center bg-slate-900/40 p-0 backdrop-blur-sm sm:items-center sm:p-4" onClick={onClose}>
      <div className="w-full max-w-md rounded-t-3xl bg-white p-5 shadow-xl sm:rounded-3xl" onClick={(e) => e.stopPropagation()}>
        <div className="mb-3 flex items-center justify-between"><h3 className="text-lg font-bold text-slate-900">{skill ? "Редагувати ціль" : "Нова ціль"}</h3><button onClick={onClose} className="rounded-md p-1 text-slate-400"><X className="h-5 w-5" /></button></div>
        <input value={name} onChange={(e) => setName(e.target.value)} placeholder="Що вчу / прокачую (напр. System Design)" className="mb-2 w-full rounded-xl border border-slate-300 px-3 py-2 text-sm font-semibold focus:border-rose-400 focus:outline-none" />
        <input value={target} onChange={(e) => setTarget(e.target.value)} placeholder="Ціль (напр. пройти курс, зробити pet-проєкт)" className="mb-2 w-full rounded-xl border border-slate-300 px-3 py-2 text-sm focus:border-rose-400 focus:outline-none" />
        <label className="mb-3 block"><span className="mb-1 block text-xs text-slate-500">Дедлайн (необов'язково)</span><input type="date" value={deadline} onChange={(e) => setDeadline(e.target.value)} className="w-full rounded-lg border border-slate-300 px-2 py-1.5 text-sm" /></label>
        <button onClick={() => { if (name.trim()) onSave({ name: name.trim(), target: target.trim(), deadline }); }} className="w-full rounded-2xl bg-rose-500 py-3 font-bold text-white hover:bg-rose-600">Зберегти</button>
      </div>
    </div>
  );
}

/* ---------- Career: Duolingo-style learning path ---------- */
const PATH_STEP_META = {
  learn:   { icon: BookOpen,   ring: "bg-rose-500",  label: "Урок" },
  read:    { icon: BookMarked, ring: "bg-sky-500",     label: "Читання" },
  explain: { icon: Lightbulb,  ring: "bg-amber-500",   label: "Поясни" },
  build:   { icon: Wrench,     ring: "bg-emerald-500", label: "Практика" },
  quiz:    { icon: HelpCircle, ring: "bg-pink-500",  label: "Тест" },
};

// Learning content is baked in (versioned in code); only per-step completion is stored.
const PM_PATH = [
  {
    "slug": "pm-basics",
    "emoji": "🧭",
    "title": "Продуктові основи PM",
    "intro": "Пройшовши цей шлях, ти зрозумієш, як мислить продуктовий менеджер: від правильних питань клієнтам до пріоритизації, MVP, PRD і метрик — і зможеш впевнено вести технічний продукт від ідеї до вимірюваного результату.",
    "steps": [
      {
        "type": "learn",
        "title": "Хто такий PM і як він мислить",
        "body": "Product Manager (PM, продуктовий менеджер) — це людина, яка відповідає за те, ЩО і ЧОМУ команда будує, тоді як інженери вирішують ЯК, а дизайнери — який досвід матиме користувач. PM не начальник над командою і не має прямої влади над розробниками. Його реальний інструмент — вплив через ясність: він приносить контекст, дані й пріоритети, а не накази.\n\nКласична модель описує PM на перетині трьох сил: цінність для бізнесу (viability), потреба користувача (desirability) і технічна здійсненність (feasibility). Хороше рішення живе там, де всі три збігаються. Якщо фіча корисна користувачу, але руйнує економіку компанії — це не продукт. Якщо технічно геніальна, але нікому не потрібна — теж ні.\n\nГоловний зсув мислення для новачка: PM думає не в термінах «зробити фічу», а в термінах «розв'язати проблему й досягти результату». Це різниця між output (ми випустили кнопку) і outcome (користувачі стали на 15% швидше завершувати замовлення). Продукт — це засіб, а не мета.\n\nДля технічного PM додається ще один вимір: ти маєш достатньо розуміти технології (API, бази даних, обмеження моделей ШІ), щоб вести змістовну розмову з інженерами, оцінювати складність і не обіцяти неможливого. Але твоя цінність — не писати код, а перетворювати хаос вимог, даних і думок на чіткий напрямок.",
        "example": "Ситуація: продажі просять «додати експорт у PDF». Слабкий PM ставить це в беклог як є. Сильний питає: яку проблему це розв'язує? Виявляється — бухгалтери клієнта не можуть здати звіт у свою систему. Можливо, правильне рішення — не PDF, а інтеграція через API. Ти рухаєшся від output («кнопка PDF») до outcome («клієнт здає звіт без ручної роботи»).",
        "keyPoints": [
          "PM відповідає за ЩО і ЧОМУ, команда — за ЯК",
          "Рішення живе на перетині цінності, потреби й здійсненності",
          "Думай про outcome (результат), а не output (випущену фічу)",
          "Вплив PM — через ясність і контекст, а не через владу"
        ],
        "resource": {
          "label": "Martin Eriksson — What, exactly, is a Product Manager?",
          "url": "https://www.mindtheproduct.com/what-exactly-is-a-product-manager/"
        }
      },
      {
        "type": "learn",
        "title": "Discovery: перестань вгадувати",
        "body": "Product discovery (продуктове дослідження) — це системна робота з перевірки того, що проблема реальна й варта розв'язання, ПЕРШ НІЖ команда витратить місяці на розробку. Протилежність discovery — «фабрика фіч», де беклог наповнюють ідеями керівництва й здогадками, а потім дивуються, чому ніхто не користується.\n\nКлючова ідея: більшість продуктових ідей помиляються. Дослідження показують, що значна частка фіч не дає очікуваного ефекту. Тому завдання PM — не героїчно вгадати, а швидко й дешево відсіяти погані ідеї. Discovery відповідає на два питання: чи є в людей ця проблема (value risk) і чи розв'яже наше рішення її так, що вони будуть цим користуватися (usability risk).\n\nDiscovery складається з двох тактів. Спершу — розуміння проблеми: інтерв'ю з користувачами, аналіз поведінки, підтримки, даних. Потім — перевірка рішення: прототипи, макети, фейкові кнопки, невеликі експерименти. Обидва такти йдуть паралельно й безперервно, а не «раз на рік перед новим кварталом».\n\nВажлива дисципліна: розділяй проблему і рішення. Новачки закохуються у своє рішення й підганяють під нього факти. Зрілий PM спершу глибоко розуміє проблему, а рішень тримає кілька й обирає найдешевше, що дає результат. «Закохуйся в проблему, а не в рішення» — це не гасло, а робоча звичка.",
        "example": "Команда хотіла будувати складний дашборд аналітики (3 місяці роботи). Замість цього PM за тиждень провів 6 інтерв'ю й показав клікабельний макет у Figma. Виявилося: користувачам потрібні лише 3 метрики й сповіщення на пошту, а не дашборд. Зекономили ~2,5 місяці й зробили те, чим справді користуються.",
        "keyPoints": [
          "Discovery = перевірити проблему до розробки, а не після",
          "Дві загрози: чи є проблема і чи розв'яже її рішення",
          "Розділяй проблему й рішення — закохуйся в проблему",
          "Дешеві експерименти (макети, прототипи) економлять місяці"
        ]
      },
      {
        "type": "read",
        "title": "The Mom Test: як питати клієнтів",
        "body": "The Mom Test — це книга Роба Фіцпатріка й набір простих правил, як говорити з користувачами так, щоб отримувати правду, а не ввічливу брехню. Назва звідси: питання треба ставити так, щоб навіть твоя мама, яка любить тебе й хоче підбадьорити, не змогла збрехати. Бо якщо спитати «Тобі подобається моя ідея?» — усі скажуть «так», і ти нічого не дізнаєшся.\n\nТри правила. Перше: говори про ЇХНЄ життя й минулий досвід, а не про твою ідею. Друге: питай про конкретні події в минулому, а не про абстрактне майбутнє («Розкажи, як ти робив це востаннє» замість «Ти б користувався цим?»). Третє: менше говори, більше слухай.\n\nЧому це критично: люди схильні бути ввічливими й давати компліменти. Компліменти — це отрута для дослідження, бо вони приємні й нічого не значать. Так само небезпечні гіпотетичні («я б точно це купив») і загальні заяви («я зазвичай…»). Тобі потрібні факти: що людина реально робила, скільки часу й грошей витратила, які костилі вигадала.\n\nНайсильніший сигнал — не слова, а дії й зобов'язання. Якщо проблема справжня, людина вже намагалася її розв'язати: платила за щось, вигадувала обхідні шляхи, витрачала час. «Ми обов'язково спробуємо ваш продукт» нічого не варте; «дайте доступ, я познайомлю вас із нашим бухгалтером просто зараз» — варте багато.",
        "example": "Погане питання: «Чи користувався б ти застосунком для планування витрат?» — відповідь завжди «так». Питання за Mom Test: «Розкажи, як ти стежив за витратами минулого місяця. Що використовував? Що дратувало найбільше?» Відповідь: «Веду в Excel, але забуваю вносити, тому в кінці місяця не сходиться» — це реальна проблема й реальний костиль.",
        "keyPoints": [
          "Питай про минулий досвід і факти, не про майбутні гіпотези",
          "Компліменти й «так, класно» — безкорисні дані",
          "Найсильніший сигнал — реальні дії, гроші й час, витрачені раніше",
          "Говори про їхнє життя, а не про свою ідею"
        ],
        "resource": {
          "label": "The Mom Test (офіційний сайт книги)",
          "url": "http://momtestbook.com/"
        }
      },
      {
        "type": "quiz",
        "title": "Перевірка: питання клієнту",
        "body": "Ти проводиш інтерв'ю за The Mom Test. Обери питання, яке дасть найнадійніші дані.",
        "quiz": {
          "question": "Яке питання дасть найнадійніші дані про реальну проблему?",
          "options": [
            "Тобі сподобалася б функція автоматичного нагадування?",
            "Ти купив би наш застосунок за 5 доларів на місяць?",
            "Розкажи, як ти вирішував це завдання минулого разу — що використовував і що дратувало?",
            "Як думаєш, багато людей мали б таку проблему?"
          ],
          "answerIndex": 2,
          "explanation": "Лише варіант про конкретний минулий досвід дає факти; решта — гіпотези та здогади, на які люди відповідають ввічливо."
        }
      },
      {
        "type": "learn",
        "title": "JTBD: робота, яку наймають продукт",
        "body": "Jobs To Be Done (JTBD, «робота, яку треба зробити») — це спосіб дивитися на продукт очима результату, якого прагне користувач, а не через його демографію чи набір фіч. Головна метафора Клейтона Крістенсена: люди не купують продукт — вони «наймають» його, щоб виконати певну роботу у своєму житті, і «звільняють», коли з'являється щось краще.\n\nКласичний приклад — молочний коктейль. Мережа хотіла підняти продажі й опитувала: густіший? солодший? Нічого не спрацювало. Коли придивилися, коли люди купують коктейль, виявилося: вранці, поодинці, у машині по дорозі на роботу. «Робота»: зробити нудну довгу дорогу цікавішою й не бути голодним до обіду. Конкурент коктейлю — не інший коктейль, а банан чи бублик. Це повністю змінює, як його покращувати.\n\nJTBD-формулювання зазвичай має структуру: «Коли [ситуація], я хочу [мотивація], щоб [очікуваний результат]». Воно навмисне не згадує жодного рішення чи технології — лише прогрес, якого людина прагне досягти. Це захищає тебе від передчасної фіксації на фічі.\n\nЧому PM це цінує: JTBD тримає фокус на стабільній потребі, а не на мінливих рішеннях. Технології змінюються, а роботи живуть десятиліттями. Формулювання роботи також підказує, з ким ти насправді конкуруєш і за яким критерієм користувач оцінює успіх.",
        "example": "Робота для сервісу відеодзвінків: «Коли я працюю віддалено й не можу зайти до колеги, я хочу швидко узгодити рішення, щоб не чекати години на листування». Звідси видно: конкурент — не лише Zoom, а й чат і лист; критерій успіху — швидкість узгодження, а не якість відео.",
        "keyPoints": [
          "Користувачі «наймають» продукт заради прогресу, а не заради фіч",
          "Формула: коли [ситуація] — хочу [мотивація] — щоб [результат]",
          "Робота стабільна, рішення мінливі — фокусуйся на роботі",
          "JTBD показує справжніх конкурентів і критерій успіху"
        ]
      },
      {
        "type": "explain",
        "title": "Поясни своїми словами: JTBD",
        "body": "Час на teach-back — найкращий спосіб закріпити знання це пояснити його вголос, ніби колезі.\n\nЗавдання: уяви, що товариш робить мобільний застосунок для вивчення слів іноземної мови. Поясни йому концепцію Jobs To Be Done і сформулюй одну ймовірну «роботу» для його користувача за формулою «коли [ситуація] — хочу [мотивація] — щоб [результат]». Потім назви, з ким насправді конкурує його застосунок (підказка: не лише інші застосунки для слів).\n\nЩо має покривати хороша відповідь: (1) чітке пояснення, що люди «наймають» продукт заради прогресу в конкретній ситуації, а не заради самих функцій; (2) коректно побудована формула роботи без згадки технологій усередині мотивації; (3) визначення справжнього критерію успіху для користувача (наприклад, «за 10 хвилин у транспорті запам'ятати достатньо, щоб не соромитися в розмові»); (4) названі неочевидні конкуренти — подкасти, паперові картки, YouTube, навіть «нічого не робити». Якщо ти легко проговорюєш усі чотири пункти без підглядання — концепцію засвоєно.",
        "keyPoints": [
          "Проговори вголос — так видно прогалини в розумінні",
          "Не згадуй технологію всередині формули роботи",
          "Обов'язково назви неочевидних конкурентів і критерій успіху"
        ]
      },
      {
        "type": "learn",
        "title": "User stories й критерії приймання",
        "body": "User story (користувацька історія) — це короткий опис потреби з погляду користувача, який задає напрямок розробки, не диктуючи технічну реалізацію. Канонічна форма: «Як [роль], я хочу [дію], щоб [цінність/мета]». Історія навмисне коротка — вона не специфікація, а «обіцянка розмови» між PM, дизайном і розробкою.\n\nЧому саме така форма. «Як [роль]» тримає фокус на тому, для кого ми робимо. «Я хочу [дію]» описує потребу. «Щоб [мета]» — найважливіша частина: вона пояснює ЧОМУ, і саме вона часто підказує простіше рішення, ніж просив користувач. Якщо ти не можеш заповнити «щоб…» — ти ще не розумієш задачу.\n\nАле сама історія не каже, коли роботу зроблено правильно. Для цього є критерії приймання (acceptance criteria) — конкретні, перевірювані умови, за яких історія вважається виконаною. Популярний формат — Given/When/Then (Дано/Коли/Тоді): дано певний стан, коли користувач щось робить, тоді система поводиться так-то. Критерії прибирають двозначність і стають основою для тестів.\n\nХороша історія відповідає ознакам INVEST: незалежна, обговорювана, цінна для користувача, оцінювана, мала (влазить у спринт) і тестована. Якщо історія «збудувати платіжну систему» — вона занадто велика; її треба розбити. Для технічного PM критерії приймання — місце, де ти фіксуєш крайні випадки: помилки, порожні стани, обмеження, поведінку офлайн.",
        "example": "Історія: «Як покупець, я хочу зберегти картку, щоб не вводити її щоразу». Критерії: Дано, що я авторизований і на сторінці оплати; Коли я ставлю галочку 'зберегти картку' й успішно плачу; Тоді картка з'являється у профілі як •••• 1234, а наступна оплата пропонує її за замовчуванням. Плюс крайній випадок: якщо оплата не пройшла — картка НЕ зберігається.",
        "keyPoints": [
          "Формула: як [роль] — хочу [дію] — щоб [мета]",
          "Частина «щоб» пояснює ЧОМУ й часто підказує простіше рішення",
          "Критерії приймання (Given/When/Then) роблять «готово» перевірюваним",
          "INVEST: історія має бути малою, цінною і тестованою"
        ]
      },
      {
        "type": "build",
        "title": "Практика: напиши історію з критеріями",
        "body": "Час застосувати теорію руками. Твоє завдання — написати одну повноцінну user story з критеріями приймання для реальної маленької фічі.\n\nВізьми знайомий продукт (наприклад, застосунок доставки їжі) і фічу «повторити минуле замовлення». Зроби так:\n\n1) Напиши історію за формулою «Як [роль], я хочу [дію], щоб [мета]». Перевір, що частина «щоб» пояснює справжню цінність, а не переказує дію.\n\n2) Додай щонайменше 3 критерії приймання у форматі Дано/Коли/Тоді. Один з них має описувати нормальний сценарій (happy path), а решта — крайні випадки: що робити, якщо якоїсь страви вже немає в меню? що, якщо змінилася ціна?\n\n3) Перевір історію за INVEST: чи вона достатньо мала для одного спринту? чи цінна для користувача, а не технічне завдання?\n\n4) Як технічний PM, додай один рядок про технічне обмеження або залежність (наприклад: «залежить від наявності товару в актуальному меню через API ресторану»).\n\nРезультат — це шаблон, який ти зможеш переносити на будь-яку майбутню фічу. Збережи його собі.",
        "keyPoints": [
          "Обов'язково опиши крайні випадки, не лише happy path",
          "Перевір, що «щоб» — про цінність, а не переказ дії",
          "Технічний PM фіксує залежності й обмеження в критеріях"
        ]
      },
      {
        "type": "learn",
        "title": "Пріоритизація: RICE і MoSCoW",
        "body": "Пріоритизація — головна щоденна робота PM: ідей завжди більше, ніж рук, і твоя цінність у тому, ЩО ти вирішуєш НЕ робити. Дві найпопулярніші рамки — RICE (для ранжування списку ідей за балами) і MoSCoW (для узгодження обсягу конкретного релізу).\n\nRICE — чотири фактори. Reach (охоплення): скільки людей це зачепить за період. Impact (вплив): наскільки сильно на кожного, за шкалою (3=масивний, 2=високий, 1=середній, 0.5=низький). Confidence (впевненість): наскільки ти віриш оцінкам, у відсотках (100%=є дані, 50%=здогад). Effort (зусилля): обсяг роботи в людино-місяцях. Формула: RICE = (Reach × Impact × Confidence) / Effort. Що більший бал, то привабливіша ідея. Ключова цінність не в точності числа, а в тому, що Confidence карає красиві ідеї без доказів, а Effort змушує рахувати вартість.\n\nMoSCoW — якісна рамка для меж релізу. Must have (мусить бути): без цього реліз не має сенсу. Should have (мало б бути): важливе, але реліз без нього ще життєздатний. Could have (могло б бути): приємні дрібниці, якщо лишається час. Won't have (цього разу не буде): свідомо винесене за межі — і це найважливіша, найчастіше пропущена категорія, бо саме вона захищає від scope creep (тихого розповзання обсягу).\n\nRICE і MoSCoW не конкурують: RICE допомагає ранжувати беклог, MoSCoW — узгодити межі релізу з командою. Стережися, щоб «Must» не роздувся: якщо все раптом «Must» — пріоритизації не відбулося.",
        "example": "RICE. Ідея А (реєстрація): Reach 5000, Impact 2, Confidence 80%, Effort 2 → (5000×2×0.8)/2 = 4000. Ідея Б (нова тема): Reach 800, Impact 1, Confidence 100%, Effort 1 → 800. Беремо А. MoSCoW для MVP запису на послугу — Must: вибір послуги, часу, підтвердження; Should: нагадування на пошту; Could: вибір майстра; Won't: онлайн-оплата, відгуки.",
        "keyPoints": [
          "RICE = (Reach × Impact × Confidence) / Effort",
          "Confidence карає ідеї без доказів, Effort — за вартість",
          "MoSCoW: Must / Should / Could / Won't — межі релізу",
          "Явне «Won't have» захищає від розповзання обсягу"
        ]
      },
      {
        "type": "quiz",
        "title": "Перевірка: пріоритизація",
        "body": "Ти рахуєш RICE для двох ідей — застосуй формулу й порівняй бали.",
        "quiz": {
          "question": "Ідея X: Reach 1000, Impact 3, Confidence 50%, Effort 3. Ідея Y: Reach 1000, Impact 1, Confidence 100%, Effort 1. Що показує RICE?",
          "options": [
            "X вигідніша: (1000×3×0.5)/3 = 500 проти (1000×1×1)/1 = 1000",
            "Y вигідніша за балом: 1000 проти 500 у X",
            "Вони рівні, бо в обох Reach однаковий",
            "RICE не можна рахувати без грошової оцінки"
          ],
          "answerIndex": 1,
          "explanation": "X = (1000×3×0.5)/3 = 500; Y = (1000×1×1)/1 = 1000. Менші зусилля й повна впевненість роблять Y вигіднішою попри слабший вплив."
        }
      },
      {
        "type": "learn",
        "title": "MVP: мінімально життєздатний продукт",
        "body": "MVP (Minimum Viable Product, мінімально життєздатний продукт) — це найменша версія продукту, яка дозволяє чогось НАВЧИТИСЯ про реальних користувачів із мінімальними витратами. Ключове слово тут — «навчитися», а не «дешево зліпити». MVP — це інструмент перевірки гіпотези, а не урізана бідна версія фінального продукту.\n\nНайпоширеніша помилка новачків — розуміти MVP як «продукт, але поганий». Насправді мета MVP — відповісти на найризикованіше запитання: чи є в людей ця проблема і чи розв'язує її наш підхід. Відома ілюстрація (Хенрік Кніберг): якщо мета — транспорт, не будуй спершу колесо, потім кузов, потім авто — усе це нічого не везе, поки не готове. Краще: спершу скейтборд, потім самокат, потім велосипед — кожен етап уже везе людину й дає зворотний зв'язок.\n\nMVP не обов'язково означає код. Буває «консьєрж-MVP», де послугу спершу надають руками, вдаючи автоматизацію. Буває «фейкові двері» — кнопка фічі, яка веде на сторінку «скоро буде», щоб виміряти попит без розробки. Landing page з формою теж MVP, якщо перевіряє готовність платити.\n\nЩоб MVP був чесним, визнач заздалегідь: яку гіпотезу перевіряєш, який сигнал вважатимеш успіхом (наприклад, «20% тих, хто побачив кнопку, натиснули») і що зробиш за кожного результату. Без цього MVP перетворюється на «просто запустили й дивимось» — і ти знову вгадуєш.",
        "example": "Стартап хотів автоматично підбирати страви за фото холодильника (складний ML). MVP: користувач надсилає фото в чат, а людина за лаштунками вручну пише 3 рецепти за 10 хвилин. За тиждень стало ясно, що ідея потрібна лише 2 з 30 — зекономили місяці на ML-моделі, якої ніхто б не чекав.",
        "keyPoints": [
          "MVP — інструмент навчання, а не «продукт, але поганий»",
          "Кожен етап має вже давати цінність (скейтборд, не колесо)",
          "MVP може бути без коду: консьєрж, фейкові двері, лендинг",
          "Заздалегідь визнач гіпотезу й сигнал успіху"
        ],
        "resource": {
          "label": "Henrik Kniberg — Making sense of MVP",
          "url": "https://blog.crisp.se/2016/01/25/henrikkniberg/making-sense-of-mvp"
        }
      },
      {
        "type": "learn",
        "title": "PRD: документ вимог до продукту",
        "body": "PRD (Product Requirements Document, документ вимог до продукту) — це письмовий артефакт, який відповідає на питання «що ми будуємо, для кого й чому», щоб уся команда мала одну спільну версію правди. Це не бюрократія заради бюрократії: PRD змушує тебе продумати задачу до кінця й уникає ситуації, коли інженер, дизайнер і PM тримають у голові три різні продукти.\n\nСучасний PRD короткий і живий (часто в Notion/Confluence), а не 40-сторінковий фоліант. Типова структура: (1) Проблема й контекст — яку біль розв'язуємо, які дані це підтверджують; (2) Цілі й метрики успіху — як зрозуміємо, що вийшло; (3) Не-цілі (out of scope) — що свідомо не робимо; (4) Користувацькі історії й вимоги з критеріями приймання; (5) Відкриті питання й залежності.\n\nНайважливіша й найчастіше слабка частина — «Проблема» і «Метрики успіху». Якщо PRD одразу стрибає в опис рішення, оминаючи проблему, — це ознака фабрики фіч. Розділ «Не-цілі» так само критичний: він гасить суперечки в середині розробки, бо межі вже узгоджені письмово.\n\nДля технічного PM у PRD додаються технічні міркування: залежності від інших сервісів, обмеження API, вимоги до даних, приватність, поведінка за помилок. Важлива дисципліна: PRD описує ЩО і ЧОМУ, а не диктує ЯК — конкретну архітектуру лишай інженерам. Хороший PRD — живий документ: його оновлюють у міру того, як discovery приносить нове знання.",
        "example": "Міні-PRD фічі 'збереження картки'. Проблема: 40% користувачів кидають повторну оплату, бо щоразу вводять картку (дані з аналітики). Мета: скоротити час оплати; метрика — частка повторних оплат за <30 сек зросте з 45% до 70%. Не-цілі: криптовалюти, розстрочка. Залежності: платіжний провайдер має підтримувати токенізацію.",
        "keyPoints": [
          "PRD = спільна версія правди: що, для кого, чому",
          "Найважливіше — розділи 'Проблема' і 'Метрики успіху'",
          "Розділ 'Не-цілі' гасить суперечки про межі",
          "PRD описує ЩО і ЧОМУ, архітектуру (ЯК) лишай інженерам"
        ]
      },
      {
        "type": "learn",
        "title": "Roadmap і метрики успіху фічі",
        "body": "Roadmap (дорожня карта) — це документ, що показує напрямок продукту в часі: над чим працюємо зараз, що далі й куди рухаємось. Ключова зрілість PM тут — розуміти, що roadmap це виклад намірів і пріоритетів, а НЕ обіцянка точних дат на рік уперед. Роздача жорстких дедлайнів на далеке майбутнє — класична пастка, бо світ і знання змінюються.\n\nСучасний підхід — roadmap, орієнтований на результати (outcome-based), а не на список фіч із датами. Замість «у березні — чат, у квітні — сповіщення» краще: «Ціль кварталу: знизити відтік нових користувачів у перший тиждень; гіпотези-напрямки: онбординг, нагадування». Формат Now / Next / Later (зараз / далі / потім) чесно передає, що ближнє визначене краще за дальнє.\n\nДруга половина зрілості — метрики успіху фічі. Перед запуском ти маєш визначити, як виміряєш, що фіча спрацювала. Розрізняй: North Star (головна метрика цінності продукту), метрики фічі (adoption — скільки спробували, retention — скільки повертаються, engagement — як часто) і guardrail-метрики (запобіжники: те, що НЕ має погіршитися, наприклад швидкість завантаження чи кількість скарг).\n\nПоширена помилка — міряти лише «випустили й скільки клікнули». Сильний PM формулює метрику як гіпотезу з числом ДО запуску («adoption нової оплати досягне 30% за місяць»), а після — чесно дивиться, підтвердилось чи ні. Метрика без цілі наперед — це прикраса, а не рішення.",
        "example": "Roadmap Now/Next/Later. Now: спрощення онбордингу (ціль — активація в 1-й день +10пп). Next: пуш-сповіщення. Later: командні акаунти. Для онбордингу метрики: adoption — % тих, хто завершив; guardrail — час завантаження не зросте понад 2 сек; North Star — тижнева активна аудиторія.",
        "keyPoints": [
          "Roadmap — виклад намірів, а не обіцянка точних дат",
          "Формат Now/Next/Later чесніший за фічі з дедлайнами",
          "Метрики фічі: adoption, retention, engagement + guardrail",
          "Визнач метрику з числом ДО запуску, не після"
        ]
      },
      {
        "type": "explain",
        "title": "Поясни: MVP проти 'сирої версії'",
        "body": "Ще один teach-back, щоб закріпити найлегше плутане поняття курсу.\n\nЗавдання: колега каже — «Наш MVP це просто продукт з мінімумом фіч, зробимо швидко й косо, а потім допиляємо». Поясни йому, чому це нерозуміння MVP, і запропонуй кращий підхід на конкретному прикладі його ідеї (нехай це буде застосунок для пошуку репетиторів).\n\nЩо має покривати сильна відповідь: (1) MVP — це про НАВЧАННЯ й перевірку найризикованішої гіпотези, а не про «поганий продукт»; (2) метафора скейтборд→велосипед: кожен етап уже дає користувачу реальну цінність; (3) конкретний дешевий MVP для репетиторів, який перевіряє попит без повної розробки — наприклад, лендинг чи ручний «консьєрж», де ти сам зводиш учня й репетитора через месенджер; (4) заздалегідь визначений сигнал успіху (скільки заявок за тиждень вважаємо підтвердженням). Якщо ти впевнено пояснюєш різницю між «мало фіч, але цінно» і «просто зроблено погано» — тема твоя.",
        "keyPoints": [
          "MVP міряється навчанням, а не кількістю фіч",
          "Кожен етап має вже давати реальну цінність",
          "Назви конкретний сигнал успіху наперед"
        ]
      },
      {
        "type": "quiz",
        "title": "Мілстоун: типові помилки новачка",
        "body": "Фінальна перевірка перед підсумком — розпізнай найпоширенішу пастку PM-початківця.",
        "quiz": {
          "question": "Що з цього — найтиповіша помилка PM-новачка, про яку попереджав увесь курс?",
          "options": [
            "Закохатися в конкретне рішення й будувати його, не перевіривши, чи реальна проблема",
            "Провести забагато інтерв'ю з користувачами перед стартом",
            "Записати розділ 'Не-цілі' в PRD",
            "Визначити метрику успіху з числом до запуску фічі"
          ],
          "answerIndex": 0,
          "explanation": "Головна пастка — закоханість у рішення без перевірки проблеми (фабрика фіч); решта варіантів — це здорові практики, яких навчав курс."
        }
      },
      {
        "type": "build",
        "title": "Фінал: збери мініпродуктовий пакет",
        "body": "Це фінальне завдання-віха, де ти складаєш усе разом в один невеликий, але цілісний продуктовий пакет. Візьми будь-яку просту ідею продукту (свою або, наприклад, «застосунок для спільних покупок сусідів по квартирі») і пройди повний шлях PM.\n\nЗроби шість кроків, спираючись на пройдені уроки:\n\n1) Discovery: сформулюй 3 питання за The Mom Test, які ти б поставив реальним користувачам, щоб перевірити проблему (лише про минулий досвід, без гіпотез).\n\n2) JTBD: запиши одну «роботу» за формулою «коли [ситуація] — хочу [мотивація] — щоб [результат]».\n\n3) Історія + критерії: напиши одну user story з 3 критеріями приймання (Дано/Коли/Тоді), включно з одним крайнім випадком.\n\n4) Пріоритизація: візьми 3 можливі фічі, порахуй RICE для кожної й розклади їх за MoSCoW (обов'язково заповни 'Won't have').\n\n5) MVP: опиши найдешевший MVP, який перевіряє головну гіпотезу, і назви сигнал успіху з числом.\n\n6) Метрики: назви одну adoption-метрику й одну guardrail-метрику для запуску.\n\nЗбережи результат як власний шаблон. Це і є базовий робочий цикл технічного PM — від правильних питань до вимірюваного запуску. Пройшовши його самостійно бодай раз, ти вже вмієш більше, ніж багато хто на співбесіді на джуніор-PM.",
        "keyPoints": [
          "Пройди повний цикл: discovery → JTBD → історія → пріоритет → MVP → метрики",
          "Обов'язково заповни 'Won't have' і сигнал успіху з числом",
          "Збережи це як власний багаторазовий шаблон PM"
        ]
      }
    ]
  },
  {
    "slug": "tech-depth",
    "emoji": "⚙️",
    "title": "Технічна глибина",
    "intro": "Пройшовши цей шлях, ти зможеш впевнено сидіти на технічних зустрічах, розуміти, як влаштований твій продукт «під капотом», ставити інженерам розумні запитання і приймати рішення, спираючись на реальні обмеження системи, а не на здогадки.",
    "steps": [
      {
        "type": "learn",
        "title": "Клієнт і сервер: хто кому дзвонить",
        "body": "Майже кожен цифровий продукт — це діалог двох сторін. **Клієнт** — це програма, якою користується людина: застосунок на телефоні, вкладка браузера, десктопна програма. **Сервер** — це комп'ютер (частіше багато комп'ютерів) десь у дата-центрі, який зберігає дані, виконує логіку й відповідає на запити клієнтів. Клієнт «дзвонить» серверу, ставить запит (request), сервер обробляє його й повертає відповідь (response).\n\nУяви ресторан. Ти (клієнт) не заходиш на кухню сам — ти кажеш офіціантові, що хочеш. Офіціант (мережа) несе замовлення на кухню (сервер), кухня готує і повертає страву. Ти не бачиш, як саме працює кухня — тобі важливий лише результат і швидкість. Так само браузер не знає, як влаштований сервер: він просто надсилає запит і чекає відповідь.\n\nЧому це важливо для PM? Бо коли ти обговорюєш фічу, майже завжди постає питання «це робимо на клієнті чи на сервері?». Логіка на клієнті — швидша для користувача, але їй не можна довіряти (її легко зламати або підробити) і вона різна на різних пристроях. Логіка на сервері — надійніша, її контролюєш ти, але кожен виклик коштує часу на мережу. Розуміння цього розподілу допомагає тобі оцінювати, чому щось «дорого» або «швидко» реалізувати, і де ховаються ризики безпеки.",
        "example": "Приклад: у застосунку є знижка «-20% для перших 100 покупців». Якщо лічильник покупців рахувати на клієнті, кожен телефон бачитиме свою цифру і система роздасть тисячі знижок. Тому лічильник має жити на сервері — єдине джерело правди, яке бачать усі клієнти однаково.",
        "keyPoints": [
          "Клієнт = те, чим користується людина; сервер = те, що зберігає дані й правила",
          "Спілкування = запит (request) від клієнта → відповідь (response) від сервера",
          "Клієнту не можна довіряти важливу логіку (ціни, права, лічильники) — вона живе на сервері"
        ]
      },
      {
        "type": "learn",
        "title": "HTTP і коди статусів",
        "body": "**HTTP** (HyperText Transfer Protocol) — це «мова», якою клієнт і сервер домовляються про формат запиту та відповіді в інтернеті. Кожен запит містить: метод (що зробити), адресу (URL, з чим), заголовки (headers — службова інформація на кшталт «хто я» чи «який формат я приймаю») і, іноді, тіло (body — самі дані). Відповідь містить код статусу, заголовки й тіло.\n\n**Код статусу** — тризначне число, яким сервер коротко каже, чим усе закінчилося. Їх групують за першою цифрою. **2xx — успіх** (200 OK — усе добре; 201 Created — щось створено). **3xx — перенаправлення** (301 — сторінка переїхала назавжди). **4xx — помилка на боці клієнта**: він щось зробив не так (400 Bad Request — кривий запит; 401 Unauthorized — ти не залогінений; 403 Forbidden — залогінений, але прав нема; 404 Not Found — такого немає; 429 Too Many Requests — забагато запитів). **5xx — помилка на боці сервера**: клієнт усе зробив правильно, але серверу стало погано (500 Internal Server Error — щось зламалося; 503 Service Unavailable — сервіс перевантажений або на техобслуговуванні).\n\nЧому PM це має знати? Бо в баг-репортах, логах і розмовах з підтримкою ці числа звучать постійно. Різниця між 401 і 403 — це різниця між «користувач не увійшов» і «увійшов, але йому не можна» — зовсім різні сценарії й різні фікси. Різниця між 4xx і 5xx каже, чия це проблема: наша чи користувача. Це напряму впливає на пріоритет інциденту.",
        "example": "Приклад тріажу: користувачі скаржаться, що не можуть зберегти профіль. У логах бачиш масові 500 — це наш сервер падає, критичний інцидент, будимо інженерів. Якби це були 400 — значить, клієнт шле кривий запит (можливо, баг у формі), теж наше, але інша команда. А 403 означало б, що це питання прав доступу, а не поломки.",
        "keyPoints": [
          "2xx — успіх, 3xx — перенаправлення, 4xx — винен клієнт, 5xx — винен сервер",
          "401 = не залогінений; 403 = залогінений, але без прав; 404 = не існує",
          "4xx проти 5xx одразу підказує, на чиєму боці шукати причину"
        ],
        "resource": {
          "label": "MDN: HTTP response status codes",
          "url": "https://developer.mozilla.org/en-US/docs/Web/HTTP/Reference/Status"
        }
      },
      {
        "type": "quiz",
        "title": "Квіз: читаємо статус-коди",
        "body": "Перевіримо чуття на кодах статусів — це щоденний інструмент PM.",
        "quiz": {
          "question": "Користувач залогінений, але при спробі відкрити чужий рахунок отримує помилку. Який код найдоречніший?",
          "options": [
            "401 Unauthorized",
            "403 Forbidden",
            "404 Not Found",
            "500 Internal Server Error"
          ],
          "answerIndex": 1,
          "explanation": "403 — особа впізнана (залогінена), але прав на цей ресурс немає. 401 було б, якби вона взагалі не увійшла."
        }
      },
      {
        "type": "learn",
        "title": "REST API: endpoints, методи, JSON, помилки",
        "body": "**API** — набір «дверей», через які одна програма звертається до іншої за даними чи діями. **REST** — найпоширеніший стиль вебʼ-API: усе, чим оперує система, — це **ресурси** (користувачі, замовлення), у кожного є адреса, а дії виражають HTTP-методами. **Endpoint** — адреса ресурсу: `/users` (усі) чи `/users/42` (конкретний). **Методи**: **GET** — прочитати (нічого не змінює); **POST** — створити; **PUT/PATCH** — оновити; **DELETE** — видалити. Читається як речення: `GET /orders/17` — «дай замовлення 17».\n\nВажлива властивість — **ідемпотентність**: GET, PUT і DELETE можна безпечно повторити, результат той самий. А POST — ні: два однакові POST створять два замовлення. Саме тому подвійний клік на «Оплатити» іноді породжує дубль-платіж.\n\nДані клієнт і сервер зазвичай передають у **JSON** — простому тексті з пар «ключ: значення» у фігурних дужках, який читає і людина, і машина. Вміти читати приклад відповіді критично: видно, які поля є, а яких бракує.\n\nОкрема тема — **помилки**. Гарне API повертає не лише код статусу (400), а й тіло з деталями: машиночитабельний код (`INSUFFICIENT_FUNDS`) і зрозуміле повідомлення. Саме з цих кодів клієнт вирішує, який екран показати. Без деталей користувач бачить безлике «Щось пішло не так».",
        "example": "Приклад дизайну API «списку бажань»:\nGET /wishlist — мій список\nPOST /wishlist {productId: 88} — додати товар 88\nDELETE /wishlist/88 — прибрати товар 88\n\nПриклад тіла помилки:\n{ \"error\": { \"code\": \"INSUFFICIENT_FUNDS\", \"message\": \"Недостатньо коштів\", \"balance\": 120, \"required\": 500 } }\nЗавдяки полям code і balance застосунок покаже точний екран замість загального збою.",
        "keyPoints": [
          "Ресурс має адресу (endpoint), метод описує дію: GET/POST/PUT/DELETE",
          "POST не ідемпотентний → повтор створює дублікати (звідси подвійні платежі)",
          "JSON — спільна мова інтеграцій; читай приклад відповіді, щоб бачити поля",
          "Хороша помилка = код статусу + машинний код + людське повідомлення"
        ],
        "resource": {
          "label": "MDN: Вступ до REST / HTTP-методів",
          "url": "https://developer.mozilla.org/en-US/docs/Web/HTTP/Reference/Methods"
        }
      },
      {
        "type": "explain",
        "title": "Поясни інженеру: що таке endpoint",
        "body": "Час на teach-back — найкращий тест розуміння. Уяви, що новий колега-дизайнер питає тебе: «А що взагалі таке цей endpoint і чим GET відрізняється від POST?» Поясни своїми словами, вголос або текстом, за 4–6 речень, без жаргону, з побутовою аналогією.\n\nДобра відповідь охоплює: (1) endpoint — це адреса конкретного ресурсу в системі, як адреса дверей; (2) метод — це намір, що ти хочеш зробити за цими дверима; (3) GET лише читає й нічого не змінює, тому його можна безпечно повторювати; (4) POST створює щось нове, тому повтор породжує дублікат; (5) одна проста аналогія (напр. дошка оголошень: GET — прочитати оголошення, POST — приколоти нове). Якщо ти змогла це пояснити без слова «ну, це технічне» — розділ засвоєно."
      },
      {
        "type": "learn",
        "title": "Бази даних: SQL vs NoSQL",
        "body": "**База даних** (БД) — це організоване сховище, де живуть дані продукту: користувачі, замовлення, повідомлення. Сервер звертається до БД, щоб прочитати або записати інформацію. Два великі світи баз — SQL і NoSQL.\n\n**SQL** (реляційні БД: PostgreSQL, MySQL) зберігають дані у **таблицях** — як в Excel: рядки (записи) і колонки (поля) з чіткою **схемою** (заздалегідь визначеною структурою). Їхня суперсила — звʼязки й узгодженість: можна надійно поєднати таблицю замовлень із таблицею користувачів і гарантувати, що не буде замовлення без покупця. Ідеально там, де дані структуровані й важлива точність — фінанси, облік.\n\n**NoSQL** (MongoDB, DynamoDB, Redis) — парасолька для нереляційних баз. Найчастіше це «документи» (гнучкі JSON-подібні об'єкти) без жорсткої схеми. Їхня суперсила — гнучкість і масштаб: легко зберігати різнорідні дані й швидко читати величезні обсяги. Ідеально для стрічок, каталогів, логів, кешу, персоналізації.\n\nЧому PM це важливо? Бо вибір бази впливає на те, що легко, а що дорого зробити. Додати нове поле в NoSQL — тривіально; у SQL — це зміна схеми (міграція), яку треба планувати. Але порахувати точний фінансовий звіт простіше й безпечніше в SQL. Коли інженер каже «це складно, бо в нас NoSQL і немає join-ів» — тепер ти розумієш, про що йдеться.",
        "example": "Приклад: банківський застосунок тримає рахунки й транзакції в SQL (потрібна залізна узгодженість — гроші не можна «загубити»), а стрічку рекомендацій і кеш сесій — у NoSQL (потрібні швидкість і гнучкість, ідеальна точність не критична). Одна компанія — обидва типи баз під різні задачі.",
        "keyPoints": [
          "SQL = таблиці + жорстка схема + надійні звʼязки; добре для грошей і обліку",
          "NoSQL = гнучкі документи без схеми + масштаб; добре для стрічок, каталогів, кешу",
          "Зміна структури: у NoSQL легко, у SQL — запланована міграція",
          "Багато продуктів використовують обидва типи під різні задачі"
        ],
        "resource": {
          "label": "SQLBolt — інтерактивний курс SQL",
          "url": "https://sqlbolt.com"
        }
      },
      {
        "type": "learn",
        "title": "Індекси: чому пошук буває повільним",
        "body": "Уяви книжку на 800 сторінок без покажчика. Щоб знайти всі згадки слова «латентність», доведеться гортати кожну сторінку. Саме так база даних шукає запис без **індексу** — вона робить **повне сканування таблиці** (full table scan), перебираючи всі рядки один за одним. Коли рядків мільйони, це повільно.\n\n**Індекс** — це окрема допоміжна структура (як алфавітний покажчик у кінці книжки), яка дозволяє базі одразу «стрибнути» до потрібних рядків, не перебираючи всі. Створюєш індекс на колонці `email` — і пошук користувача за поштою стає майже миттєвим. Це часто різниця між запитом на 2 секунди й на 2 мілісекунди.\n\nАле індекси не безкоштовні. По-перше, вони займають місце. По-друге — і це важливіше для продукту — кожен запис або оновлення даних тепер має ще й оновити індекс, тобто запис стає трохи повільнішим. Тому індексують вибірково: колонки, за якими часто шукають або сортують, а не всі підряд.\n\nЧому PM це важливо? Бо «додати індекс» — одна з найчастіших відповідей на скаргу «сторінка вантажиться вічність». Коли інженер каже «цей екран повільний, бо запит не має індексу», ти розумієш і причину, і що фікс зазвичай дешевий. А ще ти навчишся ставити правильне питання: «за якими полями користувачі найчастіше фільтрують?» — бо саме вони кандидати на індекс.",
        "example": "Приклад: адмінка тормозить при пошуку замовлень за номером телефону клієнта. Інженер додає індекс на колонку phone — час відповіді падає з 4 с до 30 мс. Компроміс: створення нового замовлення стало на частку мілісекунди повільнішим, бо оновлюється ще й індекс. Для цього продукту — очевидно вигідний обмін.",
        "keyPoints": [
          "Без індексу база перебирає всі рядки (full table scan) — повільно на великих обсягах",
          "Індекс = покажчик, що дозволяє стрибнути одразу до потрібних рядків",
          "Ціна індексу: більше місця + трохи повільніший запис/оновлення",
          "Індексують колонки, за якими часто шукають чи сортують — не всі підряд"
        ]
      },
      {
        "type": "quiz",
        "title": "Квіз: бази й індекси",
        "body": "Закріпимо різницю між типами баз і роль індексів.",
        "quiz": {
          "question": "Команда скаржиться, що екран пошуку клієнтів за прізвищем «думає» по кілька секунд на великій базі. Який фікс найімовірніший і найдешевший?",
          "options": [
            "Переписати весь продукт із SQL на NoSQL",
            "Додати індекс на колонку з прізвищем",
            "Купити більший монітор для адмінки",
            "Видалити половину клієнтів, щоб таблиця стала меншою"
          ],
          "answerIndex": 1,
          "explanation": "Повільний пошук за конкретною колонкою — класичний симптом відсутнього індексу; додати його дешево й швидко."
        }
      },
      {
        "type": "learn",
        "title": "Кеш і черги: швидше й спокійніше",
        "body": "Дві опори, на яких тримається продуктивність систем, — **кеш** і **черги**.\n\n**Кеш** (cache) — це тимчасове сховище готових відповідей поблизу, щоб не рахувати те саме двічі. Якщо якісь дані запитують часто, а змінюються рідко (курс валют, профіль, головна сторінка), система рахує їх один раз, кладе у швидку пам'ять і наступні тисячі разів віддає звідти — миттєво, не турбуючи базу. Головний виклик кешу — **інвалідація**: коли дані змінилися, старий кеш треба вчасно «протухнути», інакше користувач побачить застарілу інформацію. Звідси відомий жарт інженерів, що дві найскладніші речі — це інвалідація кешу й придумування назв.\n\n**Черга** (queue) — це буфер між тим, хто створює роботу, і тим, хто її виконує. Замість того щоб робити важку задачу прямо зараз і змушувати користувача чекати, система кладе задачу в чергу й повертає відповідь одразу, а окремий «робітник» (worker) виконає її трохи згодом. Так обробляють відправку листів, генерацію звітів, обробку відео. Черга ще й згладжує піки: якщо раптом прийшло 10 000 задач за секунду, вони спокійно чекають, а не кладуть систему.\n\nЧому це важливо для PM? Кеш пояснює, чому дані іноді «застаріли» на кілька хвилин — і це часто свідомий компроміс, а не баг. Черги пояснюють, чому лист «надіслано» приходить не миттєво. Розуміючи це, ти приймаєш кращі рішення про очікування користувача: що показати одразу, а що може відбутися «у фоні».",
        "example": "Приклад: користувач натискає «Згенерувати річний звіт». Без черги — застосунок зависає на 40 секунд. З чергою — одразу з'являється «Звіт готуємо, надішлемо на пошту», задача йде в чергу, worker її рахує, і за хвилину лист із файлом приходить. Користувач не сидить перед крутілкою.",
        "keyPoints": [
          "Кеш = готова відповідь під рукою → миттєва віддача часто запитуваних даних",
          "Плата за кеш — ризик застарілих даних; ключова проблема — вчасна інвалідація",
          "Черга = буфер, що виконує важкі задачі у фоні й згладжує піки навантаження",
          "Разом вони пояснюють «чому дані трохи застаріли» і «чому лист прийшов не миттєво»"
        ]
      },
      {
        "type": "learn",
        "title": "Латентність: звідки береться очікування",
        "body": "**Латентність** (latency) — це час між дією і реакцією: користувач натиснув кнопку — і скільки минуло до появи результату. Її плутають із **пропускною здатністю** (throughput) — скільки запитів система обробляє за секунду. Аналогія з трубою: латентність — скільки одна крапля летить від крана до раковини; throughput — скільки літрів витікає за хвилину. Можна мати товсту трубу (високий throughput) із довгим шляхом (висока латентність).\n\nЛатентність складається з багатьох доданків: час на дорогу мережею (запит фізично летить до сервера й назад — і швидкість світла реальна межа: сервер в іншому континенті — це вже десятки мілісекунд лише на дорогу), час роботи сервера, час запиту до бази, час рендерингу на екрані. Погана новина: доданки додаються. Хороша: зазвичай один-два з них домінують, і саме їх варто чинити.\n\nВажливо мислити не «середньою» латентністю, а **перцентилями**. p50 (медіана) — досвід типового користувача; p95 і p99 — досвід найповільніших 5% і 1%. Часто саме «хвіст» (p99) руйнує враження: середнє «0,3 с» звучить чудово, але якщо кожен сотий запит займає 8 секунд, це тисячі роздратованих людей на масштабі.\n\nЧому PM це важливо? Бо латентність — прямий фактор конверсії й утримання: люди кидають повільні застосунки. Коли ти ставиш ціль «швидко», перетвори її на вимірне: «p95 відкриття стрічки < 1 с». І питай не про середнє, а про хвіст.",
        "example": "Приклад: команда каже «в середньому пошук — 200 мс, усе гаразд». Ти просиш глянути p99 — і виявляється, що він 6 секунд. Копаєте: кожен сотий запит іде без кешу й робить важкий join. Фіксуєте цей хвіст — і зникають найгучніші скарги, хоча «середнє» майже не змінилося.",
        "keyPoints": [
          "Латентність = затримка однієї відповіді; throughput = скільки відповідей за секунду",
          "Затримки додаються, але зазвичай домінує один-два доданки — їх і чинимо",
          "Дивись на перцентилі (p95/p99), а не на середнє: хвіст псує враження",
          "Формулюй ціль вимірно: «p95 < 1 с», а не просто «швидко»"
        ]
      },
      {
        "type": "explain",
        "title": "Поясни другові: латентність проти throughput",
        "body": "Ще один teach-back — на цей раз про продуктивність. Уяви, що нетехнічний друг питає: «А чому кажуть, що система швидка, але при цьому лагає?» Поясни різницю між латентністю й пропускною здатністю простими словами, з однією побутовою аналогією, і чому середнього значення недостатньо.\n\nДобра відповідь торкається: (1) латентність — це скільки чекаєш на одну відповідь; (2) throughput — скільки відповідей система віддає за секунду, і одне не гарантує іншого; (3) аналогія (труба, черга в касу, конвеєр); (4) чому дивимось на p95/p99, а не на середнє — бо повільний «хвіст» псує досвід багатьом людям, навіть коли середнє гарне. Якщо ти чітко розводиш «швидко для одного» і «багато для всіх» — концепт твій."
      },
      {
        "type": "learn",
        "title": "System design і як читати діаграму архітектури",
        "body": "**System design** — це мистецтво скласти з простих цеглинок систему, яка витримає навантаження. Тобі не треба проєктувати її самій, але треба розуміти типові цеглинки й компроміси. Типовий шлях запиту: клієнт → **балансувальник навантаження** (load balancer, розподіляє запити між копіями сервера) → один із багатьох однакових серверів → перевірка **кешу** → якщо нема, звернення до **бази** → важкі задачі йдуть у **чергу**. Кожна цеглинка тобі вже знайома з попередніх уроків.\n\nКлючова ідея — **масштабування**. **Вертикальне** — зробити один сервер потужнішим; просто, але має стелю й одну точку відмови. **Горизонтальне** — багато однакових серверів за балансувальником; складніше, зате майже безмежно й стійко до збоїв. Звідси й **надлишковість** (redundancy): дублюємо критичні частини, щоб падіння однієї не зупинило продукт.\n\nРано чи пізно інженер покаже **діаграму архітектури** — коробочки, стрілочки, підписи. Це карта системи. **Коробки** — компоненти (циліндр традиційно означає базу, прямокутник — сервіс, хмаринка — зовнішній сервіс). **Стрілки** — потік даних і напрям виклику. Читай від точки входу (клієнта) за стрілками до кінця й шукай ризики: єдину базу (вузьке місце?), зовнішні залежності (що впаде?), відсутнє дублювання (одна точка відмови?).\n\nЧому PM це важливо? Бо на діаграмі видно ризики й вартість фічі ще до коду. «Ця фіча додає виклик до зовнішнього API платежів» — це нова стрілка до сторонньої хмаринки, тобто нова залежність.",
        "example": "Приклад читання: бачиш [Мобільний застосунок] → [API-сервіс] → [циліндр: База], збоку [API-сервіс] → [хмаринка: Платіжний провайдер]. Читаєш уголос: застосунок бʼє в наш API, той пише в базу, а для оплат ходить до зовнішнього провайдера. Одразу питаєш: «Що бачить користувач, якщо провайдер недоступний?» — продуктове питання прямо з картинки.",
        "keyPoints": [
          "Запит іде: клієнт → балансувальник → сервер → кеш → база; важке — у чергу",
          "Вертикальне масштабування = потужніший сервер; горизонтальне = багато серверів",
          "Коробки = компоненти (циліндр — база), стрілки = напрям виклику й потік даних",
          "Шукай ризики: єдину базу, зовнішні залежності, відсутнє дублювання"
        ]
      },
      {
        "type": "build",
        "title": "Практика: намалюй архітектуру знайомого застосунку",
        "body": "Час застосувати все на практиці. Візьми продукт, який добре знаєш як користувач (наприклад, застосунок доставки їжі, месенджер або стрічку соцмережі) і намалюй його спрощену архітектуру — на папері, у Figma, Excalidraw чи навіть у нотатках стрілками.\n\nЗавдання по кроках:\n1. Намалюй **клієнта** (застосунок користувача) як точку входу.\n2. Додай **сервер/API**, до якого він звертається, і підпиши 3–4 ключові endpoint-и методами (напр. `GET /restaurants`, `POST /orders`).\n3. Додай **базу даних** і виріши: SQL чи NoSQL для замовлень? Обґрунтуй одним реченням.\n4. Додай хоча б один **кеш** (що саме кешуєш і чому?) і одну **чергу** (яка важка задача йде у фон?).\n5. Додай одну **зовнішню залежність** (карта, оплата, пуші) окремою коробкою-хмаринкою.\n6. Признач одне «гаряче» місце й запиши питання, яке б ти поставила інженеру про його надійність чи латентність.\n\nМета — не технічна досконалість, а вправляння в мисленні цеглинками й компромісами. Якщо змогла обґрунтувати вибір бази й назвати одне вузьке місце — вправу зараховано.",
        "keyPoints": [
          "Склади систему з вивчених цеглинок: клієнт, API, база, кеш, черга, зовнішній сервіс",
          "Обґрунтовуй вибори одним реченням — саме так це роблять на реальних зустрічах",
          "Фінал вправи — сформульоване питання інженеру про ризик або латентність"
        ],
        "resource": {
          "label": "Excalidraw — безкоштовний інструмент для схем",
          "url": "https://excalidraw.com"
        }
      },
      {
        "type": "quiz",
        "title": "Квіз: збираємо все докупи",
        "body": "Підсумковий квіз на розуміння компромісів у дизайні систем.",
        "quiz": {
          "question": "Користувачі скаржаться, що при натисканні «Надіслати звіт» застосунок зависає на 30 секунд. Який підхід найкращий продуктово?",
          "options": [
            "Додати індекс на таблицю звітів",
            "Покласти генерацію звіту в чергу й одразу показати «Готуємо, надішлемо на пошту»",
            "Перевести всю базу з SQL на NoSQL",
            "Збільшити throughput сервера вдвічі"
          ],
          "answerIndex": 1,
          "explanation": "Важку довгу задачу правильно винести у фонову чергу й звільнити користувача від очікування — класичний патерн черг."
        }
      },
      {
        "type": "learn",
        "title": "Словник для розмов з інженерами",
        "body": "Технічна глибина — це не лише концепти, а й спільна мова. Ось компактний словник, який зробить тебе «своєю» на технічних зустрічах. Не зубри — повертайся як до шпаргалки.\n\n**Про роботу й потік:** *бекенд* — серверна частина (логіка, дані); *фронтенд* — те, що бачить користувач; *деплой* — викотити нову версію в продакшн; *прод (prod)* — «бойове» середовище з реальними користувачами, на противагу *staging* (тестовому); *реліз* — випуск версії; *фіче-флаг* — вимикач, що вмикає фічу окремим користувачам без нового деплою; *ролбек* — відкат до попередньої версії, коли щось пішло не так.\n\n**Про дані й надійність:** *міграція* — зміна структури бази; *даунтайм* — час, коли система недоступна; *SLA/SLO* — обіцяний/цільовий рівень надійності (напр. «99,9% часу доступний»); *технічний борг* — накопичені «зрізані кути» в коді, що згодом сповільнюють розробку; *рейт-ліміт* — обмеження на кількість запитів; *таймаут* — коли операція здалася й перестала чекати.\n\n**Фрази-питання, що працюють на зустрічах:** «Це на клієнті чи на сервері?», «Яка очікувана латентність на p95?», «Ця фіча додає нову зовнішню залежність?», «Що бачить користувач, якщо цей сервіс недоступний?», «Це блокуюча операція чи можна у фон?». Такі питання показують інженерам, що ти думаєш про систему, а не лише про макети — і різко піднімають довіру до тебе.",
        "example": "Приклад на грумінгу: інженер каже «зробимо через фіче-флаг, щоб уникнути ризикованого деплою, а якщо латентність на p95 виросте — зробимо ролбек». Ти киваєш і додаєш: «А ця фіча не додає нової зовнішньої залежності на платіжний провайдер?» — і розмова одразу йде по суті, без перекладу.",
        "keyPoints": [
          "Бекенд/фронтенд, деплой, prod/staging, фіче-флаг, ролбек — базовий словник потоку",
          "Міграція, даунтайм, SLA/SLO, техборг, рейт-ліміт, таймаут — словник надійності",
          "Готові питання («на клієнті чи сервері?», «латентність на p95?») роблять тебе учасником",
          "Мета словника — довіра інженерів і розмова без перекладу"
        ]
      },
      {
        "type": "build",
        "title": "Milestone: технічний розбір своєї фічі",
        "body": "Фінальна віха, яка збирає весь модуль в один робочий артефакт. Візьми реальну (або вигадану, але конкретну) фічу зі свого продукту й напиши для неї «технічний one-pager» на пів сторінки, використовуючи всю вивчену лексику. Це саме той документ, який відрізняє технічно грамотного PM.\n\nСтруктура one-pager-а:\n1. **Фіча одним реченням** — що і для кого.\n2. **API** — які 2–4 endpoint-и з методами потрібні (напр. `POST /bookings`, `GET /bookings/{id}`).\n3. **Дані** — де зберігаємо, SQL чи NoSQL і чому; чи потрібні нові індекси й за якими полями.\n4. **Продуктивність** — де ризик латентності; що варто кешувати; чи є важка задача під чергу.\n5. **Надійність** — які зовнішні залежності; що показуємо користувачу при їх збої; чи потрібен фіче-флаг для обережного розкату.\n6. **Питання до інженерів** — 3 конкретні запитання, які ти винесеш на грумінг.\n\nКоли цей документ написано, зроби головне: віднеси його інженеру або ментору й попроси відгук. Якщо людина каже «о, ти справді розумієш, як це працює» — вітаю, ти пройшла шлях «Технічна глибина» і тепер говориш з командою однією мовою.",
        "keyPoints": [
          "One-pager зшиває воєдино API, дані, індекси, кеш, черги, залежності й ризики",
          "Це реальний робочий артефакт, а не навчальна вправа — використовуй у роботі",
          "Головний крок — показати чернетку інженеру й зібрати відгук",
          "Мета всього модуля: говорити з командою однією мовою і мати їхню довіру"
        ]
      }
    ]
  },
  {
    "slug": "ai-literacy",
    "emoji": "🤖",
    "title": "AI-грамотність",
    "intro": "Пройшовши цей шлях, ти зможеш впевнено говорити з інженерами про LLM, приймати обґрунтовані продуктові рішення щодо AI-функцій і не боятися слів «токени», «RAG» чи «галюцинації» — вони стануть твоїми робочими інструментами.",
    "steps": [
      {
        "type": "learn",
        "title": "Що таке LLM і як він генерує текст",
        "body": "LLM (Large Language Model, велика мовна модель) — це нейронна мережа, натренована на величезних обсягах тексту з однією простою метою: передбачити наступне слово. Уяви, що модель прочитала мільярди речень і навчилася вловлювати закономірності мови. Коли ти даєш їй початок фрази, вона не «розуміє» його як людина — вона обчислює ймовірність кожного можливого продовження і обирає найправдоподібніше. Потім додає це слово до тексту й повторює цикл знову й знову. Саме тому кажуть, що LLM — це «дуже потужне автодоповнення».\n\nВажливо зрозуміти інтуїцію: модель генерує текст послідовно, шматок за шматком (ці шматки називають токенами). Вона не планує весь абзац наперед, як письменник із планом. Кожне наступне слово залежить від усього, що вже написано — і від запиту, і від власної відповіді, яку вона будує на ходу. Тому іноді модель «заговорюється» або суперечить сама собі: вона просто йшла за найімовірнішим продовженням.\n\nЧому це важливо для PM? Бо це змінює твої очікування від продукту. LLM — не база даних із точними фактами, а машина ймовірностей. Вона блискуче переформульовує, узагальнює, генерує варіанти — але може впевнено вигадувати. Розуміючи природу «передбачення наступного слова», ти краще формулюєш вимоги, знаєш, де потрібні перевірки, і не обіцяєш стейкхолдерам стовідсоткової точності там, де її архітектурно не буває.",
        "example": "Запит: «Столиця Франції —». Модель бачить цей контекст, обчислює: слово «Париж» має ймовірність ~97%, «місто» ~1%, інше ~2%. Обирає «Париж», додає до тексту, і якщо треба продовжити — рахує наступне слово вже з урахуванням «Столиця Франції — Париж».",
        "keyPoints": [
          "LLM передбачає наступний токен, а не «знає» факти",
          "Текст генерується послідовно, без плану наперед",
          "Це машина ймовірностей — блискуча в мові, ненадійна у фактах",
          "Очікування продукту треба будувати навколо цієї природи"
        ]
      },
      {
        "type": "learn",
        "title": "Токени, контекстне вікно, температура",
        "body": "Три поняття, без яких неможливо говорити про LLM предметно. Токен — це шматочок тексту, яким оперує модель. Це не завжди ціле слово: часто це частина слова або кілька символів. Приблизне правило для англійської — один токен ≈ 0,75 слова, або ~4 символи. Українська зазвичай «важча»: одне слово може розбиватися на кілька токенів. Токени важливі, бо саме за них ти платиш і саме ними вимірюються обмеження.\n\nКонтекстне вікно — це максимальна кількість токенів, яку модель може «тримати в голові» одночасно: твій запит плюс її відповідь. Якщо вікно 128 000 токенів, це весь бюджет на діалог. Коли розмова довша — старіші частини «випадають», і модель їх ніби забуває. Для продукту це критично: якщо ти будуєш чат-бота підтримки, який має пам'ятати всю історію клієнта, треба продумати, як вміщати найважливіше в обмежене вікно.\n\nТемпература — параметр, що керує випадковістю. При температурі 0 модель майже завжди обирає найімовірніше слово: відповіді стабільні, передбачувані, але прісні. При високій (0,8–1,2) вона частіше обирає менш очевидні варіанти: текст стає креативнішим, різноманітнішим, але й ризикованішим. PM має свідомо обирати: для юридичного асистента чи вилучення даних — низька температура; для генерації рекламних слоганів чи мозкового штурму — вища.",
        "example": "Промпт «Придумай назву для кав'ярні». Температура 0 щоразу дає «Coffee House». Температура 1,0 дасть «Ранкова Орбіта», «Зерно і Тиша», «Пар над містом» — різні щоразу. Для брейнштормінгу друге корисніше; для генерації SQL-запиту — навпаки, потрібен нуль.",
        "keyPoints": [
          "Токен ≈ 4 символи; українська дорожча за англійську в токенах",
          "Контекстне вікно = запит + відповідь; понад ліміт — модель «забуває»",
          "Низька температура = стабільність; висока = креативність і ризик",
          "PM обирає температуру під задачу, а не за замовчуванням"
        ]
      },
      {
        "type": "quiz",
        "title": "Перевірка: основи LLM",
        "body": "Швидка перевірка перших двох уроків перед тим, як рухатися далі.",
        "quiz": {
          "question": "Твоя команда будує асистента, що генерує SQL-запити з описів користувача. Яке налаштування температури найдоречніше?",
          "options": [
            "Висока (1,0), щоб запити були різноманітними",
            "Низька (близько 0), щоб відповіді були стабільними й точними",
            "Температура не впливає на код",
            "Середня (0,7), бо це універсальне значення"
          ],
          "answerIndex": 1,
          "explanation": "Для точних, повторюваних завдань як генерація коду потрібна низька температура — креативність тут шкодить, вона додає непотрібні варіації й помилки."
        }
      },
      {
        "type": "learn",
        "title": "Промпт-інжиніринг: роль, приклади, формат",
        "body": "Промпт — це інструкція, яку ти даєш моделі. Промпт-інжиніринг — мистецтво формулювати цю інструкцію так, щоб отримати надійний, корисний результат. Це найдешевший спосіб покращити якість AI-функції: не треба міняти модель чи тренувати щось, достатньо краще написати запит. Для PM це надважлива навичка, бо ти часто перший, хто описує бажану поведінку продукту.\n\nТри найпотужніші прийоми. Перший — роль (persona): скажи моделі, ким їй бути. «Ти досвідчений податковий консультант» задає тон, лексику й глибину відповіді. Другий — приклади (few-shot): покажи 2–3 зразки «вхід → бажаний вихід», і модель підхопить патерн набагато точніше, ніж від абстрактного опису. Це називають few-shot prompting, на противагу zero-shot (без прикладів). Третій — формат: чітко скажи, як має виглядати відповідь — «поверни JSON з полями name і price», «відповідай трьома буліт-пунктами», «максимум 50 слів».\n\nЧому це продуктова навичка, а не інженерна? Бо промпт кодує вимоги до поведінки. Коли ти пишеш специфікацію AI-фічі, добре сформульований промпт — це половина ТЗ. Він визначає тон бренду, обмеження, формат для інтеграції з UI. Погано написаний промпт дає непередбачувану фічу, і жодний фронтенд це не врятує.",
        "example": "Слабкий промпт: «Опиши цей товар». Сильний: «Ти копірайтер бренду преміум-косметики. Напиши опис товару для картки в магазині. Тон: теплий, впевнений. Формат: рівно 2 речення, до 30 слів. Приклад: Вхід: крем для рук → Вихід: Ніжна формула, що вбирається за секунди. Догляд, який хочеться повторювати щодня.»",
        "keyPoints": [
          "Промпт-інжиніринг — найдешевший важіль якості AI-фічі",
          "Роль задає тон і глибину; приклади (few-shot) задають патерн",
          "Явно вимагай формат виходу — це рятує інтеграцію з UI",
          "Добрий промпт = половина специфікації фічі"
        ]
      },
      {
        "type": "build",
        "title": "Практика: напиши структурований промпт",
        "body": "Час застосувати теорію. Уяви фічу: користувач вставляє відгук клієнта, а система має витягти з нього структуровані дані для аналітики.\n\nЗавдання: напиши повний промпт, який використовує всі три прийоми з попереднього уроку. Він має:\n1) Задати роль (наприклад, «Ти аналітик клієнтського досвіду»).\n2) Чітко описати задачу: класифікувати тональність (позитивна / нейтральна / негативна) і виділити головну проблему.\n3) Задати строгий формат виходу — JSON з полями sentiment, main_issue, urgency (low/medium/high).\n4) Містити хоча б один приклад «вхід → вихід» (few-shot).\n\nНапиши цей промпт повністю, ніби віддаєш його інженеру як частину специфікації. Потім перевір себе: чи зрозуміє модель, що робити з відгуком без явної проблеми? Чи однозначний формат для парсингу? Додай у промпт правило на випадок, коли проблеми немає (наприклад, main_issue: null). Це і є реальна робота PM над AI-фічею.",
        "keyPoints": [
          "Роль + задача + формат + приклад = надійний промпт",
          "Завжди продумуй крайні випадки (порожній/неоднозначний вхід)",
          "Строгий JSON-формат робить вихід придатним для коду"
        ]
      },
      {
        "type": "learn",
        "title": "Embeddings і векторний пошук",
        "body": "Embedding (вкладення) — це спосіб перетворити текст на список чисел (вектор), який передає його зміст. Ідея геніально проста: тексти зі схожим значенням отримують схожі вектори, тобто опиняються «поруч» у багатовимірному просторі. «Кошеня» і «маленька кішка» будуть близько; «кошеня» і «податкова декларація» — далеко. Модель-embedder навчена так, щоб відстань між векторами відображала смислову близькість, а не збіг слів.\n\nЦе відкриває семантичний пошук. Класичний пошук шукає збіг ключових слів: запит «як повернути гроші» не знайде статтю «процедура рефанду», бо слова різні. Векторний пошук порівнює зміст: він перетворює і запит, і всі документи на вектори, а потім знаходить документи, чиї вектори найближчі до вектора запиту. Так система розуміє, що «повернути гроші» і «рефанд» — про одне й те саме. Вектори зберігають у спеціальній векторній базі даних (Pinecone, Weaviate, pgvector), оптимізованій для швидкого пошуку найближчих сусідів.\n\nДля PM це фундамент багатьох AI-фіч: розумний пошук по довідці, рекомендації схожих товарів, дедуплікація тікетів, підбір релевантних документів. І, найважливіше, embeddings — це технічна основа RAG, про який далі. Розуміючи embeddings, ти можеш обґрунтовано обіцяти «пошук за змістом, а не за словами» і розумієш, звідки береться його якість і вартість.",
        "example": "Користувач питає бота підтримки: «Мій платіж не пройшов». У базі немає статті з такими словами, але є «Помилки під час оплати карткою». Векторний пошук бачить, що їхні embeddings близькі (обидва про невдалу оплату), і підтягує потрібну статтю — тоді як пошук за ключовими словами повернув би нуль результатів.",
        "keyPoints": [
          "Embedding = вектор чисел, що кодує зміст тексту",
          "Схожий зміст → близькі вектори (навіть за різних слів)",
          "Векторний пошук = семантичний, а не за ключовими словами",
          "Це технічна основа розумного пошуку й RAG"
        ]
      },
      {
        "type": "explain",
        "title": "Поясни своїми словами: семантичний пошук",
        "body": "Найкращий спосіб перевірити розуміння — пояснити концепцію іншому. Уяви, що до тебе прийшов колега-маркетолог, який нічого не знає про AI, і питає: «Чому наш новий пошук знаходить статті, у яких немає слів із мого запиту? Це магія?»\n\nПоясни йому за 4–5 речень, без жаргону, що таке embeddings і векторний пошук. Спробуй знайти власну аналогію (наприклад, карта міста, де схожі за змістом тексти — це будинки на сусідніх вулицях). Уникай слів «вектор» і «розмірність» — або, якщо вживаєш, одразу поясни на пальцях.\n\nЩо має покрити хороша відповідь: (1) текст перетворюється на числа, які передають зміст, а не букви; (2) схожі за змістом тексти отримують схожі числа; (3) пошук порівнює зміст запиту зі змістом статей і бере найближчі; (4) тому «повернути гроші» знаходить «рефанд». Якщо колега після твого пояснення киває й каже «а, тобто воно шукає за сенсом» — ти впорався. Запиши своє пояснення й порівняй із цими пунктами.",
        "keyPoints": [
          "Пояснення без жаргону — тест на справжнє розуміння",
          "Хороша аналогія цінніша за точний термін",
          "Ціль: слухач вловлює «пошук за сенсом, а не за словами»"
        ]
      },
      {
        "type": "learn",
        "title": "RAG: пошук + генерація разом",
        "body": "RAG (Retrieval-Augmented Generation, генерація з доповненням пошуком) — це, мабуть, найважливіша архітектура для продуктових AI-фіч сьогодні. Проблема, яку вона вирішує: LLM знає лише те, що було в тренувальних даних, і нічого — про твою компанію, свіжі документи чи приватну базу знань. Плюс вона схильна вигадувати. RAG елегантно розв'язує обидві біди.\n\nЯк це працює покроково. (1) Ти заздалегідь перетворюєш свої документи на embeddings і кладеш у векторну базу. (2) Коли надходить запит користувача, система робить векторний пошук і знаходить кілька найрелевантніших фрагментів. (3) Ці фрагменти вставляються прямо в промпт як контекст: «Ось релевантні документи: [...]. На їх основі відповідай на питання: [...]». (4) LLM генерує відповідь, спираючись на надані факти, а не на пам'ять. Простими словами: замість того, щоб питати модель «з голови», ти спершу даєш їй потрібну сторінку підручника, а тоді просиш відповісти.\n\nЧому PM це обожнюють. RAG дає актуальність (оновив документ — оновилась відповідь, без перетренування), знижує галюцинації (модель спирається на реальні джерела), дозволяє показувати посилання на джерело (довіра користувача!) і працює з приватними даними, які ніколи не були в моделі. Майже кожен «чат з вашими документами» чи розумний асистент підтримки — це RAG під капотом.",
        "example": "Внутрішній HR-бот. Співробітник питає: «Скільки днів відпустки лишилось у новачків після випробувального?» Модель сама цього не знає. RAG знаходить у HR-політиці абзац про відпустки, вставляє його в промпт, і бот відповідає точно за документом — ще й додає посилання «Джерело: HR-політика, розділ 4.2».",
        "keyPoints": [
          "RAG = знайти релевантні документи + вставити їх у промпт",
          "Дає актуальність без перетренування моделі",
          "Знижує галюцинації і дозволяє показувати джерела",
          "Основа більшості корпоративних AI-асистентів"
        ]
      },
      {
        "type": "quiz",
        "title": "Перевірка: RAG",
        "body": "Переконаймося, що механіка RAG вклалася.",
        "quiz": {
          "question": "У чому головна перевага RAG порівняно з тим, щоб просто запитати LLM «з пам'яті»?",
          "options": [
            "RAG робить модель швидшою і дешевшою за визначенням",
            "RAG дозволяє моделі спиратися на актуальні, приватні документи й показувати джерела",
            "RAG повністю усуває будь-які галюцинації назавжди",
            "RAG замінює потребу писати промпти"
          ],
          "answerIndex": 1,
          "explanation": "RAG підтягує релевантні зовнішні документи в промпт, даючи актуальність, роботу з приватними даними й посилання на джерела. Галюцинації він зменшує, але не усуває повністю."
        }
      },
      {
        "type": "learn",
        "title": "Fine-tuning vs RAG vs промпт",
        "body": "Коли AI-фіча не працює як треба, у команди є три важелі, і плутати їх — дорога помилка. Розберімо кожен і, головне, коли який обирати.\n\nПромпт-інжиніринг — змінюєш лише інструкцію. Найдешевше, найшвидше, миттєвий результат. Підходить, коли модель у принципі здатна виконати задачу, її треба лише правильно направити: тон, формат, кроки міркування. Завжди починай звідси. RAG — даєш моделі зовнішні знання під час запиту. Обирай, коли проблема у нестачі інформації: модель не знає твоїх фактів, даних, свіжих новин. RAG про знання «що», яке легко оновлювати. Fine-tuning (донавчання) — береш готову модель і додатково тренуєш її на своїх прикладах, змінюючи самі ваги. Це дорого, повільно, потребує якісного датасету й інженерів. Обирай, коли треба вбудувати стабільну поведінку чи стиль, який важко описати словами: специфічний формат, галузевий тон, складна класифікація, де промпт роздувається до безкінечності.\n\nПроста ментальна модель. Промпт — це інструкція працівнику. RAG — це видати працівнику довідник на час завдання. Fine-tuning — це відправити працівника на курси, щоб навички засіли назавжди. Ключове правило PM: пробуй у цьому порядку — промпт, потім RAG, і лише як крайній засіб fine-tuning. Дуже часто команди кидаються тренувати модель там, де вистачило б кращого промпта чи RAG — і спалюють бюджет і місяці даремно.",
        "example": "Бот-юрист дає відповіді надто розмито. Питання: у чому проблема? Якщо він не знає свіжих законів — це RAG (підтягнути актуальні норми). Якщо знає, але пише не тим тоном і форматом — спершу промпт. Якщо треба, щоб він завжди відповідав у строгому форматі судових документів попри все — тоді, можливо, fine-tuning на сотнях прикладів.",
        "keyPoints": [
          "Промпт: змінюєш інструкцію — дешево, швидко, пробуй першим",
          "RAG: додаєш зовнішні знання — коли бракує інформації",
          "Fine-tuning: міняєш ваги — коли треба вбудувати стійку поведінку/стиль",
          "Порядок вибору: промпт → RAG → fine-tuning (крайній засіб)"
        ]
      },
      {
        "type": "quiz",
        "title": "Перевірка: який важіль обрати",
        "body": "Класична продуктова дилема — обери правильний інструмент.",
        "quiz": {
          "question": "Ваш чат-бот дає застарілі відповіді про ціни на тарифи, бо вони змінюються щомісяця. Що обрати насамперед?",
          "options": [
            "Fine-tuning на нових цінах щомісяця",
            "Підняти температуру, щоб відповіді були свіжішими",
            "RAG, що підтягує актуальний прайс із бази під час запиту",
            "Нічого, це неможливо виправити"
          ],
          "answerIndex": 2,
          "explanation": "Проблема — у нестачі актуальної інформації, тому RAG ідеальний: оновив прайс у джерелі, і відповіді свіжі. Fine-tuning щомісяця був би абсурдно дорогим і повільним."
        }
      },
      {
        "type": "learn",
        "title": "Галюцинації: чому модель впевнено бреше",
        "body": "Галюцинація — це коли LLM видає інформацію, яка звучить переконливо й правдоподібно, але є вигаданою чи хибною. Модель не бреше зі злим наміром і навіть не «знає», що помиляється. Згадай перший урок: вона передбачає найімовірніше продовження. Якщо правдоподібна за формою відповідь ще й фактично неправильна — модель усе одно її видасть, бо оптимізована на зв'язність, а не на істину. Вона однаково впевнено скаже і правду, і вигадку.\n\nНайпідступніше — упевнений тон. Модель не каже «я не певна»: вона вигадує неіснуючу судову справу з номером і датою, неправильне API-поле чи фейкове наукове дослідження з правдоподібними авторами. Для користувача це виглядає авторитетно, і саме тому галюцинації небезпечні у продукті. У медичній, юридичній чи фінансовій сфері помилка з упевненим тоном може коштувати дуже дорого.\n\nЩо PM може з цим робити. По-перше, RAG і посилання на джерела — коли відповідь спирається на реальний документ, простір для вигадки звужується. По-друге, дизайн інтерфейсу: показувати джерела, додавати дисклеймери, давати кнопку «повідомити про помилку», не подавати AI-відповідь як абсолютну істину. По-третє, обмежувати зону застосування: не пускати AI туди, де ціна помилки надто висока без людської перевірки. Галюцинації не усуваються повністю — ними керують. Твоя задача як PM — спроєктувати продукт так, щоб неминучі помилки були помітні, зворотні й недорогі.",
        "example": "Юрист попросив AI навести прецеденти. Модель видала «Іваненко проти Держказначейства, 2019, справа №821/455» — з ідеальним форматом і впевненістю. Такої справи не існує: модель зібрала правдоподібний за виглядом номер. Без перевірки в реєстрі юрист міг би послатися на неіснуючий прецедент у суді.",
        "keyPoints": [
          "Галюцинація = правдоподібна, але вигадана відповідь",
          "Причина — оптимізація на зв'язність, а не на істину",
          "Найнебезпечніший — упевнений тон без сумніву",
          "Керуй через RAG, джерела, дизайн UI та обмеження зон застосування"
        ]
      },
      {
        "type": "learn",
        "title": "Evals: як вимірювати якість AI",
        "body": "Evals (evaluations, оцінювання) — це систематичний спосіб виміряти, наскільки добре працює твоя AI-фіча. Це AI-аналог тестів у розробці. Проблема унікальна: у LLM немає одного «правильного» виходу, як у звичайній функції. На той самий запит можливі десятки хороших відповідей і сотні поганих. Тому не можна просто написати assert. Потрібен продуманий підхід до вимірювання якості — і це напряму робота PM, бо саме ти визначаєш, що означає «добре» для продукту.\n\nПрактична основа evals — це eval-датасет: набір репрезентативних вхідних запитів разом з очікуваними властивостями відповіді. Далі є кілька способів оцінити вихід. Автоматичні метрики: чи є у відповіді потрібне ключове слово, чи валідний JSON, чи збігається з еталоном. LLM-as-judge: інша модель оцінює відповідь за критеріями («чи ввічливо? чи відповіла на питання? чи спиралася на джерело?») — швидко й дешево масштабується. Людська оцінка: найдорожча, але найнадійніша для тонких речей на кшталт тону чи корисності. Зазвичай комбінують усі три.\n\nЧому без evals не можна запускати AI-фічу. Без них ти покращуєш продукт наосліп: змінив промпт — стало краще чи гірше? Без метрик це вгадування. Evals дають регресійний захист (нова версія не зламала старе), об'єктивне порівняння варіантів і чіткий критерій готовності до релізу. Хороший PM починає думати про evals ще на етапі дизайну фічі: «Як я дізнаюся, що це працює?» — це питання має бути в кожному AI-ТЗ.",
        "example": "Фіча — бот підтримки. Eval-датасет: 50 реальних питань клієнтів + для кожного критерії «дав правильну відповідь / був ввічливим / послався на джерело». Проганяєш нову версію промпта на всіх 50, LLM-суддя ставить оцінки, ти бачиш: точність зросла з 82% до 91%, ввічливість не впала. Тепер рішення про реліз — на цифрах, а не на відчуттях.",
        "keyPoints": [
          "Evals = систематичне вимірювання якості AI-фічі",
          "Основа — eval-датасет із репрезентативних кейсів",
          "Методи: автометрики, LLM-as-judge, людська оцінка (комбінуй)",
          "«Як я дізнаюся, що це працює?» — питання PM ще на дизайні"
        ]
      },
      {
        "type": "build",
        "title": "Практика: спроєктуй eval для фічі",
        "body": "Застосуймо evals на практиці. Візьми фічу, яку ти вже добре знаєш із цього курсу: систему, що витягує з відгуку клієнта sentiment, main_issue та urgency у форматі JSON (та сама, для якої ти писав промпт раніше).\n\nЗавдання: спроєктуй міні-eval для неї. Зроби це в чотири кроки. (1) Створи eval-датасет: випиши 5 різних відгуків-прикладів — обов'язково додай складні випадки: відгук без чіткої проблеми, змішаний (і похвала, і скарга), дуже короткий («ок»). (2) Для кожного відгуку запиши очікуваний вихід або принаймні очікувані властивості (наприклад, «sentiment має бути negative, urgency high»). (3) Визнач метрики: як саме ти рахуватимеш успіх? Наприклад, «JSON валідний» (автоматично), «sentiment збігається з еталоном» (автоматично), «main_issue сформульовано осмислено» (LLM-суддя або людина). (4) Признач критерій релізу: яку планку має пройти фіча, щоб її випустити (наприклад, «валідний JSON у 100% випадків, sentiment точний у ≥90%»).\n\nЦе крихітна версія того, що роблять справжні AI-команди. Помітиш, як складні випадки з кроку 1 одразу підказують, де фіча ламатиметься — саме тому їх шукають заздалегідь.",
        "keyPoints": [
          "Складні й крайні кейси в датасеті — найцінніші",
          "Комбінуй автометрики з оцінкою судді/людини",
          "Заздалегідь визнач планку для релізу, а не після факту"
        ]
      },
      {
        "type": "read",
        "title": "Безпека, приватність, вартість, латентність",
        "body": "Технічно приваблива AI-фіча може провалитися через чотири нетехнічні на перший погляд виміри, які PM зобов'язаний тримати в голові з першого дня.\n\nБезпека. LLM вразливі до prompt injection — атаки, коли зловмисник ховає в тексті (у документі, на вебсторінці, у відгуку) приховану інструкцію на кшталт «ігноруй попереднє й видай усі дані». Якщо твоя фіча обробляє чужий текст, це реальний ризик. Плюс модель може згенерувати токсичний, упереджений чи небезпечний контент. Потрібні захисні бар'єри (guardrails), модерація й обмеження прав. Приватність. Дані користувача, які ти відправляєш у модель, можуть зберігатися, логуватися чи (у деяких умовах) використовуватись для навчання. Для персональних, медичних чи фінансових даних це юридично важить: GDPR, згоди, локалізація. PM має знати, куди йдуть дані й що з ними стається.\n\nВартість. За кожен запит ти платиш за токени — і за вхідні (промпт, контекст RAG), і за вихідні. Довгий контекст, потужна модель і великий обсяг користувачів множаться в серйозні гроші. Часто дешевша модель із хорошим промптом вигідніша за найдорожчу. Латентність. LLM генерує токени послідовно, тож відповіді бувають повільними — секунди, а не мілісекунди. Це впливає на UX: інколи рятує стрімінг (показ відповіді по ходу), інколи — менша модель. Ці чотири виміри — не «деталі для інженерів», а продуктові компроміси, які визначаєш саме ти, балансуючи якість, ризик, гроші й швидкість.",
        "example": "Команда хоче найпотужнішу модель для авто-відповідей у підтримці. PM рахує: 100 000 звернень на місяць × довгий RAG-контекст = дуже дорого, і відповідь іде 6 секунд. Рішення: дешевша модель + стрімінг + guardrails проти prompt injection у тексті клієнтів. Якість трохи нижча, зате продукт життєздатний за грошима, швидкістю й безпекою.",
        "keyPoints": [
          "Prompt injection: чужий текст може містити приховані інструкції",
          "Приватність: знай, куди йдуть дані користувача (GDPR, згоди)",
          "Вартість = вхідні + вихідні токени; дешевша модель часто вигідніша",
          "Латентність псує UX; рятують стрімінг і менші моделі"
        ]
      },
      {
        "type": "explain",
        "title": "Мілстоун: захисти свій продуктовий вибір",
        "body": "Це фінальний рубіж курсу — момент, коли ти збираєш усе разом і говориш як технічний PM з AI-навичками. Уяви, що ти на продуктовому рев'ю. Стейкхолдер каже: «Давайте зробимо AI-асистента для нашої бази знань. Просто візьмемо найрозумнішу модель і донавчимо її на всіх наших документах». Твоя задача — аргументовано запропонувати кращий підхід.\n\nПідготуй усну відповідь (4–6 речень), яка спирається на весь курс. Хороша відповідь має: (1) пояснити, чому fine-tuning тут не перший вибір, а RAG доречніший — документи змінюються, RAG дає актуальність без перетренування й показує джерела; (2) згадати про галюцинації й те, як RAG та посилання на джерела їх зменшують; (3) підняти evals — «як ми виміряємо, що асистент реально корисний перед релізом?»; (4) назвати хоча б два з чотирьох нефункціональних вимірів — вартість, латентність, приватність, безпеку — як речі, що впливають на вибір моделі.\n\nПроговори або запиши цю відповідь повністю. Якщо ти можеш зв'язно провести стейкхолдера від «просто донавчимо модель» до обґрунтованого «почнімо з RAG, задаймо evals і зважмо вартість та приватність» — ти опанувала AI-грамотність на рівні, що реально цінується в продуктових командах. Це і є твій диплом цього модуля.",
        "keyPoints": [
          "Технічний PM обґрунтовує вибір архітектури, а не йде за модою",
          "RAG перед fine-tuning; evals перед релізом; завжди зважуй вартість/приватність",
          "Уміння вести стейкхолдера до правильного рішення — і є AI-грамотність"
        ],
        "resource": {
          "label": "Google People + AI Guidebook — практики дизайну AI-продуктів",
          "url": "https://pair.withgoogle.com/guidebook/"
        }
      }
    ]
  },
  {
    "slug": "build-ai",
    "emoji": "🛠️",
    "title": "Будувати з AI",
    "intro": "Пройшовши цей шлях, ти зможеш власноруч зібрати робочу AI-фічу від ідеї до демо: розумітимеш, як влаштований виклик LLM, писатимеш промпт як специфікацію, скрутиш прототип на no-code, поставиш базові гардрейли й розкажеш про фічу так, щоб команда захотіла її будувати.",
    "steps": [
      {
        "type": "learn",
        "title": "Що взагалі означає «будувати з AI»",
        "body": "Коли ми кажемо «AI-фіча», у 2026 році майже завжди мова про виклик LLM — великої мовної моделі (Large Language Model), як-от Claude чи GPT. Модель — це не база даних з готовими відповідями. Це система, яка передбачає найімовірніше продовження тексту на основі того, що ти їй дав. Через це вона гнучка (розуміє будь-яке формулювання) і водночас непередбачувана (може «вигадати» факт). Технічний PM мусить тримати обидві ці властивості в голові одночасно.\n\nБудувати з AI — не означає навчитися програмувати нейромережі. Для PM це означає інше: вміти зібрати фічу з готових «цеглинок». Цеглинки такі: (1) сама модель, до якої ти звертаєшся через API; (2) промпт — інструкція, яку ти їй даєш; (3) дані користувача, які ти в цю інструкцію підставляєш; (4) обгортка навколо — інтерфейс, перевірки, логіка. У сучасному світі PM може зібрати перші робочі версії всього цього сам, без інженера, за пів дня.\n\nЧому це критично для твоєї кар'єри? Тому що AI зсунув межу того, що PM може перевірити власноруч. Раніше ти писав специфікацію й чекав тижні, щоб побачити, чи ідея взагалі працює. Тепер ти можеш за годину зібрати прототип, показати його трьом користувачам і зрозуміти, чи варто взагалі кликати інженерів. Це називають «prototype your way to a spec» — прототип стає частиною специфікації.",
        "example": "Уяви фічу: користувач вставляє відгук клієнта, а продукт повертає його настрій (позитивний/негативний) і головну скаргу. Розклади на цеглинки: модель = Claude; промпт = «Визнач тон і головну проблему в цьому відгуку, поверни JSON»; дані = текст відгуку; обгортка = проста форма з полем вводу і кнопкою. Усе. Це вже AI-фіча, і жодного рядка нейромережі ти не написав.",
        "keyPoints": [
          "LLM передбачає текст, а не дістає готові відповіді — звідси і гнучкість, і ризик вигадок",
          "AI-фіча = модель + промпт + дані користувача + обгортка навколо",
          "PM тепер може зібрати перший робочий прототип сам, до залучення інженерів"
        ]
      },
      {
        "type": "learn",
        "title": "Як влаштований виклик LLM API",
        "body": "API (Application Programming Interface) — це спосіб, у який одна програма звертається до іншої за чіткими правилами. Виклик LLM через API — це коли твій продукт надсилає моделі запит і отримує назад відповідь. Щоб зрозуміти фічу зсередини, розберемо чотири частини цього обміну: ключ, запит, відповідь і токени.\n\nКлюч (API key) — це секретний рядок, який доводить провайдеру (Anthropic, OpenAI), що звертаєшся саме ти, і за яким тобі рахують рахунок. Ключ = гроші й доступ, тому його ніколи не кладуть у код фронтенду й не світять у браузері — тільки на сервері. Запит (request) — це те, що ти надсилаєш: назва моделі, твої повідомлення (промпт + дані) і налаштування, наприклад «максимум 500 токенів у відповіді». Відповідь (response) — те, що модель повертає: згенерований текст плюс службова інформація, скільки токенів витрачено.\n\nТокен — базова одиниця, якою модель «міряє» текст. Приблизно один токен ≈ 4 символи англійською, або ~0,75 слова. Важливо, бо оплата й ліміти рахуються в токенах: і те, що ти надіслав (input), і те, що модель згенерувала (output). Довгий промпт з великим документом коштує реальних грошей на кожному виклику — PM мусить це закладати в економіку фічі.\n\nЧому PM це знати? Бо з цих деталей складаються три речі, за які ти відповідаєш: вартість (токени × ціна × кількість користувачів), швидкість (довша відповідь = довше чекати) і безпека (де живе ключ).",
        "example": "Спрощений виклик виглядає так: POST на api.anthropic.com/v1/messages, у тілі — {model: 'claude-...', max_tokens: 500, messages: [{role: 'user', content: 'Підсумуй цей відгук: ...'}]}, а в заголовку — секретний x-api-key. Назад приходить {content: 'Клієнт незадоволений термінами доставки', usage: {input_tokens: 120, output_tokens: 18}}. Ті 138 токенів — це те, за що тобі виставлять рахунок.",
        "keyPoints": [
          "Ключ = гроші й доступ: тільки на сервері, ніколи у браузері чи в коді фронтенду",
          "Запит несе модель, повідомлення й ліміти; відповідь несе текст і лічильник токенів",
          "Токени рахуються і на вхід, і на вихід — це пряма економіка фічі",
          "1 токен ≈ 4 символи ≈ 0,75 слова англійською"
        ]
      },
      {
        "type": "learn",
        "title": "Стрімінг: чому відповідь «друкується» на очах",
        "body": "Ти помічала, що ChatGPT чи Claude не показують відповідь одразу цілком, а «друкують» її слово за словом? Це стрімінг (streaming) — режим, у якому модель віддає токени по одному, щойно їх згенерувала, а не чекає, поки складеться вся відповідь. Технічно це потік дрібних шматочків даних, які інтерфейс склеює на екрані в реальному часі.\n\nПротилежність — «блокуючий» режим (non-streaming): ти надсилаєш запит і чекаєш повну відповідь одним куснем. Різниця не в тому, скільки часу займе вся генерація — час майже однаковий. Різниця у сприйнятті. При стрімінгу користувач бачить перше слово вже за пів секунди й розуміє, що система працює. При блокуючому режимі він дивиться на порожній екран 8 секунд і думає, що все зависло.\n\nДля PM це рішення про досвід, а не про технологію. Стрімінг майже завжди краще для довгих відповідей у чаті, бо знижує відчуття очікування й дає користувачу почати читати раніше. Але він не безкоштовний: складніше показати помилку посеред потоку, важче порахувати й показати підсумок (наприклад, «знайдено 3 проблеми»), і не можна перевірити всю відповідь гардрейлами до того, як користувач її вже побачив. Тому для коротких структурованих відповідей (класифікація, JSON, «так/ні») стрімінг часто зайвий.\n\nТвоя робота як PM — свідомо вибрати режим під конкретну фічу, а не брати той, що «за замовчуванням».",
        "example": "Дві фічі, два рішення. Чат-асистент, що пише розгорнуту відповідь на 200 слів → стрімінг, бо інакше 6 секунд порожнечі. Фіча «визнач категорію тікета», що повертає одне слово 'Білінг' → без стрімінгу: відповідь коротка, зате ти встигаєш перевірити, що категорія валідна, перш ніж показати її.",
        "keyPoints": [
          "Стрімінг віддає токени по одному — покращує відчуття швидкості, не саму швидкість",
          "Кращий для довгих чат-відповідей; часто зайвий для коротких структурованих",
          "Мінус стрімінгу: важче перевірити відповідь гардрейлами до показу користувачу"
        ]
      },
      {
        "type": "quiz",
        "title": "Перевірка: API та токени",
        "body": "Швидка перевірка, чи міцно вкладаються основи виклику LLM.",
        "quiz": {
          "question": "PM планує фічу, де до кожного запиту користувача підставляється великий документ на 3000 слів як контекст. Що з цього найбільш точне?",
          "options": [
            "Розмір документа не впливає на вартість — платиш тільки за відповідь моделі",
            "Документ збільшує input-токени на кожному виклику, тож прямо піднімає вартість і затримку",
            "Щоб зекономити, API-ключ треба перенести у код браузера",
            "Токени рахуються лише тоді, коли ввімкнено стрімінг"
          ],
          "answerIndex": 1,
          "explanation": "Платиш і за input, і за output; великий контекст на кожному виклику — це реальні токени, гроші й час."
        }
      },
      {
        "type": "learn",
        "title": "Промпт як специфікація",
        "body": "Найважливіша ідея цього модуля: промпт — це не «чарівні слова», а специфікація продукту, написана природною мовою. У класичному продукті ти пишеш вимоги для інженера, а він перетворює їх на код. З LLM ти пишеш вимоги, і модель виконує їх безпосередньо. Тобто якість твоєї фічі напряму дорівнює якості твого промпту. Нечіткий промпт = нечіткий продукт.\n\nХороший промпт як специфікація має кілька частин. Роль — ким модель має бути («Ти — асистент підтримки»). Задача — що саме зробити, конкретно й однозначно. Контекст — які дані підставляються і що вони означають. Формат виводу — точна форма відповіді (JSON з такими полями, максимум 3 речення, тільки українською). Обмеження й крайні випадки — що робити, якщо даних бракує, якщо запит поза темою, якщо відповіді немає. Приклади — один-два зразки «вхід → правильний вихід», які часто працюють краще за будь-які пояснення.\n\nЧому саме PM має це вміти? Бо промпт — це те місце, де живе продуктова логіка. Рішення «якщо відгук неоднозначний, познач його як \"потребує людини\", а не вгадуй» — це продуктове рішення, не інженерне. Ти найкраще знаєш користувача, крайні випадки й ризики, тож промпт — природно твоя зона.\n\nСприймай промпт як живий документ. Перша версія завжди недосконала; ти будеш її уточнювати на реальних прикладах — про це наступні кроки.",
        "example": "Слабкий промпт: «Підсумуй відгук». Промпт-специфікація: «Ти аналізуєш відгуки клієнтів. Для тексту нижче поверни JSON з полями: sentiment (тільки \"позитивний\"/\"негативний\"/\"нейтральний\"), main_issue (одне речення), needs_human (true, якщо відгук містить погрозу піти чи юридичні претензії). Якщо тексту замало для висновку — sentiment = \"нейтральний\", main_issue = \"недостатньо даних\". Відповідай лише українською. Текст: {відгук}».",
        "keyPoints": [
          "Промпт — це специфікація продукту, а не заклинання: якість фічі = якість промпту",
          "Структура: роль, задача, контекст, формат виводу, крайні випадки, приклади",
          "Продуктова логіка (що робити в спірних випадках) живе у промпті — це зона PM",
          "1–2 приклади «вхід → вихід» часто сильніші за абзац пояснень"
        ]
      },
      {
        "type": "explain",
        "title": "Поясни: чому промпт — це специфікація",
        "body": "Час на teach-back — поясни ідею вголос чи напиши другові, ніби він ніколи не чув про LLM.\n\nЗавдання: за 4–6 речень поясни, чому досвідчені PM кажуть, що «промпт — це специфікація», а не просто питання до чат-бота. Спробуй вжити слова «вимоги», «крайній випадок» і «формат виводу».\n\nЩо має бути в сильній відповіді: (1) думка, що модель виконує інструкцію напряму, тому неточність у промпті = дефект у продукті; (2) що промпт містить ті самі елементи, що й нормальна специфікація — задачу, формат, поведінку в крайніх випадках; (3) чому це саме робота PM, а не інженера — бо тут живуть продуктові рішення про те, як фіча має поводитися з користувачем і його даними. Якщо ти змогла це сказати своїми словами без підглядання — концепт засвоєно."
      },
      {
        "type": "learn",
        "title": "No-code інструменти: зібрати фічу без інженера",
        "body": "No-code — це інструменти, які дають зібрати робочий продукт із візуальних блоків, без написання коду. Для AI-фіч це надсила PM: ти можеш підключити LLM до реальних даних і дій за годину. Розберемо три сімейства.\n\nКастомні GPTs / асистенти (у ChatGPT, а також Projects у Claude) — найшвидший спосіб. Ти по суті зберігаєш свій промпт-специфікацію як окремого «асистента», додаєш кілька файлів як контекст — і маєш робочий інтерфейс, яким можна поділитися й дати людям потестувати. Нуль інтеграцій, хвилини роботи. Ідеально для перевірки самої ідеї й якості промпту.\n\nZapier і Make — інструменти автоматизації: вони з'єднують сервіси за логікою «коли сталася подія X — зроби дію Y». Обидва вміють крок «виклик AI» посередині. Тобто: прийшов новий рядок у Google-таблиці (тригер) → відправ його в LLM з твоїм промптом (AI-крок) → запиши відповідь назад у таблицю або надішли в Slack (дія). Різниця спрощено: Zapier простіший і швидший для лінійних сценаріїв, Make дає більше контролю й розгалужень, але має крутішу криву навчання.\n\nЧому це стратегічна навичка PM, а не «іграшка»? Бо no-code-прототип відповідає на найдорожче питання — «чи варто це взагалі будувати?» — за годину й майже безкоштовно, замість тижнів інженерної роботи. Ти приносиш команді не слайд, а працюючу річ.\n\nМежа чесності: no-code добрий для прототипів і внутрішніх інструментів. Він зазвичай не тягне навантаження мільйонів користувачів чи суворі вимоги безпеки — там уже потрібні інженери. Але щоб довести ідею — це ідеальний інструмент.",
        "example": "Сценарій у Zapier без єдиного рядка коду: тригер — новий тікет підтримки в Gmail; AI-крок — Claude отримує промпт «визнач терміновість (низька/середня/висока) і категорію»; дія — якщо терміновість = висока, надіслати повідомлення в Slack-канал команди. За 30 хвилин у тебе робоча система тріажу тікетів для демо.",
        "keyPoints": [
          "Кастомні GPTs / Projects — найшвидший спосіб перетворити промпт на щось, чим можна ділитися",
          "Zapier/Make: «тригер → AI-крок → дія» з'єднують LLM з реальними даними",
          "No-code відповідає на питання «чи варто будувати?» за годину, а не за тижні",
          "Межа: прототипи й внутрішні інструменти так, великий масштаб і сувора безпека — ні"
        ]
      },
      {
        "type": "quiz",
        "title": "Перевірка: вибір no-code інструмента",
        "body": "Обери найдоречніший інструмент під ситуацію.",
        "quiz": {
          "question": "PM хоче за один день перевірити, чи корисна ідея асистента, який відповідає на питання про внутрішню політику компанії, і дати 5 колегам його потестувати. Найшвидший розумний перший крок?",
          "options": [
            "Написати з інженерами повноцінний бекенд з базою даних і власним UI",
            "Створити кастомний GPT / Project, завантажити документи політики і поділитися посиланням",
            "Одразу будувати складний сценарій у Make з десятьма розгалуженнями",
            "Відкласти ідею, поки не буде затверджено річний бюджет на розробку"
          ],
          "answerIndex": 1,
          "explanation": "Кастомний GPT/Project дає робочий інтерфейс для тесту за хвилини — ідеально, щоб перевірити ідею перед інвестиціями."
        }
      },
      {
        "type": "build",
        "title": "Збери свою першу AI-фічу",
        "body": "Час зробити руками. Мета — мати робочу AI-фічу, якою можна поділитися, до кінця цього кроку. Не ідеальну — робочу.\n\nКроки:\n\n1. Обери маленьку, чітку задачу з одним входом і одним виходом. Приклади: «вставити відгук → отримати настрій і головну скаргу», «вставити чернетку листа → отримати ввічливішу версію», «вставити опис бага → отримати структурований тікет». Уникай багатокрокових сценаріїв — поки що одна дія.\n\n2. Напиши промпт-специфікацію за структурою з кроку про промпт: роль, задача, формат виводу, поведінка в крайньому випадку, один приклад «вхід → вихід».\n\n3. Обери інструмент: найпростіше — створити кастомний GPT або Project у Claude, вставити туди промпт як інструкцію. Хочеш ближче до «продукту» — збери сценарій у Zapier: тригер (форма чи новий рядок таблиці) → AI-крок з твоїм промптом → дія (записати відповідь).\n\n4. Прогони щонайменше 5 різних реальних вхідних прикладів, включно з одним «поганим» (порожній, дивний, поза темою).\n\n5. Запиши, де фіча спрацювала добре, а де зламалася чи вигадала. Цей список — паливо для наступного кроку про ітерацію.\n\nКритерій готовності: у тебе є посилання чи екран, який приймає вхід і повертає осмислений вихід, і нотатки про 5 прогонів. Не переходь далі, поки цього немає — решта модуля спирається на цю живу фічу.",
        "example": "Мінімальний робочий приклад за 20 хвилин: Project у Claude під назвою «Тріаж багів». Інструкція: «Перетвори опис бага від користувача на тікет: поверни JSON з title, steps_to_reproduce (список), severity (low/medium/high), і missing_info (чого бракує, щоб відтворити). Якщо опис надто розмитий — severity = low і чесно напиши, що саме незрозуміло». Тестуєш на п'яти реальних скаргах — і фіча готова до демо."
      },
      {
        "type": "learn",
        "title": "Ітерація по промпту: як робити фічу кращою",
        "body": "Перша версія промпту ніколи не буває фінальною — і це нормально. Ітерація по промпту (prompt iteration) — це системний цикл покращення: прогнати фічу на реальних прикладах, знайти, де вона помилилася, змінити промпт під конкретну помилку, прогнати знову. Ключове слово — системний. Аматор міняє формулювання навмання й сподівається; PM працює за методом.\n\nМетод такий. По-перше, збери набір тестових прикладів (evaluation set, або «evals») — 10–20 реальних входів, для яких ти знаєш правильний вихід. Обов'язково включи важкі й крайні випадки: порожній ввід, двозначний текст, спробу збити фічу з теми. Цей набір — твій вимірювальний прилад. Без нього ти не знаєш, чи зміна промпту зробила краще, чи просто інакше.\n\nПо-друге, міняй одне за раз. Якщо переписав три речення промпту одночасно і стало краще — ти не знаєш, яке саме спрацювало, і не зможеш повторити успіх. По-третє, цілься у конкретну помилку. Модель вигадує факти? Додай у промпт «якщо інформації немає в наданому тексті — напиши \"немає даних\", не вигадуй». Плутає формат? Дай точний приклад бажаного виводу.\n\nЧому це серцевина роботи PM з AI? Бо тут ти напряму формуєш поведінку продукту, спираючись на дані, а не на думки. «Мені здається, стало краще» замінюється на «на нашому наборі з 20 прикладів точність зросла з 14 до 18 правильних». Це вже продуктова інженерія, і вона твоя.",
        "example": "Було: фіча-тріаж інколи ставила severity='high' на дрібниці. Гіпотеза: модель не має визначення рівнів. Зміна (одна!): додаю в промпт «high = блокує роботу або втрата даних; medium = незручно, але є обхід; low = косметика». Прогоняю ті самі 15 прикладів: завищень severity стало 1 замість 6. Зафіксував зміну, беруся за наступну помилку.",
        "keyPoints": [
          "Спершу збери eval-набір з 10–20 прикладів із відомими правильними відповідями",
          "Міняй одне за раз, інакше не знатимеш, що саме спрацювало",
          "Цілься у конкретну помилку конкретною правкою промпту",
          "Заміни «здається, краще» на виміряне «X з 20 правильних»"
        ]
      },
      {
        "type": "explain",
        "title": "Поясни: цикл ітерації по промпту",
        "body": "Ще один teach-back — тепер про метод покращення.\n\nЗавдання: уяви, що новий колега-PM каже: «Я просто переписую промпт, поки відповідь не сподобається». Поясни йому за 5–7 речень, чому це ненадійно і як робити правильно. Обов'язково згадай eval-набір і правило «міняй одне за раз».\n\nСильна відповідь торкнеться: (1) чому «поки не сподобається» оманливе — ти дивишся на 1–2 приклади й не бачиш, що зламав інші; (2) навіщо потрібен фіксований набір прикладів із відомими правильними відповідями — щоб вимірювати, а не відчувати; (3) чому зміна одного за раз дає причинність — інакше не знаєш, яка правка допомогла; (4) що ціль — виміряне покращення на наборі, а не суб'єктивне враження. Якщо змогла це пояснити просто — ти володієш найважливішою щоденною навичкою AI-PM."
      },
      {
        "type": "learn",
        "title": "Базові гардрейли: щоб фіча не нашкодила",
        "body": "Гардрейли (guardrails) — це запобіжники, які тримають AI-фічу в безпечних межах: щоб вона не видавала шкідливого, не вигадувала фактів як правду й не робила того, чого не має. LLM за своєю природою може помилятися впевнено, тому гардрейли — не «приємний бонус», а частина мінімально відповідальної фічі. PM відповідає за них так само, як за саму функцію.\n\nРозберемо базовий набір. Обмеження області (scoping): у промпті прямо задай, що фіча робить і чого не робить — «відповідай лише на питання про наш продукт; на інше ввічливо відмовся». Це знижує ризик, що асистент почне давати юридичні чи медичні поради. Контроль вигадок (галюцинацій): інструктуй модель спиратися тільки на надані дані й чесно казати «не знаю», а не імпровізувати. Галюцинація — це коли модель впевнено видає неправдиву інформацію; для продукту це прямий репутаційний ризик.\n\nПеревірка виходу: якщо фіча має повертати структуру (JSON, одну з трьох категорій), перевіряй відповідь програмно перед показом — і май запасний варіант, якщо модель відповіла не так. Людина в контурі (human-in-the-loop): для важливих чи ризикових рішень фіча не діє сама, а готує чернетку, яку затверджує людина. Прозорість: користувач має знати, що спілкується з AI і що той може помилятися.\n\nЧому це PM, а не «служба безпеки»? Бо рішення про рівень ризику — продуктове. Ти вирішуєш, де достатньо промпт-обмеження, а де потрібна людина в контурі, зважаючи на ціну помилки для користувача. Це компроміс між швидкістю, вартістю й безпекою — класична територія PM.",
        "example": "Фіча-асистент для банку. Гардрейли: scoping — «відповідай лише про тарифи й функції застосунку; питання про конкретні транзакції клієнта переадресуй людині». Анти-галюцинація — «якщо тарифу немає в наданому списку, скажи \"уточніть у підтримці\", не називай цифру». Human-in-the-loop — асистент готує відповідь про закриття рахунку, але надсилає її тільки після підтвердження оператором. Прозорість — плашка «Це AI-асистент, можливі помилки».",
        "keyPoints": [
          "Гардрейли — обов'язкова частина фічі, а не бонус: LLM помиляється впевнено",
          "Базовий набір: обмеження області, контроль галюцинацій, перевірка виходу, людина в контурі, прозорість",
          "Галюцинація = впевнено подана неправда; проти неї — «спирайся лише на надані дані, кажи \"не знаю\"»",
          "Рівень гардрейлів — продуктове рішення PM, залежить від ціни помилки"
        ]
      },
      {
        "type": "quiz",
        "title": "Перевірка: гардрейли",
        "body": "Обери найсильніший гардрейл під ситуацію.",
        "quiz": {
          "question": "AI-асистент у медичному застосунку інколи впевнено називає дозування ліків, яких немає в перевіреній базі. Який гардрейл б'є прямо в цю проблему?",
          "options": [
            "Увімкнути стрімінг, щоб відповідь з'являлася швидше",
            "Інструктувати модель відповідати лише на основі наданої бази і писати «зверніться до лікаря», якщо даних немає, плюс людина в контурі для дозувань",
            "Зробити відповіді моделі довшими й детальнішими",
            "Прибрати з інтерфейсу згадку, що це AI, щоб не лякати користувачів"
          ],
          "answerIndex": 1,
          "explanation": "Проти впевнених вигадок працює зв'язка «спирайся лише на перевірені дані + кажи не знаю + людина в контурі» для ризикових рішень."
        }
      },
      {
        "type": "learn",
        "title": "Демо і storytelling: продати фічу за 3 хвилини",
        "body": "Зібрати фічу — половина справи. Друга половина — зробити так, щоб команда, керівництво чи користувачі зрозуміли її цінність за лічені хвилини. Демо — це не «показ кнопок», це історія, у якій твоя фіча розв'язує чийсь реальний біль. PM, який уміє розповісти, отримує ресурс на розробку; PM, який просто «показує екран», — ні.\n\nРобоча структура демо коротка. Спершу — біль: чиясь конкретна проблема в одному-двох реченнях («Оператори підтримки вручну сортують 200 тікетів на день, це 3 години рутини»). Далі — обіцянка: що змінюється («Уяви, що тікети сортуються самі за 2 секунди»). Потім — жива демонстрація: покажи фічу на реальному, впізнаваному прикладі, не на «ідеальному» вигаданому. Далі — чесність: одним реченням визнай межі й ризики («На двозначних тікетах точність нижча, тому їх ми віддаємо людині»). І фінал — заклик: що саме ти просиш («Прошу два тижні інженера, щоб довести до продакшену»).\n\nКілька правил, які рятують демо. Показуй, а не розказуй — жива фіча переконує в рази сильніше за слайд. Обери приклад, який болить саме цій аудиторії. Не ховай недоліки: якщо фіча інколи помиляється — покажи це сама й скажи, як з цим працюватимеш; так ти виглядаєш чесним експертом, а не продавцем. І май запасний план — записаний ролик чи скріншоти на випадок, якщо жива демка зламається.\n\nЧому це серцевина ролі? Бо продукт, який ніхто не зрозумів, не існує. Уміння перетворити технічну штуку на зрозумілу історію з ясним «навіщо» — це те, що відрізняє PM від просто виконавця.",
        "example": "Демо тріажу тікетів за 90 секунд: «Наша команда підтримки вручну сортує 200 тікетів щодня — 3 години рутини (біль). Ось що я зібрав за день (обіцянка). Дивіться: беру реальний вчорашній тікет, вставляю — за 2 секунди маю категорію й терміновість (жива демка). На нечітких тікетах точність падає, тому такі тікети фіча позначає \"на людину\" (чесність). Прошу два тижні інженера, щоб інтегрувати це в наш helpdesk (заклик)».",
        "keyPoints": [
          "Демо — це історія «біль → обіцянка → жива демка → чесність про межі → заклик до дії»",
          "Показуй живу фічу на реальному впізнаваному прикладі, не на ідеальному вигаданому",
          "Сама назви недоліки — це додає довіри, а не забирає її",
          "Завжди май запасний план (запис, скріншоти) на випадок збою демки"
        ]
      },
      {
        "type": "build",
        "title": "Майлстоун: доведи й покажи свою AI-фічу",
        "body": "Фінальний крок, який зшиває весь модуль. Мета — взяти фічу, яку ти зібрала раніше, довести її ітерацією, захистити гардрейлами і представити як історію. Це твій завершений артефакт для портфоліо AI-PM.\n\nЗавдання:\n\n1. Візьми фічу зі свого попереднього build-кроку.\n\n2. Ітерація: збери eval-набір з 10+ реальних прикладів (включно з крайніми), прогони, знайди 2–3 типи помилок і виправ їх точковими правками промпту — по одній за раз. Запиши «було → стало» у цифрах (напр. «12/15 → 14/15 правильних»).\n\n3. Гардрейли: додай у промпт щонайменше два — обмеження області і контроль галюцинацій («спирайся лише на надані дані»). Якщо фіча повертає структуру — продумай, що робити, коли модель відповіла не в тому форматі.\n\n4. Демо: напиши сценарій демо на 2–3 хвилини за структурою «біль → обіцянка → жива демка → чесність про межі → заклик». Прогони його вголос на реальному прикладі.\n\n5. Збери все в одну коротку сторінку: що за фіча, для кого, як покращилась у цифрах, які гардрейли, і сценарій демо.\n\nКритерій готовності: у тебе є (а) робоча фіча, (б) докази покращення в цифрах, (в) щонайменше два гардрейли, (г) відрепетируваний 3-хвилинний сценарій демо. Це повний цикл AI-PM — від ідеї до історії. Саме такий артефакт варто показувати на співбесіді: він доводить не «я знаю про AI», а «я вмію будувати з AI».",
        "example": "Готовий майлстоун-артефакт (одна сторінка): «Фіча: авто-тріаж багів для команди підтримки. Ітерація: набір з 15 реальних багів; після трьох правок промпту точність severity зросла з 9/15 до 13/15, завищень high — з 6 до 1. Гардрейли: (1) якщо опис розмитий — severity=low + перелік того, чого бракує; (2) не вигадувати кроки відтворення, яких немає в тексті. Демо-скрипт: біль (3 год рутини на день) → обіцянка → жива демка на вчорашньому тікеті → чесність (нечіткі баги йдуть на людину) → заклик (2 тижні інженера)». Цей документ — доказ навички, а не слова про неї.",
        "keyPoints": [
          "Майлстоун зшиває все: фіча + виміряна ітерація + гардрейли + демо-історія",
          "Показуй покращення в цифрах — це мова, якій довіряють стейкхолдери",
          "Готовий артефакт доводить уміння будувати, а не просто знання про AI",
          "Саме таку одну сторінку варто нести на співбесіду AI-PM"
        ]
      }
    ]
  },
  {
    "slug": "analytics",
    "emoji": "📊",
    "title": "Аналітика й метрики",
    "intro": "Пройшовши цей шлях, ти навчишся дивитися на продукт очима даних: обереш одну головну метрику, зрозумієш, де саме користувачі відвалюються, напишеш перші SQL-запити, сплануєш чесний A/B-тест і навчишся ставити дашборду правильні запитання — усе те, що відрізняє сильного Technical PM від того, хто просто вірить інтуїції.",
    "steps": [
      {
        "type": "learn",
        "title": "North-Star метрика",
        "body": "Метрика — це число, яке ти вимірюєш регулярно, щоб відповісти на конкретне питання про продукт. Але коли метрик стає багато, команда починає тягнути продукт у різні боки: маркетинг оптимізує завантаження, продажі — угоди, розробка — час відгуку. North-Star Metric (NSM), або \"полярна зоря\", — це одна головна метрика, яка найкраще відображає ту цінність, що продукт дає користувачам, і навколо якої вирівнюється вся компанія.\n\nВажливо розрізняти метрики результату (output, vanity) і метрики цінності (outcome). Vanity-метрики показують обсяг активності: кількість завантажень, лайків, переглядів. Вони приємні, але часто оманливі — мільйон завантажень нічого не каже, якщо ніхто не повертається. Outcome-метрики показують, чи отримує користувач реальну користь: чи повертається, чи виконує ключову дію, чи платить.\n\nГарна NSM має три властивості. По-перше, вона відображає цінність для користувача, а не лише дохід: якщо користувачам добре, гроші прийдуть услід. По-друге, вона випереджальна (leading) — зростає раніше, ніж дохід у звіті, що дає змогу коригувати курс завчасно. По-третє, на неї можна впливати продуктом. Класичні приклади: у Spotify — \"час прослуховування\", у Airbnb — \"кількість заброньованих ночей\", у WhatsApp — \"кількість надісланих повідомлень\". Під зорею будують дерево з 3-5 вхідних метрик (input metrics) — саме на них команди впливають щодня, і їхня сума рухає зорю. Роль PM — обрати правильну зорю й не дати команді забути, заради чого вона працює.",
        "example": "Уяви застосунок для доставки їжі. Кандидати в NSM: (а) кількість завантажень — vanity; (б) виручка — наслідок; (в) \"кількість успішно доставлених замовлень на тиждень\" — саме те. Дерево під нею: активні ресторани × конверсія в замовлення × повторюваність × середня частота. Команда логістики впливає на своєчасність доставки, і це видно в зорі.",
        "keyPoints": [
          "NSM — одна метрика, що вирівнює всю компанію навколо цінності",
          "Vanity показує обсяг, outcome — реальну користь",
          "Гарна NSM: цінність для юзера, випереджальна, керована продуктом",
          "Під зорею — дерево з 3-5 input-метрик, на які команди впливають щодня"
        ],
        "resource": {
          "label": "Amplitude: North Star Playbook (безкоштовно)",
          "url": "https://amplitude.com/north-star"
        }
      },
      {
        "type": "quiz",
        "title": "Перевірка: обери North-Star",
        "body": "Команда навчального застосунку обирає north-star метрику. Яка найкраща?",
        "quiz": {
          "question": "Який із варіантів найкраще підходить на роль North-Star Metric для застосунку вивчення мов?",
          "options": [
            "Кількість завантажень застосунку за місяць",
            "Загальна виручка компанії за квартал",
            "Кількість користувачів, які завершили хоча б один урок протягом тижня",
            "Кількість зірок у сторі застосунку"
          ],
          "answerIndex": 2,
          "explanation": "Завершений урок — це момент реальної цінності, метрика випереджальна й керована продуктом. Завантаження й зірки — vanity, виручка — наслідок."
        }
      },
      {
        "type": "learn",
        "title": "AARRR-воронка (піратські метрики)",
        "body": "AARRR — це модель, яка розкладає весь шлях користувача на п'ять етапів. Назву жартома читають як піратське \"Аррр!\", тому її ще звуть \"піратськими метриками\". Вона допомагає PM побачити, на якому саме кроці губляться люди, і не плутати різні проблеми між собою.\n\nП'ять етапів такі. Acquisition (залучення) — як людина вперше дізналася про продукт і зайшла. Activation (активація) — чи отримала вона перше \"ага!\", перший момент цінності. Retention (утримання) — чи повертається вона знову. Revenue (дохід) — чи платить. Referral (рекомендація) — чи приводить друзів. Порядок не випадковий: немає сенсу лити гроші в залучення, якщо люди не активуються й одразу йдуть — ти просто наповнюєш діряве відро.\n\nДля PM AARRR — це діагностична карта. Коли бізнес каже \"у нас проблема зі зростанням\", перше питання: на якому етапі воронки? Якщо люди заходять, але не активуються — проблема в онбордингу. Якщо активуються, але не повертаються — продукт не тримає. Якщо повертаються, але не платять — питання цінності чи упаковки. Кожен етап має свою метрику й свої важелі, і плутати їх — типова помилка початківця. Далі ми детально розберемо два найважливіші етапи: активацію та утримання.",
        "example": "Онлайн-магазин: Acquisition — 10 000 візитів; Activation — 4 000 створили кошик; Retention — 1 500 повернулися за тиждень; Revenue — 900 купили; Referral — 120 поділилися. Видно, що найбільший провал між візитом і кошиком (60% відвалюються) — отже, працювати треба над першим враженням, а не над рекламою.",
        "keyPoints": [
          "AARRR: Acquisition, Activation, Retention, Revenue, Referral",
          "Порядок важливий: не лий гроші в залучення при дірявому утриманні",
          "Воронка — діагностична карта: спершу знайди етап, де губляться люди"
        ],
        "resource": {
          "label": "Lenny's Newsletter (публічні пости про growth)",
          "url": "https://www.lennysnewsletter.com/"
        }
      },
      {
        "type": "learn",
        "title": "Activation: перший момент цінності",
        "body": "Активація — це момент, коли новий користувач уперше відчуває реальну користь від продукту, те саме \"ага!\". Це найкритичніший етап воронки, бо якщо людина не зрозуміла цінності в перші хвилини, вона піде й майже ніколи не повернеться. Уся робота маркетингу з залучення згорає саме тут, якщо активація слабка.\n\nЩоб працювати з активацією, PM визначає activation event — конкретну вимірювану дію, яка сильно корелює з довгостроковим утриманням. Легендарний приклад: Facebook свого часу виявив, що користувачі, які додали 7 друзів за 10 днів, залишалися надовго. Це стало їхньою \"північною зіркою активації\". У Slack таким сигналом були 2000 надісланих повідомлень командою, у Dropbox — покладений хоча б один файл у папку. Знайти такий поріг можна, порівнявши поведінку тих, хто залишився, з тими, хто пішов: яка рання дія найкраще їх розділяє?\n\nДля PM це означає конкретну роботу. По-перше, скоротити шлях (time-to-value) до цього моменту: прибрати зайві кроки реєстрації, показати цінність до вимоги даних. По-друге, вимірювати activation rate — частку нових користувачів, які досягли події. По-третє, проєктувати онбординг так, щоб вести людину прямо до \"ага\", а не показувати весь функціонал одразу. Активація — це не про красиві екрани привітання, а про те, щоб людина якнайшвидше зробила ту дію, заради якої прийшла.",
        "example": "Застосунок для нотаток визначає activation event як \"створив 3 нотатки в перший день\". Аналіз показав: хто перетнув цей поріг, у 4 рази частіше активний через місяць. PM спрощує онбординг: замість туру з 6 екранів одразу відкриває порожню нотатку з підказкою \"Запиши першу думку\". Activation rate зростає з 22% до 35%.",
        "keyPoints": [
          "Активація = перший момент реальної цінності (\"ага!\")",
          "Activation event — рання дія, що корелює з довгим утриманням",
          "PM скорочує time-to-value і веде онбординг прямо до цієї дії"
        ]
      },
      {
        "type": "learn",
        "title": "Retention і churn",
        "body": "Retention (утримання) — це частка користувачів, які повертаються до продукту через певний час. Churn (відтік) — дзеркальна метрика: частка тих, хто пішов. Якщо retention за місяць 70%, то місячний churn — 30%. Це, мабуть, найважливіша пара метрик у продукті: залучення нових користувачів дороге, а утримання наявних — це те, що робить бізнес життєздатним. Продукт із поганим утриманням схожий на відро з діркою: скільки не наливай, воно не наповниться.\n\nВажливо розрізняти типи утримання. N-day retention — чи повернувся користувач саме на N-ту добу (жорстка метрика, добре для щоденних продуктів). Rolling retention — чи повернувся він у будь-який день починаючи з N-го (м'якша, для рідших сценаріїв). Ще будують криву утримання: по осі X — дні від реєстрації, по осі Y — частка активних. Здоровий продукт має криву, яка після падіння виходить на плато (\"усмішку утримання\") — це ядро лояльних користувачів. Крива, що падає до нуля, означає, що продукт нікого не тримає.\n\nДля PM churn — це діагностика. Треба питати: коли саме люди йдуть (одразу після реєстрації чи через місяць)? Хто йде (який сегмент)? Чому — через ціну, брак цінності, баги, кращого конкурента? Розрізняють добровільний відтік (людина свідомо пішла) і мимовільний (наприклад, не пройшов платіж). Зниження churn навіть на кілька відсотків часто дає більший ефект на зростання, ніж збільшення залучення.",
        "example": "SaaS-сервіс має 5% місячного churn. Здається мало, але за рік це означає, що зі 100 клієнтів залишається лише ~54 (0.95^12). PM сегментує відтік і бачить: клієнти, які не інтегрували продукт із Slack у перший тиждень, ідуть удвічі частіше. Рішення — підштовхнути цю інтеграцію в онбордингу.",
        "keyPoints": [
          "Retention + churn = 100%; утримати дешевше, ніж залучити",
          "N-day (жорстка) vs rolling (м'яка) — обирай за частотою продукту",
          "Здорова крива виходить на плато (\"усмішка\"), а не падає в нуль",
          "Churn — це діагностика: коли, хто і чому йде"
        ]
      },
      {
        "type": "explain",
        "title": "Поясни: чому утримання важливіше",
        "body": "Час перевірити розуміння, пояснивши концепцію своїми словами (техніка teach-back — якщо можеш пояснити просто, значить справді зрозумів).\n\nУяви, що до тебе прийшов засновник стартапу й каже: \"У нас чудове зростання — щомісяця 10 000 нових реєстрацій! Треба вкласти ще більше в рекламу.\" Твоє завдання — пояснити йому простими словами, чому спершу варто подивитися на retention, а не заливати гроші в залучення.\n\nСпробуй сформулювати відповідь на 4-6 речень. Хороша відповідь має торкнутися таких ідей: метафора дірявого відра (нові користувачі витікають, якщо утримання слабке); те, що залучення коштує грошей, а поганий retention їх знецінює; що спершу треба подивитися на криву утримання й перевірити, чи виходить вона на плато; і що іноді правильніше спочатку полагодити активацію та утримання, а вже потім масштабувати залучення. Бонус — згадати, як саме ти б це виміряв (когортна крива утримання).",
        "keyPoints": [
          "Teach-back: пояснення простими словами виявляє прогалини в розумінні",
          "Сильна відповідь поєднує метафору, метрику й конкретний наступний крок"
        ]
      },
      {
        "type": "learn",
        "title": "Когортний аналіз",
        "body": "Когорта — це група користувачів, об'єднаних спільною ознакою в часі, найчастіше — місяцем реєстрації. Когортний аналіз означає, що ми стежимо за поведінкою кожної такої групи окремо в часі, замість того щоб дивитися на один усереднений показник по всіх. Це один із найпотужніших інструментів PM, бо середні числа брешуть.\n\nЧому брешуть? Уяви, що загальний retention продукту тримається на 40% уже пів року — здається, стабільно. Але якщо розкласти по когортах, може виявитися, що старі когорти деградують до 20%, а нові стартують із 60% завдяки покращеному онбордингу. Середнє маскує обидва тренди. Когорти показують правду: чи стає продукт кращим для нових користувачів із часом.\n\nКласична форма подання — когортна таблиця утримання. Рядки — когорти (наприклад, \"зареєструвалися в січні\", \"у лютому\"). Стовпці — скільки часу минуло (тиждень 0, 1, 2...). У клітинках — частка активних. Читаючи по рядку, ти бачиш, як конкретна когорта згасає; читаючи по стовпцю вниз, порівнюєш, чи новіші когорти тримаються краще за старі. Для PM це основний спосіб довести, що зміна в продукті справді покращила утримання: якщо когорти після релізу нового онбордингу тримаються вище — фіча спрацювала. Когорти можна будувати не лише за датою, а й за каналом залучення, планом підписки чи країною.",
        "example": "Когортна таблиця (частка активних):\n\nКогорта | Тиж0 | Тиж1 | Тиж2 | Тиж3\nСічень | 100% | 45% | 30% | 22%\nЛютий  | 100% | 48% | 33% | 25%\nБерезень| 100% | 58% | 44% | 38%\n\nБерезнева когорта тримається помітно краще — саме тоді викотили новий онбординг. Це доказ, що зміна спрацювала, якого не видно в загальному середньому.",
        "keyPoints": [
          "Когорта — група юзерів за спільною ознакою в часі (частіше місяць реєстрації)",
          "Середні метрики маскують тренди; когорти показують правду",
          "Таблиця: рядки — когорти, стовпці — час, клітинки — % активних",
          "Головний спосіб довести, що зміна покращила утримання"
        ]
      },
      {
        "type": "learn",
        "title": "SQL для PM: SELECT, WHERE",
        "body": "PM не мусить бути інженером даних, але вміння самому дістати число з бази — суперсила. Ти перестаєш чекати аналітика по три дні заради простого питання й починаєш ставити дашбордам правильні запитання, бо розумієш, звідки беруться числа. SQL (Structured Query Language) — це мова запитів до реляційних баз даних, де дані лежать у таблицях зі стовпцями й рядками.\n\nБазовий запит складається з трьох частин. SELECT — які стовпці ти хочеш побачити. FROM — з якої таблиці. WHERE — яку умову мають задовольняти рядки. Читається майже як англійське речення. Наприклад: SELECT name, country FROM users WHERE country = 'UA' — \"дай мені імена й країни з таблиці users, де країна дорівнює UA\".\n\nКілька важливих деталей. Рядкові значення беруться в одинарні лапки ('UA'), числа — ні (age > 18). Умови в WHERE комбінують через AND та OR: WHERE country = 'UA' AND age >= 18. Зірочка SELECT * означає \"всі стовпці\" — зручно для швидкого погляду, але в реальних запитах краще перелічувати потрібні. Оператор LIMIT 10 обмежує вивід десятьма рядками, щоб не витягувати мільйони. Дати порівнюють так само: WHERE created_at >= '2026-01-01'. Це фундамент — навчившись фільтрувати рядки, ти вже можеш відповісти на масу питань на кшталт \"скільки користувачів із Польщі зареєструвалися цього року?\".",
        "example": "Питання: \"Покажи email усіх користувачів із України, які зареєструвалися 2026 року.\"\n\nSELECT email, created_at\nFROM users\nWHERE country = 'UA'\n  AND created_at >= '2026-01-01';\n\nЧитається: візьми стовпці email і created_at з таблиці users, залиш лише рядки, де країна — UA і дата реєстрації від початку 2026 року.",
        "keyPoints": [
          "SELECT — стовпці, FROM — таблиця, WHERE — умова фільтрації рядків",
          "Текст у одинарних лапках, числа без лапок; умови через AND/OR",
          "LIMIT обмежує вивід; уміти дістати число самому — суперсила PM"
        ],
        "resource": {
          "label": "SQLBolt — інтерактивний безкоштовний курс SQL",
          "url": "https://sqlbolt.com/"
        }
      },
      {
        "type": "learn",
        "title": "SQL: GROUP BY, агрегати, JOIN",
        "body": "Фільтрувати рядки — це половина справи. Найчастіше PM потрібне не окреме значення, а зведення: скільки, скільки в середньому, разом. Для цього є агрегатні функції — вони стискають багато рядків в одне число: COUNT (кількість), SUM (сума), AVG (середнє), MIN/MAX (мінімум/максимум).\n\nGROUP BY групує рядки за значенням стовпця й рахує агрегат для кожної групи. Наприклад: SELECT country, COUNT(*) FROM users GROUP BY country — \"скільки користувачів у кожній країні\". Правило: усе, що в SELECT і не є агрегатом, має бути в GROUP BY. Щоб відфільтрувати вже за результатом агрегації, використовують HAVING (не WHERE): HAVING COUNT(*) > 100 залишить лише країни з понад 100 користувачами. Різниця ключова: WHERE фільтрує рядки до групування, HAVING — групи після.\n\nJOIN з'єднує дві таблиці за спільним ключем. Дані зазвичай розкидані: користувачі в одній таблиці, їхні замовлення — в іншій, пов'язані через user_id. INNER JOIN бере лише ті рядки, для яких є збіг в обох таблицях; LEFT JOIN бере всі рядки з лівої таблиці, а з правої підставляє збіг або порожнечу (NULL) — зручно, щоб знайти, наприклад, користувачів без жодного замовлення. Синтаксис: FROM users u JOIN orders o ON u.id = o.user_id. Літери u та o — аліаси, короткі імена таблиць. Опанувавши GROUP BY і JOIN, ти вже можеш самостійно рахувати виручку по країнах чи будувати основу для когортного аналізу.",
        "example": "Питання: \"Яка сумарна виручка по кожній країні, лише для країн із виручкою понад 1000?\"\n\nSELECT u.country, SUM(o.amount) AS revenue\nFROM users u\nJOIN orders o ON u.id = o.user_id\nGROUP BY u.country\nHAVING SUM(o.amount) > 1000\nORDER BY revenue DESC;\n\nЗ'єднали users з orders, згрупували за країною, підсумували суми, лишили великі й відсортували за спаданням.",
        "keyPoints": [
          "Агрегати COUNT/SUM/AVG/MIN/MAX стискають рядки в одне число",
          "GROUP BY групує; WHERE фільтрує до, HAVING — після агрегації",
          "JOIN з'єднує таблиці за ключем; LEFT JOIN зберігає рядки без збігу (NULL)"
        ],
        "resource": {
          "label": "Mode SQL Tutorial — безкоштовний, з практикою",
          "url": "https://mode.com/sql-tutorial/"
        }
      },
      {
        "type": "quiz",
        "title": "Перевірка: WHERE чи HAVING",
        "body": "Ти хочеш побачити країни, у яких понад 50 користувачів. Як правильно?",
        "quiz": {
          "question": "Який запит коректно поверне лише країни з понад 50 користувачами?",
          "options": [
            "SELECT country, COUNT(*) FROM users WHERE COUNT(*) > 50 GROUP BY country",
            "SELECT country, COUNT(*) FROM users GROUP BY country HAVING COUNT(*) > 50",
            "SELECT country, COUNT(*) FROM users GROUP BY country WHERE COUNT(*) > 50",
            "SELECT country, SUM(*) FROM users GROUP BY country HAVING SUM(*) > 50"
          ],
          "answerIndex": 1,
          "explanation": "Фільтр за результатом агрегації робиться через HAVING після GROUP BY. WHERE працює лише з окремими рядками до групування, а COUNT там неприпустимий."
        }
      },
      {
        "type": "build",
        "title": "Практика: напиши 3 запити",
        "body": "Час попрацювати руками. Уяви базу з двома таблицями: users(id, email, country, created_at) та events(id, user_id, event_name, created_at). Твоє завдання — написати три SQL-запити (перевір їх на будь-якому онлайн-пісочнику, наприклад db-fiddle).\n\nЗапит 1 (фільтр): вибери email усіх користувачів з Польщі ('PL'), які зареєструвалися 2026 року. Використай SELECT + WHERE з AND.\n\nЗапит 2 (агрегація): порахуй кількість користувачів у розрізі країн, відсортуй за спаданням кількості. Використай COUNT, GROUP BY, ORDER BY.\n\nЗапит 3 (JOIN + фільтр агрегата): для кожного користувача порахуй, скільки подій із назвою 'lesson_completed' він згенерував, і залиш лише тих, у кого таких подій 5 або більше. Використай JOIN, WHERE (для event_name), GROUP BY та HAVING.\n\nНапиши всі три запити, спробуй виконати їх у пісочниці й переконайся, що синтаксис правильний. Це базовий набір, який покриває 80% реальних питань PM до бази даних.",
        "example": "Скелет для Запиту 3, який треба доповнити:\n\nSELECT u.email, COUNT(*) AS lessons\nFROM users u\nJOIN events e ON u.id = e.user_id\nWHERE e.event_name = 'lesson_completed'\nGROUP BY u.email\nHAVING COUNT(*) >= 5\nORDER BY lessons DESC;",
        "keyPoints": [
          "Перевіряй запити в реальній пісочниці (SQLBolt, db-fiddle), а не в голові",
          "Три патерни — фільтр, агрегація, JOIN+HAVING — покривають більшість задач PM"
        ],
        "resource": {
          "label": "DB Fiddle — безкоштовна онлайн SQL-пісочниця",
          "url": "https://www.db-fiddle.com/"
        }
      },
      {
        "type": "learn",
        "title": "A/B-тести й статзначущість",
        "body": "A/B-тест (спліт-тест) — це контрольований експеримент, де ти випадково ділиш користувачів на дві групи: контрольну (A, бачить старий варіант) і тестову (B, бачить зміну). Потім порівнюєш обрану метрику між групами. Це золотий стандарт доказу, що саме твоя зміна, а не сезон чи випадковість, вплинула на поведінку. Випадковий розподіл — ключ: він гарантує, що групи в середньому однакові за всіма іншими ознаками.\n\nГоловна пастка — випадковість. Якщо в групі B конверсія 11%, а в A — 10%, це справжнє покращення чи просто шум? На це відповідає статистична значущість. p-value — це ймовірність побачити таку різницю (або більшу) чисто випадково, якби зміна насправді нічого не робила. Прийнято поріг p < 0.05: якщо ймовірність випадковості менша за 5%, різницю вважають статистично значущою. Пов'язане поняття — довірчий інтервал: діапазон, у якому з певною певністю лежить справжнє значення.\n\nДля PM критично важливі три речі. Перше: розмір вибірки й тривалість треба порахувати заздалегідь (power analysis) — на маленькій вибірці навіть велика різниця не буде значущою. Друге: не можна \"підглядати\" й зупиняти тест, щойно з'явився бажаний результат (peeking) — це роздуває хибні спрацювання. Третє: статистична значущість — не те саме, що практична: різниця може бути значущою, але настільки мала, що не варта впровадження. І пам'ятай: одночасна перевірка десятків метрик майже гарантує, що якась \"вистрелить\" випадково.",
        "example": "Тестуєш нову кнопку. Контроль A: 2000 юзерів, 200 конверсій (10%). Варіант B: 2000 юзерів, 240 конверсій (12%). Калькулятор значущості дає p = 0.03 (< 0.05) — різниця значуща, ефект +2 п.п. Але якби було по 100 юзерів у групі з тією ж часткою, p був би ~0.6 — вибірка замала, і той самий результат нічого не доводить.",
        "keyPoints": [
          "A/B-тест: випадковий розподіл ізолює ефект саме твоєї зміни",
          "p < 0.05 — умовний поріг, що різниця навряд чи випадкова",
          "Рахуй розмір вибірки заздалегідь; не зупиняй тест через peeking",
          "Статистична значущість ≠ практична важливість ефекту"
        ],
        "resource": {
          "label": "Evan Miller — безкоштовний калькулятор A/B значущості",
          "url": "https://www.evanmiller.org/ab-testing/"
        }
      },
      {
        "type": "quiz",
        "title": "Перевірка: читаємо A/B-тест",
        "body": "Команда провела A/B-тест і отримала p-value = 0.20 при бажаному порозі 0.05. Що це означає?",
        "quiz": {
          "question": "A/B-тест дав різницю в конверсії, але p-value = 0.20. Який висновок правильний?",
          "options": [
            "Зміна точно покращує конверсію, можна викочувати на всіх",
            "Спостережувана різниця може бути випадковим шумом; доказу ефекту немає",
            "Зміна точно шкідлива, її треба відкинути назавжди",
            "Тест зламаний, бо p-value має бути рівно 0.05"
          ],
          "answerIndex": 1,
          "explanation": "p = 0.20 означає 20% імовірності побачити таку різницю чисто випадково — це вище порогу 0.05, тож ефект не доведено. Це не доказ шкоди, а відсутність доказу користі (можливо, замала вибірка)."
        }
      },
      {
        "type": "learn",
        "title": "Закон Гудхарта: як не обманути себе",
        "body": "\"Коли метрика стає ціллю, вона перестає бути хорошою метрикою\" — це закон Гудхарта, найважливіше застереження для будь-кого, хто працює з даними. Щойно ти прив'язуєш до числа премії, KPI чи гордість команди, люди починають оптимізувати саме це число — часто в обхід реальної цінності, заради якої воно існувало.\n\nМеханізм простий і підступний. Метрика — це завжди спрощення, проксі для чогось складнішого й важливішого. \"Час на сайті\" — проксі для \"користувач отримує цінність\". Але якщо зробити метою час на сайті, команда може навмисно ускладнити навігацію, щоб люди довше блукали — число зросте, а досвід погіршиться. Служба підтримки з ціллю \"закривати тікети швидко\" почне закривати їх без вирішення проблеми. Це не злий умисел, а природна реакція на стимул.\n\nЯк PM захищається? По-перше, парні метрики (guardrails): до метрики швидкості додай метрику якості, до метрики зростання — метрику утримання. Пара не дає грати в одні ворота. По-друге, пам'ятай різницю: vanity-метрики (гарні числа, що тішать) vs actionable-метрики (ті, що ведуть до дій). По-третє, регулярно питай \"а що ця метрика НЕ показує?\" і \"як цю метрику можна накрутити, не створивши цінності?\". Якщо відповідь на друге питання проста — метрика вразлива. Здоровий скептицизм до власних чисел — ознака зрілого продуктового мислення, а не слабкості.",
        "example": "Команда контенту отримала KPI \"кількість опублікованих статей на місяць\". Число злетіло — але статті стали короткими й порожніми, бо кількість важливіша за якість. Guardrail врятував би: поряд із кількістю відстежувати \"середній час читання\" та \"частку статей, дочитаних до кінця\". Тоді накрутити обсяг без шкоди для якості вже не вийде.",
        "keyPoints": [
          "Закон Гудхарта: метрика-ціль перестає бути хорошою метрикою",
          "Кожна метрика — лише проксі; її легко накрутити в обхід цінності",
          "Захист — парні guardrail-метрики (швидкість+якість, зростання+утримання)",
          "Питай: як це число можна накрутити, не створивши реальної цінності?"
        ]
      },
      {
        "type": "explain",
        "title": "Поясни: пастка метрики",
        "body": "Ще один teach-back, щоб закріпити найтонше з усього модуля.\n\nСитуація: керівник пропонує поставити команді підтримки єдину ціль — \"зменшити середній час відповіді на звернення\". Звучить розумно: швидка підтримка = щасливі клієнти. Твоє завдання як PM — пояснити, у чому ризик цієї метрики за законом Гудхарта і що ти запропонував би натомість.\n\nСформулюй відповідь на 5-7 речень. Сильна відповідь має: (1) назвати закон Гудхарта й пояснити його простими словами; (2) показати конкретно, як команда може \"накрутити\" час відповіді, не допомігши клієнту насправді (наприклад, надсилати миттєві шаблонні відписки або закривати складні тікети без розв'язання); (3) запропонувати guardrail-метрику в пару — наприклад, CSAT (задоволеність клієнта) чи частку повторних звернень із того самого питання; (4) підсумувати думкою, що жодна метрика поодинці не має ставати єдиною ціллю. Спробуй проговорити це вголос — якщо пояснення звучить переконливо для нетехнічного керівника, ти справді засвоїв ідею.",
        "keyPoints": [
          "Назви механізм (Гудхарт), покажи конкретний спосіб накрутки, дай guardrail",
          "Мета teach-back — пояснити переконливо навіть нетехнічному керівнику"
        ]
      },
      {
        "type": "build",
        "title": "Мілстоун: аналітичний one-pager",
        "body": "Фінальне завдання, що зшиває весь модуль. Обери будь-який знайомий тобі продукт (реальний застосунок, сервіс, або вигаданий стартап) і склади короткий аналітичний one-pager. Це той артефакт, який Technical PM приносить на зустріч зі стейкхолдерами.\n\nСтруктура, яку треба заповнити:\n\n1. North-Star метрика: назви одну головну метрику продукту й поясни в 2-3 реченнях, чому саме вона відображає цінність (перевір: чи вона випереджальна й керована, чи не vanity).\n\n2. AARRR-розклад: під кожен із п'яти етапів (Acquisition, Activation, Retention, Revenue, Referral) впиши одну конкретну метрику, яку ти б відстежував.\n\n3. Activation event: сформулюй гіпотезу — яка рання дія користувача корелює з довгим утриманням, і як би ти це перевірив.\n\n4. Один SQL-запит: напиши запит, який дістав би твою activation rate або retention із бази (використай GROUP BY чи JOIN).\n\n5. Один A/B-тест: опиши гіпотезу, метрику успіху та guardrail-метрику проти закону Гудхарта.\n\n6. Три питання до дашборда: сформулюй три критичні питання, які ти б поставив, побачивши, що north-star метрика зросла на 20%.\n\nПройдися по всіх шести пунктах письмово. Якщо зможеш заповнити їх упевнено — ти засвоїв аналітичне мислення PM на робочому рівні й готовий вести розмову про метрики з будь-якою командою.",
        "example": "Міні-зразок пункту 5 для застосунку подкастів:\nГіпотеза: додавання кнопки \"продовжити з місця зупинки\" на головному екрані підвищить повернення.\nМетрика успіху: 7-day retention нових користувачів.\nGuardrail: середній час прослуховування на сесію (щоб не вийшло, що люди повертаються, але слухають менше).\nПоріг: p < 0.05, вибірка порахована заздалегідь на 2 тижні.",
        "keyPoints": [
          "One-pager із метрик — реальний робочий артефакт Technical PM",
          "Мілстоун зшиває все: NSM, AARRR, активацію, SQL, A/B і guardrails",
          "Якщо впевнено заповнюєш усі 6 пунктів — мислиш метриками на робочому рівні"
        ]
      }
    ]
  },
  {
    "slug": "stakeholders",
    "emoji": "🤝",
    "title": "Стейкхолдери й комунікація",
    "intro": "Пройшовши цей шлях, ти навчишся перетворювати хаос думок founder-а, інженерів і дизайну на спільне «так» — щоб твої рішення просувалися, а не тонули в нескінченних узгодженнях.",
    "steps": [
      {
        "type": "learn",
        "title": "Стейкхолдери й мапа впливу",
        "body": "Стейкхолдер (stakeholder) — це будь-яка людина чи група, чиї інтереси зачіпає твій продукт: founder, інженери, дизайн, sales, підтримка, юристи, користувачі. Головне, що треба зрозуміти про роль PM: у тебе майже немає формальної влади — ти нікого не наймаєш і не звільняєш, не роздаєш накази. Твій єдиний важіль — вплив через довіру й комунікацію. Тому «м'які навички» для PM насправді не такі вже й м'які: це твоя основна робоча компетенція, від якої залежить, чи взагалі рухаються твої рішення.\n\nНе всі стейкхолдери потребують однакової уваги. Щоб не розпорошитися й не вигоріти, розподіляй зусилля за матрицею «влада / інтерес» (power/interest grid). По вертикалі — скільки влади людина має над проєктом (може заблокувати чи схвалити), по горизонталі — наскільки їй цікаво те, що ти робиш. Виходять чотири квадранти зі своїми стратегіями: висока влада + високий інтерес — «керуй тісно» (глибоко залучай, часто синхронізуйся); висока влада + низький інтерес — «тримай задоволеними» (не вантаж деталями, але жодних сюрпризів); низька влада + високий інтерес — «тримай поінформованими» (часто це інженери й підтримка, які хочуть бути в курсі); низька влада + низький інтерес — «моніторь» (мінімум зусиль).\n\nГоловна пастка новачка — вважати мапу разовою. Влада й інтерес змінюються: юрист сьогодні байдужий, а перед релізом з питань приватності стає ключовим блокером. Тому перемальовуй мапу на старті кожної великої ініціативи — п'ять хвилин роботи економлять місяці політичних сюрпризів.",
        "example": "Проєкт: AI-чат підтримки.\n- CTO: висока влада, низький інтерес → «тримай задоволеним», короткі апдейти про ризики й вартість.\n- Head of Support: висока влада + високий інтерес → «керуй тісно», щотижневі 1:1.\n- Інженер бекенду: низька влада, високий інтерес → «тримай поінформованим», клич на демо.\n- Юрист: влади зараз мало, але щодо приватності даних вона висока → познач, щоб залучити перед релізом.",
        "keyPoints": [
          "Стейкхолдер — той, хто щось виграє чи втрачає від твого рішення",
          "У PM немає формальної влади — лише вплив через довіру",
          "Чотири стратегії: керуй тісно / задоволеними / поінформованими / моніторь",
          "Перемальовуй мапу на старті кожної великої ініціативи"
        ],
        "resource": {
          "label": "Mind Tools — Stakeholder Analysis (power/interest grid)",
          "url": "https://www.mindtools.com/aol0rms/stakeholder-analysis/"
        }
      },
      {
        "type": "learn",
        "title": "RACI: хто вирішує, а хто в курсі",
        "body": "Найчастіша причина зірваних проєктів — не технічна складність, а плутанина «хто це взагалі вирішує?». RACI — проста матриця ролей, яка цю плутанину вбиває. Її суть: для кожного важливого рішення чи задачі ти явно розписуєш, хто в ньому бере участь і в якій саме ролі. Абревіатура — це чотири типи участі.\n\nR — Responsible (виконавець): той, хто реально робить роботу; їх може бути кілька. A — Accountable (відповідальний): той ОДИН, хто має фінальне слово й відповідає за результат. Ключове правило: на кожне рішення — рівно один «A». Якщо їх двоє, рішення застрягає в нескінченному пінг-понгу, бо ніхто не може поставити крапку. C — Consulted (консультований): експерти, чию думку питають ДО рішення, у форматі діалогу. I — Informed (поінформований): ті, кому просто повідомляють ПІСЛЯ, без обговорення.\n\nЧому це критично саме для PM? Ти часто виступаєш як «A» за продуктове рішення, тоді як «R» — інженери, які його втілюють. А головне джерело образ у командах — плутанина «C проти I»: людина була впевнена, що з нею порадяться (C), а її поставили перед фактом (I). Тому ролі варто проговорювати вголос, а не тримати в голові. Важлива засторога: RACI не замінює живої розмови. Це не бюрократичний штамп, а карта, яка фіксує домовленість, до якої ти все одно приходиш через діалог із людьми.",
        "example": "Рішення: «Яку LLM-модель узяти для функції?»\n- A (фінальне слово): PM (ти).\n- R (виконавці): ML-інженер, бекенд-інженер.\n- C (до рішення): Security (приватність), Finance (вартість токенів).\n- I (після): Sales, Support, CEO.\nТепер ML-інженер знає, що його спитають, а Sales — що дізнається постфактум, і ніхто не ображений.",
        "keyPoints": [
          "R — робить, A — відповідає (рівно один!), C — питають до, I — кажуть після",
          "Два «A» = зависле рішення; завжди один відповідальний",
          "Плутанина C і I — головне джерело образ; проговорюй ролі явно"
        ]
      },
      {
        "type": "quiz",
        "title": "Квіз: перевір RACI",
        "body": "Швидка перевірка розуміння матриці RACI.",
        "quiz": {
          "question": "Команда сперечається два тижні про назву функції, і рішення не рухається. З погляду RACI, яка найімовірніша причина?",
          "options": [
            "Ніхто не призначений як Responsible",
            "На рішення припадає більше ніж один Accountable",
            "Забули когось додати в Informed",
            "Занадто багато людей у ролі Consulted"
          ],
          "answerIndex": 1,
          "explanation": "Коли фінальне слово формально мають двоє (два «A»), рішення застрягає в пінг-понгу. Правило: рівно один Accountable на рішення."
        }
      },
      {
        "type": "learn",
        "title": "Доносити «чому»",
        "body": "Найпоширеніша помилка початківця-PM — приходити до команди з готовим списком фіч: «треба кнопку X, потім екран Y». Це вбиває і мотивацію, і інженерну творчість: людей перетворюють на руки, що виконують чужі рішення. Сильний PM натомість доносить «чому» (the why) — яку проблему користувача чи бізнесу ми вирішуємо і чому саме зараз.\n\nПричина глибша за мотивацію, вона про масштаб. Коли команда розуміє «чому», вона щодня приймає сотні дрібних рішень без тебе — і приймає їх правильно, бо звіряється зі спільною метою. Коли ж вона знає лише «що», кожна дрібниця повертається до тебе на узгодження, і ти перетворюєшся на вузьке горло, крізь яке мусить пройти геть усе. «Чому» — це те, що масштабує тебе за межі однієї голови.\n\nЯк донести «чому» переконливо? Тримайся структури: проблема → доказ (дані або цитати користувачів) → чому зараз → як виглядає успіх (конкретна метрика). Уникай формулювання «бо CEO попросив»: це наказ, а не «чому», і воно не дає команді орієнтира для власних рішень. Натомість спирайся на реальний біль користувача — одна справжня цитата з інтерв'ю переконує сильніше за десять слайдів із логікою. І пам'ятай: «чому» треба повторювати. Те, що очевидне тобі після місяців занурення, команда чує вперше, тож повертайся до мети на кожному етапі, а не лише на старті.",
        "example": "Слабко: «Робимо кнопку експорту в CSV».\nСильно: «34% користувачів кидають звіти на етапі експорту (Amplitude). У 12 інтерв'ю люди казали: „Хочу просто в Excel і не воювати з системою“. Експорт у один клік → ціль підняти завершення звітів із 66% до 80% за квартал». Тепер інженер сам запропонує кращий формат.",
        "keyPoints": [
          "Продавай проблему й мету, а не готовий список фіч",
          "«Чому» дозволяє команді вирішувати без тебе",
          "Структура: проблема → доказ → чому зараз → як виглядає успіх",
          "Повторюй «чому» на кожному етапі, а не лише на старті"
        ]
      },
      {
        "type": "explain",
        "title": "Teach-back: одне рішення — три аудиторії",
        "body": "Навчання через пояснення. Уяви, що команда вирішила відкласти запуск AI-функції на місяць, щоб додати захист від «галюцинацій» моделі. Твоє завдання — пояснити ЦЕ САМЕ рішення трьом різним стейкхолдерам, підлаштувавши мову.\n\nПоясни:\n1. Founder-у, який поспішає до раунду інвестицій.\n2. Lead-інженеру, який і так перевантажений.\n3. Head of Sales, який уже пообіцяв функцію двом клієнтам.\n\nГарне пояснення адресує головний страх кожного й говорить його мовою. Founder — ризик репутації й довіри інвесторів. Інженер — технічний борг і реалістичність. Sales — що й коли він чесно пообіцяє клієнту. Одне рішення — три «упаковки», але однакова суть. Перевір себе: якщо суть у них різна — ти маніпулюєш, і це вибухне, щойно вони поговорять між собою.",
        "example": "Для Sales: «Розумію, ти пообіцяв це Acme і Globex. Ось що дам: за 4 тижні буде версія, яку не соромно показати живим клієнтам, бо вона не вигадуватиме фактів. Якщо запустимо зараз і модель збреше на демо Acme — втратимо угоду й довіру назавжди. Дам тобі дату, під яку ти спокійно підпишешся?»"
      },
      {
        "type": "learn",
        "title": "Презентувати roadmap і рішення",
        "body": "Roadmap (дорожня карта) — це не список фіч із датами, а історія про те, куди йде продукт і чому. Коли PM презентує roadmap, він продає напрямок і пріоритети, а не обіцяє точні дати. Це принципова різниця: якщо стейкхолдери сприймуть roadmap як контракт із дедлайнами, то кожне природне зміщення строків читатиметься як «зрада» й підриватиме довіру до тебе.\n\nЩоб цього уникнути, показуй roadmap за горизонтами впевненості, а не за календарем. Найпоширеніша модель — «Now / Next / Later» (Зараз / Далі / Потім). «Now» — те, над чим працюємо просто зараз, висока впевненість і чіткий обсяг. «Next» — напрямок ясний, але деталі й строки ще пливуть. «Later» — це гіпотези й ставки без жодних обіцянок. Така структура чесно комунікує, що впевненість спадає з віддаленістю в майбутнє, і захищає тебе від війни за конкретні дати.\n\nОкремо — як презентувати саме рішення. Тримай структуру: контекст і мета → головне повідомлення однією фразою (BLUF, Bottom Line Up Front — суть на самому початку) → обґрунтування → деталі. Люди з владою часто йдуть із зустрічі раніше або відволікаються, тож вони мають почути головне в перші дві хвилини, а не наприкінці. І завжди готуй відповідь на запитання «а що ми свідомо НЕ робимо і чому» — уміння показати відрізане не менш важливе, ніж показати заплановане, бо саме воно доводить, що ти пріоритезував, а не намагаєшся встигнути все одразу.",
        "example": "Замість таблиці з 20 фічами:\nNOW: AI-резюме дзвінків (закриває біль №1 із support-тикетів).\nNEXT: авто-теги тем розмов (економія часу менеджерів).\nLATER: передбачення відтоку (гіпотеза).\nІ окремим слайдом: «Свідомо НЕ робимо голосового бота цього року — ринок не готовий, ризик високий».",
        "keyPoints": [
          "Roadmap продає напрямок і пріоритети, а не точні дати",
          "«Now / Next / Later» чесно передає спадання впевненості",
          "BLUF: головне — в перші дві хвилини",
          "Показуй, що ви свідомо НЕ робите і чому"
        ],
        "resource": {
          "label": "Lenny's Newsletter — публічні пости про product roadmaps",
          "url": "https://www.lennysnewsletter.com/"
        }
      },
      {
        "type": "learn",
        "title": "Писати чіткі специфікації (PRD)",
        "body": "Специфікація, або PRD (Product Requirements Document) — це письмовий документ про те, що ми будуємо, для кого й як зрозуміти, що вийшло вдало. Головна помилка — сприймати PRD як гору вимог, які інженери мусять виконати. Насправді хороший PRD — це інструмент вирівнювання: він робить думки PM видимими, щоб інженери й дизайн могли оскаржити їх ДО того, як написано код, коли зміни ще дешеві.\n\nРобочий каркас документа: (1) Проблема й контекст — навіщо це взагалі. (2) Мета й метрики успіху — конкретне число, а не «покращити UX». (3) Користувацькі сценарії — хто і як цим користуватиметься. (4) Обсяг: що входить і, головне, що НЕ входить (non-goals). (5) Відкриті питання. Деталі інтерфейсу зазвичай живуть у макетах дизайну, а не в PRD.\n\nДва принципи, що відрізняють живий PRD від мертвого. Перший: пиши для оскарження, а не для враження. Саме секції non-goals і open questions запрошують команду сперечатися й ловити діри — без них документ лише імітує ясність. Другий: специфічність замість двозначності. Фраза «має працювати швидко» породжує безкінечну суперечку, а «<2 секунди у 95% випадків» — ні, бо її можна перевірити. Для AI-функцій це критично вдвічі: вимога «має бути розумним» — прямий рецепт провалу; заміни її на вимірювані критерії якості, щоб було ясно, коли функція готова, а коли ще ні.",
        "example": "Погана вимога: «Пошук має бути розумним і швидким».\nЧітка: «Пошук повертає результати <300 мс для 95% запитів. Розуміє друкарські помилки (до 1 помилки на слово). Якщо нічого не знайдено — показує 3 схожі запити. Non-goal: голосовий пошук цієї ітерації».",
        "keyPoints": [
          "PRD — інструмент вирівнювання, а не список вимог для враження",
          "Каркас: проблема → метрики → сценарії → обсяг і non-goals → відкриті питання",
          "Заміни «швидко/розумно» на вимірювані критерії"
        ],
        "resource": {
          "label": "The Mom Test — як чесно питати користувачів",
          "url": "https://www.momtestbook.com/"
        }
      },
      {
        "type": "build",
        "title": "Напиши одностраничний PRD",
        "body": "Застосуй каркас на практиці. Вибери одну AI-функцію (наприклад, «розумні відповіді-заготовки в чаті підтримки») і напиши PRD на ОДНУ сторінку. Обмеження навмисне: воно змушує відсіяти зайве й лишити суть.\n\nОбов'язкові секції:\n1. Проблема й контекст (2-3 речення + один доказ).\n2. Мета та 1-2 метрики (конкретні числа).\n3. 2-3 сценарії «Як [роль], я хочу [дію], щоб [вигода]».\n4. Обсяг: 3-5 пунктів «входить» і 2-3 non-goals.\n5. Для AI — секція «Критерії якості й запобіжники»: що робимо, коли модель не впевнена.\n6. 2-3 відкриті питання.\n\nКритерій успіху: дай прочитати другові, який не в контексті. Якщо він перекаже, ЩО і НАВІЩО ви будуєте, й назве одну річ, яку ви свідомо НЕ робите — PRD працює. Якщо перепитує «а що тут головне?» — перепиши з BLUF.",
        "example": "Секція якості для AI:\n«Заготовка показується лише за впевненості моделі >0.7. Нижче — порожнє поле, без нав'язування. Оператор завжди може відредагувати текст перед відправкою. Non-goal: автовідправка без підтвердження»."
      },
      {
        "type": "learn",
        "title": "Конфлікт пріоритетів і «ні» красиво",
        "body": "Конфлікт пріоритетів — це не збій, а нормальний стан. Sales хоче фічу під конкретного клієнта, підтримка кричить про баги, founder мріє про нову AI-можливість — і всі щиро впевнені, що саме їхнє «терміново». Робота PM тут — не догодити всім (це неможливо), а зробити компроміс видимим і справедливим, щоб рішення сприймалося як чесне, навіть коли комусь відмовили.\n\nПерший інструмент — спільна рамка пріоритезації. «Ні», сказане на основі особистої думки PM, звучить як каприз і породжує опір. «Ні» на основі узгодженого фреймворку — це вже система, а не сваволя. Класичний фреймворк — RICE: Reach × Impact × Confidence ÷ Effort (охоплення × вплив × впевненість ÷ зусилля). Він переносить суперечку з площини «хто голосніше кричить» у площину «які в нас критерії», і це знімає з тебе особисту образу опонента.\n\nДругий інструмент — мова компромісу, а не відмови. Формула красивого «ні»: визнай мотив людини → поясни рішення через спільну мету, а не через себе → запропонуй шлях уперед. Замість «ні, не встигнемо» кажи «так, і ось ціна: якщо беремо Acme цього спринту, фікс багів зсувається на два тижні — що обираємо?». Це не відмова, а повернення рішення тим, хто володіє пріоритетами. Окремий випадок — «ні» вгору, керівнику: його формулюють як запитання, а не як стіну: «Так, зробити можемо. Що з поточного плану знімаємо, щоб звільнити для цього місце?»",
        "example": "Дві термінові вимоги, один спринт. RICE:\n- Фіча Acme: Reach 1, Impact 3, Confidence 80%, Effort 5 → низький бал.\n- Фікс падіння оплат: Reach — усі платні, Impact 3, Confidence 100%, Effort 2 → високий.\nРозмова не про особистості: «За нашим фреймворком фікс оплат вищий. Acme — наступний спринт. Заперечення?»",
        "keyPoints": [
          "Зроби компроміс видимим і справедливим, не намагайся догодити всім",
          "Спільний фреймворк (RICE) → «хто голосніше» стає «які критерії»",
          "«Ні» = визнай мотив → спільна мета → шлях уперед",
          "Говори «так, і ось ціна»; «ні» вгору — «що прибрати натомість?»"
        ]
      },
      {
        "type": "quiz",
        "title": "Квіз: конфлікт пріоритетів",
        "body": "Перевірка розуміння управління пріоритетами.",
        "quiz": {
          "question": "Sales тисне взяти фічу для одного великого клієнта просто зараз. Яка реакція PM найздоровіша?",
          "options": [
            "Погодитися одразу — клієнт великий",
            "Відмовити навідріз: «немає часу»",
            "Показати спільний фреймворк пріоритетів і озвучити ціну: що саме зсунеться, якщо взяти цю фічу зараз",
            "Тихо додати в беклог і сподіватися, що забудуть"
          ],
          "answerIndex": 2,
          "explanation": "PM робить компроміс видимим через спільні критерії й показує вартість «так», повертаючи рішення власникам пріоритетів — не капризує й не прогинається."
        }
      },
      {
        "type": "learn",
        "title": "Апдейти статусу, що будують довіру",
        "body": "Регулярний апдейт статусу — це не бюрократія, а головний інструмент, яким PM без формальної влади щотижня заробляє довіру. Логіка проста: стейкхолдер у невіданні тривожний, а тривожний стейкхолдер починає мікроменеджити, смикати тебе питаннями й скликати позапланові зустрічі. Передбачувані апдейти знімають цю тривогу — людина знає, що вчасно дізнається все важливе, і відпускає контроль.\n\nШаблон сильного апдейту, який сканується за 30 секунд: (1) Світлофор — 🟢 в графіку / 🟡 є ризик / 🔴 заблоковано. (2) Що зроблено з минулого разу. (3) Що далі. (4) Блокери й де потрібна допомога — тут ти явно й адресно просиш дій. (5) Зміни в термінах. Головне правило усього жанру: погані новини повідомляй рано й сам. PM, який ховає червоний статус до самого дедлайну, сподіваючись, що «розсмокчеться», втрачає довіру назавжди — а той, хто попереджає завчасно, навпаки, її зміцнює.\n\nІ останнє — підлаштовуй глибину під аудиторію. CEO потрібен один рядок і колір світлофора; команді — деталі, блокери й контекст. Не розсилай усім однакову довгу «портянку»: перевантажений керівник, отримавши стіну тексту, просто перестане її читати, і твій найважливіший апдейт загубиться саме там, де його мали побачити.",
        "example": "«🟡 AI-резюме дзвінків — є ризик.\nЗроблено: інтеграція з моделлю працює, точність 82% на тестах.\nДалі: тестуємо на реальних дзвінках.\nБлокер: потрібен доступ до анонімізованих записів — @Legal, чи можемо до п'ятниці?\nТерміни: реліз усе ще 15-те, але якщо доступ затягнеться — зсунемось на 22-ге. Попереджаю заздалегідь».",
        "keyPoints": [
          "Апдейти = валюта довіри для PM без формальної влади",
          "Шаблон: світлофор → зроблено → далі → блокери → зміни термінів",
          "Погані новини — рано й від тебе",
          "Різна глибина для CEO і для команди"
        ]
      },
      {
        "type": "explain",
        "title": "Teach-back: апдейт про зрив терміну",
        "body": "Найскладніший апдейт — той, де ти повідомляєш погану новину. Сценарій: ключовий інженер захворів на два тижні, і реліз AI-функції, обіцяний на 15-те, тепер реалістичний лише на 29-те. Про це вже питали founder і sales.\n\nНапиши ОДИН короткий апдейт (5-8 рядків) для стейкхолдерів. Він має: чесно назвати проблему без виправдань і паніки; показати, що ти вже контролюєш ситуацію (є план); дати нову реалістичну дату з запасом; явно сказати, що потрібно від читачів.\n\nПеревір себе: чи назвав новину в першому рядку (BLUF)? Чи звучиш як власник ситуації, а не жертва? Чи є конкретний наступний крок? Чи не пообіцяв знову нереалістичну дату? Гарний апдейт про зрив парадоксально ЗМІЦНЮЄ довіру.",
        "example": "Орієнтир (напиши свій):\n«🔴 Реліз AI-резюме зсувається з 15-го на 29-те. Причина: Олег, єдиний із контекстом по ML, на лікарняному ~2 тижні. Що роблю: Марія підхоплює бекенд сьогодні, критичний шлях переглянутий — 29-те з запасом. Прошу: @Sales, узгодимо, як переказати дату клієнту — маю чернетку, потрібні 15 хвилин сьогодні»."
      },
      {
        "type": "learn",
        "title": "Вирівнювати founder, інженерів, дизайн",
        "body": "Вирівнювання (alignment) — це стан, коли всі однаково розуміють спільну мету й тягнуть в один бік. Складність у тому, що кожна група боїться свого: founder — втратити швидкість і візію, інженери — потонути в технічному боргу, дизайн — випустити зламаний досвід. І кожен має рацію в межах своєї правди, тому «переконати» когось, що його страх дурний, не вийде — страх треба визнати й врахувати.\n\nПерша помилка новачка — намагатися вирівняти всіх в одній великій зустрічі. Великі мітинги — це місце, де конфлікти не розв'язуються, а застигають: люди публічно захищають позиції й не можуть відступити, не втративши обличчя. Досвідчений PM робить pre-wire — проговорює гострі питання з кожним наодинці ДО спільної зустрічі. Приватно люди набагато відвертіші, тож там ти знімаєш найбільші заперечення, а на загальний мітинг виносиш майже готове рішення. Зустріч стає точкою фіксації домовленості, а не полем бою.\n\nДругий інструмент — повертати всіх до спільного «чому» й до користувача. Коли founder і lead-інженер зіштовхуються, найслабший хід PM — стати на чийсь бік, бо тоді ти наживаєш ворога й програєш роль арбітра. Найсильніший — спитати: «Яке рішення краще служить меті X і болю користувача Y?» Це переводить суперечку з «хто з нас правий» на «що об'єктивно правильно», де арбітром стають дані про користувача, а не статус чи гучність голосу.",
        "example": "Founder хоче «вже, сирим», lead-інженер — ще два тижні на стабільність. Замість вибору сторони:\n1. До зустрічі питаєш кожного, що для нього неприйнятний ризик.\n2. Founder боїться пропустити вікно на ринку, інженер — публічного фейлу.\n3. На спільній: «Закритий бета-запуск на 50 лояльних юзерів за тиждень — founder має швидкість і сигнал, інженер не ризикує репутацією на всій базі. Ок?»",
        "keyPoints": [
          "Alignment = всі однаково розуміють мету й тягнуть в один бік",
          "Pre-wire: знімай заперечення 1:1 ДО великої зустрічі",
          "Не ставай на бік — арбітр = спільна мета й дані користувача",
          "Кожна група має свій страх — почни з його визнання"
        ]
      },
      {
        "type": "quiz",
        "title": "Квіз: вирівнювання й pre-wire",
        "body": "Перевірка розуміння, як досягати вирівнювання.",
        "quiz": {
          "question": "Ти очікуєш, що founder і lead-інженер жорстко зіткнуться щодо дати релізу. Як діяти найрозумніше?",
          "options": [
            "Звести їх у велику зустріч — правда народиться в суперечці",
            "Проговорити питання з кожним окремо заздалегідь (pre-wire), зняти заперечення, а на спільну винести майже готове рішення",
            "Вирішити самому й повідомити обом постфактум",
            "Уникати теми, поки конфлікт не вибухне сам"
          ],
          "answerIndex": 1,
          "explanation": "Pre-wire знімає гострі заперечення приватно, тож велика зустріч стає точкою фіксації рішення, а не полем публічної битви."
        }
      },
      {
        "type": "learn",
        "title": "Складні розмови віч-на-віч",
        "body": "Рано чи пізно PM потрапляє в розмову, якої хочеться уникнути: сказати founder-у, що його улюблена ідея не працює за даними; дати інженеру зрозуміти, що його оцінки регулярно зривають плани. Спокуса — відкласти. Але уникання — тиха вбивця: проблема не зникає, а натомість тихо гнояться і стосунки, і сам проєкт.\n\nПерший принцип — починай із фактів, а не з висновків. Мозок співрозмовника захищається від оцінок («ти зриваєш дедлайни» — і людина вже в обороні), але спокійно приймає нейтральне спостереження («останні три спринти розійшлися з планом на 40%»). Тут допомагає розділяти факт і власну «історію»: факт — це те, що об'єктивно сталося; історія — те, що ти сам про це надумав і чому це приписав. Дуже часто історія помилкова, і саме її варто перевірити питанням, а не подавати як вирок.\n\nДругий принцип — створи безпеку. Люди говорять чесно лише тоді, коли не бояться наслідків. Тому почни з підтвердження спільної мети й поваги до людини, постав щире відкрите запитання й реально слухай відповідь, а не чекай своєї черги говорити. Тримай тон цікавості, а не звинувачення. Мета складної розмови — не «перемогти» опонента й довести свою правоту, а вийти з неї зі збереженим стосунком і спільним рішенням, яке обидва готові виконувати.",
        "example": "Замість: «Твої оцінки — обман, через тебе ми зриваємо релізи».\nСпробуй: «Хочу звірити, бо мені важлива і команда, і ти. Останні три спринти реальні строки виходили на 40% довші за оцінки. Не знаю причини — можливо, недооцінюємо тестування, або на тебе вішають зайве. Допоможи зрозуміти, що насправді відбувається?» — і слухаєш.",
        "keyPoints": [
          "Уникання гноїть і проєкт, і стосунки",
          "Починай з фактів-спостережень, не з оцінок",
          "Розділяй факт і свою «історію» — часто вона помилкова",
          "Створи безпеку; став запитання й справді слухай"
        ],
        "resource": {
          "label": "Crucial Conversations — офіційна сторінка книги",
          "url": "https://cruciallearning.com/crucial-conversations-book/"
        }
      },
      {
        "type": "build",
        "title": "Мілстоун: комунікаційний план проєкту",
        "body": "Фінальний рубіж — зведи все вивчене в один робочий артефакт, який реально використовують PM: комунікаційний план для AI-проєкту від старту до релізу.\n\nЗбери документ із частин:\n1. Мапа стейкхолдерів (влада/інтерес) — 5-6 людей зі стратегією.\n2. RACI на 3-4 ключові рішення (модель, приватність даних, дата релізу, обсяг MVP).\n3. Ритм комунікації: хто, що, як часто, яким каналом.\n4. «Чому» проєкту одним абзацом (проблема → доказ → мета).\n5. Зразок тижневого статус-апдейту за шаблоном світлофора.\n6. 2-3 передбачувані конфлікти пріоритетів і твоя заготовлена реакція («так, і ось ціна»).\n\nКритерій успіху: віддай план людині поза контекстом. За 5 хвилин вона має зрозуміти, навіщо проєкт, хто що вирішує, і як інформація тектиме до кожного. Якщо так — ти вже думаєш як Technical PM. Збережи — це готовий зразок для портфоліо.",
        "example": "Фрагмент ритму комунікації:\n| Стейкхолдер | Що отримує | Частота | Канал |\n| CEO | Світлофор + 1 рядок ризику | Тиждень | Email |\n| Команда | Повний статус + блокери | 2× на тиждень | Standup+Slack |\n| Sales | Що й коли обіцяти клієнтам | Тиждень | 15-хв синк |\n| Legal | Приватність даних | Перед релізом | 1:1 |",
        "keyPoints": [
          "Комунікаційний план зводить усі навички в один артефакт",
          "Містить: мапу, RACI, ритм, «чому», зразок апдейту, план на конфлікти",
          "Тест: сторонній розуміє проєкт за 5 хвилин",
          "Збережи як зразок для портфоліо"
        ]
      }
    ]
  },
  {
    "slug": "portfolio",
    "emoji": "💼",
    "title": "Портфоліо AI-PM",
    "intro": "Пройшовши цей шлях, ти зберешся від нуля до готового портфоліо з двох сильних AI-кейсів і резюме, які показують тебе як Technical AI-PM — навіть якщо в тебе ще не було «справжніх» юзерів чи продакшн-запусків.",
    "steps": [
      {
        "type": "learn",
        "title": "Навіщо AI-PM портфоліо",
        "body": "Портфоліо — це не «галерея скріншотів», а доказ того, що ти вмієш думати як продакт-менеджер. Резюме каже, що ти щось робила; портфоліо показує, ЯК ти ухвалювала рішення. Для AI-PM (продакт-менеджера, що будує продукти на основі AI та LLM — великих мовних моделей) це критично, бо роль нова, і рекрутери часто не мають готового «чекліста». Вони шукають докази трьох речей: продуктове мислення (бачиш проблему користувача, а не лише фічу), технічну грамотність (розумієш, як працює модель, промпт, дані, обмеження) і вміння вимірювати вплив.\n\nВажливо зрозуміти головний страх новачка: «в мене немає досвіду в компанії, отже, немає що показати». Це хибно. Наймачі на junior та mid AI-PM ролі рідко очікують кейси зі стартапів-єдинорогів. Вони очікують ЯКІСНЕ мислення на будь-якому матеріалі — навіть на pet-проєкті (особистому проєкті, зробленому для себе, без бюджету й команди). Один сильний pet-кейс, добре розказаний, б'є десять рядків «відповідав за беклог» без жодної деталі.\n\nТому наша мета в цьому модулі — не «зробити багато проєктів», а взяти 2 проєкти й розказати їх так, щоб кожен показував повний цикл продуктового мислення. Якість важливіша за кількість, глибина важливіша за ширину.",
        "example": "Порівняй два рядки. Слабкий: «Зробив чат-бота на GPT для рецептів». Сильний: «Помітив, що люди відкидають рецепти через відсутні інгредієнти; побудував бота, який переписує рецепт під те, що є в холодильнику; на 12 тестерах час до \"готую це\" впав з ~4 хв пошуку до <40 сек». Другий показує проблему, інсайт, рішення й метрику — і саме це купує рекрутер.",
        "keyPoints": [
          "Резюме = що ти робила; портфоліо = ЯК ти думала.",
          "AI-PM оцінюють за трьома осями: продукт, техніка, вплив.",
          "2 глибокі кейси сильніші за 6 поверхових."
        ]
      },
      {
        "type": "learn",
        "title": "Як обрати 2 сильні pet-проєкти",
        "body": "Не кожна ідея стає хорошим кейсом. Сильний AI pet-проєкт для портфоліо має чотири ознаки. Перша — реальна проблема, яку ти сама відчувала або спостерігала (це дає щиру історію й захищає тебе на співбесіді, бо ти знаєш контекст до дрібниць). Друга — AI тут не «прикраса», а суть рішення: якщо ту саму задачу легко закрити звичайною формою чи фільтром, це не демонструє AI-мислення. Третя — вимірюваний результат: має бути хоч якась метрика, навіть на 10 тестерах. Четверта — обмежений обсяг: проєкт, який реально довести до робочого прототипу за 1–2 тижні, а не «застосунок мрії».\n\nОбирай два проєкти так, щоб вони НЕ дублювали одне одного, а показували різні грані. Наприклад: один — про генерацію (бот, що пише чернетки листів), другий — про класифікацію чи пошук (система, що сортує вхідні звернення за темою). Так ти демонструєш ширину: і роботу з відкритим текстом, і роботу зі структурою та точністю. Якщо обидва кейси — «ще один чат із GPT», рекрутер бачить одну навичку двічі.\n\nЩоб не потонути в ідеях, використай простий фільтр: для кожного кандидата запиши одним реченням проблему й одним реченням, чому саме AI. Якщо друге речення звучить силувано («бо зараз усі роблять на AI»), відклади ідею. Обирай ті дві, де відповідь на «чому AI» очевидна: невизначеність, природна мова, велика варіативність вхідних даних — усе, де жорсткі правила ламаються.",
        "example": "Фільтр у дії. Ідея A: «трекер витрат» → чому AI? Слабко, це таблиця з категоріями. Ідея B: «бот, що з фото чека витягує суму й категорію» → чому AI? Сильно: фото різні, текст брудний, правила ламаються — саме кейс для моделі. Ідея C: «асистент, що з голосової нотатки робить структуроване завдання» → сильно й доповнює B іншою модальністю. Обираєш B і C.",
        "keyPoints": [
          "Чотири ознаки: реальна проблема, AI по суті, метрика, обмежений обсяг.",
          "Два кейси мають показувати РІЗНІ грані (генерація vs класифікація/пошук).",
          "Тест «чому саме AI»: якщо звучить силувано — відклади ідею."
        ]
      },
      {
        "type": "quiz",
        "title": "Квіз: вибір проєктів",
        "body": "Ти обираєш pet-проєкти для портфоліо AI-PM. Який набір найсильніший?",
        "quiz": {
          "question": "Які два проєкти найкраще покажуть тебе як AI-PM?",
          "options": [
            "Два чат-боти на GPT: один для рецептів, один для порад із подорожей",
            "Бот, що переписує холодні листи під тон компанії, і система, що класифікує вхідні звернення за терміновістю",
            "Лендінг про AI та презентація трендів генеративного AI",
            "Клон ChatGPT і клон Midjourney, зроблені за туторіалами"
          ],
          "answerIndex": 1,
          "explanation": "Правильний набір показує різні грані (генерація + класифікація) на реальних проблемах; решта або дублюють одну навичку, або не містять власного продуктового рішення."
        }
      },
      {
        "type": "learn",
        "title": "Анатомія кейсу: 6 блоків",
        "body": "Кожен сильний кейс будується за одним каркасом із шести блоків: контекст → проблема → інсайт → рішення → метрика → навчене. Це не бюрократія, а логічна історія, яку рекрутер прочитає за 90 секунд і зрозуміє хід твоєї думки.\n\nКонтекст (2–3 речення) — хто користувач і в якій ситуації; ставить сцену. Проблема — конкретний біль, бажано з «доказом», що він існує (спостереження, цитата, власний досвід). Інсайт — НЕ очевидне спостереження, яке змінює підхід: те, що ти зрозуміла й чого не бачить конкурент; це серце кейсу і найцінніший сигнал продуктового мислення. Рішення — що саме ти побудувала й ЧОМУ такий дизайн (тут для AI-PM важливо назвати ключові технічні рішення: яку модель, який промпт-підхід, як обробляла помилки моделі). Метрика — як ти виміряла, що стало краще. Навчене — чесний висновок: що спрацювало, що ні, що зробила б інакше.\n\nГоловна помилка новачків — писати лише «рішення» («я зробив бота, який…») і пропускати проблему та інсайт. Але саме проблема й інсайт відрізняють PM від виконавця. Рекрутер уже бачив тисячу ботів; він не бачив твого способу помічати біль і формулювати непряму гіпотезу. Тому пиши блоки в цьому порядку й свідомо не перестрибуй одразу до рішення.",
        "example": "Міні-кейс у шести рядках. Контекст: фрилансери-дизайнери ведуть листування з клієнтами. Проблема: відповіді на правки забирають 20–30 хв, бо треба стриматись і бути ввічливим. Інсайт: складність не в тексті, а в емоції — люди гальмують, бо бояться прозвучати різко. Рішення: бот, що бере чернетку «як є» й переписує в спокійний тон, зберігаючи суть; додала кнопку «м'якше/твердіше». Метрика: 9 із 11 тестерів надіслали відповідь удвічі швидше. Навчене: тон вирішує, але сліпе пом'якшення злило користувачів — треба контроль.",
        "keyPoints": [
          "Порядок: контекст → проблема → інсайт → рішення → метрика → навчене.",
          "Інсайт — серце кейсу: неочевидне спостереження, що змінює підхід.",
          "Не перестрибуй одразу до «я зробив бота» — це думка виконавця, не PM."
        ]
      },
      {
        "type": "build",
        "title": "Завдання: каркас першого кейсу",
        "body": "Час зробити руками. Візьми свій перший обраний проєкт і напиши чернетку кейсу строго за шістьма блоками. Мета зараз — не краса, а повнота: кожен із шести блоків має існувати хоча б одним реченням, навіть якщо чорновим.\n\nЗаповни цей шаблон (скопіюй у нотатки й допиши):\n1) КОНТЕКСТ: Хто користувач і в якій ситуації? (2–3 речення)\n2) ПРОБЛЕМА: Який конкретний біль? Який доказ, що він реальний?\n3) ІНСАЙТ: Що НЕОЧЕВИДНЕ ти зрозуміла? Почни фразою «Я помітила, що насправді…»\n4) РІШЕННЯ: Що побудувала + одне технічне рішення (модель / промпт-підхід / як ловила помилки моделі) + чому саме так.\n5) МЕТРИКА: Що і як виміряла? Скільки тестерів? Число до/після.\n6) НАВЧЕНЕ: Що спрацювало, що ні, що змінила б.\n\nПеревір себе двома питаннями. Перше: чи можна прочитати лише блоки «проблема» й «інсайт» і вже захотіти дізнатись рішення? Якщо ні — інсайт слабкий, попрацюй над ним. Друге: чи є в блоці «рішення» бодай одне речення, яке міг би написати лише той, хто розуміє, як працює модель? Якщо ні — додай технічну деталь. Збережи чернетку: вона стане основою оформлення в наступних кроках.",
        "example": "Приклад заповненого блоку ІНСАЙТ (щоб відчути глибину): «Я помітила, що насправді користувачі не хочуть \"кращий переклад\" — вони бояться надіслати повідомлення іноземною й зганьбитись. Тому цінність не в точності, а у впевненості: бот має не лише перекласти, а й позначити, наскільки природно це звучить для носія». Такий інсайт одразу задає інший продукт, ніж «ще один перекладач»."
      },
      {
        "type": "learn",
        "title": "Як показати вплив без «справжніх» юзерів",
        "body": "Найбільший блок для новачка: «в мене не було продакшену й реальних користувачів, звідки метрики?». Відповідь: метрика не мусить бути мільйонною — вона мусить бути чесною й доречною. Рекрутер оцінює не масштаб, а те, чи ти ВЗАГАЛІ мислиш у категоріях вимірювання. П'ять тестерів із чесним числом сильніші за «покращив досвід» без жодної цифри.\n\nДжерела метрик, коли немає аудиторії. Перше — маленький user-тест: дай 5–12 людям (друзі, спільноти, Reddit) виконати задачу й заміряй час, кількість кроків або частку успіху. Друге — до/після на собі чи невеликій вибірці («знайти рецепт: 4 хв → 40 сек»). Третє — якість роботи моделі: візьми 20–30 реальних прикладів, розміть «правильно/неправильно» вручну й порахуй точність (напр., «бот вірно визначив категорію в 24 з 30 = 80%»). Четверте — proxy-метрика (непряма): якщо не можеш виміряти утримання, виміряй те, що його передбачає, — наприклад, частку відповідей, які тестер надіслав без правок.\n\nЗавжди чесно підписуй умови: «на вибірці 11 тестерів», «на 30 розмічених прикладах». Це не слабкість, а зрілість — так роблять і в компаніях на етапі раннього прототипу. І уникай вигаданих великих чисел («+40% конверсії»): досвідчений інтерв'юер одним питанням «як саме ти це виміряла?» зруйнує весь кейс. Мале, але справжнє число захищає тебе; велике вигадане — топить.",
        "example": "Три чесні формулювання впливу. (1) Точність моделі: «На 30 вручну розмічених чеках бот вірно витяг суму в 27 (90%) і категорію в 22 (73%) — категорія стала фокусом наступної ітерації». (2) Час задачі: «Медіанний час оформлення завдання з голосу впав з 55 до 18 сек на 8 тестерах». (3) Proxy: «6 з 9 тестерів надіслали згенерований лист без єдиної правки — сигнал, що тон влучив».",
        "keyPoints": [
          "Метрика має бути чесною й доречною, а не великою.",
          "Джерела без аудиторії: user-тест, до/після, точність на розмічених прикладах, proxy.",
          "Завжди підписуй розмір вибірки; вигадані відсотки топлять кейс на співбесіді."
        ]
      },
      {
        "type": "quiz",
        "title": "Квіз: вплив і метрики",
        "body": "У тебе pet-проєкт без реальної аудиторії. Як найкраще показати вплив?",
        "quiz": {
          "question": "Який спосіб продемонструвати результат найпереконливіший і найчесніший?",
          "options": [
            "Написати «покращив користувацький досвід на 40%», бо звучить солідно",
            "Розмітити 30 реальних прикладів вручну й вказати точність моделі на цій вибірці",
            "Не додавати метрик узагалі — все одно юзерів не було",
            "Вказати кількість годин, які ти витратила на розробку"
          ],
          "answerIndex": 1,
          "explanation": "Точність на 30 розмічених прикладах — чесна, відтворювана метрика; вигадані відсотки руйнуються першим уточнювальним питанням, а витрачені години не вимірюють вплив."
        }
      },
      {
        "type": "explain",
        "title": "Teach-back: проблема vs інсайт",
        "body": "Поясни вголос або напиши, ніби навчаєш подругу, яка плутає «проблему» й «інсайт» у кейсі. Це найчастіша плутанина, і вміння її розвести — сильний сигнал продуктової зрілості.\n\nСформулюй відповідь на три питання: (1) У чому різниця між проблемою й інсайтом одним реченням кожен? (2) Чому кейс, у якому є проблема, але немає інсайту, виглядає як робота виконавця, а не PM? (3) Наведи власний приклад пари «проблема → інсайт» зі свого проєкту.\n\nХороша відповідь охоплює: проблема — це спостережуваний біль користувача («людям довго X»); інсайт — неочевидна причина або переформулювання, яке відкриває краще рішення («насправді біль не в швидкості, а в страху помилитись»). Проблему бачать усі, хто дивиться; інсайт — це те, що побачила саме ти, і саме він виправдовує вибір конкретного дизайну. Кейс без інсайту зводиться до «була задача — я зробив фічу», що не показує мислення. Твій приклад має демонструвати стрибок від очевидного болю до неочевидної причини.",
        "example": "Взірцева пара для орієнтиру. Проблема: «Користувачі кидають онбординг бота на третьому питанні». Поганий «інсайт» (насправді просто повтор проблеми): «їм не подобається довгий онбординг». Справжній інсайт: «Вони кидають не через довжину, а бо не розуміють, ЩО бот зможе зробити краще за них — цінність не показана до першого питання». Звідси інше рішення: показати миттєвий результат до онбордингу."
      },
      {
        "type": "learn",
        "title": "Оформлення в Notion",
        "body": "Notion — стандарт де-факто для PM-портфоліо: безкоштовний, дає чисту публічну сторінку за посиланням і не вимагає верстки. Мета оформлення — щоб рекрутер за 10 секунд зрозумів структуру й захотів читати далі. Візуальна ясність тут працює на тебе так само, як чистий інтерфейс на продукт.\n\nРекомендована структура простору. Головна сторінка = коротке інтро (хто ти, у чому фокус як AI-PM, 1–2 рядки) + галерея з двох карток-кейсів. Кожен кейс — окрема підсторінка з обкладинкою й емодзі-іконкою (Notion це вміє «з коробки»), а всередині — ті самі шість блоків із заголовками H2, щоб очима легко стрибати. Додай угорі кейсу міні-«факт-бокс»: Роль / Строк / Стек (напр., GPT-4o, Python, Streamlit) / Ключова метрика — це дає швидкий контекст перед читанням.\n\nПрактичні правила. Роби короткі абзаци й буліти замість стін тексту. Додавай візуал: скріншот прототипу, схему потоку «запит → модель → відповідь», приклад «до/після» відповіді бота — навіть один скрін різко піднімає довіру. Обов'язково натисни Share → Publish і відкрий доступ за посиланням, потім перевір сторінку в режимі інкогніто (типова помилка — надіслати рекрутеру приватне посилання, яке відкриває лише ти). Тримай мову однією (укр або англ) по всьому портфоліо й став посилання на демо/репозиторій, якщо вони є.",
        "example": "Шаблон факт-боксу вгорі кейсу:\nРоль: сам-собі PM + прототипувальник\nСтрок: 9 днів\nСтек: GPT-4o-mini, промпт-шаблони, Streamlit\nКористувачі в тесті: 11\nКлючова метрика: час задачі −55%\nОдин погляд — і рекрутер має контекст ще до першого абзацу.",
        "keyPoints": [
          "Головна = інтро + галерея з 2 кейсів; кожен кейс — окрема підсторінка.",
          "Факт-бокс угорі (Роль/Строк/Стек/Метрика) + шість блоків під H2.",
          "Publish, потім перевір посилання в інкогніто — приватне посилання = провал."
        ],
        "resource": {
          "label": "Notion — безкоштовний план і публічні сторінки",
          "url": "https://www.notion.com/"
        }
      },
      {
        "type": "learn",
        "title": "Сторітелінг: щоб кейс читали",
        "body": "Факти без історії не запам'ятовуються. Сторітелінг у портфоліо — це не «прикрашання», а спосіб провести читача так, щоб він відчув напругу проблеми до того, як побачить рішення. PM-роль на 50% про комунікацію, тож те, як ти розповідаєш кейс, саме по собі є демонстрацією навички.\n\nБазова структура історії проста: гачок → напруга → поворот → розв'язка. Гачок — жива перша фраза, що кидає в ситуацію («Я вп'яте переписувала одне й те саме повідомлення клієнту, боячись прозвучати грубо»). Напруга — чому наявні рішення не працюють (щоб читач подумав «і справді, а як же тут бути?»). Поворот — твій інсайт, момент «ага». Розв'язка — рішення й результат. Ця дуга природно лягає на твої шість блоків: гачок і напруга живуть у «контексті» й «проблемі», поворот — це «інсайт», розв'язка — «рішення» й «метрика».\n\nКонкретні прийоми. Пиши від першої особи й активним станом («я помітила», «я вирішила») — це показує суб'єктність. Використовуй конкретику замість абстракцій: не «користувачі відчували тертя», а «тестер тричі перепитав, куди тиснути». Наводь одну живу цитату тестера — вона варта абзацу опису. І прибирай усе, що не рухає історію: кожне речення або задає напругу, або її знімає. Якщо речення просто «є» — виріж його.",
        "example": "Два початки того самого кейсу. Прісний: «Метою проєкту було покращення процесу написання ділових листів за допомогою AI». Живий: «\"Я витрачаю більше часу на те, щоб не образити клієнта, ніж на саму відповідь\", — сказала мені знайома дизайнерка. Я почала помічати те саме за собою». Другий одразу створює напругу й особисту ставку — і його дочитають.",
        "keyPoints": [
          "Дуга: гачок → напруга → поворот (інсайт) → розв'язка.",
          "Перша особа + активний стан + конкретика замість абстракцій.",
          "Одна жива цитата тестера сильніша за абзац опису."
        ]
      },
      {
        "type": "explain",
        "title": "Teach-back: перепиши прісний вступ",
        "body": "Практика сторітелінгу через переписування. Візьми навмисно нудний вступ до кейсу й поясни, як ти його оживиш, а тоді напиши покращену версію.\n\nОсь матеріал для роботи (прісний вступ): «Даний проєкт присвячений розробці AI-рішення для оптимізації процесу обробки вхідних звернень користувачів служби підтримки з метою підвищення ефективності».\n\nТвоя відповідь має містити: (1) назвати щонайменше три вади цього вступу (напр., пасивний стан, канцелярит, немає живої людини, немає напруги); (2) пояснити, який прийом сторітелінгу застосуєш до кожної вади; (3) написати переписаний вступ на 2–3 речення, який починається з гачка, вводить конкретну людину або ситуацію й натякає на напругу. Гарна відповідь перетворює абстрактну «оптимізацію процесу» на конкретну сцену з живим болем, зберігаючи професійність (без надмірної драми).",
        "example": "Орієнтир якісного переписування: «На підтримці невеликого сервісу щоранку чекало 200+ звернень в одній купі — і \"поверніть гроші\" лежало поряд із \"як змінити аватар\". Оператор витрачав перші дві години просто на сортування, поки термінові скарги чекали. Я захотіла зняти саме це сортування з людини». Активний стан, конкретна сцена, чітка напруга — і жодного канцеляриту."
      },
      {
        "type": "learn",
        "title": "Що НЕ писати в портфоліо",
        "body": "Іноді додати щось — гірше, ніж пропустити. Є типові антипатерни, які знецінюють навіть хороший кейс, і досвідчений рекрутер зчитує їх миттєво.\n\nПерше — вигадані або «надуті» метрики («+300% залученості») без пояснення, як виміряно: це червоний прапорець, бо одне уточнювальне питання все руйнує. Друге — технічний туман: сторінки про архітектуру й бібліотеки без жодного слова про користувача та рішення; AI-PM ≠ інженер, тебе наймають за продуктові рішення, а техніка — контекст, не головний герой. Третє — приховування невдач: кейс, де «все спрацювало ідеально», виглядає несправжнім; блок «навчене» з чесним промахом додає довіри, а не забирає. Четверте — плагіат туторіалу: якщо ти зібрала проєкт «крок у крок за відео», не видавай чужі рішення за свої продуктові вибори — краще чесно опиши, що саме ТИ додала чи змінила.\n\nОкремо для AI-специфіки. Не замовчуй обмеження й ризики моделі: зрілий AI-PM свідомо пише про галюцинації (коли модель впевнено вигадує факти), приватність даних, хибні спрацювання й те, як ти їх стримувала. Відсутність цієї рефлексії читається як «не розуміє, з чим працює». І не перевантажуй жаргоном заради вигляду: якщо не можеш пояснити «embedding» простими словами, або поясни, або прибери. Правило-фільтр просте: якщо речення не додає ні довіри, ні ясності, ні історії — воно зайве.",
        "example": "Що прибрати vs що лишити. Прибрати: «Використано векторну базу з косинусною близькістю для семантичного retrieval-пайплайну» (жаргон без сенсу для читача). Лишити: «Щоб бот відповідав лише з нашої бази знань і не вигадував, я обмежила його документами компанії й додала фолбек \"не знаю\" — частка вигаданих відповідей на 30 тестах впала з 8 до 1». Друге показує і техніку, і продуктову зрілість.",
        "keyPoints": [
          "Уникай: надутих метрик, технічного туману, «все ідеально», видавання туторіалу за своє.",
          "Обов'язково згадай обмеження AI (галюцинації, приватність) і як ти їх стримувала.",
          "Фільтр: речення без довіри/ясності/історії — виріж."
        ]
      },
      {
        "type": "quiz",
        "title": "Квіз: чого уникати",
        "body": "Ти вичитуєш свій кейс перед публікацією. Що з переліченого — червоний прапорець, який варто прибрати чи переробити?",
        "quiz": {
          "question": "Який елемент найімовірніше зашкодить твоєму AI-PM кейсу?",
          "options": [
            "Чесний блок «навчене» з описом того, що не спрацювало",
            "Метрика «+300% ефективності» без пояснення, як її виміряно",
            "Згадка про те, як ти стримувала галюцинації моделі",
            "Скріншот прототипу з прикладом «до/після»"
          ],
          "answerIndex": 1,
          "explanation": "Надута метрика без методу вимірювання — червоний прапорець, що руйнується першим уточненням; решта варіантів, навпаки, підвищують довіру до кейсу."
        }
      },
      {
        "type": "read",
        "title": "Резюме під AI-PM",
        "body": "Портфоліо показує глибину, а резюме — це квиток, який має пройти повз 6-секундний перегляд і часто ATS (Applicant Tracking System — систему, що автоматично фільтрує резюме за ключовими словами). Тому резюме AI-PM працює інакше, ніж загальне PM-резюме: воно має явно сигналити і продуктовий, і AI-технічний бік.\n\nСтруктура bullet-а — формула XYZ (яку популяризували рекрутери Google і Laszlo Bock у книзі «Work Rules!»): «Досяг X, вимірявши як Y, роблячи Z». Тобто результат + метрика + дія. Для pet-проєктів це працює так само чесно, як для роботи, якщо метрика справжня. Приклад: «Скоротив час оформлення завдання на 67% (55→18 сек на 8 тестерах), побудувавши голосового асистента на GPT-4o-mini з валідацією структури відповіді». Такий рядок за секунду показує вплив, число й технічну суть.\n\nЩо додати саме AI-PM. Окремий блок навичок із двома групами: продуктові (discovery, метрики, A/B-тести, roadmap) і AI-технічні (промпт-інжиніринг, RAG, оцінка якості моделі, робота з даними, базовий SQL/Python). Природно вплітай ключові слова з опису вакансії — LLM, evals, prompt engineering — бо саме їх шукає ATS, але лише ті, що ти справді розумієш. Секцію pet-проєктів оформ як міні-кейси з посиланням на Notion: 1 рядок опису + 1 XYZ-bullet + лінк. І тримай усе на одну сторінку: для junior/mid одна щільна сторінка сильніша за дві розмиті.",
        "example": "Три XYZ-bullet-и для AI-PM:\n• «Підняв точність категоризації чеків до 90% (27/30 розмічених прикладів), спроєктувавши промпт із few-shot прикладами й фолбеком \"невизначено\"».\n• «Знизив частку вигаданих відповідей бота з 8 до 1 на 30 тестах, обмеживши модель базою знань компанії й додавши відповідь \"не знаю\"».\n• «Прискорив написання клієнтських листів удвічі для 9/11 тестерів, зробивши AI-переписувач тону з контролем \"м'якше/твердіше\"».",
        "keyPoints": [
          "Формула bullet-а XYZ: результат X, виміряний як Y, через дію Z.",
          "Дві групи навичок: продуктові + AI-технічні; вплітай ключові слова вакансії чесно.",
          "Одна сторінка; pet-проєкти = міні-кейси з лінком на Notion."
        ],
        "resource": {
          "label": "Google X-Y-Z resume formula (CNBC)",
          "url": "https://www.cnbc.com/2018/08/16/google-recruiters-share-the-3-part-formula-for-stronger-resumes.html"
        }
      },
      {
        "type": "build",
        "title": "Мілстоун: збери повне портфоліо",
        "body": "Фінальне завдання зводить усе докупи. Мета — щоб наприкінці в тебе було реальне, доступне за посиланням портфоліо з двох кейсів і оновлений AI-PM рядок у резюме. Це не вправа «на потім» — це артефакт, який ти вже зможеш надіслати рекрутеру.\n\nПройди чекліст по кроках:\n1) Notion: створи головну сторінку (інтро на 1–2 рядки + галерея на 2 кейси).\n2) Кейс №1: перенеси свій каркас із шести блоків, додай факт-бокс угорі, перепиши вступ за дугою сторітелінгу (гачок → напруга → поворот → розв'язка), встав хоча б один візуал (скрін або схему потоку).\n3) Кейс №2 (інша грань — генерація vs класифікація/пошук): той самий каркас; переконайся, що метрика чесна й підписана розміром вибірки.\n4) Вичитка «що НЕ писати»: пройдися фільтром — прибери надуті метрики, технічний туман, додай чесний рядок про обмеження AI в кожному кейсі.\n5) Publish + перевір посилання в інкогніто.\n6) Резюме: додай секцію pet-проєктів із двома XYZ-bullet-ами й лінком на Notion; онови блок навичок двома групами.\n\nКритерій готовності (definition of done): стороння людина відкриває посилання, за 90 секунд переказує проблему й інсайт кожного кейсу й називає, що ти виміряла. Якщо переказує — портфоліо працює. Попроси одного знайомого зробити саме цей тест і запиши, де він спіткнувся: це твій беклог на наступну ітерацію.",
        "example": "Міні definition of done для самоперевірки:\n☐ 2 кейси, різні грані AI\n☐ у кожному всі 6 блоків + чесна підписана метрика\n☐ вступ кожного починається з гачка, не з канцеляриту\n☐ у кожному є рядок про обмеження AI\n☐ посилання відкривається в інкогніто\n☐ у резюме 2 XYZ-bullet-и + лінк\n☐ знайомий за 90 сек переказав проблему + інсайт + метрику"
      }
    ]
  },
  {
    "slug": "interview",
    "emoji": "🎯",
    "title": "Підготовка до співбесід",
    "intro": "Пройшовши цей шлях, ти зайдеш у будь-яку PM-співбесіду з чіткими рамками для кожного типу питань, готовими історіями про свої проєкти й упевненістю технічного продакта, який вміє говорити про AI.",
    "steps": [
      {
        "type": "learn",
        "title": "Карта раундів PM-співбесіди",
        "body": "Перш ніж готуватися, треба розуміти, з чого складається сам процес. PM-співбесіда — це не одна розмова, а серія раундів, і кожен перевіряє окрему групу навичок. Якщо ти знаєш, який раунд попереду, ти готуєшся точково, а не панікуєш загалом.\n\nТипова структура така. Recruiter screen — коротка розмова з рекрутером про твій досвід, мотивацію й очікування щодо зарплати; тут відсіюють невідповідність базовим вимогам. Product sense (або product design) — тобі дають продукт чи проблему («покращ Google Maps для велосипедистів») і дивляться, як ти думаєш про користувача. Analytical / metrics — кейси на метрики та естімейти («скільки запитів на день обробляє пошук?», «яка метрика для успіху Stories?»). Behavioral — питання про твій минулий досвід («розкажи про конфлікт із розробником»), де перевіряють, як ти працюєш із людьми. Technical / execution — для технічного PM це питання про API, дані, архітектуру, а тепер і про AI/ML. Нарешті, часто є фінальний раунд з керівником або крос-функційними колегами.\n\nЧому це важливо саме тобі як технічному PM з AI-навичками: технічні та AI-раунди — твоя перевага. Багато кандидатів сильні в product sense, але «пливуть», коли треба пояснити, як працює рекомендаційна система чи де в LLM-фічі ризик галюцинацій. Саме там ти можеш виділитися.\n\nГотуйся до кожного раунду окремою рамкою — далі в курсі ми пройдемо їх по черзі.",
        "example": "Приклад розкладу онсайту в продуктовій компанії: 10:00 Product sense (45 хв) → 11:00 Analytical/metrics (45 хв) → 13:00 Behavioral (45 хв) → 14:00 Technical + AI (45 хв) → 15:00 Bar-raiser / hiring manager (30 хв). Помітно, що 4 з 5 раундів — це різні типи мислення, і готувати їх однаково не можна.",
        "keyPoints": [
          "Співбесіда — це серія раундів, кожен перевіряє свою навичку.",
          "П'ять основних типів: recruiter, product sense, metrics, behavioral, technical/AI.",
          "Для технічного PM саме technical/AI-раунд — можливість виділитися.",
          "Дізнайся розклад у рекрутера заздалегідь і готуйся точково."
        ]
      },
      {
        "type": "learn",
        "title": "Рамка CIRCLES для product sense",
        "body": "Product sense — раунд, де тебе просять придумати або покращити продукт. Головна помилка новачків — одразу кидатися пропонувати фічі. Інтерв'юер натомість хоче побачити структуру мислення. Найвідоміша рамка для цього — CIRCLES від Льюїса Ліна.\n\nCIRCLES — це сім кроків. Comprehend the situation — уточни, про що взагалі йдеться, постав уточнювальні питання. Identify the customer — визнач сегменти користувачів і обери один. Report customer needs — сформулюй потреби цього сегмента як «user needs» (наприклад, у форматі «Як користувач, я хочу…»). Cut through prioritization — відбери, які потреби найважливіші, за критеріями (охоплення, біль, відповідність цілям бізнесу). List solutions — згенеруй кілька рішень, а не одне. Evaluate tradeoffs — порівняй рішення за плюсами й мінусами. Summarize recommendation — дай чітку рекомендацію й поясни чому.\n\nЧому PM це критично: рамка показує, що ти не «фонтануєш ідеями», а йдеш від користувача й даних до рішення. Це і є суть роботи продакта. Навіть якщо в тебе слабша ідея, але сильна структура — ти виграєш у кандидата з геніальною ідеєю без структури.\n\nПрактична порада: не проговорюй абревіатуру вголос як робот. Використовуй її як внутрішній чекліст. Витрать перші 5 хвилин на кроки C-I-R (зрозуміти, обрати користувача, потреби) — саме тут кандидати найчастіше зривають раунд, пропускаючи уточнення.",
        "example": "Питання: «Спроєктуй будильник для незрячих людей».\nC: уточнюю — це мобільний застосунок чи фізичний пристрій? Припустимо, застосунок.\nI: сегменти — повністю незрячі, слабозорі, літні люди. Обираю повністю незрячих.\nR: потреби — надійно прокинутися, виставити час без зору, зрозуміти, що будильник вимкнувся.\nC: пріоритет — «виставити час без зору» найбільший біль.\nL: рішення — голосове введення; тактильні кнопки; інтеграція з розумним домом.\nE: голос швидкий, але шумно вночі; тактильні кнопки надійні, але повільніші.\nS: рекомендую голосове введення з тактильним підтвердженням.",
        "keyPoints": [
          "CIRCLES = Comprehend, Identify, Report needs, Cut/prioritize, List, Evaluate, Summarize.",
          "Структура важливіша за геніальність окремої ідеї.",
          "Перші хвилини витрать на уточнення й вибір користувача.",
          "Використовуй рамку як внутрішній чекліст, а не зачитуй вголос."
        ]
      },
      {
        "type": "explain",
        "title": "Teach-back: проведи CIRCLES вголос",
        "body": "Час перевірити, чи справді ти засвоїв рамку — найкращий тест це пояснити її вголос і застосувати.\n\nЗавдання: уяви, що інтерв'юер щойно спитав тебе «Покращ застосунок Spotify для людей, які бігають». Промовляй відповідь вголос (або запиши на диктофон) протягом 5–7 хвилин, свідомо проходячи всі сім кроків CIRCLES. На кожному кроці називай, що ти робиш («Спершу уточню ситуацію…»), а потім роби це.\n\nЩо має бути в хорошій відповіді: (1) щонайменше два-три уточнювальні питання на старті; (2) явний вибір ОДНОГО сегмента бігунів (наприклад, аматори проти марафонців) із поясненням чому; (3) 2–3 сформульовані потреби; (4) явна пріоритизація з критерієм; (5) щонайменше три різні рішення, а не одне; (6) чесний розбір компромісів; (7) фінальна однозначна рекомендація. Якщо якийсь крок ти проскочив — це і є твоя зона росту. Повтори вправу, доки всі сім кроків не звучатимуть природно.",
        "keyPoints": [
          "Проговорювання вголос виявляє прогалини, яких не видно «в голові».",
          "Слідкуй, щоб не проскочити пріоритизацію та компроміси — їх забувають найчастіше.",
          "Записуй себе: почуєш слова-паразити й невпевненість."
        ]
      },
      {
        "type": "learn",
        "title": "Естімейти та guesstimate-кейси",
        "body": "Estimation-питання (їх ще звуть guesstimate або market-sizing) звучать так: «Скільки піц з'їдають у Києві за день?» або «Скільки серверів потрібно YouTube?». Інтерв'юера не цікавить точна цифра — він перевіряє, чи вмієш ти розбивати велику невідому проблему на менші відомі шматки й міркувати вголос логічно.\n\nУніверсальний підхід — top-down декомпозиція. Почни з популяції або загального обсягу, який приблизно знаєш, і послідовно звужуй за допомогою припущень. Кожне припущення озвучуй і обґрунтовуй. Округлюй числа, щоб рахувати легко (10 млн, а не 9,7 млн). Наприкінці зроби sanity-check — чи виглядає результат правдоподібно.\n\nЧому це важливо PM: щодня продакт оцінює речі без повних даних — розмір ринку для нової фічі, навантаження на систему, потенційний дохід. Естімейт-кейс — це модель того самого мислення. Технічному PM це особливо близько: оцінити QPS (queries per second), обсяг сховища чи вартість інференсу LLM — щоденна задача.\n\nГоловна пастка — застрягнути на «правильному» числі. Правильного числа немає. Є прозорий ланцюг припущень, який інтерв'юер може простежити й погодитися. Якщо ти помилився в одному припущенні, але логіка чиста — це успіх.",
        "example": "«Скільки кав продають у Києві за день?»\nНаселення Києва ≈ 3 млн.\nПрипустимо, п'ють каву поза домом ≈ 40% → 1,2 млн людей.\nВ середньому 1,5 чашки на таку людину в день → 1,8 млн чашок.\nДодам гостей міста й офісні кавомашини ≈ +20% → ≈ 2,2 млн чашок/день.\nSanity-check: це ≈ 0,7 чашки на кожного жителя міста в день — правдоподібно.\nВідповідь: близько 2 мільйонів чашок на день, і ось мої припущення.",
        "keyPoints": [
          "Мета — прозорий ланцюг припущень, а не точна цифра.",
          "Іди top-down: від великого відомого числа до дрібніших через припущення.",
          "Округлюй агресивно й рахуй уголос.",
          "Завжди роби sanity-check наприкінці."
        ]
      },
      {
        "type": "quiz",
        "title": "Квіз: що оцінює guesstimate",
        "body": "Перевіримо головну ідею естімейт-кейсів.",
        "quiz": {
          "question": "Ти оцінюєш кількість таксі-поїздок у місті й отримав число, що вдвічі відрізняється від реального. Що найбільше впливає на оцінку інтерв'юера?",
          "options": [
            "Точність фінального числа — воно має збігтися з реальним",
            "Прозорість і логічність твоїх припущень та вміння рахувати вголос",
            "Швидкість — хто швидше назве число, той сильніший",
            "Кількість формул, які ти згадав напам'ять"
          ],
          "answerIndex": 1,
          "explanation": "Estimation перевіряє структуру мислення й обґрунтованість припущень, а не влучання в точну цифру."
        }
      },
      {
        "type": "learn",
        "title": "Метрики-кейси: як обрати правильну метрику",
        "body": "У metrics-раунді тебе просять: «Як ти виміряєш успіх функції X?» або «Метрика впала на 5% — що робиш?». Тут перевіряють, чи вмієш ти зв'язувати продукт із числами й діагностувати проблеми даними.\n\nПочни з мети функції. Метрика має відображати цінність для користувача та бізнесу, а не просто активність. Розрізняй типи: North Star metric — головна метрика цінності (напр., для Spotify — час прослуховування); driver metrics — те, що на неї впливає (кількість створених плейлистів); guardrail metrics — запобіжники, які не мають погіршитися (churn, скарги, латентність). Класична рамка для життєвого циклу — AARRR (піратські метрики): Acquisition, Activation, Retention, Referral, Revenue.\n\nЧому PM це критично: метрика керує командою. Якщо обрати неправильну (напр., «кількість кліків» без урахування задоволеності), команда оптимізуватиме не те й може нашкодити продукту. Хороший PM завжди додає guardrail-метрику, щоб не «виграти битву, програвши війну».\n\nОкремий піджанр — «метрика впала». Тут потрібен структурований діагноз: чи це реальне падіння чи баг у логуванні? Сегментуй (регіон, платформа, версія, новий/старий користувач). Перевір зовнішні чинники (свято, реліз конкурента, сезонність). Звужуй, доки не знайдеш сегмент, що дає падіння.",
        "example": "Функція: Instagram запускає «Нотатки» (короткі статуси).\nNorth Star: кількість активних діалогів, започаткованих через Нотатки на тиждень.\nDriver-метрики: % користувачів, які створили нотатку; середня к-сть відповідей на нотатку.\nGuardrail: не має зрости скарг на спам; час у Direct не має впасти.\nДіагноз падіння: якщо перегляди Нотаток впали, сегментую — виявляю, що падіння лише на Android v12 після релізу → ймовірно баг рендерингу, а не втрата інтересу.",
        "keyPoints": [
          "Метрику виводь із мети функції та цінності, не з активності.",
          "Розрізняй North Star, driver і guardrail-метрики.",
          "Завжди додавай guardrail, щоб не нашкодити оптимізацією.",
          "При падінні метрики — сегментуй, доки не локалізуєш причину."
        ]
      },
      {
        "type": "learn",
        "title": "Поведінкові питання та метод STAR",
        "body": "Behavioral-раунд перевіряє, як ти працюєш насправді: як приймаєш рішення, розв'язуєш конфлікти, справляєшся з провалами. Питання починаються з «Розкажи про час, коли…». Оцінюють не сам факт, а твою роль, мислення й уроки.\n\nСтандарт відповіді — метод STAR. Situation — коротко опиши контекст (де, коли, яка команда). Task — що конкретно було твоїм завданням чи викликом. Action — що саме ЗРОБИВ ТИ (не команда — говори «я»), детально й по кроках. Result — чим завершилося, бажано з числами й уроком. Більшість часу відповіді має припадати на Action — саме там інтерв'юер бачить тебе.\n\nЧому PM це критично: продакт не має прямої влади над командою, тому впливає через комунікацію, аргументи й довіру. Поведінкові історії — єдиний спосіб показати ці «м'які» навички. Слабкий кандидат говорить абстрактно («я завжди слухаю команду»), сильний — дає конкретну історію з деталями.\n\nПідготуйся заздалегідь: склади банк із 6–8 історій, які покривають ключові теми — конфлікт, провал/помилка, лідерство без влади, складне рішення з даними, вплив на стейкхолдерів, розставляння пріоритетів. Одна історія часто закриває кілька тем. Уникай пастки «ми»: якщо після історії незрозуміло, що зробив саме ти — вона не зарахована.",
        "example": "Питання: «Розкажи про конфлікт із інженером».\nS: У проєкті платіжного модуля ми з тимлідом розробки не погоджувалися щодо термінів.\nT: Мені треба було встигнути до релізу, йому — не накопичити технічний борг.\nA: Я зібрав дані про вплив затримки на дохід, провів окрему зустріч, вислухав його ризики, і ми домовились розбити реліз на дві фази — MVP вчасно, рефакторинг у наступному спринті. Я оновив roadmap і узгодив зі стейкхолдерами.\nR: Реліз вийшов вчасно, техборг закрили за 2 тижні, а з тимлідом ми відтоді працюємо злагоджено. Урок: конфлікт часто про різні цілі, і дані знімають емоції.",
        "keyPoints": [
          "STAR = Situation, Task, Action, Result.",
          "Більшість відповіді — це Action, і говори «я», а не «ми».",
          "Result краще з числами й з уроком наприкінці.",
          "Заздалегідь склади банк із 6–8 історій під ключові теми."
        ]
      },
      {
        "type": "build",
        "title": "Побудуй свій банк STAR-історій",
        "body": "Це практичне завдання, яке дасть тобі готову зброю на реальну співбесіду.\n\nКрок 1. Створи таблицю (Google Sheet або Notion) з колонками: Назва історії | Situation | Task | Action | Result | Які теми закриває.\n\nКрок 2. Згадай і запиши щонайменше 6 реальних історій зі свого досвіду (робота, пет-проєкт, навчання, волонтерство — усе рахується). Обов'язково покрий ці теми: конфлікт у команді; провал або помилка, з якої ти навчився; момент лідерства без формальної влади; рішення, ухвалене на основі даних; робота зі складним стейкхолдером; ситуація, де довелося жорстко пріоритизувати.\n\nКрок 3. Для кожної історії пропиши Action у 3–5 конкретних кроках від першої особи й додай у Result хоча б одне число та один урок.\n\nКрок 4. Перечитай і викресли всі «ми», де можна поставити «я». Проговори дві історії вголос — кожна має вкладатися у 2 хвилини. Це твій фундамент: перед кожною співбесідою достатньо буде освіжити таблицю.",
        "keyPoints": [
          "6+ історій у форматі таблиці — багаторазова заготовка на всі співбесіди.",
          "Одна історія може закривати кілька поведінкових тем.",
          "Кожна історія — до 2 хвилин вголос, з числом і уроком у Result."
        ]
      },
      {
        "type": "read",
        "title": "AI та технічні питання для PM",
        "body": "Для технічного PM з AI-навичками цей раунд — момент, де ти обганяєш конкурентів. Питання діляться на дві групи: загальнотехнічні (як влаштований продукт) і AI/ML-специфічні.\n\nЗагальнотехнічні: що таке API і як через нього спілкуються сервіси; різниця між клієнтом і сервером; що таке база даних і чому запити бувають повільними; що таке latency (затримка) і throughput (пропускна здатність); як влаштований кеш. Тобі не треба писати код — треба вміти пояснити ці поняття простими словами й розуміти компроміси, бо саме про компроміси ти домовлятимешся з інженерами.\n\nAI/ML-специфічні (зараз майже завжди питають): у чому різниця між класичним ML і генеративними моделями; що таке LLM і що означає «токен» та «контекстне вікно»; що таке галюцинація й чому вона виникає; що таке промпт-інжиніринг, RAG (retrieval-augmented generation — коли модель підтягує зовнішні дані перед відповіддю) та fine-tuning; як оцінити якість AI-фічі (offline-метрики проти live-тестів, human evaluation); етичні ризики — упередженість даних, приватність, безпека.\n\nЧому PM це критично: продакт AI-фічі має ставити реалістичні очікування. Якщо ти не розумієш, що LLM недетермінований і може вигадувати факти, ти пообіцяєш бізнесу те, чого модель не дає. Хороший AI-PM знає, де провести межу: що автоматизувати моделлю, а де лишити людину в контурі (human-in-the-loop).\n\nЧудове безкоштовне джерело для системного погляду на дизайн AI-продуктів — Google People + AI Guidebook.",
        "example": "Питання: «Ми хочемо чат-бота підтримки на базі LLM. Які ризики й як їх міряти?»\nСильна відповідь: головний ризик — галюцинації (бот впевнено вигадає політику повернень), тому потрібен RAG на реальній базі знань і guardrails. Міряю: точність відповідей через human evaluation на вибірці, % ескалацій до людини, CSAT, і guardrail — частота небезпечних/вигаданих відповідей. Для чутливих випадків лишаю human-in-the-loop. Спершу запускаю на 5% трафіку з логуванням.",
        "resource": {
          "label": "Google People + AI Guidebook",
          "url": "https://pair.withgoogle.com/guidebook/"
        },
        "keyPoints": [
          "Треба не кодити, а пояснювати поняття й компроміси простими словами.",
          "Ключові AI-теми: LLM, токени, галюцинації, RAG, fine-tuning, оцінка якості.",
          "AI-PM ставить реалістичні очікування й знає, де потрібен human-in-the-loop.",
          "Завжди май відповідь про етику: упередженість, приватність, безпека."
        ]
      },
      {
        "type": "quiz",
        "title": "Квіз: RAG проти fine-tuning",
        "body": "Перевіримо розуміння двох способів адаптувати LLM під продукт.",
        "quiz": {
          "question": "Компанії потрібен бот, що відповідає на питання за свіжою внутрішньою документацією, яка оновлюється щотижня. Що зазвичай доречніше як перший крок?",
          "options": [
            "Fine-tuning моделі на всій документації щотижня",
            "RAG — підтягувати релевантні документи в контекст під час запиту",
            "Збільшити кількість параметрів моделі",
            "Нічого, LLM і так усе знає з тренування"
          ],
          "answerIndex": 1,
          "explanation": "RAG дає доступ до актуальних даних без перетренування й простіше оновлюється, тому це логічний перший крок для мінливих знань."
        }
      },
      {
        "type": "explain",
        "title": "Teach-back: поясни галюцинації бабусі",
        "body": "Уміння пояснити технічне поняття простою мовою — головна навичка технічного PM. Ти постійно перекладаєш між інженерами й бізнесом.\n\nЗавдання: уголос (2–3 хвилини) поясни, що таке галюцинація LLM, так, ніби розповідаєш людині без технічного бекграунду — умовній бабусі або нетехнічному керівнику. Потім, у другій частині, поясни те саме інженеру-скептику, який хоче почути, ЧОМУ це відбувається й що з цим робити.\n\nХороше пояснення для нетехнічної людини: використовує аналогію (модель — як дуже начитаний студент, що соромиться сказати «не знаю» і тому іноді впевнено вигадує правдоподібну відповідь), уникає жаргону, наводить приклад із життя. Хороше пояснення для інженера: згадує, що модель прогнозує ймовірний наступний токен, а не звіряється з базою фактів; що причина — у природі генерації та прогалинах у даних; і що пом'якшують це через RAG, guardrails, human-in-the-loop та оцінку якості. Якщо ти можеш вільно перемкнутися між двома реєстрами — ти готовий до технічного раунду.",
        "keyPoints": [
          "Технічний PM постійно перекладає між інженерами й бізнесом.",
          "Для нетехнічних — аналогія без жаргону; для інженерів — механізм і рішення.",
          "Вміння перемикати реєстр пояснення — сильний сигнал на співбесіді."
        ]
      },
      {
        "type": "learn",
        "title": "Як говорити про свої проєкти",
        "body": "Питання «Розкажи про продукт, яким ти пишаєшся» або «Проведи мене через свій проєкт» звучить майже завжди. Це твій шанс показати продуктове мислення на реальному прикладі — але багато кандидатів перетворюють його на нудний перелік функцій.\n\nРозказуй як історію з продуктовою логікою. Проведи інтерв'юера ланцюгом: яку проблему й для кого ти вирішував → чому ця проблема варта уваги (дані, інсайти від користувачів) → які були альтернативи й чому ти обрав це рішення → як ти вимірював успіх → що вийшло й чого навчився. Це фактично та сама продуктова структура, тільки на прикладі з життя. Обов'язково показуй компроміси: що ти вирішив НЕ робити й чому — це ознака зрілості.\n\nЧому PM це критично: інтерв'юер хоче зрозуміти, чи ти ухвалюєш рішення як продакт — від проблеми й даних, а не від «хотілося зробити круту фічу». Навіть невеликий пет-проєкт, розказаний із чіткою логікою «проблема → рішення → результат → урок», б'є масштабний проєкт, розказаний хаотично.\n\nПідготуй 1–2 проєкти заздалегідь і відрепетируй розповідь на 3–4 хвилини. Май напоготові відповіді на дошпилювальні питання: «Чому саме так?», «Що б ти зробив інакше?», «Як зрозумів, що це спрацювало?». Для технічного PM додай технічний вимір: які були обмеження, як ти працював з інженерами, які метрики й дані використовував.",
        "example": "Слабко: «Я зробив застосунок для трекінгу звичок. Там був онбординг, нагадування, статистика, темна тема…» (перелік фіч).\nСильно: «Я помітив, що друзі кидають нові звички за тиждень. Опитав 15 людей — головний біль не мотивація, а забування. Тому я сфокусувався на одному: розумних нагадуваннях у момент звички, а НЕ на красивій статистиці, від якої свідомо відмовився в MVP. Метрика — 7-денне утримання. Вийшло 40% проти ~20% у типових трекерів. Урок: одна добре обрана проблема сильніша за десять фіч.»",
        "keyPoints": [
          "Розповідай як історію: проблема → дані → рішення → результат → урок.",
          "Обов'язково покажи, що вирішив НЕ робити і чому.",
          "Чітка логіка на малому проєкті б'є хаос на великому.",
          "Відрепетируй 3–4 хвилини й підготуй відповіді на «чому саме так?»."
        ]
      },
      {
        "type": "learn",
        "title": "Питання, які ставиш ти",
        "body": "Наприкінці майже кожного раунду звучить «Які у вас питання до нас?». Це не формальність — це частина оцінки. Слабкі або відсутні питання читаються як брак інтересу; продумані показують, що ти мислиш як продакт і серйозно обираєш роботу.\n\nХороші питання діляться на групи. Про продукт і стратегію: «Яка головна продуктова проблема команди на найближчі пів року?», «Як ви вирішуєте, що НЕ робити?». Про роботу PM тут: «Як виглядає взаємодія продакта з інженерами й дизайном?», «Як приймаються рішення, коли думки розходяться?». Про метрики й успіх: «За якими метриками ви оцінюєте успіх цієї команди?», «Як виглядатиме успіх для мене через 6 місяців?». Для технічного/AI-контексту: «Наскільки продакт залучений у технічні й AI-рішення?», «Як ви балансуєте швидкість і якість даних для AI-фіч?».\n\nЧому це важливо: співбесіда двостороння. Ти теж обираєш, і хороші питання допомагають зрозуміти, чи не потрапиш ти в токсичну чи хаотичну команду. Водночас вони — останнє враження про тебе в раунді.\n\nПастки, яких уникай: питання, відповідь на які є на сайті («а чим ви займаєтесь?»); питання лише про зарплату й відпустку на ранніх раундах; жодного питання взагалі. Підготуй 5–7 питань заздалегідь, бо частину з них інтерв'юер може закрити сам по ходу розмови — тобі потрібен запас.",
        "example": "Сильний набір під фінал з майбутнім керівником:\n1. «Яка найбільша продуктова ставка команди цього року і що може завадити?»\n2. «Як ви вирішуєте, які ідеї відкинути?»\n3. «Наскільки PM тут занурений у технічні та AI-рішення?»\n4. «Як виглядатиме мій успіх через 3 і 6 місяців?»\n5. «Що вам самому подобається й що складно в роботі тут?»\nОстаннє питання по-людськи розкриває реальну культуру краще за будь-який лендинг.",
        "keyPoints": [
          "Твої питання — частина оцінки, а не формальність.",
          "Пиши про стратегію, роль PM, метрики й культуру, не лише про пільги.",
          "Не питай того, що є на сайті; май запас 5–7 питань.",
          "Співбесіда двостороння — ти теж перевіряєш команду."
        ]
      },
      {
        "type": "quiz",
        "title": "Квіз: поведінкова відповідь",
        "body": "Перевіримо, чи вловив ти суть методу STAR.",
        "quiz": {
          "question": "На питання про конфлікт кандидат каже: «Ми в команді завжди все обговорюємо й знаходимо компроміс». Що головне не так із цією відповіддю?",
          "options": [
            "Занадто довга і детальна",
            "Це загальне твердження без конкретної ситуації та без ролі саме кандидата ('я')",
            "Не згадана назва компанії",
            "Немає технічних термінів"
          ],
          "answerIndex": 1,
          "explanation": "STAR вимагає конкретної історії з чіткими діями від першої особи; абстрактне 'ми завжди' не показує, що зробив саме кандидат."
        }
      },
      {
        "type": "read",
        "title": "Mock-інтерв'ю та системна підготовка",
        "body": "Знати рамки — недостатньо; на співбесіді вирішує здатність застосувати їх під тиском і вголос. Тому центральний інструмент підготовки — mock-інтерв'ю (пробні співбесіди).\n\nЯк організувати. Знайди партнера (колегу, спільноту PM, друга з іншим бекграундом) і по черзі грайте інтерв'юера й кандидата. Інтерв'юер ставить питання, мовчки слухає, робить нотатки й дає структурований фідбек: що було сильно, де загубилася структура, де бракувало «я», де не було sanity-check. Записуй сесії на відео — ти помітиш слова-паразити, невпевненість, забігання наперед. Якщо партнера немає, проговорюй відповіді вголос сам собі й записуй; це все одно набагато краще, ніж «продумувати в голові».\n\nСистема підготовки на 3–4 тижні. Тиждень 1: закрий теорію по кожному типу раунду (цей курс) і збери банк STAR-історій. Тиждень 2: практикуй product sense і естімейти по одному кейсу на день вголос. Тиждень 3: метрики й технічні/AI-питання; проговорюй пояснення понять простими словами. Тиждень 4: щоденні повноцінні mock-и, що імітують реальний розклад раундів.\n\nЧому це критично: PM-співбесіда перевіряє мислення в реальному часі. Кандидат, який 10 разів проговорив кейси вголос, звучить спокійно й структуровано; той, хто лише читав теорію, «плаває», щойно вмикається стрес. Різницю чути з перших хвилин.\n\nДля тренування навички user-інтерв'ю та розмов з користувачами (корисно і для product sense) — класика жанру The Mom Test.",
        "example": "Приклад одного mock-раунду (product sense, 45 хв):\n0–5 хв: інтерв'юер дає кейс, кандидат уточнює.\n5–35 хв: кандидат проходить CIRCLES вголос, інтерв'юер лише зрідка штовхає питаннями.\n35–40 хв: дошпилювальні питання («а якби бюджет був нульовий?»).\n40–45 хв: фідбек — «структура була чітка, але ти проскочив пріоритизацію й не назвав guardrail-метрику». Кандидат записує це як фокус на завтра.",
        "resource": {
          "label": "The Mom Test (офіційний сайт книги)",
          "url": "http://momtestbook.com/"
        },
        "keyPoints": [
          "Mock-інтерв'ю вголос — головний інструмент; теорії самої недостатньо.",
          "Записуй сесії й збирай структурований фідбек після кожної.",
          "Готуйся системно 3–4 тижні: теорія → кейси → метрики/AI → повні mock-и.",
          "Проговорена вголос відповідь звучить спокійніше за 'продуману в голові'."
        ]
      },
      {
        "type": "build",
        "title": "Мілстоун: повний mock-цикл",
        "body": "Фінальне завдання, що зв'язує весь курс докупи. Ти проведеш собі повноцінний імітований онсайт і оціниш готовність.\n\nКрок 1. Признач конкретну дату й час, як справжню співбесіду. Підготуй список: 1 product-sense кейс, 1 естімейт, 1 метрик-кейс, 3 поведінкові питання, 3 технічні/AI-питання. Візьми їх зі свого досвіду або згенеруй заздалегідь, щоб не бачити відповідей наперед.\n\nКрок 2. Проведи всі раунди поспіль (≈2 години), як на реальному онсайті, вголос і на запис. Product sense — через CIRCLES; естімейт — через top-down декомпозицію з sanity-check; метрики — з North Star і guardrail; поведінкові — через STAR з банку історій; технічні — з поясненням понять простими словами.\n\nКрок 3. У кінці кожного раунду постав «свої» питання інтерв'юеру зі свого списку 5–7.\n\nКрок 4. Передивись запис і заповни чек-лист: Чи уточнював на старті? Чи була структура? Чи говорив «я» в поведінкових? Чи робив sanity-check? Чи додавав guardrail-метрику? Чи звучав спокійно? Виділи 2–3 найслабші місця — це твій план на наступний тиждень практики. Пройшовши цей цикл кілька разів, ти зайдеш на реальну співбесіду підготовленим не в теорії, а в реальному часі.",
        "keyPoints": [
          "Мілстоун — повний імітований онсайт із усіх типів раундів поспіль.",
          "Застосуй усі рамки курсу: CIRCLES, top-down, North Star+guardrail, STAR.",
          "Оціни себе за чек-листом і виділи 2–3 зони росту.",
          "Повторюй цикл — готовність будується практикою, а не читанням."
        ]
      }
    ]
  }
];

// Book course (18 modules from "Managing Information Technology"), spliced at build time.
const BOOK_PATH = []; /*__BOOK_PATH__*/

const pathStepId = (slug, i) => `${slug}:${i}`;

function CareerPath({ modules = PM_PATH, heading = "Шлях: технічний PM з AI", progress, onDone }) {
  const [open, setOpen] = useState(null); // { modIdx, stepIdx }
  const totalSteps = modules.reduce((s, m) => s + m.steps.length, 0);
  const doneCount = modules.reduce((s, m) => s + m.steps.filter((_, i) => progress[pathStepId(m.slug, i)]).length, 0);
  const openMod = open ? modules[open.modIdx] : null;
  const openStep = openMod ? openMod.steps[open.stepIdx] : null;

  if (!modules.length) return <div className="rounded-2xl bg-white py-10 text-center text-sm text-slate-400 ring-1 ring-rose-50">Курс готується… (контент генерується — зайди трохи згодом)</div>;

  return (
    <div className="space-y-5">
      <div className="rounded-3xl bg-gradient-to-br from-rose-500 to-pink-500 p-5 text-white shadow-sm">
        <div className="flex items-center gap-2 text-sm font-semibold text-white/90"><GraduationCap className="h-4 w-4" /> {heading}</div>
        <div className="mt-1 text-2xl font-extrabold tabular-nums">{doneCount} / {totalSteps} кроків</div>
        <div className="mt-2 h-2.5 overflow-hidden rounded-full bg-white/25"><div className="h-full rounded-full bg-white transition-all" style={{ width: `${totalSteps ? (doneCount / totalSteps) * 100 : 0}%` }} /></div>
        <div className="mt-1.5 text-xs leading-relaxed text-white/80">Маленькі кроки: вивчи → поясни своїми словами → збери → пройди тест. Роби по одному на день. 💪</div>
      </div>

      {modules.map((mod, mi) => {
        const mDone = mod.steps.filter((_, i) => progress[pathStepId(mod.slug, i)]).length;
        const modDone = mDone === mod.steps.length;
        return (
          <div key={mod.slug} className="rounded-3xl bg-white p-4 shadow-sm ring-1 ring-rose-50">
            <div className="mb-3 flex items-center gap-3">
              <span className="grid h-11 w-11 shrink-0 place-items-center rounded-2xl bg-rose-100 text-2xl">{mod.emoji}</span>
              <div className="min-w-0 flex-1">
                <div className="flex items-center gap-2"><span className="font-bold text-slate-800">{mod.title}</span>{modDone && <Trophy className="h-4 w-4 shrink-0 text-amber-400" />}</div>
                <div className="text-xs leading-snug text-slate-400">{mod.intro}</div>
              </div>
              <span className="shrink-0 text-xs font-bold tabular-nums text-rose-500">{mDone}/{mod.steps.length}</span>
            </div>
            <div className="relative">
              <span className="pointer-events-none absolute bottom-4 left-[18px] top-4 w-0.5 bg-slate-100" aria-hidden />
              <div className="relative space-y-1">
                {mod.steps.map((st, si) => {
                  const sid = pathStepId(mod.slug, si);
                  const isDone = !!progress[sid];
                  const prevDone = si === 0 || !!progress[pathStepId(mod.slug, si - 1)];
                  const locked = !isDone && !prevDone;
                  const meta = PATH_STEP_META[st.type] || PATH_STEP_META.learn;
                  const NodeIcon = isDone ? Check : locked ? Lock : meta.icon;
                  return (
                    <div key={sid} className="flex items-center gap-3">
                      <button disabled={locked} onClick={() => setOpen({ modIdx: mi, stepIdx: si })} className={`relative z-10 grid h-9 w-9 shrink-0 place-items-center rounded-full text-white shadow-sm transition ${isDone ? "bg-emerald-500" : locked ? "bg-slate-200 text-slate-400" : `${meta.ring} hover:scale-105`}`}><NodeIcon className="h-4 w-4" /></button>
                      <button disabled={locked} onClick={() => setOpen({ modIdx: mi, stepIdx: si })} className="min-w-0 flex-1 py-1.5 text-left disabled:cursor-default">
                        <div className={`text-sm font-semibold ${locked ? "text-slate-300" : isDone ? "text-slate-400" : "text-slate-700"}`}>{st.title}</div>
                        <div className="text-[10px] font-bold uppercase tracking-wide text-slate-300">{meta.label}</div>
                      </button>
                      {isDone && <Check className="mr-1 h-4 w-4 shrink-0 text-emerald-400" />}
                    </div>
                  );
                })}
              </div>
            </div>
            {modDone && <div className="mt-3 rounded-2xl bg-amber-50 px-3 py-2 text-center text-sm font-semibold text-amber-700">🎉 Модуль пройдено!</div>}
          </div>
        );
      })}
      <p className="px-2 pb-2 text-center text-xs leading-relaxed text-slate-400">Це стартова база. Проходь у своєму темпі, повертайся до уроків будь-коли. Прогрес зберігається і синхронізується.</p>

      {openStep && <PathLessonModal step={openStep} done={!!progress[pathStepId(openMod.slug, open.stepIdx)]} onClose={() => setOpen(null)} onDone={(v) => { onDone(pathStepId(openMod.slug, open.stepIdx), v); setOpen(null); }} />}
    </div>
  );
}

function PathLessonModal({ step, done, onClose, onDone }) {
  const meta = PATH_STEP_META[step.type] || PATH_STEP_META.learn;
  const StepIcon = meta.icon;
  const [picked, setPicked] = useState(null);
  const [note, setNote] = useState("");
  const isQuiz = step.type === "quiz" && step.quiz;
  const correct = isQuiz && picked === step.quiz.answerIndex;
  const canFinish = done || !isQuiz || correct;
  return (
    <div className="fixed inset-0 z-50 flex items-end justify-center bg-slate-900/40 p-0 backdrop-blur-sm sm:items-center sm:p-4" onClick={onClose}>
      <div className="max-h-[92vh] w-full max-w-md overflow-y-auto rounded-t-3xl bg-white p-5 shadow-xl sm:rounded-3xl" onClick={(e) => e.stopPropagation()}>
        <div className="mb-3 flex items-center gap-2">
          <span className={`grid h-9 w-9 shrink-0 place-items-center rounded-xl text-white ${meta.ring}`}><StepIcon className="h-4 w-4" /></span>
          <div className="min-w-0 flex-1"><div className="text-[10px] font-bold uppercase tracking-wide text-slate-400">{meta.label}</div><h3 className="text-lg font-bold leading-tight text-slate-900">{step.title}</h3></div>
          <button onClick={onClose} className="shrink-0 rounded-md p-1 text-slate-400 hover:bg-slate-100"><X className="h-5 w-5" /></button>
        </div>

        {step.body && <p className="whitespace-pre-wrap text-sm leading-relaxed text-slate-600">{step.body}</p>}

        {step.example && (
          <div className="mt-3 rounded-xl bg-amber-50 px-3 py-2.5 text-sm leading-relaxed text-amber-900 ring-1 ring-amber-100">
            <div className="mb-0.5 text-[11px] font-bold uppercase tracking-wide text-amber-600">Приклад</div>
            {step.example}
          </div>
        )}

        {step.keyPoints?.length > 0 && (
          <div className="mt-3 rounded-xl bg-slate-50 px-3 py-2.5">
            <div className="mb-1 text-[11px] font-bold uppercase tracking-wide text-slate-400">Запам'ятати</div>
            <ul className="space-y-1">
              {step.keyPoints.map((k, i) => <li key={i} className="flex gap-2 text-sm text-slate-600"><span className="mt-0.5 shrink-0 text-rose-400">•</span><span>{k}</span></li>)}
            </ul>
          </div>
        )}

        {step.resource?.url && (
          <a href={step.resource.url} target="_blank" rel="noreferrer" className="mt-3 flex items-center gap-2 rounded-xl bg-rose-50 px-3 py-2 text-sm font-semibold text-rose-600 hover:bg-rose-100"><BookOpen className="h-4 w-4 shrink-0" /> <span className="min-w-0 flex-1 truncate">{step.resource.label || "Джерело"}</span> <ArrowRight className="h-4 w-4 shrink-0" /></a>
        )}

        {step.type === "explain" && (
          <textarea value={note} onChange={(e) => setNote(e.target.value)} rows={3} placeholder="Поясни своїми словами (для себе — не зберігається)…" className="mt-3 w-full rounded-xl border border-slate-300 px-3 py-2 text-sm focus:border-rose-400 focus:outline-none" />
        )}

        {isQuiz && (
          <div className="mt-3 space-y-2">
            <p className="text-sm font-semibold text-slate-700">{step.quiz.question}</p>
            {step.quiz.options.map((opt, i) => {
              const reveal = picked != null;
              const isRight = i === step.quiz.answerIndex;
              const chosen = picked === i;
              const cls = reveal ? (isRight ? "border-emerald-400 bg-emerald-50 text-emerald-800" : chosen ? "border-red-300 bg-red-50 text-red-700" : "border-slate-200 text-slate-400") : "border-slate-200 text-slate-700 hover:border-rose-300";
              return <button key={i} disabled={reveal && correct} onClick={() => setPicked(i)} className={`flex w-full items-center gap-2 rounded-xl border px-3 py-2 text-left text-sm transition ${cls}`}><span className="grid h-5 w-5 shrink-0 place-items-center rounded-full border text-[11px] font-bold">{String.fromCharCode(65 + i)}</span><span className="min-w-0 flex-1">{opt}</span>{reveal && isRight && <Check className="h-4 w-4 shrink-0" />}</button>;
            })}
            {picked != null && <div className={`rounded-xl px-3 py-2 text-sm ${correct ? "bg-emerald-50 text-emerald-700" : "bg-amber-50 text-amber-700"}`}>{correct ? "✓ " : "Майже. "}{step.quiz.explanation}</div>}
          </div>
        )}

        <button disabled={!canFinish} onClick={() => onDone(true)} className="mt-4 w-full rounded-2xl bg-rose-500 py-3 font-bold text-white transition hover:bg-rose-600 disabled:bg-slate-200 disabled:text-slate-400">{done ? "Пройдено ✓ · закрити" : isQuiz && !correct ? "Обери правильну відповідь" : "Готово ✓"}</button>
        {done && <button onClick={() => onDone(false)} className="mt-2 w-full text-center text-xs font-semibold text-slate-400">Скинути крок</button>}
      </div>
    </div>
  );
}

function ToolkitSection({ name, onRename }) {
  const [loading, setLoading] = useState(true);
  const [tool, setTool] = useState("hub"); // hub | anxiety | chores
  const [settings, setSettings] = useState({ name: "Toolkit" });
  const [favorites, setFavorites] = useState([]);
  const [tried, setTried] = useState([]);
  const [weekFocus, setWeekFocus] = useState(null);
  const [members, setMembers] = useState(null);
  const [list, setList] = useState([]);
  const [ratings, setRatings] = useState({});
  const [assignments, setAssignments] = useState({});
  const [meta, setMeta] = useState({ step: 1, useScores: false, finished: false });
  const [renaming, setRenaming] = useState(false);
  const [nameDraft, setNameDraft] = useState(name);

  const reload = useCallback(async () => {
    const d = await loadToolkitData();
    setSettings(d.settings); setFavorites(d.favorites); setTried(d.tried); setWeekFocus(d.weekFocus);
    setMembers(d.members || [{ id: ruid("m"), name: "Людина 1" }, { id: ruid("m"), name: "Людина 2" }]);
    setList(d.list); setRatings(d.ratings); setAssignments(d.assignments); setMeta(d.meta);
    setLoading(false);
  }, []);
  useEffect(() => {
    reload();
    const onReset = () => { setFavorites([]); setTried([]); setWeekFocus(null); setMembers([{ id: ruid("m"), name: "Людина 1" }, { id: ruid("m"), name: "Людина 2" }]); setList([]); setRatings({}); setAssignments({}); setMeta({ step: 1, useScores: false, finished: false }); setTool("hub"); };
    window.addEventListener("toolkit-reset", onReset);
    return () => window.removeEventListener("toolkit-reset", onReset);
  }, [reload]);

  const saveSettings = useCallback(async (patch) => { const n = { ...settings, ...patch }; setSettings(n); await store.set(TKEYS.settings, n); }, [settings]);
  const saveFav = useCallback(async (n) => { setFavorites(n); await store.set(TKEYS.anxFav, n); }, []);
  const saveTried = useCallback(async (n) => { setTried(n); await store.set(TKEYS.anxTried, n); }, []);
  const saveWeek = useCallback(async (id) => { setWeekFocus(id); await store.set(TKEYS.anxWeek, id); }, []);
  const saveMembers = useCallback(async (n) => { setMembers(n); await store.set(TKEYS.members, n); }, []);
  const saveList = useCallback(async (n) => { setList(n); await store.set(TKEYS.list, n); }, []);
  const saveRatings = useCallback(async (n) => { setRatings(n); await store.set(TKEYS.ratings, n); }, []);
  const saveAssignments = useCallback(async (n) => { setAssignments(n); await store.set(TKEYS.assignments, n); }, []);
  const saveMeta = useCallback(async (patch) => { const n = { ...meta, ...patch }; setMeta(n); await store.set(TKEYS.meta, n); }, [meta]);

  if (loading) return <div className="flex flex-1 items-center justify-center text-slate-400"><div className="flex flex-col items-center gap-3"><Wrench className="h-8 w-8 animate-pulse text-rose-400" /><span className="text-sm">Завантаження…</span></div></div>;

  return (
    <div className="min-h-screen flex-1 bg-gradient-to-b from-slate-50 to-white">
      <header className="sticky top-0 z-20 border-b border-slate-200 bg-white/85 backdrop-blur">
        <div className="mx-auto flex h-14 w-full max-w-3xl items-center gap-2 px-4">
          {renaming ? (
            <input autoFocus value={nameDraft} onChange={(e) => setNameDraft(e.target.value)} onBlur={() => { onRename(nameDraft); setRenaming(false); }} onKeyDown={(e) => { if (e.key === "Enter") { onRename(nameDraft); setRenaming(false); } }} className="mr-auto w-32 rounded-lg border border-rose-200 px-2 py-1 text-base font-semibold focus:outline-none" />
          ) : (
            <button onClick={() => { setNameDraft(name); setRenaming(true); }} className="mr-auto text-base font-semibold text-slate-900">{name} <Pencil className="ml-0.5 inline h-3.5 w-3.5 text-slate-300" /></button>
          )}
          {tool !== "hub" && <button onClick={() => setTool("hub")} className="inline-flex items-center gap-1.5 rounded-lg px-3 py-1.5 text-sm font-medium text-slate-500 hover:bg-slate-100"><ChevronLeft className="h-4 w-4" /> Інструменти</button>}
        </div>
      </header>

      <main className="mx-auto w-full max-w-3xl px-4 py-6">
        {tool === "hub" && (
          <div className="space-y-4">
            <h1 className="text-2xl font-extrabold text-slate-900">Інструменти</h1>
            <p className="text-sm text-slate-500">Практичні помічники для щоденного життя.</p>
            <button onClick={() => setTool("chores")} className="flex w-full items-center gap-4 rounded-2xl border border-slate-200 bg-white p-5 text-left shadow-sm transition hover:border-rose-200 hover:shadow-md">
              <span className="grid h-14 w-14 shrink-0 place-items-center rounded-2xl bg-rose-100 text-2xl">🧹</span>
              <span className="min-w-0 flex-1"><span className="block text-lg font-bold text-slate-800">Chore Splitter</span><span className="block text-sm text-slate-400">7-крокова система, щоб чесно поділити хатні справи — по складності, а не по кількості.</span></span>
              <ChevronRight className="h-5 w-5 text-slate-300" />
            </button>
          </div>
        )}
        {tool === "chores" && <ChoreSplitter members={members} list={list} ratings={ratings} assignments={assignments} meta={meta} saveMembers={saveMembers} saveList={saveList} saveRatings={saveRatings} saveAssignments={saveAssignments} saveMeta={saveMeta} />}
      </main>
    </div>
  );
}

/* ---------- Anxiety toolkit (library) ---------- */
function AnxietyToolkit({ favorites, tried, weekFocus, settings, onFav, onTried, onWeek, onDismissNote }) {
  const toggle = (arr, id, fn) => fn(arr.includes(id) ? arr.filter((x) => x !== id) : [...arr, id]);
  const focus = ANX_TECHNIQUES.find((t) => t.id === weekFocus);
  const ordered = [...ANX_TECHNIQUES].sort((a, b) => (favorites.includes(b.id) ? 1 : 0) - (favorites.includes(a.id) ? 1 : 0));
  return (
    <div className="space-y-4">
      <div>
        <h1 className="text-2xl font-extrabold text-slate-900">Anxiety Toolkit</h1>
        <p className="text-sm text-slate-500">Спокійні техніки, до яких можна повертатись. Це бібліотека — інтерактивні версії живуть у вкладці Calm.</p>
      </div>

      {focus && (
        <div className="flex items-center gap-3 rounded-2xl bg-gradient-to-r from-teal-400 to-sky-400 p-4 text-white shadow-sm">
          <span className="text-2xl">{focus.emoji}</span>
          <div className="flex-1"><div className="text-xs font-semibold uppercase tracking-wide text-white/80">Фокус тижня</div><div className="font-bold">{focus.name}</div></div>
          <button onClick={() => onWeek(null)} className="rounded-full bg-white/20 p-1.5 hover:bg-white/30"><X className="h-4 w-4" /></button>
        </div>
      )}

      <div className="space-y-2.5">
        {ordered.map((t) => {
          const fav = favorites.includes(t.id); const isTried = tried.includes(t.id); const isFocus = weekFocus === t.id;
          return (
            <div key={t.id} className={`rounded-2xl border p-4 shadow-sm transition ${isFocus ? "border-teal-300 bg-teal-50/50" : "border-slate-100 bg-white"}`}>
              <div className="flex items-start gap-3">
                <span className="grid h-11 w-11 shrink-0 place-items-center rounded-xl bg-slate-100 text-xl">{t.emoji}</span>
                <div className="min-w-0 flex-1">
                  <div className="flex items-center gap-2"><span className="font-bold text-slate-800">{t.name}</span></div>
                  <p className="mt-0.5 text-sm text-slate-600">{t.what}</p>
                  <p className="mt-1 text-xs text-slate-400"><b className="font-semibold text-slate-500">Чому помагає:</b> {t.why}</p>
                </div>
                <button onClick={() => toggle(favorites, t.id, onFav)} title="В улюблене" className={`shrink-0 rounded-full p-1.5 transition ${fav ? "text-amber-400" : "text-slate-300 hover:text-amber-400"}`}><Star className={`h-5 w-5 ${fav ? "fill-amber-400" : ""}`} /></button>
              </div>
              <div className="mt-3 flex flex-wrap items-center gap-2">
                <button onClick={() => toggle(tried, t.id, onTried)} className={`inline-flex items-center gap-1.5 rounded-full px-3 py-1 text-xs font-semibold transition ${isTried ? "bg-green-100 text-green-700" : "bg-slate-100 text-slate-500 hover:bg-slate-200"}`}>{isTried ? <><Check className="h-3.5 w-3.5" /> Пробувала</> : "Позначити як спробувала"}</button>
                <button onClick={() => onWeek(isFocus ? null : t.id)} className={`inline-flex items-center gap-1.5 rounded-full px-3 py-1 text-xs font-semibold transition ${isFocus ? "bg-teal-500 text-white" : "bg-slate-100 text-slate-500 hover:bg-slate-200"}`}>{isFocus ? "Фокус тижня ✓" : "Зробити фокусом тижня"}</button>
              </div>
            </div>
          );
        })}
      </div>

      {!settings.noteSeen && (
        <div className="flex items-start gap-2 rounded-2xl bg-slate-100/70 px-4 py-3 text-sm text-slate-500">
          <Info className="mt-0.5 h-4 w-4 shrink-0 text-slate-400" />
          <p className="flex-1">Ці техніки доповнюють, але не замінюють професійну допомогу. Якщо тривога сильна або триває тижнями — варто поговорити з терапевтом. 💛</p>
          <button onClick={onDismissNote} className="rounded-full p-0.5 text-slate-300 hover:text-slate-500"><X className="h-4 w-4" /></button>
        </div>
      )}
    </div>
  );
}

/* ---------- Chore Splitter (7-step wizard) ---------- */
const CHORE_STEPS = ["Учасники", "Справи", "Оцінка", "Разом", "Бали", "Розподіл", "Баланс"];
function ChoreSplitter(props) {
  const { members, list, ratings, assignments, meta, saveMembers, saveList, saveRatings, saveAssignments, saveMeta } = props;
  const step = meta.step || 1;
  const setStep = (s) => saveMeta({ step: Math.max(1, Math.min(7, s)) });

  if (meta.finished) return <ChoreBoard {...props} onEdit={() => saveMeta({ finished: false, step: 7 })} />;

  return (
    <div>
      <div className="mb-1 flex items-center gap-2"><span className="text-2xl">🧹</span><h1 className="text-2xl font-extrabold text-slate-900">Chore Splitter</h1></div>
      {/* stepper */}
      <div className="mb-5 flex items-center gap-1 overflow-x-auto pb-1">
        {CHORE_STEPS.map((s, i) => {
          const n = i + 1, active = n === step, done = n < step;
          return (
            <button key={s} onClick={() => setStep(n)} className="flex items-center gap-1">
              {i > 0 && <span className={`h-0.5 w-3 ${done || active ? "bg-rose-400" : "bg-slate-200"}`} />}
              <span className={`grid h-7 w-7 shrink-0 place-items-center rounded-full text-xs font-bold transition ${active ? "bg-rose-600 text-white" : done ? "bg-rose-100 text-rose-600" : "bg-slate-100 text-slate-400"}`}>{done ? "✓" : n}</span>
              <span className={`hidden whitespace-nowrap text-xs font-medium sm:inline ${active ? "text-rose-700" : "text-slate-400"}`}>{s}</span>
            </button>
          );
        })}
      </div>

      <div className="rounded-2xl border border-slate-200 bg-white p-4 shadow-sm sm:p-5">
        {step === 1 && <StepMembers members={members} onSave={saveMembers} />}
        {step === 2 && <StepBrainstorm list={list} onSave={saveList} />}
        {step === 3 && <StepRatings members={members} list={list} ratings={ratings} onSave={saveRatings} />}
        {step === 4 && <StepMerge members={members} list={list} ratings={ratings} />}
        {step === 5 && <StepScores members={members} list={list} ratings={ratings} meta={meta} onSaveRatings={saveRatings} onSaveMeta={saveMeta} />}
        {step === 6 && <StepAssign members={members} list={list} ratings={ratings} assignments={assignments} meta={meta} onSave={saveAssignments} />}
        {step === 7 && <StepBalance members={members} list={list} ratings={ratings} assignments={assignments} meta={meta} onSave={saveAssignments} onFinish={() => saveMeta({ finished: true })} />}
      </div>

      <div className="mt-4 flex items-center justify-between">
        <button disabled={step <= 1} onClick={() => setStep(step - 1)} className="inline-flex items-center gap-1.5 rounded-xl border border-slate-300 bg-white px-4 py-2 text-sm font-medium text-slate-700 transition hover:bg-slate-50 disabled:opacity-40"><ChevronLeft className="h-4 w-4" /> Назад</button>
        <span className="text-xs text-slate-400">Крок {step} / 7</span>
        {step < 7
          ? <button disabled={step === 2 && list.length === 0} onClick={() => setStep(step + 1)} className="inline-flex items-center gap-1.5 rounded-xl bg-rose-600 px-5 py-2 text-sm font-semibold text-white transition hover:bg-rose-700 disabled:opacity-40">Далі <ChevronRight className="h-4 w-4" /></button>
          : <span />}
      </div>
    </div>
  );
}

function StepMembers({ members, onSave }) {
  return (
    <div>
      <h2 className="mb-1 text-lg font-bold text-slate-900">Хто ділить справи?</h2>
      <p className="mb-4 text-sm text-slate-500">Додай учасників дому. Імена можна змінювати.</p>
      <div className="space-y-2">
        {members.map((m, i) => (
          <div key={m.id} className="flex items-center gap-2">
            <span className="grid h-9 w-9 shrink-0 place-items-center rounded-lg bg-rose-100 text-sm font-bold text-rose-600">{i + 1}</span>
            <input value={m.name} onChange={(e) => onSave(members.map((x) => x.id === m.id ? { ...x, name: e.target.value } : x))} className="flex-1 rounded-lg border border-slate-300 px-3 py-2 text-sm focus:border-rose-400 focus:outline-none" />
            {members.length > 2 && <button onClick={() => onSave(members.filter((x) => x.id !== m.id))} className="rounded p-1.5 text-slate-300 hover:text-red-500"><Trash2 className="h-4 w-4" /></button>}
          </div>
        ))}
      </div>
      <button onClick={() => onSave([...members, { id: ruid("m"), name: `Людина ${members.length + 1}` }])} className="mt-3 inline-flex items-center gap-1.5 rounded-lg border border-slate-300 px-3 py-1.5 text-sm font-medium text-slate-600 hover:bg-slate-50"><Plus className="h-4 w-4" /> Додати учасника</button>
    </div>
  );
}

function StepBrainstorm({ list, onSave }) {
  const [text, setText] = useState("");
  const has = (t) => list.some((c) => c.text.toLowerCase() === t.toLowerCase());
  const add = (t) => { const v = t.trim(); if (v && !has(v)) onSave([...list, { id: ruid("ch"), text: v }]); };
  return (
    <div>
      <h2 className="mb-1 text-lg font-bold text-slate-900">Всі хатні справи в один список</h2>
      <p className="mb-3 text-sm text-slate-500">💡 Розбивай великі справи на маленькі: «прання» → «завантажити» + «розвісити» + «скласти». Так нічого не загубиться й легше ділити.</p>

      <div className="mb-3 flex items-center gap-2">
        <input value={text} onChange={(e) => setText(e.target.value)} onKeyDown={(e) => { if (e.key === "Enter") { add(text); setText(""); } }} placeholder="Додати справу…" className="flex-1 rounded-lg border border-slate-300 px-3 py-2 text-sm focus:border-rose-400 focus:outline-none" />
        <button onClick={() => { add(text); setText(""); }} className="rounded-lg bg-rose-600 px-3 py-2 text-sm font-semibold text-white hover:bg-rose-700"><Plus className="h-4 w-4" /></button>
      </div>

      <div className="mb-4 space-y-3">
        <div className="text-xs font-semibold uppercase tracking-wide text-slate-400">Готовий чек-лист по кімнатах — тапни, щоб додати</div>
        {CHORE_ROOMS.map((r) => (
          <div key={r.room}>
            <div className="mb-1 flex items-center gap-1.5 text-sm font-semibold text-slate-600"><Home className="h-3.5 w-3.5 text-slate-400" /> {r.room}</div>
            <div className="flex flex-wrap gap-1.5">
              {r.items.map((it) => { const added = has(it); return (
                <button key={it} onClick={() => added ? onSave(list.filter((c) => c.text.toLowerCase() !== it.toLowerCase())) : add(it)} className={`rounded-full px-3 py-1 text-xs font-medium transition ${added ? "bg-rose-600 text-white" : "bg-slate-100 text-slate-600 hover:bg-slate-200"}`}>{added ? "✓ " : "+ "}{it}</button>
              ); })}
            </div>
          </div>
        ))}
      </div>

      {list.length > 0 && (
        <div>
          <div className="mb-1.5 text-xs font-semibold uppercase tracking-wide text-slate-400">Твій список · {list.length}</div>
          <div className="flex flex-wrap gap-1.5">
            {list.map((c) => (
              <span key={c.id} className="inline-flex items-center gap-1 rounded-full bg-rose-50 px-3 py-1 text-sm text-rose-700">{c.text}<button onClick={() => onSave(list.filter((x) => x.id !== c.id))} className="text-rose-300 hover:text-red-500"><X className="h-3.5 w-3.5" /></button></span>
            ))}
          </div>
        </div>
      )}
    </div>
  );
}

function StepRatings({ members, list, ratings, onSave }) {
  const [mi, setMi] = useState(0);
  const member = members[mi];
  const setAtt = (choreId, attitude) => {
    const next = { ...ratings, [choreId]: { ...(ratings[choreId] || {}), [member.id]: { ...(ratings[choreId]?.[member.id] || {}), attitude } } };
    onSave(next);
  };
  const ratedCount = (m) => list.filter((c) => ratings[c.id]?.[m.id]?.attitude).length;
  return (
    <div>
      <h2 className="mb-1 text-lg font-bold text-slate-900">Приватна оцінка ставлення</h2>
      <p className="mb-3 rounded-lg bg-amber-50 px-3 py-2 text-xs text-amber-800">🔒 Оцінюй чесно й наодинці. Відповіді іншого тут <b>не показуються</b> — по черзі, по одній людині.</p>
      <div className="mb-3 flex flex-wrap gap-2">
        {members.map((m, i) => (
          <button key={m.id} onClick={() => setMi(i)} className={`rounded-full px-3 py-1.5 text-sm font-semibold transition ${i === mi ? "bg-rose-600 text-white" : "bg-slate-100 text-slate-500 hover:bg-slate-200"}`}>{m.name} <span className="opacity-70">({ratedCount(m)}/{list.length})</span></button>
        ))}
      </div>
      <div className="space-y-2">
        {list.map((c) => {
          const cur = ratings[c.id]?.[member.id]?.attitude;
          return (
            <div key={c.id} className="flex items-center gap-2 rounded-xl border border-slate-100 p-2">
              <span className="min-w-0 flex-1 truncate text-sm font-medium text-slate-700">{c.text}</span>
              <div className="flex shrink-0 gap-1">
                {ATT_ORDER.map((a) => (
                  <button key={a} onClick={() => setAtt(c.id, a)} title={ATT[a].label} className="grid h-9 w-9 place-items-center rounded-lg text-lg transition" style={cur === a ? { backgroundColor: ATT[a].color + "22", boxShadow: `0 0 0 2px ${ATT[a].color}` } : { backgroundColor: "#f8fafc" }}>{ATT[a].emoji}</button>
                ))}
              </div>
            </div>
          );
        })}
      </div>
    </div>
  );
}

function StepMerge({ members, list, ratings }) {
  return (
    <div>
      <h2 className="mb-1 text-lg font-bold text-slate-900">Спільна картина</h2>
      <p className="mb-3 text-sm text-slate-500">Ставлення кожного до кожної справи. Поки без розподілу — просто дивимось разом.</p>
      <div className="overflow-x-auto rounded-xl border border-slate-200">
        <table className="w-full text-left text-sm">
          <thead className="bg-slate-50 text-xs uppercase tracking-wide text-slate-400"><tr><th className="px-3 py-2 font-medium">Справа</th>{members.map((m) => <th key={m.id} className="px-3 py-2 font-medium">{m.name}</th>)}</tr></thead>
          <tbody className="divide-y divide-slate-100">
            {list.map((c) => (
              <tr key={c.id}><td className="px-3 py-2 text-slate-700">{c.text}</td>{members.map((m) => { const a = ratings[c.id]?.[m.id]?.attitude; return <td key={m.id} className="px-3 py-2">{a ? <span className="inline-flex items-center gap-1"><span>{ATT[a].emoji}</span><span className="text-xs" style={{ color: ATT[a].color }}>{ATT[a].label}</span></span> : <span className="text-slate-300">—</span>}</td>; })}</tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

function StepScores({ members, list, ratings, meta, onSaveRatings, onSaveMeta }) {
  const use = meta.useScores;
  const setScore = (choreId, memberId, score) => {
    const s = Math.max(1, Math.min(10, Number(score) || 1));
    onSaveRatings({ ...ratings, [choreId]: { ...(ratings[choreId] || {}), [memberId]: { ...(ratings[choreId]?.[memberId] || {}), score: s } } });
  };
  return (
    <div>
      <h2 className="mb-1 text-lg font-bold text-slate-900">Бали 1–10 <span className="text-sm font-normal text-slate-400">(необовʼязково)</span></h2>
      <p className="mb-3 text-sm text-slate-500">Точніше за «подобається/терпимо/ненавиджу»: <b>1</b> = щиро подобається, <b>10</b> = «краще перееду, ніж робитиму це». Можна пропустити.</p>
      <label className="mb-4 flex items-center justify-between rounded-xl border border-slate-200 bg-slate-50 px-3 py-2.5">
        <span className="text-sm font-medium text-slate-700">Використати бали 1–10</span>
        <input type="checkbox" checked={use} onChange={(e) => onSaveMeta({ useScores: e.target.checked })} className="h-4 w-4 accent-rose-600" />
      </label>
      {use && (
        <div className="overflow-x-auto rounded-xl border border-slate-200">
          <table className="w-full text-left text-sm">
            <thead className="bg-slate-50 text-xs uppercase tracking-wide text-slate-400"><tr><th className="px-3 py-2 font-medium">Справа</th>{members.map((m) => <th key={m.id} className="px-3 py-2 font-medium">{m.name}</th>)}</tr></thead>
            <tbody className="divide-y divide-slate-100">
              {list.map((c) => (
                <tr key={c.id}><td className="px-3 py-2 text-slate-700">{c.text}</td>{members.map((m) => { const r = ratings[c.id]?.[m.id]; const val = r?.score ?? ATT[r?.attitude]?.weight ?? ""; return <td key={m.id} className="px-3 py-2"><input type="number" min="1" max="10" value={val} onChange={(e) => setScore(c.id, m.id, e.target.value)} className="w-16 rounded-lg border border-slate-300 px-2 py-1 text-sm" /></td>; })}</tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
      {!use && <p className="text-sm text-slate-400">Пропускаєш — розподіл рахуватиметься по «подобається/терпимо/ненавиджу».</p>}
    </div>
  );
}

function StepAssign({ members, list, ratings, assignments, meta, onSave }) {
  const auto = () => {
    const next = { ...assignments };
    for (const c of list) {
      let best = null, bestScore = Infinity;
      for (const m of members) { const s = choreScore(ratings, c.id, m.id, meta.useScores); if (s != null && s < bestScore) { bestScore = s; best = m.id; } }
      if (best) next[c.id] = { memberId: best, delegate: assignments[c.id]?.delegate || false };
    }
    onSave(next);
  };
  const setAssignee = (choreId, memberId) => onSave({ ...assignments, [choreId]: { ...(assignments[choreId] || {}), memberId } });
  const toggleDeleg = (choreId) => onSave({ ...assignments, [choreId]: { ...(assignments[choreId] || {}), delegate: !assignments[choreId]?.delegate } });
  return (
    <div>
      <h2 className="mb-1 text-lg font-bold text-slate-900">Розподіл</h2>
      <p className="mb-3 text-sm text-slate-500">Кожну справу — тому, кому вона <b>найменш неприємна</b> (менший бал/ставлення). «Можна делегувати?» — підстрахування на важкий день.</p>
      <button onClick={auto} className="mb-3 inline-flex items-center gap-1.5 rounded-lg bg-rose-600 px-4 py-2 text-sm font-semibold text-white hover:bg-rose-700"><SparklesIcon className="h-4 w-4" /> Призначити автоматично</button>
      <div className="space-y-2">
        {list.map((c) => {
          const a = assignments[c.id];
          return (
            <div key={c.id} className="rounded-xl border border-slate-100 p-3">
              <div className="mb-2 flex items-center justify-between gap-2">
                <span className="min-w-0 flex-1 text-sm font-medium text-slate-800">{c.text}</span>
                <label className="flex shrink-0 items-center gap-1.5 text-xs text-slate-500"><input type="checkbox" checked={!!a?.delegate} onChange={() => toggleDeleg(c.id)} className="h-3.5 w-3.5 accent-green-600" /> Можна делегувати</label>
              </div>
              <div className="flex flex-wrap gap-1.5">
                {members.map((m) => { const s = choreScore(ratings, c.id, m.id, meta.useScores); const sel = a?.memberId === m.id; return (
                  <button key={m.id} onClick={() => setAssignee(c.id, m.id)} className={`inline-flex items-center gap-1 rounded-lg px-2.5 py-1.5 text-sm font-medium transition ${sel ? "bg-rose-600 text-white" : "bg-slate-100 text-slate-600 hover:bg-slate-200"}`}>{m.name}{s != null && <span className={`text-xs ${sel ? "text-white/70" : "text-slate-400"}`}>· {s}</span>}</button>
                ); })}
              </div>
            </div>
          );
        })}
      </div>
    </div>
  );
}

function memberTotals(members, list, ratings, assignments, useScores) {
  const totals = {}; for (const m of members) totals[m.id] = 0;
  for (const c of list) { const a = assignments[c.id]; if (!a?.memberId) continue; const s = choreScore(ratings, c.id, a.memberId, useScores); if (s != null && totals[a.memberId] != null) totals[a.memberId] += s; }
  return totals;
}

function StepBalance({ members, list, ratings, assignments, meta, onSave, onFinish }) {
  const totals = memberTotals(members, list, ratings, assignments, meta.useScores);
  const max = Math.max(1, ...Object.values(totals));
  const move = (choreId, toId) => onSave({ ...assignments, [choreId]: { ...(assignments[choreId] || {}), memberId: toId } });
  const colors = ["#6366f1", "#ec4899", "#14b8a6", "#f59e0b", "#8b5cf6"];
  return (
    <div>
      <h2 className="mb-1 text-lg font-bold text-slate-900">Баланс по складності</h2>
      <p className="mb-3 text-sm text-slate-500">Рівняй за <b>сумарною складністю</b>, а не за кількістю: одна «10» переважує п'ять «2». Переноси справи, поки суми не стануть чесними.</p>

      <div className="mb-4 space-y-2">
        {members.map((m, i) => (
          <div key={m.id}>
            <div className="mb-1 flex items-center justify-between text-sm"><span className="font-semibold text-slate-700">{m.name}</span><span className="font-bold tabular-nums" style={{ color: colors[i % colors.length] }}>{totals[m.id]}</span></div>
            <div className="h-3 overflow-hidden rounded-full bg-slate-100"><div className="h-full rounded-full transition-all" style={{ width: `${(totals[m.id] / max) * 100}%`, backgroundColor: colors[i % colors.length] }} /></div>
          </div>
        ))}
      </div>

      <div className="space-y-3">
        {members.map((m) => (
          <div key={m.id} className="rounded-xl border border-slate-100 p-3">
            <div className="mb-1.5 text-sm font-bold text-slate-700">{m.name}</div>
            <div className="space-y-1">
              {list.filter((c) => assignments[c.id]?.memberId === m.id).map((c) => {
                const s = choreScore(ratings, c.id, m.id, meta.useScores);
                const others = members.filter((x) => x.id !== m.id);
                return (
                  <div key={c.id} className="flex items-center gap-2 text-sm">
                    <span className="min-w-0 flex-1 truncate text-slate-700">{c.text} {s != null && <span className="text-xs text-slate-400">· {s}</span>} {assignments[c.id]?.delegate && <span className="text-xs text-green-600">↔</span>}</span>
                    {others.map((o) => <button key={o.id} onClick={() => move(c.id, o.id)} title={`Перенести до ${o.name}`} className="shrink-0 rounded-md bg-slate-100 px-2 py-0.5 text-xs font-medium text-slate-500 hover:bg-slate-200">→ {o.name}</button>)}
                  </div>
                );
              })}
              {list.filter((c) => assignments[c.id]?.memberId === m.id).length === 0 && <div className="text-xs text-slate-300">поки нічого</div>}
            </div>
          </div>
        ))}
      </div>

      <button onClick={onFinish} className="mt-4 w-full rounded-2xl bg-rose-600 py-3 font-bold text-white shadow-lg shadow-rose-500/20 hover:bg-rose-700">Готово — показати дошку 📋</button>
    </div>
  );
}

function ChoreBoard({ members, list, ratings, assignments, meta, saveMeta, onEdit }) {
  const totals = memberTotals(members, list, ratings, assignments, meta.useScores);
  const colors = ["#6366f1", "#ec4899", "#14b8a6", "#f59e0b", "#8b5cf6"];
  return (
    <div>
      <div className="mb-1 flex items-center gap-2"><span className="text-2xl">📋</span><h1 className="text-2xl font-extrabold text-slate-900">Наша дошка справ</h1></div>
      <p className="mb-4 rounded-lg bg-rose-50 px-3 py-2 text-sm text-rose-800">📌 Повісь це на видному місці у спільному просторі (холодильник, коридор). З СДУГ: <b>з очей — з голови</b>. Зелене — можна делегувати на важкий день.</p>

      <div className="grid gap-3 sm:grid-cols-2">
        {members.map((m, i) => {
          const mine = list.filter((c) => assignments[c.id]?.memberId === m.id);
          return (
            <div key={m.id} className="rounded-2xl border-2 bg-white p-4 shadow-sm" style={{ borderColor: colors[i % colors.length] + "55" }}>
              <div className="mb-2 flex items-center justify-between">
                <span className="text-lg font-extrabold" style={{ color: colors[i % colors.length] }}>{m.name}</span>
                <span className="rounded-full bg-slate-100 px-2 py-0.5 text-xs font-bold tabular-nums text-slate-500">склад. {totals[m.id]}</span>
              </div>
              <ul className="space-y-1.5">
                {mine.map((c) => { const deleg = assignments[c.id]?.delegate; return (
                  <li key={c.id} className={`flex items-center gap-2 rounded-lg px-2.5 py-1.5 text-sm ${deleg ? "bg-green-50 text-green-800" : "bg-slate-50 text-slate-700"}`}>
                    <span className={`h-2 w-2 shrink-0 rounded-full ${deleg ? "bg-green-500" : "bg-slate-300"}`} />
                    <span className="flex-1">{c.text}</span>
                    {deleg && <span className="text-[10px] font-semibold uppercase text-green-600">делег.</span>}
                  </li>
                ); })}
                {mine.length === 0 && <li className="px-2 py-1 text-sm text-slate-300">—</li>}
              </ul>
            </div>
          );
        })}
      </div>

      <div className="mt-5 flex flex-wrap gap-3">
        <button onClick={onEdit} className="inline-flex items-center gap-2 rounded-lg border border-slate-300 bg-white px-4 py-2 text-sm font-medium text-slate-700 hover:bg-slate-50"><Pencil className="h-4 w-4" /> Редагувати розподіл</button>
        <button onClick={() => saveMeta({ finished: false, step: 1 })} className="inline-flex items-center gap-2 rounded-lg border border-slate-300 bg-white px-4 py-2 text-sm font-medium text-slate-700 hover:bg-slate-50"><RefreshCw className="h-4 w-4" /> Пройти заново</button>
      </div>
    </div>
  );
}
