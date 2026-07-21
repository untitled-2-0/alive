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
import { cloudPush, cloudRemove, isSignedIn as cloudSignedIn, currentEmail, sendCode, verifyCode, signOutCloud, refreshSession, syncNow } from "./cloud.js";

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
  return { fears: d.fears, thoughts: d.thoughts, sessions: d.sessions, settings: d.settings, recovery };
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
      let learn = 0, review = 0, newTotal = 0;
      const stages = {}; // stageId -> count, across ALL cards (interval breakdown)
      for (const c of cards) {
        if (c.state === "new") newTotal += 1;
        else if (c.state === "learning" && c.due <= endTs) learn += 1;
        else if (c.state === "review" && c.due <= endTs) review += 1;
        const sid = stageForCard(c).id;
        stages[sid] = (stages[sid] || 0) + 1;
      }
      // Strict "due today" = only cards the schedule actually placed today or earlier.
      // New cards are NOT padded in — they're studied via the "only new"/"all" modes.
      out[d.id] = { learn, review, newTotal, newDue: newTotal, total: cards.length, due: learn + review, stages };
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
    if (card && deck?.autoPlay && !session.flipped) say(card.front);
  }, [card?.id]); // eslint-disable-line react-hooks/exhaustive-deps
  useEffect(() => {
    if (card && deck?.autoPlay && session.flipped) say(card.back);
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
  const onTouchMove = (e) => { if (!dragRef.current) return; const t = e.touches[0]; dragRef.current.dx = t.clientX - dragRef.current.x0; dragRef.current.dy = t.clientY - dragRef.current.y0; if (session.flipped) setDrag(dragRef.current.dx); };
  const onTouchEnd = () => {
    const st = dragRef.current; dragRef.current = null;
    if (!st) return;
    if (!session.flipped) { if (Math.abs(st.dx) > 45 || Math.abs(st.dy) > 45) onFlip(); setDrag(0); return; }
    if (st.dx > 85) onAnswer("good");
    else if (st.dx < -85) onAnswer("again");
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
        {imgs.front && (
          <img
            src={imgs.front}
            alt=""
            onClick={(e) => { e.stopPropagation(); setLightbox(imgs.front); }}
            className="mb-4 max-h-56 w-auto cursor-zoom-in rounded-lg border border-slate-200 object-contain"
          />
        )}
        <div className="flex items-center gap-2">
          <div className="text-2xl font-semibold leading-snug text-slate-900" style={{ textWrap: "balance" }}>
            {card.front}
          </div>
          <SpeakerButton text={card.front} lang={lang} onFallback={() => setVoiceHint(true)} />
        </div>

        {session.flipped ? (
          <>
            <div className="my-6 h-px w-24 bg-slate-200" />
            {imgs.back && (
              <img
                src={imgs.back}
                alt=""
                onClick={(e) => { e.stopPropagation(); setLightbox(imgs.back); }}
                className="mb-4 max-h-56 w-auto cursor-zoom-in rounded-lg border border-slate-200 object-contain"
              />
            )}
            <div className="flex items-center gap-2">
              <div className="text-2xl font-medium text-rose-700" style={{ textWrap: "balance" }}>{card.back}</div>
              <SpeakerButton text={card.back} lang={lang} onFallback={() => setVoiceHint(true)} />
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
    <div className="fixed inset-0 z-40 flex items-end justify-center bg-slate-900/40 p-0 backdrop-blur-sm sm:items-center sm:p-4" onClick={onClose}>
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
    <div className="fixed inset-0 z-40 flex items-end justify-center bg-slate-900/40 p-0 backdrop-blur-sm sm:items-center sm:p-4" onClick={onClose}>
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
    <div className="fixed inset-0 z-40 flex items-end justify-center bg-slate-900/40 p-0 backdrop-blur-sm sm:items-center sm:p-4" onClick={onClose}>
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
              <span className="w-full text-xs text-slate-400">Лист надіслано на {input} (перевір і спам). Введи код — або просто клікни посилання в листі, воно поверне тебе сюди вже залогіненою.</span>
            </div>
          )}
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
        <div className="fixed inset-0 z-40 flex items-end justify-center bg-slate-900/30 p-0 backdrop-blur-sm sm:items-center sm:p-4" onClick={() => setMoodOpen(false)}>
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
    <div className="fixed inset-0 z-40 flex items-end justify-center bg-slate-900/40 p-0 backdrop-blur-sm sm:items-center sm:p-4" onClick={onClose}>
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
    <div className="fixed inset-0 z-40 flex items-end justify-center bg-slate-900/40 p-0 backdrop-blur-sm sm:items-center sm:p-4" onClick={onClose}>
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

        <div className="mt-6 flex items-center justify-end gap-3">
          <button onClick={onClose} className="rounded-lg px-4 py-2 text-sm font-medium text-slate-500 hover:text-slate-800">Cancel</button>
          <button onClick={save} disabled={!title.trim()} className="inline-flex items-center gap-2 rounded-xl bg-pink-500 px-5 py-2.5 font-semibold text-white shadow-sm transition hover:bg-pink-600 disabled:bg-slate-300">
            <Check className="h-4 w-4" /> {task ? "Save task" : "Add task"}
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
    <div className="fixed inset-0 z-40 flex items-end justify-center bg-slate-900/40 p-0 backdrop-blur-sm sm:items-center sm:p-4" onClick={onClose}>
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
          <button onClick={() => setCview("beforework")} className="mb-4 flex w-full items-center gap-3 rounded-3xl bg-gradient-to-r from-teal-400 to-sky-400 p-4 text-left text-white shadow-lg shadow-teal-500/20 transition hover:from-teal-500 hover:to-sky-500">
            <span className="grid h-12 w-12 shrink-0 place-items-center rounded-2xl bg-white/20"><Sparkle className="h-6 w-6" /></span>
            <span className="flex-1"><span className="block text-lg font-bold">Перед роботою</span><span className="block text-sm text-white/90">2 хв дихання → заземлення. Одне натискання — і ти в ресурсі.</span></span>
            <ArrowRight className="h-5 w-5" />
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

  const sorted = [...fears].sort((a, b) => a.intensity - b.intensity);
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
      <p className="mb-4 rounded-2xl bg-amber-50 px-3 py-2 text-xs leading-relaxed text-amber-800">Починай знизу — з найлегшої сходинки. Піднімайся вище лише тоді, коли сходинка відчувається спокійно й посильно. Поспіху немає, і немає «неправильного» темпу. Пробуй кожну в розслабленому стані.</p>

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
          {[...sorted].reverse().map((f, ri) => {
            const step = sorted.length - ri;
            const atts = f.attempts || [];
            const first = atts[0]?.after, last = atts[atts.length - 1]?.after;
            const drop = atts.length >= 1 ? (f.intensity - (last ?? f.intensity)) : 0;
            return (
              <div key={f.id} className="rounded-2xl bg-white p-4 shadow-sm ring-1 ring-amber-50" style={{ marginLeft: Math.min(ri, 6) * 6 }}>
                <div className="flex items-center gap-2">
                  <span className="grid h-8 w-8 shrink-0 place-items-center rounded-lg bg-amber-100 text-sm font-bold text-amber-700">{step}</span>
                  <div className="min-w-0 flex-1"><div className="truncate font-bold text-slate-800">{f.title}</div>
                    <div className="text-xs text-slate-400">Зараз {last ?? f.intensity}/100{atts.length ? ` · ${atts.length} ${atts.length === 1 ? "спроба" : atts.length < 5 ? "спроби" : "спроб"}` : ""}{drop > 0 ? ` · ↓${drop}` : ""}</div></div>
                  <button onClick={() => { setLogId(f.id); setBefore(last ?? f.intensity); }} className="rounded-full bg-amber-500 px-3 py-1.5 text-xs font-semibold text-white hover:bg-amber-600">Записати спробу</button>
                  <button onClick={() => { if (confirm("Прибрати цей страх?")) removeFear(f.id); }} className="text-slate-300 hover:text-red-500"><Trash2 className="h-4 w-4" /></button>
                </div>
                {atts.length > 1 && (
                  <div className="mt-2 flex items-end gap-1">
                    {atts.slice(-12).map((a, i) => <div key={i} title={`${a.date}: ${a.before}→${a.after}`} className="flex-1 rounded-t bg-amber-300" style={{ height: Math.max(4, (a.after / 100) * 36) }} />)}
                  </div>
                )}
              </div>
            );
          })}
        </div>
      )}

      {logFear && (
        <div className="fixed inset-0 z-40 flex items-end justify-center bg-slate-900/40 p-0 backdrop-blur-sm sm:items-center sm:p-4" onClick={() => setLogId(null)}>
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

  const startFast = useCallback(async () => {
    await saveCurrent({ startTs: Date.now(), targetHrs: protocol.hrs, protocol: protocol.id });
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
        <button onClick={onStart} className="mt-6 rounded-2xl bg-orange-500 px-12 py-3.5 font-bold text-white shadow-lg shadow-orange-500/25 transition hover:bg-orange-600">Почати голодування</button>
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
    <div className="fixed inset-0 z-40 flex items-end justify-center bg-slate-900/40 p-0 backdrop-blur-sm sm:items-center sm:p-4" onClick={onClose}>
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
  const budgets = await store.get(BKEYS.budgets, {});
  const month = await store.get(BKEYS.month, budDefaultMonth());
  const history = await store.get(BKEYS.history, []);
  const settings = await store.get(BKEYS.settings, { name: "Budget" });
  return { cats, items, bought, stock, budgets, month, history, settings };
}
async function collectBudgetExport() {
  const d = await loadBudgetData();
  return { cats: d.cats || [], items: d.items || [], bought: d.bought, stock: d.stock, budgets: d.budgets, month: d.month, history: d.history, settings: d.settings };
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
    setCats(d.cats || []); setItems(d.items || []); setBought(d.bought || {}); setStock(d.stock || {}); setBudgets(d.budgets || {});
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
  const setBoughtItem = (id) => {
    const turningOn = !boughtMap[id];
    setBought((prev) => { const mm = { ...(prev[month] || {}) }; if (mm[id]) delete mm[id]; else mm[id] = true; const next = { ...prev, [month]: mm }; store.set(BKEYS.bought, next); return next; });
    // куплено ↔ в запасі взаємовиключні: якщо позначили купленим, знімаємо «в запасі»
    if (turningOn) setStock((prev) => { const mm = { ...(prev[month] || {}) }; if (!mm[id]) return prev; delete mm[id]; const next = { ...prev, [month]: mm }; store.set(BKEYS.stock, next); return next; });
  };
  const setStockItem = (id) => {
    const turningOn = !stockMap[id];
    setStock((prev) => { const mm = { ...(prev[month] || {}) }; if (mm[id]) delete mm[id]; else mm[id] = true; const next = { ...prev, [month]: mm }; store.set(BKEYS.stock, next); return next; });
    if (turningOn) setBought((prev) => { const mm = { ...(prev[month] || {}) }; if (!mm[id]) return prev; delete mm[id]; const next = { ...prev, [month]: mm }; store.set(BKEYS.bought, next); return next; });
  };

  const itemsOf = (cid) => items.filter((it) => it.catId === cid);
  const planned = useMemo(() => items.reduce((s, it) => s + budLineSum(it), 0), [items]);
  const spent = useMemo(() => items.reduce((s, it) => s + (boughtMap[it.id] ? budLineSum(it) : 0), 0), [items, boughtMap]);
  const inStockSum = useMemo(() => items.reduce((s, it) => s + (stockMap[it.id] ? budLineSum(it) : 0), 0), [items, stockMap]);
  const toBuy = useMemo(() => items.reduce((s, it) => s + ((boughtMap[it.id] || stockMap[it.id]) ? 0 : budLineSum(it)), 0), [items, boughtMap, stockMap]);
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
              <CategoryBlock key={c.id} cat={c} items={itemsOf(c.id)} boughtMap={boughtMap} stockMap={stockMap} collapsed={!!collapsed[c.id]}
                onToggleCollapse={() => setCollapsed((m) => ({ ...m, [c.id]: !m[c.id] }))}
                onToggleBought={setBoughtItem} onToggleStock={setStockItem} onAddItem={() => setItemEditor({ item: null, catId: c.id })}
                onEditItem={(it) => setItemEditor({ item: it, catId: c.id })} onDeleteItem={deleteItem} />
            ))}
            <button onClick={() => setCatMgr(true)} className="flex w-full items-center justify-center gap-1.5 rounded-2xl border border-dashed border-emerald-300 py-3 text-sm font-semibold text-emerald-600 hover:bg-emerald-50"><Plus className="h-4 w-4" /> Категорія</button>
          </div>
        )}

        {bview === "shop" && (
          <ShoppingView cats={cats} items={items} boughtMap={boughtMap} stockMap={stockMap} onToggle={setBoughtItem} onToggleStock={setStockItem}
            filter={shopFilter} setFilter={setShopFilter} flat={shopFlat} setFlat={setShopFlat} />
        )}

        {bview === "chart" && (
          <BudgetChart cats={cats} items={items} boughtMap={boughtMap} mode={chartMode} setMode={setChartMode} planned={planned} spent={spent} history={history} />
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

function CategoryBlock({ cat, items, boughtMap, stockMap, collapsed, onToggleCollapse, onToggleBought, onToggleStock, onAddItem, onEditItem, onDeleteItem }) {
  const subtotal = items.reduce((s, it) => s + budLineSum(it), 0);
  const boughtCount = items.filter((it) => boughtMap[it.id]).length;
  const stockCount = items.filter((it) => stockMap[it.id]).length;
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
          {items.map((it) => <BudgetItemRow key={it.id} item={it} bought={!!boughtMap[it.id]} inStock={!!stockMap[it.id]} onToggle={() => onToggleBought(it.id)} onToggleStock={() => onToggleStock(it.id)} onEdit={() => onEditItem(it)} onDelete={() => onDeleteItem(it.id)} />)}
          <button onClick={onAddItem} className="flex w-full items-center gap-1.5 px-4 py-2.5 text-left text-sm font-medium text-emerald-600 hover:bg-emerald-50"><Plus className="h-4 w-4" /> Товар</button>
        </div>
      )}
    </div>
  );
}

function BudgetItemRow({ item, bought, inStock, onToggle, onToggleStock, onEdit, onDelete }) {
  const [showNotes, setShowNotes] = useState(false);
  const nameCls = bought ? "text-slate-400 line-through" : inStock ? "text-amber-600" : "text-slate-800";
  return (
    <div className="group px-4 py-2.5">
      <div className="flex items-center gap-3">
        <button onClick={onToggle} className={`grid h-6 w-6 shrink-0 place-items-center rounded-md border-2 transition ${bought ? "border-transparent bg-emerald-500 text-white" : "border-slate-300 hover:border-emerald-400"}`}>{bought && <Check className="h-4 w-4" />}</button>
        <button onClick={onEdit} className="min-w-0 flex-1 text-left">
          <div className={`flex items-center gap-1.5 truncate text-sm font-medium ${nameCls}`}>{item.name}{inStock && <span className="shrink-0 rounded-full bg-amber-100 px-1.5 py-0.5 text-[10px] font-bold text-amber-700">в запасі</span>}</div>
          <div className="text-xs text-slate-400 tabular-nums">{item.qty} {item.unit} × {budFmt(item.price)} = <span className="font-semibold text-slate-500">{budFmt(budLineSum(item))}</span></div>
          {budFreqHint(item.qty) && <div className="text-[11px] text-slate-400">{budFreqHint(item.qty)}</div>}
        </button>
        <button onClick={onToggleStock} title="Вже є в запасі — докуповувати не треба" className={`shrink-0 rounded-md p-1 transition ${inStock ? "text-amber-500" : "text-slate-300 hover:text-amber-500"}`}><Package className="h-4 w-4" /></button>
        {item.notes && <button onClick={() => setShowNotes((v) => !v)} title="Нотатки" className={`shrink-0 rounded-md p-1 ${showNotes ? "text-emerald-500" : "text-slate-300 hover:text-slate-500"}`}><Info className="h-4 w-4" /></button>}
        <button onClick={onDelete} className="shrink-0 rounded-md p-1 text-slate-300 hover:text-red-500 sm:opacity-0 sm:group-hover:opacity-100"><Trash2 className="h-4 w-4" /></button>
      </div>
      {showNotes && item.notes && <div className="mt-1 rounded-lg bg-slate-50 px-3 py-1.5 text-xs text-slate-500">{item.notes}</div>}
    </div>
  );
}

function ShoppingView({ cats, items, boughtMap, stockMap, onToggle, onToggleStock, filter, setFilter, flat, setFlat }) {
  const total = items.length;
  const done = items.filter((it) => boughtMap[it.id]).length;
  const stockCount = items.filter((it) => stockMap[it.id]).length;
  const cartTotal = items.reduce((s, it) => s + (boughtMap[it.id] ? budLineSum(it) : 0), 0);
  // "Лишилось купити" ховає і куплене, і те, що вже є в запасі
  const visible = (list) => (filter ? list.filter((it) => !boughtMap[it.id] && !stockMap[it.id]) : list);
  const Row = (it) => {
    const isBought = !!boughtMap[it.id], isStock = !!stockMap[it.id];
    return (
      <div key={it.id} className={`flex items-center gap-3 rounded-2xl p-4 shadow-sm ring-1 transition ${isBought ? "bg-emerald-50 ring-emerald-100" : isStock ? "bg-amber-50/70 ring-amber-100" : "bg-white ring-slate-100 hover:ring-emerald-200"}`}>
        <button onClick={() => onToggle(it.id)} className={`grid h-9 w-9 shrink-0 place-items-center rounded-full border-2 transition ${isBought ? "border-transparent bg-emerald-500 text-white" : "border-slate-300 hover:border-emerald-400"}`}>{isBought && <Check className="h-5 w-5" />}</button>
        <button onClick={() => onToggle(it.id)} className="min-w-0 flex-1 text-left"><span className={`flex items-center gap-1.5 truncate font-bold ${isBought ? "text-slate-400 line-through" : isStock ? "text-amber-600" : "text-slate-800"}`}>{it.name}{isStock && <span className="shrink-0 rounded-full bg-amber-100 px-1.5 py-0.5 text-[10px] font-bold text-amber-700">в запасі</span>}</span><span className="block text-xs text-slate-400 tabular-nums">{it.qty} {it.unit} × {budFmt(it.price)}{budFreqHint(it.qty) ? ` · ${budFreqHint(it.qty)}` : ""}</span></button>
        <button onClick={() => onToggleStock(it.id)} title="Вже є в запасі" className={`shrink-0 rounded-md p-1.5 transition ${isStock ? "text-amber-500" : "text-slate-300 hover:text-amber-500"}`}><Package className="h-4 w-4" /></button>
        <span className="shrink-0 text-sm font-bold tabular-nums text-slate-600">{budFmt(budLineSum(it))}</span>
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
function BudgetChart({ cats, items, boughtMap, mode, setMode, planned, spent, history }) {
  const data = cats.map((c, i) => {
    const list = items.filter((it) => it.catId === c.id);
    const val = mode === "planned" ? list.reduce((s, it) => s + budLineSum(it), 0) : list.reduce((s, it) => s + (boughtMap[it.id] ? budLineSum(it) : 0), 0);
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
const CAREERKEYS = { skills: "career:skills", achievements: "career:achievements", reviews: "career:reviews", path: "career:pathProgress" };
async function collectCareerExport() { return { skills: await store.get(CAREERKEYS.skills, []), achievements: await store.get(CAREERKEYS.achievements, []), reviews: await store.get(CAREERKEYS.reviews, []), path: await store.get(CAREERKEYS.path, {}) }; }
async function clearCareerData() { for (const k of Object.values(CAREERKEYS)) await store.remove(k); await store.remove("career:seeded"); }

function CareerView() {
  const today = dateKey(Date.now());
  const week = finWeekStart(today);
  const [tab, setTab] = useState("path"); // path | skills | wins | review
  const [skills, setSkills] = useState([]);
  const [wins, setWins] = useState([]);
  const [reviews, setReviews] = useState([]);
  const [pathProg, setPathProg] = useState({});
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
    const rv = await store.get(CAREERKEYS.reviews, []); setReviews(rv);
    setReviewText((rv.find((r) => r.week === week) || {}).text || "");
  })(); }, [week]);
  const setStepDone = (stepId, done) => setPathProg((prev) => { const n = { ...prev }; if (done) n[stepId] = true; else delete n[stepId]; store.set(CAREERKEYS.path, n); return n; });
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

      {tab === "path" && <CareerPath progress={pathProg} onDone={setStepDone} />}

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
    "intro": "Пройшовши цей шлях, ти зможеш впевнено вести продукт від ідеї до MVP: досліджувати потреби користувачів, писати PRD і user stories, пріоритизувати фічі й будувати роадмапу — усе, що потрібно технічному PM.",
    "steps": [
      {
        "type": "learn",
        "title": "Хто такий PM і що таке продукт",
        "body": "Продакт-менеджер (PM) відповідає за те, ЩО і ЧОМУ команда будує, а не за те, ЯК це кодиться. Його робота — знайти реальну проблему користувача, вирішити, що будувати в першу чергу, і донести це до команди. PM стоїть на перетині трьох сил: бізнес (чи це вигідно), користувач (чи це потрібно) і технології (чи це реально зробити). Наприклад, замість «додаймо чат-бота, бо це модно» PM питає «яку проблему користувача це вирішує і як ми виміряємо успіх». Для технічного PM це особливо важливо: розуміння технологій допомагає ставити реалістичні задачі, але фокус завжди на цінності для користувача."
      },
      {
        "type": "learn",
        "title": "Discovery та інтерв'ю з користувачами",
        "body": "Discovery (дискавері) — це етап досліджень ПЕРЕД тим, як щось будувати, щоб переконатися, що проблема справжня. Головний інструмент — глибинне інтерв'ю з користувачами: розмова 1-на-1, де ти слухаєш про їхній реальний досвід, а не питаєш «чи сподобалась би вам така фіча». Ключове правило: питай про минуле й теперішнє («розкажи, як ти востаннє це робив»), а не про гіпотетичне майбутнє — люди погано прогнозують свою поведінку. Уникай навідних питань і не продавай своє рішення. Мета discovery — зменшити ризик побудувати те, що нікому не потрібно.",
        "resource": {
          "label": "The Mom Test — офіційний сайт книги про інтерв'ю з користувачами",
          "url": "https://www.momtestbook.com/"
        }
      },
      {
        "type": "quiz",
        "title": "Перевірка: інтерв'ю",
        "body": "Яке питання найкраще підходить для інтерв'ю з користувачем?",
        "quiz": {
          "question": "Яке питання дасть найнадійнішу інформацію під час discovery-інтерв'ю?",
          "options": [
            "Розкажіть, як ви востаннє вирішували цю задачу?",
            "Чи купили б ви нашу нову фічу за 10 доларів?",
            "Вам би сподобався додаток із такою функцією?",
            "Скільки б ви платили за ідеальне рішення?"
          ],
          "answerIndex": 0,
          "explanation": "Питання про реальний минулий досвід дає факти, а не здогади; решта — навідні або гіпотетичні."
        }
      },
      {
        "type": "learn",
        "title": "JTBD і user stories",
        "body": "JTBD (Jobs To Be Done) — підхід, за яким люди «наймають» продукт, щоб виконати певну «роботу» у своєму житті. Класика: люди купують не дриль, а «дірку в стіні» для полиці. Формула роботи: «Коли [ситуація], я хочу [мотивація], щоб [результат]». Коли JTBD відома, її перекладають у user stories — короткі описи фіч за формулою «Як [роль], я хочу [дію], щоб [цінність]». Наприклад: «Як новий користувач, я хочу увійти через Google, щоб не створювати ще один пароль». Сторі фіксує потребу й цінність, а не технічну реалізацію, і має критерії приймання — умови, за яких вона вважається виконаною."
      },
      {
        "type": "explain",
        "title": "Поясни JTBD своїми словами",
        "body": "Поясни своїми словами, чим JTBD («робота») відрізняється від фічі, і як вона перетворюється на user story. Хороша відповідь: наводить приклад за формулою «Коли…, я хочу…, щоб…» і показує, що продукт «наймають» заради результату, а не заради самої функції."
      },
      {
        "type": "read",
        "title": "Пріоритизація: RICE і MoSCoW",
        "body": "Ідей завжди більше, ніж ресурсів, тому PM мусить пріоритизувати. RICE рахує бал за формулою (Reach × Impact × Confidence) / Effort: Reach — скільки людей це зачепить за період, Impact — наскільки сильно (напр. 3=масово, 1=слабо), Confidence — твоя впевненість у оцінках (у %), Effort — трудовитрати в людино-місяцях. Що вищий бал — то раніше береш фічу. MoSCoW простіший: розкладаєш задачі на Must have (без цього реліз не відбудеться), Should have (важливо, але можна відкласти), Could have (приємно мати) і Won't have (не зараз). RICE добре працює для порівняння багатьох фіч за числами, MoSCoW — для швидкого узгодження обсягу релізу з командою.",
        "resource": {
          "label": "Intercom: стаття про RICE-фреймворк",
          "url": "https://www.intercom.com/blog/rice-simple-prioritization-for-product-managers/"
        }
      },
      {
        "type": "quiz",
        "title": "Перевірка: RICE",
        "body": "Як зміниться пріоритет фічі в RICE?",
        "quiz": {
          "question": "У формулі RICE = (Reach × Impact × Confidence) / Effort — що станеться з балом фічі, якщо Effort зросте, а решта не зміниться?",
          "options": [
            "Бал зменшиться, тож пріоритет фічі впаде",
            "Бал зросте, тож пріоритет фічі підвищиться",
            "Бал не зміниться, бо Effort не впливає",
            "Фіча автоматично стане Must have"
          ],
          "answerIndex": 0,
          "explanation": "Effort у знаменнику: більше зусиль при тій самій користі — нижчий бал і нижчий пріоритет."
        }
      },
      {
        "type": "learn",
        "title": "MVP, PRD і Roadmap: від ідеї до плану",
        "body": "MVP (мінімально життєздатний продукт) — найменша версія продукту, що вже дає цінність користувачу й дозволяє перевірити ключову гіпотезу, не будуючи все одразу. PRD (Product Requirements Document) — документ, який описує проблему, цільового користувача, мету, обсяг (scope), user stories й метрики успіху; він синхронізує команду щодо того, ЩО будуємо і ЧОМУ. Roadmap (роадмапа) — стратегічний план у часі: що й приблизно коли команда планує зробити, згрупована навколо цілей, а не список точних дат. Разом це логічний ланцюг: discovery виявляє проблему → PRD її фіксує → пріоритизація визначає порядок → MVP перевіряє гіпотезу → roadmap показує рух далі."
      },
      {
        "type": "build",
        "title": "Мілстоун: міні-PRD для MVP",
        "body": "Фінальне завдання — збери все разом. Напиши міні-PRD на одну сторінку для MVP застосунку замовлення кави: (1) Проблема й цільовий користувач (спираючись на JTBD), (2) Мета й 1-2 метрики успіху, (3) Scope MVP — 3 фічі, які увійдуть, і 2, які свідомо НЕ увійдуть, (4) Пріоритизуй ці 3 фічі через RICE або MoSCoW, (5) Додай 2-3 user stories за формулою «Як [роль], я хочу [дію], щоб [цінність]». Це твій перший повний продуктовий документ — вітаю з проходженням модуля."
      }
    ]
  },
  {
    "slug": "tech-depth",
    "emoji": "⚙️",
    "title": "Технічна глибина",
    "intro": "Пройшовши цей шлях, ти впевнено говоритимеш з інженерами: розумітимеш, як працює веб, API, бази даних і базовий system design — і ставитимеш точні запитання на будь-якому дизайн-рев'ю.",
    "steps": [
      {
        "type": "learn",
        "title": "Клієнт і сервер",
        "body": "Веб працює за моделлю клієнт-сервер. Клієнт (браузер чи мобільний застосунок) надсилає запит, а сервер його обробляє й повертає відповідь. Наприклад, коли ти відкриваєш профіль у застосунку, клієнт просить сервер: «дай мені дані користувача 42», сервер бере їх із бази й відправляє назад. Клієнт відповідає за те, що бачить людина (UI), а сервер — за логіку й дані. Для PM це базова карта: майже кожна фіча — це діалог між клієнтом і сервером, і розуміння, де саме «живе» логіка, допомагає оцінювати складність."
      },
      {
        "type": "learn",
        "title": "Що таке HTTP",
        "body": "HTTP — це протокол (набір правил), яким клієнт і сервер обмінюються повідомленнями у вебі. Кожна взаємодія — це запит (request) від клієнта і відповідь (response) від сервера. Запит містить метод (що робити), адресу (URL) і часто тіло з даними; відповідь містить статус-код і дані. HTTPS — це той самий HTTP, але зашифрований, тому дані не можна перехопити. Для PM важливо: коли інженер каже «повільний респонс» або «падає на 500-ці», йдеться саме про цей обмін запит-відповідь."
      },
      {
        "type": "learn",
        "title": "API та REST",
        "body": "API — це «меню» можливостей, які сервер надає іншим програмам: набір операцій, які можна викликати. REST — популярний стиль побудови API, де кожен ресурс має свою адресу (endpoint), наприклад /users/42. Дії задаються HTTP-методами: GET (отримати дані), POST (створити), PUT/PATCH (оновити), DELETE (видалити). Так замість «зайди на сайт руками» одна система може автоматично запитати дані в іншої. Для PM API — це часто і продукт (те, що продаємо партнерам), і спосіб оцінити, чи фіча вимагає нового endpoint, чи вистачить наявного."
      },
      {
        "type": "quiz",
        "title": "Методи HTTP",
        "body": "Перевіримо методи REST.",
        "quiz": {
          "question": "Команда хоче додати можливість створити нове замовлення через API. Який HTTP-метод це, найімовірніше, буде?",
          "options": [
            "GET",
            "POST",
            "DELETE",
            "HEAD"
          ],
          "answerIndex": 1,
          "explanation": "POST використовують для створення нового ресурсу; GET лише читає дані, DELETE видаляє."
        }
      },
      {
        "type": "learn",
        "title": "Статус-коди та JSON",
        "body": "У відповідь на кожен запит сервер повертає статус-код — трицифрове число про результат. Головні групи: 2xx — успіх (200 OK, 201 Created), 4xx — помилка клієнта (404 Not Found — не знайдено, 401 — не авторизований, 403 — заборонено), 5xx — помилка сервера (500 Internal Server Error). Самі дані найчастіше передаються у форматі JSON — це текст із пар «ключ: значення», наприклад {\"id\": 42, \"name\": \"Оля\"}. JSON легко читає і людина, і програма. Для PM це прямий інструмент: у логах чи в інструментах на кшталт Postman ти сама побачиш, 4xx це чи 5xx, і зрозумієш, чий це баг — клієнта чи сервера."
      },
      {
        "type": "quiz",
        "title": "Читаємо статус-коди",
        "body": "Класика, яку варто знати напам'ять.",
        "quiz": {
          "question": "Користувачі скаржаться, що сторінка віддає помилку 500. Про що це насамперед сигналізує?",
          "options": [
            "Користувач ввів неправильну адресу",
            "Сталася помилка на боці сервера",
            "Все добре, запит успішний",
            "Користувач не авторизований"
          ],
          "answerIndex": 1,
          "explanation": "5xx — це помилки сервера; неправильна адреса дала б 404, а неавторизований доступ — 401/403."
        }
      },
      {
        "type": "read",
        "title": "Бази даних: SQL vs NoSQL",
        "body": "База даних — це де застосунок зберігає дані надовго. SQL-бази (PostgreSQL, MySQL) зберігають дані в таблицях зі суворою структурою (рядки й колонки) і чудові там, де важливі зв'язки й точність — фінанси, замовлення. NoSQL-бази (MongoDB, Redis) гнучкіші: зберігають документи, ключ-значення чи інші формати, добре масштабуються й підходять для великих обсягів простих даних чи швидкої зміни структури. Спрощено: SQL — про порядок і зв'язки, NoSQL — про гнучкість і масштаб. PM не обирає базу сам, але має розуміти компроміс, коли інженер каже «тут нам потрібні транзакції» або «дані надто розкидані для реляційної моделі».",
        "resource": {
          "label": "MongoDB: SQL vs NoSQL (просте пояснення)",
          "url": "https://www.mongodb.com/resources/basics/databases/nosql-explained"
        }
      },
      {
        "type": "explain",
        "title": "Поясни своїми словами",
        "body": "Поясни своїми словами, що відбувається «під капотом», коли ти в застосунку натискаєш «Зберегти профіль». Хороша відповідь проходить весь ланцюжок: клієнт формує HTTP-запит (метод POST/PUT) з даними у JSON → сервер приймає його через API-endpoint → перевіряє й записує в базу даних → повертає статус-код (напр. 200) і відповідь клієнту."
      },
      {
        "type": "build",
        "title": "System design: спроєктуй фічу",
        "body": "Візьми знайому фічу — стрічку новин у застосунку — і застосуй два ключові інструменти масштабування. Кеш (напр. Redis) — це швидке тимчасове сховище «під рукою»: часто запитувані дані тримають тут, щоб не смикати повільну базу щоразу (як пам'ятати відповідь, а не гуглити знову). Черга (напр. Kafka, RabbitMQ) — для задач, які не треба робити миттєво: запит кладуть у чергу, і фоновий процес обробляє його потім (так надсилають листи чи сповіщення). Твоє завдання, 3-4 речення: 1) назви одні дані у стрічці, які варто кешувати, і який ризик «застарілих» даних це створює; 2) назви одну дію (напр. розсилку сповіщень підписникам), яку варто винести в чергу, і що від цього виграє користувач. Це готова чернетка твого аргументу на дизайн-рев'ю."
      }
    ]
  },
  {
    "slug": "ai-literacy",
    "emoji": "🤖",
    "title": "AI-грамотність",
    "intro": "Пройшовши цей шлях, ти впевнено говоритимеш мовою AI-команди: зрозумієш, як працюють LLM, скільки коштує кожен виклик і як не запустити в продакшн модель, що вигадує факти.",
    "steps": [
      {
        "type": "learn",
        "title": "Що таке LLM і токени",
        "body": "LLM (велика мовна модель, як GPT-4 чи Claude) — це модель, натренована передбачати наступний фрагмент тексту на основі попереднього. Вона не «розуміє» світ, а статистично вгадує найімовірніше продовження. Текст модель бачить не як слова, а як токени — шматочки по ~3-4 символи; наприклад «промпт-інжиніринг» може розбитися на 5-6 токенів. Це критично для PM, бо і ліміти, і ціна рахуються саме в токенах, а не в словах чи символах. Груба оцінка: 1000 токенів ≈ 750 англійських слів (для української — менше, бо кирилиця «дорожча»)."
      },
      {
        "type": "quiz",
        "title": "Перевірка: токени",
        "body": "Швидка перевірка розуміння токенів.",
        "quiz": {
          "question": "У чому вимірюються і ліміти, і вартість роботи з LLM?",
          "options": [
            "У словах",
            "У токенах",
            "У символах",
            "У реченнях"
          ],
          "answerIndex": 1,
          "explanation": "Модель обробляє текст як токени, тому і контекстні ліміти, і біллінг рахуються саме в них."
        }
      },
      {
        "type": "learn",
        "title": "Контекстне вікно і температура",
        "body": "Контекстне вікно — це максимальна кількість токенів, яку модель тримає «в голові» за один виклик: і твій промпт (вхід), і її відповідь (вихід) разом. Якщо діалог перевищує вікно (наприклад 128k токенів ≈ велика книжка), старіші токени «випадають» і модель їх більше не бачить. Температура — це параметр (зазвичай 0-1), що керує випадковістю: низька (0-0.3) дає передбачувані, стабільні відповіді (класифікація, витяг даних, код), висока (0.7-1) — різноманіття й креатив (брейнштормінг, тексти). Важливо: температура не впливає на правдивість — модель може впевнено помилятися і на 0. Як PM, для фічі «витягни суму з рахунка» ти береш низьку, а для «запропонуй назви продукту» — вищу."
      },
      {
        "type": "learn",
        "title": "Промпт-інжиніринг",
        "body": "Промпт-інжиніринг — це мистецтво так сформулювати інструкцію, щоб модель дала потрібний результат. Базові прийоми: чітко задати роль і задачу, дати приклади бажаного формату (few-shot), попросити міркувати покроково для складних задач і явно описати формат виходу (наприклад, «поверни JSON з полями name і price»). Поганий промпт «напиши про продукт» дає розмите есе; хороший «Ти копірайтер. Напиши 3 варіанти заголовка (до 60 символів) для лендингу B2B-CRM, тон діловий» дає готовий до використання результат. Для PM це найдешевший важіль якості — часто кращий промпт вирішує проблему без дотренування моделі."
      },
      {
        "type": "build",
        "title": "Напиши структурований промпт",
        "body": "Візьми уявну фічу «AI-помічник підсумовує відгуки користувачів». Напиши для неї промпт із чотирма елементами: (1) роль моделі, (2) конкретна задача, (3) один приклад входу й бажаного виходу, (4) явний формат виходу (наприклад, JSON з полями `summary`, `sentiment`, `top_issues`). Мета — щоб інженер міг взяти твій промпт і вставити в код майже без змін."
      },
      {
        "type": "read",
        "title": "Embeddings і RAG",
        "body": "Embedding — це перетворення тексту на вектор чисел, де близькі за змістом фрази опиняються поруч у просторі. Це дозволяє шукати за сенсом, а не за точним збігом слів: запит «як повернути кошти» знайде документ «політика рефандів». RAG (Retrieval-Augmented Generation) будується на цьому: коли надходить питання, система спершу знаходить релевантні шматки твоїх документів (через embeddings + векторну базу), а потім вкладає їх у промпт як контекст, і вже тоді модель відповідає. Так LLM відповідає на основі твоїх актуальних даних, а не лише того, що «пам'ятає» з тренування — і рідше вигадує. Для PM RAG — стандартний спосіб зробити «чат із нашою документацією» без дорогого перетренування моделі.",
        "resource": {
          "label": "OpenAI: What are embeddings (гайд)",
          "url": "https://developers.openai.com/api/docs/guides/embeddings"
        }
      },
      {
        "type": "explain",
        "title": "Поясни RAG своїми словами",
        "body": "Поясни своїми словами, як працює RAG і навіщо він потрібен — так, ніби розповідаєш колезі-дизайнеру. Хороша відповідь торкається трьох речей: (1) спершу система знаходить релевантні фрагменти твоїх документів через пошук за змістом (embeddings), (2) потім підкладає їх у промпт як контекст, (3) і завдяки цьому модель відповідає на основі актуальних даних і менше галюцинує."
      },
      {
        "type": "learn",
        "title": "Галюцинації, evals, безпека й вартість",
        "body": "Галюцинація — це коли модель впевнено видає вигадку за факт (неіснуюче джерело, фейкову цифру); це не баг, а наслідок того, що вона передбачає правдоподібний текст, а не перевіряє істину. Тому для будь-якої фічі з фактами потрібні evals — набір тестів, що вимірюють якість відповідей на прикладах, як юніт-тести для моделі. Безпека: не клади в промпт зайві персональні дані, зважай на prompt injection (коли зловмисний текст у вхідних даних перехоплює інструкцію) і тримай людину в контурі для важливих рішень. Вартість: платиш за токени входу + виходу, тож довгі промпти й великий контекст = дорожче; RAG, коротші промпти й дешевші моделі для простих задач помітно ріжуть рахунок."
      },
      {
        "type": "quiz",
        "title": "Мілстоун: збери все разом",
        "body": "Фінальна перевірка ключових понять шляху.",
        "quiz": {
          "question": "Команда будує «чат із базою знань компанії». Що з переліченого найкраще зменшує галюцинації й тримає відповіді актуальними, не перетреновуючи модель?",
          "options": [
            "Підняти температуру до максимуму",
            "Використати RAG: підвантажувати релевантні документи в контекст перед відповіддю",
            "Просто взяти більшу модель",
            "Прибрати evals, щоб пришвидшити реліз"
          ],
          "answerIndex": 1,
          "explanation": "RAG підкладає актуальні документи у промпт, тож модель відповідає на основі реальних даних і рідше вигадує — без дорогого перетренування."
        }
      }
    ]
  },
  {
    "slug": "build-ai",
    "emoji": "🛠️",
    "title": "Будувати з AI",
    "intro": "Пройшовши цей шлях, ти зможеш власноруч зібрати й показати одну робочу AI-фічу — від ключа API до демо, яке не соромно показати команді.",
    "steps": [
      {
        "type": "learn",
        "title": "Що таке LLM API",
        "body": "API — це спосіб, яким твоя програма «розмовляє» з моделлю (наприклад, GPT від OpenAI чи Claude від Anthropic) через інтернет. Ти надсилаєш запит із текстом (промпт), а модель повертає відповідь — теж текстом. Три ключові речі в кожному виклику: API-ключ (твій секретний пароль для доступу), запит (що ти питаєш) і відповідь (що модель згенерувала). Для PM це важливо, бо саме тут народжується AI-фіча: ти описуєш поведінку словами, а не кодом. Розуміючи механіку виклику, ти можеш реалістично оцінювати, що модель здатна зробити, скільки це коштує і де межі."
      },
      {
        "type": "learn",
        "title": "API-ключ і безпека",
        "body": "API-ключ — це рядок символів, який ідентифікує тебе перед сервісом і прив'язаний до оплати. Ти отримуєш його в кабінеті розробника (platform.openai.com або console.anthropic.com) і вставляєш у свою програму. Головне правило: ключ ніколи не можна публікувати — ні в коді на GitHub, ні у фронтенді, ні в скріншотах. Якщо ключ витік, будь-хто може витрачати твої гроші, тому його зберігають у змінних середовища (environment variables) або секретах. Для PM це не дрібниця: витік ключа — реальний інцидент безпеки й неконтрольовані витрати.",
        "resource": {
          "label": "OpenAI: Production best practices (API keys)",
          "url": "https://developers.openai.com/api/docs/guides/production-best-practices"
        }
      },
      {
        "type": "quiz",
        "title": "Перевірка: API-ключ",
        "body": "Де НЕ можна зберігати API-ключ?",
        "quiz": {
          "question": "Ти інтегруєш AI-фічу. Куди точно НЕ можна класти API-ключ?",
          "options": [
            "У змінну середовища на сервері",
            "У код фронтенду, який бачить браузер користувача",
            "У менеджер секретів хостингу",
            "У локальний файл .env, який не потрапляє в git"
          ],
          "answerIndex": 1,
          "explanation": "Усе, що потрапляє у фронтенд, видно користувачу — ключ звідти вкрадуть; виклики до LLM роблять із бекенду."
        }
      },
      {
        "type": "learn",
        "title": "Анатомія запиту й відповіді",
        "body": "Типовий запит до LLM містить кілька частин: модель (наприклад, gpt-4o чи claude-sonnet), повідомлення (роль system задає поведінку, роль user — саме запит користувача) і параметри на кшталт temperature (наскільки «творчою» буде відповідь) та max_tokens (обмеження довжини). Токен — це шматочок тексту (приблизно 4 символи), і ти платиш за токени і на вході, і на виході. Відповідь повертається структуровано (зазвичай JSON), де сам згенерований текст лежить у полі content. Розуміючи це, PM може писати чіткіші вимоги: «system-промпт задає тон», «обмеж відповідь 200 токенами», «temperature низька для фактів»."
      },
      {
        "type": "learn",
        "title": "No-code інструменти",
        "body": "Щоб зібрати AI-фічу, PM не обов'язково писати код із нуля — є no-code та low-code інструменти. Zapier і Make дозволяють з'єднати тригер (нове повідомлення, рядок у таблиці) із викликом LLM і дією без програмування. Bubble чи Retool допомагають зібрати простий інтерфейс поверх API. Такі інструменти ідеальні для прототипу або внутрішнього демо: ти перевіряєш ідею за години, а не тижні, і показуєш команді щось клікабельне. Обмеження — менше контролю й гнучкості, ніж у власному коді, але для валідації гіпотези цього зазвичай достатньо.",
        "resource": {
          "label": "Zapier: AI",
          "url": "https://zapier.com/ai"
        }
      },
      {
        "type": "read",
        "title": "Промпт як специфікація",
        "body": "Для звичайного софту специфікацію пишуть у документі, а розробник її кодує. З LLM специфікація і є промптом: те, що ти напишеш словами, безпосередньо стає поведінкою фічі. Хороший промпт-специфікація має роль («Ти — асистент підтримки»), чіткі інструкції («відповідай лише про наш продукт»), формат виводу («поверни список із 3 пунктів»), обмеження («не вигадуй фактів») і приклади бажаної відповіді (few-shot). Це означає, що PM може прототипувати й уточнювати поведінку продукту напряму, без розробника — але й відповідальність за якість вимог тепер на тобі. Версіонуй промпти, як код: маленька зміна формулювання може помітно змінити результат.",
        "resource": {
          "label": "Anthropic: Prompt engineering overview",
          "url": "https://platform.claude.com/docs/en/docs/build-with-claude/prompt-engineering/overview"
        }
      },
      {
        "type": "explain",
        "title": "Поясни: промпт як специфікація",
        "body": "Поясни своїми словами, чому промпт для LLM можна вважати «специфікацією продукту». Гарна відповідь торкається того, що в промпті ти словами описуєш бажану поведінку (роль, тон, формат, обмеження, приклади), і саме він визначає, що фіча робитиме — тобто це продуктова вимога, а не просто текст."
      },
      {
        "type": "build",
        "title": "Напиши промпт-специфікацію",
        "body": "Обери одну маленьку AI-фічу (наприклад, «бот, що коротко переказує відгуки клієнтів»). Напиши для неї повний промпт-специфікацію з 5 частин: (1) роль моделі, (2) чітка інструкція завдання, (3) формат виводу, (4) 1-2 обмеження («не вигадуй», «максимум 3 речення»), (5) один приклад «вхід → бажаний вихід». Це і буде технічна вимога твоєї фічі."
      },
      {
        "type": "quiz",
        "title": "Перевірка: параметри",
        "body": "Що робить параметр temperature?",
        "quiz": {
          "question": "Тобі потрібні максимально стабільні, фактичні відповіді без вигадок. Яке значення temperature обрати?",
          "options": [
            "Високе (близько 1.0), щоб модель була креативнішою",
            "Низьке (близько 0), щоб відповіді були передбачуванішими",
            "Temperature не впливає на стабільність, лише на швидкість",
            "Максимальне, бо так модель точніша"
          ],
          "answerIndex": 1,
          "explanation": "Низька temperature робить вибір слів детермінованішим — менше варіативності й вигадок; висока додає креативу й ризику."
        }
      }
    ]
  },
  {
    "slug": "analytics",
    "emoji": "📊",
    "title": "Аналітика й метрики",
    "intro": "Пройшовши цей шлях, ти навчишся ставити метрики так, як це робить сильний технічний PM: обереш North-star, читатимеш воронку AARRR, самостійно витягнеш дані з бази через SQL і зрозумієш, чи справді твій A/B-тест переміг.",
    "steps": [
      {
        "type": "learn",
        "title": "North-star метрика",
        "body": "North-star (провідна) метрика — це одне число, що найкраще відображає цінність, яку продукт дає користувачам, і навколо якого команда узгоджує свої рішення. Хороша NSM показує саме отриману цінність, а не просто активність: у Spotify це не «кількість реєстрацій», а «час прослуховування», бо саме він відображає задоволення. Її обирають так, щоб зростання метрики тягнуло за собою й зростання бізнесу. Погана NSM легко «накручується» без реальної користі (наприклад, кількість кліків). Для PM це компас: коли є суперечка про пріоритети, питаєш «що з цього більше зрушить North-star?»."
      },
      {
        "type": "quiz",
        "title": "Перевірка: NSM",
        "body": "Обери найкращу North-star метрику для застосунку доставки їжі.",
        "quiz": {
          "question": "Яка метрика найкраще підходить як North-star для сервісу доставки їжі?",
          "options": [
            "Кількість завантажень застосунку за місяць",
            "Кількість успішно доставлених замовлень на активного користувача",
            "Кількість показів банерів у застосунку",
            "Кількість натискань на кнопку «Кошик»"
          ],
          "answerIndex": 1,
          "explanation": "Доставлені замовлення відображають реальну отриману цінність; завантаження, покази й кліки легко ростуть без користі для клієнта."
        }
      },
      {
        "type": "learn",
        "title": "Воронка AARRR",
        "body": "AARRR («піратські метрики») — це п'ять стадій життя користувача: Acquisition (як він тебе знайшов), Activation (перший успішний досвід, «ага-момент»), Retention (чи повертається), Referral (чи радить іншим), Revenue (чи платить). Ідея в тому, щоб міряти конверсію між сусідніми стадіями й шукати, де найбільше «протікає». Наприклад, якщо 1000 людей зареєструвались, але лише 200 дійшли до активації — це твоє вузьке місце. Порядок пріоритетів зазвичай: спершу лагодять Retention і Activation, бо без них лити трафік (Acquisition) — марно."
      },
      {
        "type": "learn",
        "title": "Activation і Retention",
        "body": "Activation — це момент, коли новий користувач уперше отримує ключову цінність продукту (наприклад, у Facebook історично це «7 друзів за 10 днів»). Retention — частка користувачів, що повертаються через певний час; його часто дивляться як криву утримання по когортах (день 1, день 7, день 30). Здорова retention-крива з часом виходить на плато, а не падає в нуль — це ознака product-market fit. Сильна активація майже завжди піднімає утримання, бо людина швидше зрозуміла, навіщо їй продукт. PM визначає чіткий «момент активації» через дані й будує онбординг так, щоб довести до нього якомога більше новачків."
      },
      {
        "type": "explain",
        "title": "Поясни своїми словами",
        "body": "Поясни своїми словами різницю між activation і retention та чому підвищувати retention зазвичай важливіше, ніж лити більше трафіку. Хороша відповідь згадує: активація — це перший «ага-момент» цінності, retention — повернення користувачів у часі, а без утримання новий трафік просто «витікає» з дірявої воронки."
      },
      {
        "type": "read",
        "title": "Базовий SQL для PM",
        "body": "SQL дозволяє PM самостійно діставати відповіді з бази, не чекаючи аналітика. Чотири основні частини: SELECT (які стовпці показати), WHERE (фільтр рядків за умовою), GROUP BY (згрупувати рядки й порахувати агрегати на кшталт COUNT/SUM/AVG), JOIN (з'єднати дві таблиці за спільним ключем). Приклад: `SELECT country, COUNT(*) FROM users WHERE created_at >= '2026-01-01' GROUP BY country` порахує нових користувачів по країнах. JOIN потрібен, коли дані розкидані по таблицях: наприклад, з'єднати `users` з `orders` за `user_id`. Важливо: WHERE фільтрує окремі рядки ще до групування, а не готові групи.",
        "resource": {
          "label": "SQLBolt — інтерактивні уроки SQL (безкоштовно)",
          "url": "https://sqlbolt.com"
        }
      },
      {
        "type": "build",
        "title": "Напиши свій SQL",
        "body": "Уяви таблицю `events(user_id, event_name, created_at)`. Напиши SQL-запит, який покаже, скільки унікальних користувачів зробили подію 'purchase' у липні 2026 року. Підказка: використай WHERE для фільтра за event_name і датою, а COUNT(DISTINCT user_id) — щоб не рахувати одного користувача двічі. Спробуй сам перед тим, як дивитись приклад: `SELECT COUNT(DISTINCT user_id) FROM events WHERE event_name = 'purchase' AND created_at >= '2026-07-01' AND created_at < '2026-08-01'`."
      },
      {
        "type": "learn",
        "title": "A/B-тести й статзначущість",
        "body": "A/B-тест — це експеримент, де користувачів випадково ділять на групу A (контроль) і групу B (нова версія), щоб чесно порівняти вплив зміни. Статистична значущість (часто p-value < 0.05) відповідає на питання «чи різниця між групами реальна, чи це просто випадковість?». Щоб помітити невеликий ефект, потрібен достатній розмір вибірки, інакше тест нічого не покаже. Часта помилка PM — «підглядати» в результати й зупиняти тест, щойно бачиш перемогу: це роздуває хибнопозитивні результати. І пам'ятай: статистична значущість не дорівнює практичній — приріст на 0.1% може бути «значущим», але не вартим впровадження."
      },
      {
        "type": "quiz",
        "title": "Мілстоун: читаємо дашборд",
        "body": "Фінальна перевірка, що збирає все разом. NSM впала на 8%, а паралельний A/B-тест дав p-value = 0.5. Як діяти?",
        "quiz": {
          "question": "На дашборді North-star впала на 8%; воронка показує стабільні Acquisition і Activation, але просів Retention день-30; A/B-тест нового онбордингу має p-value = 0.5. Який перший крок найлогічніший?",
          "options": [
            "Зупинити A/B-тест як причину падіння North-star",
            "Копати в Retention день-30 по когортах, бо саме там «протікає» воронка",
            "Негайно впровадити версію B онбордингу на всіх",
            "Збільшити бюджет на Acquisition, щоб компенсувати падіння"
          ],
          "answerIndex": 1,
          "explanation": "Acquisition і Activation стабільні, а Retention просів — це вузьке місце; A/B-тест з p-value = 0.5 незначущий, тож не є ні причиною, ні готовим рішенням."
        }
      }
    ]
  },
  {
    "slug": "stakeholders",
    "emoji": "🤝",
    "title": "Стейкхолдери й комунікація",
    "intro": "Пройшовши цей шлях, ти навчишся впевнено презентувати roadmap, вирівнювати founder, інженерів і дизайн навколо спільної мети та перетворювати конфлікти пріоритетів на ясні, задокументовані рішення.",
    "steps": [
      {
        "type": "learn",
        "title": "Хто такі стейкхолдери",
        "body": "Стейкхолдер — це будь-хто, чиї інтереси зачіпає твій продукт або хто впливає на рішення щодо нього: founder, інженери, дизайнери, продажі, підтримка, юристи, а також самі користувачі. Для PM це не просто «начальники», а люди з різними цілями, які часто конфліктують: founder хоче швидкості, інженер — якості коду, дизайнер — цілісного досвіду. Твоя робота — не догодити всім, а зробити ці інтереси видимими й узгодити їх навколо спільної мети продукту. Простий інструмент — карта стейкхолдерів за двома осями: рівень впливу та рівень зацікавленості. Тих, у кого високий вплив і висока зацікавленість, тримай найближче: залучай рано й часто."
      },
      {
        "type": "learn",
        "title": "Доносити «чому», а не лише «що»",
        "body": "Слабкий PM каже команді «робимо фічу X до п'ятниці». Сильний PM спершу пояснює «чому»: яку проблему користувача чи бізнес-ціль ми вирішуємо і як зрозуміємо, що вдалося. Коли команда розуміє «чому», вона ухвалює кращі мікрорішення без тебе й не потребує мікроменеджменту. Це також будує довіру: люди підтримують те, у творенні сенсу чого брали участь. Практичне правило — формулюй мету через результат для користувача та метрику, а не через список завдань: не «додати онбординг-екран», а «підняти активацію нових користувачів з 40% до 55%, бо вони губляться на першому кроці»."
      },
      {
        "type": "quiz",
        "title": "Перевірка: чому «чому»",
        "body": "Обери найсильніше формулювання цілі для команди.",
        "quiz": {
          "question": "Яке формулювання задачі для команди найкраще доносить «чому»?",
          "options": [
            "«Додаємо онбординг-екран до п'ятниці»",
            "«Робимо так, як просив founder»",
            "«Піднімаємо активацію нових користувачів з 40% до 55%, бо вони губляться на першому кроці»",
            "«Треба закрити багато тікетів цього спринту»"
          ],
          "answerIndex": 2,
          "explanation": "Лише цей варіант називає проблему користувача та вимірювану ціль, а не просто список завдань чи джерело наказу."
        }
      },
      {
        "type": "read",
        "title": "RACI: хто за що відповідає",
        "body": "RACI — матриця ролей для задачі чи рішення за чотирма буквами. R (Responsible) — хто безпосередньо виконує роботу. A (Accountable) — хто ухвалює фінальне рішення й відповідає за результат; таких має бути рівно один, інакше рішення зависають. C (Consulted) — з ким радимося до рішення (двобічний діалог). I (Informed) — кого просто повідомляємо після (однобічно). Класична пастка — двоє «A» на одну задачу або плутанина між C та I: люди ображаються, коли їх поставили в «I», хоча очікували, що з ними порадяться. RACI особливо рятує в крос-командних рішеннях, де незрозуміло, чиє останнє слово.",
        "resource": {
          "label": "Atlassian: RACI charts",
          "url": "https://www.atlassian.com/work-management/project-management/raci-chart"
        }
      },
      {
        "type": "quiz",
        "title": "Перевірка: RACI",
        "body": "Пригадай значення літер RACI.",
        "quiz": {
          "question": "Скільки осіб має бути в ролі «Accountable» (A) для одного рішення і чому?",
          "options": [
            "Якнайбільше — щоб відповідальність була спільною",
            "Рівно одна — щоб було зрозуміло, за ким фінальне слово",
            "Жодної — рішення ухвалює вся команда голосуванням",
            "Стільки ж, скільки й Responsible"
          ],
          "answerIndex": 1,
          "explanation": "Accountable має бути рівно один: єдина точка відповідальності не дає рішенням зависати."
        }
      },
      {
        "type": "build",
        "title": "Накидай RACI-матрицю",
        "body": "Візьми реальне рішення: «Обрати платіжного провайдера для нової підписки». Випиши 4-5 стейкхолдерів (напр. PM, ліда інженерії, дизайнера, фінансиста, founder) і для кожного признач рівно одну роль R, A, C або I. Перевір себе: чи є рівно один A? Чи не сплутав ти тих, з ким радишся (C), з тими, кого лише повідомляєш (I)? Запиши одним рядком, чому саме цей стейкхолдер отримав A."
      },
      {
        "type": "learn",
        "title": "Презентувати roadmap",
        "body": "Roadmap — це не список фіч із датами, а розповідь про напрямок: які проблеми ми вирішуємо найближчим часом і чому саме в такому порядку. Групуй роботу за темами чи цілями (наприклад «Now / Next / Later»), а не за точними дедлайнами — це знижує тиск обіцянок і лишає простір для навчання. Під різну аудиторію — різний рівень деталізації: founder-у покажи зв'язок зі стратегією та метриками, інженерам — технічні залежності, продажам — що зможуть обіцяти клієнтам. Завжди готуйся до питання «а чому не X раніше?»: тримай видимими критерії пріоритезації, щоб roadmap виглядав як усвідомлений вибір, а не випадковий список."
      },
      {
        "type": "learn",
        "title": "Керувати конфліктом пріоритетів",
        "body": "Конфлікт пріоритетів — норма, а не збій: у founder, продажів та інженерії різні цілі, і всі вони хочуть «своє першим». Твоя сила PM — перевести суперечку з площини думок і статусів у площину спільних критеріїв. Домовтеся про рамку пріоритезації (напр. RICE: Reach, Impact, Confidence, Effort) і оцінюйте ідеї за нею відкрито — тоді сперечаються з цифрами, а не з людьми. Роби компроміс явним: «беремо A зараз, B — наступним; ось на що це впливає». Найгірше рішення — тихо намагатися вмістити все: команда вигорає, а довіра падає, бо ніхто не бачить логіки вибору."
      },
      {
        "type": "explain",
        "title": "Поясни своїми словами",
        "body": "Уяви, що founder і лід продажів у чаті вимагають протилежних фіч «на вчора». Поясни своїми словами, як ти, PM, вирівняєш їх навколо спільного рішення. Хороша відповідь згадує: спершу винести на поверхню «чому» кожної сторони (яка мета за проханням), застосувати спільну рамку пріоритезації (напр. RICE) відкрито, і зробити компроміс та його наслідки явними й задокументованими — щоб рішення виглядало як спільний вибір, а не перемога одного над іншим."
      }
    ]
  },
  {
    "slug": "portfolio",
    "emoji": "💼",
    "title": "Портфоліо AI-PM",
    "intro": "Пройшовши цей шлях, ти зможеш зібрати портфоліо з двох AI pet-проєктів, які показують рекрутеру не код, а продуктове мислення — і впевнено розповісти їх історію на співбесіді.",
    "steps": [
      {
        "type": "learn",
        "title": "Навіщо AI-PM портфоліо",
        "body": "Портфоліо — це доказ продуктового мислення, а не колекція скріншотів. Для Technical PM з AI-навичками воно замінює досвід, якого ще немає у трудовій книжці: показує, що ти вмієш побачити проблему, сформулювати гіпотезу, побудувати рішення з AI і виміряти результат. Рекрутер за 60 секунд шукає відповідь на одне питання: \"Чи мислить ця людина продуктом?\" Два сильних кейси відповідають на нього краще, ніж десять слабких. Головне — не \"я зробив чат-бота\", а \"я вирішив конкретну проблему конкретних людей і ось цифра\"."
      },
      {
        "type": "learn",
        "title": "Як обрати 2 pet-проєкти",
        "body": "Бери саме два проєкти, а не один — так ти показуєш діапазон, і не десять, щоб не розпорошити глибину. Хороший AI pet-проєкт відповідає трьом критеріям: (1) є реальна проблема живих користувачів, навіть якщо це ти сам чи 5 друзів; (2) AI тут доречний, а не приклеєний заради хайпу; (3) є хоч якась метрика результату. Ідеальна пара — різні за типом: наприклад, один B2C-інструмент (AI-помічник для нотаток) і один процесний/внутрішній (авто-класифікація тікетів підтримки). Уникай проєктів \"обгортка над ChatGPT без задачі\" — вони не показують продуктового рішення."
      },
      {
        "type": "quiz",
        "title": "Перевірка: вибір проєкту",
        "body": "Який pet-проєкт найсильніший для портфоліо AI-PM?",
        "quiz": {
          "question": "Який варіант найкраще підходить як AI pet-проєкт для портфоліо PM?",
          "options": [
            "Красивий лендинг про AI без користувачів і без задачі",
            "Бот, що сумарізує довгі робочі чати — бо колеги втрачали до 40 хв/день на пошук рішень у переписці",
            "Клон ChatGPT, зроблений щоб потренувати навички промптингу",
            "Список з 10 різних напівзроблених AI-експериментів"
          ],
          "answerIndex": 1,
          "explanation": "Сильний кейс = реальна проблема реальних людей + доречний AI + вимірюваний результат (40 хв/день); решта не мають задачі або глибини."
        }
      },
      {
        "type": "learn",
        "title": "Структура кейсу: 5 блоків",
        "body": "Кожен кейс описуй за каркасом із п'яти блоків: Проблема → Інсайт → Рішення → Метрика → Навчене. Проблема — чия біль і чому вона болить (з контекстом і, якщо є, цифрою). Інсайт — неочевидне відкриття з ресерчу чи інтерв'ю, яке змінило підхід (\"користувачі не читали саммарі, бо не довіряли AI\"). Рішення — що саме ти побудував і які продуктові рішення прийняв (чому так, а не інакше). Метрика — результат у числах (retention, час, конверсія, NPS). Навчене — що б ти зробив інакше. Цей каркас перетворює \"я зробив штуку\" на історію продуктового мислення."
      },
      {
        "type": "read",
        "title": "Метрика: як не збрехати цифрою",
        "body": "Метрика має бути чесною і прив'язаною до проблеми. Погано: \"застосунком скористалося 100 людей\" (це трафік, не цінність). Добре: \"з 20 користувачів 14 повернулися на другий тиждень\" або \"середній час на задачу впав з 8 до 3 хв на вибірці 12 людей\". Навіть на pet-проєкті з 5 користувачами краще мала, але справжня цифра, ніж велика вигадана — на співбесіді розкол буде миттєвим. Якщо кількісної метрики немає, використай якісну: 3 показові цитати користувачів. Завжди вказуй розмір вибірки — це ознака зрілості PM.",
        "resource": {
          "label": "Lenny's Newsletter — блог про продукт і метрики",
          "url": "https://www.lennysnewsletter.com/"
        }
      },
      {
        "type": "build",
        "title": "Напиши один кейс за каркасом",
        "body": "Візьми свій сильніший AI-проєкт (або ідею) і напиши кейс на 150-200 слів рівно за п'ятьма блоками: Проблема, Інсайт, Рішення, Метрика, Навчене. Обмеження: у блоці 'Метрика' — конкретне число з розміром вибірки (навіть якщо вибірка = 5). У 'Інсайт' — одне неочевидне речення, яке починається з 'Виявилось, що…'. Не пиши про технології більше одного речення — фокус на продуктовому рішенні, а не на стеку."
      },
      {
        "type": "explain",
        "title": "Поясни: проблема vs рішення",
        "body": "Поясни своїми словами, чому в кейсі блок 'Проблема' має бути сильнішим і конкретнішим за блок 'Рішення'. Хороша відповідь торкається того, що рекрутери оцінюють мислення (чи бачиш ти справжню біль користувача), а не технічну реалізацію, і що будь-яке рішення знецінюється, якщо проблема надумана або розмита."
      },
      {
        "type": "learn",
        "title": "Оформлення в Notion",
        "body": "Notion — найшвидший спосіб зробити портфоліо PM без дизайнера. Структура: головна сторінка з коротким \"про мене\" (1 рядок: хто ти + напрям) і галереєю з 2 кейсів у вигляді карток. Кожен кейс — окрема сторінка з обкладинкою, одним реченням-хуком угорі та п'ятьма блоками каркасу як заголовками H2. Додай 1-2 візуали: скріншот продукту, схему флоу або графік метрики. Зроби сторінку публічною (Share → Publish) і скороти посилання. Правило: рекрутер має зрозуміти суть кейсу за 30 секунд скролу, тому виноси головне вгору і використовуй буліти, а не полотна тексту."
      },
      {
        "type": "quiz",
        "title": "Фінал: storytelling кейсу",
        "body": "Як найкраще почати усну розповідь кейсу на співбесіді?",
        "quiz": {
          "question": "Ти презентуєш AI-кейс інтерв'юеру. З чого почати розповідь?",
          "options": [
            "Зі стеку: 'Я використав GPT-4, LangChain і векторну базу…'",
            "З проблеми й людини: 'Люди Х витрачали Y часу на Z — і ось що я з цим зробив'",
            "З переліку всіх фіч, які ти встиг реалізувати",
            "З того, як складно було технічно все налаштувати"
          ],
          "answerIndex": 1,
          "explanation": "Storytelling PM починається з болю користувача (проблема → людина), бо це створює контекст і показує продуктове мислення; стек і фічі — потім і коротко."
        }
      }
    ]
  },
  {
    "slug": "interview",
    "emoji": "🎯",
    "title": "Підготовка до співбесід",
    "intro": "Пройшовши цей шлях, ти впевнено відповідатимеш на product sense, guesstimate, поведінкові та AI-кейси — і зможеш чітко й переконливо розповідати про свої проєкти на співбесіді на Technical PM.",
    "steps": [
      {
        "type": "learn",
        "title": "Мапа PM-співбесіди",
        "body": "Співбесіда на Product Manager зазвичай складається з кількох типів раундів, і до кожного готуються по-різному. Основні: product sense (спроєктуй продукт або фічу), estimation/guesstimate (оціни величину без даних), поведінкові питання (розкажи про досвід через STAR), технічні/AI-кейси (як працює система, як застосувати ML), та розмова про твої проєкти. Мета інтерв'юера — не почути «правильну відповідь», а побачити твій хід думок: чи структуровано ти міркуєш, чи ставиш уточнюючі питання, чи думаєш про користувача та бізнес одночасно. Для Technical PM додається акцент на розумінні технічних обмежень і на тому, як ти спілкуєшся з інженерами. Знаючи цю карту, ти заздалегідь розумієш, який «жанр» питання перед тобою і за якою рамкою відповідати."
      },
      {
        "type": "learn",
        "title": "Product sense: рамка CIRCLES",
        "body": "Product sense перевіряє, чи вмієш ти проєктувати продукт від потреби користувача до рішення. Популярна рамка — CIRCLES: Comprehend (зрозумій ситуацію), Identify customer (обери сегмент), Report needs (випиши потреби/болі), Cut through prioritization (пріоритизуй), List solutions (згенеруй рішення), Evaluate tradeoffs (зваж компроміси), Summarize (підсумуй рекомендацію). Ключове правило: завжди починай з уточнюючих питань і явно назви, для кого й яку проблему ти вирішуєш, перш ніж пропонувати фічі. Наприклад, на питання «Спроєктуй будильник для незрячих» спершу спитай про контекст, обери сегмент, назви болі (не бачить екран, потрібен тактильний/звуковий зворотний зв'язок), і лише тоді пропонуй рішення. Так ти показуєш user-centric мислення, а не стрибаєш одразу до фіч."
      },
      {
        "type": "quiz",
        "title": "Перевірка: product sense",
        "body": "Тебе просять: «Спроєктуй додаток для доставки продуктів для літніх людей». З чого почати?",
        "quiz": {
          "question": "Який перший крок у сильній product-sense відповіді?",
          "options": [
            "Одразу перелічити 5 крутих фіч, які спадають на думку",
            "Поставити уточнюючі питання і визначити сегмент користувача та його головну проблему",
            "Назвати технологічний стек, на якому будуватимеш додаток",
            "Оцінити, скільки грошей додаток зароблятиме за рік"
          ],
          "answerIndex": 1,
          "explanation": "Спершу зрозумій контекст і болі користувача — рішення й фічі йдуть уже після цього."
        }
      },
      {
        "type": "learn",
        "title": "Guesstimate: розклади на множники",
        "body": "Estimation-питання (наприклад, «Скільки таксі в Києві?» чи «Який ринок кав'ярень?») перевіряють структуроване мислення, а не точну цифру. Метод: розклади велике невідоме на добуток кількох оцінюваних величин, проговорюючи припущення вголос. Для «скільки замовлень піци в місті на день» можна взяти: населення → частка тих, хто їсть піцу → середня частота замовлень → поділити на дні. Завжди озвучуй кожне припущення («припустимо, місто має ~3 млн жителів»), рахуй круглими числами, а в кінці зроби sanity-check: чи результат правдоподібний за порядком величини. Інтерв'юер оцінює логіку декомпозиції та прозорість припущень — навіть якщо фінальна цифра відрізняється від реальної вдвічі, це нормально."
      },
      {
        "type": "build",
        "title": "Практика guesstimate",
        "body": "Візьми питання «Скільки чашок кави випивають у Львові за один день?» і напиши свою оцінку зверху вниз. Обов'язково: (1) почни з населення міста, (2) познач частку людей, що п'ють каву, і скільки чашок на день, (3) додай приблизний внесок туристів/кав'ярень, (4) перемнож і назви фінальне число, (5) зроби sanity-check одним реченням. Записуй кожне припущення окремим рядком — саме структура, а не точність, є метою."
      },
      {
        "type": "learn",
        "title": "Поведінкові питання: метод STAR",
        "body": "Поведінкові питання («Розкажи про конфлікт у команді», «Опиши провал») оцінюють за реальним минулим досвідом. Рамка STAR структурує відповідь: Situation (контекст — де й коли), Task (твоє завдання чи виклик), Action (що конкретно зробила саме ти — деталі й рішення), Result (вимірюваний результат, бажано в цифрах, і чого навчилася). Найпоширеніша помилка — забагато часу на Situation і замало на Action; інтерв'юер хоче почути твій особистий внесок, тому кажи «я», а не «ми». Готуй 5–7 історій наперед, які можна переповісти під різні питання (лідерство, конфлікт, провал, вплив на метрику). Result із конкретним числом («скоротили час онбордингу на 30%») робить історію переконливою й пам'ятною. Ця ж рамка працює, коли тебе просять розказати про проєкт, яким ти пишаєшся."
      },
      {
        "type": "read",
        "title": "AI/технічні кейси для PM",
        "body": "Для Technical PM з AI-фокусом можуть спитати: «Як би ти застосувала ML для рекомендацій?» або «Як працює цей продукт під капотом?». Тут не потрібно писати код — потрібно показати, що ти розумієш, коли ML доречний, які дані потрібні й як оцінити успіх. Хороша рамка відповіді: (1) чи це взагалі ML-задача, чи вистачило б простих правил; (2) яку задачу вирішуємо — класифікація, ранжування, генерація, прогноз; (3) які потрібні дані та звідки їх взяти; (4) як виміряти якість (precision/recall, а поверх — продуктова метрика типу CTR чи retention); (5) ризики — упередженість даних, помилкові спрацювання, приватність, галюцинації LLM. Завжди зв'язуй технічне рішення з користувацькою цінністю та бізнес-метрикою — саме цього чекають від PM, а не від інженера.",
        "resource": {
          "label": "Google People + AI Guidebook",
          "url": "https://pair.withgoogle.com/guidebook/"
        }
      },
      {
        "type": "quiz",
        "title": "Перевірка: AI-кейс",
        "body": "PM пропонують «додати ML» у продукт. Що зробити першим?",
        "quiz": {
          "question": "Який найкращий перший крок, коли пропонують рішення на базі ML?",
          "options": [
            "Одразу обрати нейромережу — вона потужніша за прості моделі",
            "Перевірити, яку задачу вирішуємо і чи не досить простих правил замість ML",
            "Порахувати, скільки GPU знадобиться для навчання",
            "Запустити A/B-тест ще до того, як визначено метрику успіху"
          ],
          "answerIndex": 1,
          "explanation": "Спершу визнач задачу й перевір, чи ML взагалі потрібен — часто прості правила дешевші й достатні."
        }
      },
      {
        "type": "explain",
        "title": "Мілстоун: mock-інтерв'ю",
        "body": "Проведи собі повне mock-інтерв'ю вголос (або з другом чи таймером на 30 хв) і поясни вголос, як ти пройшла кожен блок. Візьми по одному питанню кожного типу: одне product sense (за CIRCLES), один guesstimate (декомпозиція + sanity-check), одне поведінкове (STAR), один AI-кейс (задача → дані → метрика → ризики), і розповідь про свій проєкт за STAR. Гарне проходження: у кожному блоці ти вголос називаєш рамку, ставиш уточнюючі питання, структуруєш відповідь і завершуєш чітким підсумком чи рекомендацією. Це фінальний чекпоінт готовності — якщо всі п'ять блоків звучать структуровано й впевнено, ти готова до реальної співбесіди."
      }
    ]
  }
];

const pathStepId = (slug, i) => `${slug}:${i}`;

function CareerPath({ progress, onDone }) {
  const [open, setOpen] = useState(null); // { modIdx, stepIdx }
  const totalSteps = PM_PATH.reduce((s, m) => s + m.steps.length, 0);
  const doneCount = PM_PATH.reduce((s, m) => s + m.steps.filter((_, i) => progress[pathStepId(m.slug, i)]).length, 0);
  const openMod = open ? PM_PATH[open.modIdx] : null;
  const openStep = openMod ? openMod.steps[open.stepIdx] : null;

  if (!PM_PATH.length) return <div className="rounded-2xl bg-white py-10 text-center text-sm text-slate-400 ring-1 ring-rose-50">Шлях готується…</div>;

  return (
    <div className="space-y-5">
      <div className="rounded-3xl bg-gradient-to-br from-rose-500 to-pink-500 p-5 text-white shadow-sm">
        <div className="flex items-center gap-2 text-sm font-semibold text-white/90"><GraduationCap className="h-4 w-4" /> Шлях: технічний PM з AI</div>
        <div className="mt-1 text-2xl font-extrabold tabular-nums">{doneCount} / {totalSteps} кроків</div>
        <div className="mt-2 h-2.5 overflow-hidden rounded-full bg-white/25"><div className="h-full rounded-full bg-white transition-all" style={{ width: `${totalSteps ? (doneCount / totalSteps) * 100 : 0}%` }} /></div>
        <div className="mt-1.5 text-xs leading-relaxed text-white/80">Маленькі кроки: вивчи → поясни своїми словами → збери → пройди тест. Роби по одному на день. 💪</div>
      </div>

      {PM_PATH.map((mod, mi) => {
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
