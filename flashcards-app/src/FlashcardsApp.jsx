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
  Search, HandHeart, LifeBuoy, Pin, Flower2,
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

  return { ...card, ef, interval, reps, lapses, stepIdx, due, state, lastReviewed: now };
}

/* ------------------------------------------------------------------ */
/* Interval / scheduling visualisation                                 */
/* ------------------------------------------------------------------ */
// Ordered stages from short (hot) to long (cool). `max` is in days.
const INTERVAL_STAGES = [
  { id: "learning", label: "Learning", max: 0, dot: "#e11d48", bg: "bg-rose-100", text: "text-rose-700" },
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
  { key: "again", label: "Again", hint: "1", cls: "bg-rose-600 hover:bg-rose-700", ring: "ring-rose-300" },
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
};
const cKey = (date) => `routine:completions:${date}`;
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
  const completions = {};
  for (const d of cindex) {
    const doc = await store.get(cKey(d), null);
    if (doc) completions[d] = doc;
  }
  return { tasks, categories, cindex, completions, streak, moods };
}

async function collectRoutineExport() {
  const d = await loadRoutineData();
  return { tasks: d.tasks || [], categories: d.categories || [], completions: d.completions, streak: d.streak, moods: d.moods };
}

async function clearRoutineData() {
  const cindex = await store.get(RKEYS.cindex, []);
  for (const d of cindex) await store.remove(cKey(d));
  for (const k of Object.values(RKEYS)) await store.remove(k);
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
  techFav: "calm:tech:favorites", // array of technique numbers ♥
  techTried: "calm:tech:tried",   // { [num]: "YYYY-MM-DD" }
  techWeek: "calm:tech:week",     // array of 1–2 pinned technique numbers
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
  const techFav = await store.get(CKEYS.techFav, []);
  const techTried = await store.get(CKEYS.techTried, {});
  const techWeek = await store.get(CKEYS.techWeek, []);
  return { fears, thoughts, sessions, settings, techFav, techTried, techWeek };
}
async function collectCalmExport() {
  const d = await loadCalmData();
  return { fears: d.fears, thoughts: d.thoughts, sessions: d.sessions, settings: d.settings,
    techFav: d.techFav, techTried: d.techTried, techWeek: d.techWeek };
}
async function clearCalmData() {
  for (const k of Object.values(CKEYS)) await store.remove(k);
}

/* ---- 101 anti-anxiety techniques (Таня Пітерсон), loaded from /anxiety-101.json ---- */
let _anx101 = null;
async function loadAnx101() {
  if (_anx101) return _anx101;
  const res = await fetch("/anxiety-101.json");
  _anx101 = await res.json();
  return _anx101;
}
// Deterministic "technique of the day": stable per calendar day.
function techniqueOfDay(techniques, today = dateKey(Date.now())) {
  if (!techniques || !techniques.length) return null;
  let h = 0;
  for (let i = 0; i < today.length; i++) h = (h * 31 + today.charCodeAt(i)) >>> 0;
  return techniques[h % techniques.length];
}
const PART_STYLES = [
  { accent: "#0ea5e9", soft: "#e0f2fe", icon: Brain },     // 1 thoughts
  { accent: "#14b8a6", soft: "#ccfbf1", icon: Briefcase }, // 2 work/study
  { accent: "#f472b6", soft: "#fce7f3", icon: Heart },     // 3 relationships
  { accent: "#8b5cf6", soft: "#ede9fe", icon: Waves },     // 4 daily/nightly
  { accent: "#10b981", soft: "#d1fae5", icon: Flower2 },   // 5 free long-term
];

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
      const ui = await store.get("ui:prefs", { section: "studying", sidebarCollapsed: false });
      const calmSettings = await store.get(CKEYS.settings, { name: "Спокій" });
      const fastingSettings = await store.get(FKEYS.settings, { name: "Fasting" });
      const mgmtSettings = await store.get("mgmt:settings", { name: "Менеджмент" });
      const toolkitSettings = await store.get(TKEYS.settings, { name: "Toolkit" });
      const st = await store.get("stats", {
        history: {},
        settings: { newPerDay: DEFAULT_NEW_PER_DAY },
      });
      if (!alive) return;
      setSection(["routine", "calm", "fasting", "management", "toolkit"].includes(ui.section) ? ui.section : "studying");
      setSidebarCollapsed(!!ui.sidebarCollapsed);
      setCalmName(calmSettings?.name && calmSettings.name !== "Calm" ? calmSettings.name : "Спокій");
      setFastingName(fastingSettings?.name || "Fasting");
      setMgmtName(mgmtSettings?.name || "Менеджмент");
      setToolkitName(toolkitSettings?.name || "Toolkit");
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

  /* ---------- derived: per-deck due summary ---------- */
  const deckSummary = useMemo(() => {
    const now = Date.now();
    const todayKey = dateKey(now);
    const introducedToday = stats.history?.[todayKey]?.newIntroduced || 0;
    const newRemainingGlobal = Math.max(0, newPerDay - introducedToday);
    const out = {};
    for (const d of decks) {
      const cards = cardsByDeck[d.id] || [];
      let learn = 0, review = 0, newTotal = 0;
      const stages = {}; // stageId -> count, across ALL cards (interval breakdown)
      for (const c of cards) {
        if (c.state === "new") newTotal += 1;
        else if (c.state === "learning" && c.due <= now) learn += 1;
        else if (c.state === "review" && c.due <= now) review += 1;
        const sid = stageForCard(c).id;
        stages[sid] = (stages[sid] || 0) + 1;
      }
      const newDue = Math.min(newTotal, newRemainingGlobal);
      out[d.id] = { learn, review, newTotal, newDue, total: cards.length, due: learn + review + newDue, stages };
    }
    return out;
  }, [decks, cardsByDeck, stats, newPerDay]);

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
        // "due" — the scheduled queue, respecting the new-per-day cap
        const review = all.filter((x) => isReview(x) && dueNow(x));
        const learn = all.filter((x) => isLearn(x) && dueNow(x));
        const news = all.filter(isNew).slice(0, newRemaining);
        const mixed = [];
        const a = [...review], b = [...news];
        while (a.length || b.length) {
          if (a.length) mixed.push(a.shift());
          if (b.length) mixed.push(b.shift());
        }
        out = [...mixed, ...learn];
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
    await store.remove("stats");
    await clearRoutineData();
    await clearCalmData();
    await clearFastingData();
    await store.remove("mgmt:settings");
    await clearToolkitData();
    setDecks([]);
    setGroups([]);
    setCardsByDeck({});
    setStats({ history: {}, settings: { newPerDay: DEFAULT_NEW_PER_DAY } });
    setView("home");
    setCalmName("Спокій");
    setFastingName("Fasting");
    setMgmtName("Менеджмент");
    setToolkitName("Toolkit");
    flash("All data reset");
    window.dispatchEvent(new CustomEvent("routine-reset"));
    window.dispatchEvent(new CustomEvent("calm-reset"));
    window.dispatchEvent(new CustomEvent("fasting-reset"));
    window.dispatchEvent(new CustomEvent("toolkit-reset"));
  }, [decks, cardsByDeck, flash]);

  const exportAll = useCallback(async () => {
    const routine = await collectRoutineExport();
    const calm = await collectCalmExport();
    const fasting = await collectFastingExport();
    const mgmt = await store.get("mgmt:settings", { name: "Менеджмент" });
    const toolkit = await collectToolkitExport();
    const payload = {
      exportedAt: new Date().toISOString(),
      version: 6,
      decks,
      groups,
      cards: cardsByDeck,
      stats,
      routine,
      calm,
      fasting,
      mgmt,
      toolkit,
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
          <Brain className="h-8 w-8 animate-pulse text-indigo-500" />
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
        cloud={cloudState}
        onSyncNow={doSyncNow}
      />

      <div className="flex min-h-screen min-w-0 flex-1 flex-col">
        {section === "routine" ? (
          <RoutineSection />
        ) : section === "calm" ? (
          <CalmSection name={calmName} onRename={renameCalm} />
        ) : section === "fasting" ? (
          <FastingSection name={fastingName} onRename={renameFasting} />
        ) : section === "management" ? (
          <ManagementSection name={mgmtName} onRename={renameMgmt} />
        ) : section === "toolkit" ? (
          <ToolkitSection name={toolkitName} onRename={renameToolkit} />
        ) : (
        <>
      {/* studying top bar */}
      <header className="sticky top-0 z-20 border-b border-slate-200 bg-white/90 backdrop-blur">
        <div className="mx-auto flex h-14 w-full max-w-5xl items-center gap-2 px-4">
          <button onClick={toggleSidebar} className="mr-1 rounded-md p-1.5 text-slate-400 hover:bg-slate-100 hover:text-slate-700 lg:hidden" title="Menu">
            <Menu className="h-5 w-5" />
          </button>
          <button
            onClick={() => { setView("home"); setSession(null); }}
            className="mr-auto flex items-center gap-2 font-semibold tracking-tight text-slate-900"
          >
            <span className="text-base">Studying</span>
            {totalDue > 0 && view === "home" && (
              <CountPill n={totalDue} cls="bg-indigo-100 text-indigo-700 ml-1" />
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
function NavButton({ active, onClick, icon: Icon, children }) {
  return (
    <button
      onClick={onClick}
      className={`flex items-center gap-1.5 rounded-lg px-3 py-1.5 text-sm font-medium transition ${
        active ? "bg-indigo-50 text-indigo-700" : "text-slate-500 hover:bg-slate-100 hover:text-slate-800"
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
function Sidebar({ section, collapsed, onSection, onToggle, studyingDue, calmName, fastingName, mgmtName, toolkitName, cloud, onSyncNow }) {
  const items = [
    { id: "studying", label: "Studying", icon: GraduationCap, badge: studyingDue },
    { id: "routine", label: "My Routine", icon: Sun, badge: 0 },
    { id: "calm", label: calmName || "Спокій", icon: Leaf, badge: 0 },
    { id: "fasting", label: fastingName || "Fasting", icon: Hourglass, badge: 0 },
    { id: "management", label: mgmtName || "Менеджмент", icon: Briefcase, badge: 0 },
    { id: "toolkit", label: toolkitName || "Toolkit", icon: Wrench, badge: 0 },
  ];
  const wide = !collapsed;
  return (
    <aside className={`sticky top-0 flex h-screen shrink-0 flex-col border-r border-slate-200 bg-white transition-all ${collapsed ? "w-16" : "w-16 lg:w-60"}`}>
      <div className="flex h-14 items-center gap-2 px-4">
        <span className="grid h-8 w-8 shrink-0 place-items-center rounded-lg bg-indigo-600 text-white">
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
                active ? "bg-indigo-50 text-indigo-700" : "text-slate-500 hover:bg-slate-100 hover:text-slate-800"
              }`}
            >
              <span className="relative shrink-0">
                <it.icon className="h-5 w-5" />
                {it.badge > 0 && collapsed && <span className="absolute -right-1 -top-1 h-2 w-2 rounded-full bg-indigo-500" />}
              </span>
              {wide && <span className="hidden flex-1 text-left lg:inline">{it.label}</span>}
              {wide && it.badge > 0 && <span className="hidden lg:inline"><CountPill n={it.badge} cls="bg-indigo-100 text-indigo-700" /></span>}
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
        <div className="mx-auto grid h-16 w-16 place-items-center rounded-2xl bg-indigo-600 text-white">
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
            <button onClick={onNewDeck} className="inline-flex items-center gap-2 rounded-lg bg-indigo-600 px-5 py-2.5 font-semibold text-white shadow-sm transition hover:bg-indigo-700">
              <Plus className="h-4 w-4" /> New deck
            </button>
            <button onClick={onImport} className="inline-flex items-center gap-2 rounded-lg border border-slate-300 bg-white px-5 py-2.5 font-semibold text-slate-700 shadow-sm transition hover:bg-slate-50">
              <Upload className="h-4 w-4" /> Import
            </button>
          </div>
          <button onClick={onSample} className="inline-flex items-center gap-2 text-sm font-medium text-slate-500 hover:text-indigo-600">
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
        <StatTile icon={Target} label="Due today" value={totalDue} tint={totalDue ? "text-indigo-600" : "text-slate-400"} sub={totalDue ? "cards waiting" : "all caught up"} />
        <StatTile icon={Check} label="Studied today" value={studiedToday} tint="text-slate-700" sub="reviews done" />
        <StatTile icon={Flame} label="Streak" value={streak} tint={streak ? "text-orange-500" : "text-slate-400"} sub={streak === 1 ? "day" : "days"} />
        <StatTile icon={Layers} label="Decks" value={decks.length} tint="text-slate-700" />
      </div>

      {/* actions */}
      <div className="flex flex-wrap items-center gap-2">
        <button onClick={onStudyAll} className="inline-flex items-center gap-2 rounded-lg bg-indigo-600 px-4 py-2 text-sm font-semibold text-white shadow-sm transition hover:bg-indigo-700">
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
        <div {...dropProps("ungrouped")} className={`rounded-xl transition ${overTarget === "ungrouped" ? "ring-2 ring-indigo-300 ring-offset-2" : ""}`}>
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
        <span className="inline-flex items-center gap-1"><span className="h-2 w-2 rounded-full bg-rose-400" /> learning</span>
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
    <div {...dropProps} className={`overflow-hidden rounded-xl border bg-white shadow-sm transition ${highlight ? "border-indigo-400 ring-2 ring-indigo-200" : "border-slate-200"}`}>
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
          <button onClick={() => onStudyGroup(group.id)} className="inline-flex shrink-0 items-center gap-1 rounded-lg bg-indigo-600 px-3 py-1.5 text-sm font-semibold text-white transition hover:bg-indigo-700">
            <Play className="h-3.5 w-3.5" /> {r.due}
          </button>
        )}
        <button onClick={() => onEditGroup(group)} className="rounded-md p-1.5 text-slate-300 transition hover:bg-slate-100 hover:text-slate-600" title="Edit group">
          <Pencil className="h-4 w-4" />
        </button>
        <button
          onClick={() => { if (confirm(`Delete group “${group.name}”? Its decks are kept and moved to Ungrouped.`)) onDeleteGroup(group.id); }}
          className="rounded-md p-1.5 text-slate-300 transition hover:bg-rose-50 hover:text-rose-500" title="Delete group"
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
      className={`group relative flex items-center gap-3 rounded-xl border border-slate-200 bg-white p-4 shadow-sm transition hover:border-indigo-200 hover:shadow ${dragging ? "opacity-40" : ""}`}
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
            <CountPill n={sum.learn} cls="bg-rose-100 text-rose-700" />
            <CountPill n={sum.review} cls="bg-green-100 text-green-700" />
            <span className="text-slate-400">{sum.total} card{sum.total === 1 ? "" : "s"}</span>
          </div>
        </div>
      </button>
      <div className="flex shrink-0 items-center gap-1">
        {sum.due > 0 ? (
          <button onClick={() => onStudy(deck.id)} className="inline-flex items-center gap-1 rounded-lg bg-indigo-600 px-3 py-1.5 text-sm font-semibold text-white transition hover:bg-indigo-700" title="Study">
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
          className="rounded-md p-1.5 text-slate-300 transition hover:bg-rose-50 hover:text-rose-500 sm:opacity-0 sm:group-hover:opacity-100" title="Delete deck"
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
                className={`flex w-full items-center gap-2 px-3 py-2 text-left text-sm hover:bg-slate-50 ${deck.groupId === g.id ? "font-semibold text-indigo-600" : "text-slate-700"}`}>
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
      className={`grid ${dim} shrink-0 place-items-center rounded-full bg-slate-100 text-slate-500 transition hover:bg-indigo-100 hover:text-indigo-600`}
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
        <button onClick={onExit} className="mt-8 inline-flex items-center gap-2 rounded-lg bg-indigo-600 px-5 py-2.5 font-semibold text-white transition hover:bg-indigo-700">
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
            <div className="h-full rounded-full bg-indigo-600 transition-all duration-300" style={{ width: `${progress}%` }} />
          </div>
        </div>
      </div>

      {/* card (div, not button — it contains speaker buttons) */}
      <div
        role="button"
        tabIndex={0}
        onClick={session.flipped ? undefined : onFlip}
        className={`flex min-h-[300px] w-full flex-col items-center justify-center rounded-2xl border border-slate-200 bg-white p-8 text-center shadow-sm transition ${session.flipped ? "" : "cursor-pointer"}`}
      >
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
              <div className="text-2xl font-medium text-indigo-700" style={{ textWrap: "balance" }}>{card.back}</div>
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
            className="w-full appearance-none rounded-lg border border-slate-300 bg-white px-3 py-2.5 pr-9 text-sm font-medium focus:border-indigo-400 focus:outline-none focus:ring-2 focus:ring-indigo-100"
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
                  active ? "border-indigo-500 bg-indigo-50 ring-2 ring-indigo-100" : "border-slate-200 bg-white hover:border-slate-300"
                }`}
              >
                <div className={`grid h-9 w-9 shrink-0 place-items-center rounded-lg ${active ? "bg-indigo-600 text-white" : "bg-slate-100 text-slate-500"}`}>
                  <Icon className="h-4.5 w-4.5" style={{ width: 18, height: 18 }} />
                </div>
                <div className="min-w-0 flex-1">
                  <div className="flex items-center justify-between gap-2">
                    <span className="font-semibold text-slate-800">{m.label}</span>
                    <span className={`shrink-0 text-xs font-semibold tabular-nums ${n ? "text-indigo-600" : "text-slate-300"}`}>{n.toLocaleString()}</span>
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
            className="w-24 rounded-lg border border-slate-300 px-3 py-1.5 text-right text-sm tabular-nums focus:border-indigo-400 focus:outline-none focus:ring-2 focus:ring-indigo-100"
          />
          <span className="text-xs text-slate-400">due cards come first</span>
        </div>
      )}

      <div className="flex items-center gap-3 border-t border-slate-100 pt-4">
        <button
          onClick={() => onStart(setup)}
          disabled={!startCount}
          className="inline-flex items-center gap-2 rounded-lg bg-indigo-600 px-6 py-2.5 font-semibold text-white shadow-sm transition hover:bg-indigo-700 disabled:cursor-not-allowed disabled:bg-slate-300"
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
            <span className="mb-1 block text-xs font-medium text-slate-500">Name <span className="text-rose-500">*</span></span>
            <input
              autoFocus
              value={name}
              onChange={(e) => setName(e.target.value)}
              placeholder="e.g. Spanish verbs"
              className="w-full rounded-lg border border-slate-300 bg-white px-3 py-2 text-sm focus:border-indigo-400 focus:outline-none focus:ring-2 focus:ring-indigo-100"
            />
          </label>

          <label className="block">
            <span className="mb-1 block text-xs font-medium text-slate-500">Topic / category</span>
            <input
              value={topic}
              onChange={(e) => setTopic(e.target.value)}
              list="topic-options"
              placeholder="Type or pick — e.g. Languages"
              className="w-full rounded-lg border border-slate-300 bg-white px-3 py-2 text-sm focus:border-indigo-400 focus:outline-none focus:ring-2 focus:ring-indigo-100"
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
              className="w-full rounded-lg border border-slate-300 bg-white px-3 py-2 text-sm focus:border-indigo-400 focus:outline-none focus:ring-2 focus:ring-indigo-100"
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
                className={`grid h-9 w-9 place-items-center rounded-lg border text-slate-400 transition ${emoji === "" ? "border-indigo-500 bg-indigo-50" : "border-slate-200 hover:bg-slate-50"}`}
                title="No icon"
              >
                <GraduationCap className="h-4 w-4" />
              </button>
              {DECK_EMOJIS.map((e) => (
                <button
                  key={e}
                  onClick={() => setEmoji(e)}
                  className={`grid h-9 w-9 place-items-center rounded-lg border text-lg transition ${emoji === e ? "border-indigo-500 bg-indigo-50" : "border-slate-200 hover:bg-slate-50"}`}
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
                <select value={groupId} onChange={(e) => setGroupId(e.target.value)} className="w-full appearance-none rounded-lg border border-slate-300 bg-white px-3 py-2 pr-9 text-sm focus:border-indigo-400 focus:outline-none focus:ring-2 focus:ring-indigo-100">
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
                <select value={language} onChange={(e) => setLanguage(e.target.value)} className="w-full appearance-none rounded-lg border border-slate-300 bg-white px-3 py-2 pr-9 text-sm focus:border-indigo-400 focus:outline-none focus:ring-2 focus:ring-indigo-100">
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
            <input type="checkbox" checked={autoPlay} onChange={(e) => setAutoPlay(e.target.checked)} className="h-4 w-4 accent-indigo-600" />
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
                    className={`flex items-start gap-2.5 rounded-xl border p-3 text-left transition ${active ? "border-indigo-500 bg-indigo-50 ring-2 ring-indigo-100" : "border-slate-200 bg-white hover:border-slate-300"}`}
                  >
                    <div className={`grid h-8 w-8 shrink-0 place-items-center rounded-lg ${active ? "bg-indigo-600 text-white" : "bg-slate-100 text-slate-500"}`}>
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
            className="inline-flex items-center gap-2 rounded-lg bg-indigo-600 px-5 py-2 font-semibold text-white shadow-sm transition hover:bg-indigo-700 disabled:cursor-not-allowed disabled:bg-slate-300"
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
            <span className="mb-1 block text-xs font-medium text-slate-500">Name <span className="text-rose-500">*</span></span>
            <input autoFocus value={name} onChange={(e) => setName(e.target.value)} placeholder="e.g. Languages" className="w-full rounded-lg border border-slate-300 bg-white px-3 py-2 text-sm focus:border-indigo-400 focus:outline-none focus:ring-2 focus:ring-indigo-100" />
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
              <button onClick={() => setEmoji("")} className={`grid h-9 w-9 place-items-center rounded-lg border text-slate-400 transition ${emoji === "" ? "border-indigo-500 bg-indigo-50" : "border-slate-200 hover:bg-slate-50"}`}><Folder className="h-4 w-4" /></button>
              {DECK_EMOJIS.map((e) => (
                <button key={e} onClick={() => setEmoji(e)} className={`grid h-9 w-9 place-items-center rounded-lg border text-lg transition ${emoji === e ? "border-indigo-500 bg-indigo-50" : "border-slate-200 hover:bg-slate-50"}`}>{e}</button>
              ))}
            </div>
          </div>
        </div>

        <div className="mt-6 flex items-center justify-end gap-3">
          <button onClick={onClose} className="rounded-lg px-4 py-2 text-sm font-medium text-slate-500 hover:text-slate-800">Cancel</button>
          <button onClick={save} disabled={!name.trim()} className="inline-flex items-center gap-2 rounded-lg bg-indigo-600 px-5 py-2 font-semibold text-white shadow-sm transition hover:bg-indigo-700 disabled:cursor-not-allowed disabled:bg-slate-300">
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
          <button onClick={onStudy} disabled={!cards.length} className="inline-flex items-center gap-1.5 rounded-lg bg-indigo-600 px-4 py-2 text-sm font-semibold text-white shadow-sm transition hover:bg-indigo-700 disabled:bg-slate-300">
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
              className="w-40 rounded-lg border border-slate-300 bg-white px-3 py-1.5 text-sm focus:border-indigo-400 focus:outline-none focus:ring-2 focus:ring-indigo-100 sm:w-56"
            />
          </div>
          <button onClick={onAddCard} className="inline-flex items-center gap-1.5 rounded-lg bg-indigo-600 px-3 py-1.5 text-sm font-semibold text-white transition hover:bg-indigo-700">
            <Plus className="h-4 w-4" /> Add card
          </button>
        </div>

        {cards.length === 0 ? (
          <div className="rounded-xl border border-dashed border-slate-300 bg-white py-12 text-center text-slate-500">
            <p className="font-medium">No cards yet</p>
            <button onClick={onAddCard} className="mt-3 inline-flex items-center gap-1.5 rounded-lg bg-indigo-600 px-4 py-2 text-sm font-semibold text-white transition hover:bg-indigo-700">
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
                </div>
                <StageBadge card={c} />
                <SpeakerButton text={c.front} lang={c.lang || lang} />
                <button onClick={() => onEditCard(c)} className="rounded-md p-1.5 text-slate-300 transition hover:bg-slate-100 hover:text-slate-600 sm:opacity-0 sm:group-hover:opacity-100" title="Edit card">
                  <Pencil className="h-4 w-4" />
                </button>
                <button onClick={() => { if (confirm("Delete this card?")) onDeleteCard(c.id); }} className="rounded-md p-1.5 text-slate-300 transition hover:bg-rose-50 hover:text-rose-500 sm:opacity-0 sm:group-hover:opacity-100" title="Delete card">
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
      className={`rounded-lg border-2 border-dashed p-2 transition ${pasteTarget === side ? "border-indigo-400 bg-indigo-50/40" : "border-slate-200"}`}
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
              className="font-medium text-indigo-600 hover:text-indigo-700"
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

        {err && <div className="mb-3 rounded-lg border border-rose-200 bg-rose-50 px-3 py-2 text-sm text-rose-700">{err}</div>}

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
                className="w-full rounded-lg border border-slate-300 bg-white px-3 py-2 text-sm focus:border-indigo-400 focus:outline-none focus:ring-2 focus:ring-indigo-100"
              />
              <ImageSlot side={side} />
            </div>
          ))}
        </div>

        <label className="mt-4 block">
          <span className="mb-1 block text-xs font-medium text-slate-500">Audio language override (optional)</span>
          <div className="relative sm:w-64">
            <select value={lang} onChange={(e) => setLang(e.target.value)} className="w-full appearance-none rounded-lg border border-slate-300 bg-white px-3 py-2 pr-9 text-sm focus:border-indigo-400 focus:outline-none focus:ring-2 focus:ring-indigo-100">
              {LANGUAGES.map((l) => <option key={l.code || "default"} value={l.code}>{l.code ? l.label : `Deck default (${DECK_LANGUAGES.find((x) => x.code === deck?.language)?.label || deck?.language || "—"})`}</option>)}
            </select>
            <ChevronDown className="pointer-events-none absolute right-3 top-1/2 h-4 w-4 -translate-y-1/2 text-slate-400" />
          </div>
        </label>

        <div className="mt-6 flex items-center justify-end gap-3">
          {busy && <span className="mr-auto text-xs text-slate-400">Processing image…</span>}
          <button onClick={onClose} className="rounded-lg px-4 py-2 text-sm font-medium text-slate-500 hover:text-slate-800">Cancel</button>
          <button onClick={save} disabled={busy} className="inline-flex items-center gap-2 rounded-lg bg-indigo-600 px-5 py-2 font-semibold text-white shadow-sm transition hover:bg-indigo-700 disabled:bg-slate-300">
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
        <div className="flex items-start gap-2 rounded-lg border border-rose-200 bg-rose-50 px-3 py-2 text-sm text-rose-700">
          <X className="mt-0.5 h-4 w-4 shrink-0" /> {error}
        </div>
      )}

      {mode === "file" && !parsed && !sheets && (
        <div>
          <button
            onClick={() => fileRef.current?.click()}
            className="flex w-full flex-col items-center gap-2 rounded-xl border-2 border-dashed border-slate-300 bg-white py-12 text-slate-500 transition hover:border-indigo-400 hover:text-indigo-600"
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
              <button onClick={() => setAllSheets(true)} className="text-indigo-600 hover:text-indigo-700">All</button>
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
                  className="h-4 w-4 shrink-0 accent-indigo-600"
                />
                <input
                  value={s.deckName}
                  onChange={(e) => setSheet(i, { deckName: e.target.value })}
                  className="min-w-0 flex-1 rounded-md border border-transparent bg-transparent px-2 py-1 text-sm font-medium text-slate-800 hover:border-slate-200 focus:border-indigo-400 focus:bg-white focus:outline-none focus:ring-2 focus:ring-indigo-100"
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
              className="inline-flex items-center gap-2 rounded-lg bg-indigo-600 px-5 py-2.5 font-semibold text-white shadow-sm transition hover:bg-indigo-700 disabled:cursor-not-allowed disabled:bg-slate-300"
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
                    {label} {required && <span className="text-rose-500">*</span>}
                  </span>
                  <select
                    value={mapping[field]}
                    onChange={(e) => setMapping((m) => ({ ...m, [field]: e.target.value }))}
                    className="w-full rounded-lg border border-slate-300 bg-white px-3 py-2 text-sm focus:border-indigo-400 focus:outline-none focus:ring-2 focus:ring-indigo-100"
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
              className="inline-flex items-center gap-2 rounded-lg bg-indigo-600 px-5 py-2.5 font-semibold text-white shadow-sm transition hover:bg-indigo-700 disabled:cursor-not-allowed disabled:bg-slate-300"
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
              className="w-full rounded-lg border border-slate-300 bg-white px-3 py-2 font-mono text-sm focus:border-indigo-400 focus:outline-none focus:ring-2 focus:ring-indigo-100"
            />
          </div>
          <button
            onClick={commitPaste}
            disabled={!pasteText.trim()}
            className="inline-flex items-center gap-2 rounded-lg bg-indigo-600 px-5 py-2.5 font-semibold text-white shadow-sm transition hover:bg-indigo-700 disabled:cursor-not-allowed disabled:bg-slate-300"
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
        active ? "bg-indigo-600 text-white" : "text-slate-500 hover:text-slate-800"
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
            className="w-full appearance-none rounded-lg border border-slate-300 bg-white px-3 py-2 pr-9 text-sm focus:border-indigo-400 focus:outline-none focus:ring-2 focus:ring-indigo-100"
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
            className="flex-1 rounded-lg border border-slate-300 bg-white px-3 py-2 text-sm focus:border-indigo-400 focus:outline-none focus:ring-2 focus:ring-indigo-100"
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
        <StatTile icon={Check} label="Studied today" value={studiedToday} tint="text-indigo-600" />
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
                    className="w-full rounded-t bg-indigo-500 transition-all"
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
            className="w-20 rounded-lg border border-slate-300 px-3 py-1.5 text-right text-sm tabular-nums focus:border-indigo-400 focus:outline-none focus:ring-2 focus:ring-indigo-100"
          />
        </label>
        <div className="mt-4 flex flex-wrap gap-3">
          <button onClick={onExport} className="inline-flex items-center gap-2 rounded-lg border border-slate-300 bg-white px-4 py-2 text-sm font-medium text-slate-700 transition hover:border-slate-400 hover:bg-slate-50">
            <Download className="h-4 w-4" /> Export backup (JSON)
          </button>
          <button
            onClick={() => { if (confirm("Reset ALL data — decks, cards, stats AND your routine/habits? This cannot be undone.")) onReset(); }}
            className="inline-flex items-center gap-2 rounded-lg border border-rose-200 bg-white px-4 py-2 text-sm font-medium text-rose-600 transition hover:bg-rose-50"
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
        <h2 className="text-sm font-semibold text-slate-700">Cloud sync (cross-device)</h2>
      </div>

      {signed ? (
        <div className="mt-2 flex items-center justify-between gap-3">
          <p className="text-sm text-slate-500">Synced as <span className="font-semibold text-slate-700">{email}</span>. Your data is backed up and syncs across devices.</p>
          <button onClick={doSignOut} disabled={busy} className="inline-flex shrink-0 items-center gap-1.5 rounded-lg border border-slate-300 px-3 py-1.5 text-sm font-medium text-slate-600 hover:bg-slate-50"><LogOut className="h-4 w-4" /> Sign out</button>
        </div>
      ) : (
        <>
          <p className="mb-3 text-sm text-slate-500">Optional. Sign in with your email to back up everything (Studying, Routine, Calm, Fasting) and sync it across your phone and computer. A 6-digit code will be emailed to you.</p>
          {step !== "code" ? (
            <div className="flex flex-wrap items-center gap-2">
              <div className="relative flex-1 min-w-[200px]">
                <Mail className="pointer-events-none absolute left-3 top-1/2 h-4 w-4 -translate-y-1/2 text-slate-400" />
                <input type="email" value={input} onChange={(e) => setInput(e.target.value)} onKeyDown={(e) => { if (e.key === "Enter") doSend(); }} placeholder="you@email.com" className="w-full rounded-lg border border-slate-300 py-2 pl-9 pr-3 text-sm focus:border-indigo-400 focus:outline-none focus:ring-2 focus:ring-indigo-100" />
              </div>
              <button onClick={doSend} disabled={busy || !input.trim()} className="rounded-lg bg-indigo-600 px-4 py-2 text-sm font-semibold text-white transition hover:bg-indigo-700 disabled:bg-slate-300">{busy ? "…" : "Send code"}</button>
            </div>
          ) : (
            <div className="flex flex-wrap items-center gap-2">
              <input value={code} onChange={(e) => setCode(e.target.value.replace(/\D/g, "").slice(0, 6))} onKeyDown={(e) => { if (e.key === "Enter") doVerify(); }} placeholder="6-digit code" inputMode="numeric" className="w-32 rounded-lg border border-slate-300 px-3 py-2 text-center text-sm tracking-widest focus:border-indigo-400 focus:outline-none focus:ring-2 focus:ring-indigo-100" />
              <button onClick={doVerify} disabled={busy || code.length < 6} className="rounded-lg bg-indigo-600 px-4 py-2 text-sm font-semibold text-white transition hover:bg-indigo-700 disabled:bg-slate-300">{busy ? "…" : "Verify & sync"}</button>
              <button onClick={() => { setStep("idle"); setCode(""); setErr(""); }} className="text-sm font-medium text-slate-400 hover:text-slate-600">Change email</button>
              <span className="w-full text-xs text-slate-400">Code sent to {input}. Check your inbox (and spam).</span>
            </div>
          )}
          {err && <p className="mt-2 text-sm text-rose-600">{err}</p>}
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
    setTasks(d.tasks || []); setCategories(d.categories || []); setCompletions(d.completions);
    setCindex(d.cindex); setStreakMeta(d.streak || { best: 0, lastCelebrated: "" }); setMoods(d.moods || {});
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
  const persistCompletion = useCallback(async (date, doc) => {
    setCompletions((m) => ({ ...m, [date]: doc }));
    await store.set(cKey(date), doc);
    setCindex((prev) => { if (prev.includes(date)) return prev; const next = [...prev, date].sort(); store.set(RKEYS.cindex, next); return next; });
  }, []);

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
    const doc = { ...(completions[selDate] || {}) };
    doc.tasks = { ...(doc.tasks || {}) };
    const willComplete = !doc.tasks[taskId];
    if (willComplete) doc.tasks[taskId] = true; else delete doc.tasks[taskId];
    const nextCompletions = { ...completions, [selDate]: doc };
    await persistCompletion(selDate, doc);
    if (willComplete) maybeCelebrate(nextCompletions);
  }, [completions, selDate, persistCompletion, maybeCelebrate]);

  const toggleSubtask = useCallback(async (taskId, subId) => {
    const task = tasks.find((t) => t.id === taskId);
    const doc = { ...(completions[selDate] || {}) };
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
    await persistCompletion(selDate, doc);
    if (willComplete) maybeCelebrate(nextCompletions);
  }, [tasks, completions, selDate, persistCompletion, maybeCelebrate]);

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
        const doc = { ...(completions[selDate] || {}) };
        doc.goal = { ...(doc.goal || {}) };
        doc.goal[cur.taskId] = cur.elapsed;
        doc.tasks = { ...(doc.tasks || {}) };
        let willComplete = false;
        if (cur.elapsed >= cur.target && !doc.tasks[cur.taskId]) { doc.tasks[cur.taskId] = true; willComplete = true; }
        persistCompletion(selDate, doc);
        if (willComplete) maybeCelebrate({ ...completions, [selDate]: doc });
      }
      return null;
    });
  }, [completions, selDate, persistCompletion, maybeCelebrate]);

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
          const doc = { ...(completions[selDate] || {}) };
          doc.goal = { ...(doc.goal || {}), [cur.taskId]: elapsed };
          doc.tasks = { ...(doc.tasks || {}), [cur.taskId]: true };
          persistCompletion(selDate, doc);
          maybeCelebrate({ ...completions, [selDate]: doc });
          flash("Goal complete! 🎉");
          return null;
        }
        return { ...cur, elapsed };
      });
    }, 1000);
  }, [completions, selDate, persistCompletion, maybeCelebrate, flash]);

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

  if (loading) {
    return <div className="flex flex-1 items-center justify-center text-pink-400"><div className="flex flex-col items-center gap-3"><Sun className="h-8 w-8 animate-pulse" /><span className="text-sm">Loading your routine…</span></div></div>;
  }

  const detailTask = detailId ? tasks.find((t) => t.id === detailId) : null;

  return (
    <div className="min-h-screen flex-1 bg-gradient-to-b from-rose-50 to-pink-50/40" style={{ fontFamily: "inherit" }}>
      {rview === "today" && (
        <div className="mx-auto w-full max-w-2xl px-4 pb-28 pt-5">
          {/* header */}
          <div className="mb-4 flex items-center justify-between">
            <div>
              <h1 className="text-2xl font-extrabold text-slate-900">{isToday ? "Today" : prettyDate(selDate).split(",")[0]}</h1>
              <p className="text-xs font-medium text-slate-400">{prettyDate(selDate)}</p>
            </div>
            <div className="flex items-center gap-2">
              <div className="flex items-center gap-1 rounded-full bg-white px-3 py-1.5 shadow-sm ring-1 ring-rose-100">
                <span className="text-lg">🔥</span>
                <span className="text-sm font-bold tabular-nums text-orange-500">{streak.current}</span>
              </div>
              <div className="relative">
                <button onClick={() => setMenuOpen((v) => !v)} className="grid h-9 w-9 place-items-center rounded-full bg-white text-slate-500 shadow-sm ring-1 ring-rose-100 hover:text-slate-700"><Menu className="h-4 w-4" /></button>
                {menuOpen && (<>
                  <div className="fixed inset-0 z-10" onClick={() => setMenuOpen(false)} />
                  <div className="absolute right-0 top-11 z-20 w-48 overflow-hidden rounded-xl border border-slate-200 bg-white py-1 shadow-lg">
                    <button onClick={() => { setRview("stats"); setMenuOpen(false); }} className="flex w-full items-center gap-2 px-3 py-2 text-left text-sm text-slate-700 hover:bg-slate-50"><BarChart3 className="h-4 w-4 text-slate-400" /> Stats & profile</button>
                    <button onClick={() => { setCatManager(true); setMenuOpen(false); }} className="flex w-full items-center gap-2 px-3 py-2 text-left text-sm text-slate-700 hover:bg-slate-50"><Tag className="h-4 w-4 text-slate-400" /> Manage categories</button>
                  </div>
                </>)}
              </div>
            </div>
          </div>

          {/* week strip */}
          <WeekStrip selDate={selDate} today={today} completions={completions} tasks={tasks} onPick={setSelDate} />

          {/* mood banner */}
          {isToday && !bannerDismissed && moods[today] == null && (
            <div className="mt-4 flex items-center gap-3 rounded-2xl bg-white/80 p-3 shadow-sm ring-1 ring-rose-100">
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
            <button onClick={() => setMoodOpen(true)} className="mt-4 flex w-full items-center gap-3 rounded-2xl bg-white/80 p-3 text-left shadow-sm ring-1 ring-rose-100">
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
            <div className="mt-4 flex items-center gap-3 rounded-2xl bg-white/70 p-3 ring-1 ring-rose-100">
              <ProgressRing pct={dayPct} size={44} stroke={5}><span className="text-[11px] font-bold text-pink-600">{Math.round(dayPct * 100)}%</span></ProgressRing>
              <div className="text-sm font-medium text-slate-600">{dayDone} of {dayTasks.length} done{dayPct >= 1 ? " — all clear! 🎉" : ""}</div>
            </div>
          )}

          {/* task list */}
          <div className="mt-4 space-y-2.5">
            {dayTasks.length === 0 ? (
              <div className="rounded-2xl bg-white/70 py-12 text-center text-sm text-slate-400">Nothing here yet — tap the + to add a task.</div>
            ) : dayTasks.map((t) => (
              <TaskCard key={t.id} task={t} done={!!doc.tasks?.[t.id]} doc={doc}
                timer={timer?.taskId === t.id ? timer : null}
                onToggle={() => toggleTask(t.id)} onOpen={() => setDetailId(t.id)} onStartTimer={() => startTimer(t)} onStopTimer={() => stopTimer(true)} />
            ))}
          </div>
        </div>
      )}

      {rview === "stats" && (
        <RoutineStats tasks={tasks} completions={completions} moods={moods} streak={streak} best={Math.max(streakMeta.best || 0, streak.best)} onBack={() => setRview("today")} />
      )}

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
          onStartTimer={() => startTimer(detailTask)} onStopTimer={() => stopTimer(true)} />
      )}

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

function TaskCard({ task, done, doc, timer, onToggle, onOpen, onStartTimer, onStopTimer }) {
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
          {task.reminder && <span className="inline-flex items-center gap-0.5 opacity-70"><Clock className="h-3 w-3" />{task.reminder}</span>}
        </div>
      </button>
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

function TaskDetail({ task, doc, category, timer, onClose, onEdit, onDelete, onToggle, onToggleSub, onStartTimer, onStopTimer }) {
  const p = getPastel(task.color);
  const done = !!doc.tasks?.[task.id];
  const timed = task.goal?.type === "timed";
  const goalSecs = timer ? timer.elapsed : (doc.goal?.[task.id] || 0);
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
          <button onClick={onDelete} className="rounded-xl border border-rose-200 px-3 py-2.5 text-rose-500 hover:bg-rose-50"><Trash2 className="h-4 w-4" /></button>
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
                  <button onClick={() => setSubs((arr) => arr.filter((x) => x.id !== s.id))} className="rounded p-1 text-slate-300 hover:text-rose-500"><X className="h-4 w-4" /></button>
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
              <button onClick={() => { if (confirm(`Delete “${c.name}”? Its tasks stay, just uncategorized.`)) onDelete(c.id); }} className="rounded p-1 text-slate-300 hover:text-rose-500"><Trash2 className="h-4 w-4" /></button>
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
        <button onClick={onBack} className="grid h-9 w-9 place-items-center rounded-full bg-white text-slate-500 shadow-sm ring-1 ring-rose-100"><ArrowLeft className="h-4 w-4" /></button>
        <h1 className="text-2xl font-extrabold text-slate-900">Stats</h1>
      </div>

      <div className="grid grid-cols-3 gap-3">
        <div className="rounded-2xl bg-white p-4 text-center shadow-sm ring-1 ring-rose-100"><div className="text-3xl">🔥</div><div className="text-2xl font-extrabold tabular-nums text-orange-500">{streak.current}</div><div className="text-[11px] text-slate-400">day streak</div></div>
        <div className="rounded-2xl bg-white p-4 text-center shadow-sm ring-1 ring-rose-100"><div className="text-3xl">🏆</div><div className="text-2xl font-extrabold tabular-nums text-amber-500">{best}</div><div className="text-[11px] text-slate-400">best streak</div></div>
        <div className="rounded-2xl bg-white p-4 text-center shadow-sm ring-1 ring-rose-100"><div className="text-3xl">✅</div><div className="text-2xl font-extrabold tabular-nums text-green-500">{total}</div><div className="text-[11px] text-slate-400">completed</div></div>
      </div>

      {/* completion calendar */}
      <div className="mt-4 rounded-2xl bg-white p-5 shadow-sm ring-1 ring-rose-100">
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
      <div className="mt-4 rounded-2xl bg-white p-5 shadow-sm ring-1 ring-rose-100">
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
      <div className="mt-4 rounded-2xl bg-white p-5 shadow-sm ring-1 ring-rose-100">
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
  const [tab, setTab] = useState("tools");     // tools | library
  const [anxOpen, setAnxOpen] = useState(false);
  const [focusTod, setFocusTod] = useState(0); // bump to scroll library to technique-of-day
  const [fears, setFears] = useState([]);
  const [thoughts, setThoughts] = useState([]);
  const [sessions, setSessions] = useState([]);
  const [settings, setSettings] = useState({ name: "Спокій", tick: true, pattern: "box" });
  const [techData, setTechData] = useState(null);
  const [techFav, setTechFav] = useState([]);
  const [techTried, setTechTried] = useState({});
  const [techWeek, setTechWeek] = useState([]);
  const [toast, setToast] = useState(null);
  const [renaming, setRenaming] = useState(false);
  const [nameDraft, setNameDraft] = useState(name);

  const today = dateKey(Date.now());
  const flash = useCallback((m) => { setToast(m); window.clearTimeout(flash._t); flash._t = window.setTimeout(() => setToast(null), 2200); }, []);

  const reload = useCallback(async () => {
    const d = await loadCalmData();
    setFears(d.fears); setThoughts(d.thoughts); setSessions(d.sessions);
    setSettings({ tick: true, pattern: "box", ...d.settings });
    setTechFav(d.techFav || []); setTechTried(d.techTried || {}); setTechWeek(d.techWeek || []);
    setLoading(false);
    try { setTechData(await loadAnx101()); } catch (e) { /* offline: library shows loader */ }
  }, []);
  useEffect(() => {
    reload();
    const onReset = () => { setFears([]); setThoughts([]); setSessions([]); setTechFav([]); setTechTried({}); setTechWeek([]); setCview("hub"); setTab("tools"); };
    window.addEventListener("calm-reset", onReset);
    return () => window.removeEventListener("calm-reset", onReset);
  }, [reload]);

  const saveTechFav = useCallback((n) => {
    setTechFav((cur) => { const next = cur.includes(n) ? cur.filter((x) => x !== n) : [...cur, n]; store.set(CKEYS.techFav, next); return next; });
  }, []);
  const saveTechTried = useCallback((n) => {
    setTechTried((cur) => { const next = { ...cur }; if (next[n]) { delete next[n]; } else { next[n] = today; flash("Молодець 🌿"); } store.set(CKEYS.techTried, next); return next; });
  }, [today, flash]);
  const saveTechWeek = useCallback((n) => {
    setTechWeek((cur) => {
      if (cur.includes(n)) { const next = cur.filter((x) => x !== n); store.set(CKEYS.techWeek, next); return next; }
      if (cur.length >= 2) { flash("Тримай фокус на 1–2 — спершу відкріпи одну 💛"); return cur; }
      const next = [...cur, n]; store.set(CKEYS.techWeek, next); return next;
    });
  }, [flash]);

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
          <button onClick={() => setAnxOpen(true)} className="mb-4 flex w-full items-center gap-3 rounded-3xl bg-gradient-to-r from-rose-300 via-teal-300 to-sky-300 p-4 text-left text-white shadow-lg shadow-teal-500/20 transition hover:brightness-105">
            <span className="grid h-12 w-12 shrink-0 place-items-center rounded-2xl bg-white/25"><HandHeart className="h-6 w-6" /></span>
            <span className="flex-1"><span className="block text-lg font-bold">Мені зараз тривожно</span><span className="block text-sm text-white/90">Тисни — і я підкажу, з чого почати.</span></span>
            <ArrowRight className="h-5 w-5" />
          </button>

          {/* internal nav */}
          <div className="mb-4 flex gap-2 rounded-2xl bg-white p-1 shadow-sm ring-1 ring-teal-100">
            <button onClick={() => setTab("tools")} className={`flex flex-1 items-center justify-center gap-1.5 rounded-xl py-2 text-sm font-bold transition ${tab === "tools" ? "bg-teal-500 text-white shadow" : "text-slate-500 hover:text-slate-700"}`}><Leaf className="h-4 w-4" /> Інструменти</button>
            <button onClick={() => setTab("library")} className={`flex flex-1 items-center justify-center gap-1.5 rounded-xl py-2 text-sm font-bold transition ${tab === "library" ? "bg-teal-500 text-white shadow" : "text-slate-500 hover:text-slate-700"}`}><BookOpen className="h-4 w-4" /> 101 техніка</button>
          </div>

          {tab === "tools" && (<>
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
          </>)}

          {tab === "library" && (
            techData
              ? <TechLibrary data={techData} favs={techFav} tried={techTried} week={techWeek}
                  onFav={saveTechFav} onTried={saveTechTried} onPin={saveTechWeek} focusTod={focusTod} />
              : <div className="flex items-center justify-center py-16 text-teal-400"><Leaf className="mr-2 h-5 w-5 animate-pulse" /> Завантажую 101 техніку…</div>
          )}

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
                { icon: Sparkles, color: "#10b981", label: "Дай пораду зараз", desc: "Одна техніка, яку можна зробити одразу", go: () => { setTab("library"); setFocusTod((x) => x + 1); } },
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

      {toast && <div className="fixed bottom-6 left-1/2 z-40 -translate-x-1/2 rounded-full bg-slate-900 px-4 py-2 text-sm text-white shadow-lg">{toast}</div>}
    </div>
  );
}

/* ---------- 101 anti-anxiety techniques library ---------- */
function TechCard({ t, style, fav, triedDate, pinned, onFav, onTried, onPin, highlight }) {
  return (
    <div className={`rounded-2xl bg-white p-4 shadow-sm ring-1 transition ${highlight ? "ring-2 ring-teal-400" : "ring-teal-50"}`}>
      <div className="flex items-start gap-3">
        <span className="grid h-8 w-8 shrink-0 place-items-center rounded-lg text-sm font-extrabold tabular-nums text-white" style={{ backgroundColor: style.accent }}>{t.n}</span>
        <h3 className="min-w-0 flex-1 pt-1 font-bold leading-snug text-slate-800">{t.title}</h3>
      </div>
      <p className="mt-2 text-sm leading-relaxed text-slate-600">{t.idea}</p>
      <div className="mt-3 rounded-xl border border-amber-200 bg-amber-50/70 p-3">
        <div className="mb-1 flex items-center gap-1.5 text-[11px] font-bold uppercase tracking-wide text-amber-600"><Sparkles className="h-3.5 w-3.5" /> Зроби зараз</div>
        <p className="text-sm leading-relaxed text-amber-900">{t.lifehack}</p>
      </div>
      <div className="mt-3 flex flex-wrap items-center gap-2">
        <button onClick={onFav} title="До улюблених"
          className={`flex items-center gap-1.5 rounded-full px-3 py-1.5 text-xs font-semibold ring-1 transition ${fav ? "bg-rose-50 text-rose-600 ring-rose-200" : "bg-white text-slate-500 ring-slate-200 hover:text-rose-500"}`}>
          <Heart className={`h-3.5 w-3.5 ${fav ? "fill-rose-500 text-rose-500" : ""}`} /> Улюблене
        </button>
        <button onClick={onTried} title="Позначити, що спробувала"
          className={`flex items-center gap-1.5 rounded-full px-3 py-1.5 text-xs font-semibold ring-1 transition ${triedDate ? "bg-emerald-50 text-emerald-600 ring-emerald-200" : "bg-white text-slate-500 ring-slate-200 hover:text-emerald-500"}`}>
          <Check className="h-3.5 w-3.5" /> {triedDate ? `Спробувала ${triedDate}` : "Спробувала"}
        </button>
        <button onClick={onPin} title="Практикувати цього тижня"
          className={`ml-auto flex items-center gap-1.5 rounded-full px-3 py-1.5 text-xs font-semibold ring-1 transition ${pinned ? "bg-sky-50 text-sky-600 ring-sky-200" : "bg-white text-slate-500 ring-slate-200 hover:text-sky-500"}`}>
          <Pin className={`h-3.5 w-3.5 ${pinned ? "fill-sky-500 text-sky-500" : ""}`} /> {pinned ? "Цього тижня" : "На тиждень"}
        </button>
      </div>
    </div>
  );
}

function TechLibrary({ data, favs, tried, week, onFav, onTried, onPin, focusTod }) {
  const [query, setQuery] = useState("");
  const [part, setPart] = useState(0);        // 0 = all parts
  const [mode, setMode] = useState("all");    // all | fav | untried
  const techs = data.techniques;
  const parts = data.parts;
  const favSet = useMemo(() => new Set(favs), [favs]);
  const weekSet = useMemo(() => new Set(week), [week]);
  const today = dateKey(Date.now());
  const tod = useMemo(() => techniqueOfDay(techs, today), [techs, today]);
  const triedCount = Object.keys(tried).length;

  const styleFor = (n) => PART_STYLES[(n || 1) - 1] || PART_STYLES[0];
  const cardProps = (t) => ({
    t, style: styleFor(t.part), fav: favSet.has(t.n), triedDate: tried[t.n], pinned: weekSet.has(t.n),
    onFav: () => onFav(t.n), onTried: () => onTried(t.n), onPin: () => onPin(t.n),
  });

  const q = query.trim().toLowerCase();
  const filtered = techs.filter((t) => {
    if (part && t.part !== part) return false;
    if (mode === "fav" && !favSet.has(t.n)) return false;
    if (mode === "untried" && tried[t.n]) return false;
    if (q && !(t.title.toLowerCase().includes(q) || t.idea.toLowerCase().includes(q) || t.lifehack.toLowerCase().includes(q))) return false;
    return true;
  });
  const groups = parts.map((p) => ({ p, items: filtered.filter((t) => t.part === p.n) })).filter((g) => g.items.length);
  const browsing = !q && mode === "all" && part === 0; // show week + tech-of-day only when not searching/filtering

  const pinnedTechs = week.map((n) => techs.find((t) => t.n === n)).filter(Boolean);
  const todRef = useRef(null);
  useEffect(() => { if (focusTod && todRef.current) todRef.current.scrollIntoView({ behavior: "smooth", block: "center" }); }, [focusTod]);

  return (
    <div className="space-y-4">
      {/* progress */}
      <div className="flex items-center justify-between rounded-2xl bg-white p-4 shadow-sm ring-1 ring-teal-100">
        <div>
          <div className="text-sm font-bold text-slate-700">Твій прогрес</div>
          <div className="text-xs text-slate-400">Пробуй по одній — поспіху немає</div>
        </div>
        <div className="text-right"><div className="text-2xl font-extrabold tabular-nums text-teal-600">{triedCount}<span className="text-base text-slate-300">/{techs.length}</span></div><div className="text-[11px] text-slate-400">спробувано</div></div>
      </div>
      <div className="h-2 overflow-hidden rounded-full bg-teal-50"><div className="h-full rounded-full bg-gradient-to-r from-teal-400 to-emerald-400 transition-all" style={{ width: `${Math.round((triedCount / techs.length) * 100)}%` }} /></div>

      {/* this week */}
      {browsing && pinnedTechs.length > 0 && (
        <div className="rounded-2xl border border-sky-200 bg-sky-50/60 p-4">
          <div className="mb-2 flex items-center gap-1.5 text-sm font-bold text-sky-700"><Pin className="h-4 w-4" /> Цього тижня практикую</div>
          <div className="space-y-3">{pinnedTechs.map((t) => <TechCard key={t.n} {...cardProps(t)} />)}</div>
        </div>
      )}

      {/* technique of the day */}
      {browsing && tod && (
        <div ref={todRef}>
          <div className="mb-2 flex items-center gap-1.5 text-sm font-bold text-teal-700"><Sparkle className="h-4 w-4" /> Техніка дня</div>
          <TechCard {...cardProps(tod)} highlight={!!focusTod} />
        </div>
      )}

      {/* search */}
      <div className="relative">
        <Search className="pointer-events-none absolute left-3 top-1/2 h-4 w-4 -translate-y-1/2 text-slate-400" />
        <input value={query} onChange={(e) => setQuery(e.target.value)} placeholder="Пошук серед 101 техніки…"
          className="w-full rounded-2xl border border-teal-100 bg-white py-2.5 pl-10 pr-9 text-sm text-slate-700 shadow-sm focus:border-teal-300 focus:outline-none" />
        {query && <button onClick={() => setQuery("")} className="absolute right-3 top-1/2 -translate-y-1/2 text-slate-400 hover:text-slate-600"><X className="h-4 w-4" /></button>}
      </div>

      {/* filters */}
      <div className="flex flex-wrap gap-2">
        {[["all", "Усі"], ["fav", "Улюблені"], ["untried", "Ще не пробувала"]].map(([m, label]) => (
          <button key={m} onClick={() => setMode(m)} className={`rounded-full px-3 py-1.5 text-xs font-semibold ring-1 transition ${mode === m ? "bg-teal-500 text-white ring-teal-500" : "bg-white text-slate-500 ring-slate-200 hover:ring-teal-200"}`}>{label}</button>
        ))}
      </div>
      <div className="-mx-1 flex gap-2 overflow-x-auto px-1 pb-1">
        <button onClick={() => setPart(0)} className={`shrink-0 rounded-full px-3 py-1.5 text-xs font-semibold ring-1 transition ${part === 0 ? "bg-slate-800 text-white ring-slate-800" : "bg-white text-slate-500 ring-slate-200"}`}>Усі частини</button>
        {parts.map((p) => { const st = styleFor(p.n); const on = part === p.n; return (
          <button key={p.n} onClick={() => setPart(on ? 0 : p.n)} className="shrink-0 rounded-full px-3 py-1.5 text-xs font-semibold ring-1 transition"
            style={on ? { backgroundColor: st.accent, color: "#fff", borderColor: st.accent, boxShadow: `0 0 0 1px ${st.accent}` } : { backgroundColor: "#fff", color: "#64748b" }}>
            {p.n}. {p.title}
          </button>
        ); })}
      </div>

      {/* results */}
      {groups.length === 0 ? (
        <div className="rounded-2xl bg-white p-8 text-center text-sm text-slate-400 ring-1 ring-teal-50">Нічого не знайшлося. Спробуй інший запит чи фільтр.</div>
      ) : groups.map((g) => {
        const st = styleFor(g.p.n);
        return (
          <div key={g.p.n} className="space-y-3">
            <div className="flex items-center gap-2 pt-1">
              <span className="grid h-7 w-7 place-items-center rounded-lg text-white" style={{ backgroundColor: st.accent }}><st.icon className="h-4 w-4" /></span>
              <h2 className="text-sm font-extrabold text-slate-700">Частина {g.p.n}. {g.p.title}</h2>
              <span className="ml-auto text-xs text-slate-400">{g.items.length}</span>
            </div>
            {g.items.map((t) => <TechCard key={t.n} {...cardProps(t)} />)}
          </div>
        );
      })}
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
          <div key={s.key} className="rounded-2xl bg-white p-4 shadow-sm ring-1 ring-indigo-50">
            <div className="mb-2 flex items-center gap-2"><span className="text-xl">{s.emoji}</span><span className="font-bold text-slate-800">{s.n} {s.label}</span></div>
            <div className="flex flex-wrap gap-2">
              {Array.from({ length: s.n }).map((_, i) => {
                const on = (filled[s.key] || []).includes(i);
                return <button key={i} onClick={() => tap(s.key, i)} className={`grid h-10 w-10 place-items-center rounded-full border-2 transition ${on ? "border-transparent bg-indigo-500 text-white" : "border-indigo-200 text-indigo-300 hover:border-indigo-400"}`}>{on ? <Check className="h-5 w-5" /> : <Circle className="h-4 w-4" />}</button>;
              })}
            </div>
          </div>
        ))}
      </div>
      {complete && (
        <button onClick={() => onDone(Math.round((Date.now() - startRef.current) / 1000))} className="mt-5 w-full rounded-2xl bg-indigo-500 py-3.5 font-bold text-white shadow-lg shadow-indigo-500/20 hover:bg-indigo-600">Готово — я більше тут 🌿</button>
      )}
    </div>
  );
}

/* ---------- Thought record (CBT) ---------- */
function ThoughtRecord({ thoughts, onExit, onSave, onDelete }) {
  const [open, setOpen] = useState(false);
  const [f, setF] = useState({ situation: "", thought: "", emotion: "", intensity: 60, forEv: "", against: "", balanced: "" });
  const startRef = useRef(Date.now());
  const save = () => {
    if (!f.thought.trim() && !f.situation.trim()) return;
    onSave({ id: ruid("tr"), date: dateKey(Date.now()), ...f }, Math.round((Date.now() - startRef.current) / 1000));
    setOpen(false); setF({ situation: "", thought: "", emotion: "", intensity: 60, forEv: "", against: "", balanced: "" });
  };
  const Field = ({ label, hint, k, rows = 2 }) => (
    <label className="block"><span className="mb-1 block text-sm font-semibold text-slate-700">{label}</span>{hint && <span className="mb-1 block text-xs text-slate-400">{hint}</span>}
      <textarea value={f[k]} onChange={(e) => setF((s) => ({ ...s, [k]: e.target.value }))} rows={rows} className="w-full rounded-xl border border-slate-300 bg-white px-3 py-2 text-sm focus:border-violet-400 focus:outline-none focus:ring-2 focus:ring-violet-100" /></label>
  );

  if (open) return (
    <div className="mx-auto w-full max-w-lg px-4 pb-16 pt-6">
      <CalmHeader title="Журнал думок" onExit={() => setOpen(false)} />
      <div className="space-y-3">
        <Field label="Ситуація" hint="Що відбувалося?" k="situation" />
        <Field label="Автоматична думка" hint="Що промайнуло в голові?" k="thought" />
        <div className="rounded-xl bg-white p-3 ring-1 ring-slate-100">
          <div className="mb-1 flex items-center justify-between"><span className="text-sm font-semibold text-slate-700">Емоція</span><span className="text-sm font-bold text-violet-600 tabular-nums">{f.intensity}%</span></div>
          <input value={f.emotion} onChange={(e) => setF((s) => ({ ...s, emotion: e.target.value }))} placeholder="напр. тривога, сум" className="mb-2 w-full rounded-lg border border-slate-300 px-3 py-1.5 text-sm focus:border-violet-400 focus:outline-none" />
          <input type="range" min={0} max={100} value={f.intensity} onChange={(e) => setF((s) => ({ ...s, intensity: +e.target.value }))} className="w-full accent-violet-500" />
        </div>
        <Field label="Докази за цю думку" k="forEv" />
        <Field label="Докази проти неї" k="against" />
        <Field label="Врівноважена думка" hint="Добріший, реалістичніший погляд" k="balanced" />
        <button onClick={save} className="w-full rounded-2xl bg-violet-500 py-3 font-bold text-white hover:bg-violet-600">Зберегти запис</button>
      </div>
    </div>
  );

  return (
    <div className="mx-auto w-full max-w-lg px-4 pb-16 pt-6">
      <CalmHeader title="Журнал думок" onExit={onExit} right={<button onClick={() => { startRef.current = Date.now(); setOpen(true); }} className="inline-flex items-center gap-1 rounded-full bg-violet-500 px-3 py-1.5 text-sm font-semibold text-white hover:bg-violet-600"><Plus className="h-4 w-4" /> Новий</button>} />
      {thoughts.length === 0 ? (
        <div className="rounded-2xl bg-white py-12 text-center text-sm text-slate-400 ring-1 ring-violet-50">Записів ще немає. Злови тривожну думку і розплутай її.</div>
      ) : (
        <div className="space-y-3">
          {thoughts.map((t) => (
            <div key={t.id} className="rounded-2xl bg-white p-4 shadow-sm ring-1 ring-violet-50">
              <div className="mb-1 flex items-center justify-between"><span className="text-xs font-medium text-slate-400">{t.date}{t.emotion ? ` · ${t.emotion} ${t.intensity}%` : ""}</span><button onClick={() => onDelete(t.id)} className="text-slate-300 hover:text-rose-500"><Trash2 className="h-4 w-4" /></button></div>
              {t.thought && <div className="font-semibold text-slate-800">“{t.thought}”</div>}
              {t.balanced && <div className="mt-1 rounded-lg bg-violet-50 px-3 py-2 text-sm text-violet-800">↪ {t.balanced}</div>}
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
                  <button onClick={() => { if (confirm("Прибрати цей страх?")) removeFear(f.id); }} className="text-slate-300 hover:text-rose-500"><Trash2 className="h-4 w-4" /></button>
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
      <div className="my-10 grid h-64 w-64 place-items-center rounded-full bg-gradient-to-br from-pink-300 to-rose-300 text-white"><div className="text-6xl font-extrabold tabular-nums">{fmtClock(remaining)}</div></div>
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
        {fview === "ladder" && <ProtocolLadder goals={goals} diary={diary} onSet={(id) => { saveGoals({ protocol: id }); flash(`Протокол: ${getProtocol(id).label}`); setFview("timer"); }} />}
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
        <circle cx={size / 2} cy={size / 2} r={r} fill="none" stroke={color} strokeWidth={stroke} strokeDasharray={c} strokeDashoffset={off} strokeLinecap="round" style={{ transition: "stroke-dashoffset .8s ease" }} />
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
        <div className="text-4xl font-extrabold tabular-nums text-slate-900">{fmtHM(elapsedMs)}</div>
        <div className="mt-1 text-sm text-slate-400">ціль {targetH} год{current ? ` · ${Math.round(pct * 100)}%` : ""}</div>
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
function ProtocolLadder({ goals, diary, onSet }) {
  const cur = getProtocol(goals.protocol);
  const doneAtCurrent = diary.filter((r) => r.protocol === goals.protocol && r.goalMet).length;
  const nextP = PROTOCOLS.find((p) => p.level === cur.level + 1);
  const readyToStep = doneAtCurrent >= 5 && nextP;

  return (
    <div>
      <h1 className="text-xl font-extrabold text-slate-900">Драбина протоколів</h1>
      <p className="mt-1 rounded-2xl bg-amber-50 px-3 py-2 text-sm leading-relaxed text-amber-800">
        Починай з найм'якшого щабля й піднімайся <b>поступово</b> — тільки коли поточний рівень дається легко й приємно. Немає «єдино правильного» рівня; сталість важливіша за інтенсивність. Жодного поспіху.
      </p>

      {readyToStep && (
        <div className="mt-3 flex items-center gap-3 rounded-2xl border border-green-200 bg-green-50 px-4 py-3">
          <Sparkle className="h-5 w-5 shrink-0 text-green-500" />
          <p className="text-sm text-green-800">Ти комфортно витримала <b>{cur.label}</b> вже {doneAtCurrent} разів. Якщо почуваєшся добре — можна спробувати <b>{nextP.label}</b>. Але без тиску, коли сама відчуєш готовність.</p>
        </div>
      )}

      <div className="mt-4 space-y-2">
        {[...PROTOCOLS].reverse().map((p) => {
          const active = p.id === goals.protocol;
          const color = protocolColor(p.level);
          return (
            <div key={p.id} className={`flex items-start gap-3 rounded-2xl border p-4 transition ${active ? "shadow-md" : "border-slate-100 bg-white"}`} style={active ? { borderColor: color, backgroundColor: color + "12" } : {}}>
              <span className="grid h-11 w-11 shrink-0 place-items-center rounded-xl text-sm font-extrabold text-white" style={{ backgroundColor: color }}>{p.hrs}г</span>
              <div className="min-w-0 flex-1">
                <div className="flex flex-wrap items-center gap-2">
                  <span className="font-bold text-slate-900">{p.label}</span>
                  <span className="text-xs text-slate-400">вікно {p.window} год · {p.freq}</span>
                  {active && <span className="rounded-full px-2 py-0.5 text-[10px] font-bold text-white" style={{ backgroundColor: color }}>мій зараз</span>}
                </div>
                <p className="mt-0.5 text-sm text-slate-500">{p.note}</p>
              </div>
              {!active && <button onClick={() => onSet(p.id)} className="shrink-0 rounded-full bg-slate-800 px-3 py-1.5 text-xs font-semibold text-white hover:bg-slate-900">Обрати</button>}
            </div>
          );
        })}
      </div>
    </div>
  );
}

/* ---------- Overview: goals + metrics + charts ---------- */
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
              <span onClick={(e) => { e.stopPropagation(); if (confirm("Видалити запис?")) onDelete(r.id); }} className="rounded p-1 text-slate-300 hover:text-rose-500"><Trash2 className="h-4 w-4" /></span>
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
  const okColor = { yes: "bg-green-50", warn: "bg-amber-50", no: "bg-rose-50" };
  return (
    <div className="space-y-5">
      <h1 className="text-xl font-extrabold text-slate-900">Довідник</h1>

      {/* SAFETY — prominent, first */}
      <div className="rounded-2xl border-2 border-rose-200 bg-rose-50/60 p-4">
        <div className="mb-2 flex items-center gap-2 text-rose-700"><ShieldAlert className="h-5 w-5" /><h2 className="font-bold">Застереження та безпека</h2></div>
        <div className="text-sm font-semibold text-slate-700">Кому НЕ можна голодувати (або лише під наглядом лікаря):</div>
        <ul className="mt-1 space-y-1">{NO_FAST.map((x, i) => <li key={i} className="flex gap-2 text-sm text-slate-600"><span className="text-rose-400">•</span>{x}</li>)}</ul>
        <div className="mt-4 overflow-hidden rounded-xl ring-1 ring-rose-100">
          <div className="grid grid-cols-2 bg-rose-100/70 px-3 py-1.5 text-xs font-bold text-rose-700"><span>Симптом</span><span>Що робити</span></div>
          {SIDE_FX.map(([s, w], i) => <div key={i} className={`grid grid-cols-2 gap-2 px-3 py-2 text-sm ${i % 2 ? "bg-white" : "bg-rose-50/40"}`}><span className="font-semibold text-slate-700">{s}</span><span className="text-slate-600">{w}</span></div>)}
        </div>
        <p className="mt-3 rounded-lg bg-rose-600 px-3 py-2 text-sm font-semibold text-white">❗ Негайно припини голодування і звернись по допомогу при: сильній слабкості, плутанині свідомості, серцебитті, непритомності.</p>
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
    else if (m[2] != null) out.push(<code key={keyBase + k++} className="rounded bg-slate-100 px-1 py-0.5 text-[0.85em] text-indigo-700">{m[2]}</code>);
    else out.push(<span key={keyBase + k++} className="font-medium text-indigo-600">{m[3]}</span>);
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
    if (list.type === "ul") els.push(<ul key={els.length} className="mb-3 space-y-1.5">{list.items.map((it, i) => <li key={i} className="flex gap-2.5 text-slate-600"><span className="mt-[9px] h-1.5 w-1.5 shrink-0 rounded-full bg-indigo-300" /><span className="leading-relaxed">{mdInline(it, els.length + "u" + i)}</span></li>)}</ul>);
    else els.push(<ol key={els.length} className="mb-3 space-y-2">{list.items.map((it, i) => <li key={i} className="flex gap-2.5 text-slate-600"><span className="grid h-5 w-5 shrink-0 place-items-center rounded-full bg-indigo-100 text-[11px] font-bold text-indigo-600">{i + 1}</span><span className="leading-relaxed">{mdInline(it, els.length + "o" + i)}</span></li>)}</ol>);
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
  if (!doc) return <div className="flex flex-1 items-center justify-center text-indigo-400"><div className="flex flex-col items-center gap-3"><BookMarked className="h-8 w-8 animate-pulse" /><span className="text-sm">Завантаження конспекту…</span></div></div>;

  const chapters = doc.chapters;
  const open = (i) => setView(i);
  const chapter = typeof view === "number" ? chapters[view] : null;

  return (
    <div className="min-h-screen flex-1 bg-gradient-to-b from-indigo-50/50 via-slate-50 to-white">
      <header className="sticky top-0 z-20 border-b border-slate-200 bg-white/85 backdrop-blur">
        <div className="mx-auto flex h-14 w-full max-w-3xl items-center gap-2 px-4">
          {renaming ? (
            <input autoFocus value={nameDraft} onChange={(e) => setNameDraft(e.target.value)} onBlur={() => { onRename(nameDraft); setRenaming(false); }} onKeyDown={(e) => { if (e.key === "Enter") { onRename(nameDraft); setRenaming(false); } }} className="mr-auto w-40 rounded-lg border border-indigo-200 px-2 py-1 text-base font-semibold focus:outline-none" />
          ) : (
            <button onClick={() => { setNameDraft(name); setRenaming(true); }} className="mr-auto text-base font-semibold text-slate-900">{name} <Pencil className="ml-0.5 inline h-3.5 w-3.5 text-slate-300" /></button>
          )}
          {view !== "toc" && <button onClick={() => setView("toc")} className="inline-flex items-center gap-1.5 rounded-lg px-3 py-1.5 text-sm font-medium text-slate-500 hover:bg-slate-100 hover:text-slate-800"><ListTree className="h-4 w-4" /> Зміст</button>}
        </div>
      </header>

      <main className="mx-auto w-full max-w-3xl px-4 py-6">
        {view === "toc" ? (
          <>
            {/* book hero from intro */}
            <div className="mb-6 rounded-3xl border border-indigo-100 bg-white p-6 shadow-sm">
              <div className="flex items-start gap-4">
                <span className="grid h-14 w-14 shrink-0 place-items-center rounded-2xl bg-indigo-600 text-white"><BookMarked className="h-7 w-7" /></span>
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
                  <button key={i} onClick={() => open(i)} className={`flex w-full items-center gap-3 rounded-2xl border p-4 text-left shadow-sm transition hover:shadow-md ${special ? "border-amber-200 bg-amber-50/60 hover:border-amber-300" : "border-slate-100 bg-white hover:border-indigo-200"}`}>
                    <span className={`grid h-10 w-10 shrink-0 place-items-center rounded-xl text-sm font-extrabold ${special ? "bg-amber-400 text-white" : "bg-indigo-100 text-indigo-700"}`}>{special ? "🎯" : c.num}</span>
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
        {chapter.num && <span className="grid h-11 w-11 shrink-0 place-items-center rounded-2xl bg-indigo-600 text-lg font-extrabold text-white">{chapter.num}</span>}
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
        <button disabled={index >= total - 1} onClick={() => onNav(index + 1)} className="inline-flex items-center gap-1.5 rounded-xl bg-indigo-600 px-4 py-2 text-sm font-semibold text-white transition hover:bg-indigo-700 disabled:opacity-40">Далі <ChevronRight className="h-4 w-4" /></button>
      </div>
    </div>
  );
}

/* ================================================================== */
/* TOOLKIT — section UI                                               */
/* ================================================================== */
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

  if (loading) return <div className="flex flex-1 items-center justify-center text-slate-400"><div className="flex flex-col items-center gap-3"><Wrench className="h-8 w-8 animate-pulse text-indigo-400" /><span className="text-sm">Завантаження…</span></div></div>;

  return (
    <div className="min-h-screen flex-1 bg-gradient-to-b from-slate-50 to-white">
      <header className="sticky top-0 z-20 border-b border-slate-200 bg-white/85 backdrop-blur">
        <div className="mx-auto flex h-14 w-full max-w-3xl items-center gap-2 px-4">
          {renaming ? (
            <input autoFocus value={nameDraft} onChange={(e) => setNameDraft(e.target.value)} onBlur={() => { onRename(nameDraft); setRenaming(false); }} onKeyDown={(e) => { if (e.key === "Enter") { onRename(nameDraft); setRenaming(false); } }} className="mr-auto w-32 rounded-lg border border-indigo-200 px-2 py-1 text-base font-semibold focus:outline-none" />
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
            <button onClick={() => setTool("chores")} className="flex w-full items-center gap-4 rounded-2xl border border-slate-200 bg-white p-5 text-left shadow-sm transition hover:border-indigo-200 hover:shadow-md">
              <span className="grid h-14 w-14 shrink-0 place-items-center rounded-2xl bg-indigo-100 text-2xl">🧹</span>
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
              {i > 0 && <span className={`h-0.5 w-3 ${done || active ? "bg-indigo-400" : "bg-slate-200"}`} />}
              <span className={`grid h-7 w-7 shrink-0 place-items-center rounded-full text-xs font-bold transition ${active ? "bg-indigo-600 text-white" : done ? "bg-indigo-100 text-indigo-600" : "bg-slate-100 text-slate-400"}`}>{done ? "✓" : n}</span>
              <span className={`hidden whitespace-nowrap text-xs font-medium sm:inline ${active ? "text-indigo-700" : "text-slate-400"}`}>{s}</span>
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
          ? <button disabled={step === 2 && list.length === 0} onClick={() => setStep(step + 1)} className="inline-flex items-center gap-1.5 rounded-xl bg-indigo-600 px-5 py-2 text-sm font-semibold text-white transition hover:bg-indigo-700 disabled:opacity-40">Далі <ChevronRight className="h-4 w-4" /></button>
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
            <span className="grid h-9 w-9 shrink-0 place-items-center rounded-lg bg-indigo-100 text-sm font-bold text-indigo-600">{i + 1}</span>
            <input value={m.name} onChange={(e) => onSave(members.map((x) => x.id === m.id ? { ...x, name: e.target.value } : x))} className="flex-1 rounded-lg border border-slate-300 px-3 py-2 text-sm focus:border-indigo-400 focus:outline-none" />
            {members.length > 2 && <button onClick={() => onSave(members.filter((x) => x.id !== m.id))} className="rounded p-1.5 text-slate-300 hover:text-rose-500"><Trash2 className="h-4 w-4" /></button>}
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
        <input value={text} onChange={(e) => setText(e.target.value)} onKeyDown={(e) => { if (e.key === "Enter") { add(text); setText(""); } }} placeholder="Додати справу…" className="flex-1 rounded-lg border border-slate-300 px-3 py-2 text-sm focus:border-indigo-400 focus:outline-none" />
        <button onClick={() => { add(text); setText(""); }} className="rounded-lg bg-indigo-600 px-3 py-2 text-sm font-semibold text-white hover:bg-indigo-700"><Plus className="h-4 w-4" /></button>
      </div>

      <div className="mb-4 space-y-3">
        <div className="text-xs font-semibold uppercase tracking-wide text-slate-400">Готовий чек-лист по кімнатах — тапни, щоб додати</div>
        {CHORE_ROOMS.map((r) => (
          <div key={r.room}>
            <div className="mb-1 flex items-center gap-1.5 text-sm font-semibold text-slate-600"><Home className="h-3.5 w-3.5 text-slate-400" /> {r.room}</div>
            <div className="flex flex-wrap gap-1.5">
              {r.items.map((it) => { const added = has(it); return (
                <button key={it} onClick={() => added ? onSave(list.filter((c) => c.text.toLowerCase() !== it.toLowerCase())) : add(it)} className={`rounded-full px-3 py-1 text-xs font-medium transition ${added ? "bg-indigo-600 text-white" : "bg-slate-100 text-slate-600 hover:bg-slate-200"}`}>{added ? "✓ " : "+ "}{it}</button>
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
              <span key={c.id} className="inline-flex items-center gap-1 rounded-full bg-indigo-50 px-3 py-1 text-sm text-indigo-700">{c.text}<button onClick={() => onSave(list.filter((x) => x.id !== c.id))} className="text-indigo-300 hover:text-rose-500"><X className="h-3.5 w-3.5" /></button></span>
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
          <button key={m.id} onClick={() => setMi(i)} className={`rounded-full px-3 py-1.5 text-sm font-semibold transition ${i === mi ? "bg-indigo-600 text-white" : "bg-slate-100 text-slate-500 hover:bg-slate-200"}`}>{m.name} <span className="opacity-70">({ratedCount(m)}/{list.length})</span></button>
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
        <input type="checkbox" checked={use} onChange={(e) => onSaveMeta({ useScores: e.target.checked })} className="h-4 w-4 accent-indigo-600" />
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
      <button onClick={auto} className="mb-3 inline-flex items-center gap-1.5 rounded-lg bg-indigo-600 px-4 py-2 text-sm font-semibold text-white hover:bg-indigo-700"><SparklesIcon className="h-4 w-4" /> Призначити автоматично</button>
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
                  <button key={m.id} onClick={() => setAssignee(c.id, m.id)} className={`inline-flex items-center gap-1 rounded-lg px-2.5 py-1.5 text-sm font-medium transition ${sel ? "bg-indigo-600 text-white" : "bg-slate-100 text-slate-600 hover:bg-slate-200"}`}>{m.name}{s != null && <span className={`text-xs ${sel ? "text-white/70" : "text-slate-400"}`}>· {s}</span>}</button>
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

      <button onClick={onFinish} className="mt-4 w-full rounded-2xl bg-indigo-600 py-3 font-bold text-white shadow-lg shadow-indigo-500/20 hover:bg-indigo-700">Готово — показати дошку 📋</button>
    </div>
  );
}

function ChoreBoard({ members, list, ratings, assignments, meta, saveMeta, onEdit }) {
  const totals = memberTotals(members, list, ratings, assignments, meta.useScores);
  const colors = ["#6366f1", "#ec4899", "#14b8a6", "#f59e0b", "#8b5cf6"];
  return (
    <div>
      <div className="mb-1 flex items-center gap-2"><span className="text-2xl">📋</span><h1 className="text-2xl font-extrabold text-slate-900">Наша дошка справ</h1></div>
      <p className="mb-4 rounded-lg bg-indigo-50 px-3 py-2 text-sm text-indigo-800">📌 Повісь це на видному місці у спільному просторі (холодильник, коридор). З СДУГ: <b>з очей — з голови</b>. Зелене — можна делегувати на важкий день.</p>

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
