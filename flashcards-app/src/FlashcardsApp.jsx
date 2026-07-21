import React, { useState, useEffect, useMemo, useRef, useCallback } from "react";
import * as XLSX from "xlsx";
import Papa from "papaparse";
import {
  BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer, CartesianGrid, Cell,
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
} from "lucide-react";

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
    T({ emoji: "☀️", title: "Підйом", note: "Тонізуючий напій + вітаміни.", time: "07:00", color: "orange", categoryId: morning.id }),
    T({ emoji: "💊", title: "Ломексин на обличчя (1-й раз)", note: "Прив'яжи до ранкового ритуалу — одразу після підйому.", time: "07:10", color: "pink", categoryId: morning.id }),
    T({ emoji: "🧘", title: "Каланетика", note: "Поки Міша спить, поки тихо. Таймер — і понеслі.", time: "07:15", color: "purple", categoryId: morning.id, goal: { type: "timed", minutes: 30 } }),
    T({ emoji: "🚿", title: "Душ", note: "Кетоконазол шампунь — вт/пт, залишити на 5 хв.", time: "07:45", color: "teal", categoryId: morning.id }),
    T({ emoji: "🍳", title: "Сніданок + Силібор + ліки від каменю", note: "Їжа вдома — не доставка. Щось просте і швидке.", time: "08:05", color: "yellow", categoryId: morning.id }),
    T({ emoji: "🧹", title: "Прибирання — одна зона", note: "Кухня → ванна → коридор → вітальня → спальня. Таймер і стоп.", time: "08:25", color: "green", categoryId: morning.id, goal: { type: "timed", minutes: 20 } }),

    T({ emoji: "💻", title: "Початок роботи", note: "Перші 25 хв — найважливіша задача без відволікань. Телефон вбік.", time: "09:00", color: "teal", categoryId: work.id, repeat: wd }),
    T({ emoji: "🍽️", title: "Обід + Ломексин (2-й раз)", note: "Прив'яжи другий прийом ломексину до обіду — легше пам'ятати.", time: "13:00", color: "orange", categoryId: work.id, repeat: wd }),
    T({ emoji: "✋", title: "Кінець роботи — закрити ноут", note: "Не тягнути роботу у вечір. Межа важлива.", time: "18:00", color: "pink", categoryId: work.id, repeat: wd }),

    T({ emoji: "📚", title: "Англійська", note: "Одразу після роботи. Duolingo, відео — регулярне.", time: "18:00", color: "purple", categoryId: evening.id, goal: { type: "timed", minutes: 15 } }),
    T({ emoji: "🍲", title: "Приготування вечері + вечеря", note: "Готуємо вдома. Смачно і без доставки.", time: "18:30", color: "green", categoryId: evening.id }),
    T({ emoji: "🎉", title: "Вільний час", note: "Без почуття провини. Ти вже зробила все головне за день.", time: "19:30", color: "pink", categoryId: evening.id }),
    T({ emoji: "💊", title: "Ліки від пролактину", note: "Тільки в понеділок.", time: null, color: "yellow", categoryId: evening.id, repeat: { type: "weekdays", days: [1] } }),
    T({ emoji: "😴", title: "Сон", note: "Телефон геть. Книга або щось легке. До 00:00 вже спати.", time: "23:00", reminder: "23:00", color: "purple", categoryId: evening.id }),

    T({ emoji: "💧", title: "Вода 1.5–2 л", note: "Протягом усього дня, не залпом.", time: null, color: "teal", categoryId: morning.id }),
    T({ emoji: "💊", title: "Золофт", note: "На тумбочці з вечора — щоб не пропустити.", time: "21:00", reminder: "21:00", color: "pink", categoryId: evening.id }),
  ];
  return { categories, tasks };
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
  const [section, setSection] = useState("studying"); // studying | routine
  const [sidebarCollapsed, setSidebarCollapsed] = useState(false);
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
      const st = await store.get("stats", {
        history: {},
        settings: { newPerDay: DEFAULT_NEW_PER_DAY },
      });
      if (!alive) return;
      setSection(ui.section === "routine" ? "routine" : "studying");
      setSidebarCollapsed(!!ui.sidebarCollapsed);
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
    setDecks([]);
    setGroups([]);
    setCardsByDeck({});
    setStats({ history: {}, settings: { newPerDay: DEFAULT_NEW_PER_DAY } });
    setView("home");
    flash("All data reset");
    window.dispatchEvent(new CustomEvent("routine-reset"));
  }, [decks, cardsByDeck, flash]);

  const exportAll = useCallback(async () => {
    const routine = await collectRoutineExport();
    const payload = {
      exportedAt: new Date().toISOString(),
      version: 2,
      decks,
      groups,
      cards: cardsByDeck,
      stats,
      routine,
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
      />

      <div className="flex min-h-screen min-w-0 flex-1 flex-col">
        {section === "routine" ? (
          <RoutineSection />
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
          <ImportView decks={decks} onImport={importCards} onCancel={() => setView("home")} />
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
function Sidebar({ section, collapsed, onSection, onToggle, studyingDue }) {
  const items = [
    { id: "studying", label: "Studying", icon: GraduationCap, badge: studyingDue },
    { id: "routine", label: "My Routine", icon: Sun, badge: 0 },
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

      <button
        onClick={onToggle}
        className="m-2 hidden items-center gap-2 rounded-lg px-2.5 py-2 text-xs font-medium text-slate-400 transition hover:bg-slate-100 hover:text-slate-600 lg:flex"
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
function ImportView({ decks, onImport, onCancel }) {
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
      <button onClick={onOpen} className="grid h-11 w-11 shrink-0 place-items-center rounded-xl bg-white/70 text-xl">{task.emoji || "⭐"}</button>
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
            <span className="grid h-12 w-12 place-items-center rounded-2xl text-2xl" style={{ backgroundColor: p.card }}>{task.emoji || "⭐"}</span>
            <div>
              <h2 className="text-lg font-bold text-slate-900">{task.title}</h2>
              {category && <span className="rounded-full px-2 py-0.5 text-[11px] font-semibold text-white" style={{ backgroundColor: getPastel(category.color).dot }}>{category.name}</span>}
            </div>
          </div>
          <button onClick={onClose} className="rounded-md p-1 text-slate-400 hover:bg-slate-100"><X className="h-5 w-5" /></button>
        </div>

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
      emoji, title: title.trim().slice(0, 50), note: note.trim(), color, date,
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

        {/* emoji + name */}
        <div className="mb-3 flex items-center gap-3">
          <span className="grid h-14 w-14 shrink-0 place-items-center rounded-2xl text-3xl" style={{ backgroundColor: p.card }}>{emoji}</span>
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
