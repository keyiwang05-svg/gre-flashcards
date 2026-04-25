// @ts-nocheck
import React, { useEffect, useMemo, useRef, useState } from "react";
import { AnimatePresence, motion } from "framer-motion";
import * as XLSX from "xlsx";
import {
  BookOpen,
  Brain,
  CheckCircle2,
  Download,
  Eye,
  Heart,
  HeartHandshake,
  PencilLine,
  RotateCcw,
  Search,
  Shuffle,
  Target,
  Trash2,
  Upload,
  XCircle,
} from "lucide-react";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Badge } from "@/components/ui/badge";
import { Progress } from "@/components/ui/progress";
import { Tabs, TabsList, TabsTrigger } from "@/components/ui/tabs";
import { track, trackPageView } from "@/lib/analytics";

const STORAGE_KEY = "gre_flashcards_app_v11";
const APP_DB_NAME = "gre-flashcards-db";
const APP_DB_VERSION = 1;
const APP_STATE_STORE = "app_state";
const APP_STATE_ID = "primary";
const FLASHCARD_UI_STATE_KEY = "gre_flashcards_flashcard_ui_v1";

const DEFAULT_SESSION_STATS = {
  reviewed: 0,
  known: 0,
  unknown: 0,
  quizCorrect: 0,
  quizWrong: 0,
  bbPairCorrect: 0,
  bbPairWrong: 0,
};

const DEFAULT_TASK_GOALS = { due: 30, wrong: 15, pairs: 20 };
const DEFAULT_DAILY_PLAN = {
  wordFinishDays: 30,
  wordReviewCount: 40,
  pairFinishDays: 20,
  pairReviewCount: 30,
  quizTarget: 85,
};

const LOAD_SAVED_DATA_ERROR_MESSAGE = "读取本地进度失败，但你仍然可以继续使用。\n建议先导出一次数据备份。";
const IDB_FALLBACK_MESSAGE = "当前浏览器不支持大容量离线缓存，已退回 localStorage 保存。大词库下仍可能超限，建议随时导出备份。";
const PERSIST_DATA_ERROR_MESSAGE = "当前进度未成功写入本地缓存。先点 Export Progress 备份，再继续学习更稳。";

function openAppDb() {
  return new Promise((resolve, reject) => {
    if (typeof window === "undefined" || !window.indexedDB) {
      reject(new Error("IndexedDB is not available"));
      return;
    }

    const request = window.indexedDB.open(APP_DB_NAME, APP_DB_VERSION);

    request.onupgradeneeded = () => {
      const db = request.result;
      if (!db.objectStoreNames.contains(APP_STATE_STORE)) {
        db.createObjectStore(APP_STATE_STORE, { keyPath: "id" });
      }
    };

    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error || new Error("Failed to open IndexedDB"));
  });
}

async function idbGetAppState() {
  const db = await openAppDb();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(APP_STATE_STORE, "readonly");
    const store = tx.objectStore(APP_STATE_STORE);
    const request = store.get(APP_STATE_ID);

    request.onsuccess = () => resolve(request.result?.payload || null);
    request.onerror = () => reject(request.error || new Error("Failed to read IndexedDB state"));
    tx.oncomplete = () => db.close();
    tx.onerror = () => reject(tx.error || new Error("IndexedDB read transaction failed"));
  });
}

async function idbSetAppState(payload) {
  const db = await openAppDb();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(APP_STATE_STORE, "readwrite");
    tx.objectStore(APP_STATE_STORE).put({ id: APP_STATE_ID, payload, updatedAt: Date.now() });

    tx.oncomplete = () => {
      db.close();
      resolve(true);
    };
    tx.onerror = () => reject(tx.error || new Error("IndexedDB write transaction failed"));
  });
}

const sampleWords = [
  {
    id: crypto.randomUUID(),
    word: "abate",
    phonetic: "",
    corePos: "v.",
    shortMeaning: "减弱；减轻；减少",
    sentiment: "neutral",
    memoryTip: "a-bate → 强度慢慢被打下去",
    tags: ["GRE", "verbs", "high-frequency"],
    stats: { seen: 0, known: 0, unknown: 0, quizCorrect: 0, quizWrong: 0 },
    srs: { interval: 1, due: Date.now(), streak: 0 },
    senses: [
      {
        label: "考法1",
        pos: "v.",
        zh: "减弱，减轻，减少",
        en: "to become weaker; to reduce in degree or intensity",
        exampleZh: "暴风雨没有减弱的迹象。",
        exampleEn: "The storm showed no signs of abating.",
        synonyms: ["mitigate", "diminish", "ebb", "subside"],
        antonyms: ["intensify", "escalate", "aggravate"],
        related: ["abatement"],
        collocations: [],
        sentiment: "neutral",
      },
    ],
  },
  {
    id: crypto.randomUUID(),
    word: "laconic",
    phonetic: "",
    corePos: "adj.",
    shortMeaning: "言简意赅的；寡言的",
    sentiment: "neutral",
    memoryTip: "像那种只回两三个词的人",
    tags: ["GRE", "adjectives"],
    stats: { seen: 0, known: 0, unknown: 0, quizCorrect: 0, quizWrong: 0 },
    srs: { interval: 1, due: Date.now(), streak: 0 },
    senses: [
      {
        label: "考法1",
        pos: "adj.",
        zh: "言简意赅的；简洁的",
        en: "using very few words",
        exampleZh: "他只给了一个很简短的回答。",
        exampleEn: "He gave a laconic reply.",
        synonyms: ["brief", "concise", "terse"],
        antonyms: ["verbose", "wordy"],
        related: [],
        collocations: ["laconic reply"],
        sentiment: "neutral",
      },
    ],
  },
];

const sampleSixChoicePairs = [
  { id: "sample_pair_1", a: "abate", b: "mitigate", zh: "减轻；缓和", source: "Built-in sample", frequency: "超高频" },
  { id: "sample_pair_2", a: "laconic", b: "brief", zh: "简洁的；简短的", source: "Built-in sample", frequency: "超高频" },
  { id: "sample_pair_3", a: "spurious", b: "specious", zh: "貌似合理但虚假的", source: "Built-in sample", frequency: "超超高频" },
  { id: "sample_pair_4", a: "ubiquitous", b: "pervasive", zh: "无处不在的；普遍存在的", source: "Built-in sample", frequency: "超高频" },
  { id: "sample_pair_5", a: "onerous", b: "arduous", zh: "艰巨的；费力的", source: "Built-in sample", frequency: "超高频" },
  { id: "sample_pair_6", a: "alluring", b: "enticing", zh: "诱人的；吸引人的", source: "Built-in sample", frequency: "超高频" },
];

function cleanText(value) {
  return String(value ?? "").trim();
}

function splitMultiValue(value) {
  if (Array.isArray(value)) return value.filter(Boolean).map(String).map((s) => s.trim()).filter(Boolean);
  return String(value || "")
    .split(/[;,；、|]/)
    .map((s) => s.trim())
    .filter(Boolean);
}

function uniq(arr) {
  return Array.from(new Set(arr.filter(Boolean)));
}

function formatDate(ts) {
  const d = new Date(ts);
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
}

function getStartOfLocalDay(timestamp = Date.now()) {
  const d = new Date(timestamp);
  d.setHours(0, 0, 0, 0);
  return d.getTime();
}

function hashString(input) {
  let hash = 2166136261;
  for (let i = 0; i < input.length; i += 1) {
    hash ^= input.charCodeAt(i);
    hash = Math.imul(hash, 16777619);
  }
  return hash >>> 0;
}

function seededShuffleArray(items, seedText) {
  const arr = [...items];
  let seed = hashString(seedText) || 1;

  function random() {
    seed += 0x6d2b79f5;
    let t = seed;
    t = Math.imul(t ^ (t >>> 15), t | 1);
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  }

  for (let i = arr.length - 1; i > 0; i -= 1) {
    const j = Math.floor(random() * (i + 1));
    [arr[i], arr[j]] = [arr[j], arr[i]];
  }

  return arr;
}

function sentimentLabel(value) {
  if (value === "positive") return "偏褒义";
  if (value === "negative") return "偏贬义";
  if (value === "neutral") return "中性";
  return "未标注";
}

function inferSentiment(text) {
  const t = cleanText(text).toLowerCase();
  if (!t) return "";
  const negativeMarkers = ["坏", "恶", "贬", "罪", "错", "异常", "病", "衰", "减", "废", "有害", "负面", "bad", "evil", "wrong", "abnormal", "harm", "decline", "reduce", "negative"];
  const positiveMarkers = ["好", "善", "赞", "益", "积极", "鼓励", "支持", "优秀", "benefit", "good", "positive", "encourage", "support", "generous", "noble"];
  if (negativeMarkers.some((m) => t.includes(m))) return "negative";
  if (positiveMarkers.some((m) => t.includes(m))) return "positive";
  return "neutral";
}

function senseKey(sense) {
  return [sense?.label, sense?.pos, sense?.zh, sense?.en].map((item) => cleanText(item).toLowerCase()).join("|");
}

function normalizeSense(raw = {}) {
  const syns = splitMultiValue(raw.synonyms || raw["近义词"]);
  return {
    id: raw.id || crypto.randomUUID(),
    favorite: Boolean(raw.favorite),
    label: cleanText(raw.label || raw.examName || raw["考法名称"]),
    pos: cleanText(raw.pos || raw.partOfSpeech || raw["单词词性"]),
    zh: cleanText(raw.zh || raw.zhMeaning || raw.meaningZh || raw["中文释义"]),
    en: cleanText(raw.en || raw.enMeaning || raw.meaningEn || raw["英文释义"]),
    exampleZh: cleanText(raw.exampleZh || raw["例句-中文"]),
    exampleEn: cleanText(raw.exampleEn || raw["例句-英文"]),
    synonyms: syns,
    antonyms: splitMultiValue(raw.antonyms || raw["反义词"]),
    related: splitMultiValue(raw.related || raw["相关词"] || raw.wordForms || raw["词形变换"]),
    collocations: splitMultiValue(raw.collocations || raw["词组搭配"]),
    sentiment: cleanText(raw.sentiment || raw["情感色彩"] || ""),
  };
}

function normalizeWord(raw = {}) {
  const rawSenses = Array.isArray(raw.senses)
    ? raw.senses.map(normalizeSense).filter((s) => s.zh || s.en)
    : [normalizeSense(raw)].filter((s) => s.zh || s.en);

  const senseMap = new Map();
  rawSenses.forEach((sense) => {
    const key = senseKey(sense);
    if (!key) return;
    if (!senseMap.has(key)) {
      senseMap.set(key, sense);
      return;
    }
    const prev = senseMap.get(key);
    senseMap.set(key, {
      ...prev,
      ...sense,
      favorite: Boolean(prev.favorite || sense.favorite),
    });
  });

  const senses = Array.from(senseMap.values());
  const firstSense = senses[0] || {};
  const sentiment = cleanText(raw.sentiment || raw["情感色彩"] || firstSense.sentiment) || inferSentiment(`${raw.shortMeaning || ""} ${firstSense.zh || ""} ${firstSense.en || ""}`);
  const defaultPending = Boolean((raw.stats?.unknown || 0) > 0 || (raw.stats?.quizWrong || 0) > 0 || raw.reviewState?.pending);

  return {
    id: raw.id || crypto.randomUUID(),
    favorite: Boolean(raw.favorite),
    word: cleanText(raw.word || raw["单词"]),
    phonetic: cleanText(raw.phonetic || raw["音标"]),
    corePos: cleanText(raw.corePos || raw["单词词性"] || firstSense.pos),
    shortMeaning: cleanText(raw.shortMeaning || raw.meaning || raw["单词词性及基本释义"] || firstSense.zh),
    memoryTip: cleanText(raw.memoryTip || raw.memorytip || raw["记忆方法"]),
    tags: splitMultiValue(raw.tags || raw["tags"]),
    stats: {
      seen: raw.stats?.seen || 0,
      known: raw.stats?.known || 0,
      unknown: raw.stats?.unknown || 0,
      quizCorrect: raw.stats?.quizCorrect || 0,
      quizWrong: raw.stats?.quizWrong || 0,
    },
    reviewState: {
      pending: raw.reviewState?.pending ?? defaultPending,
      correctStreak: raw.reviewState?.correctStreak || 0,
      wrongStreak: raw.reviewState?.wrongStreak || (defaultPending ? 1 : 0),
      priority: raw.reviewState?.priority || (defaultPending ? 2 : 0),
      lastWrongAt: raw.reviewState?.lastWrongAt || null,
      lastReviewedAt: raw.reviewState?.lastReviewedAt || null,
    },
    srs: raw.srs || { interval: 1, due: Date.now(), streak: 0 },
    sentiment,
    senses,
  };
}

function normalizeFrequency(value, fallback = "") {
  const raw = cleanText(value || fallback);
  const compact = raw.replace(/\s+/g, "");
  const upper = compact.toUpperCase();
  if (!compact) return "";

  if (
    compact.includes("超超超高频") ||
    upper === "SSS" ||
    upper.includes("TOP450") ||
    compact.includes("450对") ||
    /(^|[^0-9])450([^0-9]|$)/.test(compact)
  ) {
    return "超超超高频";
  }

  if (
    compact.includes("超超高频") ||
    upper === "SS" ||
    upper.includes("TOP300") ||
    compact.includes("300对") ||
    /(^|[^0-9])300([^0-9]|$)/.test(compact)
  ) {
    return "超超高频";
  }

  if (
    compact.includes("超高频") ||
    upper === "S" ||
    upper.includes("TOP254") ||
    compact.includes("254对") ||
    /(^|[^0-9])254([^0-9]|$)/.test(compact)
  ) {
    return "超高频";
  }

  return raw;
}

function frequencyRank(value) {
  const normalized = normalizeFrequency(value);
  if (normalized === "超超超高频") return 3;
  if (normalized === "超超高频") return 2;
  if (normalized === "超高频") return 1;
  return 0;
}

function normalizePair(raw = {}) {
  const defaultPending = Boolean((raw.stats?.wrong || 0) > 0 || raw.reviewState?.pending);
  const source = cleanText(raw.source || raw.source_pdf || raw["来源"]);
  const normalizedFrequency = normalizeFrequency(raw.frequency || raw["频率"], source);
  return {
    id: raw.id || crypto.randomUUID(),
    a: cleanText(raw.a || raw.word1),
    b: cleanText(raw.b || raw.word2),
    zh: cleanText(raw.zh || raw.pair_zh || raw["中文释义"]),
    source,
    frequency: normalizedFrequency,
    page: raw.page ?? null,
    stats: {
      seen: raw.stats?.seen || 0,
      correct: raw.stats?.correct || 0,
      wrong: raw.stats?.wrong || 0,
      lastWrongAt: raw.stats?.lastWrongAt || null,
      lastCorrectAt: raw.stats?.lastCorrectAt || null,
    },
    reviewState: {
      pending: raw.reviewState?.pending ?? defaultPending,
      correctStreak: raw.reviewState?.correctStreak || 0,
      wrongStreak: raw.reviewState?.wrongStreak || (defaultPending ? 1 : 0),
      priority: raw.reviewState?.priority || (defaultPending ? 2 : 0),
      lastWrongAt: raw.reviewState?.lastWrongAt || raw.stats?.lastWrongAt || null,
      lastReviewedAt: raw.reviewState?.lastReviewedAt || null,
    },
  };
}

function wordKey(word) {
  return cleanText(word?.word).toLowerCase();
}

function pairKey(pair) {
  return [cleanText(pair?.a).toLowerCase(), cleanText(pair?.b).toLowerCase()].sort().join("=");
}

function pickBetterFrequency(a, b) {
  return frequencyRank(b) > frequencyRank(a) ? normalizeFrequency(b) : normalizeFrequency(a);
}

function mergeWordRecords(existing, incoming) {
  const senseMap = new Map();
  [...(existing.senses || []), ...(incoming.senses || [])].forEach((sense) => {
    const normalized = normalizeSense(sense);
    const key = senseKey(normalized);
    if (!senseMap.has(key)) {
      senseMap.set(key, normalized);
      return;
    }
    const prev = senseMap.get(key);
    senseMap.set(key, {
      ...prev,
      ...normalized,
      favorite: Boolean(prev.favorite || normalized.favorite),
    });
  });

  return normalizeWord({
    ...incoming,
    id: existing.id,
    favorite: Boolean(existing.favorite || incoming.favorite),
    stats: existing.stats,
    reviewState: existing.reviewState,
    srs: existing.srs,
    senses: Array.from(senseMap.values()),
  });
}

function mergeWords(existingWords, incomingWords, mode = "append") {
  if (mode === "replace") return incomingWords.map(normalizeWord);

  const map = new Map(existingWords.map((word) => [wordKey(word), normalizeWord(word)]));
  incomingWords.map(normalizeWord).forEach((incoming) => {
    const key = wordKey(incoming);
    if (!key) return;
    if (!map.has(key)) {
      map.set(key, incoming);
      return;
    }
    map.set(key, mergeWordRecords(map.get(key), incoming));
  });

  return Array.from(map.values());
}

function dedupePairs(pairs) {
  const map = new Map();
  pairs.map(normalizePair).forEach((pair) => {
    const key = pairKey(pair);
    if (!key) return;
    if (!map.has(key)) {
      map.set(key, pair);
      return;
    }

    const existing = map.get(key);
    map.set(key, {
      ...existing,
      ...pair,
      source: frequencyRank(existing.frequency) >= frequencyRank(pair.frequency) ? existing.source || pair.source : pair.source || existing.source,
      frequency: pickBetterFrequency(existing.frequency, pair.frequency),
      stats: {
        seen: (existing.stats?.seen || 0) + (pair.stats?.seen || 0),
        correct: (existing.stats?.correct || 0) + (pair.stats?.correct || 0),
        wrong: (existing.stats?.wrong || 0) + (pair.stats?.wrong || 0),
        lastWrongAt: Math.max(existing.stats?.lastWrongAt || 0, pair.stats?.lastWrongAt || 0) || null,
        lastCorrectAt: Math.max(existing.stats?.lastCorrectAt || 0, pair.stats?.lastCorrectAt || 0) || null,
      },
      reviewState: {
        pending: Boolean(existing.reviewState?.pending || pair.reviewState?.pending),
        correctStreak: Math.max(existing.reviewState?.correctStreak || 0, pair.reviewState?.correctStreak || 0),
        wrongStreak: Math.max(existing.reviewState?.wrongStreak || 0, pair.reviewState?.wrongStreak || 0),
        priority: Math.max(existing.reviewState?.priority || 0, pair.reviewState?.priority || 0),
        lastWrongAt: Math.max(existing.reviewState?.lastWrongAt || 0, pair.reviewState?.lastWrongAt || 0) || null,
        lastReviewedAt: Math.max(existing.reviewState?.lastReviewedAt || 0, pair.reviewState?.lastReviewedAt || 0) || null,
      },
    });
  });
  return Array.from(map.values());
}

function groupRowsToWords(rows) {
  const words = [];
  let current = null;

  rows.forEach((row) => {
    const wordValue = cleanText(row["单词"] || row.word);
    const isNewWord = Boolean(wordValue);

    if (isNewWord) {
      if (current) words.push(normalizeWord(current));
      current = {
        word: wordValue,
        corePos: cleanText(row["单词词性"]),
        shortMeaning: cleanText(row["单词词性及基本释义"]),
        memoryTip: cleanText(row["记忆方法"]),
        tags: [cleanText(row["List"]), cleanText(row["Unit"]), "GRE"].filter(Boolean),
        senses: [],
      };
    }

    if (!current) return;

    current.senses.push({
      label: row["考法名称"],
      pos: row["单词词性"],
      zh: row["中文释义"],
      en: row["英文释义"],
      exampleZh: row["例句-中文"],
      exampleEn: row["例句-英文"],
      synonyms: row["近义词"],
      antonyms: row["反义词"],
      related: [row["词形变换"], row["相关词"]].filter(Boolean).join(";"),
      collocations: row["词组搭配"],
      sentiment: row["情感色彩"],
    });
  });

  if (current) words.push(normalizeWord(current));
  return words.filter((w) => w.word);
}

function splitCSVLine(line) {
  const result = [];
  let current = "";
  let inQuotes = false;

  for (let i = 0; i < line.length; i += 1) {
    const char = line[i];
    const next = line[i + 1];
    if (char === '"') {
      if (inQuotes && next === '"') {
        current += '"';
        i += 1;
      } else {
        inQuotes = !inQuotes;
      }
    } else if (char === "," && !inQuotes) {
      result.push(current);
      current = "";
    } else {
      current += char;
    }
  }

  result.push(current);
  return result;
}

function parseCSV(text) {
  const lines = text.split(/\r?\n/).filter((line) => line.trim());
  if (!lines.length) return [];

  const headers = splitCSVLine(lines[0]).map((h) => h.trim());
  const rows = lines.slice(1).map((line) => {
    const values = splitCSVLine(line);
    const obj = {};
    headers.forEach((h, i) => {
      obj[h] = values[i] ?? "";
    });
    return obj;
  });

  if (headers.includes("单词") || headers.includes("考法名称")) return groupRowsToWords(rows);
  return rows.map(normalizeWord).filter((item) => item.word);
}

function shuffleArray(arr) {
  const copy = [...arr];
  for (let i = copy.length - 1; i > 0; i -= 1) {
    const j = Math.floor(Math.random() * (i + 1));
    [copy[i], copy[j]] = [copy[j], copy[i]];
  }
  return copy;
}

function pickRandom(items) {
  if (!items.length) return null;
  return items[Math.floor(Math.random() * items.length)];
}

async function getPdfJs() {
  const pdfjsLib = await import("pdfjs-dist");
  if (pdfjsLib.GlobalWorkerOptions && !pdfjsLib.GlobalWorkerOptions.workerSrc) {
    pdfjsLib.GlobalWorkerOptions.workerSrc = `https://cdnjs.cloudflare.com/ajax/libs/pdf.js/${pdfjsLib.version}/pdf.worker.min.mjs`;
  }
  return pdfjsLib;
}

function buildLinesFromPdfTextContent(textContent) {
  const buckets = new Map();
  textContent.items.forEach((item) => {
    const str = cleanText(item.str);
    if (!str) return;
    const y = Math.round((item.transform?.[5] || 0) * 10) / 10;
    const x = item.transform?.[4] || 0;
    if (!buckets.has(y)) buckets.set(y, []);
    buckets.get(y).push({ str, x });
  });

  return Array.from(buckets.entries())
    .sort((a, b) => Number(b[0]) - Number(a[0]))
    .map(([, items]) => items.sort((a, b) => a.x - b.x).map((item) => item.str).join(" ").replace(/\s+/g, " ").trim())
    .filter(Boolean);
}

function splitBlocksBySeparator(lines) {
  const blocks = [];
  let current = [];

  const flush = () => {
    const text = current.join("\n").trim();
    if (text) blocks.push(text);
    current = [];
  };

  lines.forEach((line) => {
    const cleaned = cleanText(line);
    const dashOnly = cleaned.replace(/-/g, "").trim() === "" && cleaned.length >= 20;
    if (dashOnly) {
      flush();
      return;
    }
    current.push(cleaned);
  });
  flush();
  return blocks;
}

function parseBbPairBlock(blockText) {
  const lines = blockText
    .split("\n")
    .map((line) => cleanText(line))
    .filter(Boolean)
    .filter((line) => !/^记忆法/.test(line))
    .filter((line) => !/^bibiji\.site/i.test(line))
    .filter((line) => !/^\d+$/.test(line));

  const relationLine = lines.find((line) => line.includes("=") && line.split("=").length >= 3) || "";
  const words = lines
    .filter((line) => !line.includes("="))
    .map((line) => line.match(/^([A-Za-z][A-Za-z'\- ]{0,80})/))
    .filter(Boolean)
    .map((m) => cleanText(m[1]))
    .slice(0, 2);

  if (words.length < 2 && relationLine) {
    const parts = relationLine.split("=").map((part) => cleanText(part)).filter(Boolean);
    if (parts.length >= 2) {
      if (!words[0]) words[0] = parts[0];
      if (!words[1]) words[1] = parts[1];
    }
  }

  if (!words[0] || !words[1]) return null;
  const zh = relationLine ? cleanText(relationLine.split("=").pop()).replace(/^[“"]|[”"]$/g, "") : "";
  return { a: words[0], b: words[1], zh };
}

async function extractPairsFromPdf(arrayBuffer, filename = "PDF") {
  const pdfjsLib = await getPdfJs();
  const loadingTask = pdfjsLib.getDocument({ data: arrayBuffer, disableWorker: true });
  const pdf = await loadingTask.promise;
  const pairs = [];
  const frequency = normalizeFrequency(filename, filename);

  for (let pageNum = 1; pageNum <= pdf.numPages; pageNum += 1) {
    const page = await pdf.getPage(pageNum);
    const textContent = await page.getTextContent();
    const lines = buildLinesFromPdfTextContent(textContent);
    const blocks = splitBlocksBySeparator(lines);
    blocks.forEach((block) => {
      const parsed = parseBbPairBlock(block);
      if (!parsed) return;
      pairs.push({
        id: crypto.randomUUID(),
        a: parsed.a,
        b: parsed.b,
        zh: parsed.zh,
        source: filename,
        frequency,
        page: pageNum,
      });
    });
  }

  return dedupePairs(pairs);
}

function getPlanStatusMeta(status) {
  if (status === "done") return { label: "已完成", badgeClass: "border-emerald-200 bg-emerald-100 text-emerald-700" };
  if (status === "warning") return { label: "差一点", badgeClass: "border-amber-200 bg-amber-100 text-amber-700" };
  return { label: "未完成", badgeClass: "border-rose-200 bg-rose-100 text-rose-700" };
}

function buildQuizQuestion(wordPool, mode, pairPool) {
  if (mode === "bb_pairs") {
    const target = pickRandom(pairPool.filter((pair) => pair.a && pair.b));
    if (!target) return null;
    const correctAnswer = Math.random() > 0.5 ? target.a : target.b;
    const promptWord = correctAnswer === target.a ? target.b : target.a;
    const distractorPool = uniq(pairPool.filter((p) => p.id !== target.id).flatMap((p) => [p.a, p.b]));
    return {
      questionType: "bb_pairs",
      pairId: target.id,
      promptWord,
      correctAnswer,
      choices: shuffleArray([
        correctAnswer,
        ...shuffleArray(distractorPool).filter((item) => item.toLowerCase() !== correctAnswer.toLowerCase() && item.toLowerCase() !== promptWord.toLowerCase()).slice(0, 5),
      ]).slice(0, 6),
      promptZh: target.zh || "",
      explanation: `${target.a} = ${target.b}`,
      frequency: target.frequency || "",
    };
  }

  const validWords = wordPool.filter((w) => w.senses.some((s) => (mode === "equivalence" ? (s.synonyms || []).length : (s.antonyms || []).length)));
  const target = pickRandom(validWords);
  if (!target) return null;

  const targetSense = target.senses.find((s) => (mode === "equivalence" ? (s.synonyms || []).length : (s.antonyms || []).length)) || target.senses[0];
  const correctPool = mode === "equivalence" ? targetSense.synonyms || [] : targetSense.antonyms || [];
  const correctAnswer = pickRandom(correctPool);
  if (!correctAnswer) return null;

  const distractors = uniq(
    validWords
      .filter((w) => w.id !== target.id)
      .flatMap((w) =>
        w.senses.flatMap((s) => [
          ...(mode === "equivalence" ? s.synonyms || [] : s.antonyms || []),
          ...(mode === "equivalence" ? s.antonyms || [] : s.synonyms || []),
        ])
      )
  ).filter((item) => item.toLowerCase() !== correctAnswer.toLowerCase());

  return {
    questionType: mode,
    targetWordId: target.id,
    promptWord: target.word,
    promptZh: targetSense.zh || "",
    promptEn: targetSense.en || "",
    correctAnswer,
    choices: shuffleArray([correctAnswer, ...shuffleArray(distractors).slice(0, 3)]),
    explanation: mode === "equivalence" ? `正确项和 ${target.word} 在这个考法下语义最接近。` : `正确项和 ${target.word} 在这个考法下构成反义关系。`,
    frequency: "",
  };
}

async function loadBuiltinLibraries() {
  const [wordsRes, pairsRes] = await Promise.all([
    fetch("/data/builtinWords.json"),
    fetch("/data/builtinPairs.json"),
  ]);

  if (!wordsRes.ok || !pairsRes.ok) {
    throw new Error("Failed to load built-in libraries");
  }

  const wordsJson = await wordsRes.json();
  const pairsJson = await pairsRes.json();

  return {
    words: Array.isArray(wordsJson) ? wordsJson.map(normalizeWord) : [],
    pairs: Array.isArray(pairsJson) ? pairsJson.map(normalizePair) : [],
  };
}
export default function GREFlashcardsApp() {
  const [words, setWords] = useState([]);
  const [sixChoicePairs, setSixChoicePairs] = useState([]);
  const [currentIndex, setCurrentIndex] = useState(0);
  const [flashcardCompleteMessage, setFlashcardCompleteMessage] = useState("");
  const [flipped, setFlipped] = useState(false);
  const [mode, setMode] = useState("all");
  const [studyView, setStudyView] = useState("flashcards");
  const [flashcardMode, setFlashcardMode] = useState("recognition");
  const [importWordMode, setImportWordMode] = useState("append");
  const [pairImportMode, setPairImportMode] = useState("append");
  const [quizMode, setQuizMode] = useState("equivalence");
  const [quizHintMode, setQuizHintMode] = useState("study");
  const [search, setSearch] = useState("");
  const [pairSearch, setPairSearch] = useState("");
  const [pairReviewMode, setPairReviewMode] = useState("mix");
  const [sessionStats, setSessionStats] = useState(DEFAULT_SESSION_STATS);
  const [dailyStats, setDailyStats] = useState({});
  const [taskGoals, setTaskGoals] = useState(DEFAULT_TASK_GOALS);
  const [dailyPlan, setDailyPlan] = useState(DEFAULT_DAILY_PLAN);
  const [orderMode, setOrderMode] = useState("ordered");
  const [flashcardFilter, setFlashcardFilter] = useState("all");
  const [revealLevel, setRevealLevel] = useState(0);
  const [quizQuestion, setQuizQuestion] = useState(null);
  const [selectedChoice, setSelectedChoice] = useState("");
  const [quizChecked, setQuizChecked] = useState(false);
  const [retrievalInput, setRetrievalInput] = useState("");
  const [storageReady, setStorageReady] = useState(false);
  const [storageMessage, setStorageMessage] = useState("");
  const [shuffleSeed, setShuffleSeed] = useState(0);
  const [reviewSummaryTab, setReviewSummaryTab] = useState("words");
  const [reviewSummarySearch, setReviewSummarySearch] = useState("");
  const fileRef = useRef(null);
  const persistTimeoutRef = useRef(null);
  const flashcardSessionRef = useRef<{
  id: string;
  mode: string;
  filter: string;
  flashcardMode: string;
  sessionSize: number;
  startedAt: number;
  wordsSeen: number;
  knownCount: number;
  unknownCount: number;
  uniqueWordIds: Set<string>;
  flushed: boolean;
} | null>(null);

const restoredFlashcardStateRef = useRef(null);
const shouldRebuildDeckRef = useRef(true);

  useEffect(() => {
    trackPageView("home");
  }, []);
  useEffect(() => {
  const handleBeforeUnload = () => {
    flushFlashcardSession("page_unload");
  };

  window.addEventListener("beforeunload", handleBeforeUnload);

  return () => {
    window.removeEventListener("beforeunload", handleBeforeUnload);
    flushFlashcardSession("component_unmount");
  };
}, []);

  useEffect(() => {
    if (studyView !== "quiz") return;
    track("start_quiz", {
      quiz_mode: quizMode,
      pool_size: quizMode === "bb_pairs" ? filteredPairs.length : filteredWords.length,
    });
  }, [studyView, quizMode]);

  useEffect(() => {
  let cancelled = false;

  async function loadSavedState() {
    try {
      let parsed = null;

      try {
        parsed = await idbGetAppState();
      } catch (idbError) {
        console.warn("IndexedDB load failed, falling back to localStorage", idbError);
      }

      if (!parsed) {
        const saved =
          typeof window !== "undefined" ? window.localStorage.getItem(STORAGE_KEY) : null;
        if (saved) parsed = JSON.parse(saved);
      }

      if (cancelled) return;

      if (parsed) {
        if (Array.isArray(parsed.words) && parsed.words.length) {
          setWords(parsed.words.map(normalizeWord));
        } else {
          const builtin = await loadBuiltinLibraries();
          if (!cancelled) setWords(builtin.words);
        }

        if (Array.isArray(parsed.sixChoicePairs) && parsed.sixChoicePairs.length) {
          setSixChoicePairs(parsed.sixChoicePairs.map(normalizePair));
        } else {
          const builtin = await loadBuiltinLibraries();
          if (!cancelled) setSixChoicePairs(builtin.pairs);
        }

        setSessionStats(parsed.sessionStats || DEFAULT_SESSION_STATS);
        setDailyStats(parsed.dailyStats || {});
        setTaskGoals(parsed.taskGoals || DEFAULT_TASK_GOALS);
        setDailyPlan({ ...DEFAULT_DAILY_PLAN, ...(parsed.dailyPlan || {}) });
        try {
  const savedFlashcardState =
    typeof window !== "undefined"
      ? window.localStorage.getItem(FLASHCARD_UI_STATE_KEY)
      : null;

  if (savedFlashcardState) {
    const flashcardState = JSON.parse(savedFlashcardState);

    if (flashcardState.mode) setMode(flashcardState.mode);
    if (flashcardState.studyView) setStudyView(flashcardState.studyView);
    if (flashcardState.flashcardMode) setFlashcardMode(flashcardState.flashcardMode);
    if (flashcardState.flashcardFilter) setFlashcardFilter(flashcardState.flashcardFilter);
    if (flashcardState.orderMode) setOrderMode(flashcardState.orderMode);
    if (typeof flashcardState.search === "string") setSearch(flashcardState.search);
    if (Number.isFinite(flashcardState.shuffleSeed)) setShuffleSeed(flashcardState.shuffleSeed);

    if (Array.isArray(flashcardState.sessionOrderIds) && flashcardState.sessionOrderIds.length) {
      const restoredIndex = Number.isFinite(flashcardState.currentIndex)
        ? flashcardState.currentIndex
        : 0;

      setSessionOrderIds(flashcardState.sessionOrderIds);
      setCurrentIndex(restoredIndex);
      restoredFlashcardStateRef.current = {
        sessionOrderIds: flashcardState.sessionOrderIds,
        currentIndex: restoredIndex,
        currentWordId: flashcardState.currentWordId || flashcardState.sessionOrderIds[restoredIndex] || null,
      };
      shouldRebuildDeckRef.current = false;
    }
  }
} catch (flashcardStateError) {
  console.warn("Failed to restore flashcard UI state", flashcardStateError);
}
      } else {
        const builtin = await loadBuiltinLibraries();
        if (cancelled) return;

        setWords(builtin.words);
        setSixChoicePairs(builtin.pairs);
      }
    } catch (e) {
      console.error("Failed to load saved data", e);

      try {
        const builtin = await loadBuiltinLibraries();
        if (!cancelled) {
          setWords(builtin.words);
          setSixChoicePairs(builtin.pairs);
        }
      } catch (builtinError) {
        console.error("Failed to load built-in libraries", builtinError);
        if (!cancelled) {
          setWords(sampleWords.map(normalizeWord));
          setSixChoicePairs(sampleSixChoicePairs.map(normalizePair));
          setStorageMessage(LOAD_SAVED_DATA_ERROR_MESSAGE);
        }
      }
    } finally {
      if (!cancelled) setStorageReady(true);
    }
  }

  loadSavedState();

  return () => {
    cancelled = true;
  };
}, []);

  useEffect(() => {
    if (!storageReady) return;
    if (persistTimeoutRef.current) clearTimeout(persistTimeoutRef.current);

    persistTimeoutRef.current = setTimeout(() => {
      const payload = { words, sixChoicePairs, sessionStats, dailyStats, taskGoals, dailyPlan };

      Promise.resolve()
        .then(() => idbSetAppState(payload))
        .then(() => {
          try {
            if (typeof window !== "undefined") {
              window.localStorage.setItem(`${STORAGE_KEY}_meta`, JSON.stringify({ wordsCount: words.length, pairsCount: sixChoicePairs.length, updatedAt: Date.now() }));
            }
          } catch (metaError) {
            console.warn("Meta save skipped", metaError);
          }
          setStorageMessage("");
        })
        .catch((idbError) => {
          console.error("Failed to persist data", idbError);
          try {
            if (typeof window !== "undefined") {
              window.localStorage.setItem(STORAGE_KEY, JSON.stringify(payload));
              setStorageMessage(IDB_FALLBACK_MESSAGE);
              return;
            }
          } catch (fallbackError) {
            console.error("Local fallback also failed", fallbackError);
          }
          setStorageMessage(PERSIST_DATA_ERROR_MESSAGE);
        });
    }, 600);

    return () => {
      if (persistTimeoutRef.current) clearTimeout(persistTimeoutRef.current);
    };
  }, [words, sixChoicePairs, sessionStats, dailyStats, taskGoals, dailyPlan, storageReady]);

  const wrongWordBook = useMemo(() => words.filter((w) => (w.stats?.unknown || 0) > 0 || (w.stats?.quizWrong || 0) > 0), [words]);

  const todayReviewTargetActual = useMemo(() => {
    const todayStart = getStartOfLocalDay(Date.now());

    const eligibleReviewWords = wrongWordBook.filter((w) => {
      const lastWrongAt = w.reviewState?.lastWrongAt || 0;
      const hasWrongHistory =
        (w.stats?.unknown || 0) > 0 ||
        (w.stats?.quizWrong || 0) > 0 ||
        Boolean(w.reviewState?.pending);

      return hasWrongHistory && lastWrongAt > 0 && lastWrongAt < todayStart;
    });

    return Math.min(
      eligibleReviewWords.length,
      Math.max(0, dailyPlan.wordReviewCount || 0)
    );
  }, [wrongWordBook, dailyPlan.wordReviewCount]);
  const stubbornWordBook = useMemo(() => words.filter((w) => (w.reviewState?.priority || 0) >= 4 || (w.reviewState?.wrongStreak || 0) >= 2), [words]);
  const favoriteWords = useMemo(() => words.filter((w) => w.favorite), [words]);
  const favoriteSenseCount = useMemo(() => words.reduce((sum, w) => sum + w.senses.filter((s) => s.favorite).length, 0), [words]);
  const wrongPairBook = useMemo(() => sixChoicePairs.filter((p) => (p.stats?.wrong || 0) > 0), [sixChoicePairs]);

  const wrongWordSummary = useMemo(() => {
    return wrongWordBook
      .map((w) => {
        const firstSense = w.senses?.[0] || {};
        const totalWrong = (w.stats?.unknown || 0) + (w.stats?.quizWrong || 0);

        return {
          id: w.id,
          word: w.word,
          meaning: w.shortMeaning || firstSense.zh || "",
          en: firstSense.en || "",
          unknownCount: w.stats?.unknown || 0,
          quizWrongCount: w.stats?.quizWrong || 0,
          totalWrong,
          lastWrongAt: w.reviewState?.lastWrongAt || 0,
        };
      })
      .sort((a, b) => b.totalWrong - a.totalWrong || (b.lastWrongAt || 0) - (a.lastWrongAt || 0));
  }, [wrongWordBook]);

  const wrongPairSummary = useMemo(() => {
    return wrongPairBook
      .map((p) => ({
        id: p.id,
        a: p.a,
        b: p.b,
        zh: p.zh || "",
        wrongCount: p.stats?.wrong || 0,
        lastWrongAt: p.stats?.lastWrongAt || p.reviewState?.lastWrongAt || 0,
        frequency: p.frequency || "",
        source: p.source || "",
      }))
      .sort((a, b) => b.wrongCount - a.wrongCount || (b.lastWrongAt || 0) - (a.lastWrongAt || 0));
  }, [wrongPairBook]);

  const filteredWrongWordSummary = useMemo(() => {
    const q = reviewSummarySearch.trim().toLowerCase();
    if (!q) return wrongWordSummary;

    return wrongWordSummary.filter((item) =>
      [item.word, item.meaning, item.en].join(" | ").toLowerCase().includes(q)
    );
  }, [wrongWordSummary, reviewSummarySearch]);

  const filteredWrongPairSummary = useMemo(() => {
    const q = reviewSummarySearch.trim().toLowerCase();
    if (!q) return wrongPairSummary;

    return wrongPairSummary.filter((item) =>
      [item.a, item.b, item.zh, item.frequency, item.source].join(" | ").toLowerCase().includes(q)
    );
  }, [wrongPairSummary, reviewSummarySearch]);

  const pairFrequencyStats = useMemo(() => ({
    sss: sixChoicePairs.filter((pair) => normalizeFrequency(pair.frequency) === "超超超高频").length,
    ss: sixChoicePairs.filter((pair) => normalizeFrequency(pair.frequency) === "超超高频").length,
    s: sixChoicePairs.filter((pair) => normalizeFrequency(pair.frequency) === "超高频").length,
  }), [sixChoicePairs]);
  const highFrequencyPairCount = useMemo(() => pairFrequencyStats.sss + pairFrequencyStats.ss + pairFrequencyStats.s, [pairFrequencyStats]);


  const filteredWords = useMemo(() => {
    const now = Date.now();
    const untouchedWords = words.filter((w) => w.stats.seen === 0 && (w.stats.quizCorrect || 0) === 0 && (w.stats.quizWrong || 0) === 0);
    const todayTaskNewCount = Math.ceil((untouchedWords.length || words.length || 0) / Math.max(1, dailyPlan.wordFinishDays || 1));
    const todayTaskReviewCount = Math.max(0, dailyPlan.wordReviewCount || 0);
    const randomizedWrongWords = shuffleArray(wrongWordBook);

    let base = [...words];
    if (mode === "wrong") base = wrongWordBook;
    if (mode === "stubborn") base = stubbornWordBook;
    if (mode === "favorite_words") base = words.filter((w) => w.favorite);
    if (mode === "favorite_senses") base = words.filter((w) => w.senses.some((s) => s.favorite));
    if (mode === "new") base = untouchedWords;
    if (mode === "hard") base = words.filter((w) => w.senses.length >= 2 || (w.stats.unknown || 0) >= 2 || (w.stats.quizWrong || 0) >= 2);

    if (flashcardFilter === "review") {
      base = randomizedWrongWords.slice(0, todayTaskReviewCount);
    }

    if (flashcardFilter === "task") {
      const reviewSlice = randomizedWrongWords.slice(0, todayTaskReviewCount);
      const reviewIds = new Set(reviewSlice.map((w) => w.id));
      const newSlice = untouchedWords.filter((w) => !reviewIds.has(w.id)).slice(0, todayTaskNewCount);
      const taskIds = new Set([...reviewSlice, ...newSlice].map((w) => w.id));
      base = words.filter((w) => taskIds.has(w.id));
    }

    if (flashcardFilter === "today_new") {
      base = untouchedWords.slice(0, todayTaskNewCount);
    }

    if (flashcardFilter === "new") {
      base = untouchedWords;
    }

    const q = search.trim().toLowerCase();
    if (q) {
      base = base.filter((w) => [
        w.word,
        w.shortMeaning,
        w.memoryTip,
        ...(w.tags || []),
        ...w.senses.flatMap((s) => [s.zh, s.en, s.exampleZh, s.exampleEn, ...(s.synonyms || []), ...(s.antonyms || [])]),
      ].join(" | ").toLowerCase().includes(q));
    }

    return base.sort((a, b) => {
      const pendingDelta = Number(Boolean(b.reviewState?.pending)) - Number(Boolean(a.reviewState?.pending));
      if (pendingDelta) return pendingDelta;
      const priorityDelta = (b.reviewState?.priority || 0) - (a.reviewState?.priority || 0);
      if (priorityDelta) return priorityDelta;
      return (b.reviewState?.lastWrongAt || 0) - (a.reviewState?.lastWrongAt || 0);
    });
  }, [words, wrongWordBook, stubbornWordBook, mode, flashcardFilter, search, dailyPlan.wordFinishDays, dailyPlan.wordReviewCount, shuffleSeed, orderMode]);

  const filteredPairs = useMemo(() => {
    let base = sixChoicePairs;
    if (pairReviewMode === "wrong") base = wrongPairBook;
    if (pairReviewMode === "mix") base = sixChoicePairs.filter((p) => frequencyRank(p.frequency) > 0);
    if (pairReviewMode === "sss") base = sixChoicePairs.filter((p) => normalizeFrequency(p.frequency) === "超超超高频");
    if (pairReviewMode === "ss") base = sixChoicePairs.filter((p) => normalizeFrequency(p.frequency) === "超超高频");
    if (pairReviewMode === "s") base = sixChoicePairs.filter((p) => normalizeFrequency(p.frequency) === "超高频");

    const q = pairSearch.trim().toLowerCase();
    if (!q) return base;
    return base.filter((p) => [p.a, p.b, p.zh, p.source, p.frequency].join(" | ").toLowerCase().includes(q));
  }, [sixChoicePairs, wrongPairBook, pairReviewMode, pairSearch]);

  const filteredWordMembershipSignature = useMemo(() => [...filteredWords.map((w) => w.id)].sort().join("|"), [filteredWords]);
  const [sessionOrderIds, setSessionOrderIds] = useState([]);
  useEffect(() => {
  if (!storageReady) return;

  try {
    if (typeof window !== "undefined") {
      window.localStorage.setItem(
        FLASHCARD_UI_STATE_KEY,
        JSON.stringify({
          currentIndex,
          sessionOrderIds,
          mode,
          studyView,
          flashcardMode,
          flashcardFilter,
          orderMode,
          shuffleSeed,
          search,
          savedAt: Date.now(),
        })
      );
    }
  } catch (flashcardStateError) {
    console.warn("Failed to save flashcard UI state", flashcardStateError);
  }
}, [
  currentIndex,
  sessionOrderIds,
  mode,
  studyView,
  flashcardMode,
  flashcardFilter,
  orderMode,
  shuffleSeed,
  search,
  storageReady,
]);

  useEffect(() => {
    setCurrentIndex((prev) => {
      if (!sessionOrderIds.length) return 0;
      return Math.min(prev, sessionOrderIds.length - 1);
    });
  }, [sessionOrderIds]);

  useEffect(() => {
    if (studyView !== "quiz") return;
    const q = buildQuizQuestion(filteredWords.length ? filteredWords : words, quizMode, filteredPairs.length ? filteredPairs : sixChoicePairs);
    setQuizQuestion(q);
    setSelectedChoice("");
    setQuizChecked(false);
  }, [studyView, quizMode, filteredWords, words, filteredPairs, sixChoicePairs]);

  useEffect(() => {
  const ids = filteredWords.map((w) => w.id);

  if (!ids.length) {
    setSessionOrderIds([]);
    setCurrentIndex(0);
    shouldRebuildDeckRef.current = false;
    restoredFlashcardStateRef.current = null;
    return;
  }

  const restoredState = restoredFlashcardStateRef.current;

  if (restoredState) {
    const validIdSet = new Set(words.map((w) => w.id));
    const restoredIds = Array.isArray(restoredState.sessionOrderIds)
      ? restoredState.sessionOrderIds.filter((id) => validIdSet.has(id))
      : [];
    const nextIds = restoredIds.length
      ? restoredIds
      : orderMode === "ordered"
        ? ids
        : shuffleArray(ids);

    const preferredWordId = restoredState.currentWordId || nextIds[restoredState.currentIndex || 0] || null;
    const preferredIndex = preferredWordId ? nextIds.findIndex((id) => id === preferredWordId) : -1;
    const nextIndex = preferredIndex >= 0
      ? preferredIndex
      : Math.min(Math.max(restoredState.currentIndex || 0, 0), nextIds.length - 1);

    setSessionOrderIds(nextIds);
    setCurrentIndex(nextIndex);
    setFlipped(false);
    setRevealLevel(0);
    setRetrievalInput("");

    restoredFlashcardStateRef.current = null;
    shouldRebuildDeckRef.current = false;
    return;
  }

  if (!shouldRebuildDeckRef.current && sessionOrderIds.length) {
    const validIdSet = new Set(words.map((w) => w.id));
    const keptIds = sessionOrderIds.filter((id) => validIdSet.has(id));

    if (keptIds.length) {
      if (keptIds.length !== sessionOrderIds.length) {
        setSessionOrderIds(keptIds);
      }
      setCurrentIndex((prev) => Math.min(Math.max(prev, 0), keptIds.length - 1));
      return;
    }
  }

  const nextIds = orderMode === "ordered" ? ids : shuffleArray(ids);
  setSessionOrderIds(nextIds);
  setCurrentIndex(0);
  setFlipped(false);
  setRevealLevel(0);
  setRetrievalInput("");

  shouldRebuildDeckRef.current = false;
}, [words.length, orderMode, shuffleSeed, mode, flashcardFilter, search, storageReady]);

  const sessionOrder = useMemo(() => {
    if (!sessionOrderIds.length) return [];
    const wordMap = new Map(words.map((w) => [w.id, w]));
    return sessionOrderIds.map((id) => wordMap.get(id)).filter(Boolean);
  }, [sessionOrderIds, words]);

  const currentWord = sessionOrder[currentIndex] || null;
  const displayedSenses = useMemo(() => {
    if (!currentWord) return [];
    return mode === "favorite_senses" ? currentWord.senses.filter((s) => s.favorite) : currentWord.senses;
  }, [currentWord, mode]);
  const promptSense = displayedSenses[0] || currentWord?.senses?.[0] || null;
  const revealSteps = ["中文词义", "英义 + 核心近义词", "反义词 + 情感色彩", "相关词 + 例句"];
  const nextRevealLabel = !flipped ? revealSteps[0] : revealLevel >= 3 ? "已全部展开" : revealSteps[revealLevel + 1];
  const progress = sessionOrder.length ? ((currentIndex + 1) / sessionOrder.length) * 100 : 0;
  const progressLabel = sessionOrder.length ? `${progress < 1 ? progress.toFixed(1) : Math.round(progress)}%` : "0%";

  function updateDailyPlan(key, value) {
    const numeric = Number(value);
    setDailyPlan((prev) => ({ ...prev, [key]: Number.isFinite(numeric) ? Math.max(0, Math.floor(numeric)) : 0 }));
  }

  function recordDailyProgress(delta) {
    const key = formatDate(Date.now());
    setDailyStats((prev) => {
      const current = prev[key] || {};
      return {
        ...prev,
        [key]: {
          ...current,
          reviewed: (current.reviewed || 0) + (delta.reviewed || 0),
          newReviewed: (current.newReviewed || 0) + (delta.newReviewed || 0),
          reviewReviewed: (current.reviewReviewed || 0) + (delta.reviewReviewed || 0),
          known: (current.known || 0) + (delta.known || 0),
          unknown: (current.unknown || 0) + (delta.unknown || 0),
          quizCorrect: (current.quizCorrect || 0) + (delta.quizCorrect || 0),
          quizWrong: (current.quizWrong || 0) + (delta.quizWrong || 0),
          bbPairCorrect: (current.bbPairCorrect || 0) + (delta.bbPairCorrect || 0),
          bbPairWrong: (current.bbPairWrong || 0) + (delta.bbPairWrong || 0),
          updatedAt: Date.now(),
        },
      };
    });
  }

  function makeFlashcardSessionId() {
  return `fc_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`;
}

function flushFlashcardSession(endReason = "manual") {
  const session = flashcardSessionRef.current;

  if (!session || session.flushed || session.wordsSeen === 0) return;

  session.flushed = true;

  track("flashcard_session_summary", {
    session_id: session.id,
    mode: session.mode,
    filter: session.filter,
    flashcard_mode: session.flashcardMode,
    session_size: session.sessionSize,
    words_seen: session.wordsSeen,
    unique_words_seen: session.uniqueWordIds.size,
    known_count: session.knownCount,
    unknown_count: session.unknownCount,
    duration_sec: Math.round((Date.now() - session.startedAt) / 1000),
    end_reason: endReason,
  });

  flashcardSessionRef.current = null;
}

function startFlashcardAnalyticsSession(args: {
  mode: string;
  filter: string;
  flashcardMode: string;
  sessionSize: number;
}) {
  flushFlashcardSession("new_session_started");

  flashcardSessionRef.current = {
    id: makeFlashcardSessionId(),
    mode: args.mode,
    filter: args.filter,
    flashcardMode: args.flashcardMode,
    sessionSize: args.sessionSize,
    startedAt: Date.now(),
    wordsSeen: 0,
    knownCount: 0,
    unknownCount: 0,
    uniqueWordIds: new Set(),
    flushed: false,
  };

  track("start_flashcard_session", {
    mode: args.mode,
    filter: args.filter,
    flashcard_mode: args.flashcardMode,
    session_size: args.sessionSize,
  });
}

function recordFlashcardResult(
  word: { id?: string; word?: string } | null | undefined,
  result: "known" | "unknown"
) {
  const session = flashcardSessionRef.current;
  if (!session) return;

  session.wordsSeen += 1;

  if (word?.id) {
    session.uniqueWordIds.add(word.id);
  } else if (word?.word) {
    session.uniqueWordIds.add(word.word);
  }

  if (result === "known") {
    session.knownCount += 1;
  } else {
    session.unknownCount += 1;
  }

  if (session.wordsSeen >= session.sessionSize) {
    flushFlashcardSession("completed");
  }
}

  function openFlashcardDeck(nextMode = "all", options = {}) {
    shouldRebuildDeckRef.current = true;
    setSessionOrderIds([]);
    setFlashcardCompleteMessage("");

    startFlashcardAnalyticsSession({
      mode: nextMode,
      filter: options.filter || "all",
      flashcardMode: options.flashcardMode || "recognition",
      sessionSize: filteredWords.length,
    });
    setStudyView("flashcards");
    setMode(nextMode);
    setFlashcardFilter(options.filter || "all");
    setFlashcardMode(options.flashcardMode || "recognition");
    setCurrentIndex(0);
    setFlashcardCompleteMessage("");
    setSessionOrderIds([]);
    setFlipped(false);
    setRevealLevel(0);
    setRetrievalInput("");
  }

  function goPrev() {
    if (!sessionOrder.length) return;
    setFlashcardCompleteMessage("");
    setCurrentIndex((prev) => (prev - 1 < 0 ? sessionOrder.length - 1 : prev - 1));
    setFlipped(false);
    setRevealLevel(0);
    setRetrievalInput("");
  }

  function getFlashcardCompleteMessage() {
    if (flashcardFilter === "task") return "太棒啦，今日总任务全部完成咯！";
    if (flashcardFilter === "review") return "太棒啦，今日复习结束哦！";
    if (flashcardFilter === "today_new" || flashcardFilter === "new") return "太棒啦，今日新词任务完成咯！";
    return "太棒啦，这组闪卡完成咯！";
  }

  function goNext() {
    if (!sessionOrder.length) return;

    if (currentIndex + 1 >= sessionOrder.length) {
      setFlashcardCompleteMessage(getFlashcardCompleteMessage());
      flushFlashcardSession("completed");
      setFlipped(false);
      setRevealLevel(0);
      setRetrievalInput("");
      return;
    }

    setCurrentIndex((prev) => prev + 1);
    setFlipped(false);
    setRevealLevel(0);
    setRetrievalInput("");
  }

  function updateWordResult(type) {
  if (!currentWord) return;

  recordFlashcardResult(currentWord, type);

    const now = Date.now();
    const todayStart = getStartOfLocalDay(now);
    const isReviewCard = Boolean(
      (currentWord.reviewState?.pending ||
        (currentWord.stats?.unknown || 0) > 0 ||
        (currentWord.stats?.quizWrong || 0) > 0) &&
      (currentWord.reviewState?.lastWrongAt || 0) > 0 &&
      (currentWord.reviewState?.lastWrongAt || 0) < todayStart
    );

    setWords((prev) => prev.map((w) => {
      if (w.id !== currentWord.id) return w;
      const seen = w.stats.seen + 1;
      const known = w.stats.known + (type === "known" ? 1 : 0);
      const unknown = w.stats.unknown + (type === "unknown" ? 1 : 0);
      let nextInterval = w.srs.interval || 1;
      let streak = w.srs.streak || 0;
      if (type === "known") {
        streak += 1;
        nextInterval = Math.min(Math.round(nextInterval * 1.8 + 1), 30);
      } else {
        streak = 0;
        nextInterval = 1;
      }

      const prevReview = w.reviewState || { pending: false, correctStreak: 0, wrongStreak: 0, priority: 0, lastWrongAt: null, lastReviewedAt: null };
      let nextReview = { ...prevReview, lastReviewedAt: now };
      if (type === "unknown") {
        nextReview = {
          ...nextReview,
          pending: true,
          correctStreak: 0,
          wrongStreak: (prevReview.wrongStreak || 0) + 1,
          priority: Math.min((prevReview.priority || 0) + 2, 8),
          lastWrongAt: now,
        };
      } else if (prevReview.pending) {
        const correctStreak = (prevReview.correctStreak || 0) + 1;
        nextReview = {
          ...nextReview,
          pending: correctStreak < 2,
          correctStreak,
          wrongStreak: 0,
          priority: Math.max((prevReview.priority || 0) - 2, 0),
        };
      } else {
        nextReview = { ...nextReview, correctStreak: 0, wrongStreak: 0, priority: Math.max((prevReview.priority || 0) - 1, 0) };
      }

      return {
        ...w,
        stats: { ...w.stats, seen, known, unknown },
        reviewState: nextReview,
        srs: { interval: nextInterval, streak, due: now + nextInterval * 24 * 60 * 60 * 1000 },
      };
    }));

    const delta = {
      reviewed: 1,
      newReviewed: isReviewCard ? 0 : 1,
      reviewReviewed: isReviewCard ? 1 : 0,
      known: type === "known" ? 1 : 0,
      unknown: type === "unknown" ? 1 : 0,
    };
    setSessionStats((prev) => ({ ...prev, reviewed: prev.reviewed + 1, known: prev.known + delta.known, unknown: prev.unknown + delta.unknown }));
    recordDailyProgress(delta);

    if (currentIndex + 1 >= sessionOrder.length) {
      setFlashcardCompleteMessage(getFlashcardCompleteMessage());
      flushFlashcardSession("completed");
      setFlipped(false);
      setRevealLevel(0);
      setRetrievalInput("");
      return;
    }

    goNext();
  }

  function toggleFavoriteWord(wordId) {
    setWords((prev) => prev.map((w) => (w.id === wordId ? { ...w, favorite: !w.favorite } : w)));
  }

  function toggleFavoriteSense(wordId, senseId) {
    setWords((prev) => prev.map((w) => (w.id !== wordId ? w : { ...w, senses: w.senses.map((s) => (s.id === senseId ? { ...s, favorite: !s.favorite } : s)) })));
  }

  function clearWrongWordRecord(wordId) {
    const target = wrongWordSummary.find((item) => item.id === wordId);
    track("clear_wrong_word_item", {
      word: target?.word || "",
      before_total_wrong: target?.totalWrong || 0,
    });
    setWords((prev) => prev.map((w) => {
      if (w.id !== wordId) return w;
      return {
        ...w,
        stats: {
          ...w.stats,
          unknown: 0,
          quizWrong: 0,
        },
        reviewState: {
          ...(w.reviewState || {}),
          pending: false,
          correctStreak: 0,
          wrongStreak: 0,
          priority: 0,
          lastWrongAt: null,
        },
      };
    }));
  }

  function clearAllWrongWords() {
    if (typeof window !== "undefined" && !window.confirm("确定清空当前单词错题表吗？这会移除这些单词的错题记录。")) return;
    track("clear_wrong_book_words_all", {
      count: wrongWordSummary.length,
    });
    setWords((prev) => prev.map((w) => {
      const hasWrong = (w.stats?.unknown || 0) > 0 || (w.stats?.quizWrong || 0) > 0;
      if (!hasWrong) return w;
      return {
        ...w,
        stats: {
          ...w.stats,
          unknown: 0,
          quizWrong: 0,
        },
        reviewState: {
          ...(w.reviewState || {}),
          pending: false,
          correctStreak: 0,
          wrongStreak: 0,
          priority: 0,
          lastWrongAt: null,
        },
      };
    }));
  }

  function clearWrongPairRecord(pairId) {
    const target = wrongPairSummary.find((item) => item.id === pairId);
    track("clear_wrong_pair_item", {
      pair_a: target?.a || "",
      pair_b: target?.b || "",
      before_wrong_count: target?.wrongCount || 0,
    });
    setSixChoicePairs((prev) => prev.map((pair) => {
      if (pair.id !== pairId) return pair;
      return {
        ...pair,
        stats: {
          ...pair.stats,
          wrong: 0,
          lastWrongAt: null,
        },
        reviewState: {
          ...(pair.reviewState || {}),
          pending: false,
          correctStreak: 0,
          wrongStreak: 0,
          priority: 0,
          lastWrongAt: null,
        },
      };
    }));
  }

  function clearAllWrongPairs() {
    if (typeof window !== "undefined" && !window.confirm("确定清空当前六选二错题表吗？这会移除这些词对的错题记录。")) return;
    track("clear_wrong_book_pairs_all", {
      count: wrongPairSummary.length,
    });
    setSixChoicePairs((prev) => prev.map((pair) => {
      const hasWrong = (pair.stats?.wrong || 0) > 0;
      if (!hasWrong) return pair;
      return {
        ...pair,
        stats: {
          ...pair.stats,
          wrong: 0,
          lastWrongAt: null,
        },
        reviewState: {
          ...(pair.reviewState || {}),
          pending: false,
          correctStreak: 0,
          wrongStreak: 0,
          priority: 0,
          lastWrongAt: null,
        },
      };
    }));
  }

  function checkQuizAnswer() {
    if (!quizQuestion || !selectedChoice || quizChecked) return;
    const isCorrect = selectedChoice === quizQuestion.correctAnswer;
    setQuizChecked(true);

    track(isCorrect ? "quiz_answer_correct" : "quiz_answer_wrong", {
      quiz_mode: quizMode,
      prompt_word: quizQuestion?.promptWord || "",
      selected_choice: selectedChoice,
      correct_answer: quizQuestion?.correctAnswer || "",
      frequency: quizQuestion?.frequency || "",
    });

    if (quizQuestion.questionType === "bb_pairs") {
      setSixChoicePairs((prev) => prev.map((pair) => {
        if (pair.id !== quizQuestion.pairId) return pair;
        const prevReview = pair.reviewState || { pending: false, correctStreak: 0, wrongStreak: 0, priority: 0, lastWrongAt: null, lastReviewedAt: null };
        let nextReview = { ...prevReview, lastReviewedAt: Date.now() };
        if (!isCorrect) {
          nextReview = {
            ...nextReview,
            pending: true,
            correctStreak: 0,
            wrongStreak: (prevReview.wrongStreak || 0) + 1,
            priority: Math.min((prevReview.priority || 0) + 2, 8),
            lastWrongAt: Date.now(),
          };
        } else if (prevReview.pending) {
          const correctStreak = (prevReview.correctStreak || 0) + 1;
          nextReview = { ...nextReview, pending: correctStreak < 2, correctStreak, wrongStreak: 0, priority: Math.max((prevReview.priority || 0) - 2, 0) };
        }
        return {
          ...pair,
          stats: {
            ...pair.stats,
            seen: (pair.stats?.seen || 0) + 1,
            correct: (pair.stats?.correct || 0) + (isCorrect ? 1 : 0),
            wrong: (pair.stats?.wrong || 0) + (isCorrect ? 0 : 1),
            lastCorrectAt: isCorrect ? Date.now() : pair.stats?.lastCorrectAt || null,
            lastWrongAt: isCorrect ? pair.stats?.lastWrongAt || null : Date.now(),
          },
          reviewState: nextReview,
        };
      }));
      const delta = { bbPairCorrect: isCorrect ? 1 : 0, bbPairWrong: isCorrect ? 0 : 1 };
      setSessionStats((prev) => ({ ...prev, bbPairCorrect: prev.bbPairCorrect + delta.bbPairCorrect, bbPairWrong: prev.bbPairWrong + delta.bbPairWrong }));
      recordDailyProgress(delta);
      return;
    }

    setWords((prev) => prev.map((w) => {
      if (w.id !== quizQuestion.targetWordId) return w;
      const prevReview = w.reviewState || { pending: false, correctStreak: 0, wrongStreak: 0, priority: 0, lastWrongAt: null, lastReviewedAt: null };
      let nextReview = { ...prevReview, lastReviewedAt: Date.now() };
      if (!isCorrect) {
        nextReview = {
          ...nextReview,
          pending: true,
          correctStreak: 0,
          wrongStreak: (prevReview.wrongStreak || 0) + 1,
          priority: Math.min((prevReview.priority || 0) + 2, 8),
          lastWrongAt: Date.now(),
        };
      } else if (prevReview.pending) {
        const correctStreak = (prevReview.correctStreak || 0) + 1;
        nextReview = { ...nextReview, pending: correctStreak < 2, correctStreak, wrongStreak: 0, priority: Math.max((prevReview.priority || 0) - 2, 0) };
      }
      return { ...w, stats: { ...w.stats, quizCorrect: (w.stats.quizCorrect || 0) + (isCorrect ? 1 : 0), quizWrong: (w.stats.quizWrong || 0) + (isCorrect ? 0 : 1) }, reviewState: nextReview };
    }));
    const delta = { quizCorrect: isCorrect ? 1 : 0, quizWrong: isCorrect ? 0 : 1 };
    setSessionStats((prev) => ({ ...prev, quizCorrect: prev.quizCorrect + delta.quizCorrect, quizWrong: prev.quizWrong + delta.quizWrong }));
    recordDailyProgress(delta);
  }

  function nextQuizQuestion() {
    track("next_quiz_question", {
      quiz_mode: quizMode,
    });
    const q = buildQuizQuestion(filteredWords.length ? filteredWords : words, quizMode, filteredPairs.length ? filteredPairs : sixChoicePairs);
    setQuizQuestion(q);
    setSelectedChoice("");
    setQuizChecked(false);
  }

  async function handleImport(event) {
    const files = Array.from(event.target.files || []);
    if (!files.length) return;

    const processFile = (file) => new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = async (e) => {
        try {
          let imported = [];
          let importedPairs = [];
          const result = e.target?.result;

          if (file.name.endsWith(".json")) {
            const parsed = JSON.parse(String(result || "{}"));
            if (Array.isArray(parsed)) {
              if (parsed.length && parsed[0] && (parsed[0].a || parsed[0].b || parsed[0].pair_zh)) importedPairs = parsed.map(normalizePair);
              else imported = parsed.map(normalizeWord).filter((item) => item.word);
            } else if (parsed && typeof parsed === "object") {
              if (Array.isArray(parsed.words)) imported = parsed.words.map(normalizeWord).filter((item) => item.word);
              if (Array.isArray(parsed.sixChoicePairs)) importedPairs = parsed.sixChoicePairs.map(normalizePair);
            }
          } else if (file.name.endsWith(".csv")) {
            imported = parseCSV(String(result || ""));
          } else if (file.name.endsWith(".xlsx") || file.name.endsWith(".xls")) {
            const workbook = XLSX.read(result, { type: "array" });
            const preferredName = workbook.SheetNames.find((name) => name.includes("GRE")) || workbook.SheetNames[0];
            const rows = XLSX.utils.sheet_to_json(workbook.Sheets[preferredName], { defval: "" });
            imported = rows[0]?.["单词"] || rows.some((r) => r["考法名称"]) ? groupRowsToWords(rows) : rows.map(normalizeWord).filter((item) => item.word);
          } else if (file.name.endsWith(".pdf")) {
            importedPairs = await extractPairsFromPdf(result, file.name);
          } else {
            reject(new Error("Unsupported file format"));
            return;
          }

          resolve({ imported, importedPairs });
        } catch (error) {
          reject(error);
        }
      };

      if (file.name.endsWith(".xlsx") || file.name.endsWith(".xls") || file.name.endsWith(".pdf")) reader.readAsArrayBuffer(file);
      else reader.readAsText(file, "utf-8");
    });

    try {
      const results = await Promise.all(files.map(processFile));
      const mergedWords = results.flatMap((r) => r.imported || []);
      const mergedPairs = results.flatMap((r) => r.importedPairs || []);

      if (mergedWords.length) {
        track("upload_word_list", {
          file_type: "mixed",
          import_mode: importWordMode,
          word_count: mergedWords.length,
        });
        setWords((prev) => mergeWords(prev, mergedWords, importWordMode));
        openFlashcardDeck("all", { onlyDue: false });
      }

      if (mergedPairs.length) {
        track("upload_pairs_pdf", {
          file_type: "pdf_or_mixed",
          pair_import_mode: pairImportMode,
          pair_count: mergedPairs.length,
        });
        if (pairImportMode === "rebuild") setSixChoicePairs(dedupePairs(mergedPairs));
        else setSixChoicePairs((prev) => dedupePairs([...prev, ...mergedPairs]));
      }

      if (!mergedWords.length && !mergedPairs.length) {
        alert("No valid words or six-choice pairs found in the uploaded files.");
      }
    } catch (error) {
      track("import_failed", {
        error_stage: "handle_import",
      });
      console.error(error);
      alert("Import failed. Please check your file format.");
    }
  }

  function exportData() {
    track("export_progress", {
      word_count: words.length,
      pair_count: sixChoicePairs.length,
      wrong_word_count: wrongWordBook.length,
      wrong_pair_count: wrongPairBook.length,
    });
    try {
      const blob = new Blob([JSON.stringify({ words, sixChoicePairs, sessionStats, dailyStats, taskGoals, dailyPlan }, null, 2)], { type: "application/json;charset=utf-8" });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = `gre_trainer_progress_${formatDate(Date.now())}.json`;
      a.style.display = "none";
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      setStorageMessage("学习数据已开始导出。如果浏览器拦截下载，请检查下载栏或浏览器权限。");
      setTimeout(() => URL.revokeObjectURL(url), 1000);
    } catch (error) {
      console.error("Export failed", error);
      setStorageMessage("导出失败。当前环境可能拦截了下载，你可以先用浏览器打开完整页面再导出。");
    }
  }

  function resetProgress() {
    setWords((prev) => prev.map((w) => ({
      ...w,
      stats: { seen: 0, known: 0, unknown: 0, quizCorrect: 0, quizWrong: 0 },
      reviewState: { pending: false, correctStreak: 0, wrongStreak: 0, priority: 0, lastWrongAt: null, lastReviewedAt: null },
      srs: { interval: 1, due: Date.now(), streak: 0 },
    })));
    setSixChoicePairs((prev) => prev.map((p) => ({
      ...p,
      stats: { ...p.stats, seen: 0, correct: 0, wrong: 0, lastWrongAt: null, lastCorrectAt: null },
      reviewState: { pending: false, correctStreak: 0, wrongStreak: 0, priority: 0, lastWrongAt: null, lastReviewedAt: null },
    })));
    setDailyStats({});
    setSessionStats(DEFAULT_SESSION_STATS);
    setCurrentIndex(0);
    setFlashcardCompleteMessage("");
    setSessionOrderIds([]);
    setFlipped(false);
    setRevealLevel(0);
  }

  function resetBbPairLibrary() {
    setSixChoicePairs(sampleSixChoicePairs.map(normalizePair));
    setPairReviewMode("mix");
    setPairSearch("");
    if (quizMode === "bb_pairs") nextQuizQuestion();
  }

  function shuffleSession() {
    shouldRebuildDeckRef.current = true;
    setFlashcardCompleteMessage("");
    setShuffleSeed((prev) => prev + 1);
  }

  const dueCount = words.filter((w) => (w.srs?.due || 0) <= Date.now()).length;
  const newWordCount = words.filter((w) => w.stats.seen === 0 && (w.stats.quizCorrect || 0) === 0 && (w.stats.quizWrong || 0) === 0).length;
  const pairNewCount = sixChoicePairs.filter((pair) => (pair.stats?.seen || 0) === 0).length;
  const dailyWordNewTarget = Math.ceil((newWordCount || words.length || 0) / Math.max(1, dailyPlan.wordFinishDays || 1));
  const dailyPairNewTarget = Math.ceil((pairNewCount || sixChoicePairs.length || 0) / Math.max(1, dailyPlan.pairFinishDays || 1));
  const todayKey = formatDate(Date.now());
  const todayStat = dailyStats[todayKey] || {};
  const todayNewWordStudyCount = todayStat.newReviewed || 0;
  const todayReviewWordStudyCount = todayStat.reviewReviewed || 0;
  const todayWordStudyCount = todayStat.reviewed || 0;
  const todayPairPracticeCount = (todayStat.bbPairCorrect || 0) + (todayStat.bbPairWrong || 0);
  const todayQuizCount = (todayStat.quizCorrect || 0) + (todayStat.quizWrong || 0);
  const todayQuizAccuracy = todayQuizCount ? Math.round(((todayStat.quizCorrect || 0) / todayQuizCount) * 100) : null;
  const todayNewWordTarget = dailyWordNewTarget;
  const todayReviewWordTarget = todayReviewTargetActual;
  const todayWordTarget = todayNewWordTarget + todayReviewWordTarget;
  const todayPairTarget = dailyPairNewTarget + dailyPlan.pairReviewCount;
  const todayNewWordRemaining = Math.max(0, todayNewWordTarget - todayNewWordStudyCount);
  const todayReviewWordRemaining = Math.max(0, todayReviewWordTarget - todayReviewWordStudyCount);
  const todayWordRemaining = todayNewWordRemaining + todayReviewWordRemaining;
  const todayPairRemaining = Math.max(0, todayPairTarget - todayPairPracticeCount);
  const todayNewWordProgressPct = todayNewWordTarget ? Math.min(100, Math.round((todayNewWordStudyCount / todayNewWordTarget) * 100)) : 100;
  const todayReviewWordProgressPct = todayReviewWordTarget ? Math.min(100, Math.round((todayReviewWordStudyCount / todayReviewWordTarget) * 100)) : 100;
  const todayWordProgressPct = todayWordTarget ? Math.min(100, Math.round((todayWordStudyCount / todayWordTarget) * 100)) : 100;
  const todayPairProgressPct = todayPairTarget ? Math.min(100, Math.round((todayPairPracticeCount / todayPairTarget) * 100)) : 0;
  const todayQuizProgressPct = todayQuizAccuracy === null ? 0 : Math.min(100, Math.round((todayQuizAccuracy / Math.max(1, dailyPlan.quizTarget)) * 100));
  const wrongCount = wrongWordBook.length;
  const newWordPlanStatus = todayNewWordRemaining === 0 ? "done" : todayNewWordProgressPct >= 70 ? "warning" : "pending";
  const reviewWordPlanStatus = todayReviewWordRemaining === 0 ? "done" : todayReviewWordProgressPct >= 70 ? "warning" : "pending";
  const wordPlanStatus = todayWordRemaining === 0 ? "done" : todayWordProgressPct >= 70 ? "warning" : "pending";
  const pairPlanStatus = todayPairRemaining === 0 ? "done" : todayPairProgressPct >= 70 ? "warning" : "pending";
  const quizPlanStatus = todayQuizAccuracy === null ? "pending" : todayQuizAccuracy >= dailyPlan.quizTarget ? "done" : todayQuizAccuracy >= Math.max(0, dailyPlan.quizTarget - 10) ? "warning" : "pending";
  const overallPlanStatus = [wordPlanStatus, pairPlanStatus, quizPlanStatus].every((status) => status === "done")
    ? "done"
    : [wordPlanStatus, pairPlanStatus, quizPlanStatus].some((status) => status === "warning" || status === "done")
      ? "warning"
      : "pending";
  const overallPlanMeta = getPlanStatusMeta(overallPlanStatus);
  const wordPlanMeta = getPlanStatusMeta(wordPlanStatus);
  const newWordPlanMeta = getPlanStatusMeta(newWordPlanStatus);
  const reviewWordPlanMeta = getPlanStatusMeta(reviewWordPlanStatus);
  const pairPlanMeta = getPlanStatusMeta(pairPlanStatus);
  const quizPlanMeta = getPlanStatusMeta(quizPlanStatus);
  
  return (
    <div className="min-h-screen bg-slate-50 p-4 md:p-8">
      <div className="mx-auto max-w-7xl space-y-6">
        {storageMessage ? (
          <Card className="rounded-2xl border-amber-300 bg-amber-50 shadow-sm">
            <CardContent className="p-5 text-sm text-amber-900 whitespace-pre-line">{storageMessage}</CardContent>
          </Card>
        ) : null}


        <div className="grid gap-4 md:grid-cols-5">
          <Card className="rounded-2xl shadow-sm"><CardContent className="p-5"><div className="text-sm text-slate-500">Total Words</div><div className="mt-2 text-3xl font-semibold">{words.length}</div></CardContent></Card>
          <Card className="rounded-2xl shadow-sm"><CardContent className="p-5"><div className="text-sm text-slate-500">Due Today</div><div className="mt-2 text-3xl font-semibold">{dueCount}</div></CardContent></Card>
          <Card className="rounded-2xl shadow-sm"><CardContent className="p-5"><div className="text-sm text-slate-500">Wrong-Words</div><div className="mt-2 text-3xl font-semibold">{wrongCount}</div></CardContent></Card>
          <Card className="rounded-2xl shadow-sm"><CardContent className="p-5"><div className="text-sm text-slate-500">Favorites</div><div className="mt-2 text-3xl font-semibold">{favoriteWords.length + favoriteSenseCount}</div></CardContent></Card>
          <Card className="rounded-2xl shadow-sm"><CardContent className="p-5"><div className="text-sm text-slate-500">BB Pairs</div><div className="mt-2 text-3xl font-semibold">{sixChoicePairs.length}</div></CardContent></Card>
        </div>

        <div className="grid gap-6 lg:grid-cols-[320px_1fr]">
          <Card className="order-2 rounded-2xl shadow-sm lg:order-1">
            <CardHeader>
              <CardTitle className="flex items-center gap-2 text-lg"><BookOpen className="h-5 w-5" /> GRE Trainer</CardTitle>
            </CardHeader>
            <CardContent className="space-y-4">
              <input ref={fileRef} type="file" accept=".json,.csv,.xlsx,.xls,.pdf" multiple className="hidden" onChange={handleImport} />

              <div className="space-y-2">
                <div className="text-sm font-medium">Training View</div>
                <Tabs value={studyView} onValueChange={setStudyView} className="w-full">
                  <TabsList className="grid w-full grid-cols-2 gap-2 h-auto bg-transparent p-0">
                    <TabsTrigger value="flashcards" className="rounded-xl border px-3 py-2">Flashcards</TabsTrigger>
                    <TabsTrigger value="quiz" className="rounded-xl border px-3 py-2">GRE Quiz</TabsTrigger>
                  </TabsList>
                </Tabs>
              </div>

              {studyView === "flashcards" ? (
                <div className="space-y-3">
                  <div>
                    <div className="text-sm font-medium">闪卡模式</div>
                    <div className="mt-2 flex flex-wrap gap-2">
                      <Button variant={flashcardMode === "recognition" ? "default" : "outline"} size="sm" className="rounded-full" onClick={() => setFlashcardMode("recognition")}>识别模式</Button>
                      <Button variant={flashcardMode === "retrieval" ? "default" : "outline"} size="sm" className="rounded-full" onClick={() => setFlashcardMode("retrieval")}>提取模式</Button>
                    </div>
                  </div>

                  <div>
                    <div className="text-sm font-medium">学习范围</div>
                    <div className="mt-2 flex flex-wrap gap-2">
                      <Button variant={flashcardFilter === "task" ? "default" : "outline"} size="sm" className="rounded-full" onClick={() => openFlashcardDeck("all", { filter: "task", flashcardMode })}>今日总任务</Button>
                      <Button variant={flashcardFilter === "review" ? "default" : "outline"} size="sm" className="rounded-full" onClick={() => openFlashcardDeck("all", { filter: "review", flashcardMode: "recognition" })}>只看今日复习</Button>
                      <Button variant={flashcardFilter === "today_new" ? "default" : "outline"} size="sm" className="rounded-full" onClick={() => openFlashcardDeck("all", { filter: "today_new", flashcardMode })}>只看今日新词</Button>
                      <Button variant={flashcardFilter === "new" ? "default" : "outline"} size="sm" className="rounded-full" onClick={() => openFlashcardDeck("all", { filter: "new", flashcardMode })}>继续学习新词</Button>
                    </div>
                  </div>

                  <div>
                    <div className="text-sm font-medium">排列方式</div>
                    <div className="mt-2 flex flex-wrap gap-2">
                      <Button variant={orderMode === "ordered" ? "default" : "outline"} size="sm" className="rounded-full" onClick={() => setOrderMode("ordered")}>正序</Button>
                      <Button variant={orderMode === "shuffled" ? "default" : "outline"} size="sm" className="rounded-full" onClick={() => setOrderMode("shuffled")}>乱序</Button>
                      <Button variant="outline" size="sm" className="rounded-full" disabled={orderMode !== "shuffled"} onClick={shuffleSession}><Shuffle className="mr-2 h-4 w-4" /> 重新洗牌</Button>
                      <Button variant="outline" size="sm" className="rounded-full" onClick={resetProgress}><RotateCcw className="mr-2 h-4 w-4" /> 重置进度</Button>
                    </div>
                  </div>
                </div>
              ) : (
                <div className="space-y-3">
                  <div className="space-y-2">
                    <div className="text-sm font-medium">Quiz Type</div>
                    <Tabs value={quizMode} onValueChange={setQuizMode} className="w-full">
                      <TabsList className="grid w-full grid-cols-3 gap-2 h-auto bg-transparent p-0">
                        <TabsTrigger value="equivalence" className="rounded-xl border px-3 py-2">等价词</TabsTrigger>
                        <TabsTrigger value="antonym" className="rounded-xl border px-3 py-2">反义关系</TabsTrigger>
                        <TabsTrigger value="bb_pairs" className="rounded-xl border px-3 py-2">六选二</TabsTrigger>
                      </TabsList>
                    </Tabs>
                  </div>
                  <div className="space-y-2">
                    <div className="text-sm font-medium">Quiz Hint Mode</div>
                    <div className="flex flex-wrap gap-2">
                      <Button variant={quizHintMode === "study" ? "default" : "outline"} size="sm" className="rounded-full" onClick={() => setQuizHintMode("study")}>学习模式</Button>
                      <Button variant={quizHintMode === "exam" ? "default" : "outline"} size="sm" className="rounded-full" onClick={() => setQuizHintMode("exam")}>考试模式</Button>
                    </div>
                  </div>
                </div>
              )}

              {studyView === "quiz" && quizMode === "bb_pairs" ? (
                <>
                  <div className="space-y-2">
                    <div className="text-sm font-medium">六选二题库</div>
                    <div className="flex flex-wrap gap-2">
                      <Button variant={pairReviewMode === "mix" ? "default" : "outline"} size="sm" className="rounded-full" onClick={() => setPairReviewMode("mix")}>高频混刷</Button>
                      <Button variant={pairReviewMode === "sss" ? "default" : "outline"} size="sm" className="rounded-full" onClick={() => setPairReviewMode("sss")}>超超超高频</Button>
                      <Button variant={pairReviewMode === "ss" ? "default" : "outline"} size="sm" className="rounded-full" onClick={() => setPairReviewMode("ss")}>超超高频</Button>
                      <Button variant={pairReviewMode === "s" ? "default" : "outline"} size="sm" className="rounded-full" onClick={() => setPairReviewMode("s")}>超高频</Button>
                      <Button variant={pairReviewMode === "wrong" ? "default" : "outline"} size="sm" className="rounded-full" onClick={() => setPairReviewMode("wrong")}>六选二错题本</Button>
                      <Button variant={pairReviewMode === "all" ? "default" : "outline"} size="sm" className="rounded-full" onClick={() => setPairReviewMode("all")}>全部题库</Button>
                    </div>
                  </div>

                  <div className="space-y-2">
                    <div className="text-sm font-medium">六选二库检索</div>
                    <div className="relative">
                      <Search className="absolute left-3 top-1/2 h-4 w-4 -translate-y-1/2 text-slate-400" />
                      <Input value={pairSearch} onChange={(e) => setPairSearch(e.target.value)} placeholder="Search pair, 中文解释, source..." className="rounded-2xl pl-9" />
                    </div>
                    <div className="text-xs text-slate-500">Current BB pair bank: {filteredPairs.length} / {sixChoicePairs.length}</div>
                    <div className="text-xs text-slate-500">频级统计：超超超高频 {pairFrequencyStats.sss} · 超超高频 {pairFrequencyStats.ss} · 超高频 {pairFrequencyStats.s}</div>
                  </div>
                </>
              ) : null}

              <div className="space-y-2">
                <div className="text-sm font-medium">Search</div>
                <div className="relative">
                  <Search className="absolute left-3 top-1/2 h-4 w-4 -translate-y-1/2 text-slate-400" />
                  <Input value={search} onChange={(e) => setSearch(e.target.value)} placeholder="Search word, 中文义, EN meaning, synonym..." className="rounded-2xl pl-9" />
                </div>
              </div>

              {studyView === "flashcards" ? (
                <div>
                  <div className="text-sm font-medium">词库分组</div>
                  <div className="mt-2 flex flex-wrap gap-2">
                    <Button variant={mode === "all" && flashcardFilter === "all" ? "default" : "outline"} size="sm" className="rounded-full" onClick={() => openFlashcardDeck("all", { filter: "all" })}>全部词库</Button>
                    <Button variant={mode === "wrong" && flashcardFilter === "all" ? "default" : "outline"} size="sm" className="rounded-full" onClick={() => {
                          track("open_wrong_book_words", { wrong_count: wrongWordSummary.length });
                          openFlashcardDeck("wrong", { filter: "all", flashcardMode: "recognition" });
                        }}>单词错题本</Button>
                    <Button variant={mode === "stubborn" && flashcardFilter === "all" ? "default" : "outline"} size="sm" className="rounded-full" onClick={() => openFlashcardDeck("stubborn", { filter: "all", flashcardMode: "retrieval" })}>顽固错词本</Button>
                    <Button variant={mode === "favorite_words" && flashcardFilter === "all" ? "default" : "outline"} size="sm" className="rounded-full" onClick={() => openFlashcardDeck("favorite_words", { filter: "all" })}>收藏单词本</Button>
                    <Button variant={mode === "favorite_senses" && flashcardFilter === "all" ? "default" : "outline"} size="sm" className="rounded-full" onClick={() => openFlashcardDeck("favorite_senses", { filter: "all" })}>收藏词义本</Button>
                  </div>
                </div>
              ) : null}

              <Button className="w-full rounded-2xl" onClick={() => fileRef.current?.click()}><Upload className="mr-2 h-4 w-4" /> Import Word List / PDF</Button>
              <Button variant="outline" className="w-full rounded-2xl" onClick={exportData}><Download className="mr-2 h-4 w-4" /> Export Progress</Button>

              <div className="space-y-2 pt-2 border-t">
                <div className="text-sm font-medium">Word Import Mode</div>
                <div className="flex flex-wrap gap-2">
                  <Button variant={importWordMode === "append" ? "default" : "outline"} size="sm" className="rounded-full" onClick={() => setImportWordMode("append")}>追加导入</Button>
                  <Button variant={importWordMode === "replace" ? "default" : "outline"} size="sm" className="rounded-full" onClick={() => setImportWordMode("replace")}>替换单词库</Button>
                </div>
                <div className="text-xs text-slate-500">{importWordMode === "append" ? "新导入的单词会追加进现有词库；重名单词会合并，并尽量保留你已有的学习记录。" : "当前导入会替换整套单词库。"}</div>
              </div>

              <div className="space-y-2">
                <div className="text-sm font-medium">BB Pair Import Mode</div>
                <div className="flex flex-wrap gap-2">
                  <Button variant={pairImportMode === "append" ? "default" : "outline"} size="sm" className="rounded-full" onClick={() => setPairImportMode("append")}>追加到现有题库</Button>
                  <Button variant={pairImportMode === "rebuild" ? "default" : "outline"} size="sm" className="rounded-full" onClick={() => setPairImportMode("rebuild")}>从零重建题库</Button>
                </div>
                <div className="text-xs text-slate-500">{pairImportMode === "append" ? "新导入的词对会和现有六选二库合并，并按更高频级保留。" : "重新导入 BB 源文件时，将只使用这次导入的数据从零建库。"}</div>
                <Button variant="outline" size="sm" className="rounded-2xl w-full" onClick={resetBbPairLibrary}>清空当前 BB 题库到默认样例</Button>
              </div>
            </CardContent>
          </Card>

          <div className="order-1 space-y-6 lg:order-2">
            {studyView === "flashcards" ? (
              <>
                <Card className="rounded-2xl shadow-sm">
                  <CardContent className="p-5">
                    <div className="mb-3 flex items-center justify-between gap-3 text-sm text-slate-600">
                      <span>{sessionOrder.length ? `Card ${Math.min(currentIndex + 1, sessionOrder.length)} / ${sessionOrder.length}` : "No cards found"}</span>
                      <div className="flex items-center gap-2">
                        {(mode !== "all" || flashcardFilter !== "all") ? <Button variant="outline" size="sm" className="rounded-full" onClick={() => openFlashcardDeck("all", { filter: "all" })}>返回全部闪卡</Button> : null}
                        <span>{progressLabel}</span>
                      </div>
                    </div>
                    <Progress value={progress} className="h-2" />
                  </CardContent>
                </Card>

                {flashcardCompleteMessage ? (
                  <Card className="rounded-[28px] border-violet-100 bg-white/90 shadow-sm">
                    <CardContent className="flex min-h-[560px] flex-col items-center justify-center gap-5 p-8 text-center">
                      <div className="rounded-full bg-violet-50 px-5 py-2 text-sm font-medium text-violet-700">Completed</div>
                      <h2 className="text-3xl font-semibold tracking-tight text-slate-900 md:text-4xl">{flashcardCompleteMessage}</h2>
                      <p className="max-w-md text-sm leading-7 text-slate-500">可以休息一下，或者回到今日计划继续其他任务。</p>
                      <div className="mt-2 flex flex-wrap justify-center gap-3">
                        <Button className="rounded-2xl" onClick={() => openFlashcardDeck("all", { filter: "task", flashcardMode })}>查看今日总任务</Button>
                        <Button variant="outline" className="rounded-2xl" onClick={() => openFlashcardDeck("all", { filter: "today_new", flashcardMode })}>继续今日新词</Button>
                        <Button variant="outline" className="rounded-2xl" onClick={() => openFlashcardDeck("all", { filter: "review", flashcardMode: "recognition" })}>查看今日复习</Button>
                      </div>
                    </CardContent>
                  </Card>
                ) : currentWord ? (
                  <AnimatePresence mode="wait">
                    <motion.div key={`${currentWord.id}-${flipped}-${revealLevel}-${currentIndex}`} initial={{ opacity: 0, y: 12 }} animate={{ opacity: 1, y: 0 }} exit={{ opacity: 0, y: -12 }} transition={{ duration: 0.2 }}>
                      <Card className="min-h-[560px] cursor-pointer rounded-[28px] shadow-sm" onClick={() => { if (!flipped) setFlipped(true); else setRevealLevel((v) => Math.min(v + 1, 3)); }}>
                        <CardContent className="flex min-h-[560px] flex-col justify-between p-6 md:p-8">
                          {!flipped ? (
                            <div className="flex h-full min-h-[420px] flex-col items-center justify-center text-center">
                              {flashcardMode === "recognition" ? (
                                <>
                                  <div className="text-sm uppercase tracking-[0.2em] text-slate-400">Flashcard</div>
                                  <h1 className="mt-6 text-5xl font-semibold tracking-tight md:text-7xl">{currentWord.word}</h1>
                                </>
                              ) : (
                                <>
                                  <div className="text-sm uppercase tracking-[0.2em] text-slate-400">Retrieval</div>
                                  <div className="mt-6 w-full max-w-2xl rounded-[28px] border bg-white p-6 text-left shadow-sm">
                                    <div className="text-xs font-medium text-slate-400">根据词义回忆单词</div>
                                    <div className="mt-3 text-sm font-medium text-slate-500">中文词义</div>
                                    <div className="mt-1 text-2xl font-semibold leading-9 text-slate-900">{promptSense?.zh || currentWord.shortMeaning || "—"}</div>
                                    <div className="mt-4 text-sm font-medium text-slate-500">English meaning</div>
                                    <div className="mt-1 text-base leading-7 text-slate-700">{promptSense?.en || "—"}</div>
                                    <div className="mt-5 space-y-2">
                                      <div className="text-sm font-medium text-slate-500">把你想到的单词打出来</div>
                                      <Input value={retrievalInput} onChange={(e) => setRetrievalInput(e.target.value)} placeholder="type the word you recalled" className="rounded-2xl" onClick={(e) => e.stopPropagation()} />
                                    </div>
                                  </div>
                                </>
                              )}
                            </div>
                          ) : (
                            <div className="space-y-6">
                              <div className="flex flex-wrap items-center gap-3">
                                <h2 className="text-3xl font-semibold">{currentWord.word}</h2>
                                <Button variant="outline" size="sm" className="rounded-full" onClick={(e) => { e.stopPropagation(); toggleFavoriteWord(currentWord.id); }}>
                                  <Heart className={`mr-2 h-4 w-4 ${currentWord.favorite ? "fill-current text-rose-500" : "text-slate-500"}`} />
                                  {currentWord.favorite ? "已收藏单词" : "收藏单词"}
                                </Button>
                                {currentWord.corePos ? <Badge variant="outline" className="rounded-full">{currentWord.corePos}</Badge> : null}
                                <Badge variant="secondary" className="rounded-full">{sentimentLabel(currentWord.sentiment)}</Badge>
                                <Badge variant="secondary" className="rounded-full">Step {revealLevel + 1}/4 · {revealSteps[revealLevel]}</Badge>
                              </div>

                              {currentWord.memoryTip ? (
                                <div className="rounded-2xl bg-slate-100 p-4">
                                  <div className="mb-1 flex items-center gap-2 text-sm font-medium text-slate-500"><HeartHandshake className="h-4 w-4" /> 辅助记忆</div>
                                  <div className="text-base leading-7 text-slate-700">{currentWord.memoryTip}</div>
                                </div>
                              ) : null}

                              {displayedSenses.map((sense, idx) => (
                                <div key={`${currentWord.id}-${idx}`} className="rounded-2xl border p-4 space-y-4">
                                  <div className="flex flex-wrap items-center gap-2">
                                    <Badge className="rounded-full">{sense.label || `Sense ${idx + 1}`}</Badge>
                                    <Button variant="outline" size="sm" className="rounded-full" onClick={(e) => { e.stopPropagation(); toggleFavoriteSense(currentWord.id, sense.id); }}>
                                      <Heart className={`mr-2 h-4 w-4 ${sense.favorite ? "fill-current text-rose-500" : "text-slate-500"}`} />
                                      {sense.favorite ? "已收藏词义" : "收藏词义"}
                                    </Button>
                                  </div>
                                  <div className="grid gap-4 md:grid-cols-2">
                                    <div><div className="mb-1 text-sm font-medium text-slate-500">中文词义</div><div className="text-base leading-7 text-slate-800">{sense.zh || "—"}</div></div>
                                    {revealLevel >= 1 ? <div><div className="mb-1 text-sm font-medium text-slate-500">English meaning</div><div className="text-base leading-7 text-slate-800">{sense.en || "—"}</div></div> : null}
                                  </div>
                                  {revealLevel >= 1 ? <div><div className="mb-1 text-sm font-medium text-slate-500">核心近义词</div><div className="flex flex-wrap gap-2">{(sense.synonyms || []).length ? sense.synonyms.map((a) => <Badge key={a} variant="secondary" className="rounded-full">{a}</Badge>) : "—"}</div></div> : null}
                                  {revealLevel >= 2 ? <div><div className="mb-1 text-sm font-medium text-slate-500">反义词</div><div className="flex flex-wrap gap-2">{(sense.antonyms || []).length ? sense.antonyms.map((a) => <Badge key={a} variant="secondary" className="rounded-full">{a}</Badge>) : "—"}</div></div> : null}
                                  {revealLevel >= 3 ? <div className="grid gap-4 md:grid-cols-2"><div><div className="mb-1 text-sm font-medium text-slate-500">例句-中文</div><div className="text-sm leading-7 text-slate-700">{sense.exampleZh || "—"}</div></div><div><div className="mb-1 text-sm font-medium text-slate-500">Example</div><div className="text-sm leading-7 text-slate-700">{sense.exampleEn || "—"}</div></div></div> : null}
                                </div>
                              ))}
                            </div>
                          )}

                          <div className="mt-8 flex flex-wrap items-center justify-between gap-3">
                            <div className="flex gap-2">
                              <Button variant="outline" className="rounded-2xl" onClick={(e) => { e.stopPropagation(); goPrev(); }}>Previous</Button>
                              <Button variant="outline" className="rounded-2xl" onClick={(e) => { e.stopPropagation(); if (!flipped) setFlipped(true); else setRevealLevel((v) => Math.min(v + 1, 3)); }}><Eye className="mr-2 h-4 w-4" /> {nextRevealLabel}</Button>
                              <Button variant="outline" className="rounded-2xl" onClick={(e) => { e.stopPropagation(); goNext(); }}>Next</Button>
                            </div>
                            <div className="flex gap-2">
                              <Button variant="outline" className="rounded-2xl" onClick={(e) => { e.stopPropagation(); updateWordResult("unknown"); }}><XCircle className="mr-2 h-4 w-4" /> Don’t Know</Button>
                              <Button className="rounded-2xl" onClick={(e) => { e.stopPropagation(); updateWordResult("known"); }}><CheckCircle2 className="mr-2 h-4 w-4" /> Know</Button>
                            </div>
                          </div>
                        </CardContent>
                      </Card>
                    </motion.div>
                  </AnimatePresence>
                ) : (
                  <Card className="rounded-[28px] shadow-sm"><CardContent className="flex min-h-[460px] flex-col items-center justify-center p-8 text-center"><div className="rounded-full bg-slate-100 p-4"><Trash2 className="h-8 w-8 text-slate-500" /></div><h3 className="mt-4 text-2xl font-semibold">No cards in this filter</h3></CardContent></Card>
                )}
              </>
            ) : (
              <>
                <Card className="rounded-2xl shadow-sm">
                  <CardContent className="p-5 flex flex-wrap items-center justify-between gap-3">
                    <div>
                      <div className="text-sm text-slate-500">GRE Quiz Mode</div>
                      <div className="mt-1 text-lg font-semibold">{quizMode === "equivalence" ? "等价词训练" : quizMode === "antonym" ? "反义关系训练" : "BB 六选二训练"}</div>
                    </div>
                    <Badge variant="outline" className="rounded-full px-3 py-1">{quizMode === "bb_pairs" ? `${filteredPairs.length || sixChoicePairs.length} pairs in pool` : `${filteredWords.length || words.length} words in pool`}</Badge>
                  </CardContent>
                </Card>

                {quizQuestion ? (
                  <Card className="min-h-[560px] rounded-[28px] shadow-sm">
                    <CardContent className="p-6 md:p-8 space-y-6">
                      <div className="flex flex-wrap items-center gap-3">
                        <Badge className="rounded-full"><Target className="mr-2 h-4 w-4" /> {quizMode === "equivalence" ? "Text Completion / Equivalence" : quizMode === "antonym" ? "Opposite Logic" : "BB Six-Choice Pairing"}</Badge>
                        {quizHintMode === "study" && quizQuestion.frequency ? <Badge variant="outline" className="rounded-full">{quizQuestion.frequency}</Badge> : null}
                      </div>
                      <div className="space-y-3">
                        <div className="text-sm text-slate-500">Prompt word</div>
                        <div className="text-4xl font-semibold tracking-tight">{quizQuestion.promptWord}</div>
                        <div className="rounded-2xl bg-slate-100 p-4 text-slate-700 leading-7">
                          {quizHintMode === "study" ? <div><span className="font-medium">中文提示：</span>{quizQuestion.promptZh || "—"}</div> : null}
                          {quizHintMode === "study" && quizMode !== "bb_pairs" ? <div className="mt-2"><span className="font-medium">English hint：</span>{quizQuestion.promptEn || "—"}</div> : null}
                        </div>
                      </div>
                      <div className={`grid gap-3 ${quizMode === "bb_pairs" ? "md:grid-cols-3" : "md:grid-cols-2"}`}>
                        {quizQuestion.choices.map((choice) => {
                          const isSelected = selectedChoice === choice;
                          const isCorrect = quizChecked && choice === quizQuestion.correctAnswer;
                          const isWrongSelected = quizChecked && isSelected && choice !== quizQuestion.correctAnswer;
                          return (
                            <button key={choice} type="button" disabled={quizChecked} onClick={() => setSelectedChoice(choice)} className={`rounded-2xl border px-4 py-4 text-left transition ${isCorrect ? "border-emerald-500 bg-emerald-50" : isWrongSelected ? "border-rose-500 bg-rose-50" : isSelected ? "border-slate-900 bg-slate-100" : "border-slate-200 bg-white hover:bg-slate-50"}`}>
                              <div className="flex items-center gap-3"><PencilLine className="h-4 w-4 text-slate-500" /><span className="text-base font-medium">{choice}</span></div>
                            </button>
                          );
                        })}
                      </div>
                      {quizChecked ? <div className={`rounded-2xl p-4 ${selectedChoice === quizQuestion.correctAnswer ? "bg-emerald-50 text-emerald-800" : "bg-rose-50 text-rose-800"}`}><div className="font-medium">{selectedChoice === quizQuestion.correctAnswer ? "答对了" : "答错了"}</div><div className="mt-1">正确答案：{quizQuestion.correctAnswer}</div><div className="mt-1">{quizQuestion.explanation}</div></div> : null}
                      <div className="flex flex-wrap items-center justify-between gap-3 pt-4">
                        <Button variant="outline" className="rounded-2xl" onClick={nextQuizQuestion}>Skip</Button>
                        <div className="flex gap-2">
                          <Button variant="outline" className="rounded-2xl" disabled={!selectedChoice || quizChecked} onClick={checkQuizAnswer}>Check</Button>
                          <Button className="rounded-2xl" onClick={nextQuizQuestion}>Next Question</Button>
                        </div>
                      </div>
                    </CardContent>
                  </Card>
                ) : (
                  <Card className="rounded-[28px] shadow-sm"><CardContent className="flex min-h-[460px] flex-col items-center justify-center p-8 text-center"><div className="rounded-full bg-slate-100 p-4"><Trash2 className="h-8 w-8 text-slate-500" /></div><h3 className="mt-4 text-2xl font-semibold">Quiz pool is too small</h3></CardContent></Card>
                )}
              </>
            )}

            <Card className="rounded-2xl shadow-sm">
              <CardHeader className="space-y-4">
                <div className="flex flex-col gap-3 md:flex-row md:items-center md:justify-between">
                  <CardTitle className="flex items-center gap-2 text-lg">
                    <Brain className="h-5 w-5" />
                    错题汇总
                  </CardTitle>

                  <div className="flex flex-wrap gap-2">
                    <Button
                      variant="outline"
                      size="sm"
                      className="rounded-full"
                      onClick={() => {
                          track("open_wrong_book_words", { wrong_count: wrongWordSummary.length });
                          openFlashcardDeck("wrong", { filter: "all", flashcardMode: "recognition" });
                        }}
                    >
                      去刷单词错题
                    </Button>
                    <Button
                      variant="outline"
                      size="sm"
                      className="rounded-full"
                      onClick={() => {
                        setStudyView("quiz");
                        setQuizMode("bb_pairs");
                        setPairReviewMode("wrong");
                      }}
                    >
                      去刷六选二错题
                    </Button>
                    <Button
                      variant="outline"
                      size="sm"
                      className="rounded-full"
                      onClick={() => {
                        if (reviewSummaryTab === "words") clearAllWrongWords();
                        else clearAllWrongPairs();
                      }}
                      disabled={reviewSummaryTab === "words" ? !wrongWordSummary.length : !wrongPairSummary.length}
                    >
                      <Trash2 className="mr-2 h-4 w-4" />
                      清空当前错题表
                    </Button>
                  </div>
                </div>

                <Tabs value={reviewSummaryTab} onValueChange={setReviewSummaryTab} className="w-full">
                  <TabsList className="grid w-full grid-cols-2 gap-2 h-auto bg-transparent p-0">
                    <TabsTrigger value="words" className="rounded-xl border px-3 py-2">
                      单词错题表（{wrongWordSummary.length}）
                    </TabsTrigger>
                    <TabsTrigger value="pairs" className="rounded-xl border px-3 py-2">
                      六选二错题表（{wrongPairSummary.length}）
                    </TabsTrigger>
                  </TabsList>
                </Tabs>

                <div className="relative">
                  <Search className="absolute left-3 top-1/2 h-4 w-4 -translate-y-1/2 text-slate-400" />
                  <Input
                    value={reviewSummarySearch}
                    onChange={(e) => setReviewSummarySearch(e.target.value)}
                    placeholder={reviewSummaryTab === "words" ? "搜索单词 / 中文义 / 英文义..." : "搜索词对 / 中文释义 / 频率 / 来源..."}
                    className="rounded-2xl pl-9"
                  />
                </div>
              </CardHeader>

              <CardContent>
                {reviewSummaryTab === "words" ? (
                  filteredWrongWordSummary.length ? (
                    <div className="space-y-2">
                      {filteredWrongWordSummary.map((item, index) => (
                        <div key={item.id} className="rounded-2xl border bg-slate-50 px-4 py-3">
                          <div className="flex flex-col gap-2 md:flex-row md:items-start md:justify-between">
                            <div className="min-w-0">
                              <div className="text-base font-semibold">
                                {index + 1}. {item.word}
                              </div>
                              <div className="mt-1 text-sm text-slate-700">{item.meaning || "未填写中文义"}</div>
                              {item.en ? <div className="mt-1 text-xs text-slate-500">{item.en}</div> : null}
                            </div>

                            <div className="flex flex-wrap items-center gap-2">
                              <Badge variant="outline">总错 {item.totalWrong}</Badge>
                              <Badge variant="outline">闪卡不认识 {item.unknownCount}</Badge>
                              <Badge variant="outline">Quiz 错误 {item.quizWrongCount}</Badge>
                              <Button
                                variant="outline"
                                size="sm"
                                className="rounded-full"
                                onClick={() => clearWrongWordRecord(item.id)}
                              >
                                <Trash2 className="mr-2 h-4 w-4" />
                                清空这条
                              </Button>
                            </div>
                          </div>
                        </div>
                      ))}
                    </div>
                  ) : (
                    <div className="rounded-2xl border border-dashed px-4 py-8 text-center text-sm text-slate-500">
                      这里还没有单词错题。
                    </div>
                  )
                ) : filteredWrongPairSummary.length ? (
                  <div className="space-y-2">
                    {filteredWrongPairSummary.map((item, index) => (
                      <div key={item.id} className="rounded-2xl border bg-slate-50 px-4 py-3">
                        <div className="flex flex-col gap-2 md:flex-row md:items-start md:justify-between">
                          <div className="min-w-0">
                            <div className="text-base font-semibold">
                              {index + 1}. {item.a} = {item.b}
                            </div>
                            <div className="mt-1 text-sm text-slate-700">{item.zh || "未填写中文释义"}</div>
                            <div className="mt-1 text-xs text-slate-500">
                              {[item.frequency, item.source].filter(Boolean).join(" · ")}
                            </div>
                          </div>

                          <div className="flex flex-wrap items-center gap-2">
                            <Badge variant="outline">错误次数 {item.wrongCount}</Badge>
                            <Button
                              variant="outline"
                              size="sm"
                              className="rounded-full"
                              onClick={() => clearWrongPairRecord(item.id)}
                            >
                              <Trash2 className="mr-2 h-4 w-4" />
                              清空这条
                            </Button>
                          </div>
                        </div>
                      </div>
                    ))}
                  </div>
                ) : (
                  <div className="rounded-2xl border border-dashed px-4 py-8 text-center text-sm text-slate-500">
                    这里还没有六选二错题。
                  </div>
                )}
              </CardContent>
            </Card>

            <Card className="rounded-2xl shadow-sm">
              <CardHeader>
                <CardTitle className="text-lg">今日计划模块</CardTitle>
              </CardHeader>
              <CardContent className="space-y-5">
                <div className="rounded-2xl border p-4 space-y-4">
                  <div className="flex flex-wrap items-center justify-between gap-3">
                    <div>
                      <div className="text-sm text-slate-500">今日执行状态</div>
                      
                    </div>
                    <Badge variant="outline" className={`rounded-full ${overallPlanMeta.badgeClass}`}>{overallPlanMeta.label}</Badge>
                  </div>
                  <div className="grid gap-3 md:grid-cols-4">
                    <div className="rounded-2xl bg-slate-50 p-3 space-y-2">
                      <div className="flex items-center justify-between text-sm"><span className="font-medium text-slate-700">今日新词</span><Badge variant="outline" className={`rounded-full ${newWordPlanMeta.badgeClass}`}>{newWordPlanMeta.label}</Badge></div>
                      <Progress value={todayNewWordProgressPct} className="h-2" />
                      <div className="text-xs text-slate-500">{todayNewWordStudyCount} / {todayNewWordTarget}，{todayNewWordRemaining > 0 ? `还差 ${todayNewWordRemaining} 个` : "已达标"}</div>
                    </div>
                    <div className="rounded-2xl bg-slate-50 p-3 space-y-2">
                      <div className="flex items-center justify-between text-sm"><span className="font-medium text-slate-700">今日复习</span><Badge variant="outline" className={`rounded-full ${reviewWordPlanMeta.badgeClass}`}>{reviewWordPlanMeta.label}</Badge></div>
                      <Progress value={todayReviewWordProgressPct} className="h-2" />
                      <div className="text-xs text-slate-500">{todayReviewWordStudyCount} / {todayReviewWordTarget}，{todayReviewWordRemaining > 0 ? `还差 ${todayReviewWordRemaining} 个` : "已达标"}</div>
                    </div>
                    <div className="rounded-2xl bg-slate-50 p-3 space-y-2">
                      <div className="flex items-center justify-between text-sm"><span className="font-medium text-slate-700">六选二任务</span><Badge variant="outline" className={`rounded-full ${pairPlanMeta.badgeClass}`}>{pairPlanMeta.label}</Badge></div>
                      <Progress value={todayPairProgressPct} className="h-2" />
                      <div className="text-xs text-slate-500">{todayPairPracticeCount} / {todayPairTarget}，{todayPairRemaining > 0 ? `还差 ${todayPairRemaining} 对` : "已达标"}</div>
                    </div>
                    <div className="rounded-2xl bg-slate-50 p-3 space-y-2">
                      <div className="flex items-center justify-between text-sm"><span className="font-medium text-slate-700">Quiz 状态</span><Badge variant="outline" className={`rounded-full ${quizPlanMeta.badgeClass}`}>{quizPlanMeta.label}</Badge></div>
                      <Progress value={todayQuizProgressPct} className="h-2" />
                      <div className="text-xs text-slate-500">{todayQuizAccuracy === null ? "今天还没开始 quiz" : `当前 ${todayQuizAccuracy}% / 目标 ${dailyPlan.quizTarget}%`}</div>
                    </div>
                  </div>
                </div>

                <div className="grid gap-4 md:grid-cols-3">
                  <div className="rounded-2xl border p-4 space-y-2">
                    <div className="text-sm font-medium text-slate-700">完成节奏</div>
                    <div className="grid gap-3 sm:grid-cols-2">
                      <div><div className="mb-1 text-xs text-slate-500">单词几天背完</div><Input type="number" min="1" value={dailyPlan.wordFinishDays} onChange={(e) => updateDailyPlan("wordFinishDays", e.target.value)} className="rounded-xl" /></div>
                      <div><div className="mb-1 text-xs text-slate-500">六选二几天背完</div><Input type="number" min="1" value={dailyPlan.pairFinishDays} onChange={(e) => updateDailyPlan("pairFinishDays", e.target.value)} className="rounded-xl" /></div>
                    </div>
                    <div className="text-xs text-slate-600">按当前库存，建议每天新背 <span className="font-semibold">{dailyWordNewTarget}</span> 个单词、<span className="font-semibold">{dailyPairNewTarget}</span> 对六选二。</div>
                  </div>
                  <div className="rounded-2xl border p-4 space-y-2">
                    <div className="text-sm font-medium text-slate-700">执行配额</div>
                    <div className="grid gap-3 sm:grid-cols-2">
                      <div><div className="mb-1 text-xs text-slate-500">单词每日复习</div><Input type="number" min="0" value={dailyPlan.wordReviewCount} onChange={(e) => updateDailyPlan("wordReviewCount", e.target.value)} className="rounded-xl" /></div>
                      <div><div className="mb-1 text-xs text-slate-500">六选二每日复习</div><Input type="number" min="0" value={dailyPlan.pairReviewCount} onChange={(e) => updateDailyPlan("pairReviewCount", e.target.value)} className="rounded-xl" /></div>
                    </div>
                    <div className="space-y-2 pt-1">
                      <div><div className="flex items-center justify-between text-xs text-slate-500"><span>今日单词总量</span><span>{todayWordStudyCount} / {todayWordTarget}</span></div><Progress value={Math.min(100, todayWordTarget ? (todayWordStudyCount / todayWordTarget) * 100 : 0)} className="mt-1 h-2" /><div className="mt-1 text-xs text-slate-600">{todayWordRemaining > 0 ? `今天还差 ${todayWordRemaining} 个单词学习量` : "今天的单词学习量已达标"}</div></div>
                      <div><div className="flex items-center justify-between text-xs text-slate-500"><span>今日六选二总量</span><span>{todayPairPracticeCount} / {todayPairTarget}</span></div><Progress value={Math.min(100, todayPairTarget ? (todayPairPracticeCount / todayPairTarget) * 100 : 0)} className="mt-1 h-2" /><div className="mt-1 text-xs text-slate-600">{todayPairRemaining > 0 ? `今天还差 ${todayPairRemaining} 对六选二练习量` : "今天的六选二练习量已达标"}</div></div>
                    </div>
                  </div>
                  <div className="rounded-2xl border p-4 space-y-2">
                    <div className="text-sm font-medium text-slate-700">质量目标</div>
                    <div><div className="mb-1 text-xs text-slate-500">Quiz 成绩目标</div><Input type="number" min="0" max="100" value={dailyPlan.quizTarget} onChange={(e) => updateDailyPlan("quizTarget", e.target.value)} className="rounded-xl" /></div>
                    <div className="rounded-xl bg-slate-50 p-3 text-sm text-slate-700 space-y-1">
                      <div>今日单词 quiz：{todayQuizCount ? `${todayStat.quizCorrect || 0} / ${todayQuizCount}` : "还没开始"}</div>
                      <div>当前准确率：{todayQuizAccuracy === null ? "暂无" : `${todayQuizAccuracy}%`}</div>
                      <div>{todayQuizAccuracy === null ? "今天至少做一轮 quiz，才能判断是否达标。" : todayQuizAccuracy >= dailyPlan.quizTarget ? "今天 quiz 准确率已达标。" : `今天 quiz 还没到目标，还差 ${dailyPlan.quizTarget - todayQuizAccuracy}%`}</div>
                    </div>
                  </div>
                </div>
              </CardContent>
            </Card>
          </div>
        </div>
      </div>
    </div>
  );
}
