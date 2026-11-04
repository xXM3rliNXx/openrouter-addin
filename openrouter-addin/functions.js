/* global CustomFunctions */
/*
 * OpenRouter-Funktionen fuer Excel
 *
 *   =OPENROUTER.PROMPT(Prompt; [Modell]; [Systemprompt])
 *   =OPENROUTER.BEWERTEN(Bereich; Kriterium; [Modell])
 *
 * - laeuft asynchron (Zelle zeigt #BUSY!, Excel bleibt bedienbar)
 * - begrenzt gleichzeitige Anfragen (Einstellung "Parallele Anfragen")
 * - wiederholt bei 429/5xx automatisch mit Wartezeit
 * - Antworten werden lokal gecacht (auch ueber Excel-Neustarts), Fehler nicht
 */

const API_URL = "https://openrouter.ai/api/v1/chat/completions";
const CACHE_STORAGE_KEY = "openrouter.cache.v1";
const MAX_CELL_LENGTH = 32767;
const MAX_RETRIES = 3;

// ---------- Einstellungen (werden im Aufgabenbereich gesetzt) ----------
const ORSettings = {
  get apiKey() { return localStorage.getItem("openrouter.apiKey") || ""; },
  set apiKey(v) { localStorage.setItem("openrouter.apiKey", v || ""); },
  get model() { return localStorage.getItem("openrouter.model") || "openrouter/auto"; },
  set model(v) { localStorage.setItem("openrouter.model", v || ""); },
  get maxParallel() {
    const n = parseInt(localStorage.getItem("openrouter.maxParallel"), 10);
    return n > 0 ? n : 8;
  },
  set maxParallel(v) { localStorage.setItem("openrouter.maxParallel", String(v)); }
};

// ---------- Cache ----------
const ORCache = (() => {
  let map;
  try {
    map = new Map(JSON.parse(localStorage.getItem(CACHE_STORAGE_KEY) || "[]"));
  } catch (e) {
    map = new Map();
  }
  let timer = null;

  function persist() {
    clearTimeout(timer);
    timer = setTimeout(() => {
      try {
        localStorage.setItem(CACHE_STORAGE_KEY, JSON.stringify([...map]));
      } catch (e) {
        // Speicher voll: aelteste Haelfte verwerfen und nochmal versuchen
        const entries = [...map];
        map = new Map(entries.slice(Math.floor(entries.length / 2)));
        try { localStorage.setItem(CACHE_STORAGE_KEY, JSON.stringify([...map])); } catch (e2) { /* egal */ }
      }
    }, 1000);
  }

  return {
    has: (k) => map.has(k),
    get: (k) => map.get(k),
    set: (k, v) => { map.set(k, v); persist(); },
    clear: () => { map = new Map(); persist(); },
    get size() { return map.size; }
  };
})();

// ---------- Begrenzung gleichzeitiger Anfragen ----------
let active = 0;
const waiting = [];

function acquire() {
  return new Promise((resolve) => {
    if (active < ORSettings.maxParallel) {
      active++;
      resolve();
    } else {
      waiting.push(resolve);
    }
  });
}

function release() {
  const next = waiting.shift();
  if (next) next();      // Slot direkt weitergeben
  else active--;
}

const ORStatus = {
  get running() { return active; },
  get queued() { return waiting.length; }
};

// ---------- API-Aufruf ----------
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
const errText = (e) => (e && e.message ? e.message : String(e));

async function callOpenRouter(model, system, prompt, extra) {
  const messages = [];
  if (system) messages.push({ role: "system", content: system });
  messages.push({ role: "user", content: prompt });

  for (let attempt = 0; ; attempt++) {
    const res = await fetch(API_URL, {
      method: "POST",
      headers: {
        "Authorization": "Bearer " + ORSettings.apiKey,
        "Content-Type": "application/json",
        "X-Title": "Excel OpenRouter Add-in"
      },
      body: JSON.stringify(Object.assign({ model, messages }, extra || {}))
    });

    if ((res.status === 429 || res.status >= 500) && attempt < MAX_RETRIES) {
      await sleep(2000 * Math.pow(2, attempt));   // 2s, 4s, 8s
      continue;
    }

    let data = {};
    try { data = await res.json(); } catch (e) { /* kein JSON */ }

    if (!res.ok || data.error) {
      const msg = (data.error && data.error.message) || res.statusText || "unbekannter Fehler";
      throw new Error(res.status + ": " + msg);
    }

    const content = data.choices && data.choices[0] && data.choices[0].message
      ? data.choices[0].message.content
      : null;
    if (content == null) throw new Error("Keine Antwort im Response");
    return String(content).trim();
  }
}

// Cache + gleiche laufende Anfragen zusammenfassen + Parallelitaet begrenzen.
// Fehler werden nicht gecacht, sondern an den Aufrufer weitergereicht.
const inflight = new Map();

function runCached(key, task) {
  if (ORCache.has(key)) return Promise.resolve(ORCache.get(key));
  if (inflight.has(key)) return inflight.get(key);

  const job = (async () => {
    await acquire();
    try {
      const value = await task();
      ORCache.set(key, value);
      return value;
    } finally {
      release();
      inflight.delete(key);
    }
  })();

  inflight.set(key, job);
  return job;
}

const NO_KEY = "#FEHLER: Kein API-Key - bitte unter Start > OpenRouter eintragen";

function pickModel(model) {
  return model == null || String(model).trim() === "" ? ORSettings.model : String(model).trim();
}

// ================== =OPENROUTER.PROMPT ==================

/**
 * Sendet einen Prompt an OpenRouter und gibt die Antwort zurueck.
 */
async function prompt(prompt, model, system) {
  prompt = prompt == null ? "" : String(prompt);
  if (prompt.trim() === "") return "";
  model = pickModel(model);
  system = system == null ? "" : String(system);
  if (!ORSettings.apiKey) return NO_KEY;

  const key = JSON.stringify([model, system, prompt]);
  try {
    return await runCached(key, async () =>
      (await callOpenRouter(model, system, prompt)).slice(0, MAX_CELL_LENGTH));
  } catch (e) {
    return "#FEHLER " + errText(e);
  }
}

// ================== =OPENROUTER.BEWERTEN ==================

const BEWERTEN_SYSTEM =
  "Du bewertest Eintraege vergleichend nach einem vorgegebenen Kriterium. " +
  "Lies zuerst alle Eintraege und vergleiche sie miteinander. " +
  "Gib dann jedem Eintrag eine ganze Zahl von 0 bis 100: Der Eintrag, der das Kriterium am besten erfuellt, " +
  "bekommt 100, der Eintrag, der es am schlechtesten erfuellt, bekommt 0. Die uebrigen liegen dazwischen, " +
  "entsprechend ihrem Abstand zum besten und schlechtesten Eintrag. Gleich gute Eintraege bekommen den gleichen Wert. " +
  "Antworte ausschliesslich mit einem JSON-Array aus Zahlen in derselben Reihenfolge wie die Eintraege, " +
  "ohne Erklaerung und ohne weiteren Text.";

function parseScores(text, n) {
  const a = text.indexOf("[");
  const b = text.lastIndexOf("]");
  if (a < 0 || b <= a) return null;
  let arr;
  try { arr = JSON.parse(text.slice(a, b + 1)); } catch (e) { return null; }
  if (!Array.isArray(arr) || arr.length !== n) return null;
  const nums = arr.map((x) => (x && typeof x === "object" ? (x.score ?? x.wert ?? x.value) : x)).map(Number);
  if (nums.some((x) => !isFinite(x))) return null;
  return nums.map((x) => Math.round(Math.max(0, Math.min(100, x))));
}

async function scoreItems(model, kriterium, items) {
  const n = items.length;
  const user =
    "Kriterium: " + kriterium + "\n\n" +
    "Eintraege (JSON-Array mit " + n + " Eintraegen):\n" + JSON.stringify(items) + "\n\n" +
    "Antworte mit einem JSON-Array aus genau " + n + " Zahlen.";

  for (let attempt = 0; attempt < 2; attempt++) {
    const text = await callOpenRouter(model, BEWERTEN_SYSTEM, user, { temperature: 0 });
    const scores = parseScores(text, n);
    if (scores) return scores;
  }
  throw new Error("Antwort des Modells nicht auswertbar (erwartet: " + n + " Zahlen)");
}

function emptyLike(grid) {
  return grid.map((row) => row.map(() => ""));
}

function errorGrid(grid, msg) {
  const out = emptyLike(grid);
  out[0][0] = msg;
  return out;
}

/**
 * Bewertet alle Zellen eines Bereichs vergleichend nach einem Kriterium.
 * Gibt einen Bereich gleicher Form zurueck (100 = am besten, 0 = am schlechtesten).
 */
async function bewerten(bereich, kriterium, model) {
  const grid = Array.isArray(bereich) ? bereich : [[bereich]];
  kriterium = kriterium == null ? "" : String(kriterium).trim();
  if (!kriterium) return errorGrid(grid, "#FEHLER: Kriterium fehlt");
  if (!ORSettings.apiKey) return errorGrid(grid, NO_KEY);
  model = pickModel(model);

  const items = [];
  const pos = [];
  grid.forEach((row, i) => row.forEach((v, j) => {
    const s = v == null ? "" : String(v).trim();
    if (s) { items.push(s); pos.push([i, j]); }
  }));

  const out = emptyLike(grid);
  if (items.length === 0) return out;
  if (items.length === 1) { out[pos[0][0]][pos[0][1]] = 100; return out; }

  const key = JSON.stringify(["BEWERTEN", model, kriterium, items]);
  let scores;
  try {
    scores = await runCached(key, () => scoreItems(model, kriterium, items));
  } catch (e) {
    return errorGrid(grid, "#FEHLER " + errText(e));
  }

  pos.forEach(([i, j], k) => { out[i][j] = scores[k]; });
  return out;
}

// ---------- Registrierung ----------
if (typeof CustomFunctions !== "undefined") {
  CustomFunctions.associate("PROMPT", prompt);
  CustomFunctions.associate("BEWERTEN", bewerten);
}

// fuer Tests unter Node
if (typeof module !== "undefined") {
  module.exports = { prompt, bewerten, parseScores, ORSettings, ORCache, ORStatus };
}
