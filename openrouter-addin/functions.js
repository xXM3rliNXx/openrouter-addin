/* global CustomFunctions */
/*
 * OpenRouter-Funktion fuer Excel: =OPENROUTER.PROMPT(Prompt; [Modell]; [Systemprompt])
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

async function callOpenRouter(model, system, prompt) {
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
      body: JSON.stringify({ model, messages })
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

// ---------- Zellfunktion ----------
const inflight = new Map();

/**
 * Sendet einen Prompt an OpenRouter und gibt die Antwort zurueck.
 * @param {any} prompt   Prompt (Text oder Zellbezug)
 * @param {any} [model]  Modell-ID, leer = Standardmodell aus den Einstellungen
 * @param {any} [system] Systemprompt
 * @returns {Promise<string>}
 */
async function prompt(prompt, model, system) {
  prompt = prompt == null ? "" : String(prompt);
  if (prompt.trim() === "") return "";
  model = model == null || String(model).trim() === "" ? ORSettings.model : String(model).trim();
  system = system == null ? "" : String(system);

  if (!ORSettings.apiKey) {
    return "#FEHLER: Kein API-Key - bitte unter Start > OpenRouter eintragen";
  }

  const key = JSON.stringify([model, system, prompt]);
  if (ORCache.has(key)) return ORCache.get(key);
  if (inflight.has(key)) return inflight.get(key);   // gleiche Anfrage laeuft schon

  const job = (async () => {
    await acquire();
    try {
      const answer = (await callOpenRouter(model, system, prompt)).slice(0, MAX_CELL_LENGTH);
      ORCache.set(key, answer);
      return answer;
    } catch (e) {
      return "#FEHLER " + (e && e.message ? e.message : e);
    } finally {
      release();
      inflight.delete(key);
    }
  })();

  inflight.set(key, job);
  return job;
}

if (typeof CustomFunctions !== "undefined") {
  CustomFunctions.associate("PROMPT", prompt);
}

// fuer Tests unter Node
if (typeof module !== "undefined") {
  module.exports = { prompt, ORSettings, ORCache, ORStatus };
}
