// ======================
// Version & Diagnostics
// ======================
const __APP_VERSION__ = "1.0.0";

function logDiag(msg, bad = false) {
  console[bad ? "error" : "log"]("[PW]", msg);
  const box = document.getElementById("diagBox");
  if (box) {
    const div = document.createElement("div");
    div.className = "diag" + (bad ? " bad" : "");
    div.textContent = msg;
    box.appendChild(div);
  }
}

// ===============
// Data Services
// ===============
const STORAGE_KEY = "pw_u_entries_v1";
const SETTINGS_KEY = "pw_u_settings_v1";
const HISTORY_KEY = "pw_u_history_v1";
const GRID_KEY = "pw_u_grid_v1"; // last loaded raw grid (array of rows)
const PRIZES_KEY = "pw_u_prizes_v1"; // prize list (ordered)
const PRIZE_IDX_KEY = "pw_u_prize_idx_v1"; // next prize index

const DEFAULT_SETTINGS = {
  stopGapMs: 320,
  noRepeat: true,
  spinDurationMs: 3200,
  perReelStaggerMs: 160,
  confettiLevel: "medium",
  sound: true,
  fanfare: true,
  tick: true,
  fanfareSample: true,
  seed: "",
  autoFullscreen: false,
  gsUrl: "",
  gsAutoRefreshSec: 0,
  gsFilter: { enabled: false, colIndex: null, value: "" },
  map: { card: 0, name: 1 },
  status: { col: null, value: "" },
  showProgress: false,
  displayMode: "spinner",
  // Google workbook connection
  gsWorkbookId: "",
  gsSheets: [], // [{gid,title}]
  attSheetGid: "",
  attIgnoreHeader: false,
  prizeSheetGid: "",
  prizeIgnoreHeader: false,
};

// IndexedDB-backed, with in-memory cache and graceful localStorage fallback.
const DataService = (() => {
  const DB_NAME = "pw_store";
  const DB_VER = 1;
  const STORE = "kv";
  let db = null;
  let ready = false;
  const cache = {
    [STORAGE_KEY]: [],
    [SETTINGS_KEY]: {},
    [HISTORY_KEY]: [],
    [GRID_KEY]: [],
    [PRIZES_KEY]: [],
    [PRIZE_IDX_KEY]: 0,
  };

  function openDB() {
    return new Promise((resolve, reject) => {
      if (!("indexedDB" in window)) return resolve(null);
      const req = indexedDB.open(DB_NAME, DB_VER);
      req.onupgradeneeded = (e) => {
        const d = req.result;
        if (!d.objectStoreNames.contains(STORE))
          d.createObjectStore(STORE, { keyPath: "key" });
      };
      req.onsuccess = () => resolve(req.result);
      req.onerror = () => reject(req.error);
    });
  }
  function idbGet(key) {
    return new Promise((resolve) => {
      if (!db) return resolve(undefined);
      const tx = db.transaction(STORE, "readonly");
      const st = tx.objectStore(STORE);
      const rq = st.get(key);
      rq.onsuccess = () => resolve(rq.result ? rq.result.val : undefined);
      rq.onerror = () => resolve(undefined);
    });
  }
  function idbSet(key, val) {
    return new Promise((resolve) => {
      if (!db) return resolve(false);
      const tx = db.transaction(STORE, "readwrite");
      const st = tx.objectStore(STORE);
      st.put({ key, val });
      tx.oncomplete = () => resolve(true);
      tx.onerror = () => resolve(false);
    });
  }
  function lsGet(key) {
    try {
      const raw = localStorage.getItem(key);
      return raw ? JSON.parse(raw) : undefined;
    } catch {
      return undefined;
    }
  }
  function lsSet(key, val) {
    try {
      localStorage.setItem(key, JSON.stringify(val));
    } catch {}
  }

  async function init() {
    if (ready) return;
    try {
      db = await openDB();
    } catch {
      db = null;
    }
    // Load/migrate values into cache
    const keys = [
      SETTINGS_KEY,
      STORAGE_KEY,
      HISTORY_KEY,
      GRID_KEY,
      PRIZES_KEY,
      PRIZE_IDX_KEY,
    ];
    for (const k of keys) {
      let v = await idbGet(k);
      if (v === undefined) {
        // migrate from localStorage if present
        v = lsGet(k);
        if (v !== undefined) await idbSet(k, v);
      }
      if (v === undefined) {
        // Defaults
        if (k === SETTINGS_KEY) v = DEFAULT_SETTINGS;
        else v = [];
      }
      cache[k] = v;
    }
    ready = true;
  }

  function getEntries() {
    return Array.isArray(cache[STORAGE_KEY]) ? cache[STORAGE_KEY] : [];
  }
  function setEntries(rows) {
    cache[STORAGE_KEY] = Array.isArray(rows) ? rows : [];
    idbSet(STORAGE_KEY, cache[STORAGE_KEY]) ||
      lsSet(STORAGE_KEY, cache[STORAGE_KEY]);
  }
  function clearEntries() {
    setEntries([]);
  }

  function getSettings() {
    return { ...DEFAULT_SETTINGS, ...(cache[SETTINGS_KEY] || {}) };
  }
  function setSettings(s) {
    cache[SETTINGS_KEY] = s || {};
    idbSet(SETTINGS_KEY, cache[SETTINGS_KEY]) ||
      lsSet(SETTINGS_KEY, cache[SETTINGS_KEY]);
  }

  function getHistory() {
    return Array.isArray(cache[HISTORY_KEY]) ? cache[HISTORY_KEY] : [];
  }
  function setHistory(h) {
    cache[HISTORY_KEY] = Array.isArray(h) ? h : [];
    idbSet(HISTORY_KEY, cache[HISTORY_KEY]) ||
      lsSet(HISTORY_KEY, cache[HISTORY_KEY]);
  }
  function clearHistory() {
    setHistory([]);
  }

  function getGrid() {
    return Array.isArray(cache[GRID_KEY]) ? cache[GRID_KEY] : [];
  }
  function setGrid(grid) {
    cache[GRID_KEY] = Array.isArray(grid) ? grid : [];
    idbSet(GRID_KEY, cache[GRID_KEY]) || lsSet(GRID_KEY, cache[GRID_KEY]);
  }

  function getPrizes() {
    return Array.isArray(cache[PRIZES_KEY]) ? cache[PRIZES_KEY] : [];
  }
  function setPrizes(list) {
    cache[PRIZES_KEY] = Array.isArray(list) ? list : [];
    idbSet(PRIZES_KEY, cache[PRIZES_KEY]) ||
      lsSet(PRIZES_KEY, cache[PRIZES_KEY]);
  }
  function getPrizeIndex() {
    return Number.isInteger(cache[PRIZE_IDX_KEY]) ? cache[PRIZE_IDX_KEY] : 0;
  }
  function setPrizeIndex(i) {
    cache[PRIZE_IDX_KEY] = Math.max(0, i | 0);
    idbSet(PRIZE_IDX_KEY, cache[PRIZE_IDX_KEY]) ||
      lsSet(PRIZE_IDX_KEY, cache[PRIZE_IDX_KEY]);
  }

  return {
    init,
    getEntries,
    setEntries,
    clearEntries,
    getSettings,
    setSettings,
    getHistory,
    setHistory,
    clearHistory,
    getGrid,
    setGrid,
    getPrizes,
    setPrizes,
    getPrizeIndex,
    setPrizeIndex,
  };
})();

// =============
// RNG helpers
// =============
function xorshift32(seed) {
  let x = seed | 0;
  return () => {
    x ^= x << 13;
    x ^= x >>> 17;
    x ^= x << 5;
    return (x >>> 0) / 0xffffffff;
  };
}
function seededRng(seedStr) {
  let h = 2166136261;
  for (let i = 0; i < seedStr.length; i++) {
    h ^= seedStr.charCodeAt(i);
    h = Math.imul(h, 16777619);
  }
  h = (h ^ (h >>> 16)) >>> 0;
  return xorshift32(h || 1);
}
function pickRandomIndex(n, settings) {
  if (n <= 0) return -1;
  if (settings.seed && settings.seed.trim()) {
    const r = seededRng(
      settings.seed + ":" + (DataService.getHistory().length || 0)
    );
    return Math.floor(r() * n);
  }
  if (window.crypto && crypto.getRandomValues) {
    const b = new Uint32Array(1);
    crypto.getRandomValues(b);
    return b[0] % n;
  }
  return Math.floor(Math.random() * n);
}

// =================
// Audio (safe)
// =================
const AudioFX = (() => {
  let ctx = null;
  try {
    ctx = new (window.AudioContext || window.webkitAudioContext)();
  } catch (e) {
    return { beep() {}, whoosh() {}, fanfare() {} };
  }
  const safe = (fn) => {
    if (ctx.state === "suspended") ctx.resume().finally(fn);
    else fn();
  };
  // Optional fanfare sample support
  let fanfareUrl = "resources/mixkit-winning-notification-2018.wav";
  let fanfareBuf = null;
  let fanfareLoading = null;
  function _decode(arrayBuffer) {
    // Prefer feature-detect on function arity to avoid double-decoding
    if (ctx.decodeAudioData.length === 1) {
      return ctx.decodeAudioData(arrayBuffer);
    }
    return new Promise((resolve, reject) => {
      ctx.decodeAudioData(arrayBuffer, resolve, reject);
    });
  }
  function preloadFanfare(url = fanfareUrl) {
    if (fanfareBuf) return Promise.resolve(fanfareBuf);
    if (fanfareLoading) return fanfareLoading;
    fanfareLoading = fetch(url)
      .then((r) => {
        if (!r.ok) throw new Error(`Failed to fetch fanfare: ${r.status}`);
        return r.arrayBuffer();
      })
      .then((ab) => _decode(ab))
      .then((buf) => {
        fanfareBuf = buf;
        return buf;
      })
      .catch((err) => {
        console.warn("[PW] fanfare preload failed", err);
        fanfareLoading = null;
        return null;
      });
    return fanfareLoading;
  }
  function setFanfareUrl(url) {
    fanfareUrl = url;
    fanfareBuf = null;
    fanfareLoading = null;
  }
  function _playBufferAt(t0, buf, { gain = 0.9, pan = 0, rate = 1 } = {}) {
    const src = ctx.createBufferSource();
    src.buffer = buf;
    src.playbackRate.setValueAtTime(rate, t0);
    const g = ctx.createGain();
    g.gain.setValueAtTime(gain, t0);
    const p = ctx.createStereoPanner?.() || null;
    if (p && p.pan?.setValueAtTime) p.pan.setValueAtTime(pan, t0);
    if (p) src.connect(g).connect(p).connect(ctx.destination);
    else src.connect(g).connect(ctx.destination);
    src.start(t0);
    return src;
  }

  // Simple ADSR-style envelope helper. Returns a function to schedule release.
  function _env(
    gainNode,
    t0,
    {
      attack = 0.005,
      decay = 0.03,
      sustain = 0.7,
      release = 0.06,
      peak = 0.05,
    } = {}
  ) {
    const p = gainNode.gain || gainNode; // allow passing GainNode or AudioParam-like
    p.cancelScheduledValues(t0);
    p.setValueAtTime(0, t0);
    p.linearRampToValueAtTime(peak, t0 + attack);
    p.linearRampToValueAtTime(peak * sustain, t0 + attack + decay);
    return (tEnd) => {
      p.linearRampToValueAtTime(0, tEnd + release);
    };
  }

  // Low-frequency oscillator for vibrato/tremolo
  function _lfo(
    t0,
    param,
    {
      type = "sine",
      freq = 6,
      depth = 5 /* cents for freq, or raw for gain */,
      unit = "cents",
    } = {}
  ) {
    if (!param) return { stop() {} };
    const o = ctx.createOscillator();
    const g = ctx.createGain();
    o.type = type;
    o.frequency.setValueAtTime(freq, t0);
    let depthValue = depth;
    if (unit === "cents") {
      // approximate cents->Hz around nominal; caller can pre-scale
      // For small modulation, scale relative to 440; caller may pre-scale for specific freq
      depthValue = (depth / 1200) * Math.log(2) * 440;
    }
    g.gain.setValueAtTime(Math.max(0, depthValue), t0);
    o.connect(g).connect(param);
    o.start(t0);
    return { stop: (t) => o.stop(t) };
  }

  // Core beep scheduler used by both public beep() and fanfare()
  function _playBeepAt(
    t0,
    {
      freq = 880,
      dur = 0.1,
      type = "sine",
      gain = 0.05,
      attack = 0.005,
      decay = 0.03,
      sustain = 0.8,
      release = 0.06,
      sweep = { to: 500 }, // { to: number }
      vibrato = null, // { freq: number, depth: number (cents) }
      pan = 0,
    } = {}
  ) {
    const osc = ctx.createOscillator();
    const g = ctx.createGain();
    const p = ctx.createStereoPanner?.() || null;

    osc.type = type;
    osc.frequency.setValueAtTime(freq, t0);
    if (sweep && sweep.to != null) {
      osc.frequency.linearRampToValueAtTime(sweep.to, t0 + dur);
    }

    if (p && p.pan?.setValueAtTime) p.pan.setValueAtTime(pan, t0);

    const scheduleRelease = _env(g, t0, {
      attack,
      decay,
      sustain,
      release,
      peak: gain,
    });
    let vib;
    if (vibrato) {
      const depthHz = ((vibrato.depth || 5) / 1200) * Math.log(2) * freq; // cents->Hz near freq
      vib = _lfo(t0, osc.frequency, {
        freq: vibrato.freq || 6,
        depth: depthHz,
        unit: "hz",
      });
    }

    if (p) osc.connect(g).connect(p).connect(ctx.destination);
    else osc.connect(g).connect(ctx.destination);

    const tEnd = t0 + dur;
    osc.start(t0);
    scheduleRelease(tEnd);
    const tStop = tEnd + release + 0.01;
    osc.stop(tStop);
    if (vib) vib.stop(tStop);
  }

  // Backward-compatible signature: (freq, dur, type, gain) or options object
  function beep(a = 10, b = 0.04, c = "square", d = 0.03) {
    const opt =
      typeof a === "object" ? a : { freq: a, dur: b, type: c, gain: d };
    const { startAt = null } = opt;
    safe(() => {
      const t0 = startAt != null ? startAt : ctx.currentTime + 0.001;
      _playBeepAt(t0, opt);
    });
  }

  // Backward-compatible whoosh: (dur) or options
  function whoosh(a = 0.8) {
    const opt = typeof a === "object" ? a : { dur: a };
    const {
      dur = 0.9,
      gain = 0.25,
      pan = 0,
      band = { type: "bandpass", start: 200, end: 2000, Q: 1.2 },
    } = opt;
    safe(() => {
      const t0 = ctx.currentTime + 0.001;
      const len = Math.max(0.05, dur);
      const buffer = ctx.createBuffer(
        1,
        Math.floor(ctx.sampleRate * len),
        ctx.sampleRate
      );
      const data = buffer.getChannelData(0);
      for (let i = 0; i < data.length; i++) {
        const x = Math.random() * 2 - 1;
        const e = 1 - i / data.length;
        data[i] = x * e; // fade-out shaped white noise
      }
      const src = ctx.createBufferSource();
      src.buffer = buffer;
      const filt = ctx.createBiquadFilter();
      filt.type = band.type || "bandpass";
      filt.Q.setValueAtTime(band.Q ?? 1, t0);
      filt.frequency.setValueAtTime(band.start ?? 200, t0);
      if (band.end)
        filt.frequency.exponentialRampToValueAtTime(
          Math.max(50, band.end),
          t0 + dur
        );

      const g = ctx.createGain();
      g.gain.setValueAtTime(0, t0);
      g.gain.linearRampToValueAtTime(gain, t0 + 0.02);
      g.gain.exponentialRampToValueAtTime(0.0001, t0 + dur);

      const p = ctx.createStereoPanner?.() || null;
      if (p && p.pan?.setValueAtTime) p.pan.setValueAtTime(pan, t0);

      if (p) src.connect(filt).connect(g).connect(p).connect(ctx.destination);
      else src.connect(filt).connect(g).connect(ctx.destination);

      src.start(t0);
      src.stop(t0 + dur + 0.02);
    });
  }

  function fanfare(opts = {}) {
    const o = typeof opts === "object" ? opts : {};
    const notes = o.notes || [523.25, 659.25, 783.99, 1046.5];
    const tempo = o.tempo || 180; // bpm
    const type = o.type || "triangle";
    const gain = o.gain != null ? o.gain : 0.06;
    const gap = o.gap != null ? o.gap : 0.04; // seconds between notes
    const sustain = o.sustain != null ? o.sustain : 0.8;
    const attack = o.attack != null ? o.attack : 0.005;
    const decay = o.decay != null ? o.decay : 0.03;
    const release = o.release != null ? o.release : 0.04;
    const useSample = o.sample !== false; // default true
    safe(() => {
      const tNow = ctx.currentTime + 0.01;
      if (useSample && fanfareBuf) {
        _playBufferAt(tNow, fanfareBuf, {
          gain: o.sampleGain ?? 0.9,
          pan: o.samplePan ?? 0,
          rate: o.sampleRate ?? 1,
        });
      } else {
        // Start preload in background if desired, but play synth immediately
        if (useSample)
          try {
            preloadFanfare(o.sampleUrl || fanfareUrl);
          } catch {}
        synth();
      }
      function synth() {
        const beat = 60 / tempo;
        let t = tNow;
        for (const f of notes) {
          _playBeepAt(t, {
            freq: f,
            dur: beat * 0.45,
            type,
            gain,
            attack,
            decay,
            sustain,
            release,
          });
          t += beat + gap;
        }
      }
    });
  }

  return { beep, whoosh, fanfare, preloadFanfare, setFanfareUrl };
})();

// ========================
// Confetti (guarded)
// ========================
const Confetti = (() => {
  const canvas = document.getElementById("confetti");
  if (!canvas) return { burst: () => {} };
  const ctx = canvas.getContext("2d");
  let raf = 0,
    running = false,
    parts = [],
    last = 0;
  function size() {
    const dpr = Math.max(1, window.devicePixelRatio || 1);
    const w = Math.floor(window.innerWidth),
      h = Math.floor(window.innerHeight);
    const cw = Math.floor(w * dpr),
      ch = Math.floor(h * dpr);
    if (canvas.width !== cw || canvas.height !== ch) {
      canvas.width = cw;
      canvas.height = ch;
    }
    ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
  }
  function start() {
    if (running) return;
    running = true;
    last = performance.now();
    const loop = (now) => {
      if (!running) {
        ctx.clearRect(0, 0, canvas.width, canvas.height);
        return;
      }
      const dt = Math.min(40, now - last);
      last = now;
      const g = 0.0025 * dt,
        drag = 0.0008 * dt;
      const W = window.innerWidth,
        H = window.innerHeight;
      ctx.clearRect(0, 0, W, H);
      let alive = 0;
      for (let i = 0; i < parts.length; i++) {
        const p = parts[i];
        if (!p.active) continue;
        p.vy += g;
        p.vx *= 1 - drag;
        p.x += p.vx * (dt * 0.06);
        p.y += p.vy * (dt * 0.06);
        p.rot += p.vr * (dt * 0.06);
        if (p.y > H * 0.8) p.alpha -= 0.015 * (dt / 16);
        if (p.alpha > 0 && p.y < H + 40) alive++;
        ctx.save();
        ctx.globalAlpha = Math.max(0, p.alpha);
        ctx.translate(p.x, p.y);
        ctx.rotate(p.rot);
        ctx.fillStyle = `hsl(${p.hue},90%,60%)`;
        ctx.fillRect(-p.w / 2, -p.h / 2, p.w, p.h);
        ctx.restore();
      }
      if (alive === 0) {
        parts.length = 0;
        running = false;
        ctx.clearRect(0, 0, W, H);
      } else {
        raf = requestAnimationFrame(loop);
      }
    };
    raf = requestAnimationFrame(loop);
  }
  function burst(level = "medium") {
    size();
    const count = level === "light" ? 80 : level === "heavy" ? 240 : 150;
    const W = window.innerWidth,
      H = window.innerHeight;
    for (let i = 0; i < count; i++) {
      parts.push({
        active: true,
        x: Math.random() * W,
        y: -20 - Math.random() * H * 0.3,
        vx: -1.2 + Math.random() * 2.4,
        vy: 2 + Math.random() * 3.2,
        w: 6 + Math.random() * 6,
        h: 10 + Math.random() * 12,
        rot: Math.random() * Math.PI,
        vr: -0.25 + Math.random() * 0.5,
        hue: Math.floor(Math.random() * 360),
        alpha: 1,
      });
    }
    start();
  }
  window.addEventListener("resize", size);
  size();
  return { burst };
})();

// ==============================
// Canvas Odometer (no DOM jitter)
// ==============================
class OdometerCanvas {
  constructor(canvas) {
    this.canvas = canvas;
    this.ctx = canvas.getContext("2d");
    this.cols = 7;
    this.alphaNum = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ".split("");
    this.digits = "0123456789".split("");
    this.reels = new Array(this.cols).fill(0).map((_, i) => ({
      symbols: i === 0 ? this.alphaNum : this.digits,
      offset: 0,
      targetIndex: 0,
      spinning: false,
      startTime: 0,
      duration: 0,
      startOffset: 0,
      endOffset: 0,
    }));
    this.reelWidth = 0;
    this.reelHeight = 0;
    this.glyphSize = 0;
    this.raf = 0;
    this.fontFamily =
      "ui-monospace, SFMono-Regular, Menlo, Consolas, 'Liberation Mono', 'Roboto Mono', monospace";
    this.size();
    window.addEventListener("resize", () => this.size());
  }
  size() {
    const dpr = Math.max(1, window.devicePixelRatio || 1);
    const cssW = Math.min(window.innerWidth * 0.9, 1100),
      cssH = Math.min(Math.max(120, window.innerHeight * 0.28), 220);
    const cw = Math.floor(cssW * dpr),
      ch = Math.floor(cssH * dpr);
    if (this.canvas.width !== cw || this.canvas.height !== ch) {
      this.canvas.width = cw;
      this.canvas.height = ch;
    }
    this.ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
    this.reelWidth = Math.floor((cssW - 6 * 16) / this.cols);
    this.reelHeight = Math.floor(cssH);
    this.glyphSize = Math.floor(this.reelHeight * 0.75);
    this.ctx.font = `${this.glyphSize}px ${this.fontFamily}`;
    this.ctx.textAlign = "center";
    this.ctx.textBaseline = "middle";
  }
  draw() {
    const ctx = this.ctx,
      W = this.canvas.clientWidth,
      H = this.canvas.clientHeight;
    ctx.clearRect(0, 0, W, H);
    for (let i = 0; i < this.cols; i++) {
      const r = this.reels[i],
        x = i * (this.reelWidth + 16),
        rx = 14;
      ctx.save();
      // slot bg
      ctx.beginPath();
      ctx.moveTo(x + rx, 0);
      ctx.arcTo(x + this.reelWidth, 0, x + this.reelWidth, this.reelHeight, rx);
      ctx.arcTo(x + this.reelWidth, this.reelHeight, x, this.reelHeight, rx);
      ctx.arcTo(x, this.reelHeight, x, 0, rx);
      ctx.arcTo(x, 0, x + this.reelWidth, 0, rx);
      ctx.closePath();
      ctx.fillStyle = "#0e1633";
      ctx.fill();
      ctx.strokeStyle = "rgba(255,255,255,.08)";
      ctx.stroke();
      // center symbol + neighbors
      const centerY = this.reelHeight / 2,
        rowH = this.reelHeight;
      const sy = r.offset % (r.symbols.length * rowH);
      const baseIndex =
        ((Math.floor(sy / rowH) % r.symbols.length) + r.symbols.length) %
        r.symbols.length;
      for (let k = -1; k <= 1; k++) {
        const idx = (baseIndex + k + r.symbols.length) % r.symbols.length,
          ch = r.symbols[idx];
        const y = centerY + k * rowH - (sy % rowH);
        ctx.shadowColor = "rgba(110,231,255,.25)";
        ctx.shadowBlur = 10;
        ctx.fillStyle = "#e9f4ff";
        ctx.font = `${this.glyphSize}px ${this.fontFamily}`;
        ctx.fillText(ch, x + this.reelWidth / 2, y);
        ctx.shadowBlur = 0;
      }
      // vignettes
      const gt = ctx.createLinearGradient(0, 0, 0, this.reelHeight * 0.25);
      gt.addColorStop(0, "rgba(4,8,18,.85)");
      gt.addColorStop(1, "rgba(4,8,18,0)");
      ctx.fillStyle = gt;
      ctx.fillRect(x, 0, this.reelWidth, this.reelHeight * 0.25);
      const gb = ctx.createLinearGradient(
        0,
        this.reelHeight * 0.75,
        0,
        this.reelHeight
      );
      gb.addColorStop(0, "rgba(4,8,18,0)");
      gb.addColorStop(1, "rgba(4,8,18,.85)");
      ctx.fillStyle = gb;
      ctx.fillRect(
        x,
        this.reelHeight * 0.75,
        this.reelWidth,
        this.reelHeight * 0.25
      );
      ctx.restore();
    }
  }
  spinTo(cardStr, settings, onTick, onDone) {
    if (this.raf) {
      cancelAnimationFrame(this.raf);
      this.raf = 0;
    }
    const now = performance.now();
    const chars = cardStr.split("");
    for (let i = 0; i < this.cols; i++) {
      const r = this.reels[i];
      const targetChar = chars[i];
      const idx = r.symbols.indexOf(targetChar);
      r.targetIndex = idx;
      // All reels spin forward (same direction) and stop left->right one by one
      const baseDur = settings.spinDurationMs || 3200;
      const gap = settings.stopGapMs != null ? settings.stopGapMs : 320; // ms between stops
      r.startTime = now; // start together for drama
      r.duration = baseDur + i * gap; // stop sequentially left-to-right
      // Force extra full cycles so even unchanged digits still spin
      const rowH = this.reelHeight;
      const ring = r.symbols.length * rowH;
      const curRow = Math.round(r.offset / rowH);
      const len = r.symbols.length;
      const curIndex = ((curRow % len) + len) % len;
      const curCycle = Math.floor((curRow - curIndex) / len);
      const extraSpins = 10 + i; // at least 10 full rotations, increasing slightly per reel
      const targetRow = (curCycle + extraSpins) * len + idx;
      r.startOffset = r.offset;
      r.endOffset = targetRow * rowH; // strictly greater than startOffset
      r.spinning = true;
    }
    AudioFX.whoosh(1.0);
    const animate = (t) => {
      let allDone = true;
      for (let i = 0; i < this.cols; i++) {
        const r = this.reels[i];
        if (!r.spinning) continue;
        const p = Math.max(0, Math.min(1, (t - r.startTime) / r.duration));
        const eased = 1 - Math.pow(1 - p, 3);
        const prevRow = Math.floor(r.offset / this.reelHeight);
        const raw = r.startOffset + (r.endOffset - r.startOffset) * eased;
        r.offset = Math.round(raw);
        const row = Math.floor(r.offset / this.reelHeight);
        if (row !== prevRow && onTick) onTick(i);
        if (p >= 1) {
          r.offset = r.endOffset;
          r.spinning = false;
        } else {
          allDone = false;
        }
      }
      this.draw();
      if (!allDone) {
        this.raf = requestAnimationFrame(animate);
      } else {
        this.raf = 0;
        if (onDone) onDone();
      }
    };
    this.raf = requestAnimationFrame(animate);
  }
}

// ========================
// Presentation bootstrap
// ========================
function initPresentation() {
  const settings = DataService.getSettings();
  const data = DataService.getEntries();
  // Wire settings modal open/close
  const openBtn = document.getElementById("openSettings");
  const modal = document.getElementById("settingsModal");
  const closeBtn = document.getElementById("closeSettings");
  const showModal = () => {
    if (modal) {
      modal.classList.remove("hidden");
      modal.setAttribute("aria-hidden", "false");
    }
  };
  const hideModal = () => {
    if (modal) {
      modal.classList.add("hidden");
      modal.setAttribute("aria-hidden", "true");
    }
  };
  openBtn?.addEventListener("click", showModal);
  closeBtn?.addEventListener("click", hideModal);
  modal?.addEventListener("click", (e) => {
    if (e.target === modal) hideModal();
  });
  document.addEventListener("keydown", (e) => {
    if (e.key === "Escape") hideModal();
  });

  // Sidebar wiring
  const sidebar = document.getElementById("sidebar");
  const sidebarOverlay = document.getElementById("sidebarOverlay");
  const hotspot = document.getElementById("sidebarHotspot");
  const menuBtn = document.getElementById("menuBtn");
  const closeSidebarBtn = document.getElementById("closeSidebar");
  const versionLabel = document.getElementById("versionLabel");
  if (versionLabel) versionLabel.textContent = __APP_VERSION__;
  function openSidebar() {
    if (sidebar) {
      sidebar.classList.add("open");
      sidebar.setAttribute("aria-hidden", "false");
    }
    if (sidebarOverlay) {
      sidebarOverlay.classList.add("show");
    }
  }
  function closeSidebar() {
    if (sidebar) {
      sidebar.classList.remove("open");
      sidebar.setAttribute("aria-hidden", "true");
    }
    if (sidebarOverlay) {
      sidebarOverlay.classList.remove("show");
    }
  }
  hotspot?.addEventListener("click", openSidebar);
  menuBtn?.addEventListener("click", openSidebar);
  closeSidebarBtn?.addEventListener("click", closeSidebar);
  sidebarOverlay?.addEventListener("click", closeSidebar);
  document.addEventListener("keydown", (e) => {
    if (e.key.toLowerCase() === "m") {
      openSidebar();
    }
    if (e.key === "Escape") {
      closeSidebar();
    }
  });

  // Spinner vs Progress toggle (in sidebar)
  let toggle = document.getElementById("toggleProgress");
  const spinnerView =
    document.getElementById("spinnerView") ||
    (() => {
      const el = document.getElementById("reelsCanvas")?.parentElement;
      return el || null;
    })();
  let progressView = document.getElementById("progressView");
  if (!progressView && spinnerView && spinnerView.parentElement) {
    const pv = document.createElement("div");
    pv.id = "progressView";
    pv.className = "progress-view hidden";
    pv.setAttribute("aria-live", "polite");
    pv.innerHTML =
      '<div class="progress-bar" role="progressbar" aria-valuemin="0" aria-valuemax="100" aria-valuenow="0">\
      <div class="progress-fill" id="progressFill"></div>\
      <div class="progress-stripe"></div>\
      <div class="progress-label" id="progressLabel">0%</div>\
    </div>\
    <div class="progress-meta">\
      <span id="progressChecked">0 checked-in</span>\
      <span>•</span>\
      <span id="progressNot">0 not checked-in</span>\
      <span>•</span>\
      <span id="progressTotal">0 total</span>\
    </div>';
    spinnerView.parentElement.insertBefore(pv, spinnerView.nextSibling);
    progressView = pv;
  }
  const leaderView = document.getElementById("leaderView");
  const prizeCardsView = document.getElementById("prizeCardsView");
  const prizeDisplay = document.querySelector(".prize-display");
  function setMode(mode) {
    const m =
      mode ||
      DataService.getSettings().displayMode ||
      (DataService.getSettings().showProgress ? "progress" : "spinner");
    if (spinnerView) spinnerView.classList.toggle("hidden", m !== "spinner");
    if (progressView) progressView.classList.toggle("hidden", m !== "progress");
    if (leaderView) leaderView.classList.toggle("hidden", m !== "leader");
    if (prizeCardsView)
      prizeCardsView.classList.toggle("hidden", m !== "prizeCards");
    if (prizeDisplay)
      prizeDisplay.classList.toggle(
        "hidden",
        m === "leader" || m === "progress"
      );
    const s = DataService.getSettings();
    s.displayMode = m;
    s.showProgress = m === "progress";
    DataService.setSettings(s);
    if (m === "progress") {
      updateProgressFromGrid();
      renderPrizeCardsProgress();
    }
    if (m === "leader") renderLeaderBoard();
    if (m === "prizeCards") renderPrizeCards();
    if (toggle) toggle.checked = m === "progress";
  }
  toggle?.addEventListener("change", () =>
    setMode(toggle.checked ? "progress" : "spinner")
  );
  setMode(
    DataService.getSettings().displayMode ||
      (settings.showProgress ? "progress" : "spinner")
  );

  // Progress computation from stored grid + status mapping
  function updateProgressFromGrid() {
    try {
      const grid = DataService.getGrid();
      const s = DataService.getSettings();
      const sIdx = parseInt(s.status?.col != null ? s.status.col : NaN, 10);
      const sVal = s.status?.value || "";
      if (!Array.isArray(grid) || grid.length === 0) {
        renderProgress(0, 0, 0);
        return;
      }
      const start = grid[0] && grid[0].some((h) => h) ? 1 : 0;
      let total = 0,
        checked = 0;
      for (let i = start; i < grid.length; i++) {
        const row = grid[i] || [];
        total++;
        if (!Number.isNaN(sIdx) && sVal) {
          const v = (row[sIdx] ?? "").toString().trim();
          if (v === sVal) checked++;
        }
      }
      renderProgress(total, checked, Math.max(0, total - checked));
    } catch {
      renderProgress(0, 0, 0);
    }
  }
  function renderProgress(total, checked, notC) {
    const fill = document.getElementById("progressFill");
    const label = document.getElementById("progressLabel");
    const elC = document.getElementById("progressChecked");
    const elN = document.getElementById("progressNot");
    const elT = document.getElementById("progressTotal");
    const pct = total > 0 ? Math.round((checked * 1000) / total) / 10 : 0;
    if (fill) {
      fill.style.width = `${pct}%`;
      fill.parentElement?.setAttribute(
        "aria-valuenow",
        String(Math.round(pct))
      );
    }
    if (label) label.textContent = `${pct}%`;
    if (elC) elC.textContent = `${checked} checked-in`;
    if (elN) elN.textContent = `${notC} not checked-in`;
    if (elT) elT.textContent = `${total} total`;
  }
  if (settings.showProgress) updateProgressFromGrid();

  // Leader Board rendering (group by wins)
  function renderLeaderBoard() {
    const list = document.getElementById("boardList");
    if (!list) return;
    list.innerHTML = "";
    const hist = DataService.getHistory();
    const map = new Map(); // card -> {name, wins, last}
    for (const h of hist) {
      const key = (h.card || "").toString();
      const prev = map.get(key) || {
        name: h.name || "",
        wins: 0,
        last: 0,
        lastPrize: "",
      };
      prev.name = h.name || prev.name;

      // Only count actual prize wins, not "no prize claimed"
      if (h.prize && h.prize !== "no prize claimed") {
        prev.wins += 1;
      }

      const t = Date.parse(h.ts || "") || 0;
      if (t > prev.last) {
        prev.last = t;
        // Show the most recent actual prize, not "no prize claimed"
        if (h.prize && h.prize !== "no prize claimed") {
          prev.lastPrize = h.prize;
        }
      }
      map.set(key, prev);
    }
    const rows = Array.from(map.entries()).map(([card, info]) => ({
      card,
      ...info,
    }));
    rows.sort(
      (a, b) =>
        b.wins - a.wins || b.last - a.last || a.name.localeCompare(b.name)
    );
    let rank = 1;
    const fmt = new Intl.DateTimeFormat(undefined, {
      dateStyle: "medium",
      timeStyle: "short",
    });
    if (rows.length === 0) {
      const empty = document.createElement("div");
      empty.className = "board-card";
      empty.innerHTML = `
        <div class="board-top">
          <div class="rank-badge">–</div>
          <div class="board-name">No winners yet</div>
          
          <span class="wins-pill">0 wins</span>
        </div>
        
        <div class="board-meta">
          <span>Last: —</span>
          <span class="prize-pill">Prize: —</span>
        </div>`;
      list.appendChild(empty);
      return;
    }
    for (const r of rows) {
      const card = document.createElement("div");
      card.className = "board-card";
      const last = r.last ? fmt.format(new Date(r.last)) : "—";
      card.innerHTML = `
        <div class="board-top">
          <div class="rank-badge">${rank++}</div>
          <div class="board-name">${r.name || ""}</div>
          
          <div class="wins-pill">${r.wins} win${r.wins === 1 ? "" : "s"}</div>
        </div>
        
        <div class="board-meta">
          <span>${last}</span>
          <span class="prize-pill">Prize: ${r.lastPrize || "—"}</span>
        </div>`;
      list.appendChild(card);
    }
  }

  function renderPrizeCards() {
    const list = document.getElementById("prizeCardsList");
    if (!list) return;

    list.innerHTML = "";
    const prizes = DataService.getPrizes();

    if (!prizes || prizes.length === 0) {
      const empty = document.createElement("div");
      empty.className = "prize-card-display";
      empty.innerHTML = `
        <div class="prize-image-placeholder">🎁</div>
        <div class="prize-card-title">No prizes configured</div>
        <div class="prize-card-status available">Setup Required</div>`;
      list.appendChild(empty);
      return;
    }

    // Find claimed prizes by checking history for confirmed wins
    const hist = DataService.getHistory();
    const confirmedPrizeIndices = new Set(
      hist
        .filter(
          (h) =>
            h.prize &&
            h.prize !== "—" &&
            h.prize.trim() !== "" &&
            h.prize !== "no prize claimed"
        )
        .map((h) => {
          // Find which prize index this winner received
          for (let i = 0; i < prizes.length; i++) {
            if (prizes[i] === h.prize) return i;
          }
          return -1;
        })
        .filter((idx) => idx >= 0)
    );

    for (let i = 0; i < prizes.length; i++) {
      const prize = prizes[i];
      const isClaimed = confirmedPrizeIndices.has(i);

      const card = document.createElement("div");
      card.className = `prize-card-display ${isClaimed ? "claimed" : ""}`;

      card.innerHTML = `
        <div class="prize-image-placeholder">🖼️</div>
        <div class="prize-card-title">${prize}</div>
        <div class="prize-card-status ${isClaimed ? "claimed" : "available"}">
          ${isClaimed ? "Claimed" : "Available"}
        </div>`;

      list.appendChild(card);
    }
  }

  function renderPrizeCardsProgress() {
    const list = document.getElementById("prizeCardsListProgress");
    if (!list) return;

    list.innerHTML = "";
    const prizes = DataService.getPrizes();

    if (!prizes || prizes.length === 0) {
      const empty = document.createElement("div");
      empty.className = "prize-card-display";
      empty.innerHTML = `
        <div class="prize-image-placeholder">🎁</div>
        <div class="prize-card-title">No prizes configured</div>
        <div class="prize-card-status available">Setup Required</div>`;
      list.appendChild(empty);
      return;
    }

    // Find claimed prizes by checking history for confirmed wins
    const hist = DataService.getHistory();
    const confirmedPrizeIndices = new Set(
      hist
        .filter(
          (h) =>
            h.prize &&
            h.prize !== "—" &&
            h.prize.trim() !== "" &&
            h.prize !== "no prize claimed"
        )
        .map((h) => {
          // Find which prize index this winner received
          for (let i = 0; i < prizes.length; i++) {
            if (prizes[i] === h.prize) return i;
          }
          return -1;
        })
        .filter((idx) => idx >= 0)
    );

    for (let i = 0; i < prizes.length; i++) {
      const prize = prizes[i];
      const isClaimed = confirmedPrizeIndices.has(i);

      const card = document.createElement("div");
      card.className = `prize-card-display ${isClaimed ? "claimed" : ""}`;

      card.innerHTML = `
        <div class="prize-image-placeholder">🖼️</div>
        <div class="prize-card-title">${prize}</div>
        <div class="prize-card-status ${isClaimed ? "claimed" : "available"}">
          ${isClaimed ? "Claimed" : "Available"}
        </div>`;

      list.appendChild(card);
    }
  }

  // Sidebar navigation actions
  const navLive = document.getElementById("navLive");
  const navCard = document.getElementById("navCardDraw");
  const navBoard = document.getElementById("navLeaderboard");
  const navPrizeShowcase = document.getElementById("navPrizeShowcase");
  const navRaffle = document.getElementById("navRaffle");
  const navSettings = document.getElementById("navSettings");
  const navGoogle = document.getElementById("navGoogle");
  const navPrizes = document.getElementById("navPrizes");
  const navAttendees = document.getElementById("navAttendees");
  const navHistory = document.getElementById("navHistory");
  const navDiagnostics = document.getElementById("navDiagnostics");
  const navHelp = document.getElementById("navHelp");
  const navEula = document.getElementById("navEula");
  const helpModal = document.getElementById("helpModal");
  const closeHelp = document.getElementById("closeHelp");
  const eulaModal = document.getElementById("eulaModal");
  const closeEula = document.getElementById("closeEula");
  const googleModal = document.getElementById("googleModal");
  const closeGoogle = document.getElementById("closeGoogle");
  const attendeesModal = document.getElementById("attendeesModal");
  const closeAttendees = document.getElementById("closeAttendees");
  const historyModal = document.getElementById("historyModal");
  const closeHistory = document.getElementById("closeHistory");
  const diagnosticsModal = document.getElementById("diagnosticsModal");
  const closeDiagnostics = document.getElementById("closeDiagnostics");
  const startBtn = document.getElementById("startSpin");
  const resetBtn = document.getElementById("resetSpin");
  const winnerName = document.getElementById("winnerName");
  const winnerBadge = document.getElementById("winnerBadge");
  const confirmBtn = document.getElementById("confirmWin");
  const nextPrizeEl = document.getElementById("nextPrize");
  const remainingCountEl = document.getElementById("remainingCount");
  const totalCountEl = document.getElementById("totalCount");
  const prizeProgressFillEl = document.getElementById("prizeProgressFill");
  const prizeCard = document.getElementById("prizeCard");
  const winnerNameCard = document.getElementById("winnerNameCard");
  const winnerCardDisplay = document.getElementById("winnerCardDisplay");
  function openModal(el) {
    if (el) {
      el.classList.remove("hidden");
      el.setAttribute("aria-hidden", "false");
      document.body.classList.add("modal-open");
    }
  }
  function closeModal(el) {
    if (el) {
      el.classList.add("hidden");
      el.setAttribute("aria-hidden", "true");
      // Check if any other modals are still open
      const openModals = document.querySelectorAll(".modal:not(.hidden)");
      if (openModals.length === 0) {
        document.body.classList.remove("modal-open");
      }
    }
  }
  navLive?.addEventListener("click", () => {
    setMode("progress");
    closeSidebar();
  });
  navCard?.addEventListener("click", () => {
    setMode("spinner");
    closeSidebar();
  });
  navBoard?.addEventListener("click", () => {
    setMode("leader");
    closeSidebar();
  });
  navPrizeShowcase?.addEventListener("click", () => {
    setMode("prizeCards");
    closeSidebar();
  });
  navRaffle?.addEventListener("click", () => {
    logDiag("Raffle Ticket Draw is coming soon.");
    closeSidebar();
  });
  navSettings?.addEventListener("click", () => {
    showModal();
    closeSidebar();
  });
  navHelp?.addEventListener("click", () => {
    openModal(helpModal);
    closeSidebar();
  });
  navEula?.addEventListener("click", () => {
    openModal(eulaModal);
    closeSidebar();
  });
  navGoogle?.addEventListener("click", () => {
    openModal(googleModal);
    closeSidebar();

    // Initialize tab functionality (only once)
    if (!googleModal.dataset.tabsInitialized) {
      initDataPreviewTabs();
      googleModal.dataset.tabsInitialized = "true";
    }

    // Update the connection status and UI
    renderGoogleConn();
  });
  closeGoogle?.addEventListener("click", () => closeModal(googleModal));

  const prizesModal = document.getElementById("prizesModal");
  const closePrizes = document.getElementById("closePrizes");

  closeHelp?.addEventListener("click", () => closeModal(helpModal));
  closeEula?.addEventListener("click", () => closeModal(eulaModal));
  navPrizes?.addEventListener("click", () => {
    openModal(prizesModal);
    closeSidebar();
    renderPrizes();
  });
  closePrizes?.addEventListener("click", () => closeModal(prizesModal));

  navAttendees?.addEventListener("click", () => {
    openModal(attendeesModal);
    closeSidebar();
    showSettings();
  });
  closeAttendees?.addEventListener("click", () => closeModal(attendeesModal));

  navHistory?.addEventListener("click", () => {
    openModal(historyModal);
    closeSidebar();
    renderHistory();
  });
  closeHistory?.addEventListener("click", () => closeModal(historyModal));

  navDiagnostics?.addEventListener("click", () => {
    openModal(diagnosticsModal);
    closeSidebar();
  });
  closeDiagnostics?.addEventListener("click", () =>
    closeModal(diagnosticsModal)
  );

  // Next prize + pending flow
  function showNextPrize() {
    const prizes = DataService.getPrizes();

    // If no prizes loaded, show empty state
    if (!prizes || prizes.length === 0) {
      if (nextPrizeEl) nextPrizeEl.textContent = "No available prizes";
      updateRemainingPrizesCounter();
      return;
    }

    // Find the next available prize by checking history for confirmed wins
    const hist = DataService.getHistory();
    const confirmedPrizeIndices = new Set(
      hist
        .filter(
          (h) =>
            h.prize &&
            h.prize !== "—" &&
            h.prize.trim() !== "" &&
            h.prize !== "no prize claimed"
        )
        .map((h) => {
          // Find which prize index this winner received
          for (let i = 0; i < prizes.length; i++) {
            if (prizes[i] === h.prize) return i;
          }
          return -1;
        })
        .filter((idx) => idx >= 0)
    );

    // Find the lowest unclaimed prize index
    let nextAvailableIndex = 0;
    while (
      nextAvailableIndex < prizes.length &&
      confirmedPrizeIndices.has(nextAvailableIndex)
    ) {
      nextAvailableIndex++;
    }

    const prize =
      nextAvailableIndex < prizes.length
        ? prizes[nextAvailableIndex]
        : "All prizes claimed";

    console.log("showNextPrize:", {
      prizes: prizes,
      prizeCount: prizes.length,
      confirmedIndices: Array.from(confirmedPrizeIndices),
      nextAvailableIndex: nextAvailableIndex,
      currentPrize: prize,
      nextPrizeEl: !!nextPrizeEl,
    });

    if (nextPrizeEl) nextPrizeEl.textContent = prize;
    updateRemainingPrizesCounter();
  }

  function updateRemainingPrizesCounter() {
    const prizes = DataService.getPrizes();

    if (!prizes || prizes.length === 0) {
      if (remainingCountEl) remainingCountEl.textContent = "0";
      if (totalCountEl) totalCountEl.textContent = "0";
      if (prizeProgressFillEl) prizeProgressFillEl.style.width = "0%";
      return;
    }

    // Count confirmed wins by checking history
    const hist = DataService.getHistory();
    const confirmedPrizeIndices = new Set(
      hist
        .filter(
          (h) =>
            h.prize &&
            h.prize !== "—" &&
            h.prize.trim() !== "" &&
            h.prize !== "no prize claimed"
        )
        .map((h) => {
          // Find which prize index this winner received
          for (let i = 0; i < prizes.length; i++) {
            if (prizes[i] === h.prize) return i;
          }
          return -1;
        })
        .filter((idx) => idx >= 0)
    );

    const totalPrizes = prizes.length;
    const claimedCount = confirmedPrizeIndices.size;
    const remaining = Math.max(0, totalPrizes - claimedCount);
    const progressPercent =
      totalPrizes > 0 ? (claimedCount / totalPrizes) * 100 : 0;

    // Update counter with animation
    if (remainingCountEl) {
      const currentValue = parseInt(remainingCountEl.textContent) || 0;
      if (currentValue !== remaining) {
        remainingCountEl.classList.add("updating");
        setTimeout(() => {
          remainingCountEl.textContent = remaining.toString();
          remainingCountEl.classList.remove("updating");

          // Special animation when all prizes are awarded
          if (remaining === 0 && totalPrizes > 0) {
            remainingCountEl.style.color = "var(--good)";
            remainingCountEl.style.textShadow =
              "0 0 25px rgba(124, 252, 0, 0.8)";
          } else {
            remainingCountEl.style.color = "";
            remainingCountEl.style.textShadow = "";
          }
        }, 100);
      }
    }

    if (totalCountEl) totalCountEl.textContent = totalPrizes.toString();
    if (prizeProgressFillEl)
      prizeProgressFillEl.style.width = `${progressPercent}%`;
  }

  // Card flip animation functions
  function flipToWinner(winnerName, winnerCard) {
    if (!prizeCard || !winnerNameCard) return;

    // Update winner info
    winnerNameCard.textContent = winnerName || "Unknown Winner";
    if (winnerCardDisplay) {
      winnerCardDisplay.textContent = `Card #${winnerCard || "---"}`;
    }

    // Add winner mode for expanded layout and effects
    // NO FLIPPING - just reveal the winner card on top and add border effect
    prizeCard.classList.add("winner-mode");
    prizeCard.classList.add("winner-reveal");
    prizeCard.classList.add("flipped"); // Show winner side immediately
  }

  function flipToPrizes() {
    if (!prizeCard) return;

    // Remove winner effects immediately - no flip animation needed
    prizeCard.classList.remove("flipped");
    prizeCard.classList.remove("winner-reveal");
    prizeCard.classList.remove("winner-mode");

    // Add rolling animation to numbers
    if (remainingCountEl) remainingCountEl.classList.add("rolling");
    if (totalCountEl) totalCountEl.classList.add("rolling");

    // Remove rolling animation after it completes
    setTimeout(() => {
      if (remainingCountEl) remainingCountEl.classList.remove("rolling");
      if (totalCountEl) totalCountEl.classList.remove("rolling");
    }, 1000);

    // Update the prize display
    showNextPrize();
    updateRemainingPrizesCounter();
  }

  showNextPrize();
  let pending = null;
  function setPending(p) {
    pending = p;
    if (confirmBtn) {
      confirmBtn.classList.toggle("hidden", !p);
    }
    if (startBtn) {
      startBtn.classList.toggle("hidden", !!p);
    }
  }
  function resetSpinView() {
    winnerName?.classList.remove("show");
    if (winnerName) winnerName.textContent = "";
    if (winnerBadge) winnerBadge.textContent = "Awaiting spin…";
    try {
      if (od && od.reels) {
        od.reels.forEach((r) => {
          r.offset = 0;
        });
        od.draw();
      }
    } catch {}
  }

  resetBtn?.addEventListener("click", () => {
    // Log the unclaimed prize if there was a pending winner
    if (pending) {
      const hist = DataService.getHistory();
      hist.push({
        ts: new Date().toISOString(),
        card: pending.entry.card,
        name: pending.entry.name,
        prize: "no prize claimed",
      });
      DataService.setHistory(hist);
      try {
        window.dispatchEvent(new Event("historyUpdated"));
      } catch {}
    }

    setPending(null);
    resetSpinView();

    // Flip card back to prizes
    flipToPrizes();

    if (startBtn) {
      startBtn.classList.remove("hidden");
      startBtn.disabled = false;
    }
    if (confirmBtn) {
      confirmBtn.classList.add("hidden");
    }
  });
  confirmBtn?.addEventListener("click", () => {
    if (!pending) return;
    const s = DataService.getSettings();
    const prizes = DataService.getPrizes();

    // Find the next available prize by checking history
    const hist = DataService.getHistory();
    let nextAvailablePrize = "";

    if (prizes && prizes.length > 0) {
      const confirmedPrizeIndices = new Set(
        hist
          .filter(
            (h) =>
              h.prize &&
              h.prize !== "—" &&
              h.prize.trim() !== "" &&
              h.prize !== "no prize claimed"
          )
          .map((h) => {
            // Find which prize index this winner received
            for (let i = 0; i < prizes.length; i++) {
              if (prizes[i] === h.prize) return i;
            }
            return -1;
          })
          .filter((idx) => idx >= 0)
      );

      // Find the lowest unclaimed prize index
      let nextAvailableIndex = 0;
      while (
        nextAvailableIndex < prizes.length &&
        confirmedPrizeIndices.has(nextAvailableIndex)
      ) {
        nextAvailableIndex++;
      }

      nextAvailablePrize =
        nextAvailableIndex < prizes.length ? prizes[nextAvailableIndex] : "";
    }

    hist.push({
      ts: new Date().toISOString(),
      card: pending.entry.card,
      name: pending.entry.name,
      prize: nextAvailablePrize,
    });
    DataService.setHistory(hist);
    try {
      window.dispatchEvent(new Event("historyUpdated"));
    } catch {}
    Confetti.burst(s.confettiLevel || "medium");

    // Flip card back to prizes after confirmation
    setTimeout(() => {
      flipToPrizes();
    }, 1000); // Short delay to let confetti play

    if (s.noRepeat) {
      const used = new Set(hist.map((h) => h.card));
      const remaining = DataService.getEntries().filter(
        (r) => !used.has(r.card)
      ).length;
      const countBadge = document.getElementById("countBadge");
      if (countBadge)
        countBadge.textContent = `${remaining} entries (eligible)`;
    }
    setPending(null);
  });

  // Start spin now sets pending winner (confirmation required)
  startBtn?.addEventListener("click", () => {
    const s = DataService.getSettings();
    function chooseEntry() {
      let rows = DataService.getEntries();
      if (rows.length === 0) return null;
      if (s.noRepeat) {
        const used = new Set(DataService.getHistory().map((h) => h.card));
        rows = rows.filter((r) => !used.has(r.card));
        if (rows.length === 0) return null;
      }
      return rows[pickRandomIndex(rows.length, s)];
    }
    const entry = chooseEntry();
    if (!entry) {
      alert("No available entries. Load data in Settings or clear history.");
      return;
    }
    resetSpinView();
    if (winnerBadge) winnerBadge.textContent = "Spinning…";
    const targets = entry.card.trim().toUpperCase();
    const tickFn = (i) => {
      if (s.sound && s.tick) AudioFX.beep(720 + i * 25, 0.03, "square", 0.02);
    };
    od.spinTo(targets, s, tickFn, () => {
      if (winnerBadge) winnerBadge.textContent = entry.card;
      if (winnerName) {
        winnerName.textContent = entry.name;
        winnerName.classList.add("show");
      }

      // Flip card to show winner
      setTimeout(() => {
        flipToWinner(entry.name, entry.card);
      }, 500); // Small delay for dramatic effect

      // Fire confetti + fanfare immediately for showmanship
      if (s.sound && s.fanfare) AudioFX.fanfare({ sample: !!s.fanfareSample });
      Confetti.burst(s.confettiLevel || "medium");
      setPending({ entry });
    });
  });

  // Defensive: delegate clicks in case elements are replaced or not bound
  document.addEventListener("click", (ev) => {
    const t = ev.target.closest?.(
      "#navLive, #navCardDraw, #navLeaderboard, #navSettings, #navHelp, #navEula, #navPrizes, #navAttendees, #navHistory, #navDiagnostics"
    );
    if (!t) return;
    switch (t.id) {
      case "navLive":
        setMode("progress");
        closeSidebar();
        break;
      case "navCardDraw":
        setMode("spinner");
        closeSidebar();
        break;
      case "navLeaderboard":
        setMode("leader");
        closeSidebar();
        break;
      case "navPrizes":
        openModal(prizesModal);
        renderPrizes();
        closeSidebar();
        break;
      case "navSettings":
        showModal();
        closeSidebar();
        break;
      case "navHelp":
        openModal(helpModal);
        closeSidebar();
        break;
      case "navEula":
        openModal(eulaModal);
        closeSidebar();
        break;
      case "navAttendees":
        openModal(attendeesModal);
        showSettings();
        closeSidebar();
        break;
      case "navHistory":
        openModal(historyModal);
        renderHistory();
        closeSidebar();
        break;
      case "navDiagnostics":
        openModal(diagnosticsModal);
        closeSidebar();
        break;
    }
  });

  // Prizes modal logic
  const prizeName = document.getElementById("prizeName");
  const prizeCount = document.getElementById("prizeCount");
  const addPrize = document.getElementById("addPrize");
  const prizeTbody = document.getElementById("prizeTbody");
  const clearPrizes = document.getElementById("clearPrizes");
  const prizesSheetUrl = document.getElementById("prizesSheetUrl");
  const loadPrizesSheet = document.getElementById("loadPrizesSheet");
  const loadPrizesConnected = document.getElementById("loadPrizesConnected");
  const gsWorkbookUrl = document.getElementById("gsWorkbookUrl");
  const connectWorkbook = document.getElementById("connectWorkbook");
  const addManualSheet = document.getElementById("addManualSheet");
  const gsConnStatus = document.getElementById("gsConnStatus");
  const attSheetSelect = document.getElementById("attSheetSelect");
  const prizeSheetSelect = document.getElementById("prizeSheetSelect");
  const attIgnoreHeaderCk = document.getElementById("attIgnoreHeader");
  const prizeIgnoreHeaderCk = document.getElementById("prizeIgnoreHeader");
  const countBadge = document.getElementById("countBadge");

  function renderGoogleConn() {
    const s = DataService.getSettings();
    if (gsConnStatus) {
      gsConnStatus.textContent = s.gsWorkbookId ? "Connected" : "Not connected";
      gsConnStatus.className = "badge " + (s.gsWorkbookId ? "ok" : "");
    }
    if (gsWorkbookUrl)
      gsWorkbookUrl.value = s.gsWorkbookId
        ? `https://docs.google.com/spreadsheets/d/${s.gsWorkbookId}/edit`
        : "";

    // Show/hide sheet selection card
    const sheetSelectionCard = document.getElementById("sheetSelectionCard");
    if (sheetSelectionCard) {
      sheetSelectionCard.style.display = s.gsWorkbookId ? "block" : "none";
    }

    function fill(select, items) {
      if (!select) return;
      select.innerHTML = "";
      const opt = document.createElement("option");
      opt.value = "";
      opt.textContent = "(Select a sheet)";
      select.appendChild(opt);
      (items || []).forEach((x) => {
        const o = document.createElement("option");
        o.value = String(x.gid);
        o.textContent = x.title;
        select.appendChild(o);
      });
    }
    fill(attSheetSelect, s.gsSheets);
    fill(prizeSheetSelect, s.gsSheets);
    if (attSheetSelect && s.attSheetGid)
      attSheetSelect.value = String(s.attSheetGid);
    if (prizeSheetSelect && s.prizeSheetGid)
      prizeSheetSelect.value = String(s.prizeSheetGid);
    if (attIgnoreHeaderCk) attIgnoreHeaderCk.checked = !!s.attIgnoreHeader;
    if (prizeIgnoreHeaderCk)
      prizeIgnoreHeaderCk.checked = !!s.prizeIgnoreHeader;
  }

  // Tab switching functionality
  function initDataPreviewTabs() {
    const tabBtns = document.querySelectorAll("[data-tab]");
    const tabContents = document.querySelectorAll(".tab-content");

    tabBtns.forEach((btn) => {
      btn.addEventListener("click", () => {
        const targetTab = btn.dataset.tab;

        // Update button states
        tabBtns.forEach((b) => b.classList.remove("active"));
        btn.classList.add("active");

        // Update content visibility
        tabContents.forEach((content) => {
          content.classList.remove("active");
          if (content.id === `${targetTab}Preview`) {
            content.classList.add("active");
          }
        });
      });
    });
  }

  // Data preview functions
  function showDataPreview() {
    const previewCard = document.getElementById("dataPreviewCard");
    if (previewCard) {
      previewCard.style.display = "block";
    }
  }

  function switchToTab(tabName) {
    const tabBtns = document.querySelectorAll("[data-tab]");
    const tabContents = document.querySelectorAll(".tab-content");

    // Update button states
    tabBtns.forEach((btn) => {
      btn.classList.remove("active");
      if (btn.dataset.tab === tabName) {
        btn.classList.add("active");
      }
    });

    // Update content visibility
    tabContents.forEach((content) => {
      content.classList.remove("active");
      if (content.id === `${tabName}Preview`) {
        content.classList.add("active");
      }
    });
  }

  function renderAttendancePreview(data, hasHeader = false) {
    const head = document.getElementById("attendancePreviewHead");
    const body = document.getElementById("attendancePreviewBody");
    const info = document.getElementById("attendanceDataInfo");

    if (!head || !body || !info) return;

    head.innerHTML = "";
    body.innerHTML = "";

    if (!data || data.length === 0) {
      info.textContent = "No data loaded";
      info.className = "badge";
      return;
    }

    const startRow = hasHeader ? 1 : 0;
    const dataRows = data.slice(startRow);
    const previewRows = dataRows.slice(0, 5); // Show first 5 data rows

    // Create header
    const headerRow = document.createElement("tr");
    if (hasHeader && data[0]) {
      data[0].forEach((header, i) => {
        const th = document.createElement("th");
        th.textContent = header || `Column ${i + 1}`;
        headerRow.appendChild(th);
      });
    } else {
      // Generate column headers
      const maxCols = Math.max(...data.map((row) => row.length));
      for (let i = 0; i < maxCols; i++) {
        const th = document.createElement("th");
        th.textContent = `Column ${i + 1}`;
        headerRow.appendChild(th);
      }
    }
    head.appendChild(headerRow);

    // Create preview rows
    previewRows.forEach((row) => {
      const tr = document.createElement("tr");
      const maxCols = headerRow.children.length;
      for (let i = 0; i < maxCols; i++) {
        const td = document.createElement("td");
        td.textContent = row[i] || "";
        tr.appendChild(td);
      }
      body.appendChild(tr);
    });

    // Update info badge
    info.textContent = `${dataRows.length} records (showing first ${Math.min(
      5,
      dataRows.length
    )})`;
    info.className = "badge ok";

    // Switch to attendance tab and show preview
    switchToTab("attendance");
    showDataPreview();
  }

  function renderPrizesPreview(data, hasHeader = false) {
    const head = document.getElementById("prizesPreviewHead");
    const body = document.getElementById("prizesPreviewBody");
    const info = document.getElementById("prizesDataInfo");

    if (!head || !body || !info) {
      console.warn("Prizes preview elements not found:", {
        head: !!head,
        body: !!body,
        info: !!info,
      });
      return;
    }

    head.innerHTML = "";
    body.innerHTML = "";

    if (!data || data.length === 0) {
      info.textContent = "No data loaded";
      info.className = "badge";
      return;
    }

    const startRow = hasHeader ? 1 : 0;
    const dataRows = data.slice(startRow);
    const previewRows = dataRows.slice(0, 5); // Show first 5 data rows

    // Create header
    const headerRow = document.createElement("tr");
    const th = document.createElement("th");
    th.textContent = hasHeader && data[0] && data[0][0] ? data[0][0] : "Prize";
    headerRow.appendChild(th);
    head.appendChild(headerRow);

    // Create preview rows
    previewRows.forEach((row) => {
      const tr = document.createElement("tr");
      const td = document.createElement("td");
      td.textContent = row[0] || "";
      tr.appendChild(td);
      body.appendChild(tr);
    });

    // Update info badge
    const prizeCount = dataRows.reduce(
      (count, row) => count + (row[0] ? 1 : 0),
      0
    );
    info.textContent = `${prizeCount} prizes (showing first ${Math.min(
      5,
      prizeCount
    )})`;
    info.className = "badge ok";

    // Switch to prizes tab and show preview
    switchToTab("prizes");
    showDataPreview();
  }

  async function connectToWorkbook(url) {
    try {
      const u = new URL(url.trim());
      const m = u.pathname.match(/\/spreadsheets\/d\/([^/]+)/);
      const id = m ? m[1] : "";
      if (!id) throw new Error("Sheet ID not found");

      // Try multiple methods to get sheet metadata
      let sheets = [];

      // Method 1: Try the metadata endpoint (may fail due to CORS)
      try {
        const metaUrl = `https://docs.google.com/spreadsheets/d/${id}/gviz/sheetmetadata?tqx=out:json`;
        const r = await fetch(metaUrl, { mode: "cors" });
        const t = await r.text();
        const json = JSON.parse(t.substring(t.indexOf("{"))); // strip JSONP
        if (Array.isArray(json.sheets)) {
          sheets = json.sheets
            .map((s) => ({
              gid: s.properties?.sheetId,
              title: s.properties?.title,
            }))
            .filter((x) => x.gid != null);
        }
      } catch (metaError) {
        console.warn("Metadata endpoint failed:", metaError.message);

        // Method 2: Try to discover sheets by attempting to access common sheet IDs
        // This is a fallback when metadata isn't available
        const commonGids = [0, 1, 2, 3, 4, 5]; // Sheet1 is usually gid=0, Sheet2=1, etc.
        const discoveredSheets = [];

        for (const gid of commonGids) {
          try {
            const testUrl = `https://docs.google.com/spreadsheets/d/${id}/export?format=csv&gid=${gid}`;
            const testResponse = await fetch(testUrl, { mode: "cors" });
            if (testResponse.ok) {
              // Try to get the sheet name from the CSV header or use a default
              const csvText = await testResponse.text();
              const lines = csvText.split("\n");
              let sheetTitle = `Sheet${gid === 0 ? "1" : gid + 1}`; // Default naming

              // If the sheet has data, use a more descriptive name
              if (lines.length > 0 && lines[0].trim()) {
                const firstRow = lines[0].split(",");
                if (firstRow.length > 0 && firstRow[0].trim()) {
                  // Use first column header as hint for sheet name
                  const hint = firstRow[0]
                    .replace(/[^a-zA-Z0-9\s]/g, "")
                    .trim();
                  if (hint) {
                    sheetTitle = `${hint.substring(0, 20)}... (gid:${gid})`;
                  }
                }
              }

              discoveredSheets.push({
                gid: gid,
                title: sheetTitle,
              });
            }
          } catch (sheetError) {
            // Sheet doesn't exist or isn't accessible, continue
            continue;
          }
        }

        if (discoveredSheets.length > 0) {
          sheets = discoveredSheets;
          console.log(
            `Discovered ${sheets.length} accessible sheets via fallback method`
          );
        } else {
          // Method 3: Manual sheet entry fallback
          const manualEntry = confirm(
            "Unable to automatically discover sheets. Would you like to manually specify sheet details?\n\n" +
              "You can find sheet GIDs by looking at the URL when you click on different sheet tabs.\n" +
              "Example: #gid=123456789"
          );

          if (manualEntry) {
            const gidInput = prompt(
              "Enter the GID (numeric ID) of your first sheet (usually 0 for Sheet1):"
            );
            const titleInput = prompt("Enter a name for this sheet:");

            if (gidInput !== null && titleInput !== null) {
              sheets = [
                {
                  gid: parseInt(gidInput) || 0,
                  title: titleInput || "Sheet1",
                },
              ];
            }
          }
        }
      }

      const s = DataService.getSettings();
      s.gsWorkbookId = id;
      s.gsSheets = sheets;
      DataService.setSettings(s);
      renderGoogleConn();

      if (sheets.length > 0) {
        alert(
          `Connected to workbook! Found ${sheets.length} accessible sheet(s).`
        );
      } else {
        alert(
          "Connected to workbook, but no sheets were discoverable. You may need to manually configure sheet GIDs."
        );
      }
    } catch (err) {
      alert("Connect failed: " + (err?.message || err));
    }
  }

  // Helper function to manually add sheets when auto-discovery fails
  function handleAddManualSheet() {
    const s = DataService.getSettings();
    if (!s.gsWorkbookId) {
      alert("Please connect to a workbook first.");
      return;
    }

    const gidInput = prompt(
      "Enter the GID (numeric ID) of the sheet.\n\n" +
        "To find the GID:\n" +
        "1. Open your Google Sheet\n" +
        "2. Click on the sheet tab you want\n" +
        "3. Look at the URL - the number after '#gid=' is your GID\n" +
        "4. If no gid is shown, the GID is 0\n\n" +
        "Enter GID:"
    );

    if (gidInput === null) return;

    const titleInput = prompt("Enter a display name for this sheet:");
    if (titleInput === null) return;

    const gid = parseInt(gidInput) || 0;
    const title = titleInput.trim() || `Sheet (gid:${gid})`;

    // Test if the sheet is accessible
    const testUrl = `https://docs.google.com/spreadsheets/d/${s.gsWorkbookId}/export?format=csv&gid=${gid}`;

    fetch(testUrl, { mode: "cors" })
      .then((response) => {
        if (!response.ok) {
          throw new Error(`Sheet not accessible (HTTP ${response.status})`);
        }

        // Add the sheet to the list
        const existingSheets = s.gsSheets || [];
        const newSheet = { gid, title };

        // Check if this GID already exists
        const existingIndex = existingSheets.findIndex(
          (sheet) => sheet.gid === gid
        );
        if (existingIndex >= 0) {
          existingSheets[existingIndex] = newSheet; // Update existing
        } else {
          existingSheets.push(newSheet); // Add new
        }

        s.gsSheets = existingSheets;
        DataService.setSettings(s);
        renderGoogleConn();
        alert(`Sheet "${title}" added successfully!`);
      })
      .catch((error) => {
        alert(
          `Failed to access sheet: ${error.message}\n\nPlease check:\n- The GID is correct\n- The sheet exists\n- The workbook is publicly accessible`
        );
      });
  }

  connectWorkbook?.addEventListener("click", () => {
    const url = gsWorkbookUrl?.value || "";
    if (!url.trim()) {
      alert("Enter workbook URL");
      return;
    }
    connectToWorkbook(url);
  });

  addManualSheet?.addEventListener("click", handleAddManualSheet);

  function selectedSheetGid(select) {
    return select && select.value ? select.value : "";
  }
  attSheetSelect?.addEventListener("change", () => {
    const s = DataService.getSettings();
    s.attSheetGid = selectedSheetGid(attSheetSelect);
    DataService.setSettings(s);
  });
  prizeSheetSelect?.addEventListener("change", () => {
    const s = DataService.getSettings();
    s.prizeSheetGid = selectedSheetGid(prizeSheetSelect);
    DataService.setSettings(s);
  });
  attIgnoreHeaderCk?.addEventListener("change", () => {
    const s = DataService.getSettings();
    s.attIgnoreHeader = !!attIgnoreHeaderCk.checked;
    DataService.setSettings(s);
  });
  prizeIgnoreHeaderCk?.addEventListener("change", () => {
    const s = DataService.getSettings();
    s.prizeIgnoreHeader = !!prizeIgnoreHeaderCk.checked;
    DataService.setSettings(s);
  });

  loadGsheetConnected?.addEventListener("click", async () => {
    const s = DataService.getSettings();
    if (!s.gsWorkbookId || !s.attSheetGid) {
      alert("Connect to Google and select an attendance sheet.");
      return;
    }
    const csv = `https://docs.google.com/spreadsheets/d/${
      s.gsWorkbookId
    }/export?format=csv&gid=${encodeURIComponent(s.attSheetGid)}`;
    try {
      const r = await fetch(csv, { mode: "cors" });
      if (!r.ok) throw new Error(`HTTP ${r.status}`);
      let text = await r.text();

      // Parse the full CSV for preview
      const fullGrid = parseCSV(text);
      const hasHeader = !s.attIgnoreHeader;
      renderAttendancePreview(fullGrid, hasHeader);

      if (s.attIgnoreHeader) {
        const lines = text.split(/\r?\n/);
        lines.shift();
        text = lines.join("\n");
      }
      // Parse locally (same logic as settings tryParse)
      const grid = parseCSV(text);
      const header = grid[0]?.map?.((h) => (h || "").toLowerCase()) || [];
      const headerDetected =
        header.includes("card number") ||
        header.includes("card") ||
        header.includes("name");
      const start = headerDetected ? 1 : 0;
      const out = [];
      for (let i = start; i < grid.length; i++) {
        const row = grid[i];
        if (!row || row.length === 0) continue;
        const card = normalizeCard(row[0] || "");
        const name = (row[1] === undefined ? "" : row[1]).toString().trim();
        if (validateCard(card)) out.push({ card, name });
      }
      DataService.setEntries(out);
      if (countBadge) countBadge.textContent = `${out.length} entries`;
      alert(`Attendance loaded (${out.length} valid entries).`);
    } catch (err) {
      alert("Failed to load: " + (err?.message || err));
    }
  });

  // Load Google Sheets data in Attendees Modal
  const loadGsheetConnectedAtt = document.getElementById(
    "loadGsheetConnectedAtt"
  );
  loadGsheetConnectedAtt?.addEventListener("click", async () => {
    const s = DataService.getSettings();
    if (!s.gsWorkbookId || !s.attSheetGid) {
      alert("Connect to Google and select an attendance sheet.");
      return;
    }
    const csv = `https://docs.google.com/spreadsheets/d/${
      s.gsWorkbookId
    }/export?format=csv&gid=${encodeURIComponent(s.attSheetGid)}`;
    try {
      const r = await fetch(csv, { mode: "cors" });
      if (!r.ok) throw new Error(`HTTP ${r.status}`);
      let text = await r.text();

      if (s.attIgnoreHeader) {
        const lines = text.split(/\r?\n/);
        lines.shift(); // Remove first line
        text = lines.join("\n");
      }
      // Parse locally (same logic as settings tryParse)
      const grid = parseCSV(text);
      const header = grid[0]?.map?.((h) => (h || "").toLowerCase()) || [];
      const headerDetected =
        header.includes("card number") ||
        header.includes("card") ||
        header.includes("name");
      const start = headerDetected ? 1 : 0;
      const out = [];
      for (let i = start; i < grid.length; i++) {
        const row = grid[i];
        if (!row || row.length === 0) continue;
        const card = normalizeCard(row[0] || "");
        const name = (row[1] === undefined ? "" : row[1]).toString().trim();
        if (validateCard(card)) out.push({ card, name });
      }
      DataService.setEntries(out);
      if (countBadge) countBadge.textContent = `${out.length} entries`;

      // Trigger settings update to refresh the attendees display
      showSettings();

      alert(`Attendance loaded (${out.length} valid entries).`);
    } catch (err) {
      alert("Failed to load: " + (err?.message || err));
    }
  });

  loadPrizesConnected?.addEventListener("click", async () => {
    const s = DataService.getSettings();
    if (!s.gsWorkbookId || !s.prizeSheetGid) {
      alert("Connect to Google and select a prizes sheet.");
      return;
    }
    const csv = `https://docs.google.com/spreadsheets/d/${
      s.gsWorkbookId
    }/export?format=csv&gid=${encodeURIComponent(s.prizeSheetGid)}`;
    try {
      const r = await fetch(csv, { mode: "cors" });
      if (!r.ok) throw new Error(`HTTP ${r.status}`);
      let text = await r.text();

      // Parse the full CSV for preview
      const fullGrid = parseCSV(text);
      const hasHeader = !s.prizeIgnoreHeader;
      renderPrizesPreview(fullGrid, hasHeader);

      // Use proper CSV parsing instead of manual splitting
      const grid = parseCSV(text);
      const start = s.prizeIgnoreHeader ? 0 : 1; // If there's a header, skip it (start at 1)
      const list = [];
      console.log("Prize loading:", {
        gridLength: grid.length,
        hasHeader: !s.prizeIgnoreHeader,
        startIndex: start,
        firstFewRows: grid.slice(0, 3),
      });

      for (let i = start; i < grid.length; i++) {
        const row = grid[i];
        if (!row || row.length === 0) continue;
        const prize = row[0]?.trim();
        if (prize) list.push(prize);
      }
      console.log("Parsed prizes:", list);
      DataService.setPrizes(list);
      DataService.setPrizeIndex(0); // Reset to first prize
      renderPrizes();
      showNextPrize(); // Update the next prize display
      alert(`Prizes loaded (${list.length} prizes).`);
    } catch (err) {
      alert("Failed to load prizes: " + (err?.message || err));
    }
  });

  function savePrizes(list) {
    DataService.setPrizes(list);
    DataService.setPrizeIndex(0); // Reset to first prize when prizes change
    showNextPrize(); // Update display when prizes change
  }
  function getPrizes() {
    return DataService.getPrizes();
  }
  function renderPrizes() {
    if (!prizeTbody) return;
    const list = getPrizes();
    prizeTbody.innerHTML = "";
    list.forEach((p, i) => {
      const tr = document.createElement("tr");
      tr.dataset.index = String(i);
      tr.innerHTML = `<td>${i + 1}</td><td>${p}</td><td>
        <button class="mini-btn" data-act="up">▲</button>
        <button class="mini-btn" data-act="down">▼</button>
        <button class="mini-btn" data-act="dup">Duplicate</button>
        <button class="mini-btn" data-act="del">Delete</button>
      </td>`;
      prizeTbody.appendChild(tr);
    });
    // Update the remaining prizes counter when prizes are modified
    updateRemainingPrizesCounter();
  }
  function addPrizeItems(name, count) {
    if (!name || !name.trim()) return;
    const list = getPrizes();
    for (let i = 0; i < (count || 1); i++) list.push(name.trim());
    savePrizes(list);
    renderPrizes();
  }
  addPrize?.addEventListener("click", () => {
    const n = prizeName?.value || "";
    const c = parseInt(prizeCount?.value || "1", 10) || 1;
    addPrizeItems(n, c);
    if (prizeName) prizeName.value = "";
  });
  clearPrizes?.addEventListener("click", () => {
    if (confirm("Clear all prizes?")) {
      savePrizes([]);
      renderPrizes();
    }
  });
  prizeTbody?.addEventListener("click", (e) => {
    const btn = e.target.closest?.("button");
    if (!btn) return;
    const act = btn.dataset.act;
    const row = btn.closest("tr");
    if (!row) return;
    const idx = parseInt(row.dataset.index || "-1", 10);
    if (isNaN(idx)) return;
    const list = getPrizes();
    if (act === "up" && idx > 0) {
      const [it] = list.splice(idx, 1);
      list.splice(idx - 1, 0, it);
    } else if (act === "down" && idx < list.length - 1) {
      const [it] = list.splice(idx, 1);
      list.splice(idx + 1, 0, it);
    } else if (act === "dup") {
      list.splice(idx + 1, 0, list[idx]);
    } else if (act === "del") {
      list.splice(idx, 1);
    }
    savePrizes(list);
    renderPrizes();
  });
  loadPrizesSheet?.addEventListener("click", async () => {
    const url = prizesSheetUrl?.value || "";
    if (!url.trim()) {
      alert("Enter a Google Sheet URL");
      return;
    }
    try {
      const u = new URL(url.trim());
      const m = u.pathname.match(/\/spreadsheets\/d\/([^/]+)/);
      const id = m ? m[1] : null;
      const hashParams = new URLSearchParams(u.hash.replace(/^#/, ""));
      const gid = hashParams.get("gid") || u.searchParams.get("gid") || "0";
      const csvUrl = id
        ? `https://docs.google.com/spreadsheets/d/${id}/export?format=csv&gid=${encodeURIComponent(
            gid
          )}`
        : url;
      const r = await fetch(csvUrl, { mode: "cors" });
      if (!r.ok) throw new Error(`HTTP ${r.status}`);
      const text = await r.text();
      const grid = parseCSV(text);
      const start = grid.length > 0 && grid[0].some((h) => h) ? 1 : 0;
      const list = [];
      for (let i = start; i < grid.length; i++) {
        const name = (grid[i][0] || "").toString().trim();
        if (name) list.push(name);
      }
      if (list.length === 0) {
        alert("No prizes found in first column.");
        return;
      }
      savePrizes(list);
      renderPrizes();
    } catch (err) {
      alert("Failed to load prizes: " + (err?.message || err));
    }
  });

  // Listen for settings/grid changes from the Settings modal
  window.addEventListener("settingsChanged", () => {
    if (DataService.getSettings().showProgress) updateProgressFromGrid();
  });
  window.addEventListener("gridUpdated", (e) => {
    if (DataService.getSettings().showProgress) updateProgressFromGrid();
    // show toasts for newly checked-in
    const names = (e.detail && e.detail.names) || [];
    if (names.length > 0) showCheckinToasts(names);
  });
  window.addEventListener("historyUpdated", () => {
    renderLeaderBoard();
    renderHistory();
  });

  function showCheckinToasts(names) {
    const layer = document.getElementById("toastLayer");
    if (!layer) return;
    let delay = 0;
    names.slice(0, 10).forEach((n, i) => {
      setTimeout(() => {
        const div = document.createElement("div");
        div.className = "toast";
        div.textContent = `Checked in: ${n}`;
        layer.appendChild(div);
        try {
          AudioFX.beep({
            freq: 880,
            dur: 0.08,
            type: "sine",
            gain: 0.05,
            sweep: { to: 1320 },
          });
        } catch {}
        setTimeout(() => {
          div.classList.add("hide");
          setTimeout(() => div.remove(), 400);
        }, 2800);
      }, delay);
      delay += 300;
    });
  }
  // Warm up fanfare sample (non-blocking)
  try {
    if (
      settings.sound &&
      settings.fanfare &&
      settings.fanfareSample &&
      AudioFX.preloadFanfare
    )
      AudioFX.preloadFanfare("resources/mixkit-winning-notification-2018.wav");
  } catch {}
  const canvas = document.getElementById("reelsCanvas");
  if (!canvas) {
    logDiag("reelsCanvas missing", true);
    return;
  }
  const od = new OdometerCanvas(canvas);
  od.draw();
  if (countBadge) countBadge.textContent = `${data.length} entries`;

  // optional auto fullscreen
  if (settings.autoFullscreen && document.documentElement.requestFullscreen) {
    document.documentElement.requestFullscreen().catch(() => {});
  }
}
function TrueFalse(v) {
  return !!v;
}

// ========================
// CSV Parsing (robust)
// ========================
function parseCSV(text) {
  if (!text) return [];
  // Remove BOM, normalize newlines
  text = text.replace(/^\uFEFF/, "").replace(/\r/g, "");
  const lines = text.split("\n").filter((l) => l && l.trim().length > 0);
  const delim = lines[0] && lines[0].includes("\t") ? "\t" : ",";
  const rows = [];
  for (const line of lines) {
    const cells = [];
    let cur = "",
      q = false;
    for (let i = 0; i < line.length; i++) {
      const ch = line[i];
      if (ch === '"') {
        if (q && line[i + 1] === '"') {
          cur += '"';
          i++;
        } else q = !q;
      } else if (ch === delim && !q) {
        cells.push(cur);
        cur = "";
      } else {
        cur += ch;
      }
    }
    cells.push(cur);
    rows.push(cells.map((c) => c.trim()));
  }
  return rows;
}
function normalizeCard(c) {
  return (c || "").toString().trim().toUpperCase().replace(/\s+/g, "");
}
function validateCard(card) {
  return /^[A-Z0-9][0-9]{6}$/.test(card);
}

// ========================
// Settings bootstrap
// ========================
function initSettings() {
  // Controls
  const noRepeat = document.getElementById("noRepeat");
  const spinMs = document.getElementById("spinMs");
  const stopGap = document.getElementById("stopGap");
  const confetti = document.getElementById("confettiLevel");
  const sound = document.getElementById("sound");
  const fanfare = document.getElementById("fanfare");
  const fanfareSample = document.getElementById("fanfareSample");
  const tick = document.getElementById("tick");
  const seed = document.getElementById("seed");
  const autoFs = document.getElementById("autoFs");
  // Google Sheet + table controls (declare early so we can set initial values)
  const gsheetUrl = document.getElementById("gsheetUrl");
  const gsAutoRefresh = document.getElementById("gsAutoRefresh");
  const gsRefreshNow = document.getElementById("gsRefreshNow");
  const gsFilterEnable = document.getElementById("gsFilterEnable");
  const gsFilterColumn = document.getElementById("gsFilterColumn");
  const gsFilterValue = document.getElementById("gsFilterValue");
  const mapCard = document.getElementById("mapCard");
  const mapName = document.getElementById("mapName");
  const statusCol = document.getElementById("statusCol");
  const statusValue = document.getElementById("statusValue");
  const entriesHead = document.getElementById("entriesHead");
  const entriesBody = document.getElementById("entriesBody");
  const kpiTotal = document.getElementById("kpiTotal");
  const kpiChecked = document.getElementById("kpiChecked");
  const kpiCheckedPct = document.getElementById("kpiCheckedPct");
  const kpiNot = document.getElementById("kpiNot");
  const kpiNotPct = document.getElementById("kpiNotPct");

  const s = DataService.getSettings();
  if (noRepeat) noRepeat.checked = s.noRepeat;
  if (spinMs) spinMs.value = s.spinDurationMs;
  if (stopGap) stopGap.value = s.stopGapMs;
  if (confetti) confetti.value = s.confettiLevel;
  if (sound) sound.checked = s.sound;
  if (fanfare) fanfare.checked = s.fanfare;
  if (fanfareSample) fanfareSample.checked = s.fanfareSample !== false;
  if (tick) tick.checked = s.tick;
  if (seed) seed.value = s.seed || "";
  if (autoFs) autoFs.checked = s.autoFullscreen;

  // Use centralized Google Sheets URL instead of old gsUrl field
  if (gsheetUrl) {
    const centralizedUrl = getCurrentGoogleSheetsUrl();
    gsheetUrl.value = centralizedUrl || s.gsUrl || "";
  }

  if (gsAutoRefresh) gsAutoRefresh.value = String(s.gsAutoRefreshSec || 0);
  // Filter UI will be completed after grid is loaded
  if (gsFilterEnable)
    gsFilterEnable.checked = !!(s.gsFilter && s.gsFilter.enabled);

  function saveFromUI() {
    const s2 = {
      noRepeat: noRepeat?.checked ?? s.noRepeat,
      spinDurationMs: Math.max(
        1200,
        parseInt(spinMs?.value || s.spinDurationMs, 10) || s.spinDurationMs
      ),
      stopGapMs: Math.max(
        80,
        parseInt(stopGap?.value || s.stopGapMs, 10) || s.stopGapMs
      ),
      perReelStaggerMs: s.perReelStaggerMs,
      confettiLevel: confetti?.value || s.confettiLevel,
      sound: sound?.checked ?? s.sound,
      fanfare: fanfare?.checked ?? s.fanfare,
      fanfareSample: fanfareSample?.checked ?? s.fanfareSample,
      tick: tick?.checked ?? s.tick,
      seed: (seed?.value || "").trim(),
      autoFullscreen: autoFs?.checked ?? s.autoFullscreen,

      // Keep existing centralized Google connection settings
      gsWorkbookId: s.gsWorkbookId,
      gsSheets: s.gsSheets,
      attSheetGid: s.attSheetGid,
      attIgnoreHeader: s.attIgnoreHeader,
      prizeSheetGid: s.prizeSheetGid,
      prizeIgnoreHeader: s.prizeIgnoreHeader,

      gsAutoRefreshSec: parseInt(gsAutoRefresh?.value || "0", 10) || 0,
      gsFilter: {
        enabled: !!gsFilterEnable?.checked,
        colIndex:
          gsFilterColumn && gsFilterColumn.value !== ""
            ? parseInt(gsFilterColumn.value, 10)
            : null,
        value: gsFilterValue?.value || "",
      },
      map: {
        card:
          typeof mapCard?.value !== "undefined" && mapCard.value !== ""
            ? parseInt(mapCard.value, 10)
            : 0,
        name:
          typeof mapName?.value !== "undefined" && mapName.value !== ""
            ? parseInt(mapName.value, 10)
            : 1,
      },
      status: {
        col:
          typeof statusCol?.value !== "undefined" && statusCol.value !== ""
            ? parseInt(statusCol.value, 10)
            : null,
        value: statusValue?.value || "",
      },
    };
    DataService.setSettings(s2);
    try {
      window.dispatchEvent(new Event("settingsChanged"));
    } catch {}
  }
  // Defer listener wiring until after all elements are defined

  // Data import/export
  const fileInput = document.getElementById("csvFile");
  const pasteArea = document.getElementById("pasteArea");
  const importBtn = document.getElementById("importBtn");
  const saveBtn = document.getElementById("saveBtn");
  const clearBtn = document.getElementById("clearBtn");
  const sampleBtn = document.getElementById("sampleBtn");
  const loadGsheet = document.getElementById("loadGsheet");
  const loadGsheetConnected = document.getElementById("loadGsheetConnected");
  const info = document.getElementById("info");
  const exportHist = document.getElementById("exportHistory");
  const clearHist = document.getElementById("clearHistory");
  const histBody = document.getElementById("histBody");
  const histBodySettings = document.getElementById("histBodySettings");
  const histBodyMain = document.getElementById("histBodyMain");
  const histInfo = document.getElementById("histInfo");

  // Additional references for History modal buttons
  const exportHistMain = document.getElementById("exportHistoryMain");
  const clearHistMain = document.getElementById("clearHistoryMain");
  const histInfoMain = document.getElementById("histInfoMain");

  // Now that all controls are declared, wire change listeners to persist settings
  [
    noRepeat,
    spinMs,
    stopGap,
    confetti,
    sound,
    fanfare,
    fanfareSample,
    tick,
    seed,
    autoFs,
    // gsheetUrl removed - now managed centrally in Google modal
    gsAutoRefresh,
    gsFilterEnable,
    gsFilterColumn,
    gsFilterValue,
    mapCard,
    mapName,
    statusCol,
    statusValue,
  ]
    .filter(Boolean)
    .forEach((el) => {
      el.addEventListener("change", saveFromUI);
      el.addEventListener("input", saveFromUI);
    });

  function renderNamesList(grid) {
    if (!entriesHead || !entriesBody) return;
    const sNow = DataService.getSettings();
    entriesHead.innerHTML = "";
    entriesBody.innerHTML = "";
    const trh = document.createElement("tr");
    trh.innerHTML = `<th>#</th><th>Status</th><th>Card</th><th>Name</th>`;
    entriesHead.appendChild(trh);

    const g = grid && Array.isArray(grid) ? grid : DataService.getGrid();
    if (!g || g.length === 0) {
      const rows = DataService.getEntries();
      const histSet = new Set(DataService.getHistory().map((h) => h.card));
      rows.forEach((r, idx) => {
        const isPrev = histSet.has(r.card);
        const statusLabel = isPrev ? "Previous Winner" : "Unknown";
        const statusClass = isPrev ? "pill-prev" : "pill-unknown";
        const tr = document.createElement("tr");
        tr.innerHTML = `<td>${
          idx + 1
        }</td><td><span class="pill ${statusClass}">${statusLabel}</span></td><td><code>${
          r.card
        }</code></td><td>${r.name || ""}</td>`;
        entriesBody.appendChild(tr);
      });
      if (info) info.textContent = `${rows.length} entries`;
      return;
    }

    const headerRow = g[0] || [];
    const hasHeader =
      headerRow.length > 0 &&
      headerRow.some((h) => h && String(h).trim() !== "");
    const start = hasHeader ? 1 : 0;
    const doFilter = !!gsFilterEnable?.checked;
    const colIndex = parseInt(gsFilterColumn?.value || "", 10);
    const filterValue = gsFilterValue?.value || "";
    const cIdx = sNow?.map?.card ?? 0;
    const nIdx = sNow?.map?.name ?? 1;
    const sIdx = sNow?.status?.col ?? NaN;
    const sVal = sNow?.status?.value || "";
    const histSet = new Set(DataService.getHistory().map((h) => h.card));
    let shown = 0,
      eligible = 0,
      ineligible = 0,
      prev = 0;
    for (let i = start; i < g.length; i++) {
      const row = g[i] || [];
      if (doFilter && !Number.isNaN(colIndex) && filterValue) {
        const val = (row[colIndex] ?? "").toString().trim();
        if (val !== filterValue) continue;
      }
      const card = normalizeCard(row[cIdx] || "");
      const name = (row[nIdx] === undefined ? "" : row[nIdx]).toString().trim();
      if (!validateCard(card)) continue;
      let statusLabel = "Ineligible";
      let statusClass = "pill-ineligible";
      if (histSet.has(card)) {
        statusLabel = "Previous Winner";
        statusClass = "pill-prev";
        prev++;
      } else if (
        !Number.isNaN(sIdx) &&
        sVal &&
        (row[sIdx] ?? "").toString().trim() === sVal
      ) {
        statusLabel = "Eligible";
        statusClass = "pill-eligible";
        eligible++;
      } else {
        ineligible++;
      }
      const tr = document.createElement("tr");
      tr.innerHTML = `<td>${++shown}</td><td><span class="pill ${statusClass}">${statusLabel}</span></td><td><code>${card}</code></td><td>${name}</td>`;
      entriesBody.appendChild(tr);
    }
    if (info)
      info.textContent = `${shown} listed — ${eligible} eligible, ${ineligible} ineligible, ${prev} previous winners`;
  }

  let parsedRows = DataService.getEntries();
  renderNamesList(DataService.getGrid());

  // If we have a previously loaded grid, restore full preview, selections, and KPIs
  try {
    const storedGrid = DataService.getGrid();
    if (Array.isArray(storedGrid) && storedGrid.length > 0) {
      lastGrid = storedGrid;
      // Populate columns then restore selections
      populateColumns(lastGrid);
      // Restore mapping and filter/status column pointers first
      if (gsFilterColumn && s?.gsFilter?.colIndex != null)
        gsFilterColumn.value = String(s.gsFilter.colIndex);
      if (mapCard && s?.map?.card != null) mapCard.value = String(s.map.card);
      if (mapName && s?.map?.name != null) mapName.value = String(s.map.name);
      if (statusCol && s?.status?.col != null)
        statusCol.value = String(s.status.col);
      // Populate values lists based on restored columns
      populateValues(lastGrid);
      // Restore selected values
      if (gsFilterValue && s?.gsFilter?.value)
        gsFilterValue.value = s.gsFilter.value;
      if (statusValue && s?.status?.value) statusValue.value = s.status.value;
      // Render preview + derive entries + KPIs
      applyFilterAndRender();
      updateKPIs();
    }
  } catch {}

  // Helpers for Google Sheet filtering UI
  function setOptions(select, items) {
    if (!select) return;
    select.innerHTML = "";
    items.forEach(({ value, label }) => {
      const opt = document.createElement("option");
      opt.value = value;
      opt.textContent = label;
      select.appendChild(opt);
    });
  }
  function populateColumns(grid) {
    if (!gsFilterColumn) return;
    const header = grid[0] || [];
    const hasHeader = header.length > 0 && header.some((h) => h);
    const cols = hasHeader ? header : header.map((_, i) => `Column ${i + 1}`);
    const items = [{ value: "", label: "(Select a column)" }].concat(
      cols.map((h, i) => ({
        value: String(i),
        label: String(h || `Column ${i + 1}`),
      }))
    );
    setOptions(gsFilterColumn, items);
    const mapItems = cols.map((h, i) => ({
      value: String(i),
      label: String(h || `Column ${i + 1}`),
    }));
    setOptions(mapCard, mapItems);
    setOptions(mapName, mapItems);
    setOptions(statusCol, [{ value: "", label: "(None)" }].concat(mapItems));
  }
  function populateValues(grid) {
    if (!gsFilterValue && !statusValue) return;
    const start = 1; // assume first row is header when filtering via Google
    // Filter values list
    if (gsFilterValue) {
      const idx = parseInt(gsFilterColumn?.value || "", 10);
      let items = [{ value: "", label: "(Select a value)" }];
      if (!Number.isNaN(idx)) {
        const uniq = new Set();
        for (let i = start; i < grid.length; i++) {
          const row = grid[i];
          if (!row) continue;
          const v = (row[idx] ?? "").toString().trim();
          if (v) uniq.add(v);
        }
        items = items.concat(
          [...uniq].sort().map((v) => ({ value: v, label: v }))
        );
      }
      setOptions(gsFilterValue, items);
    }
    // Status values list (independent of filter column selection)
    if (statusValue) {
      const sIdx = parseInt(statusCol?.value || "", 10);
      let sItems = [{ value: "", label: "(Select)" }];
      if (!Number.isNaN(sIdx)) {
        const uniq2 = new Set();
        for (let i = start; i < grid.length; i++) {
          const row = grid[i];
          if (!row) continue;
          const v2 = (row[sIdx] ?? "").toString().trim();
          if (v2) uniq2.add(v2);
        }
        sItems = sItems.concat(
          [...uniq2].sort().map((v) => ({ value: v, label: v }))
        );
      }
      setOptions(statusValue, sItems);
    }
  }
  gsFilterColumn?.addEventListener("change", () => {
    if (lastGrid) {
      populateValues(lastGrid);
      renderNamesList(lastGrid);
    }
  });
  gsFilterEnable?.addEventListener("change", () => {
    if (lastGrid) {
      applyFilterAndRender();
    }
  });
  gsFilterColumn?.addEventListener("change", () => {
    if (lastGrid) {
      applyFilterAndRender();
    }
  });
  gsFilterValue?.addEventListener("change", () => {
    if (lastGrid) {
      applyFilterAndRender();
    }
  });
  mapCard?.addEventListener("change", () => {
    if (lastGrid) {
      applyFilterAndRender();
    }
  });
  mapName?.addEventListener("change", () => {
    if (lastGrid) {
      applyFilterAndRender();
    }
  });
  statusCol?.addEventListener("change", () => {
    if (lastGrid) {
      populateValues(lastGrid);
      updateKPIs();
      renderNamesList(lastGrid);
    }
  });
  statusValue?.addEventListener("change", () => {
    if (lastGrid) {
      updateKPIs();
      renderNamesList(lastGrid);
    }
  });

  function renderHistory() {
    const hist = DataService.getHistory();
    // newest first
    const list = [...hist].sort((a, b) => (a.ts < b.ts ? 1 : -1));

    // Function to create table row HTML
    const createRowHTML = (h, i) => {
      const dt = new Date(h.ts);
      const time = isNaN(dt) ? h.ts : dt.toLocaleString();

      // Better prize display logic
      let prizeDisplay = "No prize assigned";
      if (h.prize && h.prize.trim() !== "" && h.prize !== "—") {
        prizeDisplay = h.prize;
      } else if (h.prize === "no prize claimed") {
        prizeDisplay = "No prize claimed";
      }

      return `<td>${i + 1}</td><td>${time}</td><td><code>${
        h.card
      }</code></td><td>${h.name || ""}</td><td>${prizeDisplay}</td>`;
    };

    // Render to all history tables using innerHTML (faster and simpler)
    const tableHTML = list
      .map((h, i) => `<tr>${createRowHTML(h, i)}</tr>`)
      .join("");

    if (histBodySettings) histBodySettings.innerHTML = tableHTML;
    if (histBody) histBody.innerHTML = tableHTML;
    if (histBodyMain) histBodyMain.innerHTML = tableHTML;

    // Update info badges
    const winnerText = `${hist.length} winner${hist.length === 1 ? "" : "s"}`;
    if (histInfo) histInfo.textContent = winnerText;
    if (histInfoMain) histInfoMain.textContent = winnerText;
  }
  renderHistory();

  function tryParse(text) {
    const grid = parseCSV(text);
    if (grid.length === 0) {
      alert("No rows found.");
      return;
    }
    try {
      DataService.setGrid(grid);
    } catch {}
    const header = grid[0].map((h) => (h || "").toLowerCase());
    const hasHeader =
      header.includes("card number") ||
      header.includes("card") ||
      header.includes("name");
    const start = hasHeader ? 1 : 0;
    const out = [];
    for (let i = start; i < grid.length; i++) {
      const row = grid[i];
      if (!row || row.length === 0) continue;
      const card = normalizeCard(row[0] || "");
      const name = (row[1] === undefined ? "" : row[1]).toString().trim();
      if (validateCard(card)) out.push({ card, name });
    }
    parsedRows = out;
    DataService.setEntries(parsedRows);
    lastGrid = grid;
    renderNamesList(grid);
    logDiag(`Parsed ${out.length} rows.`);
  }

  function processGrid(grid, filter) {
    if (grid.length === 0) return [];
    const header = grid[0].map((h) => (h || "").toLowerCase());
    const hasHeader =
      header.includes("card number") ||
      header.includes("card") ||
      header.includes("name") ||
      header.some(Boolean);
    const start = hasHeader ? 1 : 0;
    const out = [];
    for (let i = start; i < grid.length; i++) {
      const row = grid[i];
      if (!row || row.length === 0) continue;
      if (filter && filter.enabled && Number.isInteger(filter.colIndex)) {
        const val = (row[filter.colIndex] ?? "").toString().trim();
        if (val !== filter.value) continue;
      }
      const sNow = DataService.getSettings();
      const cIdx = sNow?.map?.card ?? 0;
      const nIdx = sNow?.map?.name ?? 1;
      const card = normalizeCard(row[cIdx] || "");
      const name = (row[nIdx] === undefined ? "" : row[nIdx]).toString().trim();
      if (validateCard(card)) out.push({ card, name });
    }
    return out;
  }

  function updateKPIs() {
    if (!lastGrid) {
      if (kpiTotal) kpiTotal.textContent = "0";
      if (kpiChecked) kpiChecked.textContent = "0";
      if (kpiCheckedPct) kpiCheckedPct.textContent = "0%";
      if (kpiNot) kpiNot.textContent = "0";
      if (kpiNotPct) kpiNotPct.textContent = "0%";
      return;
    }
    const start = lastGrid[0] && lastGrid[0].some((h) => h) ? 1 : 0;
    const sIdx = parseInt(statusCol?.value || "", 10);
    const sVal = statusValue?.value || "";
    let total = 0,
      checked = 0;
    for (let i = start; i < lastGrid.length; i++) {
      const row = lastGrid[i] || [];
      total++;
      if (!Number.isNaN(sIdx) && sVal) {
        const v2 = (row[sIdx] ?? "").toString().trim();
        if (v2 === sVal) checked++;
      }
    }
    const notC = Math.max(0, total - checked);
    const pct = total > 0 ? Math.round((checked * 1000) / total) / 10 : 0;
    const pctNot = total > 0 ? Math.round((notC * 1000) / total) / 10 : 0;
    if (kpiTotal) kpiTotal.textContent = String(total);
    if (kpiChecked) kpiChecked.textContent = String(checked);
    if (kpiCheckedPct) kpiCheckedPct.textContent = `${pct}%`;
    if (kpiNot) kpiNot.textContent = String(notC);
    if (kpiNotPct) kpiNotPct.textContent = `${pctNot}%`;
  }

  fileInput?.addEventListener("change", async (e) => {
    try {
      const f = e.target.files[0];
      if (!f) {
        alert("No file selected.");
        return;
      }
      const text = await f.text();
      tryParse(text);
      try {
        document.getElementById("fileDetails")?.removeAttribute("open");
      } catch {}
    } catch (err) {
      alert("File read failed: " + err.message);
    }
  });
  importBtn?.addEventListener("click", () => {
    try {
      tryParse(pasteArea?.value || "");
    } catch (err) {
      alert("Parse failed: " + err.message);
    }
  });
  // Entries persist automatically on import, filter, mapping, and sheet refresh
  clearBtn?.addEventListener("click", () => {
    if (confirm("Clear all saved entries?")) {
      DataService.clearEntries();
      parsedRows = [];
      renderObjectsView([]);
    }
  });

  exportHist?.addEventListener("click", () => {
    const hist = DataService.getHistory();
    const header = "timestamp,card,name,prize";
    const body = hist
      .map(
        (h) =>
          `${h.ts},${h.card},${(h.name || "").replace(/,/g, ";")},${(
            h.prize || ""
          ).replace(/,/g, ";")}`
      )
      .join("\n");
    const blob = new Blob([header + "\n" + body], { type: "text/csv" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "winner_history.csv";
    a.click();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  });
  clearHist?.addEventListener("click", () => {
    if (
      confirm("Clear winner history? This also resets no-repeat eligibility.")
    ) {
      // Clear the single source of truth
      DataService.clearHistory();

      // Update the prize display and counter after clearing history
      showNextPrize();

      // Force update of entry count badge if no-repeat is enabled
      const s = DataService.getSettings();
      if (s.noRepeat) {
        const allEntries = DataService.getEntries().length;
        const countBadge = document.getElementById("countBadge");
        if (countBadge) {
          countBadge.textContent = `${allEntries} entries (eligible)`;
        }
      }

      // Dispatch event to update all components (leaderboard and history)
      try {
        window.dispatchEvent(new Event("historyUpdated"));
      } catch {}

      alert("History cleared.");
    }
  });

  // History modal buttons (same functionality as settings buttons)
  exportHistMain?.addEventListener("click", () => {
    const hist = DataService.getHistory();
    const header = "timestamp,card,name,prize";
    const body = hist
      .map(
        (h) =>
          `${h.ts},${h.card},${(h.name || "").replace(/,/g, ";")},${(
            h.prize || ""
          ).replace(/,/g, ";")}`
      )
      .join("\n");
    const blob = new Blob([header + "\n" + body], { type: "text/csv" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "winner_history.csv";
    a.click();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  });

  clearHistMain?.addEventListener("click", () => {
    if (
      confirm("Clear winner history? This also resets no-repeat eligibility.")
    ) {
      // Clear the single source of truth
      DataService.clearHistory();

      // Update the prize display and counter after clearing history
      showNextPrize();

      // Force update of entry count badge if no-repeat is enabled
      const s = DataService.getSettings();
      if (s.noRepeat) {
        const allEntries = DataService.getEntries().length;
        const countBadge = document.getElementById("countBadge");
        if (countBadge) {
          countBadge.textContent = `${allEntries} entries (eligible)`;
        }
      }

      // Dispatch event to update all components (leaderboard and history)
      try {
        window.dispatchEvent(new Event("historyUpdated"));
      } catch {}

      alert("History cleared.");
    }
  });

  // Diagnostics banner shows version
  logDiag(`Settings initialized. Version ${__APP_VERSION__}`);
  // Note: Auto-refresh no longer auto-starts on page load to avoid unintentionally
  // overwriting existing entries. Use "Load from Google Sheet" or "Refresh Now" to begin.

  // Google Sheet loader
  function extractSheetInfo(url) {
    try {
      const u = new URL(url.trim());
      const m = u.pathname.match(/\/spreadsheets\/d\/([^/]+)/);
      const id = m ? m[1] : null;
      // gid can be in hash (#gid=...) or query
      const hashParams = new URLSearchParams(u.hash.replace(/^#/, ""));
      const gid = hashParams.get("gid") || u.searchParams.get("gid") || "0";
      return { id, gid };
    } catch {
      return { id: null, gid: null };
    }
  }
  let lastGrid = null;
  let applying = false;
  function applyFilterAndRender() {
    if (!lastGrid) return;
    // Render names list with status
    renderNamesList(lastGrid);
    // Compute parsedRows for saving
    const doFilter = !!gsFilterEnable?.checked;
    const colIndex = parseInt(gsFilterColumn?.value || "", 10);
    const value = gsFilterValue?.value || "";
    const rows = processGrid(
      lastGrid,
      doFilter && !Number.isNaN(colIndex) && value
        ? { enabled: true, colIndex, value }
        : null
    );
    parsedRows = rows;
    // Guard: avoid overwriting with empty set unless there was nothing saved yet
    try {
      const prev = DataService.getEntries() || [];
      if (rows.length > 0 || prev.length === 0) {
        DataService.setEntries(rows);
      } else {
        logDiag(
          "Sheet produced 0 valid entries; keeping previously saved entries. Adjust filter/mapping/status if needed."
        );
      }
    } catch {
      DataService.setEntries(rows);
    }
    updateKPIs();
  }

  // Helper function to get the current Google Sheets URL from centralized settings
  function getCurrentGoogleSheetsUrl() {
    const s = DataService.getSettings();
    if (!s.gsWorkbookId || !s.attSheetGid) {
      return null;
    }
    return `https://docs.google.com/spreadsheets/d/${s.gsWorkbookId}/edit#gid=${s.attSheetGid}`;
  }

  async function loadFromGoogle(url) {
    if (!url || !url.trim()) {
      alert("Enter a Google Sheet URL first.");
      return;
    }
    const { id, gid } = extractSheetInfo(url);
    if (!id) {
      alert("Could not find Sheet ID in the URL.");
      return;
    }
    // Prefer published CSV if user pasted a published URL; otherwise use export CSV.
    // export CSV works for sheets with link-sharing set to Anyone with the link (or public).
    const exportCsv = `https://docs.google.com/spreadsheets/d/${id}/export?format=csv&gid=${encodeURIComponent(
      gid || "0"
    )}`;
    const publishedCsv = `https://docs.google.com/spreadsheets/d/e/${id}/pub?gid=${encodeURIComponent(
      gid || "0"
    )}&single=true&output=csv`;
    const candidates = [exportCsv];
    // Some users paste already-published URLs; if so, try to keep it
    if (/\/pub(?:html)?/.test(url)) {
      const pub = url.replace(
        /pub(?:html)?(?:\?.*)?$/,
        `pub?gid=${encodeURIComponent(gid || "0")}&single=true&output=csv`
      );
      candidates.unshift(pub);
    }
    // Try each candidate until one succeeds
    let lastErr;
    for (const c of candidates) {
      try {
        const r = await fetch(c, { mode: "cors" });
        if (!r.ok) throw new Error(`HTTP ${r.status}`);
        const text = await r.text();
        const grid = parseCSV(text);
        // Compute newly checked-in delta before overwriting stored grid
        const prev = DataService.getGrid();
        const sNowForDelta = DataService.getSettings();
        const sCol = sNowForDelta?.status?.col;
        const sVal = sNowForDelta?.status?.value || "";
        const cIdx = sNowForDelta?.map?.card ?? 0;
        function toSet(g) {
          if (
            !Array.isArray(g) ||
            g.length === 0 ||
            sCol == null ||
            sVal === ""
          )
            return new Set();
          const start = g[0] && g[0].some((h) => h) ? 1 : 0;
          const st = new Set();
          for (let i = start; i < g.length; i++) {
            const row = g[i] || [];
            const v = (row[sCol] ?? "").toString().trim();
            if (v === sVal) {
              const card = normalizeCard(row[cIdx] || "");
              if (card) st.add(card);
            }
          }
          return st;
        }
        const before = toSet(prev);
        const after = toSet(grid);
        const newly = [];
        if (after.size > 0) {
          for (const card of after) {
            if (!before.has(card)) newly.push(card);
          }
        }

        lastGrid = grid;
        try {
          DataService.setGrid(grid);
        } catch {}
        populateColumns(grid);
        // Restore saved selections
        const sNow = DataService.getSettings();
        if (sNow?.gsFilter) {
          if (gsFilterEnable) gsFilterEnable.checked = !!sNow.gsFilter.enabled;
          if (gsFilterColumn && sNow.gsFilter.colIndex != null) {
            gsFilterColumn.value = String(sNow.gsFilter.colIndex);
          }
        }
        if (mapCard && sNow?.map?.card != null)
          mapCard.value = String(sNow.map.card);
        if (mapName && sNow?.map?.name != null)
          mapName.value = String(sNow.map.name);
        if (statusCol && sNow?.status?.col != null)
          statusCol.value = String(sNow.status.col);
        // Populate value lists and restore
        populateValues(grid);
        if (gsFilterValue && sNow?.gsFilter?.value) {
          gsFilterValue.value = sNow.gsFilter.value;
        }
        if (statusValue && sNow?.status?.value) {
          statusValue.value = sNow.status.value;
        }
        // Render, persist and report
        applyFilterAndRender();
        updateKPIs();
        // Map newly checked-in cards to names using current mapping
        if (newly.length > 0) {
          const nameIdx = DataService.getSettings()?.map?.name ?? 1;
          const start = grid[0] && grid[0].some((h) => h) ? 1 : 0;
          const nameByCard = new Map();
          for (let i = start; i < grid.length; i++) {
            const row = grid[i] || [];
            const card = normalizeCard(row[cIdx] || "");
            if (card) nameByCard.set(card, row[nameIdx] || "");
          }
          const names = newly.map((c) => nameByCard.get(c)).filter(Boolean);
          window.dispatchEvent(
            new CustomEvent("gridUpdated", { detail: { names } })
          );
        } else {
          window.dispatchEvent(
            new CustomEvent("gridUpdated", { detail: { names: [] } })
          );
        }
        const count = parsedRows?.length ?? 0;
        const fCol =
          gsFilterColumn?.selectedOptions?.[0]?.textContent || "column";
        const fVal = gsFilterValue?.value || "";
        const fOn = !!gsFilterEnable?.checked && fVal;
        logDiag(
          `Loaded ${count} valid entries from Google Sheets${
            fOn ? ` (filtered by ${fCol} = ${fVal})` : ""
          }.`
        );
        return;
      } catch (e) {
        lastErr = e;
      }
    }
    alert(
      'Failed to load CSV from Google Sheets. Ensure the sheet is shared as "Anyone with the link can view" or is Published to the web. Last error: ' +
        (lastErr?.message || lastErr)
    );
  }
  // Auto-refresh controls
  let refreshTimer = 0;
  let loading = false;
  function scheduleAutoRefresh() {
    if (refreshTimer) {
      clearInterval(refreshTimer);
      refreshTimer = 0;
    }
    const sec = parseInt(gsAutoRefresh?.value || "0", 10) || 0;
    const url = getCurrentGoogleSheetsUrl();
    if (sec > 0 && url) {
      refreshTimer = setInterval(async () => {
        if (loading) return;
        loading = true;
        try {
          await loadFromGoogle(url);
        } finally {
          loading = false;
        }
      }, sec * 1000);
      // If no grid yet, trigger an immediate load once
      if (!lastGrid) {
        (async () => {
          loading = true;
          try {
            await loadFromGoogle(url);
          } finally {
            loading = false;
          }
        })();
      }
    }
  }
  gsAutoRefresh?.addEventListener("change", scheduleAutoRefresh);
  gsRefreshNow?.addEventListener("click", () => {
    const centralizedUrl = getCurrentGoogleSheetsUrl();
    if (centralizedUrl) {
      loadFromGoogle(centralizedUrl);
    } else {
      alert(
        "Please connect to Google Sheets first using the 'Connect to Google' button in the sidebar."
      );
    }
  });
  loadGsheet?.addEventListener("click", () => {
    const centralizedUrl = getCurrentGoogleSheetsUrl();
    if (centralizedUrl) {
      loadFromGoogle(centralizedUrl);
      scheduleAutoRefresh();
    } else {
      alert(
        "Please connect to Google Sheets first using the 'Connect to Google' button in the sidebar."
      );
    }
  });
}

// ============
// Router
// ============
document.addEventListener("DOMContentLoaded", () => {
  const isPresentation = !!document.getElementById("reelsCanvas");
  const isSettings = !!document.getElementById("csvFile");
  DataService.init()
    .then(() => {
      try {
        if (isPresentation) initPresentation();
      } catch (e) {
        logDiag("initPresentation failed: " + e.message, true);
      }
      try {
        if (isSettings) initSettings();
      } catch (e) {
        logDiag("initSettings failed: " + e.message, true);
        alert("Settings failed: " + e.message);
      }
    })
    .catch((e) => {
      logDiag("Data init failed: " + (e?.message || e), true);
      try {
        if (isPresentation) initPresentation();
      } catch {}
      try {
        if (isSettings) initSettings();
      } catch {}
    });
});
