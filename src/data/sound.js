// Short, synthesized tap sounds — no audio files. Gated entirely by the
// caller checking settings.sound.enabled; this module just knows how to
// make the sounds and (on Android) pair them with a short vibration.
//
// iPhone notes:
// - Safari only lets a page make sound after a tap, so the audio context is
//   created/"unlocked" on the first touch (see unlockAudio) and every sound
//   waits for the context to be running before it is scheduled — otherwise
//   the first taps are silently dropped.
// - After switching apps, a call, or Siri, iOS leaves the context
//   "interrupted" (not "suspended"). It must be resumed from that state too,
//   or every sound stays silent until the app is reloaded.
// - Like the iPhone's own keyboard clicks, these sounds follow the ring /
//   silent switch: with the phone on silent they can't be heard.

let ctx = null;
let noiseBuf = null;

function audioCtx() {
  if (!ctx || ctx.state === "closed") {
    const AC = window.AudioContext || window.webkitAudioContext;
    if (!AC) return null;
    ctx = new AC();
    noiseBuf = null;
  }
  return ctx;
}

// Runs fn(context) once the context is actually running.
function withRunningCtx(fn) {
  const c = audioCtx();
  if (!c) return;
  if (c.state === "running") { fn(c); return; }
  const p = c.resume();
  if (p && typeof p.then === "function") p.then(() => { if (c.state === "running") fn(c); }).catch(() => {});
}

// Called on the first touches/clicks anywhere: creates the context inside
// the gesture and plays one silent sample, which is what iOS needs before it
// allows sound.
export function unlockAudio() {
  const c = audioCtx();
  if (!c) return;
  try {
    if (c.state !== "running") c.resume();
    const b = c.createBuffer(1, 1, 22050);
    const s = c.createBufferSource();
    s.buffer = b;
    s.connect(c.destination);
    s.start(0);
  } catch { /* ignore */ }
}

function whiteNoise(c) {
  if (!noiseBuf) {
    const len = Math.floor(c.sampleRate * 0.05);
    noiseBuf = c.createBuffer(1, len, c.sampleRate);
    const d = noiseBuf.getChannelData(0);
    for (let i = 0; i < len; i++) d[i] = Math.random() * 2 - 1;
  }
  return noiseBuf;
}

// `p` is the pitch multiplier (0.6 = low … 1 = normal … 1.6 = high).
function pop(c, p, volume) {
  const t = c.currentTime + 0.005;
  const o = c.createOscillator(), g = c.createGain();
  o.type = "sine";
  o.frequency.setValueAtTime(1100 * p, t);
  o.frequency.exponentialRampToValueAtTime(320 * p, t + 0.04);
  g.gain.setValueAtTime(0.0001, t);
  g.gain.exponentialRampToValueAtTime(volume * 0.5, t + 0.004);
  g.gain.exponentialRampToValueAtTime(0.0001, t + 0.05);
  o.connect(g).connect(c.destination);
  o.start(t); o.stop(t + 0.06);
}

function bubble(c, p, volume) {
  const t = c.currentTime + 0.005;
  const o = c.createOscillator(), g = c.createGain();
  o.type = "sine";
  o.frequency.setValueAtTime(380 * p, t);
  o.frequency.exponentialRampToValueAtTime(950 * p, t + 0.07);
  g.gain.setValueAtTime(0.0001, t);
  g.gain.exponentialRampToValueAtTime(volume * 0.45, t + 0.01);
  g.gain.exponentialRampToValueAtTime(0.0001, t + 0.09);
  o.connect(g).connect(c.destination);
  o.start(t); o.stop(t + 0.1);
}

// Short, crisp mechanical-keyboard-style click — a fast pitch drop with a
// hard percussive decay, distinct from the softer pop/bubble tones.
function click(c, p, volume) {
  const t = c.currentTime + 0.005;
  const o = c.createOscillator(), g = c.createGain();
  o.type = "square";
  o.frequency.setValueAtTime(2400 * p, t);
  o.frequency.exponentialRampToValueAtTime(1200 * p, t + 0.01);
  g.gain.setValueAtTime(0.0001, t);
  g.gain.exponentialRampToValueAtTime(volume * 0.3, t + 0.002);
  g.gain.exponentialRampToValueAtTime(0.0001, t + 0.018);
  o.connect(g).connect(c.destination);
  o.start(t); o.stop(t + 0.03);
}

// Soft "tock" like the iPhone keyboard: a very short band-passed noise
// tick over a tiny low body, ~15 ms. (The real system sound can't be used
// by a web page, so this is a close synthesized match.)
function keyboard(c, p, volume) {
  const t = c.currentTime + 0.005;
  const n = c.createBufferSource();
  n.buffer = whiteNoise(c);
  const bp = c.createBiquadFilter();
  bp.type = "bandpass";
  bp.frequency.setValueAtTime(2300 * p, t);
  bp.Q.value = 1.1;
  const ng = c.createGain();
  ng.gain.setValueAtTime(0.0001, t);
  ng.gain.exponentialRampToValueAtTime(volume * 0.9, t + 0.001);
  ng.gain.exponentialRampToValueAtTime(0.0001, t + 0.014);
  n.connect(bp).connect(ng).connect(c.destination);
  n.start(t); n.stop(t + 0.02);

  const o = c.createOscillator(), og = c.createGain();
  o.type = "sine";
  o.frequency.setValueAtTime(420 * p, t);
  o.frequency.exponentialRampToValueAtTime(260 * p, t + 0.015);
  og.gain.setValueAtTime(0.0001, t);
  og.gain.exponentialRampToValueAtTime(volume * 0.35, t + 0.001);
  og.gain.exponentialRampToValueAtTime(0.0001, t + 0.02);
  o.connect(og).connect(c.destination);
  o.start(t); o.stop(t + 0.025);
}

// Water-drop "plip": a sine that bends up fast, like a droplet.
function drop(c, p, volume) {
  const t = c.currentTime + 0.005;
  const o = c.createOscillator(), g = c.createGain();
  o.type = "sine";
  o.frequency.setValueAtTime(700 * p, t);
  o.frequency.exponentialRampToValueAtTime(1600 * p, t + 0.03);
  g.gain.setValueAtTime(0.0001, t);
  g.gain.exponentialRampToValueAtTime(volume * 0.5, t + 0.003);
  g.gain.exponentialRampToValueAtTime(0.0001, t + 0.06);
  o.connect(g).connect(c.destination);
  o.start(t); o.stop(t + 0.07);
}

export const SOUND_TYPES = ["keyboard", "pop", "bubble", "click", "drop"];
const SOUNDS = { pop, bubble, click, keyboard, drop };

// Pitch steps offered in Settings (multiplier on every frequency).
export const PITCH_STEPS = [0.6, 0.8, 1, 1.25, 1.6];

function clampPitch(p) {
  const n = Number(p);
  return Number.isFinite(n) && n > 0 ? Math.min(2, Math.max(0.4, n)) : 1;
}

// Plays the given type immediately — used by the Settings "test" buttons
// regardless of whether sound is currently enabled.
export function testSound(type, pitch = 1, volume = 0.4) {
  try {
    const fn = SOUNDS[type] || pop;
    const p = clampPitch(pitch);
    withRunningCtx((c) => { try { fn(c, p, volume); } catch { /* ignore */ } });
  } catch { /* Web Audio unavailable */ }
}

// Plays only if the passed-in settings say sound is on. Vibration is
// Android-only (iOS Safari has no Vibration API — this is a silent no-op
// there) and is intentionally very short.
export function playTap(settings) {
  if (!settings?.enabled) return;
  testSound(settings.type, settings.pitch);
  try { if (navigator.vibrate) navigator.vibrate(8); } catch { /* ignore */ }
}
