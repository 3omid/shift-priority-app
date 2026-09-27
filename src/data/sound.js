// Short, synthesized tap sounds — no audio files. Gated entirely by the
// caller checking settings.sound.enabled; this module just knows how to
// make the two sounds and (on Android) pair them with a short vibration.

let ctx = null;
function audioCtx() {
  if (!ctx) ctx = new (window.AudioContext || window.webkitAudioContext)();
  if (ctx.state === "suspended") ctx.resume();
  return ctx;
}

function pop(volume = 0.4) {
  const c = audioCtx(), t = c.currentTime;
  const o = c.createOscillator(), g = c.createGain();
  o.type = "sine";
  o.frequency.setValueAtTime(1100, t);
  o.frequency.exponentialRampToValueAtTime(320, t + 0.04);
  g.gain.setValueAtTime(0.0001, t);
  g.gain.exponentialRampToValueAtTime(volume * 0.5, t + 0.004);
  g.gain.exponentialRampToValueAtTime(0.0001, t + 0.05);
  o.connect(g).connect(c.destination);
  o.start(t); o.stop(t + 0.06);
}

function bubble(volume = 0.4) {
  const c = audioCtx(), t = c.currentTime;
  const o = c.createOscillator(), g = c.createGain();
  o.type = "sine";
  o.frequency.setValueAtTime(380, t);
  o.frequency.exponentialRampToValueAtTime(950, t + 0.07);
  g.gain.setValueAtTime(0.0001, t);
  g.gain.exponentialRampToValueAtTime(volume * 0.45, t + 0.01);
  g.gain.exponentialRampToValueAtTime(0.0001, t + 0.09);
  o.connect(g).connect(c.destination);
  o.start(t); o.stop(t + 0.1);
}

const SOUNDS = { pop, bubble };

// Plays the given type immediately — used by the Settings "test" buttons
// regardless of whether sound is currently enabled.
export function testSound(type) {
  try { (SOUNDS[type] || pop)(); } catch { /* Web Audio unavailable */ }
}

// Plays only if the passed-in settings say sound is on. Vibration is
// Android-only (iOS Safari has no Vibration API — this is a silent no-op
// there) and is intentionally very short.
export function playTap(settings) {
  if (!settings?.enabled) return;
  testSound(settings.type);
  try { if (navigator.vibrate) navigator.vibrate(8); } catch { /* ignore */ }
}
