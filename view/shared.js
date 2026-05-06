// ═══════════════════════════════════════════════════════════════════
//  shared.js — Shared utilities for Kanban/Gantt organizer
// ═══════════════════════════════════════════════════════════════════

// ── HTML Escape ──────────────────────────────────────────
function esc(s) {
  const d = document.createElement('div');
  d.textContent = s || '';
  return d.innerHTML;
}

// ── Date Helpers ─────────────────────────────────────────
function pad(n) { return n < 10 ? '0' + n : '' + n; }

function fmtDate(d) {
  return d.getFullYear() + '-' + pad(d.getMonth() + 1) + '-' + pad(d.getDate());
}

function parseDate(s) {
  if (!s) return new Date();
  const p = s.split('-');
  return new Date(parseInt(p[0], 10), parseInt(p[1], 10) - 1, parseInt(p[2], 10));
}

function addDays(d, n) {
  const r = new Date(d.getTime());
  r.setDate(r.getDate() + n);
  return r;
}

function daysBetween(a, b) {
  return Math.round((b.getTime() - a.getTime()) / 86400000);
}

// ── Math ─────────────────────────────────────────────────
function clamp(v, lo, hi) { return v < lo ? lo : v > hi ? hi : v; }

// ── Color Helpers ────────────────────────────────────────
function lightenColor(hex, ratio) {
  hex = String(hex).replace(/^#/, '');
  if (hex.length === 3) hex = hex[0]+hex[0]+hex[1]+hex[1]+hex[2]+hex[2];
  const r = Math.round(parseInt(hex.substring(0,2),16) + (255 - parseInt(hex.substring(0,2),16)) * ratio);
  const g = Math.round(parseInt(hex.substring(2,4),16) + (255 - parseInt(hex.substring(2,4),16)) * ratio);
  const b = Math.round(parseInt(hex.substring(4,6),16) + (255 - parseInt(hex.substring(4,6),16)) * ratio);
  return 'rgb('+r+','+g+','+b+')';
}

function contrastColor(hex) {
  hex = String(hex).replace(/^#/, '');
  if (hex.length === 3) hex = hex[0]+hex[0]+hex[1]+hex[1]+hex[2]+hex[2];
  const r = parseInt(hex.substring(0,2),16);
  const g = parseInt(hex.substring(2,4),16);
  const b = parseInt(hex.substring(4,6),16);
  const lum = (0.299 * r + 0.587 * g + 0.114 * b) / 255;
  return lum > 0.5 ? '#1a1a2e' : '#fff';
}

// ── Modal Helpers ────────────────────────────────────────
function openModal(id) { document.getElementById(id).classList.add('open'); }
function closeModal(id) { document.getElementById(id).classList.remove('open'); }

// ── ParentBridge ─────────────────────────────────────────
class ParentBridge {
  constructor() {
    this.inIframe = (window.parent !== window);
  }

  persist(action, payload) {
    const json = JSON.stringify({ action, payload });
    if (this.inIframe && window.parent && window.parent.setPendingCommand) {
      window.parent.setPendingCommand(json);
    }
    this.setStatus('Saved');
  }

  setStatus(msg, isErr) {
    if (this.inIframe && window.parent && window.parent.setStatus) {
      window.parent.setStatus(msg, isErr);
    }
  }
}
