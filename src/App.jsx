import { useState, useEffect } from "react";

// ─── Design tokens ────────────────────────────────────────────────────────────
const C = {
  bg: "#0f1117", surface: "#181c27", surfaceAlt: "#1e2335",
  border: "#2a3050", accent: "#4f8ef7", accentDim: "#2d5bbf",
  accentGlow: "rgba(79,142,247,0.18)",
  teams: "#6264a7", teamsGlow: "rgba(98,100,167,0.2)",
  outlook: "#0072c6", outlookGlow: "rgba(0,114,198,0.18)",
  text: "#e8ecf5", textMuted: "#8892b0", textDim: "#5a6480",
  danger: "#f07070", success: "#5dd88a", warn: "#f5a623", purple: "#a78bfa",
};

const uid = () => Math.random().toString(36).slice(2, 9);

// ─── EMU → pt conversions from docx inspection ───────────────────────────────
// 139700 EMU ≈ 11pt, 165100 ≈ 13pt, 177800 ≈ 14pt, null = default ~11pt body
// Font: Calibri throughout

// ─── localStorage cache helpers ───────────────────────────────────────────────
const CACHE_KEY  = "hr_field_cache_v1";
const DRAFT_KEY  = "hr_draft_v3";
const TPL_KEY    = "hr_custom_tpls_v1";

function readCache()  { try { return JSON.parse(localStorage.getItem(CACHE_KEY) || "{}"); } catch { return {}; } }
function writeCache(obj) { localStorage.setItem(CACHE_KEY, JSON.stringify(obj)); }
function patchCache(patch) { writeCache({ ...readCache(), ...patch }); }

function loadDraft()  { try { return JSON.parse(localStorage.getItem(DRAFT_KEY) || "null"); } catch { return null; } }
function saveDraft(d) { localStorage.setItem(DRAFT_KEY, JSON.stringify(d)); }

function loadCustomTpls()   { try { return JSON.parse(localStorage.getItem(TPL_KEY) || "[]"); } catch { return []; } }
function saveCustomTpls(t)  { localStorage.setItem(TPL_KEY, JSON.stringify(t)); }

// ─── Built-in templates ───────────────────────────────────────────────────────
const BUILTIN_TPLS = [
  { id: "general",   label: "General Interview",  icon: "💼", notes: "" },
  { id: "technical", label: "Technical Round",    icon: "⚙️", notes: "" },
  { id: "hr",        label: "HR Screening",       icon: "🤝", notes: "" },
  { id: "final",     label: "Final Round",        icon: "🏆", notes: "" },
  { id: "panel",     label: "Panel Interview",    icon: "👥", notes: "" },
];

const defaultCandidate = () => ({ id: uid(), name: "", email: "", date: "", startTime: "", endTime: "", showEndTimeInEmail: false, room: "", roomUrl: "" });

function escapeHtml(s) {
  return String(s ?? "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

// Word/Outlook font sizes: 2≈10pt, 3≈12pt, 4≈14pt — paired with inline pt for paste.
const O = {
  s11: "font-size:11.0pt;font-family:Calibri,sans-serif;mso-ansi-font-size:11.0pt;mso-bidi-font-size:11.0pt;",
  s13b: "font-size:13.0pt;font-family:Calibri,sans-serif;mso-ansi-font-size:13.0pt;mso-bidi-font-size:13.0pt;font-weight:bold;",
  s14bi: "font-size:14.0pt;font-family:Calibri,sans-serif;mso-ansi-font-size:14.0pt;mso-bidi-font-size:14.0pt;font-weight:bold;font-style:italic;",
  f11: 3,
  f13: 4,
  f14: 5,
};

function outlookP(margin, style, html, wordSize) {
  const fOpen = wordSize != null ? `<font face="Calibri" size="${wordSize}">` : "";
  const fClose = wordSize != null ? "</font>" : "";
  return `<p style="margin:${margin};line-height:1.5;${style}">${fOpen}${html}${fClose}</p>`;
}

function wrapClipboardHtml(fragment) {
  return `<html xmlns:w="urn:schemas-microsoft-com:office:word"><head><meta charset="utf-8"></head><body>${fragment}</body></html>`;
}

function htmlLink(url, text, linkStyle = O.s11, wordSize = O.f11) {
  const u = (url || "").trim();
  if (!u) return escapeHtml(text || "");
  const t = (text || "").trim() || u;
  const fOpen = `<font face="Calibri" size="${wordSize}">`;
  return `<a href="${escapeHtml(u)}" style="color:#0072c6;${linkStyle}">${fOpen}${escapeHtml(t)}</font></a>`;
}

function formatPanel(interviewers) {
  const list = Array.isArray(interviewers)
    ? interviewers.map(s => String(s).trim()).filter(Boolean)
    : typeof interviewers === "string" && interviewers.trim()
      ? [interviewers.trim()]
      : [];
  if (!list.length) return { html: escapeHtml("[Panel Members]"), plain: "[Panel Members]" };
  const plain = list.join(", ");
  return { html: list.map(escapeHtml).join(", "), plain };
}

function captureFormConfig({ mode, companyName, interviewStage, role, interviewers, formLink, formLinkText, additionalNotes, cc, bcc }) {
  return { mode, companyName, interviewStage, role, interviewers, formLink, formLinkText, additionalNotes, cc, bcc };
}

function describeTemplateConfig(cfg) {
  if (!cfg) return null;
  const parts = [];
  if (cfg.role) parts.push(cfg.role);
  if (cfg.interviewStage) parts.push(cfg.interviewStage);
  parts.push(cfg.mode === "virtual" ? "Virtual" : "Physical");
  if (cfg.companyName) parts.push(cfg.companyName);
  if (cfg.interviewers?.length) parts.push(`${cfg.interviewers.length} on panel`);
  return parts.join(" · ");
}

// ─── Formatters ───────────────────────────────────────────────────────────────
function fmtDate(iso) {
  if (!iso) return "";
  const [y, m, d] = iso.split("-");
  return `${d}/${m}/${y}`;
}
function fmtTime(t) {
  if (!t) return "";
  const [h, m] = t.split(":").map(Number);
  const ap = h >= 12 ? "PM" : "AM";
  return `${String(h % 12 || 12).padStart(2, "0")}:${String(m).padStart(2, "0")} ${ap}`;
}
function firstName(full) {
  const t = (full || "").trim();
  if (!t) return "[Candidate Name]";
  return t.split(/\s+/)[0];
}

function copyViaRenderedHtml(bodyHtml) {
  try {
    const el = document.createElement("div");
    el.contentEditable = "true";
    el.setAttribute("aria-hidden", "true");
    Object.assign(el.style, {
      position: "fixed",
      left: "-9999px",
      top: "0",
      width: "7.5in",
      background: "#fff",
      color: "#000",
      fontFamily: "Calibri, sans-serif",
      lineHeight: "1.5",
    });
    el.innerHTML = bodyHtml;
    document.body.appendChild(el);
    const range = document.createRange();
    range.selectNodeContents(el);
    const sel = window.getSelection();
    sel.removeAllRanges();
    sel.addRange(range);
    const ok = document.execCommand("copy");
    document.body.removeChild(el);
    sel.removeAllRanges();
    return ok;
  } catch {
    return false;
  }
}

async function copyEmailToClipboard(bodyHtml, bodyPlain) {
  if (copyViaRenderedHtml(bodyHtml)) return "rendered";
  const html = wrapClipboardHtml(bodyHtml.trim());
  try {
    if (navigator.clipboard?.write && typeof ClipboardItem !== "undefined") {
      await navigator.clipboard.write([
        new ClipboardItem({
          "text/html": new Blob([html], { type: "text/html" }),
          "text/plain": new Blob([bodyPlain], { type: "text/plain" }),
        }),
      ]);
      return "html";
    }
  } catch { /* fallback below */ }
  try {
    await navigator.clipboard.writeText(bodyPlain);
    return "plain";
  } catch {
    return false;
  }
}

// ─── Icons ────────────────────────────────────────────────────────────────────
const Icon = {
  Teams:    () => <svg width="18" height="18" viewBox="0 0 24 24" fill="none"><path d="M17 4a3 3 0 1 1-6 0 3 3 0 0 1 6 0z" fill="#6264a7"/><path d="M20 9h-7a2 2 0 0 0-2 2v6a2 2 0 0 0 2 2h7a2 2 0 0 0 2-2v-6a2 2 0 0 0-2-2z" fill="#6264a7"/><circle cx="7" cy="8" r="2.5" fill="#8b8cc8"/><path d="M10 13H4a1.5 1.5 0 0 0-1.5 1.5v4A1.5 1.5 0 0 0 4 20h6a1.5 1.5 0 0 0 1.5-1.5V17" stroke="#8b8cc8" strokeWidth="1.2" fill="none"/></svg>,
  Outlook:  () => <svg width="18" height="18" viewBox="0 0 24 24" fill="none"><rect x="2" y="5" width="20" height="14" rx="2" fill="#0072c6" opacity="0.2" stroke="#0072c6" strokeWidth="1.5"/><path d="M2 8l10 7 10-7" stroke="#0072c6" strokeWidth="1.5" strokeLinecap="round"/></svg>,
  Plus:     () => <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round"><line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/></svg>,
  Trash:    () => <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><polyline points="3 6 5 6 21 6"/><path d="M19 6l-1 14H6L5 6"/><path d="M10 11v6M14 11v6"/><path d="M9 6V4h6v2"/></svg>,
  Send:     () => <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round"><line x1="22" y1="2" x2="11" y2="13"/><polygon points="22 2 15 22 11 13 2 9 22 2"/></svg>,
  Copy:     () => <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><rect x="9" y="9" width="13" height="13" rx="2"/><path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"/></svg>,
  Edit:     () => <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>,
  Save:     () => <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z"/><polyline points="17 21 17 13 7 13 7 21"/><polyline points="7 3 7 8 15 8"/></svg>,
  Check:    () => <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke={C.success} strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round"><polyline points="20 6 9 17 4 12"/></svg>,
  Warn:     () => <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke={C.warn} strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z"/><line x1="12" y1="9" x2="12" y2="13"/><line x1="12" y1="17" x2="12.01" y2="17"/></svg>,
  Draft:    () => <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/></svg>,
  X:        () => <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>,
  Template: () => <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><rect x="3" y="3" width="7" height="7"/><rect x="14" y="3" width="7" height="7"/><rect x="14" y="14" width="7" height="7"/><rect x="3" y="14" width="7" height="7"/></svg>,
  Clock:    () => <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round"><circle cx="12" cy="12" r="10"/><polyline points="12 6 12 12 16 14"/></svg>,
  Location: () => <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round"><path d="M21 10c0 7-9 13-9 13s-9-6-9-13a9 9 0 0 1 18 0z"/><circle cx="12" cy="10" r="3"/></svg>,
  Eye:      () => <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"/><circle cx="12" cy="12" r="3"/></svg>,
  ChevDown: () => <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round"><polyline points="6 9 12 15 18 9"/></svg>,
};

// ─── Primitives ───────────────────────────────────────────────────────────────
function Input({ label, value, onChange, placeholder, type = "text", style = {}, required, cacheKey }) {
  const [f, setF] = useState(false);
  const [suggestions, setSuggestions] = useState([]);
  const [showSug, setShowSug] = useState(false);

  // Load suggestions from cache on focus
  const onFocus = () => {
    setF(true);
    if (cacheKey) {
      const cache = readCache();
      const hist = cache[cacheKey] || [];
      setSuggestions(hist.filter(s => !value || s.toLowerCase().includes(value.toLowerCase())));
      if (hist.length) setShowSug(true);
    }
  };
  const onBlur = () => {
    setF(false);
    setTimeout(() => setShowSug(false), 150);
    if (cacheKey && value.trim()) {
      const cache = readCache();
      const hist = cache[cacheKey] || [];
      const updated = [value.trim(), ...hist.filter(s => s !== value.trim())].slice(0, 8);
      patchCache({ [cacheKey]: updated });
    }
  };
  const pickSug = (s) => { onChange(s); setShowSug(false); };

  const base = { background: C.surfaceAlt, border: `1.5px solid ${f ? C.accent : C.border}`, borderRadius: 8, padding: "9px 13px", color: C.text, fontSize: 13, outline: "none", width: "100%", boxSizing: "border-box", transition: "border-color .2s, box-shadow .2s", boxShadow: f ? `0 0 0 3px ${C.accentGlow}` : "none", fontFamily: "Calibri, Calibri Light, sans-serif" };

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 5, flex: 1, minWidth: 140, position: "relative" }}>
      {label && <label style={{ fontSize: 11, color: C.textMuted, fontWeight: 600, letterSpacing: 0.4 }}>{label}{required && <span style={{ color: C.danger }}> *</span>}</label>}
      <input type={type} value={value} placeholder={placeholder} onChange={e => onChange(e.target.value)} onFocus={onFocus} onBlur={onBlur} style={{ ...base, ...style }} autoComplete="off" />
      {showSug && suggestions.length > 0 && (
        <div style={{ position: "absolute", top: "100%", left: 0, right: 0, background: C.surface, border: `1px solid ${C.border}`, borderRadius: 8, zIndex: 200, overflow: "hidden", boxShadow: "0 8px 24px rgba(0,0,0,0.4)", marginTop: 2 }}>
          {suggestions.map((s, i) => (
            <div key={i} onMouseDown={() => pickSug(s)} style={{ padding: "9px 13px", fontSize: 13, color: C.textMuted, cursor: "pointer", borderBottom: i < suggestions.length - 1 ? `1px solid ${C.border}` : "none" }}
              onMouseEnter={e => e.currentTarget.style.background = C.surfaceAlt}
              onMouseLeave={e => e.currentTarget.style.background = "transparent"}>
              {s}
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

function Select({ label, value, onChange, options, style = {} }) {
  const [f, setF] = useState(false);
  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 5, flex: 1, minWidth: 140 }}>
      {label && <label style={{ fontSize: 11, color: C.textMuted, fontWeight: 600, letterSpacing: 0.4 }}>{label}</label>}
      <div style={{ position: "relative" }}>
        <select value={value} onChange={e => onChange(e.target.value)} onFocus={() => setF(true)} onBlur={() => setF(false)}
          style={{ background: C.surfaceAlt, border: `1.5px solid ${f ? C.accent : C.border}`, borderRadius: 8, padding: "9px 36px 9px 13px", color: C.text, fontSize: 13, outline: "none", width: "100%", boxSizing: "border-box", appearance: "none", cursor: "pointer", transition: "border-color .2s, box-shadow .2s", boxShadow: f ? `0 0 0 3px ${C.accentGlow}` : "none", fontFamily: "Calibri, sans-serif", ...style }}>
          {options.map(([v, l]) => <option key={v} value={v}>{l}</option>)}
        </select>
        <span style={{ position: "absolute", right: 12, top: "50%", transform: "translateY(-50%)", pointerEvents: "none", color: C.textDim }}><Icon.ChevDown /></span>
      </div>
    </div>
  );
}

function LinkField({ label, url, onUrlChange, text, onTextChange, urlPlaceholder, textPlaceholder, urlCacheKey, textCacheKey }) {
  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 5, flex: 1, minWidth: 280 }}>
      {label && <label style={{ fontSize: 11, color: C.textMuted, fontWeight: 600, letterSpacing: 0.4 }}>{label}</label>}
      <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
        <Input label="URL" value={url} onChange={onUrlChange} placeholder={urlPlaceholder} type="url" cacheKey={urlCacheKey} />
        <Input label="Link text" value={text} onChange={onTextChange} placeholder={textPlaceholder} cacheKey={textCacheKey} />
      </div>
    </div>
  );
}

function Textarea({ label, value, onChange, placeholder, minHeight = 90 }) {
  const [f, setF] = useState(false);
  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 5, flex: 1 }}>
      {label && <label style={{ fontSize: 11, color: C.textMuted, fontWeight: 600, letterSpacing: 0.4 }}>{label}</label>}
      <textarea value={value} placeholder={placeholder} onChange={e => onChange(e.target.value)} onFocus={() => setF(true)} onBlur={() => setF(false)}
        style={{ background: C.surfaceAlt, border: `1.5px solid ${f ? C.accent : C.border}`, borderRadius: 8, padding: "10px 13px", color: C.text, fontSize: 13, outline: "none", width: "100%", boxSizing: "border-box", resize: "vertical", minHeight, lineHeight: 1.65, fontFamily: "Calibri, sans-serif", transition: "border-color .2s, box-shadow .2s", boxShadow: f ? `0 0 0 3px ${C.accentGlow}` : "none" }}
      />
    </div>
  );
}

function TagInput({ label, values, onChange, placeholder, cacheKey }) {
  const [raw, setRaw] = useState("");
  const [suggestions, setSuggestions] = useState([]);
  const [showSug, setShowSug] = useState(false);

  const add = (val) => {
    const v = (val || raw).trim();
    if (v && !values.includes(v)) {
      const next = [...values, v];
      onChange(next);
      if (cacheKey) {
        const cache = readCache();
        const hist = cache[cacheKey] || [];
        const updated = [v, ...hist.filter(s => s !== v)].slice(0, 12);
        patchCache({ [cacheKey]: updated });
      }
    }
    setRaw(""); setShowSug(false);
  };

  const onFocus = () => {
    if (cacheKey) {
      const cache = readCache();
      const hist = (cache[cacheKey] || []).filter(s => !values.includes(s));
      setSuggestions(hist); if (hist.length) setShowSug(true);
    }
  };

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 5, flex: 1, position: "relative" }}>
      {label && <label style={{ fontSize: 11, color: C.textMuted, fontWeight: 600, letterSpacing: 0.4 }}>{label}</label>}
      <div style={{ display: "flex", flexWrap: "wrap", gap: 5, background: C.surfaceAlt, border: `1.5px solid ${C.border}`, borderRadius: 8, padding: "7px 9px", minHeight: 40, alignItems: "center" }}>
        {values.map((v, i) => (
          <span key={i} style={{ display: "flex", alignItems: "center", gap: 4, background: C.surface, border: `1px solid ${C.border}`, borderRadius: 5, padding: "2px 8px", fontSize: 12, color: C.textMuted }}>
            {v}<button onClick={() => onChange(values.filter((_, j) => j !== i))} style={{ background: "none", border: "none", cursor: "pointer", color: C.textDim, padding: 0, lineHeight: 1, fontSize: 14 }}>×</button>
          </span>
        ))}
        <input value={raw} placeholder={values.length ? "" : placeholder} onChange={e => setRaw(e.target.value)}
          onFocus={onFocus}
          onBlur={() => { setTimeout(() => setShowSug(false), 150); if (raw.trim()) add(raw.trim()); }}
          onKeyDown={e => { if (e.key === "Enter" || e.key === ",") { e.preventDefault(); add(); } }}
          style={{ border: "none", background: "transparent", outline: "none", color: C.text, fontSize: 13, flex: 1, minWidth: 130, padding: "2px 3px", fontFamily: "Calibri, sans-serif" }} />
      </div>
      <span style={{ fontSize: 11, color: C.textDim }}>Press Enter, comma, or click away to add</span>
      {showSug && suggestions.length > 0 && (
        <div style={{ position: "absolute", top: "100%", left: 0, right: 0, background: C.surface, border: `1px solid ${C.border}`, borderRadius: 8, zIndex: 200, overflow: "hidden", boxShadow: "0 8px 24px rgba(0,0,0,0.4)", marginTop: 2 }}>
          {suggestions.map((s, i) => (
            <div key={i} onMouseDown={() => add(s)} style={{ padding: "9px 13px", fontSize: 13, color: C.textMuted, cursor: "pointer", borderBottom: i < suggestions.length - 1 ? `1px solid ${C.border}` : "none" }}
              onMouseEnter={e => e.currentTarget.style.background = C.surfaceAlt}
              onMouseLeave={e => e.currentTarget.style.background = "transparent"}>
              {s}
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

function Btn({ children, onClick, variant = "primary", color, style: sx = {}, title, disabled }) {
  const [hov, setHov] = useState(false);
  const col = color || (variant === "primary" ? C.accent : variant === "danger" ? C.danger : variant === "success" ? C.success : C.textMuted);
  const base = { padding: "9px 18px", borderRadius: 8, border: "none", cursor: disabled ? "not-allowed" : "pointer", fontSize: 13, fontWeight: 600, display: "flex", alignItems: "center", gap: 6, transition: "all .15s", letterSpacing: 0.2, opacity: disabled ? 0.5 : 1, userSelect: "none" };
  const variants = {
    primary: { background: hov ? col + "dd" : col, color: "#fff" },
    ghost:   { background: hov ? col + "18" : "transparent", color: col, border: `1px solid ${col}44` },
    danger:  { background: hov ? C.danger + "18" : "transparent", color: C.danger, border: "1px solid transparent" },
    subtle:  { background: hov ? C.surfaceAlt : "transparent", color: C.textMuted, border: "1px solid transparent" },
  };
  return <button onClick={disabled ? undefined : onClick} onMouseEnter={() => setHov(true)} onMouseLeave={() => setHov(false)} style={{ ...base, ...variants[variant], ...sx }} title={title}>{children}</button>;
}

// ─── Save Template Modal ──────────────────────────────────────────────────────
function SaveTemplateModal({ initial, currentConfig, onSave, onClose }) {
  const [label, setLabel] = useState(initial?.label || "");
  const [icon, setIcon]   = useState(initial?.icon || "📋");
  const isUpdate = !!initial?.id;

  const commit = () => {
    if (!label.trim()) { alert("Template name is required."); return; }
    onSave({
      id: initial?.id || uid(),
      label: label.trim(),
      icon: icon.trim() || "📋",
      config: currentConfig,
    }, !isUpdate);
    onClose();
  };

  const summary = describeTemplateConfig(currentConfig);

  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.75)", zIndex: 620, display: "flex", alignItems: "center", justifyContent: "center", padding: 20 }}>
      <div style={{ background: C.surface, border: `1px solid ${C.border}`, borderRadius: 18, width: "100%", maxWidth: 480, overflow: "hidden" }}>
        <div style={{ padding: "22px 28px 16px", borderBottom: `1px solid ${C.border}`, display: "flex", justifyContent: "space-between", alignItems: "center" }}>
          <div>
            <div style={{ fontSize: 17, fontWeight: 700 }}>{isUpdate ? "Update Template" : "Save as Template"}</div>
            <div style={{ fontSize: 12, color: C.textMuted, marginTop: 3 }}>Saves interview type, details, panel, CC/BCC, and notes. Candidates are not included.</div>
          </div>
          <Btn variant="subtle" onClick={onClose}><Icon.X /></Btn>
        </div>
        <div style={{ padding: "20px 28px 24px" }}>
          <div style={{ display: "flex", gap: 10, marginBottom: 12 }}>
            <Input label="Icon" value={icon} onChange={setIcon} placeholder="Emoji" style={{ maxWidth: 70 }} />
            <Input label="Template Name *" value={label} onChange={setLabel} placeholder="e.g. Senior Engineer – Final Round" />
          </div>
          {summary && (
            <div style={{ padding: "10px 14px", background: C.surfaceAlt, border: `1px solid ${C.border}`, borderRadius: 8, fontSize: 12, color: C.textMuted, marginBottom: 16, lineHeight: 1.55 }}>
              <strong style={{ color: C.text }}>Will save:</strong> {summary}
            </div>
          )}
          <div style={{ display: "flex", gap: 8, justifyContent: "flex-end" }}>
            <Btn variant="ghost" onClick={onClose}>Cancel</Btn>
            <Btn variant="primary" onClick={commit}><Icon.Save /> {isUpdate ? "Update Template" : "Save Template"}</Btn>
          </div>
        </div>
      </div>
    </div>
  );
}

// ─── Template Manager ─────────────────────────────────────────────────────────
function TemplateManager({ builtins, customs, onSave, onDelete, onClose, onSelectAndClose, getCurrentConfig }) {
  const [editing, setEditing] = useState(null);
  const [tab, setTab]         = useState("custom");

  const startNew = () => setEditing({
    id: uid(), label: "", icon: "📋", notes: "", config: getCurrentConfig(), isNew: true,
  });
  const startEdit = (t) => setEditing({ ...t, notes: t.notes ?? t.config?.additionalNotes ?? "", isNew: false });
  const pullConfig = () => setEditing(e => e ? { ...e, config: getCurrentConfig() } : e);

  const commit = () => {
    if (!editing) return;
    const cfg = editing.config || getCurrentConfig();
    onSave({
      id: editing.id,
      label: editing.label || "Untitled",
      icon: editing.icon || "📋",
      notes: editing.notes,
      config: cfg,
    }, editing.isNew);
    setEditing(null);
  };

  const list = tab === "builtin" ? builtins : customs;

  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.75)", zIndex: 600, display: "flex", alignItems: "center", justifyContent: "center", padding: 20 }}>
      <div style={{ background: C.surface, border: `1px solid ${C.border}`, borderRadius: 18, width: "100%", maxWidth: 680, maxHeight: "88vh", display: "flex", flexDirection: "column", overflow: "hidden" }}>
        <div style={{ padding: "22px 28px 16px", borderBottom: `1px solid ${C.border}`, display: "flex", justifyContent: "space-between", alignItems: "center" }}>
          <div><div style={{ fontSize: 17, fontWeight: 700 }}>Template Library</div><div style={{ fontSize: 12, color: C.textMuted, marginTop: 3 }}>Save and reuse full interview setups. Built-in tags only store extra notes unless you update them.</div></div>
          <div style={{ display: "flex", gap: 8 }}><Btn variant="primary" onClick={startNew}><Icon.Plus /> New from Current</Btn><Btn variant="subtle" onClick={onClose}><Icon.X /></Btn></div>
        </div>
        {editing && (
          <div style={{ padding: "20px 28px", borderBottom: `1px solid ${C.border}`, background: C.surfaceAlt }}>
            <div style={{ fontSize: 12, color: C.accent, fontWeight: 700, letterSpacing: 1, textTransform: "uppercase", marginBottom: 12 }}>{editing.isNew ? "✦ New Template" : `Editing — ${editing.label}`}</div>
            <div style={{ display: "flex", gap: 10, marginBottom: 10 }}>
              <Input label="Icon" value={editing.icon} onChange={v => setEditing(e => ({ ...e, icon: v }))} placeholder="Emoji" style={{ maxWidth: 70 }} />
              <Input label="Template Name" value={editing.label} onChange={v => setEditing(e => ({ ...e, label: v }))} placeholder="e.g. Culture Fit Interview" />
            </div>
            <Textarea label="Additional Notes (optional)" value={editing.notes} onChange={v => setEditing(e => ({ ...e, notes: v }))} placeholder="Extra notes appended to the invite body…" minHeight={70} />
            {describeTemplateConfig(editing.config) && (
              <div style={{ fontSize: 12, color: C.textMuted, marginTop: 8, padding: "8px 12px", background: C.surface, borderRadius: 8, border: `1px solid ${C.border}` }}>
                <strong style={{ color: C.text }}>Saved settings:</strong> {describeTemplateConfig(editing.config)}
              </div>
            )}
            <div style={{ display: "flex", gap: 8, marginTop: 12, justifyContent: "flex-end", flexWrap: "wrap" }}>
              <Btn variant="ghost" onClick={pullConfig} title="Replace saved interview fields with what's on the form now">Refresh from Form</Btn>
              <Btn variant="ghost" onClick={() => setEditing(null)}>Cancel</Btn>
              <Btn variant="primary" onClick={commit}><Icon.Save /> Save Template</Btn>
            </div>
          </div>
        )}
        <div style={{ display: "flex", borderBottom: `1px solid ${C.border}`, padding: "0 28px" }}>
          {[["custom", `My Templates (${customs.length})`], ["builtin", `Built-in (${builtins.length})`]].map(([k, lbl]) => (
            <button key={k} onClick={() => setTab(k)} style={{ padding: "12px 0", marginRight: 24, background: "none", border: "none", cursor: "pointer", fontSize: 13, fontWeight: 600, color: tab === k ? C.accent : C.textDim, borderBottom: `2px solid ${tab === k ? C.accent : "transparent"}`, transition: "all .15s" }}>{lbl}</button>
          ))}
        </div>
        <div style={{ overflowY: "auto", flex: 1, padding: "16px 28px 24px" }}>
          {list.map(tpl => {
            const cfgSummary = describeTemplateConfig(tpl.config);
            const notePreview = tpl.config?.additionalNotes || tpl.notes;
            return (
              <div key={tpl.id} style={{ display: "flex", alignItems: "flex-start", gap: 12, padding: "14px 16px", borderRadius: 10, border: `1px solid ${C.border}`, marginBottom: 10, background: C.surfaceAlt }}>
                <span style={{ fontSize: 22, flexShrink: 0 }}>{tpl.icon}</span>
                <div style={{ flex: 1, minWidth: 0 }}>
                  <div style={{ fontSize: 14, fontWeight: 600, marginBottom: 4 }}>{tpl.label}</div>
                  {cfgSummary && <div style={{ fontSize: 12, color: C.accent, marginBottom: 4 }}>{cfgSummary}</div>}
                  <div style={{ fontSize: 12, color: C.textMuted, lineHeight: 1.55 }}>{notePreview || <em style={{ color: C.textDim }}>{cfgSummary ? "No extra notes" : "Notes-only tag"}</em>}</div>
                </div>
                <div style={{ display: "flex", gap: 6, flexShrink: 0 }}>
                  <Btn variant="ghost" onClick={() => onSelectAndClose(tpl)} style={{ padding: "7px 12px", fontSize: 12 }}>Use</Btn>
                  <Btn variant="ghost" onClick={() => startEdit(tpl)} style={{ padding: "7px 10px" }}><Icon.Edit /></Btn>
                  {tab === "custom" && <Btn variant="danger" onClick={() => onDelete(tpl.id)} style={{ padding: "7px 10px" }}><Icon.Trash /></Btn>}
                </div>
              </div>
            );
          })}
          {list.length === 0 && <div style={{ textAlign: "center", color: C.textDim, padding: "40px 0", fontSize: 13 }}>{tab === "custom" ? 'No saved templates yet. Fill in the form and click "Save as Template", or use "New from Current".' : "No templates."}</div>}
        </div>
      </div>
    </div>
  );
}

// ─── Candidate Card ───────────────────────────────────────────────────────────
function CandidateCard({ c, idx, mode, onChange, onRemove, canRemove, conflictIds }) {
  const timeWarn    = c.startTime && c.endTime && c.startTime >= c.endTime;
  const hasConflict = conflictIds.has(c.id);
  const borderCol   = timeWarn || hasConflict ? C.warn + "88" : C.border;
  const accentCol   = mode === "virtual" ? C.teams : C.outlook;
  return (
    <div style={{ background: C.surfaceAlt, border: `1.5px solid ${borderCol}`, borderRadius: 12, padding: "18px 20px", marginBottom: 10, transition: "border-color .2s" }}>
      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 14 }}>
        <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
          <span style={{ fontSize: 11, fontWeight: 700, letterSpacing: 1, color: accentCol, textTransform: "uppercase" }}>Candidate {idx + 1}</span>
          {timeWarn    && <span style={{ display: "flex", alignItems: "center", gap: 4, fontSize: 11, color: C.warn }}><Icon.Warn /> End time before start time</span>}
          {hasConflict && !timeWarn && <span style={{ display: "flex", alignItems: "center", gap: 4, fontSize: 11, color: C.warn }}><Icon.Warn /> Time slot overlaps</span>}
        </div>
        <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
          <span style={{ fontSize: 11, fontWeight: 700, padding: "3px 10px", borderRadius: 20, letterSpacing: 0.6, textTransform: "uppercase", background: mode === "virtual" ? C.teamsGlow : C.outlookGlow, color: mode === "virtual" ? "#8b8cc8" : "#4da6e8", border: `1px solid ${mode === "virtual" ? C.teams + "44" : C.outlook + "44"}` }}>
            {mode === "virtual" ? "Virtual · Teams" : "Physical · Outlook"}
          </span>
          {canRemove && <Btn variant="danger" onClick={onRemove} style={{ padding: "5px 8px" }}><Icon.Trash /></Btn>}
        </div>
      </div>
      <div style={{ display: "flex", gap: 12, flexWrap: "wrap", marginBottom: 10 }}>
        <Input label="Candidate Name" value={c.name} onChange={v => onChange("name", v)} placeholder="e.g. Samantha Lee" required cacheKey="cand_name" />
        <Input label="Email Address" value={c.email} onChange={v => onChange("email", v)} placeholder="candidate@email.com" type="email" required cacheKey="cand_email" />
      </div>
      <div style={{ display: "flex", gap: 12, flexWrap: "wrap" }}>
        <Input label="Date" value={c.date} onChange={v => onChange("date", v)} type="date" style={{ maxWidth: 190 }} required />
        <Input label="Start Time" value={c.startTime} onChange={v => onChange("startTime", v)} type="time" style={{ maxWidth: 150 }} required />
        <div style={{ display: "flex", flexDirection: "column", gap: 6, maxWidth: 170 }}>
          <Input label="End Time (optional)" value={c.endTime} onChange={v => { onChange("endTime", v); if (!v) onChange("showEndTimeInEmail", false); }} type="time" style={{ maxWidth: 170 }} />
          <label style={{ display: "flex", alignItems: "center", gap: 7, fontSize: 11, color: c.endTime ? C.textMuted : C.textDim, cursor: c.endTime ? "pointer" : "default", userSelect: "none", paddingLeft: 2 }}>
            <input
              type="checkbox"
              checked={!!c.showEndTimeInEmail}
              disabled={!c.endTime}
              onChange={e => onChange("showEndTimeInEmail", e.target.checked)}
              style={{ width: 14, height: 14, accentColor: accentCol, cursor: c.endTime ? "pointer" : "not-allowed" }}
            />
            Show end time in email
          </label>
        </div>
        {mode === "physical" && (
          <LinkField
            label="Interview Location"
            url={c.roomUrl || ""}
            onUrlChange={v => onChange("roomUrl", v)}
            text={c.room || ""}
            onTextChange={v => onChange("room", v)}
            urlPlaceholder="https://maps.google.com/..."
            textPlaceholder="e.g. 20th Floor, Cinnamon Life"
            urlCacheKey="cand_room_url"
            textCacheKey="cand_room"
          />
        )}
      </div>
    </div>
  );
}

// ─── Email Preview Modal ──────────────────────────────────────────────────────
// Renders with Calibri font matching the docx template exactly
function PreviewModal({ items, onClose }) {
  const [idx, setIdx] = useState(0);
  const [copied, setCopied] = useState(false);
  const total = items.length;
  const item = items[idx];
  const copy = async () => {
    const ok = await copyEmailToClipboard(item.bodyHtml, item.bodyPlain);
    setCopied(ok ? "html" : "plain");
    setTimeout(() => setCopied(false), 1600);
  };

  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.78)", zIndex: 650, display: "flex", alignItems: "center", justifyContent: "center", padding: 20 }}>
      <div style={{ background: C.surface, border: `1px solid ${C.border}`, borderRadius: 18, width: "100%", maxWidth: 700, maxHeight: "86vh", display: "flex", flexDirection: "column", overflow: "hidden" }}>
        <div style={{ padding: "18px 24px 14px", borderBottom: `1px solid ${C.border}`, display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 10 }}>
          <div>
            <div style={{ fontSize: 15, fontWeight: 700 }}>📧 Email Preview</div>
            {total > 1 && <div style={{ fontSize: 12, color: C.textMuted, marginTop: 2 }}>Candidate {idx + 1} of {total}</div>}
          </div>
          <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
            {total > 1 && (
              <>
                <Btn variant="ghost" onClick={() => setIdx(i => Math.max(0, i - 1))} disabled={idx === 0} style={{ padding: "7px 12px", fontSize: 12 }}>← Prev</Btn>
                <Btn variant="ghost" onClick={() => setIdx(i => Math.min(total - 1, i + 1))} disabled={idx === total - 1} style={{ padding: "7px 12px", fontSize: 12 }}>Next →</Btn>
              </>
            )}
            <Btn variant="ghost" onClick={copy} style={{ fontSize: 12, padding: "7px 12px" }}><Icon.Copy />{copied ? "Copied!" : "Copy for Outlook"}</Btn>
            <Btn variant="subtle" onClick={onClose}><Icon.X /></Btn>
          </div>
        </div>
        {total > 1 && (
          <div style={{ display: "flex", gap: 6, padding: "10px 24px", borderBottom: `1px solid ${C.border}`, overflowX: "auto", flexWrap: "wrap" }}>
            {items.map((it, i) => (
              <button key={i} onClick={() => setIdx(i)} style={{ padding: "6px 12px", borderRadius: 8, border: `1.5px solid ${i === idx ? C.accent : C.border}`, background: i === idx ? C.accentGlow : C.surfaceAlt, color: i === idx ? C.accent : C.textMuted, cursor: "pointer", fontSize: 12, fontWeight: 600, whiteSpace: "nowrap" }}>
                {it.label}
              </button>
            ))}
          </div>
        )}
        <div style={{ padding: "18px 24px", overflowY: "auto", flex: 1 }}>
          {item.email && <div style={{ fontSize: 12, color: C.textDim, marginBottom: 12 }}>{item.email}</div>}
          <div style={{ fontSize: 11, color: C.textMuted, fontWeight: 600, letterSpacing: 0.5, marginBottom: 4 }}>SUBJECT</div>
          <div style={{ fontSize: 13, fontFamily: "Calibri, sans-serif", color: C.text, background: C.surfaceAlt, borderRadius: 8, padding: "10px 14px", marginBottom: 16 }} dangerouslySetInnerHTML={{ __html: item.subjectHtml }} />
          <div style={{ fontSize: 11, color: C.textMuted, fontWeight: 600, letterSpacing: 0.5, marginBottom: 4 }}>BODY</div>
          <div style={{ fontSize: 14, fontFamily: "Calibri, sans-serif", color: "#1a1a1a", background: "#ffffff", borderRadius: 8, padding: "20px 24px", lineHeight: 1.7 }} dangerouslySetInnerHTML={{ __html: item.bodyHtml }} />
        </div>
      </div>
    </div>
  );
}

// ─── Dispatch Modal ───────────────────────────────────────────────────────────
function DispatchModal({ items, mode, onClose }) {
  const [current, setCurrent] = useState(0);
  const [opened, setOpened]   = useState(new Set());
  const [copied, setCopied]   = useState(null);
  const total    = items.length;
  const item     = items[current];
  const allDone  = opened.size === total;
  const modeCol  = mode === "virtual" ? C.teams : C.outlook;
  const modeGlow = mode === "virtual" ? C.teamsGlow : C.outlookGlow;

  const openDraft = async (idx) => {
    const it = items[idx];
    const copied = await copyEmailToClipboard(it.bodyHtml, it.bodyPlain);
    if (mode === "physical") {
      const enc = (s) => encodeURIComponent(s).replace(/%0A/g, "%0D%0A");
      const cc  = it.cc.length  ? `&cc=${encodeURIComponent(it.cc.join(";"))}`  : "";
      const bcc = it.bcc.length ? `&bcc=${encodeURIComponent(it.bcc.join(";"))}` : "";
      const body = copied === false ? `&body=${enc(it.bodyPlain)}` : "";
      window.location.href =
        `mailto:${encodeURIComponent(it.email)}?subject=${enc(it.subjectPlain)}${cc}${bcc}${body}`;
    } else {
      window.open("https://teams.microsoft.com/l/meeting/new", "_blank", "noopener,noreferrer");
    }
    setCopied(idx);
    setOpened(prev => new Set([...prev, idx]));
  };

  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.82)", zIndex: 700, display: "flex", alignItems: "center", justifyContent: "center", padding: 20 }}>
      <div style={{ background: C.surface, border: `1px solid ${C.border}`, borderRadius: 18, width: "100%", maxWidth: 600, display: "flex", flexDirection: "column", overflow: "hidden" }}>
        <div style={{ padding: "22px 26px 16px", borderBottom: `1px solid ${C.border}`, display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
          <div>
            <div style={{ fontSize: 17, fontWeight: 700 }}>{allDone ? "✅ All Invites Dispatched" : mode === "virtual" ? "📅 Dispatching Teams Invites" : "📨 Dispatching Outlook Invites"}</div>
            <div style={{ fontSize: 12, color: C.textMuted, marginTop: 3 }}>
              {allDone ? "All drafts opened — make sure you've sent each one." : mode === "virtual"
                ? "Formatted invite copied — paste into the Teams meeting description (Ctrl+V)."
                : "Formatted body copied and your mail app opens with To, CC, BCC, and subject — paste the body (Ctrl+V)."}
            </div>
          </div>
          <Btn variant="subtle" onClick={onClose}><Icon.X /></Btn>
        </div>
        <div style={{ height: 3, background: C.border }}>
          <div style={{ height: "100%", width: `${(opened.size / total) * 100}%`, background: modeCol, transition: "width .4s ease" }} />
        </div>
        <div style={{ padding: "18px 26px", overflowY: "auto", maxHeight: 300 }}>
          {items.map((it, i) => {
            const done = opened.has(i); const active = i === current && !allDone;
            return (
              <div key={i} onClick={() => setCurrent(i)} style={{ display: "flex", alignItems: "center", gap: 12, padding: "12px 14px", borderRadius: 10, marginBottom: 8, cursor: "pointer", border: `1.5px solid ${active ? modeCol : done ? C.success + "55" : C.border}`, background: active ? modeGlow : done ? "rgba(93,216,138,0.05)" : C.surfaceAlt, transition: "all .15s" }}>
                <div style={{ width: 28, height: 28, borderRadius: "50%", flexShrink: 0, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 13, fontWeight: 700, background: done ? "rgba(93,216,138,.15)" : active ? modeCol + "22" : C.surface, border: `1.5px solid ${done ? C.success : active ? modeCol : C.border}`, color: done ? C.success : active ? modeCol : C.textDim }}>
                  {done ? "✓" : i + 1}
                </div>
                <div style={{ flex: 1, minWidth: 0 }}>
                  <div style={{ fontSize: 14, fontWeight: 600, color: done ? C.textMuted : C.text }}>{it.name}</div>
                  <div style={{ fontSize: 12, color: C.textDim, marginTop: 1, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{it.email}</div>
                </div>
                <div style={{ fontSize: 11, color: C.textDim, flexShrink: 0, textAlign: "right" }}>
                  <div>{fmtDate(it.date)}</div>
                  <div>{fmtTime(it.startTime)}{it.endTime ? `–${fmtTime(it.endTime)}` : ""}</div>
                </div>
              </div>
            );
          })}
        </div>
        {!allDone && item && (
          <div style={{ padding: "16px 26px 22px", borderTop: `1px solid ${C.border}`, background: C.surfaceAlt }}>
            {mode === "physical" && (
              <div style={{ padding: "10px 14px", background: C.outlookGlow, border: `1px solid ${C.outlook}33`, borderRadius: 8, fontSize: 12, color: "#4da6e8", marginBottom: 12, lineHeight: 1.6 }}>
                <strong>ℹ️ How it works:</strong> The formatted invite is copied (same as Preview) and your mail app opens with recipients and subject. Click in the body and press <strong>Ctrl+V</strong>. In <strong>new Outlook</strong>, use <strong>Paste → Keep source formatting</strong> if sizes look wrong.
              </div>
            )}
            {mode === "virtual" && (
              <div style={{ padding: "10px 14px", background: C.teamsGlow, border: `1px solid ${C.teams}33`, borderRadius: 8, fontSize: 12, color: "#8b8cc8", marginBottom: 12, lineHeight: 1.6 }}>
                <strong>ℹ️ Teams limitation:</strong> Microsoft does not allow pre-filling a Teams meeting via URL. Clicking below copies the formatted invite (same as Preview) and opens a new Teams meeting. Paste into the description with Ctrl+V, set the title and time, add the attendee, then send.
              </div>
            )}
            <div style={{ fontSize: 12, color: C.textMuted, marginBottom: 12, lineHeight: 1.6 }}>
              <strong style={{ color: C.text }}>Step {current + 1} of {total}:</strong>{" "}
              {mode === "virtual"
                ? <>Copy formatted body + open Teams for <strong style={{ color: C.text }}>{item.name}</strong>. Paste into the meeting description (Ctrl+V).</>
                : <>Copy formatted invite and open mail for <strong style={{ color: C.text }}>{item.name}</strong> — paste into the body (Ctrl+V).</>}
            </div>
            <div style={{ display: "flex", gap: 10, alignItems: "center" }}>
              <Btn variant="primary" color={modeCol} onClick={() => openDraft(current)} style={{ flex: 1, justifyContent: "center", padding: "11px 0", fontSize: 14, boxShadow: `0 4px 16px ${modeGlow}` }}>
                {mode === "virtual" ? <Icon.Teams /> : <Icon.Outlook />}
                {mode === "virtual" ? `Copy & Open Teams for ${item.name}` : `Open Mail for ${item.name}`}
              </Btn>
              {opened.has(current) && current < total - 1 && (
                <Btn variant="ghost" onClick={() => setCurrent(c => c + 1)} style={{ padding: "11px 18px", fontSize: 13 }}>Next →</Btn>
              )}
            </div>
            {opened.has(current) && (
              <div style={{ fontSize: 11, color: C.success, marginTop: 10, display: "flex", alignItems: "center", gap: 5 }}>
                <Icon.Check /> {mode === "virtual" ? "Formatted body copied — paste into Teams (Ctrl+V)." : "Body copied — paste into the mail draft (Ctrl+V), then send."}{" "}
                {current < total - 1 ? "Then click Next." : "You're all done!"}
              </div>
            )}

          </div>
        )}
        {allDone && (
          <div style={{ padding: "20px 26px 24px", borderTop: `1px solid ${C.border}`, textAlign: "center" }}>
            <div style={{ fontSize: 28, marginBottom: 8 }}>🎉</div>
            <div style={{ fontSize: 15, fontWeight: 700, marginBottom: 4 }}>All {total} invite{total !== 1 ? "s" : ""} dispatched!</div>
            <div style={{ fontSize: 12, color: C.textMuted, marginBottom: 16 }}>Make sure you've sent each one before closing.</div>
            <Btn variant="ghost" onClick={onClose} style={{ margin: "0 auto" }}>Close</Btn>
          </div>
        )}
      </div>
    </div>
  );
}

// ─── Email body builders ──────────────────────────────────────────────────────
// Returns { subjectHtml, subjectPlain, bodyHtml, bodyPlain }
// Styling matches exact docx inspection:
//   Font: Calibri
//   Labels bold, values normal
//   Pre-Interview Requirements: bold 13pt
//   Bullet text / tips: 11pt normal
//   Wishing you the best: bold italic 14pt
//   Candidate name in greeting: bold

function buildEmails(c, { mode, role, interviewStage, companyName, interviewers = [], formLink, formLinkText, additionalNotes }) {
  const panel  = formatPanel(interviewers);
  const locLabel = (c.room || "").trim() || "[Interview Location]";
  const locHtml  = (c.roomUrl || "").trim()
    ? htmlLink(c.roomUrl, locLabel === "[Interview Location]" ? c.roomUrl : locLabel)
    : escapeHtml(locLabel);
  const locPlain = (c.roomUrl || "").trim()
    ? `${locLabel === "[Interview Location]" ? c.roomUrl.trim() : locLabel} (${c.roomUrl.trim()})`
    : locLabel;
  const date   = fmtDate(c.date)      || "[DD/MM/YYYY]";
  const time   = fmtTime(c.startTime) || "[HH:MM AM/PM]";
  const timeEnd = c.showEndTimeInEmail && c.endTime ? ` – ${fmtTime(c.endTime)}` : "";
  const pos    = role           || "[Applied Position Title]";
  const stage  = interviewStage || "[Interview Stage]";
  const co     = companyName    || "[Company Name]";
  const name          = c.name || "[Candidate Name]";
  const greetingName  = firstName(c.name);

  const formAnchor = (formLinkText || "").trim() || "Microsoft Form";
  const fLinkHtml  = (formLink || "").trim()
    ? htmlLink(formLink, formAnchor)
    : escapeHtml("[Microsoft Form Link]");
  const fLinkPlain = (formLink || "").trim()
    ? `${formAnchor} (${formLink.trim()})`
    : "[Microsoft Form Link]";

  // ── Subject ──
  const subjectHtml = `Subject: ${escapeHtml(stage)} | <b>${escapeHtml(pos)}</b> – <b>${escapeHtml(name)}</b> | ${escapeHtml(co)}`;
  const subjectPlain = `${stage} | ${pos} – ${name} | ${co}`;

  // ── Helpers (sizes on inner spans so Outlook keeps 11 / 13 / 14 pt on paste) ──
  const row = (label, valueHtml) =>
    outlookP("4pt 0", O.s11, `• <b>${escapeHtml(label)}</b>: ${valueHtml}`, O.f11);
  const tip = (text) =>
    outlookP("4pt 0", O.s11, `• ${escapeHtml(text)}`, O.f11);
  const notesHtml = additionalNotes
    ? outlookP("16pt 0 0 0", O.s11, escapeHtml(additionalNotes).replace(/\n/g, "<br/>"), O.f11)
    : "";

  let bodyHtml = "";
  let bodyPlain = "";

  if (mode === "physical") {

    bodyHtml = `
${outlookP("0 0 4pt 0", O.s11, `Dear <b>${escapeHtml(greetingName)}</b>,`, O.f11)}
${outlookP("0 0 4pt 0", O.s11, "Please find below the details for your interview:", O.f11)}
${row("Date", escapeHtml(date))}
${row("Time", escapeHtml(time + timeEnd))}
${row("Position", escapeHtml(pos))}
${row("Interview Stage", escapeHtml(stage))}
${row("Interview Type", "Physical (In-Person)")}
${row("Interview Location", locHtml)}
${row("Interview Panel", panel.html)}
${outlookP("16pt 0 4pt 0", O.s13b, "Pre-Interview Requirements", O.f13)}
${outlookP("4pt 0", O.s11, `• Complete the ${fLinkHtml} at least one hour before your arrival.`, O.f11)}
${outlookP("4pt 0", O.s11, "• Bring your NIC and present it at the front desk for security clearance.", O.f11)}
${outlookP("16pt 0 4pt 0", O.s11, "Here are some tips for you to make this interview a great experience,", O.f11)}
${tip("Please arrive at least 10 minutes early to allow time for security check-in and registration.")}
${tip("Bring a printed copy of your CV, and any relevant documents (if requested).")}
${tip("Dress in formal business attire appropriate for a professional interview.")}
${tip("If you have trouble finding the location, contact the recruiter in advance for assistance.")}
${tip("If you are unable to attend or will be late, kindly inform the recruiter as soon as possible.")}
${notesHtml}
${outlookP("20pt 0 0 0", O.s14bi, "Wishing you the best for your interview!", O.f14)}`;

    bodyPlain = [
      `Dear ${greetingName},`,
      "",
      "Please find below the details for your interview:",
      "",
      `- Date: ${date}`,
      `- Time: ${time}${timeEnd}`,
      `- Position: ${pos}`,
      `- Interview Stage: ${stage}`,
      `- Interview Type: Physical (In-Person)`,
      `- Interview Location: ${locPlain}`,
      `- Interview Panel: ${panel.plain}`,
      "",
      "Pre-Interview Requirements",
      "",
      `- Complete the ${fLinkPlain} at least one hour before your arrival.`,
      "- Bring your NIC and present it at the front desk for security clearance.",
      "",
      "Here are some tips for you to make this interview a great experience,",
      "",
      "- Please arrive at least 10 minutes early to allow time for security check-in and registration.",
      "- Bring a printed copy of your CV, and any relevant documents (if requested).",
      "- Dress in formal business attire appropriate for a professional interview.",
      "- If you have trouble finding the location, contact the recruiter in advance for assistance.",
      "- If you are unable to attend or will be late, kindly inform the recruiter as soon as possible.",
      ...(additionalNotes ? ["", additionalNotes] : []),
      "",
      "",
      "Wishing you the best for your interview!",
    ].join("\n");

  } else {
    // Virtual
    bodyHtml = `
${outlookP("0 0 4pt 0", O.s11, `Dear <b>${escapeHtml(greetingName)}</b>,`, O.f11)}
${outlookP("0 0 4pt 0", O.s11, "Please find below the details for your interview:", O.f11)}
${row("Date", escapeHtml(date))}
${row("Time", escapeHtml(time + timeEnd))}
${row("Position", escapeHtml(pos))}
${row("Interview Stage", escapeHtml(stage))}
${row("Interview Type", "Virtual – MS Teams")}
${row("Interview Panel", panel.html)}
${outlookP("16pt 0 4pt 0", O.s11, "Here are some tips for you to make this interview a great experience,", O.f11)}
${tip("Please join the interview on time. If you are unable to join on time, notify the recruiter in advance.")}
${tip("Ensure your internet connection, camera, and microphone are working properly before the interview.")}
${tip("Find a quiet and well-lit location with minimal distractions.")}
${tip("Keep your video turned on during the interview unless instructed otherwise.")}
${tip("Dress professionally as you would for an in-person interview.")}
${tip("If you face any technical difficulties, inform the recruiter immediately.")}
${notesHtml}
${outlookP("20pt 0 0 0", O.s14bi, "Wishing you the best for your interview!", O.f14)}`;

    bodyPlain = [
      `Dear ${greetingName},`,
      "",
      "Please find below the details for your interview:",
      "",
      `- Date: ${date}`,
      `- Time: ${time}${timeEnd}`,
      `- Position: ${pos}`,
      `- Interview Stage: ${stage}`,
      `- Interview Type: Virtual – MS Teams`,
      `- Interview Panel: ${panel.plain}`,
      "",
      "Here are some tips for you to make this interview a great experience,",
      "",
      "- Please join the interview on time. If you are unable to join on time, notify the recruiter in advance.",
      "- Ensure your internet connection, camera, and microphone are working properly before the interview.",
      "- Find a quiet and well-lit location with minimal distractions.",
      "- Keep your video turned on during the interview unless instructed otherwise.",
      "- Dress professionally as you would for an in-person interview.",
      "- If you face any technical difficulties, inform the recruiter immediately.",
      ...(additionalNotes ? ["", additionalNotes] : []),
      "",
      "",
      "Wishing you the best for your interview!",
    ].join("\n");
  }

  return { subjectHtml, subjectPlain, bodyHtml, bodyPlain };
}

// ─── Main App ─────────────────────────────────────────────────────────────────
export default function HRScheduler() {
  const cache = readCache();

  const [mode, setMode]             = useState("physical");
  const [companyName, setCompany]   = useState(cache.companyName || "");
  const [interviewStage, setStage]  = useState(cache.interviewStage || "");
  const [role, setRole]             = useState(cache.role || "");
  const [interviewers, setPanel]    = useState([]);
  const [formLink, setFormLink]         = useState(cache.formLink || "");
  const [formLinkText, setFormLinkText] = useState(cache.formLinkText || "Microsoft Form");
  const [additionalNotes, setNotes] = useState("");
  const [templateId, setTplId]      = useState("general");
  const [cc, setCc]                 = useState([]);
  const [bcc, setBcc]               = useState([]);
  const [candidates, setCands]      = useState([defaultCandidate()]);

  const [builtins, setBuiltins] = useState(() => BUILTIN_TPLS.map(t => {
    try { const s = JSON.parse(localStorage.getItem(`hr_builtin_${t.id}`) || "null"); return s || t; } catch { return t; }
  }));
  const [customs, setCustoms]   = useState(loadCustomTpls);
  const [showTplMgr, setTplMgr]   = useState(false);
  const [showSaveTpl, setShowSaveTpl] = useState(false);
  const [sentLog, setSentLog]     = useState(null);
  const [draftSaved, setDS]       = useState(false);
  const [hasDraft, setHasDraft]   = useState(() => !!loadDraft());
  const [preview, setPreview]     = useState(null);

  const allTpls = [...builtins, ...customs];
  const formConfig = () => captureFormConfig({ mode, companyName, interviewStage, role, interviewers, formLink, formLinkText, additionalNotes, cc, bcc });
  const canSaveTemplate = role.trim() && interviewStage.trim();
  const activeCustomTpl = customs.find(t => t.id === templateId);

  // Cache shared fields on change
  useEffect(() => { if (companyName)    patchCache({ companyName }); }, [companyName]);
  useEffect(() => { if (interviewStage) patchCache({ interviewStage }); }, [interviewStage]);
  useEffect(() => { if (role)           patchCache({ role }); }, [role]);
  useEffect(() => { if (formLink)     patchCache({ formLink }); }, [formLink]);
  useEffect(() => { if (formLinkText) patchCache({ formLinkText }); }, [formLinkText]);

  // Auto-save draft
  useEffect(() => {
    const t = setTimeout(() => {
      saveDraft({ mode, companyName, interviewStage, role, interviewers, formLink, formLinkText, additionalNotes, templateId, cc, bcc, candidates });
      setHasDraft(true);
    }, 800);
    return () => clearTimeout(t);
  }, [mode, companyName, interviewStage, role, interviewers, formLink, formLinkText, additionalNotes, templateId, cc, bcc, candidates]);

  // Conflict detection
  const conflictIds = new Set();
  for (let i = 0; i < candidates.length; i++) {
    for (let j = i + 1; j < candidates.length; j++) {
      const a = candidates[i], b = candidates[j];
      if (a.date && b.date && a.date === b.date && a.startTime && b.startTime) {
        const aEnd = a.endTime || "23:59";
        const bEnd = b.endTime || "23:59";
        if (a.startTime < bEnd && b.startTime < aEnd) { conflictIds.add(a.id); conflictIds.add(b.id); }
      }
    }
  }

  const updCand = (id, key, val) => setCands(cs => cs.map(c => c.id === id ? { ...c, [key]: val } : c));
  const addCand = () => setCands(cs => [...cs, defaultCandidate()]);
  const remCand = (id) => setCands(cs => cs.filter(c => c.id !== id));

  const applyTemplate = (tpl) => {
    setTplId(tpl.id);
    if (tpl.config) {
      setMode(tpl.config.mode || "physical");
      setCompany(tpl.config.companyName || "");
      setStage(tpl.config.interviewStage || "");
      setRole(tpl.config.role || "");
      setPanel(tpl.config.interviewers || []);
      setFormLink(tpl.config.formLink || "");
      setFormLinkText(tpl.config.formLinkText || "Microsoft Form");
      setNotes(tpl.config.additionalNotes || tpl.notes || "");
      setCc(tpl.config.cc || []);
      setBcc(tpl.config.bcc || []);
    } else if (tpl.notes) {
      setNotes(tpl.notes);
    }
  };
  const selectTpl = (tpl) => { applyTemplate(tpl); setTplMgr(false); };

  const normalizeTpl = (updated) => {
    const notes = updated.notes ?? updated.config?.additionalNotes ?? "";
    const tpl = { ...updated, notes };
    if (tpl.config) tpl.config = { ...tpl.config, additionalNotes: notes };
    return tpl;
  };

  const saveTpl = (updated, isNew) => {
    const tpl = normalizeTpl(updated);
    if (isNew) {
      const next = [...customs, tpl]; setCustoms(next); saveCustomTpls(next); setTplId(tpl.id);
    } else {
      const bIdx = builtins.findIndex(t => t.id === tpl.id);
      if (bIdx !== -1) {
        const next = builtins.map(t => t.id === tpl.id ? tpl : t);
        setBuiltins(next); localStorage.setItem(`hr_builtin_${tpl.id}`, JSON.stringify(tpl));
      } else {
        const next = customs.map(t => t.id === tpl.id ? tpl : t);
        setCustoms(next); saveCustomTpls(next);
      }
    }
  };
  const delCustom = (id) => { const next = customs.filter(t => t.id !== id); setCustoms(next); saveCustomTpls(next); };

  const loadDraftFn = () => {
    const d = loadDraft(); if (!d) return;
    setMode(d.mode || "physical"); setCompany(d.companyName || ""); setStage(d.interviewStage || "");
    setRole(d.role || ""); setPanel(d.interviewers || []); setFormLink(d.formLink || "");
    setFormLinkText(d.formLinkText || "Microsoft Form");
    setNotes(d.additionalNotes || ""); setTplId(d.templateId || "general");
    setCc(d.cc || []); setBcc(d.bcc || []);
    setCands(d.candidates?.length ? d.candidates : [defaultCandidate()]);
  };

  const saveDraftNow = () => {
    saveDraft({ mode, companyName, interviewStage, role, interviewers, formLink, formLinkText, additionalNotes, templateId, cc, bcc, candidates });
    setDS(true); setTimeout(() => setDS(false), 2000);
  };

  const clearAll = () => {
    setMode("physical"); setCompany(""); setStage(""); setRole(""); setPanel([]); setFormLink(""); setFormLinkText("Microsoft Form"); setNotes("");
    setTplId("general"); setCc([]); setBcc([]); setCands([defaultCandidate()]);
    localStorage.removeItem(DRAFT_KEY); setHasDraft(false);
  };

  const params = { mode, role, interviewStage, companyName, interviewers, formLink, formLinkText, additionalNotes };

  const validate = () => {
    if (!role.trim())          return "Applied Position Title is required.";
    if (!interviewStage.trim()) return "Interview Stage is required.";
    for (const c of candidates) {
      if (!c.name.trim())  return `Candidate ${candidates.indexOf(c)+1}: name required.`;
      if (!c.email.trim() || !c.email.includes("@")) return `${c.name}: valid email required.`;
      if (!c.date)         return `${c.name}: date required.`;
      if (!c.startTime) return `${c.name || `Candidate ${candidates.indexOf(c) + 1}`}: start time required.`;
      if (c.endTime && c.startTime >= c.endTime) return `${c.name || `Candidate ${candidates.indexOf(c) + 1}`}: end time must be after start time.`;
    }
    return null;
  };

  const handlePreview = () => {
    const err = validate(); if (err) { alert(err); return; }
    const items = candidates.map((c, i) => {
      const { subjectHtml, bodyHtml, bodyPlain } = buildEmails(c, params);
      const label = c.name.trim() || `Candidate ${i + 1}`;
      return { label, email: c.email, subjectHtml, bodyHtml, bodyPlain };
    });
    setPreview({ items });
  };

  const handleSend = () => {
    const err = validate(); if (err) { alert(err); return; }
    const items = candidates.map(c => {
      const { subjectPlain, bodyHtml, bodyPlain } = buildEmails(c, params);
      return { name: c.name, email: c.email, date: c.date, startTime: c.startTime, endTime: c.endTime, cc, bcc, subjectPlain, bodyHtml, bodyPlain };
    });
    setSentLog(items);
  };

  const isVirtual = mode === "virtual";
  const modeCol   = isVirtual ? C.teams : C.outlook;
  const modeGlow  = isVirtual ? C.teamsGlow : C.outlookGlow;
  const activeTpl = allTpls.find(t => t.id === templateId) || allTpls[0];

  return (
    <div style={{ minHeight: "100vh", background: C.bg, color: C.text, fontFamily: "'DM Sans','Segoe UI',sans-serif", paddingBottom: 80 }}>

      {/* Header */}
      <div style={{ background: `linear-gradient(135deg,${C.surface} 0%,#161a26 100%)`, borderBottom: `1px solid ${C.border}`, padding: "20px 36px", display: "flex", alignItems: "center", gap: 14, position: "sticky", top: 0, zIndex: 100, backdropFilter: "blur(12px)" }}>
        <div style={{ width: 38, height: 38, borderRadius: 10, background: `linear-gradient(135deg,${C.accent},${C.accentDim})`, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 18, flexShrink: 0 }}>🗓</div>
        <div style={{ flex: 1 }}>
          <div style={{ fontSize: 18, fontWeight: 700, letterSpacing: -0.2 }}>HR Interview Scheduler</div>
          <div style={{ fontSize: 12, color: C.textMuted }}>Dispatch personalised interview invites to multiple candidates at once</div>
        </div>
        <div style={{ display: "flex", alignItems: "center", gap: 10, flexShrink: 0 }}>
          <div style={{ padding: "7px 14px", borderRadius: 8, background: modeGlow, border: `1px solid ${modeCol}55`, fontSize: 12, color: isVirtual ? "#8b8cc8" : "#4da6e8", display: "flex", alignItems: "center", gap: 7, fontWeight: 600 }}>
            {isVirtual ? <Icon.Teams /> : <Icon.Outlook />}
            {isVirtual ? "Virtual – MS Teams" : "Physical – Outlook"}
          </div>
          {hasDraft && <Btn variant="ghost" onClick={loadDraftFn} style={{ fontSize: 12 }}><Icon.Draft /> Load Draft</Btn>}
          <Btn variant="ghost" onClick={saveDraftNow} style={{ fontSize: 12 }}><Icon.Save /> {draftSaved ? "Saved!" : "Save Draft"}</Btn>
        </div>
      </div>

      <div style={{ maxWidth: 960, margin: "0 auto", padding: "30px 20px 0" }}>

        {/* ── Section 1: Interview Type ── */}
        <div style={{ background: C.surface, border: `1px solid ${C.border}`, borderRadius: 14, padding: "24px 28px", marginBottom: 16 }}>
          <div style={{ fontSize: 10, fontWeight: 700, letterSpacing: 1.5, color: C.textDim, textTransform: "uppercase", marginBottom: 16 }}>Interview Type</div>
          <div style={{ display: "flex", gap: 14, flexWrap: "wrap", alignItems: "flex-end" }}>
            <Select label="Interview Format *" value={mode} onChange={setMode}
              options={[["physical", "Physical (In-Person)"], ["virtual", "Virtual – MS Teams"]]}
              style={{ maxWidth: 280 }} />
            <div style={{ padding: "9px 18px", borderRadius: 8, background: modeGlow, border: `1px solid ${modeCol}44`, fontSize: 13, color: modeCol, display: "flex", alignItems: "center", gap: 8, fontWeight: 600, marginBottom: 1 }}>
              {isVirtual ? <Icon.Teams /> : <Icon.Outlook />}
              {isVirtual ? "Invites sent via Microsoft Teams" : "Invites sent via Outlook email"}
            </div>
          </div>
        </div>

        {/* ── Section 2: Interview Details ── */}
        <div style={{ background: C.surface, border: `1px solid ${C.border}`, borderRadius: 14, padding: "24px 28px", marginBottom: 16 }}>
          <div style={{ fontSize: 10, fontWeight: 700, letterSpacing: 1.5, color: C.textDim, textTransform: "uppercase", marginBottom: 16 }}>Interview Details</div>
          <div style={{ display: "flex", gap: 12, flexWrap: "wrap", marginBottom: 12 }}>
            <Input label="Applied Position Title *" value={role} onChange={setRole} placeholder="e.g. Senior Software Engineer" required cacheKey="role" />
            <Input label="Interview Stage *" value={interviewStage} onChange={setStage} placeholder="e.g. 1st Interview, Final Round" required cacheKey="interviewStage" />
          </div>
          <div style={{ display: "flex", gap: 12, flexWrap: "wrap", marginBottom: 16 }}>
            <Input label="Company Name" value={companyName} onChange={setCompany} placeholder="e.g. Ceylon Cold Stores PLC" cacheKey="companyName" />
            {!isVirtual && (
              <LinkField
                label="Microsoft Form (Pre-Interview)"
                url={formLink}
                onUrlChange={setFormLink}
                text={formLinkText}
                onTextChange={setFormLinkText}
                urlPlaceholder="https://forms.office.com/..."
                textPlaceholder="Microsoft Form"
                urlCacheKey="formLink"
                textCacheKey="formLinkText"
              />
            )}
          </div>

          {/* Template picker */}
          <div style={{ marginBottom: 14 }}>
            <label style={{ fontSize: 11, color: C.textMuted, fontWeight: 600, letterSpacing: 0.4, display: "block", marginBottom: 6 }}>Template Tag</label>
            <div style={{ display: "flex", gap: 8, flexWrap: "wrap", alignItems: "center" }}>
              {allTpls.map(tpl => (
                <button key={tpl.id} onClick={() => selectTpl(tpl)}
                  style={{ padding: "7px 14px", borderRadius: 8, border: `1.5px solid ${templateId === tpl.id ? modeCol : C.border}`, background: templateId === tpl.id ? modeGlow : C.surfaceAlt, color: templateId === tpl.id ? modeCol : C.textMuted, cursor: "pointer", fontSize: 13, fontWeight: 600, display: "flex", alignItems: "center", gap: 5, transition: "all .15s" }}>
                  <span>{tpl.icon}</span> {tpl.label}
                </button>
              ))}
              <Btn variant="ghost" onClick={() => setTplMgr(true)} style={{ padding: "7px 12px", fontSize: 12, borderStyle: "dashed" }}>
                <Icon.Template /> Manage Templates
              </Btn>
              {canSaveTemplate && (
                <Btn variant="ghost" color={C.success} onClick={() => setShowSaveTpl(true)} style={{ padding: "7px 12px", fontSize: 12, borderStyle: "dashed" }} title="Save current interview setup as a reusable template">
                  <Icon.Save /> Save as Template
                </Btn>
              )}
            </div>
          </div>

          <Textarea label="Additional Notes (appended after tips section)" value={additionalNotes} onChange={setNotes}
            placeholder="Any extra info — special instructions, documents to bring, etc." minHeight={75} />
          <div style={{ fontSize: 11, color: C.textDim, marginTop: 6 }}>
            The full invite body is auto-generated from the fields above matching the official template. Additional notes are appended at the end.
          </div>
        </div>

        {/* ── Section 3: Panel & Distribution ── */}
        <div style={{ background: C.surface, border: `1px solid ${C.border}`, borderRadius: 14, padding: "24px 28px", marginBottom: 16 }}>
          <div style={{ fontSize: 10, fontWeight: 700, letterSpacing: 1.5, color: C.textDim, textTransform: "uppercase", marginBottom: 16 }}>Panel & Distribution</div>
          <div style={{ display: "flex", gap: 12, flexWrap: "wrap" }}>
            <TagInput label="Interview Panel Members" values={interviewers} onChange={setPanel} placeholder="Name or email — press Enter" cacheKey="panelMembers" />
            <TagInput label="CC" values={cc} onChange={setCc} placeholder="cc@company.com" cacheKey="ccEmails" />
            <TagInput label="BCC" values={bcc} onChange={setBcc} placeholder="bcc@company.com" cacheKey="bccEmails" />
          </div>
        </div>

        {/* ── Section 4: Candidates ── */}
        <div style={{ background: C.surface, border: `1px solid ${C.border}`, borderRadius: 14, padding: "24px 28px", marginBottom: 16 }}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 16 }}>
            <div>
              <div style={{ fontSize: 10, fontWeight: 700, letterSpacing: 1.5, color: C.textDim, textTransform: "uppercase", marginBottom: 4 }}>Candidates</div>
              <div style={{ fontSize: 12, color: C.textMuted }}>Each candidate gets a <strong style={{ color: C.text }}>separate, personally addressed</strong> invite.</div>
            </div>
            <Btn variant="ghost" color={modeCol} onClick={addCand}><Icon.Plus /> Add Candidate</Btn>
          </div>
          {candidates.map((c, i) => (
            <CandidateCard key={c.id} c={c} idx={i} mode={mode}
              onChange={(k, v) => updCand(c.id, k, v)}
              onRemove={() => remCand(c.id)}
              canRemove={candidates.length > 1}
              conflictIds={conflictIds} />
          ))}
          {candidates.length > 1 && (
            <div style={{ fontSize: 12, color: C.textMuted, marginTop: 8, display: "flex", alignItems: "center", gap: 6 }}>
              <Icon.Check /> {candidates.length} candidates ·{" "}
              {conflictIds.size > 0 ? <span style={{ color: C.warn }}>⚠ {conflictIds.size / 2} time slot conflict(s) detected</span> : "No time conflicts detected"}
            </div>
          )}
        </div>

        {/* ── Actions ── */}
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginTop: 10 }}>
          <Btn variant="subtle" onClick={clearAll}><Icon.Trash /> Clear All</Btn>
          <div style={{ display: "flex", gap: 10 }}>
            <Btn variant="ghost" onClick={saveDraftNow}><Icon.Draft /> {draftSaved ? "✓ Saved" : "Save Draft"}</Btn>
            <Btn variant="ghost" color={C.purple} onClick={handlePreview}><Icon.Eye /> Preview {candidates.length > 1 ? `${candidates.length} Emails` : "Email"}</Btn>
            <Btn variant="primary" color={modeCol} onClick={handleSend} style={{ padding: "11px 26px", fontSize: 14, boxShadow: `0 4px 20px ${modeGlow}` }}>
              <Icon.Send /> Dispatch {candidates.length} Invite{candidates.length !== 1 ? "s" : ""}
            </Btn>
          </div>
        </div>
      </div>

      {/* Modals */}
      {showTplMgr && <TemplateManager builtins={builtins} customs={customs} onSave={saveTpl} onDelete={delCustom} onClose={() => setTplMgr(false)} onSelectAndClose={selectTpl} getCurrentConfig={formConfig} />}
      {showSaveTpl && (
        <SaveTemplateModal
          initial={activeCustomTpl}
          currentConfig={formConfig()}
          onSave={saveTpl}
          onClose={() => setShowSaveTpl(false)}
        />
      )}
      {preview && <PreviewModal items={preview.items} onClose={() => setPreview(null)} />}
      {sentLog    && <DispatchModal items={sentLog} mode={mode} onClose={() => setSentLog(null)} />}
    </div>
  );
}