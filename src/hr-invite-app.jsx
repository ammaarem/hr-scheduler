import { useState, useEffect, useRef } from "react";

// ─── Design tokens ────────────────────────────────────────────────────────────
const C = {
  bg: "#0f1117", surface: "#181c27", surfaceAlt: "#1e2335",
  border: "#2a3050", borderHover: "#3d4d7a",
  accent: "#4f8ef7", accentDim: "#2d5bbf", accentGlow: "rgba(79,142,247,0.18)",
  teams: "#6264a7", teamsDim: "#4a4c8a", teamsGlow: "rgba(98,100,167,0.2)",
  outlook: "#0072c6", outlookDim: "#005099", outlookGlow: "rgba(0,114,198,0.18)",
  text: "#e8ecf5", textMuted: "#8892b0", textDim: "#5a6480",
  danger: "#f07070", success: "#5dd88a", warn: "#f5a623",
  purple: "#a78bfa", green: "#34d399", orange: "#fb923c",
};

const uid = () => Math.random().toString(36).slice(2, 9);

// ─── Built-in templates (these ship with the app) ─────────────────────────────
const BUILTIN_TEMPLATES = [
  { id: "general",   label: "General Interview",  icon: "💼", body: "We look forward to meeting with you to discuss this exciting opportunity. Please feel free to reach out if you have any questions prior to your interview." },
  { id: "technical", label: "Technical Round",    icon: "⚙️", body: "This interview will include a technical assessment covering relevant skills for the role. Please be prepared to walk through your experience and work on a problem-solving exercise. You may use your preferred language or tool unless otherwise specified." },
  { id: "hr",        label: "HR Screening",       icon: "🤝", body: "This is an initial HR screening call to discuss the role, your background, and to answer any questions you may have about the position and company culture. The session will be conversational — no preparation is required beyond reviewing the job description." },
  { id: "final",     label: "Final Round",        icon: "🏆", body: "Congratulations on reaching the final round of interviews! You will be meeting with senior members of our team. Please bring any materials or portfolio items you would like to share. This is also your opportunity to ask any remaining questions about the role." },
  { id: "panel",     label: "Panel Interview",    icon: "👥", body: "This is a panel interview where you will be meeting with multiple team members simultaneously. The session will cover your experience, situational responses, and role-specific competencies. Each panelist may ask questions from their area of expertise." },
];

const STORAGE_KEY = "hr_scheduler_templates_v1";
const DRAFT_KEY   = "hr_scheduler_draft_v1";

function loadCustomTemplates() {
  try { return JSON.parse(localStorage.getItem(STORAGE_KEY) || "[]"); } catch { return []; }
}
function saveCustomTemplates(tpls) {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(tpls));
}
function loadDraft() {
  try { return JSON.parse(localStorage.getItem(DRAFT_KEY) || "null"); } catch { return null; }
}
function saveDraft(data) {
  localStorage.setItem(DRAFT_KEY, JSON.stringify(data));
}

const defaultCandidate = () => ({
  id: uid(), name: "", email: "", date: "", startTime: "", endTime: "", room: "",
});

// ─── Icons ────────────────────────────────────────────────────────────────────
const Icon = {
  Teams: () => (
    <svg width="18" height="18" viewBox="0 0 24 24" fill="none">
      <path d="M17 4a3 3 0 1 1-6 0 3 3 0 0 1 6 0z" fill="#6264a7"/>
      <path d="M20 9h-7a2 2 0 0 0-2 2v6a2 2 0 0 0 2 2h7a2 2 0 0 0 2-2v-6a2 2 0 0 0-2-2z" fill="#6264a7"/>
      <circle cx="7" cy="8" r="2.5" fill="#8b8cc8"/>
      <path d="M10 13H4a1.5 1.5 0 0 0-1.5 1.5v4A1.5 1.5 0 0 0 4 20h6a1.5 1.5 0 0 0 1.5-1.5V17" stroke="#8b8cc8" strokeWidth="1.2" fill="none"/>
    </svg>
  ),
  Outlook: () => (
    <svg width="18" height="18" viewBox="0 0 24 24" fill="none">
      <rect x="2" y="5" width="20" height="14" rx="2" fill="#0072c6" opacity="0.2" stroke="#0072c6" strokeWidth="1.5"/>
      <path d="M2 8l10 7 10-7" stroke="#0072c6" strokeWidth="1.5" strokeLinecap="round"/>
    </svg>
  ),
  Plus:      () => <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round"><line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/></svg>,
  Trash:     () => <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><polyline points="3 6 5 6 21 6"/><path d="M19 6l-1 14H6L5 6"/><path d="M10 11v6M14 11v6"/><path d="M9 6V4h6v2"/></svg>,
  Send:      () => <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round"><line x1="22" y1="2" x2="11" y2="13"/><polygon points="22 2 15 22 11 13 2 9 22 2"/></svg>,
  Copy:      () => <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><rect x="9" y="9" width="13" height="13" rx="2"/><path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"/></svg>,
  Edit:      () => <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>,
  Save:      () => <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z"/><polyline points="17 21 17 13 7 13 7 21"/><polyline points="7 3 7 8 15 8"/></svg>,
  Check:     () => <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke={C.success} strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round"><polyline points="20 6 9 17 4 12"/></svg>,
  Warn:      () => <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke={C.warn} strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z"/><line x1="12" y1="9" x2="12" y2="13"/><line x1="12" y1="17" x2="12.01" y2="17"/></svg>,
  Chev:      () => <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round"><polyline points="6 9 12 15 18 9"/></svg>,
  Draft:     () => <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/></svg>,
  X:         () => <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>,
  Template:  () => <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><rect x="3" y="3" width="7" height="7"/><rect x="14" y="3" width="7" height="7"/><rect x="14" y="14" width="7" height="7"/><rect x="3" y="14" width="7" height="7"/></svg>,
  Clock:     () => <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round"><circle cx="12" cy="12" r="10"/><polyline points="12 6 12 12 16 14"/></svg>,
  Person:    () => <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round"><path d="M20 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2"/><circle cx="12" cy="7" r="4"/></svg>,
  Location:  () => <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round"><path d="M21 10c0 7-9 13-9 13s-9-6-9-13a9 9 0 0 1 18 0z"/><circle cx="12" cy="10" r="3"/></svg>,
};

// ─── Shared input primitives ──────────────────────────────────────────────────
function Input({ label, value, onChange, placeholder, type = "text", style = {}, required }) {
  const [f, setF] = useState(false);
  const base = { background: C.surfaceAlt, border: `1.5px solid ${f ? C.accent : C.border}`, borderRadius: 8, padding: "9px 13px", color: C.text, fontSize: 13, outline: "none", width: "100%", boxSizing: "border-box", transition: "border-color .2s, box-shadow .2s", boxShadow: f ? `0 0 0 3px ${C.accentGlow}` : "none" };
  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 5, flex: 1, minWidth: 140 }}>
      {label && <label style={{ fontSize: 11, color: C.textMuted, fontWeight: 600, letterSpacing: 0.4 }}>{label}{required && <span style={{ color: C.danger }}> *</span>}</label>}
      <input type={type} value={value} placeholder={placeholder} onChange={e => onChange(e.target.value)} onFocus={() => setF(true)} onBlur={() => setF(false)} style={{ ...base, ...style }} />
    </div>
  );
}

function Textarea({ label, value, onChange, placeholder, minHeight = 90 }) {
  const [f, setF] = useState(false);
  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 5, flex: 1 }}>
      {label && <label style={{ fontSize: 11, color: C.textMuted, fontWeight: 600, letterSpacing: 0.4 }}>{label}</label>}
      <textarea value={value} placeholder={placeholder} onChange={e => onChange(e.target.value)} onFocus={() => setF(true)} onBlur={() => setF(false)}
        style={{ background: C.surfaceAlt, border: `1.5px solid ${f ? C.accent : C.border}`, borderRadius: 8, padding: "10px 13px", color: C.text, fontSize: 13, outline: "none", width: "100%", boxSizing: "border-box", resize: "vertical", minHeight, lineHeight: 1.65, fontFamily: "inherit", transition: "border-color .2s, box-shadow .2s", boxShadow: f ? `0 0 0 3px ${C.accentGlow}` : "none" }}
      />
    </div>
  );
}

function TagInput({ label, values, onChange, placeholder, color }) {
  const [raw, setRaw] = useState("");
  const add = () => { const v = raw.trim(); if (v && !values.includes(v)) onChange([...values, v]); setRaw(""); };
  const accentCol = color || C.accent;
  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 5, flex: 1 }}>
      {label && <label style={{ fontSize: 11, color: C.textMuted, fontWeight: 600, letterSpacing: 0.4 }}>{label}</label>}
      <div style={{ display: "flex", flexWrap: "wrap", gap: 5, background: C.surfaceAlt, border: `1.5px solid ${C.border}`, borderRadius: 8, padding: "7px 9px", minHeight: 40, alignItems: "center" }}>
        {values.map((v, i) => (
          <span key={i} style={{ display: "flex", alignItems: "center", gap: 4, background: C.surface, border: `1px solid ${C.border}`, borderRadius: 5, padding: "2px 8px", fontSize: 12, color: C.textMuted }}>
            {v}
            <button onClick={() => onChange(values.filter((_, j) => j !== i))} style={{ background: "none", border: "none", cursor: "pointer", color: C.textDim, padding: 0, lineHeight: 1, fontSize: 14, display: "flex" }}>×</button>
          </span>
        ))}
        <input value={raw} placeholder={values.length ? "" : placeholder} onChange={e => setRaw(e.target.value)}
          onKeyDown={e => { if (e.key === "Enter" || e.key === ",") { e.preventDefault(); add(); } }} onBlur={add}
          style={{ border: "none", background: "transparent", outline: "none", color: C.text, fontSize: 13, flex: 1, minWidth: 130, padding: "2px 3px" }} />
      </div>
      <span style={{ fontSize: 11, color: C.textDim }}>Press Enter or comma to add</span>
    </div>
  );
}

function Btn({ children, onClick, variant = "primary", color, style: sx = {}, title, disabled }) {
  const [hov, setHov] = useState(false);
  const col = color || (variant === "primary" ? C.accent : variant === "danger" ? C.danger : variant === "success" ? C.success : C.textMuted);
  const base = { padding: "9px 18px", borderRadius: 8, border: "none", cursor: disabled ? "not-allowed" : "pointer", fontSize: 13, fontWeight: 600, display: "flex", alignItems: "center", gap: 6, transition: "all .15s", letterSpacing: 0.2, opacity: disabled ? 0.5 : 1, userSelect: "none" };
  const variants = {
    primary: { background: hov ? col + "dd" : col, color: variant === "success" ? "#0f1117" : "#fff" },
    ghost:   { background: hov ? col + "18" : "transparent", color: col, border: `1px solid ${col}44` },
    danger:  { background: hov ? C.danger + "18" : "transparent", color: C.danger, border: `1px solid transparent` },
    subtle:  { background: hov ? C.surfaceAlt : "transparent", color: C.textMuted, border: `1px solid transparent` },
  };
  return <button onClick={disabled ? undefined : onClick} onMouseEnter={() => setHov(true)} onMouseLeave={() => setHov(false)} style={{ ...base, ...variants[variant], ...sx }} title={title}>{children}</button>;
}

// ─── Template Manager Modal ───────────────────────────────────────────────────
function TemplateManager({ builtins, customs, onSave, onDelete, onClose, onSelectAndClose }) {
  const [editing, setEditing]   = useState(null); // null | { id, label, icon, body, isNew }
  const [newLabel, setNewLabel] = useState("");
  const [newIcon, setNewIcon]   = useState("📋");
  const [newBody, setNewBody]   = useState("");
  const [tab, setTab]           = useState("builtin"); // builtin | custom

  const startNew = () => setEditing({ id: uid(), label: "", icon: "📋", body: "", isNew: true });
  const startEdit = (tpl) => { setEditing({ ...tpl, isNew: false }); };

  const commitEdit = () => {
    if (!editing) return;
    const updated = { id: editing.id, label: editing.label || "Untitled", icon: editing.icon || "📋", body: editing.body };
    onSave(updated, editing.isNew);
    setEditing(null);
  };

  const allBuiltin = builtins;
  const allCustom  = customs;

  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.75)", zIndex: 600, display: "flex", alignItems: "center", justifyContent: "center", padding: 20 }}>
      <div style={{ background: C.surface, border: `1px solid ${C.border}`, borderRadius: 18, width: "100%", maxWidth: 680, maxHeight: "88vh", display: "flex", flexDirection: "column", overflow: "hidden" }}>
        {/* Modal header */}
        <div style={{ padding: "22px 28px 16px", borderBottom: `1px solid ${C.border}`, display: "flex", justifyContent: "space-between", alignItems: "center" }}>
          <div>
            <div style={{ fontSize: 17, fontWeight: 700 }}>Template Library</div>
            <div style={{ fontSize: 12, color: C.textMuted, marginTop: 3 }}>Edit built-in templates or create your own reusable ones.</div>
          </div>
          <div style={{ display: "flex", gap: 8 }}>
            <Btn variant="primary" onClick={startNew}><Icon.Plus /> New Template</Btn>
            <Btn variant="subtle" onClick={onClose}><Icon.X /></Btn>
          </div>
        </div>

        {/* Edit pane */}
        {editing && (
          <div style={{ padding: "20px 28px", borderBottom: `1px solid ${C.border}`, background: C.surfaceAlt }}>
            <div style={{ fontSize: 12, color: C.accent, fontWeight: 700, letterSpacing: 1, textTransform: "uppercase", marginBottom: 12 }}>
              {editing.isNew ? "✦ New Template" : `Editing — ${editing.label}`}
            </div>
            <div style={{ display: "flex", gap: 10, marginBottom: 10 }}>
              <Input label="Icon" value={editing.icon} onChange={v => setEditing(e => ({ ...e, icon: v }))} placeholder="Emoji" style={{ maxWidth: 70 }} />
              <Input label="Template Name" value={editing.label} onChange={v => setEditing(e => ({ ...e, label: v }))} placeholder="e.g. Culture Fit Interview" />
            </div>
            <Textarea label="Default Description" value={editing.body} onChange={v => setEditing(e => ({ ...e, body: v }))} placeholder="Write the default message body for this template…" minHeight={80} />
            <div style={{ display: "flex", gap: 8, marginTop: 12, justifyContent: "flex-end" }}>
              <Btn variant="ghost" onClick={() => setEditing(null)}>Cancel</Btn>
              <Btn variant="primary" onClick={commitEdit}><Icon.Save /> Save Template</Btn>
            </div>
          </div>
        )}

        {/* Tabs */}
        <div style={{ display: "flex", borderBottom: `1px solid ${C.border}`, padding: "0 28px" }}>
          {[["builtin", `Built-in (${allBuiltin.length})`], ["custom", `My Templates (${allCustom.length})`]].map(([k, lbl]) => (
            <button key={k} onClick={() => setTab(k)} style={{ padding: "12px 0", marginRight: 24, background: "none", border: "none", cursor: "pointer", fontSize: 13, fontWeight: 600, color: tab === k ? C.accent : C.textDim, borderBottom: `2px solid ${tab === k ? C.accent : "transparent"}`, transition: "all .15s" }}>{lbl}</button>
          ))}
        </div>

        {/* List */}
        <div style={{ overflowY: "auto", flex: 1, padding: "16px 28px 24px" }}>
          {(tab === "builtin" ? allBuiltin : allCustom).map(tpl => (
            <div key={tpl.id} style={{ display: "flex", alignItems: "flex-start", gap: 12, padding: "14px 16px", borderRadius: 10, border: `1px solid ${C.border}`, marginBottom: 10, background: C.surfaceAlt }}>
              <span style={{ fontSize: 22, flexShrink: 0 }}>{tpl.icon}</span>
              <div style={{ flex: 1, minWidth: 0 }}>
                <div style={{ fontSize: 14, fontWeight: 600, marginBottom: 4 }}>{tpl.label}</div>
                <div style={{ fontSize: 12, color: C.textMuted, lineHeight: 1.55, whiteSpace: "pre-wrap", overflow: "hidden", display: "-webkit-box", WebkitLineClamp: 2, WebkitBoxOrient: "vertical" }}>{tpl.body || <em style={{ color: C.textDim }}>No description set</em>}</div>
              </div>
              <div style={{ display: "flex", gap: 6, flexShrink: 0 }}>
                <Btn variant="ghost" onClick={() => onSelectAndClose(tpl)} style={{ padding: "7px 12px", fontSize: 12 }}>Use</Btn>
                <Btn variant="ghost" onClick={() => startEdit(tpl)} style={{ padding: "7px 10px" }}><Icon.Edit /></Btn>
                {tab === "custom" && <Btn variant="danger" onClick={() => onDelete(tpl.id)} style={{ padding: "7px 10px" }}><Icon.Trash /></Btn>}
              </div>
            </div>
          ))}
          {(tab === "builtin" ? allBuiltin : allCustom).length === 0 && (
            <div style={{ textAlign: "center", color: C.textDim, padding: "40px 0", fontSize: 13 }}>
              {tab === "custom" ? "No custom templates yet. Click \"New Template\" to create one." : "No templates."}
            </div>
          )}
        </div>
      </div>
    </div>
  );
}

// ─── Candidate card ───────────────────────────────────────────────────────────
function CandidateCard({ c, idx, mode, onChange, onRemove, canRemove, conflictIds }) {
  const timeWarn = c.startTime && c.endTime && c.startTime >= c.endTime;
  const hasConflict = conflictIds.has(c.id);
  const borderCol = timeWarn || hasConflict ? C.warn + "88" : C.border;
  const accentCol = mode === "teams" ? C.teams : C.outlook;

  return (
    <div style={{ background: C.surfaceAlt, border: `1.5px solid ${borderCol}`, borderRadius: 12, padding: "18px 20px", marginBottom: 10, transition: "border-color .2s" }}>
      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 14 }}>
        <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
          <span style={{ fontSize: 11, fontWeight: 700, letterSpacing: 1, color: accentCol, textTransform: "uppercase" }}>Candidate {idx + 1}</span>
          {timeWarn && <span style={{ display: "flex", alignItems: "center", gap: 4, fontSize: 11, color: C.warn }}><Icon.Warn /> End time before start time</span>}
          {hasConflict && !timeWarn && <span style={{ display: "flex", alignItems: "center", gap: 4, fontSize: 11, color: C.warn }}><Icon.Warn /> Time slot overlaps with another candidate</span>}
        </div>
        <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
          <span style={{ fontSize: 11, fontWeight: 700, padding: "3px 10px", borderRadius: 20, letterSpacing: 0.6, textTransform: "uppercase", background: mode === "teams" ? C.teamsGlow : C.outlookGlow, color: mode === "teams" ? "#8b8cc8" : "#4da6e8", border: `1px solid ${mode === "teams" ? C.teams + "44" : C.outlook + "44"}` }}>
            {mode === "teams" ? "Teams" : "Outlook"}
          </span>
          {canRemove && <Btn variant="danger" onClick={onRemove} style={{ padding: "5px 8px" }}><Icon.Trash /></Btn>}
        </div>
      </div>
      <div style={{ display: "flex", gap: 12, flexWrap: "wrap", marginBottom: 10 }}>
        <Input label="Full Name" value={c.name} onChange={v => onChange("name", v)} placeholder="e.g. Samantha Lee" required />
        <Input label="Email Address" value={c.email} onChange={v => onChange("email", v)} placeholder="candidate@email.com" type="email" required />
      </div>
      <div style={{ display: "flex", gap: 12, flexWrap: "wrap" }}>
        <Input label="Date" value={c.date} onChange={v => onChange("date", v)} type="date" style={{ maxWidth: 190 }} required />
        <Input label="Start Time" value={c.startTime} onChange={v => onChange("startTime", v)} type="time" style={{ maxWidth: 150 }} required />
        <Input label="End Time" value={c.endTime} onChange={v => onChange("endTime", v)} type="time" style={{ maxWidth: 150 }} required />
        {mode === "physical" && <Input label="Room / Location" value={c.room} onChange={v => onChange("room", v)} placeholder="e.g. Room B2, Floor 3" />}
      </div>
    </div>
  );
}

// ─── Dispatch summary modal ───────────────────────────────────────────────────
function SentLog({ items, onClose }) {
  const [copied, setCopied] = useState(null);
  const copy = (text, id) => { navigator.clipboard.writeText(text); setCopied(id); setTimeout(() => setCopied(null), 1600); };
  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.78)", zIndex: 700, display: "flex", alignItems: "center", justifyContent: "center", padding: 20 }}>
      <div style={{ background: C.surface, border: `1px solid ${C.border}`, borderRadius: 18, width: "100%", maxWidth: 660, maxHeight: "84vh", display: "flex", flexDirection: "column", overflow: "hidden" }}>
        <div style={{ padding: "22px 28px 16px", borderBottom: `1px solid ${C.border}`, display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
          <div>
            <div style={{ fontSize: 17, fontWeight: 700 }}>✅ Dispatch Summary</div>
            <div style={{ fontSize: 12, color: C.textMuted, marginTop: 3 }}>{items.length} personalised invite{items.length !== 1 ? "s" : ""} prepared — each addressed individually.</div>
          </div>
          <Btn variant="subtle" onClick={onClose}><Icon.X /></Btn>
        </div>
        <div style={{ overflowY: "auto", flex: 1, padding: "16px 28px" }}>
          {items.map((item, i) => (
            <div key={i} style={{ display: "flex", alignItems: "flex-start", gap: 12, padding: "14px 16px", borderRadius: 10, border: `1px solid ${C.border}`, marginBottom: 10, background: C.surfaceAlt }}>
              <div style={{ width: 30, height: 30, borderRadius: "50%", flexShrink: 0, background: "rgba(93,216,138,.12)", border: `1px solid ${C.success}55`, display: "flex", alignItems: "center", justifyContent: "center" }}><Icon.Check /></div>
              <div style={{ flex: 1, minWidth: 0 }}>
                <div style={{ fontSize: 14, fontWeight: 600 }}>{item.name}</div>
                <div style={{ fontSize: 12, color: C.textMuted, marginTop: 2 }}>{item.email}</div>
                <div style={{ fontSize: 12, color: C.textDim, marginTop: 3, display: "flex", gap: 12, flexWrap: "wrap" }}>
                  <span style={{ display: "flex", alignItems: "center", gap: 4 }}><Icon.Clock /> {item.date} · {item.startTime}–{item.endTime}</span>
                  {item.room && <span style={{ display: "flex", alignItems: "center", gap: 4 }}><Icon.Location /> {item.room}</span>}
                </div>
                <div style={{ fontSize: 11, color: item.mode === "teams" ? "#8b8cc8" : "#4da6e8", marginTop: 5, fontWeight: 600 }}>
                  {item.mode === "teams" ? "📅 Teams Invite" : "✉️ Outlook Email"}
                </div>
              </div>
              <Btn variant="ghost" onClick={() => copy(item.payload, i)} style={{ padding: "7px 12px", fontSize: 12 }}>
                <Icon.Copy /> {copied === i ? "Copied!" : "Payload"}
              </Btn>
            </div>
          ))}
        </div>
        <div style={{ padding: "14px 28px 20px", borderTop: `1px solid ${C.border}` }}>
          <div style={{ padding: "12px 14px", background: C.surfaceAlt, borderRadius: 8, fontSize: 12, color: C.textMuted, lineHeight: 1.65 }}>
            💡 <strong style={{ color: C.text }}>To enable live sending:</strong> Connect this app to <strong style={{ color: C.text }}>Microsoft Graph API</strong> — use <code style={{ color: C.accent }}>/v1.0/me/sendMail</code> for Outlook and <code style={{ color: C.accent }}>/v1.0/me/events</code> for Teams invites. The payloads above are ready to POST.
          </div>
        </div>
      </div>
    </div>
  );
}

// ─── Main App ─────────────────────────────────────────────────────────────────
export default function HRScheduler() {
  const [mode, setMode]               = useState("physical");
  const [role, setRole]               = useState("");
  const [department, setDepartment]   = useState("");
  const [interviewType, setIType]     = useState("onsite");
  const [description, setDescription] = useState("");
  const [templateId, setTemplateId]   = useState("general");
  const [interviewers, setInterviewers] = useState([]);
  const [cc, setCc]                   = useState([]);
  const [bcc, setBcc]                 = useState([]);
  const [candidates, setCandidates]   = useState([defaultCandidate()]);
  const [builtinTemplates, setBuiltins] = useState(
    () => BUILTIN_TEMPLATES.map(t => {
      // Check if user has edited this builtin and persisted changes
      try {
        const saved = JSON.parse(localStorage.getItem(`hr_builtin_${t.id}`) || "null");
        return saved || t;
      } catch { return t; }
    })
  );
  const [customTemplates, setCustoms] = useState(loadCustomTemplates);
  const [showTplMgr, setShowTplMgr]   = useState(false);
  const [sentLog, setSentLog]         = useState(null);
  const [draftSaved, setDraftSaved]   = useState(false);
  const [hasDraft, setHasDraft]       = useState(() => !!loadDraft());

  const allTemplates = [...builtinTemplates, ...customTemplates];

  // Auto-save draft periodically
  useEffect(() => {
    const timer = setTimeout(() => {
      saveDraft({ mode, role, department, description, templateId, interviewers, cc, bcc, candidates });
      setHasDraft(true);
    }, 800);
    return () => clearTimeout(timer);
  }, [mode, role, department, description, templateId, interviewers, cc, bcc, candidates]);

  // Conflict detection — same date, overlapping times across candidates
  const conflictIds = new Set();
  for (let i = 0; i < candidates.length; i++) {
    for (let j = i + 1; j < candidates.length; j++) {
      const a = candidates[i], b = candidates[j];
      if (a.date && b.date && a.date === b.date && a.startTime && a.endTime && b.startTime && b.endTime) {
        if (a.startTime < b.endTime && b.startTime < a.endTime) {
          conflictIds.add(a.id); conflictIds.add(b.id);
        }
      }
    }
  }

  const updateCandidate = (id, key, val) => setCandidates(cs => cs.map(c => c.id === id ? { ...c, [key]: val } : c));
  const addCandidate    = () => setCandidates(cs => [...cs, defaultCandidate()]);
  const removeCandidate = (id) => setCandidates(cs => cs.filter(c => c.id !== id));

  // Template selection
  const selectTemplate = (tpl) => {
    setTemplateId(tpl.id);
    setDescription(tpl.body || "");
    setShowTplMgr(false);
  };

  // Save edits to a template (builtin or custom)
  const saveTemplate = (updated, isNew) => {
    if (isNew) {
      const next = [...customTemplates, updated];
      setCustoms(next); saveCustomTemplates(next);
    } else {
      // Is it a builtin?
      const bIdx = builtinTemplates.findIndex(t => t.id === updated.id);
      if (bIdx !== -1) {
        const next = builtinTemplates.map(t => t.id === updated.id ? updated : t);
        setBuiltins(next);
        localStorage.setItem(`hr_builtin_${updated.id}`, JSON.stringify(updated));
      } else {
        const next = customTemplates.map(t => t.id === updated.id ? updated : t);
        setCustoms(next); saveCustomTemplates(next);
      }
    }
  };

  const deleteCustom = (id) => {
    const next = customTemplates.filter(t => t.id !== id);
    setCustoms(next); saveCustomTemplates(next);
  };

  // Load draft
  const loadDraftFn = () => {
    const d = loadDraft();
    if (!d) return;
    setMode(d.mode || "physical"); setRole(d.role || ""); setDepartment(d.department || "");
    setDescription(d.description || ""); setTemplateId(d.templateId || "general");
    setInterviewers(d.interviewers || []); setCc(d.cc || []); setBcc(d.bcc || []);
    setCandidates(d.candidates?.length ? d.candidates : [defaultCandidate()]);
  };

  const saveDraftNow = () => {
    saveDraft({ mode, role, department, description, templateId, interviewers, cc, bcc, candidates });
    setDraftSaved(true); setTimeout(() => setDraftSaved(false), 2000);
  };

  const clearAll = () => {
    setMode("physical"); setRole(""); setDepartment(""); setDescription(""); setTemplateId("general");
    setInterviewers([]); setCc([]); setBcc([]); setCandidates([defaultCandidate()]);
    localStorage.removeItem(DRAFT_KEY); setHasDraft(false);
  };

  const validate = () => {
    for (const c of candidates) {
      if (!c.name.trim())   return `Candidate ${c.name || "#" + (candidates.indexOf(c)+1)}: name is required.`;
      if (!c.email.trim() || !c.email.includes("@")) return `Candidate ${c.name}: valid email required.`;
      if (!c.date)          return `Candidate ${c.name}: interview date required.`;
      if (!c.startTime || !c.endTime) return `Candidate ${c.name}: start and end times required.`;
      if (c.startTime >= c.endTime)   return `Candidate ${c.name}: end time must be after start time.`;
    }
    return null;
  };

  const handleSend = () => {
    const err = validate();
    if (err) { alert(err); return; }
    const items = candidates.map(c => {
      const lines = [
        `Dear ${c.name},`,
        "",
        `Thank you for your interest in the${role ? ` ${role}` : ""} position${department ? ` within ${department}` : ""}.`,
        `We would like to invite you for ${mode === "teams" ? "a virtual" : "an in-person"} interview.`,
        "",
        `📅  Date: ${c.date}`,
        `⏰  Time: ${c.startTime} – ${c.endTime}`,
        mode === "physical" && c.room ? `📍  Location: ${c.room}` : mode === "teams" ? "📍  Format: Microsoft Teams (invite link attached)" : "",
        interviewers.length ? `👤  Interviewer(s): ${interviewers.join(", ")}` : "",
        "",
        description,
        "",
        "Please confirm your attendance by replying to this email. If you need to reschedule, do not hesitate to get in touch.",
        "",
        "We look forward to meeting you.",
        "",
        "Kind regards,",
        "HR Team",
      ].filter(l => l !== false && l !== undefined).join("\n");

      return {
        name: c.name, email: c.email, date: c.date,
        startTime: c.startTime, endTime: c.endTime, room: c.room, mode,
        payload: JSON.stringify({
          to: c.email, cc, bcc,
          subject: `Interview Invitation${role ? ` — ${role}` : ""}${department ? ` | ${department}` : ""} | ${c.date} at ${c.startTime}`,
          body: lines,
          type: mode === "teams" ? "teamsInvite" : "outlookEmail",
          ...(mode === "teams" ? { startDateTime: `${c.date}T${c.startTime}:00`, endDateTime: `${c.date}T${c.endTime}:00` } : {}),
        }, null, 2),
      };
    });
    setSentLog(items);
  };

  const isTeams    = mode === "teams";
  const modeColor  = isTeams ? C.teams : C.outlook;
  const modeGlow   = isTeams ? C.teamsGlow : C.outlookGlow;
  const activeTpl  = allTemplates.find(t => t.id === templateId) || allTemplates[0];

  return (
    <div style={{ minHeight: "100vh", background: C.bg, color: C.text, fontFamily: "'DM Sans','Segoe UI',sans-serif", paddingBottom: 80 }}>
      {/* ── Header ── */}
      <div style={{ background: `linear-gradient(135deg, ${C.surface} 0%, #161a26 100%)`, borderBottom: `1px solid ${C.border}`, padding: "22px 36px", display: "flex", alignItems: "center", gap: 14, position: "sticky", top: 0, zIndex: 100, backdropFilter: "blur(12px)" }}>
        <div style={{ width: 38, height: 38, borderRadius: 10, background: `linear-gradient(135deg,${C.accent},${C.accentDim})`, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 18, flexShrink: 0 }}>🗓</div>
        <div style={{ flex: 1 }}>
          <div style={{ fontSize: 18, fontWeight: 700, letterSpacing: -0.2 }}>HR Interview Scheduler</div>
          <div style={{ fontSize: 12, color: C.textMuted }}>Dispatch personalised interview invites to multiple candidates at once</div>
        </div>
        {/* Mode toggle */}
        <div style={{ display: "flex", borderRadius: 10, overflow: "hidden", border: `1.5px solid ${C.border}`, flexShrink: 0 }}>
          {[["physical","Physical","Outlook"],["teams","Virtual","Teams"]].map(([val, lbl, sub]) => (
            <button key={val} onClick={() => setMode(val)}
              style={{ padding: "9px 20px", cursor: "pointer", fontSize: 13, fontWeight: 600, border: "none", display: "flex", alignItems: "center", gap: 7, transition: "all .2s",
                background: mode === val ? (val === "teams" ? C.teams : C.outlook) : C.surfaceAlt,
                color: mode === val ? "#fff" : C.textMuted }}>
              {val === "teams" ? <Icon.Teams /> : <Icon.Outlook />} {lbl}
            </button>
          ))}
        </div>
        {/* Draft actions */}
        <div style={{ display: "flex", gap: 8, flexShrink: 0 }}>
          {hasDraft && <Btn variant="ghost" onClick={loadDraftFn} style={{ fontSize: 12 }}><Icon.Draft /> Load Draft</Btn>}
          <Btn variant="ghost" onClick={saveDraftNow} style={{ fontSize: 12 }}><Icon.Save /> {draftSaved ? "Saved!" : "Save Draft"}</Btn>
        </div>
      </div>

      <div style={{ maxWidth: 940, margin: "0 auto", padding: "30px 20px 0" }}>

        {/* Mode banner */}
        <div style={{ marginBottom: 18, padding: "10px 16px", borderRadius: 10, background: modeGlow, border: `1px solid ${modeColor}44`, fontSize: 13, color: modeColor, display: "flex", alignItems: "center", gap: 8 }}>
          {isTeams ? <Icon.Teams /> : <Icon.Outlook />}
          <strong>{isTeams ? "Virtual Interview — Microsoft Teams invite" : "Physical Interview — Outlook email invite"}</strong>
          <span style={{ color: C.textMuted, marginLeft: 4 }}>{isTeams ? "Each candidate receives a personalised Teams calendar invite." : "Each candidate receives a separate personalised Outlook email."}</span>
        </div>

        {/* ── Interview Details ── */}
        <div style={{ background: C.surface, border: `1px solid ${C.border}`, borderRadius: 14, padding: "24px 28px", marginBottom: 16 }}>
          <div style={{ fontSize: 10, fontWeight: 700, letterSpacing: 1.5, color: C.textDim, textTransform: "uppercase", marginBottom: 16 }}>Interview Details</div>
          <div style={{ display: "flex", gap: 12, flexWrap: "wrap", marginBottom: 14 }}>
            <Input label="Job Role / Position" value={role} onChange={setRole} placeholder="e.g. Senior Software Engineer" />
            <Input label="Department" value={department} onChange={setDepartment} placeholder="e.g. Engineering" style={{ maxWidth: 220 }} />
          </div>

          {/* Template picker */}
          <div style={{ marginBottom: 14 }}>
            <label style={{ fontSize: 11, color: C.textMuted, fontWeight: 600, letterSpacing: 0.4, display: "block", marginBottom: 6 }}>Interview Template</label>
            <div style={{ display: "flex", gap: 8, flexWrap: "wrap", alignItems: "center" }}>
              {allTemplates.map(tpl => (
                <button key={tpl.id} onClick={() => selectTemplate(tpl)}
                  style={{ padding: "7px 14px", borderRadius: 8, border: `1.5px solid ${templateId === tpl.id ? modeColor : C.border}`, background: templateId === tpl.id ? modeGlow : C.surfaceAlt, color: templateId === tpl.id ? modeColor : C.textMuted, cursor: "pointer", fontSize: 13, fontWeight: 600, display: "flex", alignItems: "center", gap: 5, transition: "all .15s" }}>
                  <span>{tpl.icon}</span> {tpl.label}
                </button>
              ))}
              <Btn variant="ghost" onClick={() => setShowTplMgr(true)} style={{ padding: "7px 12px", fontSize: 12, borderStyle: "dashed" }}>
                <Icon.Template /> Manage Templates
              </Btn>
            </div>
          </div>

          {/* Description — editable, synced to template but free-form */}
          <Textarea
            label={`Description / Message Body ${activeTpl ? `(${activeTpl.label} template — edit freely)` : ""}`}
            value={description}
            onChange={setDescription}
            placeholder="This text will appear in each candidate's invite. You can personalise it here — candidate names are added automatically."
            minHeight={110}
          />
          <div style={{ fontSize: 11, color: C.textDim, marginTop: 6 }}>
            Tip: Editing the description here only affects this session. Use "Manage Templates" to permanently update a template's default.
          </div>
        </div>

        {/* ── Panel & Distribution ── */}
        <div style={{ background: C.surface, border: `1px solid ${C.border}`, borderRadius: 14, padding: "24px 28px", marginBottom: 16 }}>
          <div style={{ fontSize: 10, fontWeight: 700, letterSpacing: 1.5, color: C.textDim, textTransform: "uppercase", marginBottom: 16 }}>Panel & Distribution</div>
          <div style={{ display: "flex", gap: 12, flexWrap: "wrap" }}>
            <TagInput label="Interviewer(s) — included in invite body" values={interviewers} onChange={setInterviewers} placeholder="interviewer@company.com" />
            <TagInput label="CC" values={cc} onChange={setCc} placeholder="cc@company.com" />
            <TagInput label="BCC" values={bcc} onChange={setBcc} placeholder="bcc@company.com" />
          </div>
        </div>

        {/* ── Candidates ── */}
        <div style={{ background: C.surface, border: `1px solid ${C.border}`, borderRadius: 14, padding: "24px 28px", marginBottom: 16 }}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 16 }}>
            <div>
              <div style={{ fontSize: 10, fontWeight: 700, letterSpacing: 1.5, color: C.textDim, textTransform: "uppercase", marginBottom: 4 }}>Candidates</div>
              <div style={{ fontSize: 12, color: C.textMuted }}>Each candidate gets a <strong style={{ color: C.text }}>separate, personally addressed</strong> invite. Add one card per candidate.</div>
            </div>
            <Btn variant="ghost" color={modeColor} onClick={addCandidate}><Icon.Plus /> Add Candidate</Btn>
          </div>

          {candidates.map((c, i) => (
            <CandidateCard key={c.id} c={c} idx={i} mode={mode}
              onChange={(k, v) => updateCandidate(c.id, k, v)}
              onRemove={() => removeCandidate(c.id)}
              canRemove={candidates.length > 1}
              conflictIds={conflictIds}
            />
          ))}

          {candidates.length > 1 && (
            <div style={{ fontSize: 12, color: C.textMuted, marginTop: 8, display: "flex", alignItems: "center", gap: 6 }}>
              <Icon.Check /> {candidates.length} candidates · {conflictIds.size > 0 ? <span style={{ color: C.warn }}>⚠ {conflictIds.size / 2} time slot conflict(s) detected</span> : "No time conflicts detected"}
            </div>
          )}
        </div>

        {/* ── Actions ── */}
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginTop: 10 }}>
          <Btn variant="subtle" onClick={clearAll}><Icon.Trash /> Clear All</Btn>
          <div style={{ display: "flex", gap: 10 }}>
            <Btn variant="ghost" onClick={saveDraftNow}><Icon.Draft /> {draftSaved ? "✓ Saved" : "Save Draft"}</Btn>
            <Btn variant="primary" color={modeColor} onClick={handleSend} style={{ padding: "11px 26px", fontSize: 14, boxShadow: `0 4px 20px ${modeGlow}` }}>
              <Icon.Send /> Dispatch {candidates.length} Invite{candidates.length !== 1 ? "s" : ""}
            </Btn>
          </div>
        </div>
      </div>

      {/* Modals */}
      {showTplMgr && (
        <TemplateManager
          builtins={builtinTemplates}
          customs={customTemplates}
          onSave={saveTemplate}
          onDelete={deleteCustom}
          onClose={() => setShowTplMgr(false)}
          onSelectAndClose={selectTemplate}
        />
      )}
      {sentLog && <SentLog items={sentLog} onClose={() => setSentLog(null)} />}
    </div>
  );
}
