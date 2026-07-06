import React, { useState, useEffect, useMemo, useRef } from "react";
import { createClient } from "@supabase/supabase-js";
import * as XLSX from "xlsx";
import {
  LineChart, Line, BarChart, Bar, AreaChart, Area, ComposedChart,
  XAxis, YAxis, Tooltip, ResponsiveContainer, CartesianGrid, Legend,
  RadarChart, Radar, PolarGrid, PolarAngleAxis,
  ScatterChart, Scatter, ZAxis, ReferenceLine,
  PieChart, Pie, Cell, Treemap,
} from "recharts";

// ============ Supabase cloud sync ============
const SUPABASE_URL = import.meta.env.VITE_SUPABASE_URL;
const SUPABASE_ANON_KEY = import.meta.env.VITE_SUPABASE_ANON_KEY;
const supabase = SUPABASE_URL && SUPABASE_ANON_KEY
  ? createClient(SUPABASE_URL, SUPABASE_ANON_KEY)
  : null;
const OWNER_EMAIL = (import.meta.env.VITE_OWNER_EMAIL || "").toLowerCase();
const callAdminFn = async (accessToken, action, payload = {}) => {
  const res = await fetch(`${SUPABASE_URL}/functions/v1/admin`, {
    method: "POST",
    headers: { Authorization: `Bearer ${accessToken}`, "Content-Type": "application/json" },
    body: JSON.stringify({ action, ...payload }),
  });
  const body = await res.json();
  if (!res.ok) throw new Error(body.error || "Admin request failed");
  return body;
};

// ============ constants ============
const C = {
  ivory: "#FAF5EB", card: "#FFFFFF", ink: "#272140",
  saffron: "#E07B1F", saffronSoft: "#F6DCC0",
  maroon: "#5C1A24", tulsi: "#4F6B3C", line: "#E8E0D0", faint: "#8A8298",
  gold: "#C9952C", sky: "#5B7FA6",
};
const PAL = [C.saffron, C.maroon, C.tulsi, C.gold, C.sky, "#9A6FB0", "#B5651D", "#3E7C7B"];
const TARGET = 16;
const BOOKS = ["Srimad-Bhagavatam", "Bhagavad-gita", "Caitanya-caritamrta", "Nectar of Devotion", "Krsna Book", "Sri Isopanisad", "Other"];
const WORSHIP_ITEMS = [
  ["mangala", "Mangala-arati"], ["narasimha", "Narasimha-arati"], ["tulsi", "Tulsi-arati"],
  ["guruPuja", "Guru-puja"], ["simple", "Simple offering (candle & flower)"],
  ["morningArati", "Morning arati (before work)"], ["eveningArati", "Evening arati (after work)"],
  ["cooked", "Cooked for the deities"],
];
const WD = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];
// approx London sunrise hour by month (mix GMT/BST)
const SUNRISE = [8.0, 7.5, 6.6, 6.4, 5.5, 4.8, 4.9, 5.7, 6.5, 7.3, 7.2, 7.9];

// ============ helpers ============
const todayKey = (d = new Date()) =>
  `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
const emptyDay = () => ({
  rounds: [], activeStart: null, hearing: [], reading: [], worship: {}, deityDressing: [],
  versesRevised: false, versesRecited: false, prayers: [], wakeTime: "", sleepTime: "", awakePeriods: [], note: "", journaled: false, journal: "",
});
const tsForTime = (timeStr) => {
  if (!timeStr) return new Date().toISOString();
  const [h, m] = timeStr.split(":").map(Number);
  const d = new Date();
  if (isNaN(h)) return d.toISOString();
  d.setHours(h, m || 0, 0, 0);
  return d.toISOString();
};
const mins = (r) => r.start && r.end ? Math.round((new Date(r.end) - new Date(r.start)) / 60000 * 10) / 10 : null;
const fmtT = (iso) => iso ? new Date(iso).toLocaleTimeString("en-GB", { hour: "2-digit", minute: "2-digit" }) : "—";
const fmtDate = (k) => {
  const [y, m, d] = k.split("-").map(Number);
  return new Date(y, m - 1, d).toLocaleDateString("en-GB", { weekday: "short", day: "numeric", month: "short" });
};
const hearingMin = (d) => (d?.hearing || []).reduce((a, h) => a + (+h.minutes || 0), 0);
const readingMin = (d) => (d?.reading || []).reduce((a, r) => a + (+r.minutes || 0), 0);
const readingPages = (d) => (d?.reading || []).reduce((a, r) => a + (+r.pages || 0), 0);
const prayerRows = (d) => Array.isArray(d?.prayers) ? d.prayers : [];
const prayersDone = (d) => Array.isArray(d?.prayers) ? d.prayers.length > 0 : !!d?.prayers;
const dressingRows = (d) => Array.isArray(d?.deityDressing) ? d.deityDressing : [];
const awakeRows = (d) => Array.isArray(d?.awakePeriods) ? d.awakePeriods : [];
const awakeCount = (d) => awakeRows(d).filter((p) => p?.time || p?.minutes).length;
const awakeMin = (d) => awakeRows(d).reduce((a, p) => a + (+p?.minutes || 0), 0);
const worshipPct = (d) => {
  const w = d?.worship || {};
  const core = ["morningArati", "eveningArati", "cooked"];
  const prog = ["mangala", "narasimha", "tulsi", "guruPuja"];
  const full = prog.every((k) => w[k]);
  const morning = full ? 1 : w.simple ? 0.6 : prog.filter((k) => w[k]).length / 4 * 0.9;
  return Math.min(1, morning * 0.5 + core.filter((k) => w[k]).length / core.length * 0.5);
};
const dayScore = (d) => {
  if (!d) return 0;
  let s = 0;
  s += Math.min((d.rounds?.length || 0) / TARGET, 1) * 40;
  s += Math.min(hearingMin(d) / 30, 1) * 15;
  s += Math.min(readingMin(d) / 30, 1) * 20;
  s += worshipPct(d) * 15;
  s += (d.versesRevised || d.versesRecited) ? 5 : 0;
  s += prayersDone(d) ? 5 : 0;
  return Math.round(s);
};
const mean = (a) => a.length ? a.reduce((x, y) => x + y, 0) / a.length : 0;
const std = (a) => { if (a.length < 2) return 0; const m = mean(a); return Math.sqrt(mean(a.map((x) => (x - m) ** 2))); };
const median = (a) => { if (!a.length) return 0; const s = [...a].sort((x, y) => x - y); const m = Math.floor(s.length / 2); return s.length % 2 ? s[m] : (s[m - 1] + s[m]) / 2; };
const quantile = (a, q) => { if (!a.length) return 0; const s = [...a].sort((x, y) => x - y); const p = (s.length - 1) * q; const b = Math.floor(p); return s[b] + (s[b + 1] !== undefined ? (s[b + 1] - s[b]) * (p - b) : 0); };
const corr = (xs, ys) => {
  if (xs.length < 3) return null;
  const mx = mean(xs), my = mean(ys);
  let n = 0, dx = 0, dy = 0;
  xs.forEach((x, i) => { n += (x - mx) * (ys[i] - my); dx += (x - mx) ** 2; dy += (ys[i] - my) ** 2; });
  return dx && dy ? n / Math.sqrt(dx * dy) : null;
};
const hmToH = (s) => { if (!s) return null; const [a, b] = s.split(":").map(Number); return isNaN(a) ? null : a + (b || 0) / 60; };
const fmtH = (h) => h == null ? "—" : `${Math.floor(h)}:${String(Math.round((h % 1) * 60)).padStart(2, "0")}`;
const scoreColor = (s) => s === 0 ? "#EFE9DC" : s < 40 ? C.saffronSoft : s < 70 ? "#EBA85C" : s < 90 ? C.saffron : C.maroon;
const grade = (p) => p >= 0.9 ? "A" : p >= 0.75 ? "B" : p >= 0.6 ? "C" : p >= 0.4 ? "D" : "F";

// ============ shared UI ============
const cardS = { background: C.card, borderRadius: 14, padding: 18, border: `1px solid ${C.line}` };
const h2S = { fontFamily: "'Fraunces', Georgia, serif", fontSize: 16, fontWeight: 600, margin: "0 0 12px", color: C.maroon };
const btnS = { padding: "8px 14px", borderRadius: 8, fontSize: 13, fontWeight: 600, border: `1.5px solid ${C.line}`, background: "#fff", color: C.ink, cursor: "pointer" };
const btnPri = { ...btnS, background: C.saffron, color: "#fff", border: `1.5px solid ${C.saffron}` };
const inpS = { border: `1px solid ${C.line}`, borderRadius: 8, padding: "8px 10px", fontSize: 14, fontFamily: "inherit", boxSizing: "border-box" };
const tickS = { fontSize: 10, fill: C.faint };

function Toggle({ label, value, onChange }) {
  return (
    <button onClick={() => onChange(!value)} style={{
      display: "flex", alignItems: "center", justifyContent: "space-between", width: "100%",
      padding: "10px 12px", borderRadius: 10, cursor: "pointer", fontSize: 14, fontWeight: 500,
      border: `1.5px solid ${value ? C.tulsi : C.line}`, background: value ? "#F0F4EC" : C.card, color: C.ink,
    }}>
      <span>{label}</span>
      <span style={{ width: 20, height: 20, borderRadius: "50%", background: value ? C.tulsi : C.line, color: "#fff", fontSize: 13, lineHeight: "20px", textAlign: "center" }}>{value ? "✓" : ""}</span>
    </button>
  );
}

function LabelField({ value, onChange, options, onAddLabel, placeholder, style }) {
  const [adding, setAdding] = useState(false);
  const [draft, setDraft] = useState("");
  if (adding) {
    return (
      <div style={{ display: "flex", gap: 6, ...style }}>
        <input autoFocus placeholder={`New ${placeholder.toLowerCase()}`} value={draft}
          onChange={(e) => setDraft(e.target.value)} style={{ ...inpS, flex: 1 }}
          onKeyDown={(e) => { if (e.key === "Enter") { e.preventDefault(); if (draft.trim()) { onAddLabel(draft.trim()); onChange(draft.trim()); } setDraft(""); setAdding(false); } }} />
        <button style={btnS} onClick={() => { if (draft.trim()) { onAddLabel(draft.trim()); onChange(draft.trim()); } setDraft(""); setAdding(false); }}>Add</button>
        <button style={btnS} onClick={() => { setAdding(false); setDraft(""); }}>×</button>
      </div>
    );
  }
  return (
    <select value={value} onChange={(e) => { if (e.target.value === "__add__") setAdding(true); else onChange(e.target.value); }} style={{ ...inpS, ...style }}>
      <option value="" disabled>{placeholder}</option>
      {options.map((o) => <option key={o} value={o}>{o}</option>)}
      <option value="__add__">+ Add {placeholder.toLowerCase()}…</option>
    </select>
  );
}

function DaySummary({ e }) {
  if (!e) return <Empty m="No entry logged for this day." />;
  const row = { display: "flex", justifyContent: "space-between", fontSize: 13, padding: "5px 0", borderBottom: `1px solid ${C.line}` };
  return (
    <div style={{ display: "grid", gap: 14 }}>
      <div style={{ display: "flex", gap: 16, flexWrap: "wrap" }}>
        <Stat v={e.rounds?.length || 0} l="rounds" />
        <Stat v={`${hearingMin(e)}m`} l="hearing" />
        <Stat v={`${readingMin(e)}m${readingPages(e) ? ` · ${readingPages(e)}p` : ""}`} l="reading" />
        <Stat v={dayScore(e)} l="day score" />
      </div>
      {e.rounds?.length > 0 && (
        <div>
          <div style={{ fontSize: 12, color: C.faint, marginBottom: 4 }}>Rounds</div>
          {e.rounds.map((r, i) => <div key={i} style={row}><span>#{i + 1}</span><span style={{ color: C.faint }}>{fmtT(r.start)} → {fmtT(r.end)} · {mins(r)} min</span></div>)}
        </div>
      )}
      {(e.hearing || []).length > 0 && (
        <div>
          <div style={{ fontSize: 12, color: C.faint, marginBottom: 4 }}>Hearing</div>
          {e.hearing.map((h, i) => <div key={i} style={row}><span>{h.speaker}</span><span style={{ color: C.faint }}>{h.minutes} min{h.ts ? ` · ${fmtT(h.ts)}` : ""}</span></div>)}
        </div>
      )}
      {(e.reading || []).length > 0 && (
        <div>
          <div style={{ fontSize: 12, color: C.faint, marginBottom: 4 }}>Reading</div>
          {e.reading.map((r, i) => <div key={i} style={row}><span>{r.book}{r.section ? ` · ${r.section}` : ""}</span><span style={{ color: C.faint }}>{r.minutes} min{r.pages ? ` · ${r.pages}p` : ""}</span></div>)}
        </div>
      )}
      {prayerRows(e).length > 0 && (
        <div>
          <div style={{ fontSize: 12, color: C.faint, marginBottom: 4 }}>Prayers</div>
          {prayerRows(e).map((p, i) => <div key={i} style={row}><span>{p.label}</span><span style={{ color: C.faint }}>{p.ts ? fmtT(p.ts) : ""}</span></div>)}
        </div>
      )}
      {dressingRows(e).length > 0 && (
        <div>
          <div style={{ fontSize: 12, color: C.faint, marginBottom: 4 }}>Deity dressing</div>
          {dressingRows(e).map((d, i) => <div key={i} style={row}><span>{d.label}</span><span style={{ color: C.faint }}>{d.ts ? fmtT(d.ts) : ""}</span></div>)}
        </div>
      )}
      {WORSHIP_ITEMS.some(([k]) => e.worship?.[k]) && (
        <div>
          <div style={{ fontSize: 12, color: C.faint, marginBottom: 4 }}>Worship</div>
          <div style={{ fontSize: 13 }}>{WORSHIP_ITEMS.filter(([k]) => e.worship?.[k]).map(([, l]) => l).join(" · ")}</div>
        </div>
      )}
      <div style={{ display: "flex", gap: 16, flexWrap: "wrap", fontSize: 13, color: C.faint }}>
        {e.wakeTime && <span>Woke {e.wakeTime}</span>}
        {e.sleepTime && <span>Slept {e.sleepTime}</span>}
        {awakeMin(e) > 0 && <span>Awake {awakeCount(e)}×/{awakeMin(e)}m</span>}
        {(e.versesRecited || e.versesRevised) && <span>Verses practised</span>}
        {e.journaled && <span>Journaled</span>}
      </div>
      {e.journaled && e.journal && (
        <div>
          <div style={{ fontSize: 12, color: C.faint, marginBottom: 4 }}>Journal</div>
          <div style={{ fontSize: 13, whiteSpace: "pre-wrap" }}>{e.journal}</div>
        </div>
      )}
      {e.note && (
        <div>
          <div style={{ fontSize: 12, color: C.faint, marginBottom: 4 }}>Note</div>
          <div style={{ fontSize: 13 }}>{e.note}</div>
        </div>
      )}
    </div>
  );
}

function Viz({ n, t, note, children }) {
  return (
    <section style={cardS}>
      <div style={{ display: "flex", gap: 8, alignItems: "baseline" }}>
        <span style={{ fontSize: 11, color: C.saffron, fontWeight: 700, fontFamily: "monospace" }}>#{String(n).padStart(2, "0")}</span>
        <h2 style={{ ...h2S, margin: 0 }}>{t}</h2>
      </div>
      <div style={{ marginTop: 10 }}>{children}</div>
      {note && <div style={{ fontSize: 12, color: C.faint, marginTop: 6 }}>{note}</div>}
    </section>
  );
}
const Empty = ({ m = "Not enough data yet — keep logging." }) => <div style={{ fontSize: 13, color: C.faint }}>{m}</div>;
function Stat({ v, l, sub }) {
  return (
    <div>
      <div style={{ fontFamily: "'Fraunces', Georgia, serif", fontSize: 34, fontWeight: 700, color: C.maroon }}>{v}</div>
      <div style={{ fontSize: 12, color: C.faint, textTransform: "uppercase", letterSpacing: ".06em" }}>{l}</div>
      {sub && <div style={{ fontSize: 13, color: C.ink, marginTop: 4 }}>{sub}</div>}
    </div>
  );
}
function Rank({ items, unit = "" }) {
  if (!items.length) return <Empty />;
  const max = Math.max(...items.map((i) => i[1]));
  return items.map(([l, v, extra], i) => (
    <div key={i} style={{ margin: "6px 0" }}>
      <div style={{ display: "flex", justifyContent: "space-between", fontSize: 13 }}>
        <span>{l}{extra && <span style={{ color: C.faint }}> {extra}</span>}</span>
        <b>{typeof v === "number" ? Math.round(v * 10) / 10 : v}{unit}</b>
      </div>
      <div style={{ height: 6, background: C.ivory, borderRadius: 3 }}>
        <div style={{ height: 6, width: `${max ? (v / max) * 100 : 0}%`, background: PAL[i % PAL.length], borderRadius: 3 }} />
      </div>
    </div>
  ));
}
function Gauge({ pct, label }) {
  const p = Math.max(0, Math.min(1, pct || 0));
  const ang = -180 + p * 180;
  return (
    <div style={{ textAlign: "center" }}>
      <svg viewBox="0 0 200 110" width="180">
        <path d="M10 100 A90 90 0 0 1 190 100" fill="none" stroke={C.ivory} strokeWidth="16" strokeLinecap="round" />
        <path d="M10 100 A90 90 0 0 1 190 100" fill="none" stroke={C.saffron} strokeWidth="16" strokeLinecap="round"
          strokeDasharray={`${p * 283} 283`} />
        <line x1="100" y1="100" x2={100 + 70 * Math.cos(ang * Math.PI / 180)} y2={100 + 70 * Math.sin(ang * Math.PI / 180)} stroke={C.maroon} strokeWidth="3" />
        <text x="100" y="92" textAnchor="middle" style={{ font: "700 22px Georgia", fill: C.maroon }}>{Math.round(p * 100)}%</text>
      </svg>
      <div style={{ fontSize: 12, color: C.faint }}>{label}</div>
    </div>
  );
}
function MatrixGrid({ rows, cols, value, colorFn, cell = 16, fmt }) {
  return (
    <div style={{ overflowX: "auto" }}>
      <div style={{ display: "grid", gridTemplateColumns: `64px repeat(${cols.length}, ${cell + 4}px)`, gap: 2, width: "max-content", alignItems: "center" }}>
        <div />
        {cols.map((c, j) => <div key={j} style={{ fontSize: 9, color: C.faint, textAlign: "center" }}>{c}</div>)}
        {rows.map((r, i) => (
          <React.Fragment key={i}>
            <div style={{ fontSize: 10, color: C.faint, paddingRight: 4 }}>{r}</div>
            {cols.map((c, j) => {
              const v = value(i, j);
              return <div key={j} title={fmt ? fmt(i, j, v) : v} style={{ width: cell, height: cell, borderRadius: 3, background: colorFn(v) }} />;
            })}
          </React.Fragment>
        ))}
      </div>
    </div>
  );
}
function BarsViz({ data, x, bars, h = 180, stack, dom, hideLegend }) {
  if (!data.length) return <Empty />;
  return (
    <ResponsiveContainer width="100%" height={h}>
      <BarChart data={data} margin={{ top: 6, right: 6, left: -24, bottom: 0 }}>
        <CartesianGrid stroke={C.line} strokeDasharray="3 3" />
        <XAxis dataKey={x} tick={tickS} interval="preserveStartEnd" />
        <YAxis tick={tickS} domain={dom} />
        <Tooltip />
        {bars.length > 1 && !hideLegend && <Legend wrapperStyle={{ fontSize: 11 }} />}
        {bars.map((b, i) => <Bar key={b.k} dataKey={b.k} name={b.name || b.k} stackId={stack ? "s" : undefined} fill={b.c || PAL[i]} radius={stack ? 0 : [3, 3, 0, 0]} />)}
      </BarChart>
    </ResponsiveContainer>
  );
}
function LinesViz({ data, x, lines, h = 180, dom, refY, refLabel }) {
  if (!data.length) return <Empty />;
  return (
    <ResponsiveContainer width="100%" height={h}>
      <LineChart data={data} margin={{ top: 6, right: 6, left: -24, bottom: 0 }}>
        <CartesianGrid stroke={C.line} strokeDasharray="3 3" />
        <XAxis dataKey={x} tick={tickS} interval="preserveStartEnd" />
        <YAxis tick={tickS} domain={dom} />
        <Tooltip />
        {lines.length > 1 && <Legend wrapperStyle={{ fontSize: 11 }} />}
        {refY != null && <ReferenceLine y={refY} stroke={C.tulsi} strokeDasharray="4 4" label={{ value: refLabel, fontSize: 10, fill: C.tulsi }} />}
        {lines.map((l, i) => <Line key={l.k} type="monotone" dataKey={l.k} name={l.name || l.k} stroke={l.c || PAL[i]} strokeWidth={2} dot={false} connectNulls />)}
      </LineChart>
    </ResponsiveContainer>
  );
}
function PtsViz({ data, xk, yk, h = 190, xdom, ydom, color = C.maroon, xfmt, yfmt, yrev, zk, trend }) {
  if (data.length < 2) return <Empty />;
  let tl = null;
  if (trend) {
    const xs = data.map((d) => d[xk]), ys = data.map((d) => d[yk]);
    const mx = mean(xs), my = mean(ys);
    let num = 0, den = 0;
    xs.forEach((x, i) => { num += (x - mx) * (ys[i] - my); den += (x - mx) ** 2; });
    if (den) {
      const b = num / den, a = my - b * mx;
      const x0 = Math.min(...xs), x1 = Math.max(...xs);
      tl = [{ [xk]: x0, [yk]: a + b * x0 }, { [xk]: x1, [yk]: a + b * x1 }];
    }
  }
  return (
    <ResponsiveContainer width="100%" height={h}>
      <ScatterChart margin={{ top: 6, right: 6, left: -18, bottom: 0 }}>
        <CartesianGrid stroke={C.line} strokeDasharray="3 3" />
        <XAxis type="number" dataKey={xk} domain={xdom || ["auto", "auto"]} tick={tickS} tickFormatter={xfmt} />
        <YAxis type="number" dataKey={yk} domain={ydom || ["auto", "auto"]} tick={tickS} tickFormatter={yfmt} reversed={yrev} />
        {zk && <ZAxis type="number" dataKey={zk} range={[20, 90]} />}
        <Tooltip formatter={(v) => Math.round(v * 100) / 100} />
        <Scatter data={data} fill={color} fillOpacity={0.65} />
        {tl && <Scatter data={tl} line={{ stroke: C.tulsi, strokeWidth: 2 }} shape={() => null} />}
      </ScatterChart>
    </ResponsiveContainer>
  );
}
function DonutViz({ data, h = 190 }) {
  if (!data.filter((d) => d.value > 0).length) return <Empty />;
  return (
    <ResponsiveContainer width="100%" height={h}>
      <PieChart>
        <Pie data={data} dataKey="value" nameKey="name" innerRadius="50%" outerRadius="80%" paddingAngle={2}>
          {data.map((d, i) => <Cell key={i} fill={PAL[i % PAL.length]} />)}
        </Pie>
        <Tooltip /><Legend wrapperStyle={{ fontSize: 11 }} />
      </PieChart>
    </ResponsiveContainer>
  );
}

// ============ Visual groups ============
// A = analytics object (see App). Each group renders its numbered Viz cards.

function JapaDurationViz({ A }) {
  const { dayList, allRounds, today } = A;
  const withR = dayList.filter((d) => d.rounds.length);
  // 1 violin-ish per weekday
  const wdStats = WD.map((w, wd) => {
    const ds = allRounds.filter((r) => r.wd === wd).map((r) => r.dur);
    return { w, lo: ds.length ? quantile(ds, 0.1) : null, q1: quantile(ds, 0.25), m: median(ds), q3: quantile(ds, 0.75), hi: ds.length ? quantile(ds, 0.9) : null, n: ds.length };
  });
  // 2 rolling 7d avg dur vs best week
  const roll = withR.map((d, i) => {
    const win = withR.slice(Math.max(0, i - 6), i + 1).flatMap((x) => x.rounds.map((r) => r.dur));
    return { label: d.k.slice(5), v: +mean(win).toFixed(1) };
  });
  const bestWeek = roll.length ? Math.min(...roll.map((r) => r.v)) : null;
  // 3 fatigue: last - first round duration
  const fatigue = withR.filter((d) => d.rounds.length >= 4).map((d) => ({ label: d.k.slice(5), gap: +(d.rounds[d.rounds.length - 1].dur - d.rounds[0].dur).toFixed(1) }));
  // 4 round# x avg dur
  const byIdx = Array.from({ length: TARGET }, (_, i) => {
    const ds = allRounds.filter((r) => r.i === i).map((r) => r.dur);
    return ds.length ? mean(ds) : null;
  });
  const maxIdx = Math.max(...byIdx.filter((v) => v != null), 1);
  // 5 daily japa minutes area
  const dailyMin = dayList.slice(-42).map((d) => ({ label: d.k.slice(5), min: +d.rounds.reduce((a, r) => a + r.dur, 0).toFixed(0) }));
  // 6 histogram
  const buckets = {};
  allRounds.forEach((r) => { const b = Math.min(20, Math.floor(r.dur / 2) * 2); buckets[b] = (buckets[b] || 0) + 1; });
  const hist = Object.entries(buckets).map(([b, n]) => ({ b: `${b}-${+b + 2}`, n, bb: +b })).sort((a, x) => a.bb - x.bb);
  // 7 weekly box
  const weeks = {};
  allRounds.forEach((r) => {
    const d = new Date(r.date); d.setDate(d.getDate() - d.getDay());
    const wk = todayKey(d);
    (weeks[wk] = weeks[wk] || []).push(r.dur);
  });
  const weekBox = Object.entries(weeks).sort().slice(-8).map(([wk, ds]) => ({ wk: wk.slice(5), lo: quantile(ds, 0.1), q1: quantile(ds, 0.25), m: median(ds), q3: quantile(ds, 0.75), hi: quantile(ds, 0.9) }));
  // 8 PR
  const clean = allRounds.filter((r) => r.dur >= 3);
  const pr = clean.length ? clean.reduce((a, b) => (a.dur < b.dur ? a : b)) : null;
  // 9 dur vs hour
  const dvh = allRounds.filter((r) => r.hour != null).map((r) => ({ hour: r.hour, dur: r.dur }));
  // 11 daily std
  const consist = withR.filter((d) => d.rounds.length >= 4).map((d) => ({ label: d.k.slice(5), sd: +std(d.rounds.map((r) => r.dur)).toFixed(2) }));
  // 12 ekadasi vs normal
  const ekD = allRounds.filter((r) => A.byKey[r.k]?.ek).map((r) => r.dur);
  const noD = allRounds.filter((r) => !A.byKey[r.k]?.ek).map((r) => r.dur);
  // 13 waterfall today
  const tRounds = today?.rounds || [];
  let cum = 0;
  const wf = tRounds.map((r, i) => { const s = cum; cum += r.dur; return { r: `#${i + 1}`, pad: +s.toFixed(0), dur: +r.dur.toFixed(1) }; });
  // 14 golden hour
  const last30 = allRounds.filter((r) => (new Date() - r.date) / 864e5 <= 30 && r.hour != null);
  const golden = last30.length ? last30.filter((r) => r.hour < SUNRISE[r.date.getMonth()] + 1.5).length / last30.length : null;
  return (
    <div style={{ display: "grid", gap: 14 }}>
      <Viz n={1} t="Round duration by weekday" note="Bar = 25–75th percentile, tick = median, whisker = 10–90th. Wide Mondays? That's the weekend tax.">
        {allRounds.length < 8 ? <Empty /> : (
          <svg viewBox="0 0 340 150" width="100%">
            {wdStats.map((s, i) => s.n ? (
              <g key={i} transform={`translate(${20 + i * 46},0)`}>
                <line x1="14" x2="14" y1={140 - s.hi * 6} y2={140 - s.lo * 6} stroke={C.faint} />
                <rect x="4" width="20" y={140 - s.q3 * 6} height={Math.max(2, (s.q3 - s.q1) * 6)} fill={C.saffronSoft} stroke={C.saffron} rx="3" />
                <line x1="4" x2="24" y1={140 - s.m * 6} y2={140 - s.m * 6} stroke={C.maroon} strokeWidth="2" />
                <text x="14" y="149" textAnchor="middle" style={{ font: "9px sans-serif", fill: C.faint }}>{s.w}</text>
              </g>
            ) : null)}
          </svg>
        )}
      </Viz>
      <Viz n={2} t="Distraction index" note="7-day rolling average round length. Green line = your best week — the benchmark to chase.">
        <LinesViz data={roll.slice(-60)} x="label" lines={[{ k: "v", name: "Avg min/round", c: C.maroon }]} refY={bestWeek} refLabel="best" />
      </Viz>
      <Viz n={3} t="Fatigue curve — last vs first round" note="Positive = your final rounds run longer than your first. The mind tires; the timer notices.">
        <BarsViz data={fatigue.slice(-30)} x="label" bars={[{ k: "gap", name: "Last − first (min)", c: C.sky }]} />
      </Viz>
      <Viz n={4} t="Which round number sags" note="Average duration by round position 1–16. A bulge at 10–14 is the classic mid-japa wobble.">
        <div style={{ display: "grid", gridTemplateColumns: "repeat(16, 1fr)", gap: 3 }}>
          {byIdx.map((v, i) => (
            <div key={i} style={{ textAlign: "center" }}>
              <div title={v ? `${v.toFixed(1)} min` : ""} style={{ height: 40, borderRadius: 4, background: v == null ? C.ivory : `rgba(224,123,31,${0.15 + 0.85 * (v / maxIdx)})` }} />
              <div style={{ fontSize: 8, color: C.faint }}>{i + 1}</div>
            </div>
          ))}
        </div>
      </Viz>
      <Viz n={5} t="Daily japa minutes" note="Total time on the beads each day — volume, not just count.">
        {dailyMin.length ? (
          <ResponsiveContainer width="100%" height={170}>
            <AreaChart data={dailyMin} margin={{ top: 6, right: 6, left: -24, bottom: 0 }}>
              <CartesianGrid stroke={C.line} strokeDasharray="3 3" />
              <XAxis dataKey="label" tick={tickS} interval="preserveStartEnd" /><YAxis tick={tickS} /><Tooltip />
              <Area type="monotone" dataKey="min" stroke={C.saffron} fill={C.saffronSoft} />
            </AreaChart>
          </ResponsiveContainer>
        ) : <Empty />}
      </Viz>
      <Viz n={6} t="Duration histogram" note="Where your rounds live. A tight peak is a settled mind; a long right tail is wandering.">
        <BarsViz data={hist} x="b" bars={[{ k: "n", name: "Rounds", c: C.saffron }]} />
      </Viz>
      <Viz n={7} t="Weekly spread (boxplot)" note="Shrinking boxes week-on-week = growing steadiness.">
        {weekBox.length < 2 ? <Empty /> : (
          <svg viewBox={`0 0 ${40 + weekBox.length * 40} 150`} width="100%">
            {weekBox.map((s, i) => (
              <g key={i} transform={`translate(${20 + i * 40},0)`}>
                <line x1="14" x2="14" y1={140 - s.hi * 6} y2={140 - s.lo * 6} stroke={C.faint} />
                <rect x="4" width="20" y={140 - s.q3 * 6} height={Math.max(2, (s.q3 - s.q1) * 6)} fill="#EAF0E4" stroke={C.tulsi} rx="3" />
                <line x1="4" x2="24" y1={140 - s.m * 6} y2={140 - s.m * 6} stroke={C.maroon} strokeWidth="2" />
                <text x="14" y="149" textAnchor="middle" style={{ font: "8px sans-serif", fill: C.faint }}>{s.wk}</text>
              </g>
            ))}
          </svg>
        )}
      </Viz>
      <Viz n={8} t="Personal record" note="Fastest clean round (≥3 min counted as genuine).">
        {pr ? <Stat v={`${pr.dur} min`} l="fastest round" sub={`${fmtDate(pr.k)} at ${fmtT(pr.start)}`} /> : <Empty />}
      </Viz>
      <Viz n={9} t="Duration vs time of day" note="Trend line up and to the right = later rounds are slower rounds.">
        <PtsViz data={dvh} xk="hour" yk="dur" xdom={[4, 23]} xfmt={fmtH} trend />
      </Viz>
      <Viz n={10} t="Streak spiral" note="Last 120 days wound inward — saffron = all 16 rounds. An unbroken arm is the goal.">
        <StreakSpiral A={A} />
      </Viz>
      <Viz n={11} t="Pace consistency" note="Standard deviation of round lengths within each day. Lower = more even chanting.">
        <LinesViz data={consist.slice(-30)} x="label" lines={[{ k: "sd", name: "Std dev (min)", c: C.tulsi }]} />
      </Viz>
      <Viz n={12} t="Ekadasi effect" note={`Detected from notes containing "ekadasi". Avg round length compared.`}>
        {ekD.length ? <BarsViz data={[{ d: "Ekadasi", v: +mean(ekD).toFixed(1) }, { d: "Other days", v: +mean(noD).toFixed(1) }]} x="d" bars={[{ k: "v", name: "Avg min/round", c: C.gold }]} /> : <Empty m='No Ekadasi days found — write "Ekadasi" in the day note to tag them.' />}
      </Viz>
      <Viz n={13} t="Today's waterfall" note="How each round stacked into today's total japa time.">
        {wf.length ? (
          <ResponsiveContainer width="100%" height={170}>
            <BarChart data={wf} margin={{ top: 6, right: 6, left: -24, bottom: 0 }}>
              <XAxis dataKey="r" tick={tickS} /><YAxis tick={tickS} /><Tooltip />
              <Bar dataKey="pad" stackId="w" fill="transparent" />
              <Bar dataKey="dur" stackId="w" fill={C.saffron} radius={[3, 3, 0, 0]} />
            </BarChart>
          </ResponsiveContainer>
        ) : <Empty m="No rounds today yet." />}
      </Viz>
      <Viz n={14} t="Golden hour rate" note="Share of rounds (30d) chanted within ~90 min of London sunrise.">
        {golden == null ? <Empty /> : <Gauge pct={golden} label="rounds in the golden hour" />}
      </Viz>
      <Viz n={15} t="Today's bead ring" note="The day's mala, live. Filled beads pulse gently.">
        <BeadRing n={today?.rounds.length || 0} />
      </Viz>
    </div>
  );
}

function StreakSpiral({ A }) {
  const pts = [];
  const d = new Date();
  for (let i = 0; i < 120; i++) {
    const k = todayKey(d);
    const e = A.byKey[k];
    const ang = -i * 0.32, rad = 12 + i * 0.95;
    pts.push({ x: 130 + rad * Math.cos(ang), y: 130 + rad * Math.sin(ang), full: e && e.n >= TARGET, any: !!e });
    d.setDate(d.getDate() - 1);
  }
  return (
    <svg viewBox="0 0 260 260" width="240" style={{ display: "block", margin: "0 auto" }}>
      {pts.map((p, i) => <circle key={i} cx={p.x} cy={p.y} r={i === 0 ? 5 : 3.4} fill={p.full ? C.saffron : p.any ? C.saffronSoft : "#EFE9DC"} stroke={i === 0 ? C.maroon : "none"} />)}
      <text x="130" y="134" textAnchor="middle" style={{ font: "10px sans-serif", fill: C.faint }}>today→</text>
    </svg>
  );
}
function BeadRing({ n }) {
  return (
    <svg viewBox="0 0 220 220" width="200" style={{ display: "block", margin: "0 auto" }}>
      <style>{`@keyframes pulse{0%,100%{opacity:1}50%{opacity:.55}} .pb{animation:pulse 2.4s ease-in-out infinite}`}</style>
      {Array.from({ length: TARGET }).map((_, i) => {
        const a = (i / TARGET) * 2 * Math.PI - Math.PI / 2;
        return <circle key={i} className={i < n ? "pb" : ""} cx={110 + 85 * Math.cos(a)} cy={110 + 85 * Math.sin(a)} r="11"
          fill={i < n ? C.saffron : "#fff"} stroke={i < n ? C.maroon : C.line} strokeWidth="2" style={{ animationDelay: `${i * 0.15}s` }} />;
      })}
      <text x="110" y="116" textAnchor="middle" style={{ font: "700 24px Georgia", fill: C.maroon }}>{n}/{TARGET}</text>
    </svg>
  );
}

function JapaTimeViz({ A }) {
  const { dayList, allRounds } = A;
  const withR = dayList.filter((d) => d.rounds.length);
  // 17 first start trend
  const firstT = withR.slice(-90).map((d) => ({ label: d.k.slice(5), h: d.firstH && +d.firstH.toFixed(2) }));
  // 18 last end vs work
  const lastT = withR.slice(-60).map((d) => ({ label: d.k.slice(5), h: d.lastEndH && +d.lastEndH.toFixed(2) }));
  // 19 max gap
  const gaps = withR.filter((d) => d.rounds.length >= 2).map((d) => {
    const sorted = [...d.rounds].sort((a, b) => new Date(a.start) - new Date(b.start));
    let g = 0;
    for (let i = 1; i < sorted.length; i++) g = Math.max(g, (new Date(sorted[i].start) - new Date(sorted[i - 1].end)) / 36e5);
    return { label: d.k.slice(5), g: +g.toFixed(1) };
  });
  // 20 sittings
  const sittings = withR.slice(-30).map((d) => {
    const sorted = [...d.rounds].sort((a, b) => new Date(a.start) - new Date(b.start));
    let s = sorted.length ? 1 : 0;
    for (let i = 1; i < sorted.length; i++) if ((new Date(sorted[i].start) - new Date(sorted[i - 1].end)) / 6e4 > 15) s++;
    return { label: d.k.slice(5), sittings: s };
  });
  // 21 sunrise overlay
  const sunPts = allRounds.filter((r) => r.hour != null && (new Date() - r.date) / 864e5 <= 60)
    .map((r) => ({ x: +(60 - (new Date() - r.date) / 864e5).toFixed(1), hour: r.hour, sr: SUNRISE[r.date.getMonth()] }));
  // 22 weekday vs weekend first start
  const wdF = withR.filter((d) => d.wd >= 1 && d.wd <= 5).map((d) => d.firstH).filter(Boolean);
  const weF = withR.filter((d) => d.wd === 0 || d.wd === 6).map((d) => d.firstH).filter(Boolean);
  // 23 morning completion (7d)
  const l7 = allRounds.filter((r) => (new Date() - r.date) / 864e5 <= 7 && r.hour != null);
  const morning = l7.length ? l7.filter((r) => r.hour < 8).length / l7.length : null;
  // 24 calendar by median hour
  // 25 density by hour
  const dens = Array.from({ length: 20 }, (_, i) => ({ h: `${i + 4}`, n: allRounds.filter((r) => r.hour != null && Math.floor(r.hour) === i + 4).length }));
  return (
    <div style={{ display: "grid", gap: 14 }}>
      <Viz n={16} t="Polar japa clock (7 days)" note="Each arc is a round placed on the 24-hour dial. Devotion should crowd the top-left (pre-dawn).">
        <PolarClock A={A} />
      </Viz>
      <Viz n={17} t="First-round start time — 90 days" note="The single best predictor of a complete day.">
        <LinesViz data={firstT} x="label" lines={[{ k: "h", name: "Start hour", c: C.maroon }]} dom={[4, 12]} refY={SUNRISE[new Date().getMonth()]} refLabel="sunrise" />
      </Viz>
      <Viz n={18} t="Last round finished vs work" note="Green line = 9:00. Above it means japa leaked past the start of the working day.">
        <LinesViz data={lastT} x="label" lines={[{ k: "h", name: "Last round end", c: C.sky }]} refY={9} refLabel="work" />
      </Viz>
      <Viz n={19} t="Longest japa gap per day" note="Hours between sittings. Big gaps = rounds carried around all day like unfinished business.">
        <BarsViz data={gaps.slice(-30)} x="label" bars={[{ k: "g", name: "Max gap (h)", c: C.gold }]} />
      </Viz>
      <Viz n={20} t="One sitting or many?" note="Number of distinct sittings (15+ min break splits them). One is ideal.">
        <BarsViz data={sittings} x="label" bars={[{ k: "sittings", c: C.tulsi }]} />
      </Viz>
      <Viz n={21} t="Rounds vs the sunrise line" note="Dots below the green line were chanted before sunrise — brahma-muhurta territory.">
        {sunPts.length < 2 ? <Empty /> : (
          <ResponsiveContainer width="100%" height={200}>
            <ComposedChart data={sunPts} margin={{ top: 6, right: 6, left: -18, bottom: 0 }}>
              <CartesianGrid stroke={C.line} strokeDasharray="3 3" />
              <XAxis type="number" dataKey="x" tick={tickS} tickFormatter={(v) => `${Math.round(60 - v)}d`} />
              <YAxis type="number" dataKey="hour" domain={[3, 23]} reversed tick={tickS} tickFormatter={fmtH} />
              <Tooltip formatter={(v) => fmtH(v)} />
              <Line type="monotone" dataKey="sr" stroke={C.tulsi} dot={false} name="Sunrise" strokeWidth={2} />
              <Scatter dataKey="hour" fill={C.maroon} fillOpacity={0.6} name="Round start" />
            </ComposedChart>
          </ResponsiveContainer>
        )}
      </Viz>
      <Viz n={22} t="Weekday vs weekend start" note="How much does the alarm clock owe to the office?">
        {wdF.length && weF.length ? <BarsViz data={[{ d: "Weekday", v: +mean(wdF).toFixed(2) }, { d: "Weekend", v: +mean(weF).toFixed(2) }]} x="d" bars={[{ k: "v", name: "Avg first-round hour", c: C.saffron }]} /> : <Empty />}
      </Viz>
      <Viz n={23} t="Morning completion (7d)" note="Share of rounds finished before 8am this week.">
        {morning == null ? <Empty /> : <Gauge pct={morning} label="rounds before 8am" />}
      </Viz>
      <Viz n={24} t="Calendar of median chanting hour" note="Each day coloured by when the middle round happened — dark = early, pale = late.">
        <HourCalendar A={A} />
      </Viz>
      <Viz n={25} t="Start-time density" note="All-time distribution of when rounds begin. One sharp pre-dawn spike is the ideal shape.">
        <BarsViz data={dens} x="h" bars={[{ k: "n", name: "Rounds", c: C.maroon }]} />
      </Viz>
    </div>
  );
}
function PolarClock({ A }) {
  const rounds = A.allRounds.filter((r) => r.hour != null && (new Date() - r.date) / 864e5 <= 7);
  return (
    <svg viewBox="0 0 240 240" width="230" style={{ display: "block", margin: "0 auto" }}>
      <circle cx="120" cy="120" r="100" fill="none" stroke={C.line} />
      {[0, 6, 12, 18].map((h) => {
        const a = (h / 24) * 2 * Math.PI - Math.PI / 2;
        return <text key={h} x={120 + 112 * Math.cos(a)} y={124 + 112 * Math.sin(a)} textAnchor="middle" style={{ font: "10px sans-serif", fill: C.faint }}>{h}:00</text>;
      })}
      {rounds.map((r, i) => {
        const a0 = (r.hour / 24) * 2 * Math.PI - Math.PI / 2;
        const a1 = ((r.hour + r.dur / 60) / 24) * 2 * Math.PI - Math.PI / 2;
        const rad = 55 + ((new Date() - r.date) / 864e5 / 7) * 42;
        return <path key={i} d={`M ${120 + rad * Math.cos(a0)} ${120 + rad * Math.sin(a0)} A ${rad} ${rad} 0 0 1 ${120 + rad * Math.cos(a1)} ${120 + rad * Math.sin(a1)}`}
          fill="none" stroke={C.saffron} strokeWidth="5" strokeLinecap="round" opacity="0.75" />;
      })}
      {!rounds.length && <text x="120" y="124" textAnchor="middle" style={{ font: "11px sans-serif", fill: C.faint }}>No timed rounds this week</text>}
    </svg>
  );
}
function HourCalendar({ A }) {
  const weeks = 12, cells = [], today = new Date();
  const start = new Date(today); start.setDate(today.getDate() - (weeks * 7 - 1));
  for (let i = 0; i < weeks * 7; i++) {
    const d = new Date(start); d.setDate(start.getDate() + i);
    const k = todayKey(d), e = A.byKey[k];
    const med = e && e.rounds.length ? median(e.rounds.map((r) => r.hour).filter((h) => h != null)) : null;
    cells.push({ k, med, future: d > today });
  }
  const color = (m) => m == null ? "#EFE9DC" : m < 6 ? C.maroon : m < 8 ? C.saffron : m < 12 ? "#EBA85C" : C.saffronSoft;
  return (
    <div style={{ overflowX: "auto" }}>
      <div style={{ display: "grid", gridTemplateRows: "repeat(7,13px)", gridAutoFlow: "column", gap: 3, width: "max-content" }}>
        {cells.map((c) => <div key={c.k} title={`${fmtDate(c.k)} — median ${c.med ? fmtH(c.med) : "—"}`} style={{ width: 13, height: 13, borderRadius: 3, background: c.future ? "transparent" : color(c.med) }} />)}
      </div>
      <div style={{ fontSize: 11, color: C.faint, marginTop: 6 }}>maroon &lt;6am · saffron &lt;8am · pale = later</div>
    </div>
  );
}

function HearingViz({ A }) {
  const { dayList } = A;
  const sessions = dayList.flatMap((d) => (d.e.hearing || []).map((h) => ({ ...h, k: d.k, month: d.month, ts: h.ts })));
  const bySpeaker = {};
  sessions.forEach((s) => { bySpeaker[s.speaker] = (bySpeaker[s.speaker] || 0) + (+s.minutes || 0); });
  const speakers = Object.entries(bySpeaker).sort((a, b) => b[1] - a[1]);
  // 27 month dominance
  const months = [...new Set(sessions.map((s) => s.month))].sort().slice(-6);
  const top5 = speakers.slice(0, 5).map((s) => s[0]);
  const monthDom = months.map((m) => {
    const row = { m: m.slice(2) };
    top5.forEach((sp) => { row[sp] = sessions.filter((s) => s.month === m && s.speaker === sp).reduce((a, s) => a + +s.minutes, 0); });
    row.Other = sessions.filter((s) => s.month === m && !top5.includes(s.speaker)).reduce((a, s) => a + +s.minutes, 0);
    return row;
  });
  // 28 prabhupada %
  const sp = months.map((m) => {
    const all = sessions.filter((s) => s.month === m);
    const tot = all.reduce((a, s) => a + +s.minutes, 0);
    const pp = all.filter((s) => /prabhupada/i.test(s.speaker)).reduce((a, s) => a + +s.minutes, 0);
    return { m: m.slice(2), pct: tot ? Math.round((pp / tot) * 100) : 0 };
  });
  // 29 hearing vs next-day avg round
  const lag = [];
  dayList.forEach((d, i) => { const n = dayList[i + 1]; if (n && n.avgDur && d.hm > 0) lag.push({ hm: d.hm, dur: +n.avgDur.toFixed(1) }); });
  // 30 speaker x daypart
  const part = (ts) => { if (!ts) return "Unlogged"; const h = new Date(ts).getHours(); return h < 12 ? "Morning" : h < 18 ? "Afternoon" : "Evening"; };
  const dayparts = ["Morning", "Afternoon", "Evening", "Unlogged"];
  const spDp = top5.map((spk) => {
    const row = { spk: spk.length > 12 ? spk.slice(0, 12) + "…" : spk };
    dayparts.forEach((p) => { row[p] = sessions.filter((s) => s.speaker === spk && part(s.ts) === p).reduce((a, s) => a + +s.minutes, 0); });
    return row;
  });
  // 31 cumulative
  const cumData = dayList.map((d) => {
    const row = { label: d.k.slice(5) };
    return { d, row };
  });
  const cumTotals = {};
  const cumSeries = cumData.map(({ d, row }) => {
    (d.e.hearing || []).forEach((h) => { cumTotals[h.speaker] = (cumTotals[h.speaker] || 0) + +h.minutes; });
    top5.forEach((spk) => { row[spk] = +((cumTotals[spk] || 0) / 60).toFixed(1); });
    return row;
  });
  // 32 new voices
  const seen = new Set();
  const newV = months.map((m) => {
    const fresh = new Set();
    sessions.filter((s) => s.month === m).forEach((s) => { if (!seen.has(s.speaker)) { seen.add(s.speaker); fresh.add(s.speaker); } });
    return { m: m.slice(2), n: fresh.size };
  });
  // 34 avg session per speaker
  const avgSess = speakers.slice(0, 8).map(([spk]) => {
    const ss = sessions.filter((s) => s.speaker === spk);
    return [spk, mean(ss.map((s) => +s.minutes)), `(${ss.length}×)`];
  });
  const maxMin = speakers.length ? speakers[0][1] : 1;
  return (
    <div style={{ display: "grid", gap: 14 }}>
      <Viz n={26} t="Who fills your ears" note="Total minutes per speaker, all time.">
        {speakers.length ? (
          <ResponsiveContainer width="100%" height={200}>
            <Treemap data={speakers.map(([n, v], i) => ({ name: n, size: v, fill: PAL[i % PAL.length] }))} dataKey="size" stroke="#fff" content={<TreemapCell />} />
          </ResponsiveContainer>
        ) : <Empty />}
      </Viz>
      <Viz n={27} t="Speaker mix by month" note="Stacked minutes for your top voices — is one teacher crowding the rest out?">
        <BarsViz data={monthDom} x="m" stack bars={[...top5.map((s, i) => ({ k: s, c: PAL[i] })), { k: "Other", c: "#D8D0C0" }]} h={210} />
      </Viz>
      <Viz n={28} t="Prabhupada share" note="Percent of hearing that is Srila Prabhupada directly, by month.">
        <LinesViz data={sp} x="m" lines={[{ k: "pct", name: "%", c: C.maroon }]} dom={[0, 100]} />
      </Viz>
      <Viz n={29} t="Hearing today, japa tomorrow" note="Each dot: minutes heard vs next morning's average round length. Down-slope = hearing sharpens japa.">
        <PtsViz data={lag} xk="hm" yk="dur" trend />
      </Viz>
      <Viz n={30} t="Speaker × time of day" note="When each voice gets your attention (entries logged from now on carry a timestamp).">
        <BarsViz data={spDp} x="spk" stack bars={dayparts.map((p, i) => ({ k: p, c: [C.saffron, C.gold, C.maroon, "#D8D0C0"][i] }))} h={200} />
      </Viz>
      <Viz n={31} t="Cumulative hours per speaker" note="The long race — hours of association accumulated.">
        <LinesViz data={cumSeries.slice(-90)} x="label" lines={top5.map((s, i) => ({ k: s, c: PAL[i] }))} h={210} />
      </Viz>
      <Viz n={32} t="New voices per month" note="First-time speakers discovered each month.">
        <BarsViz data={newV} x="m" bars={[{ k: "n", name: "New speakers", c: C.tulsi }]} />
      </Viz>
      <Viz n={33} t="Hearing streak calendar" note="Days with any hearing at all.">
        <BinaryCalendar A={A} test={(d) => d.hm > 0} />
      </Viz>
      <Viz n={34} t="Session length per speaker" note="Long sittings vs snack-sized clips.">
        <Rank items={avgSess} unit=" min" />
      </Viz>
      <Viz n={35} t="Voice cloud" note="Size = total minutes heard.">
        <div style={{ display: "flex", flexWrap: "wrap", gap: "6px 14px", alignItems: "baseline" }}>
          {speakers.length ? speakers.map(([n, v], i) => (
            <span key={n} style={{ fontFamily: "'Fraunces', Georgia, serif", color: PAL[i % PAL.length], fontSize: 12 + (v / maxMin) * 26, fontWeight: 600 }}>{n}</span>
          )) : <Empty />}
        </div>
      </Viz>
    </div>
  );
}
function TreemapCell({ x, y, width, height, name, fill }) {
  if (width < 4 || height < 4) return null;
  return (
    <g>
      <rect x={x} y={y} width={width} height={height} fill={fill || C.saffron} stroke="#fff" rx="3" />
      {width > 56 && height > 22 && <text x={x + 6} y={y + 16} style={{ font: "11px sans-serif", fill: "#fff" }}>{name}</text>}
    </g>
  );
}
function BinaryCalendar({ A, test }) {
  const weeks = 14, cells = [], today = new Date();
  const start = new Date(today); start.setDate(today.getDate() - (weeks * 7 - 1));
  for (let i = 0; i < weeks * 7; i++) {
    const d = new Date(start); d.setDate(start.getDate() + i);
    const k = todayKey(d), e = A.byKey[k];
    cells.push({ k, on: e ? test(e) : false, logged: !!e, future: d > today });
  }
  return (
    <div style={{ overflowX: "auto" }}>
      <div style={{ display: "grid", gridTemplateRows: "repeat(7,13px)", gridAutoFlow: "column", gap: 3, width: "max-content" }}>
        {cells.map((c) => <div key={c.k} title={fmtDate(c.k)} style={{ width: 13, height: 13, borderRadius: 3, background: c.future ? "transparent" : c.on ? C.tulsi : c.logged ? "#EFE9DC" : "#F5F0E4" }} />)}
      </div>
    </div>
  );
}

function ReadingViz({ A }) {
  const { dayList } = A;
  const sessions = dayList.flatMap((d) => (d.e.reading || []).map((r) => ({ ...r, k: d.k, month: d.month, date: d.date })));
  const byBook = {};
  sessions.forEach((s) => {
    const b = byBook[s.book] = byBook[s.book] || { min: 0, n: 0, first: s.k, last: s.k, lastSection: "" };
    b.min += +s.minutes || 0; b.n++; b.last = s.k; if (s.section) b.lastSection = s.section;
  });
  const books = Object.entries(byBook).sort((a, b) => b[1].min - a[1].min);
  const maxBook = books.length ? books[0][1].min : 1;
  // 37/38 canto & chapter coverage
  const sbCantos = new Set(), bgCh = new Set();
  sessions.forEach((s) => {
    const sb = /SB\s*(\d{1,2})/i.exec(s.section || ""); if (sb && +sb[1] >= 1 && +sb[1] <= 12) sbCantos.add(+sb[1]);
    const bg = /BG\s*(\d{1,2})/i.exec(s.section || ""); if (bg && +bg[1] >= 1 && +bg[1] <= 18) bgCh.add(+bg[1]);
  });
  // 39 weekly stacked area per book
  const wkOf = (k) => { const d = new Date(k + "T12:00"); d.setDate(d.getDate() - d.getDay()); return todayKey(d); };
  const wks = [...new Set(sessions.map((s) => wkOf(s.k)))].sort().slice(-10);
  const topB = books.slice(0, 4).map((b) => b[0]);
  const wkArea = wks.map((w) => {
    const row = { w: w.slice(5) };
    topB.forEach((b) => { row[b] = sessions.filter((s) => wkOf(s.k) === w && s.book === b).reduce((a, s) => a + +s.minutes, 0); });
    return row;
  });
  // 42 time-of-day (ts-logged entries)
  const tod = sessions.filter((s) => s.ts).map((s) => { const d = new Date(s.ts); return { hour: d.getHours() + d.getMinutes() / 60, min: +s.minutes }; });
  // 43 histogram
  const buckets = {};
  sessions.forEach((s) => { const b = Math.min(60, Math.floor(+s.minutes / 10) * 10); buckets[b] = (buckets[b] || 0) + 1; });
  const hist = Object.entries(buckets).map(([b, n]) => ({ b: `${b}+`, n, bb: +b })).sort((a, x) => a.bb - x.bb);
  // 44 rotation
  const rot = wks.map((w) => ({ w: w.slice(5), books: new Set(sessions.filter((s) => wkOf(s.k) === w).map((s) => s.book)).size }));
  // 45 monthly 100% split
  const months = [...new Set(sessions.map((s) => s.month))].sort().slice(-6);
  const splitKeys = ["Srimad-Bhagavatam", "Bhagavad-gita", "Caitanya-caritamrta"];
  const split = months.map((m) => {
    const ms = sessions.filter((s) => s.month === m);
    const tot = ms.reduce((a, s) => a + +s.minutes, 0) || 1;
    const row = { m: m.slice(2) };
    splitKeys.forEach((b) => { row[b.split("-")[0]] = Math.round(ms.filter((s) => s.book === b).reduce((a, s) => a + +s.minutes, 0) / tot * 100); });
    row.Other = 100 - splitKeys.reduce((a, b) => a + (row[b.split("-")[0]] || 0), 0);
    return row;
  });
  // 46 report card
  const year = new Date().getFullYear().toString();
  const yearMin = books.map(([b, s]) => [b, sessions.filter((x) => x.book === b && x.k.startsWith(year)).reduce((a, x) => a + +x.minutes, 0)]);
  // 47 longest streak per book
  const streakBook = books.slice(0, 6).map(([b]) => {
    const days = new Set(sessions.filter((s) => s.book === b).map((s) => s.k));
    let best = 0;
    days.forEach((k) => {
      const prev = new Date(k + "T12:00"); prev.setDate(prev.getDate() - 1);
      if (days.has(todayKey(prev))) return;
      let len = 0; const d = new Date(k + "T12:00");
      while (days.has(todayKey(d))) { len++; d.setDate(d.getDate() + 1); }
      best = Math.max(best, len);
    });
    return [b, best];
  });
  // 48 reading vs hearing ratio
  const ratio = dayList.slice(-42).map((d) => ({ label: d.k.slice(5), r: d.rm, h: d.hm }));
  // 49 revisits
  const secCount = {};
  sessions.forEach((s) => { if (s.section) secCount[s.section.trim()] = (secCount[s.section.trim()] || 0) + 1; });
  const revisits = Object.entries(secCount).filter(([, n]) => n > 1).sort((a, b) => b[1] - a[1]).slice(0, 8);
  return (
    <div style={{ display: "grid", gap: 14 }}>
      <Viz n={36} t="The bookshelf" note="Spine height = total time inside each book.">
        {books.length ? (
          <div style={{ display: "flex", alignItems: "flex-end", gap: 6, height: 150, borderBottom: `3px solid ${C.maroon}`, paddingBottom: 0 }}>
            {books.map(([b, s], i) => (
              <div key={b} title={`${b}: ${s.min} min`} style={{
                flex: 1, height: `${20 + (s.min / maxBook) * 80}%`, background: PAL[i % PAL.length],
                borderRadius: "4px 4px 0 0", display: "flex", alignItems: "flex-end", justifyContent: "center", overflow: "hidden",
              }}>
                <span style={{ writingMode: "vertical-rl", color: "#fff", fontSize: 10, padding: 4, whiteSpace: "nowrap" }}>{b}</span>
              </div>
            ))}
          </div>
        ) : <Empty />}
      </Viz>
      <Viz n={37} t="Srimad-Bhagavatam canto map" note='Cantos touched, parsed from sections like "SB 1.2.6".'>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(12,1fr)", gap: 4 }}>
          {Array.from({ length: 12 }, (_, i) => (
            <div key={i} style={{ textAlign: "center", padding: "10px 0", borderRadius: 6, background: sbCantos.has(i + 1) ? C.saffron : C.ivory, color: sbCantos.has(i + 1) ? "#fff" : C.faint, fontSize: 12, fontWeight: 600 }}>{i + 1}</div>
          ))}
        </div>
      </Viz>
      <Viz n={38} t="Bhagavad-gita chapter map" note='Chapters touched, parsed from "BG 2.13" style sections.'>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(9,1fr)", gap: 4 }}>
          {Array.from({ length: 18 }, (_, i) => (
            <div key={i} style={{ textAlign: "center", padding: "8px 0", borderRadius: 6, background: bgCh.has(i + 1) ? C.tulsi : C.ivory, color: bgCh.has(i + 1) ? "#fff" : C.faint, fontSize: 12, fontWeight: 600 }}>{i + 1}</div>
          ))}
        </div>
      </Viz>
      <Viz n={39} t="Reading velocity by book" note="Weekly minutes, stacked.">
        <BarsViz data={wkArea} x="w" stack bars={topB.map((b, i) => ({ k: b, c: PAL[i] }))} h={200} />
      </Viz>
      <Viz n={40} t="Bookmarks" note="Where you last left off in each book.">
        {books.length ? books.map(([b, s]) => (
          <div key={b} style={{ display: "flex", justifyContent: "space-between", fontSize: 13, padding: "6px 0", borderBottom: `1px solid ${C.line}` }}>
            <span><b>{b}</b>{s.lastSection && <span style={{ color: C.tulsi }}> · {s.lastSection}</span>}</span>
            <span style={{ color: C.faint }}>{fmtDate(s.last)}</span>
          </div>
        )) : <Empty />}
      </Viz>
      <Viz n={41} t="Book Gantt" note="Active span of each book, first session to latest.">
        <BookGantt books={books} />
      </Viz>
      <Viz n={42} t="When you read" note="Time-of-day of reading sessions (timestamped from now on). Dot size = minutes.">
        <PtsViz data={tod} xk="hour" yk="min" xdom={[4, 24]} xfmt={fmtH} zk="min" />
      </Viz>
      <Viz n={43} t="Deep reads vs nibbles" note="Distribution of session lengths.">
        <BarsViz data={hist} x="b" bars={[{ k: "n", name: "Sessions", c: C.gold }]} />
      </Viz>
      <Viz n={44} t="Book rotation" note="Distinct books per week. Focus or feast?">
        <BarsViz data={rot} x="w" bars={[{ k: "books", c: C.sky }]} />
      </Viz>
      <Viz n={45} t="SB / BG / CC balance" note="Monthly share of reading minutes (%).">
        <BarsViz data={split} x="m" stack bars={[{ k: "Srimad", c: PAL[0] }, { k: "Bhagavad", c: PAL[1] }, { k: "Caitanya", c: PAL[2] }, { k: "Other", c: "#D8D0C0" }]} h={200} dom={[0, 100]} />
      </Viz>
      <Viz n={46} t={`${year} report card`} note="Minutes per book this year, graded against a 1,500-min/year bar.">
        <Rank items={yearMin.map(([b, m]) => [b, m, `· ${grade(m / 1500)}`])} unit=" min" />
      </Viz>
      <Viz n={47} t="Daily-reading streaks" note="Longest unbroken run of consecutive days per book.">
        <Rank items={streakBook} unit=" days" />
      </Viz>
      <Viz n={48} t="Reading vs hearing" note="Sravanam and svadhyaya side by side.">
        <LinesViz data={ratio} x="label" lines={[{ k: "r", name: "Reading", c: C.saffron }, { k: "h", name: "Hearing", c: C.tulsi }]} />
      </Viz>
      <Viz n={49} t="Passages you return to" note="Sections logged more than once — the verses that keep calling.">
        {revisits.length ? <Rank items={revisits} unit="×" /> : <Empty m="No repeated sections yet." />}
      </Viz>
      <Viz n={50} t="Year strips per book" note="One 52-week strip per book; darker = more minutes that week.">
        <YearStrips sessions={sessions} books={books.slice(0, 5).map((b) => b[0])} />
      </Viz>
    </div>
  );
}
function BookGantt({ books }) {
  if (!books.length) return <Empty />;
  const all = books.map(([b, s]) => ({ b, f: new Date(s.first + "T12:00"), l: new Date(s.last + "T12:00") }));
  const min = Math.min(...all.map((a) => a.f)), max = Math.max(...all.map((a) => a.l), Date.now());
  const span = Math.max(max - min, 864e5);
  return all.map((a, i) => (
    <div key={a.b} style={{ margin: "8px 0" }}>
      <div style={{ fontSize: 11, color: C.faint }}>{a.b}</div>
      <div style={{ height: 10, background: C.ivory, borderRadius: 5, position: "relative" }}>
        <div style={{ position: "absolute", left: `${((a.f - min) / span) * 100}%`, width: `${Math.max(2, ((a.l - a.f) / span) * 100)}%`, top: 0, bottom: 0, background: PAL[i % PAL.length], borderRadius: 5 }} />
      </div>
    </div>
  ));
}
function YearStrips({ sessions, books }) {
  if (!books.length) return <Empty />;
  const year = new Date().getFullYear();
  return books.map((b, bi) => {
    const wkMin = Array(53).fill(0);
    sessions.filter((s) => s.book === b && s.k.startsWith(String(year))).forEach((s) => {
      const d = new Date(s.k + "T12:00");
      const wk = Math.floor((d - new Date(year, 0, 1)) / (7 * 864e5));
      wkMin[Math.min(52, wk)] += +s.minutes;
    });
    const mx = Math.max(...wkMin, 1);
    return (
      <div key={b} style={{ marginBottom: 8 }}>
        <div style={{ fontSize: 11, color: C.faint }}>{b}</div>
        <div style={{ display: "flex", gap: 1 }}>
          {wkMin.map((m, i) => <div key={i} style={{ flex: 1, height: 12, borderRadius: 2, background: m ? `rgba(92,26,36,${0.2 + 0.8 * (m / mx)})` : C.ivory }} />)}
        </div>
      </div>
    );
  });
}

function WorshipViz({ A }) {
  const { dayList } = A;
  const prog = ["mangala", "narasimha", "tulsi", "guruPuja"];
  const level = (e) => {
    const w = e.e.worship || {};
    if (prog.every((k) => w[k])) return 3;
    if (prog.some((k) => w[k])) return 2;
    if (w.simple) return 1;
    return 0;
  };
  // 51 ladder
  const ladder = dayList.slice(-30).map((d) => ({ label: d.k.slice(5), lvl: level(d) }));
  // 52 donut
  const lvls = [0, 0, 0, 0];
  dayList.forEach((d) => lvls[level(d)]++);
  // 53 cooked streak
  let cook = 0;
  { const d = new Date(); for (;;) { const e = A.byKey[todayKey(d)]; if (e && e.e.worship?.cooked) { cook++; d.setDate(d.getDate() - 1); } else if (todayKey(d) === todayKey()) { d.setDate(d.getDate() - 1); } else break; if (cook > 999) break; } }
  // 54 8x7 grid this week
  const wk7 = A.last7;
  // 55 before vs after work
  const logged = dayList.filter((d) => Object.keys(d.e.worship || {}).length);
  const rel = (k) => logged.length ? logged.filter((d) => d.e.worship?.[k]).length / logged.length : 0;
  // 56 trend with ekadasi
  const wTrend = dayList.slice(-45).map((d) => ({ label: d.k.slice(5), pct: Math.round(d.wp * 100), ek: d.ek ? Math.round(d.wp * 100) : null }));
  // 57 first dropped
  const dropped = {};
  logged.filter((d) => d.wp < 0.99).forEach((d) => WORSHIP_ITEMS.forEach(([k, l]) => { if (!d.e.worship?.[k] && k !== "simple") dropped[l] = (dropped[l] || 0) + 1; }));
  const dropRank = Object.entries(dropped).sort((a, b) => b[1] - a[1]);
  // 59 weekend vs weekday
  const wdW = mean(dayList.filter((d) => d.wd >= 1 && d.wd <= 5).map((d) => d.wp)) * 100;
  const weW = mean(dayList.filter((d) => d.wd === 0 || d.wd === 6).map((d) => d.wp)) * 100;
  // 60 cooked vs score
  const cookedS = mean(dayList.filter((d) => d.e.worship?.cooked).map((d) => d.score));
  const notS = mean(dayList.filter((d) => !d.e.worship?.cooked).map((d) => d.score));
  return (
    <div style={{ display: "grid", gap: 14 }}>
      <Viz n={51} t="Morning programme ladder" note="0 = none · 1 = simple offering · 2 = partial programme · 3 = full programme.">
        <BarsViz data={ladder} x="label" bars={[{ k: "lvl", name: "Level", c: C.maroon }]} dom={[0, 3]} />
      </Viz>
      <Viz n={52} t="Full programme vs simple offering" note="The shape of your mornings, all time.">
        <DonutViz data={[{ name: "Full programme", value: lvls[3] }, { name: "Partial", value: lvls[2] }, { name: "Simple offering", value: lvls[1] }, { name: "None logged", value: lvls[0] }]} />
      </Viz>
      <Viz n={53} t="Cooking streak" note="Consecutive days an offering was cooked for the deities.">
        <Stat v={`${cook} ${"🔥".repeat(Math.min(5, Math.ceil(cook / 7)))}`} l="days cooking for Krishna" />
      </Viz>
      <Viz n={54} t="This week's arati matrix" note="Every item × every day.">
        <MatrixGrid rows={WORSHIP_ITEMS.map(([, l]) => l)} cols={wk7.map((k) => fmtDate(k).slice(0, 3))} cell={18}
          value={(i, j) => A.byKey[wk7[j]]?.e.worship?.[WORSHIP_ITEMS[i][0]] ? 1 : 0}
          colorFn={(v) => v ? C.tulsi : "#EFE9DC"} fmt={(i, j, v) => `${WORSHIP_ITEMS[i][1]} — ${v ? "done" : "—"}`} />
      </Viz>
      <Viz n={55} t="Bookends of the working day" note="Reliability of the before-work and after-work arati.">
        <BarsViz data={[{ d: "Before work", v: Math.round(rel("morningArati") * 100) }, { d: "After work", v: Math.round(rel("eveningArati") * 100) }]} x="d" bars={[{ k: "v", name: "% of days", c: C.saffron }]} dom={[0, 100]} />
      </Viz>
      <Viz n={56} t="Worship score over time" note="Maroon dots mark Ekadasi days (from notes).">
        {wTrend.length ? (
          <ResponsiveContainer width="100%" height={180}>
            <ComposedChart data={wTrend} margin={{ top: 6, right: 6, left: -24, bottom: 0 }}>
              <CartesianGrid stroke={C.line} strokeDasharray="3 3" />
              <XAxis dataKey="label" tick={tickS} interval="preserveStartEnd" /><YAxis tick={tickS} domain={[0, 100]} /><Tooltip />
              <Line type="monotone" dataKey="pct" stroke={C.gold} strokeWidth={2} dot={false} name="Worship %" />
              <Scatter dataKey="ek" fill={C.maroon} name="Ekadasi" />
            </ComposedChart>
          </ResponsiveContainer>
        ) : <Empty />}
      </Viz>
      <Viz n={57} t="What gets dropped first" note="On incomplete days, which item is most often missing.">
        <Rank items={dropRank} unit=" days" />
      </Viz>
      <Viz n={58} t="Worship constellation" note="Each star is an item; brightness = how often it's offered.">
        <Constellation A={A} logged={logged} />
      </Viz>
      <Viz n={59} t="Weekday vs weekend worship" note="Does the altar feel the office too?">
        <BarsViz data={[{ d: "Weekday", v: Math.round(wdW || 0) }, { d: "Weekend", v: Math.round(weW || 0) }]} x="d" bars={[{ k: "v", name: "Avg worship %", c: C.tulsi }]} dom={[0, 100]} />
      </Viz>
      <Viz n={60} t="The cooking dividend" note="Average day score on days you cooked for the deities vs not.">
        <BarsViz data={[{ d: "Cooked", v: Math.round(cookedS || 0) }, { d: "Didn't", v: Math.round(notS || 0) }]} x="d" bars={[{ k: "v", name: "Avg score", c: C.maroon }]} dom={[0, 100]} />
      </Viz>
    </div>
  );
}
function Constellation({ logged }) {
  const pos = [[40, 36], [110, 22], [180, 40], [240, 28], [70, 90], [150, 80], [220, 92], [120, 130]];
  return (
    <svg viewBox="0 0 280 160" width="100%" style={{ background: "#221E38", borderRadius: 10 }}>
      {WORSHIP_ITEMS.map(([k, l], i) => {
        const freq = logged.length ? logged.filter((d) => d.e.worship?.[k]).length / logged.length : 0;
        return (
          <g key={k}>
            <circle cx={pos[i][0]} cy={pos[i][1]} r={3 + freq * 6} fill="#FFE9C2" opacity={0.25 + freq * 0.75} />
            <text x={pos[i][0]} y={pos[i][1] + 18} textAnchor="middle" style={{ font: "8px sans-serif", fill: "#B9B2D6" }}>{l.split(" ")[0]}</text>
          </g>
        );
      })}
    </svg>
  );
}

function VersesViz({ A, verses }) {
  const { dayList } = A;
  const mem = verses.filter((v) => v.status === "memorised");
  // 62 days to memorise
  const dtm = verses.filter((v) => v.addedAt && v.memorisedAt)
    .map((v) => [v.ref, Math.max(1, Math.round((new Date(v.memorisedAt) - new Date(v.addedAt)) / 864e5))]);
  // 63/64 recency
  const recency = verses.map((v) => {
    const lastPractice = [v.lastRevised, v.lastRecited].filter(Boolean).sort().slice(-1)[0];
    const days = lastPractice ? Math.round((new Date() - new Date(lastPractice + "T12:00")) / 864e5) : null;
    return { ...v, days, lastPractice };
  }).sort((a, b) => (b.days ?? 999) - (a.days ?? 999));
  // 65 cumulative memorised
  const memDates = mem.filter((v) => v.memorisedAt).map((v) => v.memorisedAt.slice(0, 10)).sort();
  let cum = 0;
  const cumMem = memDates.map((d) => ({ label: d.slice(5), n: ++cum }));
  // 66 source split
  const src = {};
  verses.forEach((v) => { const m = /^([A-Za-z]+)/.exec(v.ref || ""); const s = v.book || (m ? m[1].toUpperCase() : "?"); src[s] = (src[s] || 0) + 1; });
  // 67 sparklines from history
  // 69 revised vs avg round dur
  const rv = dayList.filter((d) => d.avgDur != null).map((d) => ({ x: versePracticed(d.e) ? 1 : 0, dur: d.avgDur }));
  const revD = rv.filter((r) => r.x).map((r) => r.dur), noRevD = rv.filter((r) => !r.x).map((r) => r.dur);
  return (
    <div style={{ display: "grid", gap: 14 }}>
      <Viz n={61} t="Verse practice funnel" note="Verses added → learning/reciting → memorised solidly.">
        <BarsViz data={[{ s: "Added", n: verses.length }, { s: "Learning", n: verses.length - mem.length }, { s: "Memorised", n: mem.length }]} x="s" bars={[{ k: "n", c: C.maroon }]} />
      </Viz>
      <Viz n={62} t="Days to memorise" note="From adding a verse to marking it memorised (tracked from now on).">
        {dtm.length ? <Rank items={dtm} unit=" days" /> : <Empty m="Will populate as you memorise newly-added verses." />}
      </Viz>
      <Viz n={63} t="Practice recency heat" note="Days since each verse was last revised or recited — pale is fresh, deep maroon is overdue.">
        <div style={{ display: "flex", flexWrap: "wrap", gap: 6 }}>
          {verses.length ? recency.map((v) => (
            <div key={v.ref} title={`${v.ref}: ${v.days == null ? "never practised" : v.days + "d ago"}`}
              style={{ padding: "6px 10px", borderRadius: 8, fontSize: 12, color: "#fff", background: v.days == null ? C.ink : v.days <= 1 ? C.tulsi : v.days <= 3 ? C.saffron : C.maroon }}>
              {v.ref}
            </div>
          )) : <Empty />}
        </div>
      </Viz>
      <Viz n={64} t="Forgetting-risk queue" note="Revise from the top down — spaced repetition by neglect.">
        <Rank items={recency.slice(0, 8).map((v) => [v.ref, v.days ?? 99, v.status === "memorised" ? "· memorised" : ""])} unit="d" />
      </Viz>
      <Viz n={65} t="Cumulative verses memorised" note="The staircase of śāstra in the heart.">
        {cumMem.length ? <LinesViz data={cumMem} x="label" lines={[{ k: "n", name: "Memorised", c: C.maroon }]} /> : <Stat v={mem.length} l="memorised so far" sub="Step chart builds as future verses get memorised-dates." />}
      </Viz>
      <Viz n={66} t="Where your verses come from" note="By scripture prefix in the reference.">
        <DonutViz data={Object.entries(src).map(([name, value]) => ({ name, value }))} />
      </Viz>
      <Viz n={67} t="Per-verse practice sparklines" note="Last 30 days of revision or recitation history per verse.">
        {verses.filter((v) => v.history?.length).length ? verses.filter((v) => v.history?.length).map((v) => (
          <div key={v.ref} style={{ display: "flex", alignItems: "center", gap: 8, margin: "5px 0" }}>
            <span style={{ fontSize: 12, width: 70 }}>{v.ref}</span>
            <div style={{ display: "flex", gap: 1 }}>
              {Array.from({ length: 30 }, (_, i) => {
                const d = new Date(); d.setDate(d.getDate() - 29 + i);
                return <div key={i} style={{ width: 6, height: 14, borderRadius: 1, background: ((v.history || []).includes(todayKey(d)) || (v.recitationHistory || []).includes(todayKey(d))) ? C.tulsi : C.ivory }} />;
              })}
            </div>
          </div>
        )) : <Empty m="Revise or recite verses on the Today tab and the history fills in." />}
      </Viz>
      <Viz n={68} t="Prayer calendar" note="Days prayers were recited.">
        <BinaryCalendar A={A} test={(d) => prayersDone(d.e)} />
      </Viz>
      <Viz n={69} t="Journaling calendar" note="Days you journaled.">
        <BinaryCalendar A={A} test={(d) => !!d.e.journaled} />
      </Viz>
      <Viz n={70} t="Do verses sharpen japa?" note="Average round length on verse-practice days vs non-practice days.">
        {revD.length && noRevD.length ? <BarsViz data={[{ d: "Practised", v: +mean(revD).toFixed(1) }, { d: "Didn't", v: +mean(noRevD).toFixed(1) }]} x="d" bars={[{ k: "v", name: "Avg min/round", c: C.sky }]} /> : <Empty />}
      </Viz>
      <Viz n={71} t="The śloka tree" note="One leaf per memorised verse; buds are still learning.">
        <SlokaTree verses={verses} />
      </Viz>
    </div>
  );
}
function SlokaTree({ verses }) {
  return (
    <svg viewBox="0 0 280 200" width="100%">
      <rect x="132" y="120" width="16" height="80" rx="4" fill="#7A5230" />
      <path d="M140 130 Q100 110 70 80 M140 125 Q180 105 212 78 M140 118 Q140 80 140 55" stroke="#7A5230" strokeWidth="6" fill="none" strokeLinecap="round" />
      {verses.map((v, i) => {
        const branch = i % 3;
        const t = 0.35 + (Math.floor(i / 3) % 4) * 0.18;
        const x = branch === 0 ? 140 - 70 * t : branch === 1 ? 140 + 72 * t : 140 + (i % 2 ? 10 : -10);
        const y = branch === 2 ? 118 - 63 * t : (branch === 0 ? 130 : 125) - 48 * t;
        const memd = v.status === "memorised";
        return (
          <g key={i}>
            <ellipse cx={x} cy={y} rx={memd ? 9 : 5} ry={memd ? 12 : 7} fill={memd ? C.tulsi : "#B9C9A8"} transform={`rotate(${(i * 37) % 60 - 30} ${x} ${y})`} />
            {memd && <text x={x} y={y + 24} textAnchor="middle" style={{ font: "7px sans-serif", fill: C.faint }}>{v.ref}</text>}
          </g>
        );
      })}
      {!verses.length && <text x="140" y="40" textAnchor="middle" style={{ font: "11px sans-serif", fill: C.faint }}>Add verses to grow the tree</text>}
    </svg>
  );
}

function SleepViz({ A }) {
  const { dayList } = A;
  const sl = dayList.filter((d) => d.sleepH != null || d.wake != null);
  // 71 bands
  const bands = sl.slice(-14);
  // 72 snooze tax
  const tax = dayList.filter((d) => d.wake != null && d.firstH != null)
    .map((d) => ({ label: d.k.slice(5), lag: Math.max(0, +((d.firstH - d.wake) * 60).toFixed(0)) }));
  // 73 sleep vs round dur
  const svd = dayList.filter((d) => d.sleepH != null && d.avgDur != null).map((d) => ({ sleep: d.sleepH, dur: d.avgDur }));
  // 74 brahma muhurta
  const wakes = dayList.filter((d) => d.wake != null);
  const bm = wakes.length ? wakes.filter((d) => d.wake <= 4.75).length / wakes.length : null;
  // 75 bedtime by weekday
  const bedWd = WD.map((w, i) => {
    const bs = dayList.filter((d) => d.wd === i && d.sleep != null).map((d) => d.sleep < 12 ? d.sleep + 24 : d.sleep);
    return { w, bed: bs.length ? +mean(bs).toFixed(2) : null };
  });
  // 76 regularity
  const wkReg = [];
  for (let i = 6; i < wakes.length; i++) wkReg.push({ label: wakes[i].k.slice(5), sd: +std(wakes.slice(i - 6, i + 1).map((d) => d.wake)).toFixed(2) });
  // 77 sleep vs score
  const svs = dayList.filter((d) => d.sleepH != null).map((d) => ({ sleep: d.sleepH, score: d.score }));
  // 78 quadrant
  const quad = dayList.filter((d) => d.wake != null).map((d) => ({ wake: d.wake, score: d.score }));
  const medWake = median(quad.map((q) => q.wake)), medScore = median(quad.map((q) => q.score));
  // 79 sleep debt weekly
  const wkOf = (k) => { const d = new Date(k + "T12:00"); d.setDate(d.getDate() - d.getDay()); return todayKey(d); };
  const debtWks = {};
  sl.forEach((d) => { if (d.sleepH != null) { const w = wkOf(d.k); debtWks[w] = (debtWks[w] || 0) + (7 - d.sleepH); } });
  const debt = Object.entries(debtWks).sort().slice(-10).map(([w, v]) => ({ w: w.slice(5), debt: +v.toFixed(1) }));
  // 81 night awakenings
  const awakeLoad = dayList.filter((d) => d.awakeCount || d.awakeMin).slice(-30).map((d) => ({
    label: d.k.slice(5),
    times: d.awakeCount,
    minutes: +(d.awakeMin || 0),
  }));
  return (
    <div style={{ display: "grid", gap: 14 }}>
      <Viz n={71} t="Sleep bands" note="Each column spans sleep → wake. Short, early, consistent columns are the win.">
        {bands.length ? (
          <div style={{ overflowX: "auto", paddingBottom: 4 }}>
            <svg viewBox={`0 0 ${Math.max(300, bands.length * 20 + 34)} 210`} width="100%" style={{ minWidth: Math.min(520, Math.max(320, bands.length * 20 + 34)) }}>
            {[20, 24, 28, 32].map((h) => <text key={h} x="2" y={(h - 19) * 14} style={{ font: "8px sans-serif", fill: C.faint }}>{h % 24}:00</text>)}
            {bands.map((d, i) => {
              if (d.sleep == null || d.wake == null) return null;
              const s = d.sleep < 12 ? d.sleep + 24 : d.sleep;
              const w = d.wake + 24;
              return <rect key={i} x={30 + i * 20} y={Math.max(6, (s - 19) * 14)} width="13" height={Math.max(4, Math.min(190, (w - s) * 14))} rx="4"
                fill={d.wake <= 4.75 ? C.tulsi : C.sky} opacity="0.85"><title>{fmtDate(d.k)}: {d.e.sleepTime}→{d.e.wakeTime}{d.awakeMin ? ` · awake ${d.awakeMin} min` : ""}</title></rect>;
            })}
          </svg>
          </div>
        ) : <Empty m="Log sleep and wake times on the Today tab." />}
      </Viz>
      <Viz n={72} t="The snooze tax" note="Minutes between waking and the first round starting.">
        <BarsViz data={tax.slice(-14)} x="label" bars={[{ k: "lag", name: "Wake → japa (min)", c: C.gold }]} h={190} />
      </Viz>
      <Viz n={73} t="Sleep vs japa quality" note="Hours slept vs next morning's avg round length. There is a sweet spot.">
        <PtsViz data={svd} xk="sleep" yk="dur" trend />
      </Viz>
      <Viz n={74} t="Brahma-muhurta rate" note="Share of logged days waking by 4:45am.">
        {bm == null ? <Empty /> : <Gauge pct={bm} label="days up for brahma-muhurta" />}
      </Viz>
      <Viz n={75} t="Bedtime by weekday" note="Friday drift is visible from orbit.">
        <BarsViz data={bedWd.filter((b) => b.bed != null)} x="w" bars={[{ k: "bed", name: "Avg bedtime", c: C.maroon }]} dom={[20, 30]} />
      </Viz>
      <Viz n={76} t="Wake regularity" note="7-day rolling std-dev of wake time. Under 0.5h is monastic.">
        <LinesViz data={wkReg.slice(-21)} x="label" lines={[{ k: "sd", name: "Std dev (h)", c: C.sky }]} h={190} />
      </Viz>
      <Viz n={77} t="Sleep vs day score" note="Total sadhana score against hours slept.">
        <PtsViz data={svs} xk="sleep" yk="score" trend />
      </Viz>
      <Viz n={78} t="Early-rising quadrant" note="Left of the line = earlier than your median wake; upper = above-median score. Live in the top-left.">
        {quad.length > 3 ? (
          <ResponsiveContainer width="100%" height={200}>
            <ScatterChart margin={{ top: 6, right: 6, left: -18, bottom: 0 }}>
              <CartesianGrid stroke={C.line} strokeDasharray="3 3" />
              <XAxis type="number" dataKey="wake" tick={tickS} tickFormatter={fmtH} domain={["auto", "auto"]} />
              <YAxis type="number" dataKey="score" tick={tickS} domain={[0, 100]} />
              <Tooltip formatter={(v, n) => n === "wake" ? fmtH(v) : v} />
              <ReferenceLine x={medWake} stroke={C.faint} strokeDasharray="4 4" />
              <ReferenceLine y={medScore} stroke={C.faint} strokeDasharray="4 4" />
              <Scatter data={quad} fill={C.maroon} fillOpacity={0.6} />
            </ScatterChart>
          </ResponsiveContainer>
        ) : <Empty />}
      </Viz>
      <Viz n={79} t="Weekly sleep debt" note="Hours short of 7h/night accumulated per week. Negative is surplus.">
        <BarsViz data={debt} x="w" bars={[{ k: "debt", name: "Debt (h)", c: C.sky }]} />
      </Viz>
      <Viz n={80} t="Wake-time ribbon" note="The year, one sliver per day, coloured by wake hour (maroon = pre-dawn).">
        <WakeRibbon A={A} />
      </Viz>
      <Viz n={81} t="Night awakenings" note="Times awake and total minutes awake during the night. Net sleep subtracts these minutes.">
        <BarsViz data={awakeLoad.slice(-14)} x="label" bars={[{ k: "minutes", name: "Minutes awake", c: C.sky }]} h={190} />
        <div style={{ fontSize: 12, color: C.faint, marginTop: 6 }}>{awakeLoad.slice(-14).map((d) => `${d.label}: ${d.times}×`).join(" · ")}</div>
      </Viz>
    </div>
  );
}
function WakeRibbon({ A }) {
  const year = new Date().getFullYear();
  const days = [];
  const d = new Date(year, 0, 1);
  while (d.getFullYear() === year && d <= new Date()) { days.push(todayKey(d)); d.setDate(d.getDate() + 1); }
  const color = (w) => w == null ? "#EFE9DC" : w <= 4.75 ? C.maroon : w <= 6 ? C.saffron : w <= 8 ? "#EBA85C" : C.saffronSoft;
  return (
    <div style={{ display: "flex", height: 38, borderRadius: 6, overflow: "hidden" }}>
      {days.map((k) => <div key={k} title={`${fmtDate(k)} — ${A.byKey[k]?.e.wakeTime || "—"}`} style={{ flex: 1, background: color(A.byKey[k]?.wake) }} />)}
    </div>
  );
}

function CompositeViz({ A, verses }) {
  const { dayList } = A;
  // limb pcts per day
  const limbs = dayList.map((d) => ({
    ...d,
    cP: Math.min(d.n / TARGET, 1), hP: Math.min(d.hm / 30, 1), rP: Math.min(d.rm / 30, 1), wP: d.wp,
    vP: versePracticed(d.e) ? 1 : 0, pP: prayersDone(d.e) ? 1 : 0,
  }));
  const LIMB = [["cP", "Chanting"], ["hP", "Hearing"], ["rP", "Reading"], ["wP", "Worship"], ["vP", "Verses"], ["pP", "Prayers"]];
  // 81 stacked minutes
  const stackMin = dayList.slice(-30).map((d) => ({
    label: d.k.slice(5),
    Chanting: +d.rounds.reduce((a, r) => a + r.dur, 0).toFixed(0),
    Hearing: d.hm, Reading: d.rm, "Worship (est)": Math.round(d.wp * 30),
  }));
  // 82 correlation matrix
  const corrM = LIMB.map(([a]) => LIMB.map(([b]) => {
    const pairs = limbs.filter((d) => true);
    return corr(pairs.map((d) => d[a]), pairs.map((d) => d[b]));
  }));
  // 83 domino
  const bad = limbs.filter((d) => d.n < TARGET), good = limbs.filter((d) => d.n >= TARGET);
  const domino = LIMB.slice(1).map(([k, name]) => ({
    limb: name, "16-round days": +(mean(good.map((d) => d[k])) * 100).toFixed(0), "Short days": +(mean(bad.map((d) => d[k])) * 100).toFixed(0),
  }));
  // 84 waterfall today
  const t = A.today;
  const parts = t ? [
    ["Chanting", Math.min(t.n / TARGET, 1) * 40], ["Hearing", Math.min(t.hm / 30, 1) * 15],
    ["Reading", Math.min(t.rm / 30, 1) * 20], ["Worship", t.wp * 15],
    ["Verses", versePracticed(t.e) ? 5 : 0], ["Prayers", prayersDone(t.e) ? 5 : 0],
  ] : [];
  let acc = 0;
  const wf = parts.map(([n, v]) => { const p = acc; acc += v; return { n, pad: +p.toFixed(0), v: +v.toFixed(1) }; });
  // 85 report card this month
  const mo = new Date().toISOString().slice(0, 7);
  const moDays = limbs.filter((d) => d.k.startsWith(mo));
  const card = LIMB.map(([k, name]) => { const p = mean(moDays.map((d) => d[k])); return [name, Math.round(p * 100), `· ${grade(p)}`]; });
  // 86 bollinger
  const boll = [];
  for (let i = 6; i < dayList.length; i++) {
    const win = dayList.slice(i - 6, i + 1).map((d) => d.score);
    boll.push({ label: dayList[i].k.slice(5), m: +mean(win).toFixed(0), hi: +(mean(win) + std(win)).toFixed(0), lo: +(mean(win) - std(win)).toFixed(0) });
  }
  // 87 best day
  const best = dayList.length ? dayList.reduce((a, b) => (b.score > a.score ? b : a)) : null;
  // 90 momentum
  const mom = [];
  for (let i = 30; i < dayList.length; i++) {
    const w1 = mean(dayList.slice(i - 29, i + 1).map((d) => d.score));
    const w0 = mean(dayList.slice(Math.max(0, i - 36), i - 6).map((d) => d.score));
    mom.push({ label: dayList[i].k.slice(5), m: +(w1 - w0).toFixed(1) });
  }
  // 91 MVS
  const mvs = limbs.length ? limbs.filter((d) => d.cP > 0 && d.hP > 0 && d.rP > 0 && d.wP > 0).length / limbs.length : null;
  // 92 consistency vs intensity per month
  const moList = [...new Set(dayList.map((d) => d.month))].sort();
  const ci = moList.map((m) => {
    const ds = limbs.filter((d) => d.month === m);
    return { m, logged: ds.length, avg: +mean(ds.map((d) => d.score)).toFixed(0) };
  });
  // 93 sharpe
  const s30 = dayList.slice(-30).map((d) => d.score);
  const sharpe = s30.length > 4 && std(s30) > 0 ? (mean(s30) / std(s30)).toFixed(2) : null;
  // 94 weekday fingerprint
  const wdGrid = LIMB.map(([k]) => WD.map((w, wd) => { const ds = limbs.filter((d) => d.wd === wd); return ds.length ? mean(ds.map((d) => d[k])) : null; }));
  // 95 festival timeline
  const fl = dayList.slice(-90).map((d) => ({ label: d.k.slice(5), score: d.score, mark: d.ek || d.fest ? d.score : null }));
  // 96 recovery
  const recoveries = [];
  for (let i = 1; i < dayList.length; i++) {
    if (dayList[i - 1].n >= TARGET && dayList[i].n < TARGET) {
      let j = i;
      while (j < dayList.length && dayList[j].n < TARGET) j++;
      if (j < dayList.length) recoveries.push(j - i);
    }
  }
  // 97 allocation
  const totalMin = limbs.reduce((a, d) => a + d.rounds.reduce((x, r) => x + r.dur, 0) + d.hm + d.rm, 0) || 1;
  const actual = {
    Chanting: limbs.reduce((a, d) => a + d.rounds.reduce((x, r) => x + r.dur, 0), 0) / totalMin * 100,
    Hearing: limbs.reduce((a, d) => a + d.hm, 0) / totalMin * 100,
    Reading: limbs.reduce((a, d) => a + d.rm, 0) / totalMin * 100,
  };
  const alloc = [
    { d: "Actual", Chanting: +actual.Chanting.toFixed(0), Hearing: +actual.Hearing.toFixed(0), Reading: +actual.Reading.toFixed(0) },
    { d: "Score weights", Chanting: 53, Hearing: 20, Reading: 27 },
  ];
  // 98 streak ring
  // 99 compound
  let cr = 0, ch = 0, crd = 0;
  const compound = dayList.map((d) => { cr += d.n; ch += d.hm / 60; crd += d.rm / 60; return { label: d.k.slice(5), rounds: cr, "hours heard": +ch.toFixed(1), "hours read": +crd.toFixed(1) }; });
  return (
    <div style={{ display: "grid", gap: 14 }}>
      <Viz n={81} t="Where the minutes go" note="Daily minutes per limb (worship estimated at 30 min when full).">
        <BarsViz data={stackMin} x="label" stack bars={[{ k: "Chanting", c: PAL[0] }, { k: "Hearing", c: PAL[2] }, { k: "Reading", c: PAL[1] }, { k: "Worship (est)", c: PAL[3] }]} h={210} />
      </Viz>
      <Viz n={82} t="Limb correlation matrix" note="Green = practices that rise together; red = trade-offs. Diagonal is always 1.">
        <MatrixGrid rows={LIMB.map(([, n]) => n)} cols={LIMB.map(([, n]) => n.slice(0, 3))} cell={26}
          value={(i, j) => corrM[i][j]}
          colorFn={(v) => v == null ? "#EFE9DC" : v > 0 ? `rgba(79,107,60,${Math.abs(v)})` : `rgba(92,26,36,${Math.abs(v)})`}
          fmt={(i, j, v) => `${LIMB[i][1]} × ${LIMB[j][1]}: ${v == null ? "—" : v.toFixed(2)}`} />
      </Viz>
      <Viz n={83} t="The domino effect" note="When chanting falls short of 16, how far do the other limbs fall with it?">
        <BarsViz data={domino} x="limb" bars={[{ k: "16-round days", c: C.tulsi }, { k: "Short days", c: C.maroon }]} h={200} dom={[0, 100]} />
      </Viz>
      <Viz n={84} t="Today's score, decomposed" note="Each limb's contribution stacking to the day's total.">
        {wf.length ? (
          <ResponsiveContainer width="100%" height={180}>
            <BarChart data={wf} margin={{ top: 6, right: 6, left: -24, bottom: 0 }}>
              <XAxis dataKey="n" tick={tickS} /><YAxis tick={tickS} domain={[0, 100]} /><Tooltip />
              <Bar dataKey="pad" stackId="w" fill="transparent" />
              <Bar dataKey="v" stackId="w" fill={C.saffron} radius={[3, 3, 0, 0]} />
            </BarChart>
          </ResponsiveContainer>
        ) : <Empty />}
      </Viz>
      <Viz n={85} t="This month's report card" note="Average completion per limb, graded.">
        <Rank items={card} unit="%" />
      </Viz>
      <Viz n={86} t="Score with volatility bands" note="7-day mean ± 1σ. A narrow channel climbing slowly beats a wild zig-zag.">
        {boll.length ? (
          <ResponsiveContainer width="100%" height={200}>
            <ComposedChart data={boll.slice(-60)} margin={{ top: 6, right: 6, left: -24, bottom: 0 }}>
              <CartesianGrid stroke={C.line} strokeDasharray="3 3" />
              <XAxis dataKey="label" tick={tickS} interval="preserveStartEnd" /><YAxis tick={tickS} domain={[0, 100]} /><Tooltip />
              <Area type="monotone" dataKey="hi" stroke="none" fill={C.saffronSoft} />
              <Area type="monotone" dataKey="lo" stroke="none" fill="#fff" />
              <Line type="monotone" dataKey="m" stroke={C.maroon} strokeWidth={2} dot={false} />
            </ComposedChart>
          </ResponsiveContainer>
        ) : <Empty />}
      </Viz>
      <Viz n={87} t="Your best day, replayed" note="The highest-scoring day on record and what made it.">
        {best ? (
          <div>
            <Stat v={`${best.score}/100`} l={fmtDate(best.k)} sub={`${best.n} rounds · ${best.hm}m heard · ${best.rm}m read · worship ${Math.round(best.wp * 100)}%`} />
            <div style={{ position: "relative", height: 22, background: C.ivory, borderRadius: 6, marginTop: 10, overflow: "hidden" }}>
              {best.rounds.map((r, i) => {
                const toPct = (h) => Math.max(0, Math.min(100, ((h - 4) / 19) * 100));
                return <div key={i} style={{ position: "absolute", left: `${toPct(r.hour)}%`, width: `${Math.max(0.8, (r.dur / 60 / 19) * 100)}%`, top: 3, bottom: 3, background: C.saffron, borderRadius: 3 }} />;
              })}
            </div>
            <div style={{ display: "flex", justifyContent: "space-between", fontSize: 10, color: C.faint }}><span>4am</span><span>11pm</span></div>
          </div>
        ) : <Empty />}
      </Viz>
      <Viz n={88} t="Year in pixels" note="Every day of the year, coloured by score.">
        <YearPixels A={A} />
      </Viz>
      <Viz n={89} t="Three-month radar overlay" note="Limb balance per month — watch the shape grow outward.">
        <MonthRadars limbs={limbs} LIMB={LIMB} />
      </Viz>
      <Viz n={90} t="Momentum" note="30-day average score vs the prior 30 days. Above zero = compounding.">
        <LinesViz data={mom.slice(-60)} x="label" lines={[{ k: "m", name: "Δ score", c: C.tulsi }]} refY={0} />
      </Viz>
      <Viz n={91} t="Minimum-viable-sadhana rate" note="Days where all four limbs got at least something.">
        {mvs == null ? <Empty /> : <Gauge pct={mvs} label="days touching all four limbs" />}
      </Viz>
      <Viz n={92} t="Consistency vs intensity" note="Each dot is a month: days logged (x) vs average score (y). Top-right is the saint's corner.">
        <PtsViz data={ci.map((c) => ({ logged: c.logged, avg: c.avg }))} xk="logged" yk="avg" ydom={[0, 100]} />
      </Viz>
      <Viz n={93} t="Sadhana Sharpe ratio" note="Mean ÷ volatility of the last 30 day-scores. You of all people know: risk-adjusted returns.">
        {sharpe ? <Stat v={sharpe} l="30-day Sharpe" sub={`mean ${Math.round(mean(s30))} · σ ${std(s30).toFixed(1)}`} /> : <Empty />}
      </Viz>
      <Viz n={94} t="Weekday fingerprint" note="Average completion of each limb by day of week.">
        <MatrixGrid rows={LIMB.map(([, n]) => n)} cols={WD} cell={22}
          value={(i, j) => wdGrid[i][j]}
          colorFn={(v) => v == null ? "#EFE9DC" : `rgba(224,123,31,${0.12 + 0.88 * v})`}
          fmt={(i, j, v) => `${LIMB[i][1]} · ${WD[j]}: ${v == null ? "—" : Math.round(v * 100) + "%"}`} />
      </Viz>
      <Viz n={95} t="Score timeline with holy days" note='Maroon dots = days whose note mentions Ekadasi or a festival.'>
        {fl.length ? (
          <ResponsiveContainer width="100%" height={190}>
            <ComposedChart data={fl} margin={{ top: 6, right: 6, left: -24, bottom: 0 }}>
              <CartesianGrid stroke={C.line} strokeDasharray="3 3" />
              <XAxis dataKey="label" tick={tickS} interval="preserveStartEnd" /><YAxis tick={tickS} domain={[0, 100]} /><Tooltip />
              <Line type="monotone" dataKey="score" stroke={C.gold} strokeWidth={2} dot={false} />
              <Scatter dataKey="mark" fill={C.maroon} />
            </ComposedChart>
          </ResponsiveContainer>
        ) : <Empty />}
      </Viz>
      <Viz n={96} t="Recovery time" note="After a broken 16-round day, how many days to get back. Faster recovery > never falling.">
        {recoveries.length ? <Stat v={`${mean(recoveries).toFixed(1)} days`} l="avg recovery" sub={`fastest ${Math.min(...recoveries)} · slowest ${Math.max(...recoveries)} · ${recoveries.length} recoveries`} /> : <Empty m="No broken-and-recovered streaks yet. May it stay that way." />}
      </Viz>
      <Viz n={97} t="Time allocation vs design" note="How your actual minutes split across the three timed limbs, against the scoring weights you chose.">
        <BarsViz data={alloc} x="d" stack bars={[{ k: "Chanting", c: PAL[0] }, { k: "Hearing", c: PAL[2] }, { k: "Reading", c: PAL[1] }]} dom={[0, 100]} />
      </Viz>
      <Viz n={98} t="100-day vow" note="Current 16-round streak against a 100-day target.">
        <ProgressRing pct={Math.min(1, A.streak / 100)} label={`${A.streak} / 100 days`} />
      </Viz>
      <Viz n={99} t="Compound devotion" note="Lifetime cumulative rounds, hours heard, hours read. The only chart that never goes down.">
        <LinesViz data={compound.slice(-180)} x="label" lines={[{ k: "rounds", c: C.maroon }, { k: "hours heard", c: C.tulsi }, { k: "hours read", c: C.saffron }]} h={210} />
      </Viz>
      <Viz n={100} t="The year mandala" note="365 petals around the centre, coloured by each day's score.">
        <Mandala A={A} />
      </Viz>
    </div>
  );
}
function YearPixels({ A }) {
  const year = new Date().getFullYear();
  const days = [];
  const d = new Date(year, 0, 1);
  while (d.getFullYear() === year) { days.push(todayKey(d)); d.setDate(d.getDate() + 1); }
  return (
    <div style={{ display: "grid", gridTemplateColumns: "repeat(31, 1fr)", gap: 2 }}>
      {days.map((k) => {
        const future = k > todayKey();
        return <div key={k} title={`${fmtDate(k)} — ${A.byKey[k]?.score ?? 0}`} style={{ aspectRatio: "1", borderRadius: 2, background: future ? "#F5F0E4" : scoreColor(A.byKey[k]?.score ?? 0) }} />;
      })}
    </div>
  );
}
function MonthRadars({ limbs, LIMB }) {
  const mos = [...new Set(limbs.map((d) => d.month))].sort().slice(-3);
  if (!mos.length) return <Empty />;
  const data = LIMB.map(([k, name]) => {
    const row = { axis: name };
    mos.forEach((m) => { const ds = limbs.filter((d) => d.month === m); row[m] = ds.length ? Math.round(mean(ds.map((d) => d[k])) * 100) : 0; });
    return row;
  });
  return (
    <ResponsiveContainer width="100%" height={260}>
      <RadarChart data={data} outerRadius="72%">
        <PolarGrid stroke={C.line} />
        <PolarAngleAxis dataKey="axis" tick={{ fontSize: 11, fill: C.ink }} />
        <Legend wrapperStyle={{ fontSize: 11 }} />
        {mos.map((m, i) => <Radar key={m} dataKey={m} stroke={PAL[i]} fill={PAL[i]} fillOpacity={0.18} />)}
      </RadarChart>
    </ResponsiveContainer>
  );
}
function ProgressRing({ pct, label }) {
  const r = 70, circ = 2 * Math.PI * r;
  return (
    <div style={{ textAlign: "center" }}>
      <svg viewBox="0 0 180 180" width="170">
        <circle cx="90" cy="90" r={r} fill="none" stroke={C.ivory} strokeWidth="14" />
        <circle cx="90" cy="90" r={r} fill="none" stroke={C.saffron} strokeWidth="14" strokeLinecap="round"
          strokeDasharray={`${pct * circ} ${circ}`} transform="rotate(-90 90 90)" />
        <text x="90" y="96" textAnchor="middle" style={{ font: "700 20px Georgia", fill: C.maroon }}>{Math.round(pct * 100)}%</text>
      </svg>
      <div style={{ fontSize: 12, color: C.faint }}>{label}</div>
    </div>
  );
}
function Mandala({ A }) {
  const year = new Date().getFullYear();
  const days = [];
  const d = new Date(year, 0, 1);
  while (d.getFullYear() === year) { days.push(todayKey(d)); d.setDate(d.getDate() + 1); }
  return (
    <svg viewBox="0 0 320 320" width="300" style={{ display: "block", margin: "0 auto" }}>
      <circle cx="160" cy="160" r="26" fill={C.maroon} />
      <text x="160" y="165" textAnchor="middle" style={{ font: "700 13px Georgia", fill: "#FFE9C2" }}>{year}</text>
      {days.map((k, i) => {
        const a = (i / days.length) * 2 * Math.PI - Math.PI / 2;
        const future = k > todayKey();
        const s = A.byKey[k]?.score ?? 0;
        const len = 36 + (future ? 6 : (s / 100) * 88);
        return <line key={k} x1={160 + 34 * Math.cos(a)} y1={160 + 34 * Math.sin(a)}
          x2={160 + len * Math.cos(a)} y2={160 + len * Math.sin(a)}
          stroke={future ? "#F0EADC" : scoreColor(s)} strokeWidth="2.4" strokeLinecap="round">
          <title>{fmtDate(k)} — {s}</title></line>;
      })}
    </svg>
  );
}


// ============ Excel export helpers ============
const addDays = (d, n) => {
  const x = new Date(d);
  x.setDate(x.getDate() + n);
  return x;
};
const startOfWeek = (d = new Date()) => {
  const x = new Date(d);
  x.setHours(0, 0, 0, 0);
  const day = x.getDay(); // Sunday = 0
  const diff = day === 0 ? -6 : 1 - day; // Monday-start week
  x.setDate(x.getDate() + diff);
  return x;
};
const endOfDay = (d) => {
  const x = new Date(d);
  x.setHours(23, 59, 59, 999);
  return x;
};
const round1 = (x) => x == null || x === "" ? "" : Math.round(Number(x) * 10) / 10;
const round2 = (x) => x == null || x === "" ? "" : Math.round(Number(x) * 100) / 100;
const yesNo = (v) => v ? "Yes" : "No";
const versePracticed = (e) => !!(e?.versesRevised || e?.versesRecited);
const safeSheet = (name) => String(name).replace(/[\\/?*\[\]:]/g, " ").slice(0, 31);
const sheetFromRows = (rows, emptyMessage = "No data for this period") =>
  XLSX.utils.json_to_sheet(rows.length ? rows : [{ Message: emptyMessage }]);
const appendSheet = (wb, rows, name, emptyMessage) => {
  XLSX.utils.book_append_sheet(wb, sheetFromRows(rows, emptyMessage), safeSheet(name));
};
const periodWindow = (period) => {
  const now = new Date();
  if (period === "week") {
    const start = startOfWeek(now);
    const end = endOfDay(addDays(start, 6));
    return { start, end, label: `${todayKey(start)}_to_${todayKey(end)}`, title: "Current week" };
  }
  if (period === "month") {
    const start = new Date(now.getFullYear(), now.getMonth(), 1);
    const end = endOfDay(new Date(now.getFullYear(), now.getMonth() + 1, 0));
    return { start, end, label: `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, "0")}`, title: "Current month" };
  }
  if (period === "year") {
    const start = new Date(now.getFullYear(), 0, 1);
    const end = endOfDay(new Date(now.getFullYear(), 11, 31));
    return { start, end, label: `${now.getFullYear()}`, title: "Current year" };
  }
  return { start: null, end: null, label: "all-data", title: "All data" };
};

const getStoredPayload = async (key) => {
  try {
    if (window.storage?.get) return await window.storage.get(key);
  } catch (e) {}
  try {
    const value = window.localStorage?.getItem(key);
    return value ? { value } : null;
  } catch (e) {}
  return null;
};
const setStoredPayload = async (key, value) => {
  try {
    if (window.storage?.set) return await window.storage.set(key, value);
  } catch (e) {}
  if (window.localStorage) {
    window.localStorage.setItem(key, value);
    return true;
  }
  throw new Error("No browser storage available");
};

const DEFAULT_LABELS = { speakers: ["Srila Prabhupada"], books: [...BOOKS], prayers: [], deities: [] };
const dedupOrdered = (...lists) => {
  const seen = new Set(); const out = [];
  lists.flat().forEach((v) => { if (v && !seen.has(v)) { seen.add(v); out.push(v); } });
  return out;
};
const normaliseLabels = (labels) => ({
  speakers: dedupOrdered(DEFAULT_LABELS.speakers, Array.isArray(labels?.speakers) ? labels.speakers : []),
  books: dedupOrdered(DEFAULT_LABELS.books, Array.isArray(labels?.books) ? labels.books : []),
  prayers: dedupOrdered(Array.isArray(labels?.prayers) ? labels.prayers : []),
  deities: dedupOrdered(Array.isArray(labels?.deities) ? labels.deities : []),
});
const normalisePayload = (payload) => ({
  days: payload?.days && typeof payload.days === "object" ? payload.days : {},
  verses: Array.isArray(payload?.verses) ? payload.verses : [],
  labels: normaliseLabels(payload?.labels),
});
const mergeUnique = (...lists) => [...new Set(lists.flat().filter(Boolean))].sort();
const verseKey = (v, i) => (v?.ref || v?.text || v?.id || `verse-${i}`).trim().toLowerCase();
const mergePayloads = (cloudPayload, localPayload) => {
  const cloud = normalisePayload(cloudPayload);
  const local = normalisePayload(localPayload);
  const days = { ...cloud.days, ...local.days }; // local browser wins if the same date exists in both
  const verseMap = new Map();
  [...cloud.verses, ...local.verses].forEach((v, i) => {
    if (!v) return;
    const key = verseKey(v, i);
    const prev = verseMap.get(key) || {};
    verseMap.set(key, {
      ...prev, ...v,
      history: mergeUnique(prev.history || [], v.history || []),
      recitationHistory: mergeUnique(prev.recitationHistory || [], v.recitationHistory || []),
    });
  });
  const labels = {
    speakers: dedupOrdered(cloud.labels.speakers, local.labels.speakers),
    books: dedupOrdered(cloud.labels.books, local.labels.books),
    prayers: dedupOrdered(cloud.labels.prayers, local.labels.prayers),
    deities: dedupOrdered(cloud.labels.deities, local.labels.deities),
  };
  return { days, verses: [...verseMap.values()], labels };
};
const payloadHasContent = (payload) => {
  const p = normalisePayload(payload);
  return Object.keys(p.days).length > 0 || p.verses.length > 0;
};
const fetchCloudState = async (userId) => {
  if (!supabase || !userId) return null;
  const { data, error } = await supabase
    .from("sadhana_app_state")
    .select("data")
    .eq("user_id", userId)
    .maybeSingle();
  if (error) throw error;
  return data?.data ? normalisePayload(data.data) : null;
};
const saveCloudState = async (userId, payload) => {
  if (!supabase || !userId) return false;
  const { error } = await supabase
    .from("sadhana_app_state")
    .upsert({ user_id: userId, data: normalisePayload(payload) }, { onConflict: "user_id" });
  if (error) throw error;
  return true;
};

// ============ MAIN APP ============
export default function App() {
  const [data, setData] = useState({});
  const [verses, setVerses] = useState([]);
  const [labels, setLabels] = useState(DEFAULT_LABELS);
  const [loaded, setLoaded] = useState(false);
  const [tab, setTab] = useState("today");
  const [vizTab, setVizTab] = useState("jdur");
  const [save, setSave] = useState("idle");
  const [session, setSession] = useState(null);
  const [sync, setSync] = useState("local");
  const [cloudLoadedFor, setCloudLoadedFor] = useState("");
  const [authEmail, setAuthEmail] = useState("");
  const [authPassword, setAuthPassword] = useState("");
  const [authMode, setAuthMode] = useState("signin");
  const [authMsg, setAuthMsg] = useState("");
  const [showSignup, setShowSignup] = useState(false);
  const [signupEmail, setSignupEmail] = useState("");
  const [signupPassword, setSignupPassword] = useState("");
  const [signupConfirm, setSignupConfirm] = useState("");
  const [signupMsg, setSignupMsg] = useState("");
  const [adminUsers, setAdminUsers] = useState(null);
  const [adminMsg, setAdminMsg] = useState("");
  const [adminSelected, setAdminSelected] = useState(null);
  const [adminSelectedData, setAdminSelectedData] = useState(null);
  const [adminNewPassword, setAdminNewPassword] = useState("");
  const [backupMsg, setBackupMsg] = useState("");
  const [importMode, setImportMode] = useState("merge");
  const fileInputRef = useRef(null);
  const [, tick] = useState(0);
  const tk = todayKey();
  const day = data[tk] || emptyDay();
  const user = session?.user || null;
  const cloudEnabled = !!supabase;

  useEffect(() => {
    (async () => {
      try {
        // migrate from v2 key if present
        let raw = null;
        raw = await getStoredPayload("sadhana-v3");
        if (!raw?.value) raw = await getStoredPayload("sadhana-v2");
        const norm = normalisePayload(raw?.value ? JSON.parse(raw.value) : null);
        setData(norm.days);
        setVerses(norm.verses);
        setLabels(norm.labels);
      } catch (e) {}
      setLoaded(true);
    })();
  }, []);

  useEffect(() => {
    if (!supabase) return;
    let alive = true;
    supabase.auth.getSession().then(({ data }) => {
      if (alive) setSession(data.session || null);
    });
    const { data: sub } = supabase.auth.onAuthStateChange((_event, nextSession) => {
      setSession(nextSession);
      setCloudLoadedFor("");
    });
    return () => {
      alive = false;
      sub?.subscription?.unsubscribe?.();
    };
  }, []);

  useEffect(() => {
    if (!loaded || !user || cloudLoadedFor === user.id) return;
    (async () => {
      try {
        setSync("loading");
        const localPayload = { days: data, verses, labels };
        const cloudPayload = await fetchCloudState(user.id);
        const merged = cloudPayload ? mergePayloads(cloudPayload, localPayload) : normalisePayload(localPayload);
        if (cloudPayload || payloadHasContent(localPayload)) {
          setData(merged.days);
          setVerses(merged.verses);
          setLabels(merged.labels);
          await setStoredPayload("sadhana-v3", JSON.stringify(merged));
          await saveCloudState(user.id, merged);
        }
        setCloudLoadedFor(user.id);
        setSync("synced");
      } catch (e) {
        console.error(e);
        setSync("error");
      }
    })();
  }, [loaded, user?.id, cloudLoadedFor]);

  useEffect(() => {
    if (!day.activeStart) return;
    const id = setInterval(() => tick((n) => n + 1), 1000);
    return () => clearInterval(id);
  }, [day.activeStart]);

  const persist = async (days, vs = verses, lbls = labels) => {
    const payload = { days, verses: vs, labels: lbls };
    setData(days); setVerses(vs); setLabels(lbls); setSave("saving");
    try {
      await setStoredPayload("sadhana-v3", JSON.stringify(payload));
      if (user) {
        setSync("saving");
        await saveCloudState(user.id, payload);
        setSync("synced");
      }
      setSave("saved"); setTimeout(() => setSave("idle"), 1200);
    } catch (e) { console.error(e); setSave("error"); setSync(user ? "error" : "local"); }
  };
  const update = (patch) => persist({ ...data, [tk]: { ...day, ...patch } });
  const addLabel = (kind, value) => {
    const v = (value || "").trim();
    if (!v || (labels[kind] || []).includes(v)) return;
    persist(data, verses, { ...labels, [kind]: [...(labels[kind] || []), v] });
  };
  const setAwakeCount = (raw) => {
    const n = Math.max(0, Math.min(12, Number(raw) || 0));
    const current = awakeRows(day);
    const next = Array.from({ length: n }, (_, i) => current[i] || { time: "", minutes: "" });
    update({ awakePeriods: next });
  };
  const updateAwakePeriod = (i, patch) => {
    const next = awakeRows(day).map((p, j) => j === i ? { ...p, ...patch } : p);
    update({ awakePeriods: next });
  };

  // japa
  const startRound = () => update({ activeStart: new Date().toISOString() });
  const finishRound = () => {
    const rounds = [...day.rounds, { start: day.activeStart, end: new Date().toISOString() }];
    update({ rounds, activeStart: autochant ? new Date().toISOString() : null });
  };
  const cancelRound = () => update({ activeStart: null });
  const removeRound = (i) => update({ rounds: day.rounds.filter((_, j) => j !== i) });
  const logManualRound = (m, endTime) => {
    const end = new Date(tsForTime(endTime)), start = new Date(end - m * 60000);
    update({ rounds: [...day.rounds, { start: start.toISOString(), end: end.toISOString() }] });
  };

  const [hForm, setHForm] = useState({ speaker: "", minutes: "", time: "" });
  const [rForm, setRForm] = useState({ book: BOOKS[0], section: "", minutes: "", pages: "", time: "" });
  const [pForm, setPForm] = useState({ label: "", time: "" });
  const [dForm, setDForm] = useState({ label: "", time: "" });
  const [vForm, setVForm] = useState({ book: BOOKS[0], ref: "", text: "" });
  const [manageVerses, setManageVerses] = useState(false);
  const [historyDate, setHistoryDate] = useState("");
  const [manualMin, setManualMin] = useState("");
  const [manualEnd, setManualEnd] = useState("");
  const [autochant, setAutochant] = useState(() => {
    try { return window.localStorage.getItem("sadhana-autochant") === "1"; } catch { return false; }
  });
  const toggleAutochant = (v) => {
    setAutochant(v);
    try { window.localStorage.setItem("sadhana-autochant", v ? "1" : "0"); } catch { /* localStorage unavailable */ }
  };
  const [exportOptions, setExportOptions] = useState({
    dataSummary: true, dataDaily: true, dataJapa: true, dataHearing: true, dataReading: true,
    dataWorship: true, dataSleep: true, dataVerses: true, dataVersePractice: true,
    vizOverview: true, vizJapa: false, vizHearingReading: false, vizSleep: true, vizWorship: false, vizComposite: false,
  });
  const setExportOption = (key, value) => setExportOptions((o) => ({ ...o, [key]: value }));

  // ============ analytics ============
  const A = useMemo(() => {
    const keys = Object.keys(data).sort();
    const dayList = keys.map((k) => {
      const e = data[k];
      const rounds = (e.rounds || []).map((r, i) => ({
        ...r, i, dur: mins(r),
        hour: r.start ? new Date(r.start).getHours() + new Date(r.start).getMinutes() / 60 : null,
      })).filter((r) => r.dur != null && r.dur >= 0);
      const date = new Date(k + "T12:00:00");
      const wake = hmToH(e.wakeTime), sleep = hmToH(e.sleepTime);
      const nightAwake = awakeRows(e);
      const awakeTotalMin = awakeMin(e);
      let grossSleepH = null, sleepH = null;
      if (wake != null && sleep != null) {
        let d = wake - sleep;
        if (d <= 0) d += 24;
        if (d <= 14) { grossSleepH = d; sleepH = Math.max(0, d - awakeTotalMin / 60); }
      }
      return {
        k, e, date, wd: date.getDay(), month: k.slice(0, 7), rounds, n: rounds.length,
        score: dayScore(e), hm: hearingMin(e), rm: readingMin(e), wp: worshipPct(e),
        firstH: rounds.length ? Math.min(...rounds.map((r) => r.hour).filter((h) => h != null)) : null,
        lastEndH: rounds.length ? Math.max(...rounds.map((r) => { const d2 = new Date(r.end); return d2.getHours() + d2.getMinutes() / 60; })) : null,
        avgDur: rounds.length ? mean(rounds.map((r) => r.dur)) : null,
        wake, sleep, sleepH, grossSleepH, awakePeriods: nightAwake, awakeCount: awakeCount(e), awakeMin: awakeTotalMin,
        ek: /ekadasi|ekadashi/i.test(e.note || ""),
        fest: /festival|janmastami|janmashtami|gaura|rama navami|nrsimha|narasimha caturdasi|radhastami/i.test(e.note || ""),
      };
    });
    const byKey = Object.fromEntries(dayList.map((d) => [d.k, d]));
    const allRounds = dayList.flatMap((d) => d.rounds.map((r) => ({ ...r, k: d.k, wd: d.wd, date: d.date })));
    const last7 = [];
    { const d = new Date(); for (let i = 6; i >= 0; i--) { const dd = new Date(d); dd.setDate(d.getDate() - i); last7.push(todayKey(dd)); } }
    let streak = 0;
    { const d = new Date(); for (;;) { const k = todayKey(d), e = byKey[k]; if (e && e.n >= TARGET) { streak++; d.setDate(d.getDate() - 1); } else if (k === tk) { d.setDate(d.getDate() - 1); } else break; if (streak > 3650) break; } }
    return { keys, dayList, byKey, allRounds, last7, streak, today: byKey[tk] };
  }, [data, tk]);

  const roundDurations = day.rounds.map(mins).filter((m) => m != null);
  const avgRound = roundDurations.length ? (roundDurations.reduce((a, b) => a + b, 0) / roundDurations.length).toFixed(1) : null;

  const trend = useMemo(() => {
    const out = []; const d = new Date(); d.setDate(d.getDate() - 29);
    for (let i = 0; i < 30; i++) {
      const k = todayKey(d), e = A.byKey[k];
      out.push({
        label: `${d.getDate()}/${d.getMonth() + 1}`,
        rounds: e?.n || 0, avgRound: e?.avgDur ? +e.avgDur.toFixed(1) : null,
        reading: e?.rm || 0, hearing: e?.hm || 0, score: e?.score || 0,
      });
      d.setDate(d.getDate() + 1);
    }
    return out;
  }, [A]);

  const weekStats = useMemo(() => {
    const es = A.last7.map((k) => A.byKey[k]).filter(Boolean);
    return {
      logged: es.length, full: es.filter((e) => e.n >= TARGET).length,
      hearing: es.reduce((a, e) => a + e.hm, 0), reading: es.reduce((a, e) => a + e.rm, 0),
      cooked: es.filter((e) => e.e.worship?.cooked).length,
      avgScore: es.length ? Math.round(mean(es.map((e) => e.score))) : 0,
    };
  }, [A]);
  const radar = useMemo(() => {
    const es = A.last7.map((k) => A.byKey[k]).filter(Boolean);
    const n = es.length || 1;
    const pct = (f) => Math.round(es.reduce((a, e) => a + f(e), 0) / n * 100);
    return [
      { axis: "Chanting", v: pct((e) => Math.min(e.n / TARGET, 1)) },
      { axis: "Hearing", v: pct((e) => Math.min(e.hm / 30, 1)) },
      { axis: "Reading", v: pct((e) => Math.min(e.rm / 30, 1)) },
      { axis: "Worship", v: pct((e) => e.wp) },
      { axis: "Verses", v: pct((e) => (versePracticed(e.e) ? 1 : 0)) },
      { axis: "Prayers", v: pct((e) => (prayersDone(e.e) ? 1 : 0)) },
    ];
  }, [A]);
  const speakerStats = useMemo(() => {
    const m = {};
    A.last7.forEach((k) => (A.byKey[k]?.e.hearing || []).forEach((h) => { m[h.speaker] = (m[h.speaker] || 0) + (+h.minutes || 0); }));
    return Object.entries(m).sort((a, b) => b[1] - a[1]);
  }, [A]);

  const signInOrUp = async (mode = authMode) => {
    if (!supabase) { setAuthMsg("Supabase is not configured. Check your .env.local file."); return; }
    if (!authEmail || !authPassword) { setAuthMsg("Enter an email and password first."); return; }
    setAuthMode(mode);
    setAuthMsg(mode === "signup" ? "Creating account…" : "Signing in…");
    try {
      const args = { email: authEmail.trim(), password: authPassword };
      const { error } = mode === "signup"
        ? await supabase.auth.signUp(args)
        : await supabase.auth.signInWithPassword(args);
      if (error) throw error;
      setAuthMsg(mode === "signup" ? "Account created. If email confirmation is enabled, confirm your email, then sign in." : "Signed in. Cloud sync is now active.");
      setAuthPassword("");
    } catch (e) { setAuthMsg(e.message || "Auth failed."); }
  };
  const signOut = async () => {
    if (!supabase) return;
    await supabase.auth.signOut();
    setSession(null);
    setSync("local");
    setAuthMsg("Signed out. This device will keep using local storage.");
  };
  const createAccount = async () => {
    if (!supabase) { setSignupMsg("Supabase is not configured."); return; }
    if (!signupEmail || !signupPassword) { setSignupMsg("Enter an email and password."); return; }
    if (signupPassword !== signupConfirm) { setSignupMsg("Passwords don't match."); return; }
    if (signupPassword.length < 6) { setSignupMsg("Password must be at least 6 characters."); return; }
    setSignupMsg("Creating account…");
    try {
      const { error } = await supabase.auth.signUp({ email: signupEmail.trim(), password: signupPassword });
      if (error) throw error;
      setSignupMsg("Account created. If email confirmation is enabled, confirm your email, then sign in.");
      setSignupPassword(""); setSignupConfirm("");
    } catch (e) { setSignupMsg(e.message || "Could not create account."); }
  };
  const sendPasswordReset = async () => {
    if (!supabase) { setAuthMsg("Supabase is not configured."); return; }
    if (!authEmail) { setAuthMsg("Enter your email above first, then tap Forgot password."); return; }
    setAuthMsg("Sending reset link…");
    try {
      const { error } = await supabase.auth.resetPasswordForEmail(authEmail.trim(), { redirectTo: window.location.href });
      if (error) throw error;
      setAuthMsg("If that email has an account, a reset link is on its way.");
    } catch (e) { setAuthMsg(e.message || "Could not send reset link."); }
  };
  const isOwner = !!user && OWNER_EMAIL && (user.email || "").toLowerCase() === OWNER_EMAIL;
  const loadAdminUsers = async () => {
    if (!session?.access_token) return;
    setAdminMsg("Loading users…");
    try {
      const { users } = await callAdminFn(session.access_token, "listUsers");
      setAdminUsers(users);
      setAdminMsg("");
    } catch (e) { setAdminMsg(e.message || "Could not load users."); }
  };
  const loadAdminUserData = async (u) => {
    setAdminSelected(u);
    setAdminSelectedData(null);
    setAdminNewPassword("");
    setAdminMsg("Loading data…");
    try {
      const { data } = await callAdminFn(session.access_token, "getUserData", { userId: u.id });
      setAdminSelectedData(data);
      setAdminMsg("");
    } catch (e) { setAdminMsg(e.message || "Could not load that user's data."); }
  };
  const resetAdminUserPassword = async () => {
    if (!adminSelected || !adminNewPassword) return;
    setAdminMsg("Setting new password…");
    try {
      await callAdminFn(session.access_token, "resetPassword", { userId: adminSelected.id, newPassword: adminNewPassword });
      setAdminMsg(`Password updated for ${adminSelected.email}. Share the new password with them directly.`);
      setAdminNewPassword("");
    } catch (e) { setAdminMsg(e.message || "Could not reset password."); }
  };
  const uploadLocalToCloud = async () => {
    if (!user) return;
    try {
      setSync("saving");
      await saveCloudState(user.id, { days: data, verses, labels });
      setSync("synced");
      setAuthMsg("Uploaded this device's data to Supabase.");
    } catch (e) { setSync("error"); setAuthMsg(e.message || "Cloud upload failed."); }
  };
  const downloadCloudToDevice = async () => {
    if (!user) return;
    try {
      setSync("loading");
      const cloudPayload = await fetchCloudState(user.id);
      if (!cloudPayload) { setAuthMsg("No cloud data found yet."); setSync("synced"); return; }
      setData(cloudPayload.days);
      setVerses(cloudPayload.verses);
      setLabels(cloudPayload.labels);
      await setStoredPayload("sadhana-v3", JSON.stringify(cloudPayload));
      setSync("synced");
      setAuthMsg("Downloaded cloud data to this device.");
    } catch (e) { setSync("error"); setAuthMsg(e.message || "Cloud download failed."); }
  };
  const mergeCloudAndDevice = async () => {
    if (!user) return;
    try {
      setSync("loading");
      const cloudPayload = await fetchCloudState(user.id);
      const merged = mergePayloads(cloudPayload || {}, { days: data, verses, labels });
      setData(merged.days);
      setVerses(merged.verses);
      setLabels(merged.labels);
      await setStoredPayload("sadhana-v3", JSON.stringify(merged));
      await saveCloudState(user.id, merged);
      setSync("synced");
      setAuthMsg("Merged cloud and this device, then saved the merged version.");
    } catch (e) { setSync("error"); setAuthMsg(e.message || "Cloud merge failed."); }
  };

  const exportBackupJson = () => {
    const payload = normalisePayload({ days: data, verses, labels });
    const blob = new Blob([JSON.stringify(payload, null, 2)], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `sadhana-backup-${todayKey()}.json`;
    a.click();
    URL.revokeObjectURL(url);
  };
  const triggerImport = (mode) => {
    setImportMode(mode);
    fileInputRef.current?.click();
  };
  const importBackupJson = async (file, mode) => {
    try {
      const text = await file.text();
      const parsed = normalisePayload(JSON.parse(text));
      if (!payloadHasContent(parsed)) { setBackupMsg("That file has no sadhana data in it."); return; }
      const next = mode === "replace" ? parsed : mergePayloads(parsed, { days: data, verses, labels });
      await persist(next.days, next.verses, next.labels);
      setBackupMsg(mode === "replace" ? "Replaced current data with the backup file." : "Merged the backup file into your current data.");
    } catch (e) {
      setBackupMsg(e.message ? `Could not read that backup file: ${e.message}` : "Could not read that backup file.");
    }
  };
  const handleImportFile = (e) => {
    const file = e.target.files?.[0];
    e.target.value = "";
    if (file) importBackupJson(file, importMode);
  };

  const exportExcel = (period) => {
    const opt = exportOptions;
    const { start, end, label, title } = periodWindow(period);
    const inRange = (d) => !start || (d.date >= start && d.date <= end);
    const days = A.dayList.filter(inRange);

    if (!days.length) {
      window.alert(`No logged days found for ${title.toLowerCase()}.`);
      return;
    }

    const totalJapaMin = days.reduce((a, d) => a + d.rounds.reduce((x, r) => x + r.dur, 0), 0);
    const sleepDays = days.filter((d) => d.sleepH != null);
    const fullRoundDays = days.filter((d) => d.n >= TARGET).length;

    const summaryRows = [
      { Metric: "Export period", Value: title },
      { Metric: "Start date", Value: start ? todayKey(start) : "First logged day" },
      { Metric: "End date", Value: end ? todayKey(end) : "Latest logged day" },
      { Metric: "Days logged", Value: days.length },
      { Metric: "16-round days", Value: fullRoundDays },
      { Metric: "Average day score", Value: days.length ? Math.round(mean(days.map((d) => d.score))) : 0 },
      { Metric: "Total rounds", Value: days.reduce((a, d) => a + d.n, 0) },
      { Metric: "Total japa minutes", Value: round1(totalJapaMin) },
      { Metric: "Total hearing minutes", Value: days.reduce((a, d) => a + d.hm, 0) },
      { Metric: "Total reading minutes", Value: days.reduce((a, d) => a + d.rm, 0) },
      { Metric: "Average net sleep hours", Value: sleepDays.length ? round2(mean(sleepDays.map((d) => d.sleepH))) : "" },
      { Metric: "Total night-awake minutes", Value: days.reduce((a, d) => a + (d.awakeMin || 0), 0) },
      { Metric: "Verse practice days", Value: days.filter((d) => versePracticed(d.e)).length },
      { Metric: "Verse recitation days", Value: days.filter((d) => d.e.versesRecited).length },
      { Metric: "Prayers recited days", Value: days.filter((d) => prayersDone(d.e)).length },
      { Metric: "Cooked for deities days", Value: days.filter((d) => d.e.worship?.cooked).length },
      { Metric: "Journaled days", Value: days.filter((d) => d.e.journaled).length },
    ];

    const dailyRows = days.map((d) => ({
      Date: d.k,
      Day: WD[d.wd],
      Score: d.score,
      Rounds: d.n,
      "Completed 16 rounds": yesNo(d.n >= TARGET),
      "Total japa minutes": round1(d.rounds.reduce((a, r) => a + r.dur, 0)),
      "Average round minutes": d.avgDur == null ? "" : round2(d.avgDur),
      "First round start": d.firstH == null ? "" : fmtH(d.firstH),
      "Last round end": d.lastEndH == null ? "" : fmtH(d.lastEndH),
      "Hearing minutes": d.hm,
      "Reading minutes": d.rm,
      "Worship %": Math.round(d.wp * 100),
      "Verses revised": yesNo(d.e.versesRevised),
      "Verses recited": yesNo(d.e.versesRecited),
      "Any verse practice": yesNo(versePracticed(d.e)),
      "Prayers recited": yesNo(prayersDone(d.e)),
      "Sleep time": d.e.sleepTime || "",
      "Wake time": d.e.wakeTime || "",
      "Gross sleep hours": d.grossSleepH == null ? "" : round2(d.grossSleepH),
      "Times awake": d.awakeCount || 0,
      "Awake minutes": d.awakeMin || 0,
      "Net sleep hours": d.sleepH == null ? "" : round2(d.sleepH),
      "Cooked for deities": yesNo(d.e.worship?.cooked),
      "Morning arati": yesNo(d.e.worship?.morningArati),
      "Evening arati": yesNo(d.e.worship?.eveningArati),
      Ekadasi: yesNo(d.ek),
      Festival: yesNo(d.fest),
      Note: d.e.note || "",
    }));

    const roundRows = days.flatMap((d) => d.rounds.map((r, i) => ({
      Date: d.k,
      "Round #": i + 1,
      Start: fmtT(r.start),
      End: fmtT(r.end),
      "Duration minutes": round2(r.dur),
    })));

    const hearingRows = days.flatMap((d) => (d.e.hearing || []).map((h, i) => ({
      Date: d.k,
      "Session #": i + 1,
      Speaker: h.speaker || "",
      Minutes: +h.minutes || 0,
      "Logged at": h.ts ? fmtT(h.ts) : "",
    })));

    const readingRows = days.flatMap((d) => (d.e.reading || []).map((r, i) => ({
      Date: d.k,
      "Session #": i + 1,
      Book: r.book || "",
      Section: r.section || "",
      Minutes: +r.minutes || 0,
      "Logged at": r.ts ? fmtT(r.ts) : "",
    })));

    const worshipRows = days.map((d) => {
      const row = { Date: d.k };
      WORSHIP_ITEMS.forEach(([key, label]) => { row[label] = yesNo(d.e.worship?.[key]); });
      return row;
    });

    const awakeRowsForExport = days.flatMap((d) => (d.awakePeriods || []).map((p, i) => ({
      Date: d.k,
      "Awakening #": i + 1,
      "Time awake": p.time || "",
      "Minutes awake": +p.minutes || 0,
    })));

    const verseRows = verses.map((v) => {
      const hist = v.history || [];
      const recHist = v.recitationHistory || [];
      const periodHist = start ? hist.filter((k) => k >= todayKey(start) && k <= todayKey(end)) : hist;
      const periodRecHist = start ? recHist.filter((k) => k >= todayKey(start) && k <= todayKey(end)) : recHist;
      return {
        Reference: v.ref || "",
        "First words": v.text || "",
        Status: v.status || "",
        "Added at": v.addedAt ? new Date(v.addedAt).toLocaleString("en-GB") : "",
        "Memorised at": v.memorisedAt ? new Date(v.memorisedAt).toLocaleString("en-GB") : "",
        "Last revised": v.lastRevised || "",
        "Last recited": v.lastRecited || "",
        "Practices in period": periodHist.length + periodRecHist.length,
        "Revisions in period": periodHist.length,
        "Recitations in period": periodRecHist.length,
        "Total revisions": hist.length,
        "Total recitations": recHist.length,
        "Revision dates in period": periodHist.join(", "),
        "Recitation dates in period": periodRecHist.join(", "),
      };
    });

    const versePracticeRows = verses.flatMap((v) => [
      ...(v.history || []).map((k) => ({ Date: k, Reference: v.ref || "", Type: "Revision" })),
      ...(v.recitationHistory || []).map((k) => ({ Date: k, Reference: v.ref || "", Type: "Recitation" })),
    ]).filter((r) => !start || (r.Date >= todayKey(start) && r.Date <= todayKey(end))).sort((a, b) => a.Date.localeCompare(b.Date));

    const visualOverviewRows = days.map((d) => ({
      Date: d.k,
      Score: d.score,
      Rounds: d.n,
      "Hearing minutes": d.hm,
      "Reading minutes": d.rm,
      "Net sleep hours": d.sleepH == null ? "" : round2(d.sleepH),
      "Wake time hour": d.wake == null ? "" : round2(d.wake),
      "Worship %": Math.round(d.wp * 100),
      "Verse practice": yesNo(versePracticed(d.e)),
    }));
    const visualJapaRows = days.map((d) => ({
      Date: d.k,
      Rounds: d.n,
      "Total japa minutes": round1(d.rounds.reduce((a, r) => a + r.dur, 0)),
      "Average round minutes": d.avgDur == null ? "" : round2(d.avgDur),
      "First round hour": d.firstH == null ? "" : round2(d.firstH),
      "Last round end hour": d.lastEndH == null ? "" : round2(d.lastEndH),
    }));
    const visualHearingReadingRows = days.map((d) => ({ Date: d.k, "Hearing minutes": d.hm, "Reading minutes": d.rm }));
    const visualSleepRows = days.map((d) => ({
      Date: d.k,
      "Sleep time": d.e.sleepTime || "",
      "Wake time": d.e.wakeTime || "",
      "Gross sleep hours": d.grossSleepH == null ? "" : round2(d.grossSleepH),
      "Net sleep hours": d.sleepH == null ? "" : round2(d.sleepH),
      "Awake count": d.awakeCount || 0,
      "Awake minutes": d.awakeMin || 0,
      Score: d.score,
    }));
    const visualWorshipRows = days.map((d) => ({ Date: d.k, "Worship %": Math.round(d.wp * 100), ...Object.fromEntries(WORSHIP_ITEMS.map(([key, label]) => [label, yesNo(d.e.worship?.[key])])) }));
    const visualCompositeRows = days.map((d) => ({
      Date: d.k,
      Chanting: Math.round(Math.min(d.n / TARGET, 1) * 100),
      Hearing: Math.round(Math.min(d.hm / 30, 1) * 100),
      Reading: Math.round(Math.min(d.rm / 30, 1) * 100),
      Worship: Math.round(d.wp * 100),
      Verses: versePracticed(d.e) ? 100 : 0,
      Prayers: prayersDone(d.e) ? 100 : 0,
      Score: d.score,
    }));

    const wb = XLSX.utils.book_new();
    wb.Props = {
      Title: `Sadhana export - ${title}`,
      Subject: "Daily sadhana tracker export",
      Author: "Sundar Chaitanya Das",
      CreatedDate: new Date(),
    };

    let sheetsAdded = 0;
    const addIf = (on, rows, name, empty) => { if (on) { appendSheet(wb, rows, name, empty); sheetsAdded++; } };
    addIf(opt.dataSummary, summaryRows, "Summary");
    addIf(opt.dataDaily, dailyRows, "Daily log");
    addIf(opt.dataJapa, roundRows, "Japa rounds", "No japa rounds in this period");
    addIf(opt.dataHearing, hearingRows, "Hearing", "No hearing sessions in this period");
    addIf(opt.dataReading, readingRows, "Reading", "No reading sessions in this period");
    addIf(opt.dataWorship, worshipRows, "Worship");
    addIf(opt.dataSleep, awakeRowsForExport, "Night awakenings", "No night awakenings in this period");
    addIf(opt.dataVerses, verseRows, "Verses", "No verses saved yet");
    addIf(opt.dataVersePractice, versePracticeRows, "Verse practice", "No verse practice in this period");
    addIf(opt.vizOverview, visualOverviewRows, "Viz overview");
    addIf(opt.vizJapa, visualJapaRows, "Viz japa");
    addIf(opt.vizHearingReading, visualHearingReadingRows, "Viz hearing reading");
    addIf(opt.vizSleep, visualSleepRows, "Viz sleep");
    addIf(opt.vizWorship, visualWorshipRows, "Viz worship");
    addIf(opt.vizComposite, visualCompositeRows, "Viz composite");
    if (!sheetsAdded) appendSheet(wb, [{ Message: "No export sections selected." }], "Selection");

    XLSX.writeFile(wb, `sadhana-${period}-${label}.xlsx`);
  };

  if (!loaded) return <div style={{ minHeight: "100vh", background: C.ivory, display: "flex", alignItems: "center", justifyContent: "center", color: C.faint, fontFamily: "system-ui" }}>Loading your sadhana…</div>;

  const liveSec = day.activeStart ? Math.floor((Date.now() - new Date(day.activeStart)) / 1000) : 0;
  const VIZ_TABS = [
    ["jdur", "Japa · duration"], ["jtime", "Japa · time"], ["hear", "Hearing"], ["read", "Reading"],
    ["wor", "Worship"], ["ver", "Verses"], ["sleep", "Sleep"], ["comp", "Composite"],
  ];

  return (
    <div style={{ minHeight: "100vh", background: C.ivory, color: C.ink, fontFamily: "system-ui, -apple-system, sans-serif" }}>
      <style>{`@import url('https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,400;9..144,600;9..144,700&display=swap');
        input,select,textarea{accent-color:${C.saffron};}`}</style>
      <div style={{ maxWidth: 700, margin: "0 auto", padding: "26px 14px 60px" }}>

        <header style={{ marginBottom: 18 }}>
          <div style={{ fontSize: 11, letterSpacing: ".14em", textTransform: "uppercase", color: C.tulsi, fontWeight: 600 }}>Sundar Chaitanya Das · Daily Sadhana</div>
          <div style={{ display: "flex", alignItems: "baseline", justifyContent: "space-between", flexWrap: "wrap", gap: 8 }}>
            <h1 style={{ fontFamily: "'Fraunces', Georgia, serif", fontWeight: 700, fontSize: 32, margin: "4px 0 0", color: C.maroon }}>Sadhana</h1>
            <div style={{ fontSize: 13, color: C.faint }}>
              {A.streak > 0 ? <><b style={{ color: C.saffron }}>{A.streak}</b>-day streak at 16 rounds</> : "Begin the streak today"}
              {save === "saving" && " · saving…"}{save === "saved" && " · saved ✓"}{save === "error" && " · save failed"}
              {user && <span> · cloud {sync === "synced" ? "✓" : sync === "saving" ? "saving…" : sync === "loading" ? "loading…" : sync === "error" ? "error" : "on"}</span>}
            </div>
          </div>
        </header>

        <nav style={{ display: "flex", gap: 6, marginBottom: 16, flexWrap: "wrap" }}>
          {[["today", "Today"], ["japa", "Japa"], ["trends", "Trends"], ["week", "Review"], ["account", "Account"], ["export", "Export"], ["viz", "Visuals"], ["about", "About"], ...(isOwner ? [["admin", "Admin"]] : [])].map(([k, l]) => (
            <button key={k} onClick={() => setTab(k)} style={{
              padding: "8px 14px", borderRadius: 999, fontSize: 13, cursor: "pointer", fontWeight: 600,
              border: `1.5px solid ${tab === k ? C.maroon : C.line}`,
              background: tab === k ? C.maroon : C.card, color: tab === k ? "#fff" : C.ink,
            }}>{l}</button>
          ))}
        </nav>

        {/* ===== TODAY ===== */}
        {tab === "today" && (
          <div style={{ display: "grid", gap: 14 }}>
            <section style={{ ...cardS, display: "flex", gap: 12, flexWrap: "wrap", alignItems: "center", justifyContent: "space-between" }}>
              <div style={{ fontFamily: "'Fraunces', Georgia, serif", fontSize: 15, color: C.maroon, fontWeight: 600 }}>
                Chanting {day.rounds.length}/{TARGET}{avgRound && <span style={{ color: C.faint, fontWeight: 400 }}> · avg {avgRound} min/round</span>}
              </div>
              <button onClick={() => setTab("japa")} style={btnPri}>Open japa timer →</button>
            </section>

            <section style={cardS}>
              <h2 style={h2S}>Hearing — {hearingMin(day)} min</h2>
              {day.hearing.map((h, i) => (
                <div key={i} style={{ display: "flex", justifyContent: "space-between", fontSize: 14, padding: "6px 0", borderBottom: `1px solid ${C.line}` }}>
                  <span>{h.speaker}{h.ts && <span style={{ color: C.faint }}> · {fmtT(h.ts)}</span>}</span>
                  <span style={{ color: C.faint }}>{h.minutes} min
                    <button onClick={() => update({ hearing: day.hearing.filter((_, j) => j !== i) })} style={{ marginLeft: 10, border: "none", background: "none", color: C.maroon, cursor: "pointer" }}>×</button>
                  </span>
                </div>
              ))}
              <div style={{ display: "flex", gap: 8, marginTop: 10, flexWrap: "wrap" }}>
                <LabelField value={hForm.speaker} onChange={(v) => setHForm({ ...hForm, speaker: v })}
                  options={labels.speakers} onAddLabel={(v) => addLabel("speakers", v)}
                  placeholder="Speaker" style={{ flex: 2, minWidth: 140 }} />
                <input placeholder="Min" type="number" value={hForm.minutes} onChange={(e) => setHForm({ ...hForm, minutes: e.target.value })} style={{ ...inpS, width: 70 }} />
                <input title="Time (optional, defaults to now)" type="time" value={hForm.time} onChange={(e) => setHForm({ ...hForm, time: e.target.value })} style={{ ...inpS, width: 100 }} />
                <button style={btnS} onClick={() => {
                  if (!hForm.speaker || !hForm.minutes) return;
                  update({ hearing: [...day.hearing, { speaker: hForm.speaker, minutes: hForm.minutes, ts: tsForTime(hForm.time) }] });
                  setHForm({ speaker: "", minutes: "", time: "" });
                }}>Add</button>
              </div>
            </section>

            <section style={cardS}>
              <h2 style={h2S}>Reading — {readingMin(day)} min{readingPages(day) > 0 && <span style={{ color: C.faint, fontWeight: 400 }}> · {readingPages(day)} pages</span>}</h2>
              {day.reading.map((r, i) => (
                <div key={i} style={{ display: "flex", justifyContent: "space-between", fontSize: 14, padding: "6px 0", borderBottom: `1px solid ${C.line}` }}>
                  <span>{r.book}{r.section && <span style={{ color: C.faint }}> · {r.section}</span>}</span>
                  <span style={{ color: C.faint }}>{r.minutes} min{r.pages ? ` · ${r.pages}p` : ""}
                    <button onClick={() => update({ reading: day.reading.filter((_, j) => j !== i) })} style={{ marginLeft: 10, border: "none", background: "none", color: C.maroon, cursor: "pointer" }}>×</button>
                  </span>
                </div>
              ))}
              <div style={{ display: "flex", gap: 8, marginTop: 10, flexWrap: "wrap" }}>
                <LabelField value={rForm.book} onChange={(v) => setRForm({ ...rForm, book: v })}
                  options={labels.books} onAddLabel={(v) => addLabel("books", v)}
                  placeholder="Book" style={{ flex: 1, minWidth: 130 }} />
                <input placeholder="Section (e.g. SB 1.2.6)" value={rForm.section} onChange={(e) => setRForm({ ...rForm, section: e.target.value })} style={{ ...inpS, flex: 1, minWidth: 130 }} />
                <input placeholder="Min" type="number" value={rForm.minutes} onChange={(e) => setRForm({ ...rForm, minutes: e.target.value })} style={{ ...inpS, width: 70 }} />
                <input placeholder="Pages" type="number" value={rForm.pages} onChange={(e) => setRForm({ ...rForm, pages: e.target.value })} style={{ ...inpS, width: 70 }} />
                <input title="Time (optional, defaults to now)" type="time" value={rForm.time} onChange={(e) => setRForm({ ...rForm, time: e.target.value })} style={{ ...inpS, width: 100 }} />
                <button style={btnS} onClick={() => {
                  if (!rForm.minutes) return;
                  update({ reading: [...day.reading, { book: rForm.book, section: rForm.section, minutes: rForm.minutes, pages: rForm.pages, ts: tsForTime(rForm.time) }] });
                  setRForm({ book: rForm.book, section: "", minutes: "", pages: "", time: "" });
                }}>Add</button>
              </div>
            </section>

            <section style={{ ...cardS, display: "grid", gap: 8 }}>
              <h2 style={h2S}>Worship</h2>
              <div style={{ fontSize: 12, color: C.faint, marginBottom: 2 }}>Morning programme</div>
              {WORSHIP_ITEMS.slice(0, 5).map(([k, l]) => (
                <Toggle key={k} label={l} value={!!day.worship[k]} onChange={(v) => update({ worship: { ...day.worship, [k]: v } })} />
              ))}
              <div style={{ fontSize: 12, color: C.faint, margin: "6px 0 2px" }}>Through the day</div>
              {WORSHIP_ITEMS.slice(5).map(([k, l]) => (
                <Toggle key={k} label={l} value={!!day.worship[k]} onChange={(v) => update({ worship: { ...day.worship, [k]: v } })} />
              ))}
              <div style={{ fontSize: 12, color: C.faint, margin: "10px 0 2px" }}>Deity dressing (if any)</div>
              {dressingRows(day).map((dr, i) => (
                <div key={i} style={{ display: "flex", justifyContent: "space-between", fontSize: 14, padding: "6px 0", borderBottom: `1px solid ${C.line}` }}>
                  <span>{dr.label}{dr.ts && <span style={{ color: C.faint }}> · {fmtT(dr.ts)}</span>}</span>
                  <button onClick={() => update({ deityDressing: dressingRows(day).filter((_, j) => j !== i) })} style={{ border: "none", background: "none", color: C.maroon, cursor: "pointer" }}>×</button>
                </div>
              ))}
              <div style={{ display: "flex", gap: 8, marginTop: 4, flexWrap: "wrap" }}>
                <LabelField value={dForm.label} onChange={(v) => setDForm({ ...dForm, label: v })}
                  options={labels.deities} onAddLabel={(v) => addLabel("deities", v)}
                  placeholder="Deity" style={{ flex: 2, minWidth: 140 }} />
                <input title="Time (optional, defaults to now)" type="time" value={dForm.time} onChange={(e) => setDForm({ ...dForm, time: e.target.value })} style={{ ...inpS, width: 100 }} />
                <button style={btnS} onClick={() => {
                  if (!dForm.label) return;
                  update({ deityDressing: [...dressingRows(day), { label: dForm.label, ts: tsForTime(dForm.time) }] });
                  setDForm({ label: "", time: "" });
                }}>Add</button>
              </div>
            </section>

            <section style={cardS}>
              <h2 style={h2S}>Verses & recitation{verses.length > 0 && <span style={{ color: C.faint, fontWeight: 400, fontSize: 13 }}> — {verses.filter((v) => v.lastRevised === tk || v.lastRecited === tk).length}/{verses.length} practised today</span>}</h2>
              {verses.length === 0 && <div style={{ fontSize: 13, color: C.faint, marginBottom: 8 }}>Add the verses you're memorising or reciting — practise them here each day.</div>}
              {verses.map((v, i) => {
                const revised = v.lastRevised === tk;
                const recited = v.lastRecited === tk;
                return (
                  <div key={i} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "7px 0", borderBottom: `1px solid ${C.line}`, fontSize: 14, gap: 8, flexWrap: "wrap" }}>
                    <span style={{ minWidth: 120, flex: 1 }}>
                      <b>{v.ref}</b>
                      {v.status === "memorised" && <span style={{ color: C.tulsi, fontSize: 11, marginLeft: 6 }}>MEMORISED</span>}
                      {v.text && <div style={{ color: C.faint, fontSize: 12 }}>{v.text}</div>}
                    </span>
                    <span style={{ display: "flex", gap: 6, flexWrap: "wrap", justifyContent: "flex-end" }}>
                      <button onClick={() => {
                        const hist = new Set(v.recitationHistory || []);
                        if (recited) hist.delete(tk); else hist.add(tk);
                        const vs = verses.map((x, j) => j === i ? { ...x, lastRecited: recited ? null : tk, recitationHistory: [...hist] } : x);
                        const anyRevised = vs.some((x) => x.lastRevised === tk);
                        const anyRecited = vs.some((x) => x.lastRecited === tk);
                        persist({ ...data, [tk]: { ...day, versesRevised: anyRevised, versesRecited: anyRecited } }, vs);
                      }} style={{ ...btnS, padding: "5px 12px", fontSize: 12, background: recited ? "#F0F4EC" : "#fff", borderColor: recited ? C.tulsi : C.line, color: recited ? C.tulsi : C.ink }}>
                        {recited ? "Recited ✓" : "Recite"}
                      </button>
                      <button onClick={() => {
                        const hist = new Set(v.history || []);
                        if (revised) hist.delete(tk); else hist.add(tk);
                        const vs = verses.map((x, j) => j === i ? { ...x, lastRevised: revised ? null : tk, history: [...hist] } : x);
                        const anyRevised = vs.some((x) => x.lastRevised === tk);
                        const anyRecited = vs.some((x) => x.lastRecited === tk);
                        persist({ ...data, [tk]: { ...day, versesRevised: anyRevised, versesRecited: anyRecited } }, vs);
                      }} style={{ ...btnS, padding: "5px 12px", fontSize: 12, background: revised ? "#F0F4EC" : "#fff", borderColor: revised ? C.tulsi : C.line, color: revised ? C.tulsi : C.ink }}>
                        {revised ? "Revised ✓" : "Revise"}
                      </button>
                      {manageVerses && (
                        <>
                          <button onClick={() => persist(data, verses.map((x, j) => j === i ? { ...x, status: x.status === "memorised" ? "learning" : "memorised", memorisedAt: x.status === "memorised" ? null : new Date().toISOString() } : x))}
                            style={{ ...btnS, padding: "5px 10px", fontSize: 12 }}>{v.status === "memorised" ? "↺" : "✓ Mem"}</button>
                          <button onClick={() => persist(data, verses.filter((_, j) => j !== i))} style={{ border: "none", background: "none", color: C.maroon, cursor: "pointer" }}>×</button>
                        </>
                      )}
                    </span>
                  </div>
                );
              })}
              <div style={{ display: "flex", gap: 8, marginTop: 10, flexWrap: "wrap", alignItems: "center" }}>
                {manageVerses && (
                  <>
                    <LabelField value={vForm.book} onChange={(v) => setVForm({ ...vForm, book: v })}
                      options={labels.books} onAddLabel={(v) => addLabel("books", v)}
                      placeholder="Book" style={{ width: 130 }} />
                    <input placeholder="Ref (2.13)" value={vForm.ref} onChange={(e) => setVForm({ ...vForm, ref: e.target.value })} style={{ ...inpS, width: 90 }} />
                    <input placeholder="First words" value={vForm.text} onChange={(e) => setVForm({ ...vForm, text: e.target.value })} style={{ ...inpS, flex: 1, minWidth: 120 }} />
                    <button style={btnS} onClick={() => {
                      if (!vForm.ref) return;
                      const fullRef = vForm.book ? `${vForm.book} ${vForm.ref}` : vForm.ref;
                      persist(data, [...verses, { book: vForm.book, ref: fullRef, text: vForm.text, status: "learning", lastRevised: null, lastRecited: null, addedAt: new Date().toISOString(), history: [], recitationHistory: [] }]);
                      setVForm({ book: vForm.book, ref: "", text: "" });
                    }}>Add</button>
                  </>
                )}
                <button onClick={() => setManageVerses(!manageVerses)} style={{ ...btnS, marginLeft: "auto", fontSize: 12, color: C.faint }}>{manageVerses ? "Done" : "Manage"}</button>
              </div>
            </section>

            <section style={{ ...cardS, display: "grid", gap: 8 }}>
              <h2 style={h2S}>Prayers & rest</h2>
              {prayerRows(day).map((p, i) => (
                <div key={i} style={{ display: "flex", justifyContent: "space-between", fontSize: 14, padding: "6px 0", borderBottom: `1px solid ${C.line}` }}>
                  <span>{p.label}{p.ts && <span style={{ color: C.faint }}> · {fmtT(p.ts)}</span>}</span>
                  <button onClick={() => update({ prayers: prayerRows(day).filter((_, j) => j !== i) })} style={{ border: "none", background: "none", color: C.maroon, cursor: "pointer" }}>×</button>
                </div>
              ))}
              <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
                <LabelField value={pForm.label} onChange={(v) => setPForm({ ...pForm, label: v })}
                  options={labels.prayers} onAddLabel={(v) => addLabel("prayers", v)}
                  placeholder="Prayer" style={{ flex: 2, minWidth: 140 }} />
                <input title="Time (optional, defaults to now)" type="time" value={pForm.time} onChange={(e) => setPForm({ ...pForm, time: e.target.value })} style={{ ...inpS, width: 100 }} />
                <button style={btnS} onClick={() => {
                  if (!pForm.label) return;
                  update({ prayers: [...prayerRows(day), { label: pForm.label, ts: tsForTime(pForm.time) }] });
                  setPForm({ label: "", time: "" });
                }}>Add</button>
              </div>
              <div style={{ display: "flex", gap: 10, marginTop: 4, flexWrap: "wrap" }}>
                <label style={{ fontSize: 13, color: C.faint, flex: 1, minWidth: 130 }}>Woke up
                  <input type="time" value={day.wakeTime} onChange={(e) => update({ wakeTime: e.target.value })} style={{ ...inpS, width: "100%", marginTop: 4 }} />
                </label>
                <label style={{ fontSize: 13, color: C.faint, flex: 1, minWidth: 130 }}>Slept (last night)
                  <input type="time" value={day.sleepTime} onChange={(e) => update({ sleepTime: e.target.value })} style={{ ...inpS, width: "100%", marginTop: 4 }} />
                </label>
              </div>

              <div style={{ display: "grid", gap: 8, marginTop: 6 }}>
                <label style={{ fontSize: 13, color: C.faint, maxWidth: 180 }}>Times awake
                  <input type="number" min="0" max="12" value={awakeRows(day).length} onChange={(e) => setAwakeCount(e.target.value)} style={{ ...inpS, width: "100%", marginTop: 4 }} />
                </label>
                {awakeRows(day).map((p, i) => (
                  <div key={i} style={{ display: "flex", gap: 8, flexWrap: "wrap", alignItems: "center" }}>
                    <span style={{ fontSize: 12, color: C.faint, width: 74 }}>Awake #{i + 1}</span>
                    <label style={{ fontSize: 12, color: C.faint, flex: 1, minWidth: 120 }}>Time
                      <input type="time" value={p.time || ""} onChange={(e) => updateAwakePeriod(i, { time: e.target.value })} style={{ ...inpS, width: "100%", marginTop: 3 }} />
                    </label>
                    <label style={{ fontSize: 12, color: C.faint, flex: 1, minWidth: 120 }}>Length awake (min)
                      <input type="number" min="0" value={p.minutes || ""} onChange={(e) => updateAwakePeriod(i, { minutes: e.target.value })} style={{ ...inpS, width: "100%", marginTop: 3 }} />
                    </label>
                  </div>
                ))}
                {awakeMin(day) > 0 && <div style={{ fontSize: 12, color: C.faint }}>Awake total: <b style={{ color: C.maroon }}>{awakeMin(day)} min</b>. Net sleep adjusts automatically.</div>}
              </div>

              <textarea value={day.note} onChange={(e) => update({ note: e.target.value })} placeholder='Realisation / note… (write "Ekadasi" to tag holy days)'
                style={{ ...inpS, width: "100%", minHeight: 56, resize: "vertical", marginTop: 4 }} />
              <div style={{ fontSize: 14 }}>Day score: <b style={{ color: C.saffron, fontSize: 18 }}>{dayScore(day)}</b>/100</div>
            </section>

            <section style={{ ...cardS, display: "grid", gap: 8 }}>
              <h2 style={h2S}>Journal</h2>
              <Toggle label="Journaled today" value={!!day.journaled} onChange={(v) => update({ journaled: v })} />
              <textarea value={day.journal || ""} onChange={(e) => update({ journal: e.target.value, journaled: e.target.value ? true : day.journaled })}
                placeholder="Today's journal entry…" style={{ ...inpS, width: "100%", minHeight: 90, resize: "vertical" }} />
            </section>
          </div>
        )}

        {/* ===== JAPA ===== */}
        {tab === "japa" && (
          <div style={{ display: "grid", gap: 14 }}>
            <section style={{ ...cardS, textAlign: "center" }}>
              <h2 style={h2S}>Round {Math.min(day.rounds.length + 1, TARGET)} of {TARGET}</h2>
              {day.activeStart ? (
                <>
                  <div style={{ fontFamily: "'Fraunces', Georgia, serif", fontSize: 48, fontWeight: 600, color: C.maroon }}>
                    {String(Math.floor(liveSec / 60)).padStart(2, "0")}:{String(liveSec % 60).padStart(2, "0")}
                  </div>
                  <div style={{ fontSize: 13, color: C.faint, marginBottom: 14 }}>Started {fmtT(day.activeStart)}</div>
                  <div style={{ display: "flex", gap: 10, justifyContent: "center" }}>
                    <button style={{ ...btnPri, padding: "12px 28px", fontSize: 15 }} onClick={finishRound}>Finish round</button>
                    <button style={btnS} onClick={cancelRound}>Cancel</button>
                  </div>
                </>
              ) : (
                <>
                  <button style={{ ...btnPri, padding: "14px 32px", fontSize: 16 }} onClick={startRound}>Start round</button>
                  <div style={{ marginTop: 14, display: "flex", gap: 8, justifyContent: "center", alignItems: "center", flexWrap: "wrap" }}>
                    <span style={{ fontSize: 13, color: C.faint }}>or log a finished round:</span>
                    <input type="number" placeholder="min" value={manualMin} onChange={(e) => setManualMin(e.target.value)} style={{ ...inpS, width: 70 }} />
                    <span style={{ fontSize: 13, color: C.faint }}>ending at</span>
                    <input title="Defaults to now" type="time" value={manualEnd} onChange={(e) => setManualEnd(e.target.value)} style={{ ...inpS, width: 100 }} />
                    <button style={btnS} onClick={() => { if (+manualMin > 0) { logManualRound(+manualMin, manualEnd); setManualMin(""); setManualEnd(""); } }}>Log</button>
                  </div>
                </>
              )}
            </section>

            <section style={cardS}>
              <Toggle label="Autochant — auto-start the next round when one finishes" value={autochant} onChange={toggleAutochant} />
            </section>

            <section style={cardS}>
              <h2 style={h2S}>Today's rounds{avgRound && <span style={{ color: C.faint, fontWeight: 400, fontSize: 14 }}> — avg {avgRound} min</span>}</h2>
              {day.rounds.length === 0 && <div style={{ fontSize: 13, color: C.faint }}>No rounds yet. Long rounds usually mean distraction — the timer keeps you honest.</div>}
              {day.rounds.map((r, i) => {
                const m = mins(r);
                const worst = roundDurations.length > 1 && m === Math.max(...roundDurations);
                return (
                  <div key={i} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "7px 0", borderBottom: `1px solid ${C.line}`, fontSize: 14 }}>
                    <span><b>#{i + 1}</b> <span style={{ color: C.faint }}>{fmtT(r.start)} → {fmtT(r.end)}</span></span>
                    <span>
                      <b style={{ color: worst ? C.maroon : C.tulsi }}>{m} min</b>
                      <button onClick={() => removeRound(i)} style={{ marginLeft: 10, border: "none", background: "none", color: C.maroon, cursor: "pointer" }}>×</button>
                    </span>
                  </div>
                );
              })}
              {roundDurations.length > 0 && (
                <>
                  <div style={{ marginTop: 14 }}>
                    <div style={{ fontSize: 12, color: C.faint, marginBottom: 4 }}>When in the day (4:00 → 23:00)</div>
                    <div style={{ position: "relative", height: 26, background: C.ivory, borderRadius: 6, overflow: "hidden" }}>
                      {day.rounds.map((r, i) => {
                        const s = new Date(r.start), e = new Date(r.end);
                        const toPct = (d) => Math.max(0, Math.min(100, ((d.getHours() + d.getMinutes() / 60 - 4) / 19) * 100));
                        const l = toPct(s), w = Math.max(0.8, toPct(e) - l);
                        return <div key={i} title={`#${i + 1} ${fmtT(r.start)}–${fmtT(r.end)}`} style={{ position: "absolute", left: `${l}%`, width: `${w}%`, top: 3, bottom: 3, background: C.saffron, borderRadius: 3 }} />;
                      })}
                    </div>
                    <div style={{ display: "flex", justifyContent: "space-between", fontSize: 10, color: C.faint }}>
                      <span>4am</span><span>9am</span><span>2pm</span><span>7pm</span><span>11pm</span>
                    </div>
                  </div>
                  <ResponsiveContainer width="100%" height={150}>
                    <BarChart data={roundDurations.map((m, i) => ({ r: `#${i + 1}`, min: m }))} margin={{ top: 14, right: 4, left: -26, bottom: 0 }}>
                      <XAxis dataKey="r" tick={tickS} /><YAxis tick={tickS} /><Tooltip />
                      <Bar dataKey="min" fill={C.saffron} radius={[4, 4, 0, 0]} />
                    </BarChart>
                  </ResponsiveContainer>
                </>
              )}
            </section>
          </div>
        )}

        {/* ===== TRENDS ===== */}
        {tab === "trends" && (
          <div style={{ display: "grid", gap: 14 }}>
            <section style={cardS}>
              <h2 style={h2S}>Rounds & round length — 30 days</h2>
              <ResponsiveContainer width="100%" height={200}>
                <ComposedChart data={trend} margin={{ top: 6, right: 6, left: -24, bottom: 0 }}>
                  <CartesianGrid stroke={C.line} strokeDasharray="3 3" />
                  <XAxis dataKey="label" tick={tickS} interval="preserveStartEnd" />
                  <YAxis tick={tickS} />
                  <Tooltip />
                  <Legend wrapperStyle={{ fontSize: 11 }} />
                  <Bar dataKey="rounds" name="Rounds" fill={C.saffronSoft} radius={[3, 3, 0, 0]} />
                  <Line type="monotone" dataKey="avgRound" name="Avg min/round" stroke={C.maroon} strokeWidth={2} dot={false} connectNulls />
                </ComposedChart>
              </ResponsiveContainer>
            </section>
            <section style={cardS}>
              <h2 style={h2S}>Hearing & reading — 30 days</h2>
              <LinesViz data={trend} x="label" lines={[{ k: "hearing", name: "Hearing (min)", c: C.tulsi }, { k: "reading", name: "Reading (min)", c: C.saffron }]} h={200} />
            </section>
            <section style={cardS}>
              <h2 style={h2S}>Day score — 30 days</h2>
              <LinesViz data={trend} x="label" lines={[{ k: "score", name: "Score", c: C.maroon }]} h={190} dom={[0, 100]} />
            </section>
            <section style={cardS}>
              <h2 style={h2S}>16-round heatmap — 16 weeks</h2>
              <BinaryCalendar A={A} test={(d) => d.n >= TARGET} />
              <div style={{ fontSize: 11, color: C.faint, marginTop: 6 }}>Green = all 16 rounds chanted.</div>
            </section>
          </div>
        )}

        {/* ===== REVIEW ===== */}
        {tab === "week" && (
          <div style={{ display: "grid", gap: 14 }}>
            <section style={cardS}>
              <h2 style={h2S}>This week</h2>
              <div style={{ display: "grid", gridTemplateColumns: "repeat(3,1fr)", gap: 12 }}>
                <Stat v={`${weekStats.full}/7`} l="16-round days" />
                <Stat v={weekStats.avgScore} l="avg score" />
                <Stat v={`${weekStats.cooked}/7`} l="cooked" />
                <Stat v={`${Math.round(weekStats.hearing / 60 * 10) / 10}h`} l="heard" />
                <Stat v={`${Math.round(weekStats.reading / 60 * 10) / 10}h`} l="read" />
                <Stat v={weekStats.logged} l="days logged" />
              </div>
            </section>
            <section style={cardS}>
              <h2 style={h2S}>Limb balance — 7 days</h2>
              <ResponsiveContainer width="100%" height={250}>
                <RadarChart data={radar} outerRadius="72%">
                  <PolarGrid stroke={C.line} />
                  <PolarAngleAxis dataKey="axis" tick={{ fontSize: 11, fill: C.ink }} />
                  <Radar dataKey="v" stroke={C.saffron} fill={C.saffron} fillOpacity={0.25} />
                </RadarChart>
              </ResponsiveContainer>
            </section>
            <section style={cardS}>
              <h2 style={h2S}>Voices this week</h2>
              <Rank items={speakerStats} unit=" min" />
              {!speakerStats.length && <Empty m="No hearing logged this week." />}
            </section>
            <section style={cardS}>
              <h2 style={h2S}>Sleep & rise — 7 days</h2>
              {A.last7.map((k) => {
                const e = A.byKey[k];
                return (
                  <div key={k} style={{ display: "flex", justifyContent: "space-between", fontSize: 13, padding: "5px 0", borderBottom: `1px solid ${C.line}` }}>
                    <span>{fmtDate(k)}</span>
                    <span style={{ color: C.faint }}>
                      {e?.e.sleepTime || "—"} → {e?.e.wakeTime || "—"}
                      {e?.awakeMin > 0 && <span style={{ color: C.maroon, marginLeft: 8 }}>awake {e.awakeCount}×/{e.awakeMin}m</span>}
                      {e?.sleepH != null && <b style={{ color: e.sleepH < 6 ? C.maroon : C.tulsi, marginLeft: 8 }}>{e.sleepH.toFixed(1)}h net</b>}
                    </span>
                  </div>
                );
              })}
            </section>
            <section style={cardS}>
              <h2 style={h2S}>Recent notes</h2>
              {A.dayList.slice(-7).reverse().filter((d) => d.e.note).map((d) => (
                <div key={d.k} style={{ padding: "7px 0", borderBottom: `1px solid ${C.line}` }}>
                  <div style={{ fontSize: 11, color: C.faint }}>{fmtDate(d.k)}{d.ek && <span style={{ color: C.maroon }}> · Ekadasi</span>}</div>
                  <div style={{ fontSize: 13 }}>{d.e.note}</div>
                </div>
              ))}
              {!A.dayList.slice(-7).some((d) => d.e.note) && <Empty m="No notes this week." />}
            </section>
            <section style={cardS}>
              <h2 style={h2S}>History — view a past day</h2>
              <p style={{ fontSize: 12, color: C.faint, marginTop: 0 }}>Read-only — to change a past entry, use Today's tab on the day itself.</p>
              <input type="date" value={historyDate} max={tk} onChange={(e) => setHistoryDate(e.target.value)} style={{ ...inpS, marginBottom: 14 }} />
              {historyDate && historyDate !== tk && <DaySummary e={data[historyDate]} />}
              {historyDate === tk && <div style={{ fontSize: 13, color: C.faint }}>That's today — see the Today tab to edit it.</div>}
            </section>
          </div>
        )}

        {/* ===== ACCOUNT ===== */}
        {tab === "account" && (
          <div style={{ display: "grid", gap: 14 }}>
            <section style={cardS}>
              <h2 style={h2S}>Account & cloud sync</h2>
              {!cloudEnabled && (
                <div style={{ fontSize: 13, color: C.maroon, lineHeight: 1.5 }}>
                  Supabase is not connected yet. Check that <code>.env.local</code> contains <code>VITE_SUPABASE_URL</code> and <code>VITE_SUPABASE_ANON_KEY</code>, then restart <code>npm.cmd run dev</code>.
                </div>
              )}
              {cloudEnabled && !user && (
                <div style={{ display: "grid", gap: 10 }}>
                  <p style={{ fontSize: 13, color: C.faint, margin: 0 }}>
                    Sign in to sync the same sadhana data across your laptop and phone. Your browser will still keep a local/offline copy.
                  </p>
                  <input type="email" placeholder="Email" value={authEmail} onChange={(e) => setAuthEmail(e.target.value)} style={inpS} />
                  <input type="password" placeholder="Password" value={authPassword} onChange={(e) => setAuthPassword(e.target.value)} style={inpS} />
                  <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
                    <button style={btnPri} onClick={() => signInOrUp("signin")}>Sign in</button>
                    <button style={btnS} onClick={() => { setShowSignup(true); setSignupEmail(authEmail); setSignupMsg(""); }}>Create account</button>
                  </div>
                  <button style={{ ...btnS, width: "fit-content", fontSize: 12, color: C.faint, border: "none", background: "none", padding: 0 }} onClick={sendPasswordReset}>Forgot password?</button>
                </div>
              )}
              {cloudEnabled && user && (
                <div style={{ display: "grid", gap: 10 }}>
                  <div style={{ fontSize: 13, color: C.ink }}>
                    Signed in as <b>{user.email}</b>. Status: <b style={{ color: sync === "error" ? C.maroon : C.tulsi }}>{sync}</b>.
                  </div>
                  <div style={{ display: "grid", gap: 8 }}>
                    <button style={btnPri} onClick={uploadLocalToCloud}>Upload this device to cloud</button>
                    <button style={btnS} onClick={downloadCloudToDevice}>Download cloud to this device</button>
                    <button style={btnS} onClick={mergeCloudAndDevice}>Merge cloud + this device</button>
                    <button style={{ ...btnS, color: C.maroon }} onClick={signOut}>Sign out</button>
                  </div>
                </div>
              )}
              {authMsg && <div style={{ fontSize: 12, color: C.faint, marginTop: 10 }}>{authMsg}</div>}
            </section>
          </div>
        )}

        {/* ===== ABOUT ===== */}
        {tab === "about" && (
          <div style={{ display: "grid", gap: 14 }}>
            <section style={cardS}>
              <h2 style={h2S}>Sundar Chaitanya Das</h2>
              <p style={{ fontSize: 13, color: C.faint, marginTop: -6 }}>formerly known as Sumedh Brahmadevara</p>
              <div style={{ display: "grid", gap: 12, fontSize: 14, lineHeight: 1.7 }}>
                <p style={{ margin: 0 }}>
                  Sundar Chaitanya Das, formerly known as Sumedh Brahmadevara, was born into the Kṛṣṇa conscious movement. From a young age, he was particularly intrigued by philosophy and Sanskrit ślokas: he would sit in lectures making notes and memorised ten chapters of the Bhagavad-gītā by the age of eight.
                </p>
                <p style={{ margin: 0 }}>
                  He later began serving in kīrtanas and deepened his spiritual practice through increased scriptural study of the Bhagavad-gītā, Śrī Īśopaniṣad, Bhakti-rasāmṛta-sindhu, Tattva-sandarbha and various other Gauḍīya Vaiṣṇava texts. He completed his Bhakti-śāstrī qualification at the age of fifteen.
                </p>
                <p style={{ margin: 0 }}>
                  He is a qualified Bhakti-śāstrī teacher and has delivered classes and presentations on Kṛṣṇa consciousness for groups of different ages over the years. He took initiation from Radhanath Swami at the age of eighteen. He has led the Pandava Sena Manchester youth group and studied Economics at the University of Cambridge, where he also served through KCSoc.
                </p>
              </div>
              <div style={{ display: "flex", gap: 10, flexWrap: "wrap", marginTop: 18 }}>
                <a href="https://www.instagram.com/sundarchaitanyadas/" target="_blank" rel="noopener noreferrer" style={{ ...btnS, textDecoration: "none", display: "inline-block" }}>Instagram →</a>
                <a href="https://www.youtube.com/@sundarchaitanyadas" target="_blank" rel="noopener noreferrer" style={{ ...btnS, textDecoration: "none", display: "inline-block" }}>YouTube →</a>
                <a href="https://substack.com/@sundarchaitanyadas" target="_blank" rel="noopener noreferrer" style={{ ...btnS, textDecoration: "none", display: "inline-block" }}>Substack →</a>
                <a href="mailto:sundarchaitanyadas@gmail.com" style={{ ...btnS, textDecoration: "none", display: "inline-block" }}>Email →</a>
              </div>
            </section>

            <section style={cardS}>
              <h2 style={h2S}>The importance of sadhana</h2>
              <div style={{ display: "grid", gap: 12, fontSize: 14, lineHeight: 1.7 }}>
                <p style={{ margin: 0 }}>
                  Sādhana is the daily practice that turns philosophy into realisation. Chanting, hearing, reading, worship and remembrance are not separate boxes to tick — each limb supports the others, and consistency across all of them, done a little every day, does more for the heart than intensity practised occasionally.
                </p>
                <p style={{ margin: 0 }}>
                  This app exists because what gets measured gets attended to. Tracking sadhana honestly — including the days it falls short — builds the self-awareness needed to actually improve, rather than relying on vague impressions of "how it's going."
                </p>
              </div>
            </section>
          </div>
        )}

        {/* ===== ADMIN ===== */}
        {tab === "admin" && isOwner && (
          <div style={{ display: "grid", gap: 14 }}>
            <section style={cardS}>
              <h2 style={h2S}>All users</h2>
              <p style={{ fontSize: 12, color: C.faint, marginTop: 0 }}>
                Requires the <code>admin</code> Supabase Edge Function to be deployed. Passwords are never stored or shown — you can set a new one for a locked-out user below.
              </p>
              <button style={btnPri} onClick={loadAdminUsers}>Load users</button>
              {adminMsg && <div style={{ fontSize: 12, color: C.faint, marginTop: 10 }}>{adminMsg}</div>}
              {adminUsers && (
                <div style={{ marginTop: 12, display: "grid", gap: 6 }}>
                  {adminUsers.map((u) => (
                    <button key={u.id} onClick={() => loadAdminUserData(u)}
                      style={{ ...btnS, textAlign: "left", background: adminSelected?.id === u.id ? C.saffronSoft : "#fff" }}>
                      {u.email} <span style={{ color: C.faint, fontSize: 11 }}> · last in {u.lastSignInAt ? fmtDate(u.lastSignInAt.slice(0, 10)) : "never"}</span>
                    </button>
                  ))}
                  {adminUsers.length === 0 && <Empty m="No users yet." />}
                </div>
              )}
            </section>

            {adminSelected && (
              <section style={cardS}>
                <h2 style={h2S}>{adminSelected.email}</h2>
                <div style={{ display: "grid", gap: 8, marginBottom: 14 }}>
                  <div style={{ fontSize: 13, color: C.faint }}>
                    {adminSelectedData ? `${Object.keys(adminSelectedData.days || {}).length} days logged, ${(adminSelectedData.verses || []).length} verses.` : "Loading…"}
                  </div>
                </div>
                <div style={{ display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap" }}>
                  <input type="password" placeholder="New password (min 6 chars)" value={adminNewPassword} onChange={(e) => setAdminNewPassword(e.target.value)} style={{ ...inpS, flex: 1, minWidth: 180 }} />
                  <button style={btnS} onClick={resetAdminUserPassword}>Set new password</button>
                </div>
              </section>
            )}
          </div>
        )}

        {showSignup && (
          <div style={{ position: "fixed", inset: 0, background: "rgba(39,33,64,0.55)", display: "flex", alignItems: "center", justifyContent: "center", padding: 16, zIndex: 100 }}
            onClick={(e) => { if (e.target === e.currentTarget) setShowSignup(false); }}>
            <div style={{ ...cardS, width: "100%", maxWidth: 360 }}>
              <h2 style={h2S}>Create account</h2>
              <div style={{ display: "grid", gap: 10 }}>
                <input type="email" placeholder="Email" value={signupEmail} onChange={(e) => setSignupEmail(e.target.value)} style={inpS} />
                <input type="password" placeholder="Password" value={signupPassword} onChange={(e) => setSignupPassword(e.target.value)} style={inpS} />
                <input type="password" placeholder="Confirm password" value={signupConfirm} onChange={(e) => setSignupConfirm(e.target.value)} style={inpS} />
                <div style={{ display: "flex", gap: 8 }}>
                  <button style={{ ...btnPri, flex: 1 }} onClick={createAccount}>Create account</button>
                  <button style={btnS} onClick={() => setShowSignup(false)}>Cancel</button>
                </div>
                {signupMsg && <div style={{ fontSize: 12, color: C.faint }}>{signupMsg}</div>}
              </div>
            </div>
          </div>
        )}

        {/* ===== EXPORT ===== */}
        {tab === "export" && (
          <div style={{ display: "grid", gap: 14 }}>
            <section style={cardS}>
              <h2 style={h2S}>Backup & restore (JSON)</h2>
              <p style={{ fontSize: 13, color: C.faint, marginTop: 0 }}>
                A full, lossless backup of every day, japa round, and verse — unlike the Excel exports below, this file can be loaded straight back into the app. Use it to carry all your data into a new version of the app, or to restore after a device change.
              </p>
              <div style={{ display: "grid", gap: 8 }}>
                <button style={{ ...btnPri, width: "100%" }} onClick={exportBackupJson}>Download full backup (.json) →</button>
                <button style={{ ...btnS, width: "100%" }} onClick={() => triggerImport("merge")}>Restore backup — merge with current data</button>
                <button style={{ ...btnS, width: "100%", color: C.maroon }} onClick={() => triggerImport("replace")}>Restore backup — replace current data</button>
              </div>
              <input ref={fileInputRef} type="file" accept="application/json" onChange={handleImportFile} style={{ display: "none" }} />
              {backupMsg && <div style={{ fontSize: 12, color: C.faint, marginTop: 10 }}>{backupMsg}</div>}
            </section>

            <section style={cardS}>
              <h2 style={h2S}>Excel exports</h2>
              <p style={{ fontSize: 13, color: C.faint, marginTop: 0 }}>
                Choose the data sheets and chart-ready visual packs you want, then download an Excel workbook for the current week, month, year, or all data.
              </p>
              <div style={{ display: "grid", gap: 10 }}>
                <button style={{ ...btnPri, width: "100%" }} onClick={() => exportExcel("week")}>Download selected — this week →</button>
                <button style={{ ...btnPri, width: "100%", background: C.tulsi, borderColor: C.tulsi }} onClick={() => exportExcel("month")}>Download selected — this month →</button>
                <button style={{ ...btnPri, width: "100%", background: C.maroon, borderColor: C.maroon }} onClick={() => exportExcel("year")}>Download selected — this year →</button>
                <button style={{ ...btnS, width: "100%" }} onClick={() => exportExcel("all")}>Download selected — all data →</button>
              </div>
            </section>

            <section style={cardS}>
              <h2 style={h2S}>Choose data sheets</h2>
              <div style={{ display: "grid", gap: 8 }}>
                {[
                  ["dataSummary", "Summary"], ["dataDaily", "Daily log"], ["dataJapa", "Japa rounds"],
                  ["dataHearing", "Hearing sessions"], ["dataReading", "Reading sessions"], ["dataWorship", "Worship checklist"],
                  ["dataSleep", "Night awakenings"], ["dataVerses", "Verses"], ["dataVersePractice", "Verse practice history"],
                ].map(([key, label]) => <Toggle key={key} label={label} value={!!exportOptions[key]} onChange={(v) => setExportOption(key, v)} />)}
              </div>
            </section>

            <section style={cardS}>
              <h2 style={h2S}>Choose visual packs</h2>
              <p style={{ fontSize: 12, color: C.faint, marginTop: 0 }}>
                These export chart-ready tables behind the dashboard visuals, so you can recreate or analyse the charts in Excel.
              </p>
              <div style={{ display: "grid", gap: 8 }}>
                {[
                  ["vizOverview", "Overview visual data"], ["vizJapa", "Japa visual data"], ["vizHearingReading", "Hearing/reading visual data"],
                  ["vizSleep", "Sleep visual data"], ["vizWorship", "Worship visual data"], ["vizComposite", "Composite score visual data"],
                ].map(([key, label]) => <Toggle key={key} label={label} value={!!exportOptions[key]} onChange={(v) => setExportOption(key, v)} />)}
              </div>
              <div style={{ display: "flex", gap: 8, flexWrap: "wrap", marginTop: 12 }}>
                <button style={btnS} onClick={() => setExportOptions(Object.fromEntries(Object.keys(exportOptions).map((k) => [k, true])))}>Select all</button>
                <button style={btnS} onClick={() => setExportOptions({ ...exportOptions, vizOverview: true, vizJapa: false, vizHearingReading: false, vizSleep: true, vizWorship: false, vizComposite: false })}>Core visuals</button>
                <button style={btnS} onClick={() => setExportOptions(Object.fromEntries(Object.keys(exportOptions).map((k) => [k, false])))}>Clear all</button>
              </div>
            </section>

            <section style={cardS}>
              <h2 style={h2S}>Export notes</h2>
              <div style={{ fontSize: 13, color: C.ink, lineHeight: 1.6 }}>
                Weekly exports use Monday–Sunday. Monthly and yearly exports use the current calendar month/year.<br />
                Verse practice now separates <b>revision</b> from <b>recitation</b>, while the daily verse score counts either as verse practice.
              </div>
            </section>
          </div>
        )}

        {/* ===== VISUALS ===== */}
        {tab === "viz" && (
          <div>
            <div style={{ display: "flex", gap: 6, marginBottom: 14, flexWrap: "wrap" }}>
              {VIZ_TABS.map(([k, l]) => (
                <button key={k} onClick={() => setVizTab(k)} style={{
                  padding: "6px 12px", borderRadius: 999, fontSize: 12, cursor: "pointer", fontWeight: 600,
                  border: `1.5px solid ${vizTab === k ? C.saffron : C.line}`,
                  background: vizTab === k ? C.saffron : C.card, color: vizTab === k ? "#fff" : C.ink,
                }}>{l}</button>
              ))}
            </div>
            {vizTab === "jdur" && <JapaDurationViz A={A} />}
            {vizTab === "jtime" && <JapaTimeViz A={A} />}
            {vizTab === "hear" && <HearingViz A={A} />}
            {vizTab === "read" && <ReadingViz A={A} />}
            {vizTab === "wor" && <WorshipViz A={A} />}
            {vizTab === "ver" && <VersesViz A={A} verses={verses} />}
            {vizTab === "sleep" && <SleepViz A={A} />}
            {vizTab === "comp" && <CompositeViz A={A} verses={verses} />}
          </div>
        )}

      </div>
    </div>
  );
}
