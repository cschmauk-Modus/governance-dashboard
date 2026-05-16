// @ts-nocheck
"use client";"use client";
import React, { useMemo, useState } from "react";
import * as XLSX from "xlsx";
import {
  ResponsiveContainer,
  RadarChart,
  PolarGrid,
  PolarAngleAxis,
  PolarRadiusAxis,
  Radar,
  BarChart,
  Bar,
  XAxis,
  YAxis,
  Tooltip,
  ScatterChart,
  Scatter,
  ZAxis,
  Cell,
  AreaChart,
  Area,
} from "recharts";
import { motion } from "framer-motion";
import { Upload, ShieldCheck, Activity, AlertTriangle, Target, Building2, Cpu, Gauge, Layers3 } from "lucide-react";

const TEAL = "#0D9488";
const DARK = "#07111f";
const CARD = "rgba(255,255,255,0.92)";
const INK = "#0f172a";
const MUTED = "#64748b";
const YELLOW = "#F59E0B";
const RED = "#DC2626";
const GREEN = "#10B981";
const BLUE = "#2563EB";

const demoWorkbook = {
  Empresa: [
    { Campo: "Empresa", Valor: "Cenco Malls" },
    { Campo: "Industria", Valor: "Retail Inmobiliario / Centros Comerciales" },
    { Campo: "Países", Valor: "Chile, Perú y Colombia" },
    { Campo: "Año análisis", Valor: "2025" },
    { Campo: "Propósito declarado", Valor: "Liderar la generación de espacios y experiencias para un mundo mejor" },
    { Campo: "SAI global", Valor: 3.6 },
    { Campo: "Madurez organizacional", Valor: "Medio-Alto" },
    { Campo: "Juicio ejecutivo", Valor: "Propósito instalado con buena coherencia discursiva, pero con brechas de institucionalización y accountability." },
  ],
  Scores_SAI: [
    { Dimensión: "Propósito y Dirección Estratégica", Score: 4.1, Estado: "🟢", "Lectura Ejecutiva": "Propósito claro y alineado al negocio" },
    { Dimensión: "Gobierno Corporativo y Accountability", Score: 3.2, Estado: "🟡", "Lectura Ejecutiva": "Gobierno formal robusto pero oversight parcial" },
    { Dimensión: "Consistencia Decisional", Score: 4.0, Estado: "🟢", "Lectura Ejecutiva": "Decisiones coherentes con narrativa estratégica" },
    { Dimensión: "ESG y Sostenibilidad Integrada", Score: 3.6, Estado: "🟡", "Lectura Ejecutiva": "Integración operativa sólida, no sistémica" },
    { Dimensión: "Stakeholders y Legitimidad", Score: 3.6, Estado: "🟡", "Lectura Ejecutiva": "Madurez media-alta con brechas territoriales" },
    { Dimensión: "Cultura y Activación", Score: 3.3, Estado: "🟡", "Lectura Ejecutiva": "Activación en progreso" },
    { Dimensión: "Riesgo y Resiliencia", Score: 3.2, Estado: "🟡", "Lectura Ejecutiva": "Requiere mayor trazabilidad" },
  ],
  "Indicadores IAP": [
    { COMPANY: "CENCO", NODO: "IAP TOTAL", LABEL: "Activación Ecosistema Completo", DESCRIPCION: "Activación de Propósito General Score", Score: 3.34, "Descripcion Score": "Activación en progreso", Sugerencias: "Potenciar SROI real y trade off de experiencia de stakeholders." },
    { COMPANY: "CENCO", NODO: "SIP TOTAL", LABEL: "Activación por actividades por nodos completa", DESCRIPCION: "Propósito Integrado en acción score", Score: 3.03, "Descripcion Score": "Activación en progreso", Sugerencias: "Revisar coherencia de ejecución, proyectos y ecosistema." },
    { COMPANY: "CENCO", NODO: "SAI TOTAL", LABEL: "Coherencia Discursiva Completa", DESCRIPCION: "Realidad Discursiva Score", Score: 3.5, "Descripcion Score": "Operativo / Integrado", Sugerencias: "Afinar relación discursiva en documentos y comunicaciones formales." },
    { COMPANY: "CENCO", NODO: "IAP PLANETA", LABEL: "IAP Planeta", DESCRIPCION: "Activación Propósito Planeta", Score: 3.22, "Descripcion Score": "Activación en progreso", Sugerencias: "Gestión ambiental buena, resiliencia incompleta." },
    { COMPANY: "CENCO", NODO: "IAP PERSONAS", LABEL: "IAP Personas", DESCRIPCION: "Activación Personas", Score: 3.18, "Descripcion Score": "Activación en progreso", Sugerencias: "Homologar cultura e incentivos." },
    { COMPANY: "CENCO", NODO: "IAP ECOSISTEMA", LABEL: "IAP Ecosistema", DESCRIPCION: "Activación Ecosistema", Score: 3.12, "Descripcion Score": "Activación en progreso", Sugerencias: "Mejorar legitimidad territorial." },
  ],
  "ESG Radar": [
    { COMPANY: "CENCO", ID: "PAIESG-000", Category: "ESG", Gobernance: 4.0, Social: 3.2, Environment: 4.0, Comentarios: "Excelente desarrollo de ESG" },
    { COMPANY: "CENCO", ID: "PAIESG-001", Category: "IAP", Gobernance: 3.11, Social: 3.09, Environment: 3.32, Comentarios: "Apostar en apalancamiento de activación de propósito." },
  ],
  Riesgos: [
    { Riesgo: "Purpose washing", Intensidad: "Alta", Impacto: "Alto", Comentario: "Narrativa podría avanzar más rápido que institucionalización" },
    { Riesgo: "Déficit accountability ESG", Intensidad: "Alta", Impacto: "Alto", Comentario: "Falta trazabilidad e integración" },
    { Riesgo: "Legitimidad territorial", Intensidad: "Alta", Impacto: "Alto", Comentario: "Expansión urbana genera tensiones" },
    { Riesgo: "Riesgo climático", Intensidad: "Alta", Impacto: "Alto", Comentario: "Portafolio expuesto a estrés hídrico e inundaciones" },
    { Riesgo: "Gobernanza IA y privacidad", Intensidad: "Media-Alta", Impacto: "Alto", Comentario: "Gobernanza IA aún insuficiente" },
  ],
  Tensiones: [
    { Tensión: "Rentabilidad vs sostenibilidad", Intensidad: "Alta", Impacto: "Alto", Comentario: "Internalización parcial de externalidades" },
    { Tensión: "Crecimiento vs comunidad", Intensidad: "Alta", Impacto: "Alto", Comentario: "Riesgo territorial y reputacional" },
    { Tensión: "Experiencia cliente vs privacidad", Intensidad: "Alta", Impacto: "Alto", Comentario: "Digitalización aumenta exposición" },
    { Tensión: "Legado fundador vs institucionalización", Intensidad: "Media-Alta", Impacto: "Alto", Comentario: "Dependencia cultural histórica" },
    { Tensión: "Velocidad innovación vs gobernanza", Intensidad: "Alta", Impacto: "Alto", Comentario: "IA y datos avanzan más rápido que gobierno" },
  ],
  Prioridades: [
    { "Prioridad Estratégica": "Institucionalizar curaduría estratégica del ecosistema", Horizonte: "12-24 meses", "Impacto Esperado": "Alto" },
    { "Prioridad Estratégica": "Instalar accountability directorio sobre propósito y legitimidad", Horizonte: "6-12 meses", "Impacto Esperado": "Alto" },
    { "Prioridad Estratégica": "Convertir ESG en filtro de decisión", Horizonte: "12 meses", "Impacto Esperado": "Alto" },
    { "Prioridad Estratégica": "Formalizar gobernanza IA y datos", Horizonte: "6-12 meses", "Impacto Esperado": "Medio-Alto" },
    { "Prioridad Estratégica": "Fortalecer resiliencia climática y territorial", Horizonte: "12-24 meses", "Impacto Esperado": "Alto" },
  ],
  Nodos: [
    { Nodo: "Personas", "Fortaleza Principal": "Experiencia y seguridad", "Brecha Principal": "Incentivos y cultura homogénea", Nivel: "Medio-Alto" },
    { Nodo: "Gobierno Corporativo", "Fortaleza Principal": "Estructura formal sólida", "Brecha Principal": "Oversight ecosistema", Nivel: "Medio" },
    { Nodo: "Planeta", "Fortaleza Principal": "Medición y gestión ambiental", "Brecha Principal": "Resiliencia climática", Nivel: "Medio-Alto" },
    { Nodo: "Propiedad", "Fortaleza Principal": "Transformación activos y experiencia", "Brecha Principal": "Curaduría gobernada", Nivel: "Alto" },
    { Nodo: "Ecosistema", "Fortaleza Principal": "Relación stakeholders", "Brecha Principal": "Legitimidad territorial", Nivel: "Medio" },
  ],
};

function valueOf(rows: any[], key: string, fallback: any = "—") {
  const found = rows?.find((r) => String(r.Campo || "").trim().toLowerCase() === key.toLowerCase());
  return found?.Valor ?? fallback;
}

function n(v: any, fallback: number = 0) {
  const x = Number(String(v ?? "").replace(",", "."));
  return Number.isFinite(x) ? x : fallback;
}

function levelToScore(level: any) {
  const l = String(level || "").toLowerCase();
  if (l.includes("alto") && !l.includes("medio")) return 4.3;
  if (l.includes("medio-alto")) return 3.6;
  if (l.includes("medio")) return 3.0;
  if (l.includes("bajo")) return 2.0;
  return 3.2;
}

function colorByScore(score: number) {
  if (score >= 3.8) return GREEN;
  if (score >= 3.0) return YELLOW;
  return RED;
}

function riskNumeric(v: any) {
  const s = String(v || "").toLowerCase();
  if (s.includes("muy")) return 5;
  if (s.includes("alta") || s.includes("alto")) return 4.4;
  if (s.includes("media-alta") || s.includes("medio-alto")) return 3.7;
  if (s.includes("media") || s.includes("medio")) return 3;
  if (s.includes("baja") || s.includes("bajo")) return 2;
  return 3.2;
}

function normalizeWorkbook(raw: any) {
  const empresaRows = raw.Empresa || [];
  const scores = raw.Scores_SAI || raw["Scores SAI"] || [];
  const iap = raw["Indicadores IAP"] || raw.Indicadores_IAP || [];
  const esg = raw["ESG Radar"] || raw.ESG_Radar || [];
  const riesgos = raw.Riesgos || [];
  const tensiones = raw.Tensiones || [];
  const prioridades = raw.Prioridades || [];
  const nodos = raw.Nodos || [];

  const saiFromScores = scores.length ? scores.reduce((a, r) => a + n(r.Score), 0) / scores.length : 0;
  const saiGlobal = n(valueOf(empresaRows, "SAI global", saiFromScores), saiFromScores);
  const iapTotal = iap.find((r) => String(r.NODO || r.Dimensión || "").toUpperCase().includes("IAP TOTAL"));
  const iapScore = n(iapTotal?.Score ?? iapTotal?.score, iap.length ? iap.reduce((a, r) => a + n(r.Score ?? r.score), 0) / iap.length : 3);
  const madurez = valueOf(empresaRows, "Madurez organizacional", "Medio-Alto");

  const activation = Math.min(5, Math.max(0, (iapScore * 0.7 + levelToScore(madurez) * 0.3)));
  const coherence = Math.min(5, Math.max(0, saiGlobal));

  const scorecard = [
    { label: "Propósito", score: scores.find((r) => String(r.Dimensión).toLowerCase().includes("propósito"))?.Score ?? saiGlobal },
    { label: "Gobierno", score: scores.find((r) => String(r.Dimensión).toLowerCase().includes("gobierno"))?.Score ?? 3.2 },
    { label: "ESG", score: scores.find((r) => String(r.Dimensión).toLowerCase().includes("esg"))?.Score ?? 3.5 },
    { label: "Stakeholders", score: scores.find((r) => String(r.Dimensión).toLowerCase().includes("stake"))?.Score ?? 3.3 },
    { label: "Cultura", score: scores.find((r) => String(r.Dimensión).toLowerCase().includes("cultura"))?.Score ?? activation },
    { label: "Riesgo", score: scores.find((r) => String(r.Dimensión).toLowerCase().includes("riesgo"))?.Score ?? 3.1 },
    { label: "Consistencia", score: scores.find((r) => String(r.Dimensión).toLowerCase().includes("consistencia"))?.Score ?? coherence },
  ];

  const radar = ["Gobernance", "Social", "Environment"].map((axis) => {
    const current = esg.find((r) => String(r.Category || "").toUpperCase() === "IAP") || esg[0] || {};
    const bench = esg.find((r) => String(r.Category || "").toUpperCase() === "ESG") || esg[1] || {};
    return {
      dimension: axis === "Gobernance" ? "Governance" : axis,
      actual: n(current[axis]),
      benchmark: n(bench[axis]),
      tendencia: n(current[axis]) >= n(bench[axis]) ? "↗" : "→",
    };
  });

  return {
    empresa: {
      nombre: valueOf(empresaRows, "Empresa", "Empresa"),
      industria: valueOf(empresaRows, "Industria", "Industria"),
      paises: valueOf(empresaRows, "Países", "—"),
      fecha: valueOf(empresaRows, "Año análisis", new Date().getFullYear()),
      proposito: valueOf(empresaRows, "Propósito declarado", "Propósito no informado"),
      madurez,
      juicio: valueOf(empresaRows, "Juicio ejecutivo", "Lectura ejecutiva pendiente de completar."),
      saiGlobal,
      iapScore,
      coherence,
      activation,
    },
    scores,
    iap,
    radar,
    riesgos,
    tensiones,
    prioridades,
    nodos,
    scorecard,
  };
}

function Panel({ title, icon: Icon, children, className = "" }) {
  return (
    <motion.section
      initial={{ opacity: 0, y: 16 }}
      animate={{ opacity: 1, y: 0 }}
      transition={{ duration: 0.45 }}
      className={`rounded-3xl border border-slate-200/80 bg-white/90 p-4 shadow-[0_20px_60px_rgba(2,8,23,0.10)] backdrop-blur ${className}`}
    >
      {title && (
        <div className="mb-3 flex items-center justify-between gap-3">
          <div className="flex items-center gap-2">
            {Icon && <Icon className="h-4 w-4 text-teal-600" />}
            <h3 className="text-xs font-bold uppercase tracking-[0.18em] text-slate-600">{title}</h3>
          </div>
          <div className="h-1.5 w-1.5 rounded-full bg-teal-500 shadow-[0_0_14px_#0D9488]" />
        </div>
      )}
      {children}
    </motion.section>
  );
}

function GaugePanel({ score, maturity, estado }) {
  const pct = Math.max(0, Math.min(100, (score / 5) * 100));
  const color = colorByScore(score);
  return (
    <Panel title="Gauge SAI Global" icon={Gauge}>
      <div className="flex flex-col items-center">
        <div className="relative grid h-44 w-44 place-items-center rounded-full" style={{ background: `conic-gradient(${color} ${pct}%, #e2e8f0 ${pct}% 100%)`, boxShadow: `0 0 32px ${color}55` }}>
          <div className="grid h-32 w-32 place-items-center rounded-full bg-white shadow-inner">
            <div className="text-center">
              <div className="text-4xl font-black text-slate-900">{score.toFixed(2)}</div>
              <div className="text-[10px] font-bold uppercase tracking-widest text-slate-500">/ 5.00 SAI</div>
            </div>
          </div>
        </div>
        <div className="mt-4 w-full rounded-2xl bg-slate-50 p-3 text-center">
          <div className="text-sm font-bold text-slate-900">{maturity}</div>
          <div className="text-xs text-slate-500">{estado || "Madurez ejecutiva"}</div>
        </div>
      </div>
    </Panel>
  );
}

function ActivationMatrix({ empresa }) {
  const x = (empresa.coherence / 5) * 100;
  const y = 100 - (empresa.activation / 5) * 100;
  const quadrant = empresa.coherence >= 3 && empresa.activation >= 3 ? "Propósito Integrado" : empresa.coherence >= 3 ? "Narrativa sin Activación" : empresa.activation >= 3 ? "Acción Desalineada" : "Propósito Irrelevante";
  return (
    <Panel title="Matriz de Activación del Propósito" icon={Target} className="min-h-[520px]">
      <div className="mb-4 rounded-2xl bg-slate-900 p-4 text-white shadow-[0_0_40px_rgba(13,148,136,0.18)]">
        <div className="flex flex-wrap items-center justify-between gap-3">
          <div>
            <div className="text-xs font-bold uppercase tracking-[0.22em] text-teal-300">SAI = Semantic Alignment Intelligence</div>
            <div className="mt-1 text-sm text-slate-200">Modelo de evaluación de integración estratégica del propósito y coherencia organizacional.</div>
          </div>
          <div className="rounded-full bg-teal-400/10 px-4 py-2 text-sm font-bold text-teal-200 ring-1 ring-teal-300/30">{quadrant}</div>
        </div>
      </div>
      <div className="relative h-[390px] overflow-hidden rounded-3xl border border-slate-200 bg-[radial-gradient(circle_at_50%_45%,rgba(13,148,136,0.16),transparent_35%),linear-gradient(135deg,#ffffff,#f8fafc)] p-5">
        <div className="absolute left-1/2 top-0 h-full w-px bg-slate-300" />
        <div className="absolute left-0 top-1/2 h-px w-full bg-slate-300" />
        <div className="absolute left-6 top-6 rounded-2xl bg-emerald-50 px-3 py-2 text-xs font-bold text-emerald-700">Propósito Integrado</div>
        <div className="absolute right-6 top-6 rounded-2xl bg-amber-50 px-3 py-2 text-xs font-bold text-amber-700">Narrativa sin Activación</div>
        <div className="absolute left-6 bottom-10 rounded-2xl bg-blue-50 px-3 py-2 text-xs font-bold text-blue-700">Acción Desalineada</div>
        <div className="absolute right-6 bottom-10 rounded-2xl bg-rose-50 px-3 py-2 text-xs font-bold text-rose-700">Propósito Irrelevante</div>
        <div className="absolute bottom-3 left-1/2 -translate-x-1/2 text-[11px] font-bold uppercase tracking-widest text-slate-500">Coherencia Estratégica →</div>
        <div className="absolute left-2 top-1/2 -translate-y-1/2 -rotate-90 text-[11px] font-bold uppercase tracking-widest text-slate-500">Nivel de Activación →</div>
        <motion.div
          initial={{ scale: 0.65, opacity: 0 }}
          animate={{ scale: 1, opacity: 1, left: `${x}%`, top: `${y}%` }}
          transition={{ type: "spring", stiffness: 90, damping: 14 }}
          className="absolute z-10 -translate-x-1/2 -translate-y-1/2"
        >
          <div className="relative">
            <div className="absolute -inset-6 rounded-full bg-teal-400/30 blur-xl" />
            <div className="relative h-7 w-7 rounded-full border-4 border-white bg-teal-500 shadow-[0_0_28px_#0D9488]" />
            <div className="absolute left-8 top-1/2 w-60 -translate-y-1/2 rounded-2xl bg-slate-900/95 p-3 text-white shadow-2xl">
              <div className="text-sm font-black">{empresa.nombre}</div>
              <div className="text-xs text-slate-300">SAI {empresa.coherence.toFixed(2)} · IAP {empresa.activation.toFixed(2)}</div>
              <div className="mt-1 text-[11px] text-teal-200">{empresa.juicio}</div>
            </div>
          </div>
        </motion.div>
      </div>
    </Panel>
  );
}

function FileLoader({ onLoad }) {
  async function handleFile(e) {
    const file = e.target.files?.[0];
    if (!file) return;
    const data = await file.arrayBuffer();
    const wb = XLSX.read(data, { type: "array" });
    const parsed = {};
    wb.SheetNames.forEach((name) => {
      parsed[name] = XLSX.utils.sheet_to_json(wb.Sheets[name], { defval: "" });
    });
    onLoad(parsed);
  }
  return (
    <label className="group flex cursor-pointer items-center gap-2 rounded-2xl border border-white/20 bg-white/10 px-4 py-2 text-sm font-semibold text-white backdrop-blur transition hover:bg-white/20">
      <Upload className="h-4 w-4" /> Cargar Excel
      <input type="file" accept=".xlsx,.xls" className="hidden" onChange={handleFile} />
    </label>
  );
}

export default function GovernanceIntelligenceDashboard() {
  const [raw, setRaw] = useState(demoWorkbook);
  const data = useMemo(() => normalizeWorkbook(raw), [raw]);

  const saiBars = data.scores.map((r) => ({ name: r.Dimensión || r.dimension, score: n(r.Score ?? r.score), estado: r.Estado || r.estado }));
  const riskData = data.riesgos.map((r, idx) => ({ name: r.Riesgo || r.riesgo, impacto: riskNumeric(r.Impacto), probabilidad: riskNumeric(r.Probabilidad || r.Intensidad), intensidad: riskNumeric(r.Intensidad), comentario: r.Comentario, idx }));
  const iapCards = data.iap.slice(0, 8).map((r) => ({ label: r.LABEL || r.KPI || r.Dimensión || r.NODO, score: n(r.Score ?? r.score), desc: r.DESCRIPCION || r["Descripcion Score"] || r.Sugerencias || "" }));

  return (
    <div className="min-h-screen bg-[#F1F5F9] p-4 text-slate-900 md:p-6">
      <div className="mx-auto max-w-[1800px] overflow-hidden rounded-[2rem] border border-slate-200 bg-[#F1F5F9] shadow-2xl">
        <header className="relative overflow-hidden bg-slate-950 px-6 py-5 text-white">
          <div className="absolute inset-0 bg-[radial-gradient(circle_at_15%_0%,rgba(13,148,136,0.32),transparent_26%),radial-gradient(circle_at_85%_20%,rgba(37,99,235,0.25),transparent_30%)]" />
          <div className="relative flex flex-wrap items-center justify-between gap-4">
            <div>
              <div className="flex items-center gap-3">
                <div className="grid h-11 w-11 place-items-center rounded-2xl bg-teal-500 shadow-[0_0_28px_rgba(13,148,136,0.75)]"><Layers3 className="h-6 w-6" /></div>
                <div>
                  <h1 className="text-2xl font-black tracking-tight md:text-3xl">Strategic Governance Intelligence</h1>
                  <p className="text-sm text-slate-300">Control tower ejecutivo · Purpose Governance · ESG/IAP Oversight</p>
                </div>
              </div>
            </div>
            <FileLoader onLoad={setRaw} />
          </div>
        </header>

        <main className="grid grid-cols-1 gap-4 p-4 xl:grid-cols-[330px_minmax(620px,1fr)_360px]">
          <aside className="space-y-4">
            <Panel title="Empresa" icon={Building2}>
              <div className="rounded-3xl bg-slate-950 p-4 text-white shadow-[0_0_32px_rgba(13,148,136,0.16)]">
                <div className="text-2xl font-black">{data.empresa.nombre}</div>
                <div className="mt-1 text-xs text-slate-300">{data.empresa.industria}</div>
                <div className="mt-4 grid grid-cols-2 gap-2 text-xs">
                  <div className="rounded-2xl bg-white/10 p-3"><b>Países</b><br />{data.empresa.paises}</div>
                  <div className="rounded-2xl bg-white/10 p-3"><b>Fecha</b><br />{data.empresa.fecha}</div>
                </div>
                <div className="mt-3 rounded-2xl border border-teal-300/20 bg-teal-400/10 p-3 text-sm font-bold text-teal-100">Madurez: {data.empresa.madurez}</div>
              </div>
            </Panel>

            <Panel title="Propósito" icon={ShieldCheck}>
              <p className="text-lg font-semibold leading-snug text-slate-900">“{data.empresa.proposito}”</p>
              <p className="mt-3 text-xs leading-relaxed text-slate-500">Lectura directiva: evaluar si el propósito gobierna decisiones, trade-offs, asignación de capital, legitimidad territorial y accountability.</p>
            </Panel>

            <Panel title="Score por dimensión SAI" icon={Activity}>
              <div className="h-64">
                <ResponsiveContainer width="100%" height="100%">
                  <BarChart data={saiBars} layout="vertical" margin={{ left: 8, right: 20, top: 4, bottom: 4 }}>
                    <XAxis type="number" domain={[0, 5]} hide />
                    <YAxis type="category" dataKey="name" width={150} tick={{ fontSize: 10, fill: MUTED }} />
                    <Tooltip formatter={(v) => Number(v).toFixed(2)} />
                    <Bar dataKey="score" radius={[0, 10, 10, 0]} barSize={14}>
                      {saiBars.map((entry, index) => <Cell key={index} fill={colorByScore(entry.score)} />)}
                    </Bar>
                  </BarChart>
                </ResponsiveContainer>
              </div>
            </Panel>
          </aside>

          <section className="space-y-4">
            <ActivationMatrix empresa={data.empresa} />

            <div className="grid grid-cols-1 gap-4 2xl:grid-cols-2">
              <Panel title="Brechas críticas" icon={AlertTriangle}>
                <div className="grid gap-2">
                  {["Narrativa vs evidencia", "ESG vs gobernanza", "Crecimiento vs resiliencia", "Accountability", "Legitimidad"].map((b, i) => {
                    const risk = Math.min(5, 2.8 + i * 0.35 + (5 - data.empresa.activation) * 0.22);
                    return (
                      <div key={b} className="grid grid-cols-[1fr_90px] items-center gap-3 rounded-2xl bg-slate-50 p-3">
                        <div className="text-sm font-semibold text-slate-800">{b}</div>
                        <div className="h-2 overflow-hidden rounded-full bg-slate-200">
                          <div className="h-full rounded-full" style={{ width: `${(risk / 5) * 100}%`, background: colorByScore(5 - risk + 2) }} />
                        </div>
                      </div>
                    );
                  })}
                </div>
              </Panel>

              <Panel title="Radar ESG / IAP" icon={Cpu}>
                <div className="h-72">
                  <ResponsiveContainer width="100%" height="100%">
                    <RadarChart data={data.radar} outerRadius="72%">
                      <PolarGrid />
                      <PolarAngleAxis dataKey="dimension" tick={{ fontSize: 12, fill: INK }} />
                      <PolarRadiusAxis angle={90} domain={[0, 5]} tick={{ fontSize: 10 }} />
                      <Radar name="Score actual" dataKey="actual" stroke={TEAL} fill={TEAL} fillOpacity={0.28} strokeWidth={3} />
                      <Radar name="Benchmark" dataKey="benchmark" stroke={BLUE} fill={BLUE} fillOpacity={0.12} strokeWidth={2} />
                      <Tooltip />
                    </RadarChart>
                  </ResponsiveContainer>
                </div>
              </Panel>
            </div>

            <div className="grid grid-cols-1 gap-4 2xl:grid-cols-2">
              <Panel title="Mapa de riesgos" icon={AlertTriangle}>
                <div className="h-72">
                  <ResponsiveContainer width="100%" height="100%">
                    <ScatterChart margin={{ top: 20, right: 20, bottom: 18, left: 8 }}>
                      <XAxis type="number" dataKey="probabilidad" name="Probabilidad" domain={[0, 5]} tick={{ fontSize: 10 }} />
                      <YAxis type="number" dataKey="impacto" name="Impacto" domain={[0, 5]} tick={{ fontSize: 10 }} />
                      <ZAxis type="number" dataKey="intensidad" range={[110, 700]} />
                      <Tooltip cursor={{ strokeDasharray: "3 3" }} formatter={(v) => Number(v).toFixed(1)} content={({ active, payload }) => active && payload?.length ? <div className="rounded-2xl bg-slate-950 p-3 text-xs text-white shadow-xl"><b>{payload[0].payload.name}</b><br />{payload[0].payload.comentario}</div> : null} />
                      <Scatter data={riskData}>
                        {riskData.map((entry, index) => <Cell key={index} fill={entry.intensidad >= 4 ? RED : YELLOW} />)}
                      </Scatter>
                    </ScatterChart>
                  </ResponsiveContainer>
                </div>
              </Panel>

              <Panel title="KPI IAP" icon={Activity}>
                <div className="grid grid-cols-2 gap-3">
                  {iapCards.map((k, i) => (
                    <div key={`${k.label}-${i}`} className="rounded-2xl border border-slate-100 bg-slate-50 p-3 transition hover:-translate-y-0.5 hover:shadow-lg">
                      <div className="flex items-center justify-between gap-2">
                        <div className="line-clamp-1 text-xs font-bold text-slate-700">{k.label}</div>
                        <div className="rounded-full px-2 py-0.5 text-[10px] font-bold text-white" style={{ background: colorByScore(k.score) }}>{k.score.toFixed(2)}</div>
                      </div>
                      <div className="mt-2 h-1.5 overflow-hidden rounded-full bg-slate-200"><div className="h-full rounded-full" style={{ width: `${(k.score / 5) * 100}%`, background: colorByScore(k.score) }} /></div>
                      <div className="mt-2 line-clamp-2 text-[10px] leading-snug text-slate-500">{k.desc}</div>
                    </div>
                  ))}
                </div>
              </Panel>
            </div>
          </section>

          <aside className="space-y-4">
            <GaugePanel score={data.empresa.saiGlobal} maturity={data.empresa.madurez} estado={data.empresa.juicio} />

            <Panel title="Scorecard ejecutivo" icon={ShieldCheck}>
              <div className="space-y-2">
                {data.scorecard.map((s) => (
                  <div key={s.label} className="rounded-2xl bg-slate-50 p-3">
                    <div className="mb-1 flex justify-between text-xs font-bold"><span>{s.label}</span><span style={{ color: colorByScore(n(s.score)) }}>{n(s.score).toFixed(2)}</span></div>
                    <div className="h-2 overflow-hidden rounded-full bg-slate-200"><div className="h-full rounded-full" style={{ width: `${(n(s.score) / 5) * 100}%`, background: colorByScore(n(s.score)) }} /></div>
                  </div>
                ))}
              </div>
            </Panel>

            <Panel title="Prioridades estratégicas" icon={Target}>
              <div className="space-y-2">
                {data.prioridades.slice(0, 5).map((p, i) => (
                  <div key={i} className="rounded-2xl border border-slate-100 bg-white p-3 shadow-sm">
                    <div className="flex items-start gap-3">
                      <div className="grid h-7 w-7 shrink-0 place-items-center rounded-full bg-slate-950 text-xs font-black text-white">{i + 1}</div>
                      <div>
                        <div className="text-sm font-bold text-slate-900">{p["Prioridad Estratégica"] || p.Prioridad || p.prioridad}</div>
                        <div className="mt-1 flex gap-2 text-[10px] font-bold uppercase tracking-wide text-slate-500">
                          <span>{p.Horizonte || p.horizonte}</span>
                          <span className="text-teal-600">{p["Impacto Esperado"] || p["impacto esperado"] || p.Impacto}</span>
                        </div>
                      </div>
                    </div>
                  </div>
                ))}
              </div>
            </Panel>

            <Panel title="Nodos de institucionalización" icon={Layers3}>
              <div className="space-y-2">
                {data.nodos.map((nodo, i) => (
                  <div key={i} className="rounded-2xl bg-slate-50 p-3">
                    <div className="flex justify-between gap-2"><b className="text-sm">{nodo.Nodo || nodo.nodo}</b><span className="text-xs font-bold text-teal-600">{nodo.Nivel || nodo["nivel de madurez"]}</span></div>
                    <div className="mt-1 text-[11px] text-slate-500"><b>Fortaleza:</b> {nodo["Fortaleza Principal"] || nodo.fortalezas}</div>
                    <div className="text-[11px] text-slate-500"><b>Brecha:</b> {nodo["Brecha Principal"] || nodo.brechas}</div>
                  </div>
                ))}
              </div>
            </Panel>
          </aside>
        </main>
      </div>
    </div>
  );
}
