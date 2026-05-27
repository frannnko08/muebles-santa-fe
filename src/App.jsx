import { useState, useEffect, useCallback } from "react";
import { supabase } from '../lib/supabase'
import { obtenerCotizaciones, crearCotizacion, aceptarCotizacion, obtenerNotasVenta, editarNumeroNotaVenta } from '../lib/cotizaciones'

const COLORS = {
  bg: "#0f0e0c", surface: "#1a1916", card: "#222018", border: "#2e2b24",
  accent: "#c8a45a", accentDim: "#8a6e38", success: "#5a9e6f",
  danger: "#c05a5a", text: "#f0ead8", muted: "#8a8270", subtle: "#3a3628",
  warning: "#c8943a",
};


const INITIAL_COTIZACIONES = [
  {id:1,numero:"11127",cliente:"DEMOVI",fecha:"2026-05-11",total:109179},
  {id:2,numero:"12113",cliente:"(sin nombre)",fecha:"2026-05-04",total:65000},
  {id:3,numero:"12114",cliente:"QCLASS S.A.",fecha:"2026-05-04",total:2130089},
  {id:4,numero:"12115",cliente:"LUIS NUÑEZ",fecha:"2026-05-05",total:423348},
  {id:5,numero:"12116",cliente:"SUR-SERVI SPA",fecha:"2026-05-05",total:234161},
  {id:6,numero:"12117",cliente:"MUEBLES ASENJO",fecha:"2026-05-06",total:1713600},
  {id:7,numero:"12118",cliente:"HECTOR VALENZUELA",fecha:"2026-05-06",total:1463461},
  {id:8,numero:"12119",cliente:"MUEBLES ASENJO",fecha:"2026-05-07",total:1590827},
  {id:9,numero:"12120",cliente:"MAXIMILIANO MORALES",fecha:"2026-05-07",total:316310},
  {id:10,numero:"12121",cliente:"OSVALDO IRIBARREN",fecha:"2026-05-08",total:110688},
  {id:11,numero:"12122",cliente:"MUEBLES ASENJO",fecha:"2026-05-11",total:237931},
  {id:12,numero:"12123",cliente:"GUILLERMO MUÑOZ",fecha:"2026-05-11",total:166564},
  {id:13,numero:"12124",cliente:"(sin nombre)",fecha:"2026-05-11",total:205269},
  {id:14,numero:"12125",cliente:"(sin nombre)",fecha:"2026-05-11",total:134583},
  {id:15,numero:"12126",cliente:"GUSTAVO",fecha:"2026-05-11",total:140242},
  {id:16,numero:"12127",cliente:"MUEBLES ASENJO",fecha:"2026-05-11",total:1280547},
  {id:17,numero:"12128",cliente:"DIEGO TOLEDO",fecha:"2026-05-11",total:275756},
  {id:18,numero:"12129",cliente:"JEOVANY GARRIDO",fecha:"2026-05-12",total:240166},
  {id:19,numero:"12129b",cliente:"ROSA MARIA VASQUEZ",fecha:"2026-05-12",total:199356},
  {id:20,numero:"12131",cliente:"JORGE BORBARAN",fecha:"2026-05-12",total:98341},
  {id:21,numero:"12131b",cliente:"DECOMADERA",fecha:"2026-05-12",total:26355},
  {id:22,numero:"12133",cliente:"VICTOR CABRERA",fecha:"2026-05-13",total:478004},
  {id:23,numero:"12134",cliente:"JOSE MARIO",fecha:"2026-05-14",total:148599},
  {id:24,numero:"12135",cliente:"VALE EVANS",fecha:"2026-05-14",total:59067},
  {id:25,numero:"12137",cliente:"NICOLAS SEPULVEDA",fecha:"2026-05-14",total:93002},
  {id:26,numero:"12138",cliente:"TARIM",fecha:"2026-05-14",total:975414},
  {id:27,numero:"12138b",cliente:"MIRIAM ACUÑA",fecha:"2026-05-14",total:253964},
  {id:28,numero:"12140",cliente:"JORGE BORBARAN",fecha:"2026-05-14",total:128554},
  {id:29,numero:"12141",cliente:"GERO MATTE",fecha:"2026-05-14",total:401831},
  {id:30,numero:"12142",cliente:"MUEBLERIA VALENTINA",fecha:"2026-05-14",total:419564},
  {id:31,numero:"12142b",cliente:"(sin nombre)",fecha:"2026-05-15",total:102475},
  {id:32,numero:"12143",cliente:"MUEBLES ASENJO",fecha:"2026-05-15",total:633474},
  {id:33,numero:"12144",cliente:"CLAUDIO ARELLANO",fecha:"2026-05-15",total:341710},
  {id:34,numero:"12146",cliente:"(sin nombre)",fecha:"2026-05-15",total:228639},
  {id:35,numero:"12147",cliente:"ESPACIO HABITADO",fecha:"2026-05-18",total:96744},
  {id:36,numero:"12158",cliente:"CARLOS SALINAS",fecha:"2026-05-18",total:176515},
  {id:37,numero:"12159",cliente:"MARTIN TEJEDA",fecha:"2026-05-18",total:7055679},
  {id:38,numero:"12160",cliente:"MARCELA BRAVO",fecha:"2026-05-19",total:396685},
  {id:39,numero:"12161",cliente:"ANDRES",fecha:"2026-05-18",total:165705},
  {id:40,numero:"12162",cliente:"ENRIQUE GUTIERREZ",fecha:"2026-05-19",total:106481},
  {id:41,numero:"12163",cliente:"THE ROCK SPA",fecha:"2026-05-19",total:80325},
  {id:42,numero:"12164",cliente:"CARLOS ROJAS",fecha:"2026-05-20",total:55569},
];

const INITIAL_NOTAS = [
  {id:1,numero:"7442",cliente:"HECTOR ROMERO",fecha:"2026-05-04",total:193953,cotizacion:null},
  {id:2,numero:"7443",cliente:"CLAUDIO BECERRA",fecha:"2026-05-04",total:180396,cotizacion:null},
  {id:3,numero:"7444",cliente:"OSVALDO IRIBARREN",fecha:"2026-05-04",total:109900,cotizacion:null},
  {id:4,numero:"7445",cliente:"CLAUDIO BECERRA",fecha:"2026-05-04",total:39000,cotizacion:null},
  {id:5,numero:"7446",cliente:"SANDRA TORRES",fecha:"2026-05-06",total:39000,cotizacion:null},
  {id:6,numero:"7447",cliente:"SUR-SERVI SPA",fecha:"2026-05-05",total:234161,cotizacion:null},
  {id:7,numero:"7448",cliente:"JUAN CARLOS YAÑEZ",fecha:"2026-05-05",total:143514,cotizacion:null},
  {id:8,numero:"7449",cliente:"MUEBLES ASENJO LTDA.",fecha:"2026-05-07",total:1713600,cotizacion:"12117"},
  {id:9,numero:"7450",cliente:"MAXIMILIANO MORALES",fecha:"2026-05-08",total:316169,cotizacion:null},
  {id:10,numero:"7451",cliente:"MUEBLES ASENJO LTDA.",fecha:"2026-05-11",total:1590827,cotizacion:"12119"},
  {id:11,numero:"7452",cliente:"MUEBLES ASENJO LTDA.",fecha:"2026-05-11",total:237931,cotizacion:"12122"},
  {id:12,numero:"7453",cliente:"DHOME DESIGN SPA",fecha:"2026-05-12",total:109179,cotizacion:"11127"},
  {id:13,numero:"7454",cliente:"PABLO MATURANA",fecha:"2026-05-12",total:119900,cotizacion:null},
  {id:14,numero:"7455",cliente:"OSVALDO IRIBARREN",fecha:"2026-05-13",total:110688,cotizacion:null},
  {id:15,numero:"7456",cliente:"GUILLERMO MUÑOZ",fecha:"2026-05-12",total:166565,cotizacion:"12123"},
  {id:16,numero:"7457",cliente:"GUSTAVO",fecha:"2026-05-12",total:65000,cotizacion:"12126"},
  {id:17,numero:"7458",cliente:"VALE EVANS",fecha:"2026-05-12",total:65000,cotizacion:"12135"},
  {id:18,numero:"7459",cliente:"DIEGO TOLEDO",fecha:"2026-05-12",total:275627,cotizacion:"12128"},
  {id:19,numero:"7460",cliente:"SANCHEZ SPA",fecha:"2026-05-14",total:65000,cotizacion:"12113"},
  {id:20,numero:"7461",cliente:"JUAN CARLOS YAÑEZ",fecha:"2026-05-05",total:146785,cotizacion:null},
  {id:21,numero:"7462",cliente:"NICOLAS SEPULVEDA",fecha:"2026-05-15",total:147980,cotizacion:null},
  {id:22,numero:"7463",cliente:"JOSE MARIO",fecha:"2026-05-15",total:148598,cotizacion:null},
  {id:23,numero:"7464",cliente:"JOSE MARIO",fecha:"2026-05-15",total:148598,cotizacion:null},
  {id:24,numero:"7465",cliente:"RENGO",fecha:"2026-05-18",total:102498,cotizacion:null},
  {id:25,numero:"7466",cliente:"DECOMADERA",fecha:"2026-05-15",total:26355,cotizacion:null},
  {id:26,numero:"7467",cliente:"CARLOS SALINAS",fecha:"2026-05-18",total:176515,cotizacion:null},
  {id:27,numero:"7468",cliente:"MUEBLES ASENJO LTDA.",fecha:"2026-05-19",total:633474,cotizacion:"12143"},
  {id:28,numero:"7469",cliente:"MUEBLES ASENJO LTDA.",fecha:"2026-05-19",total:1213047,cotizacion:"12127"},
  {id:29,numero:"7470",cliente:"THE ROCK SPA",fecha:"2026-05-19",total:67500,cotizacion:"12127"},
  {id:30,numero:"7471",cliente:"JORGE BORBARAN",fecha:"2026-05-19",total:128555,cotizacion:"12140"},
];

function businessDaysSince(dateStr) {
  const start = new Date(dateStr);
  const today = new Date();
  let count = 0;
  const cur = new Date(start);
  cur.setDate(cur.getDate() + 1);
  while (cur <= today) {
    if (cur.getDay() !== 0 && cur.getDay() !== 6) count++;
    cur.setDate(cur.getDate() + 1);
  }
  return count;
}

function getStatus(q, convertedNums) {
  if (convertedNums.has(q.numero)) return "vendida";
  const d = businessDaysSince(q.fecha);
  if (d > 10) return "vencida";
  if (d >= 7) return "urgente";
  return "activa";
}

function fmt(n) { return "$" + Number(n).toLocaleString("es-CL"); }
function fmtDate(iso) {
  if (!iso) return "";
  const [y,m,d] = iso.split("-");
  return `${d}/${m}/${y}`;
}

const STATUS_CONFIG = {
  vendida: { bg:"#1a3325", color:"#5a9e6f", border:"#2d5040", label:"✓ VENDIDA" },
  vencida: { bg:"#1e1525", color:"#9a7aaa", border:"#4a2a5a", label:"✕ VENCIDA" },
  urgente: { bg:"#2a1f0a", color:"#c8943a", border:"#5a3a10", label:"⚡ SEGUIMIENTO" },
  activa:  { bg:"#1a1f2a", color:"#5a8abe", border:"#2a3a5a", label:"● ACTIVA" },
};
const LEFT_COLOR = { vendida:COLORS.success, vencida:"#7a5a8a", urgente:COLORS.warning, activa:"#5a8abe" };

function StatusBadge({ status }) {
  const c = STATUS_CONFIG[status];
  return <span style={{ background:c.bg, color:c.color, border:`1px solid ${c.border}`, borderRadius:20, padding:"2px 10px", fontSize:11, fontWeight:600, whiteSpace:"nowrap" }}>{c.label}</span>;
}

function StatCard({ label, value, sub, color, icon }) {
  return (
    <div style={{ background:COLORS.card, border:`1px solid ${COLORS.border}`, borderRadius:12, padding:"18px 16px", display:"flex", flexDirection:"column", gap:5, position:"relative", overflow:"hidden" }}>
      <div style={{ position:"absolute", top:0, left:0, right:0, height:3, background:color||COLORS.accent }} />
      <span style={{ fontSize:22 }}>{icon}</span>
      <span style={{ fontSize:10, color:COLORS.muted, textTransform:"uppercase", letterSpacing:1 }}>{label}</span>
      <span style={{ fontSize:22, fontWeight:700, color:color||COLORS.accent, fontFamily:"Georgia,serif", lineHeight:1 }}>{value}</span>
      {sub && <span style={{ fontSize:10, color:COLORS.muted }}>{sub}</span>}
    </div>
  );
}

function NotasModal({ cotizacion, seguimiento, onSave, onClose }) {
  const [texto, setTexto] = useState("");
  const existing = seguimiento[cotizacion.numero] || [];

  const handleAdd = () => {
    if (!texto.trim()) return;
    const nueva = { fecha: new Date().toISOString().split("T")[0], texto: texto.trim() };
    onSave(cotizacion.numero, [...existing, nueva]);
    setTexto("");
  };

  const handleDelete = (idx) => {
    onSave(cotizacion.numero, existing.filter((_, i) => i !== idx));
  };

  return (
    <div style={{ position:"fixed", inset:0, background:"rgba(0,0,0,0.75)", display:"flex", alignItems:"center", justifyContent:"center", zIndex:100 }} onClick={onClose}>
      <div style={{ background:COLORS.card, border:`1px solid ${COLORS.border}`, borderRadius:16, padding:28, width:440, maxWidth:"95vw", maxHeight:"80vh", display:"flex", flexDirection:"column" }} onClick={e=>e.stopPropagation()}>
        <div style={{ marginBottom:16 }}>
          <h3 style={{ margin:"0 0 4px", color:COLORS.accent, fontFamily:"Georgia,serif", fontSize:17 }}>Seguimiento</h3>
          <span style={{ fontSize:12, color:COLORS.muted }}>#{cotizacion.numero} · {cotizacion.cliente}</span>
        </div>
        <div style={{ flex:1, overflowY:"auto", marginBottom:16 }}>
          {existing.length === 0 && <p style={{ color:COLORS.muted, fontSize:12, textAlign:"center", padding:"20px 0" }}>Sin notas aún.</p>}
          {existing.map((n, i) => (
            <div key={i} style={{ background:COLORS.surface, border:`1px solid ${COLORS.border}`, borderRadius:8, padding:"10px 12px", marginBottom:8, position:"relative" }}>
              <div style={{ fontSize:10, color:COLORS.muted, marginBottom:4 }}>📅 {fmtDate(n.fecha)}</div>
              <div style={{ fontSize:13, color:COLORS.text }}>{n.texto}</div>
              <button onClick={() => handleDelete(i)} style={{ position:"absolute", top:8, right:8, background:"none", border:"none", color:COLORS.danger, cursor:"pointer", fontSize:14 }}>✕</button>
            </div>
          ))}
        </div>
        <div style={{ display:"flex", gap:8 }}>
          <textarea placeholder="Ej: Llamé, no contestó. / Dijo que lo ve esta semana." value={texto} onChange={e=>setTexto(e.target.value)}
            onKeyDown={e=>{ if(e.key==="Enter"&&!e.shiftKey){e.preventDefault();handleAdd();}}}
            rows={2} style={{ flex:1, background:COLORS.surface, border:`1px solid ${COLORS.border}`, borderRadius:8, padding:"9px 12px", color:COLORS.text, fontSize:13, outline:"none", resize:"none", fontFamily:"inherit" }}/>
          <button onClick={handleAdd} style={{ background:COLORS.accent, border:"none", borderRadius:8, padding:"0 16px", color:"#0f0e0c", fontWeight:700, cursor:"pointer", fontSize:20 }}>+</button>
        </div>
        <button onClick={onClose} style={{ marginTop:10, background:COLORS.subtle, border:`1px solid ${COLORS.border}`, borderRadius:8, padding:"9px", color:COLORS.muted, cursor:"pointer", fontSize:13 }}>Cerrar</button>
      </div>
    </div>
  );
}

export default function App() {
  useEffect(() => {
  async function cargarDatos() {
    const dataCotizaciones = await obtenerCotizaciones();

    const cotizacionesFormateadas = dataCotizaciones.map((c) => ({
      id: `supabase-${c.id}`,
      numero: String(c.numero),
      cliente: c.cliente,
      total: c.total,
      fecha: c.fecha_creacion?.split("T")[0] || new Date().toISOString().split("T")[0],
    }));

    setCotizaciones(cotizacionesFormateadas);

    const dataNotas = await obtenerNotasVenta();

    const notasFormateadas = dataNotas.map((n) => ({
      id: `supabase-${n.id}`,
      numero: String(n.numero),
      cliente: n.cotizaciones?.cliente || n.cliente || "(sin cliente)",
      fecha: n.fecha?.split("T")[0] || new Date().toISOString().split("T")[0],
     total: n.cotizaciones?.total || n.total || 0,
      cotizacion: n.cotizaciones?.numero ? String(n.cotizaciones.numero) : null,
}));

    setNotas(notasFormateadas);
  }

  cargarDatos();
}, []);
  const [tab, setTab] = useState("dashboard");
  const [filter, setFilter] = useState("");
  const [showVencidas, setShowVencidas] = useState(false);
  const [seguimiento, setSeguimiento] = useState({});
  const [cotizaciones, setCotizaciones] = useState([]);
  const [notas, setNotas] = useState([]);
  const [modalCot, setModalCot] = useState(null);
  const [nuevoCliente, setNuevoCliente] = useState("");
  const [nuevoTotal, setNuevoTotal] = useState("");

  // Load seguimiento from localStorage
  useEffect(() => {
    try {
      const saved = localStorage.getItem("sf-seguimiento");
      if (saved) setSeguimiento(JSON.parse(saved));
    } catch(e) {}
  }, []);

  const saveSeguimiento = useCallback((numero, entries) => {
    const updated = { ...seguimiento, [numero]: entries };
    setSeguimiento(updated);
    try { localStorage.setItem("sf-seguimiento", JSON.stringify(updated)); } catch(e) {}
  }, [seguimiento]);

  const convertedNums = new Set(notas.filter(s => s.cotizacion).map(s => s.cotizacion));
  const sinCotizacion = notas.filter(s => !s.cotizacion);
  const withStatus = cotizaciones.map(q => ({ ...q, status: getStatus(q, convertedNums) }));
  const vendidas = withStatus.filter(q => q.status === "vendida");
  const urgentes = withStatus.filter(q => q.status === "urgente");
  const activas  = withStatus.filter(q => q.status === "activa");
  const vencidas = withStatus.filter(q => q.status === "vencida");

  const rate = cotizaciones.length > 0
    ? (vendidas.length / cotizaciones.length * 100).toFixed(1)
    : 0;
  const totalSold = notas.reduce((s,n) => s + n.total, 0);
  const totalActiva = [...activas,...urgentes].reduce((s,q) => s + q.total, 0);
  const totalVencida = vencidas.reduce((s,q) => s + q.total, 0);
  const totalQuoted = cotizaciones.reduce((s,q) => s + q.total, 0);

  const filteredQuotes = withStatus.filter(q =>
    (q.numero.toLowerCase().includes(filter.toLowerCase()) || q.cliente.toLowerCase().includes(filter.toLowerCase())) &&
    (showVencidas || q.status !== "vencida")
  ).sort((a,b) => { const o={urgente:0,activa:1,vendida:2,vencida:3}; return o[a.status]-o[b.status]; });

  const r2=58,cx=80,cy=76;
  const toRad=d=>d*Math.PI/180;
  const startA=200,sweepA=140;
  const arcPath=(s,e)=>{
    const p1={x:cx+r2*Math.cos(toRad(s)),y:cy+r2*Math.sin(toRad(s))};
    const p2={x:cx+r2*Math.cos(toRad(e)),y:cy+r2*Math.sin(toRad(e))};
    return `M ${p1.x} ${p1.y} A ${r2} ${r2} 0 ${e-s>180?1:0} 1 ${p2.x} ${p2.y}`;
  };
  const fillEnd=startA+sweepA*(rate/100);

  const tabs=[
    {key:"dashboard",label:"📊 Resumen"},
    {key:"quotes",label:`📋 Cotizaciones (${cotizaciones.length})`},
    {key:"sales",label:`✅ Notas de Venta (${notas.length})`},
    {key:"sinmatch",label:`⚠ Sin cruzar (${sinCotizacion.length})`},
  ];

  return (
    <div style={{ minHeight:"100vh", background:COLORS.bg, color:COLORS.text, fontFamily:"'Trebuchet MS',sans-serif", paddingBottom:60 }}>
 <div style={{ margin:20, padding:16, background:COLORS.card, border:`1px solid ${COLORS.border}`, borderRadius:12, display:"flex", gap:10, flexWrap:"wrap" }}>
  <input
    placeholder="Nombre cliente"
    value={nuevoCliente}
    onChange={(e) => setNuevoCliente(e.target.value)}
    style={{ padding:10, borderRadius:8, border:`1px solid ${COLORS.border}`, background:COLORS.surface, color:COLORS.text }}
  />

  <input
    type="number"
    placeholder="Total"
    value={nuevoTotal}
    onChange={(e) => setNuevoTotal(e.target.value)}
    style={{ padding:10, borderRadius:8, border:`1px solid ${COLORS.border}`, background:COLORS.surface, color:COLORS.text }}
  />

  <button
    onClick={async () => {
      if (!nuevoCliente || !nuevoTotal) {
        alert("Completa cliente y total");
        return;
      }

      await crearCotizacion(nuevoCliente, Number(nuevoTotal));

const dataActualizada = await obtenerCotizaciones();

const cotizacionesFormateadas = dataActualizada.map((c) => ({
  id: c.id,
  numero: String(c.numero),
  cliente: c.cliente,
  total: c.total,
  fecha: c.fecha_creacion?.split("T")[0] || new Date().toISOString().split("T")[0],
}));

setCotizaciones([...INITIAL_COTIZACIONES, ...cotizacionesFormateadas]);

alert("Cotización creada");
setNuevoCliente("");
setNuevoTotal("");
    }}
    style={{ padding:"10px 18px", background:COLORS.accent, border:"none", borderRadius:8, fontWeight:700, cursor:"pointer" }}
  >
    Crear Cotización
  </button>
</div>
      <div style={{ background:COLORS.surface, borderBottom:`1px solid ${COLORS.border}`, padding:"14px 24px", display:"flex", alignItems:"center", gap:12 }}>
        <span style={{ fontSize:24 }}>🪑</span>
        <div>
          <h1 style={{ margin:0, fontSize:17, color:COLORS.accent, fontFamily:"Georgia,serif" }}>Muebles Santa Fe · Panel de Conversión</h1>
          <span style={{ fontSize:11, color:COLORS.muted }}>Mayo 2026 · {cotizaciones.length} cotizaciones · {notas.length} notas de venta</span>
        </div>
      </div>

      <div style={{ display:"flex", gap:2, padding:"12px 24px 0", borderBottom:`1px solid ${COLORS.border}`, overflowX:"auto" }}>
        {tabs.map(t=>(
          <button key={t.key} onClick={()=>setTab(t.key)} style={{ background:tab===t.key?COLORS.card:"transparent", border:`1px solid ${tab===t.key?COLORS.border:"transparent"}`, borderBottom:tab===t.key?`2px solid ${COLORS.accent}`:"2px solid transparent", borderRadius:"8px 8px 0 0", padding:"8px 16px", color:tab===t.key?COLORS.accent:COLORS.muted, cursor:"pointer", fontSize:12, fontWeight:tab===t.key?700:400, whiteSpace:"nowrap" }}>{t.label}</button>
        ))}
      </div>

      <div style={{ padding:"22px 24px", maxWidth:1100, margin:"0 auto" }}>
        {tab==="dashboard" && (
          <div>
            <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fit,minmax(140px,1fr))", gap:12, marginBottom:20 }}>
              <StatCard label="Conversión" value={`${rate}%`} sub={`${vendidas.length} de ${cotizaciones.length}`} icon="🎯" color={COLORS.accent}/>
              <StatCard label="Vendidas" value={vendidas.length} sub={fmt(totalSold)} icon="✅" color={COLORS.success}/>
              <StatCard label="Activas" value={activas.length} sub={fmt(totalActiva)} icon="●" color="#5a8abe"/>
              <StatCard label="Seguimiento" value={urgentes.length} sub="7-10 días hábiles" icon="⚡" color={COLORS.warning}/>
              <StatCard label="Vencidas" value={vencidas.length} sub={fmt(totalVencida)} icon="✕" color="#9a7aaa"/>
              <StatCard label="Total Vendido" value={fmt(totalSold)} sub="30 notas" icon="💰" color={COLORS.success}/>
            </div>

            {urgentes.length>0 && (
              <div style={{ background:"#1e1500", border:`1px solid ${COLORS.warning}`, borderRadius:10, padding:"12px 16px", marginBottom:16 }}>
                <div style={{ fontSize:12, fontWeight:700, color:COLORS.warning, marginBottom:8 }}>⚡ {urgentes.length} cotizaciones requieren seguimiento urgente</div>
                {urgentes.map(q=>(
                  <div key={q.id} style={{ fontSize:12, color:COLORS.text, marginBottom:4, display:"flex", justifyContent:"space-between", alignItems:"center" }}>
                    <span><span style={{ color:COLORS.accent }}>#{q.numero}</span> {q.cliente} — {fmt(q.total)}</span>
                    <div style={{ display:"flex", alignItems:"center", gap:8 }}>
                      <span style={{ color:COLORS.muted, fontSize:11 }}>{businessDaysSince(q.fecha)} días háb.</span>
                      <button onClick={()=>setModalCot(q)} style={{ background:COLORS.warning, border:"none", borderRadius:6, padding:"3px 10px", color:"#0f0e0c", fontWeight:700, cursor:"pointer", fontSize:11 }}>+ Nota</button>
                   {q.status !== "vendida" && (
  <button
    onClick={async () => {
      const idReal = String(q.id).replace("supabase-", "");
      const nv = await aceptarCotizacion(idReal);

      if (nv) {
        alert(`Cotización aceptada. Nota de venta creada: ${nv}`);
        window.location.reload();
      }
    }}
    style={{
      background: COLORS.success,
      border: "none",
      borderRadius: 6,
      padding: "4px 10px",
      color: "#fff",
      cursor: "pointer",
      fontSize: 12,
      fontWeight: 700
    }}
  >
    Aceptar
  </button>
)}
                    </div>
                  </div>
                ))}
              </div>
            )}

            <div style={{ display:"grid", gridTemplateColumns:"200px 1fr", gap:16 }}>
              <div style={{ background:COLORS.card, border:`1px solid ${COLORS.border}`, borderRadius:12, padding:18, display:"flex", flexDirection:"column", alignItems:"center" }}>
                <span style={{ fontSize:11, color:COLORS.muted, textTransform:"uppercase", letterSpacing:1, marginBottom:4 }}>Conversión</span>
                <svg width="160" height="92" viewBox="0 0 160 92">
                  <path d={arcPath(startA,startA+sweepA)} stroke={COLORS.subtle} strokeWidth="11" fill="none" strokeLinecap="round"/>
                  <path d={arcPath(startA,fillEnd)} stroke={Number(rate)>=40?COLORS.success:COLORS.warning} strokeWidth="11" fill="none" strokeLinecap="round"/>
                </svg>
                <div style={{ fontSize:38, fontWeight:800, color:COLORS.accent, fontFamily:"Georgia,serif", lineHeight:1, marginTop:-10 }}>{rate}%</div>
                <div style={{ fontSize:11, color:COLORS.muted, marginTop:4, textAlign:"center" }}>{vendidas.length} vendidas · {vencidas.length} vencidas</div>
                <div style={{ marginTop:10, width:"100%", display:"flex", flexDirection:"column", gap:5 }}>
                  {[["✓ Vendidas",vendidas.length,COLORS.success],["● Activas",activas.length,"#5a8abe"],["⚡ Seguimiento",urgentes.length,COLORS.warning],["✕ Vencidas",vencidas.length,"#9a7aaa"]].map(([l,c,col])=>(
                    <div key={l} style={{ display:"flex", justifyContent:"space-between" }}>
                      <span style={{ fontSize:11, color:col }}>{l}</span><span style={{ fontSize:11, color:COLORS.muted }}>{c}</span>
                    </div>
                  ))}
                </div>
              </div>

              <div style={{ background:COLORS.card, border:`1px solid ${COLORS.border}`, borderRadius:12, padding:18, maxHeight:420, overflowY:"auto" }}>
                <span style={{ fontSize:11, color:COLORS.muted, textTransform:"uppercase", letterSpacing:1 }}>Estado por Cotización</span>
                <div style={{ marginTop:14 }}>
                  {withStatus.slice().sort((a,b)=>{ const o={urgente:0,activa:1,vendida:2,vencida:3}; return o[a.status]-o[b.status]||b.total-a.total; }).map(q=>{
                    const nvs=notas.filter(s=>s.cotizacion===q.numero);
                    const pct=totalQuoted>0?q.total/totalQuoted*100:0;
                    const nCount=(seguimiento[q.numero]||[]).length;
                    return (
                      <div key={q.id} style={{ marginBottom:11, opacity:q.status==="vencida"?0.5:1 }}>
                        <div style={{ display:"flex", justifyContent:"space-between", marginBottom:3 }}>
                          <span style={{ fontSize:12, color:COLORS.text }}><b style={{ color:COLORS.accent }}>#{q.numero}</b> {q.cliente}</span>
                          <span style={{ fontSize:12, fontWeight:600, color:q.status==="vendida"?COLORS.success:COLORS.muted }}>{fmt(q.total)}</span>
                        </div>
                        <div style={{ height:5, background:COLORS.subtle, borderRadius:3, overflow:"hidden" }}>
                          <div style={{ height:"100%", width:`${pct}%`, background:LEFT_COLOR[q.status], borderRadius:3 }}/>
                        </div>
                        <div style={{ display:"flex", justifyContent:"space-between", marginTop:2, alignItems:"center" }}>
                          <div style={{ display:"flex", alignItems:"center", gap:6 }}>
                            <StatusBadge status={q.status}/>
                            {nvs.length>0 && <span style={{ fontSize:10, color:COLORS.muted }}>→ {nvs.map(n=>`NV#${n.numero}`).join(", ")}</span>}
                            {nCount>0 && <span style={{ fontSize:10, color:COLORS.accent }}>📝{nCount}</span>}
                          </div>
                          <button onClick={()=>setModalCot(q)} style={{ background:"none", border:`1px solid ${COLORS.border}`, borderRadius:6, padding:"2px 8px", color:COLORS.muted, cursor:"pointer", fontSize:11 }}>📝</button>
                        </div>
                      </div>
                    );
                  })}
                </div>
              </div>
            </div>
          </div>
        )}

        {tab==="quotes" && (
          <div>
            <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:14, flexWrap:"wrap", gap:10 }}>
              <h2 style={{ margin:0, fontFamily:"Georgia,serif", color:COLORS.accent, fontSize:17 }}>Cotizaciones Mayo 2026</h2>
              <label style={{ display:"flex", alignItems:"center", gap:8, fontSize:12, color:COLORS.muted, cursor:"pointer" }}>
                <input type="checkbox" checked={showVencidas} onChange={e=>setShowVencidas(e.target.checked)}/> Mostrar vencidas ({vencidas.length})
              </label>
            </div>
            <input placeholder="Buscar por número o cliente..." value={filter} onChange={e=>setFilter(e.target.value)}
              style={{ width:"100%", boxSizing:"border-box", marginBottom:12, background:COLORS.surface, border:`1px solid ${COLORS.border}`, borderRadius:8, padding:"9px 14px", color:COLORS.text, fontSize:13, outline:"none" }}/>
            <div style={{ display:"flex", flexDirection:"column", gap:7 }}>
              {filteredQuotes.map(q=>{
                const nvs=notas.filter(s=>s.cotizacion===q.numero);
                const seg=seguimiento[q.numero]||[];
                const last=seg[seg.length-1];
                return (
                  <div key={q.id} style={{ background:COLORS.card, border:`1px solid ${STATUS_CONFIG[q.status].border}`, borderLeft:`4px solid ${LEFT_COLOR[q.status]}`, borderRadius:10, padding:"11px 14px", opacity:q.status==="vencida"?0.55:1 }}>
                    <div style={{ display:"flex", alignItems:"center", justifyContent:"space-between", flexWrap:"wrap", gap:8 }}>
                      <div>
                        <span style={{ fontWeight:700, color:COLORS.accent, marginRight:8 }}>#{q.numero}</span>
                        <span style={{ color:COLORS.text }}>{q.cliente}</span>
                        <span style={{ color:COLORS.muted, marginLeft:8, fontSize:11 }}>{q.fecha} · {businessDaysSince(q.fecha)}d háb.</span>
                        {nvs.length>0 && <span style={{ marginLeft:8, fontSize:11, color:COLORS.success }}>→ {nvs.map(n=>`NV#${n.numero}`).join(", ")}</span>}
                      </div>
                      <div style={{ display:"flex", alignItems:"center", gap:10 }}>
                        <span style={{ fontWeight:700, color:COLORS.text }}>{fmt(q.total)}</span>
                        <StatusBadge status={q.status}/>
                        <button onClick={()=>setModalCot(q)} style={{ background:seg.length>0?COLORS.subtle:"none", border:`1px solid ${COLORS.border}`, borderRadius:6, padding:"4px 10px", color:seg.length>0?COLORS.accent:COLORS.muted, cursor:"pointer", fontSize:12 }}>
                          📝 {seg.length>0?seg.length:""}
                        </button>
                        {q.status !== "vendida" && (
  <button
    onClick={async () => {
      const idReal = String(q.id).replace("supabase-", "");
      const nv = await aceptarCotizacion(idReal);

      if (nv) {
        alert(`Cotización aceptada. Nota de venta creada: ${nv}`);
      }
    }}
    style={{
      background: COLORS.success,
      border: "none",
      borderRadius: 6,
      padding: "4px 10px",
      color: "#fff",
      cursor: "pointer",
      fontSize: 12,
      fontWeight: 700
    }}
  >
    Aceptar
  </button>
)}
                      </div>
                    </div>
                    {last && (
                      <div style={{ marginTop:8, background:COLORS.surface, borderRadius:6, padding:"6px 10px", fontSize:11, color:COLORS.muted }}>
                        <span style={{ color:COLORS.accent }}>Último seguimiento {fmtDate(last.fecha)}:</span> {last.texto}
                      </div>
                    )}
                  </div>
                );
              })}
            </div>
          </div>
        )}

        {tab==="sales" && (
          <div>
            <h2 style={{ margin:"0 0 14px", fontFamily:"Georgia,serif", color:COLORS.success, fontSize:17 }}>Notas de Venta Mayo 2026</h2>
            <div style={{ marginBottom:14, background:COLORS.subtle, borderRadius:10, padding:"10px 16px", display:"flex", gap:24 }}>
              <div><span style={{ fontSize:11, color:COLORS.muted }}>TOTAL VENDIDO</span><div style={{ fontSize:18, fontWeight:700, color:COLORS.success }}>{fmt(totalSold)}</div></div>
              <div><span style={{ fontSize:11, color:COLORS.muted }}>NOTAS</span><div style={{ fontSize:18, fontWeight:700, color:COLORS.success }}>{notas.length}</div></div>
            </div>
            <div style={{ display:"flex", flexDirection:"column", gap:7 }}>
              {notas.map(s=>(
                <div key={s.id} style={{ background:COLORS.card, border:`1px solid ${s.cotizacion?"#2d5040":COLORS.border}`, borderLeft:`4px solid ${s.cotizacion?COLORS.success:COLORS.warning}`, borderRadius:10, padding:"11px 14px", display:"flex", alignItems:"center", justifyContent:"space-between", flexWrap:"wrap", gap:8 }}>
                  <div>
                    <span style={{ fontWeight:700, color:COLORS.success, marginRight:8 }}>NV#{s.numero}</span>
                    <span style={{ color:COLORS.text }}>{s.cliente}</span>
                    {s.cotizacion ? <span style={{ marginLeft:8, fontSize:11, color:COLORS.muted, background:COLORS.subtle, borderRadius:4, padding:"2px 7px" }}>← COT#{s.cotizacion}</span>
                      : <span style={{ marginLeft:8, fontSize:11, color:COLORS.warning, background:"#2a1f0a", borderRadius:4, padding:"2px 7px", border:`1px solid ${COLORS.warning}` }}>sin cotización</span>}
                    <span style={{ color:COLORS.muted, marginLeft:8, fontSize:11 }}>{s.fecha}</span>
                  </div>
                  <div style={{ display:"flex", alignItems:"center", gap:8 }}>
  <span style={{ fontWeight:700, color:COLORS.success }}>{fmt(s.total)}</span>

  {String(s.id).startsWith("supabase-") && (
    <button
      onClick={async () => {
        const nuevoNumero = prompt("Nuevo número de Nota de Venta:", s.numero);

        if (!nuevoNumero) return;

        const cotizacionRelacionada = cotizaciones.find(
          (q) => String(q.numero) === String(s.cotizacion)
        );

        if (!cotizacionRelacionada) {
          alert("No se encontró la cotización relacionada");
          return;
        }

        const idRealCotizacion = String(cotizacionRelacionada.id).replace("supabase-", "");

        const ok = await editarNumeroNotaVenta(idRealCotizacion, Number(nuevoNumero));

        if (ok) {
          alert("Número de Nota de Venta actualizado");
          window.location.reload();
        }
      }}
      style={{
        background: COLORS.subtle,
        border: `1px solid ${COLORS.border}`,
        borderRadius: 6,
        padding: "4px 8px",
        color: COLORS.accent,
        cursor: "pointer",
        fontSize: 11
      }}
    >
      Editar NV
    </button>
  )}
</div>
                </div>
              ))}
            </div>
          </div>
        )}

        {tab==="sinmatch" && (
          <div>
            <h2 style={{ margin:"0 0 6px", fontFamily:"Georgia,serif", color:COLORS.warning, fontSize:17 }}>⚠ Ventas sin cotización cruzada</h2>
            <p style={{ margin:"0 0 16px", fontSize:12, color:COLORS.muted }}>Notas sin cotización correspondiente en Mayo 2026.</p>
            <div style={{ display:"flex", flexDirection:"column", gap:7 }}>
              {sinCotizacion.map(s=>(
                <div key={s.id} style={{ background:COLORS.card, border:`1px solid #3a2a10`, borderLeft:`4px solid ${COLORS.warning}`, borderRadius:10, padding:"11px 14px", display:"flex", alignItems:"center", justifyContent:"space-between", flexWrap:"wrap", gap:8 }}>
                  <div>
                    <span style={{ fontWeight:700, color:COLORS.warning, marginRight:8 }}>NV#{s.numero}</span>
                    <span style={{ color:COLORS.text }}>{s.cliente}</span>
                    <span style={{ color:COLORS.muted, marginLeft:8, fontSize:11 }}>{s.fecha}</span>
                  </div>
                  <span style={{ fontWeight:700, color:COLORS.warning }}>{fmt(s.total)}</span>
                </div>
              ))}
            </div>
            <div style={{ marginTop:16, background:COLORS.subtle, borderRadius:10, padding:14 }}>
              <div style={{ fontSize:11, color:COLORS.muted, marginBottom:4 }}>TOTAL SIN COTIZACIÓN CRUZADA</div>
              <div style={{ fontSize:22, fontWeight:700, color:COLORS.warning }}>{fmt(sinCotizacion.reduce((s,n)=>s+n.total,0))}</div>
            </div>
          </div>
        )}
      </div>

      {modalCot && <NotasModal cotizacion={modalCot} seguimiento={seguimiento} onSave={saveSeguimiento} onClose={()=>setModalCot(null)}/>}
    </div>
  );
}
