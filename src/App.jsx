import { useState, useEffect, useCallback } from "react";
import { supabase } from '../lib/supabase'
import { obtenerCotizaciones, obtenerNotasVenta, editarNumeroNotaVenta, importarNotaVentaExcel, eliminarNotaVenta, importarCotizacionExcel, obtenerDetallesCotizaciones } from '../lib/cotizaciones'
import * as XLSX from 'xlsx'

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

function NotasModal({ cotizacion, detalles = [], seguimiento, onSave, onClose }) {
  const [texto, setTexto] = useState("");
  const existing = seguimiento[cotizacion.numero] || [];
  const totalCotizacion = Number(cotizacion.total) || 0;
  const netoCotizacion = Math.round(totalCotizacion / 1.19);
  const ivaCotizacion = totalCotizacion - netoCotizacion;
  useEffect(() => {
  const cerrarConEsc = (e) => {
    if (e.key === "Escape") onClose();
  };

  window.addEventListener("keydown", cerrarConEsc);

  return () => {
    window.removeEventListener("keydown", cerrarConEsc);
  };
}, [onClose]);

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
    <div
  onClick={onClose}
  style={{
    position:"fixed",
    inset:0,
    background:"rgba(0,0,0,0.75)",
    display:"flex",
    alignItems:"flex-start",
    justifyContent:"center",
    zIndex:100,
    overflowY:"auto",
    padding:"20px 10px"
  }}
>
      <div
  onClick={(e) => e.stopPropagation()}
  style={{
    background:COLORS.card,
    border:`1px solid ${COLORS.border}`,
    borderRadius:16,
    padding:28,
    width:440,
    maxWidth:"95vw",
    maxHeight:"90vh",
    overflowY:"auto",
    display:"flex",
    flexDirection:"column",
    boxSizing:"border-box"
  }}
>
              <div style={{ marginBottom:16 }}>
          <h3 style={{ margin:"0 0 4px", color:COLORS.accent, fontFamily:"Georgia,serif", fontSize:17 }}>Seguimiento</h3>
          <span style={{ fontSize:12, color:COLORS.muted }}>#{cotizacion.numero} · {cotizacion.cliente}</span>
        <div style={{
  marginTop:14,
  background:COLORS.surface,
  border:`1px solid ${COLORS.border}`,
  borderRadius:10,
  padding:12
}}>
  <div style={{ fontSize:12, fontWeight:700, color:COLORS.accent, marginBottom:8 }}>
    Detalle de cotización
  </div>

  {detalles.length === 0 ? (
    <div style={{ fontSize:12, color:COLORS.muted }}>
      Sin detalle guardado.
    </div>
  ) : (
    detalles.map((d) => (
      <div key={d.id} style={{
        fontSize:12,
        color:COLORS.text,
        padding:"7px 0",
        borderBottom:`1px solid ${COLORS.border}`
      }}>
        <b>{d.unidad}</b> x {d.tipo} · {d.largo} x {d.ancho} · {d.color}
        <br />
        <span style={{ color:COLORS.muted }}>
          Valor: {fmt(d.valor)} · Total: {fmt(d.total)}
        </span>
      </div>
    ))
    
  )}
  <div style={{
  marginTop:12,
  paddingTop:10,
  borderTop:`1px solid ${COLORS.border}`,
  display:"grid",
  gap:4,
  fontSize:12
}}>
  <div style={{ display:"flex", justifyContent:"space-between" }}>
    <span style={{ color:COLORS.muted }}>Neto</span>
    <b>{fmt(netoCotizacion)}</b>
  </div>

  <div style={{ display:"flex", justifyContent:"space-between" }}>
    <span style={{ color:COLORS.muted }}>IVA 19%</span>
    <b>{fmt(ivaCotizacion)}</b>
  </div>

  <div style={{
    display:"flex",
    justifyContent:"space-between",
    color:COLORS.accent,
    fontSize:14,
    marginTop:4
  }}>
    <span><b>Total</b></span>
    <b>{fmt(totalCotizacion)}</b>
  </div>
</div>
</div>
        </div>
        <div style={{ flex:1, overflowY:"auto", marginBottom:16 }}>
          {existing.length === 0 && <p style={{ color:COLORS.muted, fontSize:12, textAlign:"center", padding:"20px 0" }}>Sin notas aún.</p>}
          {existing.map((n, i) => (
            <div key={i} style={{ background:COLORS.surface, border:`1px solid ${COLORS.border}`, borderRadius:8, padding:"10px 12px", marginBottom:8, position:"relative" }}>
              <div style={{ fontSize:10, color:COLORS.muted, marginBottom:4 }}>📅 {fmtDate(n.fecha)}</div>
              <div style={{ fontSize:16, color:COLORS.text }}>{n.texto}</div>
              <button onClick={() => handleDelete(i)} style={{ position:"absolute", top:8, right:8, background:"none", border:"none", color:COLORS.danger, cursor:"pointer", fontSize:14 }}>✕</button>
            </div>
          ))}
        </div>
        <div style={{ display:"flex", gap:8 }}>
          <textarea
  placeholder="Ej: Llamé, no contestó. / Dijo que lo ve esta semana."
  value={texto}
  onChange={e=>setTexto(e.target.value)}
            onKeyDown={e=>{ if(e.key==="Enter"&&!e.shiftKey){e.preventDefault();handleAdd();}}}
            rows={2} style={{ flex:1, background:COLORS.surface, border:`1px solid ${COLORS.border}`, borderRadius:8, padding:"9px 12px", color:COLORS.text, fontSize:13, outline:"none", resize:"none", fontFamily:"inherit" }}/>
          <button onClick={handleAdd} style={{ background:COLORS.accent, border:"none", borderRadius:8, padding:"0 16px", color:"#0f0e0c", fontWeight:700, cursor:"pointer", fontSize:20 }}>+</button>
        </div>
        <button onClick={onClose} style={{ marginTop:10, background:COLORS.subtle, border:`1px solid ${COLORS.border}`, borderRadius:8, padding:"9px", color:COLORS.muted, cursor:"pointer", fontSize:13 }}>Cerrar</button>
      </div>
    </div>
  );
}
function ProduccionModal({ nota, detalles = [], onClose, onSave }) {
  const [fechaEntrega, setFechaEntrega] = useState(nota.fecha_entrega_estimada || "");
  const [observaciones, setObservaciones] = useState(nota.produccion_observaciones || "");
  useEffect(() => {
  const cerrarConEsc = (e) => {
    if (e.key === "Escape") onClose();
  };

  window.addEventListener("keydown", cerrarConEsc);

  return () => {
    window.removeEventListener("keydown", cerrarConEsc);
  };
}, [onClose]);

const guardarProduccionModal = () => {
  onSave({
    ...nota,
    fecha_entrega_estimada: fechaEntrega,
    produccion_observaciones: observaciones,
    ...procesos
  });
};

  const [procesos, setProcesos] = useState({
    mdf_cortado: nota.mdf_cortado || false,
    lamina_cortada: nota.lamina_cortada || false,
    tupizado: nota.tupizado || false,
    armado: nota.armado || false,
    pegado: nota.pegado || false,
    postformado: nota.postformado || false,
  });

  const ordenProcesos = [
    "mdf_cortado",
    "lamina_cortada",
    "tupizado",
    "armado",
    "pegado",
    "postformado",
  ];

  const nombresProcesos = {
    mdf_cortado: "MDF cortado",
    lamina_cortada: "Lámina cortada",
    tupizado: "Tupizado",
    armado: "Armado",
    pegado: "Pegado",
    postformado: "Postformado",
  };

  const toggleProceso = (key) => {
    const index = ordenProcesos.indexOf(key);

    setProcesos(prev => {
      const nuevo = { ...prev };
      const nuevoValor = !prev[key];

      if (nuevoValor) {
        for (let i = 0; i <= index; i++) {
          nuevo[ordenProcesos[i]] = true;
        }
      } else {
        nuevo[key] = false;
      }

      return nuevo;
    });
  };

  return (
    <div
  onClick={onClose}
  style={{
    position:"fixed",
    inset:0,
    background:"rgba(0,0,0,0.65)",
    display:"flex",
    alignItems:"flex-start",
    justifyContent:"center",
    zIndex:9999,
    overflowY:"auto",
    padding:"20px 10px"
  }}
>
      <div
  onClick={(e) => e.stopPropagation()}
  style={{
    background:COLORS.card,
    border:`1px solid ${COLORS.border}`,
    borderRadius:14,
    padding:22,
    width:"460px",
    maxWidth:"92%",
    maxHeight:"90vh",
    overflowY:"auto",
    color:COLORS.text,
    boxSizing:"border-box"
  }}
>
        <h2 style={{ marginTop:0, color:COLORS.accent }}>
          Producción NV#{nota.numero}
        </h2>

        <p><b>Cliente:</b> {nota.cliente}</p>
        <div style={{
  background:COLORS.surface,
  border:`1px solid ${COLORS.border}`,
  borderRadius:10,
  padding:12,
  margin:"10px 0 14px"
}}>
  <h3 style={{ margin:"0 0 10px", color:COLORS.accent, fontSize:15 }}>
    Detalle de producción
  </h3>

  {detalles.length === 0 ? (
    <p style={{ margin:0, color:COLORS.muted, fontSize:13 }}>
      Esta nota de venta no tiene detalle de producción importado.
    </p>
  ) : (
    <div style={{ display:"flex", flexDirection:"column", gap:8 }}>
      {detalles.map((d) => (
        <div key={d.id || d.orden} style={{
          border:`1px solid ${COLORS.border}`,
          borderRadius:8,
          padding:10,
          background:COLORS.card
        }}>
          <div style={{ fontWeight:700, color:COLORS.text }}>
            {d.cantidad || 0} x {d.descripcion || "Sin descripción"}
          </div>

          <div style={{ fontSize:13, color:COLORS.muted, marginTop:4 }}>
            Material: <b style={{ color:COLORS.text }}>{d.material || "-"}</b>
          </div>

          <div style={{ fontSize:13, color:COLORS.muted, marginTop:4 }}>
            Medida: <b style={{ color:COLORS.text }}>{d.alto || 0} x {d.ancho || 0}</b>
          </div>

          <div style={{ fontSize:13, color:COLORS.muted, marginTop:4 }}>
            Color: <b style={{ color:COLORS.text }}>{d.color || "-"}</b>
          </div>
        </div>
      ))}
    </div>
  )}
</div>
        <label>Fecha estimada de entrega</label>
        <input
          type="date"
          value={fechaEntrega}
          onChange={(e) => setFechaEntrega(e.target.value)}
          style={{
            width:"100%",
            padding:10,
            margin:"6px 0 14px",
            borderRadius:8,
            border:`1px solid ${COLORS.border}`,
            background:COLORS.surface,
            color:COLORS.text
          }}
        />

        <div style={{ display:"grid", gap:8, marginBottom:14 }}>
          {ordenProcesos.map(key => (
            <label key={key} style={{
              display:"flex",
              alignItems:"center",
              gap:8,
              background:COLORS.surface,
              border:`1px solid ${COLORS.border}`,
              borderRadius:8,
              padding:10,
              cursor:"pointer"
            }}>
              <input
                type="checkbox"
                checked={procesos[key]}
                onChange={() => toggleProceso(key)}
              />
              {nombresProcesos[key]}
            </label>
          ))}
        </div>

        <label>Observaciones producción</label>
        <textarea
  value={observaciones}
  onChange={(e) => setObservaciones(e.target.value)}
  onKeyDown={(e) => {
    if (e.key === "Enter" && !e.shiftKey) {
      e.preventDefault();
      guardarProduccionModal();
    }
  }}
          placeholder="Ej: pedido urgente, falta lámina, cliente pidió cambio..."
          style={{
            width:"100%",
            minHeight:90,
            padding:10,
            margin:"6px 0 14px",
            borderRadius:8,
            background:COLORS.surface,
            color:COLORS.text,
            border:`1px solid ${COLORS.border}`,
fontSize:16,
resize:"none",
boxSizing:"border-box"
          }}
        />

        <div style={{ display:"flex", justifyContent:"flex-end", gap:10 }}>
          <button onClick={onClose}>Cancelar</button>

          <button
            onClick={guardarProduccionModal}
            style={{
              background:COLORS.success,
              color:"#fff",
              border:"none",
              borderRadius:8,
              padding:"9px 14px",
              cursor:"pointer",
              fontWeight:700
            }}
          >
            Guardar
          </button>
        </div>
      </div>
    </div>
  );
}
function GestionNVModal({ nota, abonos = [], onClose, onSave }) {
  const [nuevoAbono, setNuevoAbono] = useState("");
  const [observacionAbono, setObservacionAbono] = useState("");
  useEffect(() => {
  const cerrarConEsc = (e) => {
    if (e.key === "Escape") onClose();
  };

  window.addEventListener("keydown", cerrarConEsc);

  return () => {
    window.removeEventListener("keydown", cerrarConEsc);
  };
}, [onClose]);

const guardarGestionNV = () => {
  onSave({
    ...nota,
    nuevoAbono,
    observacionAbono,
    materiales,
    proceso,
    observaciones
  });
};
  const [materiales, setMateriales] = useState(nota.materiales || "falta");
  const [proceso, setProceso] = useState(nota.proceso || "en espera");
  const [observaciones, setObservaciones] = useState(nota.observaciones || "");

  const total = Number(nota.total) || 0;
  const totalAbonado = abonos.reduce((sum, a) => sum + Number(a.monto || 0), 0);
  const saldo = Math.max(total - totalAbonado, 0);

  return (
    <div
  onClick={onClose}
  style={{
    position:"fixed",
    inset:0,
    background:"rgba(0,0,0,0.65)",
    display:"flex",
    alignItems:"flex-start",
    justifyContent:"center",
    zIndex:9999,
    overflowY:"auto",
    padding:"20px 10px"
  }}
>
      <div
  onClick={(e) => e.stopPropagation()}
  style={{
    background:COLORS.card,
    border:`1px solid ${COLORS.border}`,
    borderRadius:14,
    padding:22,
    width:"420px",
    maxWidth:"92%",
    maxHeight:"90vh",
    overflowY:"auto",
    color:COLORS.text,
    boxSizing:"border-box"
  }}
>
        <h2 style={{ marginTop:0, color:COLORS.accent }}>
          Gestión NV #{nota.numero}
        </h2>

        <p><b>Cliente:</b> {nota.cliente}</p>
        <p><b>Total:</b> {fmt(total)}</p>

        <label>Nuevo abono</label>
<input
  type="number"
  value={nuevoAbono}
  onChange={(e) => setNuevoAbono(e.target.value)}
  placeholder="Ej: 50000"
  style={{
    width:"100%",
    padding:10,
    margin:"6px 0 12px",
    borderRadius:8,
    border:`1px solid ${COLORS.border}`,
    background:COLORS.surface,
    color:COLORS.text
  }}
/>

<label>Observación del abono</label>
<input
  type="text"
  value={observacionAbono}
  onChange={(e) => setObservacionAbono(e.target.value)}
  placeholder="Ej: transferencia, efectivo, banco..."
  style={{
    width:"100%",
    padding:10,
    margin:"6px 0 12px",
    borderRadius:8,
    border:`1px solid ${COLORS.border}`,
    background:COLORS.surface,
    color:COLORS.text
  }}
/>
<p><b>Total abonado:</b> {fmt(totalAbonado)}</p>

<div style={{
  margin:"12px 0",
  padding:12,
  borderRadius:10,
  background:COLORS.surface,
  border:`1px solid ${COLORS.border}`
}}>
  <b>Historial de abonos</b>

  {abonos.length === 0 ? (
    <p style={{ color:COLORS.muted, fontSize:13 }}>
      Sin abonos registrados
    </p>
  ) : (
    <div style={{ marginTop:8, display:"grid", gap:6 }}>
      {abonos.map(a => (
        <div key={a.id} style={{
          display:"flex",
          justifyContent:"space-between",
          gap:10,
          fontSize:13,
          borderBottom:`1px solid ${COLORS.border}`,
          paddingBottom:5
        }}>
          <span>{a.fecha}</span>
          <span style={{ fontWeight:700 }}>{fmt(a.monto)}</span>
          <span style={{ color:COLORS.muted }}>
            {a.observacion || "Sin observación"}
          </span>
        </div>
      ))}
    </div>
  )}
</div>
        <p>
          <b>Saldo pendiente:</b>{" "}
          <span style={{ color: saldo === 0 ? COLORS.success : COLORS.warning }}>
            {fmt(saldo)}
          </span>
        </p>

        <label>Materiales</label>
        <select
          value={materiales}
          onChange={(e) => setMateriales(e.target.value)}
          style={{
            width:"100%",
            padding:10,
            margin:"6px 0 12px",
            borderRadius:8,
            background:COLORS.surface,
            color:COLORS.text,
            border:`1px solid ${COLORS.border}`
          }}
        >
          <option value="falta">Falta</option>
          <option value="comprados">Comprados</option>
        </select>

        <label>Proceso</label>
        <select
          value={proceso}
          onChange={(e) => setProceso(e.target.value)}
          style={{
            width:"100%",
            padding:10,
            margin:"6px 0 12px",
            borderRadius:8,
            background:COLORS.surface,
            color:COLORS.text,
            border:`1px solid ${COLORS.border}`
          }}
        >
          <option value="en espera">En espera</option>
          <option value="en producción">En producción</option>
          <option value="terminado">Terminado</option>
          <option value="entregado">Entregado</option>
        </select>

        <label>Observaciones</label>
        <textarea
  value={observaciones}
  onChange={(e) => setObservaciones(e.target.value)}
  onKeyDown={(e) => {
    if (e.key === "Enter" && !e.shiftKey) {
      e.preventDefault();
      guardarGestionNV();
    }
  }}
          placeholder="Escribe una observación..."
          style={{
            width:"100%",
            minHeight:90,
            padding:10,
            margin:"6px 0 14px",
            borderRadius:8,
            background:COLORS.surface,
            color:COLORS.text,
            border:`1px solid ${COLORS.border}`,
fontSize:16,
resize:"none",
boxSizing:"border-box"
          }}
        />

        <div style={{ display:"flex", justifyContent:"flex-end", gap:10 }}>
          <button onClick={onClose}>
            Cancelar
          </button>

          <button
            onClick={guardarGestionNV}
            style={{
              background:COLORS.success,
              color:"#fff",
              border:"none",
              borderRadius:8,
              padding:"9px 14px",
              cursor:"pointer",
              fontWeight:700
            }}
          >
            Guardar
          </button>
        </div>
      </div>
    </div>
  );
}
function SobrantesLaminadoModal({ producto, sobrantes, onClose, onSave, onUsar }) {
  const [filas, setFilas] = useState([{ largo:"", ancho:"" }]);
  const [largoBuscado, setLargoBuscado] = useState("");
  const [anchoBuscado, setAnchoBuscado] = useState("");
  const [buscarTodos, setBuscarTodos] = useState(false);
  useEffect(() => {
    const cerrarConEsc = (e) => {
      if (e.key === "Escape") onClose();
    };

    window.addEventListener("keydown", cerrarConEsc);

    return () => {
      window.removeEventListener("keydown", cerrarConEsc);
    };
  }, [onClose]);

  const cambiarFila = (index, campo, valor) => {
    setFilas(prev =>
      prev.map((fila, i) =>
        i === index ? { ...fila, [campo]: valor } : fila
      )
    );
  };

  const agregarFila = () => {
    setFilas(prev => [...prev, { largo:"", ancho:"" }]);
  };

  const sobrantesDisponibles = sobrantes.filter(s => !s.usado);
  const sobrantesCompatibles = sobrantesDisponibles.filter((s) => {
  const largoNecesario = Number(largoBuscado) || 0;
  const anchoNecesario = Number(anchoBuscado) || 0;

  if (largoNecesario <= 0 || anchoNecesario <= 0) return false;

  return Number(s.largo) >= largoNecesario && Number(s.ancho) >= anchoNecesario;
});
  return (
    <div
      onClick={onClose}
      style={{
        position:"fixed",
        inset:0,
        background:"rgba(0,0,0,0.65)",
        display:"flex",
        alignItems:"flex-start",
        justifyContent:"center",
        zIndex:9999,
        overflowY:"auto",
        padding:"20px 10px"
      }}
    >
      <div
        onClick={(e) => e.stopPropagation()}
        style={{
          background:COLORS.card,
          border:`1px solid ${COLORS.border}`,
          borderRadius:14,
          padding:22,
          width:"560px",
          maxWidth:"92%",
          maxHeight:"90vh",
          overflowY:"auto",
          color:COLORS.text,
          boxSizing:"border-box"
        }}
      >
        <h2 style={{ margin:"0 0 8px", color:COLORS.accent }}>
          Sobrantes
        </h2>

        <p style={{ margin:"0 0 14px" }}>
          <b>{producto.nombre}</b>
        </p>
        
        <div style={{
  border:`1px solid ${COLORS.border}`,
  borderRadius:10,
  padding:12,
  background:COLORS.surface,
  marginBottom:16
}}>
  <h3 style={{ fontSize:14, margin:"0 0 10px", color:COLORS.accent }}>
    Buscar sobrante compatible
  </h3>

  <div style={{
    display:"grid",
    gridTemplateColumns:"1fr 1fr",
    gap:8,
    marginBottom:10
  }}>
    <input
      type="number"
      placeholder="Largo necesario"
      value={largoBuscado}
      onChange={(e) => setLargoBuscado(e.target.value)}
      style={{
        padding:"9px",
        borderRadius:8,
        border:`1px solid ${COLORS.border}`,
        background:COLORS.card,
        color:COLORS.text,
        fontSize:16,
        boxSizing:"border-box",
        width:"100%"
      }}
    />

    <input
      type="number"
      placeholder="Ancho necesario"
      value={anchoBuscado}
      onChange={(e) => setAnchoBuscado(e.target.value)}
      style={{
        padding:"9px",
        borderRadius:8,
        border:`1px solid ${COLORS.border}`,
        background:COLORS.card,
        color:COLORS.text,
        fontSize:16,
        boxSizing:"border-box",
        width:"100%"
      }}
    />
  </div>

  {Number(largoBuscado) > 0 && Number(anchoBuscado) > 0 && (
    <div>
      <p style={{ margin:"0 0 8px", fontSize:13, color:COLORS.muted }}>
        Resultado:
      </p>

      {sobrantesCompatibles.length === 0 ? (
        <p style={{ margin:0, color:COLORS.danger, fontSize:13, fontWeight:700 }}>
          No hay sobrantes compatibles.
        </p>
      ) : (
        <div style={{ display:"flex", flexDirection:"column", gap:6 }}>
          {sobrantesCompatibles.map((s) => (
            <div
              key={s.id}
              style={{
                border:`1px solid ${COLORS.border}`,
                borderRadius:8,
                padding:8,
                background:COLORS.card,
                display:"flex",
                justifyContent:"space-between",
                gap:10
              }}
            >
              <span>
                {Number(s.largo)} x {Number(s.ancho)}
              </span>

              <b style={{ color:COLORS.success }}>
                Sirve
              </b>
            </div>
          ))}
        </div>
      )}
    </div>
  )}
</div>
        <h3 style={{ fontSize:14, margin:"0 0 8px", color:COLORS.success }}>
          Sobrantes disponibles
        </h3>

        {sobrantesDisponibles.length === 0 ? (
          <p style={{ color:COLORS.muted, fontSize:13 }}>
            No hay sobrantes registrados.
          </p>
        ) : (
          <div style={{ display:"flex", flexDirection:"column", gap:8, marginBottom:18 }}>
            {sobrantesDisponibles.map((s) => (
              <div
                key={s.id}
                style={{
                  border:`1px solid ${COLORS.border}`,
                  borderRadius:8,
                  padding:10,
                  background:COLORS.surface,
                  display:"flex",
                  justifyContent:"space-between",
                  alignItems:"center",
                  gap:10
                }}
              >
                <span>
                  {Number(s.largo)} x {Number(s.ancho)}
                </span>

                <button
                  onClick={() => onUsar(s)}
                  style={{
                    padding:"6px 8px",
                    borderRadius:8,
                    border:"none",
                    background:COLORS.danger,
                    color:"#fff",
                    fontWeight:700,
                    cursor:"pointer"
                  }}
                >
                  Usado
                </button>
              </div>
            ))}
          </div>
        )}

        <h3 style={{ fontSize:14, margin:"12px 0 8px", color:COLORS.accent }}>
          Añadir sobrantes
        </h3>

        {filas.map((fila, index) => (
          <div
            key={index}
            style={{
              display:"grid",
              gridTemplateColumns:"1fr 1fr",
              gap:8,
              marginBottom:8
            }}
          >
            <input
              type="number"
              placeholder="Largo"
              value={fila.largo}
              onChange={(e) => cambiarFila(index, "largo", e.target.value)}
              style={{
                padding:"9px",
                borderRadius:8,
                border:`1px solid ${COLORS.border}`,
                background:COLORS.surface,
                color:COLORS.text,
                fontSize:16,
                boxSizing:"border-box",
                width:"100%"
              }}
            />

            <input
              type="number"
              placeholder="Ancho"
              value={fila.ancho}
              onChange={(e) => cambiarFila(index, "ancho", e.target.value)}
              style={{
                padding:"9px",
                borderRadius:8,
                border:`1px solid ${COLORS.border}`,
                background:COLORS.surface,
                color:COLORS.text,
                fontSize:16,
                boxSizing:"border-box",
                width:"100%"
              }}
            />
          </div>
        ))}

        <button
          onClick={agregarFila}
          style={{
            width:"100%",
            padding:"8px",
            borderRadius:8,
            border:`1px solid ${COLORS.border}`,
            background:COLORS.surface,
            color:COLORS.text,
            fontWeight:700,
            cursor:"pointer",
            marginBottom:12
          }}
        >
          + Agregar otra fila
        </button>

        <div style={{ display:"flex", gap:10, justifyContent:"flex-end" }}>
          <button onClick={onClose} style={{ padding:"9px 12px", borderRadius:8 }}>
            Cancelar
          </button>

          <button
            onClick={() => onSave({ producto, sobrantes:filas })}
            style={{
              padding:"9px 14px",
              borderRadius:8,
              border:"none",
              background:COLORS.success,
              color:"#fff",
              fontWeight:700,
              cursor:"pointer"
            }}
          >
            Guardar sobrantes
          </button>
        </div>
      </div>
    </div>
  );
}
function PreviewOCModal({ productos, onClose, onConfirm }) {
  useEffect(() => {
    const cerrarConEsc = (e) => {
      if (e.key === "Escape") onClose();
    };

    window.addEventListener("keydown", cerrarConEsc);

    return () => {
      window.removeEventListener("keydown", cerrarConEsc);
    };
  }, [onClose]);

  return (
    <div
      onClick={onClose}
      style={{
        position:"fixed",
        inset:0,
        background:"rgba(0,0,0,0.65)",
        display:"flex",
        alignItems:"flex-start",
        justifyContent:"center",
        zIndex:9999,
        overflowY:"auto",
        padding:"20px 10px"
      }}
    >
      <div
        onClick={(e) => e.stopPropagation()}
        style={{
          background:COLORS.card,
          border:`1px solid ${COLORS.border}`,
          borderRadius:14,
          padding:22,
          width:"520px",
          maxWidth:"92%",
          maxHeight:"90vh",
          overflowY:"auto",
          color:COLORS.text,
          boxSizing:"border-box"
        }}
      >
        <h2 style={{ margin:"0 0 10px", color:COLORS.accent }}>
          Vista previa OC
        </h2>

        <p style={{ color:COLORS.muted, fontSize:13 }}>
          Revisa los productos antes de ingresarlos al inventario.
        </p>

        <div style={{ display:"flex", flexDirection:"column", gap:8, marginTop:12 }}>
          {productos.map((p, index) => (
            <div
              key={index}
              style={{
                border:`1px solid ${COLORS.border}`,
                borderRadius:8,
                padding:10,
                background:COLORS.surface,
                display:"flex",
                justifyContent:"space-between",
                gap:10
              }}
            >
              <span>{p.nombre}</span>
              <b style={{ color:COLORS.accent }}>{p.cantidad}</b>
            </div>
          ))}
        </div>

        <div style={{ display:"flex", gap:10, justifyContent:"flex-end", marginTop:16 }}>
          <button onClick={onClose} style={{ padding:"9px 12px", borderRadius:8 }}>
            Cancelar
          </button>

          <button
            onClick={onConfirm}
            style={{
              padding:"9px 14px",
              borderRadius:8,
              border:"none",
              background:COLORS.success,
              color:"#fff",
              fontWeight:700,
              cursor:"pointer"
            }}
          >
            Confirmar ingreso
          </button>
        </div>
      </div>
    </div>
  );
}
function EditarProductoModal({ producto, onClose, onSave }) {
  const [nombre, setNombre] = useState(producto.nombre || "");
  const [categoria, setCategoria] = useState(producto.categoria || "Tableros");
  const [unidad, setUnidad] = useState(producto.unidad || "unidades");
  const [stockMinimo, setStockMinimo] = useState(producto.stock_minimo || "");
  const [proveedor, setProveedor] = useState(producto.proveedor || "");
  const [esLaminado, setEsLaminado] = useState(producto.es_laminado || false);
  const [observaciones, setObservaciones] = useState(producto.observaciones || "");

  useEffect(() => {
    const cerrarConEsc = (e) => {
      if (e.key === "Escape") onClose();
    };

    window.addEventListener("keydown", cerrarConEsc);

    return () => {
      window.removeEventListener("keydown", cerrarConEsc);
    };
  }, [onClose]);

  const estiloInput = {
    width:"100%",
    padding:"9px 10px",
    borderRadius:8,
    border:`1px solid ${COLORS.border}`,
    background:COLORS.surface,
    color:COLORS.text,
    fontSize:16,
    boxSizing:"border-box"
  };

  const guardar = () => {
    if (!nombre.trim()) {
      alert("Debes ingresar el nombre del producto.");
      return;
    }

    onSave({
      ...producto,
      nombre: nombre.trim().toUpperCase(),
      categoria,
      unidad,
      stock_minimo: Number(stockMinimo) || 0,
      proveedor: proveedor.trim(),
      es_laminado: esLaminado,
      observaciones: observaciones.trim()
    });
  };

  return (
    <div
      onClick={onClose}
      style={{
        position:"fixed",
        inset:0,
        background:"rgba(0,0,0,0.65)",
        display:"flex",
        alignItems:"flex-start",
        justifyContent:"center",
        zIndex:9999,
        overflowY:"auto",
        padding:"20px 10px"
      }}
    >
      <div
        onClick={(e) => e.stopPropagation()}
        style={{
          background:COLORS.card,
          border:`1px solid ${COLORS.border}`,
          borderRadius:14,
          padding:22,
          width:"460px",
          maxWidth:"92%",
          maxHeight:"90vh",
          overflowY:"auto",
          color:COLORS.text,
          boxSizing:"border-box"
        }}
      >
        <h2 style={{ margin:"0 0 14px", color:COLORS.accent }}>
          Editar producto
        </h2>

        <label>Nombre del producto</label>
        <input
          value={nombre}
          onChange={(e) => setNombre(e.target.value)}
          style={{ ...estiloInput, margin:"6px 0 12px" }}
        />

        <label>Categoría</label>
        <select
          value={categoria}
          onChange={(e) => setCategoria(e.target.value)}
          style={{ ...estiloInput, margin:"6px 0 12px" }}
        >
          <option>Tableros</option>
          <option>Laminados</option>
          <option>Pegamentos</option>
          <option>Embalaje</option>
          <option>Tornillos</option>
          <option>Herrajes</option>
          <option>Insumos generales</option>
        </select>

        <label>Unidad</label>
        <select
          value={unidad}
          onChange={(e) => setUnidad(e.target.value)}
          style={{ ...estiloInput, margin:"6px 0 12px" }}
        >
          <option>planchas</option>
          <option>láminas</option>
          <option>tarros</option>
          <option>bolsas</option>
          <option>unidades</option>
          <option>rollos</option>
          <option>metros</option>
          <option>kilos</option>
          <option>litros</option>
        </select>

        <label>Stock mínimo</label>
        <input
          type="number"
          value={stockMinimo}
          onChange={(e) => setStockMinimo(e.target.value)}
          style={{ ...estiloInput, margin:"6px 0 12px" }}
        />

        <label>Proveedor habitual</label>
        <input
          value={proveedor}
          onChange={(e) => setProveedor(e.target.value)}
          placeholder="Ej: Imperial, Merino, Latam"
          style={{ ...estiloInput, margin:"6px 0 12px" }}
        />

        <label style={{ display:"flex", gap:8, alignItems:"center", margin:"8px 0 12px" }}>
          <input
            type="checkbox"
            checked={esLaminado}
            onChange={(e) => {
              setEsLaminado(e.target.checked);
              if (e.target.checked) {
                setCategoria("Laminados");
                setUnidad("láminas");
              }
            }}
          />
          Es laminado
        </label>

        <label>Observaciones</label>
        <textarea
          value={observaciones}
          onChange={(e) => setObservaciones(e.target.value)}
          style={{
            ...estiloInput,
            minHeight:70,
            resize:"none",
            margin:"6px 0 16px"
          }}
        />

        <div style={{ display:"flex", gap:10, justifyContent:"flex-end" }}>
          <button onClick={onClose} style={{ padding:"9px 12px", borderRadius:8 }}>
            Cancelar
          </button>

          <button
            onClick={guardar}
            style={{
              padding:"9px 14px",
              borderRadius:8,
              border:"none",
              background:COLORS.accent,
              color:"#111",
              fontWeight:700,
              cursor:"pointer"
            }}
          >
            Guardar cambios
          </button>
        </div>
      </div>
    </div>
  );
}
function MovimientoInventarioModal({ data, onClose, onSave }) {
  const [cantidad, setCantidad] = useState("");

  const producto = data.producto;
  const tipo = data.tipo;

  useEffect(() => {
    const cerrarConEsc = (e) => {
      if (e.key === "Escape") onClose();
    };

    window.addEventListener("keydown", cerrarConEsc);

    return () => {
      window.removeEventListener("keydown", cerrarConEsc);
    };
  }, [onClose]);

  const titulo = tipo === "entrada" ? "Entrada de stock" : "Salida de stock";

  return (
    <div
      onClick={onClose}
      style={{
        position:"fixed",
        inset:0,
        background:"rgba(0,0,0,0.65)",
        display:"flex",
        alignItems:"flex-start",
        justifyContent:"center",
        zIndex:9999,
        overflowY:"auto",
        padding:"20px 10px"
      }}
    >
      <div
        onClick={(e) => e.stopPropagation()}
        style={{
          background:COLORS.card,
          border:`1px solid ${COLORS.border}`,
          borderRadius:14,
          padding:22,
          width:"420px",
          maxWidth:"92%",
          color:COLORS.text,
          boxSizing:"border-box"
        }}
      >
        <h2 style={{ margin:"0 0 14px", color:COLORS.accent }}>
          {titulo}
        </h2>

        <p style={{ margin:"0 0 12px" }}>
          <b>Producto:</b> {producto.nombre}
        </p>

        <p style={{ margin:"0 0 12px", color:COLORS.muted }}>
          Stock actual: <b style={{ color:COLORS.text }}>{Number(producto.stock_actual || 0)}</b>
        </p>

        <label>Cantidad</label>
        <input
          type="number"
          value={cantidad}
          onChange={(e) => setCantidad(e.target.value)}
          onKeyDown={(e) => {
            if (e.key === "Enter") {
              onSave({ producto, tipo, cantidad });
            }
          }}
          style={{
            width:"100%",
            padding:"10px",
            borderRadius:8,
            border:`1px solid ${COLORS.border}`,
            background:COLORS.surface,
            color:COLORS.text,
            fontSize:16,
            boxSizing:"border-box",
            margin:"6px 0 16px"
          }}
        />

        <div style={{ display:"flex", gap:10, justifyContent:"flex-end" }}>
          <button onClick={onClose} style={{ padding:"9px 12px", borderRadius:8 }}>
            Cancelar
          </button>

          <button
            onClick={() => onSave({ producto, tipo, cantidad })}
            style={{
              padding:"9px 14px",
              borderRadius:8,
              border:"none",
              background: tipo === "entrada" ? COLORS.success : COLORS.danger,
              color:"#fff",
              fontWeight:700,
              cursor:"pointer"
            }}
          >
            Guardar {tipo}
          </button>
        </div>
      </div>
    </div>
  );
}

function HistorialInventarioModal({ producto, movimientos, onClose }) {
  useEffect(() => {
    const cerrarConEsc = (e) => {
      if (e.key === "Escape") onClose();
    };

    window.addEventListener("keydown", cerrarConEsc);

    return () => {
      window.removeEventListener("keydown", cerrarConEsc);
    };
  }, [onClose]);

  const historial = movimientos.filter(m => m.producto_id === producto.id);

  return (
    <div
      onClick={onClose}
      style={{
        position:"fixed",
        inset:0,
        background:"rgba(0,0,0,0.65)",
        display:"flex",
        alignItems:"flex-start",
        justifyContent:"center",
        zIndex:9999,
        overflowY:"auto",
        padding:"20px 10px"
      }}
    >
      <div
        onClick={(e) => e.stopPropagation()}
        style={{
          background:COLORS.card,
          border:`1px solid ${COLORS.border}`,
          borderRadius:14,
          padding:22,
          width:"520px",
          maxWidth:"92%",
          maxHeight:"90vh",
          overflowY:"auto",
          color:COLORS.text,
          boxSizing:"border-box"
        }}
      >
        <h2 style={{ margin:"0 0 8px", color:COLORS.accent }}>
          Historial
        </h2>

        <p style={{ margin:"0 0 14px" }}>
          <b>{producto.nombre}</b>
        </p>

        {historial.length === 0 ? (
          <p style={{ color:COLORS.muted }}>
            Este producto todavía no tiene movimientos.
          </p>
        ) : (
          <div style={{ display:"flex", flexDirection:"column", gap:8 }}>
            {historial.map((m) => (
              <div
                key={m.id}
                style={{
                  border:`1px solid ${COLORS.border}`,
                  borderRadius:10,
                  padding:10,
                  background:COLORS.surface
                }}
              >
                <div style={{ display:"flex", justifyContent:"space-between", gap:10 }}>
                  <b style={{
                    color: m.tipo === "entrada" ? COLORS.success : COLORS.danger
                  }}>
                    {m.tipo === "entrada" ? "Entrada" : "Salida"} {m.tipo === "entrada" ? "+" : "-"}{Number(m.cantidad || 0)}
                  </b>

                  <span style={{ color:COLORS.muted, fontSize:12 }}>
                    {new Date(m.created_at).toLocaleString("es-CL")}
                  </span>
                </div>

                <div style={{ color:COLORS.muted, fontSize:12, marginTop:5 }}>
                  Stock: {Number(m.stock_anterior || 0)} → {Number(m.stock_nuevo || 0)}
                </div>
              </div>
            ))}
          </div>
        )}

        <button
          onClick={onClose}
          style={{
            marginTop:16,
            width:"100%",
            padding:"9px 12px",
            borderRadius:8,
            border:"none",
            background:COLORS.accent,
            color:"#111",
            fontWeight:700,
            cursor:"pointer"
          }}
        >
          Cerrar
        </button>
      </div>
    </div>
  );
}
function NuevoProductoModal({ onClose, onSave }) {
  const [nombre, setNombre] = useState("");
  const [categoria, setCategoria] = useState("Tableros");
  const [unidad, setUnidad] = useState("planchas");
  const [stockActual, setStockActual] = useState("");
  const [stockMinimo, setStockMinimo] = useState("");
  const [proveedor, setProveedor] = useState("");
  const [esLaminado, setEsLaminado] = useState(false);
  const [observaciones, setObservaciones] = useState("");

  useEffect(() => {
    const cerrarConEsc = (e) => {
      if (e.key === "Escape") onClose();
    };

    window.addEventListener("keydown", cerrarConEsc);

    return () => {
      window.removeEventListener("keydown", cerrarConEsc);
    };
  }, [onClose]);

  const guardar = () => {
    if (!nombre.trim()) {
      alert("Debes ingresar el nombre del producto.");
      return;
    }

    onSave({
      nombre: nombre.trim().toUpperCase(),
      categoria,
      unidad,
      stock_actual: Number(stockActual) || 0,
      stock_minimo: Number(stockMinimo) || 0,
      proveedor: proveedor.trim(),
      es_laminado: esLaminado,
      activo: true,
      observaciones: observaciones.trim()
    });
  };

  const estiloInput = {
    width:"100%",
    padding:"9px 10px",
    borderRadius:8,
    border:`1px solid ${COLORS.border}`,
    background:COLORS.surface,
    color:COLORS.text,
    fontSize:16,
    boxSizing:"border-box"
  };

  return (
    <div
      onClick={onClose}
      style={{
        position:"fixed",
        inset:0,
        background:"rgba(0,0,0,0.65)",
        display:"flex",
        alignItems:"flex-start",
        justifyContent:"center",
        zIndex:9999,
        overflowY:"auto",
        padding:"20px 10px"
      }}
    >
      <div
        onClick={(e) => e.stopPropagation()}
        style={{
          background:COLORS.card,
          border:`1px solid ${COLORS.border}`,
          borderRadius:14,
          padding:22,
          width:"460px",
          maxWidth:"92%",
          maxHeight:"90vh",
          overflowY:"auto",
          color:COLORS.text,
          boxSizing:"border-box"
        }}
      >
        <h2 style={{ margin:"0 0 14px", color:COLORS.accent }}>
          Agregar producto
        </h2>

        <label>Nombre del producto</label>
        <input
          value={nombre}
          onChange={(e) => setNombre(e.target.value)}
          placeholder="Ej: MDF 18MM 122X244"
          style={{ ...estiloInput, margin:"6px 0 12px" }}
        />

        <label>Categoría</label>
        <select
          value={categoria}
          onChange={(e) => setCategoria(e.target.value)}
          style={{ ...estiloInput, margin:"6px 0 12px" }}
        >
          <option>Tableros</option>
          <option>Laminados</option>
          <option>Pegamentos</option>
          <option>Embalaje</option>
          <option>Tornillos</option>
          <option>Herrajes</option>
          <option>Insumos generales</option>
        </select>

        <label>Unidad</label>
        <select
          value={unidad}
          onChange={(e) => setUnidad(e.target.value)}
          style={{ ...estiloInput, margin:"6px 0 12px" }}
        >
          <option>planchas</option>
          <option>láminas</option>
          <option>tarros</option>
          <option>bolsas</option>
          <option>unidades</option>
          <option>rollos</option>
          <option>metros</option>
          <option>kilos</option>
          <option>litros</option>
        </select>

        <label>Stock actual</label>
        <input
          type="number"
          value={stockActual}
          onChange={(e) => setStockActual(e.target.value)}
          style={{ ...estiloInput, margin:"6px 0 12px" }}
        />

        <label>Stock mínimo</label>
        <input
          type="number"
          value={stockMinimo}
          onChange={(e) => setStockMinimo(e.target.value)}
          style={{ ...estiloInput, margin:"6px 0 12px" }}
        />

        <label>Proveedor habitual</label>
        <input
          value={proveedor}
          onChange={(e) => setProveedor(e.target.value)}
          placeholder="Ej: Imperial, Merino, Latam"
          style={{ ...estiloInput, margin:"6px 0 12px" }}
        />

        <label style={{ display:"flex", gap:8, alignItems:"center", margin:"8px 0 12px" }}>
          <input
            type="checkbox"
            checked={esLaminado}
            onChange={(e) => {
              setEsLaminado(e.target.checked);
              if (e.target.checked) {
                setCategoria("Laminados");
                setUnidad("láminas");
              }
            }}
          />
          Es laminado
        </label>

        <label>Observaciones</label>
        <textarea
          value={observaciones}
          onChange={(e) => setObservaciones(e.target.value)}
          style={{
            ...estiloInput,
            minHeight:70,
            resize:"none",
            margin:"6px 0 16px"
          }}
        />

        <div style={{ display:"flex", gap:10, justifyContent:"flex-end" }}>
          <button onClick={onClose} style={{ padding:"9px 12px", borderRadius:8 }}>
            Cancelar
          </button>

          <button
            onClick={guardar}
            style={{
              padding:"9px 14px",
              borderRadius:8,
              border:"none",
              background:COLORS.accent,
              color:"#111",
              fontWeight:700,
              cursor:"pointer"
            }}
          >
            Guardar producto
          </button>
        </div>
      </div>
    </div>
  );
}
export default function App() {
  const obtenerDetalleProduccionDesdeExcel = (workbook) => {
  const hoja = workbook.Sheets["CUBIERTA"];

  if (!hoja) return [];

  const filas = XLSX.utils.sheet_to_json(hoja, {
    header: 1,
    defval: ""
  });

  let inicioDetalle = -1;

  for (let i = 0; i < filas.length; i++) {
    const textoFila = filas[i].join(" ").toUpperCase();

    if (
      textoFila.includes("MATERIAL") &&
      textoFila.includes("CANTIDAD") &&
      textoFila.includes("DESCRIPCIÓN") &&
      textoFila.includes("COLOR")
    ) {
      inicioDetalle = i + 1;
      break;
    }
  }

  if (inicioDetalle === -1) return [];

  const detalles = [];

  for (let i = inicioDetalle; i < filas.length; i++) {
    const fila = filas[i];

    const material = String(fila[0] || "").trim();
    const cantidad = fila[1];
    const descripcion = String(fila[2] || "").trim();
    const alto = fila[3];
    const ancho = fila[4];
    const color = String(fila[5] || "").trim();

    const filaVacia =
      !material &&
      !cantidad &&
      !descripcion &&
      !alto &&
      !ancho &&
      !color;

    if (filaVacia) break;

    detalles.push({
      material,
      cantidad: Number(cantidad) || 0,
      descripcion,
      alto: Number(alto) || 0,
      ancho: Number(ancho) || 0,
      color,
      orden: detalles.length + 1
    });
  }

  return detalles;
};
  const importarCotizacion = async (event) => {
  const file = event.target.files[0];

  if (!file) return;

  const data = await file.arrayBuffer();
  const workbook = XLSX.read(data);
  const sheet = workbook.Sheets["RESUMEN"];
  const nombresHojas = workbook.SheetNames;

  const hojaDetalleNombre = nombresHojas.find(
  n =>
    n !== "RESUMEN" &&
    n !== "NOTA DE VENTA" &&
    n !== "PRODUCCION" &&
    n !== "SEGUIMIENTO"
);

  const hojaDetalle = workbook.Sheets[hojaDetalleNombre];

  const detalleJson = XLSX.utils.sheet_to_json(hojaDetalle, {
  header: 1,
  defval: ""
});
  let inicioDetalle = -1;

for (let i = 0; i < detalleJson.length; i++) {
  const filaTexto = detalleJson[i].join(" ").toUpperCase();

  if (
    filaTexto.includes("UNIDAD") &&
    filaTexto.includes("TIPO")
  ) {
    inicioDetalle = i + 1;
    break;
  }
}

const detalles = [];

if (inicioDetalle !== -1) {
  for (let i = inicioDetalle; i < detalleJson.length; i++) {
    const fila = detalleJson[i];

    const textoFila = fila.join(" ").toUpperCase();

    if (
      textoFila.includes("NETO") ||
      textoFila.includes("IVA") ||
      textoFila.includes("TOTAL")
    ) {
      break;
    }

    if (!fila[1]) continue;

    detalles.push({
  unidad: Number(fila[1]) || 0,
  tipo: String(fila[2] || ""),
  largo: Number(fila[3]) || 0,
  ancho: Number(fila[4]) || 0,
  color: String(fila[5] || ""),
  valor: Number(fila[6]) || 0,
  total: Number(fila[7]) || 0,
});
  }
}

  if (!sheet) {
    alert("No existe hoja RESUMEN");
    return;
  }

  const numero = sheet["A2"]?.v;
  const cliente = sheet["B2"]?.v;
  const fecha = sheet["C2"]?.v;
  const total = sheet["D2"]?.v;

  const ok = await importarCotizacionExcel({
  numero,
  cliente,
  fecha,
  total,
  detalles
});

  if (ok) {
    window.location.reload();
  }
};
  const importarExcel = async (event) => {
    const file = event.target.files[0]

    if (!file) return

    const data = await file.arrayBuffer()

    const workbook = XLSX.read(data)

    const sheet = workbook.Sheets['RESUMEN']

    if (!sheet) {
      alert('No existe hoja RESUMEN')
      return
    }
    

    const cotizacion = sheet['A2']?.v
    const notaVenta = sheet['B2']?.v
    const cliente = sheet['C2']?.v
    const fecha = sheet['D2']?.v
    const total = sheet['E2']?.v

    console.log({
      cotizacion,
      notaVenta,
      cliente,
      fecha,
      total
    })
    
    const detallesProduccion = obtenerDetalleProduccionDesdeExcel(workbook);

const ok = await importarNotaVentaExcel({
  cotizacion,
  notaVenta,
  cliente,
  fecha,
  total
});

if (ok) {
  await supabase
    .from("detalles_notas_venta_produccion")
    .delete()
    .eq("nota_venta_numero", String(notaVenta));

  if (detallesProduccion.length > 0) {
    const detallesParaGuardar = detallesProduccion.map((d) => ({
      nota_venta_numero: String(notaVenta),
      cotizacion_numero: String(cotizacion || ""),
      cliente: String(cliente || ""),
      material: d.material,
      cantidad: d.cantidad,
      descripcion: d.descripcion,
      alto: d.alto,
      ancho: d.ancho,
      color: d.color,
      orden: d.orden
    }));

    const { error: errorDetalles } = await supabase
      .from("detalles_notas_venta_produccion")
      .insert(detallesParaGuardar);

    if (errorDetalles) {
      console.error(errorDetalles);
      alert("La nota se importó, pero hubo un error guardando el detalle de producción.");
      return;
    }
  }

  window.location.reload();
}
  }
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
    const detallesData = await obtenerDetallesCotizaciones();
    setDetallesCotizaciones(detallesData);

    const dataNotas = await obtenerNotasVenta();

    const notasFormateadas = dataNotas.map((n) => ({
      id: `supabase-${n.id}`,
      numero: String(n.numero),
      cliente: n.cotizaciones?.cliente || n.cliente || "(sin cliente)",
      fecha: n.fecha?.split("T")[0] || new Date().toISOString().split("T")[0],
     total: n.cotizaciones?.total || n.total || 0,
      cotizacion: n.cotizaciones?.numero ? String(n.cotizaciones.numero) : null,
      abono: n.abono || 0,
      estado_pago: n.estado_pago || "pendiente",
      materiales: n.materiales || "falta",
      proceso: n.proceso || "en espera",
      observaciones: n.observaciones || "",
      fecha_entrega_estimada: n.fecha_entrega_estimada || "",
      produccion_observaciones: n.produccion_observaciones || "",
      mdf_cortado: n.mdf_cortado || false,
      lamina_cortada: n.lamina_cortada || false,
      tupizado: n.tupizado || false,
      armado: n.armado || false,
      pegado: n.pegado || false,
      postformado: n.postformado || false,
}));

    setNotas(notasFormateadas);
    const { data: dataAbonos, error: errorAbonos } = await supabase
  .from("abonos_nv")
  .select("*")
  .order("fecha", { ascending: false });

if (!errorAbonos) {
  setAbonosNV(dataAbonos || []);
}
const { data: dataDetallesNV, error: errorDetallesNV } = await supabase
  .from("detalles_notas_venta_produccion")
  .select("*")
  .order("orden", { ascending: true });

if (!errorDetallesNV) {
  setDetallesNotasVenta(dataDetallesNV || []);
}
const { data: dataInventario, error: errorInventario } = await supabase
  .from("inventario_productos")
  .select("*")
  .order("nombre", { ascending: true });

if (!errorInventario) {
  setProductosInventario(dataInventario || []);
}
const { data: dataMovimientosInventario, error: errorMovimientosInventario } = await supabase
  .from("inventario_movimientos")
  .select("*")
  .order("created_at", { ascending: false });

if (!errorMovimientosInventario) {
  setMovimientosInventario(dataMovimientosInventario || []);
}
const { data: dataSobrantesLaminados, error: errorSobrantesLaminados } = await supabase
  .from("inventario_laminado_sobrantes")
  .select("*")
  .order("created_at", { ascending: false });

if (!errorSobrantesLaminados) {
  setSobrantesLaminados(dataSobrantesLaminados || []);
}
  }

  cargarDatos();
}, []);
  const [tab, setTab] = useState("dashboard");
  const [filter, setFilter] = useState("");
  const [fechaDesde, setFechaDesde] = useState("");
  const [fechaHasta, setFechaHasta] = useState("");
  const [mesFiltro, setMesFiltro] = useState("");
  const [showVencidas, setShowVencidas] = useState(false);
  const [seguimiento, setSeguimiento] = useState({});
  const [cotizaciones, setCotizaciones] = useState([]);
  const [detallesCotizaciones, setDetallesCotizaciones] = useState([]);
  const [detallesNotasVenta, setDetallesNotasVenta] = useState([]);
  const [productosInventario, setProductosInventario] = useState([]);
  const [sobrantesLaminados, setSobrantesLaminados] = useState([]);
  const [modalNuevoProducto, setModalNuevoProducto] = useState(false);
  const [movimientosInventario, setMovimientosInventario] = useState([]);
  const [modalMovimientoInventario, setModalMovimientoInventario] = useState(null);
  const [modalHistorialProducto, setModalHistorialProducto] = useState(null);
  const [modalEditarProducto, setModalEditarProducto] = useState(null);
  const [busquedaInventario, setBusquedaInventario] = useState("");
  const [filtroCategoriaInventario, setFiltroCategoriaInventario] = useState("todos");
  const [previewOC, setPreviewOC] = useState([]);
  const [modalPreviewOC, setModalPreviewOC] = useState(false);
  const [modalSobrantesLaminado, setModalSobrantesLaminado] = useState(null);
  const [busquedaLaminadoNombre, setBusquedaLaminadoNombre] = useState("");
  const [busquedaSobranteLargo, setBusquedaSobranteLargo] = useState("");
  const [busquedaSobranteAncho, setBusquedaSobranteAncho] = useState("");
  const [notas, setNotas] = useState([]);
  const [abonosNV, setAbonosNV] = useState([]);
  const [modalCot, setModalCot] = useState(null);
  const [modalNV, setModalNV] = useState(null);
  const [modalProduccion, setModalProduccion] = useState(null);
  const [filtroProduccion, setFiltroProduccion] = useState("todos");
  const [busquedaCliente, setBusquedaCliente] = useState("")
  

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
  const cotizacionesFiltradas = withStatus.filter((q) =>
  q.cliente
    ?.toLowerCase()
    .includes(busquedaCliente.toLowerCase())
);
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

  const cumpleFecha = (fecha) => {
  if (!fecha) return true;

  const f = new Date(fecha);
  const desde = fechaDesde ? new Date(fechaDesde) : null;
  const hasta = fechaHasta ? new Date(fechaHasta) : null;

  if (desde && f < desde) return false;
  if (hasta && f > hasta) return false;

  return true;
};

const cumpleMes = (fecha) => {
  if (!mesFiltro) return true;
  if (!fecha) return false;

  const fechaObj = new Date(fecha);
  const year = fechaObj.getFullYear();
  const month = String(fechaObj.getMonth() + 1).padStart(2, "0");

  return `${year}-${month}` === mesFiltro;
};
  const limpiarFiltros = () => {
  setSearch("");
  setFechaDesde("");
  setFechaHasta("");
  setMesFiltro("");
};

const filteredQuotes = withStatus.filter(q =>
  (q.numero.toLowerCase().includes(filter.toLowerCase()) || q.cliente.toLowerCase().includes(filter.toLowerCase())) &&
  cumpleFecha(q.fecha) &&
  cumpleMes(q.fecha) &&
  (showVencidas || q.status !== "vencida")
).sort((a,b) => { 
  const o={urgente:0,activa:1,vendida:2,vencida:3}; 
  return o[a.status]-o[b.status]; 
});

const filteredNotas = notas.filter(s =>
  (s.numero.toLowerCase().includes(filter.toLowerCase()) || s.cliente.toLowerCase().includes(filter.toLowerCase())) &&
  cumpleFecha(s.fecha) &&
  cumpleMes(s.fecha)
);

  const r2=58,cx=80,cy=76;
  const toRad=d=>d*Math.PI/180;
  const startA=200,sweepA=140;
  const arcPath=(s,e)=>{
    const p1={x:cx+r2*Math.cos(toRad(s)),y:cy+r2*Math.sin(toRad(s))};
    const p2={x:cx+r2*Math.cos(toRad(e)),y:cy+r2*Math.sin(toRad(e))};
    return `M ${p1.x} ${p1.y} A ${r2} ${r2} 0 ${e-s>180?1:0} 1 ${p2.x} ${p2.y}`;
  };
  const fillEnd=startA+sweepA*(rate/100);
  const guardarGestionNV = async (nvActualizada) => {
  const idReal = Number(String(nvActualizada.id).replace("supabase-", ""));

  const nuevoAbono = Number(nvActualizada.nuevoAbono) || 0;

  if (nuevoAbono > 0) {
    const { error: errorInsert } = await supabase
      .from("abonos_nv")
      .insert({
        nota_venta_id: idReal,
        monto: nuevoAbono,
        fecha: new Date().toISOString().split("T")[0],
        observacion: nvActualizada.observacionAbono || ""
      });

    if (errorInsert) {
      alert("Error al guardar el abono");
      console.error(errorInsert);
      return;
    }
  }

  const { data: abonosActualizados, error: errorAbonos } = await supabase
    .from("abonos_nv")
    .select("*")
    .eq("nota_venta_id", idReal);

  if (errorAbonos) {
    alert("Error al calcular abonos");
    console.error(errorAbonos);
    return;
  }

  const totalAbonado = (abonosActualizados || []).reduce(
    (sum, a) => sum + Number(a.monto || 0),
    0
  );

  const total = Number(nvActualizada.total) || 0;

  let estadoPago = "pendiente";

  if (totalAbonado >= total && total > 0) {
    estadoPago = "pagada";
  } else if (totalAbonado > 0) {
    estadoPago = "abonada";
  }

  const { error } = await supabase
    .from("notas_venta")
    .update({
      abono: totalAbonado,
      estado_pago: estadoPago,
      materiales: nvActualizada.materiales,
      proceso: nvActualizada.proceso,
      observaciones: nvActualizada.observaciones,
    })
    .eq("id", idReal);

  if (error) {
    alert("Error al guardar gestión de NV");
    console.error(error);
    return;
  }

  setAbonosNV(prev => {
    const otros = prev.filter(a => a.nota_venta_id !== idReal);
    return [...otros, ...(abonosActualizados || [])];
  });

  setNotas(notas.map(n =>
    n.id === nvActualizada.id
      ? { ...nvActualizada, abono: totalAbonado, estado_pago: estadoPago }
      : n
  ));

  setModalNV(null);
};
  
const guardarProduccion = async (notaActualizada) => {
  const idReal = Number(String(notaActualizada.id).replace("supabase-", ""));

  const { error } = await supabase
    .from("notas_venta")
    .update({
      fecha_entrega_estimada: notaActualizada.fecha_entrega_estimada || null,
      produccion_observaciones: notaActualizada.produccion_observaciones || "",
      mdf_cortado: notaActualizada.mdf_cortado,
      lamina_cortada: notaActualizada.lamina_cortada,
      tupizado: notaActualizada.tupizado,
      armado: notaActualizada.armado,
      pegado: notaActualizada.pegado,
      postformado: notaActualizada.postformado,
    })
    .eq("id", idReal);

  if (error) {
    alert("Error al guardar producción");
    console.error(error);
    return;
  }

  setNotas(notas.map(n =>
    n.id === notaActualizada.id
      ? { ...n, ...notaActualizada }
      : n
  ));

  setModalProduccion(null);
};
const fechaLocal = (fecha) => {
  if (!fecha) return null;

  const [year, month, day] = fecha.split("-").map(Number);
  return new Date(year, month - 1, day);
};

const estaAtrasada = (fecha) => {
  if (!fecha) return false;

  const hoy = new Date();
  hoy.setHours(0, 0, 0, 0);

  const entrega = fechaLocal(fecha);
  entrega.setHours(0, 0, 0, 0);

  return entrega < hoy;
};

const venceHoy = (fecha) => {
  if (!fecha) return false;

  const hoy = new Date();
  hoy.setHours(0, 0, 0, 0);

  const entrega = fechaLocal(fecha);
  entrega.setHours(0, 0, 0, 0);

  return entrega.getTime() === hoy.getTime();
};
const venceManana = (fecha) => {
  if (!fecha) return false;

  const manana = new Date();
  manana.setDate(manana.getDate() + 1);
  manana.setHours(0, 0, 0, 0);

  const entrega = fechaLocal(fecha);
  entrega.setHours(0, 0, 0, 0);

  return entrega.getTime() === manana.getTime();
};
const estadoProduccion = (n) => {
  if (estaAtrasada(n.fecha_entrega_estimada) && !n.postformado) {
    return {
      texto: "🔴 Atrasada",
      color: COLORS.danger
    };
  }

  if (n.postformado) {
    return {
      texto: "🟢 Lista para entregar",
      color: COLORS.success
    };
  }

  if (
    n.mdf_cortado ||
    n.lamina_cortada ||
    n.tupizado ||
    n.armado ||
    n.pegado
  ) {
    return {
      texto: "🔵 En proceso",
      color: "#60a5fa"
    };
  }

  return {
    texto: "🟡 Sin iniciar",
    color: COLORS.warning
  };
};
const coincideFiltroProduccion = (n, filtro) => {
  const estado = estadoProduccion(n).texto;

  if (filtro === "todos") return true;
  if (filtro === "atrasadas") return estado.includes("Atrasada");
  if (filtro === "hoy") return venceHoy(n.fecha_entrega_estimada) && !n.postformado;
  if (filtro === "manana") return venceManana(n.fecha_entrega_estimada) && !n.postformado;
  if (filtro === "proceso") return estado.includes("En proceso");
  if (filtro === "sin_iniciar") return estado.includes("Sin iniciar");
  if (filtro === "listas") return estado.includes("Lista");

  return true;
};

const prioridadProduccion = (n) => {
  if (estaAtrasada(n.fecha_entrega_estimada) && !n.postformado) return 1;
  if (venceHoy(n.fecha_entrega_estimada) && !n.postformado) return 2;
  if (venceManana(n.fecha_entrega_estimada) && !n.postformado) return 3;
  if (estadoProduccion(n).texto.includes("En proceso")) return 4;
  if (estadoProduccion(n).texto.includes("Sin iniciar")) return 5;
  if (estadoProduccion(n).texto.includes("Lista")) return 6;

  return 7;
};
const estadoStockInventario = (producto) => {
  const stock = Number(producto.stock_actual || 0);
  const minimo = Number(producto.stock_minimo || 0);

  if (stock <= 0) {
    return {
      texto: "🔴 Sin stock",
      color: COLORS.danger
    };
  }

  if (minimo > 0 && stock <= minimo) {
    return {
      texto: "🟡 Bajo stock",
      color: COLORS.warning
    };
  }

  return {
    texto: "🟢 OK",
    color: COLORS.success
  };
};
const productosInventarioFiltrados = productosInventario.filter((p) => {
  const textoBusqueda = busquedaInventario.trim().toLowerCase();

  const coincideBusqueda =
    !textoBusqueda ||
    String(p.nombre || "").toLowerCase().includes(textoBusqueda);

  const coincideCategoria =
    filtroCategoriaInventario === "todos" ||
    String(p.categoria || "").toLowerCase() === filtroCategoriaInventario;

  return p.activo !== false && coincideBusqueda && coincideCategoria;
});
const sobrantesGlobalesCompatibles = sobrantesLaminados
  .filter(s => !s.usado)
  .map(s => {
    const producto = productosInventario.find(p => p.id === s.producto_id);
    return {
      ...s,
      producto_nombre: producto?.nombre || "Laminado sin nombre"
    };
  })
  .filter(s => {
    const nombreBuscado = busquedaLaminadoNombre.trim().toLowerCase();
    const largoNecesario = Number(busquedaSobranteLargo) || 0;
    const anchoNecesario = Number(busquedaSobranteAncho) || 0;

    const coincideNombre =
      !nombreBuscado ||
      String(s.producto_nombre || "").toLowerCase().includes(nombreBuscado);

    const coincideMedida =
      largoNecesario > 0 &&
      anchoNecesario > 0 &&
      Number(s.largo) >= largoNecesario &&
      Number(s.ancho) >= anchoNecesario;

    return coincideNombre && coincideMedida;
  });
const guardarNuevoProducto = async (producto) => {
  const { data, error } = await supabase
    .from("inventario_productos")
    .insert([producto])
    .select();

  if (error) {
    console.error(error);
    alert("Error al guardar el producto.");
    return;
  }

  setProductosInventario(prev => [...prev, data[0]].sort((a,b) => a.nombre.localeCompare(b.nombre)));
  setModalNuevoProducto(false);
};
const guardarMovimientoInventario = async ({ producto, tipo, cantidad }) => {
  const cantidadNumero = Number(cantidad);

  if (!cantidadNumero || cantidadNumero <= 0) {
    alert("Debes ingresar una cantidad válida.");
    return;
  }

  const stockAnterior = Number(producto.stock_actual || 0);

  if (tipo === "salida" && cantidadNumero > stockAnterior) {
    alert("No puedes descontar más stock del disponible.");
    return;
  }

  const stockNuevo = tipo === "entrada"
    ? stockAnterior + cantidadNumero
    : stockAnterior - cantidadNumero;

  const { error: errorUpdate } = await supabase
    .from("inventario_productos")
    .update({ stock_actual: stockNuevo })
    .eq("id", producto.id);

  if (errorUpdate) {
    console.error(errorUpdate);
    alert("Error al actualizar el stock.");
    return;
  }

  const movimiento = {
    producto_id: producto.id,
    tipo,
    cantidad: cantidadNumero,
    stock_anterior: stockAnterior,
    stock_nuevo: stockNuevo,
    origen: "manual"
  };

  const { data, error: errorMovimiento } = await supabase
    .from("inventario_movimientos")
    .insert([movimiento])
    .select();

  if (errorMovimiento) {
    console.error(errorMovimiento);
    alert("El stock cambió, pero hubo un error guardando el historial.");
    return;
  }

  setProductosInventario(prev =>
    prev.map(p =>
      p.id === producto.id ? { ...p, stock_actual: stockNuevo } : p
    )
  );

  setMovimientosInventario(prev => [data[0], ...prev]);
  setModalMovimientoInventario(null);
};
const guardarSobrantesLaminado = async ({ producto, sobrantes }) => {
  const sobrantesValidos = sobrantes
    .filter(s => Number(s.largo) > 0 && Number(s.ancho) > 0)
    .map(s => ({
      producto_id: producto.id,
      largo: Number(s.largo),
      ancho: Number(s.ancho),
      observaciones: "",
      usado: false
    }));

  if (sobrantesValidos.length === 0) {
    alert("Debes ingresar al menos un sobrante válido.");
    return;
  }

  const { data, error } = await supabase
    .from("inventario_laminado_sobrantes")
    .insert(sobrantesValidos)
    .select();

  if (error) {
    console.error(error);
    alert("Error al guardar los sobrantes.");
    return;
  }

  setSobrantesLaminados(prev => [...(data || []), ...prev]);
  setModalSobrantesLaminado(null);
};

const marcarSobranteUsado = async (sobrante) => {
  const { data, error } = await supabase
    .from("inventario_laminado_sobrantes")
    .update({ usado: true })
    .eq("id", sobrante.id)
    .select();

  if (error) {
    console.error(error);
    alert("Error al marcar el sobrante como usado.");
    return;
  }

  setSobrantesLaminados(prev =>
    prev.map(s => s.id === sobrante.id ? data[0] : s)
  );
};
const guardarEdicionProducto = async (productoEditado) => {
  const { data, error } = await supabase
    .from("inventario_productos")
    .update({
      nombre: productoEditado.nombre,
      categoria: productoEditado.categoria,
      unidad: productoEditado.unidad,
      stock_minimo: productoEditado.stock_minimo,
      proveedor: productoEditado.proveedor,
      observaciones: productoEditado.observaciones,
      es_laminado: productoEditado.es_laminado
    })
    .eq("id", productoEditado.id)
    .select();

  if (error) {
    console.error(error);
    alert("Error al editar el producto.");
    return;
  }

  setProductosInventario(prev =>
    prev
      .map(p => p.id === productoEditado.id ? data[0] : p)
      .sort((a,b) => a.nombre.localeCompare(b.nombre))
  );

  setModalEditarProducto(null);
};
const importarOrdenCompraExcel = async (e) => {
  const file = e.target.files?.[0];
  if (!file) return;

  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: "array" });

  const hojaNombre = workbook.SheetNames[0];
  const hoja = workbook.Sheets[hojaNombre];

  const filas = XLSX.utils.sheet_to_json(hoja, {
    header: 1,
    defval: ""
  });

  let filaEncabezado = -1;
  let colCantidad = -1;
  let colProducto = -1;

  for (let i = 0; i < filas.length; i++) {
    const fila = filas[i].map(c => String(c).trim().toUpperCase());

    const cantidadIndex = fila.findIndex(c => c.includes("CANTIDAD"));
    const productoIndex = fila.findIndex(c => c.includes("PRODUCTO"));

    if (cantidadIndex !== -1 && productoIndex !== -1) {
      filaEncabezado = i;
      colCantidad = cantidadIndex;
      colProducto = productoIndex;
      break;
    }
  }

  if (filaEncabezado === -1) {
    alert("No se encontró la tabla de productos en la orden de compra.");
    return;
  }

  const productosEncontrados = [];

  for (let i = filaEncabezado + 1; i < filas.length; i++) {
    const fila = filas[i];

    const cantidad = Number(fila[colCantidad]) || 0;
    const nombre = String(fila[colProducto] || "").trim().toUpperCase();

    if (!nombre && !cantidad) break;

    if (!nombre || cantidad <= 0) continue;

    productosEncontrados.push({
      nombre,
      cantidad
    });
  }

  if (productosEncontrados.length === 0) {
    alert("No se encontraron productos válidos en la orden de compra.");
    return;
  }

  const agrupados = productosEncontrados.reduce((acc, item) => {
    const existente = acc.find(p => p.nombre === item.nombre);

    if (existente) {
      existente.cantidad += item.cantidad;
    } else {
      acc.push({ ...item });
    }

    return acc;
  }, []);

  setPreviewOC(agrupados);
  setModalPreviewOC(true);

  e.target.value = "";
};
const detectarCategoriaProducto = (nombre) => {
  const n = nombre.toUpperCase();

  if (n.includes("MDF") || n.includes("MELAMIL")) return "Tableros";
  if (n.includes("LAM") || n.includes("LAMINADO")) return "Laminados";
  if (n.includes("AGOREX") || n.includes("COLA")) return "Pegamentos";
  if (n.includes("FILM") || n.includes("CARTON") || n.includes("CARTÓN")) return "Embalaje";
  if (n.includes("TORNILLO")) return "Tornillos";
  if (n.includes("RIEL") || n.includes("TIRADOR") || n.includes("BISAGRA")) return "Herrajes";

  return "Insumos generales";
};

const detectarUnidadProducto = (nombre) => {
  const n = nombre.toUpperCase();

  if (n.includes("MDF") || n.includes("MELAMIL")) return "planchas";
  if (n.includes("LAM") || n.includes("LAMINADO")) return "láminas";
  if (n.includes("AGOREX")) return "tarros";
  if (n.includes("COLA")) return "bolsas";
  if (n.includes("FILM")) return "rollos";
  if (n.includes("CARTON") || n.includes("CARTÓN")) return "rollos";

  return "unidades";
};

const confirmarImportacionOC = async () => {
  if (previewOC.length === 0) return;

  for (const item of previewOC) {
    const nombre = item.nombre.trim().toUpperCase();
    const cantidad = Number(item.cantidad) || 0;

    const productoExistente = productosInventario.find(
      p => String(p.nombre || "").trim().toUpperCase() === nombre
    );

    if (productoExistente) {
      const stockAnterior = Number(productoExistente.stock_actual || 0);
      const stockNuevo = stockAnterior + cantidad;

      const { data, error } = await supabase
        .from("inventario_productos")
        .update({ stock_actual: stockNuevo })
        .eq("id", productoExistente.id)
        .select();

      if (error) {
        console.error(error);
        alert(`Error actualizando ${nombre}`);
        return;
      }

      await supabase
        .from("inventario_movimientos")
        .insert([{
          producto_id: productoExistente.id,
          tipo: "entrada",
          cantidad,
          stock_anterior: stockAnterior,
          stock_nuevo: stockNuevo,
          origen: "orden_compra"
        }]);

      setProductosInventario(prev =>
        prev.map(p => p.id === productoExistente.id ? data[0] : p)
      );

    } else {
      const categoria = detectarCategoriaProducto(nombre);
      const unidad = detectarUnidadProducto(nombre);
      const esLaminado = categoria === "Laminados";

      const { data, error } = await supabase
        .from("inventario_productos")
        .insert([{
          nombre,
          categoria,
          unidad,
          stock_actual: cantidad,
          stock_minimo: 0,
          proveedor: "",
          es_laminado: esLaminado,
          activo: true
        }])
        .select();

      if (error) {
        console.error(error);
        alert(`Error creando ${nombre}`);
        return;
      }

      const nuevoProducto = data[0];

      await supabase
        .from("inventario_movimientos")
        .insert([{
          producto_id: nuevoProducto.id,
          tipo: "entrada",
          cantidad,
          stock_anterior: 0,
          stock_nuevo: cantidad,
          origen: "orden_compra"
        }]);

      setProductosInventario(prev =>
        [...prev, nuevoProducto].sort((a,b) => a.nombre.localeCompare(b.nombre))
      );
    }
  }

  setPreviewOC([]);
  setModalPreviewOC(false);
  alert("Orden de compra importada correctamente.");
};
  const tabs=[
    {key:"dashboard",label:"📊 Resumen"},
    {key:"quotes",label:`📋 Cotizaciones (${cotizaciones.length})`},
    {key:"sales",label:`✅ Notas de Venta (${notas.length})`},
    {key:"sinmatch",label:`⚠ Sin cruzar (${sinCotizacion.length})`},
    {key:"produccion",label:`🏭 Producción`},
    {key:"inventario",label:`📦 Inventario`},
  ];

  return (
    <div style={{ minHeight:"100vh", background:COLORS.bg, color:COLORS.text, fontFamily:"'Trebuchet MS',sans-serif", paddingBottom:60 }}>
 <div style={{ margin:20, padding:16, background:COLORS.card, border:`1px solid ${COLORS.border}`, borderRadius:12, display:"flex", gap:10, flexWrap:"wrap" }}>
  

  <label
  style={{
    padding: "10px 18px",
    background: COLORS.success,
    border: "none",
    borderRadius: 8,
    fontWeight: 700,
    cursor: "pointer",
    color: "#fff"
  }}
>
  Importar NV
  <input
    type="file"
    accept=".xlsx,.xls"
    onChange={importarExcel}
    style={{ display: "none" }}
  />
</label>
<label
  style={{
    padding: "10px 18px",
    background: COLORS.accent,
    borderRadius: 8,
    fontWeight: 700,
    cursor: "pointer",
    color: "#fff",
    display: "inline-block",
    marginLeft: 10
  }}
>
  Importar Cotización
  <input
    type="file"
    accept=".xlsx,.xls"
    onChange={importarCotizacion}
    style={{ display: "none" }}
  />
</label>
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
              <input
  type="text"
  placeholder="Buscar cliente..."
  value={busquedaCliente}
  onChange={(e) => setBusquedaCliente(e.target.value)}
  style={{
    padding: "10px",
    marginBottom: "15px",
    width: "250px",
    borderRadius: "8px",
    border: "1px solid #444"
  }}
/>
              <h2 style={{ margin:0, fontFamily:"Georgia,serif", color:COLORS.accent, fontSize:17 }}>Cotizaciones Mayo 2026</h2>
              <label style={{ display:"flex", alignItems:"center", gap:8, fontSize:12, color:COLORS.muted, cursor:"pointer" }}>
                <input type="checkbox" checked={showVencidas} onChange={e=>setShowVencidas(e.target.checked)}/> Mostrar vencidas ({vencidas.length})
              </label>
            </div>
            <input placeholder="Buscar por número o cliente..." value={filter} onChange={e=>setFilter(e.target.value)}
              style={{ width:"100%", boxSizing:"border-box", marginBottom:12, background:COLORS.surface, border:`1px solid ${COLORS.border}`, borderRadius:8, padding:"9px 14px", color:COLORS.text, fontSize:13, outline:"none" }}/>
            <div style={{ display:"flex", gap:10, marginBottom:12, flexWrap:"wrap" }}>
  <input
    type="date"
    value={fechaDesde}
    onChange={e=>setFechaDesde(e.target.value)}
    style={{ background:COLORS.surface, border:`1px solid ${COLORS.border}`, borderRadius:8, padding:"9px 14px", color:COLORS.text }}
  />

  <input
    type="date"
    value={fechaHasta}
    onChange={e=>setFechaHasta(e.target.value)}
    style={{ background:COLORS.surface, border:`1px solid ${COLORS.border}`, borderRadius:8, padding:"9px 14px", color:COLORS.text }}
  />
  <select
  value={mesFiltro}
  onChange={(e) => setMesFiltro(e.target.value)}
  style={{
    background:COLORS.surface,
    border:`1px solid ${COLORS.border}`,
    borderRadius:8,
    padding:"9px 14px",
    color:COLORS.text
  }}
>
  <option value="">Todos los meses</option>

  {[...new Set([...cotizaciones, ...notas]
    .map(item => item.fecha)
    .filter(Boolean)
    .map(fecha => {
      const f = new Date(fecha);
      const year = f.getFullYear();
      const month = String(f.getMonth() + 1).padStart(2, "0");
      return `${year}-${month}`;
    })
  )]
    .sort((a, b) => b.localeCompare(a))
    .map(mes => {
      const [year, month] = mes.split("-");
      const nombreMes = new Date(year, Number(month) - 1)
        .toLocaleDateString("es-CL", { month:"long", year:"numeric" });

      return (
        <option key={mes} value={mes}>
          {nombreMes.charAt(0).toUpperCase() + nombreMes.slice(1)}
        </option>
      );
    })}
</select>
<button
  onClick={limpiarFiltros}
  style={{
    background:COLORS.danger,
    color:"#fff",
    border:"none",
    borderRadius:8,
    padding:"9px 14px",
    cursor:"pointer",
    fontWeight:700
  }}
>
  Limpiar filtros
</button>
</div>
            <div style={{ display:"flex", flexDirection:"column", gap:7 }}>
              {filteredQuotes.map(q=>{
                const nvs=notas.filter(s=>s.cotizacion===q.numero);
                const seg=seguimiento[q.numero]||[];
                const last=seg[seg.length-1];
                return (
                  <div key={q.id} style={{ background:COLORS.card, border:`1px solid ${STATUS_CONFIG[q.status].border}`, borderLeft:`4px solid ${LEFT_COLOR[q.status]}`, borderRadius:10, padding:"11px 14px", opacity:q.status==="vencida"?0.55:1 }}>
                    <div style={{ display:"flex", alignItems:"center", justifyContent:"space-between", flexWrap:"wrap", gap:8 }}>
                      <div>
                        <span
  onClick={() => setModalCot(q)}
  style={{
    fontWeight:700,
    color:COLORS.accent,
    marginRight:8,
    cursor:"pointer"
  }}
>
  #{q.numero}
</span>

<span
  onClick={() => setModalCot(q)}
  style={{
    color:
      q.status === "vendida"
        ? COLORS.success
        : q.status === "vencida"
        ? COLORS.danger
        : q.status === "urgente"
        ? COLORS.warning
        : COLORS.text,
    cursor:"pointer",
    textDecoration:"underline"
  }}
>
  {q.cliente}
</span>
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
              {filteredNotas.map(s=>(
                <div
                key={s.id}
                onClick={() => setModalNV(s)}
                style={{background:COLORS.card, border:`1px solid ${s.cotizacion?"#2d5040":COLORS.border}`, borderLeft:`4px solid ${s.cotizacion?COLORS.success:COLORS.warning}`, borderRadius:10, padding:"11px 14px", display:"flex", alignItems:"center", justifyContent:"space-between", flexWrap:"wrap", gap:8 }}>
                  <div>
  <div>
    <span style={{ fontWeight:700, color:COLORS.success, marginRight:8 }}>NV#{s.numero}</span>
    <span style={{ color:COLORS.text }}>{s.cliente}</span>
    {s.cotizacion ? <span style={{ marginLeft:8, fontSize:11, color:COLORS.muted, background:COLORS.subtle, borderRadius:4, padding:"2px 7px" }}>← COT#{s.cotizacion}</span>
      : <span style={{ marginLeft:8, fontSize:11, color:COLORS.warning, background:"#2a1f0a", borderRadius:4, padding:"2px 7px", border:`1px solid ${COLORS.warning}` }}>sin cotización</span>}
    <span style={{ color:COLORS.muted, marginLeft:8, fontSize:11 }}>{s.fecha}</span>
  </div>

  <div style={{ display:"flex", gap:6, flexWrap:"wrap", marginTop:7 }}>
    <span style={{
      padding:"3px 8px",
      borderRadius:999,
      fontSize:11,
      fontWeight:700,
      background:
        s.estado_pago === "pagada"
          ? "rgba(34,197,94,.18)"
          : s.estado_pago === "abonada"
          ? "rgba(250,204,21,.18)"
          : "rgba(239,68,68,.18)",
      color:
        s.estado_pago === "pagada"
          ? "#4ade80"
          : s.estado_pago === "abonada"
          ? "#fde047"
          : "#f87171"
    }}>
      {s.estado_pago === "pagada"
        ? "🟢 Pagada"
        : s.estado_pago === "abonada"
        ? "🟡 Abonada"
        : "🔴 Pendiente"}
    </span>

    <span style={{
      padding:"3px 8px",
      borderRadius:999,
      fontSize:11,
      fontWeight:700,
      background:"rgba(59,130,246,.15)",
      color:"#60a5fa"
    }}>
      {s.proceso === "en espera" && "⏳ En espera"}
      {s.proceso === "en producción" && "📦 En producción"}
      {s.proceso === "terminado" && "✅ Terminado"}
      {s.proceso === "entregado" && "🚚 Entregado"}
    </span>

    <span style={{
      padding:"3px 8px",
      borderRadius:999,
      fontSize:11,
      fontWeight:700,
      background:
        s.materiales === "comprados"
          ? "rgba(34,197,94,.18)"
          : "rgba(239,68,68,.18)",
      color:
        s.materiales === "comprados"
          ? "#4ade80"
          : "#f87171"
    }}>
      {s.materiales === "comprados"
        ? "✅ Materiales"
        : "❌ Falta material"}
    </span>
  </div>
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
  <button
  onClick={async () => {
    const confirmar = confirm(`¿Eliminar NV ${s.numero}?`);

    if (!confirmar) return;

    const ok = await eliminarNotaVenta(s.id);

    if (ok) {
      window.location.reload();
    }
  }}
  style={{
    background: COLORS.danger,
    border: "none",
    borderRadius: 6,
    padding: "4px 8px",
    color: "#fff",
    cursor: "pointer",
    fontSize: 11
  }}
>
  Eliminar
</button>
</div>
                </div>
              ))}
            </div>
          </div>
        )}
{tab==="produccion" && (
  <div>
    <h2 style={{ margin:"0 0 14px", fontFamily:"Georgia,serif", color:COLORS.accent, fontSize:17 }}>
      🏭 Producción
    </h2>
<div style={{ display:"flex", gap:8, flexWrap:"wrap", marginBottom:14 }}>
  {[
    ["todos", "Todos"],
    ["atrasadas", "Atrasadas"],
    ["hoy", "Para hoy"],
    ["manana", "Para mañana"],
    ["proceso", "En proceso"],
    ["sin_iniciar", "Sin iniciar"],
    ["listas", "Listas"]
  ].map(([valor, texto]) => (
    <button
      key={valor}
      onClick={() => setFiltroProduccion(valor)}
      style={{
        border:`1px solid ${filtroProduccion === valor ? COLORS.accent : COLORS.border}`,
        background:filtroProduccion === valor ? COLORS.accent : COLORS.card,
        color:filtroProduccion === valor ? "#111" : COLORS.text,
        borderRadius:999,
        padding:"6px 10px",
        cursor:"pointer",
        fontSize:12,
        fontWeight:700
      }}
    >
      {texto}
    </button>
  ))}
</div>
    <div style={{ display:"grid", gap:12 }}>
      {notas
  .filter(n => n.proceso !== "entregado")
  .filter(n => coincideFiltroProduccion(n, filtroProduccion))
  .sort((a,b) => {
    const prioridadA = prioridadProduccion(a);
    const prioridadB = prioridadProduccion(b);

    if (prioridadA !== prioridadB) {
      return prioridadA - prioridadB;
    }

    if (!a.fecha_entrega_estimada) return 1;
    if (!b.fecha_entrega_estimada) return -1;

    return new Date(a.fecha_entrega_estimada) - new Date(b.fecha_entrega_estimada);
  })
  .map(n => (
          <div
           key={n.id}
            onClick={() => setModalProduccion(n)}
           style={{
            cursor:"pointer",
            background:COLORS.card,
            border: estaAtrasada(n.fecha_entrega_estimada)
  ? `1px solid ${COLORS.danger}`
  : venceHoy(n.fecha_entrega_estimada)
  ? `1px solid ${COLORS.warning}`
  : `1px solid ${COLORS.border}`,
            borderRadius:12,
            padding:14
          }}
          >
            <div style={{ display:"flex", justifyContent:"space-between", gap:12, flexWrap:"wrap" }}>
              <div>
  <b style={{ color:COLORS.success }}>NV#{n.numero}</b>
  <span style={{ marginLeft:8 }}>{n.cliente}</span>

  <span style={{
    marginLeft:10,
    padding:"3px 8px",
    borderRadius:999,
    fontSize:11,
    fontWeight:700,
    background:estadoProduccion(n).color + "22",
    color:estadoProduccion(n).color,
    border:`1px solid ${estadoProduccion(n).color}`
  }}>
    {estadoProduccion(n).texto}
  </span>
</div>

              <span style={{ color:COLORS.muted }}>
                Entrega estimada: {n.fecha_entrega_estimada || "Sin fecha"}
{" "}
{estaAtrasada(n.fecha_entrega_estimada) && (
  <b style={{ color:COLORS.danger }}>(Atrasada)</b>
)}
{venceHoy(n.fecha_entrega_estimada) && (
  <b style={{ color:COLORS.warning }}>(Para hoy)</b>
)}
{venceManana(n.fecha_entrega_estimada) && (
  <b style={{ color:"#60a5fa" }}>(Para mañana)</b>
)}
              </span>
            </div>

            <div style={{ marginTop:8, color:COLORS.muted, fontSize:13 }}>
              Proceso: <b style={{ color:COLORS.text }}>{n.proceso}</b>
            </div>

            <div style={{ marginTop:10, display:"flex", gap:8, flexWrap:"wrap", fontSize:13 }}>
              {n.mdf_cortado && <span>✅ MDF cortado</span>}
              {n.lamina_cortada && <span>✅ Lámina cortada</span>}
              {n.tupizado && <span>✅ Tupizado</span>}
              {n.armado && <span>✅ Armado</span>}
              {n.pegado && <span>✅ Pegado</span>}
              {n.postformado && <span>✅ Postformado</span>}
            </div>

            {n.produccion_observaciones && (
              <p style={{ marginTop:10, color:COLORS.warning }}>
                Obs: {n.produccion_observaciones}
              </p>
            )}
          </div>
        ))}
    </div>
  </div>
)}
{tab==="inventario" && (
  <div>
    <h2 style={{ margin:"0 0 14px", fontFamily:"Georgia,serif", color:COLORS.accent, fontSize:17 }}>
      📦 Inventario
    </h2>
    <button
  onClick={() => setModalNuevoProducto(true)}
  style={{
    marginBottom:14,
    padding:"9px 12px",
    borderRadius:8,
    border:"none",
    background:COLORS.accent,
    color:"#111",
    fontWeight:700,
    cursor:"pointer"
  }}
>
  + Agregar producto
</button>
<label
  style={{
    display:"inline-block",
    marginLeft:8,
    marginBottom:14,
    padding:"9px 12px",
    borderRadius:8,
    border:`1px solid ${COLORS.border}`,
    background:COLORS.surface,
    color:COLORS.text,
    fontWeight:700,
    cursor:"pointer"
  }}
>
  📥 Importar OC
  <input
    type="file"
    accept=".xls,.xlsx"
    onChange={importarOrdenCompraExcel}
    style={{ display:"none" }}
  />
</label>
<div style={{
  background:COLORS.card,
  border:`1px solid ${COLORS.border}`,
  borderRadius:12,
  padding:12,
  marginBottom:16
}}>
  <input
    value={busquedaInventario}
    onChange={(e) => setBusquedaInventario(e.target.value)}
    placeholder="Buscar producto..."
    style={{
      width:"100%",
      padding:"10px",
      borderRadius:8,
      border:`1px solid ${COLORS.border}`,
      background:COLORS.surface,
      color:COLORS.text,
      fontSize:16,
      boxSizing:"border-box",
      marginBottom:10
    }}
  />

  <div style={{ display:"flex", gap:8, flexWrap:"wrap" }}>
    {[
      ["todos", "Todos"],
      ["tableros", "Tableros"],
      ["laminados", "Laminados"],
      ["pegamentos", "Pegamentos"],
      ["embalaje", "Embalaje"],
      ["tornillos", "Tornillos"],
      ["herrajes", "Herrajes"],
      ["insumos generales", "Insumos"]
    ].map(([valor, texto]) => (
      <button
        key={valor}
        onClick={() => setFiltroCategoriaInventario(valor)}
        style={{
          border:`1px solid ${filtroCategoriaInventario === valor ? COLORS.accent : COLORS.border}`,
          background:filtroCategoriaInventario === valor ? COLORS.accent : COLORS.surface,
          color:filtroCategoriaInventario === valor ? "#111" : COLORS.text,
          borderRadius:999,
          padding:"6px 10px",
          cursor:"pointer",
          fontSize:12,
          fontWeight:700
        }}
      >
        {texto}
      </button>
    ))}
  </div>
</div>
    <div style={{
      background:COLORS.card,
      border:`1px solid ${COLORS.border}`,
      borderRadius:12,
      padding:14,
      marginBottom:18
    }}>
      <p style={{ margin:"0 0 6px", color:COLORS.text, fontWeight:700 }}>
        Inventario general
      </p>
      <p style={{ margin:0, color:COLORS.muted, fontSize:13 }}>
        Aquí se controlarán MDF, Melamil, pegamentos, cartón, film, tornillos, herrajes y otros productos.
      </p>
    </div>

    <h3 style={{ color:COLORS.success, fontSize:15, margin:"0 0 10px" }}>
      Productos generales
    </h3>

    <div style={{
      display:"grid",
      gridTemplateColumns:"repeat(auto-fit, minmax(230px, 1fr))",
      gap:12,
      marginBottom:24
    }}>
      {productosInventarioFiltrados.filter(p => !p.es_laminado).length === 0 ? (
        <p style={{ color:COLORS.muted }}>
          Todavía no hay productos generales cargados.
        </p>
      ) : (
        productosInventarioFiltrados
  .filter(p => !p.es_laminado)
          .map((p) => {
            const estado = estadoStockInventario(p);

            return (
              <div key={p.id} style={{
                background:COLORS.card,
                border:`1px solid ${COLORS.border}`,
                borderRadius:12,
                padding:14
              }}>
                <div style={{
  display:"flex",
  justifyContent:"space-between",
  alignItems:"center",
  gap:10
}}>
  <b style={{
    color:COLORS.text,
    fontSize:14,
    lineHeight:1.2
  }}>
    {p.nombre}
  </b>

  <span style={{
    fontSize:22,
    fontWeight:800,
    color:COLORS.accent,
    whiteSpace:"nowrap"
  }}>
    {Number(p.stock_actual || 0)}
  </span>
</div>

<div style={{
  marginTop:8,
  color:estado.color,
  fontSize:12,
  fontWeight:700
}}>
  {estado.texto}
</div>
                <div style={{
  display:"grid",
  gridTemplateColumns:"repeat(4, 1fr)",
  gap:6,
  marginTop:12
}}>
  <button
    title="Entrada"
    onClick={() => setModalMovimientoInventario({ producto:p, tipo:"entrada" })}
    style={{
      padding:"6px",
      borderRadius:8,
      border:"none",
      background:COLORS.success,
      color:"#fff",
      fontWeight:700,
      cursor:"pointer"
    }}
  >
    ➕
  </button>

  <button
    title="Salida"
    onClick={() => setModalMovimientoInventario({ producto:p, tipo:"salida" })}
    style={{
      padding:"6px",
      borderRadius:8,
      border:"none",
      background:COLORS.danger,
      color:"#fff",
      fontWeight:700,
      cursor:"pointer"
    }}
  >
    ➖
  </button>

  <button
    title="Historial"
    onClick={() => setModalHistorialProducto(p)}
    style={{
      padding:"6px",
      borderRadius:8,
      border:`1px solid ${COLORS.border}`,
      background:COLORS.surface,
      color:COLORS.text,
      fontWeight:700,
      cursor:"pointer"
    }}
  >
    📜
  </button>

  <button
    title="Editar"
    onClick={() => setModalEditarProducto(p)}
    style={{
      padding:"6px",
      borderRadius:8,
      border:`1px solid ${COLORS.border}`,
      background:COLORS.surface,
      color:COLORS.text,
      fontWeight:700,
      cursor:"pointer"
    }}
  >
    ✏️
  </button>
</div>
              </div>
            );
          })
      )}
    </div>

    <h3 style={{ color:COLORS.accent, fontSize:15, margin:"0 0 10px" }}>
      Laminados
    </h3>
<div style={{
  background:COLORS.card,
  border:`1px solid ${COLORS.border}`,
  borderRadius:12,
  padding:12,
  marginBottom:14
}}>
  <p style={{ margin:"0 0 10px", color:COLORS.text, fontWeight:700 }}>
    Buscar sobrantes disponibles
  </p>

  <input
    value={busquedaLaminadoNombre}
    onChange={(e) => setBusquedaLaminadoNombre(e.target.value)}
    placeholder="Filtrar por color o nombre del laminado..."
    style={{
      width:"100%",
      padding:"10px",
      borderRadius:8,
      border:`1px solid ${COLORS.border}`,
      background:COLORS.surface,
      color:COLORS.text,
      fontSize:16,
      boxSizing:"border-box",
      marginBottom:8
    }}
  />

  <div style={{
    display:"grid",
    gridTemplateColumns:"1fr 1fr",
    gap:8
  }}>
    <input
      type="number"
      value={busquedaSobranteLargo}
      onChange={(e) => setBusquedaSobranteLargo(e.target.value)}
      placeholder="Largo necesario"
      style={{
        width:"100%",
        padding:"10px",
        borderRadius:8,
        border:`1px solid ${COLORS.border}`,
        background:COLORS.surface,
        color:COLORS.text,
        fontSize:16,
        boxSizing:"border-box"
      }}
    />

    <input
      type="number"
      value={busquedaSobranteAncho}
      onChange={(e) => setBusquedaSobranteAncho(e.target.value)}
      placeholder="Ancho necesario"
      style={{
        width:"100%",
        padding:"10px",
        borderRadius:8,
        border:`1px solid ${COLORS.border}`,
        background:COLORS.surface,
        color:COLORS.text,
        fontSize:16,
        boxSizing:"border-box"
      }}
    />
  </div>

  {Number(busquedaSobranteLargo) > 0 && Number(busquedaSobranteAncho) > 0 && (
    <div style={{ marginTop:12 }}>
      {sobrantesGlobalesCompatibles.length === 0 ? (
        <p style={{ margin:0, color:COLORS.danger, fontSize:13, fontWeight:700 }}>
          No hay sobrantes compatibles.
        </p>
      ) : (
        <div style={{ display:"flex", flexDirection:"column", gap:8 }}>
          {sobrantesGlobalesCompatibles.map((s) => (
            <div
              key={s.id}
              style={{
                border:`1px solid ${COLORS.border}`,
                borderRadius:8,
                padding:10,
                background:COLORS.surface
              }}
            >
              <div style={{
                display:"flex",
                justifyContent:"space-between",
                gap:10
              }}>
                <b style={{ color:COLORS.text }}>
                  {s.producto_nombre}
                </b>

                <span style={{ color:COLORS.success, fontWeight:700 }}>
                  {Number(s.largo)} x {Number(s.ancho)}
                </span>
              </div>

              <div style={{ color:COLORS.muted, fontSize:12, marginTop:4 }}>
                Sobrante compatible
              </div>
            </div>
          ))}
        </div>
      )}
    </div>
  )}
</div>
    <div style={{
      display:"grid",
      gridTemplateColumns:"repeat(auto-fit, minmax(230px, 1fr))",
      gap:12
    }}>
      {productosInventarioFiltrados.filter(p => p.es_laminado).length === 0 ? (
        <p style={{ color:COLORS.muted }}>
          Todavía no hay laminados cargados.
        </p>
      ) : (
        productosInventarioFiltrados
  .filter(p => p.es_laminado)
          .map((p) => {
            const estado = estadoStockInventario(p);

            return (
              <div key={p.id} style={{
                background:COLORS.card,
                border:`1px solid ${COLORS.border}`,
                borderRadius:12,
                padding:14
              }}>
                <div style={{
  display:"flex",
  justifyContent:"space-between",
  alignItems:"center",
  gap:10
}}>
  <b style={{
    color:COLORS.text,
    fontSize:14,
    lineHeight:1.2
  }}>
    {p.nombre}
  </b>

  <span style={{
    fontSize:22,
    fontWeight:800,
    color:COLORS.accent,
    whiteSpace:"nowrap"
  }}>
    {Number(p.stock_actual || 0)}
  </span>
</div>

<div style={{
  marginTop:8,
  color:estado.color,
  fontSize:12,
  fontWeight:700
}}>
  {estado.texto}
</div>
                <div style={{
  display:"grid",
  gridTemplateColumns:"repeat(4, 1fr)",
  gap:6,
  marginTop:12
}}>
  <button
    title="Entrada"
    onClick={() => setModalMovimientoInventario({ producto:p, tipo:"entrada" })}
    style={{
      padding:"6px",
      borderRadius:8,
      border:"none",
      background:COLORS.success,
      color:"#fff",
      fontWeight:700,
      cursor:"pointer"
    }}
  >
    ➕
  </button>

  <button
    title="Salida"
    onClick={() => setModalMovimientoInventario({ producto:p, tipo:"salida" })}
    style={{
      padding:"6px",
      borderRadius:8,
      border:"none",
      background:COLORS.danger,
      color:"#fff",
      fontWeight:700,
      cursor:"pointer"
    }}
  >
    ➖
  </button>

  <button
    title="Historial"
    onClick={() => setModalHistorialProducto(p)}
    style={{
      padding:"6px",
      borderRadius:8,
      border:`1px solid ${COLORS.border}`,
      background:COLORS.surface,
      color:COLORS.text,
      fontWeight:700,
      cursor:"pointer"
    }}
  >
    📜
  </button>

  <button
    title="Editar"
    onClick={() => setModalEditarProducto(p)}
    style={{
      padding:"6px",
      borderRadius:8,
      border:`1px solid ${COLORS.border}`,
      background:COLORS.surface,
      color:COLORS.text,
      fontWeight:700,
      cursor:"pointer"
    }}
  >
    ✏️
  </button>
</div>

                <button
  onClick={() => setModalSobrantesLaminado(p)}
  style={{
    marginTop:12,
    width:"100%",
    padding:"8px 10px",
    borderRadius:8,
    border:`1px solid ${COLORS.border}`,
    background:COLORS.surface,
    color:COLORS.text,
    fontWeight:700,
    cursor:"pointer"
  }}
>
  Sobrantes ({sobrantesLaminados.filter(s => s.producto_id === p.id && !s.usado).length})
</button>
              </div>
            );
          })
      )}
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

      {modalCot && (
  <NotasModal
    cotizacion={modalCot}
    detalles={detallesCotizaciones.filter(d =>
      d.cotizacion_id === Number(String(modalCot.id).replace("supabase-", ""))
    )}
    seguimiento={seguimiento}
    onSave={saveSeguimiento}
    onClose={() => setModalCot(null)}
  />
)}
{modalNuevoProducto && (
  <NuevoProductoModal
    onClose={() => setModalNuevoProducto(false)}
    onSave={guardarNuevoProducto}
  />
)}
{modalMovimientoInventario && (
  <MovimientoInventarioModal
    data={modalMovimientoInventario}
    onClose={() => setModalMovimientoInventario(null)}
    onSave={guardarMovimientoInventario}
  />
)}

{modalHistorialProducto && (
  <HistorialInventarioModal
    producto={modalHistorialProducto}
    movimientos={movimientosInventario}
    onClose={() => setModalHistorialProducto(null)}
  />
)}
{modalEditarProducto && (
  <EditarProductoModal
    producto={modalEditarProducto}
    onClose={() => setModalEditarProducto(null)}
    onSave={guardarEdicionProducto}
  />
)}
{modalSobrantesLaminado && (
  <SobrantesLaminadoModal
    producto={modalSobrantesLaminado}
    sobrantes={sobrantesLaminados.filter(s => s.producto_id === modalSobrantesLaminado.id)}
    onClose={() => setModalSobrantesLaminado(null)}
    onSave={guardarSobrantesLaminado}
    onUsar={marcarSobranteUsado}
  />
)}
{modalPreviewOC && (
  <PreviewOCModal
    productos={previewOC}
    onClose={() => setModalPreviewOC(false)}
    onConfirm={confirmarImportacionOC}
  />
)}
{modalProduccion && (
  <ProduccionModal
  nota={modalProduccion}
  detalles={detallesNotasVenta.filter(d =>
    String(d.nota_venta_numero) === String(modalProduccion.numero)
  )}
  onClose={() => setModalProduccion(null)}
  onSave={guardarProduccion}
/>
)}
    {modalNV && (
  <GestionNVModal
  nota={modalNV}
  abonos={abonosNV.filter(a => a.nota_venta_id === Number(String(modalNV.id).replace("supabase-", "")))}
  onClose={() => setModalNV(null)}
  onSave={guardarGestionNV}
/>
)}
    </div>
  );
}
