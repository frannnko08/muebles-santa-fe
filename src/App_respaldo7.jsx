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
  const hoy = new Date().toISOString().split("T")[0];
  const [fecha, setFecha] = useState(hoy);
  const [documento, setDocumento] = useState("");
  const [proveedor, setProveedor] = useState("");
  const [totalCompra, setTotalCompra] = useState("");
  const [estadoPago, setEstadoPago] = useState("pagado");

  useEffect(() => {
    const cerrarConEsc = (e) => {
      if (e.key === "Escape") onClose();
    };

    window.addEventListener("keydown", cerrarConEsc);

    return () => {
      window.removeEventListener("keydown", cerrarConEsc);
    };
  }, [onClose]);

  const inputStyle = {
    width:"100%",
    padding:"9px 10px",
    borderRadius:8,
    border:`1px solid ${COLORS.border}`,
    background:COLORS.surface,
    color:COLORS.text,
    boxSizing:"border-box"
  };

  const confirmar = () => {
    onConfirm({
      fecha,
      documento: documento.trim(),
      proveedor: proveedor.trim(),
      totalCompra: Number(totalCompra || 0),
      estadoPago
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
          width:"620px",
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
          Revisa los productos antes de ingresarlos al inventario. Si ingresas el total de la compra, la app también creará el asiento contable automáticamente.
        </p>

        <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fit, minmax(160px, 1fr))", gap:10, margin:"12px 0" }}>
          <div>
            <label style={{ fontSize:12, color:COLORS.muted }}>Fecha</label>
            <input type="date" value={fecha} onChange={(e)=>setFecha(e.target.value)} style={inputStyle} />
          </div>
          <div>
            <label style={{ fontSize:12, color:COLORS.muted }}>Factura / OC</label>
            <input value={documento} onChange={(e)=>setDocumento(e.target.value)} placeholder="Ej: F-12345" style={inputStyle} />
          </div>
          <div>
            <label style={{ fontSize:12, color:COLORS.muted }}>Proveedor</label>
            <input value={proveedor} onChange={(e)=>setProveedor(e.target.value)} placeholder="Ej: Imperial" style={inputStyle} />
          </div>
          <div>
            <label style={{ fontSize:12, color:COLORS.muted }}>Total compra IVA incluido</label>
            <input type="number" value={totalCompra} onChange={(e)=>setTotalCompra(e.target.value)} placeholder="Ej: 119000" style={inputStyle} />
          </div>
        </div>

        <div style={{ display:"flex", gap:10, flexWrap:"wrap", margin:"0 0 14px" }}>
          <button
            onClick={() => setEstadoPago("pagado")}
            style={{
              padding:"8px 12px",
              borderRadius:8,
              border:`1px solid ${estadoPago === "pagado" ? COLORS.success : COLORS.border}`,
              background:estadoPago === "pagado" ? COLORS.success : COLORS.surface,
              color:estadoPago === "pagado" ? "#fff" : COLORS.text,
              fontWeight:700,
              cursor:"pointer"
            }}
          >
            Pagado con banco
          </button>
          <button
            onClick={() => setEstadoPago("pendiente")}
            style={{
              padding:"8px 12px",
              borderRadius:8,
              border:`1px solid ${estadoPago === "pendiente" ? COLORS.warn : COLORS.border}`,
              background:estadoPago === "pendiente" ? COLORS.warn : COLORS.surface,
              color:estadoPago === "pendiente" ? "#111" : COLORS.text,
              fontWeight:700,
              cursor:"pointer"
            }}
          >
            Pendiente proveedor
          </button>
        </div>

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
            onClick={confirmar}
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

function calcularResumenVentaLaminas(venta) {
  const ventaTotal = Number(venta?.total_venta || 0);
  const costoTotal = Number(venta?.costo_compra_total || 0);

  const ventaNeta = Math.round(ventaTotal / 1.19);
  const ivaVenta = ventaTotal - ventaNeta;

  const costoNeto = Math.round(costoTotal / 1.19);
  const ivaCompra = costoTotal - costoNeto;

  const utilidadNeta = ventaNeta - costoNeto;
  const ivaProvisionar = ivaVenta - ivaCompra;

  return {
    ventaTotal,
    costoTotal,
    ventaNeta,
    ivaVenta,
    costoNeto,
    ivaCompra,
    utilidadNeta,
    ivaProvisionar
  };
}

function VentaLaminasModal({ venta, detalles = [], onClose, onSaveCosto }) {
  const [costoCompra, setCostoCompra] = useState(venta.costo_compra_total || "");
  const resumen = calcularResumenVentaLaminas({ ...venta, costo_compra_total: costoCompra });

  useEffect(() => {
    const cerrarConEsc = (e) => {
      if (e.key === "Escape") onClose();
    };
    window.addEventListener("keydown", cerrarConEsc);
    return () => window.removeEventListener("keydown", cerrarConEsc);
  }, [onClose]);

  const guardar = () => {
    onSaveCosto(venta, Number(costoCompra) || 0);
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
        zIndex:120,
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
          padding:24,
          width:760,
          maxWidth:"95vw",
          maxHeight:"90vh",
          overflowY:"auto",
          boxSizing:"border-box"
        }}
      >
        <div style={{ display:"flex", justifyContent:"space-between", gap:12, alignItems:"flex-start", marginBottom:16 }}>
          <div>
            <h3 style={{ margin:"0 0 4px", color:COLORS.accent, fontFamily:"Georgia,serif", fontSize:18 }}>
              Venta de láminas #{venta.numero}
            </h3>
            <div style={{ fontSize:12, color:COLORS.muted }}>
              {venta.cliente} · {fmtDate(venta.fecha)}
            </div>
          </div>
          <button onClick={onClose} style={{ background:"transparent", border:`1px solid ${COLORS.border}`, color:COLORS.muted, borderRadius:8, padding:"6px 10px", cursor:"pointer" }}>✕</button>
        </div>

        <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fit,minmax(140px,1fr))", gap:10, marginBottom:16 }}>
          <StatCard label="Venta c/IVA" value={fmt(resumen.ventaTotal)} icon="💰" color={COLORS.success}/>
          <StatCard label="Costo c/IVA" value={fmt(resumen.costoTotal)} icon="🧾" color={COLORS.warning}/>
          <StatCard label="Utilidad neta" value={fmt(resumen.utilidadNeta)} icon="📈" color={resumen.utilidadNeta >= 0 ? COLORS.success : COLORS.danger}/>
          <StatCard label="IVA a provisionar" value={fmt(resumen.ivaProvisionar)} icon="🏛️" color={resumen.ivaProvisionar >= 0 ? COLORS.accent : COLORS.danger}/>
        </div>

        <div style={{ background:COLORS.surface, border:`1px solid ${COLORS.border}`, borderRadius:12, padding:14, marginBottom:16 }}>
          <label style={{ display:"block", fontSize:11, color:COLORS.muted, marginBottom:6, textTransform:"uppercase", letterSpacing:1 }}>
            Costo compra IVA incluido
          </label>
          <div style={{ display:"flex", gap:8, flexWrap:"wrap" }}>
            <input
              type="number"
              value={costoCompra}
              onChange={(e) => setCostoCompra(e.target.value)}
              placeholder="Ej: 80000"
              style={{ flex:"1 1 180px", minWidth:0, background:COLORS.bg, border:`1px solid ${COLORS.border}`, color:COLORS.text, borderRadius:8, padding:"10px 12px" }}
            />
            <button onClick={guardar} style={{ background:COLORS.accent, color:"#111", border:"none", borderRadius:8, padding:"10px 14px", fontWeight:700, cursor:"pointer" }}>
              Guardar costo
            </button>
          </div>
          <div style={{ marginTop:10, fontSize:12, color:COLORS.muted }}>
            Venta neta: {fmt(resumen.ventaNeta)} · IVA venta: {fmt(resumen.ivaVenta)} · Costo neto: {fmt(resumen.costoNeto)} · IVA compra: {fmt(resumen.ivaCompra)}
          </div>
        </div>

        <div style={{ background:COLORS.surface, border:`1px solid ${COLORS.border}`, borderRadius:12, padding:14 }}>
          <div style={{ fontSize:13, fontWeight:700, color:COLORS.accent, marginBottom:10 }}>Detalle importado desde Excel</div>
          {detalles.length === 0 ? (
            <div style={{ fontSize:12, color:COLORS.muted }}>No hay detalle guardado para esta venta.</div>
          ) : (
            <div style={{ overflowX:"auto" }}>
              <table style={{ width:"100%", borderCollapse:"collapse", fontSize:12 }}>
                <thead>
                  <tr style={{ color:COLORS.muted, textAlign:"left" }}>
                    <th style={{ padding:"8px", borderBottom:`1px solid ${COLORS.border}` }}>Cant.</th>
                    <th style={{ padding:"8px", borderBottom:`1px solid ${COLORS.border}` }}>Tipo</th>
                    <th style={{ padding:"8px", borderBottom:`1px solid ${COLORS.border}` }}>Medida</th>
                    <th style={{ padding:"8px", borderBottom:`1px solid ${COLORS.border}` }}>Color</th>
                    <th style={{ padding:"8px", borderBottom:`1px solid ${COLORS.border}`, textAlign:"right" }}>Total</th>
                  </tr>
                </thead>
                <tbody>
                  {detalles.map((d) => (
                    <tr key={d.id || d.orden}>
                      <td style={{ padding:"8px", borderBottom:`1px solid ${COLORS.border}` }}>{d.cantidad}</td>
                      <td style={{ padding:"8px", borderBottom:`1px solid ${COLORS.border}` }}>{d.tipo}</td>
                      <td style={{ padding:"8px", borderBottom:`1px solid ${COLORS.border}` }}>{d.largo} x {d.ancho}</td>
                      <td style={{ padding:"8px", borderBottom:`1px solid ${COLORS.border}`, color:COLORS.accent }}>{d.color}</td>
                      <td style={{ padding:"8px", borderBottom:`1px solid ${COLORS.border}`, textAlign:"right", fontWeight:700 }}>{fmt(d.total)}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}
        </div>
      </div>
    </div>
  );
}


function ContabilidadModal({ asiento, onClose, onSave }) {
  const hoy = new Date().toISOString().slice(0, 10);
  const cuentasRapidas = [
    "BANCO",
    "CAJA",
    "CLIENTES / CXC",
    "PROVEEDORES",
    "COMPRAS / MATERIALES",
    "VENTAS",
    "IVA CF",
    "IVA DF",
    "REMUNERACIONES",
    "AFP POR PAGAR",
    "SALUD POR PAGAR",
    "CESANTÍA POR PAGAR",
    "MUTUAL POR PAGAR",
    "ARRIENDO",
    "COMBUSTIBLE",
    "OTROS"
  ];

  const nuevaFila = (base = {}) => ({
    fecha: base.fecha || hoy,
    detalle: base.detalle || "",
    desglose: base.desglose || "",
    documento: base.documento || "",
    definicion: base.definicion || "",
    debe: base.debe || "",
    haber: base.haber || ""
  });

  const [form, setForm] = useState({
    fecha: asiento?.fecha || hoy,
    detalle: asiento?.detalle || "",
    desglose: asiento?.desglose || "",
    documento: asiento?.documento || "",
    definicion: asiento?.definicion || "",
    debe: asiento?.debe || "",
    haber: asiento?.haber || ""
  });

  const [cabecera, setCabecera] = useState({
    fecha: hoy,
    glosa: "",
    documento: "",
    definicion: ""
  });

  const [filas, setFilas] = useState([
    nuevaFila(),
    nuevaFila(),
    nuevaFila()
  ]);

  useEffect(() => {
    const cerrarConEsc = (e) => {
      if (e.key === "Escape") onClose();
    };
    window.addEventListener("keydown", cerrarConEsc);
    return () => window.removeEventListener("keydown", cerrarConEsc);
  }, [onClose]);

  const setCampo = (campo, valor) => setForm(prev => ({ ...prev, [campo]: valor }));
  const setCampoCabecera = (campo, valor) => setCabecera(prev => ({ ...prev, [campo]: valor }));
  const setCampoFila = (index, campo, valor) => {
    setFilas(prev => prev.map((fila, i) => i === index ? { ...fila, [campo]: valor } : fila));
  };

  const agregarFila = () => {
    setFilas(prev => [...prev, nuevaFila({
      fecha: cabecera.fecha,
      desglose: cabecera.glosa,
      documento: cabecera.documento,
      definicion: cabecera.definicion
    })]);
  };

  const eliminarFila = (index) => {
    setFilas(prev => prev.length <= 1 ? prev : prev.filter((_, i) => i !== index));
  };

  const aplicarCabeceraATodas = () => {
    setFilas(prev => prev.map(fila => ({
      ...fila,
      fecha: cabecera.fecha || fila.fecha,
      desglose: cabecera.glosa || fila.desglose,
      documento: cabecera.documento || fila.documento,
      definicion: cabecera.definicion || fila.definicion
    })));
  };

  const copiarFilaAnterior = (index) => {
    if (index <= 0) return;
    setFilas(prev => prev.map((fila, i) => {
      if (i !== index) return fila;
      const anterior = prev[index - 1];
      return {
        ...fila,
        fecha: anterior.fecha,
        desglose: anterior.desglose,
        documento: anterior.documento,
        definicion: anterior.definicion
      };
    }));
  };

  const limpiarFilasVacias = () => filas
    .map(fila => ({
      fecha: fila.fecha,
      detalle: String(fila.detalle || "").trim().toUpperCase(),
      desglose: String(fila.desglose || "").trim(),
      documento: String(fila.documento || "").trim(),
      definicion: String(fila.definicion || "").trim(),
      debe: Number(fila.debe || 0),
      haber: Number(fila.haber || 0)
    }))
    .filter(fila => fila.fecha || fila.detalle || fila.desglose || fila.documento || fila.definicion || fila.debe || fila.haber);

  const filasValidas = limpiarFilasVacias();
  const totalDebe = filasValidas.reduce((sum, fila) => sum + Number(fila.debe || 0), 0);
  const totalHaber = filasValidas.reduce((sum, fila) => sum + Number(fila.haber || 0), 0);
  const diferencia = totalDebe - totalHaber;
  const estaCuadrado = filasValidas.length > 0 && Math.abs(diferencia) === 0;

  const guardar = () => {
    if (asiento?.id) {
      if (!form.fecha) {
        alert("Debes ingresar una fecha.");
        return;
      }
      if (!form.detalle.trim()) {
        alert("Debes ingresar el detalle/cuenta del asiento.");
        return;
      }

      const debe = Number(form.debe) || 0;
      const haber = Number(form.haber) || 0;

      if (debe <= 0 && haber <= 0) {
        alert("Debes ingresar un monto en Debe o Haber.");
        return;
      }

      if (debe > 0 && haber > 0) {
        alert("Un asiento no debe tener Debe y Haber al mismo tiempo. Deja uno de los dos en cero.");
        return;
      }

      onSave({
        ...asiento,
        fecha: form.fecha,
        detalle: form.detalle.trim().toUpperCase(),
        desglose: form.desglose.trim(),
        documento: form.documento.trim(),
        definicion: form.definicion.trim(),
        debe,
        haber
      });
      return;
    }

    if (filasValidas.length < 2) {
      alert("Debes ingresar al menos 2 líneas para un asiento manual.");
      return;
    }

    for (const fila of filasValidas) {
      if (!fila.fecha) {
        alert("Todas las filas con datos deben tener fecha.");
        return;
      }
      if (!fila.detalle) {
        alert("Todas las filas con datos deben tener cuenta/detalle.");
        return;
      }
      if (fila.debe <= 0 && fila.haber <= 0) {
        alert("Cada fila debe tener monto en Debe o en Haber.");
        return;
      }
      if (fila.debe > 0 && fila.haber > 0) {
        alert("Una fila no puede tener Debe y Haber al mismo tiempo.");
        return;
      }
    }

    if (!estaCuadrado) {
      alert("El asiento no está cuadrado. El total Debe debe ser igual al total Haber.");
      return;
    }

    onSave(filasValidas);
  };

  const inputStyle = {
    width:"100%",
    boxSizing:"border-box",
    background:COLORS.bg,
    border:`1px solid ${COLORS.border}`,
    color:COLORS.text,
    borderRadius:8,
    padding:"10px 12px"
  };

  const tableInputStyle = {
    width:"100%",
    boxSizing:"border-box",
    background:COLORS.bg,
    border:`1px solid ${COLORS.border}`,
    color:COLORS.text,
    borderRadius:7,
    padding:"8px 9px",
    fontSize:12
  };

  const labelStyle = { display:"block", fontSize:11, color:COLORS.muted, marginBottom:6, textTransform:"uppercase", letterSpacing:0.8 };

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
        zIndex:130,
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
          padding:22,
          width:asiento?.id ? 720 : 1180,
          maxWidth:"98vw",
          boxSizing:"border-box"
        }}
      >
        <div style={{ display:"flex", justifyContent:"space-between", gap:12, alignItems:"flex-start", marginBottom:16 }}>
          <div>
            <h3 style={{ margin:"0 0 4px", color:COLORS.accent, fontFamily:"Georgia,serif", fontSize:18 }}>
              {asiento?.id ? "Editar asiento contable" : "Asiento manual tipo tabla"}
            </h3>
            <div style={{ fontSize:12, color:COLORS.muted }}>
              {asiento?.id ? "Edita una línea del libro diario." : "Ingresa varias líneas de una sola vez. Solo se puede guardar si Debe y Haber cuadran."}
            </div>
          </div>
          <button onClick={onClose} style={{ background:"transparent", border:`1px solid ${COLORS.border}`, color:COLORS.muted, borderRadius:8, padding:"6px 10px", cursor:"pointer" }}>✕</button>
        </div>

        {asiento?.id ? (
          <>
            <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fit,minmax(180px,1fr))", gap:12 }}>
              <div>
                <label style={labelStyle}>Fecha</label>
                <input type="date" value={form.fecha} onChange={(e) => setCampo("fecha", e.target.value)} style={inputStyle} />
              </div>
              <div>
                <label style={labelStyle}>Detalle / Cuenta</label>
                <input value={form.detalle} onChange={(e) => setCampo("detalle", e.target.value)} placeholder="Ej: BANCO, CAJA, IVA CF" style={inputStyle} />
              </div>
              <div>
                <label style={labelStyle}>NV / Factura</label>
                <input value={form.documento} onChange={(e) => setCampo("documento", e.target.value)} placeholder="Ej: F1234 / NV 7500" style={inputStyle} />
              </div>
              <div>
                <label style={labelStyle}>Definición</label>
                <input value={form.definicion} onChange={(e) => setCampo("definicion", e.target.value)} placeholder="Ej: COCINA, LÁMINAS" style={inputStyle} />
              </div>
            </div>

            <div style={{ marginTop:12 }}>
              <label style={labelStyle}>Desglose</label>
              <input value={form.desglose} onChange={(e) => setCampo("desglose", e.target.value)} placeholder="Ej: Cliente, proveedor o explicación" style={inputStyle} />
            </div>

            <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fit,minmax(160px,1fr))", gap:12, marginTop:12 }}>
              <div>
                <label style={labelStyle}>Debe</label>
                <input type="number" value={form.debe} onChange={(e) => setCampo("debe", e.target.value)} placeholder="0" style={inputStyle} />
              </div>
              <div>
                <label style={labelStyle}>Haber</label>
                <input type="number" value={form.haber} onChange={(e) => setCampo("haber", e.target.value)} placeholder="0" style={inputStyle} />
              </div>
            </div>
          </>
        ) : (
          <>
            <div style={{ background:COLORS.subtle, border:`1px solid ${COLORS.border}`, borderRadius:12, padding:14, marginBottom:14 }}>
              <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fit,minmax(180px,1fr))", gap:12 }}>
                <div>
                  <label style={labelStyle}>Fecha general</label>
                  <input type="date" value={cabecera.fecha} onChange={(e) => setCampoCabecera("fecha", e.target.value)} style={inputStyle} />
                </div>
                <div>
                  <label style={labelStyle}>Glosa / Detalle general</label>
                  <input value={cabecera.glosa} onChange={(e) => setCampoCabecera("glosa", e.target.value)} placeholder="Ej: Compra MDF Imperial" style={inputStyle} />
                </div>
                <div>
                  <label style={labelStyle}>NV / Factura general</label>
                  <input value={cabecera.documento} onChange={(e) => setCampoCabecera("documento", e.target.value)} placeholder="Ej: F-12345" style={inputStyle} />
                </div>
                <div>
                  <label style={labelStyle}>Definición general</label>
                  <input value={cabecera.definicion} onChange={(e) => setCampoCabecera("definicion", e.target.value)} placeholder="Ej: MATERIALES" style={inputStyle} />
                </div>
              </div>
              <div style={{ display:"flex", justifyContent:"space-between", gap:10, flexWrap:"wrap", alignItems:"center", marginTop:12 }}>
                <div style={{ fontSize:11, color:COLORS.muted }}>
                  Usa el botón para rellenar automáticamente fecha, glosa, documento y definición en todas las filas.
                </div>
                <button onClick={aplicarCabeceraATodas} style={{ padding:"9px 12px", borderRadius:8, border:"none", background:COLORS.accent, color:"#111", fontWeight:700, cursor:"pointer" }}>
                  Rellenar todas las filas
                </button>
              </div>
            </div>

            <div style={{ overflowX:"auto", border:`1px solid ${COLORS.border}`, borderRadius:12 }}>
              <table style={{ width:"100%", borderCollapse:"collapse", minWidth:1050, fontSize:12 }}>
                <thead>
                  <tr style={{ color:COLORS.muted, textAlign:"left", background:COLORS.surface }}>
                    <th style={{ padding:"8px", borderBottom:`1px solid ${COLORS.border}` }}>Fecha</th>
                    <th style={{ padding:"8px", borderBottom:`1px solid ${COLORS.border}` }}>Cuenta</th>
                    <th style={{ padding:"8px", borderBottom:`1px solid ${COLORS.border}` }}>Desglose</th>
                    <th style={{ padding:"8px", borderBottom:`1px solid ${COLORS.border}` }}>NV/Factura</th>
                    <th style={{ padding:"8px", borderBottom:`1px solid ${COLORS.border}` }}>Definición</th>
                    <th style={{ padding:"8px", borderBottom:`1px solid ${COLORS.border}` }}>Debe</th>
                    <th style={{ padding:"8px", borderBottom:`1px solid ${COLORS.border}` }}>Haber</th>
                    <th style={{ padding:"8px", borderBottom:`1px solid ${COLORS.border}`, textAlign:"right" }}>Acción</th>
                  </tr>
                </thead>
                <tbody>
                  {filas.map((fila, index) => (
                    <tr key={index}>
                      <td style={{ padding:"7px", borderBottom:`1px solid ${COLORS.border}` }}>
                        <input type="date" value={fila.fecha} onChange={(e) => setCampoFila(index, "fecha", e.target.value)} style={tableInputStyle} />
                      </td>
                      <td style={{ padding:"7px", borderBottom:`1px solid ${COLORS.border}` }}>
                        <input list="cuentas-contables" value={fila.detalle} onChange={(e) => setCampoFila(index, "detalle", e.target.value)} placeholder="BANCO" style={tableInputStyle} />
                      </td>
                      <td style={{ padding:"7px", borderBottom:`1px solid ${COLORS.border}` }}>
                        <input value={fila.desglose} onChange={(e) => setCampoFila(index, "desglose", e.target.value)} placeholder="Compra MDF" style={tableInputStyle} />
                      </td>
                      <td style={{ padding:"7px", borderBottom:`1px solid ${COLORS.border}` }}>
                        <input value={fila.documento} onChange={(e) => setCampoFila(index, "documento", e.target.value)} placeholder="F-123" style={tableInputStyle} />
                      </td>
                      <td style={{ padding:"7px", borderBottom:`1px solid ${COLORS.border}` }}>
                        <input value={fila.definicion} onChange={(e) => setCampoFila(index, "definicion", e.target.value)} placeholder="MATERIALES" style={tableInputStyle} />
                      </td>
                      <td style={{ padding:"7px", borderBottom:`1px solid ${COLORS.border}` }}>
                        <input type="number" value={fila.debe} onChange={(e) => setCampoFila(index, "debe", e.target.value)} placeholder="0" style={tableInputStyle} />
                      </td>
                      <td style={{ padding:"7px", borderBottom:`1px solid ${COLORS.border}` }}>
                        <input type="number" value={fila.haber} onChange={(e) => setCampoFila(index, "haber", e.target.value)} placeholder="0" style={tableInputStyle} />
                      </td>
                      <td style={{ padding:"7px", borderBottom:`1px solid ${COLORS.border}`, textAlign:"right", whiteSpace:"nowrap" }}>
                        <button onClick={() => copiarFilaAnterior(index)} disabled={index === 0} title="Copiar datos de la fila anterior" style={{ marginRight:6, padding:"7px 8px", borderRadius:7, border:`1px solid ${COLORS.border}`, background:index === 0 ? COLORS.bg : COLORS.surface, color:index === 0 ? COLORS.muted : COLORS.text, cursor:index === 0 ? "not-allowed" : "pointer" }}>
                          Copiar ↑
                        </button>
                        <button onClick={() => eliminarFila(index)} style={{ padding:"7px 8px", borderRadius:7, border:"none", background:COLORS.danger, color:"#fff", cursor:"pointer" }}>
                          🗑
                        </button>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
              <datalist id="cuentas-contables">
                {cuentasRapidas.map(cuenta => <option key={cuenta} value={cuenta} />)}
              </datalist>
            </div>

            <div style={{ display:"flex", justifyContent:"space-between", gap:12, alignItems:"center", flexWrap:"wrap", marginTop:14 }}>
              <button onClick={agregarFila} style={{ padding:"9px 12px", borderRadius:8, border:`1px solid ${COLORS.border}`, background:COLORS.surface, color:COLORS.text, fontWeight:700, cursor:"pointer" }}>
                + Agregar fila
              </button>

              <div style={{ display:"flex", gap:10, flexWrap:"wrap", justifyContent:"flex-end" }}>
                <div style={{ background:COLORS.surface, border:`1px solid ${COLORS.border}`, borderRadius:10, padding:"9px 12px" }}>
                  <div style={{ fontSize:10, color:COLORS.muted }}>Total Debe</div>
                  <div style={{ fontWeight:700, color:COLORS.success }}>{fmt(totalDebe)}</div>
                </div>
                <div style={{ background:COLORS.surface, border:`1px solid ${COLORS.border}`, borderRadius:10, padding:"9px 12px" }}>
                  <div style={{ fontSize:10, color:COLORS.muted }}>Total Haber</div>
                  <div style={{ fontWeight:700, color:COLORS.warning }}>{fmt(totalHaber)}</div>
                </div>
                <div style={{ background:COLORS.surface, border:`1px solid ${Math.abs(diferencia) === 0 ? COLORS.border : COLORS.danger}`, borderRadius:10, padding:"9px 12px" }}>
                  <div style={{ fontSize:10, color:COLORS.muted }}>Diferencia</div>
                  <div style={{ fontWeight:700, color:Math.abs(diferencia) === 0 ? COLORS.success : COLORS.danger }}>{fmt(diferencia)}</div>
                </div>
              </div>
            </div>
          </>
        )}

        <div style={{ marginTop:16, display:"flex", justifyContent:"flex-end", gap:10, flexWrap:"wrap" }}>
          <button onClick={onClose} style={{ padding:"10px 14px", borderRadius:8, border:`1px solid ${COLORS.border}`, background:COLORS.surface, color:COLORS.text, cursor:"pointer" }}>
            Cancelar
          </button>
          <button
            onClick={guardar}
            disabled={!asiento?.id && !estaCuadrado}
            style={{
              padding:"10px 16px",
              borderRadius:8,
              border:"none",
              background:!asiento?.id && !estaCuadrado ? COLORS.surface : COLORS.accent,
              color:!asiento?.id && !estaCuadrado ? COLORS.muted : "#111",
              fontWeight:700,
              cursor:!asiento?.id && !estaCuadrado ? "not-allowed" : "pointer"
            }}
          >
            {asiento?.id ? "Guardar asiento" : "Guardar asiento cuadrado"}
          </button>
        </div>
      </div>
    </div>
  );
}

function GestionRapidaContabilidadModal({ tipo, onClose, onSave }) {
  const hoy = new Date().toISOString().slice(0, 10);
  const [form, setForm] = useState({
    fecha: hoy,
    nombre: "",
    documento: "",
    definicion: "",
    descripcion: "",
    total: "",
    formaPago: "Banco",
    estado: "Contado",
    origen: "Banco",
    destino: "Caja"
  });

  useEffect(() => {
    const cerrarConEsc = (e) => {
      if (e.key === "Escape") onClose();
    };
    window.addEventListener("keydown", cerrarConEsc);
    return () => window.removeEventListener("keydown", cerrarConEsc);
  }, [onClose]);

  const setCampo = (campo, valor) => setForm(prev => ({ ...prev, [campo]: valor }));

  const inputStyle = {
    width:"100%",
    boxSizing:"border-box",
    background:COLORS.bg,
    border:`1px solid ${COLORS.border}`,
    color:COLORS.text,
    borderRadius:8,
    padding:"10px 12px"
  };

  const labelStyle = { display:"block", fontSize:11, color:COLORS.muted, marginBottom:6, textTransform:"uppercase", letterSpacing:0.8 };

  const titulos = {
    compra: "Registrar compra",
    venta: "Registrar venta",
    pago_cliente: "Pago de cliente",
    pago_proveedor: "Pago a proveedor",
    movimiento: "Movimiento caja/banco"
  };

  const total = Math.round(Number(form.total || 0));
  const neto = Math.round(total / 1.19);
  const iva = total - neto;

  const construirAsientos = () => {
    if (!form.fecha) {
      alert("Debes ingresar una fecha.");
      return null;
    }
    if (total <= 0) {
      alert("Debes ingresar un monto mayor a cero.");
      return null;
    }

    const nombre = form.nombre.trim().toUpperCase();
    const documento = form.documento.trim();
    const definicion = form.definicion.trim().toUpperCase();
    const descripcion = form.descripcion.trim();
    const desgloseBase = [nombre, descripcion].filter(Boolean).join(" - ");
    const base = { fecha: form.fecha, documento, definicion, desglose: desgloseBase };

    if (tipo === "compra") {
      const pago = form.formaPago.toUpperCase();
      const asientos = [
        { ...base, detalle:"COMPRAS / MATERIALES", debe: neto, haber: 0 },
        { ...base, detalle:"IVA CF", debe: iva, haber: 0 },
        { ...base, detalle:"PROVEEDORES", debe: 0, haber: total }
      ];

      if (form.formaPago !== "Pendiente") {
        asientos.push(
          { ...base, detalle:"PROVEEDORES", debe: total, haber: 0 },
          { ...base, detalle:pago, debe: 0, haber: total }
        );
      }

      return asientos;
    }

    if (tipo === "venta") {
      const cuentaCobro = form.estado === "Crédito" ? "CLIENTES / CXC" : form.formaPago.toUpperCase();
      return [
        { ...base, detalle:cuentaCobro, debe: total, haber: 0 },
        { ...base, detalle:"VENTAS", debe: 0, haber: neto },
        { ...base, detalle:"IVA DF", debe: 0, haber: iva }
      ];
    }

    if (tipo === "pago_cliente") {
      return [
        { ...base, detalle:form.formaPago.toUpperCase(), debe: total, haber: 0 },
        { ...base, detalle:"CLIENTES / CXC", debe: 0, haber: total }
      ];
    }

    if (tipo === "pago_proveedor") {
      return [
        { ...base, detalle:"PROVEEDORES", debe: total, haber: 0 },
        { ...base, detalle:form.formaPago.toUpperCase(), debe: 0, haber: total }
      ];
    }

    if (tipo === "movimiento") {
      if (form.origen === form.destino) {
        alert("El origen y destino no pueden ser iguales.");
        return null;
      }
      return [
        { ...base, detalle:form.destino.toUpperCase(), debe: total, haber: 0 },
        { ...base, detalle:form.origen.toUpperCase(), debe: 0, haber: total }
      ];
    }

    return null;
  };

  const guardar = () => {
    const asientos = construirAsientos();
    if (!asientos) return;
    onSave(asientos);
  };

  const vistaPrevia = construirAsientosNoAlert();

  function construirAsientosNoAlert() {
    if (!total || total <= 0) return [];
    const nombre = form.nombre.trim().toUpperCase();
    const documento = form.documento.trim();
    const definicion = form.definicion.trim().toUpperCase();
    const descripcion = form.descripcion.trim();
    const desgloseBase = [nombre, descripcion].filter(Boolean).join(" - ");
    const base = { fecha: form.fecha, documento, definicion, desglose: desgloseBase };
    if (tipo === "compra") {
      const pago = form.formaPago.toUpperCase();
      const arr = [
        { ...base, detalle:"COMPRAS / MATERIALES", debe: neto, haber: 0 },
        { ...base, detalle:"IVA CF", debe: iva, haber: 0 },
        { ...base, detalle:"PROVEEDORES", debe: 0, haber: total }
      ];
      if (form.formaPago !== "Pendiente") arr.push({ ...base, detalle:"PROVEEDORES", debe: total, haber: 0 }, { ...base, detalle:pago, debe: 0, haber: total });
      return arr;
    }
    if (tipo === "venta") return [
      { ...base, detalle:form.estado === "Crédito" ? "CLIENTES / CXC" : form.formaPago.toUpperCase(), debe: total, haber: 0 },
      { ...base, detalle:"VENTAS", debe: 0, haber: neto },
      { ...base, detalle:"IVA DF", debe: 0, haber: iva }
    ];
    if (tipo === "pago_cliente") return [
      { ...base, detalle:form.formaPago.toUpperCase(), debe: total, haber: 0 },
      { ...base, detalle:"CLIENTES / CXC", debe: 0, haber: total }
    ];
    if (tipo === "pago_proveedor") return [
      { ...base, detalle:"PROVEEDORES", debe: total, haber: 0 },
      { ...base, detalle:form.formaPago.toUpperCase(), debe: 0, haber: total }
    ];
    if (tipo === "movimiento") return [
      { ...base, detalle:form.destino.toUpperCase(), debe: total, haber: 0 },
      { ...base, detalle:form.origen.toUpperCase(), debe: 0, haber: total }
    ];
    return [];
  }

  return (
    <div onClick={onClose} style={{ position:"fixed", inset:0, background:"rgba(0,0,0,0.75)", display:"flex", alignItems:"flex-start", justifyContent:"center", zIndex:135, overflowY:"auto", padding:"20px 10px" }}>
      <div onClick={(e) => e.stopPropagation()} style={{ background:COLORS.card, border:`1px solid ${COLORS.border}`, borderRadius:16, padding:22, width:760, maxWidth:"95vw", boxSizing:"border-box" }}>
        <div style={{ display:"flex", justifyContent:"space-between", gap:12, alignItems:"flex-start", marginBottom:16 }}>
          <div>
            <h3 style={{ margin:"0 0 4px", color:COLORS.accent, fontFamily:"Georgia,serif", fontSize:18 }}>{titulos[tipo] || "Gestión contable"}</h3>
            <div style={{ fontSize:12, color:COLORS.muted }}>Ingresas una gestión y la app genera los asientos automáticamente.</div>
          </div>
          <button onClick={onClose} style={{ background:"transparent", border:`1px solid ${COLORS.border}`, color:COLORS.muted, borderRadius:8, padding:"6px 10px", cursor:"pointer" }}>✕</button>
        </div>

        <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fit,minmax(180px,1fr))", gap:12 }}>
          <div><label style={labelStyle}>Fecha</label><input type="date" value={form.fecha} onChange={(e) => setCampo("fecha", e.target.value)} style={inputStyle} /></div>
          <div><label style={labelStyle}>{tipo === "compra" || tipo === "pago_proveedor" ? "Proveedor" : tipo === "movimiento" ? "Responsable / detalle" : "Cliente"}</label><input value={form.nombre} onChange={(e) => setCampo("nombre", e.target.value)} placeholder="Nombre" style={inputStyle} /></div>
          <div><label style={labelStyle}>NV / Factura</label><input value={form.documento} onChange={(e) => setCampo("documento", e.target.value)} placeholder="Ej: F1234 / NV 7500" style={inputStyle} /></div>
          <div><label style={labelStyle}>Definición</label><input value={form.definicion} onChange={(e) => setCampo("definicion", e.target.value)} placeholder="Ej: MATERIALES, COCINA" style={inputStyle} /></div>
        </div>

        <div style={{ marginTop:12 }}><label style={labelStyle}>Descripción</label><input value={form.descripcion} onChange={(e) => setCampo("descripcion", e.target.value)} placeholder="Ej: Compra de MDF, venta cubierta, abono cliente..." style={inputStyle} /></div>

        <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fit,minmax(180px,1fr))", gap:12, marginTop:12 }}>
          <div><label style={labelStyle}>Monto IVA incluido</label><input type="number" value={form.total} onChange={(e) => setCampo("total", e.target.value)} placeholder="0" style={inputStyle} /></div>
          {(tipo === "compra" || tipo === "venta" || tipo === "pago_cliente" || tipo === "pago_proveedor") && (
            <div><label style={labelStyle}>Forma de pago</label><select value={form.formaPago} onChange={(e) => setCampo("formaPago", e.target.value)} style={inputStyle}>
              <option>Banco</option><option>Caja</option>{tipo === "compra" && <option>Pendiente</option>}
            </select></div>
          )}
          {tipo === "venta" && (
            <div><label style={labelStyle}>Estado</label><select value={form.estado} onChange={(e) => setCampo("estado", e.target.value)} style={inputStyle}><option>Contado</option><option>Crédito</option></select></div>
          )}
          {tipo === "movimiento" && (<>
            <div><label style={labelStyle}>Origen</label><select value={form.origen} onChange={(e) => setCampo("origen", e.target.value)} style={inputStyle}><option>Banco</option><option>Caja</option></select></div>
            <div><label style={labelStyle}>Destino</label><select value={form.destino} onChange={(e) => setCampo("destino", e.target.value)} style={inputStyle}><option>Caja</option><option>Banco</option></select></div>
          </>)}
        </div>

        {(tipo === "compra" || tipo === "venta") && total > 0 && (
          <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fit,minmax(150px,1fr))", gap:10, marginTop:12 }}>
            <StatCard label="Neto" value={fmt(neto)} icon="📄" color={COLORS.success}/>
            <StatCard label={tipo === "compra" ? "IVA CF" : "IVA DF"} value={fmt(iva)} icon="🧾" color={COLORS.accent}/>
            <StatCard label="Total" value={fmt(total)} icon="💰" color={COLORS.warning}/>
          </div>
        )}

        <div style={{ marginTop:16, background:COLORS.surface, border:`1px solid ${COLORS.border}`, borderRadius:12, padding:12 }}>
          <div style={{ fontSize:13, fontWeight:700, color:COLORS.accent, marginBottom:8 }}>Vista previa de asientos</div>
          {vistaPrevia.length === 0 ? <div style={{ fontSize:12, color:COLORS.muted }}>Ingresa un monto para ver los asientos que se crearán.</div> : (
            <div style={{ overflowX:"auto" }}><table style={{ width:"100%", borderCollapse:"collapse", fontSize:12, minWidth:520 }}>
              <thead><tr style={{ color:COLORS.muted, textAlign:"left" }}><th style={{ padding:7, borderBottom:`1px solid ${COLORS.border}` }}>Cuenta</th><th style={{ padding:7, borderBottom:`1px solid ${COLORS.border}`, textAlign:"right" }}>Debe</th><th style={{ padding:7, borderBottom:`1px solid ${COLORS.border}`, textAlign:"right" }}>Haber</th></tr></thead>
              <tbody>{vistaPrevia.map((a, i) => <tr key={i}><td style={{ padding:7, borderBottom:`1px solid ${COLORS.border}`, color:COLORS.accent, fontWeight:700 }}>{a.detalle}</td><td style={{ padding:7, borderBottom:`1px solid ${COLORS.border}`, textAlign:"right" }}>{a.debe ? fmt(a.debe) : "-"}</td><td style={{ padding:7, borderBottom:`1px solid ${COLORS.border}`, textAlign:"right" }}>{a.haber ? fmt(a.haber) : "-"}</td></tr>)}</tbody>
            </table></div>
          )}
        </div>

        <div style={{ marginTop:16, display:"flex", justifyContent:"flex-end", gap:10, flexWrap:"wrap" }}>
          <button onClick={onClose} style={{ padding:"10px 14px", borderRadius:8, border:`1px solid ${COLORS.border}`, background:COLORS.surface, color:COLORS.text, cursor:"pointer" }}>Cancelar</button>
          <button onClick={guardar} style={{ padding:"10px 16px", borderRadius:8, border:"none", background:COLORS.accent, color:"#111", fontWeight:700, cursor:"pointer" }}>Crear asientos</button>
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

  await crearAsientoVentaNVAutomatica({
    fecha: excelDateToISO(fecha),
    numero: notaVenta,
    cliente,
    totalVenta: total
  });

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
const { data: dataVentasLaminas, error: errorVentasLaminas } = await supabase
  .from("ventas_laminas")
  .select("*")
  .order("fecha", { ascending: false });

if (!errorVentasLaminas) {
  setVentasLaminas(dataVentasLaminas || []);
}

const { data: dataDetallesVentasLaminas, error: errorDetallesVentasLaminas } = await supabase
  .from("detalles_ventas_laminas")
  .select("*")
  .order("orden", { ascending: true });

if (!errorDetallesVentasLaminas) {
  setDetallesVentasLaminas(dataDetallesVentasLaminas || []);
}
const { data: dataAsientosContables, error: errorAsientosContables } = await supabase
  .from("asientos_contables")
  .select("*")
  .order("fecha", { ascending: false })
  .order("created_at", { ascending: false });

if (!errorAsientosContables) {
  setAsientosContables(dataAsientosContables || []);
}
  }

  cargarDatos();
}, []);
  const [tab, setTab] = useState("dashboard");
  const [filter, setFilter] = useState("");
  const [fechaDesde, setFechaDesde] = useState("");
  const [fechaHasta, setFechaHasta] = useState("");
  const [mesFiltro, setMesFiltro] = useState(new Date().toISOString().slice(0, 7));
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
  const [ventasLaminas, setVentasLaminas] = useState([]);
  const [detallesVentasLaminas, setDetallesVentasLaminas] = useState([]);
  const [modalVentaLaminas, setModalVentaLaminas] = useState(null);
  const [modalCxcClientes, setModalCxcClientes] = useState(false);
  const [mesVentaLaminas, setMesVentaLaminas] = useState(new Date().toISOString().slice(0, 7));
  const [topLaminasCantidad, setTopLaminasCantidad] = useState(10);
  const [asientosContables, setAsientosContables] = useState([]);
  const [modalContabilidad, setModalContabilidad] = useState(null);
  const [modalGestionContable, setModalGestionContable] = useState(null);
  const [mesContabilidad, setMesContabilidad] = useState(new Date().toISOString().slice(0, 7));
  const [busquedaContabilidad, setBusquedaContabilidad] = useState("");
  

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

const obtenerMesFecha = (fecha) => {
  if (!fecha) return "";

  const texto = String(fecha).trim();
  const match = texto.match(/^(\d{4})-(\d{2})/);
  if (match) return `${match[1]}-${match[2]}`;

  const fechaObj = new Date(fecha);
  if (Number.isNaN(fechaObj.getTime())) return "";

  const year = fechaObj.getFullYear();
  const month = String(fechaObj.getMonth() + 1).padStart(2, "0");
  return `${year}-${month}`;
};

const nombreMes = (mes) => {
  if (!mes) return "Todos los meses";
  const [year, month] = mes.split("-");
  const nombre = new Date(Number(year), Number(month) - 1, 1)
    .toLocaleDateString("es-CL", { month:"long", year:"numeric" });
  return nombre.charAt(0).toUpperCase() + nombre.slice(1);
};

const cumpleMes = (fecha) => {
  if (!mesFiltro) return true;
  return obtenerMesFecha(fecha) === mesFiltro;
};
  const limpiarFiltros = () => {
  setSearch("");
  setFechaDesde("");
  setFechaHasta("");
  setMesFiltro(new Date().toISOString().slice(0, 7));
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

const totalSoldFiltrado = filteredNotas.reduce((s,n) => s + Number(n.total || 0), 0);
const mesSeleccionadoTexto = nombreMes(mesFiltro);

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

  if (nuevoAbono > 0) {
    await crearAsientoAbonoNVAutomatico({
      fecha: new Date().toISOString().split("T")[0],
      numero: nvActualizada.numero,
      cliente: nvActualizada.cliente,
      monto: nuevoAbono
    });
  }

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
const calcularNetoIvaDesdeTotal = (totalConIva) => {
  const total = Math.round(Number(totalConIva || 0));
  const neto = Math.round(total / 1.19);
  const iva = total - neto;
  return { total, neto, iva };
};

const guardarAsientosAutomaticos = async (asientos, opciones = {}) => {
  const payload = asientos
    .map(a => ({
      fecha: a.fecha || new Date().toISOString().split("T")[0],
      detalle: a.detalle || "",
      desglose: a.desglose || "",
      documento: a.documento || "",
      definicion: a.definicion || "",
      debe: Number(a.debe || 0),
      haber: Number(a.haber || 0)
    }))
    .filter(a => a.detalle && (a.debe > 0 || a.haber > 0));

  if (payload.length === 0) return true;

  const totalDebe = payload.reduce((sum, a) => sum + Number(a.debe || 0), 0);
  const totalHaber = payload.reduce((sum, a) => sum + Number(a.haber || 0), 0);

  if (totalDebe !== totalHaber) {
    alert("No se generó el asiento automático porque Debe y Haber no cuadran.");
    return false;
  }

  if (opciones.reemplazarDocumento && opciones.reemplazarDesglose) {
    await supabase
      .from("asientos_contables")
      .delete()
      .eq("documento", opciones.reemplazarDocumento)
      .eq("desglose", opciones.reemplazarDesglose);

    setAsientosContables(prev => prev.filter(a =>
      !(a.documento === opciones.reemplazarDocumento && a.desglose === opciones.reemplazarDesglose)
    ));
  }

  const { data, error } = await supabase
    .from("asientos_contables")
    .insert(payload)
    .select();

  if (error) {
    console.error(error);
    alert("La operación se guardó, pero no se pudo crear el asiento contable automático.");
    return false;
  }

  setAsientosContables(prev => [...(data || []), ...prev]);
  return true;
};

const crearAsientoCompraAutomatica = async ({ fecha, documento, proveedor, totalCompra, estadoPago, detalleBase }) => {
  const { total, neto, iva } = calcularNetoIvaDesdeTotal(totalCompra);
  if (!total) return true;

  const doc = documento || `COMPRA-${Date.now()}`;
  const detalle = detalleBase || `Compra ${proveedor || "proveedor"}`;
  const cuentaHaber = estadoPago === "pendiente" ? "Proveedores" : "Banco";

  return guardarAsientosAutomaticos([
    { fecha, detalle, desglose:"Compra automática", documento:doc, definicion:"Compras", debe:neto, haber:0 },
    { fecha, detalle, desglose:"Compra automática", documento:doc, definicion:"IVA CF", debe:iva, haber:0 },
    { fecha, detalle, desglose:"Compra automática", documento:doc, definicion:cuentaHaber, debe:0, haber:total }
  ]);
};

const crearAsientoVentaNVAutomatica = async ({ fecha, numero, cliente, totalVenta }) => {
  const { total, neto, iva } = calcularNetoIvaDesdeTotal(totalVenta);
  if (!total) return true;

  const documento = `NV-${numero}`;

  return guardarAsientosAutomaticos([
    { fecha, detalle:`Venta NV ${numero} - ${cliente || "cliente"}`, desglose:"Venta nota de venta", documento, definicion:"Clientes", debe:total, haber:0 },
    { fecha, detalle:`Venta NV ${numero} - ${cliente || "cliente"}`, desglose:"Venta nota de venta", documento, definicion:"Ventas", debe:0, haber:neto },
    { fecha, detalle:`Venta NV ${numero} - ${cliente || "cliente"}`, desglose:"Venta nota de venta", documento, definicion:"IVA DF", debe:0, haber:iva }
  ], {
    reemplazarDocumento: documento,
    reemplazarDesglose: "Venta nota de venta"
  });
};

const crearAsientoAbonoNVAutomatico = async ({ fecha, numero, cliente, monto }) => {
  const total = Math.round(Number(monto || 0));
  if (!total) return true;

  const documento = `ABONO-NV-${numero}-${Date.now()}`;

  return guardarAsientosAutomaticos([
    { fecha, detalle:`Abono NV ${numero} - ${cliente || "cliente"}`, desglose:"Abono cliente", documento, definicion:"Banco", debe:total, haber:0 },
    { fecha, detalle:`Abono NV ${numero} - ${cliente || "cliente"}`, desglose:"Abono cliente", documento, definicion:"Clientes", debe:0, haber:total }
  ]);
};

const crearAsientoVentaLaminasAutomatica = async ({ fecha, numero, cliente, totalVenta }) => {
  const { total, neto, iva } = calcularNetoIvaDesdeTotal(totalVenta);
  if (!total) return true;

  const documento = `VL-${numero}`;

  return guardarAsientosAutomaticos([
    { fecha, detalle:`Venta láminas ${numero} - ${cliente || "cliente"}`, desglose:"Venta láminas", documento, definicion:"Banco", debe:total, haber:0 },
    { fecha, detalle:`Venta láminas ${numero} - ${cliente || "cliente"}`, desglose:"Venta láminas", documento, definicion:"Ventas", debe:0, haber:neto },
    { fecha, detalle:`Venta láminas ${numero} - ${cliente || "cliente"}`, desglose:"Venta láminas", documento, definicion:"IVA DF", debe:0, haber:iva }
  ], {
    reemplazarDocumento: documento,
    reemplazarDesglose: "Venta láminas"
  });
};

const crearAsientoCostoVentaLaminasAutomatico = async ({ fecha, numero, cliente, costoCompra }) => {
  const { total, neto, iva } = calcularNetoIvaDesdeTotal(costoCompra);
  if (!total) return true;

  const documento = `VL-${numero}`;

  return guardarAsientosAutomaticos([
    { fecha, detalle:`Costo láminas ${numero} - ${cliente || "cliente"}`, desglose:"Costo venta láminas", documento, definicion:"Compras", debe:neto, haber:0 },
    { fecha, detalle:`Costo láminas ${numero} - ${cliente || "cliente"}`, desglose:"Costo venta láminas", documento, definicion:"IVA CF", debe:iva, haber:0 },
    { fecha, detalle:`Costo láminas ${numero} - ${cliente || "cliente"}`, desglose:"Costo venta láminas", documento, definicion:"Banco", debe:0, haber:total }
  ], {
    reemplazarDocumento: documento,
    reemplazarDesglose: "Costo venta láminas"
  });
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

const confirmarImportacionOC = async (datosCompra = {}) => {
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

  if (Number(datosCompra.totalCompra || 0) > 0) {
    await crearAsientoCompraAutomatica({
      fecha: datosCompra.fecha || new Date().toISOString().split("T")[0],
      documento: datosCompra.documento || `OC-${Date.now()}`,
      proveedor: datosCompra.proveedor || "Proveedor",
      totalCompra: datosCompra.totalCompra,
      estadoPago: datosCompra.estadoPago || "pagado",
      detalleBase: `Compra OC ${datosCompra.documento || "sin documento"} ${datosCompra.proveedor || ""}`.trim()
    });
  }

  setPreviewOC([]);
  setModalPreviewOC(false);
  alert("Orden de compra importada correctamente.");
};

const excelDateToISO = (value) => {
  if (!value) return new Date().toISOString().split("T")[0];

  if (typeof value === "number") {
    const parsed = XLSX.SSF.parse_date_code(value);
    if (parsed) {
      return `${parsed.y}-${String(parsed.m).padStart(2, "0")}-${String(parsed.d).padStart(2, "0")}`;
    }
  }

  const fecha = new Date(value);
  if (!isNaN(fecha.getTime())) return fecha.toISOString().split("T")[0];

  return new Date().toISOString().split("T")[0];
};

const importarVentaLaminasExcel = async (event) => {
  const file = event.target.files?.[0];
  if (!file) return;

  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, { type:"array" });
  const hojaNombre = workbook.SheetNames[0];
  const hoja = workbook.Sheets[hojaNombre];

  if (!hoja) {
    alert("No se pudo leer la primera hoja del Excel.");
    return;
  }

  const filas = XLSX.utils.sheet_to_json(hoja, { header:1, defval:"" });

  const buscarCeldaPorTexto = (textoBuscado) => {
    const buscado = textoBuscado.toUpperCase();

    for (const ref in hoja) {
      if (ref.startsWith("!")) continue;

      const valor = String(hoja[ref]?.v || "").toUpperCase().trim();

      if (valor.includes(buscado)) {
        return ref;
      }
    }

    return null;
  };

  const leerCelda = (ref) => hoja[ref]?.v || "";

  const numeroRef = buscarCeldaPorTexto("N°") || buscarCeldaPorTexto("Nº");
  const numeroTexto = numeroRef
    ? String(leerCelda(numeroRef)).replace("N°", "").replace("Nº", "").trim()
    : "";

  const numero = numeroTexto || String(file.name || "").split(" ")[0];
  const cliente = String(hoja["C10"]?.v || "").trim() || "(sin cliente)";
  const fecha = excelDateToISO(hoja["G10"]?.v);

  let totalVenta = 0;

  for (let i = 0; i < filas.length; i++) {
    const fila = filas[i];

    for (let j = 0; j < fila.length; j++) {
      const texto = String(fila[j] || "").toUpperCase().trim();

      if (texto === "TOTAL") {
        const posibleTotal = Number(fila[j + 1]) || 0;

        if (posibleTotal > 0) {
          totalVenta = Math.round(posibleTotal);
        }
      }
    }
  }

  if (!numero || !totalVenta) {
    alert("No se pudo leer el número o total de la cotización de láminas.");
    event.target.value = "";
    return;
  }

  const duplicada = ventasLaminas.some(v => String(v.numero) === String(numero));
  if (duplicada) {
    alert("Esta venta de láminas ya existe.");
    event.target.value = "";
    return;
  }

  let filaEncabezado = -1;
  let colCantidad = -1;
  let colTipo = -1;
  let colLargo = -1;
  let colAncho = -1;
  let colColor = -1;
  let colValor = -1;
  let colTotal = -1;

  for (let i = 0; i < filas.length; i++) {
    const fila = filas[i].map(c => String(c).trim().toUpperCase());

    const cantidadIndex = fila.findIndex(c => c.includes("STOCK") || c.includes("CANTIDAD"));
    const tipoIndex = fila.findIndex(c => c.includes("TIPO"));
    const largoIndex = fila.findIndex(c => c.includes("LARGO"));
    const anchoIndex = fila.findIndex(c => c.includes("ANCHO"));
    const colorIndex = fila.findIndex(c => c.includes("COLOR"));
    const valorIndex = fila.findIndex(c => c.includes("VALOR"));
    const totalIndex = fila.findIndex(c => c.includes("TOTAL"));

    if (cantidadIndex !== -1 && tipoIndex !== -1 && colorIndex !== -1 && totalIndex !== -1) {
      filaEncabezado = i;
      colCantidad = cantidadIndex;
      colTipo = tipoIndex;
      colLargo = largoIndex;
      colAncho = anchoIndex;
      colColor = colorIndex;
      colValor = valorIndex;
      colTotal = totalIndex;
      break;
    }
  }

  if (filaEncabezado === -1) {
    alert("No se encontró el detalle de láminas en el Excel.");
    event.target.value = "";
    return;
  }

  const detalles = [];

  for (let i = filaEncabezado + 1; i < filas.length; i++) {
    const fila = filas[i];
    const textoFila = fila.join(" ").toUpperCase();

    if (textoFila.includes("NETO") || textoFila.includes("I.V.A") || textoFila.includes("IVA") || textoFila.includes("TRANSFERIR")) break;

    const cantidad = Number(fila[colCantidad]) || 0;
    const tipo = String(fila[colTipo] || "").trim();
    const color = String(fila[colColor] || "").trim();
    const total = Math.round(Number(fila[colTotal]) || 0);

    if (!cantidad && !tipo && !color && !total) continue;
    if (!cantidad || !color) continue;

    detalles.push({
      cantidad,
      tipo,
      largo: Number(fila[colLargo]) || 0,
      ancho: Number(fila[colAncho]) || 0,
      color,
      valor_unitario: Math.round(Number(fila[colValor]) || 0),
      total,
      orden: detalles.length + 1
    });
  }

  const { data: ventaInsertada, error: errorVenta } = await supabase
    .from("ventas_laminas")
    .insert([{
      numero,
      cliente,
      fecha,
      total_venta: totalVenta,
      costo_compra_total: 0
    }])
    .select();

  if (errorVenta) {
    console.error(errorVenta);
    alert("Error al guardar la venta de láminas. Revisa si las tablas están creadas en Supabase.");
    event.target.value = "";
    return;
  }

  const nuevaVenta = ventaInsertada[0];

  if (detalles.length > 0) {
    const detallesParaGuardar = detalles.map(d => ({
      venta_lamina_id: nuevaVenta.id,
      cantidad: d.cantidad,
      tipo: d.tipo,
      largo: d.largo,
      ancho: d.ancho,
      color: d.color,
      valor_unitario: d.valor_unitario,
      total: d.total,
      orden: d.orden
    }));

    const { data: detallesInsertados, error: errorDetalles } = await supabase
      .from("detalles_ventas_laminas")
      .insert(detallesParaGuardar)
      .select();

    if (errorDetalles) {
      console.error(errorDetalles);
      alert("La venta se guardó, pero hubo un error guardando el detalle.");
      event.target.value = "";
      return;
    }

    setDetallesVentasLaminas(prev => [...prev, ...(detallesInsertados || [])]);
  }

  setVentasLaminas(prev => [nuevaVenta, ...prev]);

  await crearAsientoVentaLaminasAutomatica({
    fecha,
    numero,
    cliente,
    totalVenta
  });

  setTab("venta_laminas");
  event.target.value = "";
  alert("Venta de láminas importada correctamente.");
};

const guardarCostoVentaLaminas = async (venta, costoCompraTotal) => {
  const { data, error } = await supabase
    .from("ventas_laminas")
    .update({ costo_compra_total: costoCompraTotal })
    .eq("id", venta.id)
    .select();

  if (error) {
    console.error(error);
    alert("Error al guardar el costo de compra.");
    return;
  }

  const actualizada = data[0];
  setVentasLaminas(prev => prev.map(v => v.id === actualizada.id ? actualizada : v));

  await crearAsientoCostoVentaLaminasAutomatico({
    fecha: actualizada.fecha || new Date().toISOString().split("T")[0],
    numero: actualizada.numero,
    cliente: actualizada.cliente,
    costoCompra: costoCompraTotal
  });

  setModalVentaLaminas(actualizada);
};


const guardarAsientoContable = async (asiento) => {
  const payload = {
    fecha: asiento.fecha,
    detalle: asiento.detalle,
    desglose: asiento.desglose || "",
    documento: asiento.documento || "",
    definicion: asiento.definicion || "",
    debe: Number(asiento.debe || 0),
    haber: Number(asiento.haber || 0)
  };

  if (asiento.id) {
    const { data, error } = await supabase
      .from("asientos_contables")
      .update(payload)
      .eq("id", asiento.id)
      .select();

    if (error) {
      console.error(error);
      alert("No se pudo actualizar el asiento contable.");
      return;
    }

    const actualizado = data[0];
    setAsientosContables(prev => prev.map(a => a.id === actualizado.id ? actualizado : a));
  } else {
    const { data, error } = await supabase
      .from("asientos_contables")
      .insert([payload])
      .select();

    if (error) {
      console.error(error);
      alert("No se pudo guardar el asiento contable. Revisa que la tabla exista en Supabase.");
      return;
    }

    setAsientosContables(prev => [data[0], ...prev]);
  }

  setModalContabilidad(null);
};


const guardarGestionContable = async (asientos) => {
  const payload = asientos.map(a => ({
    fecha: a.fecha,
    detalle: a.detalle,
    desglose: a.desglose || "",
    documento: a.documento || "",
    definicion: a.definicion || "",
    debe: Number(a.debe || 0),
    haber: Number(a.haber || 0)
  }));

  const totalDebe = payload.reduce((sum, a) => sum + Number(a.debe || 0), 0);
  const totalHaber = payload.reduce((sum, a) => sum + Number(a.haber || 0), 0);

  if (Math.abs(totalDebe - totalHaber) !== 0) {
    alert("Los asientos no cuadran. Revisa los montos antes de guardar.");
    return;
  }

  const { data, error } = await supabase
    .from("asientos_contables")
    .insert(payload)
    .select();

  if (error) {
    console.error(error);
    alert("No se pudieron guardar los asientos automáticos.");
    return;
  }

  setAsientosContables(prev => [...data, ...prev]);
  setModalGestionContable(null);
  alert("Gestión contable registrada correctamente.");
};

const eliminarAsientoContable = async (asiento) => {
  const confirmar = window.confirm("¿Seguro que quieres eliminar este asiento contable?\n\nEsta acción no se puede deshacer.");
  if (!confirmar) return;

  const { error } = await supabase
    .from("asientos_contables")
    .delete()
    .eq("id", asiento.id);

  if (error) {
    console.error(error);
    alert("No se pudo eliminar el asiento contable.");
    return;
  }

  setAsientosContables(prev => prev.filter(a => a.id !== asiento.id));
};

const asientosContablesFiltrados = asientosContables.filter(a => {
  const cumpleMesContabilidad = !mesContabilidad || obtenerMesFecha(a.fecha) === mesContabilidad;
  const texto = `${a.detalle || ""} ${a.desglose || ""} ${a.documento || ""} ${a.definicion || ""}`.toLowerCase();
  const cumpleBusqueda = !busquedaContabilidad || texto.includes(busquedaContabilidad.toLowerCase());
  return cumpleMesContabilidad && cumpleBusqueda;
});

const resumenContabilidad = asientosContablesFiltrados.reduce((acc, a) => {
  const detalle = String(a.detalle || "").toUpperCase().trim();
  const definicion = String(a.definicion || "").toUpperCase().trim();
  const debe = Number(a.debe || 0);
  const haber = Number(a.haber || 0);

  acc.debe += debe;
  acc.haber += haber;

  if (definicion === "IVA CF" || definicion.includes("IVA CREDITO") || definicion.includes("IVA CRÉDITO")) acc.ivaCf += debe - haber;
  if (definicion === "IVA DF" || definicion.includes("IVA DEBITO") || definicion.includes("IVA DÉBITO")) acc.ivaDf += haber - debe;
  if (definicion === "BANCO" || definicion.includes("BANCO")) acc.banco += debe - haber;
  if (definicion === "CAJA" || definicion.includes("CAJA")) acc.caja += debe - haber;
  if (definicion === "CLIENTES" || definicion.includes("CLIENTE") || definicion.includes("CXC") || definicion.includes("CUENTAS X COBRAR")) acc.cxc += debe - haber;

  return acc;
}, { debe:0, haber:0, ivaCf:0, ivaDf:0, banco:0, caja:0, cxc:0 });

const diferenciaContabilidad = resumenContabilidad.debe - resumenContabilidad.haber;
const detalleCxcNotas = notas
  .filter(n => !mesContabilidad || obtenerMesFecha(n.fecha) === mesContabilidad)
  .map(n => {
    const idNota = Number(String(n.id).replace("supabase-", ""));

    const total = Number(n.total || 0);

    const abonadoDesdeNota = Number(n.abono || n.abonado || 0);

    const abonadoDesdeHistorial = abonosNV
      .filter(a => Number(a.nota_venta_id) === idNota)
      .reduce((sum, a) => sum + Number(a.monto || 0), 0);

    const abonado = Math.max(abonadoDesdeNota, abonadoDesdeHistorial);

    const saldo = Math.max(total - abonado, 0);

    return {
      numero: n.numero || n.nota_venta || n.id,
      cliente: n.cliente || "Sin cliente",
      saldo
    };
  })
  .filter(n => n.saldo > 0);

const totalCxcNotas = detalleCxcNotas.reduce((sum, n) => sum + Number(n.saldo || 0), 0);
const eliminarVentaLaminas = async (venta) => {
  const confirmar = window.confirm(
    `¿Seguro que quieres eliminar la venta de láminas #${venta.numero}?

Esta acción no se puede deshacer.`
  );

  if (!confirmar) return;

  const { error } = await supabase
    .from("ventas_laminas")
    .delete()
    .eq("id", venta.id);

  if (error) {
    console.error(error);
    alert("No se pudo eliminar la venta de láminas.");
    return;
  }

  setVentasLaminas(prev => prev.filter(v => v.id !== venta.id));
  setDetallesVentasLaminas(prev => prev.filter(d => d.venta_lamina_id !== venta.id));

  if (modalVentaLaminas?.id === venta.id) {
    setModalVentaLaminas(null);
  }

  alert("Venta de láminas eliminada correctamente.");
};

const ventasLaminasDelMes = ventasLaminas.filter(v => {
  if (!mesVentaLaminas) return true;
  return obtenerMesFecha(v.fecha) === mesVentaLaminas;
});

const resumenVentasLaminasMes = ventasLaminasDelMes.reduce((acc, venta) => {
  const r = calcularResumenVentaLaminas(venta);
  acc.venta += r.ventaTotal;
  acc.costo += r.costoTotal;
  acc.utilidad += r.utilidadNeta;
  acc.iva += r.ivaProvisionar;
  return acc;
}, { venta:0, costo:0, utilidad:0, iva:0 });

const rankingLaminas = Object.values(
  detallesVentasLaminas
    .filter(d => ventasLaminasDelMes.some(v => v.id === d.venta_lamina_id))
    .reduce((acc, d) => {
      const color = String(d.color || "Sin color").trim().toUpperCase();
      if (!acc[color]) acc[color] = { color, cantidad:0, total:0 };
      acc[color].cantidad += Number(d.cantidad || 0);
      acc[color].total += Number(d.total || 0);
      return acc;
    }, {})
)
  .sort((a,b) => b.cantidad - a.cantidad)
  .slice(0, topLaminasCantidad);

const maxRankingCantidad = Math.max(...rankingLaminas.map(r => r.cantidad), 1);

  const tabs=[
    {key:"dashboard",label:"📊 Resumen"},
    {key:"quotes",label:`📋 Cotizaciones (${filteredQuotes.length})`},
    {key:"sales",label:`✅ Notas de Venta (${filteredNotas.length})`},
    {key:"sinmatch",label:`⚠ Sin cruzar (${sinCotizacion.length})`},
    {key:"produccion",label:`🏭 Producción`},
    {key:"inventario",label:`📦 Inventario`},
    {key:"venta_laminas",label:`🧾 Venta de Láminas (${ventasLaminasDelMes.length})`},
    {key:"contabilidad",label:`📚 Contabilidad (${asientosContablesFiltrados.length})`},
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
<label
  style={{
    padding: "10px 18px",
    background: COLORS.warning,
    borderRadius: 8,
    fontWeight: 700,
    cursor: "pointer",
    color: "#111",
    display: "inline-block",
    marginLeft: 10
  }}
>
  Importar Venta Láminas
  <input
    type="file"
    accept=".xlsx,.xls"
    onChange={importarVentaLaminasExcel}
    style={{ display: "none" }}
  />
</label>
</div>
      

      <div style={{ display:"flex", gap:2, padding:"12px 24px 0", borderBottom:`1px solid ${COLORS.border}`, overflowX:"auto" }}>
        {tabs.map(t=>(
          <button key={t.key} onClick={()=>setTab(t.key)} style={{ background:tab===t.key?COLORS.card:"transparent", border:`1px solid ${tab===t.key?COLORS.border:"transparent"}`, borderBottom:tab===t.key?`2px solid ${COLORS.accent}`:"2px solid transparent", borderRadius:"8px 8px 0 0", padding:"8px 16px", color:tab===t.key?COLORS.accent:COLORS.muted, cursor:"pointer", fontSize:12, fontWeight:tab===t.key?700:400, whiteSpace:"nowrap" }}>{t.label}</button>
        ))}
      </div>

      <div style={{ padding:"clamp(12px, 4vw, 22px)", maxWidth:1100, margin:"0 auto" }}>
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
              <h2 style={{ margin:0, fontFamily:"Georgia,serif", color:COLORS.accent, fontSize:17 }}>Cotizaciones {mesSeleccionadoTexto}</h2>
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

  {[...new Set([
    new Date().toISOString().slice(0, 7),
    ...[...cotizaciones, ...notas]
      .map(item => item.fecha)
      .filter(Boolean)
      .map(fecha => obtenerMesFecha(fecha))
      .filter(Boolean)
  ])]
    .sort((a, b) => b.localeCompare(a))
    .map(mes => (
      <option key={mes} value={mes}>{nombreMes(mes)}</option>
    ))}
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
            <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", gap:10, flexWrap:"wrap", marginBottom:14 }}>
              <h2 style={{ margin:0, fontFamily:"Georgia,serif", color:COLORS.success, fontSize:17 }}>Notas de Venta {mesSeleccionadoTexto}</h2>
              <input
                type="month"
                value={mesFiltro}
                onChange={(e) => setMesFiltro(e.target.value)}
                style={{ background:COLORS.surface, border:`1px solid ${COLORS.border}`, color:COLORS.text, borderRadius:8, padding:"8px 10px" }}
              />
            </div>
            <div style={{ marginBottom:14, background:COLORS.subtle, borderRadius:10, padding:"10px 16px", display:"flex", gap:24 }}>
              <div><span style={{ fontSize:11, color:COLORS.muted }}>TOTAL VENDIDO</span><div style={{ fontSize:18, fontWeight:700, color:COLORS.success }}>{fmt(totalSoldFiltrado)}</div></div>
              <div><span style={{ fontSize:11, color:COLORS.muted }}>NOTAS</span><div style={{ fontSize:18, fontWeight:700, color:COLORS.success }}>{filteredNotas.length}</div></div>
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

        {tab==="venta_laminas" && (
          <div>
            <div style={{ display:"flex", justifyContent:"space-between", gap:12, alignItems:"flex-end", flexWrap:"wrap", marginBottom:16 }}>
              <div>
                <h2 style={{ margin:"0 0 6px", fontFamily:"Georgia,serif", color:COLORS.accent, fontSize:17 }}>🧾 Venta de Láminas</h2>
                <p style={{ margin:0, fontSize:12, color:COLORS.muted }}>Cotizaciones de láminas revendidas, utilidad neta e IVA a provisionar.</p>
              </div>
              <div>
                <label style={{ display:"block", fontSize:10, color:COLORS.muted, marginBottom:5 }}>Mes</label>
                <input
                  type="month"
                  value={mesVentaLaminas}
                  onChange={(e) => setMesVentaLaminas(e.target.value)}
                  style={{ background:COLORS.surface, border:`1px solid ${COLORS.border}`, color:COLORS.text, borderRadius:8, padding:"8px 10px" }}
                />
              </div>
            </div>

            <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fit,minmax(150px,1fr))", gap:12, marginBottom:18 }}>
              <StatCard label="Ventas del mes" value={fmt(resumenVentasLaminasMes.venta)} sub={`${ventasLaminasDelMes.length} registros`} icon="💰" color={COLORS.success}/>
              <StatCard label="Costos del mes" value={fmt(resumenVentasLaminasMes.costo)} icon="🧾" color={COLORS.warning}/>
              <StatCard label="Utilidad neta" value={fmt(resumenVentasLaminasMes.utilidad)} icon="📈" color={resumenVentasLaminasMes.utilidad >= 0 ? COLORS.success : COLORS.danger}/>
              <StatCard label="IVA a provisionar" value={fmt(resumenVentasLaminasMes.iva)} icon="🏛️" color={resumenVentasLaminasMes.iva >= 0 ? COLORS.accent : COLORS.danger}/>
            </div>

            <div style={{ display:"grid", gridTemplateColumns:"minmax(0,1.2fr) minmax(280px,0.8fr)", gap:16 }}>
              <div style={{ background:COLORS.card, border:`1px solid ${COLORS.border}`, borderRadius:12, padding:14 }}>
                <div style={{ fontSize:13, fontWeight:700, color:COLORS.accent, marginBottom:10 }}>Ventas importadas</div>
                {ventasLaminasDelMes.length === 0 ? (
                  <div style={{ color:COLORS.muted, fontSize:12 }}>No hay ventas de láminas para este mes.</div>
                ) : (
                  <div style={{ display:"flex", flexDirection:"column", gap:8 }}>
                    {ventasLaminasDelMes.map(v => {
                      const r = calcularResumenVentaLaminas(v);
                      return (
                        <div
                          key={v.id}
                          onClick={() => setModalVentaLaminas(v)}
                          style={{ background:COLORS.surface, border:`1px solid ${COLORS.border}`, borderLeft:`4px solid ${COLORS.accent}`, borderRadius:10, padding:"11px 12px", cursor:"pointer", display:"flex", justifyContent:"space-between", gap:10, flexWrap:"wrap" }}
                        >
                          <div>
                            <div style={{ fontWeight:700, color:COLORS.text }}>#{v.numero} · {v.cliente}</div>
                            <div style={{ fontSize:11, color:COLORS.muted }}>{fmtDate(v.fecha)} · Costo: {r.costoTotal > 0 ? fmt(r.costoTotal) : "pendiente"}</div>
                          </div>
                          <div style={{ textAlign:"right", display:"flex", flexDirection:"column", alignItems:"flex-end", gap:6 }}>
                            <div>
                              <div style={{ fontWeight:700, color:COLORS.success }}>{fmt(r.ventaTotal)}</div>
                              <div style={{ fontSize:11, color:r.utilidadNeta >= 0 ? COLORS.accent : COLORS.danger }}>Utilidad: {fmt(r.utilidadNeta)}</div>
                            </div>
                            <button
                              onClick={(e) => {
                                e.stopPropagation();
                                eliminarVentaLaminas(v);
                              }}
                              style={{ background: COLORS.danger, border: "none", borderRadius: 6, padding: "5px 9px", color: "#fff", fontSize: 11, fontWeight: 700, cursor: "pointer" }}
                            >
                              Eliminar
                            </button>
                          </div>
                        </div>
                      );
                    })}
                  </div>
                )}
              </div>

              <div style={{ background:COLORS.card, border:`1px solid ${COLORS.border}`, borderRadius:12, padding:14 }}>
                <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", gap:8, marginBottom:12 }}>
                  <div style={{ fontSize:13, fontWeight:700, color:COLORS.accent }}>Colores más vendidos</div>
                  <select
                    value={topLaminasCantidad}
                    onChange={(e) => setTopLaminasCantidad(Number(e.target.value))}
                    style={{ background:COLORS.surface, border:`1px solid ${COLORS.border}`, color:COLORS.text, borderRadius:8, padding:"6px 8px", fontSize:12 }}
                  >
                    <option value={10}>Top 10</option>
                    <option value={50}>Top 50</option>
                  </select>
                </div>

                {rankingLaminas.length === 0 ? (
                  <div style={{ color:COLORS.muted, fontSize:12 }}>Todavía no hay detalle para generar ranking.</div>
                ) : (
                  <div style={{ display:"flex", flexDirection:"column", gap:9 }}>
                    {rankingLaminas.map((r, idx) => (
                      <div key={r.color}>
                        <div style={{ display:"flex", justifyContent:"space-between", gap:8, fontSize:12, marginBottom:4 }}>
                          <span style={{ color:COLORS.text, fontWeight:700 }}>{idx + 1}. {r.color}</span>
                          <span style={{ color:COLORS.accent }}>{r.cantidad} láminas</span>
                        </div>
                        <div style={{ height:8, background:COLORS.surface, borderRadius:20, overflow:"hidden", border:`1px solid ${COLORS.border}` }}>
                          <div style={{ height:"100%", width:`${Math.max(4, (r.cantidad / maxRankingCantidad) * 100)}%`, background:COLORS.accent }} />
                        </div>
                        <div style={{ fontSize:10, color:COLORS.muted, marginTop:3 }}>Vendido: {fmt(r.total)}</div>
                      </div>
                    ))}
                  </div>
                )}
              </div>
            </div>
          </div>
        )}


        {tab==="contabilidad" && (
          <div>
            <div style={{ display:"flex", justifyContent:"space-between", gap:12, alignItems:"flex-end", flexWrap:"wrap", marginBottom:16 }}>
              <div>
                <h2 style={{ margin:"0 0 6px", fontFamily:"Georgia,serif", color:COLORS.accent, fontSize:17 }}>📚 Contabilidad</h2>
                <p style={{ margin:0, fontSize:12, color:COLORS.muted }}>Libro diario simple con Debe, Haber, IVA CF/DF, Banco, Caja y CxC Clientes.</p>
              </div>
              <div style={{ display:"flex", gap:8, flexWrap:"wrap", alignItems:"flex-end" }}>
                <div>
                  <label style={{ display:"block", fontSize:10, color:COLORS.muted, marginBottom:5 }}>Mes</label>
                  <input
                    type="month"
                    value={mesContabilidad}
                    onChange={(e) => setMesContabilidad(e.target.value)}
                    style={{ background:COLORS.surface, border:`1px solid ${COLORS.border}`, color:COLORS.text, borderRadius:8, padding:"8px 10px" }}
                  />
                </div>
                <button onClick={() => setModalGestionContable("compra")} style={{ background:COLORS.accent, color:"#111", border:"none", borderRadius:8, padding:"10px 12px", fontWeight:700, cursor:"pointer" }}>+ Compra</button>
                <button onClick={() => setModalGestionContable("venta")} style={{ background:COLORS.success, color:"#fff", border:"none", borderRadius:8, padding:"10px 12px", fontWeight:700, cursor:"pointer" }}>+ Venta</button>
                <button onClick={() => setModalGestionContable("pago_cliente")} style={{ background:COLORS.surface, color:COLORS.text, border:`1px solid ${COLORS.border}`, borderRadius:8, padding:"10px 12px", fontWeight:700, cursor:"pointer" }}>+ Pago cliente</button>
                <button onClick={() => setModalGestionContable("pago_proveedor")} style={{ background:COLORS.surface, color:COLORS.text, border:`1px solid ${COLORS.border}`, borderRadius:8, padding:"10px 12px", fontWeight:700, cursor:"pointer" }}>+ Pago proveedor</button>
                <button onClick={() => setModalGestionContable("movimiento")} style={{ background:COLORS.surface, color:COLORS.text, border:`1px solid ${COLORS.border}`, borderRadius:8, padding:"10px 12px", fontWeight:700, cursor:"pointer" }}>+ Caja/Banco</button>
                <button
                  onClick={() => setModalContabilidad({})}
                  style={{ background:"transparent", color:COLORS.muted, border:`1px solid ${COLORS.border}`, borderRadius:8, padding:"10px 12px", fontWeight:700, cursor:"pointer" }}
                >
                  + Asiento manual
                </button>
              </div>
            </div>

            <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fit,minmax(150px,1fr))", gap:12, marginBottom:14 }}>
              <StatCard label="Debe total" value={fmt(resumenContabilidad.debe)} icon="⬅️" color={COLORS.success}/>
              <StatCard label="Haber total" value={fmt(resumenContabilidad.haber)} icon="➡️" color={COLORS.warning}/>
              <StatCard label="Diferencia" value={fmt(diferenciaContabilidad)} sub={Math.abs(diferenciaContabilidad) === 0 ? "Cuadrado" : "No cuadrado"} icon="⚖️" color={Math.abs(diferenciaContabilidad) === 0 ? COLORS.success : COLORS.danger}/>
              <StatCard label="IVA CF" value={fmt(resumenContabilidad.ivaCf)} icon="🧾" color={COLORS.accent}/>
              <StatCard label="IVA DF" value={fmt(resumenContabilidad.ivaDf)} icon="🏛️" color={COLORS.accent}/>
              <StatCard label="IVA a pagar" value={fmt(resumenContabilidad.ivaDf - resumenContabilidad.ivaCf)} icon="📌" color={(resumenContabilidad.ivaDf - resumenContabilidad.ivaCf) >= 0 ? COLORS.warning : COLORS.success}/>
            </div>

            <div
  onClick={() => setModalCxcClientes(true)}
  style={{ cursor: "pointer" }}
>
  <StatCard label="CxC Clientes" value={fmt(totalCxcNotas)} icon="👥" color={COLORS.warning}/>
</div>

            <div style={{ background:COLORS.card, border:`1px solid ${COLORS.border}`, borderRadius:12, padding:14 }}>
              <div style={{ display:"flex", justifyContent:"space-between", gap:10, flexWrap:"wrap", alignItems:"center", marginBottom:12 }}>
                <div style={{ fontSize:13, fontWeight:700, color:COLORS.accent }}>Asientos del mes</div>
                <input
                  value={busquedaContabilidad}
                  onChange={(e) => setBusquedaContabilidad(e.target.value)}
                  placeholder="Buscar detalle, desglose, factura..."
                  style={{ minWidth:220, flex:"0 1 320px", background:COLORS.surface, border:`1px solid ${COLORS.border}`, color:COLORS.text, borderRadius:8, padding:"8px 10px" }}
                />
              </div>

              {asientosContablesFiltrados.length === 0 ? (
                <div style={{ color:COLORS.muted, fontSize:12 }}>No hay asientos contables para este mes.</div>
              ) : (
                <div style={{ overflowX:"auto" }}>
                  <table style={{ width:"100%", borderCollapse:"collapse", fontSize:12, minWidth:780 }}>
                    <thead>
                      <tr style={{ color:COLORS.muted, textAlign:"left" }}>
                        <th style={{ padding:"8px", borderBottom:`1px solid ${COLORS.border}` }}>Fecha</th>
                        <th style={{ padding:"8px", borderBottom:`1px solid ${COLORS.border}` }}>Detalle</th>
                        <th style={{ padding:"8px", borderBottom:`1px solid ${COLORS.border}` }}>Desglose</th>
                        <th style={{ padding:"8px", borderBottom:`1px solid ${COLORS.border}` }}>NV/Factura</th>
                        <th style={{ padding:"8px", borderBottom:`1px solid ${COLORS.border}` }}>Definición</th>
                        <th style={{ padding:"8px", borderBottom:`1px solid ${COLORS.border}`, textAlign:"right" }}>Debe</th>
                        <th style={{ padding:"8px", borderBottom:`1px solid ${COLORS.border}`, textAlign:"right" }}>Haber</th>
                        <th style={{ padding:"8px", borderBottom:`1px solid ${COLORS.border}`, textAlign:"right" }}>Acciones</th>
                      </tr>
                    </thead>
                    <tbody>
                      {asientosContablesFiltrados.map(a => (
                        <tr key={a.id}>
                          <td style={{ padding:"8px", borderBottom:`1px solid ${COLORS.border}` }}>{fmtDate(a.fecha)}</td>
                          <td style={{ padding:"8px", borderBottom:`1px solid ${COLORS.border}`, fontWeight:700, color:COLORS.accent }}>{a.detalle}</td>
                          <td style={{ padding:"8px", borderBottom:`1px solid ${COLORS.border}` }}>{a.desglose}</td>
                          <td style={{ padding:"8px", borderBottom:`1px solid ${COLORS.border}` }}>{a.documento}</td>
                          <td style={{ padding:"8px", borderBottom:`1px solid ${COLORS.border}` }}>{a.definicion}</td>
                          <td style={{ padding:"8px", borderBottom:`1px solid ${COLORS.border}`, textAlign:"right", color:COLORS.success, fontWeight:700 }}>{Number(a.debe || 0) > 0 ? fmt(a.debe) : "-"}</td>
                          <td style={{ padding:"8px", borderBottom:`1px solid ${COLORS.border}`, textAlign:"right", color:COLORS.warning, fontWeight:700 }}>{Number(a.haber || 0) > 0 ? fmt(a.haber) : "-"}</td>
                          <td style={{ padding:"8px", borderBottom:`1px solid ${COLORS.border}`, textAlign:"right" }}>
                            <button onClick={() => setModalContabilidad(a)} style={{ marginRight:6, background:COLORS.surface, color:COLORS.text, border:`1px solid ${COLORS.border}`, borderRadius:7, padding:"5px 8px", cursor:"pointer" }}>Editar</button>
                            <button onClick={() => eliminarAsientoContable(a)} style={{ background:COLORS.danger, color:"#fff", border:"none", borderRadius:7, padding:"5px 8px", cursor:"pointer" }}>Eliminar</button>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
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
{modalVentaLaminas && (
  <VentaLaminasModal
    venta={modalVentaLaminas}
    detalles={detallesVentasLaminas.filter(d => d.venta_lamina_id === modalVentaLaminas.id)}
    onClose={() => setModalVentaLaminas(null)}
    onSaveCosto={guardarCostoVentaLaminas}
  />
)}

{modalGestionContable && (
  <GestionRapidaContabilidadModal
    tipo={modalGestionContable}
    onClose={() => setModalGestionContable(null)}
    onSave={guardarGestionContable}
  />
)}
{modalContabilidad && (
  <ContabilidadModal
    asiento={modalContabilidad?.id ? modalContabilidad : null}
    onClose={() => setModalContabilidad(null)}
    onSave={modalContabilidad?.id ? guardarAsientoContable : guardarGestionContable}
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
{modalCxcClientes && (
  <div
    onClick={() => setModalCxcClientes(false)}
    style={{
      position: "fixed",
      inset: 0,
      background: "rgba(0,0,0,.45)",
      display: "flex",
      alignItems: "center",
      justifyContent: "center",
      zIndex: 9999,
      padding: 16
    }}
  >
    <div
      onClick={(e) => e.stopPropagation()}
      style={{
        background: COLORS.surface,
        borderRadius: 14,
        width: "min(720px, 96vw)",
        maxHeight: "85vh",
        overflow: "auto",
        padding: 18,
        border: `1px solid ${COLORS.border}`
      }}
    >
      <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", gap:12, marginBottom:14 }}>
        <div>
          <h2 style={{ margin:0, color:COLORS.accent }}>Detalle CxC Clientes</h2>
          <div style={{ fontSize:12, color:COLORS.muted }}>
            Notas de venta con saldo pendiente
          </div>
        </div>

        <button
          onClick={() => setModalCxcClientes(false)}
          style={{
            border:"none",
            background:COLORS.danger,
            color:"#fff",
            borderRadius:8,
            padding:"7px 10px",
            fontWeight:700,
            cursor:"pointer"
          }}
        >
          Cerrar
        </button>
      </div>

      {detalleCxcNotas.length === 0 ? (
        <div style={{ color:COLORS.muted, padding:12 }}>
          No hay notas de venta con saldo pendiente.
        </div>
      ) : (
        <div style={{ overflowX:"auto" }}>
          <table style={{ width:"100%", borderCollapse:"collapse", fontSize:13 }}>
            <thead>
              <tr style={{ background:COLORS.bg }}>
                <th style={{ textAlign:"left", padding:8, borderBottom:`1px solid ${COLORS.border}` }}>NV</th>
                <th style={{ textAlign:"left", padding:8, borderBottom:`1px solid ${COLORS.border}` }}>Cliente</th>
                <th style={{ textAlign:"right", padding:8, borderBottom:`1px solid ${COLORS.border}` }}>Saldo</th>
              </tr>
            </thead>
            <tbody>
              {detalleCxcNotas.map((n, idx) => (
                <tr key={idx}>
                  <td style={{ padding:8, borderBottom:`1px solid ${COLORS.border}` }}>{n.numero}</td>
                  <td style={{ padding:8, borderBottom:`1px solid ${COLORS.border}` }}>{n.cliente}</td>
                  <td style={{ padding:8, textAlign:"right", fontWeight:700, borderBottom:`1px solid ${COLORS.border}` }}>
                    {fmt(n.saldo)}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}

      <div style={{ marginTop:14, textAlign:"right", fontWeight:800, color:COLORS.accent }}>
        Total CxC: {fmt(totalCxcNotas)}
      </div>
    </div>
  </div>
)}
    </div>
  );
}
